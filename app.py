import streamlit as st
import pandas as pd
import pdfplumber
import zipfile
import io

st.set_page_config(page_title="見積抽出ツール", layout="wide")
st.title("見積抽出ツール")
st.write("PDFまたはZIPをアップロードしてください。まずは文字抽出結果をCSV化します。")

uploaded_files = st.file_uploader(
    "PDFまたはZIPをアップロード",
    type=["pdf", "zip"],
    accept_multiple_files=True
)

def extract_text_from_pdf_bytes(file_bytes: bytes):
    results = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            results.append({
                "page": i,
                "raw_text": text.strip()
            })
    return results

def process_uploaded_file(uploaded_file):
    extracted_rows = []

    if uploaded_file.name.lower().endswith(".pdf"):
        file_bytes = uploaded_file.read()
        pages = extract_text_from_pdf_bytes(file_bytes)
        for p in pages:
            extracted_rows.append({
                "file_name": uploaded_file.name,
                "page": p["page"],
                "major_category": "",
                "item_name": "",
                "spec": "",
                "quantity": "",
                "unit": "",
                "unit_price": "",
                "amount": "",
                "needs_review": "",
                "exclude_flag": "",
                "exclude_reason": "",
                "raw_text": p["raw_text"]
            })

    elif uploaded_file.name.lower().endswith(".zip"):
        zip_bytes = uploaded_file.read()
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as z:
            for name in z.namelist():
                if name.lower().endswith(".pdf"):
                    file_bytes = z.read(name)
                    pages = extract_text_from_pdf_bytes(file_bytes)
                    for p in pages:
                        extracted_rows.append({
                            "file_name": name,
                            "page": p["page"],
                            "major_category": "",
                            "item_name": "",
                            "spec": "",
                            "quantity": "",
                            "unit": "",
                            "unit_price": "",
                            "amount": "",
                            "needs_review": "",
                            "exclude_flag": "",
                            "exclude_reason": "",
                            "raw_text": p["raw_text"]
                        })

    return extracted_rows

if uploaded_files:
    all_rows = []

    for uploaded_file in uploaded_files:
        all_rows.extend(process_uploaded_file(uploaded_file))

    df = pd.DataFrame(all_rows)

    st.subheader("抽出結果プレビュー")
    st.dataframe(df, use_container_width=True, height=500)

    csv_data = df.to_csv(index=False).encode("utf-8-sig")

    st.download_button(
        label="CSVダウンロード",
        data=csv_data,
        file_name="estimate_extract_raw_text.csv",
        mime="text/csv"
    )
else:
    st.info("まずPDFまたはZIPをアップロードしてください。")
