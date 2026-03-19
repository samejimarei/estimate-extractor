import streamlit as st
import pandas as pd

st.title("見積抽出ツール")

uploaded_files = st.file_uploader(
    "PDFまたはZIPをアップロードしてください",
    type=["pdf","zip"],
    accept_multiple_files=True
)

if uploaded_files:
    st.write(f"{len(uploaded_files)}個のファイルがアップロードされました")

    data = []
    for f in uploaded_files:
        data.append({
            "file_name": f.name,
            "item_name": "",
            "unit": "",
            "unit_price": ""
        })

    df = pd.DataFrame(data)

    st.dataframe(df)

    csv = df.to_csv(index=False).encode("utf-8-sig")

    st.download_button(
        "CSVダウンロード",
        csv,
        "result.csv",
        "text/csv"
    )
