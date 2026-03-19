import io
import re
import zipfile
from typing import List, Dict
import pandas as pd
import pdfplumber
import streamlit as st

st.set_page_config(page_title="見積抽出ツール", layout="wide")
st.title("見積抽出ツール")

UNIT_CANDIDATES = {
    "式","台","枚","本","箇所","箇","ｍ","m","m2","㎡","㎥",
    "日","セット","ｾｯﾄ","人工","個","ヶ所","箱","巻","丁","脚","面"
}

DATE_PATTERN = re.compile(
    r"(20\d{2})[年/\.](\d{1,2})[月/\.](\d{1,2})日?"
)

DETAIL_COLUMNS = [
    "file_name",
    "estimate_date",
    "page",
    "major_category",
    "no",
    "item_spec",
    "work_type_manual",
    "quantity",
    "unit",
    "unit_price",
    "amount",
    "raw_row",
    "needs_review",
]

def clean(t):
    return re.sub(r"\s+", " ", t.replace("\u3000"," ")).strip()

def extract_estimate_date(pdf):
    for page in pdf.pages[:2]:
        text = page.extract_text() or ""
        for m in DATE_PATTERN.finditer(text):
            y,mn,d = m.groups()
            return f"{y}-{int(mn):02d}-{int(d):02d}"
    return ""

def is_detail_page(text):
    score = 0
    if "NO." in text: score += 1
    if "単価" in text: score += 1
    if "金額" in text: score += 1
    return score >= 2

def detect_major_category(lines):
    for l in lines[:5]:
        if "工事" in l and len(l) < 25:
            return clean(re.sub(r"\d+$","",l))
    return ""

def parse_line(line):
    t = clean(line)

    if not t:
        return None

    if any(x in t for x in ["小計","御見積","消費税","PAGE"]):
        return None

    m = re.match(
        r"^(?P<no>\d+)\s+(?P<left>.+?)\s+(?P<qty>-?\d+(?:\.\d+)?)\s+(?P<unit>\S+)\s+(?P<unit_price>-?[\d,]+)\s+(?P<amount>-?[\d,]+)$",
        t
    )

    if not m:
        return {
            "no":"",
            "item_spec":"",
            "quantity":"",
            "unit":"",
            "unit_price":"",
            "amount":"",
            "raw_row":t,
            "needs_review":1
        }

    unit = clean(m.group("unit"))

    if unit not in UNIT_CANDIDATES:
        return {
            "no":"",
            "item_spec":"",
            "quantity":"",
            "unit":"",
            "unit_price":"",
            "amount":"",
            "raw_row":t,
            "needs_review":1
        }

    return {
        "no":m.group("no"),
        "item_spec":clean(m.group("left")),
        "quantity":m.group("qty"),
        "unit":unit,
        "unit_price":m.group("unit_price"),
        "amount":m.group("amount"),
        "raw_row":t,
        "needs_review":0
    }

def process_pdf(name, bytes_):
    rows = []
    with pdfplumber.open(io.BytesIO(bytes_)) as pdf:

        estimate_date = extract_estimate_date(pdf)

        for pageno,page in enumerate(pdf.pages,1):

            text = page.extract_text() or ""
            lines = [clean(x) for x in text.split("\n")]

            if not is_detail_page(text):
                continue

            major = detect_major_category(lines)

            for line in lines:
                parsed = parse_line(line)
                if not parsed:
                    continue

                rows.append({
                    "file_name":name,
                    "estimate_date":estimate_date,
                    "page":pageno,
                    "major_category":major,
                    "work_type_manual":"",
                    **parsed
                })

    return rows

def process_upload(file):
    rows = []
    if file.name.lower().endswith(".pdf"):
        rows += process_pdf(file.name, file.read())

    if file.name.lower().endswith(".zip"):
        with zipfile.ZipFile(io.BytesIO(file.read())) as z:
            for n in z.namelist():
                if n.lower().endswith(".pdf"):
                    rows += process_pdf(n, z.read(n))
    return rows

files = st.file_uploader("PDFまたはZIP", type=["pdf","zip"], accept_multiple_files=True)

if files:
    allrows = []
    for f in files:
        allrows += process_upload(f)

    df = pd.DataFrame(allrows, columns=DETAIL_COLUMNS)

    st.write("抽出行数", len(df))
    st.dataframe(df, height=600)

    st.download_button(
        "CSVダウンロード",
        df.to_csv(index=False).encode("utf-8-sig"),
        "estimate_extract.csv",
        "text/csv"
    )
    
