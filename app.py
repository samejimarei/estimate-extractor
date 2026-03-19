import io
import re
import zipfile
from typing import List, Dict, Optional

import pandas as pd
import pdfplumber
import streamlit as st


st.set_page_config(page_title="見積抽出ツール", layout="wide")
st.title("見積抽出ツール")
st.write("PDFまたはZIPをアップロードして、見積明細をCSV化します。")


DETAIL_COLUMNS = [
    "file_name",
    "page",
    "page_type",
    "major_category_raw",
    "major_category",
    "sub_category",
    "line_no",
    "item_name",
    "spec",
    "quantity",
    "unit",
    "unit_price",
    "amount",
    "raw_row",
    "needs_review",
    "exclude_flag",
    "exclude_reason",
]

UNIT_CANDIDATES = {
    "式", "台", "枚", "本", "箇所", "箇", "ｍ", "m", "m2", "㎡", "㎥", "日",
    "セット", "ｾｯﾄ", "人工", "個", "ヶ所", "箱", "巻", "丁", "脚", "面"
}

HEADER_KEYWORDS = [
    "NO.", "項目", "仕様・規格/型番", "数量", "単位", "単価", "金額"
]

EXCLUDE_PATTERNS = [
    r"^小\s*計",
    r"^小計",
    r"^御見積金額",
    r"^内消費税",
    r"^PAGE\.",
    r"^\d+\s*$",
    r"^E\s*S\s*T\s*I\s*M\s*A\s*T\s*E$",
]

SUMMARY_EXCLUDE_PATTERNS = [
    r"^小\s*計",
    r"^小計",
    r"^御見積金額",
    r"^内消費税",
]

SUBCATEGORY_PATTERN = re.compile(r"^[\-－ー].+[\-－ー]$")
NUMERIC_PATTERN = re.compile(r"^-?[\d,]+(?:\.\d+)?$")


def clean_text(text: str) -> str:
    text = text.replace("\u3000", " ")
    text = re.sub(r"[ \t]+", " ", text)
    return text.strip()


def normalize_major_category(text: str) -> str:
    text = clean_text(text)
    text = re.sub(r"\s*[①②③④⑤⑥⑦⑧⑨⑩]$", "", text)
    text = re.sub(r"\s*\d+$", "", text)
    return text.strip()


def is_excluded_line(text: str) -> Optional[str]:
    t = clean_text(text)
    if not t:
        return "blank"
    for p in EXCLUDE_PATTERNS:
        if re.search(p, t):
            return "meta"
    if any(k in t for k in HEADER_KEYWORDS):
        return "header"
    return None


def is_summary_excluded_line(text: str) -> Optional[str]:
    t = clean_text(text)
    if not t:
        return "blank"
    for p in SUMMARY_EXCLUDE_PATTERNS:
        if re.search(p, t):
            return "meta"
    if any(k in t for k in HEADER_KEYWORDS):
        return "header"
    return None


def extract_lines_from_page(page) -> List[str]:
    words = page.extract_words(
        keep_blank_chars=False,
        use_text_flow=True,
        x_tolerance=2,
        y_tolerance=2,
    )

    if not words:
        page_text = page.extract_text() or ""
        return [clean_text(x) for x in page_text.splitlines() if clean_text(x)]

    words = sorted(words, key=lambda w: (round(w["top"], 1), w["x0"]))

    grouped = []
    current = []
    current_top = None
    tolerance = 3

    for w in words:
        top = w["top"]
        if current_top is None:
            current = [w]
            current_top = top
            continue

        if abs(top - current_top) <= tolerance:
            current.append(w)
        else:
            grouped.append(current)
            current = [w]
            current_top = top

    if current:
        grouped.append(current)

    lines = []
    for group in grouped:
        group = sorted(group, key=lambda w: w["x0"])
        line = " ".join(w["text"] for w in group)
        line = clean_text(line)
        if line:
            lines.append(line)

    return lines


def detect_page_type(lines: List[str], page_num: int) -> str:
    joined = " ".join(lines)
    if "NO." in joined and "単価" in joined and "金額" in joined:
        return "detail"
    if page_num == 1 and "御見積書" in joined:
        return "summary"
    if page_num == 1:
        return "summary"
    return "unknown"


def detect_major_category(lines: List[str], page_type: str, page_num: int) -> str:
    if page_type == "detail":
        for line in lines[:5]:
            t = clean_text(line)
            if not t:
                continue
            if any(k in t for k in HEADER_KEYWORDS):
                continue
            if "PAGE." in t:
                continue
            if len(t) <= 30:
                return t
    if page_type == "summary" and page_num == 1:
        return ""
    return ""


def looks_like_subcategory(line: str) -> bool:
    t = clean_text(line)
    if not t:
        return False
    if SUBCATEGORY_PATTERN.match(t):
        return True
    return False


def parse_summary_line(line: str) -> Optional[Dict]:
    """
    例:
    1 解体工事 1 式 182,000 182,000
    """
    t = clean_text(line)
    if is_summary_excluded_line(t):
        return None

    m = re.match(
        r"^(?P<line_no>\d+)\s+(?P<item_name>.+?)\s+(?P<quantity>\d+(?:\.\d+)?)\s+(?P<unit>\S+)\s+(?P<unit_price>-?[\d,]+)\s+(?P<amount>-?[\d,]+)$",
        t
    )
    if not m:
        return None

    item_name = clean_text(m.group("item_name"))
    if item_name == "調整値引き":
        return {
            "line_no": m.group("line_no"),
            "item_name": item_name,
            "spec": "",
            "quantity": m.group("quantity"),
            "unit": m.group("unit"),
            "unit_price": m.group("unit_price"),
            "amount": m.group("amount"),
            "needs_review": 0,
            "exclude_flag": 1,
            "exclude_reason": "discount",
        }

    return {
        "line_no": m.group("line_no"),
        "item_name": item_name,
        "spec": "",
        "quantity": m.group("quantity"),
        "unit": m.group("unit"),
        "unit_price": m.group("unit_price"),
        "amount": m.group("amount"),
        "needs_review": 0,
        "exclude_flag": 0,
        "exclude_reason": "",
    }


def parse_detail_line(line: str) -> Optional[Dict]:
    """
    基本戦略:
    右から quantity / unit / unit_price / amount を抜く
    左側はひとまず item_name に寄せる
    """
    t = clean_text(line)
    exclude_reason = is_excluded_line(t)
    if exclude_reason:
        return {
            "parsed": False,
            "exclude_flag": 1,
            "exclude_reason": exclude_reason,
            "raw_row": t
        }

    if looks_like_subcategory(t):
        return {
            "parsed": False,
            "exclude_flag": 1,
            "exclude_reason": "subcategory_marker",
            "raw_row": t
        }

    # 基本形
    # 1 現場養生費 共用部・EV内部 1.0 式 22,000 22,000
    m = re.match(
        r"^(?P<line_no>\d+)\s+(?P<left>.+?)\s+(?P<quantity>-?\d+(?:\.\d+)?)\s+(?P<unit>\S+)\s+(?P<unit_price>-?[\d,]+)\s+(?P<amount>-?[\d,]+)$",
        t
    )

    if not m:
        return {
            "parsed": False,
            "exclude_flag": 0,
            "exclude_reason": "",
            "raw_row": t
        }

    left = clean_text(m.group("left"))
    unit = clean_text(m.group("unit"))

    if unit not in UNIT_CANDIDATES and not re.match(r"^[A-Za-z0-9㎡㎥ｍm/]+$", unit):
        return {
            "parsed": False,
            "exclude_flag": 0,
            "exclude_reason": "",
            "raw_row": t
        }

    item_name = left
    spec = ""

    if item_name == "調整値引き":
        return {
            "parsed": True,
            "line_no": m.group("line_no"),
            "item_name": item_name,
            "spec": spec,
            "quantity": m.group("quantity"),
            "unit": unit,
            "unit_price": m.group("unit_price"),
            "amount": m.group("amount"),
            "needs_review": 0,
            "exclude_flag": 1,
            "exclude_reason": "discount",
            "raw_row": t
        }

    needs_review = 0
    if len(item_name) > 40:
        needs_review = 1

    return {
        "parsed": True,
        "line_no": m.group("line_no"),
        "item_name": item_name,
        "spec": spec,
        "quantity": m.group("quantity"),
        "unit": unit,
        "unit_price": m.group("unit_price"),
        "amount": m.group("amount"),
        "needs_review": needs_review,
        "exclude_flag": 0,
        "exclude_reason": "",
        "raw_row": t
    }


def process_pdf(file_name: str, file_bytes: bytes) -> List[Dict]:
    rows = []

    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        current_subcategory = ""

        for page_num, page in enumerate(pdf.pages, start=1):
            lines = extract_lines_from_page(page)
            page_type = detect_page_type(lines, page_num)
            major_category_raw = detect_major_category(lines, page_type, page_num)
            major_category = normalize_major_category(major_category_raw) if major_category_raw else ""

            if page_type == "summary":
                for line in lines:
                    parsed = parse_summary_line(line)
                    if not parsed:
                        continue

                    rows.append({
                        "file_name": file_name,
                        "page": page_num,
                        "page_type": "summary",
                        "major_category_raw": "",
                        "major_category": "",
                        "sub_category": "",
                        "line_no": parsed["line_no"],
                        "item_name": parsed["item_name"],
                        "spec": parsed["spec"],
                        "quantity": parsed["quantity"],
                        "unit": parsed["unit"],
                        "unit_price": parsed["unit_price"],
                        "amount": parsed["amount"],
                        "raw_row": line,
                        "needs_review": parsed["needs_review"],
                        "exclude_flag": parsed["exclude_flag"],
                        "exclude_reason": parsed["exclude_reason"],
                    })

            elif page_type == "detail":
                for line in lines:
                    if looks_like_subcategory(line):
                        current_subcategory = clean_text(line).strip("-－ー")
                        continue

                    parsed = parse_detail_line(line)
                    if not parsed:
                        continue

                    if not parsed.get("parsed"):
                        # 抽出できなかったが残しておきたい行
                        reason = parsed.get("exclude_reason", "")
                        exclude_flag = parsed.get("exclude_flag", 0)

                        # header/meta系は保存しない
                        if reason in {"blank", "meta", "header", "subcategory_marker"}:
                            continue

                        rows.append({
                            "file_name": file_name,
                            "page": page_num,
                            "page_type": "detail",
                            "major_category_raw": major_category_raw,
                            "major_category": major_category,
                            "sub_category": current_subcategory,
                            "line_no": "",
                            "item_name": "",
                            "spec": "",
                            "quantity": "",
                            "unit": "",
                            "unit_price": "",
                            "amount": "",
                            "raw_row": parsed["raw_row"],
                            "needs_review": 1,
                            "exclude_flag": exclude_flag,
                            "exclude_reason": reason or "unparsed_line",
                        })
                        continue

                    rows.append({
                        "file_name": file_name,
                        "page": page_num,
                        "page_type": "detail",
                        "major_category_raw": major_category_raw,
                        "major_category": major_category,
                        "sub_category": current_subcategory,
                        "line_no": parsed["line_no"],
                        "item_name": parsed["item_name"],
                        "spec": parsed["spec"],
                        "quantity": parsed["quantity"],
                        "unit": parsed["unit"],
                        "unit_price": parsed["unit_price"],
                        "amount": parsed["amount"],
                        "raw_row": parsed["raw_row"],
                        "needs_review": parsed["needs_review"],
                        "exclude_flag": parsed["exclude_flag"],
                        "exclude_reason": parsed["exclude_reason"],
                    })

            else:
                # unknown はとりあえず全文を要確認として残す
                for line in lines:
                    reason = is_excluded_line(line)
                    if reason in {"blank", "meta", "header"}:
                        continue

                    rows.append({
                        "file_name": file_name,
                        "page": page_num,
                        "page_type": "unknown",
                        "major_category_raw": "",
                        "major_category": "",
                        "sub_category": "",
                        "line_no": "",
                        "item_name": "",
                        "spec": "",
                        "quantity": "",
                        "unit": "",
                        "unit_price": "",
                        "amount": "",
                        "raw_row": line,
                        "needs_review": 1,
                        "exclude_flag": 0,
                        "exclude_reason": "unknown_page_type",
                    })

    return rows


def process_uploaded_file(uploaded_file) -> List[Dict]:
    all_rows = []

    if uploaded_file.name.lower().endswith(".pdf"):
        file_bytes = uploaded_file.read()
        all_rows.extend(process_pdf(uploaded_file.name, file_bytes))

    elif uploaded_file.name.lower().endswith(".zip"):
        zip_bytes = uploaded_file.read()
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as z:
            for name in z.namelist():
                if name.lower().endswith(".pdf"):
                    file_bytes = z.read(name)
                    all_rows.extend(process_pdf(name, file_bytes))

    return all_rows


uploaded_files = st.file_uploader(
    "PDFまたはZIPをアップロード",
    type=["pdf", "zip"],
    accept_multiple_files=True
)

if uploaded_files:
    rows = []
    for uploaded_file in uploaded_files:
        rows.extend(process_uploaded_file(uploaded_file))

    df = pd.DataFrame(rows, columns=DETAIL_COLUMNS)

    st.subheader("抽出結果プレビュー")
    st.write(f"抽出行数: {len(df):,}")
    st.dataframe(df, use_container_width=True, height=600)

    csv_data = df.to_csv(index=False).encode("utf-8-sig")

    st.download_button(
        label="CSVダウンロード",
        data=csv_data,
        file_name="estimate_extract_detail.csv",
        mime="text/csv"
    )
else:
    st.info("まずPDFまたはZIPをアップロードしてください。")
