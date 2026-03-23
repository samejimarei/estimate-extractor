import io
import re
import zipfile
from typing import List, Dict, Optional, Tuple

import pandas as pd
import pdfplumber
import streamlit as st


st.set_page_config(page_title="見積抽出ツール", layout="wide")
st.title("見積抽出ツール")
st.write("PDFまたはZIPをアップロードして、見積明細を抽出します。")


OUTPUT_COLUMNS = [
    "file_name",
    "estimate_date",
    "page",
    "major_category",
    "no",
    "item_spec",
    "quantity",
    "unit",
    "unit_price",
    "amount",
    "raw_row",
]


DATE_PATTERNS = [
    re.compile(r"(20\d{2})年(\d{1,2})月(\d{1,2})日"),
    re.compile(r"(20\d{2})/(\d{1,2})/(\d{1,2})"),
    re.compile(r"(20\d{2})\.(\d{1,2})\.(\d{1,2})"),
]

HEADER_KEYWORDS = ["NO.", "項目", "数量", "単位", "単価", "金額"]


def clean_text(text: str) -> str:
    text = str(text or "")
    text = text.replace("\u3000", " ")
    text = text.replace("\xa0", " ")
    text = re.sub(r"[ \t]+", " ", text)
    return text.strip()


def normalize_date_str(text: str) -> str:
    t = clean_text(text)
    for pattern in DATE_PATTERNS:
        m = pattern.search(t)
        if m:
            y, mth, d = m.groups()
            return f"{y}-{int(mth):02d}-{int(d):02d}"
    return ""


def extract_estimate_date(pdf) -> str:
    candidates = []

    for page in pdf.pages[:2]:
        text = page.extract_text() or ""
        for line in text.splitlines():
            line = clean_text(line)
            if not line:
                continue

            if "見積作成日" in line:
                d = normalize_date_str(line)
                if d:
                    return d

            d = normalize_date_str(line)
            if d:
                candidates.append(d)

    return candidates[0] if candidates else ""


def normalize_money_token(text: str) -> str:
    t = clean_text(text)
    if not t:
        return ""

    t = t.replace("¥", "").replace("￥", "").replace(" ", "")

    if t.startswith("(") and t.endswith(")"):
        inner = t[1:-1]
        if re.fullmatch(r"[\d,]+", inner):
            return "-" + inner

    return t


def has_money_like(text: str) -> bool:
    t = normalize_money_token(text)
    if not t:
        return False
    return bool(re.search(r"\d", t))


def extract_words_from_page(page) -> List[Dict]:
    words = page.extract_words(
        keep_blank_chars=False,
        use_text_flow=True,
        x_tolerance=2,
        y_tolerance=2,
    )

    results = []
    footer_cutoff = page.height - 24

    for w in words:
        text = clean_text(w.get("text", ""))
        if not text:
            continue

        top = float(w["top"])
        bottom = float(w["bottom"])

        if top >= footer_cutoff or bottom >= footer_cutoff:
            continue

        results.append({
            "text": text,
            "x0": float(w["x0"]),
            "x1": float(w["x1"]),
            "top": top,
            "bottom": bottom,
        })

    return results


def group_words_by_row(words: List[Dict], tolerance: float = 3.0) -> List[List[Dict]]:
    if not words:
        return []

    words = sorted(words, key=lambda w: (round(w["top"], 1), w["x0"]))

    rows = []
    current = []
    current_top = None

    for w in words:
        if current_top is None:
            current = [w]
            current_top = w["top"]
            continue

        if abs(w["top"] - current_top) <= tolerance:
            current.append(w)
        else:
            rows.append(sorted(current, key=lambda x: x["x0"]))
            current = [w]
            current_top = w["top"]

    if current:
        rows.append(sorted(current, key=lambda x: x["x0"]))

    return rows


def row_to_text(row_words: List[Dict]) -> str:
    return clean_text(" ".join(w["text"] for w in sorted(row_words, key=lambda x: x["x0"])))


def is_header_text(text: str) -> bool:
    score = sum(1 for k in HEADER_KEYWORDS if k in text)
    return score >= 4


def find_header_row(rows: List[List[Dict]]) -> Optional[List[Dict]]:
    for row in rows:
        text = row_to_text(row)
        if is_header_text(text):
            return row
    return None


def get_header_pos(header_row: List[Dict], keyword: str) -> Tuple[Optional[float], Optional[float]]:
    for w in header_row:
        if keyword in w["text"]:
            return w["x0"], w["x1"]
    return None, None


def build_column_boundaries(header_row: List[Dict]) -> Dict[str, Tuple[float, float]]:
    no_x0, no_x1 = get_header_pos(header_row, "NO.")
    item_x0, item_x1 = get_header_pos(header_row, "項目")
    qty_x0, qty_x1 = get_header_pos(header_row, "数量")
    unit_x0, unit_x1 = get_header_pos(header_row, "単位")
    unit_price_x0, unit_price_x1 = get_header_pos(header_row, "単価")
    amount_x0, amount_x1 = get_header_pos(header_row, "金額")

    required = [no_x0, no_x1, item_x0, item_x1, qty_x0, qty_x1, unit_x0, unit_x1, unit_price_x0, unit_price_x1, amount_x0, amount_x1]
    if any(v is None for v in required):
        raise ValueError("ヘッダー位置の取得に失敗しました。")

    margin = 6

    return {
        "no": (max(0, no_x0 - margin), (item_x0 + no_x1) / 2),
        "item_spec": (max(0, item_x0 - margin), (qty_x0 + item_x1) / 2),
        "quantity": (max(0, qty_x0 - margin), (unit_x0 + qty_x1) / 2),
        "unit": (max(0, unit_x0 - margin), (unit_price_x0 + unit_x1) / 2),
        "unit_price": (max(0, unit_price_x0 - margin), (amount_x0 + unit_price_x1) / 2),
        "amount": (max(0, amount_x0 - margin), 9999.0),
    }


def assign_word_to_column(word: Dict, boundaries: Dict[str, Tuple[float, float]]) -> Optional[str]:
    center_x = (word["x0"] + word["x1"]) / 2

    for col, (x_min, x_max) in boundaries.items():
        if x_min <= center_x < x_max:
            return col

    return None


def row_to_record(row_words: List[Dict], boundaries: Dict[str, Tuple[float, float]]) -> Dict:
    buckets = {
        "no": [],
        "item_spec": [],
        "quantity": [],
        "unit": [],
        "unit_price": [],
        "amount": [],
    }

    for w in row_words:
        col = assign_word_to_column(w, boundaries)
        if col:
            buckets[col].append(w["text"])

    return {
        "no": clean_text(" ".join(buckets["no"])),
        "item_spec": clean_text(" ".join(buckets["item_spec"])),
        "quantity": clean_text(" ".join(buckets["quantity"])),
        "unit": clean_text(" ".join(buckets["unit"])),
        "unit_price": clean_text(" ".join(buckets["unit_price"])),
        "amount": clean_text(" ".join(buckets["amount"])),
        "raw_row": row_to_text(row_words),
        "row_words": row_words,
    }


def is_summary_row(text: str) -> bool:
    t = clean_text(text)
    if not t:
        return True
    if is_header_text(t):
        return True
    if "小計" in t or "小 計" in t:
        return True
    if re.match(r"^PAGE\.", t, flags=re.IGNORECASE):
        return True
    return False


def should_merge_records(cur: Dict, nxt: Dict) -> bool:
    cur_has_head = bool(cur["no"] or cur["item_spec"])
    cur_has_tail = bool(cur["quantity"] or cur["unit"] or cur["unit_price"] or cur["amount"])

    nxt_has_head = bool(nxt["no"])
    nxt_has_tail = bool(nxt["quantity"] or nxt["unit"] or nxt["unit_price"] or nxt["amount"])

    if not cur_has_head:
        return False

    if cur_has_tail:
        return False

    if nxt_has_head:
        return False

    if not nxt_has_tail:
        return False

    return True


def merge_two_records(cur: Dict, nxt: Dict, boundaries: Dict[str, Tuple[float, float]]) -> Dict:
    merged_words = sorted(cur["row_words"] + nxt["row_words"], key=lambda x: x["x0"])
    return row_to_record(merged_words, boundaries)


def merge_split_rows(records: List[Dict], boundaries: Dict[str, Tuple[float, float]]) -> List[Dict]:
    if not records:
        return []

    merged = []
    i = 0

    while i < len(records):
        cur = records[i]

        if i + 1 < len(records):
            nxt = records[i + 1]
            if should_merge_records(cur, nxt):
                merged.append(merge_two_records(cur, nxt, boundaries))
                i += 2
                continue

        merged.append(cur)
        i += 1

    return merged


def is_adopted_record(rec: Dict) -> bool:
    return bool(clean_text(rec["unit"]) or clean_text(rec["unit_price"]) or clean_text(rec["amount"]))


def detect_major_category_from_page_text(page) -> str:
    text = page.extract_text() or ""
    lines = [clean_text(x) for x in text.splitlines() if clean_text(x)]

    header_idx = None
    for i, line in enumerate(lines):
        if is_header_text(line):
            header_idx = i
            break

    if header_idx is None:
        return ""

    upper_lines = lines[:header_idx]

    for line in reversed(upper_lines):
        if len(line) <= 40 and not re.search(r"[¥￥]", line) and not re.search(r"見積|株式会社|TEL|FAX|〒", line):
            return line

    return ""


def process_pdf(file_name: str, file_bytes: bytes) -> List[Dict]:
    output_rows = []

    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        estimate_date = extract_estimate_date(pdf)

        for page_num, page in enumerate(pdf.pages, start=1):
            try:
                words = extract_words_from_page(page)
                rows = group_words_by_row(words)

                if not rows:
                    continue

                header_row = find_header_row(rows)
                if header_row is None:
                    continue

                boundaries = build_column_boundaries(header_row)
                header_top = min(w["top"] for w in header_row)

                body_rows = []
                for row in rows:
                    row_top = min(w["top"] for w in row)
                    if row_top <= header_top:
                        continue

                    text = row_to_text(row)
                    if is_summary_row(text):
                        continue

                    body_rows.append(row)

                records = [row_to_record(row, boundaries) for row in body_rows]
                records = merge_split_rows(records, boundaries)

                major_category = detect_major_category_from_page_text(page)

                for rec in records:
                    if not is_adopted_record(rec):
                        continue

                    output_rows.append({
                        "file_name": file_name,
                        "estimate_date": estimate_date,
                        "page": page_num,
                        "major_category": major_category,
                        "no": clean_text(rec["no"]),
                        "item_spec": clean_text(rec["item_spec"]),
                        "quantity": clean_text(rec["quantity"]),
                        "unit": clean_text(rec["unit"]),
                        "unit_price": normalize_money_token(rec["unit_price"]),
                        "amount": normalize_money_token(rec["amount"]),
                        "raw_row": clean_text(rec["raw_row"]),
                    })

            except Exception as e:
                output_rows.append({
                    "file_name": file_name,
                    "estimate_date": estimate_date,
                    "page": page_num,
                    "major_category": "",
                    "no": "",
                    "item_spec": f"[ERROR] {e}",
                    "quantity": "",
                    "unit": "",
                    "unit_price": "",
                    "amount": "",
                    "raw_row": "",
                })

    return output_rows


def process_uploaded_file(uploaded_file) -> List[Dict]:
    rows = []

    if uploaded_file.name.lower().endswith(".pdf"):
        rows.extend(process_pdf(uploaded_file.name, uploaded_file.read()))

    elif uploaded_file.name.lower().endswith(".zip"):
        zip_bytes = uploaded_file.read()
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as z:
            for name in z.namelist():
                if name.lower().endswith(".pdf"):
                    rows.extend(process_pdf(name, z.read(name)))

    return rows


def make_excel_df(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=[
            "file_name",
            "estimate_date",
            "major_category",
            "no",
            "item_spec",
            "quantity",
            "unit",
            "unit_price",
            "amount",
        ])

    out = df.copy()

    for col in OUTPUT_COLUMNS:
        if col not in out.columns:
            out[col] = ""

    out = out[
        out["unit"].astype(str).str.strip().ne("") |
        out["unit_price"].astype(str).str.strip().ne("") |
        out["amount"].astype(str).str.strip().ne("")
    ].copy()

    return out[
        ["file_name", "estimate_date", "major_category", "no", "item_spec", "quantity", "unit", "unit_price", "amount"]
    ].fillna("")


uploaded_files = st.file_uploader(
    "PDFまたはZIPをアップロード",
    type=["pdf", "zip"],
    accept_multiple_files=True
)

if uploaded_files:
    all_rows = []

    for uploaded_file in uploaded_files:
        all_rows.extend(process_uploaded_file(uploaded_file))

    df = pd.DataFrame(all_rows, columns=OUTPUT_COLUMNS)

    st.subheader("抽出結果")
    st.write(f"抽出行数: {len(df):,}")
    st.dataframe(df, use_container_width=True, height=500)

    excel_df = make_excel_df(df)

    st.subheader("Excel貼り付け用データ")
    st.write(f"対象行数: {len(excel_df):,}")
    st.dataframe(excel_df, use_container_width=True, height=350)

    csv_data = df.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        label="CSVダウンロード",
        data=csv_data,
        file_name="estimate_extract.csv",
        mime="text/csv"
    )
else:
    st.info("まずPDFまたはZIPをアップロードしてください。")
