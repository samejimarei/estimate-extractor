import io
import re
import zipfile
import html
from typing import List, Dict, Optional, Tuple

import pandas as pd
import pdfplumber
import streamlit as st
import streamlit.components.v1 as components


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

EXCEL_COLUMNS = [
    "file_name",
    "estimate_date",
    "major_category",
    "no",
    "item_spec",
    "quantity",
    "unit",
    "unit_price",
    "amount",
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


def row_to_text(row_words: List[Dict]) -> str:
    return clean_text(" ".join(w["text"] for w in sorted(row_words, key=lambda x: x["x0"])))


def is_header_text(text: str) -> bool:
    score = sum(1 for k in HEADER_KEYWORDS if k in text)
    return score >= 4


def is_summary_row(text: str) -> bool:
    t = clean_text(text)
    if not t:
        return True
    if is_header_text(t):
        return True
    if "小計" in t or "小 計" in t or "合計" in t:
        return True
    if re.match(r"^PAGE\.", t, flags=re.IGNORECASE):
        return True
    return False


def extract_all_words(page) -> List[Dict]:
    words = page.extract_words(
        keep_blank_chars=False,
        use_text_flow=True,
        x_tolerance=2,
        y_tolerance=2,
    )

    results = []
    for w in words:
        text = clean_text(w.get("text", ""))
        if not text:
            continue
        results.append({
            "text": text,
            "x0": float(w["x0"]),
            "x1": float(w["x1"]),
            "top": float(w["top"]),
            "bottom": float(w["bottom"]),
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


def build_column_boundaries(header_row: List[Dict], table_x_max: float) -> Dict[str, Tuple[float, float]]:
    no_x0, no_x1 = get_header_pos(header_row, "NO.")
    item_x0, item_x1 = get_header_pos(header_row, "項目")
    qty_x0, qty_x1 = get_header_pos(header_row, "数量")
    unit_x0, unit_x1 = get_header_pos(header_row, "単位")
    unit_price_x0, unit_price_x1 = get_header_pos(header_row, "単価")
    amount_x0, amount_x1 = get_header_pos(header_row, "金額")

    required = [
        no_x0, no_x1, item_x0, item_x1, qty_x0, qty_x1,
        unit_x0, unit_x1, unit_price_x0, unit_price_x1, amount_x0, amount_x1
    ]
    if any(v is None for v in required):
        raise ValueError("ヘッダー位置の取得に失敗しました。")

    margin = 6

    return {
        "no": (max(0, no_x0 - margin), (item_x0 + no_x1) / 2),
        "item_spec": (max(0, item_x0 - margin), (qty_x0 + item_x1) / 2),
        "quantity": (max(0, qty_x0 - margin), (unit_x0 + qty_x1) / 2),
        "unit": (max(0, unit_x0 - margin), (unit_price_x0 + unit_x1) / 2),
        "unit_price": (max(0, unit_price_x0 - margin), (amount_x0 + unit_price_x1) / 2),
        "amount": (max(0, amount_x0 - margin), table_x_max),
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
        if col is not None:
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


def should_merge_records(cur: Dict, nxt: Dict) -> bool:
    cur_has_head = bool(clean_text(cur["no"]) or clean_text(cur["item_spec"]))
    cur_has_tail = bool(clean_text(cur["quantity"]) or clean_text(cur["unit"]) or clean_text(cur["unit_price"]) or clean_text(cur["amount"]))

    nxt_has_no = bool(clean_text(nxt["no"]))
    nxt_has_tail = bool(clean_text(nxt["quantity"]) or clean_text(nxt["unit"]) or clean_text(nxt["unit_price"]) or clean_text(nxt["amount"]))

    if not cur_has_head:
        return False
    if cur_has_tail:
        return False
    if nxt_has_no:
        return False
    if not nxt_has_tail:
        return False

    return True


def merge_two_records(cur: Dict, nxt: Dict, boundaries: Dict[str, Tuple[float, float]]) -> Dict:
    merged_words = sorted(cur["row_words"] + nxt["row_words"], key=lambda x: (x["top"], x["x0"]))
    return row_to_record(merged_words, boundaries)


def merge_split_rows(records: List[Dict], boundaries: Dict[str, Tuple[float, float]]) -> List[Dict]:
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
    return bool(
        clean_text(rec["unit"]) or
        clean_text(rec["unit_price"]) or
        clean_text(rec["amount"])
    )


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


def get_candidate_vertical_lines(page) -> List[Dict]:
    candidates = []

    # page.lines
    for ln in getattr(page, "lines", []):
        x0 = float(ln["x0"])
        x1 = float(ln["x1"])
        top = float(ln["top"])
        bottom = float(ln["bottom"])
        if abs(x1 - x0) <= 1.5 and (bottom - top) >= 40:
            candidates.append({
                "x": (x0 + x1) / 2,
                "top": top,
                "bottom": bottom,
                "height": bottom - top,
                "source": "line",
            })

    # page.rects の左右辺
    for rc in getattr(page, "rects", []):
        x0 = float(rc["x0"])
        x1 = float(rc["x1"])
        top = float(rc["top"])
        bottom = float(rc["bottom"])
        if (bottom - top) >= 40:
            candidates.append({
                "x": x0,
                "top": top,
                "bottom": bottom,
                "height": bottom - top,
                "source": "rect_left",
            })
            candidates.append({
                "x": x1,
                "top": top,
                "bottom": bottom,
                "height": bottom - top,
                "source": "rect_right",
            })

    return candidates


def cluster_x_positions(xs: List[float], tolerance: float = 2.5) -> List[float]:
    if not xs:
        return []

    xs = sorted(xs)
    groups = [[xs[0]]]

    for x in xs[1:]:
        if abs(x - groups[-1][-1]) <= tolerance:
            groups[-1].append(x)
        else:
            groups.append([x])

    return [sum(g) / len(g) for g in groups]


def detect_table_region_from_lines(page, header_row: List[Dict]) -> Optional[Dict]:
    vlines = get_candidate_vertical_lines(page)
    if not vlines:
        return None

    header_top = min(w["top"] for w in header_row)
    header_bottom = max(w["bottom"] for w in header_row)

    usable = []
    for ln in vlines:
        overlap_header = ln["top"] <= header_bottom + 5 and ln["bottom"] >= header_top - 5
        if overlap_header or ln["top"] <= header_bottom <= ln["bottom"]:
            usable.append(ln)

    if len(usable) < 4:
        return None

    clustered = cluster_x_positions([ln["x"] for ln in usable], tolerance=3.0)
    if len(clustered) < 4:
        return None

    x_min = min(clustered)
    x_max = max(clustered)

    body_tops = [ln["top"] for ln in usable]
    body_bottoms = [ln["bottom"] for ln in usable]

    return {
        "x_min": x_min - 2,
        "x_max": x_max + 2,
        "header_top": header_top,
        "body_top": header_bottom,
        "body_bottom": max(body_bottoms),
        "method": "lines",
        "line_count": len(usable),
    }


def detect_table_region_from_header(page, header_row: List[Dict]) -> Dict:
    header_top = min(w["top"] for w in header_row)
    header_bottom = max(w["bottom"] for w in header_row)

    x_min = min(w["x0"] for w in header_row) - 8
    x_max = max(w["x1"] for w in header_row) + 8

    return {
        "x_min": max(0, x_min),
        "x_max": min(page.width, x_max),
        "header_top": header_top,
        "body_top": header_bottom,
        "body_bottom": page.height - 20,
        "method": "header",
        "line_count": 0,
    }


def detect_table_region(page, rows: List[List[Dict]], header_row: List[Dict]) -> Dict:
    region = detect_table_region_from_lines(page, header_row)
    if region is not None:
        return region
    return detect_table_region_from_header(page, header_row)


def filter_words_in_region(words: List[Dict], region: Dict) -> List[Dict]:
    filtered = []
    for w in words:
        cx = (w["x0"] + w["x1"]) / 2
        cy = (w["top"] + w["bottom"]) / 2

        if region["x_min"] <= cx <= region["x_max"] and region["body_top"] <= cy <= region["body_bottom"]:
            filtered.append(w)

    return filtered


def process_pdf(file_name: str, file_bytes: bytes) -> List[Dict]:
    output_rows = []

    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        estimate_date = extract_estimate_date(pdf)

        for page_num, page in enumerate(pdf.pages, start=1):
            try:
                all_words = extract_all_words(page)
                all_rows = group_words_by_row(all_words)

                if not all_rows:
                    continue

                header_row = find_header_row(all_rows)
                if header_row is None:
                    continue

                region = detect_table_region(page, all_rows, header_row)
                region_words = filter_words_in_region(all_words, region)
                body_rows = group_words_by_row(region_words)

                if not body_rows:
                    continue

                # ヘッダー行を除外
                body_rows = [r for r in body_rows if not is_header_text(row_to_text(r))]
                body_rows = [r for r in body_rows if not is_summary_row(row_to_text(r))]

                boundaries = build_column_boundaries(header_row, table_x_max=region["x_max"])
                records = [row_to_record(r, boundaries) for r in body_rows]
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
        return pd.DataFrame(columns=EXCEL_COLUMNS)

    out = df.copy()

    for col in OUTPUT_COLUMNS:
        if col not in out.columns:
            out[col] = ""

    out = out[
        out["unit"].astype(str).str.strip().ne("") |
        out["unit_price"].astype(str).str.strip().ne("") |
        out["amount"].astype(str).str.strip().ne("")
    ].copy()

    return out[EXCEL_COLUMNS].fillna("")


def render_excel_copy_button(df_for_copy: pd.DataFrame, label: str = "Excel用コピー"):
    if df_for_copy.empty:
        return

    tsv_text = df_for_copy.to_csv(sep="\t", index=False, header=False)
    safe_tsv = html.escape(tsv_text)
    safe_label = html.escape(label)

    components.html(
        f"""
        <button id="copy-tsv-btn" style="
            background:#0e7490;
            color:white;
            border:none;
            padding:10px 16px;
            border-radius:8px;
            cursor:pointer;
            font-size:14px;
            font-weight:600;">
            {safe_label}
        </button>
        <div id="copy-status" style="margin-top:8px;font-size:13px;color:#333;"></div>

        <script>
        const btn = document.getElementById("copy-tsv-btn");
        const status = document.getElementById("copy-status");
        const text = `{safe_tsv}`;

        btn.onclick = async () => {{
            try {{
                await navigator.clipboard.writeText(text);
                status.innerText = "コピーしました。ExcelでA1セルを選んで貼り付けてください。";
            }} catch (e) {{
                status.innerText = "コピーに失敗しました。ブラウザ設定をご確認ください。";
            }}
        }};
        </script>
        """,
        height=85,
    )


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

    render_excel_copy_button(excel_df, label="Excel用コピー（A1に貼り付け）")

    csv_data = df.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        label="CSVダウンロード",
        data=csv_data,
        file_name="estimate_extract.csv",
        mime="text/csv"
    )
else:
    st.info("まずPDFまたはZIPをアップロードしてください。")
