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

MAJOR_CATEGORY_NOISE_PATTERNS = [
    re.compile(r"^PAGE\.?\d*$", re.IGNORECASE),
    re.compile(r"^PAGE\.\d+$", re.IGNORECASE),
    re.compile(r"^ESTIMATE$", re.IGNORECASE),
    re.compile(r"^\d{4}/\d{1,2}/\d{1,2}$"),
    re.compile(r"^\d{4}\.\d{1,2}\.\d{1,2}$"),
    re.compile(r"^\d+$"),
]


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



def group_words_by_row(words: List[Dict], tolerance: float = 4.0) -> List[List[Dict]]:
    if not words:
        return []

    words = sorted(words, key=lambda w: (((w["top"] + w["bottom"]) / 2), w["x0"]))

    rows = []
    current = []
    current_center = None

    for w in words:
        center = (w["top"] + w["bottom"]) / 2

        if current_center is None:
            current = [w]
            current_center = center
            continue

        if abs(center - current_center) <= tolerance:
            current.append(w)
            current_center = sum((x["top"] + x["bottom"]) / 2 for x in current) / len(current)
        else:
            rows.append(sorted(current, key=lambda x: x["x0"]))
            current = [w]
            current_center = center

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



def is_number_like_token(text: str) -> bool:
    t = normalize_money_token(text).replace(",", "")
    return bool(re.fullmatch(r"-?\d+(?:\.\d+)?", t))



def apply_right_edge_fallback(rec: Dict, row_words: List[Dict]) -> Dict:
    numeric_words = [w for w in sorted(row_words, key=lambda x: x["x0"]) if is_number_like_token(w["text"])]
    if len(numeric_words) < 2:
        return rec

    tail = [w["text"] for w in numeric_words[-3:]]

    if not rec["amount"] and len(tail) >= 1:
        rec["amount"] = tail[-1]
    if not rec["unit_price"] and len(tail) >= 2:
        rec["unit_price"] = tail[-2]
    if not rec["quantity"] and len(tail) >= 3:
        rec["quantity"] = tail[-3]

    return rec



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

    rec = {
        "no": clean_text(" ".join(buckets["no"])),
        "item_spec": clean_text(" ".join(buckets["item_spec"])),
        "quantity": clean_text(" ".join(buckets["quantity"])),
        "unit": clean_text(" ".join(buckets["unit"])),
        "unit_price": clean_text(" ".join(buckets["unit_price"])),
        "amount": clean_text(" ".join(buckets["amount"])),
        "raw_row": row_to_text(row_words),
        "row_words": row_words,
    }

    return apply_right_edge_fallback(rec, row_words)



def merge_split_rows(records: List[Dict], boundaries: Dict[str, Tuple[float, float]]) -> List[Dict]:
    return records



def is_placeholder_number_row(rec: Dict) -> bool:
    no = clean_text(rec["no"])
    item_spec = clean_text(rec["item_spec"])
    quantity = clean_text(rec["quantity"])
    unit = clean_text(rec["unit"])
    unit_price = clean_text(rec["unit_price"])
    amount = clean_text(rec["amount"])
    raw_row = clean_text(rec["raw_row"])

    if not no or item_spec or quantity or unit or unit_price:
        return False
    if raw_row == no:
        return True
    if amount and amount == no:
        return True
    return False



def count_numeric_fields(rec: Dict) -> int:
    count = 0
    if clean_text(rec["quantity"]):
        count += 1
    if clean_text(rec["unit_price"]):
        count += 1
    if clean_text(rec["amount"]):
        count += 1
    return count



def is_adopted_record(rec: Dict) -> bool:
    if is_placeholder_number_row(rec):
        return False

    amount = clean_text(rec["amount"])
    numeric_count = count_numeric_fields(rec)
    unit = clean_text(rec["unit"])
    item_spec = clean_text(rec["item_spec"])

    if amount and (numeric_count >= 2 or (unit and item_spec)):
        return True
    if numeric_count >= 3:
        return True
    return False



def get_candidate_horizontal_lines(page) -> List[Dict]:
    candidates = []

    for ln in getattr(page, "lines", []):
        x0 = float(ln["x0"])
        x1 = float(ln["x1"])
        top = float(ln["top"])
        bottom = float(ln["bottom"])
        if abs(bottom - top) <= 1.5 and (x1 - x0) >= 80:
            candidates.append({
                "y": (top + bottom) / 2,
                "x0": x0,
                "x1": x1,
                "width": x1 - x0,
            })

    for rc in getattr(page, "rects", []):
        x0 = float(rc["x0"])
        x1 = float(rc["x1"])
        top = float(rc["top"])
        bottom = float(rc["bottom"])
        if (x1 - x0) >= 80:
            candidates.append({"y": top, "x0": x0, "x1": x1, "width": x1 - x0})
            candidates.append({"y": bottom, "x0": x0, "x1": x1, "width": x1 - x0})

    return candidates



def get_candidate_vertical_lines(page) -> List[Dict]:
    candidates = []

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

    hlines = get_candidate_horizontal_lines(page)
    row_lines = []
    for ln in hlines:
        crosses_table = ln["x0"] <= x_min + 10 and ln["x1"] >= x_max - 10
        if crosses_table and ln["y"] >= header_bottom - 2:
            row_lines.append(ln["y"])

    body_bottom = max(ln["bottom"] for ln in usable)
    if row_lines:
        row_lines = sorted(row_lines)
        body_bottom = row_lines[-1] + 2

    return {
        "x_min": x_min - 2,
        "x_max": x_max + 2,
        "header_top": header_top,
        "body_top": header_bottom,
        "body_bottom": body_bottom,
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



def looks_like_major_category(text: str) -> bool:
    t = clean_text(text)
    if not t:
        return False
    if len(t) > 24:
        return False
    if normalize_date_str(t):
        return False
    if any(p.fullmatch(t) for p in MAJOR_CATEGORY_NOISE_PATTERNS):
        return False
    if re.search(r"PAGE|ESTIMATE", t, flags=re.IGNORECASE):
        return False
    if re.fullmatch(r"[A-Za-z0-9./:-]+", t):
        return False
    return True



def detect_major_category_outside_table(page, all_words: List[Dict], region: Dict) -> str:
    page_left_limit = page.width * 0.45
    top_limit = region["header_top"] - 2

    candidates = []
    for w in all_words:
        text = clean_text(w["text"])
        if not text:
            continue
        if w["x0"] > page_left_limit:
            continue
        if w["bottom"] > top_limit:
            continue
        if region["x_min"] <= ((w["x0"] + w["x1"]) / 2) <= region["x_max"]:
            continue
        if not looks_like_major_category(text):
            continue
        candidates.append(w)

    if not candidates:
        return ""

    rows = group_words_by_row(candidates, tolerance=4.0)
    row_texts = [row_to_text(r) for r in rows if looks_like_major_category(row_to_text(r))]
    if not row_texts:
        return ""

    row_texts = sorted(row_texts, key=lambda t: (len(t), t))
    return row_texts[0]



def is_index_like_page(page_num: int, records: List[Dict]) -> bool:
    if page_num == 1:
        return True
    if not records:
        return True

    adopted_count = sum(1 for r in records if is_adopted_record(r))
    adopted_rate = adopted_count / max(len(records), 1)
    amount_count = sum(1 for r in records if clean_text(r["amount"]))

    if adopted_rate < 0.25 and amount_count <= 2:
        return True
    return False



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
                boundaries = build_column_boundaries(header_row, table_x_max=region["x_max"])

                region_words = filter_words_in_region(all_words, region)
                body_rows = group_words_by_row(region_words)
                body_rows = [r for r in body_rows if not is_header_text(row_to_text(r))]
                body_rows = [r for r in body_rows if not is_summary_row(row_to_text(r))]

                records = [row_to_record(r, boundaries) for r in body_rows]
                records = merge_split_rows(records, boundaries)

                if is_index_like_page(page_num, records):
                    continue

                major_category = detect_major_category_outside_table(page, all_words, region)

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



def build_debug_table_region_rows(file_name: str, file_bytes: bytes) -> List[Dict]:
    debug_rows = []

    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            try:
                all_words = extract_all_words(page)
                all_rows = group_words_by_row(all_words)
                if not all_rows:
                    debug_rows.append({
                        "file_name": file_name,
                        "page": page_num,
                        "status": "no_rows",
                        "region_method": "",
                        "x_min": "",
                        "x_max": "",
                        "body_top": "",
                        "body_bottom": "",
                        "all_words": 0,
                        "region_words": 0,
                        "header_text": "",
                        "major_category": "",
                    })
                    continue

                header_row = find_header_row(all_rows)
                if header_row is None:
                    debug_rows.append({
                        "file_name": file_name,
                        "page": page_num,
                        "status": "no_header",
                        "region_method": "",
                        "x_min": "",
                        "x_max": "",
                        "body_top": "",
                        "body_bottom": "",
                        "all_words": len(all_words),
                        "region_words": 0,
                        "header_text": "",
                        "major_category": "",
                    })
                    continue

                region = detect_table_region(page, all_rows, header_row)
                boundaries = build_column_boundaries(header_row, table_x_max=region["x_max"])
                region_words = filter_words_in_region(all_words, region)
                body_rows = group_words_by_row(region_words)
                body_rows = [r for r in body_rows if not is_header_text(row_to_text(r))]
                body_rows = [r for r in body_rows if not is_summary_row(row_to_text(r))]
                records = [row_to_record(r, boundaries) for r in body_rows]

                debug_rows.append({
                    "file_name": file_name,
                    "page": page_num,
                    "status": "index_like" if is_index_like_page(page_num, records) else "ok",
                    "region_method": region.get("method", ""),
                    "x_min": round(region.get("x_min", 0), 1),
                    "x_max": round(region.get("x_max", 0), 1),
                    "body_top": round(region.get("body_top", 0), 1),
                    "body_bottom": round(region.get("body_bottom", 0), 1),
                    "all_words": len(all_words),
                    "region_words": len(region_words),
                    "header_text": row_to_text(header_row),
                    "major_category": detect_major_category_outside_table(page, all_words, region),
                })
            except Exception as e:
                debug_rows.append({
                    "file_name": file_name,
                    "page": page_num,
                    "status": f"error: {e}",
                    "region_method": "",
                    "x_min": "",
                    "x_max": "",
                    "body_top": "",
                    "body_bottom": "",
                    "all_words": "",
                    "region_words": "",
                    "header_text": "",
                    "major_category": "",
                })

    return debug_rows



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
        out["quantity"].astype(str).str.strip().ne("") |
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

debug_region = st.checkbox("表領域デバッグを表示", value=True)

if uploaded_files:
    all_rows = []
    debug_rows = []

    for uploaded_file in uploaded_files:
        all_rows.extend(process_uploaded_file(uploaded_file))

        if debug_region:
            if uploaded_file.name.lower().endswith(".pdf"):
                debug_rows.extend(build_debug_table_region_rows(uploaded_file.name, uploaded_file.getvalue()))
            elif uploaded_file.name.lower().endswith(".zip"):
                zip_bytes = uploaded_file.getvalue()
                with zipfile.ZipFile(io.BytesIO(zip_bytes)) as z:
                    for name in z.namelist():
                        if name.lower().endswith(".pdf"):
                            debug_rows.extend(build_debug_table_region_rows(name, z.read(name)))

    df = pd.DataFrame(all_rows, columns=OUTPUT_COLUMNS)

    st.subheader("抽出結果")
    st.write(f"抽出行数: {len(df):,}")
    st.dataframe(df, use_container_width=True, height=500)

    if debug_region:
        debug_df = pd.DataFrame(debug_rows)
        st.subheader("表領域デバッグ")
        st.dataframe(debug_df, use_container_width=True, height=300)

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
