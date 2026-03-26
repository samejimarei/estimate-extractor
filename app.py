import io
import re
import zipfile
import html
from typing import List, Dict, Optional, Tuple

import pandas as pd
import pdfplumber
import streamlit as st
import streamlit.components.v1 as components


st.set_page_config(page_title="見積抽出ツール_CellGrid版", layout="wide")
st.title("見積抽出ツール CellGrid版")
st.write("罫線からセル格子を作って明細を抽出します。表外の大分類のみ拾います。")


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

HEADER_WORD_RULES = {
    "no": ["NO.", "NO", "No.", "ＮＯ"],
    "item_spec": ["項目", "仕様", "規格", "型番"],
    "quantity": ["数量"],
    "unit": ["単位"],
    "unit_price": ["単価"],
    "amount": ["金額"],
}

MAJOR_CATEGORY_NOISE = [
    re.compile(r"^PAGE\.?\d*$", re.IGNORECASE),
    re.compile(r"^ESTIMATE$", re.IGNORECASE),
    re.compile(r"^\d{4}/\d{1,2}/\d{1,2}$"),
    re.compile(r"^\d{4}\.\d{1,2}\.\d{1,2}$"),
    re.compile(r"^\d+$"),
]


# ---------- basic utils ----------

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



def is_number_like(text: str) -> bool:
    t = normalize_money_token(text).replace(",", "")
    return bool(re.fullmatch(r"-?\d+(?:\.\d+)?", t))



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


# ---------- words / rows ----------

def extract_all_words(page) -> List[Dict]:
    words = page.extract_words(
        keep_blank_chars=False,
        use_text_flow=True,
        x_tolerance=2,
        y_tolerance=2,
    )
    out = []
    for w in words:
        text = clean_text(w.get("text", ""))
        if not text:
            continue
        out.append({
            "text": text,
            "x0": float(w["x0"]),
            "x1": float(w["x1"]),
            "top": float(w["top"]),
            "bottom": float(w["bottom"]),
        })
    return out



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



def row_to_text(row_words: List[Dict]) -> str:
    return clean_text(" ".join(w["text"] for w in sorted(row_words, key=lambda x: x["x0"])))



def looks_like_header_text(text: str) -> bool:
    hits = 0
    for rules in HEADER_WORD_RULES.values():
        if any(k in text for k in rules):
            hits += 1
    return hits >= 4



def find_header_row(all_rows: List[List[Dict]]) -> Optional[List[Dict]]:
    for row in all_rows:
        if looks_like_header_text(row_to_text(row)):
            return row
    return None


# ---------- line geometry ----------

def _unique_positions(values: List[float], tolerance: float = 3.0) -> List[float]:
    if not values:
        return []
    values = sorted(values)
    groups = [[values[0]]]
    for v in values[1:]:
        if abs(v - groups[-1][-1]) <= tolerance:
            groups[-1].append(v)
        else:
            groups.append([v])
    return [sum(g) / len(g) for g in groups]



def get_vertical_segments(page) -> List[Dict]:
    out = []
    for ln in getattr(page, "lines", []):
        x0 = float(ln["x0"])
        x1 = float(ln["x1"])
        top = float(ln["top"])
        bottom = float(ln["bottom"])
        if abs(x1 - x0) <= 1.5 and bottom - top >= 20:
            out.append({"x": (x0 + x1) / 2, "top": top, "bottom": bottom})
    for rc in getattr(page, "rects", []):
        x0 = float(rc["x0"])
        x1 = float(rc["x1"])
        top = float(rc["top"])
        bottom = float(rc["bottom"])
        if bottom - top >= 20:
            out.append({"x": x0, "top": top, "bottom": bottom})
            out.append({"x": x1, "top": top, "bottom": bottom})
    return out



def get_horizontal_segments(page) -> List[Dict]:
    out = []
    for ln in getattr(page, "lines", []):
        x0 = float(ln["x0"])
        x1 = float(ln["x1"])
        top = float(ln["top"])
        bottom = float(ln["bottom"])
        if abs(bottom - top) <= 1.5 and x1 - x0 >= 40:
            out.append({"y": (top + bottom) / 2, "x0": x0, "x1": x1})
    for rc in getattr(page, "rects", []):
        x0 = float(rc["x0"])
        x1 = float(rc["x1"])
        top = float(rc["top"])
        bottom = float(rc["bottom"])
        if x1 - x0 >= 40:
            out.append({"y": top, "x0": x0, "x1": x1})
            out.append({"y": bottom, "x0": x0, "x1": x1})
    return out



def detect_table_grid(page, header_row: List[Dict]) -> Optional[Dict]:
    header_top = min(w["top"] for w in header_row)
    header_bottom = max(w["bottom"] for w in header_row)

    v_segments = get_vertical_segments(page)
    h_segments = get_horizontal_segments(page)
    if not v_segments or not h_segments:
        return None

    header_crossing_v = [
        s for s in v_segments
        if s["top"] <= header_bottom + 4 and s["bottom"] >= header_top - 4
    ]
    if len(header_crossing_v) < 4:
        return None

    xs = _unique_positions([s["x"] for s in header_crossing_v], tolerance=3.0)
    if len(xs) < 4:
        return None

    table_x_min = min(xs)
    table_x_max = max(xs)

    y_candidates = []
    for h in h_segments:
        if h["x0"] <= table_x_min + 8 and h["x1"] >= table_x_max - 8:
            if h["y"] >= header_top - 8:
                y_candidates.append(h["y"])

    ys = _unique_positions(y_candidates, tolerance=3.0)
    ys = [y for y in ys if y >= header_top - 8]
    if len(ys) < 3:
        return None

    # keep only contiguous grid-like portion around header
    header_mid = (header_top + header_bottom) / 2
    ys = sorted(ys)

    # find header band index
    header_idx = None
    for i in range(len(ys) - 1):
        if ys[i] <= header_mid <= ys[i + 1]:
            header_idx = i
            break
    if header_idx is None:
        return None

    body_ys = ys[header_idx:]
    if len(body_ys) < 3:
        return None

    return {
        "xs": xs,
        "ys": body_ys,
        "x_min": table_x_min,
        "x_max": table_x_max,
        "header_top": header_top,
        "header_bottom": header_bottom,
        "body_top": body_ys[1],
        "body_bottom": body_ys[-1],
    }


# ---------- cell extraction ----------

def cell_bbox(xs: List[float], ys: List[float], col_idx: int, row_idx: int) -> Tuple[float, float, float, float]:
    return xs[col_idx], ys[row_idx], xs[col_idx + 1], ys[row_idx + 1]



def words_in_bbox(words: List[Dict], bbox: Tuple[float, float, float, float], pad: float = 1.5) -> List[Dict]:
    x0, y0, x1, y1 = bbox
    out = []
    for w in words:
        cx = (w["x0"] + w["x1"]) / 2
        cy = (w["top"] + w["bottom"]) / 2
        if x0 - pad <= cx <= x1 + pad and y0 - pad <= cy <= y1 + pad:
            out.append(w)
    return sorted(out, key=lambda x: (x["top"], x["x0"]))



def text_in_bbox(words: List[Dict], bbox: Tuple[float, float, float, float]) -> str:
    return clean_text(" ".join(w["text"] for w in words_in_bbox(words, bbox)))



def detect_column_map(grid: Dict, all_words: List[Dict]) -> Optional[Dict[str, int]]:
    xs = grid["xs"]
    ys = grid["ys"]
    header_cells = []
    for col_idx in range(len(xs) - 1):
        bbox = cell_bbox(xs, ys, col_idx, 0)
        header_cells.append(text_in_bbox(all_words, bbox))

    mapping = {}
    for key, rules in HEADER_WORD_RULES.items():
        best_idx = None
        best_score = -1
        for idx, text in enumerate(header_cells):
            score = sum(1 for r in rules if r in text)
            if score > best_score:
                best_score = score
                best_idx = idx
        if best_score <= 0:
            return None
        mapping[key] = best_idx

    # monotonic sanity
    ordered = [mapping[k] for k in ["no", "item_spec", "quantity", "unit", "unit_price", "amount"]]
    if ordered != sorted(ordered):
        return None
    return mapping



def build_record_from_grid_row(grid: Dict, row_idx: int, all_words: List[Dict], column_map: Dict[str, int]) -> Dict:
    xs = grid["xs"]
    ys = grid["ys"]

    cell_texts = {}
    for key, col_idx in column_map.items():
        bbox = cell_bbox(xs, ys, col_idx, row_idx)
        cell_texts[key] = text_in_bbox(all_words, bbox)

    row_bbox = (xs[0], ys[row_idx], xs[-1], ys[row_idx + 1])
    raw_row = text_in_bbox(all_words, row_bbox)

    return {
        "no": cell_texts.get("no", ""),
        "item_spec": cell_texts.get("item_spec", ""),
        "quantity": cell_texts.get("quantity", ""),
        "unit": cell_texts.get("unit", ""),
        "unit_price": cell_texts.get("unit_price", ""),
        "amount": cell_texts.get("amount", ""),
        "raw_row": raw_row,
    }


# ---------- major category ----------

def looks_like_major_category(text: str) -> bool:
    t = clean_text(text)
    if not t:
        return False
    if len(t) > 24:
        return False
    if normalize_date_str(t):
        return False
    if any(p.fullmatch(t) for p in MAJOR_CATEGORY_NOISE):
        return False
    if re.search(r"PAGE|ESTIMATE", t, flags=re.IGNORECASE):
        return False
    if re.fullmatch(r"[A-Za-z0-9./:-]+", t):
        return False
    return True



def detect_major_category(page, all_words: List[Dict], grid: Dict) -> str:
    header_top = grid["header_top"]
    left_limit = page.width * 0.45
    candidates = [
        w for w in all_words
        if w["x0"] <= left_limit and w["bottom"] <= header_top + 2
    ]
    rows = group_words_by_row(candidates, tolerance=4.0)
    row_texts = [row_to_text(r) for r in rows]
    row_texts = [t for t in row_texts if looks_like_major_category(t)]
    if not row_texts:
        return ""

    # Prefer row nearest header, then longer text
    ranked = []
    for row in rows:
        t = row_to_text(row)
        if not looks_like_major_category(t):
            continue
        row_bottom = max(w["bottom"] for w in row)
        ranked.append((abs(header_top - row_bottom), -len(t), t))
    ranked.sort()
    return ranked[0][2] if ranked else ""


# ---------- record filtering ----------

def is_summary_like(rec: Dict) -> bool:
    text = clean_text(rec["raw_row"])
    return ("小計" in text) or ("合計" in text)



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
    return raw_row == no or (amount and amount == no)



def numeric_field_count(rec: Dict) -> int:
    return sum(
        1 for k in ["quantity", "unit_price", "amount"]
        if clean_text(rec[k])
    )



def is_adopted_record(rec: Dict) -> bool:
    if is_placeholder_number_row(rec):
        return False
    if is_summary_like(rec):
        return False

    qty = clean_text(rec["quantity"])
    unit = clean_text(rec["unit"])
    unit_price = clean_text(rec["unit_price"])
    amount = clean_text(rec["amount"])
    item_spec = clean_text(rec["item_spec"])

    if amount and unit_price:
        return True
    if amount and qty and item_spec:
        return True
    if numeric_field_count(rec) >= 3:
        return True
    if amount and unit and item_spec:
        return True
    return False


# ---------- main extraction ----------

def process_pdf(file_name: str, file_bytes: bytes) -> Tuple[List[Dict], List[Dict]]:
    output_rows = []
    debug_rows = []

    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        estimate_date = extract_estimate_date(pdf)

        for page_num, page in enumerate(pdf.pages, start=1):
            try:
                all_words = extract_all_words(page)
                all_rows = group_words_by_row(all_words)
                header_row = find_header_row(all_rows)

                if header_row is None:
                    debug_rows.append({
                        "file_name": file_name,
                        "page": page_num,
                        "status": "no_header",
                        "major_category": "",
                        "grid_cols": "",
                        "grid_rows": "",
                    })
                    continue

                grid = detect_table_grid(page, header_row)
                if grid is None:
                    debug_rows.append({
                        "file_name": file_name,
                        "page": page_num,
                        "status": "no_grid",
                        "major_category": "",
                        "grid_cols": "",
                        "grid_rows": "",
                    })
                    continue

                column_map = detect_column_map(grid, all_words)
                if column_map is None:
                    debug_rows.append({
                        "file_name": file_name,
                        "page": page_num,
                        "status": "bad_column_map",
                        "major_category": "",
                        "grid_cols": len(grid["xs"]) - 1,
                        "grid_rows": len(grid["ys"]) - 1,
                    })
                    continue

                if page_num == 1:
                    debug_rows.append({
                        "file_name": file_name,
                        "page": page_num,
                        "status": "skip_page1",
                        "major_category": "",
                        "grid_cols": len(grid["xs"]) - 1,
                        "grid_rows": len(grid["ys"]) - 1,
                    })
                    continue

                major_category = detect_major_category(page, all_words, grid)

                page_records = []
                for row_idx in range(1, len(grid["ys"]) - 1):
                    rec = build_record_from_grid_row(grid, row_idx, all_words, column_map)
                    if is_adopted_record(rec):
                        page_records.append(rec)

                if not page_records:
                    debug_rows.append({
                        "file_name": file_name,
                        "page": page_num,
                        "status": "no_records",
                        "major_category": major_category,
                        "grid_cols": len(grid["xs"]) - 1,
                        "grid_rows": len(grid["ys"]) - 1,
                    })
                    continue

                for rec in page_records:
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

                debug_rows.append({
                    "file_name": file_name,
                    "page": page_num,
                    "status": "ok",
                    "major_category": major_category,
                    "grid_cols": len(grid["xs"]) - 1,
                    "grid_rows": len(grid["ys"]) - 1,
                })

            except Exception as e:
                debug_rows.append({
                    "file_name": file_name,
                    "page": page_num,
                    "status": f"error: {e}",
                    "major_category": "",
                    "grid_cols": "",
                    "grid_rows": "",
                })

    return output_rows, debug_rows


# ---------- ui ----------

def make_excel_df(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=EXCEL_COLUMNS)
    out = df.copy()
    for col in OUTPUT_COLUMNS:
        if col not in out.columns:
            out[col] = ""
    return out[EXCEL_COLUMNS].fillna("")



def render_excel_copy_button(df_for_copy: pd.DataFrame, label: str = "Excel用コピー"):
    if df_for_copy.empty:
        return
    if len(df_for_copy) > 3000:
        st.caption("行数が多いため、Excel用コピーは省略しています。CSVをご利用ください。")
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
    accept_multiple_files=True,
)

debug_region = st.checkbox("デバッグ表示", value=False)

if uploaded_files:
    all_rows = []
    all_debug = []

    for uploaded_file in uploaded_files:
        if uploaded_file.name.lower().endswith(".pdf"):
            rows, debug_rows = process_pdf(uploaded_file.name, uploaded_file.read())
            all_rows.extend(rows)
            all_debug.extend(debug_rows)

        elif uploaded_file.name.lower().endswith(".zip"):
            zip_bytes = uploaded_file.read()
            with zipfile.ZipFile(io.BytesIO(zip_bytes)) as z:
                for name in z.namelist():
                    if name.lower().endswith(".pdf"):
                        rows, debug_rows = process_pdf(name, z.read(name))
                        all_rows.extend(rows)
                        all_debug.extend(debug_rows)

    df = pd.DataFrame(all_rows, columns=OUTPUT_COLUMNS)
    st.subheader("抽出結果")
    st.write(f"抽出行数: {len(df):,}")
    st.dataframe(df, use_container_width=True, height=500)

    if debug_region:
        debug_df = pd.DataFrame(all_debug)
        st.subheader("デバッグ")
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
        file_name="estimate_extract_cellgrid.csv",
        mime="text/csv",
    )
else:
    st.info("まずPDFまたはZIPをアップロードしてください。")
