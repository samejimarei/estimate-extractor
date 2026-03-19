import io
import re
import zipfile
import html
from typing import List, Dict, Optional

import pandas as pd
import pdfplumber
import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(page_title="見積抽出ツール", layout="wide")
st.title("見積抽出ツール")
st.write("PDFまたはZIPをアップロードして、見積明細を抽出します。")


DETAIL_COLUMNS = [
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
    "needs_review",
]

PASTE_COLUMNS = [
    "estimate_date",
    "major_category",
    "no",
    "item_spec",
    "quantity",
    "unit",
    "unit_price",
    "amount",
]

UNIT_CANDIDATES = {
    "式", "台", "枚", "本", "箇所", "箇", "ｍ", "m", "m2", "㎡", "㎥",
    "日", "セット", "ｾｯﾄ", "人工", "個", "ヶ所", "箱", "巻", "丁", "脚", "面"
}

DATE_PATTERNS = [
    re.compile(r"(20\d{2})年(\d{1,2})月(\d{1,2})日"),
    re.compile(r"(20\d{2})/(\d{1,2})/(\d{1,2})"),
    re.compile(r"(20\d{2})\.(\d{1,2})\.(\d{1,2})"),
]

HEADER_LINE_PATTERN = re.compile(r"NO\.\s+項目\s+仕様・規格/型番\s+数量\s+単位\s+単価\s+金額")
SUBCATEGORY_PATTERN = re.compile(r"^\d*\s*[\-－ー].+[\-－ー]$")
LINE_PARSE_PATTERN = re.compile(
    r"^(?P<no>\d+)\s+(?P<left>.+?)\s+(?P<qty>-?\d+(?:\.\d+)?)\s+"
    r"(?P<unit>\S+)\s+(?P<unit_price>-?[\d,]+)\s+(?P<amount>-?[\d,]+)$"
)


def clean_text(text: str) -> str:
    text = str(text or "").replace("\u3000", " ")
    text = re.sub(r"[ \t]+", " ", text)
    return text.strip()


def extract_estimate_date_from_text(text: str) -> str:
    text = text or ""

    # まず「見積作成日」付近を優先
    priority_lines = []
    for line in text.splitlines():
        t = clean_text(line)
        if "見積作成日" in t or ("見積" in t and "日" in t):
            priority_lines.append(t)

    for line in priority_lines:
        for pattern in DATE_PATTERNS:
            m = pattern.search(line)
            if m:
                y, mn, d = m.groups()
                return f"{y}-{int(mn):02d}-{int(d):02d}"

    # 見つからなければ先頭2ページから一般日付を拾う
    for pattern in DATE_PATTERNS:
        m = pattern.search(text)
        if m:
            y, mn, d = m.groups()
            return f"{y}-{int(mn):02d}-{int(d):02d}"

    return ""


def extract_estimate_date(pdf) -> str:
    text_parts = []
    for page in pdf.pages[:2]:
        page_text = page.extract_text() or ""
        text_parts.append(page_text)
    return extract_estimate_date_from_text("\n".join(text_parts))


def is_header_line(text: str) -> bool:
    t = clean_text(text)
    if not t:
        return False
    if HEADER_LINE_PATTERN.search(t):
        return True
    return all(k in t for k in ["NO.", "項目", "数量", "単位", "単価", "金額"])


def looks_like_detail_line(text: str) -> bool:
    t = clean_text(text)
    if not t:
        return False
    return LINE_PARSE_PATTERN.match(t) is not None


def is_detail_page(text: str) -> bool:
    lines = [clean_text(x) for x in (text or "").splitlines() if clean_text(x)]
    if not lines:
        return False

    header_found = any(is_header_line(line) for line in lines)
    detail_like_count = sum(1 for line in lines if looks_like_detail_line(line))

    # 明細ヘッダーがあり、明細らしい行が2件以上あれば明細ページ
    if header_found and detail_like_count >= 2:
        return True

    # ヘッダーがなくても、明細らしい行が多ければ明細ページ
    if detail_like_count >= 4:
        return True

    return False


def normalize_major_category(text: str) -> str:
    t = clean_text(text)
    t = re.sub(r"\s*[①②③④⑤⑥⑦⑧⑨⑩]$", "", t)
    t = re.sub(r"\s*\d+$", "", t)
    return t.strip()


def detect_major_category(lines: List[str]) -> str:
    candidates = []
    for line in lines[:8]:
        t = clean_text(line)
        if not t:
            continue
        if is_header_line(t):
            continue
        if "小計" in t or "小 計" in t:
            continue
        if t in {"御 見 積 書", "E S T I M A T E"}:
            continue
        if re.match(r"^\d+\s+.+\s+\d+(?:\.\d+)?\s+\S+\s+-?[\d,]+\s+-?[\d,]+$", t):
            continue
        if "工事" in t and len(t) <= 30:
            candidates.append(t)

    if not candidates:
        return ""

    return normalize_major_category(candidates[0])


def is_excluded_raw_line(text: str) -> bool:
    t = clean_text(text)
    if not t:
        return True

    exclude_patterns = [
        r"^御\s*見\s*積\s*書",
        r"^E\s*S\s*T\s*I\s*M\s*A\s*T\s*E$",
        r"^OXY株式会社$",
        r"^NO\.\s+項目\s+仕様・規格/型番\s+数量\s+単位\s+単価\s+金額$",
        r"^小\s*計",
        r"^内消費税",
        r"^御見積金額",
        r"^PAGE\.?",
        r"^\d+$",
    ]
    for pattern in exclude_patterns:
        if re.search(pattern, t):
            return True

    return False


def parse_detail_line(line: str) -> Optional[Dict]:
    t = clean_text(line)
    if not t:
        return None

    if is_header_line(t):
        return None

    if "小計" in t or "小 計" in t:
        return None

    m = LINE_PARSE_PATTERN.match(t)
    if not m:
        return {
            "no": "",
            "item_spec": "",
            "quantity": "",
            "unit": "",
            "unit_price": "",
            "amount": "",
            "raw_row": t,
            "needs_review": 1,
        }

    unit = clean_text(m.group("unit"))
    if unit not in UNIT_CANDIDATES and not re.match(r"^[A-Za-z0-9㎡㎥ｍm/]+$", unit):
        return {
            "no": "",
            "item_spec": "",
            "quantity": "",
            "unit": "",
            "unit_price": "",
            "amount": "",
            "raw_row": t,
            "needs_review": 1,
        }

    return {
        "no": clean_text(m.group("no")),
        "item_spec": clean_text(m.group("left")),
        "quantity": clean_text(m.group("qty")),
        "unit": unit,
        "unit_price": clean_text(m.group("unit_price")),
        "amount": clean_text(m.group("amount")),
        "raw_row": t,
        "needs_review": 0,
    }


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


def process_pdf(file_name: str, file_bytes: bytes) -> List[Dict]:
    rows = []

    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        estimate_date = extract_estimate_date(pdf)

        for page_num, page in enumerate(pdf.pages, start=1):
            page_text = page.extract_text() or ""
            lines = extract_lines_from_page(page)

            if not is_detail_page(page_text):
                continue

            major_category = detect_major_category(lines)

            for line in lines:
                parsed = parse_detail_line(line)
                if not parsed:
                    continue

                rows.append({
                    "file_name": file_name,
                    "estimate_date": estimate_date,
                    "page": page_num,
                    "major_category": major_category,
                    "no": parsed["no"],
                    "item_spec": parsed["item_spec"],
                    "quantity": parsed["quantity"],
                    "unit": parsed["unit"],
                    "unit_price": parsed["unit_price"],
                    "amount": parsed["amount"],
                    "raw_row": parsed["raw_row"],
                    "needs_review": parsed["needs_review"],
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


def is_subcategory_only(text: str) -> bool:
    t = clean_text(text)
    return bool(SUBCATEGORY_PATTERN.match(t))


def is_noise_row(row: pd.Series) -> bool:
    raw = clean_text(row.get("raw_row", ""))
    major = clean_text(row.get("major_category", ""))
    no = clean_text(row.get("no", ""))
    item_spec = clean_text(row.get("item_spec", ""))
    qty = clean_text(row.get("quantity", ""))
    unit = clean_text(row.get("unit", ""))
    unit_price = clean_text(row.get("unit_price", ""))
    amount = clean_text(row.get("amount", ""))
    needs_review = int(row.get("needs_review", 0)) if str(row.get("needs_review", "")).strip() else 0

    if not any([raw, no, item_spec, qty, unit, unit_price, amount]):
        return True

    if is_header_line(raw):
        return True

    if "小計" in raw or "小 計" in raw:
        return True

    if is_subcategory_only(raw):
        return True

    if raw == major and major != "":
        return True

    if re.fullmatch(r"\d+", raw):
        return True

    # 解析失敗で、数値列も取れていない行は貼付用から除外
    if needs_review == 1 and not any([qty, unit, unit_price, amount]):
        return True

    # 単文字・破片を除外
    if needs_review == 1 and len(raw) <= 3:
        return True

    return False


def make_paste_df(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=PASTE_COLUMNS)

    temp = df.copy()

    for col in DETAIL_COLUMNS:
        if col not in temp.columns:
            temp[col] = ""

    temp = temp[~temp.apply(is_noise_row, axis=1)].copy()

    # 明細として成立している行だけ残す
    temp = temp[
        (temp["no"].astype(str).str.strip() != "") &
        (temp["item_spec"].astype(str).str.strip() != "") &
        (temp["quantity"].astype(str).str.strip() != "") &
        (temp["unit"].astype(str).str.strip() != "") &
        (temp["unit_price"].astype(str).str.strip() != "") &
        (temp["amount"].astype(str).str.strip() != "")
    ].copy()

    temp = temp[PASTE_COLUMNS].fillna("").astype(str)

    return temp


def render_excel_copy_button(df_for_copy: pd.DataFrame, label: str = "Excel用コピー"):
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
        height=80,
    )


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

    st.subheader("抽出結果（検証用）")
    st.write(f"抽出行数: {len(df):,}")
    st.dataframe(df, use_container_width=True, height=500)

    paste_df = make_paste_df(df)

    st.subheader("Excel貼り付け用データ")
    st.write(f"貼り付け対象行数: {len(paste_df):,}")
    st.dataframe(paste_df, use_container_width=True, height=350)

    render_excel_copy_button(paste_df, label="Excel用コピー（A1に貼り付け）")

    csv_data = df.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        label="CSVダウンロード（検証用）",
        data=csv_data,
        file_name="estimate_extract_detail.csv",
        mime="text/csv"
    )
else:
    st.info("まずPDFまたはZIPをアップロードしてください。")
