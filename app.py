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

UNIT_CANDIDATES = {
    "式", "台", "枚", "本", "箇所", "箇", "ｍ", "m", "m2", "㎡", "㎥",
    "日", "セット", "ｾｯﾄ", "人工", "個", "ヶ所", "箱", "巻", "丁", "脚", "面",
    "缶"
}

HEADER_KEYWORDS = ["NO.", "項目", "仕様・規格/型番", "数量", "単位", "単価", "金額"]

DATE_PATTERNS = [
    re.compile(r"(20\d{2})年(\d{1,2})月(\d{1,2})日"),
    re.compile(r"(20\d{2})/(\d{1,2})/(\d{1,2})"),
    re.compile(r"(20\d{2})\.(\d{1,2})\.(\d{1,2})"),
]

SUBCATEGORY_PATTERN = re.compile(r"^\d*\s*[\-－ー【\[].+[\-－ー】\]]\s*$")
ONLY_NUMBER_PATTERN = re.compile(r"^\d+$")
MONEY_TOKEN_PATTERN = re.compile(r"^\(?¥?-?[\d,]+\)?$")


def clean_text(text: str) -> str:
    text = str(text or "")
    text = text.replace("\u3000", " ")
    text = text.replace("\xa0", " ")
    text = re.sub(r"[ \t]+", " ", text)
    return text.strip()


def normalize_major_category(text: str) -> str:
    t = clean_text(text)
    t = re.sub(r"\s*[①②③④⑤⑥⑦⑧⑨⑩]$", "", t)
    t = re.sub(r"\s*\d+$", "", t)
    return t.strip()


def normalize_date_str(text: str) -> str:
    t = clean_text(text)
    for pattern in DATE_PATTERNS:
        m = pattern.search(t)
        if m:
            y, mn, d = m.groups()
            return f"{y}-{int(mn):02d}-{int(d):02d}"
    return ""


def extract_estimate_date(pdf) -> str:
    candidates = []

    for page in pdf.pages[:2]:
        text = page.extract_text() or ""
        for line in text.splitlines():
            t = clean_text(line)
            if not t:
                continue
            if "見積作成日" in t:
                d = normalize_date_str(t)
                if d:
                    return d
            d = normalize_date_str(t)
            if d:
                candidates.append(d)

    return candidates[0] if candidates else ""


def normalize_money_token(text: str) -> str:
    t = clean_text(text)
    if not t:
        return ""

    t = t.replace("¥", "")
    t = t.replace("￥", "")
    t = t.replace(" ", "")

    if t.startswith("(") and t.endswith(")"):
        inner = t[1:-1]
        if re.fullmatch(r"[\d,]+", inner):
            return "-" + inner

    return t


def is_money_token(text: str) -> bool:
    t = normalize_money_token(text)
    return bool(re.fullmatch(r"-?[\d,]+", t))


def is_quantity_token(text: str) -> bool:
    t = clean_text(text).replace(",", "")
    return bool(re.fullmatch(r"-?\d+(?:\.\d+)?", t))


def is_unit_token(text: str) -> bool:
    t = clean_text(text)
    if not t:
        return False
    if t in UNIT_CANDIDATES:
        return True
    if re.fullmatch(r"[A-Za-z0-9㎡㎥ｍm/]+", t):
        return True
    return False


def is_header_line(text: str) -> bool:
    t = clean_text(text)
    if not t:
        return False
    score = sum(1 for k in ["NO.", "項目", "数量", "単位", "単価", "金額"] if k in t)
    return score >= 4


def find_header_index(lines: List[str]) -> Optional[int]:
    for i, line in enumerate(lines):
        if is_header_line(line):
            return i
    return None


def is_subcategory_only(text: str) -> bool:
    return bool(SUBCATEGORY_PATTERN.match(clean_text(text)))


def is_page_title_candidate(text: str) -> bool:
    t = clean_text(text)
    if not t:
        return False
    if len(t) < 2 or len(t) > 30:
        return False
    if is_header_line(t):
        return False
    if ONLY_NUMBER_PATTERN.match(t):
        return False
    if is_subcategory_only(t):
        return False
    if any(x in t for x in ["御 見 積 書", "E S T I M A T E", "見積作成日", "振 込 先", "工 事 件 名", "工 事 場 所", "お支払い条件", "有効期限", "小計", "小 計"]):
        return False
    if re.search(r"(株式会社|有限会社|〒|TEL|FAX|MAIL|登録番号)", t):
        return False
    if re.match(r"^\d+\s+", t):
        return False
    if is_detail_like_line(t):
        return False
    return True


def second_page_has_title_above_header(lines: List[str]) -> Tuple[bool, str]:
    header_idx = find_header_index(lines)
    if header_idx is None:
        return False, ""

    upper_lines = [clean_text(x) for x in lines[:header_idx] if clean_text(x)]
    if not upper_lines:
        return False, ""

    candidates = [x for x in upper_lines if is_page_title_candidate(x)]
    if not candidates:
        return False, ""

    title = normalize_major_category(candidates[-1])
    return True, title


def should_skip_first_page(pdf) -> bool:
    if len(pdf.pages) < 2:
        return False

    second_page_lines = extract_lines_from_page(pdf.pages[1])
    has_title, _ = second_page_has_title_above_header(second_page_lines)
    return has_title


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

    footer_cutoff = page.height - 28
    usable_words = []
    for w in words:
        txt = clean_text(w.get("text", ""))
        if not txt:
            continue
        top = float(w["top"])
        bottom = float(w["bottom"])
        if top >= footer_cutoff or bottom >= footer_cutoff:
            continue
        usable_words.append(w)

    usable_words = sorted(usable_words, key=lambda w: (round(w["top"], 1), w["x0"]))

    grouped = []
    current = []
    current_top = None
    tolerance = 3

    for w in usable_words:
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


def detect_major_category(lines: List[str]) -> str:
    header_idx = find_header_index(lines)
    if header_idx is not None:
        upper = [clean_text(x) for x in lines[:header_idx] if clean_text(x)]
        candidates = [x for x in upper if is_page_title_candidate(x)]
        if candidates:
            return normalize_major_category(candidates[-1])

    # ヘッダー上になければページ全体からの簡易救済
    for line in lines[:8]:
        t = clean_text(line)
        if is_page_title_candidate(t):
            return normalize_major_category(t)

    return ""


def is_detail_like_line(text: str) -> bool:
    t = clean_text(text)
    parsed = parse_detail_line_core(t)
    return parsed is not None


def parse_detail_line_core(line: str) -> Optional[Dict]:
    t = clean_text(line)
    if not t:
        return None

    # パターン1: no left qty unit unit_price amount
    m1 = re.match(
        r"^(?P<no>\d+)\s+(?P<left>.+?)\s+(?P<qty>-?\d+(?:\.\d+)?)\s+"
        r"(?P<unit>\S+)\s+(?P<unit_price>\(?¥?-?[\d,]+\)?)\s+(?P<amount>\(?¥?-?[\d,]+\)?)$",
        t
    )
    if m1:
        return {
            "no": clean_text(m1.group("no")),
            "item_spec": clean_text(m1.group("left")),
            "quantity": clean_text(m1.group("qty")),
            "unit": clean_text(m1.group("unit")),
            "unit_price": normalize_money_token(m1.group("unit_price")),
            "amount": normalize_money_token(m1.group("amount")),
        }

    # パターン2: no left unit qty unit_price amount
    m2 = re.match(
        r"^(?P<no>\d+)\s+(?P<left>.+?)\s+(?P<unit>\S+)\s+"
        r"(?P<qty>-?\d+(?:\.\d+)?)\s+(?P<unit_price>\(?¥?-?[\d,]+\)?)\s+(?P<amount>\(?¥?-?[\d,]+\)?)$",
        t
    )
    if m2:
        return {
            "no": clean_text(m2.group("no")),
            "item_spec": clean_text(m2.group("left")),
            "quantity": clean_text(m2.group("qty")),
            "unit": clean_text(m2.group("unit")),
            "unit_price": normalize_money_token(m2.group("unit_price")),
            "amount": normalize_money_token(m2.group("amount")),
        }

    # パターン3: no left amount amount? broken no use
    return None


def parse_detail_line(line: str) -> Optional[Dict]:
    t = clean_text(line)
    if not t:
        return None

    if is_header_line(t):
        return None

    if any(x in t for x in ["御 見 積 書", "E S T I M A T E", "見積作成日", "振 込 先", "工 事 件 名", "工 事 場 所", "お支払い条件", "有効期限"]):
        return None

    if "小計" in t or "小 計" in t:
        return None

    if re.match(r"^PAGE\.", t, flags=re.IGNORECASE):
        return None

    if ONLY_NUMBER_PATTERN.match(t):
        return None

    if is_subcategory_only(t):
        return {
            "no": "",
            "item_spec": "",
            "quantity": "",
            "unit": "",
            "unit_price": "",
            "amount": "",
            "raw_row": t,
            "needs_review": 1,
            "is_subcategory": True,
        }

    core = parse_detail_line_core(t)
    if core is None:
        return {
            "no": "",
            "item_spec": "",
            "quantity": "",
            "unit": "",
            "unit_price": "",
            "amount": "",
            "raw_row": t,
            "needs_review": 1,
            "is_subcategory": False,
        }

    if not is_quantity_token(core["quantity"]):
        return {
            "no": "",
            "item_spec": "",
            "quantity": "",
            "unit": "",
            "unit_price": "",
            "amount": "",
            "raw_row": t,
            "needs_review": 1,
            "is_subcategory": False,
        }

    if not is_unit_token(core["unit"]):
        return {
            "no": "",
            "item_spec": "",
            "quantity": "",
            "unit": "",
            "unit_price": "",
            "amount": "",
            "raw_row": t,
            "needs_review": 1,
            "is_subcategory": False,
        }

    if not is_money_token(core["unit_price"]) or not is_money_token(core["amount"]):
        return {
            "no": "",
            "item_spec": "",
            "quantity": "",
            "unit": "",
            "unit_price": "",
            "amount": "",
            "raw_row": t,
            "needs_review": 1,
            "is_subcategory": False,
        }

    return {
        "no": core["no"],
        "item_spec": core["item_spec"],
        "quantity": core["quantity"],
        "unit": core["unit"],
        "unit_price": core["unit_price"],
        "amount": core["amount"],
        "raw_row": t,
        "needs_review": 0,
        "is_subcategory": False,
    }


def merge_multiline_details(lines: List[str]) -> List[str]:
    merged = []
    i = 0

    while i < len(lines):
        current = clean_text(lines[i])
        if not current:
            i += 1
            continue

        # 先に通常パースできればそのまま
        parsed = parse_detail_line_core(current)
        if parsed is not None:
            merged.append(current)
            i += 1
            continue

        # 次行と結合してパースできるか試す
        if i + 1 < len(lines):
            nxt = clean_text(lines[i + 1])
            combined = clean_text(current + " " + nxt)
            parsed2 = parse_detail_line_core(combined)
            if parsed2 is not None:
                merged.append(combined)
                i += 2
                continue

            # 先頭行が no + left だけ、次行に数量以降
            m_left = re.match(r"^(?P<no>\d+)\s+(?P<left>.+)$", current)
            if m_left:
                combined2 = clean_text(current + " " + nxt)
                parsed3 = parse_detail_line_core(combined2)
                if parsed3 is not None:
                    merged.append(combined2)
                    i += 2
                    continue

        merged.append(current)
        i += 1

    return merged


def process_pdf(file_name: str, file_bytes: bytes) -> List[Dict]:
    rows = []

    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        estimate_date = extract_estimate_date(pdf)
        skip_first_page = should_skip_first_page(pdf)

        for page_num, page in enumerate(pdf.pages, start=1):
            if page_num == 1 and skip_first_page:
                continue

            lines = extract_lines_from_page(page)
            lines = merge_multiline_details(lines)
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

    if re.match(r"^PAGE\.", raw, flags=re.IGNORECASE):
        return True

    if is_subcategory_only(raw):
        return True

    if raw == major and major != "":
        return True

    if ONLY_NUMBER_PATTERN.match(raw):
        return True

    if raw in {"御 見 積 書", "E S T I M A T E"}:
        return True

    if needs_review == 1 and not any([qty, unit, unit_price, amount]):
        return True

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
