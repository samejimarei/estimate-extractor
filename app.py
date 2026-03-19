import io
import re
import zipfile
from typing import List, Dict, Optional, Tuple, Any

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
    "式", "台", "枚", "本", "箇所", "箇", "ｍ", "m", "m2", "㎡", "㎥",
    "日", "セット", "ｾｯﾄ", "人工", "個", "ヶ所", "箱", "巻", "丁", "脚", "面"
}

HEADER_KEYWORDS = ["NO.", "項目", "仕様", "規格", "型番", "数量", "単位", "単価", "金額"]

EXCLUDE_PATTERNS = [
    r"^小\s*計",
    r"^小計",
    r"^御見積金額",
    r"^内消費税",
    r"^PAGE\.?",
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
LINE_NO_PATTERN = re.compile(r"^\d+$")
QUANTITY_PATTERN = re.compile(r"^-?\d+(?:\.\d+)?$")
CIRCLED_NUM_PATTERN = re.compile(r"[①②③④⑤⑥⑦⑧⑨⑩]$")
TRAILING_NUM_PATTERN = re.compile(r"\s*\d+$")


def clean_text(text: str) -> str:
    text = (text or "").replace("\u3000", " ")
    text = re.sub(r"[ \t]+", " ", text)
    return text.strip()


def normalize_major_category(text: str) -> str:
    t = clean_text(text)
    t = re.sub(CIRCLED_NUM_PATTERN, "", t)
    t = re.sub(TRAILING_NUM_PATTERN, "", t)
    return t.strip()


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


def looks_like_subcategory(line: str) -> bool:
    t = clean_text(line)
    if not t:
        return False
    return bool(SUBCATEGORY_PATTERN.match(t))


def is_numeric_token(text: str) -> bool:
    return bool(NUMERIC_PATTERN.match(clean_text(text)))


def is_quantity_token(text: str) -> bool:
    return bool(QUANTITY_PATTERN.match(clean_text(text)))


def is_unit_token(text: str) -> bool:
    t = clean_text(text)
    if not t:
        return False
    if t in UNIT_CANDIDATES:
        return True
    if re.match(r"^[A-Za-z0-9㎡㎥ｍm/]+$", t):
        return True
    return False


def extract_words_structured(page) -> List[Dict[str, Any]]:
    words = page.extract_words(
        keep_blank_chars=False,
        use_text_flow=False,
        x_tolerance=1.5,
        y_tolerance=1.5,
        extra_attrs=[]
    )
    normalized = []
    for w in words or []:
        txt = clean_text(w.get("text", ""))
        if not txt:
            continue
        normalized.append({
            "text": txt,
            "x0": float(w["x0"]),
            "x1": float(w["x1"]),
            "top": float(w["top"]),
            "bottom": float(w["bottom"]),
            "doctop": float(w.get("doctop", w["top"])),
        })
    return normalized


def group_words_into_rows(words: List[Dict[str, Any]], y_tolerance: float = 3.0) -> List[List[Dict[str, Any]]]:
    if not words:
        return []

    words = sorted(words, key=lambda w: (round(w["top"], 1), w["x0"]))
    rows: List[List[Dict[str, Any]]] = []
    current: List[Dict[str, Any]] = []
    current_top: Optional[float] = None

    for w in words:
        top = w["top"]
        if current_top is None:
            current = [w]
            current_top = top
            continue

        if abs(top - current_top) <= y_tolerance:
            current.append(w)
        else:
            rows.append(sorted(current, key=lambda x: x["x0"]))
            current = [w]
            current_top = top

    if current:
        rows.append(sorted(current, key=lambda x: x["x0"]))

    return rows


def row_to_text(row_words: List[Dict[str, Any]]) -> str:
    return clean_text(" ".join(w["text"] for w in sorted(row_words, key=lambda x: x["x0"])))


def detect_page_type_from_rows(row_texts: List[str], page_num: int) -> str:
    joined = " ".join(row_texts)

    if "NO." in joined and "単価" in joined and "金額" in joined:
        return "detail"

    if page_num == 1 and "御見積書" in joined:
        return "summary"

    if page_num == 1:
        return "summary"

    return "unknown"


def detect_header_row_index(row_texts: List[str]) -> Optional[int]:
    for i, t in enumerate(row_texts[:10]):
        score = 0
        for k in ["NO.", "数量", "単位", "単価", "金額"]:
            if k in t:
                score += 1
        if score >= 3:
            return i
    return None


def detect_column_anchors(header_row_words: List[Dict[str, Any]]) -> Dict[str, float]:
    anchors: Dict[str, float] = {}

    for w in header_row_words:
        txt = w["text"]
        cx = (w["x0"] + w["x1"]) / 2

        if txt == "NO.":
            anchors["line_no"] = cx
        elif "数量" in txt:
            anchors["quantity"] = cx
        elif "単位" in txt:
            anchors["unit"] = cx
        elif "単価" in txt:
            anchors["unit_price"] = cx
        elif "金額" in txt:
            anchors["amount"] = cx

    return anchors


def split_row_by_anchors(row_words: List[Dict[str, Any]], anchors: Dict[str, float]) -> Dict[str, str]:
    if not anchors:
        return {
            "line_no": "",
            "left_text": row_to_text(row_words),
            "quantity": "",
            "unit": "",
            "unit_price": "",
            "amount": "",
        }

    qty_x = anchors.get("quantity", 10_000)
    unit_x = anchors.get("unit", 10_000)
    unit_price_x = anchors.get("unit_price", 10_000)
    amount_x = anchors.get("amount", 10_000)

    line_no_words = []
    left_words = []
    qty_words = []
    unit_words = []
    unit_price_words = []
    amount_words = []

    for w in row_words:
        cx = (w["x0"] + w["x1"]) / 2

        if "line_no" in anchors and cx < (anchors["line_no"] + qty_x) / 2 and LINE_NO_PATTERN.match(w["text"]):
            line_no_words.append(w)
        elif cx < (qty_x + unit_x) / 2:
            left_words.append(w)
        elif cx < (unit_x + unit_price_x) / 2:
            qty_words.append(w)
        elif cx < (unit_price_x + amount_x) / 2:
            unit_words.append(w)
        elif cx < amount_x + 40:
            unit_price_words.append(w)
        else:
            amount_words.append(w)

    result = {
        "line_no": row_to_text(line_no_words),
        "left_text": row_to_text(left_words),
        "quantity": row_to_text(qty_words),
        "unit": row_to_text(unit_words),
        "unit_price": row_to_text(unit_price_words),
        "amount": row_to_text(amount_words),
    }

    # amount列に単価・金額が一緒に入った場合の補正
    if not result["unit_price"] and result["amount"]:
        parts = result["amount"].split()
        if len(parts) >= 2 and all(is_numeric_token(p) for p in parts[-2:]):
            result["unit_price"] = parts[-2]
            result["amount"] = parts[-1]

    return result


def fallback_parse_detail_from_right(row_text: str) -> Dict[str, str]:
    t = clean_text(row_text)

    m = re.match(
        r"^(?P<line_no>\d+)?\s*(?P<left>.+?)\s+(?P<quantity>-?\d+(?:\.\d+)?)\s+(?P<unit>\S+)\s+(?P<unit_price>-?[\d,]+)\s+(?P<amount>-?[\d,]+)$",
        t
    )
    if not m:
        return {
            "line_no": "",
            "left_text": t,
            "quantity": "",
            "unit": "",
            "unit_price": "",
            "amount": "",
        }

    return {
        "line_no": clean_text(m.group("line_no") or ""),
        "left_text": clean_text(m.group("left")),
        "quantity": clean_text(m.group("quantity")),
        "unit": clean_text(m.group("unit")),
        "unit_price": clean_text(m.group("unit_price")),
        "amount": clean_text(m.group("amount")),
    }


def split_item_and_spec(left_text: str) -> Tuple[str, str]:
    t = clean_text(left_text)
    if not t:
        return "", ""

    # 先に全角・半角の区切りで分ける
    if "　" in left_text:
        parts = [clean_text(x) for x in left_text.split("　") if clean_text(x)]
        if len(parts) >= 2:
            return parts[0], " ".join(parts[1:])

    # 型番やメーカーが後ろに付くケースの簡易分離
    # 例: 洗面化粧台 LIXIL ピアラ
    tokens = t.split()
    if len(tokens) >= 3:
        first = tokens[0]
        rest = " ".join(tokens[1:])

        # 先頭1語を項目、残りを仕様に寄せる
        if len(first) <= 12:
            return first, rest

    return t, ""


def detect_major_category_from_rows(row_texts: List[str], page_type: str, page_num: int) -> str:
    if page_type != "detail":
        return ""

    candidates = []
    for t in row_texts[:8]:
        tt = clean_text(t)
        if not tt:
            continue
        if any(k in tt for k in HEADER_KEYWORDS):
            continue
        if "PAGE." in tt or "御見積書" in tt:
            continue
        if re.search(r"^[0-9,\.\-\s]+$", tt):
            continue
        if len(tt) <= 40:
            candidates.append(tt)

    if not candidates:
        return ""

    # 「○○工事」を優先
    for c in candidates:
        if "工事" in c:
            return c

    return candidates[0]


def parse_summary_line(line: str) -> Optional[Dict[str, Any]]:
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
    exclude_flag = 1 if item_name in {"調整値引き", "値引き"} else 0
    exclude_reason = "discount" if exclude_flag else ""

    return {
        "line_no": m.group("line_no"),
        "item_name": item_name,
        "spec": "",
        "quantity": m.group("quantity"),
        "unit": m.group("unit"),
        "unit_price": m.group("unit_price"),
        "amount": m.group("amount"),
        "needs_review": 0,
        "exclude_flag": exclude_flag,
        "exclude_reason": exclude_reason,
    }


def build_detail_records_from_page(page) -> List[Dict[str, Any]]:
    words = extract_words_structured(page)
    if not words:
        return []

    row_words_list = group_words_into_rows(words, y_tolerance=3.0)
    row_texts = [row_to_text(r) for r in row_words_list]

    header_idx = detect_header_row_index(row_texts)
    anchors = {}
    if header_idx is not None:
        anchors = detect_column_anchors(row_words_list[header_idx])

    records: List[Dict[str, Any]] = []
    started = header_idx is None

    for idx, row_words in enumerate(row_words_list):
        row_text = row_to_text(row_words)
        if not row_text:
            continue

        reason = is_excluded_line(row_text)
        if header_idx is not None and idx == header_idx:
            started = True
            continue

        if not started:
            continue

        if reason in {"blank", "meta", "header"}:
            continue

        if looks_like_subcategory(row_text):
            records.append({
                "record_type": "subcategory",
                "raw_row": row_text,
                "subcategory": clean_text(row_text).strip("-－ー").strip(),
            })
            continue

        cells = split_row_by_anchors(row_words, anchors) if anchors else fallback_parse_detail_from_right(row_text)

        line_no = clean_text(cells["line_no"])
        left_text = clean_text(cells["left_text"])
        quantity = clean_text(cells["quantity"])
        unit = clean_text(cells["unit"])
        unit_price = clean_text(cells["unit_price"])
        amount = clean_text(cells["amount"])

        # アンカー分割がうまくいかない場合は右側パースにフォールバック
        if not amount or not unit_price:
            fb = fallback_parse_detail_from_right(row_text)
            if fb["amount"] and fb["unit_price"]:
                line_no = fb["line_no"] or line_no
                left_text = fb["left_text"] or left_text
                quantity = fb["quantity"] or quantity
                unit = fb["unit"] or unit
                unit_price = fb["unit_price"] or unit_price
                amount = fb["amount"] or amount

        item_name, spec = split_item_and_spec(left_text)

        parsed_ok = (
            bool(line_no) and
            bool(item_name) and
            bool(quantity) and
            bool(unit) and
            bool(unit_price) and
            bool(amount) and
            is_quantity_token(quantity) and
            is_unit_token(unit) and
            is_numeric_token(unit_price) and
            is_numeric_token(amount)
        )

        exclude_flag = 1 if item_name in {"調整値引き", "値引き"} else 0
        exclude_reason = "discount" if exclude_flag else ""

        records.append({
            "record_type": "detail",
            "parsed": parsed_ok,
            "line_no": line_no,
            "item_name": item_name,
            "spec": spec,
            "quantity": quantity,
            "unit": unit,
            "unit_price": unit_price,
            "amount": amount,
            "raw_row": row_text,
            "exclude_flag": exclude_flag,
            "exclude_reason": exclude_reason,
        })

    return merge_continuation_rows(records)


def merge_continuation_rows(records: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    line_noが無く、数値列も無い行は直前明細の続きとして結合。
    内装工事ページの2段化対策。
    """
    merged: List[Dict[str, Any]] = []

    for rec in records:
        if rec["record_type"] != "detail":
            merged.append(rec)
            continue

        is_continuation = (
            not rec.get("line_no") and
            rec.get("raw_row") and
            not rec.get("quantity") and
            not rec.get("unit") and
            not rec.get("unit_price") and
            not rec.get("amount")
        )

        if is_continuation and merged:
            prev = merged[-1]
            if prev.get("record_type") == "detail":
                extra_text = clean_text(rec.get("raw_row", ""))
                if extra_text:
                    prev["raw_row"] = clean_text(prev.get("raw_row", "") + " " + extra_text)

                    if prev.get("spec"):
                        prev["spec"] = clean_text(prev["spec"] + " " + extra_text)
                    elif prev.get("item_name"):
                        prev["spec"] = extra_text
                    else:
                        prev["item_name"] = extra_text

                    prev["parsed"] = bool(
                        prev.get("line_no") and
                        prev.get("item_name") and
                        prev.get("quantity") and
                        prev.get("unit") and
                        prev.get("unit_price") and
                        prev.get("amount")
                    )
                    continue

        merged.append(rec)

    return merged


def extract_lines_for_unknown(page) -> List[str]:
    text = page.extract_text() or ""
    return [clean_text(x) for x in text.splitlines() if clean_text(x)]


def process_pdf(file_name: str, file_bytes: bytes) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []

    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        current_subcategory = ""

        for page_num, page in enumerate(pdf.pages, start=1):
            words = extract_words_structured(page)
            row_words_list = group_words_into_rows(words, y_tolerance=3.0)
            row_texts = [row_to_text(r) for r in row_words_list] if row_words_list else extract_lines_for_unknown(page)

            page_type = detect_page_type_from_rows(row_texts, page_num)
            major_category_raw = detect_major_category_from_rows(row_texts, page_type, page_num)
            major_category = normalize_major_category(major_category_raw) if major_category_raw else ""

            if page_type == "summary":
                for line in row_texts:
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
                detail_records = build_detail_records_from_page(page)

                for rec in detail_records:
                    if rec["record_type"] == "subcategory":
                        current_subcategory = rec["subcategory"]
                        continue

                    if rec["record_type"] != "detail":
                        continue

                    parsed = rec.get("parsed", False)
                    reason = rec.get("exclude_reason", "")
                    exclude_flag = rec.get("exclude_flag", 0)

                    needs_review = 0
                    if not parsed:
                        needs_review = 1
                    elif len(rec.get("item_name", "")) > 40 and not rec.get("spec"):
                        needs_review = 1

                    rows.append({
                        "file_name": file_name,
                        "page": page_num,
                        "page_type": "detail",
                        "major_category_raw": major_category_raw,
                        "major_category": major_category,
                        "sub_category": current_subcategory,
                        "line_no": rec.get("line_no", ""),
                        "item_name": rec.get("item_name", ""),
                        "spec": rec.get("spec", ""),
                        "quantity": rec.get("quantity", ""),
                        "unit": rec.get("unit", ""),
                        "unit_price": rec.get("unit_price", ""),
                        "amount": rec.get("amount", ""),
                        "raw_row": rec.get("raw_row", ""),
                        "needs_review": needs_review,
                        "exclude_flag": exclude_flag,
                        "exclude_reason": reason if reason else ("unparsed_line" if not parsed else ""),
                    })

            else:
                for line in row_texts:
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


def process_uploaded_file(uploaded_file) -> List[Dict[str, Any]]:
    all_rows: List[Dict[str, Any]] = []

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
    rows: List[Dict[str, Any]] = []
    for uploaded_file in uploaded_files:
        rows.extend(process_uploaded_file(uploaded_file))

    df = pd.DataFrame(rows, columns=DETAIL_COLUMNS)

    st.subheader("抽出結果プレビュー")
    st.write(f"抽出行数: {len(df):,}")
    st.dataframe(df, use_container_width=True, height=600)

    st.subheader("簡易集計")
    c1, c2, c3 = st.columns(3)
    c1.metric("needs_review件数", int((df["needs_review"] == 1).sum()) if not df.empty else 0)
    c2.metric("detail件数", int((df["page_type"] == "detail").sum()) if not df.empty else 0)
    c3.metric("unknown件数", int((df["page_type"] == "unknown").sum()) if not df.empty else 0)

    csv_data = df.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        label="CSVダウンロード",
        data=csv_data,
        file_name="estimate_extract_detail.csv",
        mime="text/csv"
    )
else:
    st.info("まずPDFまたはZIPをアップロードしてください。")
