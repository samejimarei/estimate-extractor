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
    "缶", "帖", "室", "ヶ", "双", "組", "箇口", "A工", "B工", "C工", "L", "kg",
    "回", "樽", "人", "時間", "工", "坪", "畳", "m3", "㎏", "箇所分"
}

DATE_PATTERNS = [
    re.compile(r"(20\d{2})年(\d{1,2})月(\d{1,2})日"),
    re.compile(r"(20\d{2})/(\d{1,2})/(\d{1,2})"),
    re.compile(r"(20\d{2})\.(\d{1,2})\.(\d{1,2})"),
]

HEADER_TOKENS = ["NO.", "項目", "仕様・規格/型番", "数量", "単位", "単価", "金額"]

SUBCATEGORY_ONLY_PATTERN = re.compile(r"^\s*(?:\d+\s*)?[\-－ー【\[].+[\-－ー】\]]\s*$")
ONLY_NUMBER_PATTERN = re.compile(r"^\d+$")
LINE_START_NO_PATTERN = re.compile(r"^(?P<no>\d+)\s+(?P<rest>.+)$")
MONEY_BODY_PATTERN = re.compile(r"-?[\d,]+")
RAW_MONEY_TOKEN_PATTERN = re.compile(r"^\(?[¥￥]?-?[\d,]+\)?$")


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


def normalize_major_category(text: str) -> str:
    t = clean_text(text)
    t = re.sub(r"\s*[①②③④⑤⑥⑦⑧⑨⑩]$", "", t)
    t = re.sub(r"\s*\d+$", "", t)
    return t.strip()


def is_header_line(text: str) -> bool:
    t = clean_text(text)
    if not t:
        return False
    score = sum(1 for token in ["NO.", "項目", "数量", "単位", "単価", "金額"] if token in t)
    return score >= 4


def find_header_index(lines: List[str]) -> Optional[int]:
    for i, line in enumerate(lines):
        if is_header_line(line):
            return i
    return None


def normalize_money_token(text: str) -> str:
    t = clean_text(text)
    if not t:
        return ""

    t = t.replace("¥", "").replace("￥", "").replace(" ", "")

    if t.startswith("(") and t.endswith(")"):
        inner = t[1:-1]
        if MONEY_BODY_PATTERN.fullmatch(inner):
            return "-" + inner

    return t


def is_money_token(text: str) -> bool:
    t = normalize_money_token(text)
    return bool(MONEY_BODY_PATTERN.fullmatch(t))


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


def is_subcategory_only(text: str) -> bool:
    return bool(SUBCATEGORY_ONLY_PATTERN.match(clean_text(text)))


def looks_like_cover_noise(text: str) -> bool:
    t = clean_text(text)
    if not t:
        return True

    cover_keywords = [
        "御 見 積 書", "E S T I M A T E", "見積作成日", "振 込 先", "工 事 件 名", "工 事 場 所",
        "お支払い条件", "有効期限", "代表取締役", "登録番号", "MAIL.", "TEL", "FAX", "〒",
        "株式会社", "有限会社", "御中"
    ]
    return any(k in t for k in cover_keywords)


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
    if looks_like_cover_noise(t):
        return False
    if "小計" in t or "小 計" in t:
        return False
    if RAW_MONEY_TOKEN_PATTERN.match(t):
        return False
    if t.startswith("¥") or t.startswith("￥"):
        return False
    if re.search(r"^[\(\)¥￥,\d\s\-]+$", t):
        return False
    if re.match(r"^\d+\s+", t):
        return False
    if is_subcategory_only(t):
        return False
    return True


def second_page_has_title_above_header(lines: List[str]) -> Tuple[bool, str]:
    header_idx = find_header_index(lines)
    if header_idx is None:
        return False, ""

    upper_lines = [clean_text(x) for x in lines[:header_idx] if clean_text(x)]
    candidates = [x for x in upper_lines if is_page_title_candidate(x)]
    if not candidates:
        return False, ""

    title = normalize_major_category(candidates[-1])
    return True, title


def should_skip_first_page(pdf) -> bool:
    if len(pdf.pages) < 2:
        return False

    second_lines = extract_lines_from_page(pdf.pages[1])
    has_title, _ = second_page_has_title_above_header(second_lines)
    return has_title


def detect_major_category(lines: List[str]) -> str:
    header_idx = find_header_index(lines)
    if header_idx is not None:
        upper_lines = [clean_text(x) for x in lines[:header_idx] if clean_text(x)]
        candidates = [x for x in upper_lines if is_page_title_candidate(x)]
        if candidates:
            return normalize_major_category(candidates[-1])

    return ""


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

    footer_cutoff = page.height - 24
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


def strip_inline_bracket_heading(text: str) -> str:
    t = clean_text(text)
    t = re.sub(r"^【[^】]+】\s*", "", t)
    t = re.sub(r"^\[[^\]]+\]\s*", "", t)
    return clean_text(t)


def strip_leading_inner_no(text: str) -> str:
    t = clean_text(text)
    t = re.sub(r"^(?:\d+\s+)+", "", t)
    return clean_text(t)


def parse_detail_line_core(line: str) -> Optional[Dict]:
    t = clean_text(line)
    if not t:
        return None

    t = strip_inline_bracket_heading(t)

    m1 = re.match(
        r"^(?P<no>\d+)\s+(?P<left>.+?)\s+(?P<qty>-?\d+(?:\.\d+)?)\s+"
        r"(?P<unit>\S+)\s+(?P<unit_price>\(?[¥￥]?-?[\d,]+\)?)\s+"
        r"(?P<amount>\(?[¥￥]?-?[\d,]+\)?)$",
        t
    )
    if m1:
        left = strip_leading_inner_no(m1.group("left"))
        return {
            "no": clean_text(m1.group("no")),
            "item_spec": clean_text(left),
            "quantity": clean_text(m1.group("qty")),
            "unit": clean_text(m1.group("unit")),
            "unit_price": normalize_money_token(m1.group("unit_price")),
            "amount": normalize_money_token(m1.group("amount")),
        }

    m2 = re.match(
        r"^(?P<no>\d+)\s+(?P<left>.+?)\s+(?P<unit>\S+)\s+"
        r"(?P<qty>-?\d+(?:\.\d+)?)\s+(?P<unit_price>\(?[¥￥]?-?[\d,]+\)?)\s+"
        r"(?P<amount>\(?[¥￥]?-?[\d,]+\)?)$",
        t
    )
    if m2:
        left = strip_leading_inner_no(m2.group("left"))
        return {
            "no": clean_text(m2.group("no")),
            "item_spec": clean_text(left),
            "quantity": clean_text(m2.group("qty")),
            "unit": clean_text(m2.group("unit")),
            "unit_price": normalize_money_token(m2.group("unit_price")),
            "amount": normalize_money_token(m2.group("amount")),
        }

    return None


def parse_detail_line(line: str) -> Optional[Dict]:
    t = clean_text(line)
    if not t:
        return None

    if is_header_line(t):
        return None

    if looks_like_cover_noise(t):
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


def line_starts_with_no(text: str) -> bool:
    return LINE_START_NO_PATTERN.match(clean_text(text)) is not None


def line_has_tail_values_only(text: str) -> bool:
    t = clean_text(text)
    if not t:
        return False

    if line_starts_with_no(t):
        return False

    if re.search(
        r"-?\d+(?:\.\d+)?\s+\S+\s+\(?[¥￥]?-?[\d,]+\)?\s+\(?[¥￥]?-?[\d,]+\)?$",
        t
    ):
        return True

    if re.search(
        r"\S+\s+-?\d+(?:\.\d+)?\s+\(?[¥￥]?-?[\d,]+\)?\s+\(?[¥￥]?-?[\d,]+\)?$",
        t
    ):
        return True

    return False


def should_merge_lines(current: str, nxt: str) -> bool:
    current = clean_text(current)
    nxt = clean_text(nxt)

    if not current or not nxt:
        return False

    if parse_detail_line_core(current) is not None:
        return False

    if not line_starts_with_no(current):
        return False

    if line_starts_with_no(nxt):
        return False

    if is_subcategory_only(nxt):
        return False

    if is_header_line(nxt):
        return False

    if "小計" in nxt or "小 計" in nxt:
        return False

    if not line_has_tail_values_only(nxt):
        return False

    combined = clean_text(current + " " + nxt)
    return parse_detail_line_core(combined) is not None


def merge_multiline_details(lines: List[str]) -> List[str]:
    merged = []
    i = 0

    while i < len(lines):
        current = clean_text(lines[i])
        if not current:
            i += 1
            continue

        if i + 1 < len(lines):
            nxt = clean_text(lines[i + 1])
            if should_merge_lines(current, nxt):
                merged.append(clean_text(current + " " + nxt))
                i += 2
                continue

        merged.append(current)
        i += 1

    return merged


# ============================================================
# 座標ベース本番抽出用関数群
# ============================================================

def extract_words_debug_from_page(page) -> List[Dict]:
    words = page.extract_words(
        keep_blank_chars=False,
        use_text_flow=True,
        x_tolerance=2,
        y_tolerance=2,
    )

    footer_cutoff = page.height - 24
    results = []

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

    words_sorted = sorted(words, key=lambda w: (round(w["top"], 1), w["x0"]))

    rows = []
    current = []
    current_top = None

    for w in words_sorted:
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


def row_words_to_text(row_words: List[Dict]) -> str:
    return " ".join(w["text"] for w in sorted(row_words, key=lambda x: x["x0"]))


def find_header_row_words(rows: List[List[Dict]]) -> Optional[List[Dict]]:
    header_keywords = ["NO.", "項目", "数量", "単位", "単価", "金額"]

    for row in rows:
        text = row_words_to_text(row)
        score = sum(1 for k in header_keywords if k in text)
        if score >= 5:
            return row

    return None


def get_header_word_pos(header_row: List[Dict], keyword: str) -> Tuple[Optional[float], Optional[float]]:
    for w in header_row:
        if keyword in w["text"]:
            return w["x0"], w["x1"]
    return None, None


def build_dynamic_column_boundaries(header_row: List[Dict]) -> Dict[str, Tuple[float, float]]:
    no_x0, no_x1 = get_header_word_pos(header_row, "NO.")
    item_x0, item_x1 = get_header_word_pos(header_row, "項目")
    qty_x0, qty_x1 = get_header_word_pos(header_row, "数量")
    unit_x0, unit_x1 = get_header_word_pos(header_row, "単位")
    unit_price_x0, unit_price_x1 = get_header_word_pos(header_row, "単価")
    amount_x0, amount_x1 = get_header_word_pos(header_row, "金額")

    required = [
        no_x0, no_x1, item_x0, item_x1, qty_x0, qty_x1,
        unit_x0, unit_x1, unit_price_x0, unit_price_x1, amount_x0, amount_x1
    ]
    if any(v is None for v in required):
        raise ValueError("ヘッダー語の位置取得に失敗しました。")

    margin = 6

    boundaries = {
        "no": (
            max(0, no_x0 - margin),
            (item_x0 + no_x1) / 2
        ),
        "item_spec": (
            max(0, item_x0 - margin),
            (qty_x0 + item_x1) / 2
        ),
        "quantity": (
            max(0, qty_x0 - margin),
            (unit_x0 + qty_x1) / 2
        ),
        "unit": (
            max(0, unit_x0 - margin),
            (unit_price_x0 + unit_x1) / 2
        ),
        "unit_price": (
            max(0, unit_price_x0 - margin),
            (amount_x0 + unit_price_x1) / 2
        ),
        "amount": (
            max(0, amount_x0 - margin),
            9999.0
        ),
    }

    return boundaries


def assign_word_to_dynamic_column(word: Dict, boundaries: Dict[str, Tuple[float, float]]) -> Optional[str]:
    center_x = (word["x0"] + word["x1"]) / 2

    for col_name, (x_min, x_max) in boundaries.items():
        if x_min <= center_x < x_max:
            return col_name

    return None


def row_words_to_dynamic_record(row_words: List[Dict], boundaries: Dict[str, Tuple[float, float]]) -> Dict:
    buckets = {
        "no": [],
        "item_spec": [],
        "quantity": [],
        "unit": [],
        "unit_price": [],
        "amount": [],
    }

    for w in row_words:
        col = assign_word_to_dynamic_column(w, boundaries)
        if col is not None:
            buckets[col].append(w["text"])

    return {
        "no": clean_text(" ".join(buckets["no"])),
        "item_spec": clean_text(" ".join(buckets["item_spec"])),
        "quantity": clean_text(" ".join(buckets["quantity"])),
        "unit": clean_text(" ".join(buckets["unit"])),
        "unit_price": clean_text(" ".join(buckets["unit_price"])),
        "amount": clean_text(" ".join(buckets["amount"])),
        "raw_row": clean_text(row_words_to_text(row_words)),
    }


def is_date_fragment_text(text: str) -> bool:
    t = clean_text(text)
    if not t:
        return True

    if re.fullmatch(r"\d{1,2}\s+\d{1,2}", t):
        return True

    if re.fullmatch(r"/\s*\d{1,2}", t):
        return True

    if re.fullmatch(r"\d{1,2}\s*/", t):
        return True

    if re.fullmatch(r"\d{4}", t):
        return True

    if re.fullmatch(r"\d{1,2}", t):
        return True

    if re.fullmatch(r"/", t):
        return True

    if len(t) <= 4 and re.fullmatch(r"[\d/\-]+", t):
        return True

    return False


def row_has_tail_value_pattern(record: Dict) -> bool:
    qty = clean_text(record.get("quantity", ""))
    unit = clean_text(record.get("unit", ""))
    unit_price = clean_text(record.get("unit_price", ""))
    amount = clean_text(record.get("amount", ""))

    score = 0
    if qty and (is_quantity_token(qty) or re.search(r"\d", qty)):
        score += 1
    if unit:
        score += 1
    if unit_price and is_money_token(unit_price):
        score += 1
    if amount and is_money_token(amount):
        score += 1

    return score >= 3


def is_poc_noise_row(row_words: List[Dict], boundaries: Dict[str, Tuple[float, float]]) -> bool:
    text = row_words_to_text(row_words)
    record = row_words_to_dynamic_record(row_words, boundaries)

    if is_date_fragment_text(text):
        return True

    if is_header_line(text):
        return True

    if "小計" in text or "小 計" in text:
        return True

    if looks_like_cover_noise(text):
        return True

    no_text = clean_text(record.get("no", ""))
    item_text = clean_text(record.get("item_spec", ""))
    qty = clean_text(record.get("quantity", ""))
    unit = clean_text(record.get("unit", ""))
    unit_price = clean_text(record.get("unit_price", ""))
    amount = clean_text(record.get("amount", ""))

    non_empty = [x for x in [no_text, item_text, qty, unit, unit_price, amount] if x]
    if len(non_empty) == 0:
        return True

    if len(non_empty) == 1 and len(text) <= 4:
        return True

    return False


def merge_row_words(left: List[Dict], right: List[Dict]) -> List[Dict]:
    merged = left + right
    return sorted(merged, key=lambda x: x["x0"])


def should_merge_poc_rows(current_row: List[Dict], next_row: List[Dict], boundaries: Dict[str, Tuple[float, float]]) -> bool:
    cur = row_words_to_dynamic_record(current_row, boundaries)
    nxt = row_words_to_dynamic_record(next_row, boundaries)

    cur_no = clean_text(cur["no"])
    nxt_no = clean_text(nxt["no"])

    cur_item = clean_text(cur["item_spec"])
    cur_amount = clean_text(cur["amount"])
    cur_unit_price = clean_text(cur["unit_price"])

    nxt_amount = clean_text(nxt["amount"])
    nxt_unit_price = clean_text(nxt["unit_price"])
    nxt_qty = clean_text(nxt["quantity"])
    nxt_unit = clean_text(nxt["unit"])

    if is_date_fragment_text(nxt["raw_row"]):
        return False

    if is_header_line(nxt["raw_row"]):
        return False

    if "小計" in nxt["raw_row"] or "小 計" in nxt["raw_row"]:
        return False

    current_is_stub = bool(re.match(r"^\d+$", cur_no)) and (
        not cur_amount or not cur_unit_price or len(cur_item) <= 30
    )

    next_is_cont = (
        not nxt_no and
        (
            (nxt_amount and is_money_token(nxt_amount)) or
            (nxt_unit_price and is_money_token(nxt_unit_price))
        ) and
        (nxt_qty or nxt_unit or nxt_unit_price or nxt_amount)
    )

    if current_is_stub and next_is_cont:
        return True

    if bool(cur_no) and not nxt_no and row_has_tail_value_pattern(nxt):
        return True

    return False


def merge_split_rows_for_poc(rows: List[List[Dict]], boundaries: Dict[str, Tuple[float, float]]) -> List[List[Dict]]:
    if not rows:
        return []

    merged_rows = []
    i = 0

    while i < len(rows):
        current = rows[i]

        if i + 1 < len(rows):
            nxt = rows[i + 1]
            if should_merge_poc_rows(current, nxt, boundaries):
                merged_rows.append(merge_row_words(current, nxt))
                i += 2
                continue

        merged_rows.append(current)
        i += 1

    return merged_rows


def detail_score(record: Dict) -> int:
    score = 0

    no_text = clean_text(record.get("no", ""))
    item_text = clean_text(record.get("item_spec", ""))
    qty = clean_text(record.get("quantity", ""))
    unit = clean_text(record.get("unit", ""))
    unit_price = clean_text(record.get("unit_price", ""))
    amount = clean_text(record.get("amount", ""))

    if no_text and re.match(r"^\d+", no_text):
        score += 2

    if item_text:
        score += 1

    if qty and (is_quantity_token(qty) or re.search(r"\d", qty)):
        score += 2

    if unit:
        score += 1

    if unit_price and is_money_token(unit_price):
        score += 2

    if amount and is_money_token(amount):
        score += 3

    return score


def is_poc_detail_record(record: Dict) -> bool:
    text = clean_text(record.get("raw_row", ""))
    score = detail_score(record)

    if is_date_fragment_text(text):
        return False

    if "小計" in text or "小 計" in text:
        return False

    if clean_text(record.get("amount", "")) and is_money_token(record.get("amount", "")):
        return score >= 6

    return score >= 7


def classify_record_status(record: Dict) -> str:
    score = detail_score(record)
    amount = clean_text(record.get("amount", ""))
    unit_price = clean_text(record.get("unit_price", ""))
    qty = clean_text(record.get("quantity", ""))
    unit = clean_text(record.get("unit", ""))
    no = clean_text(record.get("no", ""))
    item_spec = clean_text(record.get("item_spec", ""))

    full_ok = (
        no and item_spec and qty and unit and unit_price and amount and
        is_money_token(unit_price) and is_money_token(amount)
    )
    if full_ok:
        return "complete"

    rescued_ok = (
        item_spec and
        qty and unit and unit_price and amount and
        is_money_token(unit_price) and is_money_token(amount) and
        score >= 6
    )
    if rescued_ok:
        return "rescued"

    return "review"


def extract_page_rows_with_boundaries(page) -> Tuple[List[List[Dict]], Dict[str, Tuple[float, float]], str]:
    words = extract_words_debug_from_page(page)
    rows = group_words_by_row(words)

    if not rows:
        raise ValueError("ページ内にwordが見つかりませんでした。")

    header_row = find_header_row_words(rows)
    if header_row is None:
        raise ValueError("ヘッダー行が見つかりませんでした。")

    boundaries = build_dynamic_column_boundaries(header_row)
    header_text = row_words_to_text(header_row)
    header_top = min(w["top"] for w in header_row)

    body_rows = []
    for row in rows:
        row_top = min(w["top"] for w in row)
        if row_top <= header_top:
            continue
        body_rows.append(row)

    return body_rows, boundaries, header_text


def detect_major_category_from_page_rows(page_rows: List[List[Dict]], header_text: str = "") -> str:
    lines = []
    if header_text:
        lines.append(header_text)

    # 実際には body_rows には header より上が無いので、ここは fallback 用
    for row in page_rows[:5]:
        lines.append(row_words_to_text(row))

    # 既存ロジックに寄せた fallback
    return detect_major_category(lines)


def parse_record_to_output_row(record: Dict, file_name: str, estimate_date: str, page_num: int, major_category: str) -> Dict:
    status = classify_record_status(record)
    needs_review = 0 if status == "complete" else 1

    return {
        "file_name": file_name,
        "estimate_date": estimate_date,
        "page": page_num,
        "major_category": major_category,
        "no": clean_text(record.get("no", "")),
        "item_spec": clean_text(record.get("item_spec", "")),
        "quantity": clean_text(record.get("quantity", "")),
        "unit": clean_text(record.get("unit", "")),
        "unit_price": normalize_money_token(record.get("unit_price", "")),
        "amount": normalize_money_token(record.get("amount", "")),
        "raw_row": clean_text(record.get("raw_row", "")),
        "needs_review": needs_review,
    }


def fallback_process_page_by_text(page, file_name: str, estimate_date: str, page_num: int) -> List[Dict]:
    lines = extract_lines_from_page(page)
    lines = merge_multiline_details(lines)
    major_category = detect_major_category(lines)

    rows = []
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


def process_pdf(file_name: str, file_bytes: bytes) -> List[Dict]:
    rows = []

    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        estimate_date = extract_estimate_date(pdf)
        skip_first_page = should_skip_first_page(pdf)

        previous_major_category = ""

        for page_num, page in enumerate(pdf.pages, start=1):
            if page_num == 1 and skip_first_page:
                continue

            try:
                body_rows, boundaries, header_text = extract_page_rows_with_boundaries(page)

                filtered_rows = [row for row in body_rows if not is_poc_noise_row(row, boundaries)]
                merged_rows = merge_split_rows_for_poc(filtered_rows, boundaries)

                page_records = []
                for row in merged_rows:
                    rec = row_words_to_dynamic_record(row, boundaries)
                    rec["detail_score"] = detail_score(rec)
                    rec["parse_status"] = classify_record_status(rec)
                    page_records.append(rec)

                page_lines_for_title = extract_lines_from_page(page)
                major_category = detect_major_category(page_lines_for_title)
                if not major_category:
                    major_category = previous_major_category

                for rec in page_records:
                    if rec["parse_status"] not in {"complete", "rescued"}:
                        continue

                    rows.append(parse_record_to_output_row(
                        rec,
                        file_name=file_name,
                        estimate_date=estimate_date,
                        page_num=page_num,
                        major_category=major_category,
                    ))

                if major_category:
                    previous_major_category = major_category

            except Exception:
                fallback_rows = fallback_process_page_by_text(
                    page=page,
                    file_name=file_name,
                    estimate_date=estimate_date,
                    page_num=page_num,
                )
                rows.extend(fallback_rows)

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


# ============================================================
# PoC 検証モード用関数群
# ============================================================

def list_pdf_names_in_uploaded_file(uploaded_file) -> List[str]:
    name = uploaded_file.name

    if name.lower().endswith(".pdf"):
        return [name]

    if name.lower().endswith(".zip"):
        zip_bytes = uploaded_file.getvalue()
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as z:
            return sorted([n for n in z.namelist() if n.lower().endswith(".pdf")])

    return []


def get_pdf_bytes_from_uploaded_file(uploaded_file, pdf_name: str) -> bytes:
    name = uploaded_file.name

    if name.lower().endswith(".pdf"):
        if name != pdf_name:
            raise ValueError(f"指定PDF名が一致しません: {pdf_name}")
        return uploaded_file.getvalue()

    if name.lower().endswith(".zip"):
        zip_bytes = uploaded_file.getvalue()
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as z:
            return z.read(pdf_name)

    raise ValueError("PDFまたはZIPではないファイルです。")


def run_single_page_poc(uploaded_file, pdf_name: str, page_num: int) -> Tuple[pd.DataFrame, pd.DataFrame, Dict[str, Tuple[float, float]], str]:
    pdf_bytes = get_pdf_bytes_from_uploaded_file(uploaded_file, pdf_name)

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        if page_num < 1 or page_num > len(pdf.pages):
            raise ValueError(f"ページ番号が不正です。指定: {page_num}, 総ページ数: {len(pdf.pages)}")

        page = pdf.pages[page_num - 1]

        words = extract_words_debug_from_page(page)
        rows = group_words_by_row(words)
        header_row = find_header_row_words(rows)

        if header_row is None:
            raise ValueError("ヘッダー行が見つかりませんでした。")

        boundaries = build_dynamic_column_boundaries(header_row)
        header_text = row_words_to_text(header_row)

        words_df = pd.DataFrame(words, columns=["text", "x0", "x1", "top", "bottom"])

        header_top = min(w["top"] for w in header_row)

        body_rows = []
        for row in rows:
            row_top = min(w["top"] for w in row)
            if row_top <= header_top:
                continue
            body_rows.append(row)

        filtered_rows = [row for row in body_rows if not is_poc_noise_row(row, boundaries)]
        merged_rows = merge_split_rows_for_poc(filtered_rows, boundaries)

        records = []
        for row in merged_rows:
            row_top = min(w["top"] for w in row)
            rec = row_words_to_dynamic_record(row, boundaries)
            rec["row_top"] = row_top
            rec["detail_score"] = detail_score(rec)
            rec["is_detail_like"] = 1 if is_poc_detail_record(rec) else 0
            records.append(rec)

        record_df = pd.DataFrame(records)

        detail_df = record_df[record_df["is_detail_like"] == 1].copy() if not record_df.empty else pd.DataFrame(
            columns=["no", "item_spec", "quantity", "unit", "unit_price", "amount", "raw_row", "row_top", "detail_score", "is_detail_like"]
        )

        return words_df, detail_df, boundaries, header_text


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

    if looks_like_cover_noise(raw):
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

st.divider()
st.subheader("1ページ検証モード（PoC）")
st.write("通常抽出とは別に、1ページだけを対象に座標ベースの列抽出を検証します。")
enable_poc_mode = st.checkbox("1ページ検証モードを使う", value=False)

if uploaded_files:
    if enable_poc_mode:
        st.info("検証モードでは、アップロードしたファイルのうち1つを選び、1ページだけ座標ベースで解析します。")

        uploaded_name_map = {f.name: f for f in uploaded_files}
        selected_uploaded_name = st.selectbox(
            "検証対象ファイル",
            options=list(uploaded_name_map.keys()),
            index=0
        )
        selected_uploaded_file = uploaded_name_map[selected_uploaded_name]

        pdf_names = list_pdf_names_in_uploaded_file(selected_uploaded_file)

        if pdf_names:
            selected_pdf_name = st.selectbox(
                "ZIP内PDF名（PDF単体アップロード時はそのファイル名）",
                options=pdf_names,
                index=0
            )

            selected_page_num = st.number_input(
                "検証ページ番号",
                min_value=1,
                value=2,
                step=1
            )

            if st.button("このページを検証する"):
                try:
                    words_df, detail_df, boundaries, header_text = run_single_page_poc(
                        selected_uploaded_file,
                        selected_pdf_name,
                        int(selected_page_num)
                    )

                    st.success("1ページ検証が完了しました。")

                    st.write("検出ヘッダー")
                    st.code(header_text)

                    st.write("列境界")
                    boundary_df = pd.DataFrame(
                        [
                            {"column": k, "x_min": v[0], "x_max": v[1]}
                            for k, v in boundaries.items()
                        ]
                    )
                    st.dataframe(boundary_df, use_container_width=True, height=250)

                    st.write("word座標一覧")
                    st.dataframe(words_df, use_container_width=True, height=300)

                    st.write("仮抽出結果（座標ベース）")
                    st.dataframe(detail_df, use_container_width=True, height=350)

                except Exception as e:
                    st.error(f"検証モードでエラーが発生しました: {e}")
        else:
            st.warning("このファイルからPDFが見つかりませんでした。")

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
