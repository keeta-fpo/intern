# Streamlit Frontend for Robust OCR → Financials Excel
# ----------------------------------------------------
# Features
# - Upload PDF or image(s)
# - OCR via OpenTyphoon API **or** local Tesseract (no-cost fallback)
# - Robust parsing of messy tables → tidy DataFrames
# - Extract audit opinion
# - Balance validation + unmapped heading report
# - Preview & Download Excel

# How to run:
#   1) pip install -U streamlit pdf2image pillow pytesseract pandas rapidfuzz beautifulsoup4 lxml xlsxwriter
#   2) (Linux) apt-get install -y poppler-utils tesseract-ocr
#   3) streamlit run streamlit_frontend_app.py

import io
import re
import json
import tempfile
from io import StringIO
from typing import List, Tuple, Optional

import pandas as pd
import streamlit as st
from bs4 import BeautifulSoup
from PIL import Image
import pytesseract
from pytesseract import TesseractError, TesseractNotFoundError

# Optional fuzzy matcher
try:
    from rapidfuzz import fuzz, process
    USE_FUZZY = True
except Exception:
    USE_FUZZY = False

# ---------- OCR Backends ----------
# OpenTyphoon API
import requests

def ocr_opentyphoon(image_bytes: bytes, api_key: str, task_type: str = "default",
                    max_tokens: int = 16000, temperature: float = 0.1,
                    top_p: float = 0.6, repetition_penalty: float = 1.2,
                    pages: List[int] | None = None) -> str:
    url = "https://api.opentyphoon.ai/v1/ocr"
    files = {"file": ("page.png", image_bytes)}
    data = {
        "task_type": task_type,
        "max_tokens": str(max_tokens),
        "temperature": str(temperature),
        "top_p": str(top_p),
        "repetition_penalty": str(repetition_penalty),
    }
    if pages:
        data["pages"] = json.dumps(pages)
    headers = {"Authorization": f"Bearer {api_key}"}
    resp = requests.post(url, files=files, data=data, headers=headers, timeout=120)
    resp.raise_for_status()
    result = resp.json()
    texts = []
    for page_result in result.get("results", []):
        if page_result.get("success") and page_result.get("message"):
            content = page_result["message"]["choices"][0]["message"]["content"]
            try:
                parsed = json.loads(content)
                text = parsed.get("natural_text", content)
            except json.JSONDecodeError:
                text = content
            texts.append(text)
    return "\n".join(texts)

# Local Tesseract fallback
try:
    import pytesseract
    TESSERACT_AVAILABLE = True
except Exception:
    TESSERACT_AVAILABLE = False

from pdf2image import convert_from_bytes

# ---------- Parsing Utilities (Robust) ----------
THAI_DIGITS = str.maketrans("๐๑๒๓๔๕๖๗๘๙", "0123456789")
DASHES = {"-", "–", "—", ""}

ALIASES_REGEX = {
    "เงินสดและรายการเทียบเท่าเงินสด": [r"เงินสด(และ)?รายการเทียบเท่า(เงินสด)?", r"cash.*equivalent"],
    "ลูกหนี้การค้าและลูกหนี้อื่น": [r"ลูกหนี้การค้า(และ|,)?\s*ลูกหนี้อื่น", r"ลูกหนี้การค้า(?!.*สุทธิ)", r"ลูกหนี้อื่น(?!.*สุทธิ)"],
    "รวมสินทรัพย์หมุนเวียน": [r"^รวม\s*สินทรัพย์\s*หมุนเวียน$"],
    "ที่ดิน อาคารและอุปกรณ์": [r"(ที่ดิน|อาคาร|อุปกรณ์)"],
    "สินทรัพย์ไม่หมุนเวียนอื่น": [r"สินทรัพย์ไม่หมุนเวียน(อื่น|รวมอื่น)"],
    "รวมหนี้สินหมุนเวียน": [r"^รวม\s*หนี้สิน\s*หมุนเวียน$"],
    "เงินกู้ยืมระยะสั้น": [r"เงินกู้(ยืม)?\s*ระยะสั้น"],
    "ส่วนของหนี้สินระยะยาวที่ถึงกำหนดชำระภายในหนึ่งปี": [r"หนี้สินระยะยาว.*(ถึงกำหนด|ครบกำหนด).*หนึ่งปี"],
}

KW_ASSET = r"(สินทรัพย์|asset)"
KW_LIAB_EQUITY = r"(หนี้สิน|ส่วนของผู้ถือหุ้น|liabilit|equity)"
KW_PL = r"(รายได้|กำไร|ขาดทุน|income|profit|loss)"


def normalize_text(s: str) -> str:
    s = (s or "").lower()
    s = re.sub(r"\s+", " ", s)
    s = s.replace("ฯ", "เวียน")
    return s.strip()


def map_heading_to_canonical(item: str, threshold: int = 88) -> Tuple[str | None, int]:
    t = normalize_text(item)
    for canonical, patterns in ALIASES_REGEX.items():
        for p in patterns:
            if re.search(p, t):
                return canonical, 100
    if USE_FUZZY:
        candidates = list(ALIASES_REGEX.keys())
        best = process.extractOne(t, candidates, scorer=fuzz.token_set_ratio)
        if best and best[1] >= threshold:
            return best[0], best[1]
    return None, 0


def apply_heading_mapping(df: pd.DataFrame) -> Tuple[pd.DataFrame, list]:
    if "รายการ" not in df.columns:
        return df, []
    mapped, unmapped = [], []
    for x in df["รายการ"].astype(str).tolist():
        canon, score = map_heading_to_canonical(x)
        if canon:
            mapped.append(canon)
        else:
            mapped.append(x)
            unmapped.append(x)
    out = df.copy()
    out["หัวข้อ_มาตรฐาน"] = mapped
    return out, sorted(set(unmapped))


def parse_number(raw: str, unit_multiplier=1.0):
    if pd.isna(raw):
        return pd.NA
    s = str(raw).strip().translate(THAI_DIGITS)
    if s in DASHES:
        return pd.NA
    neg = False
    if re.match(r"^\(\s*.+\s*\)$", s):
        neg = True
        s = s.strip()[1:-1].strip()
    s = s.replace(",", "").replace("\u200b", "")
    m = re.search(r"-?\d+(\.\d+)?", s)
    if not m:
        return pd.NA
    val = float(m.group(0)) * unit_multiplier
    return -val if neg else val


def detect_unit_multiplier(extracted_text: str):
    hint = extracted_text.lower()
    if "ล้านบาท" in hint:
        return 1_000_000.0
    if "พันบาท" in hint:
        return 1_000.0
    return 1.0


def extract_year_text(s):
    m = re.search(r"(25\d{2}|20\d{2})", str(s))
    return m.group(0) if m else str(s)


def rename_year_columns(df: pd.DataFrame):
    new_cols, seen = [], {}
    for c in df.columns:
        if c in ["รายการ", "หมายเหตุ"]:
            new_cols.append(c)
            continue
        col = extract_year_text(c)
        if col in seen:
            seen[col] += 1
            col = f"{col}_{seen[col]}"
        else:
            seen[col] = 1
        new_cols.append(col)
    df.columns = new_cols
    return df


def guess_plaintext_tables(text: str):
    lines = [l for l in text.splitlines() if l.strip()]
    blocks, cur = [], []
    for ln in lines:
        if re.search(r"\s{2,}|\t", ln):
            cur.append(ln)
        else:
            if cur:
                blocks.append("\n".join(cur))
                cur = []
    if cur:
        blocks.append("\n".join(cur))
    return blocks


def plaintext_block_to_df(block: str):
    rows = []
    for ln in block.splitlines():
        parts = re.split(r"\t+|\s{2,}", ln.strip())
        if len(parts) >= 2:
            rows.append(parts)
    if not rows:
        return None
    max_len = max(len(r) for r in rows)
    for r in rows:
        if len(r) < max_len:
            r.extend([""] * (max_len - len(r)))
    cols = [f"col_{i}" for i in range(max_len)]
    df = pd.DataFrame(rows, columns=cols)
    df = df.rename(columns={df.columns[0]: "รายการ"})
    return df


def read_tables_from_html(extracted_text: str) -> List[pd.DataFrame]:
    tables: List[pd.DataFrame] = []
    # Try direct read_html
    try:
        dfs = pd.read_html(StringIO(extracted_text))
        tables.extend(dfs)
    except Exception:
        pass
    # Parse <table> tags precisely
    try:
        soup = BeautifulSoup(extracted_text, "html.parser")
        for t in soup.find_all("table"):
            try:
                dfx = pd.read_html(StringIO(str(t)))  # wrap with StringIO to silence FutureWarning
                tables.extend(dfx)
            except Exception:
                continue
    except Exception:
        pass
    # Fallback plaintext → pseudo tables
    if not tables:
        blocks = guess_plaintext_tables(extracted_text)
        for b in blocks:
            try:
                df = plaintext_block_to_df(b)
                if df is not None and df.shape[1] >= 2:
                    tables.append(df)
            except Exception:
                continue
    return tables


def tidy_table(df_raw: pd.DataFrame, unit_multiplier: float) -> pd.DataFrame:
    df = df_raw.copy()
    if df.empty or df.shape[1] == 0:
        return pd.DataFrame()
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = ["_".join([str(x) for x in tup if str(x) != "nan"]).strip() for tup in df.columns]
    # choose an item column
    text_cols = [c for c in df.columns if df[c].dtype == "object"]
    candidate_item_col = (text_cols[0] if text_cols else df.columns[0])
    if candidate_item_col != "รายการ":
        df = df.rename(columns={candidate_item_col: "รายการ"})
    df = rename_year_columns(df)
    # detect numeric/year columns
    year_cols = []
    for c in df.columns:
        if c in ["รายการ", "หมายเหตุ"]:
            continue
        if re.fullmatch(r"\d{4}(_\d+)?", str(c)):
            year_cols.append(c)
            continue
        sample = df[c].astype(str).head(20).tolist()
        numeric_like = 0
        for v in sample:
            vv = v.translate(THAI_DIGITS).replace(",", "").replace("\u200b", "").strip()
            if vv in DASHES or re.search(r"-?\d+(\.\d+)?", vv):
                numeric_like += 1
        if len(sample) > 0 and numeric_like / len(sample) >= 0.6:
            year_cols.append(c)
    for c in year_cols:
        df[c] = df[c].apply(lambda v: parse_number(v, unit_multiplier=unit_multiplier))
    if "หมายเหตุ" not in df.columns:
        df["หมายเหตุ"] = pd.NA
    has_item_col = "รายการ" in df.columns
    if year_cols:
        numeric_mask = ~df[year_cols].isna().all(axis=1)
        item_mask = df["รายการ"].notna() if has_item_col else False
        mask_keep = numeric_mask | item_mask
        df = df[mask_keep].reset_index(drop=True)
    ordered = (["รายการ"] if has_item_col else []) + year_cols + ["หมายเหตุ"]
    ordered = [c for c in ordered if c in df.columns]
    df = df[ordered]
    if "รายการ" not in df.columns:
        df.insert(0, "รายการ", pd.NA)
    return df


def classify_table(df: pd.DataFrame) -> str:
    if "รายการ" not in df.columns:
        return "unknown"
    text = " ".join(df["รายการ"].astype(str).tolist()).lower()
    score_asset = bool(re.search(KW_ASSET, text))
    score_liab  = bool(re.search(KW_LIAB_EQUITY, text))
    score_pl    = bool(re.search(KW_PL, text))
    if score_asset and not (score_liab or score_pl):
        return "asset"
    if score_liab and not score_asset:
        return "liab_eq"
    if score_pl:
        return "pl"
    return "unknown"


def extract_audit_opinion_html(extracted_text: str) -> str:
    soup = BeautifulSoup(extracted_text, "html.parser")
    full_text = soup.get_text("\n", strip=True)
    return extract_audit_opinion_text(full_text)


def extract_audit_opinion_text(full_text: str) -> str:
    txt = re.sub(r"[ \t]+", " ", full_text)
    patterns = [
        (r"(ความเห็นของผู้สอบบัญชี[^\n]*)([\s\S]+?)(งบการเงิน|งบ\s)", True),
        (r"(ความคิดเห็นของผู้สอบบัญชี[^\n]*)([\s\S]+?)(งบการเงิน|งบ\s)", True),
    ]
    for pat, _ in patterns:
        m = re.search(pat, txt)
        if m:
            head = m.group(1).strip()
            body = m.group(2).strip()
            return f"{head}\n{body}".strip()
    lines = [l.strip() for l in txt.splitlines() if l.strip()]
    found = [l for l in lines if ("ผู้สอบบัญชี" in l or "ความเห็น" in l or "ความคิดเห็น" in l)]
    return "\n".join(found) if found else "ไม่มีความคิดเห็นผู้สอบบัญชี"


def basic_balance_check(bs_df: pd.DataFrame):
    if bs_df.empty:
        return {"status": "empty", "notes": "no balance sheet rows"}
    asset_total_rows = bs_df[bs_df["รายการ"].astype(str).str.contains("รวมสินทรัพย์")]
    liab_rows = bs_df[bs_df["รายการ"].astype(str).str.contains("รวมหนี้สิน")]
    equity_rows = bs_df[bs_df["รายการ"].astype(str).str.contains("ส่วนของผู้ถือหุ้น|รวมส่วนของผู้ถือหุ้น")]
    year_cols = [c for c in bs_df.columns if re.fullmatch(r"\d{4}(_\d+)?", str(c))]
    issues = []
    for c in year_cols:
        a = asset_total_rows[c].dropna()
        l = liab_rows[c].dropna()
        e = equity_rows[c].dropna()
        if not a.empty and (not l.empty or not e.empty):
            lhs = a.iloc[-1]
            rhs = (l.iloc[-1] if not l.empty else 0) + (e.iloc[-1] if not e.empty else 0)
            if pd.notna(lhs) and pd.notna(rhs):
                if abs(lhs - rhs) > max(1e-6, abs(lhs) * 0.02):
                    issues.append(f"ปี {c}: รวมสินทรัพย์ {lhs:,.0f} != หนี้สิน+ทุน {rhs:,.0f}")
    status = "ok" if not issues else "mismatch"
    return {"status": status, "notes": "; ".join(issues) if issues else ""}


def process_extracted_text(extracted_text: str, export_path_excel="งบการเงิน_clean.xlsx"):
    unit_mult = detect_unit_multiplier(extracted_text)
    tables = read_tables_from_html(extracted_text)

    assets, liab_eq, pl, unknowns = [], [], [], []
    unmapped_all = set()

    for t in tables:
        t = tidy_table(t, unit_multiplier=unit_mult)
        if t.empty or t.shape[1] < 2:
            continue
        t2, unmapped = apply_heading_mapping(t)
        unmapped_all.update(unmapped)
        kind = classify_table(t2)
        if kind == "asset":
            assets.append(t2)
        elif kind == "liab_eq":
            liab_eq.append(t2)
        elif kind == "pl":
            pl.append(t2)
        else:
            unknowns.append(t2)

    asset_df = pd.concat(assets, ignore_index=True) if assets else pd.DataFrame()
    liab_eq_df = pd.concat(liab_eq, ignore_index=True) if liab_eq else pd.DataFrame()
    pl_df = pd.concat(pl, ignore_index=True) if pl else pd.DataFrame()

    if not asset_df.empty or not liab_eq_df.empty:
        bs_df = pd.concat([d for d in [asset_df, liab_eq_df] if not d.empty], ignore_index=True)
    else:
        bs_df = pd.DataFrame()

    try:
        audit_opinion = extract_audit_opinion_html(extracted_text)
    except Exception:
        audit_opinion = extract_audit_opinion_text(extracted_text)

    bs_check = basic_balance_check(bs_df) if not bs_df.empty else {"status": "no_bs", "notes": ""}

    with pd.ExcelWriter(export_path_excel, engine="xlsxwriter") as writer:
        if not bs_df.empty:
            bs_df.to_excel(writer, sheet_name="งบแสดงฐานะการเงิน", index=False)
        if not pl_df.empty:
            pl_df.to_excel(writer, sheet_name="งบกำไรขาดทุน", index=False)
        if unknowns:
            pd.concat(unknowns, ignore_index=True).to_excel(writer, sheet_name="unknown_tables", index=False)
        meta_rows = [
            {"key": "unit_multiplier", "value": unit_mult},
            {"key": "balance_check", "value": bs_check["status"]},
            {"key": "balance_notes", "value": bs_check["notes"]},
            {"key": "unmapped_headings", "value": "; ".join(sorted(unmapped_all)) if unmapped_all else ""},
        ]
        pd.DataFrame(meta_rows).to_excel(writer, sheet_name="meta", index=False)
        pd.DataFrame([{"audit_opinion": audit_opinion}]).to_excel(writer, sheet_name="audit_opinion", index=False)
        wb = writer.book
        num_fmt = wb.add_format({"num_format": "#,##0"})
        for sheet in ["งบแสดงฐานะการเงิน", "งบกำไรขาดทุน", "unknown_tables"]:
            if sheet in writer.sheets:
                ws = writer.sheets[sheet]
                ws.set_column(0, 0, 50)
                for col_idx in range(1, 50):
                    ws.set_column(col_idx, col_idx, 18, num_fmt)

    return {
        "balance_check": bs_check,
        "unmapped": sorted(unmapped_all),
        "export_path": export_path_excel,
        "has_bs": not bs_df.empty,
        "has_pl": not pl_df.empty,
        "unknown_count": len(unknowns),
    }

def process_image_ocr(pil_img: Image.Image, tess_langs: str = 'tha+eng') -> Optional[str]:
    """Process image with Tesseract OCR.
    
    Args:
        pil_img: PIL Image object
        tess_langs: Tesseract language string
    Returns:
        Extracted text or None if error occurs
    """
    try:
        text = pytesseract.image_to_string(pil_img, lang=tess_langs)
        return text
    except (TesseractError, TesseractNotFoundError) as e:
        st.error(f"OCR Error: {str(e)}")
        st.info("Please check if Tesseract is properly installed")
        return None

# ---------- Streamlit UI ----------
st.set_page_config(page_title="OCR → Financials Excel", layout="wide")
st.title("📄 OCR → งบการเงิน (Excel)")

with st.sidebar:
    st.header("ตั้งค่า OCR")
    ocr_backend = st.radio("เลือกวิธี OCR", ["OpenTyphoon API", "Tesseract (ออฟไลน์)"]) 
    if ocr_backend == "OpenTyphoon API":
        api_key = st.text_input("OpenTyphoon API Key", type="password")
        task_type = st.selectbox("task_type", ["default", "table", "text"] , index=0)
        max_tokens = st.number_input("max_tokens", 1000, 32000, 16000, step=1000)
        temperature = st.number_input("temperature", 0.0, 1.0, 0.1, step=0.1)
        top_p = st.number_input("top_p", 0.0, 1.0, 0.6, step=0.05)
        repetition_penalty = st.number_input("repetition_penalty", 0.5, 2.0, 1.2, step=0.1)
    else:
        if not TESSERACT_AVAILABLE:
            st.warning("⚠️ ไม่พบ pytesseract: ติดตั้งด้วย `pip install pytesseract` และลง `tesseract-ocr` ในระบบ")
        tess_langs = st.text_input("ภาษา OCR (เช่น tha+eng)", value="tha+eng")

    st.header("การแปลง PDF")
    dpi = st.slider("DPI (สำหรับ PDF → ภาพ)", 200, 500, 350, step=50)

uploaded = st.file_uploader("อัปโหลด PDF หรือรูปภาพ (PNG/JPG)", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=False)

if uploaded:
    st.info("กดปุ่ม 'ประมวลผล' เพื่อเริ่ม OCR และสร้าง Excel")
    if st.button("🚀 ประมวลผล"):
        try:
            file_bytes = uploaded.read()
            pages_images: List[Image.Image] = []
            if uploaded.type == "application/pdf":
                pages_images = convert_from_bytes(file_bytes, dpi=dpi)
            else:
                pages_images = [Image.open(io.BytesIO(file_bytes)).convert("RGB")]

            extracted_parts: List[str] = []
            prog = st.progress(0)
            for i, pil_img in enumerate(pages_images, start=1):
                buf = io.BytesIO()
                pil_img.save(buf, format="PNG")
                img_bytes = buf.getvalue()

                if ocr_backend == "OpenTyphoon API":
                    if not api_key:
                        st.error("กรุณาใส่ API Key ก่อน")
                        st.stop()
                    text = ocr_opentyphoon(img_bytes, api_key=api_key, task_type=task_type,
                                           max_tokens=max_tokens, temperature=temperature,
                                           top_p=top_p, repetition_penalty=repetition_penalty,
                                           pages=[i])
                else:
                    if not TESSERACT_AVAILABLE:
                        st.error("ยังไม่มี Tesseract ติดตั้งบนระบบนี้")
                        st.stop()
                    text = process_image_ocr(pil_img)

                extracted_parts.append(text)
                prog.progress(i / len(pages_images))

            extracted_text = "\n\n".join(extracted_parts)

            st.subheader("ตัวอย่างข้อความที่ OCR ได้ (ย่อ)")
            st.code(extracted_text[:3000] + ("..." if len(extracted_text) > 3000 else ""))

            st.subheader("🔧 แปลงเป็นตาราง + Excel")
            result = process_extracted_text(extracted_text, export_path_excel="งบการเงิน_clean.xlsx")

            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Balance check", result["balance_check"]["status"])
            with col2:
                st.metric("Unknown tables", result["unknown_count"])
            with col3:
                st.metric("Unmapped headings", len(result["unmapped"]))

            if result["unmapped"]:
                with st.expander("ดูหัวข้อที่ยังไม่แม็ป (สำหรับปรับ alias)"):
                    st.write(result["unmapped"]) 

            # Download buttons
            with open(result["export_path"], "rb") as f:
                st.download_button("⬇️ ดาวน์โหลดงบการเงิน (Excel)", data=f, file_name="งบการเงิน_clean.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            st.success("เสร็จแล้ว!")
        except Exception as e:
            st.exception(e)
else:
    st.caption("รองรับ PDF หรือรูปเดี่ยว (PNG/JPG)")
