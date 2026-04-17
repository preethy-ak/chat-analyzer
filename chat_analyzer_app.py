"""
Chat Analyzer Dashboard — Shopee & Lazada
==========================================
Graas.ai-themed Streamlit app for daily chat enquiry analysis.

Run:  streamlit run chat_analyzer_dashboard.py
Deps: pip install streamlit pandas openpyxl xlsxwriter
"""

import streamlit as st
import pandas as pd
import numpy as np
import re, io, warnings
from datetime import datetime, timedelta

warnings.filterwarnings("ignore")

# ───────────────────A──────────────────────────────────────────────────────────
# PAGE CONFIG — Graas.ai theme
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Chat Analyzer Daashboard | Graas.ai",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────────────
# CUSTOM CSS — Graas.ai brand colours (#1B2A4A navy, #00C4B4 teal, #FF6B35 orange)
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
/* ── Global ── */
html, body, [class*="css"] { font-family: 'Inter', 'Segoe UI', sans-serif; }
.main { background: #F4F6FB; }
.block-container { padding: 1.5rem 2rem; }

/* ── Top header bar ── */
.graas-header {
    background: linear-gradient(135deg, #1B2A4A 0%, #243554 100%);
    border-radius: 12px;
    padding: 1.2rem 1.8rem;
    margin-bottom: 1.5rem;
    display: flex;
    align-items: center;
    gap: 1rem;
}
.graas-header h1 { color: #fff; margin: 0; font-size: 1.5rem; font-weight: 700; }
.graas-header p  { color: #A8C0D6; margin: 0; font-size: 0.85rem; }
.graas-logo { color: #00C4B4; font-size: 2rem; }

/* ── Metric cards ── */
.metric-row { display: flex; gap: 1rem; margin-bottom: 1.5rem; flex-wrap: wrap; }
.metric-card {
    background: #fff;
    border-radius: 10px;
    padding: 1rem 1.3rem;
    flex: 1;
    min-width: 150px;
    border-left: 4px solid #00C4B4;
    box-shadow: 0 2px 8px rgba(0,0,0,0.06);
}
.metric-card.orange { border-left-color: #FF6B35; }
.metric-card.red    { border-left-color: #E74C3C; }
.metric-card.navy   { border-left-color: #1B2A4A; }
.metric-card.green  { border-left-color: #27AE60; }
.metric-val { font-size: 1.9rem; font-weight: 800; color: #1B2A4A; }
.metric-label { font-size: 0.78rem; color: #7A8EA8; font-weight: 500; text-transform: uppercase; letter-spacing: 0.5px; }
.metric-sub { font-size: 0.75rem; color: #A0AEC0; margin-top: 2px; }

/* ── Section titles ── */
.section-title {
    font-size: 1rem;
    font-weight: 700;
    color: #1B2A4A;
    border-bottom: 2px solid #00C4B4;
    padding-bottom: 0.4rem;
    margin: 1.5rem 0 1rem;
}

/* ── Priority badges ── */
.badge-high   { background:#FDECEA; color:#C0392B; padding:2px 8px; border-radius:12px; font-size:0.75rem; font-weight:600; }
.badge-medium { background:#FEF9E7; color:#D68910; padding:2px 8px; border-radius:12px; font-size:0.75rem; font-weight:600; }
.badge-low    { background:#EAF4FB; color:#2980B9; padding:2px 8px; border-radius:12px; font-size:0.75rem; font-weight:600; }

/* ── Sentiment ── */
.sent-pos { color:#27AE60; font-weight:600; }
.sent-neu { color:#7F8C8D; font-weight:600; }
.sent-neg { color:#C0392B; font-weight:600; }

/* ── Sidebar ── */
.css-1d391kg { background: #1B2A4A !important; }
section[data-testid="stSidebar"] { background: #1B2A4A !important; }
section[data-testid="stSidebar"] * { color: #D4E6F1 !important; }
section[data-testid="stSidebar"] .stSelectbox label,
section[data-testid="stSidebar"] .stMultiSelect label { color: #A8C0D6 !important; font-size: 0.82rem; }

/* ── Tabs ── */
.stTabs [data-baseweb="tab-list"] { background: #fff; border-radius:8px; padding:4px; gap:4px; }
.stTabs [data-baseweb="tab"] { border-radius:6px; padding:6px 18px; font-weight:600; color:#7A8EA8; }
.stTabs [aria-selected="true"] { background:#00C4B4 !important; color:#fff !important; }

/* ── Suggested reply box ── */
.reply-box {
    background: #F0FBF9;
    border: 1px solid #00C4B4;
    border-radius: 8px;
    padding: 0.9rem 1rem;
    font-size: 0.85rem;
    color: #1B2A4A;
    line-height: 1.6;
    margin-top: 0.5rem;
}
.reply-label { font-size:0.75rem; color:#00C4B4; font-weight:700; text-transform:uppercase; margin-bottom:4px; }

/* ── Upload area ── */
.upload-area {
    background: #fff;
    border: 2px dashed #00C4B4;
    border-radius: 12px;
    padding: 2rem;
    text-align: center;
    margin-bottom: 1rem;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────────────────────────

# Issue-type keyword mapping
ISSUE_KEYWORDS = {
    "Refund": [
        "refund", "คืนเงิน", "pengembalian dana", "dana kembali", "ibalik", "irefund",
        "bayar balik", "money back", "reimburse", "reimbursement",
    ],
    "Return": [
        "return", "คืนสินค้า", "retur", "rma", "send back", "ส่งคืน", "kembalikan",
        "return item", "product return",
    ],
    "Cancellation": [
        "cancel", "cancelled", "ยกเลิก", "batalkan", "batal", "cancellation",
        "cancel order", "ยกเลิกคำสั่งซื้อ",
    ],
    "Delay": [
        "delay", "late", "slow", "ช้า", "lambat", "belum sampai", "haven't received",
        "not arrived", "waiting", "รอนาน", "still waiting", "lama", "terlambat",
        "overdue", "not delivered yet", "ยังไม่ได้รับ", "belum diterima",
    ],
    "Damaged/Wrong Item": [
        "wrong item", "wrong product", "damaged", "broken", "defective",
        "สินค้าผิด", "ของเสีย", "ของแตก", "ของชำรุด", "rusak", "cacat",
        "salah barang", "salah produk", "not as described", "different item",
        "wrong size", "wrong colour", "wrong color", "different from picture",
    ],
    "Missing Item": [
        "missing", "not received", "didn't receive", "never received",
        "ไม่ได้รับ", "ของหาย", "hilang", "tidak diterima", "tidak ada", "kurang",
        "incomplete", "item missing", "package empty",
    ],
    "Payment Issue": [
        "payment", "ชำระเงิน", "bayar", "pembayaran", "charge", "double charge",
        "overcharged", "wrong charge", "billing", "invoice", "โอนเงิน", "จ่ายเงิน",
        "pay", "transfer", "deducted", "not paid",
    ],
    "Product Inquiry": [
        "how to", "how do", "วิธีใช้", "ราคา", "price", "size", "ขนาด",
        "สี", "colour", "color", "spec", "specification", "ingredient",
        "cara pakai", "ukuran", "warna", "harga", "stok", "stock", "available",
        "variant", "model", "version",
    ],
    "Promotion Issue": [
        "voucher", "promo", "discount", "coupon", "code", "sale", "offer",
        "โปรโมชั่น", "ส่วนลด", "โค้ด", "diskon", "kode promo", "cashback",
        "flash sale", "deal", "bundle",
    ],
    "Technical Issue": [
        "error", "bug", "cannot", "can't", "unable", "failed", "not working",
        "app issue", "website", "login", "checkout problem", "system",
        "ไม่สามารถ", "เกิดข้อผิดพลาด", "tidak bisa", "gagal", "eror",
    ],
    "Complaint": [
        "complain", "complaint", "terrible", "horrible", "awful", "worst",
        "ร้องเรียน", "ไม่พอใจ", "รำคาญ", "โกรธ", "disappointed",
        "frustrated", "unacceptable", "poor service", "bad service",
        "kecewa", "mengecewakan", "tidak puas", "buruk", "parah",
    ],
}

PRIORITY_MAP = {
    "High":   ["Refund", "Complaint", "Damaged/Wrong Item"],
    "Medium": ["Delay", "Missing Item", "Return", "Cancellation"],
    "Low":    ["Product Inquiry", "Promotion Issue", "Payment Issue", "Technical Issue"],
}

# Stalling phrases (seller still processing — NOT resolved)
STALLING_PATTERNS = [
    r"will (check|look|get back|follow up|investigate|verify|review|update)",
    r"let me (check|look into|verify|confirm|see)",
    r"(checking|looking into|investigating|following up|reviewing)",
    r"please (wait|hold on|allow us|bear with)",
    r"i will (check|get back|follow up|update)",
    r"we (are|will) (checking|looking|investigating|getting back|following up)",
    r"get back to you",
    r"bear with us",
    r"kindly (wait|allow|hold)",
    r"we'?ll? (check|look|get back|follow up)",
    r"akan (kami|segera) (cek|periksa|tindak lanjut|proses|hubungi)",
    r"mohon (tunggu|ditunggu|bersabar)",
    r"kami (sedang|akan) (cek|periksa|proses|tindak lanjut)",
    r"จะตรวจสอบ", r"กำลังตรวจสอบ", r"จะแจ้งกลับ",
    r"จะดำเนินการ", r"ขอตรวจสอบ", r"ขอเวลา",
    r"จะติดต่อกลับ", r"ติดตามให้", r"กำลังประสานงาน",
    r"escalat",
]

# Resolution phrases (conversation closed / solved)
RESOLUTION_PATTERNS = [
    r"refund (has been|was|is) (processed|completed|done|issued|approved)",
    r"(your|the) (order|item|package) (has been|was|is) (shipped|dispatched|replaced|delivered)",
    r"(issue|problem|case) (has been|was|is) (resolved|fixed|closed|sorted|handled)",
    r"(cancellation|cancel) (has been|was|is) (processed|done|completed|approved)",
    r"(we have|we've) (processed|completed|resolved|fixed|issued|sent)",
    r"please (expect|allow) (\d|few|some|a couple)",
    r"track.*link.*sent", r"tracking (number|id|code) (is|was|has been)",
    r"you (should|will) (receive|get) (it|your order|the item)",
    r"ดำเนินการเรียบร้อย", r"จัดการเรียบร้อย", r"แก้ไขเรียบร้อย",
    r"คืนเงินเรียบร้อย", r"ยกเลิกเรียบร้อย",
    r"sudah (diproses|selesai|dikirim|dikembalikan|dibatalkan)",
    r"telah (diproses|selesai|diselesaikan|dikirimkan)",
]

# Auto-reply detection
AUTO_REPLY_PATTERNS = [
    r"(thank you for contacting|thanks for reaching out).*auto",
    r"auto.?reply", r"automated (response|message|reply)",
    r"we'?ll? (get back|respond) (to you )?(within|in|shortly|soon)",
    r"our (team|agent).*(will|shall) (respond|reply|contact)",
    r"welcome to .*(official store|store).*\nhow (can|may) (we|i) help",
    r"สวัสดีค่ะ.*แอดมิน.*ยินดีให้บริการ",
    r"ยินดีต้อนรับ.*ร้าน",
    r"hi.{0,30}welcome to.{0,40}store",
]

# Positive / Negative sentiment keywords (multilingual)
POSITIVE_KWS = [
    "thank", "thanks", "great", "excellent", "awesome", "perfect", "love",
    "good", "nice", "happy", "satisfied", "wonderful", "amazing", "fantastic",
    "superb", "appreciate", "helpful", "fast", "quick", "well done", "recommend",
    "ขอบคุณ", "ดีมาก", "ประทับใจ", "พอใจ", "ยอดเยี่ยม", "ดีเลย", "ดีค่ะ", "ดีครับ",
    "terima kasih", "bagus", "mantap", "keren", "memuaskan", "puas", "oke baik",
    "salamat", "maganda", "ayos", "galing",
]

NEGATIVE_KWS = [
    "terrible", "worst", "angry", "disappointed", "frustrated", "cheated", "scam",
    "fraud", "fake", "broken", "damaged", "wrong item", "missing", "never received",
    "unacceptable", "horrible", "awful", "complain", "complaint", "refund",
    "ผิดหวัง", "โกรธ", "ไม่พอใจ", "แย่มาก", "แย่", "หลอกลวง", "ของเสีย",
    "ของปลอม", "ช้ามาก", "รอนาน", "สินค้าไม่ตรง", "ไม่ได้รับ", "ชำรุด",
    "tipu", "rusak", "cacat", "mengecewakan", "marah", "kecewa", "buruk", "parah",
    "salah", "tidak diterima", "hilang",
]

# Suggested replies per issue type
SUGGESTED_REPLIES = {
    "Refund": (
        "Thank you for reaching out, and we sincerely apologise for the inconvenience. "
        "We have reviewed your request and are pleased to confirm that your refund of [AMOUNT] "
        "has been initiated and will be reflected in your original payment method within 3–5 business days. "
        "Your order reference is [ORDER_ID]. We truly value your trust in us and hope to serve you better next time. "
        "If you have any further questions, please don't hesitate to reach out. 😊\n\n"
        "We'd love to hear your feedback — could you take a moment to rate your experience with us?"
    ),
    "Return": (
        "Thank you for contacting us about your return request. We're sorry to hear the product "
        "didn't meet your expectations. We've initiated the return process for order [ORDER_ID]. "
        "Please use the return label / return portal link we'll send to your registered email within 24 hours. "
        "Once we receive the item, the replacement or refund will be processed within 3–5 business days. "
        "We appreciate your patience and your continued support. 😊\n\n"
        "How would you rate your experience with us today?"
    ),
    "Cancellation": (
        "We've received your cancellation request for order [ORDER_ID]. We're sorry to see you go! "
        "Your order has been successfully cancelled and any payment made will be refunded within 3–5 business days. "
        "If you change your mind or need assistance with a future purchase, we're always here to help. 😊\n\n"
        "We'd appreciate your feedback — how was your experience with our team today?"
    ),
    "Delay": (
        "Thank you for your patience, and we sincerely apologise for the delay with your order [ORDER_ID]. "
        "We've checked with our logistics partner and your package is currently [STATUS]. "
        "Estimated delivery is [DATE]. We understand how frustrating delays can be and we truly appreciate your understanding. "
        "You can track your order in real time here: [TRACKING_LINK]. "
        "Please reach out if the delivery isn't received by [DATE+1] and we'll escalate immediately. 😊\n\n"
        "How was your experience with our support team today?"
    ),
    "Damaged/Wrong Item": (
        "We're truly sorry to hear that you received a damaged / incorrect item for order [ORDER_ID]. "
        "This is not the experience we want for you. To resolve this as quickly as possible, "
        "we've arranged a replacement to be dispatched within 1–2 business days. "
        "You do not need to return the incorrect / damaged item. "
        "We sincerely apologise for the inconvenience caused and will ensure this doesn't happen again. 😊\n\n"
        "Could you spare a moment to rate your support experience today?"
    ),
    "Missing Item": (
        "We're sorry to hear that your order [ORDER_ID] arrived with a missing item. "
        "We've raised an investigation with our fulfilment team and will have an update for you within 24 hours. "
        "In the meantime, we'll arrange a replacement or full refund, whichever you prefer. "
        "We apologise for this experience and truly appreciate your patience. 😊\n\n"
        "We'd love your feedback — how would you rate your experience with us today?"
    ),
    "Payment Issue": (
        "Thank you for flagging this payment concern. We've reviewed your account and order [ORDER_ID]. "
        "Our finance team has been notified and the discrepancy will be resolved within 2–3 business days. "
        "A confirmation will be sent to your registered email once completed. "
        "We apologise for any inconvenience and truly value your trust in us. 😊\n\n"
        "How was your experience with our support team today?"
    ),
    "Product Inquiry": (
        "Thank you for your interest in [PRODUCT_NAME]! "
        "Here are the details you requested: [DETAILS]. "
        "If you have more questions about specifications, sizing, or availability, "
        "please feel free to ask — we're happy to help you find the perfect product. 😊\n\n"
        "How can we assist you further today?"
    ),
    "Promotion Issue": (
        "Thank you for reaching out about the promotion. We're sorry for the confusion. "
        "We've reviewed your order [ORDER_ID] and confirmed that the discount of [AMOUNT] is applicable. "
        "The adjustment will be reflected within 24–48 hours. "
        "If the voucher code didn't apply correctly, please share it with us and we'll verify it right away. 😊\n\n"
        "How was your support experience today?"
    ),
    "Technical Issue": (
        "We apologise for the technical difficulty you're experiencing. "
        "Our team has been notified and is working on a resolution. "
        "In the meantime, please try [TROUBLESHOOTING STEP] and let us know if the issue persists. "
        "We aim to have this fully resolved within [TIMEFRAME]. "
        "Thank you for your patience — we appreciate it greatly. 😊\n\n"
        "How was your experience with our support today?"
    ),
    "Complaint": (
        "Thank you for taking the time to share your feedback, and we sincerely apologise for the experience you had. "
        "This is not the standard of service we strive for. We've escalated your case [CASE_ID] to our senior team "
        "for immediate review, and a dedicated agent will contact you within 4 hours. "
        "We take every concern seriously and are committed to making this right for you. 😊\n\n"
        "Your feedback helps us improve — how would you rate your support experience today?"
    ),
    "Other": (
        "Thank you for reaching out to us! We've reviewed your message and our team is addressing your concern. "
        "We aim to provide a resolution within 24 hours and will keep you updated throughout. "
        "We appreciate your patience and your trust in us. 😊\n\n"
        "How was your experience with our support team today?"
    ),
}

# ─────────────────────────────────────────────────────────────────────────────
# HELPER FUNCTIONS
# ─────────────────────────────────────────────────────────────────────────────

def detect_sentiment(text: str) -> str:
    """Keyword-based multilingual sentiment detector."""
    if not isinstance(text, str) or not text.strip():
        return "Neutral"
    t = text.lower()
    neg = sum(1 for kw in NEGATIVE_KWS if kw in t)
    pos = sum(1 for kw in POSITIVE_KWS if kw in t)
    if neg > pos:
        return "Negative"
    if pos > neg:
        return "Positive"
    return "Neutral"


def detect_issue_type(text: str) -> str:
    """Classify message into one of 11 issue types using keyword matching."""
    if not isinstance(text, str) or not text.strip():
        return "Other"
    t = text.lower()
    scores = {}
    for issue, kws in ISSUE_KEYWORDS.items():
        score = sum(1 for kw in kws if kw.lower() in t)
        if score > 0:
            scores[issue] = score
    if not scores:
        return "Other"
    return max(scores, key=scores.get)


def get_priority(issue_type: str) -> str:
    """Map issue type to priority level."""
    for priority, issues in PRIORITY_MAP.items():
        if issue_type in issues:
            return priority
    return "Low"


def matches_any(text: str, patterns: list) -> bool:
    """Check if text matches any regex pattern (case-insensitive)."""
    if not isinstance(text, str):
        return False
    t = text.lower()
    return any(re.search(p, t, re.IGNORECASE) for p in patterns)


def is_auto_reply(text: str) -> bool:
    return matches_any(text, AUTO_REPLY_PATTERNS)


def conversation_is_unresolved(seller_msgs: list) -> bool:
    """
    Returns True if conversation has stalling phrases without a following resolution phrase.
    Strategy: scan chronologically. If stall found but no later resolution → unresolved.
    """
    stall_found = False
    for msg in seller_msgs:
        if matches_any(msg, STALLING_PATTERNS):
            stall_found = True
        if matches_any(msg, RESOLUTION_PATTERNS):
            stall_found = False   # Resolution found after stall → mark resolved
    return stall_found


def compute_csat(sentiment: str, is_resolved: bool) -> float:
    """Proxy CSAT score 1–5 based on sentiment + resolution."""
    matrix = {
        ("Positive", True):  5.0,
        ("Positive", False): 3.5,
        ("Neutral",  True):  4.0,
        ("Neutral",  False): 3.0,
        ("Negative", True):  2.5,
        ("Negative", False): 1.0,
    }
    return matrix.get((sentiment, is_resolved), 3.0)


def generate_summary(buyer_msgs: list, issue_type: str) -> str:
    """Rule-based buyer chat summary."""
    if not buyer_msgs:
        return "No buyer messages."
    combined = " ".join([m for m in buyer_msgs if isinstance(m, str)])[:400]
    return f"[{issue_type}] Buyer enquiry: {combined[:200]}{'...' if len(combined) > 200 else ''}"


def fmt_mins(mins) -> str:
    """Format float minutes → human-readable."""
    if pd.isna(mins) or mins < 0:
        return "—"
    if mins < 60:
        return f"{int(mins)}m"
    h = int(mins // 60)
    m = int(mins % 60)
    return f"{h}h {m}m" if m else f"{h}h"


# ─────────────────────────────────────────────────────────────────────────────
# DATA LOADING
# ─────────────────────────────────────────────────────────────────────────────

@st.cache_data(show_spinner=False)
def load_data(file_bytes: bytes) -> pd.DataFrame:
    """Load Excel, combine both sheets, add PLATFORM column, parse dates."""
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    sheets_found = xl.sheet_names

    dfs = []
    platform_map = {}
    for s in sheets_found:
        name_lower = s.lower()
        if "lazada" in name_lower:
            platform_map[s] = "Lazada"
        elif "shopee" in name_lower:
            platform_map[s] = "Shopee"
        else:
            platform_map[s] = "Unknown"
        df = xl.parse(s, dtype=str)
        df["PLATFORM"] = platform_map[s]
        dfs.append(df)

    combined = pd.concat(dfs, ignore_index=True)

    # Parse MESSAGE_TIME
    combined["MESSAGE_TIME"] = pd.to_datetime(combined["MESSAGE_TIME"], errors="coerce")

    # Normalise columns
    for col in ["STORE_CODE", "SITE_NICK_NAME_ID", "COUNTRY_CODE",
                "CONVERSATION_ID", "BUYER_NAME", "MESSAGE_PARSED",
                "MESSAGE_TYPE", "SENDER"]:
        if col in combined.columns:
            combined[col] = combined[col].fillna("").astype(str).str.strip()

    # Boolean flags — vectorised to avoid pandas 3.x lambda issues
    for flag in ["IS_READ", "IS_ANSWERED"]:
        if flag in combined.columns:
            combined[flag] = (
                combined[flag].astype(str).str.strip().str.lower()
                .isin(["true", "1", "yes"])
            )

    return combined


# ─────────────────────────────────────────────────────────────────────────────
# ANALYSIS ENGINE
# ─────────────────────────────────────────────────────────────────────────────

def analyse(df: pd.DataFrame) -> pd.DataFrame:
    """
    Group by CONVERSATION_ID, compute per-conversation metrics.
    Returns one row per conversation.
    """
    rows = []

    # Sort messages chronologically within each conversation
    df_sorted = df.sort_values(["CONVERSATION_ID", "MESSAGE_TIME"])

    for conv_id, grp in df_sorted.groupby("CONVERSATION_ID", sort=False):
        buyer_msgs   = grp[grp["SENDER"].str.lower() == "buyer"]["MESSAGE_PARSED"].tolist()
        seller_msgs  = grp[grp["SENDER"].str.lower() == "seller"]["MESSAGE_PARSED"].tolist()
        all_msgs     = grp["MESSAGE_PARSED"].tolist()

        # Full text for analysis (combine buyer messages)
        full_buyer_text = " ".join([m for m in buyer_msgs if isinstance(m, str)])

        issue_type   = detect_issue_type(full_buyer_text)
        sentiment    = detect_sentiment(full_buyer_text)
        is_unresolved = conversation_is_unresolved(seller_msgs)
        is_resolved  = not is_unresolved
        priority     = get_priority(issue_type)
        csat         = compute_csat(sentiment, is_resolved)
        summary      = generate_summary(buyer_msgs, issue_type)
        reply        = SUGGESTED_REPLIES.get(issue_type, SUGGESTED_REPLIES["Other"])

        # CRT: average time (mins) between consecutive buyer→seller pairs
        crt_list = []
        grp_rows = grp.sort_values("MESSAGE_TIME").to_dict("records")
        last_buyer_time = None
        for r in grp_rows:
            if r["SENDER"].lower() == "buyer":
                last_buyer_time = r["MESSAGE_TIME"]
            elif r["SENDER"].lower() == "seller" and last_buyer_time is not None:
                delta = (r["MESSAGE_TIME"] - last_buyer_time).total_seconds() / 60
                if 0 <= delta <= 1440:   # ignore gaps > 24h (likely new day)
                    crt_list.append(delta)
                last_buyer_time = None

        avg_crt = np.mean(crt_list) if crt_list else np.nan

        # Last message time and date
        last_msg_time = grp["MESSAGE_TIME"].max()
        first_msg_time = grp["MESSAGE_TIME"].min()

        # Metadata from first row of the group
        meta = grp.iloc[0]
        rows.append({
            "CONVERSATION_ID":   conv_id,
            "PLATFORM":          meta.get("PLATFORM", ""),
            "STORE_CODE":        meta.get("STORE_CODE", ""),
            "SITE_NICK_NAME_ID": meta.get("SITE_NICK_NAME_ID", ""),
            "COUNTRY_CODE":      meta.get("COUNTRY_CODE", ""),
            "BUYER_NAME":        meta.get("BUYER_NAME", ""),
            "BUYER_ID":          meta.get("BUYER_ID", ""),
            "FIRST_MSG_TIME":    first_msg_time,
            "LAST_MSG_TIME":     last_msg_time,
            "MSG_COUNT":         len(grp),
            "BUYER_MSG_COUNT":   len(buyer_msgs),
            "SELLER_MSG_COUNT":  len(seller_msgs),
            "ISSUE_TYPE":        issue_type,
            "PRIORITY":          priority,
            "SENTIMENT":         sentiment,
            "IS_UNRESOLVED":     is_unresolved,
            "IS_RESOLVED":       is_resolved,
            "CSAT_PROXY":        round(csat, 1),
            "AVG_CRT_MINS":      round(avg_crt, 1) if not np.isnan(avg_crt) else None,
            "BUYER_SUMMARY":     summary,
            "SUGGESTED_REPLY":   reply,
            "IS_ANSWERED":       str(meta.get("IS_ANSWERED", "")).lower() == "true",
            "IS_READ":           str(meta.get("IS_READ", "")).lower() == "true",
        })

    return pd.DataFrame(rows)


# ─────────────────────────────────────────────────────────────────────────────
# EXCEL EXPORT
# ─────────────────────────────────────────────────────────────────────────────

def build_excel(conv_df: pd.DataFrame, today_str: str) -> bytes:
    """Build 4-sheet Excel output."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        wb = writer.book

        # ── Formats ──────────────────────────────────────────────────────────
        hdr_fmt  = wb.add_format({"bold": True, "bg_color": "#1B2A4A", "font_color": "#FFFFFF",
                                   "border": 1, "font_size": 10, "align": "center", "valign": "vcenter"})
        sub_fmt  = wb.add_format({"bold": True, "bg_color": "#00C4B4", "font_color": "#FFFFFF",
                                   "border": 1, "font_size": 10})
        num_fmt  = wb.add_format({"num_format": "#,##0", "border": 1})
        dec_fmt  = wb.add_format({"num_format": "0.0", "border": 1})
        cell_fmt = wb.add_format({"border": 1, "font_size": 9, "text_wrap": True, "valign": "top"})
        high_fmt = wb.add_format({"border": 1, "font_size": 9, "bg_color": "#FDECEA", "font_color": "#C0392B"})
        med_fmt  = wb.add_format({"border": 1, "font_size": 9, "bg_color": "#FEF9E7", "font_color": "#D68910"})
        low_fmt  = wb.add_format({"border": 1, "font_size": 9, "bg_color": "#EAF4FB", "font_color": "#2980B9"})

        def write_df(ws, df, start_row=0):
            for c_idx, col in enumerate(df.columns):
                ws.write(start_row, c_idx, col, hdr_fmt)
            for r_idx, row in enumerate(df.itertuples(index=False), start=start_row + 1):
                for c_idx, val in enumerate(row):
                    if val is None or (isinstance(val, float) and np.isnan(val)):
                        ws.write(r_idx, c_idx, "", cell_fmt)
                    elif isinstance(val, (int, float)):
                        ws.write_number(r_idx, c_idx, val, dec_fmt)
                    else:
                        ws.write(r_idx, c_idx, str(val), cell_fmt)

        # ── Sheet 1 : Summary Dashboard ──────────────────────────────────────
        ws1 = wb.add_worksheet("Summary Dashboard")
        writer.sheets["Summary Dashboard"] = ws1
        ws1.set_column(0, 0, 28)
        ws1.set_column(1, 1, 20)

        total      = len(conv_df)
        resolved   = conv_df["IS_RESOLVED"].sum()
        unresolved = conv_df["IS_UNRESOLVED"].sum()
        crr        = round(resolved / total * 100, 1) if total else 0
        avg_crt    = conv_df["AVG_CRT_MINS"].mean()
        avg_csat   = conv_df["CSAT_PROXY"].mean()

        today_df   = conv_df[conv_df["LAST_MSG_TIME"].dt.normalize() == pd.Timestamp(today_str)]
        hi_today   = len(today_df[today_df["PRIORITY"] == "High"])

        summary_data = [
            ["METRIC", "VALUE"],
            ["Total Conversations (7 days)", total],
            ["Today's Conversations", len(today_df)],
            ["Resolved Conversations", int(resolved)],
            ["Unresolved Conversations", int(unresolved)],
            ["Chat Resolution Rate (CRR)", f"{crr}%"],
            ["Avg Chat Response Time (CRT)", fmt_mins(avg_crt)],
            ["Avg CSAT Proxy Score (1–5)", round(avg_csat, 2) if not np.isnan(avg_csat) else "—"],
            ["Today's High Priority Chats", hi_today],
            ["Platforms", ", ".join(conv_df["PLATFORM"].unique().tolist())],
        ]

        ws1.write(0, 0, f"Chat Analyzer Summary — {today_str}", wb.add_format(
            {"bold": True, "font_size": 14, "font_color": "#1B2A4A"}))
        ws1.write(1, 0, "Generated by Graas.ai Chat Analyzer Dashboard", wb.add_format(
            {"italic": True, "font_color": "#7A8EA8"}))

        for i, (label, val) in enumerate(summary_data[1:], start=3):
            ws1.write(i, 0, label, sub_fmt)
            ws1.write(i, 1, val, cell_fmt)

        # Issue type breakdown
        ws1.write(14, 0, "ISSUE TYPE BREAKDOWN", sub_fmt)
        ws1.write(14, 1, "COUNT", hdr_fmt)
        for i, (issue, cnt) in enumerate(conv_df["ISSUE_TYPE"].value_counts().items(), start=15):
            ws1.write(i, 0, issue, cell_fmt)
            ws1.write(i, 1, int(cnt), num_fmt)

        # ── Sheet 2 : Today Priority Chats ───────────────────────────────────
        priority_cols = [
            "CONVERSATION_ID", "PLATFORM", "STORE_CODE", "BUYER_NAME",
            "ISSUE_TYPE", "PRIORITY", "SENTIMENT", "IS_UNRESOLVED",
            "CSAT_PROXY", "AVG_CRT_MINS", "BUYER_SUMMARY", "SUGGESTED_REPLY",
        ]
        today_pri = today_df.sort_values(
            "PRIORITY",
            key=lambda s: s.map({"High": 0, "Medium": 1, "Low": 2}).fillna(3)
        )[priority_cols]

        today_pri.to_excel(writer, sheet_name="Today Priority Chats", index=False)
        ws2 = writer.sheets["Today Priority Chats"]
        ws2.set_column(0, 0, 40)
        ws2.set_column(1, 5, 15)
        ws2.set_column(10, 11, 50)
        for c_idx, col in enumerate(today_pri.columns):
            ws2.write(0, c_idx, col, hdr_fmt)

        # ── Sheet 3 : Detailed Chat Analysis ─────────────────────────────────
        detail_cols = [
            "CONVERSATION_ID", "PLATFORM", "STORE_CODE", "SITE_NICK_NAME_ID",
            "COUNTRY_CODE", "BUYER_NAME", "FIRST_MSG_TIME", "LAST_MSG_TIME",
            "MSG_COUNT", "ISSUE_TYPE", "PRIORITY", "SENTIMENT",
            "IS_RESOLVED", "IS_UNRESOLVED", "CSAT_PROXY", "AVG_CRT_MINS",
            "BUYER_SUMMARY", "SUGGESTED_REPLY",
        ]
        detail = conv_df[detail_cols].copy()
        detail["FIRST_MSG_TIME"] = detail["FIRST_MSG_TIME"].dt.strftime("%Y-%m-%d %H:%M")
        detail["LAST_MSG_TIME"]  = detail["LAST_MSG_TIME"].dt.strftime("%Y-%m-%d %H:%M")
        detail.to_excel(writer, sheet_name="Detailed Chat Analysis", index=False)
        ws3 = writer.sheets["Detailed Chat Analysis"]
        ws3.set_column(0, 0, 40)
        ws3.set_column(7, 7, 18)
        ws3.set_column(16, 17, 60)
        for c_idx, col in enumerate(detail.columns):
            ws3.write(0, c_idx, col, hdr_fmt)

        # ── Sheet 4 : Unresolved Chats ────────────────────────────────────────
        unres = conv_df[conv_df["IS_UNRESOLVED"]][priority_cols].sort_values(
            "PRIORITY",
            key=lambda s: s.map({"High": 0, "Medium": 1, "Low": 2}).fillna(3)
        )
        unres.to_excel(writer, sheet_name="Unresolved Chats", index=False)
        ws4 = writer.sheets["Unresolved Chats"]
        ws4.set_column(0, 0, 40)
        ws4.set_column(10, 11, 50)
        for c_idx, col in enumerate(unres.columns):
            ws4.write(0, c_idx, col, hdr_fmt)

    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────────────────────────────────────
# UI COMPONENTS
# ─────────────────────────────────────────────────────────────────────────────

def render_header():
    st.markdown("""
    <div class="graas-header">
        <div class="graas-logo">📊</div>
        <div>
            <h1>Chat Analyzer Dashboard</h1>
            <p>Graas.ai · Shopee & Lazada · Sentiment · CSAT · Unresolved Detection · Suggested Replies</p>
        </div>
    </div>
    """, unsafe_allow_html=True)


def render_metrics(conv_df: pd.DataFrame, today_ts: pd.Timestamp):
    total      = len(conv_df)
    resolved   = int(conv_df["IS_RESOLVED"].sum())
    unresolved = int(conv_df["IS_UNRESOLVED"].sum())
    crr        = round(resolved / total * 100, 1) if total else 0
    avg_crt    = conv_df["AVG_CRT_MINS"].mean()
    avg_csat   = conv_df["CSAT_PROXY"].mean()

    # Compare at day granularity using Timestamps (pandas 3.0 safe)
    today_conv  = conv_df[conv_df["LAST_MSG_TIME"].dt.normalize() == today_ts]
    hi_today    = len(today_conv[today_conv["PRIORITY"] == "High"])

    neg_pct = round(len(conv_df[conv_df["SENTIMENT"] == "Negative"]) / total * 100, 1) if total else 0

    cols = st.columns(8)
    metrics = [
        (cols[0], "🗣️ Total Convs", f"{total:,}",   "7-day window", ""),
        (cols[1], "📅 Today",       f"{len(today_conv):,}", "conversations", "navy"),
        (cols[2], "✅ Resolved",    f"{resolved:,}", f"CRR {crr}%", "green"),
        (cols[3], "🔴 Unresolved",  f"{unresolved:,}", "need action", "red"),
        (cols[4], "⚡ CRT",         fmt_mins(avg_crt), "avg response time", "orange"),
        (cols[5], "⭐ CSAT",        f"{avg_csat:.1f}/5" if not np.isnan(avg_csat) else "—", "proxy score", ""),
        (cols[6], "😠 Negative",    f"{neg_pct}%",  "sentiment", "red"),
        (cols[7], "🔥 High Pri",    f"{hi_today}",  "today's urgent", "orange"),
    ]
    for col, label, val, sub, cls in metrics:
        with col:
            st.markdown(f"""
            <div class="metric-card {cls}">
                <div class="metric-label">{label}</div>
                <div class="metric-val">{val}</div>
                <div class="metric-sub">{sub}</div>
            </div>
            """, unsafe_allow_html=True)


def priority_badge(p: str) -> str:
    cls = {"High": "badge-high", "Medium": "badge-medium", "Low": "badge-low"}.get(p, "badge-low")
    return f'<span class="{cls}">{p}</span>'


def sentiment_span(s: str) -> str:
    cls = {"Positive": "sent-pos", "Neutral": "sent-neu", "Negative": "sent-neg"}.get(s, "sent-neu")
    icon = {"Positive": "😊", "Neutral": "😐", "Negative": "😠"}.get(s, "")
    return f'<span class="{cls}">{icon} {s}</span>'


# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR FILTERS
# ─────────────────────────────────────────────────────────────────────────────

def apply_filters(conv_df: pd.DataFrame) -> pd.DataFrame:
    st.sidebar.markdown("## 🔍 Filters")
    st.sidebar.markdown("---")

    # Platform
    platforms = ["All"] + sorted(conv_df["PLATFORM"].unique().tolist())
    sel_platform = st.sidebar.selectbox("Platform", platforms)
    if sel_platform != "All":
        conv_df = conv_df[conv_df["PLATFORM"] == sel_platform]

    # Date range — use Timestamp min/max to avoid pandas 3.0 dt.date issues
    _ts_min = conv_df["LAST_MSG_TIME"].dropna().min()
    _ts_max = conv_df["LAST_MSG_TIME"].dropna().max()
    min_date = _ts_min.date() if pd.notna(_ts_min) else datetime.today().date()
    max_date = _ts_max.date() if pd.notna(_ts_max) else datetime.today().date()
    date_range = st.sidebar.date_input(
        "Date Range",
        value=(min_date, max_date),
        min_value=min_date,
        max_value=max_date,
    )
    if isinstance(date_range, (list, tuple)) and len(date_range) == 2:
        # Compare against Timestamps for pandas 3.0 compatibility
        start_ts = pd.Timestamp(date_range[0])
        end_ts   = pd.Timestamp(date_range[1]) + pd.Timedelta(hours=23, minutes=59, seconds=59)
        conv_df = conv_df[
            (conv_df["LAST_MSG_TIME"] >= start_ts) &
            (conv_df["LAST_MSG_TIME"] <= end_ts)
        ]

    # Priority
    prio_opts = ["All", "High", "Medium", "Low"]
    sel_prio = st.sidebar.selectbox("Priority", prio_opts)
    if sel_prio != "All":
        conv_df = conv_df[conv_df["PRIORITY"] == sel_prio]

    # Sentiment
    sent_opts = ["All", "Positive", "Neutral", "Negative"]
    sel_sent = st.sidebar.selectbox("Sentiment", sent_opts)
    if sel_sent != "All":
        conv_df = conv_df[conv_df["SENTIMENT"] == sel_sent]

    # Resolved
    res_opts = ["All", "Resolved", "Unresolved"]
    sel_res = st.sidebar.selectbox("Resolution Status", res_opts)
    if sel_res == "Resolved":
        conv_df = conv_df[conv_df["IS_RESOLVED"]]
    elif sel_res == "Unresolved":
        conv_df = conv_df[conv_df["IS_UNRESOLVED"]]

    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🔎 Search")

    # STORE_CODE
    stores = sorted(conv_df["STORE_CODE"].dropna().unique().tolist())
    sel_stores = st.sidebar.multiselect("Store Code", stores)
    if sel_stores:
        conv_df = conv_df[conv_df["STORE_CODE"].isin(sel_stores)]

    # SITE_NICK_NAME_ID
    sites = sorted(conv_df["SITE_NICK_NAME_ID"].dropna().unique().tolist())
    sel_sites = st.sidebar.multiselect("Site Nickname", sites)
    if sel_sites:
        conv_df = conv_df[conv_df["SITE_NICK_NAME_ID"].isin(sel_sites)]

    # COUNTRY_CODE
    countries = sorted(conv_df["COUNTRY_CODE"].dropna().unique().tolist())
    sel_countries = st.sidebar.multiselect("Country Code", countries)
    if sel_countries:
        conv_df = conv_df[conv_df["COUNTRY_CODE"].isin(sel_countries)]

    # BUYER_NAME free text search
    buyer_search = st.sidebar.text_input("Buyer Name (search)")
    if buyer_search:
        conv_df = conv_df[conv_df["BUYER_NAME"].str.contains(buyer_search, case=False, na=False)]

    # CONVERSATION_ID free text
    conv_search = st.sidebar.text_input("Conversation ID (search)")
    if conv_search:
        conv_df = conv_df[conv_df["CONVERSATION_ID"].str.contains(conv_search, case=False, na=False)]

    # Issue type
    issue_opts = ["All"] + sorted(conv_df["ISSUE_TYPE"].unique().tolist())
    sel_issue = st.sidebar.selectbox("Issue Type", issue_opts)
    if sel_issue != "All":
        conv_df = conv_df[conv_df["ISSUE_TYPE"] == sel_issue]

    st.sidebar.markdown("---")
    st.sidebar.markdown(f"**{len(conv_df):,}** conversations match filters")
    return conv_df


# ─────────────────────────────────────────────────────────────────────────────
# MAIN APP
# ─────────────────────────────────────────────────────────────────────────────

def main():
    render_header()

    # ── File Upload ───────────────────────────────────────────────────────────
    st.markdown('<div class="section-title">📂 Upload Chat Data</div>', unsafe_allow_html=True)
    uploaded = st.file_uploader(
        "Upload Excel file with sheets: lazada_chat_enquiries & shopee_chat_enquiries",
        type=["xlsx"],
        help="Single Excel file containing both Lazada and Shopee chat sheets.",
    )

    if not uploaded:
        st.info("👆 Upload your chat enquiries Excel file to get started.")
        st.markdown("""
        **Expected Excel format:**
        - Sheet 1: `lazada_chat_enquiries`
        - Sheet 2: `shopee_chat_enquiries`
        - Columns: `STORE_CODE`, `SITE_NICK_NAME_ID`, `COUNTRY_CODE`, `CONVERSATION_ID`,
          `IS_READ`, `IS_ANSWERED`, `MESSAGE_TIME`, `BUYER_NAME`, `MESSAGE_PARSED`,
          `MESSAGE_TYPE`, `MESSAGE_ID`, `SENDER`, `BUYER_ID`
        """)
        return

    # ── Load & Filter ─────────────────────────────────────────────────────────
    with st.spinner("⏳ Loading chat data…"):
        raw_df = load_data(uploaded.read())

    # pandas 3.0 fix: .dt.date.max() fails with NaT/float mixing — use Timestamp.max() then .date()
    _max_ts = raw_df["MESSAGE_TIME"].dropna().max()
    today_date = _max_ts.date() if pd.notna(_max_ts) else datetime.today().date()
    today_str = today_date.strftime("%Y-%m-%d")
    today_ts  = pd.Timestamp(today_date)          # used throughout for comparisons

    # Filter to last 7 days — compare Timestamps, not date objects
    cutoff = today_ts - timedelta(days=6)
    df_7day = raw_df[raw_df["MESSAGE_TIME"] >= cutoff].copy()

    st.success(
        f"✅ Loaded **{len(raw_df):,}** messages across "
        f"**{raw_df['PLATFORM'].nunique()}** platforms. "
        f"Analysing last 7 days: **{cutoff.date()}** → **{today_date}**"
    )

    # ── Analyse ───────────────────────────────────────────────────────────────
    with st.spinner("🔍 Analysing conversations (sentiment · issue type · CRT · CSAT)…"):
        conv_df = analyse(df_7day)

    # ── Sidebar Filters ───────────────────────────────────────────────────────
    conv_filtered = apply_filters(conv_df)

    if conv_filtered.empty:
        st.warning("No conversations match the current filters.")
        return

    # ── Metrics Row ───────────────────────────────────────────────────────────
    st.markdown('<div class="section-title">📈 Key Metrics</div>', unsafe_allow_html=True)
    render_metrics(conv_filtered, today_ts)

    # ── Charts Row ────────────────────────────────────────────────────────────
    st.markdown('<div class="section-title">📊 Analytics</div>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)

    with c1:
        issue_counts = conv_filtered["ISSUE_TYPE"].value_counts().reset_index()
        issue_counts.columns = ["Issue Type", "Count"]
        st.markdown("**Issue Type Distribution**")
        st.bar_chart(issue_counts.set_index("Issue Type")["Count"], color="#00C4B4")

    with c2:
        sent_counts = conv_filtered["SENTIMENT"].value_counts().reset_index()
        sent_counts.columns = ["Sentiment", "Count"]
        st.markdown("**Sentiment Breakdown**")
        color_map = {"Positive": "#27AE60", "Neutral": "#7F8C8D", "Negative": "#E74C3C"}
        st.bar_chart(sent_counts.set_index("Sentiment")["Count"])

    with c3:
        # Use dt.normalize() (returns Timestamp at midnight) — pandas 3.0 safe
        daily = (
            conv_filtered
            .assign(DATE=conv_filtered["LAST_MSG_TIME"].dt.normalize())
            .groupby("DATE")
            .size()
            .reset_index(name="Conversations")
        )
        st.markdown("**Daily Conversation Volume**")
        st.line_chart(daily.set_index("DATE")["Conversations"], color="#FF6B35")

    # ── Tabs ─────────────────────────────────────────────────────────────────
    tab1, tab2, tab3, tab4 = st.tabs([
        "🔥 Today's Priority Chats",
        "📋 All Conversations",
        "🔴 Unresolved Chats",
        "💬 Suggested Replies",
    ])

    display_cols = [
        "CONVERSATION_ID", "PLATFORM", "STORE_CODE", "BUYER_NAME",
        "ISSUE_TYPE", "PRIORITY", "SENTIMENT", "IS_UNRESOLVED",
        "CSAT_PROXY", "AVG_CRT_MINS", "BUYER_SUMMARY",
    ]

    with tab1:
        # Use normalize() for date comparison — pandas 3.0 safe
        today_df = conv_filtered[conv_filtered["LAST_MSG_TIME"].dt.normalize() == today_ts]
        today_sorted = today_df.sort_values(
            "PRIORITY",
            key=lambda s: s.map({"High": 0, "Medium": 1, "Low": 2}).fillna(3)
        )
        st.markdown(f"**{len(today_sorted)} conversations today** — sorted by priority")
        if today_sorted.empty:
            st.info("No conversations found for today.")
        else:
            st.dataframe(
                today_sorted[display_cols].reset_index(drop=True),
                use_container_width=True,
                height=450,
                column_config={
                    "CSAT_PROXY":   st.column_config.NumberColumn("CSAT (1-5)", format="%.1f"),
                    "AVG_CRT_MINS": st.column_config.NumberColumn("CRT (mins)", format="%.0f"),
                    "IS_UNRESOLVED": st.column_config.CheckboxColumn("Unresolved?"),
                    "BUYER_SUMMARY": st.column_config.TextColumn("Summary", width="large"),
                },
            )

    with tab2:
        all_sorted = conv_filtered.sort_values("LAST_MSG_TIME", ascending=False)
        st.markdown(f"**{len(all_sorted)} conversations** in filtered view")
        st.dataframe(
            all_sorted[display_cols].reset_index(drop=True),
            use_container_width=True,
            height=500,
            column_config={
                "CSAT_PROXY":   st.column_config.NumberColumn("CSAT (1-5)", format="%.1f"),
                "AVG_CRT_MINS": st.column_config.NumberColumn("CRT (mins)", format="%.0f"),
                "IS_UNRESOLVED": st.column_config.CheckboxColumn("Unresolved?"),
                "BUYER_SUMMARY": st.column_config.TextColumn("Summary", width="large"),
            },
        )

    with tab3:
        unres_df = conv_filtered[conv_filtered["IS_UNRESOLVED"]].sort_values(
            "PRIORITY",
            key=lambda s: s.map({"High": 0, "Medium": 1, "Low": 2}).fillna(3)
        )
        st.markdown(
            f"**{len(unres_df)} unresolved conversations** — contain stalling phrases without resolution"
        )
        if unres_df.empty:
            st.success("🎉 No unresolved conversations found!")
        else:
            st.dataframe(
                unres_df[display_cols].reset_index(drop=True),
                use_container_width=True,
                height=450,
                column_config={
                    "CSAT_PROXY":   st.column_config.NumberColumn("CSAT (1-5)", format="%.1f"),
                    "AVG_CRT_MINS": st.column_config.NumberColumn("CRT (mins)", format="%.0f"),
                    "IS_UNRESOLVED": st.column_config.CheckboxColumn("Unresolved?"),
                    "BUYER_SUMMARY": st.column_config.TextColumn("Summary", width="large"),
                },
            )

    with tab4:
        st.markdown("### 💬 Suggested Reply Templates by Issue Type")
        st.caption(
            "Empathetic, resolution-oriented replies — replace [PLACEHOLDERS] before sending."
        )
        for issue_type, reply_text in SUGGESTED_REPLIES.items():
            if issue_type == "Other":
                continue
            priority = get_priority(issue_type)
            badge_color = {"High": "🔴", "Medium": "🟡", "Low": "🔵"}.get(priority, "⚪")
            with st.expander(f"{badge_color} {issue_type}  ({priority} Priority)"):
                st.markdown(f"""
                <div class="reply-label">Suggested Reply</div>
                <div class="reply-box">{reply_text}</div>
                """, unsafe_allow_html=True)

        # Per-conversation reply lookup
        st.markdown("---")
        st.markdown("### 🔍 Look Up Reply for a Specific Conversation")
        conv_ids = conv_filtered["CONVERSATION_ID"].tolist()
        if conv_ids:
            sel_conv = st.selectbox("Select Conversation ID", conv_ids[:500])
            row = conv_filtered[conv_filtered["CONVERSATION_ID"] == sel_conv].iloc[0]
            st.markdown(f"""
            **Issue Type:** {row['ISSUE_TYPE']}  |
            **Priority:** {row['PRIORITY']}  |
            **Sentiment:** {row['SENTIMENT']}  |
            **CSAT Proxy:** {row['CSAT_PROXY']}
            """)
            st.markdown(f"""
            <div class="reply-label">Buyer Summary</div>
            <div class="reply-box">{row['BUYER_SUMMARY']}</div>
            """, unsafe_allow_html=True)
            st.markdown(f"""
            <div class="reply-label">Suggested Reply</div>
            <div class="reply-box">{row['SUGGESTED_REPLY']}</div>
            """, unsafe_allow_html=True)

    # ── Issue Breakdown Table ─────────────────────────────────────────────────
    st.markdown('<div class="section-title">📂 Issue Type Breakdown</div>', unsafe_allow_html=True)
    ib = (
        conv_filtered
        .groupby(["ISSUE_TYPE", "PRIORITY"])
        .agg(
            Count=("CONVERSATION_ID", "count"),
            Unresolved=("IS_UNRESOLVED", "sum"),
            Avg_CSAT=("CSAT_PROXY", "mean"),
            Avg_CRT_mins=("AVG_CRT_MINS", "mean"),
        )
        .reset_index()
        .sort_values("Count", ascending=False)
    )
    ib["Avg_CSAT"] = ib["Avg_CSAT"].round(1)
    ib["Avg_CRT_mins"] = ib["Avg_CRT_mins"].round(0).fillna(0).astype(int)
    ib["Unresolved"] = ib["Unresolved"].astype(int)
    st.dataframe(ib, use_container_width=True, height=300)

    # ── Store Performance ─────────────────────────────────────────────────────
    st.markdown('<div class="section-title">🏪 Store Performance</div>', unsafe_allow_html=True)
    sp = (
        conv_filtered
        .groupby(["STORE_CODE", "PLATFORM"])
        .agg(
            Conversations=("CONVERSATION_ID", "count"),
            Unresolved=("IS_UNRESOLVED", "sum"),
            Avg_CSAT=("CSAT_PROXY", "mean"),
            Avg_CRT_mins=("AVG_CRT_MINS", "mean"),
            Negative_Sent=("SENTIMENT", lambda x: (x == "Negative").sum()),
        )
        .reset_index()
        .sort_values("Conversations", ascending=False)
    )
    sp["Avg_CSAT"] = sp["Avg_CSAT"].round(1)
    sp["Avg_CRT_mins"] = sp["Avg_CRT_mins"].round(0).fillna(0).astype(int)
    sp["Unresolved"] = sp["Unresolved"].astype(int)
    sp["CRR%"] = ((sp["Conversations"] - sp["Unresolved"]) / sp["Conversations"] * 100).round(1)
    st.dataframe(sp, use_container_width=True, height=350)

    # ── Excel Download ────────────────────────────────────────────────────────
    st.markdown('<div class="section-title">⬇️ Download Report</div>', unsafe_allow_html=True)
    with st.spinner("Building Excel report…"):
        excel_bytes = build_excel(conv_filtered, today_str)

    st.download_button(
        label="📥 Download Full Excel Report",
        data=excel_bytes,
        file_name=f"Chat_Analysis_{today_str}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    st.caption(
        "Excel contains: **Summary Dashboard** · **Today Priority Chats** · "
        "**Detailed Chat Analysis** · **Unresolved Chats**"
    )


if __name__ == "__main__":
    main()
