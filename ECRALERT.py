import os
import io
import base64
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
import pandas as pd
import streamlit as st
from datetime import date
from PIL import Image, Image as PILImage
import gspread

# 📌 แก้ไขจุดนี้: อิมพอร์ต service_account ให้ถูกต้อง
from google.oauth2 import service_account

# หลังจากอิมพอร์ตแล้ว สามารถเรียกใช้แบบนี้ได้เลยโดยไม่เกิด Error
# crecs = service_account.Credentials.from_service_account_file(...)

# =============================================================
# 🖼️ BASE64 IMAGE CONVERSION & RESIZE
# =============================================================
def convert_image_to_base64(uploaded_file, max_size=(600, 600), quality=60):
    """
    ย่อ/บีบอัดรูปภาพให้เหมาะกับ Google Sheets แล้วเก็บเป็น Data URL
    โดยพยายามให้ขนาดข้อความต่ำกว่า ~40,000 ตัวอักษร เพื่อไม่ชน cell limit
    """
    if uploaded_file is None:
        return ""

    try:
        original = Image.open(uploaded_file)
        if original.mode in ("RGBA", "P"):
            original = original.convert("RGB")

        # ลองหลายระดับจากคุณภาพสูง -> ต่ำ จนกว่าจะเล็กพอ
        attempts = [
            ((600, 600), 60),
            ((550, 550), 50),
            ((500, 500), 45),
            ((450, 450), 40),
            ((400, 400), 35),
            ((350, 350), 30),
        ]

        for size, q in attempts:
            img = original.copy()
            img.thumbnail(size, Image.Resampling.LANCZOS)

            buffer = io.BytesIO()
            img.save(buffer, format="JPEG", quality=q, optimize=True)
            encoded = base64.b64encode(buffer.getvalue()).decode()
            data_url = f"data:image/jpeg;base64,{encoded}"

            if len(data_url) <= 40000:
                print("🖼️ IMAGE COMPRESSED:", len(data_url), "chars", "size=", size, "quality=", q)
                return data_url

        # ถ้ายังใหญ่ ให้ใช้ตัวเลือกสุดท้าย
        print("⚠️ IMAGE still large after compression:", len(data_url), "chars")
        return data_url

    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการแปลงรูปภาพ: {e}")
        print(f"❌ IMAGE CONVERSION ERROR: {type(e).__name__}: {e}")
        return ""

def base64_to_image(base64_str):
    """แปลง Base64 String กลับเป็น PIL Image object เพื่อนำไปใส่ใน Excel"""
    try:
        if "," in base64_str:
            base64_str = base64_str.split(",")[1]
        image_data = base64.b64decode(base64_str)
        return Image.open(io.BytesIO(image_data))
    except Exception as e:
        return None

# =============================================================
# ตั้งค่าหน้าเว็บ Streamlit
# =============================================================
st.set_page_config(
    layout="wide",
    page_title="KFT Change Control System",
    page_icon="🔐"
)

TEMPLATE_FILE = "template_form.xlsx"
UPLOAD_DIR = "uploads"

os.makedirs(UPLOAD_DIR, exist_ok=True)

# =============================================================
# CONNECT GOOGLE SHEETS API
# =============================================================
@st.cache_resource
def get_gspread_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    
    json_path = os.path.join(os.path.dirname(__file__), "service_account.json")
    if os.path.exists(json_path):
        credentials = service_account.Credentials.from_service_account_file(
            json_path,
            scopes=scopes
        )
    else:
        creds_dict = dict(st.secrets["gcp_service_account"])
        if "private_key" in creds_dict:
            creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        credentials = service_account.Credentials.from_service_account_info(
            creds_dict,
            scopes=scopes
        )
        
    return gspread.authorize(credentials)

def get_worksheet():
    gc = get_gspread_client()
    spreadsheet_name = st.secrets.get("sheets", {}).get("spreadsheet_name", "change_control_db")
    sh = gc.open(spreadsheet_name)
    return sh.sheet1

# =============================================================
# CONFIGURATION: SMTP EMAIL SETTINGS & DEPARTMENT EMAILS
# =============================================================
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
SENDER_EMAIL = "pdd1development@gmail.com"
SENDER_PASSWORD = st.secrets.get("email", {}).get("sender_password", "")
APP_URL = "https://related-alert-erh2rywrtchautlthjlrwb.streamlit.app/"

DEPT_EMAILS = {
    "PDD": ["pdd_1@kftc.co.th", "saksiam@kftc.co.th", "manoc@kftc.co.th"],
    "QC": ["uchai@kftc.co.th", "sirirat@kftc.co.th", "pdd_1@kftc.co.th"],
    "PCD": ["pc-3@kftc.co.th", "pdd_1@kftc.co.th"],
    "PRD": ["suriya@kftc.co.th", "setthanan@kftc.co.th", "pd1center@kftc.co.th", "pdd_1@kftc.co.th"],
    "PRO": ["suriya@kftc.co.th", "setthanan@kftc.co.th", "pd1center@kftc.co.th", "pdd_1@kftc.co.th"],
    "PDD_MGR": ["manoch@kftc.co.th"],
    "QCD_MGR": ["maitree@kftc.co.th"],
    "PRD_MGR": ["suriya@kftc.co.th"],
    "PCD_MGR": ["umaporn@kftc.co.th"],
    "GM": ["mayuree@kftc.co.th"],
    "ALL": [
        "pdd_1@kftc.co.th", "saksiam@kftc.co.th", "manoc@kftc.co.th",
        "uchai@kftc.co.th", "sirirat@kftc.co.th", "pc-3@kftc.co.th",
        "suriya@kftc.co.th", "setthanan@kftc.co.th", "pd1center@kftc.co.th"
    ]
}

def send_email_notification(to_email, subject, body_content):
    recipient_list = to_email if isinstance(to_email, list) else [to_email]
    msg = MIMEMultipart()
    msg['From'] = SENDER_EMAIL
    msg['To'] = ", ".join(recipient_list)
    msg['Subject'] = subject
    msg.attach(MIMEText(body_content, 'plain', 'utf-8'))

    try:
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.sendmail(SENDER_EMAIL, recipient_list, msg.as_string())
        server.quit()
        return True
    except Exception as e:
        st.error(f"❌ [DEBUG ERROR] ส่ง Email ไม่สำเร็จ สาเหตุเกิดจาก: {type(e).__name__} - {e}")
        return False

def send_all_completed_alert_email(doc_no, customer, part_name):
    to_email = DEPT_EMAILS.get("PDD_MGR", SENDER_EMAIL)
    subject = f"✅ [Wait Manager Approval] ใบงาน {doc_no} ปิดข้อ YES ครบถ้วนแล้ว (รอ PDD MGR อนุมัติ)"
    body = (
        f"เรียน ผู้จัดการ PDD (PDD MGR),\n\n"
        f"ใบงาน Change Control เลขที่ {doc_no} (Customer: {customer}, Part: {part_name}) "
        f"ได้รับการปิดข้อรายการที่ต้องแก้ไข (YES) พร้อมลงวันที่ปิดงานจริง (Actual Close) ครบถ้วนตาม Plan แล้ว\n\n"
        f"ระบบได้เปิดให้เข้าสู่ขั้นตอนอนุมัติแล้ว รบกวนผู้จัดการเข้าสู่ระบบเพื่อพิจารณาลงนามอนุมัติเอกสารเป็นลำดับแรกครับ\n\n"
        f"🔗 เข้าสู่ระบบได้ที่: {APP_URL}\n\n"
        f"ขอแสดงความนับถือ,\nระบบ KFT Change Control Automated System"
    )
    return send_email_notification(to_email, subject, body)

def send_approval_next_step_email(doc_no, customer, part_name, approver_title, next_approver_key, next_approver_title):
    to_email = DEPT_EMAILS.get(next_approver_key, SENDER_EMAIL)
    subject = f"🖊️ [Approval Step] ใบงาน {doc_no} รอการอนุมัติจาก {next_approver_title}"
    body = (
        f"เรียน {next_approver_title},\n\n"
        f"{approver_title} ได้ทำการลงนามอนุมัติเอกสารเลขที่ {doc_no} เรียบร้อยแล้ว\n"
        f"CUSTOMER: {customer}\nPART NAME: {part_name}\n\n"
        f"ลำดับถัดไป: รบกวนท่านเข้าสู่ระบบเพื่อพิจารณาอนุมัติเอกสารในระบบครับ\n\n"
        f"🔗 เข้าสู่ระบบได้ที่: {APP_URL}\n\n"
        f"ขอแสดงความนับถือ,\nระบบ KFT Change Control Automated System"
    )
    return send_email_notification(to_email, subject, body)

def send_final_approved_email(doc_no, customer, part_name, gm_name):
    to_email = DEPT_EMAILS.get("ALL", SENDER_EMAIL)
    subject = f"🎉 [FINAL APPROVED] เอกสาร Change Control {doc_no} ผ่านการอนุมัติเสร็จสมบูรณ์"
    body = (
        f"เรียน ทีมงานที่เกี่ยวข้องทุกท่าน,\n\n"
        f"เอกสารควบคุมการเปลี่ยนแปลง (Change Control) เลขที่ {doc_no}\n"
        f"CUSTOMER: {customer}\nPART NAME: {part_name}\n"
        f"ผู้อนุมัติขั้นสุดท้าย (GM): {gm_name}\n\n"
        f"บัดนี้ เอกสารดังกล่าวได้ผ่านการอนุมัติครบถ้วนตามลำดับขั้นตอนเรียบร้อยแล้วครับ\n\n"
        f"🔗 ตรวจสอบเอกสารได้ที่: {APP_URL}\n\n"
        f"ขอแสดงความนับถือ,\nระบบ KFT Change Control Automated System"
    )
    return send_email_notification(to_email, subject, body)

# =============================================================
# 🎨 CSS
# =============================================================
st.markdown("""
<style>
    .stApp { background: linear-gradient(135deg, #e3f2fd 0%, #ffffff 50%, #f0f8ff 100%); }
    h1, h2, h3 { color: #1565c0 !important; font-weight: 700 !important; }
    .login-container {
        background: rgba(255, 255, 255, 0.95); border-radius: 20px; padding: 40px;
        box-shadow: 0 8px 32px rgba(33, 150, 243, 0.15); border: 1px solid #bbdefb;
        max-width: 450px; margin: 80px auto; text-align: center;
    }
    .login-title { color: #1565c0; font-size: 28px; font-weight: 700; margin-bottom: 8px; }
    .login-subtitle { color: #64b5f6; font-size: 14px; margin-bottom: 30px; }
    .stButton > button {
        background: linear-gradient(135deg, #42a5f5 0%, #1976d2 100%) !important;
        color: white !important; border: none !important; border-radius: 12px !important;
        padding: 12px 24px !important; font-weight: 600 !important;
        box-shadow: 0 4px 15px rgba(25, 118, 210, 0.3) !important;
    }
    .status-card {
        padding: 15px; border-radius: 10px; background-color: #ffffff;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05); border-left: 5px solid #1565c0;
        margin-bottom: 10px;
    }
</style>
""", unsafe_allow_html=True)

# =============================================================
# 👤 USERS
# =============================================================
sec_passwords = st.secrets.get("passwords", {})

USERS = {
    "Thanawat": {"password": sec_passwords.get("Thanawat", ""), "dept": "PDD (Product Design)", "name": "ENGINEER PDD"},
    "qc_user": {"password": sec_passwords.get("qc_user", ""), "dept": "QC (Quality Control)", "name": "ENGINEER QC"},
    "pcd_user": {"password": sec_passwords.get("pcd_user", ""), "dept": "PCD (Production Control)", "name": "ENGINEER PCD"},
    "prd_user": {"password": sec_passwords.get("prd_user", ""), "dept": "PRD (Production / PD)", "name": "ENGINEER Production"},
    "Manoch": {"password": sec_passwords.get("Manoch", ""), "dept": "MGR - PDD (ผู้จัดการ PDD)", "name": "ผู้จัดการ PDD"},
    "mgr_qcd": {"password": sec_passwords.get("mgr_qcd", ""), "dept": "MGR - QCD (ผู้จัดการ QC)", "name": "ผู้จัดการ QC"},
    "mgr_pcd": {"password": sec_passwords.get("mgr_pcd", ""), "dept": "MGR - PCD (ผู้จัดการ PCD)", "name": "ผู้จัดการ PCD"},
    "mgr_prd": {"password": sec_passwords.get("mgr_prd", ""), "dept": "MGR - PD (ผู้จัดการ Production)", "name": "ผู้จัดการ PRD"},
    "gm_user": {"password": sec_passwords.get("gm_user", ""), "dept": "AGM / GM (ผู้บริหารอนุมัติขั้นสุดท้าย)", "name": "ผู้บริหาร GM"},
    "print_user": {"password": sec_passwords.get("print_user", ""), "dept": "Print Form", "name": "เจ้าหน้าที่พิมพ์เอกสาร"},
}

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.current_user = None
    st.session_state.current_dept = None
    st.session_state.user_name = None

def login(username, password):
    user = USERS.get(username)
    if user and user["password"] != "" and user["password"] == password:
        st.session_state.logged_in = True
        st.session_state.current_user = username
        st.session_state.current_dept = user["dept"]
        st.session_state.user_name = user["name"]
        return True
    return False

def logout():
    st.session_state.logged_in = False
    st.session_state.current_user = None
    st.session_state.current_dept = None
    st.session_state.user_name = None
    st.rerun()

# =============================================================
# 🔐 LOGIN UI
# =============================================================
if not st.session_state.logged_in:
    st.markdown("<style>[data-testid='stSidebar'] {display: none;}</style>", unsafe_allow_html=True)
    _, col2, _ = st.columns([1, 2, 1])
    with col2:
        st.markdown("""
            <div class="login-container">
                <div style="font-size: 64px;">🔐</div>
                <div class="login-title">KFT Change Control</div>
                <div class="login-subtitle">ระบบควบคุมเอกสารการเปลี่ยนแปลง</div>
            </div>
        """, unsafe_allow_html=True)
        username = st.text_input("👤 Username", key="login_user")
        password = st.text_input("🔒 Password", type="password", key="login_pass")
        if st.button("🔓 เข้าสู่ระบบ", use_container_width=True, type="primary"):
            if login(username, password):
                st.success("✅ เข้าสู่ระบบสำเร็จ!")
                st.rerun()
            else:
                st.error("❌ Username หรือ Password ไม่ถูกต้อง")
    st.stop()

# =============================================================
# DATABASE OPERATIONS (GOOGLE SHEETS)
# =============================================================
ITEM_DEPT_MAPPING = {
    1: ("PDD", "MASTER DRAWING."), 2: ("PDD", "MATERIAL PART NO. LIST. , ACC DWG."),
    3: ("PDD", "PROCESS FLOW CHART."), 4: ("PDD", "OPERATION MANUAL."),
    5: ("PDD", "TEST RESULT."), 6: ("PDD", "FMEA"), 7: ("PDD", "TOOLING No"),
    8: ("QC", "CONTROL PLAN."), 9: ("QC", "INCOMING SHEET."),
    10: ("QC", "FINAL INSPECTION SHEET."), 11: ("QC", "W/I Out Going / TRAINING QC."),
    12: ("QC", "INSPECTION STD. + DATA CHECK."), 13: ("QC", "MSA"),
    14: ("QC", "PSW UP-DATE., PPAP APPROVAL."), 15: ("QC", "CHECKING FIXTURE."),
    16: ("PCD", "MATERIAL REQUIREMENT."), 17: ("PCD", "PACKING STANDARD."),
    18: ("PRD", "WORKING INSTRUCTION."), 19: ("PRD", "TRAINING PRODUCTION.")
}

DEPT_ITEM_RANGES = {
    "PDD": list(range(1, 8)),
    "QC": list(range(8, 16)),
    "PCD": list(range(16, 18)),
    "PRD": list(range(18, 20)),
}

DEPT_LABELS = {
    "PDD": "PDD — ข้อ 1-7",
    "QC": "QC — ข้อ 8-15",
    "PCD": "PCD — ข้อ 16-17",
    "PRD": "PRD — ข้อ 18-19",
}

def get_user_dept_code(dept_text=None):
    text = str(dept_text or st.session_state.get("current_dept", "")).upper()
    if "MGR - PDD" in text or text.startswith("PDD"):
        return "PDD"
    if "MGR - QCD" in text or text.startswith("QC") or "QUALITY" in text:
        return "QC"
    if "MGR - PCD" in text or text.startswith("PCD"):
        return "PCD"
    if "MGR - PD" in text or text.startswith("PRD") or "PRODUCTION / PD" in text or text.startswith("PRO"):
        return "PRD"
    return ""

def get_dept_completion(doc_data, dept_code):
    item_numbers = DEPT_ITEM_RANGES.get(dept_code, [])
    missing = []
    for num in item_numbers:
        rev = get_doc_value(doc_data, num, "REVISE").upper().strip()
        resp = get_doc_value(doc_data, num, "RESP").strip()
        plan = get_doc_value(doc_data, num, "PLAN").strip()
        close = get_doc_value(doc_data, num, "CLOSE").strip()
        _, title = ITEM_DEPT_MAPPING.get(num, (dept_code, ""))
        if rev not in ("YES", "NO"):
            missing.append(f"ข้อ {num}: ยังไม่ได้เลือก YES/NO ({title})")
            continue
        if rev == "YES":
            if not resp or resp == "-":
                missing.append(f"ข้อ {num}: ยังไม่ได้ระบุผู้รับผิดชอบ ({title})")
            if not plan or plan == "-":
                missing.append(f"ข้อ {num}: ยังไม่ได้ระบุ Plan Date ({title})")
            if not close or close == "-":
                missing.append(f"ข้อ {num}: ยังไม่ได้ระบุ Actual Close ({title})")
    return len(missing) == 0, missing

def get_all_dept_completion(doc_data):
    result = {}
    all_missing = []
    for dept_code in ("PDD", "QC", "PCD", "PRD"):
        ok, missing = get_dept_completion(doc_data, dept_code)
        result[dept_code] = ok
        all_missing.extend([f"[{dept_code}] {m}" for m in missing])
    return all(result.values()), result, all_missing

def get_doc_value(doc_data, num, field_type):
    if doc_data is None:
        return ""
    
    if isinstance(doc_data, pd.Series):
        doc_data = doc_data.to_dict()
    elif isinstance(doc_data, dict) and not doc_data:
        return ""

    keys_to_check = [
        f"DOC_{num}_{field_type.upper()}",
        f"DOC_{num}_{field_type.lower()}",
        f"doc_{num}_{field_type.lower()}",
        f"DOC{num}_{field_type.upper()}",
        f"DOC{num}_{field_type.lower()}"
    ]
    
    for key in keys_to_check:
        if key in doc_data and doc_data[key] is not None and not pd.isna(doc_data[key]):
            return str(doc_data[key]).strip()
            
    return ""

def normalize_header(value):
    """ทำให้ชื่อ Header จาก Google Sheet เป็นรูปแบบเดียวกันก่อนนำไปค้นหา"""
    return (
        str(value)
        .strip()
        .upper()
        .replace("\n", "")
        .replace("\r", "")
        .replace(" ", "_")
    )


def normalize_data_key(value):
    """ทำให้ Key จากข้อมูลที่ส่งเข้า Google Sheet เป็นรูปแบบเดียวกับ Header"""
    return normalize_header(value)


def get_document_data(doc_no):
    """
    ดึงข้อมูลเอกสารจาก Google Sheets
    IMPORTANT:
    - DOCUMENT_NO ใช้หาแถว
    - SUBJECT อ่านจาก Column N (index 13) โดยตรง
    - ไม่พึ่งชื่อ Header สำหรับ SUBJECT
    """
    try:
        ws = get_worksheet()
        all_rows = ws.get_all_values()

        if not all_rows or len(all_rows) < 2:
            return None

        headers = [normalize_header(c) for c in all_rows[0]]

        # หา DOCUMENT_NO จาก Header
        doc_no_idx = 0
        for idx, h in enumerate(headers):
            if h in ["DOCUMENT_NO", "DOCUMENTNO", "DOC_NO", "DOCNO"]:
                doc_no_idx = idx
                break

        search_key = str(doc_no).strip().upper()

        for row_number, row in enumerate(all_rows[1:], start=2):
            padded_row = list(row) + [""] * max(0, len(headers) - len(row))

            if doc_no_idx >= len(padded_row):
                continue

            current_doc_no = str(padded_row[doc_no_idx]).strip().upper()

            if current_doc_no != search_key:
                continue

            row_data = {}

            # อ่านข้อมูลตาม Header
            for i, header in enumerate(headers):
                if header:
                    row_data[header] = str(
                        padded_row[i] if i < len(padded_row) else ""
                    ).strip()

            # =====================================================
            # SUBJECT = GOOGLE SHEET COLUMN N
            # N = column 14 = zero-based index 13
            # =====================================================
            subject_n = ""
            if len(row) > 13:
                subject_n = str(row[13]).strip()

            # เขียนทับด้วยค่าจาก Column N โดยตรง
            row_data["SUBJECT_TEXT"] = subject_n
            row_data["SUBJECT"] = subject_n
            row_data["SUBJECT_BY_COLUMN_N"] = subject_n
            row_data["COL_13"] = subject_n

            print("=" * 80)
            print("🔎 GOOGLE SHEET DOCUMENT FOUND")
            print("ROW NUMBER          :", row_number)
            print("DOCUMENT_NO         :", repr(current_doc_no))
            print("TOTAL COLUMNS       :", len(row))
            print("COLUMN N INDEX      :", 13)
            print("COLUMN N HEADER     :", repr(
                all_rows[0][13] if len(all_rows[0]) > 13 else "NO COLUMN N"
            ))
            print("COLUMN N SUBJECT    :", repr(subject_n))
            print("SUBJECT LENGTH      :", len(subject_n))
            print("=" * 80)

            return row_data

        print(f"❌ DOCUMENT NOT FOUND: {doc_no}")
        return None

    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการอ่านข้อมูลจาก Google Sheets: {e}")
        print(f"❌ get_document_data ERROR: {type(e).__name__}: {e}")
        return None

def get_all_documents():
    try:
        ws = get_worksheet()
        records = ws.get_all_records()
        df = pd.DataFrame(records)

        if not df.empty:
            df.columns = [normalize_header(c) for c in df.columns]

        return df

    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการดึงข้อมูล Realtime: {e}")
        print(f"❌ get_all_documents ERROR: {type(e).__name__}: {e}")
        return pd.DataFrame()


def save_to_excel(data_dict):
    """
    บันทึก/อัปเดตข้อมูลลง Google Sheets

    IMPORTANT:
    SUBJECT_TEXT จะถูกบันทึกลง Column N (column 14) โดยตรง
    เพื่อไม่ให้ปัญหา Header ทำให้ Subject หาย
    """
    try:
        ws = get_worksheet()
        all_rows = ws.get_all_values()

        if not all_rows:
            st.error("❌ Google Sheet ยังว่างเปล่า")
            return False

        headers_raw = list(all_rows[0])
        headers = [normalize_header(h) for h in headers_raw]

        normalized_data = {}
        for key, value in data_dict.items():
            clean_key = normalize_data_key(key)
            normalized_data[clean_key] = "" if value is None else str(value).strip()

        # =====================================================
        # 🖼️ IMAGE FIELD ALIAS
        # Google Sheet จริงใช้ Header = SUBJECT_IMAGE_PATH
        # แต่ Form ใช้ key = IMAGE_BASE64
        # ต้อง map ให้ตรงกัน ไม่เช่นนั้นรูปจะไม่ถูกบันทึกลง Sheet
        # =====================================================
        image_value = normalized_data.get("IMAGE_BASE64", "")
        if not image_value:
            image_value = normalized_data.get("SUBJECT_IMAGE_PATH", "")

        if image_value:
            if "SUBJECT_IMAGE_PATH" in headers:
                normalized_data["SUBJECT_IMAGE_PATH"] = image_value
            if "IMAGE_BASE64" in headers:
                normalized_data["IMAGE_BASE64"] = image_value

            print("🖼️ IMAGE TO SAVE: found", len(image_value), "chars")
        else:
            print("🖼️ IMAGE TO SAVE: EMPTY")

        doc_no = normalized_data.get("DOCUMENT_NO", "").strip().upper()
        if not doc_no:
            st.error("❌ ไม่มี DOCUMENT_NO")
            return False

        # หา DOCUMENT_NO column
        doc_col_idx = 0
        for idx, h in enumerate(headers):
            if h in ["DOCUMENT_NO", "DOCUMENTNO", "DOC_NO", "DOCNO"]:
                doc_col_idx = idx
                break

        row_index = -1
        for r_idx, row in enumerate(all_rows[1:], start=2):
            if len(row) > doc_col_idx:
                existing_doc_no = str(row[doc_col_idx]).strip().upper()
                if existing_doc_no == doc_no:
                    row_index = r_idx
                    break

        if row_index == -1:
            # สร้าง row ใหม่ตามจำนวน header
            new_row = [
                normalized_data.get(header, "")
                for header in headers
            ]

            # SUBJECT = Column N โดยตรง
            if len(new_row) < 14:
                new_row += [""] * (14 - len(new_row))

            new_row[13] = normalized_data.get("SUBJECT_TEXT", "")

            # IMAGE: เขียนลง Header จริงของ Sheet
            if "SUBJECT_IMAGE_PATH" in headers:
                image_col = headers.index("SUBJECT_IMAGE_PATH")
                new_row[image_col] = normalized_data.get("SUBJECT_IMAGE_PATH", "")
            elif "IMAGE_BASE64" in headers:
                image_col = headers.index("IMAGE_BASE64")
                new_row[image_col] = normalized_data.get("IMAGE_BASE64", "")

            ws.append_row(new_row)

            print("✅ APPEND new Google Sheet row")
            print("DOCUMENT_NO SAVED :", repr(doc_no))
            print("SUBJECT_TEXT SAVED:", repr(new_row[13]))
            if "SUBJECT_IMAGE_PATH" in headers:
                print("IMAGE SAVED COLUMN : SUBJECT_IMAGE_PATH", len(new_row[headers.index("SUBJECT_IMAGE_PATH")]))
            elif "IMAGE_BASE64" in headers:
                print("IMAGE SAVED COLUMN : IMAGE_BASE64", len(new_row[headers.index("IMAGE_BASE64")]))
            return True

        # =====================================================
        # UPDATE EXISTING ROW
        # =====================================================

        cell_updates = []

        for key, value in normalized_data.items():
            if key in headers:
                col_index = headers.index(key) + 1

                cell_updates.append(
                    gspread.Cell(
                        row=row_index,
                        col=col_index,
                        value=value
                    )
                )

        # =====================================================
        # SUBJECT_TEXT = COLUMN N DIRECTLY
        # =====================================================
        subject_value = normalized_data.get("SUBJECT_TEXT", "")

        cell_updates.append(
            gspread.Cell(
                row=row_index,
                col=14,
                value=subject_value
            )
        )

        # 🖼️ IMAGE: บังคับ update ลง Column SUBJECT_IMAGE_PATH
        # เพราะ Form ใช้ IMAGE_BASE64 แต่ Google Sheet ใช้ SUBJECT_IMAGE_PATH
        if image_value:
            if "SUBJECT_IMAGE_PATH" in headers:
                image_col = headers.index("SUBJECT_IMAGE_PATH") + 1
                cell_updates.append(
                    gspread.Cell(
                        row=row_index,
                        col=image_col,
                        value=image_value
                    )
                )
                print("🖼️ IMAGE UPDATE -> SUBJECT_IMAGE_PATH", len(image_value), "chars")
            elif "IMAGE_BASE64" in headers:
                image_col = headers.index("IMAGE_BASE64") + 1
                cell_updates.append(
                    gspread.Cell(
                        row=row_index,
                        col=image_col,
                        value=image_value
                    )
                )
                print("🖼️ IMAGE UPDATE -> IMAGE_BASE64", len(image_value), "chars")

        if cell_updates:
            ws.update_cells(cell_updates)

        print("=" * 80)
        print("✅ UPDATE Google Sheet")
        print("ROW              :", row_index)
        print("DOCUMENT_NO      :", repr(doc_no))
        print("SUBJECT -> COL N :", repr(subject_value))
        print("SUBJECT LENGTH   :", len(subject_value))
        print("=" * 80)

        return True

    except Exception as e:
        st.error(f"❌ บันทึกข้อมูลลง Google Sheets ไม่สำเร็จ: {e}")
        print(f"❌ save_to_excel ERROR: {type(e).__name__}: {e}")
        return False

def check_yes_items_completed(doc_data):
    if not doc_data:
        return False, ["ไม่พบข้อมูลเอกสาร"]
    
    missing_list = []
    has_yes_item = False
    
    for num in range(1, 20):
        rev_val = get_doc_value(doc_data, num, "REVISE").upper()
        if rev_val == "YES":
            has_yes_item = True
            resp_val = get_doc_value(doc_data, num, "RESP")
            close_val = get_doc_value(doc_data, num, "CLOSE")
            dept, title = ITEM_DEPT_MAPPING.get(num, ("-", "-"))
            
            if not resp_val or resp_val == "-":
                missing_list.append(f"ข้อ {num} [{dept}]: ยังไม่ได้ลงชื่อผู้รับผิดชอบ ({title})")
            if not close_val or close_val == "-":
                missing_list.append(f"ข้อ {num} [{dept}]: ยังไม่ได้ลงวันที่ปิดเอกสารจริง ACTUAL CLOSE ({title})")
    
    if not has_yes_item:
        return True, []
        
    is_completed = (len(missing_list) == 0)
    return is_completed, missing_list

# =============================================================
# 🔍 ฟังก์ชันคำนวณตำแหน่งปัจจุบันของเอกสารแบบ Realtime
# =============================================================
def get_realtime_location(row):
    status = str(row.get('DOC_STATUS', '')).strip().upper()
    
    if status == "APPROVED" or row.get('APPR_GM'):
        return "🟢 อนุมัติเสร็จสมบูรณ์แล้ว", "อนุมัติครบถ้วน (GM Approved)", "SUCCESS"
    
    # ตรวจสถานะทั้ง 4 แผนกจากกติกาเดียวกับหน้า Form
    _, dept_completion, _ = get_all_dept_completion(row)
    pending_depts = [d for d in ("PDD", "QC", "PCD", "PRD") if not dept_completion[d]]

    if pending_depts:
        depts_str = ", ".join(pending_depts)
        return "🔵 กำลังดำเนินการ", f"ติดอยู่ที่แผนก: {depts_str} (รอกรอก/ปิดงาน checklist)", "ENGINEER"

    # Manager loop: PDD -> QC -> PCD -> PRD -> GM
    if not row.get('APPR_PDD_MGR'):
        return "🟡 รอการอนุมัติ", "อยู่ที่ PDD Manager (รอ PDD MGR ลงนาม)", "MGR"
    elif not row.get('APPR_QCD_MGR'):
        return "🟡 รอการอนุมัติ", "อยู่ที่ QC Manager (รอ QCD MGR ลงนาม)", "MGR"
    elif not row.get('APPR_PCD_MGR'):
        return "🟡 รอการอนุมัติ", "อยู่ที่ PCD Manager (รอ PCD MGR ลงนาม)", "MGR"
    elif not row.get('APPR_PRD_MGR'):
        return "🟡 รอการอนุมัติ", "อยู่ที่ PRD Manager (รอ PRD MGR ลงนาม)", "MGR"
    elif not row.get('APPR_GM'):
        return "🟡 รอการอนุมัติ", "อยู่ที่ผู้บริหาร AGM / GM (รอ GM ลงนามอนุมัติ)", "MGR"
        
    return "🔵 กำลังดำเนินการ", "อยู่ที่แผนก: PDD (รอยืนยันส่งต่อ Manager)", "ENGINEER"

# =============================================================
# 🖨️ EXPORT TO EXCEL TEMPLATE FORM (ฉบับสมบูรณ์ แก้ไขปัญหา Subject ไม่แสดง)
# =============================================================
def export_to_printed_form(doc_no):
    if not os.path.exists(TEMPLATE_FILE):
        return None, f"❌ ไม่พบไฟล์แบบฟอร์มต้นฉบับ..."
        
    raw_data = get_document_data(doc_no)
    
    # 🔍 เพิ่มบรรทัดนี้เพื่อดูว่า get_document_data ส่งอะไรออกมากันแน่
    print("🔍 DEBUG RAW_DATA:", raw_data)
    
    if not raw_data:
        return None, "❌ ไม่พบข้อมูล..."
        
    # แปลงโครงสร้างข้อมูลให้เป็น Dict เพื่อให้ค้นหาตามชื่อหัวคอลัมน์ได้ง่าย
    doc_data = {}
    if isinstance(raw_data, dict):
        doc_data = raw_data
    elif isinstance(raw_data, (list, tuple)):
        for idx, val in enumerate(raw_data):
            doc_data[f"COL_{idx}"] = val
        potential_subjs = [str(v) for v in raw_data if v and len(str(v)) > 3 and not str(v).startswith("R0")]
        if potential_subjs:
            doc_data["SUBJECT_TEXT"] = potential_subjs[0]

    # พิมพ์ตรวจสอบข้อมูลดิบใน Console (ช่วยให้เห็นว่าคีย์ตรงกันไหม)
    print(f"🔍 ข้อมูลดิบ (doc_data) ของเอกสาร {doc_no}:", doc_data)

    def get_val(*keys):
        for k in keys:
            if k in doc_data and doc_data[k] is not None:
                return str(doc_data[k])
            for actual_k in doc_data.keys():
                clean_actual = str(actual_k).upper().replace("_", "").replace(" ", "")
                clean_target = str(k).upper().replace("_", "").replace(" ", "")
                if clean_actual == clean_target:
                    val = doc_data[actual_k]
                    if val is not None:
                        return str(val)
        return ""

    try:
        wb = openpyxl.load_workbook(TEMPLATE_FILE)
        ws = wb.active 
        
        def write_cell(coordinate, value):
            target_cell = ws[coordinate]
            for merged_range in list(ws.merged_cells.ranges):
                if target_cell.coordinate in merged_range:
                    top_left = ws.cell(row=merged_range.min_row, column=merged_range.min_col)
                    top_left.value = value
                    return top_left
            target_cell.value = value
            return target_cell

        # 📌 เขียนข้อมูล Header หลัก
        write_cell("D3", get_val("PART_NAME", "PARTNAME", "COL_1"))
        write_cell("D4", get_val("PART_NO", "PARTNO", "COL_2"))
        write_cell("F5", get_val("MASTER_DWG_NO", "DWG", "COL_3"))
        write_cell("P3", get_val("MODEL", "COL_4"))
        write_cell("X3", get_val("DATE", "COL_5"))
        write_cell("X1", get_val("DOCUMENT_NO", "DOC_NO", "DOCNO"))
        write_cell("H8", get_val("REF_DOC_NO", "REF"))
        write_cell("X4", get_val("ISSUE_BY", "USER"))
        
        write_cell("W7", get_val("EFF_EVENT", "EVENT"))
        write_cell("W8", get_val("EFF_PLAN", "PLAN"))
        write_cell("W9", get_val("EFF_ACTUAL", "ACTUAL"))

        # =========================================================
        # 📌 1. SUBJECT
        # Google Sheet Column N -> Excel D12
        # =========================================================

        # เริ่มจาก doc_data
        subj_val = str(
            doc_data.get("SUBJECT_TEXT", "")
            or doc_data.get("SUBJECT_BY_COLUMN_N", "")
            or doc_data.get("SUBJECT", "")
            or doc_data.get("SUBJECTTEXT", "")
            or doc_data.get("COL_13", "")
            or ""
        ).strip()

        # =========================================================
        # ULTIMATE FALLBACK:
        # อ่าน Google Sheet ใหม่ และหา Column N โดยตรง
        # =========================================================
        if not subj_val:
            try:
                direct_ws = get_worksheet()
                direct_rows = direct_ws.get_all_values()

                if direct_rows:
                    direct_headers = [
                        normalize_header(x) for x in direct_rows[0]
                    ]

                    direct_doc_idx = 0
                    for idx, h in enumerate(direct_headers):
                        if h in [
                            "DOCUMENT_NO",
                            "DOCUMENTNO",
                            "DOC_NO",
                            "DOCNO"
                        ]:
                            direct_doc_idx = idx
                            break

                    for direct_row in direct_rows[1:]:
                        if len(direct_row) <= direct_doc_idx:
                            continue

                        direct_doc = str(
                            direct_row[direct_doc_idx]
                        ).strip().upper()

                        if direct_doc == str(doc_no).strip().upper():
                            if len(direct_row) > 13:
                                subj_val = str(
                                    direct_row[13]
                                ).strip()
                            break

                print(
                    "🔎 DIRECT EXPORT SUBJECT =",
                    repr(subj_val)
                )

            except Exception as direct_err:
                print(
                    "❌ DIRECT EXPORT SUBJECT ERROR:",
                    type(direct_err).__name__,
                    str(direct_err)
                )

        print("=" * 80)
        print("📌 EXPORT SUBJECT")
        print("DOCUMENT NO :", repr(doc_no))
        print("SUBJECT     :", repr(subj_val))
        print("LENGTH      :", len(subj_val))
        print("=" * 80)

        # =========================================================
        # Template จริง = D12:Q14
        # =========================================================
        subject_merge = None

        for rng in list(ws.merged_cells.ranges):
            if (
                rng.min_row <= 12 <= rng.max_row
                and rng.min_col <= 4 <= rng.max_col
            ):
                subject_merge = str(rng)
                break

        print("📄 TEMPLATE SUBJECT MERGE:", subject_merge)
        print("📄 EXCEL D12 BEFORE:", repr(ws["D12"].value))

        if subject_merge:
            ws.unmerge_cells(subject_merge)

        # เขียนตรง D12
        ws["D12"] = subj_val

        ws["D12"].alignment = openpyxl.styles.Alignment(
            wrap_text=True,
            vertical="top",
            horizontal="left"
        )

        print("📄 EXCEL D12 AFTER:", repr(ws["D12"].value))

        # Merge กลับตาม Template เดิม
        if subject_merge:
            ws.merge_cells(subject_merge)
        else:
            ws.merge_cells("D12:Q14")

        print("📄 EXCEL D12 FINAL:", repr(ws["D12"].value))

        # =========================================================

        # 📌 2. ดึงและฝังรูปภาพ (ATTACHED IMAGE ในพื้นที่ R12:AA15)
        img_raw = get_val("IMAGE_BASE64", "IMAGE", "IMAGE_DATA", "IMG", "SUBJECT_IMAGE_PATH")
        if img_raw and img_raw.strip():
            try:
                img_str = img_raw.strip()
                if "," in img_str:
                    img_str = img_str.split(",")[1]

                img_bytes_data = base64.b64decode(img_str)
                pil_img = PILImage.open(io.BytesIO(img_bytes_data))

                img_bytes = io.BytesIO()
                pil_img.save(img_bytes, format='PNG')
                img_bytes.seek(0)
                
                xl_img = OpenpyxlImage(img_bytes)
                xl_img.width = 280
                xl_img.height = 95
                ws.add_image(xl_img, "R12")
            except Exception as img_err:
                print(f"⚠️ เกิดข้อผิดพลาดในการโหลดรูปภาพลง Excel: {img_err}")

        # =========================================================
        # ⚠️ IMPORTANT: I12:I14 อยู่ภายใน merged range D12:Q14
        # ใน Template จริง ดังนั้นห้ามเขียน I12/I13/I14 เพราะ
        # write_cell() จะเขียนกลับไปที่ D12 และล้าง SUBJECT ทิ้ง
        # =========================================================
        # ไม่เขียน I12/I13/I14 ที่อยู่ในพื้นที่ Subject

        # I15 อยู่นอก merged range D12:Q14 จึงเขียนได้ตามปกติ
        write_cell(
            "I15",
            f"✓ ({get_val('ATTACH_OTHERS_DETAIL')})"
            if get_val("ATTACH_OTHERS") == "YES" else ""
        )

        # =========================================================
        # 🛡️ FINAL SUBJECT PROTECTION
        # เขียน Subject ซ้ำอีกครั้งหลังจากทุก field ที่อาจกระทบ D12
        # เพื่อป้องกัน Subject ถูกเขียนทับเป็นค่าว่าง
        # =========================================================
        if subject_merge:
            # D12 เป็น top-left cell ของ merged D12:Q14
            ws["D12"] = subj_val
        else:
            ws["D12"] = subj_val

        print("🛡️ FINAL SUBJECT AFTER OTHER FIELDS:", repr(ws["D12"].value))

        judgement_val = get_val("JUDGEMENT")
        write_cell("S13", "✓" if judgement_val == "FEASIBLE" else "")
        write_cell("S14", "✓" if judgement_val == "IMPROBABILITY" else "")

        # Checklist 19 รายการ
        start_row = 19
        for i in range(1, 20):
            current_row = start_row + (i - 1)
            rev_val = str(get_doc_value(doc_data, i, "REVISE")).upper().strip()
            
            if rev_val == "YES":
                write_cell(f"K{current_row}", "✓")
                write_cell(f"M{current_row}", "")
            elif rev_val == "NO":
                write_cell(f"K{current_row}", "")
                write_cell(f"M{current_row}", "✓")
            else:
                write_cell(f"K{current_row}", "")
                write_cell(f"M{current_row}", "")
                
            write_cell(f"O{current_row}", get_doc_value(doc_data, i, "RESP"))
            write_cell(f"U{current_row}", get_doc_value(doc_data, i, "PLAN"))
            write_cell(f"Y{current_row}", get_doc_value(doc_data, i, "CLOSE"))
            
        # Manager Signatures
        write_cell("O41", get_val("APPR_PDD_MGR"))
        write_cell("N44", get_val("DATE_PDD_MGR"))
        write_cell("Q41", get_val("APPR_QCD_MGR"))
        write_cell("Q44", get_val("DATE_QCD_MGR"))
        write_cell("S41", get_val("APPR_PCD_MGR"))
        write_cell("T44", get_val("DATE_PCD_MGR"))
        write_cell("V41", get_val("APPR_PRD_MGR"))
        write_cell("W44", get_val("DATE_PRD_MGR"))
        write_cell("Y41", get_val("APPR_GM"))
        write_cell("Z44", get_val("DATE_GM"))
        
        safe_doc_no = doc_no.replace("/", "_").replace("\\", "_")
        output_filename = f"Change_Control_Sheet_{safe_doc_no}.xlsx"
        wb.save(output_filename)
        wb.close()

        # 🛡️ ตรวจสอบไฟล์ที่ save แล้วว่ามีรูปจริงหรือไม่
        try:
            verify_wb = openpyxl.load_workbook(output_filename)
            verify_ws = verify_wb.active
            print("🖼️ EXCEL IMAGE COUNT AFTER SAVE:", len(verify_ws._images))
            print("🛡️ POST-SAVE SUBJECT:", repr(verify_ws["D12"].value))
            verify_wb.close()
        except Exception as verify_err:
            print("⚠️ POST-SAVE VERIFY ERROR:", type(verify_err).__name__, verify_err)

        return output_filename, None
    except Exception as e:
        return None, f"เกิดข้อผิดพลาดในการสร้างไฟล์ Excel: {str(e)}"

# Helper ฟังก์ชันสร้างปุ่ม Download Excel จาก doc_no
def render_download_excel_button(doc_no, button_label="📥 ดาวน์โหลดไฟล์ Excel ฟอร์มจริง"):
    output_file, error_msg = export_to_printed_form(doc_no)
    if error_msg:
        st.error(error_msg)
    elif output_file and os.path.exists(output_file):
        with open(output_file, "rb") as f:
            st.download_button(
                label=button_label,
                data=f.read(),
                file_name=os.path.basename(output_file),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_btn_{doc_no}_{os.urandom(4).hex()}"
            )

# =============================================================
# 📋 Sidebar Navigation
# =============================================================
with st.sidebar:
    st.markdown(f"👤 **{st.session_state.user_name}**\n\n🏢 **บทบาท:** {st.session_state.current_dept}")
    st.markdown("---")
    
    menu = st.radio("📌 เมนูใช้งาน", ["📊 Dashboard ติดตามสถานะ Realtime", "📝 บันทึก/อนุมัติ เอกสาร"])
    
    st.markdown("---")
    if st.button("🚪 ออกจากระบบ", use_container_width=True):
        logout()

selected_dept = st.session_state.current_dept

# =============================================================
# 📊 VIEW 1: DASHBOARD ติดตามสถานะแบบ REALTIME & ภาพรวม
# =============================================================
if menu == "📊 Dashboard ติดตามสถานะ Realtime":
    col_head1, col_head2 = st.columns([3, 1])
    with col_head1:
        st.title("📊 Realtime Tracking & Overview Dashboard")
        st.caption("ระบบติดตามตำแหน่งเอกสารแบบเรียลไทม์ และสรุปภาพรวมเอกสารทั้งหมดในระบบ")
    with col_head2:
        if st.button("🔄 รีเฟรชข้อมูลล่าสุด (Refresh Data)", use_container_width=True):
            st.rerun()

    df_all = get_all_documents()
    
    if df_all.empty:
        st.warning("⚠️ ยังไม่มีข้อมูลใบงานในระบบ")
    else:
        status_info = df_all.apply(get_realtime_location, axis=1)
        df_all['MAIN_STATUS'] = [s[0] for s in status_info]
        df_all['CURRENT_LOCATION'] = [s[1] for s in status_info]
        df_all['STAGE_TYPE'] = [s[2] for s in status_info]

        st.markdown("### 🔍 ค้นหาและติดตามตำแหน่งเอกสารแบบเจาะลึก (Realtime Search)")
        doc_search_input = st.text_input("กรอก DOCUMENT NO. ที่ต้องการติดตามตำแหน่ง (เช่น R001/26) :", placeholder="พิมพ์รหัสเอกสารที่นี่...").strip().upper()

        if doc_search_input:
            matched_df = df_all[df_all['DOCUMENT_NO'].astype(str).str.strip().str.upper() == doc_search_input]
            if not matched_df.empty:
                doc_row = matched_df.iloc[0]
                
                st.markdown(f"""
                <div class="status-card">
                    <h3 style="margin:0; color:#1565c0;">📄 เอกสารเลขที่: {doc_row['DOCUMENT_NO']}</h3>
                    <p style="margin:5px 0;"><b>Customer:</b> {doc_row.get('CUSTOMER_NAME', '-')} | <b>Part Name:</b> {doc_row.get('PART_NAME', '-')} | <b>Model:</b> {doc_row.get('MODEL', '-')}</p>
                    <h4 style="margin:10px 0 0 0; color:#d32f2f;">📍 สถานะปัจจุบัน: {doc_row['CURRENT_LOCATION']}</h4>
                </div>
                """, unsafe_allow_html=True)

                st.markdown("##### 📥 **ดาวน์โหลดเอกสารฟอร์มจริง (ไม่ต้องรอ Manager อนุมัติ):**")
                render_download_excel_button(doc_search_input)

                with st.expander("📌 **คลิกเพื่อดูรายละเอียดสถานะรายข้อ (Checklist 19 ข้อ & การเซ็น MGR)**", expanded=True):
                    col_t1, col_t2 = st.columns(2)
                    with col_t1:
                        st.markdown("##### 📝 **รายการที่เลือก YES และสถานะการปิดงาน:**")
                        has_yes = False
                        for num in range(1, 20):
                            rev = get_doc_value(doc_row, num, "REVISE").upper()
                            if rev == "YES":
                                has_yes = True
                                dept, title = ITEM_DEPT_MAPPING.get(num, ("-", "-"))
                                resp = get_doc_value(doc_row, num, "RESP")
                                close_dt = get_doc_value(doc_row, num, "CLOSE")
                                if resp and close_dt and close_dt != "-":
                                    st.success(f"✅ ข้อ {num} [{dept}]: {title} (ปิดงานเรียบร้อยโดย {resp} เมื่อ {close_dt})")
                                else:
                                    st.error(f"❌ ข้อ {num} [{dept}]: {title} (ยังไม่ปิดงาน - ค้างผู้รับผิดชอบ/Actual Close)")
                        if not has_yes:
                            st.info("ℹ️ ใบงานนี้ไม่มีหัวข้อที่เลือกแก้ไข (ไม่มีข้อ YES)")

                    with col_t2:
                        st.markdown("##### 🖊️ **สถานะการลงนามอนุมัติ (Manager Approval):**")
                        st.write(f"1. **PDD MGR:** {doc_row.get('APPR_PDD_MGR') if doc_row.get('APPR_PDD_MGR') else '⏳ รอการลงนาม'}")
                        st.write(f"2. **QCD MGR:** {doc_row.get('APPR_QCD_MGR') if doc_row.get('APPR_QCD_MGR') else '⏳ รอการลงนาม'}")
                        st.write(f"3. **PRD MGR:** {doc_row.get('APPR_PRD_MGR') if doc_row.get('APPR_PRD_MGR') else '⏳ รอการลงนาม'}")
                        st.write(f"4. **PCD MGR:** {doc_row.get('APPR_PCD_MGR') if doc_row.get('APPR_PCD_MGR') else '⏳ รอการลงนาม'}")
                        st.write(f"5. **AGM / GM:** {doc_row.get('APPR_GM') if doc_row.get('APPR_GM') else '⏳ รอการลงนาม'}")
            else:
                st.error(f"❌ ไม่พบเอกสารเลขที่ '{doc_search_input}' ในฐานข้อมูล")
        
        st.markdown("---")

        st.markdown("### 📈 ภาพรวมเอกสารทั้งหมดในระบบ (System Overview)")
        
        total_docs = len(df_all)
        approved_docs = len(df_all[df_all['MAIN_STATUS'].str.contains("Approved")])
        pending_mgr = len(df_all[df_all['MAIN_STATUS'].str.contains("รอการอนุมัติ")])
        in_progress = total_docs - approved_docs - pending_mgr

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("📄 ใบงานทั้งหมดในระบบ", f"{total_docs} รายการ")
        m2.metric("🔵 รอแผนกปิดข้อ YES", f"{in_progress} รายการ")
        m3.metric("🟡 รอผู้จัดการอนุมัติ (Wait MGR)", f"{pending_mgr} รายการ")
        m4.metric("🟢 อนุมัติเสร็จสมบูรณ์", f"{approved_docs} รายการ")

        st.markdown("<br>", unsafe_allow_html=True)

        st.markdown("##### 📋 **ตารางติดตามตำแหน่งเอกสารแบบเรียลไทม์ (Realtime Document List)**")
        
        c_filter1, c_filter2 = st.columns(2)
        with c_filter1:
            filter_status = st.selectbox("🎯 กรองตามสถานะหลัก:", ["ทั้งหมด", "🔵 กำลังดำเนินการ", "🟡 รอการอนุมัติ", "🟢 อนุมัติเสร็จสมบูรณ์"])
        with c_filter2:
            filter_text = st.text_input("🔍 กรองตาม Customer / Part Name / Issue By:", key="table_search")

        display_df = df_all.copy()
        
        if filter_status != "ทั้งหมด":
            display_df = display_df[display_df['MAIN_STATUS'].str.contains(filter_status[:2])]
            
        if filter_text:
            ft = filter_text.strip().upper()
            display_df = display_df[
                display_df['CUSTOMER_NAME'].astype(str).str.upper().str.contains(ft) |
                display_df['PART_NAME'].astype(str).str.upper().str.contains(ft) |
                display_df['ISSUE_BY'].astype(str).str.upper().str.contains(ft)
            ]

        show_cols = ['DOCUMENT_NO', 'CUSTOMER_NAME', 'PART_NAME', 'MODEL', 'ISSUE_BY', 'MAIN_STATUS', 'CURRENT_LOCATION']
        valid_cols = [c for c in show_cols if c in display_df.columns]

        st.dataframe(
            display_df[valid_cols].rename(columns={
                'DOCUMENT_NO': 'เลขที่เอกสาร',
                'CUSTOMER_NAME': 'ลูกค้า',
                'PART_NAME': 'ชื่อพาร์ท',
                'MODEL': 'โมเดล',
                'ISSUE_BY': 'ผู้จัดทำ',
                'MAIN_STATUS': 'สถานะหลัก',
                'CURRENT_LOCATION': '📍 ตำแหน่งปัจจุบันของเอกสาร (ผู้รับผิดชอบ)'
            }),
            use_container_width=True,
            hide_index=True
        )

        st.markdown("##### 📥 **ดาวน์โหลดฟอร์ม Excel ด่วนจากตาราง:**")
        quick_doc = st.selectbox("เลือกเลขที่เอกสารที่ต้องการโหลด Excel:", ["-"] + list(display_df['DOCUMENT_NO'].unique()))
        if quick_doc != "-":
            render_download_excel_button(quick_doc, f"📥 ดาวน์โหลดไฟล์ Excel ใบงาน {quick_doc}")

# =============================================================
# 📝 VIEW 2: หน้าบันทึก / อนุมัติ เอกสาร
# =============================================================
else:
    st.title("KFT - RELATED DOCUMENT CHANGE CONTROL SYSTEM")

    if "Print Form" in selected_dept:
        st.subheader("🖨️ เจ้าหน้าที่พิมพ์เอกสาร / Download Form")
        doc_search = st.text_input("พิมพ์เลขที่เอกสาร DOCUMENT NO. เพื่อพิมพ์ฟอร์ม (เช่น R001/26):")
        if doc_search:
            render_download_excel_button(doc_search)
    else:
        st.subheader(f"📌 การจัดการเอกสารสำหรับ: {selected_dept}")
        
        user_dept = get_user_dept_code(selected_dept)
        is_manager = "MGR -" in str(selected_dept) or "AGM / GM" in str(selected_dept)
        can_create_new = user_dept == "PDD" and not is_manager

        if can_create_new:
            action_type = st.radio(
                "เลือกการทำงาน:",
                ["สร้างเอกสารใหม่ (PDD)", "ค้นหา/แก้ไข เอกสารเดิม"],
                horizontal=True,
                key="action_type_main"
            )
        else:
            action_type = "ค้นหา/แก้ไข เอกสารเดิม"
            st.info("🔐 บัญชีนี้ค้นหาเอกสารเดิมได้ และแก้ไขได้เฉพาะ checklist ของแผนกตนเอง")

        doc_data = {}
        doc_no = ""

        if action_type == "ค้นหา/แก้ไข เอกสารเดิม":
            doc_no = st.text_input(
                "กรอก DOCUMENT NO. ที่ต้องการดึงข้อมูล:",
                key="search_doc_main",
                placeholder="เช่น R001/26"
            ).strip().upper()

            if doc_no:
                doc_data = get_document_data(doc_no) or {}
                if doc_data:
                    loaded_key = st.session_state.get("_loaded_doc_no")
                    if loaded_key != doc_no:
                        st.session_state["_loaded_doc_no"] = doc_no
                        st.session_state["doc_no_field"] = doc_data.get("DOCUMENT_NO", doc_no)
                        st.session_state["customer_field"] = doc_data.get("CUSTOMER_NAME", "")
                        st.session_state["part_name_field"] = doc_data.get("PART_NAME", "")
                        st.session_state["part_no_field"] = doc_data.get("PART_NO", "")
                        st.session_state["model_field"] = doc_data.get("MODEL", "")
                        st.session_state["dwg_field"] = doc_data.get("MASTER_DWG_NO", "")
                        st.session_state["ref_doc_field"] = doc_data.get("REF_DOC_NO", "")
                        st.session_state["issue_by_field"] = doc_data.get("ISSUE_BY", st.session_state.user_name)
                        issue_date_loaded = doc_data.get("DATE", "")
                        try:
                            st.session_state["issue_date_field"] = date.fromisoformat(str(issue_date_loaded)[:10]) if issue_date_loaded else date.today()
                        except Exception:
                            st.session_state["issue_date_field"] = date.today()
                        st.session_state["subject_field"] = doc_data.get("SUBJECT_TEXT", "")
                        for n in range(1, 20):
                            st.session_state[f"rev_{n}"] = get_doc_value(doc_data, n, "REVISE").upper() or "-"
                            st.session_state[f"resp_{n}"] = get_doc_value(doc_data, n, "RESP")
                            st.session_state[f"plan_{n}"] = get_doc_value(doc_data, n, "PLAN")
                            st.session_state[f"close_{n}"] = get_doc_value(doc_data, n, "CLOSE")
                        st.session_state.pop("uploaded_image_widget", None)
                    st.success(f"✅ โหลดข้อมูลเดิมของเอกสาร {doc_no} สำเร็จ")
                    print("📝 FORM LOAD SUBJECT_TEXT:", repr(doc_data.get("SUBJECT_TEXT", "")))
                else:
                    st.error("❌ ไม่พบข้อมูลเอกสารในระบบ")
        else:
            if st.session_state.get("_loaded_doc_no") != "__NEW__":
                st.session_state["_loaded_doc_no"] = "__NEW__"
                for key in ["doc_no_field","customer_field","part_name_field","part_no_field","model_field","dwg_field","ref_doc_field","issue_by_field","subject_field","issue_date_field"]:
                    st.session_state.pop(key, None)
                for n in range(1, 20):
                    for prefix in ("rev_", "resp_", "plan_", "close_"):
                        st.session_state.pop(f"{prefix}{n}", None)

        st.markdown("---")

        # 📌 เพิ่มปุ่มดาวน์โหลดเอกสารไว้ด้านบนสุดของแบบฟอร์มหากมีข้อมูลเอกสารแล้ว
        current_doc_id = doc_data.get("DOCUMENT_NO", doc_no)
        if current_doc_id:
            col_dl1, col_dl2 = st.columns([2, 1])
            with col_dl1:
                st.markdown(f"#### 📄 จัดการเอกสารเลขที่: **{current_doc_id}**")
            with col_dl2:
                render_download_excel_button(current_doc_id, "📥 ดาวน์โหลด Excel ใบงานนี้")
            st.markdown("---")

        st.subheader("📄 ข้อมูลทั่วไปของเอกสาร (General Information)")
        st.caption("🔐 ข้อมูลส่วนหลักของเอกสาร (Topic/General Information) เป็นข้อมูลที่ PDD เป็นผู้จัดทำ — แผนกอื่นสามารถดูได้ แต่แก้ไขไม่ได้")

        # ข้อมูลหลักของเอกสารให้ PDD เป็นเจ้าของข้อมูลเท่านั้น
        # แผนกอื่นต้องเห็นค่าที่ PDD บันทึกไว้ แต่เป็น Read-only
        pdd_can_edit_main = (user_dept == "PDD" and not is_manager)

        c1, c2, c3 = st.columns(3)
        with c1:
            doc_no_val = st.text_input("DOCUMENT NO.", key="doc_no_field", disabled=not pdd_can_edit_main)
            customer_name = st.text_input("CUSTOMER NAME", key="customer_field", disabled=not pdd_can_edit_main)
            part_name = st.text_input("PART NAME", key="part_name_field", disabled=not pdd_can_edit_main)
        with c2:
            part_no = st.text_input("PART NO.", key="part_no_field", disabled=not pdd_can_edit_main)
            model = st.text_input("MODEL", key="model_field", disabled=not pdd_can_edit_main)
            master_dwg_no = st.text_input("MASTER DWG NO.", key="dwg_field", disabled=not pdd_can_edit_main)
        with c3:
            issue_date_raw = doc_data.get("DATE", "")
            try:
                issue_date_default = date.fromisoformat(str(issue_date_raw)[:10]) if issue_date_raw else date.today()
            except Exception:
                issue_date_default = date.today()
            issue_date = st.date_input("DATE", value=issue_date_default, key="issue_date_field", disabled=not pdd_can_edit_main)
            ref_doc_no = st.text_input("REF. DOC. NO.", key="ref_doc_field", disabled=not pdd_can_edit_main)
            issue_by = st.text_input("ISSUE BY", key="issue_by_field", disabled=not pdd_can_edit_main)

        # =========================================================
        # 📌 SUBJECT / TOPIC (พื้นที่ D12:Q14) & IMAGE (พื้นที่ R12:AA14)
        # =========================================================
        st.markdown("---")
        st.subheader("📝 Topic / Subject และ รูปภาพแนบ (Subject & Attached Image)")
        if pdd_can_edit_main:
            st.info("✏️ PDD สามารถแก้ไข Topic / Subject และรูปภาพแนบได้")
        else:
            st.info("👁️ แสดงข้อมูล Topic / Subject ที่ PDD บันทึกไว้ — แผนกนี้ไม่สามารถแก้ไขได้")

        col_subj, col_img = st.columns([1, 1])
        
        with col_subj:
            st.markdown("**1. TOPIC / SUBJECT (รายละเอียดเรื่องที่เปลี่ยนแปลง - พื้นที่ D12:Q14)**")
            subject_text = st.text_area(
                "Topic / Subject:",
                height=150,
                placeholder="ระบุข้อความรายละเอียดการเปลี่ยนแปลงที่นี่...",
                key="subject_field",
                disabled=not pdd_can_edit_main
            )

        with col_img:
            st.markdown("**2. ATTACHED IMAGE (รูปภาพประกอบ - พื้นที่ R12:AA14)**")
            uploaded_image = None
            if pdd_can_edit_main:
                uploaded_image = st.file_uploader(
                    "อัปโหลดรูปภาพแนบ (JPG / PNG):",
                    type=["jpg", "jpeg", "png"],
                    key="uploaded_image_widget"
                )
            else:
                st.caption("🔒 รูปภาพแนบเป็นข้อมูลจาก PDD และไม่สามารถเปลี่ยนแปลงได้จากแผนกนี้")
            
            image_base64_str = (
                doc_data.get("IMAGE_BASE64", "")
                or doc_data.get("SUBJECT_IMAGE_PATH", "")
                or doc_data.get("IMAGE", "")
                or ""
            )
            if uploaded_image is not None:
                image_base64_str = convert_image_to_base64(uploaded_image)
                st.image(uploaded_image, caption="รูปภาพที่อัปโหลดใหม่", use_container_width=True)
            elif image_base64_str:
                st.image(image_base64_str, caption="รูปภาพปัจจุบันในระบบ", use_container_width=True)

        st.markdown("---")
        st.subheader("📋 รายการตรวจสอบและแผนการดำเนินงาน (Checklist 19 รายการ)")

        # =========================================================
        # Checklist 19 ข้อ แยกตามสิทธิ์แผนก
        # PDD 1-7 | QC 8-15 | PCD 16-17 | PRD 18-19
        # =========================================================
        checklist_results = {}
        st.markdown("### 🧭 แบ่งส่วนงาน")
        st.info("PDD: ข้อ 1-7  |  QC: ข้อ 8-15  |  PCD: ข้อ 16-17  |  PRD: ข้อ 18-19")

        for dept_code, label in DEPT_LABELS.items():
            item_numbers = DEPT_ITEM_RANGES[dept_code]
            dept_ok, dept_missing = get_dept_completion(doc_data, dept_code)
            status_text = "✅ ครบ" if dept_ok else "⏳ ยังไม่ครบ"
            with st.expander(f"{label} — {status_text}", expanded=(dept_code == user_dept or is_manager)):
                if dept_code == user_dept and not is_manager:
                    st.caption("✏️ แผนกของคุณ: แก้ไขข้อมูลได้เฉพาะข้อในส่วนนี้")
                else:
                    st.caption("🔒 อ่านอย่างเดียว: ไม่มีสิทธิ์แก้ไขรายการของแผนกนี้")

                for i in item_numbers:
                    _, title = ITEM_DEPT_MAPPING[i]
                    curr_rev = get_doc_value(doc_data, i, "REVISE").upper()
                    curr_resp = get_doc_value(doc_data, i, "RESP")
                    curr_plan = get_doc_value(doc_data, i, "PLAN")
                    curr_close = get_doc_value(doc_data, i, "CLOSE")
                    st.markdown(f"**ข้อ {i}. [{dept_code}] {title}**")

                    if dept_code == user_dept and not is_manager:
                        col_a, col_b, col_c, col_d = st.columns([1.5, 2, 2, 2])
                        with col_a:
                            rev_choice = st.radio(f"แก้ไขข้อ {i}", ["-", "YES", "NO"], key=f"rev_{i}", horizontal=True)
                        with col_b:
                            resp_choice = st.text_input(f"ผู้รับผิดชอบ ข้อ {i}", key=f"resp_{i}")
                        with col_c:
                            plan_choice = st.text_input(f"Plan Date ข้อ {i}", key=f"plan_{i}")
                        with col_d:
                            close_choice = st.text_input(f"Actual Close ข้อ {i}", key=f"close_{i}")
                        checklist_results[f"DOC_{i}_REVISE"] = rev_choice
                        checklist_results[f"DOC_{i}_RESP"] = resp_choice
                        checklist_results[f"DOC_{i}_PLAN"] = plan_choice
                        checklist_results[f"DOC_{i}_CLOSE"] = close_choice
                    else:
                        c1, c2, c3, c4 = st.columns([1.5, 2, 2, 2])
                        c1.write(f"REVISE: **{curr_rev or '-'}**")
                        c2.write(f"RESP: **{curr_resp or '-'}**")
                        c3.write(f"PLAN: **{curr_plan or '-'}**")
                        c4.write(f"CLOSE: **{curr_close or '-'}**")

                if dept_ok:
                    st.success("ส่วนงานนี้กรอกและปิดงานครบแล้ว")
                else:
                    for m in dept_missing[:8]:
                        st.warning(m)
                    if len(dept_missing) > 8:
                        st.caption(f"และอีก {len(dept_missing)-8} รายการ...")

        st.subheader("🖊️ การลงนามอนุมัติเอกสาร (Manager Approval)")

        appr_pdd = doc_data.get("APPR_PDD_MGR", "")
        appr_qcd = doc_data.get("APPR_QCD_MGR", "")
        appr_prd = doc_data.get("APPR_PRD_MGR", "")
        appr_pcd = doc_data.get("APPR_PCD_MGR", "")
        appr_gm  = doc_data.get("APPR_GM", "")

        is_completed, dept_completion, missing_reasons = get_all_dept_completion(doc_data)

        st.markdown("### 📊 สถานะการปิดงานก่อนเข้า Approval Loop")
        sc1, sc2, sc3, sc4 = st.columns(4)
        for col, dept_code in zip([sc1, sc2, sc3, sc4], ["PDD", "QC", "PCD", "PRD"]):
            with col:
                st.metric(dept_code, "✅ ครบ" if dept_completion[dept_code] else "⏳ ยังไม่ครบ")

        if not is_completed:
            st.warning("⚠️ ยังไม่เปิด Manager Approval — ทั้ง 4 แผนกต้องกรอกและปิดงานส่วนของตนเองให้ครบก่อน")
            for m in missing_reasons[:20]:
                st.write(f"- {m}")
        else:
            st.success("✅ PDD + QC + PCD + PRD ปิดงานครบแล้ว — เข้าสู่ Manager Approval Loop ได้")

        col_m1, col_m2, col_m3, col_m4, col_m5 = st.columns(5)
        with col_m1:
            st.markdown("**PDD MGR**")
            st.info(appr_pdd if appr_pdd else "รอลงนาม")
            if "MGR - PDD" in selected_dept and is_completed and not appr_pdd:
                if st.button("🖊️ ลงนามอนุมัติ (PDD MGR)"):
                    appr_pdd = st.session_state.user_name
                    date_pdd = str(date.today())
                    save_to_excel({"DOCUMENT_NO": doc_no_val, "APPR_PDD_MGR": appr_pdd, "DATE_PDD_MGR": date_pdd})
                    send_approval_next_step_email(doc_no_val, customer_name, part_name, "ผู้จัดการ PDD", "QCD_MGR", "ผู้จัดการ QC (QCD MGR)")
                    st.success("ลงนามอนุมัติสำเร็จ!")
                    st.rerun()

        with col_m2:
            st.markdown("**QCD MGR**")
            st.info(appr_qcd if appr_qcd else "รอลงนาม")
            if "MGR - QCD" in selected_dept and is_completed and appr_pdd and not appr_qcd:
                if st.button("🖊️ ลงนามอนุมัติ (QCD MGR)"):
                    appr_qcd = st.session_state.user_name
                    date_qcd = str(date.today())
                    save_to_excel({"DOCUMENT_NO": doc_no_val, "APPR_QCD_MGR": appr_qcd, "DATE_QCD_MGR": date_qcd})
                    send_approval_next_step_email(doc_no_val, customer_name, part_name, "ผู้จัดการ QC", "PCD_MGR", "ผู้จัดการ PCD (PCD MGR)")
                    st.success("ลงนามอนุมัติสำเร็จ!")
                    st.rerun()

        with col_m3:
            st.markdown("**PRD MGR**")
            st.info(appr_prd if appr_prd else "รอลงนาม")
            if "MGR - PD" in selected_dept and is_completed and appr_pcd and not appr_prd:
                if st.button("🖊️ ลงนามอนุมัติ (PRD MGR)"):
                    appr_prd = st.session_state.user_name
                    date_prd = str(date.today())
                    save_to_excel({"DOCUMENT_NO": doc_no_val, "APPR_PRD_MGR": appr_prd, "DATE_PRD_MGR": date_prd})
                    send_approval_next_step_email(doc_no_val, customer_name, part_name, "ผู้จัดการ PRD", "GM", "ผู้บริหาร (AGM / GM)")
                    st.success("ลงนามอนุมัติสำเร็จ!")
                    st.rerun()

        with col_m4:
            st.markdown("**PCD MGR**")
            st.info(appr_pcd if appr_pcd else "รอลงนาม")
            if "MGR - PCD" in selected_dept and is_completed and appr_qcd and not appr_pcd:
                if st.button("🖊️ ลงนามอนุมัติ (PCD MGR)"):
                    appr_pcd = st.session_state.user_name
                    date_pcd = str(date.today())
                    save_to_excel({"DOCUMENT_NO": doc_no_val, "APPR_PCD_MGR": appr_pcd, "DATE_PCD_MGR": date_pcd})
                    send_approval_next_step_email(doc_no_val, customer_name, part_name, "ผู้จัดการ PCD", "PRD_MGR", "ผู้จัดการ PRD (PRD MGR)")
                    st.success("ลงนามอนุมัติสำเร็จ!")
                    st.rerun()

        with col_m5:
            st.markdown("**AGM / GM**")
            st.info(appr_gm if appr_gm else "รอลงนาม")
            if "AGM / GM" in selected_dept and is_completed and appr_prd and not appr_gm:
                if st.button("🖊️ ลงนามอนุมัติขั้นสุดท้าย (GM)"):
                    appr_gm = st.session_state.user_name
                    date_gm = str(date.today())
                    save_to_excel({"DOCUMENT_NO": doc_no_val, "APPR_GM": appr_gm, "DATE_GM": date_gm, "DOC_STATUS": "APPROVED"})
                    send_final_approved_email(doc_no_val, customer_name, part_name, appr_gm)
                    st.success("อนุมัติเสร็จสมบูรณ์เรียบร้อย!")
                    st.rerun()

        st.markdown("---")
        
        # ปุ่มบันทึกข้อมูลหลัก และ ปุ่มดาวน์โหลดเอกสารท้ายฟอร์ม
        col_b1, col_b2 = st.columns([2, 1])
        with col_b1:
            if st.button("💾 บันทึกข้อมูลลงฐานข้อมูล (Save Data)", type="primary", use_container_width=True):
                if not doc_no_val:
                    st.error("❌ กรุณาระบุ DOCUMENT NO.")
                else:
                    # ค้นหาบรรทัด save_data ในส่วนการกดปุ่มบันทึก แล้วปรับให้ส่งทั้ง SUBJECT_TEXT และ SUBJECT
                    # PDD เป็นเจ้าของข้อมูล Topic / General Information
                    # แผนกอื่นส่งเฉพาะ checklist ของตนเอง เพื่อป้องกันการเขียนทับข้อมูลหลักของ PDD
                    save_data = {"DOCUMENT_NO": doc_no_val}
                    if pdd_can_edit_main:
                        save_data.update({
                            "CUSTOMER_NAME": customer_name,
                            "PART_NAME": part_name,
                            "PART_NO": part_no,
                            "MODEL": model,
                            "MASTER_DWG_NO": master_dwg_no,
                            "DATE": str(issue_date),
                            "REF_DOC_NO": ref_doc_no,
                            "ISSUE_BY": issue_by,
                            "SUBJECT_TEXT": str(subject_text).strip(),
                            "SUBJECT": str(subject_text).strip(),
                            "IMAGE_BASE64": image_base64_str
                        })
                    save_data.update(checklist_results)

                    if save_to_excel(save_data):
                        st.success(f"✅ บันทึกข้อมูลเอกสาร {doc_no_val} เรียบร้อยแล้ว!")
                        
                        # ตรวจสอบการส่ง Email แจ้งเตือนเมื่อวิศวกรปิดข้อ YES ครบ
                        check_data = get_document_data(doc_no_val)
                        completed, _ = check_yes_items_completed(check_data)
                        all_done, _, all_missing = get_all_dept_completion(check_data)
                        if all_done and not check_data.get("APPR_PDD_MGR"):
                            send_all_completed_alert_email(doc_no_val, customer_name, part_name)
                            st.info("📧 ทุกแผนกปิดงานครบแล้ว — ส่งอีเมลแจ้ง PDD Manager เพื่อเริ่ม Approval Loop")
                        elif not all_done:
                            st.info(f"ℹ️ บันทึกแล้ว แต่ยังไม่ส่งเข้า Manager Approval เพราะยังเหลือ {len(all_missing)} รายการ")
        with col_b2:
            if doc_no_val:
                render_download_excel_button(doc_no_val, "📥 ดาวน์โหลด Excel ฟอร์มจริง")
