import os
import io
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
import pandas as pd
import streamlit as st
from datetime import date
from PIL import Image
import gspread
from google.oauth2 import service_account
import base64

# =============================================================
# 🖼️ BASE64 IMAGE CONVERSION & RESIZE
# =============================================================
def convert_image_to_base64(uploaded_file, max_size=(600, 600), quality=60):
    """ย่อขนาดและบีบอัดรูปภาพก่อนแปลงเป็น Base64 เพื่อไม่ให้เกิน 50,000 ตัวอักษรใน Google Sheets"""
    if uploaded_file is not None:
        try:
            img = Image.open(uploaded_file)
            if img.mode in ("RGBA", "P"):
                img = img.convert("RGB")
            img.thumbnail(max_size, Image.Resampling.LANCZOS)
            
            buffer = io.BytesIO()
            img.save(buffer, format="JPEG", quality=quality, optimize=True)
            bytes_data = buffer.getvalue()
            
            base64_str = base64.b64encode(bytes_data).decode()
            return f"data:image/jpeg;base64,{base64_str}"
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาดในการแปลงรูปภาพ: {e}")
            return ""
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
    "prd_user": {"password": sec_passwords.get("prd_user", ""), "dept": "PRO (Production / PD)", "name": "ENGINEER Production"},
    "mgr_pdd": {"password": sec_passwords.get("mgr_pdd", ""), "dept": "MGR - PDD (ผู้จัดการ PDD)", "name": "ผู้จัดการ PDD"},
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
    18: ("PRO", "WORKING INSTRUCTION."), 19: ("PRO", "TRAINING PRODUCTION.")
}

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

def get_document_data(doc_no):
    try:
        ws = get_worksheet()
        records = ws.get_all_records()
        df = pd.DataFrame(records)
        
        if not df.empty:
            df.columns = [str(c).strip().upper() for c in df.columns]
            target_col = None
            for col in ['DOCUMENT_NO', 'DOCUMENT NO', 'DOC_NO']:
                if col in df.columns:
                    target_col = col
                    break
            
            if target_col:
                df[target_col] = df[target_col].astype(str).str.strip().str.upper()
                search_key = str(doc_no).strip().upper()
                matched = df[df[target_col] == search_key]
                if not matched.empty:
                    row_data = matched.iloc[0].to_dict()
                    return {str(k): ("" if pd.isna(v) else str(v).strip()) for k, v in row_data.items()}
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการอ่านข้อมูลจาก Google Sheets: {e}")
    return None

def get_all_documents():
    try:
        ws = get_worksheet()
        records = ws.get_all_records()
        df = pd.DataFrame(records)
        if not df.empty:
            df.columns = [str(c).strip().upper() for c in df.columns]
            return df
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการดึงข้อมูล Realtime: {e}")
    return pd.DataFrame()

def save_to_excel(data_dict):
    try:
        ws = get_worksheet()
        headers = [str(h).strip().upper() for h in ws.row_values(1)]
        if not headers:
            st.error("❌ Google Sheet ยังไม่มี Header ในบรรทัดแรก")
            return False

        records = ws.get_all_records()
        df_old = pd.DataFrame(records)
        doc_no = str(data_dict.get('DOCUMENT_NO', '')).strip().upper()

        target_col = 'DOCUMENT_NO' if 'DOCUMENT_NO' in headers else headers[0]

        if not df_old.empty and target_col in df_old.columns and doc_no in df_old[target_col].astype(str).str.strip().str.upper().values:
            df_old[target_col] = df_old[target_col].astype(str).str.strip().str.upper()
            row_index = df_old[df_old[target_col] == doc_no].index[0] + 2
            cell_updates = []
            for key, value in data_dict.items():
                clean_key = str(key).strip().upper()
                if clean_key in headers and value is not None and value != "":
                    col_index = headers.index(clean_key) + 1
                    cell_updates.append(gspread.Cell(row=row_index, col=col_index, value=str(value)))
            if cell_updates:
                ws.update_cells(cell_updates)
        else:
            new_row = [str(data_dict.get(col, data_dict.get(col.upper(), ""))) for col in headers]
            ws.append_row(new_row)
        return True
    except Exception as e:
        st.error(f"❌ บันทึกข้อมูลลง Google Sheets ไม่สำเร็จ: {e}")
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
    
    pending_depts = set()
    for num in range(1, 20):
        rev = get_doc_value(row, num, "REVISE").upper()
        if rev == "YES":
            close_val = get_doc_value(row, num, "CLOSE")
            resp_val = get_doc_value(row, num, "RESP")
            if not close_val or close_val == "-" or not resp_val or resp_val == "-":
                dept, _ = ITEM_DEPT_MAPPING.get(num, ("-", "-"))
                pending_depts.add(dept)
                
    if pending_depts:
        depts_str = ", ".join(sorted(list(pending_depts)))
        return "🔵 กำลังดำเนินการ", f"ติดอยู่ที่แผนก: {depts_str} (รอปิดข้อ YES & ลง Actual Close)", "ENGINEER"

    if not row.get('APPR_PDD_MGR'):
        return "🟡 รอการอนุมัติ", "อยู่ที่แผนก: PDD (รอ PDD Manager ลงนาม)", "MGR"
    elif not row.get('APPR_QCD_MGR'):
        return "🟡 รอการอนุมัติ", "อยู่ที่แผนก: QC (รอ QCD Manager ลงนาม)", "MGR"
    elif not row.get('APPR_PRD_MGR'):
        return "🟡 รอการอนุมัติ", "อยู่ที่แผนก: PRO/PD (รอ PRD Manager ลงนาม)", "MGR"
    elif not row.get('APPR_PCD_MGR'):
        return "🟡 รอการอนุมัติ", "อยู่ที่แผนก: PCD (รอ PCD Manager ลงนาม)", "MGR"
    elif not row.get('APPR_GM'):
        return "🟡 รอการอนุมัติ", "อยู่ที่ผู้บริหาร: AGM / GM (รอ GM ลงนามอนุมัติ)", "MGR"
        
    return "🔵 กำลังดำเนินการ", "อยู่ที่แผนก: PDD (รอยืนยันส่งต่อ Manager)", "ENGINEER"

# =============================================================
# 🖨️ EXPORT TO EXCEL TEMPLATE FORM (ปรับปรุง Subject และเพิ่มระบบฝังรูปภาพ R12:AA14)
# =============================================================
def export_to_printed_form(doc_no):
    if not os.path.exists(TEMPLATE_FILE):
        return None, f"❌ ไม่พบไฟล์แบบฟอร์มต้นฉบับ '{TEMPLATE_FILE}' ในโฟลเดอร์โปรเจกต์"
        
    doc_data = get_document_data(doc_no)
    if not doc_data:
        return None, "❌ ไม่พบข้อมูลของเอกสารเลขที่นี้ในฐานข้อมูล"
        
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

        # เขียนข้อมูล Header หลัก
        write_cell("D3", doc_data.get("PART_NAME", ""))
        write_cell("D4", doc_data.get("PART_NO", ""))
        write_cell("F5", doc_data.get("MASTER_DWG_NO", ""))
        write_cell("P3", doc_data.get("MODEL", ""))
        write_cell("X3", doc_data.get("DATE", ""))
        write_cell("X1", doc_data.get("DOCUMENT_NO", ""))
        write_cell("H8", doc_data.get("REF_DOC_NO", ""))
        write_cell("X4", doc_data.get("ISSUE_BY", ""))
        
        write_cell("W7", doc_data.get("EFF_EVENT", ""))
        write_cell("W8", doc_data.get("EFF_PLAN", ""))
        write_cell("W9", doc_data.get("EFF_ACTUAL", ""))

       # -------------------------------------------------------------
# 📌 1. จัดการดึงและเขียนค่า SUBJECT ลง D12
# -------------------------------------------------------------
subj_val = ""

# วนลูปหาจากทุก Key ใน doc_data ที่น่าจะเป็นไปได้
for k, v in doc_data.items():
    clean_k = str(k).upper().replace("_", "").replace(" ", "")
    if any(x in clean_k for x in ["SUBJECT", "DETAILS", "DETAIL"]):
        if v and str(v).strip() and str(v).strip().lower() != "none":
            subj_val = str(v).strip()
            break

# เขียนลง D12 (Top-Left Cell ของ Merged Range D12:Q15)
ws["D12"].value = subj_val
ws["D12"].alignment = openpyxl.styles.Alignment(wrap_text=True, vertical="top", horizontal="left")


# -------------------------------------------------------------
# 📌 2. จัดการแปลงและใส่อยู่ในพื้นที่ R12:AA15 (IMAGE)
# -------------------------------------------------------------
img_raw = doc_data.get("IMAGE_BASE64") or doc_data.get("IMAGE") or doc_data.get("IMAGE_DATA") or ""

if img_raw and str(img_raw).strip():
    try:
        img_str = str(img_raw).strip()
        # ตัด Prefix ออกหากมีติดมากับ Base64 (เช่น data:image/png;base64,...)
        if "," in img_str:
            img_str = img_str.split(",")[1]

        # แปลง Base64 กลับเป็น Image Bytes
        img_data = base64.b64decode(img_str)
        pil_img = PILImage.open(io.BytesIO(img_data))

        # บันทึกลง BytesIO แบบ PNG
        img_bytes = io.BytesIO()
        pil_img.save(img_bytes, format='PNG')
        img_bytes.seek(0)

        # สร้าง Openpyxl Image Object และปรับขนาดให้พอดีกับช่อง R12:AA15
        xl_img = OpenpyxlImage(img_bytes)
        xl_img.width = 280   # ความกว้างพิกเซล
        xl_img.height = 85   # ความสูงพิกเซล

        # วางลงที่ตำแหน่งเซลล์ R12
        ws.add_image(xl_img, "R12")
    except Exception as img_err:
        print(f"⚠️ เกิดข้อผิดพลาดในการโหลดรูปภาพลง Excel: {img_err}")

        # เครื่องหมายถูก Checkbox
        write_cell("I12", "✓" if doc_data.get("ATTACH_DRAWING") == "YES" else "")
        write_cell("I13", "✓" if doc_data.get("ATTACH_ECI") == "YES" else "")
        write_cell("I14", "✓" if doc_data.get("ATTACH_MEETING") == "YES" else "")
        write_cell("I15", f"✓ ({doc_data.get('ATTACH_OTHERS_DETAIL', '')})" if doc_data.get("ATTACH_OTHERS") == "YES" else "")

        judgement_val = doc_data.get("JUDGEMENT", "")
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
        write_cell("O41", doc_data.get("APPR_PDD_MGR", ""))
        write_cell("N44", doc_data.get("DATE_PDD_MGR", ""))
        write_cell("Q41", doc_data.get("APPR_QCD_MGR", ""))
        write_cell("Q44", doc_data.get("DATE_QCD_MGR", ""))
        write_cell("S41", doc_data.get("APPR_PCD_MGR", ""))
        write_cell("T44", doc_data.get("DATE_PCD_MGR", ""))
        write_cell("V41", doc_data.get("APPR_PRD_MGR", ""))
        write_cell("W44", doc_data.get("DATE_PRD_MGR", ""))
        write_cell("Y41", doc_data.get("APPR_GM", ""))
        write_cell("Z44", doc_data.get("DATE_GM", ""))
        
        safe_doc_no = doc_no.replace("/", "_").replace("\\", "_")
        output_filename = f"Change_Control_Sheet_{safe_doc_no}.xlsx"
        wb.save(output_filename)
        wb.close()
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
        
        # ส่วนเลือกสร้างเอกสารใหม่ หรือดึงเอกสารเดิมมาแก้ไข
        action_type = st.radio("เลือกการทำงาน:", ["สร้างเอกสารใหม่ (PDD)", "ค้นหา/แก้ไข เอกสารเดิม"], horizontal=True)
        
        doc_data = {}
        doc_no = ""

        if action_type == "ค้นหา/แก้ไข เอกสารเดิม":
            doc_no = st.text_input("กรอก DOCUMENT NO. ที่ต้องการดึงข้อมูล:", key="search_doc_main").strip().upper()
            if doc_no:
                doc_data = get_document_data(doc_no)
                if doc_data:
                    st.success(f"✅ พบข้อมูลเอกสารเลขที่ {doc_no}")
                else:
                    st.error("❌ ไม่พบข้อมูลเอกสารในระบบ")

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

        c1, c2, c3 = st.columns(3)
        with c1:
            doc_no_val = st.text_input("DOCUMENT NO.", value=doc_data.get("DOCUMENT_NO", doc_no))
            customer_name = st.text_input("CUSTOMER NAME", value=doc_data.get("CUSTOMER_NAME", ""))
            part_name = st.text_input("PART NAME", value=doc_data.get("PART_NAME", ""))
        with c2:
            part_no = st.text_input("PART NO.", value=doc_data.get("PART_NO", ""))
            model = st.text_input("MODEL", value=doc_data.get("MODEL", ""))
            master_dwg_no = st.text_input("MASTER DWG NO.", value=doc_data.get("MASTER_DWG_NO", ""))
        with c3:
            issue_date = st.date_input("DATE", value=date.today())
            ref_doc_no = st.text_input("REF. DOC. NO.", value=doc_data.get("REF_DOC_NO", ""))
            issue_by = st.text_input("ISSUE BY", value=doc_data.get("ISSUE_BY", st.session_state.user_name))

        # =========================================================
        # 📌 SUBJECT TEXT (พื้นที่ D12:Q14) & IMAGE UPLOAD (พื้นที่ R12:AA14)
        # =========================================================
        st.markdown("---")
        st.subheader("📝 รายละเอียด Subject และ รูปภาพแนบ (Subject & Attached Image)")
        
        col_subj, col_img = st.columns([1, 1])
        
        with col_subj:
            st.markdown("**1. SUBJECT (รายละเอียดเรื่องที่เปลี่ยนแปลง - พื้นที่ D12:Q14)**")
            subject_text = st.text_area(
                "กรอกเนื้อหา Subject / รายละเอียดการเปลี่ยนแปลง:",
                value=doc_data.get("SUBJECT_TEXT", ""),
                height=150,
                placeholder="ระบุข้อความรายละเอียดการเปลี่ยนแปลงที่นี่..."
            )

        with col_img:
            st.markdown("**2. ATTACHED IMAGE (รูปภาพประกอบ - พื้นที่ R12:AA14)**")
            uploaded_image = st.file_uploader("อัปโหลดรูปภาพแนบ (JPG / PNG):", type=["jpg", "jpeg", "png"])
            
            image_base64_str = doc_data.get("IMAGE_BASE64", "")
            if uploaded_image is not None:
                image_base64_str = convert_image_to_base64(uploaded_image)
                st.image(uploaded_image, caption="รูปภาพที่อัปโหลดใหม่", use_container_width=True)
            elif image_base64_str:
                st.image(image_base64_str, caption="รูปภาพปัจจุบันในระบบ", use_container_width=True)

        st.markdown("---")
        st.subheader("📋 รายการตรวจสอบและแผนการดำเนินงาน (Checklist 19 รายการ)")

        # วนลูปสร้าง UI Checklist สำหรับทั้ง 19 รายการ
        checklist_results = {}
        for i in range(1, 20):
            dept, title = ITEM_DEPT_MAPPING[i]
            st.markdown(f"**ข้อ {i}. [{dept}] {title}**")
            
            curr_rev = get_doc_value(doc_data, i, "REVISE").upper()
            curr_resp = get_doc_value(doc_data, i, "RESP")
            curr_plan = get_doc_value(doc_data, i, "PLAN")
            curr_close = get_doc_value(doc_data, i, "CLOSE")

            col_a, col_b, col_c, col_d = st.columns([1.5, 2, 2, 2])
            with col_a:
                rev_choice = st.radio(f"แก้ไขข้อ {i}", ["-", "YES", "NO"], index=1 if curr_rev=="YES" else (2 if curr_rev=="NO" else 0), key=f"rev_{i}", horizontal=True)
            with col_b:
                resp_choice = st.text_input(f"ผู้รับผิดชอบ ข้อ {i}", value=curr_resp, key=f"resp_{i}")
            with col_c:
                plan_choice = st.text_input(f"Plan Date ข้อ {i}", value=curr_plan, key=f"plan_{i}")
            with col_d:
                close_choice = st.text_input(f"Actual Close ข้อ {i}", value=curr_close, key=f"close_{i}")

            checklist_results[f"DOC_{i}_REVISE"] = rev_choice
            checklist_results[f"DOC_{i}_RESP"] = resp_choice
            checklist_results[f"DOC_{i}_PLAN"] = plan_choice
            checklist_results[f"DOC_{i}_CLOSE"] = close_choice

        # =========================================================
        # 🖊️ ส่วนการลงนามอนุมัติ (MANAGER APPROVAL)
        # =========================================================
        st.markdown("---")
        st.subheader("🖊️ การลงนามอนุมัติเอกสาร (Manager Approval)")

        appr_pdd = doc_data.get("APPR_PDD_MGR", "")
        appr_qcd = doc_data.get("APPR_QCD_MGR", "")
        appr_prd = doc_data.get("APPR_PRD_MGR", "")
        appr_pcd = doc_data.get("APPR_PCD_MGR", "")
        appr_gm  = doc_data.get("APPR_GM", "")

        is_completed, missing_reasons = check_yes_items_completed(doc_data)

        if "MGR" in selected_dept or "GM" in selected_dept:
            if not is_completed:
                st.warning("⚠️ ยังไม่สามารถลงนามอนุมัติได้ เนื่องจากรายการ YES ยังปิดงานไม่ครบถ้วน:")
                for m in missing_reasons:
                    st.write(f"- {m}")
            else:
                st.success("✅ ปิดงานข้อ YES ครบถ้วนแล้ว สามารถลงนามอนุมัติได้")

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
                    send_approval_next_step_email(doc_no_val, customer_name, part_name, "ผู้จัดการ QC", "PRD_MGR", "ผู้จัดการ Production (PRD MGR)")
                    st.success("ลงนามอนุมัติสำเร็จ!")
                    st.rerun()

        with col_m3:
            st.markdown("**PRD MGR**")
            st.info(appr_prd if appr_prd else "รอลงนาม")
            if "MGR - PD" in selected_dept and is_completed and appr_qcd and not appr_prd:
                if st.button("🖊️ ลงนามอนุมัติ (PRD MGR)"):
                    appr_prd = st.session_state.user_name
                    date_prd = str(date.today())
                    save_to_excel({"DOCUMENT_NO": doc_no_val, "APPR_PRD_MGR": appr_prd, "DATE_PRD_MGR": date_prd})
                    send_approval_next_step_email(doc_no_val, customer_name, part_name, "ผู้จัดการ PRD", "PCD_MGR", "ผู้จัดการ PCD (PCD MGR)")
                    st.success("ลงนามอนุมัติสำเร็จ!")
                    st.rerun()

        with col_m4:
            st.markdown("**PCD MGR**")
            st.info(appr_pcd if appr_pcd else "รอลงนาม")
            if "MGR - PCD" in selected_dept and is_completed and appr_prd and not appr_pcd:
                if st.button("🖊️ ลงนามอนุมัติ (PCD MGR)"):
                    appr_pcd = st.session_state.user_name
                    date_pcd = str(date.today())
                    save_to_excel({"DOCUMENT_NO": doc_no_val, "APPR_PCD_MGR": appr_pcd, "DATE_PCD_MGR": date_pcd})
                    send_approval_next_step_email(doc_no_val, customer_name, part_name, "ผู้จัดการ PCD", "GM", "ผู้บริหาร (AGM / GM)")
                    st.success("ลงนามอนุมัติสำเร็จ!")
                    st.rerun()

        with col_m5:
            st.markdown("**AGM / GM**")
            st.info(appr_gm if appr_gm else "รอลงนาม")
            if "AGM / GM" in selected_dept and is_completed and appr_pcd and not appr_gm:
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
                    save_data = {
                        "DOCUMENT_NO": doc_no_val,
                        "CUSTOMER_NAME": customer_name,
                        "PART_NAME": part_name,
                        "PART_NO": part_no,
                        "MODEL": model,
                        "MASTER_DWG_NO": master_dwg_no,
                        "DATE": str(issue_date),
                        "REF_DOC_NO": ref_doc_no,
                        "ISSUE_BY": issue_by,
                        "SUBJECT_TEXT": subject_text,
                        "SUBJECT": subject_text,  # 📌 เพิ่ม Key นี้รองรับกรณีตาราง Google Sheet ใช้ชื่อ SUBJECT
                        "IMAGE_BASE64": image_base64_str
                     }
                    save_data.update(checklist_results)

                    if save_to_excel(save_data):
                        st.success(f"✅ บันทึกข้อมูลเอกสาร {doc_no_val} เรียบร้อยแล้ว!")
                        
                        # ตรวจสอบการส่ง Email แจ้งเตือนเมื่อวิศวกรปิดข้อ YES ครบ
                        check_data = get_document_data(doc_no_val)
                        completed, _ = check_yes_items_completed(check_data)
                        if completed and not check_data.get("APPR_PDD_MGR"):
                            send_all_completed_alert_email(doc_no_val, customer_name, part_name)
                            st.info("📧 ส่งอีเมลแจ้งเตือนไปยัง PDD Manager เรียบร้อยแล้ว")
        with col_b2:
            if doc_no_val:
                render_download_excel_button(doc_no_val, "📥 ดาวน์โหลด Excel ฟอร์มจริง")
