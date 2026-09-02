import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import openpyxl
import pandas as pd
import streamlit as st
from datetime import date
from PIL import Image
import gspread
from google.oauth2 import service_account

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

if not os.path.exists(UPLOAD_DIR):
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

def send_next_dept_alert_email(doc_no, customer, part_name, target_dept):
    to_email = DEPT_EMAILS.get(target_dept, SENDER_EMAIL)
    subject = f"🚨 [Action Required] ใบงาน Change Control {doc_no} ถึงคิวแผนก {target_dept}"
    body = (
        f"เรียน ทีมงานแผนก {target_dept},\n\n"
        f"ใบงานควบคุมการเปลี่ยนแปลง (Change Control) มีรายละเอียดดังนี้:\n"
        f"DOCUMENT NO.: {doc_no}\nCUSTOMER: {customer}\nPART NAME: {part_name}\n\n"
        f"แผนกก่อนหน้าได้บันทึกข้อมูลเรียบร้อยแล้ว รบกวนเข้าสู่ระบบเพื่อกรอกข้อมูลในส่วนของท่านครับ\n\n"
        f"🔗 เข้าสู่ระบบได้ที่: {APP_URL}\n\n"
        f"ขอแสดงความนับถือ,\nระบบ KFT Change Control Automated System"
    )
    return send_email_notification(to_email, subject, body)

def send_all_completed_alert_email(doc_no, customer, part_name):
    to_email = DEPT_EMAILS.get("PDD_MGR", SENDER_EMAIL)
    subject = f"✅ [Form Completed] ใบงาน {doc_no} กรอกข้อมูลและปิดงานครบถ้วนแล้ว (รอ PDD MGR อนุมัติ)"
    body = (
        f"เรียน ผู้จัดการ PDD (PDD MGR),\n\n"
        f"ใบงาน Change Control เลขที่ {doc_no} (Customer: {customer}, Part: {part_name}) "
        f"ได้รับการบันทึกข้อมูลและลงวันที่ปิดงานจริง (Actual Close) ครบถ้วนแล้ว\n\n"
        f"รบกวนผู้จัดการเข้าสู่ระบบเพื่อพิจารณาลงนามอนุมัติเอกสารเป็นลำดับแรกครับ\n\n"
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
        st.error(f"เกิดข้อผิดพลาดในการดึงข้อมูลทั้งหมด: {e}")
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

# =============================================================
# 🎯 ตรวจสอบการปิดงานตาม Plan (Actual Close Date Check)
# =============================================================
def check_all_departments_completed(doc_no):
    doc_data = get_document_data(doc_no)
    if not doc_data: return False
    for num in range(1, 20):
        rev_value = str(doc_data.get(f"DOC_{num}_REVISE", "NO")).upper()
        if rev_value == "YES":
            resp_value = str(doc_data.get(f"DOC_{num}_RESP", "")).strip()
            close_value = str(doc_data.get(f"DOC_{num}_CLOSE", "")).strip()
            # ต้องมีทั้งผู้รับผิดชอบ และ วันที่ปิดงานจริง
            if resp_value == "" or resp_value == "-" or close_value == "" or close_value == "-": 
                return False
    return True 

def get_missing_items(doc_no):
    doc_data = get_document_data(doc_no)
    if not doc_data: return []
    missing_list = []
    for num in range(1, 20):
        rev_value = str(doc_data.get(f"DOC_{num}_REVISE", "NO")).upper()
        if rev_value == "YES":
            resp_value = str(doc_data.get(f"DOC_{num}_RESP", "")).strip()
            close_value = str(doc_data.get(f"DOC_{num}_CLOSE", "")).strip()
            dept, title = ITEM_DEPT_MAPPING.get(num, ("-", "-"))
            
            if resp_value == "" or resp_value == "-":
                missing_list.append(f"ข้อ {num} [{dept}]: ยังไม่ได้ระบุผู้รับผิดชอบ ({title})")
            if close_value == "" or close_value == "-":
                missing_list.append(f"ข้อ {num} [{dept}]: ยังไม่ได้ลงวันที่ปิดเอกสารจริง (Actual Close) ({title})")
    return missing_list

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
                    ws.cell(row=merged_range.min_row, column=merged_range.min_col, value=value)
                    return
            target_cell.value = value
        
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
        
        write_cell("D12", doc_data.get("SUBJECT_TEXT", ""))
        
        write_cell("I12", "X" if doc_data.get("ATTACH_DRAWING") == "YES" else "")
        write_cell("I13", "X" if doc_data.get("ATTACH_ECI") == "YES" else "")
        write_cell("I14", "X" if doc_data.get("ATTACH_MEETING") == "YES" else "")
        write_cell("I15", f"X ({doc_data.get('ATTACH_OTHERS_DETAIL', '')})" if doc_data.get("ATTACH_OTHERS") == "YES" else "")

        judgement_val = doc_data.get("JUDGEMENT", "")
        write_cell("S13", "X" if judgement_val == "FEASIBLE" else "")
        write_cell("S14", "X" if judgement_val == "IMPROBABILITY" else "")

        start_row = 19
        for i in range(1, 20):
            current_row = start_row + (i - 1)
            rev_val = doc_data.get(f"DOC_{i}_REVISE", "NO")
            if rev_val == "YES":
                write_cell(f"K{current_row}", "X")
                write_cell(f"M{current_row}", "")
            else:
                write_cell(f"K{current_row}", "")
                write_cell(f"M{current_row}", "X")
                
            write_cell(f"O{current_row}", doc_data.get(f"DOC_{i}_RESP", ""))
            write_cell(f"U{current_row}", doc_data.get(f"DOC_{i}_PLAN", ""))
            write_cell(f"Y{current_row}", doc_data.get(f"DOC_{i}_CLOSE", ""))
            
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

# =============================================================
# 📋 Sidebar Navigation
# =============================================================
with st.sidebar:
    st.markdown(f"👤 **{st.session_state.user_name}**\n\n🏢 **บทบาท:** {st.session_state.current_dept}")
    st.markdown("---")
    
    # เมนูนำทาง
    menu = st.radio("📌 เมนูใช้งาน", ["📝 บันทึก/อนุมัติ เอกสาร", "📊 Dashboard ติดตามสถานะ"])
    
    st.markdown("---")
    if st.button("🚪 ออกจากระบบ", use_container_width=True):
        logout()

selected_dept = st.session_state.current_dept

# =============================================================
# 📊 VIEW 1: DASHBOARD ติดตามสถานะตามลำดับขั้นตอน
# =============================================================
if menu == "📊 Dashboard ติดตามสถานะ":
    st.title("📊 Dashboard ติดตามสถานะเอกสาร Change Control")
    st.write("ติดตามความคืบหน้าของใบงานตั้งแต่วิศวกรเปิดเอกสาร จนถึง GM อนุมัติขั้นสุดท้าย")
    
    df_all = get_all_documents()
    if df_all.empty:
        st.warning("⚠️ ยังไม่มีข้อมูลใบงานในระบบ")
    else:
        # การจัดหมวดหมู่สถานะ
        def get_detailed_status(row):
            status = str(row.get('DOC_STATUS', '')).upper()
            if status == "APPROVED":
                return "🟢 Approved (เสร็จสิ้นสมบูรณ์)"
            elif status == "WAIT_APPROVAL" or status == "FINISH":
                if not row.get('APPR_PDD_MGR'): return "🟡 รอ PDD MGR อนุมัติ"
                elif not row.get('APPR_QCD_MGR'): return "🟡 รอ QCD MGR อนุมัติ"
                elif not row.get('APPR_PRD_MGR'): return "🟡 รอ PRD MGR อนุมัติ"
                elif not row.get('APPR_PCD_MGR'): return "🟡 รอ PCD MGR อนุมัติ"
                elif not row.get('APPR_GM'): return "🟡 รอ AGM / GM อนุมัติ"
            else:
                return "🔵 กำลังดำเนินการ (รอปิดข้อ Check List)"
            return "🔵 PENDING"

        df_all['STEP_STATUS'] = df_all.apply(get_detailed_status, axis=1)

        # สรุปภาพรวม Metrics
        total_docs = len(df_all)
        approved_docs = len(df_all[df_all['STEP_STATUS'].str.contains("Approved")])
        pending_mgr = len(df_all[df_all['STEP_STATUS'].str.contains("รอ")])
        in_progress = total_docs - approved_docs - pending_mgr

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("📄 ใบงานทั้งหมด", total_docs)
        m2.metric("🔵 อยู่ระหว่างดำเนินการ (กำลังปิดข้อ)", in_progress)
        m3.metric("🟡 รอผู้จัดการอนุมัติ (Wait MGR)", pending_mgr)
        m4.metric("🟢 อนุมัติเสร็จสมบูรณ์ (Approved)", approved_docs)
        
        st.markdown("---")
        
        # ตัวกรองค้นหา
        col_search1, col_search2 = st.columns(2)
        with col_search1:
            search_text = st.text_input("🔍 ค้นหาตาม DOCUMENT NO / Customer / Part Name:")
        with col_search2:
            status_filter = st.selectbox("🎯 กรองตามสถานะขั้นตอน:", ["ทั้งหมด", "🔵 กำลังดำเนินการ", "🟡 รอผู้จัดการอนุมัติ", "🟢 Approved (เสร็จสิ้นสมบูรณ์)"])

        filtered_df = df_all.copy()
        if search_text:
            s = search_text.strip().upper()
            filtered_df = filtered_df[
                filtered_df['DOCUMENT_NO'].astype(str).str.upper().str.contains(s) |
                filtered_df['CUSTOMER_NAME'].astype(str).str.upper().str.contains(s) |
                filtered_df['PART_NAME'].astype(str).str.upper().str.contains(s)
            ]
        if status_filter != "ทั้งหมด":
            filtered_df = filtered_df[filtered_df['STEP_STATUS'].str.contains(status_filter[:2])]

        # แสดงข้อมูลในรูปแบบตาราง
        display_cols = ['DOCUMENT_NO', 'CUSTOMER_NAME', 'PART_NAME', 'MODEL', 'ISSUE_BY', 'DATE', 'STEP_STATUS']
        available_cols = [c for c in display_cols if c in filtered_df.columns]
        
        st.dataframe(
            filtered_df[available_cols].rename(columns={
                'DOCUMENT_NO': 'เลขที่เอกสาร',
                'CUSTOMER_NAME': 'ลูกค้า',
                'PART_NAME': 'ชื่อพาร์ท',
                'MODEL': 'โมเดล',
                'ISSUE_BY': 'ผู้จัดทำ',
                'DATE': 'วันที่สร้าง',
                'STEP_STATUS': 'สถานะขั้นตอนปัจจุบัน'
            }),
            use_container_width=True,
            hide_index=True
        )

# =============================================================
# 📝 VIEW 2: หน้าบันทึก / อนุมัติ เอกสาร
# =============================================================
else:
    st.title("KFT - RELATED DOCUMENT CHANGE CONTROL SYSTEM")

    if "Print Form" in selected_dept:
        st.subheader("🖨️ ระบบดึงและพิมพ์ฟอร์มเอกสารควบคุมอัตโนมัติ (Excel Format บริษัท)")
        print_doc_no = st.text_input("กรอก DOCUMENT NO. ที่ต้องการแปลงข้อมูลออกฟอร์ม (เช่น R001/26) :", key="print_doc_input").strip().upper()
        
        if print_doc_no:
            doc_data = get_document_data(print_doc_no)
            if doc_data:
                st.success(f"พบข้อมูลของใบงานหมายเลข {print_doc_no} ในระบบ")
                st.info(f"📋 สรุปงาน -> ลูกค้า: {doc_data.get('CUSTOMER_NAME', '-')} | พาร์ท: {doc_data.get('PART_NAME', '-')}")
                
                output_file, error_msg = export_to_printed_form(print_doc_no)
                if error_msg:
                    st.error(error_msg)
                elif output_file and os.path.exists(output_file):
                    with open(output_file, "rb") as f:
                        st.download_button(
                            label="📥 คลิกตรงนี้เพื่อดาวน์โหลดแบบฟอร์มจริงสำเร็จรูป (.xlsx)",
                            data=f,
                            file_name=os.path.basename(output_file),
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="primary",
                            use_container_width=True
                        )
                    st.balloons()
            else:
                st.error("❌ ไม่พบข้อมูลรหัสใบงานนี้ในฐานข้อมูลระบบ")

    elif "MGR" in selected_dept or "GM" in selected_dept:
        st.subheader("🔒 ระบบพิจารณาลงนามดิจิทัลและอนุมัติเอกสารควบคุม (Manager Sign-Off)")
        approve_doc_no = st.text_input("ระบุ DOCUMENT NO. ที่ต้องการพิจารณาอนุมัติ :", key="appr_doc_input").strip().upper()
        
        if approve_doc_no:
            doc_data = get_document_data(approve_doc_no)
            if doc_data:
                cust = doc_data.get('CUSTOMER_NAME', '-')
                part = doc_data.get('PART_NAME', '-')
                doc_status = doc_data.get('DOC_STATUS', 'PENDING')

                st.info(f"📋 รายละเอียดใบงาน -> ลูกค้า: {cust} | ชื่อพาร์ท: {part} | สถานะเอกสาร: {doc_status}")
                st.markdown("##### **📊 ตารางเช็กสถานะการเซ็นอนุมัติในขณะนี้:**")
                
                c_pdd, c_qcd, c_pd, c_pcd, c_gm = st.columns(5)
                c_pdd.metric("1. PDD MGR.", doc_data.get("APPR_PDD_MGR") if doc_data.get("APPR_PDD_MGR") else "⏳ รออนุมัติ", doc_data.get("DATE_PDD_MGR"))
                c_qcd.metric("2. QCD MGR.", doc_data.get("APPR_QCD_MGR") if doc_data.get("APPR_QCD_MGR") else "⏳ รออนุมัติ", doc_data.get("DATE_QCD_MGR"))
                c_pd.metric("3. PRD MGR.", doc_data.get("APPR_PRD_MGR") if doc_data.get("APPR_PRD_MGR") else "⏳ รออนุมัติ", doc_data.get("DATE_PRD_MGR"))
                c_pcd.metric("4. PCD MGR.", doc_data.get("APPR_PCD_MGR") if doc_data.get("APPR_PCD_MGR") else "⏳ รออนุมัติ", doc_data.get("DATE_PCD_MGR"))
                c_gm.metric("5. AGM / GM", doc_data.get("APPR_GM") if doc_data.get("APPR_GM") else "⏳ รออนุมัติ", doc_data.get("DATE_GM"))
                
                st.markdown("---")
                current_today = date.today().strftime('%Y-%m-%d')
                mgr_name = st.text_input("พิมพ์ชื่อ-นามสกุล ของคุณเพื่อใช้ยืนยันการอนุมัติ :", key="mgr_name_input").strip()
                
                # เช็กว่าแต่ละแผนกใส่ Actual Close Date ครบหรือยัง
                all_completed = check_all_departments_completed(approve_doc_no)

                if not all_completed:
                    st.warning("⚠️ เอกสารนี้ยังไม่สามารถอนุมัติได้ เนื่องจากพนักงานยังกรอกข้อมูลหรือลงวันที่ปิดงานจริง (Actual Close) ไม่ครบถ้วน")
                    missing_items = get_missing_items(approve_doc_no)
                    if missing_items:
                        with st.expander("🔍 **คลิกเพื่อดูรายการข้อและแผนกที่ยังค้างอยู่**", expanded=True):
                            for item in missing_items:
                                st.write(f"❌ {item}")

                elif "MGR - PDD" in selected_dept:
                    if st.button("🖊️ อนุมัติในฐานะ PDD MGR.", type="primary"):
                        if mgr_name == "": st.error("❌ กรุณาพิมพ์ชื่อตัวตนของคุณก่อน")
                        elif save_to_excel({"DOCUMENT_NO": approve_doc_no, "APPR_PDD_MGR": mgr_name, "DATE_PDD_MGR": current_today}):
                            st.success("✅ บันทึกการเซ็นอนุมัติเรียบร้อย!")
                            send_approval_next_step_email(approve_doc_no, cust, part, "ผู้จัดการ PDD (PDD MGR)", "QCD_MGR", "ผู้จัดการ QC (QCD MGR)")
                            st.rerun()

                elif "MGR - QCD" in selected_dept:
                    if not doc_data.get("APPR_PDD_MGR"): st.error("⚠️ ต้องรอให้ PDD MGR ลงนามอนุมัติก่อนครับ")
                    else:
                        if st.button("🖊️ อนุมัติในฐานะ QCD MGR.", type="primary"):
                            if mgr_name == "": st.error("❌ กรุณาพิมพ์ชื่อตัวตนของคุณก่อน")
                            elif save_to_excel({"DOCUMENT_NO": approve_doc_no, "APPR_QCD_MGR": mgr_name, "DATE_QCD_MGR": current_today}):
                                st.success("✅ บันทึกการเซ็นอนุมัติเรียบร้อย!")
                                send_approval_next_step_email(approve_doc_no, cust, part, "ผู้จัดการ QC (QCD MGR)", "PRD_MGR", "ผู้จัดการ Production (PRD MGR)")
                                st.rerun()

                elif "MGR - PD" in selected_dept or "MGR - PRD" in selected_dept:
                    if not doc_data.get("APPR_QCD_MGR"): st.error("⚠️ ต้องรอให้ QCD MGR ลงนามอนุมัติก่อนครับ")
                    else:
                        if st.button("🖊️ อนุมัติในฐานะ PRD MGR.", type="primary"):
                            if mgr_name == "": st.error("❌ กรุณาพิมพ์ชื่อตัวตนของคุณก่อน")
                            elif save_to_excel({"DOCUMENT_NO": approve_doc_no, "APPR_PRD_MGR": mgr_name, "DATE_PRD_MGR": current_today}):
                                st.success("✅ บันทึกการเซ็นอนุมัติเรียบร้อย!")
                                send_approval_next_step_email(approve_doc_no, cust, part, "ผู้จัดการ Production (PRD MGR)", "PCD_MGR", "ผู้จัดการ PCD (PCD MGR)")
                                st.rerun()

                elif "MGR - PCD" in selected_dept:
                    if not doc_data.get("APPR_PRD_MGR"): st.error("⚠️ ต้องรอให้ PRD MGR ลงนามอนุมัติก่อนครับ")
                    else:
                        if st.button("🖊️ อนุมัติในฐานะ PCD MGR.", type="primary"):
                            if mgr_name == "": st.error("❌ กรุณาพิมพ์ชื่อตัวตนของคุณก่อน")
                            elif save_to_excel({"DOCUMENT_NO": approve_doc_no, "APPR_PCD_MGR": mgr_name, "DATE_PCD_MGR": current_today}):
                                st.success("✅ บันทึกการเซ็นอนุมัติเรียบร้อย!")
                                send_approval_next_step_email(approve_doc_no, cust, part, "ผู้จัดการ PCD (PCD MGR)", "GM", "ผู้บริหาร (AGM / GM)")
                                st.rerun()

                elif "AGM / GM" in selected_dept:
                    if not doc_data.get("APPR_PCD_MGR"): st.error("⚠️ ต้องรอให้ MGR ทั้ง 4 แผนก ลงนามครบถ้วนก่อนครับ")
                    else:
                        if st.button("🏆 ยืนยันปิดงานขั้นสุดท้าย (AGM / GM APPROVAL)", type="primary"):
                            if mgr_name == "": st.error("❌ กรุณาพิมพ์ชื่อตัวตนของคุณก่อน")
                            else:
                                final_data = {
                                    "DOCUMENT_NO": approve_doc_no, "APPR_GM": mgr_name, 
                                    "DATE_GM": current_today, "DOC_STATUS": "APPROVED", "EFF_ACTUAL": current_today
                                }
                                if save_to_excel(final_data): 
                                    st.success("🎉 ใบงานนี้ผ่านการอนุมัติสมบูรณ์แล้ว!")
                                    send_final_approved_email(approve_doc_no, cust, part, mgr_name)
                                    st.rerun()
            else:
                st.error("❌ ไม่พบข้อมูลรหัสเอกสารควบคุมนี้")

    else:
        st.subheader("📝 ส่วนที่ 1: รายละเอียดข้อมูลโครงสร้างวิศวกรรมทั่วไป (Header)")
        
        # 1. รับค่า DOCUMENT NO.
        doc_no = st.text_input("📝 DOCUMENT NO. *จำเป็นต้องระบุ :", key="search_doc_no").strip().upper()
        
        # 2. ค้นหาและโหลดข้อมูลใส่ Session State เมื่อมีการป้อนเลขเอกสาร
        if doc_no:
            existing_data = get_document_data(doc_no)
            if existing_data:
                if "PDD" not in selected_dept:
                    st.info(f"✨ ค้นพบใบงานเลขที่ {doc_no} ในระบบ")
                
                st.session_state["loaded_cust"] = existing_data.get("CUSTOMER_NAME", "-- เลือกชื่อลูกค้า --")
                st.session_state["loaded_pname"] = existing_data.get("PART_NAME", "")
                st.session_state["loaded_pno"] = existing_data.get("PART_NO", "")
                st.session_state["loaded_model"] = existing_data.get("MODEL", "")
                st.session_state["loaded_mdwg"] = existing_data.get("MASTER_DWG_NO", "")
                st.session_state["loaded_issue"] = existing_data.get("ISSUE_BY", st.session_state.user_name)
                st.session_state["loaded_reftype"] = existing_data.get("REF_DOC_TYPE", "CUSTOMER ECI No.")
                st.session_state["loaded_refno"] = existing_data.get("REF_DOC_NO", "")
                st.session_state["loaded_event"] = existing_data.get("EFF_EVENT", "")
                st.session_state["loaded_plan"] = existing_data.get("EFF_PLAN", "")
                st.session_state["loaded_att_dwg"] = (existing_data.get("ATTACH_DRAWING") == "YES")
                st.session_state["loaded_att_eci"] = (existing_data.get("ATTACH_ECI") == "YES")
                st.session_state["loaded_att_meet"] = (existing_data.get("ATTACH_MEETING") == "YES")
                st.session_state["loaded_att_oth"] = (existing_data.get("ATTACH_OTHERS") == "YES")
                st.session_state["loaded_att_oth_det"] = existing_data.get("ATTACH_OTHERS_DETAIL", "")
                st.session_state["loaded_judge"] = existing_data.get("JUDGEMENT", "FEASIBLE")
                st.session_state["loaded_subject"] = existing_data.get("SUBJECT_TEXT", "")
                
                for num in range(1, 20):
                    st.session_state[f"loaded_doc_{num}_rev"] = existing_data.get(f"DOC_{num}_REVISE", "NO")
                    st.session_state[f"loaded_doc_{num}_resp"] = existing_data.get(f"DOC_{num}_RESP", "")
                    st.session_state[f"loaded_doc_{num}_plan"] = existing_data.get(f"DOC_{num}_PLAN", "")
                    st.session_state[f"loaded_doc_{num}_close"] = existing_data.get(f"DOC_{num}_CLOSE", "")
            else:
                st.session_state["loaded_cust"] = "-- เลือกชื่อลูกค้า --"
                st.session_state["loaded_pname"] = ""
                st.session_state["loaded_pno"] = ""
                st.session_state["loaded_model"] = ""
                st.session_state["loaded_mdwg"] = ""
                st.session_state["loaded_issue"] = st.session_state.user_name if "PDD" in selected_dept else ""
                st.session_state["loaded_reftype"] = "CUSTOMER ECI No."
                st.session_state["loaded_refno"] = ""
                st.session_state["loaded_event"] = ""
                st.session_state["loaded_plan"] = ""
                st.session_state["loaded_att_dwg"] = False
                st.session_state["loaded_att_eci"] = False
                st.session_state["loaded_att_meet"] = False
                st.session_state["loaded_att_oth"] = False
                st.session_state["loaded_att_oth_det"] = ""
                st.session_state["loaded_judge"] = "FEASIBLE"
                st.session_state["loaded_subject"] = ""
                
                for num in range(1, 20):
                    st.session_state[f"loaded_doc_{num}_rev"] = "NO"
                    st.session_state[f"loaded_doc_{num}_resp"] = ""
                    st.session_state[f"loaded_doc_{num}_plan"] = ""
                    st.session_state[f"loaded_doc_{num}_close"] = ""

        is_disabled = False if "PDD" in selected_dept else True

        col_main_left, col_main_right = st.columns(2)
        with col_main_left:
            customer_list = ["-- เลือกชื่อลูกค้า --", "HONDA", "NISSAN", "TMA", "ADIENT", "OTHER"]
            saved_cust = st.session_state.get("loaded_cust", "-- เลือกชื่อลูกค้า --")
            default_cust_index = customer_list.index(saved_cust) if saved_cust in customer_list else 0
            customer_name = st.selectbox("👤 CUSTOMER NAME :", customer_list, index=default_cust_index, disabled=is_disabled)
            part_name = st.text_input("PART NAME :", value=st.session_state.get("loaded_pname", ""), disabled=is_disabled)
            part_no = st.text_input("PART NO. :", value=st.session_state.get("loaded_pno", ""), disabled=is_disabled)
        with col_main_right:
            model_name = st.text_input("MODEL :", value=st.session_state.get("loaded_model", ""), disabled=is_disabled)
            master_dwg = st.text_input("MASTER DWG. NO. :", value=st.session_state.get("loaded_mdwg", ""), disabled=is_disabled)
            issue_by = st.text_input("✍️ ISSUE BY (ผู้จัดทำเอกสาร) :", value=st.session_state.get("loaded_issue", ""), disabled=is_disabled)

        st.markdown("---")
        left_col, right_col = st.columns(2)
        with left_col:
            saved_ref_type = st.session_state.get("loaded_reftype", "CUSTOMER ECI No.")
            ref_types = ["CUSTOMER ECI No.", "DESIGN NOTE. No.", "PROCESS CHANGE No."]
            default_ref_index = ref_types.index(saved_ref_type) if saved_ref_type in ref_types else 0
            ref_doc_type = st.radio("เลือกประเภทเอกสารแจ้งแก้แบบ :", ref_types, index=default_ref_index, horizontal=True, disabled=is_disabled)
            ref_doc_no = st.text_input(f"กรอกเลขที่เอกสาร ({ref_doc_type}) :", value=st.session_state.get("loaded_refno", ""), disabled=is_disabled).strip().upper()
        with right_col:
            eff_event = st.text_input("EVENT (เงื่อนไขการเริ่มมีผล) :", value=st.session_state.get("loaded_event", ""), disabled=is_disabled)
            saved_plan_date = st.session_state.get("loaded_plan", "")
            try: default_plan = date.fromisoformat(saved_plan_date) if saved_plan_date else date.today()
            except ValueError: default_plan = date.today()
            eff_plan = st.date_input("PLAN (วันที่เริ่มแผนงาน) :", value=default_plan, disabled=is_disabled)

        st.markdown("---")
        col_attach, col_judge = st.columns(2)
        with col_attach:
            st.markdown("##### 📎 **ATTACH CUSTOMER'S**")
            att_dwg = st.checkbox("DRAWING", value=st.session_state.get("loaded_att_dwg", False), disabled=is_disabled)
            att_eci = st.checkbox("ECI or Design Note.", value=st.session_state.get("loaded_att_eci", False), disabled=is_disabled)
            att_meeting = st.checkbox("MEETING MINUTE", value=st.session_state.get("loaded_att_meet", False), disabled=is_disabled)
            att_others = st.checkbox("OTHERS", value=st.session_state.get("loaded_att_oth", False), disabled=is_disabled)
            att_others_detail = st.text_input("ระบุ OTHERS (ถ้ามี) :", value=st.session_state.get("loaded_att_oth_det", ""), disabled=is_disabled or not att_others)

        with col_judge:
            st.markdown("##### ⚖️ **JUDGEMENT (by PDD)**")
            saved_judge = st.session_state.get("loaded_judge", "FEASIBLE")
            judge_options = ["FEASIBLE", "IMPROBABILITY"]
            judge_idx = judge_options.index(saved_judge) if saved_judge in judge_options else 0
            judgement = st.radio("ผลการประเมินโดย PDD :", judge_options, index=judge_idx, disabled=is_disabled)

        st.markdown("---")
        subject_text = st.text_area("SUBJECT (บันทึกเนื้อหารายรายละเอียดสาเหตุการแก้ไขแบบวิศวกรรม):", value=st.session_state.get("loaded_subject", ""), disabled=is_disabled)

        # =============================================================
        # 📷 ส่วนการอัปโหลดและแสดงผลรูปภาพ
        # =============================================================
        uploaded_file = st.file_uploader("📷 อัปโหลดรูปภาพพิมพ์เขียวประกอบหัวข้อ SUBJECT (ถ้ามี):", type=["jpg", "png", "jpeg"])
        
        if uploaded_file is not None:
            safe_doc_name = doc_no.replace("/", "_").replace("\\", "_") if doc_no else "temp"
            image_filename = f"{safe_doc_name}_{uploaded_file.name}"
            image_path = os.path.join(UPLOAD_DIR, image_filename)
            with open(image_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            st.image(image_path, caption="รูปภาพประกอบรายการ", use_container_width=True)

        # =============================================================
        # 📋 รายการตรวจสอบ 19 ข้อ + ปิดงานตาม PLAN (CLOSE DATE)
        # =============================================================
        st.markdown("---")
        st.subheader("📋 รายการเอกสารที่ต้องแก้ไข ลงนาม และปิดเอกสารตาม Plan (Checklist 19 ข้อ)")
        
        form_data = {}
        for num, (dept, title) in ITEM_DEPT_MAPPING.items():
            st.markdown(f"**ข้อ {num}. [{dept}] {title}**")
            c1, c2, c3, c4 = st.columns([1, 2, 2, 2])
            
            is_item_disabled = False if (dept in selected_dept or ("PRD" in selected_dept and dept == "PRO")) else True
            
            rev_saved = st.session_state.get(f"loaded_doc_{num}_rev", "NO")
            resp_saved = st.session_state.get(f"loaded_doc_{num}_resp", "")
            plan_saved = st.session_state.get(f"loaded_doc_{num}_plan", "")
            close_saved = st.session_state.get(f"loaded_doc_{num}_close", "")

            with c1:
                rev_val = st.radio(f"แก้ไข? ({num})", ["NO", "YES"], index=1 if rev_saved == "YES" else 0, horizontal=True, disabled=is_item_disabled, key=f"ui_rev_{num}")
                form_data[f"DOC_{num}_REVISE"] = rev_val

            with c2:
                resp_val = st.text_input(f"ผู้รับผิดชอบ ({num})", value=resp_saved, disabled=is_item_disabled, key=f"ui_resp_{num}")
                form_data[f"DOC_{num}_RESP"] = resp_val

            with c3:
                plan_val = st.text_input(f"กำหนดเสร็จ PLAN ({num})", value=plan_saved, disabled=is_item_disabled, key=f"ui_plan_{num}")
                form_data[f"DOC_{num}_PLAN"] = plan_val

            with c4:
                # เพิ่มช่อง Actual Close Date สำหรับการปิดงานแต่ละข้อ
                close_val = st.text_input(f"วันที่ปิดเอกสารจริง ACTUAL CLOSE ({num})", value=close_saved, disabled=is_item_disabled, key=f"ui_close_{num}")
                form_data[f"DOC_{num}_CLOSE"] = close_val

            st.markdown("<hr style='margin:5px 0;'>", unsafe_allow_html=True)

        # =============================================================
        # 💾 บันทึกข้อมูล & ส่ง Email อัตโนมัติตามลำดับขั้นตอน
        # =============================================================
        st.markdown("---")
        if st.button("💾 บันทึกข้อมูล (Save Data)", type="primary", use_container_width=True):
            if not doc_no:
                st.error("❌ กรุณาระบุ DOCUMENT NO. ก่อนทำการบันทึกข้อมูล")
            else:
                save_payload = {
                    "DOCUMENT_NO": doc_no,
                    "CUSTOMER_NAME": customer_name,
                    "PART_NAME": part_name,
                    "PART_NO": part_no,
                    "MODEL": model_name,
                    "MASTER_DWG_NO": master_dwg,
                    "ISSUE_BY": issue_by,
                    "REF_DOC_TYPE": ref_doc_type,
                    "REF_DOC_NO": ref_doc_no,
                    "EFF_EVENT": eff_event,
                    "EFF_PLAN": eff_plan.strftime('%Y-%m-%d') if isinstance(eff_plan, date) else str(eff_plan),
                    "ATTACH_DRAWING": "YES" if att_dwg else "NO",
                    "ATTACH_ECI": "YES" if att_eci else "NO",
                    "ATTACH_MEETING": "YES" if att_meeting else "NO",
                    "ATTACH_OTHERS": "YES" if att_others else "NO",
                    "ATTACH_OTHERS_DETAIL": att_others_detail,
                    "JUDGEMENT": judgement,
                    "SUBJECT_TEXT": subject_text,
                    "DATE": date.today().strftime('%Y-%m-%d'),
                    **form_data
                }
                
                if save_to_excel(save_payload):
                    st.success("✅ บันทึกข้อมูลเข้าสู่ฐานข้อมูล Google Sheets เรียบร้อยแล้ว!")
                    
                    # 📩 แจ้งเตือนส่งต่อไปยังแผนก QC
                    if "PDD" in selected_dept:
                        send_next_dept_alert_email(doc_no, customer_name, part_name, target_dept="QC")

                    # 🎯 เช็กเงื่อนไขการส่งต่อหา Manager (ต้องกรอก Actual Close Date ครบทุกข้อ)
                    if check_all_departments_completed(doc_no):
                        save_to_excel({"DOCUMENT_NO": doc_no, "DOC_STATUS": "WAIT_APPROVAL"})
                        send_all_completed_alert_email(doc_no, customer_name, part_name)
                        st.info("🎉 กรอกข้อมูลและปิดเอกสารจริงครบถ้วนแล้ว! ส่งอีเมลแจ้งเตือน PDD MGR เพื่อพิจารณาอนุมัติเรียบร้อย")
                    else:
                        save_to_excel({"DOCUMENT_NO": doc_no, "DOC_STATUS": "IN_PROGRESS"})
                        st.warning("⚠️ ข้อมูลถูกบันทึกแล้ว แต่สถานะยังเป็น 'IN_PROGRESS' เนื่องจากยังมีบางรายการไม่ได้ลงวันที่ปิดงานจริง (Actual Close)")

                    st.rerun()
