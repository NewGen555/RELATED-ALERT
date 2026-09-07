import os
import io
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

# =============================================================
# 🎨 CSS STYLING
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
# 👤 USERS & LOGIN MANAGEMENT
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

# =============================================================
# 🔍 REALTIME LOCATION CALCULATOR
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
# 🖨️ EXPORT TO EXCEL TEMPLATE FORM
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
        
        subject_val = doc_data.get("SUBJECT_TEXT") or doc_data.get("SUBJECT", "")
        write_cell("D12", subject_val)
        
        img_path = doc_data.get("IMAGE_PATH") or doc_data.get("IMAGE", "")
        if img_path and os.path.exists(img_path):
            try:
                from openpyxl.drawing.image import Image as OpenpyxlImage
                img = OpenpyxlImage(img_path)
                img.width = 250
                img.height = 80
                ws.add_image(img, "R12")
            except Exception as img_err:
                print(f"⚠️ ไม่สามารถแทรกรูปภาพได้: {img_err}")
        
        write_cell("I12", "✓" if doc_data.get("ATTACH_DRAWING") == "YES" else "")
        write_cell("I13", "✓" if doc_data.get("ATTACH_ECI") == "YES" else "")
        write_cell("I14", "✓" if doc_data.get("ATTACH_MEETING") == "YES" else "")
        write_cell("I15", f"✓ ({doc_data.get('ATTACH_OTHERS_DETAIL', '')})" if doc_data.get("ATTACH_OTHERS") == "YES" else "")

        judgement_val = doc_data.get("JUDGEMENT", "")
        write_cell("S13", "✓" if judgement_val == "FEASIBLE" else "")
        write_cell("S14", "✓" if judgement_val == "IMPROBABILITY" else "")

        start_row = 19
        for i in range(1, 20):
            current_row = start_row + (i - 1)
            rev_val = str(get_doc_value(doc_data, i, "REVISE")).upper().strip()
            
            if rev_val == "YES":
                write_cell(f"K{current_row}", "✓")
                write_cell(f"M{current_row}", "")
                write_cell(f"O{current_row}", get_doc_value(doc_data, i, "RESP"))
                write_cell(f"U{current_row}", get_doc_value(doc_data, i, "PLAN"))
                write_cell(f"Y{current_row}", get_doc_value(doc_data, i, "CLOSE"))
            elif rev_val == "NO":
                write_cell(f"K{current_row}", "")
                write_cell(f"M{current_row}", "✓")
                write_cell(f"O{current_row}", "")
                write_cell(f"U{current_row}", "")
                write_cell(f"Y{current_row}", "")
            else:
                write_cell(f"K{current_row}", "")
                write_cell(f"M{current_row}", "")
                write_cell(f"O{current_row}", "")
                write_cell(f"U{current_row}", "")
                write_cell(f"Y{current_row}", "")
                
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
# 📋 SIDEBAR
# =============================================================
with st.sidebar:
    st.markdown(f"👤 **{st.session_state.user_name}**\n\n🏢 **บทบาท:** {st.session_state.current_dept}")
    st.markdown("---")
    menu = st.radio("📌 เมนูใช้งาน", ["📊 Dashboard ติดตามสถานะ Realtime", "📝 บันทึก/อนุมัติ เอกสาร"])
    st.markdown("---")
    if st.button("🚪 ออกจากระบบ", use_container_width=True):
        logout()

# =============================================================
# 📊 VIEW 1: DASHBOARD
# =============================================================
if menu == "📊 Dashboard ติดตามสถานะ Realtime":
    col_head1, col_head2 = st.columns([3, 1])
    with col_head1:
        st.title("📊 Realtime Tracking & Overview Dashboard")
        st.caption("ระบบติดตามตำแหน่งเอกสารแบบเรียลไทม์ และสรุปภาพรวมเอกสารทั้งหมดในระบบ")
    with col_head2:
        if st.button("🔄 รีเฟรชข้อมูลล่าสุด", use_container_width=True):
            st.rerun()

    df_all = get_all_documents()
    if df_all.empty:
        st.warning("⚠️ ยังไม่มีข้อมูลใบงานในระบบ")
    else:
        status_info = df_all.apply(get_realtime_location, axis=1)
        df_all['MAIN_STATUS'] = [s[0] for s in status_info]
        df_all['CURRENT_LOCATION'] = [s[1] for s in status_info]

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

                st.markdown("##### 📥 **ดาวน์โหลดเอกสารฟอร์มจริง:**")
                render_download_excel_button(doc_search_input)
            else:
                st.error(f"❌ ไม่พบเอกสารเลขที่ '{doc_search_input}' ในฐานข้อมูล")
        
        st.markdown("---")
        st.markdown("### 📈 ภาพรวมเอกสารทั้งหมดในระบบ (System Overview)")
        
        total_docs = len(df_all)
        approved_docs = len(df_all[df_all['MAIN_STATUS'].str.contains("Approved")])
        pending_mgr = len(df_all[df_all['MAIN_STATUS'].str.contains("รอการอนุมัติ")])
        in_progress = total_docs - approved_docs - pending_mgr

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("📄 ใบงานทั้งหมด", f"{total_docs} รายการ")
        m2.metric("🔵 รอแผนกปิดข้อ YES", f"{in_progress} รายการ")
        m3.metric("🟡 รอผู้จัดการอนุมัติ", f"{pending_mgr} รายการ")
        m4.metric("🟢 อนุมัติเสร็จสมบูรณ์", f"{approved_docs} รายการ")

        st.markdown("<br>", unsafe_allow_html=True)
        st.dataframe(df_all[['DOCUMENT_NO', 'CUSTOMER_NAME', 'PART_NAME', 'MODEL', 'ISSUE_BY', 'MAIN_STATUS', 'CURRENT_LOCATION']], use_container_width=True, hide_index=True)

# =============================================================
# 📝 VIEW 2: FULL FORM ENTRY & APPROVAL (ส่วนกรอกข้อมูลที่สมบูรณ์)
# =============================================================
elif menu == "📝 บันทึก/อนุมัติ เอกสาร":
    st.title("📝 บันทึกข้อมูล และลงนามอนุมัติเอกสาร")
    st.caption(f"ผู้ใช้งานปัจจุบัน: {st.session_state.user_name} ({st.session_state.current_dept})")

    df_all_docs = get_all_documents()
    doc_list = []
    if not df_all_docs.empty and 'DOCUMENT_NO' in df_all_docs.columns:
        doc_list = df_all_docs['DOCUMENT_NO'].dropna().astype(str).str.strip().unique().tolist()

    selected_doc = st.selectbox("📌 เลือกเลขที่เอกสารเพื่อดึงข้อมูลขึ้นมาแก้ไข/อนุมัติ (หรือเลือก 'สร้างเอกสารใหม่'):", ["-- สร้างเอกสารใหม่ --"] + doc_list)

    doc_data = {}
    if selected_doc != "-- สร้างเอกสารใหม่ --":
        doc_data = get_document_data(selected_doc) or {}
        st.info(f"📂 กำลังแก้ไข/อนุมัติ เอกสารเลขที่: **{selected_doc}**")
        render_download_excel_button(selected_doc, f"📥 ดาวน์โหลดฟอร์มจริง Excel ({selected_doc})")
    else:
        st.success("✨ กำลังสร้างเอกสาร Change Control ฉบับใหม่")

    st.markdown("---")

    # ---------------------------------------------------------
    # 1. ข้อมูลทั่วไปของเอกสาร (GENERAL INFORMATION)
    # ---------------------------------------------------------
    st.subheader("1. ข้อมูลทั่วไปของเอกสาร (General Information)")
    col1, col2, col3 = st.columns(3)
    with col1:
        doc_no_val = st.text_input("DOCUMENT NO.", value=doc_data.get("DOCUMENT_NO", ""))
        customer_val = st.text_input("CUSTOMER NAME", value=doc_data.get("CUSTOMER_NAME", ""))
        part_name_val = st.text_input("PART NAME", value=doc_data.get("PART_NAME", ""))
        part_no_val = st.text_input("PART NO.", value=doc_data.get("PART_NO", ""))
    with col2:
        model_val = st.text_input("MODEL", value=doc_data.get("MODEL", ""))
        master_dwg_val = st.text_input("MASTER DRAWING NO.", value=doc_data.get("MASTER_DWG_NO", ""))
        ref_doc_val = st.text_input("REF. DOC NO.", value=doc_data.get("REF_DOC_NO", ""))
    with col3:
        issue_by_val = st.text_input("ISSUE BY", value=doc_data.get("ISSUE_BY", st.session_state.user_name))
        date_val = st.text_input("DATE", value=doc_data.get("DATE", str(date.today())))

    st.markdown("---")

    # ---------------------------------------------------------
    # 2. เหตุผลและรายละเอียดการเปลี่ยนแปลง (SUBJECT / REASON)
    # ---------------------------------------------------------
    st.subheader("2. เหตุผลและรายละเอียดการเปลี่ยนแปลง (Subject / Reason)")
    subject_val = st.text_area("SUBJECT / DETAILS", value=doc_data.get("SUBJECT_TEXT") or doc_data.get("SUBJECT", ""), height=100)
    
    uploaded_img = st.file_uploader("📷 อัปโหลดรูปภาพแนบประกอบ Subject / Details", type=["png", "jpg", "jpeg"])
    img_base64 = doc_data.get("IMAGE", "")
    if uploaded_img is not None:
        img_base64 = convert_image_to_base64(uploaded_img)
        st.success("✅ อัปโหลดรูปภาพเรียบร้อยแล้ว")
    elif img_base64:
        st.caption("🖼️ เอกสารนี้มีรูปภาพแนบประกอบอยู่แล้ว")

    st.markdown("---")

    # ---------------------------------------------------------
    # 3. เอกสารแนบ (ATTACHMENT)
    # ---------------------------------------------------------
    st.subheader("3. เอกสารแนบ (Attachment)")
    col_att1, col_att2 = st.columns(2)
    with col_att1:
        att_dwg = st.checkbox("DRAWING / SPEC", value=(doc_data.get("ATTACH_DRAWING") == "YES"))
        att_eci = st.checkbox("ECI / ECO", value=(doc_data.get("ATTACH_ECI") == "YES"))
    with col_att2:
        att_meeting = st.checkbox("MINUTES OF MEETING", value=(doc_data.get("ATTACH_MEETING") == "YES"))
        att_others = st.checkbox("OTHERS", value=(doc_data.get("ATTACH_OTHERS") == "YES"))
        att_others_detail = st.text_input("ระบุ OTHERS DETAIL:", value=doc_data.get("ATTACH_OTHERS_DETAIL", ""))

    st.markdown("---")

    # ---------------------------------------------------------
    # 4. กำหนดการมีผลบังคับใช้ (EFFECTIVE DATE)
    # ---------------------------------------------------------
    st.subheader("4. กำหนดการมีผลบังคับใช้ (Effective Event / Timing)")
    col_eff1, col_eff2, col_eff3 = st.columns(3)
    with col_eff1:
        eff_event = st.text_input("EFFECTIVE EVENT", value=doc_data.get("EFF_EVENT", ""))
    with col_eff2:
        eff_plan = st.text_input("PLAN DATE", value=doc_data.get("EFF_PLAN", ""))
    with col_eff3:
        eff_actual = st.text_input("ACTUAL DATE", value=doc_data.get("EFF_ACTUAL", ""))

    st.markdown("---")

    # ---------------------------------------------------------
    # 5. ตาราง CHECKLIST รายการแก้ไข 19 ข้อ
    # ---------------------------------------------------------
    st.subheader("5. รายการตรวจสอบเอกสารและการดำเนินงาน (Checklist 19 รายการ)")
    st.caption("เลือก REVISE (YES/NO) และกรอกผู้รับผิดชอบ, กำหนดวัน Plan และวัน Actual Close")

    checklist_results = {}
    for num in range(1, 20):
        dept, title = ITEM_DEPT_MAPPING.get(num, ("-", "-"))
        curr_rev = get_doc_value(doc_data, num, "REVISE").upper()
        curr_resp = get_doc_value(doc_data, num, "RESP")
        curr_plan = get_doc_value(doc_data, num, "PLAN")
        curr_close = get_doc_value(doc_data, num, "CLOSE")

        with st.expander(f"ข้อ {num}. [{dept}] {title}", expanded=(curr_rev == "YES")):
            c1, c2, c3, c4 = st.columns([1.5, 2, 2, 2])
            with c1:
                rev_choice = st.selectbox(f"REVISE #{num}", ["-", "YES", "NO"], index=["-", "YES", "NO"].index(curr_rev) if curr_rev in ["YES", "NO"] else 0, key=f"rev_{num}")
            with c2:
                resp_input = st.text_input(f"RESPONSIBLE #{num}", value=curr_resp, key=f"resp_{num}")
            with c3:
                plan_input = st.text_input(f"PLAN DATE #{num}", value=curr_plan, key=f"plan_{num}")
            with c4:
                close_input = st.text_input(f"ACTUAL CLOSE #{num}", value=curr_close, key=f"close_{num}")

            checklist_results[f"DOC_{num}_REVISE"] = rev_choice
            checklist_results[f"DOC_{num}_RESP"] = resp_input
            checklist_results[f"DOC_{num}_PLAN"] = plan_input
            checklist_results[f"DOC_{num}_CLOSE"] = close_input

    st.markdown("---")

    # ---------------------------------------------------------
    # 6. การลงนามอนุมัติ (MANAGER / GM APPROVAL)
    # ---------------------------------------------------------
    st.subheader("6. การลงนามอนุมัติเอกสาร (Management Approval)")
    
    col_app1, col_app2, col_app3, col_app4, col_app5 = st.columns(5)
    with col_app1:
        st.markdown("**PDD Manager**")
        appr_pdd = st.text_input("ชื่อผู้อนุมัติ (PDD)", value=doc_data.get("APPR_PDD_MGR", ""), key="appr_pdd")
        date_pdd = st.text_input("วันที่อนุมัติ (PDD)", value=doc_data.get("DATE_PDD_MGR", ""), key="date_pdd")
    with col_app2:
        st.markdown("**QC Manager**")
        appr_qcd = st.text_input("ชื่อผู้อนุมัติ (QC)", value=doc_data.get("APPR_QCD_MGR", ""), key="appr_qcd")
        date_qcd = st.text_input("วันที่อนุมัติ (QC)", value=doc_data.get("DATE_QCD_MGR", ""), key="date_qcd")
    with col_app3:
        st.markdown("**PCD Manager**")
        appr_pcd = st.text_input("ชื่อผู้อนุมัติ (PCD)", value=doc_data.get("APPR_PCD_MGR", ""), key="appr_pcd")
        date_pcd = st.text_input("วันที่อนุมัติ (PCD)", value=doc_data.get("DATE_PCD_MGR", ""), key="date_pcd")
    with col_app4:
        st.markdown("**PRD Manager**")
        appr_prd = st.text_input("ชื่อผู้อนุมัติ (PRD)", value=doc_data.get("APPR_PRD_MGR", ""), key="appr_prd")
        date_prd = st.text_input("วันที่อนุมัติ (PRD)", value=doc_data.get("DATE_PRD_MGR", ""), key="date_prd")
    with col_app5:
        st.markdown("**AGM / GM**")
        appr_gm = st.text_input("ชื่อผู้อนุมัติ (GM)", value=doc_data.get("APPR_GM", ""), key="appr_gm")
        date_gm = st.text_input("วันที่อนุมัติ (GM)", value=doc_data.get("DATE_GM", ""), key="date_gm")

    st.markdown("---")

    # ---------------------------------------------------------
    # ปุ่มบันทึกข้อมูลหลัก (SAVE BUTTON)
    # ---------------------------------------------------------
    if st.button("💾 บันทึกข้อมูลเอกสารลงฐานข้อมูล (Save All Data)", type="primary", use_container_width=True):
        if not doc_no_val.strip():
            st.error("❌ กรุณากำหนด DOCUMENT NO. ก่อนบันทึกข้อมูล")
        else:
            save_payload = {
                "DOCUMENT_NO": doc_no_val.strip().upper(),
                "CUSTOMER_NAME": customer_val,
                "PART_NAME": part_name_val,
                "PART_NO": part_no_val,
                "MODEL": model_val,
                "MASTER_DWG_NO": master_dwg_val,
                "REF_DOC_NO": ref_doc_val,
                "ISSUE_BY": issue_by_val,
                "DATE": date_val,
                "SUBJECT_TEXT": subject_val,
                "IMAGE": img_base64,
                "ATTACH_DRAWING": "YES" if att_dwg else "NO",
                "ATTACH_ECI": "YES" if att_eci else "NO",
                "ATTACH_MEETING": "YES" if att_meeting else "NO",
                "ATTACH_OTHERS": "YES" if att_others else "NO",
                "ATTACH_OTHERS_DETAIL": att_others_detail,
                "EFF_EVENT": eff_event,
                "EFF_PLAN": eff_plan,
                "EFF_ACTUAL": eff_actual,
                "APPR_PDD_MGR": appr_pdd,
                "DATE_PDD_MGR": date_pdd,
                "APPR_QCD_MGR": appr_qcd,
                "DATE_QCD_MGR": date_qcd,
                "APPR_PCD_MGR": appr_pcd,
                "DATE_PCD_MGR": date_pcd,
                "APPR_PRD_MGR": appr_prd,
                "DATE_PRD_MGR": date_prd,
                "APPR_GM": appr_gm,
                "DATE_GM": date_gm,
            }
            # รวมรายการ Checklist 19 ข้อ
            save_payload.update(checklist_results)

            if save_to_excel(save_payload):
                st.success(f"✅ บันทึกข้อมูลเอกสารเลขที่ {doc_no_val} เรียบร้อยแล้ว!")
                st.rerun()
