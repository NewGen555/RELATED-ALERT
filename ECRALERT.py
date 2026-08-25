import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import openpyxl
import pandas as pd
import streamlit as st
from datetime import date
from PIL import Image

# =============================================================
# ตั้งค่าหน้าเว็บ Streamlit
# =============================================================
st.set_page_config(
    layout="wide",
    page_title="KFT Change Control System",
    page_icon="🔐"
)

EXCEL_FILE = "change_control_db.xlsx"
TEMPLATE_FILE = "template_form.xlsx"
UPLOAD_DIR = "uploads"

if not os.path.exists(UPLOAD_DIR):
    os.makedirs(UPLOAD_DIR)

# =============================================================
# CONFIGURATION: SMTP EMAIL SETTINGS & DEPARTMENT EMAILS
# ดึงรหัสผ่านอีเมลจาก st.secrets (หากไม่มีให้ใช้ค่าสำรอง)
# =============================================================
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
SENDER_EMAIL = "pdd1development@gmail.com"
SENDER_PASSWORD = st.secrets.get("email", {}).get("sender_password", "awynohxlypvxuxfo")
APP_URL = "https://related-alert-erh2rywrtchautlthjlrwb.streamlit.app/"

DEPT_EMAILS = {
    "PDD": ["pdd_1@kftc.co.th", "saksiam@kftc.co.th", "manoc@kftc.co.th"],
    "QC": ["uchai@kftc.co.th", "sirirat@kftc.co.th", "pdd_1@kftc.co.th"],
    "PCD": ["pc-3@kftc.co.th", "pdd_1@kftc.co.th"],
    "PRD": ["suriya@kftc.co.th", "setthanan@kftc.co.th", "pd1center@kftc.co.th", "pdd_1@kftc.co.th"],
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

# =============================================================
# FUNCTION: ส่งแจ้งเตือนผ่าน Email (SMTP)
# =============================================================
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
    subject = f"✅ [Form Completed] ใบงาน {doc_no} กรอกข้อมูลครบถ้วนแล้ว (รอ PDD MGR อนุมัติ)"
    body = (
        f"เรียน ผู้จัดการ PDD (PDD MGR),\n\n"
        f"ใบงาน Change Control เลขที่ {doc_no} (Customer: {customer}, Part: {part_name}) "
        f"ได้รับการบันทึกข้อมูลรายการตรวจสอบครบทั้ง 19 ข้อจากทุกแผนกเรียบร้อยแล้ว\n\n"
        f"ลำดับแรก: รบกวนผู้จัดการเข้าสู่ระบบเพื่อพิจารณาอนุมัติเอกสารเป็นลำดับแรกครับ\n\n"
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
# 🎨 CSS ธีมสีฟ้าอ่อน + ขาว
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
# 👤 ระบบผู้ใช้งาน (ดึง Password จาก st.secrets)
# =============================================================
sec_passwords = st.secrets.get("passwords", {})

USERS = {
    "pdd_user": {"password": sec_passwords.get("pdd_user", "pdd1234"), "dept": "PDD (Product Design)", "name": "ENGINEER PDD"},
    "qc_user": {"password": sec_passwords.get("qc_user", "qc1234"), "dept": "QC (Quality Control)", "name": "ENGINEER QC"},
    "pcd_user": {"password": sec_passwords.get("pcd_user", "pcd1234"), "dept": "PCD (Production Control)", "name": "ENGINEER PCD"},
    "prd_user": {"password": sec_passwords.get("prd_user", "prd1234"), "dept": "PRO (Production / PD)", "name": "ENGINEER Production"},
    "mgr_pdd": {"password": sec_passwords.get("mgr_pdd", "mgrpdd1"), "dept": "MGR - PDD (ผู้จัดการ PDD)", "name": "ผู้จัดการ PDD"},
    "mgr_qcd": {"password": sec_passwords.get("mgr_qcd", "mgrqc1"), "dept": "MGR - QCD (ผู้จัดการ QC)", "name": "ผู้จัดการ QC"},
    "mgr_pcd": {"password": sec_passwords.get("mgr_pcd", "mgrpcd1"), "dept": "MGR - PCD (ผู้จัดการ PCD)", "name": "ผู้จัดการ PCD"},
    "mgr_prd": {"password": sec_passwords.get("mgr_prd", "mgrprd1"), "dept": "MGR - PD (ผู้จัดการ Production)", "name": "ผู้จัดการ PRD"},
    "gm_user": {"password": sec_passwords.get("gm_user", "gm1234"), "dept": "AGM / GM (ผู้บริหารอนุมัติขั้นสุดท้าย)", "name": "ผู้บริหาร GM"},
    "print_user": {"password": sec_passwords.get("print_user", "print1"), "dept": "Print Form", "name": "เจ้าหน้าที่พิมพ์เอกสาร"},
}

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.current_user = None
    st.session_state.current_dept = None
    st.session_state.user_name = None

if "form_reset_counter" not in st.session_state:
    st.session_state.form_reset_counter = 0

def login(username, password):
    user = USERS.get(username)
    if user and user["password"] == password:
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

def clear_all_inputs():
    st.session_state.form_reset_counter += 1
    st.rerun()

# =============================================================
# 🔐 หน้า LOGIN
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
# 📋 Sidebar
# =============================================================
with st.sidebar:
    st.markdown(f"👤 **{st.session_state.user_name}**\n\n🏢 **บทบาท:** {st.session_state.current_dept}")
    st.markdown("---")
    if st.button("🚪 ออกจากระบบ", use_container_width=True):
        logout()

selected_dept = st.session_state.current_dept

# =============================================================
# 🖨️ ฟังก์ชันเปิด Template Excel
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
# ส่วนประมวลผล Excel Database
# =============================================================
ITEM_DEPT_MAPPING = {
    1: ("PDD", "MASTER DRAWING."),
    2: ("PDD", "MATERIAL PART NO. LIST. , ACC DWG."),
    3: ("PDD", "PROCESS FLOW CHART."),
    4: ("PDD", "OPERATION MANUAL."),
    5: ("PDD", "TEST RESULT."),
    6: ("PDD", "FMEA"),
    7: ("PDD", "TOOLING No"),
    8: ("QC", "CONTROL PLAN."),
    9: ("QC", "INCOMING SHEET."),
    10: ("QC", "FINAL INSPECTION SHEET."),
    11: ("QC", "W/I Out Going / TRAINING QC."),
    12: ("QC", "INSPECTION STD. + DATA CHECK."),
    13: ("QC", "MSA"),
    14: ("QC", "PSW UP-DATE., PPAP APPROVAL."),
    15: ("QC", "CHECKING FIXTURE."),
    16: ("PCD", "MATERIAL REQUIREMENT."),
    17: ("PCD", "PACKING STANDARD."),
    18: ("PRO", "WORKING INSTRUCTION."),
    19: ("PRO", "TRAINING PRODUCTION.")
}

def ตรวจเช็คและสร้างไฟล์():
    if not os.path.exists(EXCEL_FILE):
        columns = [
            "DOCUMENT_NO", "CUSTOMER_NAME", "PART_NAME", "PART_NO", "MODEL", "MASTER_DWG_NO", "DATE", "ISSUE_BY",
            "REF_DOC_TYPE", "REF_DOC_NO", "EFF_EVENT", "EFF_PLAN", "EFF_ACTUAL", "SUBJECT_TEXT", "SUBJECT_IMAGE_PATH",
            "ATTACH_DRAWING", "ATTACH_ECI", "ATTACH_MEETING", "ATTACH_OTHERS", "ATTACH_OTHERS_DETAIL", "JUDGEMENT",
            "DOC_STATUS",
            "APPR_PDD_MGR", "DATE_PDD_MGR", "APPR_QCD_MGR", "DATE_QCD_MGR", "APPR_PCD_MGR", "DATE_PCD_MGR", "APPR_PRD_MGR", "DATE_PRD_MGR", "APPR_GM", "DATE_GM"
        ]
        for num in range(1, 20):
            columns.extend([f"DOC_{num}_REVISE", f"DOC_{num}_RESP", f"DOC_{num}_PLAN", f"DOC_{num}_CLOSE"])
        df = pd.DataFrame(columns=columns)
        df.to_excel(EXCEL_FILE, index=False)

ตรวจเช็คและสร้างไฟล์()

def get_document_data(doc_no):
    if os.path.exists(EXCEL_FILE):
        try:
            df = pd.read_excel(EXCEL_FILE)
            if 'DOCUMENT_NO' in df.columns and doc_no in df['DOCUMENT_NO'].values:
                row_data = df[df['DOCUMENT_NO'] == doc_no].iloc[0]
                return {k: ("" if pd.isna(v) else v) for k, v in row_data.to_dict().items()}
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาดในการอ่านไฟล์ฐานข้อมูล: {e}")
    return None

def save_to_excel(data_dict):
    try:
        if os.path.exists(EXCEL_FILE):
            df_old = pd.read_excel(EXCEL_FILE)
            df_new = pd.DataFrame([data_dict])
            if 'DOCUMENT_NO' in df_old.columns and data_dict['DOCUMENT_NO'] in df_old['DOCUMENT_NO'].values:
                old_row = df_old[df_old['DOCUMENT_NO'] == data_dict['DOCUMENT_NO']].iloc[0].to_dict()
                for key, value in data_dict.items():
                    if pd.notna(value) and value != "": old_row[key] = value
                df_old = df_old[df_old['DOCUMENT_NO'] != data_dict['DOCUMENT_NO']]
                df_new = pd.DataFrame([old_row])
            df_combined = pd.concat([df_old, df_new], ignore_index=True)
        else:
            df_combined = pd.DataFrame([data_dict])
        df_combined.to_excel(EXCEL_FILE, index=False)
        return True
    except PermissionError:
        st.error(f"❌ บันทึกข้อมูลไม่ได้: กรุณาปิดไฟล์หลัก '{EXCEL_FILE}' บนคอมพิวเตอร์ของคุณก่อน")
        return False

def check_all_departments_completed(doc_no):
    doc_data = get_document_data(doc_no)
    if not doc_data: return False
    for num in range(1, 20):
        rev_value = str(doc_data.get(f"DOC_{num}_REVISE", "NO")).upper()
        if rev_value == "YES":
            resp_value = str(doc_data.get(f"DOC_{num}_RESP", "")).strip()
            if resp_value == "" or resp_value == "-": return False
    return True 

def get_missing_items(doc_no):
    doc_data = get_document_data(doc_no)
    if not doc_data: return []
    missing_list = []
    for num in range(1, 20):
        rev_value = str(doc_data.get(f"DOC_{num}_REVISE", "NO")).upper()
        if rev_value == "YES":
            resp_value = str(doc_data.get(f"DOC_{num}_RESP", "")).strip()
            if resp_value == "" or resp_value == "-":
                dept, title = ITEM_DEPT_MAPPING.get(num, ("-", "-"))
                missing_list.append(f"ข้อ {num} [{dept}]: {title}")
    return missing_list

# =============================================================
# UI DISPLAY
# =============================================================
st.title("KFT - RELATED DOCUMENT CHANGE CONTROL SYSTEM")

reset_id = st.session_state.form_reset_counter

if "Print Form" in selected_dept:
    st.subheader("🖨️ ระบบดึงและพิมพ์ฟอร์มเอกสารควบคุมอัตโนมัติ (Excel Format บริษัท)")
    print_doc_no = st.text_input("กรอก DOCUMENT NO. ที่ต้องการแปลงข้อมูลออกฟอร์ม (เช่น R001/26) :", key=f"print_doc_{reset_id}").strip().upper()
    
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
    approve_doc_no = st.text_input("ระบุ DOCUMENT NO. ที่ต้องการพิจารณาอนุมัติ :", key=f"appr_doc_{reset_id}").strip().upper()
    
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
            mgr_name = st.text_input("พิมพ์ชื่อ-นามสกุล ของคุณเพื่อใช้ยืนยันการอนุมัติ :", key=f"mgr_name_{reset_id}").strip()
            
            if doc_status != "FINISH" and doc_status != "APPROVED":
                missing_items = get_missing_items(approve_doc_no)
                st.warning("⚠️ เอกสารนี้ยังไม่อยู่ในสถานะ 'FINISH' เนื่องจากพนักงานยังกรอกข้อมูลไม่ครบถ้วน")
                if missing_items:
                    with st.expander("🔍 **คลิกเพื่อดูรายการข้อและแผนกที่ยังไม่ได้กรอกข้อมูล**", expanded=True):
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
                            for num in range(1, 20): final_data[f"DOC_{num}_CLOSE"] = current_today
                            if save_to_excel(final_data): 
                                st.success("🎉 ใบงานนี้ผ่านการอนุมัติสมบูรณ์แล้ว!")
                                send_final_approved_email(approve_doc_no, cust, part, mgr_name)
                                st.rerun()
        else:
            st.error("❌ ไม่พบข้อมูลรหัสเอกสารควบคุมนี้")

else:
    st.subheader("📝 ส่วนที่ 1: รายละเอียดข้อมูลโครงสร้างวิศวกรรมทั่วไป (Header)")
    doc_no = st.text_input("📝 DOCUMENT NO. *จำเป็นต้องระบุ :", key=f"doc_no_{reset_id}").strip().upper()
    
    existing_data = None
    if doc_no:
        existing_data = get_document_data(doc_no)
        if existing_data and "PDD" not in selected_dept:
            st.info(f"✨ ค้นพบใบงานเลขที่ {doc_no} ในระบบ")

    is_disabled = False if "PDD" in selected_dept else True

    def get_val(key, default=""):
        return existing_data.get(key, default) if existing_data else default

    col_main_left, col_main_right = st.columns(2)
    with col_main_left:
        customer_list = ["-- เลือกชื่อลูกค้า --", "HONDA", "NISSAN", "TMA", "ADEINT", "OTHER"]
        saved_cust = get_val("CUSTOMER_NAME", "-- เลือกชื่อลูกค้า --")
        default_cust_index = customer_list.index(saved_cust) if saved_cust in customer_list else 0
        customer_name = st.selectbox("👤 CUSTOMER NAME :", customer_list, index=default_cust_index, disabled=is_disabled, key=f"cust_{reset_id}")
        part_name = st.text_input("PART NAME :", value=get_val("PART_NAME"), disabled=is_disabled, key=f"pname_{reset_id}")
        part_no = st.text_input("PART NO. :", value=get_val("PART_NO"), disabled=is_disabled, key=f"pno_{reset_id}")
    with col_main_right:
        model_name = st.text_input("MODEL :", value=get_val("MODEL"), disabled=is_disabled, key=f"model_{reset_id}")
        master_dwg = st.text_input("MASTER DWG. NO. :", value=get_val("MASTER_DWG_NO"), disabled=is_disabled, key=f"mdwg_{reset_id}")
        saved_issue_by = get_val("ISSUE_BY", st.session_state.user_name if "PDD" in selected_dept else "")
        issue_by = st.text_input("✍️ ISSUE BY (ผู้จัดทำเอกสาร) :", value=saved_issue_by, disabled=is_disabled, key=f"issue_{reset_id}")

    st.markdown("---")
    left_col, right_col = st.columns(2)
    with left_col:
        saved_ref_type = get_val("REF_DOC_TYPE", "CUSTOMER ECI No.")
        ref_types = ["CUSTOMER ECI No.", "DESIGN NOTE. No.", "PROCESS CHANGE No."]
        default_ref_index = ref_types.index(saved_ref_type) if saved_ref_type in ref_types else 0
        ref_doc_type = st.radio("เลือกประเภทเอกสารแจ้งแก้แบบ :", ref_types, index=default_ref_index, horizontal=True, disabled=is_disabled, key=f"reftype_{reset_id}")
        ref_doc_no = st.text_input(f"กรอกเลขที่เอกสาร ({ref_doc_type}) :", value=get_val("REF_DOC_NO"), disabled=is_disabled, key=f"refno_{reset_id}").strip().upper()
    with right_col:
        eff_event = st.text_input("EVENT (เงื่อนไขการเริ่มมีผล) :", value=get_val("EFF_EVENT"), disabled=is_disabled, key=f"event_{reset_id}")
        saved_plan_date = get_val("EFF_PLAN")
        try: default_plan = date.fromisoformat(saved_plan_date) if saved_plan_date else date.today()
        except ValueError: default_plan = date.today()
        eff_plan = st.date_input("PLAN (วันที่เริ่มแผนงาน) :", value=default_plan, disabled=is_disabled, key=f"effplan_{reset_id}")

    st.markdown("---")
    col_attach, col_judge = st.columns(2)
    with col_attach:
        st.markdown("##### 📎 **ATTACH CUSTOMER'S**")
        att_dwg = st.checkbox("DRAWING", value=(get_val("ATTACH_DRAWING") == "YES"), disabled=is_disabled, key=f"adwg_{reset_id}")
        att_eci = st.checkbox("ECI or Design Note.", value=(get_val("ATTACH_ECI") == "YES"), disabled=is_disabled, key=f"aeci_{reset_id}")
        att_meeting = st.checkbox("MEETING MINUTE", value=(get_val("ATTACH_MEETING") == "YES"), disabled=is_disabled, key=f"ameet_{reset_id}")
        att_others = st.checkbox("OTHERS", value=(get_val("ATTACH_OTHERS") == "YES"), disabled=is_disabled, key=f"aoth_{reset_id}")
        att_others_detail = st.text_input("ระบุ OTHERS (ถ้ามี) :", value=get_val("ATTACH_OTHERS_DETAIL"), disabled=is_disabled or not att_others, key=f"aothdet_{reset_id}")

    with col_judge:
        st.markdown("##### ⚖️ **JUDGEMENT (by PDD)**")
        saved_judge = get_val("JUDGEMENT", "FEASIBLE")
        judge_options = ["FEASIBLE", "IMPROBABILITY"]
        judge_idx = judge_options.index(saved_judge) if saved_judge in judge_options else 0
        judgement = st.radio("ผลการประเมินโดย PDD :", judge_options, index=judge_idx, disabled=is_disabled, key=f"judge_{reset_id}")

    st.markdown("---")
    subject_text = st.text_area("SUBJECT (บันทึกเนื้อหารายละเอียดสาเหตุการแก้ไขแบบวิศวกรรม):", value=get_val("SUBJECT_TEXT"), disabled=is_disabled, key=f"subj_{reset_id}")
    image_path_to_save = get_val("SUBJECT_IMAGE_PATH")
    
    if "PDD" in selected_dept:
        uploaded_image = st.file_uploader("อัปโหลดรูปภาพพิมพ์เขียวประกอบหัวข้อ SUBJECT (ถ้ามี):", type=["png", "jpg", "jpeg"], key=f"img_{reset_id}")
        if uploaded_image is not None and doc_no != "":
            image = Image.open(uploaded_image)
            st.image(image, caption="📷 ตัวอย่างรูปภาพที่เตรียมจัดเก็บ", width=300)
            img_dir = "stored_images"
            if not os.path.exists(img_dir): os.makedirs(img_dir)
            image_path_to_save = os.path.join(img_dir, f"{doc_no}_subject.png")
            image.save(image_path_to_save)
    else:
        if image_path_to_save and os.path.exists(image_path_to_save):
            st.image(Image.open(image_path_to_save), caption="📷 รูปภาพประกอบที่ส่งต่อมาจากแผนก PDD", width=300)

    st.markdown("---")
    dept_docs_mapping = {}
    current_dept_key = ""

    if "PDD" in selected_dept:
        current_dept_key = "PDD"
        dept_docs_mapping = {
            1: "MASTER DRAWING.", 2: "MATERIAL PART NO. LIST. , ACC DWG.", 
            3: "PROCESS FLOW CHART.", 4: "OPERATION MANUAL.", 
            5: "TEST RESULT.", 6: "FMEA (APQP TEAM)", 7: "TOOLING No"
        }
    elif "QC" in selected_dept:
        current_dept_key = "QC"
        dept_docs_mapping = {
            8: "CONTROL PLAN.", 9: "INCOMING SHEET.", 
            10: "FINAL INSPECTION SHEET.", 11: "W/I Out Going / TRAINING QC.", 
            12: "INSPECTION STD. + DATA CHECK.", 13: "MSA", 
            14: "PSW UP-DATE., PPAP APPROVAL.", 15: "CHECKING FIXTURE."
        }
    elif "PCD" in selected_dept:
        current_dept_key = "PCD"
        dept_docs_mapping = {16: "MATERIAL REQUIREMENT.", 17: "PACKING STANDARD."}
    elif "PRO" in selected_dept:
        current_dept_key = "PRD"
        dept_docs_mapping = {18: "WORKING INSTRUCTION.", 19: "TRAINING PRODUCTION."}

    st.subheader(f"📋 ส่วนที่ 2: กรอกรายการตรวจสอบสำหรับแผนก {current_dept_key}")
    dept_inputs = {}

    for num, doc_name in dept_docs_mapping.items():
        with st.expander(f"📄 ข้อ {num}: {doc_name}", expanded=True):
            c1, c2, c3 = st.columns([1, 2, 2])
            saved_rev = get_val(f"DOC_{num}_REVISE", "NO")
            rev_options = ["NO", "YES"]
            rev_index = rev_options.index(saved_rev) if saved_rev in rev_options else 0

            with c1: rev = st.radio(f"REVISE", rev_options, index=rev_index, key=f"rev_{num}_{reset_id}", horizontal=True)
            with c2: resp = st.text_input(f"RESPONSIBILITY PERSON", value=get_val(f"DOC_{num}_RESP"), key=f"resp_{num}_{reset_id}")
            
            saved_plan = get_val(f"DOC_{num}_PLAN")
            try: p_date_val = date.fromisoformat(saved_plan) if saved_plan else date.today()
            except ValueError: p_date_val = date.today()
            with c3: p_date = st.date_input(f"PLAN TO FINISH", value=p_date_val, key=f"plan_{num}_{reset_id}")

            dept_inputs[f"DOC_{num}_REVISE"] = rev
            dept_inputs[f"DOC_{num}_RESP"] = resp
            dept_inputs[f"DOC_{num}_PLAN"] = p_date.strftime('%Y-%m-%d')

    # =============================================================
    # 🔘 ส่วนของปุ่มควบคุม (บันทึกข้อมูล / Clear หน้าจอ เฉพาะ PDD)
    # =============================================================
    if "PDD" in selected_dept:
        col_btn_save, col_btn_clear = st.columns([3, 1])
        with col_btn_save:
            btn_save = st.button("💾 บันทึกข้อมูลและส่งต่อขั้นตอนถัดไป", type="primary", use_container_width=True)
        with col_btn_clear:
            btn_clear = st.button("🗑️ Clear ข้อมูลทั้งหมด", use_container_width=True)
    else:
        btn_save = st.button("💾 บันทึกข้อมูลและส่งต่อขั้นตอนถัดไป", type="primary", use_container_width=True)
        btn_clear = False

    if btn_clear:
        clear_all_inputs()

    # =============================================================
    # 🔘 ส่วนบันทึกข้อมูล พร้อมระบบ Validation เช็กการกรอกข้อมูลของแผนก
    # =============================================================
    if btn_save:
        if not doc_no:
            st.error("❌ กรุณากรอก DOCUMENT NO. ก่อนบันทึกข้อมูล")
        else:
            unfilled_items = []
            for num, doc_name in dept_docs_mapping.items():
                rev_val = dept_inputs.get(f"DOC_{num}_REVISE", "NO")
                resp_val = dept_inputs.get(f"DOC_{num}_RESP", "").strip()
                
                if rev_val == "YES" and (not resp_val or resp_val == "-"):
                    unfilled_items.append(f"ข้อ {num}: {doc_name}")

            if unfilled_items:
                st.error(f"❌ แผนก {current_dept_key} ยังกรอกข้อมูลไม่ครบถ้วน! รายการที่เลือก REVISE เป็น 'YES' กรุณาระบุชื่อผู้รับผิดชอบ (RESPONSIBILITY PERSON)")
                with st.expander("🔍 **คลิกเพื่อดูรายการข้อที่ต้องระบุผู้รับผิดชอบ**", expanded=True):
                    for item in unfilled_items:
                        st.write(f"⚠️ {item}")
            else:
                save_payload = {"DOCUMENT_NO": doc_no, "DOC_STATUS": "IN_PROGRESS"}
                
                if current_dept_key == "PDD":
                    save_payload.update({
                        "CUSTOMER_NAME": customer_name, "PART_NAME": part_name, "PART_NO": part_no,
                        "MODEL": model_name, "MASTER_DWG_NO": master_dwg, "DATE": date.today().strftime('%Y-%m-%d'),
                        "ISSUE_BY": issue_by, "REF_DOC_TYPE": ref_doc_type, "REF_DOC_NO": ref_doc_no,
                        "EFF_EVENT": eff_event, "EFF_PLAN": eff_plan.strftime('%Y-%m-%d'),
                        "ATTACH_DRAWING": "YES" if att_dwg else "NO", "ATTACH_ECI": "YES" if att_eci else "NO",
                        "ATTACH_MEETING": "YES" if att_meeting else "NO", "ATTACH_OTHERS": "YES" if att_others else "NO",
                        "ATTACH_OTHERS_DETAIL": att_others_detail, "JUDGEMENT": judgement,
                        "SUBJECT_TEXT": subject_text, "SUBJECT_IMAGE_PATH": image_path_to_save
                    })
                
                save_payload.update(dept_inputs)

                if save_to_excel(save_payload):
                    st.success(f"✅ บันทึกข้อมูลของแผนก {current_dept_key} เรียบร้อยแล้ว!")
                    
                    if check_all_departments_completed(doc_no):
                        save_to_excel({"DOCUMENT_NO": doc_no, "DOC_STATUS": "FINISH"})
                        st.balloons()
                        st.success("🎉 ทุกแผนกกรอกข้อมูลครบทั้ง 19 ข้อเรียบร้อยแล้ว! เอกสารเปลี่ยนสถานะเป็น 'FINISH' รอ MGR อนุมัติ")
                        send_all_completed_alert_email(doc_no, customer_name if 'customer_name' in locals() else "", part_name if 'part_name' in locals() else "")
                    else:
                        next_dept_map = {"PDD": "QC", "QC": "PCD", "PCD": "PRD"}
                        next_target = next_dept_map.get(current_dept_key)
                        if next_target:
                            send_next_dept_alert_email(doc_no, customer_name if 'customer_name' in locals() else "", part_name if 'part_name' in locals() else "", next_target)