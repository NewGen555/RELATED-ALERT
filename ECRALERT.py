import os
import requests  # ใช้สำหรับการยิง API ไปยัง LINE Notification
import openpyxl  # ใช้สำหรับหยอดค่าลงตารางแบบฟอร์มจริง
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
UPLOAD_DIR = "uploads"  # โฟลเดอร์เก็บรูปภาพ

if not os.path.exists(UPLOAD_DIR):
    os.makedirs(UPLOAD_DIR)

# =============================================================
# CONFIGURATION: LINE MESSAGING API & GROUP SETTINGS
# =============================================================
LINE_ACCESS_TOKEN = "RBMqGMQq55Qc+ia3TCT/eZbs6Hp/8eyFSRUCy5URtFhopGRzo83Y2m+7K4JZplUgOZi13r/f9JyHm9bLg4VRfuV84l6/zktHMm2hASsDevA0brJNfeTIqhHci5K3vKgIUJ9xnIM5yJftZPD6vReKegdB04t89/1O/w1cDnyilFU="
LINE_GROUP_ID = "C66b5ef8b6a38a80fe1320cbcb346db1f"
APP_URL = "https://related-alert-erh2rywrtchautlthjlrwb.streamlit.app/"

# =============================================================
# FUNCTION: ส่งแจ้งเตือนเข้า LINE Group
# =============================================================
def send_line_group_message(message_text):
    """ส่งข้อความเข้ากลุ่ม LINE ผ่าน LINE Messaging API"""
    url = "https://api.line.me/v2/bot/message/push"
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {LINE_ACCESS_TOKEN}"
    }
    payload = {
        "to": LINE_GROUP_ID,
        "messages": [
            {
                "type": "text",
                "text": message_text
            }
        ]
    }
    try:
        response = requests.post(url, headers=headers, json=payload, timeout=5)
        print(f"LINE Response: {response.status_code} - {response.text}")
        
        if response.status_code == 200:
            return True
        else:
            st.error(f"❌ ส่ง LINE ไม่สำเร็จ (Code: {response.status_code}) - {response.text}")
            return False
    except Exception as e:
        st.error(f"⚠️ เกิดข้อผิดพลาดในการส่ง LINE: {e}")
        return False

def send_next_dept_alert_line(doc_no, customer, part_name, target_dept):
    """แจ้งเตือนไปยังแผนกถัดไปตาม Workflow"""
    msg = (
        f"🚨 [Action Required] ใบงาน Change Control ถึงคิวแผนก {target_dept}!\n"
        f"────────────────────────\n"
        f"📄 DOCUMENT NO.: {doc_no}\n"
        f"👤 CUSTOMER: {customer}\n"
        f"⚙️ PART NAME: {part_name}\n"
        f"────────────────────────\n"
        f"แผนกก่อนหน้าบันทึกข้อมูลเรียบร้อยแล้ว\n"
        f"รบกวนทีมงานแผนก **{target_dept}** เข้าสู่ระบบเพื่อกรอกข้อมูลในส่วนของท่านครับ\n\n"
        f"🔗 ลิงก์เข้าสู่ระบบ:\n{APP_URL}"
    )
    return send_line_group_message(msg)

def send_all_completed_alert_line(doc_no, customer, part_name):
    """แจ้งเตือนเมื่อทุกแผนกกรอกข้อมูลครบ 100% ให้ผู้จัดการอนุมัติ"""
    msg = (
        f"✅ [Complete] กรอกข้อมูลครบถ้วนทุกแผนกแล้ว!\n"
        f"────────────────────────\n"
        f"📄 DOCUMENT NO.: {doc_no}\n"
        f"👤 CUSTOMER: {customer}\n"
        f"⚙️ PART NAME: {part_name}\n"
        f"────────────────────────\n"
        f"ทุกแผนก (PDD, QC, PCD, PRD) บันทึกรายการตรวจสอบครบถ้วนแล้ว\n"
        f"รบกวนทีมผู้จัดการเข้าสู่ระบบเพื่อพิจารณาอนุมัติ (Sign-off) ครับ\n\n"
        f"🔗 ลิงก์เข้าสู่ระบบ:\n{APP_URL}"
    )
    return send_line_group_message(msg)

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
# 👤 ระบบผู้ใช้งาน (User Database)
# =============================================================
USERS = {
    "pdd_user": {"password": "pdd1234", "dept": "PDD (Product Design)", "name": "ENGINEER PDD"},
    "qc_user": {"password": "qc1234", "dept": "QC (Quality Control)", "name": "ENGINEER QC"},
    "pcd_user": {"password": "pcd1234", "dept": "PCD (Production Control)", "name": "ENGINEER PCD"},
    "prd_user": {"password": "prd1234", "dept": "PRO (Production / PD)", "name": "ENGINEER Production"},
    "mgr_pdd": {"password": "mgrpdd1", "dept": "MGR - PDD (ผู้จัดการ PDD)", "name": "ผู้จัดการ PDD"},
    "mgr_qcd": {"password": "mgrqc1", "dept": "MGR - QCD (ผู้จัดการ QC)", "name": "ผู้จัดการ QC"},
    "mgr_pcd": {"password": "mgrpcd1", "dept": "MGR - PCD (ผู้จัดการ PCD)", "name": "ผู้จัดการ PCD"},
    "mgr_prd": {"password": "mgrprd1", "dept": "MGR - PD (ผู้จัดการ Production)", "name": "ผู้จัดการ PRD"},
    "gm_user": {"password": "gm1234", "dept": "AGM / GM (ผู้บริหารอนุมัติขั้นสุดท้าย)", "name": "ผู้บริหาร GM"},
    "print_user": {"password": "print1", "dept": "Print Form", "name": "เจ้าหน้าที่พิมพ์เอกสาร"},
}

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.current_user = None
    st.session_state.current_dept = None
    st.session_state.user_name = None

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

# ดึงบทบาทเข้าใช้งานโดยตรงตามสิทธิ์ของผู้ใช้งานที่เข้าสู่ระบบ
selected_dept = st.session_state.current_dept

# =============================================================
# 🖨️ ฟังก์ชันเปิด Template ตารางจริง แล้วเขียนข้อมูลลงตามพิกัดช่อง Excel
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
        
        # 1. เขียนข้อมูลส่วนหัว (Header)
        write_cell("D3", doc_data.get("PART_NAME", ""))
        write_cell("D4", doc_data.get("PART_NO", ""))
        write_cell("F5", doc_data.get("MASTER_DWG_NO", ""))
        
        write_cell("P3", doc_data.get("MODEL", ""))
        write_cell("X3", doc_data.get("DATE", ""))
        write_cell("X1", doc_data.get("DOCUMENT_NO", ""))
        write_cell("H8", doc_data.get("REF_DOC_NO", ""))
        write_cell("X4", doc_data.get("ISSUE_BY", ""))  # เขียนข้อมูล ISSUE BY ลงพิกัด X4
        
        # ข้อมูลกลุ่ม Effective
        write_cell("W7", doc_data.get("EFF_EVENT", ""))
        write_cell("W8", doc_data.get("EFF_PLAN", ""))
        write_cell("W9", doc_data.get("EFF_ACTUAL", ""))
        
        # รายละเอียดเรื่อง (SUBJECT)
        write_cell("D12", doc_data.get("SUBJECT_TEXT", ""))
        
        # 2. เขียนข้อมูลตารางข้อ 1 - 19
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
            write_cell(f"T{current_row}", doc_data.get(f"DOC_{i}_PLAN", ""))
            write_cell(f"W{current_row}", doc_data.get(f"DOC_{i}_CLOSE", ""))
            
        # 3. เขียนข้อมูลการลงนาม (Sign-off)
        write_cell("O41", doc_data.get("APPR_PDD_MGR", ""))
        write_cell("O45", doc_data.get("DATE_PDD_MGR", ""))
        write_cell("Q41", doc_data.get("APPR_QCD_MGR", ""))
        write_cell("Q45", doc_data.get("DATE_QCD_MGR", ""))
        write_cell("S41", doc_data.get("APPR_PCD_MGR", ""))
        write_cell("S45", doc_data.get("DATE_PCD_MGR", ""))
        write_cell("U41", doc_data.get("APPR_PD_MGR", ""))
        write_cell("U45", doc_data.get("DATE_PD_MGR", ""))
        write_cell("W41", doc_data.get("APPR_GM", ""))
        write_cell("W45", doc_data.get("DATE_GM", ""))
        
        safe_doc_no = doc_no.replace("/", "_").replace("\\", "_")
        output_filename = f"Change_Control_Sheet_{safe_doc_no}.xlsx"
        wb.save(output_filename)
        wb.close()
        return output_filename, None
    except Exception as e:
        return None, f"เกิดข้อผิดพลาดในการสร้างไฟล์ Excel: {str(e)}"

# =============================================================
# ส่วนประมวลผลและการจัดการฐานข้อมูลหลัก (Excel Database)
# =============================================================
def ตรวจเช็คและสร้างไฟล์():
    if not os.path.exists(EXCEL_FILE):
        columns = [
            "DOCUMENT_NO", "CUSTOMER_NAME", "PART_NAME", "PART_NO", "MODEL", "MASTER_DWG_NO", "DATE", "ISSUE_BY",
            "REF_DOC_TYPE", "REF_DOC_NO", "EFF_EVENT", "EFF_PLAN", "EFF_ACTUAL", "SUBJECT_TEXT", "SUBJECT_IMAGE_PATH", "DOC_STATUS",
            "APPR_PDD_MGR", "DATE_PDD_MGR", "APPR_QCD_MGR", "DATE_QCD_MGR", "APPR_PCD_MGR", "DATE_PCD_MGR", "APPR_PD_MGR", "DATE_PD_MGR", "APPR_GM", "DATE_GM"
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
        resp_value = str(doc_data.get(f"DOC_{num}_RESP", "")).strip()
        if resp_value == "": return False 
    return True 

# =============================================================
# หน้าจอการแสดงผล UI ตามสิทธิ์แผนกที่เข้าสู่ระบบอัตโนมัติ
# =============================================================
st.title("KFT - RELATED DOCUMENT CHANGE CONTROL SYSTEM")

# ---------------------------------------------------------
# 🌟 เมนูพิเศษ: ขาดาวน์โหลดและพิมพ์แบบฟอร์มตารางจริง (Print Form)
# ---------------------------------------------------------
if "Print Form" in selected_dept:
    st.subheader("🖨️ ระบบดึงและพิมพ์ฟอร์มเอกสารควบคุมอัตโนมัติ (Excel Format บริษัท)")
    print_doc_no = st.text_input("กรอก DOCUMENT NO. ที่ต้องการแปลงข้อมูลออกฟอร์ม (เช่น R001/26) :").strip().upper()
    
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

# ---------------------------------------------------------
# 🔒 เมนู: กลุ่มหัวหน้างานและผู้บริหารลงนาม (Manager Sign-Off)
# ---------------------------------------------------------
elif "MGR" in selected_dept or "GM" in selected_dept:
    st.subheader("🔒 ระบบพิจารณาลงนามดิจิทัลและอนุมัติเอกสารควบคุม (Manager Sign-Off)")
    approve_doc_no = st.text_input("ระบุ DOCUMENT NO. ที่ต้องการพิจารณาอนุมัติ :").strip().upper()
    
    if approve_doc_no:
        doc_data = get_document_data(approve_doc_no)
        if doc_data:
            st.info(f"📋 รายละเอียดใบงาน -> ลูกค้า: {doc_data.get('CUSTOMER_NAME')} | ชื่อพาร์ท: {doc_data.get('PART_NAME')}")
            st.markdown("##### **📊 ตารางเช็กสถานะการเซ็นอนุมัติในขณะนี้:**")
            
            c_pdd, c_qcd, c_pcd, c_pd, c_gm = st.columns(5)
            c_pdd.metric("1. PDD MGR.", doc_data.get("APPR_PDD_MGR") if doc_data.get("APPR_PDD_MGR") else "⏳ รออนุมัติ", doc_data.get("DATE_PDD_MGR"))
            c_qcd.metric("2. QCD MGR.", doc_data.get("APPR_QCD_MGR") if doc_data.get("APPR_QCD_MGR") else "⏳ รออนุมัติ", doc_data.get("DATE_QCD_MGR"))
            c_pcd.metric("3. PCD MGR.", doc_data.get("APPR_PCD_MGR") if doc_data.get("APPR_PCD_MGR") else "⏳ รออนุมัติ", doc_data.get("DATE_PCD_MGR"))
            c_pd.metric("4. PD MGR.", doc_data.get("APPR_PD_MGR") if doc_data.get("APPR_PD_MGR") else "⏳ รออนุมัติ", doc_data.get("DATE_PD_MGR"))
            c_gm.metric("5. AGM / GM", doc_data.get("APPR_GM") if doc_data.get("APPR_GM") else "⏳ รออนุมัติ", doc_data.get("DATE_GM"))
            
            st.markdown("---")
            current_today = date.today().strftime('%Y-%m-%d')
            mgr_name = st.text_input("พิมพ์ชื่อ-นามสกุล ของคุณเพื่อใช้ยืนยันการอนุมัติ :").strip()
            
            if "MGR - PDD" in selected_dept:
                if st.button("🖊️ อนุมัติในฐานะ PDD MGR.", type="primary"):
                    if mgr_name == "": st.error("❌ กรุณาพิมพ์ชื่อตัวตนของคุณก่อน")
                    elif save_to_excel({"DOCUMENT_NO": approve_doc_no, "APPR_PDD_MGR": mgr_name, "DATE_PDD_MGR": current_today}):
                        st.success("✅ บันทึกการเซ็นอนุมัติของ PDD MGR. เรียบร้อย!"); st.rerun()

            elif "MGR - QCD" in selected_dept:
                if st.button("🖊️ อนุมัติในฐานะ QCD MGR.", type="primary"):
                    if mgr_name == "": st.error("❌ กรุณาพิมพ์ชื่อตัวตนของคุณก่อน")
                    elif save_to_excel({"DOCUMENT_NO": approve_doc_no, "APPR_QCD_MGR": mgr_name, "DATE_QCD_MGR": current_today}):
                        st.success("✅ บันทึกการเซ็นอนุมัติของ QCD MGR. เรียบร้อย!"); st.rerun()

            elif "MGR - PCD" in selected_dept:
                if st.button("🖊️ อนุมัติในฐานะ PCD MGR.", type="primary"):
                    if mgr_name == "": st.error("❌ กรุณาพิมพ์ชื่อตัวตนของคุณก่อน")
                    elif save_to_excel({"DOCUMENT_NO": approve_doc_no, "APPR_PCD_MGR": mgr_name, "DATE_PCD_MGR": current_today}):
                        st.success("✅ บันทึกการเซ็นอนุมัติของ PCD MGR. เรียบร้อย!"); st.rerun()

            elif "MGR - PD" in selected_dept:
                if st.button("🖊️ อนุมัติในฐานะ PD MGR.", type="primary"):
                    if mgr_name == "": st.error("❌ กรุณาพิมพ์ชื่อตัวตนของคุณก่อน")
                    elif save_to_excel({"DOCUMENT_NO": approve_doc_no, "APPR_PD_MGR": mgr_name, "DATE_PD_MGR": current_today}):
                        st.success("✅ บันทึกการเซ็นอนุมัติของ PD MGR. เรียบร้อย!"); st.rerun()

            elif "AGM / GM" in selected_dept:
                if not check_all_departments_completed(approve_doc_no):
                    st.error("⚠️ ผู้อนุมัติขั้นสุดท้ายยังไม่สามารถลงนามได้: เนื่องจากพนักงานยังบันทึกข้อตรวจสอบไม่ครบทั้ง 19 ข้อ")
                else:
                    if st.button("🏆 ยืนยันปิดงานขั้นสุดท้าย (AGM / GM APPROVAL)", type="primary"):
                        if mgr_name == "": st.error("❌ กรุณาพิมพ์ชื่อตัวตนของคุณก่อน")
                        else:
                            final_data = {"DOCUMENT_NO": approve_doc_no, "APPR_GM": mgr_name, "DATE_GM": current_today, "DOC_STATUS": "APPROVED", "EFF_ACTUAL": current_today}
                            for num in range(1, 20): final_data[f"DOC_{num}_CLOSE"] = current_today
                            if save_to_excel(final_data): st.success("🎉 ใบงานนี้ผ่านการอนุมัติสมบูรณ์และปิดเคสถาวรแล้ว!"); st.rerun()
        else:
            st.error("❌ ไม่พบข้อมูลรหัสเอกสารควบคุมนี้")

# ---------------------------------------------------------
# 📝 เมนู: ขากรอกตารางรายการแยกตามผู้ปฏิบัติงานฝ่ายต่างๆ (PDD, QC, PCD, PRO)
# ---------------------------------------------------------
else:
    st.subheader("📝 ส่วนที่ 1: รายละเอียดข้อมูลโครงสร้างวิศวกรรมทั่วไป (Header)")
    doc_no = st.text_input("📝 DOCUMENT NO. *จำเป็นต้องระบุ :").strip().upper()
    
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
        customer_list = ["-- เลือกชื่อลูกค้า --", "CUSTOMER_A", "CUSTOMER_B", "CUSTOMER_C", "CUSTOMER_D", "CUSTOMER_E"]
        saved_cust = get_val("CUSTOMER_NAME", "-- เลือกชื่อลูกค้า --")
        default_cust_index = customer_list.index(saved_cust) if saved_cust in customer_list else 0
        customer_name = st.selectbox("👤 CUSTOMER NAME :", customer_list, index=default_cust_index, disabled=is_disabled)
        part_name = st.text_input("PART NAME :", value=get_val("PART_NAME"), disabled=is_disabled)
        part_no = st.text_input("PART NO. :", value=get_val("PART_NO"), disabled=is_disabled)
    with col_main_right:
        model_name = st.text_input("MODEL :", value=get_val("MODEL"), disabled=is_disabled)
        master_dwg = st.text_input("MASTER DWG. NO. :", value=get_val("MASTER_DWG_NO"), disabled=is_disabled)
        
        # 📌 ปรับแก้ไขตรงนี้: แสดงช่อง ISSUE BY ตลอดเวลา (PDD กรอกได้ / แผนกอื่นดูได้อย่างเดียว)
        saved_issue_by = get_val("ISSUE_BY", st.session_state.user_name if "PDD" in selected_dept else "")
        issue_by = st.text_input("✍️ ISSUE BY (ผู้จัดทำเอกสาร) :", value=saved_issue_by, disabled=is_disabled)

    st.markdown("---")
    left_col, right_col = st.columns(2)
    with left_col:
        saved_ref_type = get_val("REF_DOC_TYPE", "CUSTOMER ECI No.")
        ref_types = ["CUSTOMER ECI No.", "DESIGN NOTE. No.", "PROCESS CHANGE No."]
        default_ref_index = ref_types.index(saved_ref_type) if saved_ref_type in ref_types else 0
        ref_doc_type = st.radio("เลือกประเภทเอกสารแจ้งแก้แบบ :", ref_types, index=default_ref_index, horizontal=True, disabled=is_disabled)
        ref_doc_no = st.text_input(f"กรอกเลขที่เอกสาร ({ref_doc_type}) :", value=get_val("REF_DOC_NO"), disabled=is_disabled).strip().upper()
    with right_col:
        eff_event = st.text_input("EVENT (เงื่อนไขการเริ่มมีผล) :", value=get_val("EFF_EVENT"), disabled=is_disabled)
        saved_plan_date = get_val("EFF_PLAN")
        try: default_plan = date.fromisoformat(saved_plan_date) if saved_plan_date else date.today()
        except ValueError: default_plan = date.today()
        eff_plan = st.date_input("PLAN (วันที่เริ่มแผนงาน) :", value=default_plan, disabled=is_disabled)

    st.markdown("---")
    subject_text = st.text_area("SUBJECT (บันทึกเนื้อหารายละเอียดสาเหตุการแก้ไขแบบวิศวกรรม):", value=get_val("SUBJECT_TEXT"), disabled=is_disabled)
    image_path_to_save = get_val("SUBJECT_IMAGE_PATH")
    
    if "PDD" in selected_dept:
        uploaded_image = st.file_uploader("อัปโหลดรูปภาพพิมพ์เขียวประกอบหัวข้อ SUBJECT (ถ้ามี):", type=["png", "jpg", "jpeg"])
        if uploaded_image is not None and doc_no != "":
            image = Image.open(uploaded_image)
            st.image(image, caption="📷 ตัวอย่างรูปภาพที่เตรียมจัดเก็บ", width=300)
            img_dir = "stored_images"
            if not os.path.exists(img_dir): os.makedirs(img_dir)
            image_path_to_save = os.path.join(img_dir, f"{doc_no}_subject.png")
    else:
        if image_path_to_save and os.path.exists(image_path_to_save):
            st.image(Image.open(image_path_to_save), caption="📷 รูปภาพประกอบที่ส่งต่อมาจากแผนก PDD", width=300)

    st.markdown("---")
    dept_docs_mapping = {}
    current_dept_key = ""

    if "PDD" in selected_dept:
        current_dept_key = "PDD"
        dept_docs_mapping = {1: "MASTER DRAWING.", 2: "MATERIAL PART NO. LIST. , ACC DWG.", 3: "PROCESS FLOW CHART.", 4: "OPERATION MANUAL.", 5: "TEST RESULT.", 6: "FMEA (APQP TEAM)", 7: "TOOLING No"}
    elif "QC" in selected_dept:
        current_dept_key = "QC"
        dept_docs_mapping = {8: "CONTROL PLAN.", 9: "INCOMING SHEET.", 10: "FINAL INSPECTION SHEET.", 11: "W/I Out Going / TRAINING QC.", 12: "INSPECTION STD. + DATA CHECK.", 13: "MSA", 14: "PSW UP-DATE., PPAP APPROVAL.", 15: "CHECKING FIXTURE."}
    elif "PCD" in selected_dept:
        current_dept_key = "PCD"
        dept_docs_mapping = {16: "MATERIAL REQUIREMENT.", 17: "PACKING STANDARD."}
    elif "PRO" in selected_dept:
        current_dept_key = "PRO"
        dept_docs_mapping = {18: "WORKING INSTRUCTION.", 19: "TRAINING PRODUCTION."}

    dept_inputs = {}
    for num, doc_name in dept_docs_mapping.items():
        with st.expander(f"📄 ข้อ {num}: {doc_name}", expanded=True):
            c1, c2, c3 = st.columns([1, 2, 2])
            saved_rev = get_val(f"DOC_{num}_REVISE", "NO")
            rev_options = ["NO", "YES"]
            rev_index = rev_options.index(saved_rev) if saved_rev in rev_options else 0
            saved_p_date = get_val(f"DOC_{num}_PLAN")
            try: default_p_date = date.fromisoformat(saved_p_date) if saved_p_date else date.today()
            except ValueError: default_p_date = date.today()

            with c1: rev = st.radio(f"REVISE", rev_options, index=rev_index, key=f"rev_{num}", horizontal=True)
            with c2: resp = st.text_input(f"RESPONSIBILITY PERSON", value=get_val(f"DOC_{num}_RESP"), key=f"resp_{num}")
            with c3: p_date = st.date_input(f"PLAN TO FINISH", value=default_p_date, key=f"p_{num}")
            
            dept_inputs[f"DOC_{num}_REVISE"] = rev
            dept_inputs[f"DOC_{num}_RESP"] = resp
            dept_inputs[f"DOC_{num}_PLAN"] = p_date.strftime('%Y-%m-%d')

    if st.button(f"💾 บันทึกข้อมูลของแผนก {current_dept_key} ลงฐานข้อมูล", type="primary"):
        if not doc_no: st.error("❌ กรุณาระบุกรอกรหัส DOCUMENT NO. ก่อนบันทึก")
        else:
            if "PDD" in selected_dept:
                final_data = {
                    "DOCUMENT_NO": doc_no, 
                    "CUSTOMER_NAME": customer_name, 
                    "PART_NAME": part_name, 
                    "PART_NO": part_no, 
                    "MODEL": model_name, 
                    "MASTER_DWG_NO": master_dwg, 
                    "ISSUE_BY": issue_by,  # 📌 เพิ่มการเซฟค่า ISSUE BY
                    "DATE": date.today().strftime('%Y-%m-%d'), 
                    "REF_DOC_TYPE": ref_doc_type, 
                    "REF_DOC_NO": ref_doc_no, 
                    "EFF_EVENT": eff_event, 
                    "EFF_PLAN": eff_plan.strftime('%Y-%m-%d'), 
                    "SUBJECT_TEXT": subject_text, 
                    "SUBJECT_IMAGE_PATH": image_path_to_save, 
                    "DOC_STATUS": "PENDING", 
                    **dept_inputs
                }
                if 'uploaded_image' in locals() and uploaded_image is not None: image.save(image_path_to_save)
            else:
                final_data = {"DOCUMENT_NO": doc_no, **dept_inputs}
            
            if save_to_excel(final_data):
                st.success(f"🎉 บันทึกข้อมูลของแผนก {current_dept_key} สำเร็จแล้ว!")
                
                # ตรวจเช็คว่ากรอกครบทุกแผนกแล้วหรือยัง
                if check_all_departments_completed(doc_no):
                    latest_info = get_document_data(doc_no)
                    send_all_completed_alert_line(
                        doc_no, 
                        latest_info.get("CUSTOMER_NAME", "-"), 
                        latest_info.get("PART_NAME", "-")
                    )