import os
import requests
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

# -------------------------------------------------------------
# ฟังก์ชันส่งแจ้งเตือนตามลำดับขั้นตอน (Sequential Alerts)
# -------------------------------------------------------------
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
    "prd_user": {"password": "prd1234", "dept": "PRD (Production / PRD)", "name": "ENGINEER Production"},
    "mgr_pdd": {"password": "mgrpdd1", "dept": "MGR - PDD (ผู้จัดการ PDD)", "name": "ผู้จัดการ PDD"},
    "mgr_qcd": {"password": "mgrqc1", "dept": "MGR - QCD (ผู้จัดการ QC)", "name": "ผู้จัดการ QC"},
    "mgr_pcd": {"password": "mgrpcd1", "dept": "MGR - PCD (ผู้จัดการ PCD)", "name": "ผู้จัดการ PCD"},
    "mgr_prd": {"password": "mgrprd1", "dept": "MGR - PRD (ผู้จัดการ Production)", "name": "ผู้จัดการ PRD"},
    "gm_user": {"password": "gm1234", "dept": "AGM / GM (ผู้บริหารอนุมัติขั้นสุดท้าย)", "name": "ผู้บริหาร GM"},
    "print_user": {"password": "print1", "dept": "🖨️ Print Form (ดาวน์โหลดเอกสารแบบฟอร์มจริง)", "name": "เจ้าหน้าที่พิมพ์เอกสาร"},
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
    st.markdown(f"👤 **{st.session_state.user_name}**\n\n🏢 {st.session_state.current_dept}")
    st.markdown("---")
    if st.button("🚪 ออกจากระบบ", use_container_width=True):
        logout()

selected_dept = st.session_state.current_dept if st.session_state.current_dept else ""


# =============================================================
# Helper Functions Database & Excel
# =============================================================
def init_db():
    if not os.path.exists(EXCEL_FILE):
        cols = [
            "DOCUMENT_NO", "CUSTOMER_NAME", "PART_NAME", "PART_NO", "MODEL", "MASTER_DWG_NO", "DATE",
            "REF_DOC_TYPE", "REF_DOC_NO", "EFF_EVENT", "EFF_PLAN", "EFF_ACTUAL", "SUBJECT_TEXT", "SUBJECT_IMAGE_PATH", "DOC_STATUS",
            "APPR_PDD_MGR", "DATE_PDD_MGR", "APPR_QCD_MGR", "DATE_QCD_MGR", "APPR_PCD_MGR", "DATE_PCD_MGR", "APPR_PRD_MGR", "DATE_PRD_MGR", "APPR_GM", "DATE_GM","ISSUE_BY"
        ]
        for num in range(1, 20):
            cols.extend([f"DOC_{num}_REVISE", f"DOC_{num}_RESP", f"DOC_{num}_PLAN", f"DOC_{num}_CLOSE"])
        pd.DataFrame(columns=cols).to_excel(EXCEL_FILE, index=False)

init_db()

def get_document_data(doc_no):
    if os.path.exists(EXCEL_FILE):
        df = pd.read_excel(EXCEL_FILE)
        if 'DOCUMENT_NO' in df.columns and doc_no in df['DOCUMENT_NO'].values:
            row_data = df[df['DOCUMENT_NO'] == doc_no].iloc[0]
            return {k: ("" if pd.isna(v) else v) for k, v in row_data.to_dict().items()}
    return None

def save_to_excel(data_dict):
    try:
        if os.path.exists(EXCEL_FILE):
            df_old = pd.read_excel(EXCEL_FILE)
            if 'DOCUMENT_NO' in df_old.columns and data_dict['DOCUMENT_NO'] in df_old['DOCUMENT_NO'].values:
                old_row = df_old[df_old['DOCUMENT_NO'] == data_dict['DOCUMENT_NO']].iloc[0].to_dict()
                for k, v in data_dict.items():
                    if pd.notna(v) and v != "": old_row[k] = v
                df_old = df_old[df_old['DOCUMENT_NO'] != data_dict['DOCUMENT_NO']]
                df_new = pd.DataFrame([old_row])
            else:
                df_new = pd.DataFrame([data_dict])
            df_combined = pd.concat([df_old, df_new], ignore_index=True)
        else:
            df_combined = pd.DataFrame([data_dict])
        df_combined.to_excel(EXCEL_FILE, index=False)
        return True
    except PermissionError:
        st.error(f"❌ กรุณาปิดไฟล์ '{EXCEL_FILE}' บนเครื่องก่อนบันทึก")
        return False

def check_range_completed(doc_no, start_num, end_num):
    """ฟังก์ชันเช็กว่ารายการในข้อช่วงที่ระบุกรอก RESPครบแล้วหรือยัง"""
    doc_data = get_document_data(doc_no)
    if not doc_data: return False
    for num in range(start_num, end_num + 1):
        if str(doc_data.get(f"DOC_{num}_RESP", "")).strip() == "":
            return False
    return True


# =============================================================
# 🏠 Main App
# =============================================================
st.title("🔷 KFT - CHANGE CONTROL SYSTEM")
st.caption(f"ผู้ใช้: {st.session_state.user_name} | แผนก: {st.session_state.current_dept}")
st.markdown("---")

# ---------------------------------------------------------
# 🖨️ Print Form Mode
# ---------------------------------------------------------
if "Print Form" in selected_dept:
    st.subheader("🖨️ ระบบพิมพ์ฟอร์มเอกสารควบคุม")
    print_doc_no = st.text_input("กรอก DOCUMENT NO. :").strip().upper()
    if print_doc_no:
        doc_data = get_document_data(print_doc_no)
        if doc_data:
            st.success(f"พบข้อมูลเอกสาร {print_doc_no}")
        else:
            st.error("❌ ไม่พบข้อมูลใบงานนี้")

# ---------------------------------------------------------
# 🔒 Manager Approval Mode
# ---------------------------------------------------------
elif "MGR" in selected_dept or "GM" in selected_dept:
    st.subheader("🔒 ระบบพิจารณาลงนามดิจิทัล (Manager Sign-Off)")
    approve_doc_no = st.text_input("ระบุ DOCUMENT NO. :").strip().upper()
    if approve_doc_no:
        doc_data = get_document_data(approve_doc_no)
        if doc_data:
            st.info(f"ลูกค้า: {doc_data.get('CUSTOMER_NAME')} | พาร์ท: {doc_data.get('PART_NAME')}")
            mgr_name = st.text_input("พิมพ์ชื่อ-นามสกุล ของคุณเพื่อยืนยัน :").strip()
            current_today = date.today().strftime('%Y-%m-%d')
            
            if st.button("🖊️ ยืนยันอนุมัติเอกสาร", type="primary"):
                if not mgr_name:
                    st.error("❌ กรุณาพิมพ์ชื่อของคุณก่อน")
                else:
                    update_dict = {"DOCUMENT_NO": approve_doc_no}
                    if "MGR - PDD" in selected_dept: update_dict.update({"APPR_PDD_MGR": mgr_name, "DATE_PDD_MGR": current_today})
                    elif "MGR - QCD" in selected_dept: update_dict.update({"APPR_QCD_MGR": mgr_name, "DATE_QCD_MGR": current_today})
                    elif "MGR - PCD" in selected_dept: update_dict.update({"APPR_PCD_MGR": mgr_name, "DATE_PCD_MGR": current_today})
                    elif "MGR - PRD" in selected_dept: update_dict.update({"APPR_PRD_MGR": mgr_name, "DATE_PRD_MGR": current_today})
                    elif "AGM / GM" in selected_dept: 
                        # เมื่อ GM อนุมัติเป็นขั้นตอนสุดท้าย -> ตั้งสถานะเป็น APPROVED และบันทึก EFFECTIVE DATE (ACTUAL) อัตโนมัติ
                        update_dict.update({
                            "APPR_GM": mgr_name, 
                            "DATE_GM": current_today, 
                            "DOC_STATUS": "APPROVED",
                            "EFF_ACTUAL": current_today
                        })
                    
                    if save_to_excel(update_dict):
                        st.success("✅ เซ็นอนุมัติเรียบร้อยแล้ว!")
                        st.rerun()

# ---------------------------------------------------------
# 📝 Data Entry Mode (PDD, QC, PCD, PRD)
# ---------------------------------------------------------
else:
    st.subheader("📝 กรอกข้อมูลเอกสารเปลี่ยนแปลง")
    doc_no = st.text_input("📝 DOCUMENT NO. * :", key="main_doc_no").strip().upper()
    
    existing_data = get_document_data(doc_no) if doc_no else None
    is_disabled = False if "PDD" in selected_dept else True
    
    def get_val(k, default=""): return existing_data.get(k, default) if existing_data else default

    # 1. ข้อมูลหลัก Header
    c_left, c_right = st.columns(2)
    with c_left:
        cust_list = ["-- เลือกชื่อลูกค้า --", "ADIENT", "HONDA", "NISSAN", "TOYOTA", "OTHER"]
        saved_c = get_val("CUSTOMER_NAME", "-- เลือกชื่อลูกค้า --")
        idx_c = cust_list.index(saved_c) if saved_c in cust_list else 0
        customer_name = st.selectbox("👤 CUSTOMER NAME :", cust_list, index=idx_c, disabled=is_disabled)
        part_name = st.text_input("PART NAME :", value=get_val("PART_NAME"), disabled=is_disabled)
        part_no = st.text_input("PART NO. :", value=get_val("PART_NO"), disabled=is_disabled)
    with c_right:
        model_name = st.text_input("MODEL :", value=get_val("MODEL"), disabled=is_disabled)
        master_dwg = st.text_input("MASTER DWG. NO. :", value=get_val("MASTER_DWG_NO"), disabled=is_disabled)
        issue_by = st.text_input("ISSUE BY :", value=get_val("ISSUE_BY"), disabled=is_disabled)

    st.markdown("---")

    # 2. ส่วนเอกสารอ้างอิง และ EFFECTIVE DATE
    col_ref_left, col_ref_right = st.columns(2)
    
    with col_ref_left:
        st.markdown("##### 📄 REFERENCE DOCUMENT")
        ref_options = ["CUSTOMER ECI No.", "DESIGN NOTE No.", "PROCESS CHANGE No."]
        saved_ref_type = get_val("REF_DOC_TYPE", "CUSTOMER ECI No.")
        idx_ref = ref_options.index(saved_ref_type) if saved_ref_type in ref_options else 0
        
        ref_doc_type = st.radio(
            "เลือกประเภทเอกสารอ้างอิง :",
            ref_options,
            index=idx_ref,
            disabled=is_disabled
        )
        ref_doc_no = st.text_input(
            "เลขที่เอกสารอ้างอิง (Document No.) :",
            value=get_val("REF_DOC_NO"),
            disabled=is_disabled
        )

    with col_ref_right:
        st.markdown("##### 📅 EFFECTIVE DATE")
        eff_event = st.text_input("EVENT :", value=get_val("EFF_EVENT"), disabled=is_disabled)
        eff_plan = st.text_input("PLAN :", value=get_val("EFF_PLAN"), disabled=is_disabled)
        
        # ช่อง ACTUAL จะดึงวันที่มาแสดงอัตโนมัติเมื่ออนุมัติครบแล้ว (Locked)
        actual_val = get_val("EFF_ACTUAL")
        eff_actual = st.text_input(
            "ACTUAL (อัตโนมัติเมื่ออนุมัติครบแล้ว) :", 
            value=actual_val if actual_val else "รออนุมัติครบถ้วน...", 
            disabled=True
        )

    st.markdown("---")

    # 3. ส่วน SUBJECT & INSERT PICTURE
    st.markdown("##### 📌 รายละเอียดหัวข้อการเปลี่ยนแปลง (SUBJECT & PICTURE)")
    col_sub1, col_sub2 = st.columns(2)
    with col_sub1:
        subject_text = st.text_area(
            "📝 SUBJECT / DETAILS OF CHANGE (รายละเอียดการเปลี่ยนแปลง) :",
            value=get_val("SUBJECT_TEXT"),
            height=130,
            disabled=is_disabled
        )
    
    saved_img_path = get_val("SUBJECT_IMAGE_PATH")
    image_file_path = saved_img_path  # เก็บ Path รูปเพื่อลง Excel

    with col_sub2:
        st.write("🖼️ ATTACHED IMAGE (รูปภาพประกอบ) :")
        uploaded_img = st.file_uploader(
            "อัปโหลดรูปภาพประกอบ (.png, .jpg, .jpeg)",
            type=["png", "jpg", "jpeg"],
            disabled=is_disabled,
            key="img_uploader"
        )
        
        # จัดการบันทึกไฟล์รูปภาพ
        if uploaded_img is not None and not is_disabled:
            file_ext = uploaded_img.name.split(".")[-1]
            img_filename = f"{doc_no if doc_no else 'temp'}_subject.{file_ext}"
            image_file_path = os.path.join(UPLOAD_DIR, img_filename)
            with open(image_file_path, "wb") as f:
                f.write(uploaded_img.getbuffer())
            st.image(uploaded_img, caption="รูปภาพที่อัปโหลดใหม่", use_column_width=True)
        elif saved_img_path and os.path.exists(str(saved_img_path)):
            st.image(saved_img_path, caption="รูปภาพประกอบเดิม", use_column_width=True)

    st.markdown("---")
    
    # 4. กำหนดหัวข้อการตรวจสอบตามแผนก
    dept_docs_mapping = {}
    if "PDD" in selected_dept:
        dept_docs_mapping = {1: "MASTER DRAWING.", 2: "MATERIAL PART NO. LIST.", 3: "PROCESS FLOW CHART.", 4: "OPERATION MANUAL.", 5: "TEST RESULT.", 6: "FMEA", 7: "TOOLING No"}
    elif "QC" in selected_dept:
        dept_docs_mapping = {8: "CONTROL PLAN.", 9: "INCOMING SHEET.", 10: "FINAL INSPECTION SHEET.", 11: "W/I Out Going", 12: "INSPECTION STD.", 13: "MSA", 14: "PPAP APPROVAL.", 15: "CHECKING FIXTURE."}
    elif "PCD" in selected_dept:
        dept_docs_mapping = {16: "MATERIAL REQUIREMENT.", 17: "PACKING STANDARD."}
    elif "PRD" in selected_dept:
        dept_docs_mapping = {18: "WORKING INSTRUCTION.", 19: "TRAINING PRODUCTION."}

    dept_inputs = {}
    for num, name in dept_docs_mapping.items():
        with st.expander(f"📄 ข้อ {num}: {name}", expanded=True):
            col1, col2, col3 = st.columns([1, 2, 2])
            rev = col1.radio("REVISE", ["NO", "YES"], index=0 if get_val(f"DOC_{num}_REVISE") != "YES" else 1, key=f"rev_{num}", horizontal=True)
            resp = col2.text_input("RESPONSIBLE PERSON", value=get_val(f"DOC_{num}_RESP"), key=f"resp_{num}")
            p_date = col3.date_input("PLAN DATE", key=f"p_{num}")
            
            dept_inputs[f"DOC_{num}_REVISE"] = rev
            dept_inputs[f"DOC_{num}_RESP"] = resp
            dept_inputs[f"DOC_{num}_PLAN"] = str(p_date)

    st.markdown("---")
    if st.button("💾 บันทึกข้อมูลและส่งแจ้งเตือนเข้า LINE", type="primary", use_container_width=True):
        if not doc_no:
            st.error("❌ กรุณาระบุ DOCUMENT NO.")
        else:
            save_payload = {
                "DOCUMENT_NO": doc_no, "CUSTOMER_NAME": customer_name, "PART_NAME": part_name,
                "PART_NO": part_no, "MODEL": model_name, "MASTER_DWG_NO": master_dwg, "ISSUE_BY": issue_by,
                "REF_DOC_TYPE": ref_doc_type, "REF_DOC_NO": ref_doc_no,
                "EFF_EVENT": eff_event, "EFF_PLAN": eff_plan,
                "SUBJECT_TEXT": subject_text, "SUBJECT_IMAGE_PATH": image_file_path,
                "DATE": str(date.today())
            }
            save_payload.update(dept_inputs)
            
            if save_to_excel(save_payload):
                st.success("✅ บันทึกข้อมูลเข้าฐานข้อมูลสำเร็จ!")
                
                # -------------------------------------------------
                # 🚨 ส่วนการส่ง LINE ตามเงื่อนไขแบบเป็นลำดับขั้นตอน (Sequential Flow)
                # -------------------------------------------------
                if "PDD" in selected_dept:
                    # PDD กรอกข้อ 1-7 เสร็จ -> ส่งหา QC
                    if check_range_completed(doc_no, 1, 7):
                        if send_next_dept_alert_line(doc_no, customer_name, part_name, "QC"):
                            st.info("📲 บันทึกส่วน PDD ครบแล้ว! ส่ง LINE แจ้งเตือนแผนก QC เรียบร้อย")
                    else:
                        st.warning("⚠️ บันทึกสำเร็จ แต่กรุณากรอกผู้รับผิดชอบ (RESPONSIBLE PERSON) ข้อ 1-7 ให้ครบถ้วนเพื่อส่ง LINE หา QC")

                elif "QC" in selected_dept:
                    # QC กรอกข้อ 8-15 เสร็จ -> ส่งหา PCD
                    if check_range_completed(doc_no, 8, 15):
                        if send_next_dept_alert_line(doc_no, customer_name, part_name, "PCD"):
                            st.info("📲 บันทึกส่วน QC ครบแล้ว! ส่ง LINE แจ้งเตือนแผนก PCD เรียบร้อย")
                    else:
                        st.warning("⚠️ บันทึกสำเร็จ แต่กรุณากรอกผู้รับผิดชอบ (RESPONSIBLE PERSON) ข้อ 8-15 ให้ครบถ้วนเพื่อส่ง LINE หา PCD")

                elif "PCD" in selected_dept:
                    # PCD กรอกข้อ 16-17 เสร็จ -> ส่งหา PRD
                    if check_range_completed(doc_no, 16, 17):
                        if send_next_dept_alert_line(doc_no, customer_name, part_name, "PRD"):
                            st.info("📲 บันทึกส่วน PCD ครบแล้ว! ส่ง LINE แจ้งเตือนแผนก PRD เรียบร้อย")
                    else:
                        st.warning("⚠️ บันทึกสำเร็จ แต่กรุณากรอกผู้รับผิดชอบ (RESPONSIBLE PERSON) ข้อ 16-17 ให้ครบถ้วนเพื่อส่ง LINE หา PRD")

                elif "PRD" in selected_dept:
                    # PRD กรอกข้อ 18-19 เสร็จ -> เช็กทั้ง 19 ข้อ ถ้าครบให้ยิงแจ้งเตือนรวมรออนุมัติ
                    if check_range_completed(doc_no, 1, 19):
                        if send_all_completed_alert_line(doc_no, customer_name, part_name):
                            st.info("🎉 กรอกข้อมูลครบทั้ง 19 ข้อแล้ว! ส่ง LINE แจ้งเตือนผู้จัดการเพื่อพิจารณาอนุมัติเรียบร้อย")
                    else:
                        st.warning("⚠️ บันทึกสำเร็จ แต่รายการทั้งหมด (ข้อ 1-19) ยังไม่ครบถ้วน จึงยังไม่ส่ง LINE แจ้งเตือนผู้จัดการ")