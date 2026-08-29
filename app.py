import os
import io
import re
import time
import pandas as pd
import streamlit as st
import google.generativeai as genai
from duckduckgo_search import DDGS
from pypdf import PdfReader
import docx
from email_auth import generate_otp, send_otp_email

# 구글 API 연동 패키지
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

OTP_EXPIRATION_SECONDS = 300

st.set_page_config(page_title="통합 비즈니스 분석 Gem", page_icon="📈", layout="wide")

# 다국어(I18N) 텍스트 사전
I18N = {
    "한국어": {
        "page_title": "📈 통합 비즈니스 분석 Gem",
        "login_title": "🔑 이메일 패스코드 로그인",
        "login_caption": "등록된 이메일로 패스코드를 발송받아 로그인하세요.",
        "email_label": "이메일 주소",
        "send_otp_btn": "패스코드 전송",
        "verify_otp_btn": "인증 및 로그인",
        "otp_input_label": "패스코드 6자리 입력",
        "reenter_email_btn": "이메일 다시 입력하기",
        "not_registered_err": "등록되지 않은 사용자입니다. (관리자에게 Secrets 등록을 요청하세요)",
        "otp_sent_msg": "패스코드가 발송되었습니다. 이메일을 확인하세요.",
        "otp_expired_err": "패스코드 유효시간(5분)이 만료되었습니다.",
        "otp_mismatch_err": "패스코드가 일치하지 않습니다.",
        "account_info": "👤 계정 정보",
        "current_user": "접속 계정",
        "logout_btn": "로그아웃",
        "ai_setting_header": "🤖 AI 모델 제어 설정",
        "model_select_label": "분석 제미나이 모델",
        "temp_slider_label": "답변 창의성 (Temperature)",
        "temp_help": "낮을수록 데이터 연산과 사실 기반 정밀 분석을, 높을수록 창의적인 마케팅 아이디어를 제시합니다.",
        "admin_header": "🛠️ 마케터 전용 사용자 관리",
        "admin_key_label": "관리자 마스터키",
        "admin_auth_success": "관리자 인증됨",
        "user_list_label": "현재 영구 등록된 사용자 목록:",
        "app_caption": "현재 선택된 AI 모델",
        "input_placeholder": "분석 질문을 입력하세요 (예: 2025~2026년 8월 한국, 중국, 일본의 모델별 실판매 수량 비교표와 핵심 이슈를 종합 보고서로 작성해 줘)",
        "status_start": "통합 데이터 파이프라인 가동 중...",
        "status_step1": "📊 1/3. 구글드라이브 시트 데이터 동기화 및 SQL 추출 중...",
        "status_step2": "📁 2/3. 구글드라이브 비즈니스 문서 파싱 및 분석 중...",
        "status_step3": "🌐 3/3. 외부 시장 동향 웹 검색 수행 중...",
        "status_complete": "데이터 수집 완료! 종합 보고서 작성 중...",
        "download_btn": "📄 마크다운 보고서 다운로드 (.md)",
        "system_instruction": "너는 기업 최고 경영진을 위해 데이터와 문서를 정밀 분석하는 수석 비즈니스 분석가야. 모든 답변은 정량적 실적과 정성적 이슈를 기반으로 논리적이고 명확한 마크다운 보고서로 작성해야 해.",
        "report_prompt_directive": "너는 최고 경영진을 위한 수석 비즈니스 분석가야. 아래 제공된 3가지 데이터를 종합하여 한국어로 완벽하고 정돈된 비즈니스 분석 보고서를 마크다운 형식으로 작성하라."
    },
    "English": {
        "page_title": "📈 Integrated Business Analysis Gem",
        "login_title": "🔑 Email Passcode Login",
        "login_caption": "Enter your registered email to receive a login passcode.",
        "email_label": "Email Address",
        "send_otp_btn": "Send Passcode",
        "verify_otp_btn": "Verify & Login",
        "otp_input_label": "Enter 6-digit Passcode",
        "reenter_email_btn": "Re-enter Email",
        "not_registered_err": "Unregistered user. Please ask the admin to add you in Secrets.",
        "otp_sent_msg": "Passcode sent. Please check your email inbox.",
        "otp_expired_err": "Passcode has expired (5-minute limit).",
        "otp_mismatch_err": "Passcode does not match.",
        "account_info": "👤 Account Info",
        "current_user": "Logged in as",
        "logout_btn": "Logout",
        "ai_setting_header": "🤖 AI Model Settings",
        "model_select_label": "Gemini Model",
        "temp_slider_label": "Creativity (Temperature)",
        "temp_help": "Lower values (0.0-0.2) focus on factual data/SQL precision; higher values generate creative marketing ideas.",
        "admin_header": "🛠️ Marketer Admin Console",
        "admin_key_label": "Admin Master Key",
        "admin_auth_success": "Admin Authorized",
        "user_list_label": "Permanently Registered Users:",
        "app_caption": "Active AI Model",
        "input_placeholder": "Enter analysis query (e.g., Create a sales performance table and key issue report for Korea, China, and Japan by model from 2025 to Aug 2026)",
        "status_start": "Running integrated data pipeline...",
        "status_step1": "📊 1/3. Syncing Drive sheets & executing SQL...",
        "status_step2": "📁 2/3. Parsing & analyzing Drive documents...",
        "status_step3": "🌐 3/3. Performing external web search...",
        "status_complete": "Data collection complete! Generating comprehensive report...",
        "download_btn": "📄 Download Markdown Report (.md)",
        "system_instruction": "You are a senior business analyst evaluating data and documents for C-level executives. Produce logical and clear Markdown reports based on quantitative results and qualitative issues.",
        "report_prompt_directive": "You are a senior business analyst for C-level executives. Synthesize the 3 data sources below and write a comprehensive business analysis report in English using Markdown formatting."
    }
}

# 서버 설정값(Secrets) 로드
try:
    GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
    SHEETS_FOLDER_ID = st.secrets["SHEETS_FOLDER_ID"]
    DOCS_FOLDER_ID = st.secrets["DOCS_FOLDER_ID"]
    RESEND_API_KEY = st.secrets["RESEND_API_KEY"]
    ADMIN_MASTER_KEY = st.secrets["ADMIN_MASTER_KEY"]
    
    # Secrets에서 ALLOWED_EMAILS를 가져와 리스트로 변환 (공백 제거)
    raw_emails = st.secrets.get("ALLOWED_EMAILS", "")
    ALLOWED_USER_LIST = [e.strip() for e in raw_emails.split(",") if e.strip()]
    
    # GCP 서비스 계정 정보 로드
    gcp_credentials = dict(st.secrets["gcp_service_account"])
except Exception as e:
    st.error("시스템 설정(Secrets)이 올바르게 완료되지 않았습니다.")
    st.stop()

# 등록된 사용자인지 Secrets를 기준으로 검사
def is_registered_user(email: str) -> bool:
    return email.strip() in ALLOWED_USER_LIST

genai.configure(api_key=GEMINI_API_KEY)

# 1. 사이드바 - 언어 선택
selected_lang = st.sidebar.radio("🌐 Language / 언어 선택", ["한국어", "English"], index=0)
L = I18N[selected_lang]

# 2. 사이드바 - 마케터 전용 사용자 관리 (Secrets에 등록된 영구 명단 확인용)
st.sidebar.markdown("---")
with st.sidebar.expander(L["admin_header"]):
    admin_key_input = st.text_input(L["admin_key_label"], type="password")
    if admin_key_input == ADMIN_MASTER_KEY:
        st.success(L["admin_auth_success"])
        st.markdown(f"**{L['user_list_label']}**")
        if ALLOWED_USER_LIST:
            for u in ALLOWED_USER_LIST:
                st.text(f"- {u}")
        else:
            st.text("등록된 사용자가 없습니다. Secrets에 ALLOWED_EMAILS를 추가하세요.")
        
        st.info("💡 사용자를 추가/삭제하려면 Streamlit Cloud의 [Settings] -> [Secrets] 메뉴에서 ALLOWED_EMAILS 항목을 수정하세요.")

# 3. 로그인 화면 처리
def login_screen():
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.title(L["login_title"])
        st.caption(L["login_caption"])

        if "otp_sent" not in st.session_state:
            st.session_state["otp_sent"] = False

        if not st.session_state["otp_sent"]:
            with st.form("request_otp"):
                email = st.text_input(L["email_label"], placeholder="user@company.com")
                submit = st.form_submit_button(L["send_otp_btn"], use_container_width=True)
                if submit:
                    if not is_registered_user(email):
                        st.error(L["not_registered_err"])
                    else:
                        otp_code = generate_otp()
                        is_success, err_msg = send_otp_email(RESEND_API_KEY, email, otp_code)
                        if is_success:
                            st.session_state["otp_sent"] = True
                            st.session_state["target_email"] = email
                            st.session_state["generated_otp"] = otp_code
                            st.session_state["otp_time"] = time.time()
                            st.success(L["otp_sent_msg"])
                            st.rerun()
                        else:
                            st.error(f"메일 발송 실패: {err_msg}")
        else:
            st.info(f"**{st.session_state['target_email']}**")
            with st.form("verify_otp"):
                user_otp = st.text_input(L["otp_input_label"], max_chars=6, type="password")
                verify = st.form_submit_button(L["verify_otp_btn"], use_container_width=True)
                if verify:
                    if time.time() - st.session_state["otp_time"] > OTP_EXPIRATION_SECONDS:
                        st.error(L["otp_expired_err"])
                        st.session_state["otp_sent"] = False
                    elif user_otp == st.session_state["generated_otp"]:
                        st.session_state["authenticated"] = True
                        st.session_state["user_email"] = st.session_state["target_email"]
                        st.rerun()
                    else:
                        st.error(L["otp_mismatch_err"])

            if st.button(L["reenter_email_btn"]):
                st.session_state["otp_sent"] = False
                st.rerun()

if not st.session_state.get("authenticated", False):
    login_screen()
    st.stop()

# 4. 로그인 성공 후 추가되는 사이드바 및 AI 설정
st.sidebar.title(L["account_info"])
st.sidebar.write(f"{L['current_user']}: **{st.session_state['user_email']}**")
if st.sidebar.button(L["logout_btn"], use_container_width=True):
    st.session_state["authenticated"] = False
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.header(L["ai_setting_header"])

selected_model_name = st.sidebar.selectbox(
    L["model_select_label"],
    ["models/gemini-1.5-pro", "models/gemini-1.5-flash", "models/gemini-2.0-flash"],
    index=0
)

temperature_val = st.sidebar.slider(
    L["temp_slider_label"],
    min_value=0.0,
    max_value=1.0,
    value=0.2,
    step=0.1,
    help=L["temp_help"]
)

generation_config = genai.GenerationConfig(
    temperature=temperature_val,
    top_p=0.95,
    max_output_tokens=8192
)

model = genai.GenerativeModel(
    model_name=selected_model_name,
    generation_config=generation_config,
    system_instruction=L["system_instruction"]
)

st.title(L["page_title"])
st.caption(f"{L['app_caption']}: **{selected_model_name}** (Temp: {temperature_val})")

# ========================================================
# Google Drive API (서비스 계정 연동 파이프라인)
# ========================================================

def parse_folder_id(input_str):
    if "drive.google.com" in input_str:
        match = re.search(r'folders/([a-zA-Z0-9_-]+)', input_str)
        if match: return match.group(1)
    return input_str.strip()

def get_drive_service():
    creds = service_account.Credentials.from_service_account_info(
        gcp_credentials,
        scopes=['https://www.googleapis.com/auth/drive.readonly']
    )
    return build('drive', 'v3', credentials=creds)

# [파이프라인 1] 다중 시트 동기화 및 SQL 추출
def fetch_and_load_multiple_sheets(folder_id, user_prompt):
    drive_service = get_drive_service()
    query = f"'{folder_id.strip()}' in parents and trashed = false"
    
    results = drive_service.files().list(
        q=query,
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
        fields="files(id, name, mimeType)"
    ).execute()
    files = results.get('files', [])

    conn = sqlite3.connect(':memory:')
    loaded_tables = []
    table_index = 1

    for f in files:
        file_id = f['id']
        file_name = f['name']
        mime_type = f['mimeType']
        file_name_clean = re.sub(r'[^a-zA-Z0-9_]', '_', os.path.splitext(file_name)[0])
        table_name = f"data_{table_index}_{file_name_clean[:20]}"

        try:
            if mime_type == 'application/vnd.google-apps.spreadsheet':
                request = drive_service.files().export_media(fileId=file_id, mimeType='text/csv')
            elif file_name.endswith(('.csv', '.xlsx', '.xls')):
                request = drive_service.files().get_media(fileId=file_id, supportsAllDrives=True)
            else:
                continue

            fh = io.BytesIO()
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done:
                _, done = downloader.next_chunk()
            fh.seek(0)
            
            if file_name.endswith(('.xlsx', '.xls')):
                df = pd.read_excel(fh)
            else:
                df = pd.read_csv(fh)

            df.to_sql(table_name, conn, index=False, if_exists='replace')
            cols = ", ".join(df.columns.tolist())
            loaded_tables.append(f"- Table [{table_name}] (File: {file_name}) / Columns: {cols}")
            table_index += 1
        except Exception:
            continue

    if not loaded_tables:
        raise Exception("Google Drive folder is empty or files could not be read. Ensure service account has viewer access.")

    schema_info = "\n".join(loaded_tables)
    sql_prompt = f"""
    As an SQLite expert, generate ONLY a valid SQL query based on the table schema below to answer the user request.
    
    [Schema]
    {schema_info}
    
    Request: {user_prompt}
    """
    sql_res = model.generate_content(sql_prompt).text.strip()
    clean_sql = sql_res.replace("```sql", "").replace("```", "").strip()
    
    result_df = pd.read_sql_query(clean_sql, conn)
    return clean_sql, result_df.to_markdown(index=False)

# [파이프라인 2] 문서 동기화 및 텍스트 추출
def fetch_and_extract_multi_docs(folder_id, user_prompt):
    drive_service = get_drive_service()
    query = f"'{folder_id.strip()}' in parents and trashed = false"
    
    results = drive_service.files().list(
        q=query,
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
        fields="files(id, name, mimeType)"
    ).execute()
    files = results.get('files', [])

    combined_text = ""

    for f in files:
        file_id = f['id']
        file_name = f['name']
        mime_type = f['mimeType']
        ext = os.path.splitext(file_name)[1].lower()

        file_content = ""
        try:
            if mime_type == 'application/vnd.google-apps.document':
                request = drive_service.files().export_media(fileId=file_id, mimeType='text/plain')
            else:
                request = drive_service.files().get_media(fileId=file_id, supportsAllDrives=True)

            fh = io.BytesIO()
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done:
                _, done = downloader.next_chunk()
            fh.seek(0)

            if mime_type == 'application/vnd.google-apps.document' or ext in ['.md', '.txt']:
                file_content = fh.getvalue().decode('utf-8', errors='ignore')
            elif ext == '.pdf':
                reader = PdfReader(fh)
                for page in reader.pages:
                    file_content += page.extract_text() + "\n"
            elif ext in ['.docx', '.doc']:
                doc = docx.Document(fh)
                file_content = "\n".join([p.text for p in doc.paragraphs])
        except Exception:
            continue

        if file_content.strip():
            combined_text += f"\n--- [Document: {file_name}] ---\n" + file_content

    if not combined_text:
        return "No business documents found."

    extract_prompt = f"""
    Summarize key issues, emails, and collaboration records relevant to the user request from the documents below.
    
    [Documents]
    {combined_text[:35000]}
    
    Request: {user_prompt}
    """
    return model.generate_content(extract_prompt).text

# [파이프라인 3] 웹 검색
def perform_web_search(query):
    try:
        results = DDGS().text(keywords=query, max_results=3)
        return "\n".join([f"- [{r['title']}]({r['href']}): {r['body']}" for r in results])
    except Exception as e:
        return f"Web search failed: {e}"

# UI 메인 대화창
if "messages" not in st.session_state:
    st.session_state.messages = []

for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

if prompt := st.chat_input(L["input_placeholder"]):
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    with st.chat_message("assistant"):
        status = st.status(L["status_start"], expanded=True)
        
        status.write(L["status_step1"])
        try:
            executed_sql, sheet_data = fetch_and_load_multiple_sheets(SHEETS_FOLDER_ID, prompt)
            status.write(f"└ SQL: `{executed_sql}`")
        except Exception as e:
            sheet_data = f"Sheet data extraction failed: {e}"

        status.write(L["status_step2"])
        docs_data = fetch_and_extract_multi_docs(DOCS_FOLDER_ID, prompt)

        status.write(L["status_step3"])
        web_data = perform_web_search(prompt)

        status.update(label=L["status_complete"], state="complete", expanded=False)

        report_prompt = f"""
        {L['report_prompt_directive']}

        [Source 1: Google Sheets SQL Data]
        {sheet_data}

        [Source 2: Business Documents Data]
        {docs_data}

        [Source 3: External Web Search]
        {web_data}

        User Request: {prompt}
        """
        report_md = model.generate_content(report_prompt).text
        st.markdown(report_md)
        st.session_state.messages.append({"role": "assistant", "content": report_md})

        st.download_button(
            label=L["download_btn"],
            data=report_md,
            file_name="Business_Analysis_Report.md",
            mime="text/markdown",
            use_container_width=True
        )
