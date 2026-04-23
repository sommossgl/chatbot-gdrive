import streamlit as st
import anthropic
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import pandas as pd
from dotenv import load_dotenv
import os
import gspread
from google.oauth2.service_account import Credentials
from google.auth.transport.requests import Request

load_dotenv()

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

def get_credentials():
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        return Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    except Exception as e:
        return Credentials.from_service_account_file("credentials.json", scopes=SCOPES)

# def get_gspread_client():
#     creds = get_credentials()
#     return gspread.Client(auth=creds)

def get_gspread_client():
    creds = get_credentials()
    # refresh token ก่อนใช้งาน
    creds.refresh(Request())
    return gspread.Client(auth=creds)

def get_folder_id_from_url(url):
    if "folders/" in url:
        return url.split("folders/")[1].split("?")[0]
    return url

def scan_folder(folder_id):
    creds = get_credentials()
    service = build("drive", "v3", credentials=creds)
    results = service.files().list(
        q=f"'{folder_id}' in parents and trashed=false",
        fields="files(id, name, mimeType)"
    ).execute()
    return results.get("files", [])

def read_sheet(file_id):
    client = get_gspread_client()
    workbook = client.open_by_key(file_id)
    all_sheets = ""
    for sheet in workbook.worksheets():
        try:
            values = sheet.get_all_values()
            if len(values) > 0:
                text = "\n".join(["\t".join(row) for row in values])
                all_sheets += f"\n--- Tab: {sheet.title} ---\n{text}\n"
        except Exception as e:
            all_sheets += f"\n--- Tab: {sheet.title} --- [Error: {e}]\n"
    return all_sheets

def read_doc(file_id):
    creds = get_credentials()
    service = build("docs", "v1", credentials=creds)
    doc = service.documents().get(documentId=file_id).execute()
    content = ""
    for element in doc.get("body", {}).get("content", []):
        if "paragraph" in element:
            for part in element["paragraph"].get("elements", []):
                content += part.get("textRun", {}).get("content", "")
    return content

def read_all_files(folder_id):
    files = scan_folder(folder_id)
    all_data = {}
    errors = []
    for f in files:
        try:
            if f["mimeType"] == "application/vnd.google-apps.spreadsheet":
                content = read_sheet(f['id'])
                all_data[f["name"]] = f"[Google Sheet]\n{content}"
            elif f["mimeType"] == "application/vnd.google-apps.document":
                content = read_doc(f['id'])
                all_data[f["name"]] = f"[Google Doc]\n{content}"
        except Exception as e:
            errors.append(f"{f['name']}: {str(e)}")
            all_data[f["name"]] = f"[Error: {str(e)}]"
    return all_data, files, errors

def ask_claude(question, all_data):
    try:
        api_key = st.secrets["ANTHROPIC_API_KEY"]
    except:
        api_key = os.getenv("ANTHROPIC_API_KEY")

    client = anthropic.Anthropic(api_key=api_key)
    context = "\n\n".join([f"=== {name} ===\n{content}"
                           for name, content in all_data.items()])

    message = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=1024,
        messages=[{
            "role": "user",
            "content": f"""คุณเป็น AI assistant ที่ช่วยวิเคราะห์ข้อมูล ตอบในภาษาเดียวกับคำถาม

ข้อมูลจาก Google Drive:
{context}

คำถาม: {question}"""
        }]
    )
    return message.content[0].text

st.title("🤖 Chatbot จาก Google Drive")
st.markdown("---")

folder_input = st.text_input(
    "📁 วาง Google Drive Folder URL หรือ Folder ID:",
    placeholder="https://drive.google.com/drive/folders/xxxx"
)

if folder_input:
    folder_id = get_folder_id_from_url(folder_input.strip())

    with st.spinner("กำลังสแกนไฟล์ในโฟลเดอร์..."):
        try:
            all_data, files, errors = read_all_files(folder_id)

            st.success(f"พบไฟล์ทั้งหมด {len(files)} ไฟล์")

            with st.expander("ดูไฟล์ทั้งหมด"):
                for f in files:
                    icon = "📊" if "spreadsheet" in f["mimeType"] else "📄"
                    st.write(f"{icon} {f['name']}")

            if errors:
                with st.expander("❌ Errors"):
                    for e in errors:
                        st.error(e)

            with st.expander("🔍 Debug: ข้อมูลที่อ่านได้"):
                for name, content in all_data.items():
                    st.write(f"**{name}:** {len(content)} characters")
                    st.text(content[:300])

            if "messages" not in st.session_state:
                st.session_state.messages = []
            if "folder_id" not in st.session_state or st.session_state.folder_id != folder_id:
                st.session_state.messages = []
                st.session_state.folder_id = folder_id

            for msg in st.session_state.messages:
                with st.chat_message(msg["role"]):
                    st.write(msg["content"])

            if question := st.chat_input("ถามอะไรก็ได้เกี่ยวกับข้อมูลในโฟลเดอร์นี้..."):
                st.session_state.messages.append({"role": "user", "content": question})
                with st.chat_message("user"):
                    st.write(question)

                with st.chat_message("assistant"):
                    with st.spinner("กำลังคิด..."):
                        answer = ask_claude(question, all_data)
                        st.write(answer)
                        st.session_state.messages.append({"role": "assistant", "content": answer})

        except Exception as e:
            st.error(f"เกิดข้อผิดพลาด: {e}")
            st.exception(e)
else:
    st.info("กรุณาใส่ Google Drive Folder URL เพื่อเริ่มต้น")