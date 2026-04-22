import streamlit as st
import anthropic
import gspread
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv
import os
import pandas as pd

load_dotenv()

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

def connect_google_drive():
    creds = Credentials.from_service_account_file("credentials.json", scopes=SCOPES)
    client = gspread.authorize(creds)
    return client

def get_sheet_data(sheet_url):
    client = connect_google_drive()
    sheet = client.open_by_url(sheet_url).sheet1
    data = sheet.get_all_records()
    return pd.DataFrame(data)

def ask_claude(question, data_context):
    client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))
    message = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=1024,
        messages=[
            {
                "role": "user",
                "content": f"""You are a helpful assistant. Answer in the same language as the question.
                
Here is the data from Google Sheets:
{data_context}

Question: {question}"""
            }
        ]
    )
    return message.content[0].text

st.title("🤖 Chatbot จาก Google Sheets")
st.markdown("---")

sheet_url = st.text_input("📎 วาง Google Sheets URL ที่นี่:", placeholder="https://docs.google.com/spreadsheets/d/...")

if sheet_url:
    with st.spinner("กำลังโหลดข้อมูล..."):
        try:
            df = get_sheet_data(sheet_url)
            st.success(f"โหลดข้อมูลสำเร็จ — {len(df)} แถว, {len(df.columns)} คอลัมน์")
            with st.expander("ดูข้อมูล"):
                st.dataframe(df)
            
            if "messages" not in st.session_state:
                st.session_state.messages = []

            for msg in st.session_state.messages:
                with st.chat_message(msg["role"]):
                    st.write(msg["content"])

            if question := st.chat_input("ถามอะไรก็ได้เกี่ยวกับข้อมูลนี้..."):
                st.session_state.messages.append({"role": "user", "content": question})
                with st.chat_message("user"):
                    st.write(question)

                with st.chat_message("assistant"):
                    with st.spinner("กำลังคิด..."):
                        answer = ask_claude(question, df.to_string())
                        st.write(answer)
                        st.session_state.messages.append({"role": "assistant", "content": answer})

        except Exception as e:
            st.error(f"เกิดข้อผิดพลาด: {e}")
else:
    st.info("กรุณาใส่ Google Sheets URL เพื่อเริ่มต้น")