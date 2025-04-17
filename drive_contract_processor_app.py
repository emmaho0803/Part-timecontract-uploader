#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import streamlit as st
import gspread
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
from datetime import datetime
import re
import smtplib
from email.mime.text import MIMEText

scopes = st.secrets["config"]["SCOPES"].split(",")
creds = Credentials.from_service_account_info(
    st.secrets["gcp_service_account"],
    scopes=scopes
)

# 設定
SHEET_NAME = st.secrets["config"]["SHEET_NAME"]
DRIVE_FOLDER_ID = st.secrets["config"]["DRIVE_FOLDER_ID"]
GMAIL_USER = st.secrets["email"]["GMAIL_USER"]
GMAIL_APP_PASSWORD = st.secrets["email"]["GMAIL_APP_PASSWORD"]
TO_EMAIL = st.secrets["email"]["TO_EMAIL"]

# ========== 函式區 ==========
def parse_date(datestr):
    for fmt in ("%Y-%m-%d", "%Y/%m/%d"):
        try:
            return datetime.strptime(datestr, fmt).date()
        except ValueError:
            continue
    raise ValueError(f"❌ 無法解析日期格式：{datestr}")

def check_and_send_reminders(sheet, today, gmail_user, app_password, to_email):
    data = sheet.get_all_records()
    reminder_msgs = []

    for i, row in enumerate(data):
        contract_name = row["合約名稱"]
        due_date_str = row["到期日"]
        received = row["已收金額"]
        should_receive = row["應收回饋金"]

        if not due_date_str:
            continue

        due_date = parse_date(due_date_str)

        if today > due_date and received < should_receive:
            msg = f"【催帳提醒】\n合約：「{contract_name}」已於 {due_date} 到期未全額收款\n應收：{should_receive} 元，已收：{received} 元"
            reminder_msgs.append(msg)

            sheet.update_cell(i + 2, 11, "是")
            sheet.update_cell(i + 2, 12, today.strftime("%Y-%m-%d"))

    if reminder_msgs:
        email_body = "\n\n".join(reminder_msgs)
        msg = MIMEText(email_body)
        msg["Subject"] = "📬 催帳提醒通知"
        msg["From"] = gmail_user
        msg["To"] = to_email

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(gmail_user, app_password)
            server.send_message(msg)
        return True, reminder_msgs
    return False, []

def parse_contract_filename(file_name):
    name_no_ext = file_name.replace(".pdf", "")
    title, raw_info = name_no_ext.split("__", 1)
    target_info, budget_info, date_range = raw_info.split("_")
    partner, contact = target_info.split("&")

    match = re.match(r"(\d+)\((\d+)%\)", budget_info)
    if match:
        amount = int(match.group(1))
        percent = int(match.group(2)) / 100
    else:
        raise ValueError("金額格式錯誤")

    start_str, end_str = date_range.split("-")
    start_date = datetime.strptime(start_str, "%Y%m%d").strftime("%Y-%m-%d")
    end_date = datetime.strptime(end_str, "%Y%m%d").strftime("%Y-%m-%d")

    return {
        "合約名稱": title,
        "對象": partner,
        "合作人": contact,
        "回饋金%": percent,
        "經費金額": amount,
        "應收回饋金": round(amount * percent),
        "已收金額": 0,
        "簽約日期": start_date,
        "到期日": end_date
    }

def process_drive_folder(folder_id, sheet, drive_service):
    results = drive_service.files().list(
        q=f"'{folder_id}' in parents and mimeType='application/pdf'",
        fields="files(id, name)").execute()
    files = results.get('files', [])
    existing_files = sheet.get_all_records()
    existing_pdf_ids = [
        row["PDF連結"].split("/d/")[1].split("/")[0]
        for row in existing_files if row["PDF連結"]
    ]

    new_files = 0
    for file in files:
        if file['id'] in existing_pdf_ids:
            continue
        try:
            parsed = parse_contract_filename(file['name'])
            drive_service.permissions().create(
                fileId=file['id'],
                body={'type': 'anyone', 'role': 'reader'},
                fields='id'
            ).execute()
            file_url = f"https://drive.google.com/file/d/{file['id']}/view?usp=sharing"
            row_data = [
                parsed["合約名稱"],
                parsed["對象"],
                parsed["合作人"],
                parsed["簽約日期"],
                f'{int(parsed["回饋金%"] * 100)}%',
                parsed["經費金額"],
                parsed["應收回饋金"],
                parsed["已收金額"],
                parsed["到期日"],
                file_url,
                "",
                ""
            ]
            sheet.append_row(row_data)
            new_files += 1
        except Exception as e:
            st.warning(f"❌ 錯誤：{file['name']} → {e}")
    return new_files

# ========== Streamlit UI ==========

st.set_page_config(page_title="合約上傳與催帳系統", layout="centered")
st.title("📁 合約 PDF 上傳與 📬 催帳提醒工具")

if st.button("🚀 開始處理合約 PDF 並寫入 Sheet"):
    creds = Credentials.from_service_account_file(CREDENTIAL_FILE, scopes=SCOPES)
    gc = gspread.authorize(creds)
    sheet = gc.open(SHEET_NAME).sheet1
    drive_service = build('drive', 'v3', credentials=creds)
    count = process_drive_folder(DRIVE_FOLDER_ID, sheet, drive_service)
    st.success(f"✅ 共寫入 {count} 份新合約！")

if st.button("📬 發送催帳提醒 Email"):
    creds = Credentials.from_service_account_file(CREDENTIAL_FILE, scopes=SCOPES)
    gc = gspread.authorize(creds)
    sheet = gc.open(SHEET_NAME).sheet1
    today = datetime.today().date()
    success, msgs = check_and_send_reminders(sheet, today, GMAIL_USER, GMAIL_APP_PASSWORD, TO_EMAIL)
    if success:
        st.success(f"✅ 已發送催帳通知，共 {len(msgs)} 筆！")
        for m in msgs:
            st.text(m)
    else:
        st.info("✅ 今日無需催帳通知！")

