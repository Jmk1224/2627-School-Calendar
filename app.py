import streamlit as st
import pandas as pd
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import io
from google.oauth2 import service_account
from googleapiclient.discovery import build

st.set_page_config(page_title="學校活動及排班管理系統", page_icon="📅", layout="wide")

st.title("🏫 學校活動及教師排班管理系統")

# -------------------------------------------------------------------
# Google Calendar Sync Helper Function
# -------------------------------------------------------------------
def add_to_google_calendar(event_data):
    try:
        creds_dict = st.secrets["gcp_service_account"]
        calendar_id = st.secrets["CALENDAR_ID"]
        
        scopes = ['https://www.googleapis.com/auth/calendar']
        creds = service_account.Credentials.from_service_account_info(creds_dict, scopes=scopes)
        service = build('calendar', 'v3', credentials=creds)

        teachers_str = "、".join([t.strip() for t in event_data['teachers'].split('、') if t.strip()])
        
        event_body = {
            'summary': f"{event_data['title']}（{teachers_str}）",
            'location': event_data['location'],
            'description': f"負責老師：{teachers_str}\n說明：{event_data['notes']}",
            'start': {
                'dateTime': f"{event_data['date']}T{event_data['start_time'].strftime('%H:%M:%S')}",
                'timeZone': 'Asia/Hong_Kong',
            },
            'end': {
                'dateTime': f"{event_data['date']}T{event_data['end_time'].strftime('%H:%M:%S')}",
                'timeZone': 'Asia/Hong_Kong',
            },
        }

        service.events().insert(calendarId=calendar_id, body=event_body).execute()
        return True, "成功同步至 Google Calendar！"
    except Exception as e:
        return False, f"Google Calendar 同步失敗: {str(e)}"

# Initialize Session State
if "activities" not in st.session_state:
    st.session_state.activities = []

# -------------------------------------------------------------------
# Sidebar: Add New Activity
# -------------------------------------------------------------------
st.sidebar.header("➕ 新增活動內容")
with st.sidebar.form("event_form", clear_on_submit=True):
    title = st.text_input("活動名稱 (例如: 測試日曆同步)")
    event_date = st.date_input("日期", value=datetime.today())
    col1, col2 = st.columns(2)
    with col1:
        start_time = st.time_input("開始時間", value=datetime.strptime("10:00", "%H:%M").time())
    with col2:
        end_time = st.time_input("結束時間", value=datetime.strptime("11:00", "%H:%M").time())
    location = st.text_input("地點", value="禮堂")
    teachers_raw = st.text_input("負責老師", value="校、明")
    notes = st.text_input("說明 / 備註", value="測試自動同步功能")

    submitted = st.form_submit_button("新增活動並同步")
    if submitted:
        if title and teachers_raw:
            new_id = max([a["id"] for a in st.session_state.activities], default=0) + 1
            teachers_fmt = teachers_raw.replace(" ", "、").replace(",", "、")
            
            event_obj = {
                "id": new_id,
                "title": title,
                "date": event_date,
                "start_time": start_time,
                "end_time": end_time,
                "location": location,
                "teachers": teachers_fmt,
                "notes": notes,
                "exported": True
            }
            
            # API Call to Google Calendar
            success, msg = add_to_google_calendar(event_obj)
            if success:
                st.session_state.activities.append(event_obj)
                st.sidebar.success(f"✅ {title} - {msg}")
                st.rerun()
            else:
                st.sidebar.error(f"❌ {msg}")
        else:
            st.sidebar.error("請填寫活動名稱及負責老師！")

# -------------------------------------------------------------------
# Main Content Table Display
# -------------------------------------------------------------------
st.subheader("📋 所有歷史及新增活動")
if st.session_state.activities:
    df = pd.DataFrame([
        {
            "序號": act["id"],
            "活動名稱": act["title"],
            "日期": act["date"].strftime("%Y-%m-%d"),
            "時間": f"{act['start_time'].strftime('%H:%M')} - {act['end_time'].strftime('%H:%M')}",
            "地點": act["location"],
            "負責老師": act["teachers"],
            "說明": act["notes"],
            "Google Calendar 狀態": "✅ 已即時同步"
        } for act in st.session_state.activities
    ])
    st.dataframe(df, use_container_width=True, hide_index=True)
else:
    st.info("目前沒有任何活動紀錄。請利用左側表單新增活動進行測試！")
