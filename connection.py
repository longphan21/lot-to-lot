import streamlit as st
import gspread

DEFAULT_SHEET_URL = "https://docs.google.com/spreadsheets/d/1DQIDnYgzcvmkmrlL3ohTH6SS7S19MWWVisQlT1Aq-f0/edit?gid=0#gid=0"


@st.cache_resource(show_spinner=False)
def get_gsheet_client():
    """Rẽ nhánh xác thực: Ưu tiên Streamlit Cloud Secrets, fallback về file credentials.json ở Local."""
    if "gcp_service_account" in st.secrets:
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        return gspread.service_account_from_dict(creds_dict)
    return gspread.service_account(filename="credentials.json")


@st.cache_resource(show_spinner=False)
def get_spreadsheet():
    """Mở và giữ kết nối tới file Google Sheets (chỉ gọi API mở file 1 lần duy nhất)."""
    client = get_gsheet_client()
    sheet_url = st.secrets.get("SHEET_URL", DEFAULT_SHEET_URL) if hasattr(st, "secrets") else DEFAULT_SHEET_URL
    return client.open_by_url(sheet_url)