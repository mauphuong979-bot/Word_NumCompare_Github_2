import csv
import os
import pandas as pd
import streamlit as st
from datetime import datetime, timedelta, timezone

# Use absolute path for Streamlit Cloud compatibility
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_FILE = os.path.join(BASE_DIR, "usage_log.csv")

def is_gsheet_configured():
    """Checks if Google Sheets connection is configured in st.secrets."""
    try:
        if hasattr(st, "secrets"):
            if "connections" in st.secrets and "gsheets" in st.secrets["connections"]:
                return True
            if "gsheets" in st.secrets:
                return True
    except Exception:
        pass
    return False

def get_logging_mode():
    """Returns a string indicating where logs are currently being saved."""
    if is_gsheet_configured():
        return "Google Sheets (Active)"
    return "CSV (Local Only)"

def _get_gsheet_url():
    """Extracts spreadsheet URL from secrets if present."""
    if "connections" in st.secrets and "gsheets" in st.secrets["connections"]:
        return st.secrets["connections"]["gsheets"].get("spreadsheet") or st.secrets["connections"]["gsheets"].get("url")
    if "gsheets" in st.secrets:
        return st.secrets["gsheets"].get("spreadsheet") or st.secrets["gsheets"].get("url")
    return None

def log_event(user, event_type, details):
    """Logs an event to the appropriate destination (CSV and/or GSheets)."""
    vn_tz = timezone(timedelta(hours=7))
    timestamp = datetime.now(vn_tz).strftime("%Y-%m-%d %H:%M:%S")
    
    if is_gsheet_configured():
        try:
            from streamlit_gsheets import GSheetsConnection
            ss_url = _get_gsheet_url()
            conn = st.connection("gsheets", type=GSheetsConnection)
            df = conn.read(spreadsheet=ss_url, ttl=0)
            
            new_row = pd.DataFrame([{
                "Timestamp": timestamp,
                "User": user,
                "Event Type": event_type,
                "Details": details
            }])
            
            updated_df = new_row if df.empty else pd.concat([df, new_row], ignore_index=True)
            conn.update(spreadsheet=ss_url, data=updated_df)
            st.session_state["last_log_status"] = "Success: Logged to Google Sheets & CSV"
        except Exception as e:
            print(f"CRITICAL ERROR: Failed to log to GSheet: {e}")
            st.session_state["last_log_status"] = f"Error: {str(e)}"
    else:
        st.session_state["last_log_status"] = "Success: Logged to local CSV"

    # Always write to local CSV fallback
    log_to_csv_fallback(timestamp, user, event_type, details)

def log_to_csv_fallback(timestamp, user, event_type, details):
    """Fallback helper for CSV logging."""
    file_exists = os.path.isfile(LOG_FILE)
    try:
        with open(LOG_FILE, mode='a', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            if not file_exists:
                writer.writerow(["Timestamp", "User", "Event Type", "Details"])
            writer.writerow([timestamp, user, event_type, details])
    except Exception as e:
        print(f"Error logging event to CSV: {e}")

def get_logs():
    """Reads all logs from the appropriate source."""
    if is_gsheet_configured():
        try:
            from streamlit_gsheets import GSheetsConnection
            ss_url = _get_gsheet_url()
            conn = st.connection("gsheets", type=GSheetsConnection)
            df = conn.read(spreadsheet=ss_url, ttl=0)
            if not df.empty:
                return df.to_dict('records')[::-1]
        except Exception as e:
            print(f"Error reading from GSheet: {e}")

    # Fallback to local CSV
    if not os.path.exists(LOG_FILE):
        return []
    
    try:
        logs = []
        with open(LOG_FILE, mode='r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                logs.append(row)
        return logs[::-1]
    except Exception:
        return []
