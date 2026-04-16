import csv
import os
from datetime import datetime

# Use absolute path for Streamlit Cloud compatibility
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_FILE = os.path.join(BASE_DIR, "usage_log.csv")

def log_event(user, event_type, details):
    """Logs an event to the usage_log.csv file."""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    file_exists = os.path.isfile(LOG_FILE)
    
    try:
        with open(LOG_FILE, mode='a', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            # Write header if file is new
            if not file_exists:
                writer.writerow(["Timestamp", "User", "Event Type", "Details"])
            
            writer.writerow([timestamp, user, event_type, details])
    except Exception as e:
        # Fail silently in the UI but print for debugging
        print(f"Error logging event: {e}")

def get_logs():
    """Reads all logs from the CSV file."""
    if not os.path.exists(LOG_FILE):
        return []
    
    try:
        logs = []
        with open(LOG_FILE, mode='r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                logs.append(row)
        return logs[::-1]  # Return in reverse chronological order
    except Exception:
        return []
