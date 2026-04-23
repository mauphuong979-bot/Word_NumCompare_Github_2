import streamlit as st
import pandas as pd
import io
import json
import os
import zipfile
import tempfile
import shutil
import platform
import subprocess
from extractor import extract_table_data
from processor import compare_dataframes
from usage_logger import log_event, get_logs
from datetime import datetime, timedelta, timezone

# Conditional imports for Windows-specific PDF conversion
IS_WINDOWS = platform.system() == "Windows"
if IS_WINDOWS:
    try:
        import pythoncom
        from docx2pdf import convert as docx2pdf_convert
    except ImportError:
        IS_WINDOWS = False

# Page Configuration
st.set_page_config(
    page_title="Professional Document Suite",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Load CSS
with open("style.css") as f:
    st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

# Authentication Logic
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
USERS_FILE = os.path.join(BASE_DIR, "users.json")

def load_users():
    if os.path.exists(USERS_FILE):
        with open(USERS_FILE, "r") as f:
            return json.load(f)["users"]
    return []

def check_credentials(username, password):
    users = load_users()
    for user in users:
        if user["username"] == username and user["password"] == password:
            return True, user.get("role", "user")
    return False, None

def save_user(username, password, role, auto_fill=False):
    users = load_users()
    if any(u["username"] == username for u in users):
        return False, f"Username '{username}' already exists."
    
    users.append({
        "username": username,
        "password": password,
        "role": role,
        "auto_fill": auto_fill
    })
    
    try:
        with open(USERS_FILE, "w") as f:
            json.dump({"users": users}, f, indent=4)
        return True, f"User '{username}' added successfully."
    except Exception as e:
        return False, f"Error saving user: {e}"

def remove_user(username):
    users = load_users()
    updated_users = [u for u in users if u["username"] != username]
    try:
        with open(USERS_FILE, "w") as f:
            json.dump({"users": updated_users}, f, indent=4)
        return True, f"User '{username}' deleted successfully."
    except Exception as e:
        return False, f"Error deleting user: {e}"

def update_user_data(old_username, new_username, new_password, new_role, auto_fill):
    users = load_users()
    if old_username != new_username:
        if any(u["username"] == new_username for u in users):
            return False, f"Username '{new_username}' already exists."
    
    for user in users:
        if user["username"] == old_username:
            user["username"] = new_username
            user["password"] = new_password
            user["role"] = new_role
            user["auto_fill"] = auto_fill
            break
            
    try:
        with open(USERS_FILE, "w") as f:
            json.dump({"users": users}, f, indent=4)
        return True, f"User '{new_username}' updated successfully."
    except Exception as e:
        return False, f"Error updating user: {e}"

if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False

def handle_autofill():
    users = load_users()
    selected_user = next((u for u in users if u["username"] == st.session_state.login_user), None)
    if selected_user and selected_user.get("auto_fill"):
        st.session_state.login_password = selected_user["password"]
    else:
        st.session_state.login_password = ""

def login_screen():
    st.markdown('<div class="login-container">', unsafe_allow_html=True)
    st.markdown('<div class="login-header">🔐 User Login</div>', unsafe_allow_html=True)
    
    users = load_users()
    usernames = [u["username"] for u in users]
    
    if "login_user" not in st.session_state:
        st.session_state.login_user = "user" if "user" in usernames else (usernames[0] if usernames else "")
        handle_autofill()

    st.selectbox("Select Username", usernames, key="login_user", on_change=handle_autofill)
    
    with st.form("login_form"):
        password = st.text_input("Password", key="login_password", type="password")
        submit = st.form_submit_button("Sign In")
        
        if submit:
            is_valid, role = check_credentials(st.session_state.login_user, password)
            if is_valid:
                st.session_state.authenticated = True
                st.session_state.username = st.session_state.login_user
                st.session_state.role = role
                log_event(st.session_state.username, "Login", "Successfully signed in")
                st.rerun()
            else:
                st.error("Invalid username or password")
    
    st.markdown('</div>', unsafe_allow_html=True)

if not st.session_state.authenticated:
    login_screen()
    st.stop()

# --- GLOBAL SIDEBAR (Authenticated Only) ---
with st.sidebar:
    st.markdown(f"👤 **Logged in as:** {st.session_state.username}")
    if st.button("🚪 Logout", use_container_width=True):
        st.session_state.authenticated = False
        st.rerun()
    st.divider()
    st.caption(f"v2.1 ({platform.system()} Edition)")

# App UI Header
st.markdown('<div class="app-logo">📄</div>', unsafe_allow_html=True)
st.markdown('<div class="main-header">Professional Document Suite</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Data Comparison & Document Conversion Utilities</div>', unsafe_allow_html=True)

# Main Tab Navigation
tab_titles = ["📊 Number Comparison", "📑 Word to PDF Converter"]
if st.session_state.authenticated and st.session_state.get('role') == 'admin':
    tab_titles.append("🛡️ Admin")

tabs = st.tabs(tab_titles)
tab_compare = tabs[0]
tab_pdf = tabs[1]
if len(tabs) > 2:
    tab_admin = tabs[2]
else:
    tab_admin = None

# --- TAB 1: NUMBER COMPARISON ---
with tab_compare:
    # 1. Compact Horizontal Settings Bar at the Top
    st.markdown('<div class="settings-panel">', unsafe_allow_html=True)
    st.markdown("### ⚙️ Settings")
    
    set_col1, set_col2, set_col3, set_col4 = st.columns([1.5, 1, 1, 1.5], gap="medium")
    
    with set_col1:
        st.markdown('<div class="settings-label">1. Number Formats</div>', unsafe_allow_html=True)
        s_col1, s_col2 = st.columns(2)
        with s_col1:
            st.caption("Doc 1")
            format_1 = st.radio("F1", ["Vietnam", "US"], index=0, key="f1_fmt", label_visibility="collapsed")
        with s_col2:
            st.caption("Doc 2")
            format_2 = st.radio("F2", ["Vietnam", "US"], index=1, key="f2_fmt", label_visibility="collapsed")
    
    with set_col2:
        st.markdown('<div class="settings-label">2. Target</div>', unsafe_allow_html=True)
        extract_mode = st.radio("Target", ["Number", "Text"], index=0, horizontal=True, label_visibility="collapsed")
    
    with set_col3:
        st.markdown('<div class="settings-label">3. Filter</div>', unsafe_allow_html=True)
        view_mode = st.radio("Filter", ["Mismatches", "All Results"], index=0, label_visibility="collapsed")

    with set_col4:
        st.info("💡 Header check suggested first.")
    st.markdown('</div>', unsafe_allow_html=True)

    # 2. Main Logic Area (Full Width)
    col_c1, col_c2 = st.columns(2)
    with col_c1:
        st.markdown('<div class="converter-card">', unsafe_allow_html=True)
        st.info("📂 **Upload Document 1**")
        file1 = st.file_uploader("Select first .docx file", type=["docx"], key="doc1")
        st.markdown('</div>', unsafe_allow_html=True)

    with col_c2:
        st.markdown('<div class="converter-card">', unsafe_allow_html=True)
        st.info("📂 **Upload Document 2**")
        file2 = st.file_uploader("Select second .docx file", type=["docx"], key="doc2")
        st.markdown('</div>', unsafe_allow_html=True)

    if file1 and file2:
        if st.button("🚀 Run Comparison", use_container_width=True):
            with st.spinner("Extracting and comparing data..."):
                try:
                    df1 = extract_table_data(file1, format_1, mode=extract_mode)
                    df2 = extract_table_data(file2, format_2, mode=extract_mode)
                    
                    if df1.empty or df2.empty:
                        st.warning(f"One of the documents contains no {extract_mode.lower()} table data.")
                    else:
                        merged_df, result_msg, mismatch_count = compare_dataframes(df1, df2, mode=extract_mode)
                        log_event(st.session_state.username, "Comparison", f"Files: {file1.name}, {file2.name} | Mode: {extract_mode}")
                        
                        st.divider()
                        m_col1, m_col2, m_col3 = st.columns(3)
                        m_col1.metric("Rows in Doc 1", len(df1))
                        m_col2.metric("Rows in Doc 2", len(df2))
                        delta_color = "normal" if mismatch_count == 0 else "inverse"
                        m_col3.metric("Mismatches", mismatch_count, delta=mismatch_count, delta_color=delta_color)
                        
                        if mismatch_count == 0:
                            st.success("Verification successful! All values match perfectly.")
                        else:
                            st.error(f"Found {mismatch_count} discrepancies. Please review the table below.")
                            st.markdown("""
                                <div class="word-hint">
                                    <b>💡 Tip: Quickly Find Tables in MS Word</b><br>
                                    1. Press <b>Ctrl + G</b> to open 'Go To' dialog.<br>
                                    2. Select <b>Table</b> in the list.<br>
                                    3. Enter the <b>Table Number</b> from the list below.<br>
                                    4. Click <b>Go To</b> to jump directly to the table.
                                </div>
                            """, unsafe_allow_html=True)
                        
                        display_df = merged_df.copy()
                        if extract_mode == 'Number':
                            mismatch_mask = display_df['Diff'].abs() > 1e-6
                        else:
                            mismatch_mask = display_df['Text 1'].fillna("") != display_df['Text 2'].fillna("")
                        
                        if view_mode == "Mismatches":
                            display_df = display_df[mismatch_mask]
                        
                        if extract_mode == 'Text':
                            display_df = display_df.drop(columns=['Value 1', 'Value 2', 'Diff'])
                        
                        def highlight_diff(row):
                            if extract_mode == 'Number':
                                is_diff = abs(row['Diff']) > 1e-6
                            else:
                                is_diff = str(row['Text 1']) != str(row['Text 2'])
                            return ['background-color: #fef2f2' if is_diff else '' for _ in row]

                        styled_df = display_df.style.apply(highlight_diff, axis=1)
                        if extract_mode == 'Number':
                            styled_df = styled_df.format({'Value 1': "{:,.2f}", 'Value 2': "{:,.2f}", 'Diff': "{:,.2f}"})
                        
                        st.dataframe(styled_df, use_container_width=True, height=600)
                        
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            merged_df.to_excel(writer, index=False, sheet_name='Comparison')
                        st.download_button(
                            label="📥 Download Comparison Report (Excel)",
                            data=output.getvalue(),
                            file_name="Comparison_Report.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                except Exception as e:
                    st.error(f"Error during analysis: {str(e)}")
    else:
        st.info("Please upload Word documents to start.")

# --- TAB 2: PDF CONVERTER ---
with tab_pdf:
    st.markdown("""
        <div class="guide-box">
            <h4>📖 Multi-Platform PDF Conversion</h4>
            <ul>
                <li><b>Windows:</b> Uses Microsoft Word engine (Highest fidelity).</li>
                <li><b>Cloud/Linux:</b> Uses LibreOffice engine (Cross-platform compatibility).</li>
                <li>Attach file(s), convert, and download individually or as a ZIP.</li>
            </ul>
        </div>
    """, unsafe_allow_html=True)

    uploaded_files = st.file_uploader(
        "Select Word (.docx) files to convert", 
        type=["docx"], 
        accept_multiple_files=True,
        key="pdf_uploader"
    )

    if uploaded_files:
        st.markdown('<div class="converter-card">', unsafe_allow_html=True)
        st.subheader(f"Batch Processing ({len(uploaded_files)} files)")
        
        if st.button("✨ Convert to PDF", use_container_width=True):
            progress_bar = st.progress(0)
            status_placeholder = st.empty()
            converted_files = []
            
            temp_dir = tempfile.mkdtemp()
            # Handle COM initialization for Windows
            if IS_WINDOWS:
                pythoncom.CoInitialize()

            try:
                for i, uploaded_file in enumerate(uploaded_files):
                    status_placeholder.info(f"Processing: {uploaded_file.name}...")
                    
                    input_path = os.path.join(temp_dir, uploaded_file.name)
                    with open(input_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())
                    
                    output_filename = os.path.splitext(uploaded_file.name)[0] + ".pdf"
                    output_path = os.path.join(temp_dir, output_filename)
                    
                    try:
                        if IS_WINDOWS:
                            # Use docx2pdf for Windows
                            docx2pdf_convert(input_path, output_path)
                        else:
                            # Use LibreOffice for Linux
                            subprocess.run([
                                'libreoffice', '--headless', 
                                '--convert-to', 'pdf', 
                                '--outdir', temp_dir, 
                                input_path
                            ], check=True, capture_output=True)
                        
                        if os.path.exists(output_path):
                            with open(output_path, "rb") as f:
                                pdf_data = f.read()
                            
                            converted_files.append({
                                "name": output_filename,
                                "data": pdf_data,
                                "status": "success"
                            })
                        else:
                            raise Exception("Output PDF file was not generated.")
                            
                    except Exception as e:
                        st.error(f"Failed to convert {uploaded_file.name}: {str(e)}")
                        converted_files.append({
                            "name": uploaded_file.name,
                            "status": "error",
                            "error": str(e)
                        })
                    
                    progress_bar.progress((i + 1) / len(uploaded_files))
                
                status_placeholder.success(f"Successfully processed {len(converted_files)} files!")
                st.session_state.converted_files_data = converted_files
                log_event(st.session_state.username, "PDF Conversion", f"Batch of {len(uploaded_files)} files processed")
                
            finally:
                if IS_WINDOWS:
                    pythoncom.CoUninitialize()
                shutil.rmtree(temp_dir)

        if 'converted_files_data' in st.session_state and st.session_state.converted_files_data:
            st.divider()
            st.markdown("### ⬇️ Download Your Files")
            
            for file_info in st.session_state.converted_files_data:
                col_name, col_status, col_btn = st.columns([4, 1, 2])
                with col_name:
                    st.text(f"📄 {file_info['name']}")
                with col_status:
                    if file_info['status'] == "success":
                        st.markdown('<span class="status-badge status-success">Ready</span>', unsafe_allow_html=True)
                    else:
                        st.markdown('<span class="status-badge status-error">Failed</span>', unsafe_allow_html=True)
                with col_btn:
                    if file_info['status'] == "success":
                        st.download_button(
                            label="Download",
                            data=file_info['data'],
                            file_name=file_info['name'],
                            mime="application/pdf",
                            key=f"dl_{file_info['name']}"
                        )
            
            success_files = [f for f in st.session_state.converted_files_data if f['status'] == "success"]
            if len(success_files) > 0:
                st.divider()
                st.markdown("#### Package all files")
                
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                    for f in success_files:
                        zf.writestr(f['name'], f['data'])
                
                st.download_button(
                    label="📥 Download All (ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name="Converted_PDFs.zip",
                    mime="application/zip",
                    use_container_width=True
                )
        st.markdown('</div>', unsafe_allow_html=True)
    else:
        st.info("Upload one or more Word documents to get started.")

# --- TAB 3: ADMIN (Admin Only) ---
if tab_admin:
    with tab_admin:
        admin_tabs = st.tabs(["📊 Usage Analytics", "👥 User Control"])
        
        # --- SUB-TAB: ANALYTICS & LOGS ---
        with admin_tabs[0]:
            st.markdown("### 📊 System Usage & Activity")
            
            logs = get_logs()
            if logs:
                log_df = pd.DataFrame(logs)
                
                # 1. Dashboard Metrics
                m1, m2, m3, m4 = st.columns(4)
                total_events = len(log_df)
                unique_users = log_df['User'].nunique()
                most_active = log_df['User'].value_counts().idxmax() if not log_df.empty else "N/A"
                last_24h = len(log_df) # Placeholder for real filter if needed
                
                m1.metric("Total Events", total_events)
                m2.metric("Unique Users", unique_users)
                m3.metric("Top User", most_active)
                m4.metric("Recent (All)", total_events)
                
                st.divider()
                
                # 2. Advanced Filters
                st.markdown("#### 🔍 Filter Logs")
                f_col1, f_col2, f_col3 = st.columns([2, 1, 1])
                
                with f_col1:
                    search_query = st.text_input("Search details or users...", placeholder="e.g. login, user1, comparison")
                with f_col2:
                    event_filter = st.selectbox("Event Type", ["All"] + list(log_df['Event Type'].unique()))
                with f_col3:
                    row_limit = st.selectbox("Display Limit", [100, 500, 1000, "All"], index=1)

                # Applying filters
                filtered_df = log_df.copy()
                if search_query:
                    filtered_df = filtered_df[
                        filtered_df['Details'].str.contains(search_query, case=False, na=False) | 
                        filtered_df['User'].str.contains(search_query, case=False, na=False)
                    ]
                if event_filter != "All":
                    filtered_df = filtered_df[filtered_df['Event Type'] == event_filter]
                
                if row_limit != "All":
                    filtered_df = filtered_df.head(int(row_limit))

                # 3. Enhanced Log Display
                st.markdown(f"Showing **{len(filtered_df)}** entries")
                st.dataframe(filtered_df, use_container_width=True, height=600)
                
                # 4. Actions
                output_log = io.BytesIO()
                with pd.ExcelWriter(output_log, engine='openpyxl') as writer:
                    filtered_df.to_excel(writer, index=False, sheet_name='FilteredLogs')
                
                vn_tz = timezone(timedelta(hours=7))
                st.download_button(
                    label="📥 Export Current View to Excel",
                    data=output_log.getvalue(),
                    file_name=f"System_Log_Export_{datetime.now(vn_tz).strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            else:
                st.info("No system logs available yet.")

        # --- SUB-TAB: USER MANAGEMENT ---
        with admin_tabs[1]:
            st.markdown("### 👥 User Account Management")
            
            uc_col1, uc_col2 = st.columns([1, 2], gap="large")
            
            with uc_col1:
                st.markdown('<div class="converter-card">', unsafe_allow_html=True)
                st.markdown("#### ➕ Create New Account")
                with st.form("add_user_form_v2", clear_on_submit=True):
                    new_user = st.text_input("Username")
                    new_pass = st.text_input("Password", type="password")
                    new_role = st.selectbox("Role", ["user", "admin"])
                    new_auto_fill = st.checkbox("Enable Auto-fill", value=False)
                    if st.form_submit_button("Register User", use_container_width=True):
                        if new_user and new_pass:
                            success_add, msg_add = save_user(new_user, new_pass, new_role, new_auto_fill)
                            if success_add: st.success(msg_add)
                            else: st.error(msg_add)
                st.markdown('</div>', unsafe_allow_html=True)

            with uc_col2:
                st.markdown("#### 📋 Existing Users")
                current_users_list = load_users()
                
                for idx, user_data in enumerate(current_users_list):
                    is_admin_account = (user_data["username"] == "admin")
                    with st.expander(f"👤 {user_data['username']} — Role: {user_data['role'].upper()}", expanded=False):
                        e_col1, e_col2 = st.columns(2)
                        with e_col1:
                            edit_name = st.text_input("Username", value=user_data["username"], disabled=is_admin_account, key=f"v2_un_{idx}")
                            edit_pass = st.text_input("Password", value=user_data["password"], type="password", key=f"v2_pw_{idx}")
                        with e_col2:
                            edit_role = st.selectbox("Role", ["admin", "user"], index=0 if user_data["role"]=="admin" else 1, key=f"v2_rl_{idx}")
                            edit_auto = st.checkbox("Auto-fill", value=user_data.get("auto_fill", False), key=f"v2_af_{idx}")
                        
                        st.divider()
                        ec1, ec2 = st.columns(2)
                        with ec1:
                            if st.button("💾 Save Changes", key=f"v2_upd_{idx}", use_container_width=True):
                                success_upd, msg_upd = update_user_data(user_data["username"], edit_name, edit_pass, edit_role, edit_auto)
                                if success_upd: st.rerun()
                                else: st.error(msg_upd)
                        with ec2:
                            if is_admin_account:
                                st.button("🚫 System Admin", disabled=True, key=f"v2_del_dis_{idx}", use_container_width=True)
                            else:
                                if st.button("🗑️ Delete User", key=f"v2_del_{idx}", use_container_width=True):
                                    remove_user(user_data["username"])
                                    st.rerun()

# Footer
st.divider()
st.caption("Professional Document Suite | Advanced Streamlit Application")
