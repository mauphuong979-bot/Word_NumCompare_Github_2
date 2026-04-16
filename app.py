import streamlit as st
import pandas as pd
import io
import json
import os
from extractor import extract_table_data
from processor import compare_dataframes
from usage_logger import log_event, get_logs
from datetime import datetime

# Page Configuration
st.set_page_config(
    page_title="Word Number Comparison",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Load CSS
with open("style.css") as f:
    st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

# Authentication Logic
def load_users():
    if os.path.exists("users.json"):
        with open("users.json", "r") as f:
            return json.load(f)["users"]
    return []

def check_credentials(username, password):
    users = load_users()
    for user in users:
        if user["username"] == username and user["password"] == password:
            return True
    return False

def save_user(username, password, role, auto_fill=False):
    users = load_users()
    # Check if user already exists
    if any(u["username"] == username for u in users):
        return False, f"Username '{username}' already exists."
    
    users.append({
        "username": username,
        "password": password,
        "role": role,
        "auto_fill": auto_fill
    })
    
    try:
        with open("users.json", "w") as f:
            json.dump({"users": users}, f, indent=4)
        return True, f"User '{username}' added successfully."
    except Exception as e:
        return False, f"Error saving user: {e}"

def remove_user(username):
    users = load_users()
    updated_users = [u for u in users if u["username"] != username]
    
    try:
        with open("users.json", "w") as f:
            json.dump({"users": updated_users}, f, indent=4)
        return True, f"User '{username}' deleted successfully."
    except Exception as e:
        return False, f"Error deleting user: {e}"

def update_user_data(old_username, new_username, new_password, new_role, auto_fill):
    users = load_users()
    
    # Check if new username is already taken (if changing name)
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
        with open("users.json", "w") as f:
            json.dump({"users": users}, f, indent=4)
        return True, f"User '{new_username}' updated successfully."
    except Exception as e:
        return False, f"Error updating user: {e}"

if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False

def login_screen():
    st.markdown('<div class="login-container">', unsafe_allow_html=True)
    st.markdown('<div class="login-header">🔐 User Login</div>', unsafe_allow_html=True)
    
    users = load_users()
    usernames = [u["username"] for u in users]
    # Set default to 'user' if it exists
    default_index = usernames.index("user") if "user" in usernames else 0
    username = st.selectbox("Select Username", usernames, index=default_index)
    
    # Check for auto-fill
    selected_user = next((u for u in users if u["username"] == username), None)
    auto_pass = selected_user["password"] if selected_user and selected_user.get("auto_fill") else ""
    
    with st.form("login_form"):
        password = st.text_input("Password", value=auto_pass, type="password")
        submit = st.form_submit_button("Sign In")
        
        if submit:
            if check_credentials(username, password):
                st.session_state.authenticated = True
                st.session_state.username = username
                log_event(username, "Login", "Successfully signed in")
                st.rerun()
            else:
                st.error("Invalid username or password")
    
    st.markdown('</div>', unsafe_allow_html=True)

if not st.session_state.authenticated:
    login_screen()
    st.stop()

# App UI (Authenticated Only)
st.markdown('<div class="app-logo">📊</div>', unsafe_allow_html=True)
st.markdown('<div class="main-header">Word Document Number Comparison</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Compare table data across two documents with different formats</div>', unsafe_allow_html=True)

# Sidebar Configuration
with st.sidebar:
    st.markdown(f"👤 **Logged in as:** {st.session_state.username}")
    if st.button("🚪 Logout"):
        st.session_state.authenticated = False
        st.rerun()
        
    st.divider()
    st.markdown("### ⚙️ Settings")
    
    # Combined Document Formats
    st.markdown("**1. Number Formats**")
    colA, colB = st.columns(2)
    with colA:
        st.caption("Document 1")
        format_1 = st.radio("Format 1", ["Vietnam", "US"], index=0, key="f1_fmt", label_visibility="collapsed")
    with colB:
        st.caption("Document 2")
        format_2 = st.radio("Format 2", ["Vietnam", "US"], index=1, key="f2_fmt", label_visibility="collapsed")
    
    st.divider()
    
    # Analysis Mode
    st.markdown("**2. Check Target**")
    extract_mode = st.radio("Target", ["Number", "Text"], index=0, horizontal=True, label_visibility="collapsed")
    
    st.markdown("**3. View Mode**")
    view_mode = st.radio(
        "Mode", 
        ["Show Mismatches Only", "Show All Results"], 
        index=0, 
        label_visibility="collapsed"
    )

    st.divider()
    with st.container():
        st.info("💡 **Tip:** Use 'Text' mode first to verify table headers match before checking values.")

# Main Interface
col1, col2 = st.columns(2)

with col1:
    st.info("📂 **Upload Document 1**")
    file1 = st.file_uploader("Select first .docx file", type=["docx"], key="doc1")

with col2:
    st.info("📂 **Upload Document 2**")
    file2 = st.file_uploader("Select second .docx file", type=["docx"], key="doc2")

if file1 and file2:
    if st.button("🚀 Run Comparison"):
        with st.spinner("Extracting and comparing data..."):
            try:
                # Extract
                df1 = extract_table_data(file1, format_1, mode=extract_mode)
                df2 = extract_table_data(file2, format_2, mode=extract_mode)
                
                # Check for data
                if df1.empty or df2.empty:
                    st.warning(f"One of the documents contains no {extract_mode.lower()} table data.")
                    st.stop()
                
                # Compare
                merged_df, result_msg, mismatch_count = compare_dataframes(df1, df2, mode=extract_mode)
                
                # Log Comparison
                log_event(st.session_state.username, "Comparison", f"Files: {file1.name}, {file2.name} | Mode: {extract_mode} | Mismatches: {mismatch_count}")
                
                # Display Results
                st.divider()
                
                # Metrics
                m_col1, m_col2, m_col3 = st.columns(3)
                m_col1.metric("Rows found in Document 1", len(df1))
                m_col2.metric("Rows found in Document 2", len(df2))
                
                delta_color = "normal" if mismatch_count == 0 else "inverse"
                m_col3.metric("Mismatches", mismatch_count, delta=mismatch_count, delta_color=delta_color)
                
                # result message
                if mismatch_count == 0:
                    st.success(result_msg)
                else:
                    st.error(result_msg)
                    st.markdown("""
                        <div class="word-hint">
                            <b>💡 Hướng dẫn tìm nhanh bảng trong Microsoft Word:</b><br>
                            1. Nhấn tổ hợp phím <b>Ctrl + G</b> để mở hộp thoại <i>Find and Replace (Go To)</i>.<br>
                            2. Tại mục <b>Go to what</b>, chọn <b>Table</b>.<br>
                            3. Nhập <b>Số thứ tự bảng</b> (lấy từ cột <i>Table</i> ở danh sách bên dưới) vào ô <b>Enter table number</b>.<br>
                            4. Nhấn <b>Go To</b> để di chuyển đến chính xác bảng cần kiểm tra.
                        </div>
                    """, unsafe_allow_html=True)
                
                # Results Table
                st.subheader(f"Detailed {extract_mode} Comparison")
                

                # Filter columns and identify differences
                display_df = merged_df.copy()
                
                if extract_mode == 'Number':
                    mismatch_mask = display_df['Diff'].abs() > 1e-6
                else:
                    mismatch_mask = display_df['Text 1'].fillna("") != display_df['Text 2'].fillna("")
                
                if view_mode == "Show Mismatches Only":
                    display_df = display_df[mismatch_mask]
                
                if extract_mode == 'Text':
                    # Hide numerical columns for text mode
                    display_df = display_df.drop(columns=['Value 1', 'Value 2', 'Diff'])
                
                # Styling
                def highlight_diff(row):
                    if extract_mode == 'Number':
                        is_diff = abs(row['Diff']) > 1e-6
                    else:
                        is_diff = str(row['Text 1']) != str(row['Text 2'])
                    return ['background-color: #ffebee' if is_diff else '' for _ in row]

                styled_df = display_df.style.apply(highlight_diff, axis=1)
                
                if extract_mode == 'Number':
                    styled_df = styled_df.format({
                        'Value 1': "{:,.2f}",
                        'Value 2': "{:,.2f}",
                        'Diff': "{:,.2f}"
                    })
                
                st.dataframe(styled_df, use_container_width=True, height=500)
                
                # Export to Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    merged_df.to_excel(writer, index=False, sheet_name='Comparison')
                excel_data = output.getvalue()
                
                st.download_button(
                    label="📥 Download Results (Excel)",
                    data=excel_data,
                    file_name="Number_Comparison_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            except Exception as e:
                st.error(f"An error occurred during processing: {str(e)}")
else:
    st.info("Please upload both Word documents to start the comparison.")

# Admin Section: Usage Logs
if st.session_state.authenticated and st.session_state.username == "admin":
    st.divider()
    with st.expander("📈 Usage Logs (Admin Only)", expanded=False):
        st.markdown("### Recent Tool Activity")
        logs = get_logs()
        if logs:
            log_df = pd.DataFrame(logs)
            st.dataframe(log_df, use_container_width=True)
            
            # Export Logs to Excel
            output_log = io.BytesIO()
            with pd.ExcelWriter(output_log, engine='openpyxl') as writer:
                log_df.to_excel(writer, index=False, sheet_name='UsageLogs')
            log_excel_data = output_log.getvalue()
            
            st.download_button(
                label="📥 Download Usage Logs (Excel)",
                data=log_excel_data,
                file_name=f"Usage_Log_Export_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            st.info(f"Total actions recorded: {len(logs)}")
        else:
            st.write("No logs found yet.")

    with st.expander("👤 User Management (Admin Only)", expanded=False):
        st.markdown("### Add New User")
        with st.form("add_user_form", clear_on_submit=True):
            new_user = st.text_input("New Username")
            new_pass = st.text_input("New Password", type="password")
            new_role = st.selectbox("Role", ["user", "admin"], index=0)
            new_auto_fill = st.checkbox("Auto-fill Password", value=False)
            add_btn = st.form_submit_button("➕ Add User")
            
            if add_btn:
                if new_user and new_pass:
                    success, msg = save_user(new_user, new_pass, new_role, new_auto_fill)
                    if success:
                        st.success(msg)
                        log_event(st.session_state.username, "User Management", f"Created user: {new_user} ({new_role})")
                    else:
                        st.error(msg)
                else:
                    st.warning("Please fill in both username and password.")
        
        st.divider()
        st.markdown("### Current Users (Edit/Delete)")
        
        current_users = load_users()
        
        for idx, user in enumerate(current_users):
            is_admin_user = (user["username"] == "admin")
            
            with st.container():
                # Display individual user row with unique keys
                c1, c2, c3, c4, c5, c6 = st.columns([2, 2, 1, 1, 1, 1])
                
                with c1:
                    # Admin username is protected
                    if is_admin_user:
                        edit_name = st.text_input("Name", value=user["username"], disabled=True, key=f"un_{idx}")
                    else:
                        edit_name = st.text_input("Name", value=user["username"], key=f"un_{idx}")
                
                with c2:
                    # Password is editable for everyone
                    edit_pass = st.text_input("Password", value=user["password"], type="password", key=f"pw_{idx}")
                
                with c3:
                    # Role is editable for everyone
                    edit_role = st.selectbox("Role", ["admin", "user"], index=0 if user["role"]=="admin" else 1, key=f"rl_{idx}")
                
                with c4:
                    # Auto-fill toggle
                    edit_auto = st.checkbox("Auto-fill", value=user.get("auto_fill", False), key=f"af_{idx}")

                with c5:
                    if st.button("💾 Save", key=f"upd_{idx}", use_container_width=True):
                        success, msg = update_user_data(user["username"], edit_name, edit_pass, edit_role, edit_auto)
                        if success:
                            st.success(msg)
                            log_event(st.session_state.username, "User Management", f"Updated user: {edit_name}")
                            st.rerun()
                        else:
                            st.error(msg)
                
                with c6:
                    # Admin account cannot be deleted
                    if is_admin_user:
                        st.button("🚫", disabled=True, key=f"del_dis_{idx}", help="Cannot delete main admin", use_container_width=True)
                    else:
                        if st.button("🗑️", key=f"del_{idx}", use_container_width=True, help=f"Delete {user['username']}"):
                            success, msg = remove_user(user["username"])
                            if success:
                                st.success(msg)
                                log_event(st.session_state.username, "User Management", f"Deleted user: {user['username']}")
                                st.rerun()
                            else:
                                st.error(msg)
                st.markdown("---")

# Footer
st.divider()
st.caption("Word Number Comparison Tool | Professional Streamlit Application")
