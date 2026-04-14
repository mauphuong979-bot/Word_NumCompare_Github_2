import streamlit as st
import pandas as pd
import io
from extractor import extract_table_data
from processor import compare_dataframes

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

# App UI
st.markdown('<div class="app-logo">📊</div>', unsafe_allow_html=True)
st.markdown('<div class="main-header">Word Document Number Comparison</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Compare table data across two documents with different formats</div>', unsafe_allow_html=True)

# Sidebar Configuration
with st.sidebar:
    st.markdown("### ⚙️ Control Panel")
    st.caption("Configure your comparison settings below.")
    
    # Group 1: File Formats
    with st.expander("📄 Document Formats", expanded=True):
        st.markdown("**Document 1**")
        format_1 = st.radio("Number Format", ["US", "Vietnam"], index=1, key="f1_fmt", help="US: 1,234.56 | Vietnam: 1.234,56", label_visibility="collapsed")
        
        st.divider()
        
        st.markdown("**Document 2**")
        format_2 = st.radio("Number Format", ["US", "Vietnam"], index=0, key="f2_fmt", help="US: 1,234.56 | Vietnam: 1.234,56", label_visibility="collapsed")
    
    # Group 2: Analysis Options
    with st.expander("🛠️ Analysis & View", expanded=True):
        st.markdown("**Extraction Target**")
        extract_mode = st.radio("Target", ["Number", "Non-Number"], index=0, help="Number: Extract digits | Non-Number: Extract text", label_visibility="collapsed")
        
        st.divider()
        
        st.markdown("**View Mode**")
        view_mode = st.radio(
            "Mode", 
            ["Show Mismatches Only", "Show All Results"], 
            index=0, 
            help="Toggle between seeing only errors or the full list.",
            label_visibility="collapsed"
        )

    st.divider()
    with st.container():
        st.info("💡 **Tip:** Use 'Non-Number' mode first to verify table headers match before checking values.")

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
                
                # Display Results
                st.divider()
                
                # Metrics
                m_col1, m_col2, m_col3 = st.columns(3)
                m_col1.metric("Rows found in Doc 1", len(df1))
                m_col2.metric("Rows found in Doc 2", len(df2))
                
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
                
                if extract_mode == 'Non-Number':
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

# Footer
st.divider()
st.caption("Word Number Comparison Tool | Professional Streamlit Application")
