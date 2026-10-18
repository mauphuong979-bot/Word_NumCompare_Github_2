import streamlit as st
import pandas as pd
import io
import os
import sys
import platform

# Cấu hình trang ứng dụng (phải gọi đầu tiên trước mọi lệnh Streamlit khác)
st.set_page_config(
    page_title="Đối Chiếu Số Liệu Word",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Đảm bảo thư mục hiện tại luôn nằm ở đầu sys.path (đặc biệt quan trọng trên Streamlit Cloud)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

CSS_FILE = os.path.join(BASE_DIR, "style.css")

# Nhập các mô-đun nghiệp vụ với cơ chế kiểm tra và hiển thị lỗi rõ ràng
try:
    from extractor import (
        extract_table_data,
        get_document_tables_summary,
        detect_number_format,
        format_display_number
    )
    from processor import compare_documents
    from annotator import annotate_docx, create_zip_package
except Exception as e:
    import traceback
    st.error(f"❌ Lỗi tải thư viện ứng dụng ({type(e).__name__}): {str(e)}")
    st.code(traceback.format_exc())
    st.info("💡 Hướng dẫn xử lý: Vui lòng vào thanh điều khiển 'Manage app' ở góc dưới bên phải Streamlit Cloud -> chọn dấu 3 chấm (⋮) -> 'Reboot app' để Streamlit Cloud cài đặt lại đầy đủ các gói từ requirements.txt.")
    st.stop()

# Tải định dạng CSS tùy biến
if os.path.exists(CSS_FILE):
    with open(CSS_FILE, "r", encoding="utf-8") as f:
        st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

@st.dialog("📖 Hướng Dẫn Sử Dụng", width="large")
def show_user_guide():
    st.markdown("""
### 1. Giới thiệu chức năng
Ứng dụng hỗ trợ kiểm tra, rà soát và đối chiếu tự động dữ liệu bảng biểu giữa hai tài liệu Word (`.docx`). Hệ thống giúp phát hiện sai lệch số liệu, kiểm tra biến động cấu trúc bảng và tự động đính kèm bóng thoại ghi chú (native comment) trực tiếp vào văn bản Word để phục vụ công tác kiểm toán, soát xét báo cáo tài chính và hợp đồng.

---

### 2. Các bước thực hiện
1. **Tải tài liệu:**
   * Chọn hoặc kéo thả **Tài liệu 1** (file gốc / đối chiếu) và **Tài liệu 2** (file cần kiểm tra).
   * Hệ thống tự động nhận diện chuẩn định dạng số của từng file (Chuẩn Việt Nam: `1.234,56` hoặc Chuẩn Quốc tế: `1,234.56`).
2. **Cấu hình tùy chọn (Thanh bên trái):**
   * **Chuẩn hiển thị kết quả:** Chọn định dạng số muốn hiển thị trên giao diện và file xuất ra.
   * **Mục tiêu đối chiếu:** Bật/tắt so sánh số liệu (mặc định) và so sánh chữ / nhãn chỉ tiêu.
3. **Thực hiện đối chiếu:**
   * Bấm nút **🚀 BẮT ĐẦU ĐỐI CHIẾU DỮ LIỆU**.
   * Hệ thống tự động căn chỉnh bảng biểu và gióng từng dòng theo thuật toán tối ưu, kể cả khi có dòng thừa/thiếu hoặc bảng bị gộp/tách.
4. **Xem & Tải kết quả:**
   * Xem tóm tắt nhanh trên màn hình (tỉ lệ khớp, số lượng điểm lệch, biến động cấu trúc bảng).
   * Tải trọn bộ kết quả (.zip) gồm 2 file Word đã gắn bóng thoại ghi chú (`.docx`) và báo cáo chi tiết (`.xlsx`).

---

### 3. Ý nghĩa các ký hiệu & Kết quả đối chiếu
* **❌ Sai lệch số liệu:** Giá trị số tại ô tương ứng khác nhau. Trong file Word xuất ra, ô này được tô màu nền đỏ và đính kèm bóng thoại ghi chú chi tiết chênh lệch.
* **⚠️ Thừa dòng / Dòng đệm trống:** Dòng xuất hiện ở một tài liệu nhưng không có ở tài liệu còn lại. Hệ thống tự nhận diện các dòng đệm trống định dạng để gióng đúng các dòng dữ liệu phía sau mà không gây báo lỗi dây chuyền.
* **ℹ️ Bảng gộp / tách (Đã tự động đối chiếu):** Trường hợp 2 bảng ở tài liệu này được gộp thành 1 bảng ở tài liệu kia (hoặc ngược lại). Hệ thống tự động phân tách cấu trúc để đối chiếu số liệu từng ô khớp 100% và gắn bóng thoại thông tin mà không tô màu cảnh báo nhầm lẫn.

---

### 4. An toàn & Bảo mật
* Toàn bộ quá trình đọc, xử lý và đối chiếu tài liệu diễn ra hoàn toàn cục bộ trên máy tính của bạn, không gửi dữ liệu ra máy chủ bên ngoài.
""")
    if st.button("Đóng hướng dẫn", use_container_width=True):
        st.rerun()

# --- THANH BÊN CÀI ĐẶT (SIDEBAR) ---
with st.sidebar:
    st.markdown("""
        <div class="sidebar-title-box">
            <span class="sidebar-title-icon">⚙️</span>
            <span class="sidebar-title-text">Cài Đặt Đối Chiếu</span>
        </div>
    """, unsafe_allow_html=True)
    
    # Khối 1: Định dạng số hiển thị
    with st.container(border=True):
        st.markdown('<div class="sidebar-card-label">Chuẩn hiển thị kết quả</div>', unsafe_allow_html=True)
        display_standard = st.radio(
            "Chuẩn hiển thị kết quả",
            ["Chuẩn Việt Nam (1.234,56)", "Chuẩn Quốc tế / US (1,234.56)"],
            index=0,
            label_visibility="collapsed"
        )
    
    # Khối 2: Mục tiêu đối chiếu
    with st.container(border=True):
        st.markdown('<div class="sidebar-card-label">Mục tiêu đối chiếu</div>', unsafe_allow_html=True)
        compare_numbers = st.checkbox("So sánh số liệu", value=True, help="Kiểm tra độ chính xác của các ô số liệu tài chính")
        compare_text = st.checkbox("So sánh chữ / nhãn", value=False, help="Kiểm tra nội dung văn bản và nhãn chỉ tiêu")
        if not compare_numbers and not compare_text:
            st.warning("⚠️ Chọn ít nhất 1 mục tiêu.")
    
    # Nút Xem hướng dẫn & Nút Đối chiếu tệp mới
    if st.button("📖 Hướng dẫn sử dụng", use_container_width=True):
        show_user_guide()

    if st.button("🔄 Đối chiếu tệp mới", use_container_width=True):
        st.session_state.clear()
        st.rerun()

    # Khối 3: Thông tin công nghệ
    st.markdown("""
        <div class="sidebar-info-box">
            <b>⚡ Cơ chế thông minh:</b><br>
            • Tự phát hiện chuẩn số từng file<br>
            • Gióng bảng & dòng theo quy hoạch động<br>
            • Tự động dựng và phân tách bảng gộp<br>
            • Chống lệch dây chuyền do dòng đệm
        </div>
    """, unsafe_allow_html=True)
    
    # Chân trang thanh bên
    st.markdown(f"""
        <div class="sidebar-footer">
            Đối Chiếu Số Liệu Word • v2.3<br>
            Bảo mật cục bộ ({platform.system()})
        </div>
    """, unsafe_allow_html=True)

# Tiêu đề ứng dụng
st.markdown('<div class="app-logo">📊</div>', unsafe_allow_html=True)
st.markdown('<div class="main-header">Đối Chiếu Số Liệu Word</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Hệ thống tự động rà soát sai lệch số liệu & kiểm tra cấu trúc văn bản</div>', unsafe_allow_html=True)

# 1. Khu vực Tải Tài Liệu
col_c1, col_c2 = st.columns(2, gap="medium")

with col_c1:
    with st.container(border=True):
        st.markdown("#### 📂 Tài liệu 1 (File gốc / Đối chiếu)")
        file1 = st.file_uploader("Chọn file .docx thứ nhất", type=["docx"], key="doc1", label_visibility="collapsed")
        f1_detected = None
        if file1:
            f1_bytes = file1.getvalue()
            f1_detected = detect_number_format(io.BytesIO(f1_bytes))
            css_cls = "us-format" if f1_detected['format'] == 'US' else ""
            fmt_desc = "Chuẩn Quốc tế / US (1,234.56)" if f1_detected['format'] == 'US' else "Chuẩn Việt Nam (1.234,56)"
            samples_str = " | ".join(f"<code>{s}</code>" for s in f1_detected['samples'][:3])
            st.markdown(f"""
                <div class="meta-badge {css_cls}">
                    <b>🎯 Tự động phát hiện:</b> {fmt_desc} (Độ tin cậy: {f1_detected['confidence']}%)<br>
                    <span style="font-size:0.78rem; color:#475569;">Mẫu số liệu: {samples_str}</span>
                </div>
            """, unsafe_allow_html=True)

with col_c2:
    with st.container(border=True):
        st.markdown("#### 📂 Tài liệu 2 (File cần kiểm tra)")
        file2 = st.file_uploader("Chọn file .docx thứ hai", type=["docx"], key="doc2", label_visibility="collapsed")
        f2_detected = None
        if file2:
            f2_bytes = file2.getvalue()
            f2_detected = detect_number_format(io.BytesIO(f2_bytes))
            css_cls = "us-format" if f2_detected['format'] == 'US' else ""
            fmt_desc = "Chuẩn Quốc tế / US (1,234.56)" if f2_detected['format'] == 'US' else "Chuẩn Việt Nam (1.234,56)"
            samples_str = " | ".join(f"<code>{s}</code>" for s in f2_detected['samples'][:3])
            st.markdown(f"""
                <div class="meta-badge {css_cls}">
                    <b>🎯 Tự động phát hiện:</b> {fmt_desc} (Độ tin cậy: {f2_detected['confidence']}%)<br>
                    <span style="font-size:0.78rem; color:#475569;">Mẫu số liệu: {samples_str}</span>
                </div>
            """, unsafe_allow_html=True)

format_1 = f1_detected['format'] if f1_detected else 'Vietnam'
format_2 = f2_detected['format'] if f2_detected else 'Vietnam'

# Nút thực hiện đối chiếu
if file1 and file2:
    if not compare_numbers and not compare_text:
        st.warning("Vui lòng chọn ít nhất một mục tiêu đối chiếu (So sánh số hoặc So sánh chữ) ở thanh bên trái.")
    else:
        if st.button("🚀 BẮT ĐẦU ĐỐI CHIẾU DỮ LIỆU", type="primary", use_container_width=True):
            with st.spinner("Đang tự động căn chỉnh bảng, gióng dòng thông minh và rà soát dữ liệu..."):
                try:
                    f1_bytes = file1.getvalue()
                    f2_bytes = file2.getvalue()
                    
                    results_df, structure_df, struct_stats, stats_summary, disc1, disc2 = compare_documents(
                        io.BytesIO(f1_bytes), io.BytesIO(f2_bytes), format_1, format_2,
                        compare_numbers=compare_numbers, compare_text=compare_text
                    )
                    
                    # Tạo file Word có comment bóng thoại và tô màu
                    ann1_bytes = annotate_docx(f1_bytes, disc1, 'doc1', file1.name, file2.name, display_standard=display_standard)
                    ann2_bytes = annotate_docx(f2_bytes, disc2, 'doc2', file1.name, file2.name, display_standard=display_standard)
                    
                    # Tạo báo cáo Excel 2 sheet
                    excel_buf = io.BytesIO()
                    with pd.ExcelWriter(excel_buf, engine='openpyxl') as writer:
                        structure_df.to_excel(writer, index=False, sheet_name='Tong_quan_cau_truc')
                        results_df.to_excel(writer, index=False, sheet_name='Chi_tiet_doi_chieu')
                    excel_bytes = excel_buf.getvalue()
                    
                    # Đóng gói ZIP
                    zip_bytes = create_zip_package(ann1_bytes, file1.name, ann2_bytes, file2.name, excel_bytes)
                    
                    # Lưu vào Session State
                    st.session_state['comparison_ready'] = True
                    st.session_state['results_df'] = results_df
                    st.session_state['structure_df'] = structure_df
                    st.session_state['struct_stats'] = struct_stats
                    st.session_state['stats_summary'] = stats_summary
                    st.session_state['ann1_bytes'] = ann1_bytes
                    st.session_state['ann2_bytes'] = ann2_bytes
                    st.session_state['excel_bytes'] = excel_bytes
                    st.session_state['zip_bytes'] = zip_bytes
                    st.session_state['file1_name'] = file1.name
                    st.session_state['file2_name'] = file2.name
                    st.session_state['display_standard'] = display_standard
                    st.session_state['compare_numbers'] = compare_numbers
                    st.session_state['compare_text'] = compare_text
                except Exception as e:
                    st.error(f"Lỗi trong quá trình phân tích đối chiếu: {str(e)}")

    # Hiển thị kết quả đối chiếu
    if st.session_state.get('comparison_ready'):
        results_df = st.session_state['results_df']
        structure_df = st.session_state['structure_df']
        struct_stats = st.session_state['struct_stats']
        stats_summary = st.session_state['stats_summary']
        ann1_bytes = st.session_state['ann1_bytes']
        ann2_bytes = st.session_state['ann2_bytes']
        excel_bytes = st.session_state['excel_bytes']
        zip_bytes = st.session_state['zip_bytes']
        f1_name = st.session_state['file1_name']
        f2_name = st.session_state['file2_name']

        # Phân nhóm các loại sai lệch
        num_mismatches = results_df[results_df['Trạng thái'] == "❌ Sai lệch số liệu"]
        format_mismatches = results_df[results_df['Trạng thái'] == "❌ Lệch định dạng số/chữ"]
        text_mismatches = results_df[results_df['Trạng thái'] == "❌ Lệch văn bản"]
        extra_rows = results_df[results_df['Trạng thái'].str.contains("Thừa dòng", na=False)]
        unmatched_tables = results_df[results_df['Trạng thái'].str.contains("Bảng riêng|gộp/tách", na=False)]

        real_mismatch_count = len(num_mismatches) + len(format_mismatches) + len(text_mismatches)
        unmatched_count = len(unmatched_tables)

        st.divider()

        # Chuẩn bị mô tả cấu trúc
        diff_structure_df = structure_df[structure_df['Trạng thái'] != "✅ Đồng nhất"]
        struct_diff_cnt = len(diff_structure_df)
        
        merged_tables_cnt = len(structure_df[structure_df['Trạng thái'].str.contains('gộp|tách', case=False, na=False)])
        row_diff_cnt = len(structure_df[structure_df['Trạng thái'].str.contains('dòng', case=False, na=False) & ~structure_df['Trạng thái'].str.contains('gộp|tách', case=False, na=False)])
        unmat_cnt = len(structure_df[structure_df['Trạng thái'].str.contains('Bảng riêng|Không đối ứng', case=False, na=False)])
        
        parts_sub = []
        if row_diff_cnt > 0:
            parts_sub.append(f"{row_diff_cnt} bảng lệch dòng")
        if merged_tables_cnt > 0:
            parts_sub.append(f"{merged_tables_cnt} bảng gộp/tách")
        if unmat_cnt > 0:
            parts_sub.append(f"{unmat_cnt} bảng riêng")
            
        if parts_sub:
            struct_sub = ", ".join(parts_sub)
        elif struct_diff_cnt > 0:
            struct_sub = f"{struct_diff_cnt} bảng khác biệt"
        else:
            struct_sub = "Đồng nhất 100%"

        # --- BANNER TỔNG QUAN KẾT QUẢ ---
        if real_mismatch_count > 0 and struct_diff_cnt > 0:
            banner_cls = "danger"
            banner_title = f'<span style="font-size:1.35rem; margin-right:0.35rem;">🙁</span> Phát hiện {real_mismatch_count} điểm lệch số liệu và {struct_diff_cnt} bảng có khác biệt cấu trúc'
            banner_sub = f"Đã đối chiếu <b>{stats_summary['total_checked']:,}</b> số liệu (tỉ lệ khớp <b>{stats_summary['match_rate']}%</b>) qua <b>{struct_stats['aligned_pairs_count']}</b> bảng tương ứng giữa 2 tài liệu ({struct_sub})."
        elif real_mismatch_count > 0:
            banner_cls = "danger"
            banner_title = f'<span style="font-size:1.35rem; margin-right:0.35rem;">🙁</span> Phát hiện {real_mismatch_count} điểm sai lệch số liệu'
            banner_sub = f"Đã đối chiếu <b>{stats_summary['total_checked']:,}</b> số liệu (tỉ lệ khớp <b>{stats_summary['match_rate']}%</b>). Cấu trúc các bảng biểu hoàn toàn đồng nhất."
        elif struct_diff_cnt > 0 or len(extra_rows) > 0:
            banner_cls = "warning"
            banner_title = f'<span style="font-size:1.35rem; margin-right:0.35rem;">🧐</span> Số liệu tại các bảng khớp 100%, phát hiện {struct_diff_cnt} bảng có khác biệt cấu trúc'
            banner_sub = f"Toàn bộ <b>{stats_summary['total_checked']:,}</b> số liệu đều khớp chính xác. Lưu ý: Có <b>{struct_diff_cnt}</b> bảng ({struct_sub}) cần rà soát lại."
        else:
            banner_cls = "success"
            banner_title = '<span style="font-size:1.35rem; margin-right:0.35rem;">😊</span> Xác nhận hoàn toàn trùng khớp giữa 2 tài liệu'
            banner_sub = f"Toàn bộ <b>{stats_summary['total_checked']:,}</b> số liệu và cấu trúc toàn bộ <b>{struct_stats['aligned_pairs_count']}</b> bảng biểu đều khớp chính xác 100%."

        banner_html = f'<div class="summary-banner {banner_cls}"><div class="banner-title">{banner_title}</div><div class="banner-sub">{banner_sub}</div></div>'
        st.markdown(banner_html, unsafe_allow_html=True)

        # Thẻ chỉ số tổng quát (Executive Stats Grid)
        val_class = "danger" if real_mismatch_count > 0 else ""
        mismatch_sub = "Cần kiểm tra lại" if real_mismatch_count > 0 else "Trùng khớp hoàn toàn"
        struct_class = "danger" if struct_diff_cnt > 0 else ""

        max_tables = max(struct_stats["total_tables_1"], struct_stats["total_tables_2"])
        stats_html = (
            f'<div class="stats-row">'
            f'<div class="stat-card"><div class="stat-label">Sai lệch số liệu</div><div class="stat-value {val_class}">{real_mismatch_count}</div><div class="stat-sub">{mismatch_sub}</div></div>'
            f'<div class="stat-card"><div class="stat-label">Tỉ lệ khớp số liệu</div><div class="stat-value">{stats_summary["match_rate"]}%</div><div class="stat-sub">{stats_summary["total_matched"]:,} / {stats_summary["total_checked"]:,} số liệu</div></div>'
            f'<div class="stat-card"><div class="stat-label">Bảng đã so sánh</div><div class="stat-value">{struct_stats["aligned_pairs_count"]} bảng</div><div class="stat-sub">Khớp {struct_stats["aligned_pairs_count"]}/{max_tables} bảng giữa 2 file</div></div>'
            f'<div class="stat-card"><div class="stat-label">Khác biệt cấu trúc</div><div class="stat-value {struct_class}">{struct_diff_cnt} bảng</div><div class="stat-sub">{struct_sub}</div></div>'
            f'</div>'
        )
        st.markdown(stats_html, unsafe_allow_html=True)

        # --- MỤC 1: TẢI FILE KẾT QUẢ ---
        st.divider()
        st.markdown("### 1. Tải file kết quả")

        dl_col1, dl_col2, dl_col3, dl_col4 = st.columns(4)

        with dl_col1:
            st.markdown("""
                <div class="download-card">
                    <div class="download-title">Gói tổng hợp</div>
                    <div class="download-desc">2 file Word có ghi chú + báo cáo Excel</div>
                </div>
            """, unsafe_allow_html=True)
            st.download_button(
                label="Tải trọn bộ (.zip)",
                data=zip_bytes,
                file_name="Word_Comparison_Package.zip",
                mime="application/zip",
                use_container_width=True
            )

        with dl_col2:
            st.markdown(f"""
                <div class="download-card">
                    <div class="download-title">File 1 có ghi chú</div>
                    <div class="download-desc">Tô màu ô & bóng thoại ghi chú Word</div>
                </div>
            """, unsafe_allow_html=True)
            st.download_button(
                label="Tải File 1 (.docx)",
                data=ann1_bytes,
                file_name=f"[Da_Comment]_{f1_name}",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

        with dl_col3:
            st.markdown(f"""
                <div class="download-card">
                    <div class="download-title">File 2 có ghi chú</div>
                    <div class="download-desc">Tô màu ô & bóng thoại ghi chú Word</div>
                </div>
            """, unsafe_allow_html=True)
            st.download_button(
                label="Tải File 2 (.docx)",
                data=ann2_bytes,
                file_name=f"[Da_Comment]_{f2_name}",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

        with dl_col4:
            st.markdown("""
                <div class="download-card">
                    <div class="download-title">Báo cáo Excel</div>
                    <div class="download-desc">2 sheet: cấu trúc & chi tiết đối chiếu</div>
                </div>
            """, unsafe_allow_html=True)
            st.download_button(
                label="Tải báo cáo (.xlsx)",
                data=excel_bytes,
                file_name="Bao_cao_doi_chieu.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        # --- MỤC 2: TÓM TẮT & CHI TIẾT SAI LỆCH ---
        st.divider()
        merged_tables = structure_df[structure_df['Trạng thái'].str.contains('gộp|tách', case=False, na=False)]
        total_issues_count = real_mismatch_count + len(extra_rows) + unmatched_count + len(merged_tables)
        issue_badge = f" ({total_issues_count} điểm cần lưu ý)" if total_issues_count > 0 else ""
        st.markdown(f"### 2. Tóm tắt & Chi tiết sai lệch{issue_badge}")
        
        has_any_issue = False

        # 2.1 Điểm sai lệch số liệu
        if not num_mismatches.empty:
            has_any_issue = True
            items_html = []
            for _, r in num_mismatches.iterrows():
                v1_str = format_display_number(r['Giá trị File 1'], display_standard)
                v2_str = format_display_number(r['Giá trị File 2'], display_standard)
                diff_str = format_display_number(r['Chênh lệch'], display_standard, show_sign=True)
                item = (
                    f'• <b>{r["Bảng F1"]} ↔ {r["Bảng F2"]}</b> tại <code>{r["Vị trí File 1"]}</code> / <code>{r["Vị trí File 2"]}</code>: '
                    f'Chỉ tiêu <b>{r["Chỉ tiêu / Nhãn"]}</b> lệch <b>{diff_str}</b> (File 1: <code>{v1_str}</code> | File 2: <code>{v2_str}</code>).'
                )
                items_html.append(f'<div class="summary-item">{item}</div>')
            
            body_html = "".join(items_html)
            card_html = f'<div class="summary-box danger"><div class="summary-title danger">Sai lệch số liệu ({len(num_mismatches)} vị trí):</div>{body_html}</div>'
            st.markdown(card_html, unsafe_allow_html=True)

        # 2.2 Điểm lệch định dạng số / chữ hoặc văn bản
        other_mismatches = pd.concat([format_mismatches, text_mismatches])
        if not other_mismatches.empty:
            has_any_issue = True
            items_html = []
            for _, r in other_mismatches.iterrows():
                item = (
                    f'• <b>{r["Bảng F1"]} ↔ {r["Bảng F2"]}</b>: {r["Trạng thái"]} tại chỉ tiêu <b>{r["Chỉ tiêu / Nhãn"]}</b> '
                    f'(File 1: <code>{r["Văn bản File 1"]}</code> | File 2: <code>{r["Văn bản File 2"]}</code>).'
                )
                items_html.append(f'<div class="summary-item">{item}</div>')
            
            body_html = "".join(items_html)
            card_html = f'<div class="summary-box danger"><div class="summary-title danger">Lệch kiểu dữ liệu hoặc nội dung chữ ({len(other_mismatches)} vị trí):</div>{body_html}</div>'
            st.markdown(card_html, unsafe_allow_html=True)

        # 2.3 Thừa / thiếu dòng
        if not extra_rows.empty:
            has_any_issue = True
            items_html = []
            for _, r in extra_rows.iterrows():
                loc = r['Vị trí File 1'] if r['Vị trí File 1'] != "(Không có)" else r['Vị trí File 2']
                t_f = r['Bảng F1'] if r['Bảng F1'] != "(Không đối ứng)" and r['Bảng F1'] != "(Không có)" else r['Bảng F2']
                clean_r_loc = loc
                if '_R' in loc:
                    clean_r_loc = loc.split('_')[1].replace('R', 'Dòng ')
                item = f'• {r["Trạng thái"]} tại <b>{t_f}</b> (<code>{clean_r_loc}</code>): {r["Chỉ tiêu / Nhãn"]}'
                items_html.append(f'<div class="summary-item">{item}</div>')
            
            body_html = "".join(items_html)
            card_html = f'<div class="summary-box warning"><div class="summary-title warning">Dòng thừa hoặc thiếu ({len(extra_rows)} dòng):</div>{body_html}</div>'
            st.markdown(card_html, unsafe_allow_html=True)

        # 2.4 Biến động cấu trúc: Bảng gộp / tách (Đã tự động đối chiếu số liệu)
        if not merged_tables.empty:
            has_any_issue = True
            items_html = []
            for _, r in merged_tables.iterrows():
                item = f'• <b>{r["Bảng File 1"]} ↔ {r["Bảng File 2"]}</b>: {r["Ghi chú"]} <i>(Số liệu từng ô đã được tự động dựng lại và đối chiếu khớp 100%)</i>.'
                items_html.append(f'<div class="summary-item">{item}</div>')
            body_html = "".join(items_html)
            card_html = f'<div class="summary-box warning"><div class="summary-title warning">Bảng bị gộp/tách cấu trúc ({len(merged_tables)} bảng — Đã tự động đối chiếu số):</div>{body_html}</div>'
            st.markdown(card_html, unsafe_allow_html=True)

        # 2.5 Cảnh báo cấu trúc: Bảng riêng chưa đối chiếu tự động
        if not unmatched_tables.empty:
            has_any_issue = True
            items_html = []
            for _, r in unmatched_tables.iterrows():
                t_name = r['Bảng F1'] if r['Bảng F1'] != "(Không đối ứng)" else r['Bảng F2']
                f_origin = "File 1" if r['Bảng F1'] != "(Không đối ứng)" else "File 2"
                note_text = r.get('Ghi chú & Khuyến nghị', '')
                item = f'• <b>{t_name} ({f_origin})</b>: {r["Chỉ tiêu / Nhãn"]}. <i>{note_text}</i>'
                items_html.append(f'<div class="summary-item">{item}</div>')
            
            body_html = "".join(items_html)
            footer_note = '<div style="font-size:0.8rem; color:#92400e; margin-top:0.35rem; padding-left:0.4rem;">Lưu ý: Các bảng chưa đối chiếu đã được tô màu và gắn bóng thoại ghi chú trực tiếp trong file Word tải về để hỗ trợ rà soát thủ công.</div>'
            card_html = f'<div class="summary-box warning"><div class="summary-title warning">Bảng chưa đối chiếu tự động do gộp/tách bảng ({len(unmatched_tables)} bảng):</div>{body_html}{footer_note}</div>'
            st.markdown(card_html, unsafe_allow_html=True)

        # 2.6 Hoàn toàn khớp
        if not has_any_issue:
            st.markdown('<div class="summary-box success"><div class="summary-title success"><span style="font-size:1.15rem; margin-right:0.35rem;">😊</span> Xác nhận hoàn hảo — Không phát hiện sai lệch</div><div class="summary-item">Tất cả số liệu và cấu trúc các bảng biểu giữa hai tài liệu đều khớp chính xác 100%.</div></div>', unsafe_allow_html=True)

        # 2.6 Chi tiết dòng sai lệch (gọn gàng, 7 cột, không cuộn ngang)
        bad_rows_df = results_df[results_df['Trạng thái'] != "✅ Khớp"].copy()
        if not bad_rows_df.empty:
            st.markdown(f"<div style='margin-top:1.2rem; font-weight:600; font-size:0.95rem; color:#1e293b;'>Bảng chi tiết các dòng phát sinh ({len(bad_rows_df)} dòng):</div>", unsafe_allow_html=True)
            display_rows = []
            for _, r in bad_rows_df.iterrows():
                t1 = str(r.get('Bảng F1', ''))
                t2 = str(r.get('Bảng F2', ''))
                stt = str(r.get('Trạng thái', ''))

                if t1 == '(Không đối ứng)':
                    table_str = f'{t2} (File 2)'
                elif t2 == '(Không đối ứng)':
                    table_str = f'{t1} (File 1)'
                elif t1 == t2:
                    table_str = t1
                else:
                    table_str = f'{t1} ↔ {t2}'

                loc1 = str(r.get('Vị trí File 1', ''))
                loc2 = str(r.get('Vị trí File 2', ''))
                
                def clean_loc(loc):
                    if '_R' in loc and 'C' in loc:
                        parts = loc.split('_')[1]
                        r_part = parts.split('C')[0].replace('R', 'Dòng ')
                        c_part = 'Cột ' + parts.split('C')[1]
                        return f'{r_part}, {c_part}'
                    if '_R' in loc:
                        parts = loc.split('_')[1]
                        return parts.replace('R', 'Dòng ')
                    if loc in ['(Không có)', '(Không đối ứng)']:
                        return '—'
                    return loc

                l1_clean = clean_loc(loc1)
                l2_clean = clean_loc(loc2)
                if l1_clean == l2_clean:
                    loc_str = l1_clean
                elif l1_clean == '—':
                    loc_str = l2_clean
                elif l2_clean == '—':
                    loc_str = l1_clean
                else:
                    loc_str = f'{l1_clean}'

                v1 = r.get('Giá trị File 1')
                v2 = r.get('Giá trị File 2')
                diff = r.get('Chênh lệch')

                if pd.notna(v1):
                    v1_str = format_display_number(v1, display_standard)
                else:
                    v1_str = '—'

                if pd.notna(v2):
                    v2_str = format_display_number(v2, display_standard)
                else:
                    v2_str = '—'

                if pd.notna(diff):
                    diff_str = format_display_number(diff, display_standard, show_sign=True)
                else:
                    diff_str = '—'

                label_raw = str(r.get('Chỉ tiêu / Nhãn', ''))
                if 'Bảng ' in label_raw and ':' in label_raw:
                    label_raw = label_raw.split(':', 1)[0]
                elif len(label_raw) > 35:
                    label_raw = label_raw[:32] + '...'

                if 'Sai lệch số liệu' in stt:
                    result_desc = f'Lệch {diff_str}'
                elif 'Bảng riêng' in stt or 'gộp/tách' in stt:
                    loc_str = 'Toàn bảng'
                    if t2 == '(Không đối ứng)':
                        v1_str = 'Có trong F1'
                        v2_str = '(Bị gộp ở F2)'
                    else:
                        v1_str = '(Tách thành 2 bảng ở F1)'
                        v2_str = 'Có trong F2'
                    result_desc = 'Bị gộp/tách (chưa đối chiếu)'
                elif 'Thừa dòng' in stt:
                    result_desc = stt.replace('⚠️ ', '')
                elif 'Lệch văn bản' in stt:
                    result_desc = 'Lệch văn bản'
                else:
                    result_desc = stt.replace('❌ ', '').replace('⚠️ ', '')

                display_rows.append({
                    'Bảng': table_str,
                    'Vị trí': loc_str,
                    'Chỉ tiêu': label_raw,
                    'File 1': v1_str,
                    'File 2': v2_str,
                    'Chênh lệch': diff_str,
                    'Kết quả': result_desc
                })

            compact_df = pd.DataFrame(display_rows)

            def highlight_rows(row):
                res = str(row.get('Kết quả', ''))
                if 'Lệch' in res:
                    return ['background-color: #fef2f2'] * len(row)
                elif 'Bảng riêng' in res or 'Thừa dòng' in res or 'gộp/tách' in res:
                    return ['background-color: #fffbeb'] * len(row)
                return [''] * len(row)

            styled_compact = compact_df.style.apply(highlight_rows, axis=1)
            st.dataframe(styled_compact, use_container_width=True, hide_index=True, height=min(320, 45 + len(compact_df) * 36))

        # 2.7 Bảng biến động cấu trúc (nếu có)
        diff_structure_df = structure_df[structure_df['Trạng thái'] != "✅ Đồng nhất"]
        if not diff_structure_df.empty:
            st.markdown(f"<div style='margin-top:1.2rem; font-weight:600; font-size:0.95rem; color:#1e293b;'>Bảng biến động cấu trúc ({len(diff_structure_df)} bảng):</div>", unsafe_allow_html=True)
            st.caption("Các bảng bên dưới có sự biến động về số dòng, cột hoặc phát sinh do gộp/tách bảng:")
            clean_struct = pd.DataFrame({
                'File 1 (Bảng)': diff_structure_df['Bảng File 1'],
                'File 2 (Bảng)': diff_structure_df['Bảng File 2'],
                'Tiêu đề / Mẫu': diff_structure_df['Mô tả / Tiêu đề'].apply(lambda s: str(s)[:45] + '...' if len(str(s)) > 45 else str(s)),
                'Kích thước F1': diff_structure_df['File 1 (Kích thước)'],
                'Kích thước F2': diff_structure_df['File 2 (Kích thước)'],
                'Trạng thái': diff_structure_df['Trạng thái'].str.replace('⚠️ ', ''),
                'Ghi chú': diff_structure_df['Ghi chú']
            })
            st.dataframe(clean_struct, use_container_width=True, hide_index=True, height=min(220, 45 + len(clean_struct) * 35))
else:
    st.info("Vui lòng tải lên 2 tài liệu Word (.docx) để bắt đầu đối chiếu.")

# Chân trang ứng dụng
st.divider()
st.caption("Đối Chiếu Số Liệu Word | Ứng dụng Kiểm tra Văn bản Tự động")
