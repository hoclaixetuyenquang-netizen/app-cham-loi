import streamlit as st
import pandas as pd
from datetime import date, datetime, timedelta
import sqlite3
import io
import base64

# Cấu hình trang cho di động
st.set_page_config(
    page_title="Chấm Lỗi",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# CSS tối ưu cho điện thoại
st.markdown("""
<style>
    /* Tối ưu kích thước cho mobile */
    .main {
        padding: 0.5rem;
    }
    .stButton > button {
        width: 100%;
        padding: 10px;
        font-size: 16px;
        font-weight: bold;
    }
    h1 {
        font-size: 24px;
        text-align: center;
    }
    h2 {
        font-size: 18px;
    }
    /* Nút tích lỗi sát lề trái */
    .error-button-container {
        width: 100%;
        margin-bottom: 8px;
    }
    .error-button {
        width: 100%;
        padding: 12px;
        border: 2px solid #ddd;
        border-radius: 5px;
        background-color: #f0f0f0;
        cursor: pointer;
        font-size: 15px;
        font-weight: 600;
        text-align: left;
        transition: all 0.2s ease;
        user-select: none;
    }
    .error-button:hover {
        border-color: #0066cc;
        background-color: #e6f0ff;
    }
    .error-button.active {
        background-color: #ff4444;
        color: white;
        border-color: #cc0000;
    }
    /* Định dạng bảng báo cáo */
    .total-row {
        font-weight: bold;
        background-color: #e8f4f8;
    }
</style>
""", unsafe_allow_html=True)

# Hàm phát âm thanh thông báo
def play_sound():
    """Phát âm thanh thông báo khi lưu thành công"""
    sound_html = """
    <audio autoplay>
        <source src="data:audio/wav;base64,UklGRiYAAABXQVZFZm10IBAAAAABAAEAQB8AAAB9AAACABAAZGF0YQIAAAAAAA==" type="audio/wav">
    </audio>
    """
    st.markdown(sound_html, unsafe_allow_html=True)

# 1. Khởi tạo Database SQLite cục bộ
conn = sqlite3.connect("dulieu_loi.db", check_same_thread=False)
c = conn.cursor()

# Tạo bảng nếu chưa có
columns = [
    "Khong_that_day_an_toan", "Khong_bat_xi_nhan", "Khong_quan_sat_guong", 
    "Dung_do_xe_sai_quy_dinh", "Khong_chap_hanh_bien_bao", "Mo_cua_xe_khong_an_toan", 
    "Vuot_xe_khong_an_toan", "Quay_dau_sai_quy_dinh", "Khong_giam_toc_do", 
    "Khong_chap_hanh_vach_ke", "Khong_nghe_sat_hach_vien", "Loi_khac"
]

columns_display = [
    "Không thắt dây an toàn", "Không bật xi nhan trái/phải", "Không quan sát gương",
    "Dừng, đỗ xe sai quy định", "Không chấp hành hiệu lệnh", "Mở cửa xe không an toàn",
    "Vượt xe không đảm bảo", "Quay đầu xe không đúng", "Không quan sát, giảm tốc",
    "Lỗi vạch kẻ đường", "Không thực hiện theo yêu cầu", "Lỗi khác"
]

cols_sql = ", ".join([f"{col} INTEGER" for col in columns])
c.execute(f"CREATE TABLE IF NOT EXISTS LoiThi (ngay TEXT, {cols_sql})")
conn.commit()

# Khởi tạo session state cho việc theo dõi các lỗi được chọn
if "selected_errors" not in st.session_state:
    st.session_state.selected_errors = set()

if "save_success" not in st.session_state:
    st.session_state.save_success = False

# 2. Giao diện nhập liệu
st.title("🚗 Tích Lỗi Học Viên")

# Chọn ngày
ngay_sat_hach = st.date_input("📅 Ngày sát hạch", date.today(), format="DD/MM/YYYY")

# Form tích lỗi với nút bấm
st.write("**✓ Tích vào các lỗi học viên mắc phải (nhấn để chọn/bỏ chọn):**")

# Hiển thị thông báo nếu vừa lưu thành công
if st.session_state.save_success:
    st.success("✅ Đã lưu thành công!", icon="✅")
    st.balloons()
    st.session_state.save_success = False

# Tạo các nút bấm lỗi sát lề trái
for idx, display_text in enumerate(columns_display):
    # Tạo nút bấm động với màu sắc thay đổi
    is_selected = idx in st.session_state.selected_errors
    button_color = "🔴" if is_selected else "⚪"
    
    if st.button(f"{button_color} {display_text}", key=f"error_btn_{idx}", use_container_width=True):
        # Toggle lỗi
        if idx in st.session_state.selected_errors:
            st.session_state.selected_errors.remove(idx)
        else:
            st.session_state.selected_errors.add(idx)
        st.rerun()

# Các nút hành động
st.divider()
col1, col2, col3 = st.columns(3)

with col1:
    if st.button("💾 Lưu", use_container_width=True):
        # Tạo mảng lỗi
        loi_values = [1 if idx in st.session_state.selected_errors else 0 for idx in range(len(columns))]
        
        # Lưu vào database
        placeholders = ", ".join(["?"] * (len(columns) + 1))
        c.execute(f"INSERT INTO LoiThi VALUES ({placeholders})", [str(ngay_sat_hach)] + loi_values)
        conn.commit()
        
        # Phát âm thanh
        play_sound()
        
        # Reset lỗi đã chọn
        st.session_state.selected_errors = set()
        st.session_state.save_success = True
        
        st.rerun()

with col2:
    if st.button("🔄 Xóa", use_container_width=True):
        st.session_state.selected_errors = set()
        st.rerun()

with col3:
    # Hiển thị số lỗi được chọn
    st.metric("Số lỗi chọn", len(st.session_state.selected_errors))

# 3. Xuất báo cáo tổng hợp ra Excel
st.divider()
st.subheader("📊 Báo Cáo Tổng Hợp")

# Thống kê nhanh
try:
    df_all = pd.read_sql_query("SELECT * FROM LoiThi", conn)
    if not df_all.empty:
        st.info(f"📈 Tổng số lần ghi nhận: **{len(df_all)}** | Số ngày: **{df_all['ngay'].nunique()}**")
except:
    pass

# Phần lọc dữ liệu
st.write("**Lọc dữ liệu báo cáo:**")
filter_col1, filter_col2, filter_col3 = st.columns(3)

filter_type = "all"
filter_date = None
filter_date_from = None
filter_date_to = None

with filter_col1:
    filter_type = st.radio("Chọn loại lọc:", ["Tất cả", "Một ngày", "Từ ngày đến ngày"], horizontal=True)

if filter_type == "Một ngày":
    with filter_col2:
        filter_date = st.date_input("Chọn ngày:", date.today(), format="DD/MM/YYYY", key="filter_single_date")
elif filter_type == "Từ ngày đến ngày":
    with filter_col2:
        filter_date_from = st.date_input("Từ ngày:", date.today() - timedelta(days=7), format="DD/MM/YYYY", key="filter_from_date")
    with filter_col3:
        filter_date_to = st.date_input("Đến ngày:", date.today(), format="DD/MM/YYYY", key="filter_to_date")

if st.button("📥 Tạo Báo Cáo", use_container_width=True):
    # Đọc dữ liệu từ DB
    df = pd.read_sql_query("SELECT * FROM LoiThi", conn)
    
    if df.empty:
        st.warning("⚠️ Chưa có dữ liệu để xuất.")
    else:
        # Áp dụng bộ lọc
        if filter_type == "Một ngày":
            df = df[df['ngay'] == str(filter_date)]
        elif filter_type == "Từ ngày đến ngày":
            df = df[(df['ngay'] >= str(filter_date_from)) & (df['ngay'] <= str(filter_date_to))]
        
        if df.empty:
            st.warning("⚠️ Không có dữ liệu cho khoảng thời gian đã chọn.")
        else:
            # Gom nhóm cộng dồn lỗi theo ngày
            df_tong_hop = df.groupby('ngay').sum(numeric_only=True).reset_index()
            
            # Tính tổng cộng
            total_row = df_tong_hop.sum(numeric_only=True)
            total_row['ngay'] = 'TỔNG CỘNG'
            
            # Thêm hàng tổng cộng vào đầu
            df_tong_hop = pd.concat([pd.DataFrame([total_row]), df_tong_hop], ignore_index=True)
            
            # Thêm cột STT
            df_tong_hop.insert(0, 'STT', range(0, len(df_tong_hop)))
            df_tong_hop.loc[0, 'STT'] = ''  # Hàng tổng không có STT
            
            # Reset STT cho các dòng khác
            for i in range(1, len(df_tong_hop)):
                df_tong_hop.loc[i, 'STT'] = i
            
            # Đổi tên cột
            df_tong_hop.columns = [
                "STT", "Ngày sát hạch", "Không thắt dây an toàn", 
                "Không bật xi nhan trái/phải", "Không quan sát gương", 
                "Dừng, đỗ xe sai quy định", "Không chấp hành hiệu lệnh", 
                "Mở cửa xe không an toàn", "Vượt xe không đảm bảo", 
                "Quay đầu xe không đúng", "Không quan sát, giảm tốc", 
                "Lỗi vạch kẻ đường", "Không thực hiện theo yêu cầu", "Lỗi khác"
            ]
            
            # Chuyển các cột lỗi thành số nguyên (trừ hàng tổng)
            for col in df_tong_hop.columns[2:]:
                df_tong_hop[col] = df_tong_hop[col].astype('Int64')
            
            # Xuất ra file Excel trên bộ nhớ (để tải về)
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_tong_hop.to_excel(writer, index=False, sheet_name='TongHopLoi')
            
            st.download_button(
                label="📥 Tải File Excel",
                data=output.getvalue(),
                file_name=f"Bao_cao_loi_{date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
            # Hiển thị preview dữ liệu
            with st.expander("👁️ Xem trước dữ liệu"):
                st.dataframe(df_tong_hop, use_container_width=True)

# 4. Phần quản lý dữ liệu (tùy chọn)
with st.expander("⚙️ Quản Lý Dữ Liệu"):
    st.subheader("Xem Dữ Liệu")
    col_manage1, col_manage2 = st.columns(2)
    
    with col_manage1:
        if st.button("🔍 Xem Tất Cả Dữ Liệu", use_container_width=True):
            df_view = pd.read_sql_query("SELECT * FROM LoiThi", conn)
            if not df_view.empty:
                # Đổi tên cột cho hiển thị
                df_view_display = df_view.copy()
                df_view_display.columns = [
                    "Ngày sát hạch", "Không thắt dây an toàn", 
                    "Không bật xi nhan trái/phải", "Không quan sát gương", 
                    "Dừng, đỗ xe sai quy định", "Không chấp hành hiệu lệnh", 
                    "Mở cửa xe không an toàn", "Vượt xe không đảm bảo", 
                    "Quay đầu xe không đúng", "Không quan sát, giảm tốc", 
                    "Lỗi vạch kẻ đường", "Không thực hiện theo yêu cầu", "Lỗi khác"
                ]
                st.dataframe(df_view_display, use_container_width=True)
            else:
                st.info("Không có dữ liệu")
    
    with col_manage2:
        st.subheader("Xóa Dữ Liệu")
        password_input = st.text_input("Nhập mật khẩu:", type="password", key="del_password")
        if st.button("🗑️ Xóa Tất Cả", use_container_width=True):
            if password_input == "123":
                c.execute("DELETE FROM LoiThi")
                conn.commit()
                st.success("✅ Đã xóa tất cả dữ liệu")
                st.rerun()
            else:
                st.error("❌ Mật khẩu không chính xác!")
