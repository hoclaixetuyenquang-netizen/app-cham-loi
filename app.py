import streamlit as st
import pandas as pd
from datetime import date
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
    .stCheckbox {
        padding: 8px 0;
    }
    .stCheckbox label {
        font-size: 15px;
    }
    .stDateInput input {
        font-size: 16px;
    }
    h1 {
        font-size: 24px;
        text-align: center;
    }
    h2 {
        font-size: 18px;
    }
    /* Giảm khoảng cách giữa các phần tử */
    .stForm {
        border: 1px solid #ddd;
        padding: 10px;
        border-radius: 5px;
    }
</style>
""", unsafe_allow_html=True)

# Hàm phát âm thanh thông báo
def play_sound():
    """Phát âm thanh thông báo khi lưu thành công"""
    # Tạo âm thanh beep đơn giản bằng HTML5 Audio
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
cols_sql = ", ".join([f"{col} INTEGER" for col in columns])
c.execute(f"CREATE TABLE IF NOT EXISTS LoiThi (ngay TEXT, {cols_sql})")
conn.commit()

# 2. Giao diện nhập liệu
st.title("🚗 Tích Lỗi Học Viên")

# Chọn ngày
ngay_sat_hach = st.date_input("📅 Ngày sát hạch", date.today(), format="DD/MM/YYYY")

# Form tích lỗi với 2 cột để tiết kiệm không gian
st.write("**✓ Tích vào các lỗi học viên mắc phải:**")

with st.form("form_loi"):
    # Chia thành 2 cột trên mobile
    col1, col2 = st.columns(2)
    
    loi_values = []
    for idx, col_name in enumerate(columns):
        # Phân bố đều các checkbox vào 2 cột
        target_col = col1 if idx % 2 == 0 else col2
        with target_col:
            display_text = col_name.replace("_", " ")
            # Mặc định checkbox luôn unchecked (False)
            loi_values.append(int(st.checkbox(display_text, key=f"loi_{idx}", value=False)))
    
    col_btn1, col_btn2 = st.columns(2)
    
    with col_btn1:
        submit = st.form_submit_button("💾 Lưu", use_container_width=True)
    
    with col_btn2:
        reset_btn = st.form_submit_button("🔄 Xóa", use_container_width=True)
    
    if submit:
        placeholders = ", ".join(["?"] * (len(columns) + 1))
        c.execute(f"INSERT INTO LoiThi VALUES ({placeholders})", [str(ngay_sat_hach)] + loi_values)
        conn.commit()
        
        # Phát âm thanh
        play_sound()
        
        # Hiển thị thông báo thành công
        st.success("✅ Đã lưu thành công!")
        st.balloons()
        
        # Reset form bằng cách xóa các key từ session state
        for idx in range(len(columns)):
            if f"loi_{idx}" in st.session_state:
                del st.session_state[f"loi_{idx}"]
        
        # Rerun để reset form
        st.rerun()
    
    if reset_btn:
        # Xóa tất cả checkbox
        for idx in range(len(columns)):
            if f"loi_{idx}" in st.session_state:
                del st.session_state[f"loi_{idx}"]
        st.rerun()

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

if st.button("📥 Tạo File Excel", use_container_width=True):
    # Đọc dữ liệu từ DB
    df = pd.read_sql_query("SELECT * FROM LoiThi", conn)
    
    if df.empty:
        st.warning("⚠️ Chưa có dữ liệu để xuất.")
    else:
        # Gom nhóm cộng dồn lỗi theo ngày
        df_tong_hop = df.groupby('ngay').sum(numeric_only=True).reset_index()
        df_tong_hop.insert(0, 'STT', range(1, len(df_tong_hop) + 1))
        
        # Đổi tên cột cho giống mẫu
        df_tong_hop.columns = [
            "STT", "Ngày sát hạch", "Không thắt dây an toàn", 
            "Không bật xi nhan trái/phải", "Không quan sát gương", 
            "Dừng, đỗ xe sai quy định", "Không chấp hành hiệu lệnh", 
            "Mở cửa xe không an toàn", "Vượt xe không đảm bảo", 
            "Quay đầu xe không đúng", "Không quan sát, giảm tốc", 
            "Lỗi vạch kẻ đường", "Không thực hiện theo yêu cầu", "Lỗi khác"
        ]
        
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
    col_manage1, col_manage2 = st.columns(2)
    
    with col_manage1:
        if st.button("🔍 Xem Tất Cả Dữ Liệu", use_container_width=True):
            df_view = pd.read_sql_query("SELECT * FROM LoiThi", conn)
            if not df_view.empty:
                st.dataframe(df_view, use_container_width=True)
            else:
                st.info("Không có dữ liệu")
    
    with col_manage2:
        if st.button("🗑️ Xóa Tất Cả", use_container_width=True):
            c.execute("DELETE FROM LoiThi")
            conn.commit()
            st.success("✅ Đã xóa tất cả dữ liệu")
            st.rerun()
