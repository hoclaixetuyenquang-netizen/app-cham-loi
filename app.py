import streamlit as st
import pandas as pd
from datetime import date
import sqlite3
import io

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

# 2. Giao diện nhập liệu trên điện thoại
st.set_page_config(page_title="Chấm Lỗi Đường Trường", layout="centered")
st.title("🚗 Tích lỗi học viên")

# Form tích lỗi
ngay_sat_hach = st.date_input("Ngày sát hạch", date.today())
with st.form("form_loi"):
    st.write("Tích vào các lỗi học viên mắc phải:")
    loi_values = []
    for col in columns:
        # Thay thế dấu '_' bằng khoảng trắng để hiển thị tên lỗi cho đẹp
        loi_values.append(int(st.checkbox(col.replace("_", " "))))
    
    submit = st.form_submit_button("Lưu dữ liệu học viên này")
    
    if submit:
        placeholders = ", ".join(["?"] * (len(columns) + 1))
        c.execute(f"INSERT INTO LoiThi VALUES ({placeholders})", [str(ngay_sat_hach)] + loi_values)
        conn.commit()
        st.success("✅ Đã lưu thành công!")

# 3. Xuất báo cáo tổng hợp ra Excel
st.divider()
st.subheader("📊 Xuất báo cáo tổng hợp")

if st.button("Tạo file Excel"):
    # Đọc dữ liệu từ DB
    df = pd.read_sql_query("SELECT * FROM LoiThi", conn)
    
    if df.empty:
        st.warning("Chưa có dữ liệu để xuất.")
    else:
        # Gom nhóm cộng dồn lỗi theo ngày
        df_tong_hop = df.groupby('ngay').sum().reset_index()
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
            label="📥 Tải file Excel Báo Cáo",
            data=output.getvalue(),
            file_name=f"Bao_cao_loi_{date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )