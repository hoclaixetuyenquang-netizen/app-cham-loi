import streamlit as st
import pandas as pd
from datetime import date
import sqlite3
import io

# ==========================================
# 1. TỪ ĐIỂN CÁC LỖI (TIẾNG VIỆT CHUẨN)
# ==========================================
DANH_SACH_LOI = {
    "loi_1": "Không thắt dây an toàn",
    "loi_2": "Không bật xi nhan trái xuất phát hoặc xi nhan phải kết thúc",
    "loi_3": "Không quan sát gương",
    "loi_4": "Dừng, đỗ xe sai quy định",
    "loi_5": "Không chấp hành hiệu lệnh biển báo",
    "loi_6": "Mở cửa xe không an toàn",
    "loi_7": "Vượt xe không đảm bảo an toàn",
    "loi_8": "Quay đầu xe không đúng quy định",
    "loi_9": "Không quan sát, giảm tốc độ hoặc dừng lại trong các trường hợp theo quy định",
    "loi_10": "Lỗi không chấp hành vạch kẻ đường",
    "loi_11": "Không thực hiện theo yêu cầu của Sát hạch viên",
    "loi_12": "Lỗi khác"
}

# ==========================================
# 2. KHỞI TẠO CƠ SỞ DỮ LIỆU
# ==========================================
conn = sqlite3.connect("dulieu_loi_v3.db", check_same_thread=False)
c = conn.cursor()

# Tạo bảng: Mỗi học viên là 1 dòng riêng biệt
cols_sql = ", ".join([f"{col} INTEGER" for col in DANH_SACH_LOI.keys()])
c.execute(f"""
    CREATE TABLE IF NOT EXISTS HocVien (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ngay_thi TEXT,
        sbd INTEGER,
        {cols_sql}
    )
""")
conn.commit()

# ==========================================
# 3. THIẾT KẾ GIAO DIỆN (UI)
# ==========================================
st.set_page_config(page_title="Chấm Lỗi Sát Hạch", page_icon="🚗", layout="centered")
st.title("🚗 App Chấm Lỗi Đường Trường")

# Chia 2 Tab
tab1, tab2 = st.tabs(["📝 NHẬP LỖI HỌC VIÊN", "📊 XUẤT BÁO CÁO EXCEL"])

# ------------------------------------------
# TAB 1: NHẬP LIỆU
# ------------------------------------------
with tab1:
    st.info("📌 Nhập số báo danh, tích các lỗi học viên mắc phải và bấm Lưu.")
    
    with st.form("form_nhap_loi", clear_on_submit=True):
        ngay_sat_hach = st.date_input("📅 Ngày sát hạch:", date.today())
        
        # Dùng number_input để khi bấm trên điện thoại hiện bàn phím số
        sbd = st.number_input("🔢 Số báo danh học viên:", min_value=1, step=1, value=1)
        
        st.write("### ❌ Tích chọn các lỗi vi phạm:")
        
        loi_ghi_nhan = {}
        for ma_loi, ten_loi in DANH_SACH_LOI.items():
            loi_ghi_nhan[ma_loi] = st.checkbox(ten_loi)
            
        st.markdown("<br>", unsafe_allow_html=True)
        submit_btn = st.form_submit_button("💾 LƯU KẾT QUẢ HỌC VIÊN NÀY", use_container_width=True)
        
        if submit_btn:
            gia_tri_loi = [1 if loi_ghi_nhan[ma] else 0 for ma in DANH_SACH_LOI.keys()]
            
            placeholders = ", ".join(["?"] * (len(DANH_SACH_LOI) + 2))
            query = f"INSERT INTO HocVien (ngay_thi, sbd, {', '.join(DANH_SACH_LOI.keys())}) VALUES ({placeholders})"
            
            c.execute(query, [str(ngay_sat_hach), int(sbd)] + gia_tri_loi)
            conn.commit()
            st.success(f"✅ Đã lưu kết quả học viên SBD: **{sbd}** thành công!")

# ------------------------------------------
# TAB 2: XUẤT BÁO CÁO
# ------------------------------------------
with tab2:
    st.write("### Tùy chọn xuất báo cáo")
    loai_bao_cao = st.radio("Bạn muốn xuất báo cáo theo:", ["Tất cả các ngày", "Chỉ một ngày cụ thể"])
    
    ngay_loc = None
    if loai_bao_cao == "Chỉ một ngày cụ thể":
        ngay_loc = st.date_input("Chọn ngày muốn xuất:", date.today())

    if st.button("📥 TẠO VÀ TẢI FILE EXCEL", use_container_width=True, type="primary"):
        # Lấy dữ liệu dạng danh sách từng học viên (không group by sum nữa)
        fields_query = ", ".join(DANH_SACH_LOI.keys())
        if loai_bao_cao == "Tất cả các ngày":
            query_sql = f"SELECT ngay_thi, {fields_query} FROM HocVien"
            ten_file = "Tong_Hop_Loi_Tat_Ca.xlsx"
        else:
            query_sql = f"SELECT ngay_thi, {fields_query} FROM HocVien WHERE ngay_thi = '{ngay_loc}'"
            ten_file = f"Tong_Hop_Loi_Ngay_{ngay_loc}.xlsx"
            
        df = pd.read_sql_query(query_sql, conn)
        
        if df.empty:
            st.warning("📭 Không có dữ liệu nào.")
        else:
            # Thêm cột STT chạy từ 1, 2, 3... tương ứng với từng dòng học viên
            df.insert(0, 'STT', range(1, len(df) + 1))
            
            # Đổi tên cột chuẩn xác theo form Excel mẫu
            cot_tieng_viet = ["STT", "Ngày sát hạch"] + list(DANH_SACH_LOI.values())
            df.columns = cot_tieng_viet
            
            # Xuất file Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='TongHopLoi')
            
            st.success("🎉 File Excel đã tạo xong!")
            st.download_button(
                label="⬇️ TẢI FILE EXCEL VỀ MÁY",
                data=output.getvalue(),
                file_name=ten_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
            st.markdown("---")
            st.write("👀 *Xem trước danh sách (Mỗi học viên 1 dòng):*")
            st.dataframe(df)
