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
conn = sqlite3.connect("dulieu_loi_v2.db", check_same_thread=False)
c = conn.cursor()

# Tạo bảng: Mỗi dòng là 1 học viên
cols_sql = ", ".join([f"{col} INTEGER" for col in DANH_SACH_LOI.keys()])
c.execute(f"""
    CREATE TABLE IF NOT EXISTS HocVien (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ngay_thi TEXT,
        ho_ten TEXT,
        {cols_sql}
    )
""")
conn.commit()

# ==========================================
# 3. THIẾT KẾ GIAO DIỆN (UI)
# ==========================================
st.set_page_config(page_title="Chấm Lỗi Sát Hạch", page_icon="🚗", layout="centered")
st.title("🚗 App Chấm Lỗi Đường Trường")

# Chia giao diện làm 2 Tab để dễ dùng trên điện thoại
tab1, tab2 = st.tabs(["📝 NHẬP LỖI HỌC VIÊN", "📊 XUẤT BÁO CÁO EXCEL"])

# ------------------------------------------
# TAB 1: NHẬP LIỆU (Dành cho Sát hạch viên)
# ------------------------------------------
with tab1:
    st.info("📌 Chọn ngày thi và nhập thông tin học viên. Tích vào các lỗi mắc phải và bấm Lưu.")
    
    # Form nhập liệu (clear_on_submit=True giúp xóa trắng form sau khi lưu)
    with st.form("form_nhap_loi", clear_on_submit=True):
        ngay_sat_hach = st.date_input("📅 Ngày sát hạch:", date.today())
        ho_ten = st.text_input("👤 Họ tên học viên / SBD (Bắt buộc):", placeholder="VD: Nguyễn Văn A - SBD 123")
        
        st.write("### ❌ Tích chọn các lỗi vi phạm:")
        
        # Tạo danh sách checkbox từ từ điển tiếng Việt
        loi_ghi_nhan = {}
        for ma_loi, ten_loi in DANH_SACH_LOI.items():
            loi_ghi_nhan[ma_loi] = st.checkbox(ten_loi)
            
        st.markdown("<br>", unsafe_allow_html=True)
        submit_btn = st.form_submit_button("💾 LƯU DỮ LIỆU HỌC VIÊN NÀY", use_container_width=True)
        
        if submit_btn:
            if ho_ten.strip() == "":
                st.error("⚠️ Vui lòng nhập Tên hoặc Số báo danh của học viên!")
            else:
                # Chuyển True/False từ checkbox thành 1/0
                gia_tri_loi = [1 if loi_ghi_nhan[ma] else 0 for ma in DANH_SACH_LOI.keys()]
                
                # Lưu vào DB
                placeholders = ", ".join(["?"] * (len(DANH_SACH_LOI) + 2))
                query = f"INSERT INTO HocVien (ngay_thi, ho_ten, {', '.join(DANH_SACH_LOI.keys())}) VALUES ({placeholders})"
                
                c.execute(query, [str(ngay_sat_hach), ho_ten] + gia_tri_loi)
                conn.commit()
                st.success(f"✅ Đã lưu thành công kết quả của: **{ho_ten}**")

# ------------------------------------------
# TAB 2: XUẤT BÁO CÁO (Dành cho Cán bộ tổng hợp)
# ------------------------------------------
with tab2:
    st.write("### Tùy chọn xuất báo cáo")
    loai_bao_cao = st.radio("Bạn muốn xuất báo cáo cho:", ["Tất cả các ngày", "Chỉ một ngày cụ thể"])
    
    ngay_loc = None
    if loai_bao_cao == "Chỉ một ngày cụ thể":
        ngay_loc = st.date_input("Chọn ngày muốn xuất:", date.today())

    if st.button("📥 TẠO VÀ TẢI FILE EXCEL", use_container_width=True, type="primary"):
        # Lấy dữ liệu từ DB
        if loai_bao_cao == "Tất cả các ngày":
            df = pd.read_sql_query("SELECT * FROM HocVien", conn)
            ten_file = "Tong_Hop_Loi_Tat_Ca.xlsx"
        else:
            df = pd.read_sql_query(f"SELECT * FROM HocVien WHERE ngay_thi = '{ngay_loc}'", conn)
            ten_file = f"Tong_Hop_Loi_Ngay_{ngay_loc}.xlsx"
        
        if df.empty:
            st.warning("📭 Không có dữ liệu nào trong thời gian này.")
        else:
            # Bỏ cột id và tên học viên để Gom nhóm (Cộng dồn) lỗi theo ngày thi
            df_loi_chi_tiet = df.drop(columns=['id', 'ho_ten'])
            df_tong_hop = df_loi_chi_tiet.groupby('ngay_thi').sum().reset_index()
            
            # Thêm cột STT
            df_tong_hop.insert(0, 'STT', range(1, len(df_tong_hop) + 1))
            
            # Đổi tên cột chuẩn xác theo form Excel mẫu của bạn
            cot_tieng_viet = ["STT", "Ngày sát hạch"] + list(DANH_SACH_LOI.values())
            df_tong_hop.columns = cot_tieng_viet
            
            # Đổ dữ liệu vào file Excel trên RAM (không lưu rác ra máy)
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_tong_hop.to_excel(writer, index=False, sheet_name='TongHopLoi')
            
            st.success("🎉 File Excel đã tạo xong! Bấm nút bên dưới để lưu về máy.")
            st.download_button(
                label="⬇️ TẢI FILE EXCEL VỀ MÁY",
                data=output.getvalue(),
                file_name=ten_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
            st.markdown("---")
            st.write("👀 *Xem trước một phần dữ liệu tổng hợp:*")
            st.dataframe(df_tong_hop)
