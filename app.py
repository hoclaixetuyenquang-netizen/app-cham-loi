import streamlit as st
import pandas as pd
from datetime import date, datetime
import sqlite3
import io
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

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
conn = sqlite3.connect("dulieu_loi_v5.db", check_same_thread=False)
c = conn.cursor()

cols_sql = ", ".join([f"{col} INTEGER" for col in DANH_SACH_LOI.keys()])
c.execute(f"""
    CREATE TABLE IF NOT EXISTS HocVien (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ngay_iso TEXT,
        ngay_hien_thi TEXT,
        sbd TEXT,
        {cols_sql}
    )
""")
conn.commit()

# ==========================================
# 3. HÀM TẠO FILE EXCEL GIỐNG MẪU
# ==========================================
def tao_excel_chuan_mau(data_rows, ten_cot_2, tieu_de="Số lượng lỗi do sát hạch viên trừ trong phần thi đường trường"):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    headers = ["STT", ten_cot_2] + list(DANH_SACH_LOI.values())

    # Dòng tiêu đề lớn (Tô màu xanh)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(headers))
    title_cell = ws.cell(row=2, column=1, value=tieu_de)
    title_cell.font = Font(name="Times New Roman", size=12, bold=True)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    title_cell.fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")

    # Dòng tiêu đề các cột
    for col_idx, text in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col_idx, value=text)
        cell.font = Font(name="Times New Roman", size=10, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")

    ws.row_dimensions[2].height = 25
    ws.row_dimensions[3].height = 80

    # Đổ dữ liệu
    start_row = 4
    for idx, row_data in enumerate(data_rows, 1):
        curr_row = start_row + idx - 1
        ws.cell(row=curr_row, column=1, value=idx).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=curr_row, column=2, value=row_data[0]).alignment = Alignment(horizontal="center", vertical="center")
        
        for col_idx, val in enumerate(row_data[1:], 3):
            cell = ws.cell(row=curr_row, column=col_idx, value=val)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.font = Font(name="Times New Roman", size=11)

    # Dòng TỔNG SỐ ở cuối
    last_data_row = start_row + len(data_rows) - 1
    total_row = last_data_row + 1

    ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=2)
    total_label = ws.cell(row=total_row, column=1, value="TỔNG SỐ")
    total_label.font = Font(name="Times New Roman", size=11, bold=True)
    total_label.alignment = Alignment(horizontal="center", vertical="center")

    for col_idx in range(3, len(headers) + 1):
        col_letter = get_column_letter(col_idx)
        sum_formula = f"=SUM({col_letter}{start_row}:{col_letter}{last_data_row})"
        cell = ws.cell(row=total_row, column=col_idx, value=sum_formula)
        cell.font = Font(name="Times New Roman", size=11, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Kẻ khung ô
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                         top=Side(style='thin'), bottom=Side(style='thin'))

    for row in ws.iter_rows(min_row=2, max_row=total_row, min_col=1, max_col=len(headers)):
        for cell in row:
            cell.border = thin_border

    ws.column_dimensions['A'].width = 6
    ws.column_dimensions['B'].width = 16
    for i in range(3, len(headers) + 1):
        ws.column_dimensions[get_column_letter(i)].width = 13

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# ==========================================
# 4. GIAO DIỆN ỨNG DỤNG (STREAMLIT)
# ==========================================
st.set_page_config(page_title="Chấm Lỗi Sát Hạch", page_icon="🚗", layout="centered")
st.title("🚗 App Chấm Lỗi Đường Trường")

tab1, tab2 = st.tabs(["📝 NHẬP LỖI HỌC VIÊN", "📊 XUẤT BÁO CÁO EXCEL"])

# ------------------------------------------
# TAB 1: NHẬP LIỆU
# ------------------------------------------
with tab1:
    st.info("📌 Bấm số trên bảng bên dưới để nhập SBD, tích lỗi và bấm Lưu.")
    
    ngay_sat_hach_dt = st.date_input("📅 Ngày sát hạch:", date.today())
    ngay_iso = ngay_sat_hach_dt.strftime("%Y-%m-%d")
    ngay_hien_thi = ngay_sat_hach_dt.strftime("%d/%m/%Y") # dd/mm/yyyy

    if "sbd_val" not in st.session_state:
        st.session_state["sbd_val"] = ""

    # Màn hình hiển thị SBD
    st.markdown("##### 🔢 Số báo danh học viên:")
    sbd_display = st.session_state["sbd_val"] if st.session_state["sbd_val"] else "---"
    st.markdown(f"<h2 style='text-align: center; color: #1E88E5; background-color: #E3F2FD; padding: 10px; border-radius: 8px;'>{sbd_display}</h2>", unsafe_allow_html=True)

    # BẢNG SỐ ẢO (Virtual Numpad)
    c1, c2, c3 = st.columns(3)
    with c1:
        if st.button("1", key="k1", use_container_width=True): st.session_state["sbd_val"] += "1"; st.rerun()
        if st.button("4", key="k4", use_container_width=True): st.session_state["sbd_val"] += "4"; st.rerun()
        if st.button("7", key="k7", use_container_width=True): st.session_state["sbd_val"] += "7"; st.rerun()
        if st.button("🔴 Xóa hết", key="k_clr", use_container_width=True): st.session_state["sbd_val"] = ""; st.rerun()
    with c2:
        if st.button("2", key="k2", use_container_width=True): st.session_state["sbd_val"] += "2"; st.rerun()
        if st.button("5", key="k5", use_container_width=True): st.session_state["sbd_val"] += "5"; st.rerun()
        if st.button("8", key="k8", use_container_width=True): st.session_state["sbd_val"] += "8"; st.rerun()
        if st.button("0", key="k0", use_container_width=True): st.session_state["sbd_val"] += "0"; st.rerun()
    with c3:
        if st.button("3", key="k3", use_container_width=True): st.session_state["sbd_val"] += "3"; st.rerun()
        if st.button("6", key="k6", use_container_width=True): st.session_state["sbd_val"] += "6"; st.rerun()
        if st.button("9", key="k9", use_container_width=True): st.session_state["sbd_val"] += "9"; st.rerun()
        if st.button("⌫ Xóa", key="k_back", use_container_width=True): st.session_state["sbd_val"] = st.session_state["sbd_val"][:-1]; st.rerun()

    st.markdown("---")
    st.write("### ❌ Tích chọn các lỗi vi phạm:")
    
    with st.form("form_tich_loi"):
        loi_ghi_nhan = {}
        for ma_loi, ten_loi in DANH_SACH_LOI.items():
            loi_ghi_nhan[ma_loi] = st.checkbox(ten_loi)
            
        submit_btn = st.form_submit_button("💾 LƯU KẾT QUẢ HỌC VIÊN NÀY", use_container_width=True, type="primary")
        
        if submit_btn:
            if not st.session_state["sbd_val"]:
                st.error("⚠️ Vui lòng nhập Số báo danh bằng bảng số bên trên!")
            else:
                gia_tri_loi = [1 if loi_ghi_nhan[ma] else 0 for ma in DANH_SACH_LOI.keys()]
                placeholders = ", ".join(["?"] * (len(DANH_SACH_LOI) + 3))
                query = f"INSERT INTO HocVien (ngay_iso, ngay_hien_thi, sbd, {', '.join(DANH_SACH_LOI.keys())}) VALUES ({placeholders})"
                
                c.execute(query, [ngay_iso, ngay_hien_thi, st.session_state["sbd_val"]] + gia_tri_loi)
                conn.commit()
                st.success(f"✅ Đã lưu kết quả SBD **{st.session_state['sbd_val']}** ngày **{ngay_hien_thi}**!")
                st.session_state["sbd_val"] = ""
                st.rerun()

# ------------------------------------------
# TAB 2: XUẤT BÁO CÁO
# ------------------------------------------
with tab2:
    st.write("### 📊 Chức năng xuất báo cáo")
    loai_bao_cao = st.radio(
        "Chọn chế độ xuất dữ liệu:",
        ["Chỉ 1 ngày cụ thể (Chi tiết từng SBD)", "Từ ngày... Đến ngày... (Tổng hợp cộng dồn theo ngày)"]
    )
    
    cols_loi_sql = ", ".join(DANH_SACH_LOI.keys())

    if loai_bao_cao == "Chỉ 1 ngày cụ thể (Chi tiết từng SBD)":
        ngay_chon_dt = st.date_input("Chọn ngày muốn xuất:", date.today())
        ngay_chon_iso = ngay_chon_dt.strftime("%Y-%m-%d")
        ngay_chon_str = ngay_chon_dt.strftime("%d/%m/%Y")
        
        if st.button("📥 TẠO VÀ TẢI FILE EXCEL CHI TIẾT", use_container_width=True, type="primary"):
            c.execute(f"SELECT sbd, {cols_loi_sql} FROM HocVien WHERE ngay_iso = ? ORDER BY id ASC", (ngay_chon_iso,))
            rows = c.fetchall()
            
            if not rows:
                st.warning(f"📭 Không có dữ liệu trong ngày {ngay_chon_str}.")
            else:
                excel_bytes = tao_excel_chuan_mau(rows, "Số báo danh")
                st.success(f"🎉 Tạo file thành công! Tổng cộng: {len(rows)} lượt thi.")
                st.download_button(
                    label="⬇️ TẢI FILE EXCEL BÁO CÁO",
                    data=excel_bytes,
                    file_name=f"Chi_Tiet_Loi_Ngay_{ngay_chon_dt.strftime('%d_%m_%Y')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

    else: # Từ ngày đến ngày (Cộng dồn theo ngày như mẫu)
        col_d1, col_d2 = st.columns(2)
        with col_d1:
            tu_ngay_dt = st.date_input("Từ ngày:", date.today())
        with col_d2:
            den_ngay_dt = st.date_input("Đến ngày:", date.today())
            
        tu_ngay_iso = tu_ngay_dt.strftime("%Y-%m-%d")
        den_ngay_iso = den_ngay_dt.strftime("%Y-%m-%d")

        if st.button("📥 TẠO VÀ TẢI FILE EXCEL TỔNG HỢP", use_container_width=True, type="primary"):
            sums_sql = ", ".join([f"SUM({k})" for k in DANH_SACH_LOI.keys()])
            query = f"""
                SELECT ngay_hien_thi, {sums_sql} 
                FROM HocVien 
                WHERE ngay_iso BETWEEN ? AND ? 
                GROUP BY ngay_iso 
                ORDER BY ngay_iso ASC
            """
            c.execute(query, (tu_ngay_iso, den_ngay_iso))
            rows = c.fetchall()
            
            if not rows:
                st.warning("📭 Không có dữ liệu trong khoảng thời gian đã chọn.")
            else:
                excel_bytes = tao_excel_chuan_mau(rows, "Ngày sát hạch")
                st.success(f"🎉 Tạo xong báo cáo tổng hợp cho {len(rows)} ngày thi!")
                st.download_button(
                    label="⬇️ TẢI FILE EXCEL TỔNG HỢP",
                    data=excel_bytes,
                    file_name=f"Tong_Hop_Loi_{tu_ngay_dt.strftime('%d%m%Y')}_den_{den_ngay_dt.strftime('%d%m%Y')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
