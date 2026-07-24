import streamlit as st
import pandas as pd
from datetime import date
import sqlite3
import io
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# ==========================================
# 1. CSS TÙY CHỈNH: GIAO DIỆN MOBILE & Ô BẤM
# ==========================================
st.markdown("""
<style>
    /* Ép các cột thành hàng ngang trên mobile */
    [data-testid="stHorizontalBlock"] {
        flex-direction: row !important;
        flex-wrap: nowrap !important;
        gap: 5px !important;
    }
    [data-testid="column"] {
        min-width: 0 !important;
        padding: 0 !important;
    }
    
    /* Phóng to nút bấm bảng số */
    [data-testid="stButton"] button {
        height: 60px !important;
        font-size: 24px !important;
        font-weight: bold !important;
        border-radius: 12px !important;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    
    /* Biến Checkbox thành dạng Ô BẤM (Tiles) */
    [data-testid="stCheckbox"] {
        background-color: #f8f9fa;
        padding: 12px 15px;
        border-radius: 10px;
        border: 2px solid #e9ecef;
        margin-bottom: 4px;
        transition: 0.2s;
    }
    /* Đổi màu xanh khi ô được bấm chọn */
    [data-testid="stCheckbox"]:has(input:checked) {
        background-color: #e0f2fe;
        border-color: #3b82f6;
    }
    [data-testid="stCheckbox"] label {
        cursor: pointer;
        width: 100%;
    }
    [data-testid="stCheckbox"] p {
        font-size: 16px !important;
        font-weight: 600 !important;
        color: #333;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 2. TỪ ĐIỂN LỖI & KHỞI TẠO CƠ SỞ DỮ LIỆU
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

conn = sqlite3.connect("dulieu_loi_v6.db", check_same_thread=False)
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
# 3. HÀM TẠO FILE EXCEL CHUẨN MẪU
# ==========================================
def tao_excel_chuan_mau(data_rows, ten_cot_2, tieu_de="Số lượng lỗi do sát hạch viên trừ trong phần thi đường trường"):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    headers = ["STT", ten_cot_2] + list(DANH_SACH_LOI.values())

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(headers))
    title_cell = ws.cell(row=2, column=1, value=tieu_de)
    title_cell.font = Font(name="Times New Roman", size=12, bold=True)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    title_cell.fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")

    for col_idx, text in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col_idx, value=text)
        cell.font = Font(name="Times New Roman", size=10, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")

    ws.row_dimensions[2].height = 25
    ws.row_dimensions[3].height = 80

    start_row = 4
    for idx, row_data in enumerate(data_rows, 1):
        curr_row = start_row + idx - 1
        ws.cell(row=curr_row, column=1, value=idx).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=curr_row, column=2, value=row_data[0]).alignment = Alignment(horizontal="center", vertical="center")
        
        for col_idx, val in enumerate(row_data[1:], 3):
            cell = ws.cell(row=curr_row, column=col_idx, value=val)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.font = Font(name="Times New Roman", size=11)

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

    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
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

# Quản lý trạng thái SBD bằng hàm Callback (Cập nhật SBD ngay lập tức không bị delay)
if "sbd_val" not in st.session_state:
    st.session_state.sbd_val = ""

def add_num(val): st.session_state.sbd_val += str(val)
def clear_num(): st.session_state.sbd_val = ""
def back_num(): st.session_state.sbd_val = st.session_state.sbd_val[:-1]

st.title("🚗 App Chấm Lỗi Đường Trường")
tab1, tab2 = st.tabs(["📝 NHẬP LỖI HỌC VIÊN", "📊 XUẤT BÁO CÁO EXCEL"])

# ------------------------------------------
# TAB 1: BẢNG SỐ VÀ Ô BẤM LỖI
# ------------------------------------------
with tab1:
    ngay_sat_hach_dt = st.date_input("📅 Ngày sát hạch:", date.today())
    ngay_iso = ngay_sat_hach_dt.strftime("%Y-%m-%d")
    ngay_hien_thi = ngay_sat_hach_dt.strftime("%d/%m/%Y")

    # Màn hình hiển thị SBD
    st.markdown("##### 🔢 Số báo danh học viên:")
    sbd_display = st.session_state.sbd_val if st.session_state.sbd_val else "---"
    st.markdown(f"<h2 style='text-align: center; color: #1E88E5; background-color: #E3F2FD; padding: 10px; border-radius: 8px; margin-bottom: 10px;'>{sbd_display}</h2>", unsafe_allow_html=True)

    # Bàn phím số Grid 3x4
    r1c1, r1c2, r1c3 = st.columns(3)
    r1c1.button("1", on_click=add_num, args=(1,), use_container_width=True)
    r1c2.button("2", on_click=add_num, args=(2,), use_container_width=True)
    r1c3.button("3", on_click=add_num, args=(3,), use_container_width=True)

    r2c1, r2c2, r2c3 = st.columns(3)
    r2c1.button("4", on_click=add_num, args=(4,), use_container_width=True)
    r2c2.button("5", on_click=add_num, args=(5,), use_container_width=True)
    r2c3.button("6", on_click=add_num, args=(6,), use_container_width=True)

    r3c1, r3c2, r3c3 = st.columns(3)
    r3c1.button("7", on_click=add_num, args=(7,), use_container_width=True)
    r3c2.button("8", on_click=add_num, args=(8,), use_container_width=True)
    r3c3.button("9", on_click=add_num, args=(9,), use_container_width=True)

    r4c1, r4c2, r4c3 = st.columns(3)
    r4c1.button("🔴", on_click=clear_num, use_container_width=True, help="Xóa hết")
    r4c2.button("0", on_click=add_num, args=(0,), use_container_width=True)
    r4c3.button("⌫", on_click=back_num, use_container_width=True, help="Xóa lùi")

    st.write("### ❌ Chạm vào các lỗi vi phạm:")
    
    with st.form("form_tich_loi", clear_on_submit=True):
        loi_ghi_nhan = {}
        for ma_loi, ten_loi in DANH_SACH_LOI.items():
            loi_ghi_nhan[ma_loi] = st.checkbox(ten_loi)
            
        submit_btn = st.form_submit_button("💾 LƯU KẾT QUẢ", use_container_width=True, type="primary")
        
        if submit_btn:
            if not st.session_state.sbd_val:
                st.error("⚠️ Vui lòng nhập Số báo danh bằng bảng số bên trên!")
            else:
                gia_tri_loi = [1 if loi_ghi_nhan[ma] else 0 for ma in DANH_SACH_LOI.keys()]
                placeholders = ", ".join(["?"] * (len(DANH_SACH_LOI) + 3))
                query = f"INSERT INTO HocVien (ngay_iso, ngay_hien_thi, sbd, {', '.join(DANH_SACH_LOI.keys())}) VALUES ({placeholders})"
                
                c.execute(query, [ngay_iso, ngay_hien_thi, st.session_state.sbd_val] + gia_tri_loi)
                conn.commit()
                st.success(f"✅ Đã lưu kết quả SBD **{st.session_state.sbd_val}** thành công!")
                st.session_state.sbd_val = ""

# ------------------------------------------
# TAB 2: XUẤT VÀ XEM TRƯỚC BÁO CÁO
# ------------------------------------------
with tab2:
    loai_bao_cao = st.radio(
        "Chọn chế độ báo cáo:",
        ["Chỉ 1 ngày cụ thể (Chi tiết từng SBD)", "Từ ngày... Đến ngày... (Tổng hợp cộng dồn)"]
    )
    
    cols_loi_sql = ", ".join(DANH_SACH_LOI.keys())

    if loai_bao_cao == "Chỉ 1 ngày cụ thể (Chi tiết từng SBD)":
        ngay_chon_dt = st.date_input("Chọn ngày muốn xuất:", date.today())
        ngay_chon_iso = ngay_chon_dt.strftime("%Y-%m-%d")
        
        if st.button("📥 TẢI FILE CHI TIẾT", use_container_width=True, type="primary"):
            c.execute(f"SELECT sbd, {cols_loi_sql} FROM HocVien WHERE ngay_iso = ? ORDER BY id ASC", (ngay_chon_iso,))
            rows = c.fetchall()
            
            if not rows:
                st.warning(f"📭 Không có dữ liệu trong ngày này.")
            else:
                excel_bytes = tao_excel_chuan_mau(rows, "Số báo danh")
                st.success(f"🎉 Tạo file thành công! Tổng cộng: {len(rows)} lượt thi.")
                st.download_button(
                    label="⬇️ LƯU FILE EXCEL VỀ MÁY",
                    data=excel_bytes,
                    file_name=f"Chi_Tiet_Loi_{ngay_chon_dt.strftime('%d_%m_%Y')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
                # Bảng xem trước dữ liệu
                st.markdown("---")
                st.write("👀 *Xem trước dữ liệu sẽ xuất:*")
                df_preview = pd.DataFrame(rows, columns=["Số báo danh"] + list(DANH_SACH_LOI.values()))
                df_preview.index = df_preview.index + 1
                st.dataframe(df_preview, use_container_width=True)

    else:
        col_d1, col_d2 = st.columns(2)
        with col_d1:
            tu_ngay_dt = st.date_input("Từ ngày:", date.today())
        with col_d2:
            den_ngay_dt = st.date_input("Đến ngày:", date.today())
            
        tu_ngay_iso = tu_ngay_dt.strftime("%Y-%m-%d")
        den_ngay_iso = den_ngay_dt.strftime("%Y-%m-%d")

        if st.button("📥 TẢI FILE TỔNG HỢP", use_container_width=True, type="primary"):
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
                    label="⬇️ LƯU FILE EXCEL VỀ MÁY",
                    data=excel_bytes,
                    file_name=f"Tong_Hop_Loi_{tu_ngay_dt.strftime('%d%m%Y')}_{den_ngay_dt.strftime('%d%m%Y')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
                # Bảng xem trước dữ liệu
                st.markdown("---")
                st.write("👀 *Xem trước dữ liệu sẽ xuất:*")
                df_preview = pd.DataFrame(rows, columns=["Ngày sát hạch"] + list(DANH_SACH_LOI.values()))
                df_preview.index = df_preview.index + 1
                st.dataframe(df_preview, use_container_width=True)
