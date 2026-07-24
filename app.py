import streamlit as st
import pandas as pd
from datetime import date
import sqlite3
import io
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# ==========================================
# 1. CSS KHÓA CỨNG BẢNG SỐ & LÀM GỌN GIAO DIỆN
# ==========================================
st.markdown("""
<style>
    /* Cắt bớt viền trắng thừa xung quanh app để tiết kiệm diện tích */
    .block-container {
        padding-top: 1.5rem !important;
        padding-bottom: 1rem !important;
        padding-left: 0.5rem !important;
        padding-right: 0.5rem !important;
    }

    /* CHỈ NHẮM VÀO BẢNG SỐ (CÁC HÀNG CÓ 3 CỘT): ÉP BUỘC NẰM NGANG */
    [data-testid="stHorizontalBlock"]:has(> div:nth-child(3)) {
        display: flex !important;
        flex-direction: row !important;
        flex-wrap: nowrap !important;
        gap: 5px !important;
    }
    [data-testid="stHorizontalBlock"]:has(> div:nth-child(3)) > [data-testid="column"] {
        width: 33.33% !important;
        flex: 1 1 33.33% !important;
        min-width: 0 !important;
        padding: 0 !important;
    }
    
    /* Thu nhỏ nút bấm số cho vừa tay, không chiếm chỗ */
    [data-testid="stButton"] button {
        height: 50px !important; 
        min-height: 50px !important;
        font-size: 22px !important;
        font-weight: bold !important;
        border-radius: 8px !important;
        padding: 0 !important;
        margin: 0 !important;
    }
    
    /* Thiết kế thẻ chọn lỗi chạm là đổi màu */
    [data-testid="stCheckbox"] {
        background-color: #f8f9fa;
        padding: 8px 12px;
        border-radius: 8px;
        border: 1px solid #e9ecef;
        margin-bottom: 3px;
    }
    [data-testid="stCheckbox"]:has(input:checked) {
        background-color: #e0f2fe;
        border-color: #3b82f6;
    }
    [data-testid="stCheckbox"] label { cursor: pointer; width: 100%; }
    [data-testid="stCheckbox"] p { font-size: 15px !important; font-weight: bold !important; color: #333; margin: 0;}
    
    /* Box hiển thị SBD */
    .sbd-box {
        text-align: center; 
        color: #1E88E5; 
        background-color: #E3F2FD; 
        padding: 10px; 
        border-radius: 8px;
        font-size: 24px;
        font-weight: bold;
        margin-bottom: 10px;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 2. TỪ ĐIỂN LỖI & CƠ SỞ DỮ LIỆU
# ==========================================
DANH_SACH_LOI = {
    "loi_1": "Không thắt dây an toàn",
    "loi_2": "Không bật xi nhan trái/phải",
    "loi_3": "Không quan sát gương",
    "loi_4": "Dừng, đỗ xe sai quy định",
    "loi_5": "Không chấp hành lệnh biển báo",
    "loi_6": "Mở cửa xe không an toàn",
    "loi_7": "Vượt xe không đảm bảo an toàn",
    "loi_8": "Quay đầu xe sai quy định",
    "loi_9": "Không quan sát, giảm tốc độ",
    "loi_10": "Không chấp hành vạch kẻ đường",
    "loi_11": "Không theo yêu cầu Sát hạch viên",
    "loi_12": "Lỗi khác"
}

conn = sqlite3.connect("dulieu_loi_v9.db", check_same_thread=False)
c = conn.cursor()
cols_sql = ", ".join([f"{col} INTEGER" for col in DANH_SACH_LOI.keys()])
c.execute(f"CREATE TABLE IF NOT EXISTS HocVien (id INTEGER PRIMARY KEY AUTOINCREMENT, ngay_iso TEXT, ngay_hien_thi TEXT, sbd TEXT, {cols_sql})")
conn.commit()

# ==========================================
# 3. HÀM TẠO EXCEL (GIỮ NGUYÊN)
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
        for cell in row: cell.border = thin_border

    ws.column_dimensions['A'].width = 6
    ws.column_dimensions['B'].width = 16
    for i in range(3, len(headers) + 1):
        ws.column_dimensions[get_column_letter(i)].width = 13

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# ==========================================
# 4. GIAO DIỆN CHÍNH
# ==========================================
st.set_page_config(page_title="Chấm Lỗi Sát Hạch", page_icon="🚗", layout="centered")

if "sbd_val" not in st.session_state: st.session_state.sbd_val = ""
if "nhap_xong_sbd" not in st.session_state: st.session_state.nhap_xong_sbd = False

def add_num(val): st.session_state.sbd_val += str(val)
def clear_num(): st.session_state.sbd_val = ""
def back_num(): st.session_state.sbd_val = st.session_state.sbd_val[:-1]

def xac_nhan():
    if st.session_state.sbd_val: st.session_state.nhap_xong_sbd = True
    else: st.error("⚠️ Phải nhập số báo danh trước!")

def huy_sbd(): st.session_state.nhap_xong_sbd = False

st.title("🚗 Chấm Lỗi Đường Trường")
tab1, tab2 = st.tabs(["📝 NHẬP LỖI", "📊 BÁO CÁO"])

# ------------------------------------------
# TAB 1: NHẬP LIỆU (Định dạng ngày dd/mm/yyyy)
# ------------------------------------------
with tab1:
    ngay_sat_hach_dt = st.date_input("📅 Ngày sát hạch:", date.today(), format="DD/MM/YYYY")
    ngay_iso = ngay_sat_hach_dt.strftime("%Y-%m-%d")
    ngay_hien_thi = ngay_sat_hach_dt.strftime("%d/%m/%Y")

    # ----- BẢNG SỐ SIÊU GỌN -----
    if not st.session_state.nhap_xong_sbd:
        sbd_display = st.session_state.sbd_val if st.session_state.sbd_val else "---"
        st.markdown(f"<div class='sbd-box'>SBD: {sbd_display}</div>", unsafe_allow_html=True)

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
        r4c1.button("Xóa", on_click=clear_num, use_container_width=True)
        r4c2.button("0", on_click=add_num, args=(0,), use_container_width=True)
        r4c3.button("Lùi", on_click=back_num, use_container_width=True)

        st.button("✅ TIẾP TỤC (CHỌN LỖI)", on_click=xac_nhan, use_container_width=True, type="primary")

    # ----- BẢNG LỖI -----
    else:
        colA, colB = st.columns([3, 1])
        with colA: st.markdown(f"<h3 style='color: #1E88E5; margin: 0;'>Đang chấm SBD: {st.session_state.sbd_val}</h3>", unsafe_allow_html=True)
        with colB: st.button("🔄 Sửa", on_click=huy_sbd, use_container_width=True)

        st.write("---")
        with st.form("form_tich_loi", clear_on_submit=True):
            loi_ghi_nhan = {}
            for ma_loi, ten_loi in DANH_SACH_LOI.items():
                loi_ghi_nhan[ma_loi] = st.checkbox(ten_loi)
                
            if st.form_submit_button("💾 LƯU VÀ TIẾP TỤC", use_container_width=True, type="primary"):
                gia_tri_loi = [1 if loi_ghi_nhan[ma] else 0 for ma in DANH_SACH_LOI.keys()]
                query = f"INSERT INTO HocVien (ngay_iso, ngay_hien_thi, sbd, {', '.join(DANH_SACH_LOI.keys())}) VALUES ({', '.join(['?']*(len(DANH_SACH_LOI)+3))})"
                c.execute(query, [ngay_iso, ngay_hien_thi, st.session_state.sbd_val] + gia_tri_loi)
                conn.commit()
                
                st.success(f"✅ Đã lưu thành công SBD **{st.session_state.sbd_val}**!")
                st.session_state.sbd_val = ""
                st.session_state.nhap_xong_sbd = False
                st.rerun()

# ------------------------------------------
# TAB 2: XUẤT BÁO CÁO (Định dạng ngày dd/mm/yyyy)
# ------------------------------------------
with tab2:
    loai_bao_cao = st.radio("Chọn loại báo cáo:", ["Chi tiết (1 ngày)", "Tổng hợp (Từ ngày - Đến ngày)"])
    cols_loi_sql = ", ".join(DANH_SACH_LOI.keys())

    if loai_bao_cao == "Chi tiết (1 ngày)":
        ngay_chon_dt = st.date_input("Chọn ngày xuất:", date.today(), format="DD/MM/YYYY")
        
        if st.button("📥 TẢI EXCEL", use_container_width=True, type="primary"):
            c.execute(f"SELECT sbd, {cols_loi_sql} FROM HocVien WHERE ngay_iso = ? ORDER BY id ASC", (ngay_chon_dt.strftime("%Y-%m-%d"),))
            rows = c.fetchall()
            if not rows: st.warning("📭 Không có dữ liệu.")
            else:
                excel_bytes = tao_excel_chuan_mau(rows, "Số báo danh")
                st.download_button("⬇️ LƯU FILE VỀ MÁY", data=excel_bytes, file_name=f"Chi_Tiet_{ngay_chon_dt.strftime('%d_%m_%Y')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                df_preview = pd.DataFrame(rows, columns=["SBD"] + list(DANH_SACH_LOI.values()))
                df_preview.index = df_preview.index + 1
                st.dataframe(df_preview, use_container_width=True)

    else:
        c1, c2 = st.columns(2)
        with c1: tu_ngay_dt = st.date_input("Từ ngày:", date.today(), format="DD/MM/YYYY")
        with c2: den_ngay_dt = st.date_input("Đến ngày:", date.today(), format="DD/MM/YYYY")

        if st.button("📥 TẢI EXCEL TỔNG HỢP", use_container_width=True, type="primary"):
            sums_sql = ", ".join([f"SUM({k})" for k in DANH_SACH_LOI.keys()])
            query = f"SELECT ngay_hien_thi, {sums_sql} FROM HocVien WHERE ngay_iso BETWEEN ? AND ? GROUP BY ngay_iso ORDER BY ngay_iso ASC"
            c.execute(query, (tu_ngay_dt.strftime("%Y-%m-%d"), den_ngay_dt.strftime("%Y-%m-%d")))
            rows = c.fetchall()
            if not rows: st.warning("📭 Không có dữ liệu.")
            else:
                excel_bytes = tao_excel_chuan_mau(rows, "Ngày sát hạch")
                st.download_button("⬇️ LƯU FILE VỀ MÁY", data=excel_bytes, file_name=f"Tong_Hop_{tu_ngay_dt.strftime('%d%m%Y')}_{den_ngay_dt.strftime('%d%m%Y')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                df_preview = pd.DataFrame(rows, columns=["Ngày"] + list(DANH_SACH_LOI.values()))
                df_preview.index = df_preview.index + 1
                st.dataframe(df_preview, use_container_width=True)
