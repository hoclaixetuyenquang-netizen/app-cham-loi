import streamlit as st
import pandas as pd
from datetime import date, datetime, timedelta
import io
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import os

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

# ===== ĐỌC/GHI FILE CSV =====
CSV_FILE = "dulieu_loi.csv"

def load_data():
    """Đọc dữ liệu từ CSV"""
    if os.path.exists(CSV_FILE):
        return pd.read_csv(CSV_FILE)
    return pd.DataFrame()

def save_data(df):
    """Lưu dữ liệu vào CSV"""
    df.to_csv(CSV_FILE, index=False, encoding='utf-8-sig')

# Danh sách cột lỗi
columns_display = [
    "Không thắt dây an toàn", "Không bật xi nhan trái/phải", "Không quan sát gương",
    "Dừng, đỗ xe sai quy định", "Không chấp hành hiệu lệnh", "Mở cửa xe không an toàn",
    "Vượt xe không đảm bảo", "Quay đầu xe không đúng", "Không quan sát, giảm tốc",
    "Lỗi vạch kẻ đường", "Không thực hiện theo yêu cầu", "Lỗi khác"
]

# Khởi tạo session state
if "selected_errors" not in st.session_state:
    st.session_state.selected_errors = set()

if "save_success" not in st.session_state:
    st.session_state.save_success = False

# ===== GIAO DIỆN NHẬP LIỆU =====
st.title("THỐNG KÊ LỖI ĐƯỜNG TRƯỜNG")

# Chọn ngày
ngay_sat_hach = st.date_input("📅 Ngày sát hạch", date.today(), format="DD/MM/YYYY")

# Form tích lỗi với nút bấm
st.write("**✓ Tích vào các lỗi học viên mắc phải (nhấn để chọn/bỏ chọn):**")

# Hiển thị thông báo nếu vừa lưu thành công
if st.session_state.save_success:
    st.success("✅ Đã lưu thành công!", icon="✅")
    st.balloons()
    st.session_state.save_success = False

# Tạo các nút bấm lỗi
for idx, display_text in enumerate(columns_display):
    is_selected = idx in st.session_state.selected_errors
    button_color = "🔴" if is_selected else "⚪"
    
    if st.button(f"{button_color} {display_text}", key=f"error_btn_{idx}", use_container_width=True):
        # Toggle selection
        if idx in st.session_state.selected_errors:
            st.session_state.selected_errors.remove(idx)
        else:
            st.session_state.selected_errors.add(idx)

        # Auto-save current selections to CSV immediately after toggle
        loi_values = [1 if i in st.session_state.selected_errors else 0 for i in range(len(columns_display))]
        row_data = {"Ngày": str(ngay_sat_hach)}
        for i, col in enumerate(columns_display):
            row_data[col] = loi_values[i]

        # Đọc dữ liệu cũ
        df = load_data()

        # Thêm hàng mới
        df = pd.concat([df, pd.DataFrame([row_data])], ignore_index=True)

        # Lưu vào CSV
        save_data(df)

        # Phát âm thanh và hiển thị thông báo nhỏ
        play_sound()
        # Reset selection so the UI returns to initial state for next input
        st.session_state.selected_errors = set()
        st.session_state.save_success = True
        st.rerun()

# Các nút hành động: bỏ nút Lưu và Xóa (tự động lưu khi chọn lỗi); chỉ hiển thị số lỗi chọn
st.divider()
st.metric("Số lỗi chọn", len(st.session_state.selected_errors))

# ===== XUẤT BÁO CÁO =====
st.divider()
st.subheader("📊 Báo Cáo Tổng Hợp")

# Thống kê nhanh
df_all = load_data()
if not df_all.empty:
    st.info(f"📈 Tổng số lần ghi nhận: **{len(df_all)}** | Số ngày: **{df_all['Ngày'].nunique()}**")

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

# Danh sách tên lỗi đầy đủ
full_error_names = [
    "Không thắt dây an toàn",
    "Không bật xi nhan trái/phải",
    "Không quan sát gương",
    "Dừng, đỗ xe sai quy định",
    "Không chấp hành hiệu lệnh",
    "Mở cửa xe không an toàn",
    "Vượt xe không đảm bảo",
    "Quay đầu xe không đúng",
    "Không quan sát, giảm tốc độ để hoạc dừng các trương hợp theo quy định",
    "Lỗi vạch kẻ đường",
    "Không thực hiện theo yêu cầu của Sát hạch viên",
    "Lỗi khác"
]

def create_report_excel(df_tong_hop):
    """Tạo file Excel với định dạng theo mẫu (vị trí cột động, tổng đúng cột)"""
    wb = Workbook()
    ws = wb.active
    ws.title = "TongHopLoi"

    # Fonts / styles
    title_font = Font(name='Times New Roman', size=14, bold=True)
    total_font = Font(name='Times New Roman', size=11, bold=True, color="CC0000")
    header_font = Font(name='Times New Roman', size=10, bold=True)
    data_font = Font(name='Times New Roman', size=10)

    header_fill = PatternFill(start_color="B4C7E7", end_color="B4C7E7", fill_type="solid")
    total_fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")

    center_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    border = Border(
        left=Side(style='thin', color="000000"),
        right=Side(style='thin', color="000000"),
        top=Side(style='thin', color="000000"),
        bottom=Side(style='thin', color="000000")
    )

    # Hợp nhất tiêu đề theo số cột thực tế
    last_col_letter = get_column_letter(len(df_tong_hop.columns))
    ws.merge_cells(f'A1:{last_col_letter}1')
    title_cell = ws['A1']
    title_cell.value = "Số lượng lỗi do sát hạch viên trừ trong phần thi đường trường"
    title_cell.font = title_font
    title_cell.alignment = center_alignment
    ws.row_dimensions[1].height = 30

    # Tạo mapping header -> index (1-based Excel columns)
    header_positions = {col_name: idx + 1 for idx, col_name in enumerate(df_tong_hop.columns)}

    # Hàng tổng cộng (Row 2) — ghi 'TỔNG CỘNG' ở cột A và các giá trị tổng ở đúng vị trí theo header_positions
    total_row_data = df_tong_hop.iloc[0]
    ws.cell(row=2, column=1).value = "TỔNG CỘNG"
    ws.cell(row=2, column=1).font = total_font
    ws.cell(row=2, column=1).fill = total_fill
    ws.cell(row=2, column=1).alignment = center_alignment
    ws.cell(row=2, column=1).border = border

    # Ghi tổng bắt đầu ở vị trí của mỗi header (bỏ STT và cột ngày nếu có)
    for col_name in df_tong_hop.columns:
        if col_name in ('STT',) or 'ngày' in col_name.lower():
            continue
        col_idx = header_positions[col_name]
        cell = ws.cell(row=2, column=col_idx)
        value = total_row_data[col_name]
        try:
            cell.value = int(value) if pd.notna(value) else 0
        except (ValueError, TypeError):
            cell.value = value if pd.notna(value) else 0
        cell.font = total_font
        cell.fill = total_fill
        cell.alignment = center_alignment
        cell.border = border

    # Hàng header (Row 3) theo header_positions (đảm bảo đúng thứ tự)
    for col_name, col_idx in header_positions.items():
        cell = ws.cell(row=3, column=col_idx)
        cell.value = col_name
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_alignment
        cell.border = border

    ws.row_dimensions[3].height = 40

    # Dữ liệu (từ Row 4 trở đi)
    for row_idx, (_, row_data) in enumerate(df_tong_hop.iloc[1:].iterrows(), start=4):
        for col_name, col_idx in header_positions.items():
            cell = ws.cell(row=row_idx, column=col_idx)
            value = row_data[col_name]
            if col_name == 'STT':
                cell.value = value if value != '' else ''
            elif 'ngày' in col_name.lower():
                cell.value = value
            else:
                try:
                    cell.value = int(value) if pd.notna(value) else 0
                except (ValueError, TypeError):
                    cell.value = value if pd.notna(value) else 0
            cell.font = data_font
            cell.alignment = center_alignment
            cell.border = border

    # Đặt chiều rộng cột động: STT nhỏ, Ngày vừa, các cột còn lại rộng hơn
    if 'STT' in header_positions:
        ws.column_dimensions[get_column_letter(header_positions['STT'])].width = 8
    # tìm cột ngày (bất kỳ tên chứa 'ngày')
    for name, idx in header_positions.items():
        if 'ngày' in name.lower():
            ws.column_dimensions[get_column_letter(idx)].width = 15
            break
    for name, idx in header_positions.items():
        if name not in ('STT',) and 'ngày' not in name.lower():
            ws.column_dimensions[get_column_letter(idx)].width = 18

    # Lưu vào BytesIO và trả về bytes
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()

if st.button("📥 Tạo Báo Cáo", use_container_width=True):
    df = load_data()
    
    if df.empty:
        st.warning("⚠️ Chưa có dữ liệu để xuất.")
    else:
        # Áp dụng bộ lọc
        if filter_type == "Một ngày":
            df = df[df['Ngày'] == str(filter_date)]
        elif filter_type == "Từ ngày đến ngày":
            df = df[(df['Ngày'] >= str(filter_date_from)) & (df['Ngày'] <= str(filter_date_to))]
        
        if df.empty:
            st.warning("⚠️ Không có dữ liệu cho khoảng thời gian đã chọn.")
        else:
            # Gom nhóm và tính tổng
            numeric_cols = [col for col in df.columns if col != 'Ngày']
            # Ép các cột numeric về số (nếu có giá trị không phải số sẽ thành NaN rồi fill 0)
            df[numeric_cols] = df[numeric_cols].apply(pd.to_numeric, errors='coerce').fillna(0)
            df_tong_hop = df.groupby('Ngày')[numeric_cols].sum().reset_index()
            
            # Tính tổng cộng
            total_row = {'Ngày': 'TỔNG CỘNG'}
            for col in numeric_cols:
                total_row[col] = df_tong_hop[col].sum()
            
            # Thêm hàng tổng cộng vào đầu
            df_tong_hop = pd.concat([pd.DataFrame([total_row]), df_tong_hop], ignore_index=True)
            
            # Thêm cột STT
            stt = [''] + list(range(1, len(df_tong_hop)))
            df_tong_hop.insert(0, 'STT', stt)
            
            # Tạo file Excel
            excel_data = create_report_excel(df_tong_hop)
            
            st.download_button(
                label="📥 Tải File Excel",
                data=excel_data,
                file_name=f"Bao_cao_loi_{date.today().strftime('%d_%m_%Y')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
            # Hiển thị preview
            with st.expander("👁️ Xem trước dữ liệu"):
                st.dataframe(df_tong_hop, use_container_width=True)

# ===== QUẢN LÝ DỮ LIỆU =====
with st.expander("⚙️ Quản Lý Dữ Liệu"):
    st.subheader("Xem Dữ Liệu")
    col_manage1, col_manage2 = st.columns(2)
    
    with col_manage1:
        if st.button("🔍 Xem Tất Cả Dữ Liệu", use_container_width=True):
            df_view = load_data()
            if not df_view.empty:
                st.dataframe(df_view, use_container_width=True)
            else:
                st.info("Không có dữ liệu")
    
    with col_manage2:
        st.subheader("Xóa Dữ Liệu")
        password_input = st.text_input("Nhập mật khẩu:", type="password", key="del_password")
        if st.button("🗑️ Xóa Tất Cả", use_container_width=True):
            if password_input == "Admin@1234":
                # Xóa file CSV
                if os.path.exists(CSV_FILE):
                    os.remove(CSV_FILE)
                    st.success("✅ Đã xóa tất cả dữ liệu")
                    st.rerun()
                else:
                    st.info("Không có dữ liệu để xóa")
            else:
                st.error("❌ Mật khẩu không chính xác!")
