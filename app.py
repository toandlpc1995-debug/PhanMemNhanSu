import streamlit as st
import docx
import pandas as pd
import re
import os
from io import BytesIO
from docx.document import Document as _Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
from datetime import datetime
import time

# --- CẤU HÌNH TRANG WEB ---
st.set_page_config(page_title="Phần mềm Tổng hợp Nhân sự", layout="wide", initial_sidebar_state="expanded")
st.title("PHẦN MỀM TỔNG HỢP NHÂN SỰ TỪ SƠ YẾU LÝ LỊCH")

LOCAL_DB = "DuLieu_NhanSu.csv"

def load_local_db():
    if os.path.exists(LOCAL_DB):
        df = pd.read_csv(LOCAL_DB, dtype=str)
        df['SoHieu'] = df['SoHieu'].str.strip()
        df.drop_duplicates(subset=['SoHieu'], keep='last', inplace=True)
        return df
    return pd.DataFrame(columns=["SoHieu", "HoTen", "DonVi", "KhenThuong", "KyLuat", "QuaTrinhCongTac"])

def save_local_db(df):
    df.to_csv(LOCAL_DB, index=False)

def convert_df_to_excel(df, sheet_name="BaoCao"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()

# --- THUẬT TOÁN ĐỌC WORD ---
def iter_block_items(parent):
    if isinstance(parent, _Document): parent_elm = parent.element.body
    elif isinstance(parent, _Cell): parent_elm = parent._tc
    else: raise ValueError("Lỗi cấu trúc Word")

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P): yield Paragraph(child, parent).text
        elif isinstance(child, CT_Tbl):
            table = Table(child, parent)
            for row in table.rows: yield " | ".join([cell.text for cell in row.cells])

def deduplicate_table_cells(text):
    if not text: return ""
    cells = [c.strip() for c in text.split('|') if c.strip()]
    if not cells: return ""
    seen = set()
    unique_cells = [x for x in cells if not (x in seen or seen.add(x))]
    return " - ".join(unique_cells)

def extract_data_from_docx(file):
    doc = docx.Document(file)
    text = '\n'.join(list(iter_block_items(doc)))
    data = {"SoHieu": "", "HoTen": "", "DonVi": "", "KhenThuong": "", "KyLuat": "", "QuaTrinhCongTac": ""}
    
    try:
        sh_match = re.search(r"Số hiệu:\s*([a-zA-Z0-9_\-]+)", text, re.IGNORECASE)
        if sh_match: data["SoHieu"] = sh_match.group(1).strip()
            
        ht_match = re.search(r"1-\s*Họ và tên khai sinh:(.*?)(?:Nam, nữ|2-)", text, re.IGNORECASE | re.DOTALL)
        if ht_match: 
            raw_name = re.split(r'Họ và tên khai sinh:', ht_match.group(1), flags=re.IGNORECASE)[-1]
            data["HoTen"] = deduplicate_table_cells(raw_name.replace('\n', ' '))
            
        dv_match = re.search(r"3\s*-\s*Chức danh.*?công tác:(.*?)(?=- Chức vụ|4-)", text, re.IGNORECASE | re.DOTALL)
        if dv_match: 
            raw_dv = re.split(r'đơn vị công tác:', dv_match.group(1), flags=re.IGNORECASE)[-1]
            data["DonVi"] = deduplicate_table_cells(raw_dv.replace('\n', ' '))
        
        kt_match = re.search(r"24-\s*Khen thưởng.*?:(.*?)25-\s*Kỷ luật", text, re.DOTALL | re.IGNORECASE)
        if kt_match: data["KhenThuong"] = kt_match.group(1).strip()
            
        kl_match = re.search(r"25-\s*Kỷ luật.*?:(.*?)26-\s*Đào tạo", text, re.DOTALL | re.IGNORECASE)
        if kl_match: data["KyLuat"] = kl_match.group(1).strip()
            
        qt_match = re.search(r"27-\s*Tóm tắt quá trình công tác.*?:(.*?)28-\s*Đặc điểm", text, re.DOTALL | re.IGNORECASE)
        if qt_match: data["QuaTrinhCongTac"] = qt_match.group(1).strip()
    except Exception: pass 
    return data

def parse_records(text):
    if not text or pd.isna(text): return []
    parts = re.split(r'(?i)Năm\s+(\d{4})', str(text))
    parsed = []
    for i in range(1, len(parts), 2):
        year = int(parts[i])
        content_raw = parts[i+1].strip(' \n\r\t.-,;:)(')
        sub_contents = [c.strip(' \n\r\t.-,;:)(') for c in re.split(r'\.\)|\n', content_raw)]
        for sub in sub_contents:
            if len(sub) > 2: parsed.append({"Năm": year, "Nội dung": sub})
    parsed.sort(key=lambda x: x['Năm'], reverse=True) 
    return parsed

def parse_date(date_str):
    date_str = str(date_str).strip().lower()
    # Hỗ trợ cực mạnh chữ "nay" và "hiện tại"
    if 'nay' in date_str or 'hiện tại' in date_str: return datetime.now()
    nums = re.findall(r'\d+', date_str)
    try:
        if len(nums) == 2: return datetime(int(nums[1]), int(nums[0]), 1)
        elif len(nums) >= 3: return datetime(int(nums[2]), int(nums[1]), int(nums[0]))
    except: pass
    return None

def tao_bang_gop_o(df_data, col_name, is_ca_nhan=False):
    display_data = []
    for _, row in df_data.iterrows():
        records = parse_records(row[col_name])
        if 'TuNam' in row and 'DenNam' in row:
            records = [r for r in records if row['TuNam'] <= r['Năm'] <= row['DenNam']]
        if not records: continue
        
        records_by_year = {}
        for r in records:
            records_by_year.setdefault(r['Năm'], []).append(r['Nội dung'])
            
        first_person = True
        for yr in sorted(records_by_year.keys(), reverse=True):
            contents = records_by_year[yr]
            first_year = True
            for content in contents:
                if is_ca_nhan:
                    display_data.append({"Năm": str(yr) if first_year else "", "Nội dung chi tiết": content})
                else:
                    display_data.append({
                        "Số hiệu": str(row['SoHieu']) if first_person else "",
                        "Họ và Tên": str(row['HoTen']) if first_person else "",
                        "Đơn vị": str(row['DonVi']) if first_person else "",
                        "Năm": str(yr) if first_year else "",
                        "Nội dung chi tiết": content
                    })
                first_person = False
                first_year = False
    return pd.DataFrame(display_data)

# ==========================================
# GIAO DIỆN CHÍNH
# ==========================================
df_employees = load_local_db()

st.sidebar.markdown("### 💡 Giao diện Tối / Sáng")
st.sidebar.info("Bật/tắt Dark Mode: Bấm vào **Menu (⋮)** góc phải trên cùng -> **Settings** -> Mục **Theme** chọn Dark hoặc Light.")

st.sidebar.header("📥 Nạp dữ liệu Sơ yếu lý lịch")
uploaded_files = st.sidebar.file_uploader("Tải lên file Word", type=['docx'], accept_multiple_files=True)

if uploaded_files:
    if st.sidebar.button("Đọc và Lưu vào Cơ sở dữ liệu"):
        new_data, failed_files = [], []
        for file in uploaded_files:
            emp_data = extract_data_from_docx(file)
            if emp_data["HoTen"]: new_data.append(emp_data)
            else: failed_files.append(file.name)
        
        if new_data:
            df_new = pd.DataFrame(new_data)
            df_new['SoHieu'] = df_new['SoHieu'].astype(str).str.strip()
            if not df_employees.empty:
                df_employees['SoHieu'] = df_employees['SoHieu'].astype(str).str.strip()
                df_combined = pd.concat([df_employees, df_new]).drop_duplicates(subset=['SoHieu'], keep='last')
            else: df_combined = df_new
            save_local_db(df_combined)
            st.session_state["thong_bao_success"] = f"Đã nạp thành công {len(new_data)} bộ hồ sơ!"
            time.sleep(0.5)
            st.rerun()

if "thong_bao_success" in st.session_state:
    st.sidebar.success(st.session_state["thong_bao_success"])
    del st.session_state["thong_bao_success"]

st.caption(f"Trạng thái hệ thống: Đang quản lý {len(df_employees)} nhân sự.")

danh_sach_tim_kiem = []
if not df_employees.empty:
    danh_sach_tim_kiem = df_employees.apply(lambda row: f"[{str(row['SoHieu'])}] {str(row['HoTen'])}", axis=1).tolist()

tab1, tab2, tab3 = st.tabs(["1. Khen thưởng", "2. Kỷ luật", "3. Quy đổi hệ số năm công tác"])

# ==========================================
# TAB 1: KHEN THƯỞNG
# ==========================================
with tab1:
    st.header("Báo cáo Khen thưởng")
    if not df_employees.empty:
        loai_loc_kt = st.radio("Lọc Khen thưởng theo:", ("Theo Giai đoạn (Từ năm - Đến năm)", "Theo Cán bộ / Nhân viên"), key="radio_kt")
        if loai_loc_kt == "Theo Giai đoạn (Từ năm - Đến năm)":
            col1, col2, col3 = st.columns([3, 3, 4])
            with col1: tu_nam_kt = st.number_input("Từ năm:", min_value=1950, max_value=2050, value=2015, key="tu_kt")
            with col2: den_nam_kt = st.number_input("Đến năm:", min_value=1950, max_value=2050, value=2025, key="den_kt")
            
            df_temp = df_employees.copy()
            df_temp['TuNam'], df_temp['DenNam'] = tu_nam_kt, den_nam_kt
            df_show = tao_bang_gop_o(df_temp, 'KhenThuong')
            
            with col3:
                st.write("")
                st.write("")
                if not df_show.empty:
                    st.download_button("📥 Xuất Excel (Giai đoạn)", data=convert_df_to_excel(df_show, "KhenThuong"), file_name=f"KhenThuong_{tu_nam_kt}_{den_nam_kt}.xlsx")

            if not df_show.empty: st.dataframe(df_show, use_container_width=True, height=600, hide_index=True)
            else: st.warning(f"Không có khen thưởng trong giai đoạn {tu_nam_kt} - {den_nam_kt}")
        else:
            col4, col5 = st.columns([7, 3])
            with col4: ns_chon_kt = st.selectbox("Tìm kiếm Cán bộ / Nhân viên:", danh_sach_tim_kiem, key="sb_kt")
            if ns_chon_kt:
                so_hieu_chon = ns_chon_kt.split(']')[0].replace('[', '')
                thong_tin = df_employees[df_employees['SoHieu'] == so_hieu_chon].iloc[[0]].copy()
                df_show = tao_bang_gop_o(thong_tin, 'KhenThuong', is_ca_nhan=True)
                
                with col5:
                    st.write("")
                    st.write("")
                    if not df_show.empty:
                        st.download_button("📥 Xuất Excel (Cá nhân)", data=convert_df_to_excel(df_show, "KhenThuong"), file_name=f"KhenThuong_{so_hieu_chon}.xlsx")
                
                st.info(f"**Chi tiết Khen thưởng của: {thong_tin.iloc[0]['HoTen']} (Số hiệu: {thong_tin.iloc[0]['SoHieu']})**")
                if not df_show.empty: st.dataframe(df_show, use_container_width=True, height=400, hide_index=True)
                else: st.write("Không có dữ liệu khen thưởng.")
    else: st.info("Hệ thống chưa có dữ liệu.")

# ==========================================
# TAB 2: KỶ LUẬT
# ==========================================
with tab2:
    st.header("Báo cáo Kỷ luật")
    if not df_employees.empty:
        loai_loc_kl = st.radio("Lọc Kỷ luật theo:", ("Theo Giai đoạn (Từ năm - Đến năm)", "Theo Cán bộ / Nhân viên"), key="radio_kl")
        if loai_loc_kl == "Theo Giai đoạn (Từ năm - Đến năm)":
            col6, col7, col8 = st.columns([3, 3, 4])
            with col6: tu_nam_kl = st.number_input("Từ năm:", min_value=1950, max_value=2050, value=2015, key="tu_kl")
            with col7: den_nam_kl = st.number_input("Đến năm:", min_value=1950, max_value=2050, value=2025, key="den_kl")
            
            df_temp = df_employees.copy()
            df_temp['TuNam'], df_temp['DenNam'] = tu_nam_kl, den_nam_kl
            df_show = tao_bang_gop_o(df_temp, 'KyLuat')
            
            with col8:
                st.write("")
                st.write("")
                if not df_show.empty:
                    st.download_button("📥 Xuất Excel (Giai đoạn)", data=convert_df_to_excel(df_show, "KyLuat"), file_name=f"KyLuat_{tu_nam_kl}_{den_nam_kl}.xlsx")
            
            if not df_show.empty: st.dataframe(df_show, use_container_width=True, height=600, hide_index=True)
            else: st.warning("Không có kỷ luật nào trong giai đoạn này.")
        else:
            col9, col10 = st.columns([7, 3])
            with col9: ns_chon_kl = st.selectbox("Tìm kiếm Cán bộ / Nhân viên:", danh_sach_tim_kiem, key="sb_kl")
            if ns_chon_kl:
                so_hieu_chon = ns_chon_kl.split(']')[0].replace('[', '')
                thong_tin = df_employees[df_employees['SoHieu'] == so_hieu_chon].iloc[[0]].copy()
                df_show = tao_bang_gop_o(thong_tin, 'KyLuat', is_ca_nhan=True)
                
                with col10:
                    st.write("")
                    st.write("")
                    if not df_show.empty:
                        st.download_button("📥 Xuất Excel (Cá nhân)", data=convert_df_to_excel(df_show, "KyLuat"), file_name=f"KyLuat_{so_hieu_chon}.xlsx")
                
                st.info(f"**Chi tiết Kỷ luật của: {thong_tin.iloc[0]['HoTen']} (Số hiệu: {thong_tin.iloc[0]['SoHieu']})**")
                if not df_show.empty: st.dataframe(df_show, use_container_width=True, height=400, hide_index=True)
                else: st.write("Không có dữ liệu kỷ luật.")
    else: st.info("Hệ thống chưa có dữ liệu.")

# ==========================================
# TAB 3: QUY ĐỔI HỆ SỐ NĂM CÔNG TÁC 
# ==========================================
with tab3:
    st.header("Tính toán quy đổi thời gian công tác")
    
    if not df_employees.empty:
        blacklist = ["đảng", "bí thư", "chi ủy", "chi uỷ", "cấp ủy", "cấp uỷ", "thường vụ", "uỷ viên", "ủy viên", "chi bộ", "đoàn", "đoàn thanh niên", "chi đoàn", "chiến sĩ", "bộ đội", "quân sự", "tổng cục kỹ thuật", "chi hội", "dân quân"]
        
        chuc_danh_mapping = {
            "phó chủ tịch công đoàn": "Phó Chủ tịch Công đoàn",
            "chủ tịch công đoàn": "Chủ tịch Công đoàn",
            "phó giám đốc điện lực": "Phó Giám đốc Điện lực",
            "giám đốc điện lực": "Giám đốc Điện lực",
            "phó trưởng điện lực": "Phó Trưởng Điện lực",
            "trưởng điện lực": "Trưởng Điện lực",
            "phó chánh văn phòng": "Phó Chánh Văn phòng",
            "chánh văn phòng": "Chánh Văn phòng",
            "phó chi nhánh trưởng": "Phó Chi nhánh",
            "phó chi nhánh": "Phó Chi nhánh",
            "chi nhánh trưởng": "Chi nhánh trưởng",
            "phó giám đốc": "Phó Giám đốc Công ty",
            "giám đốc": "Giám đốc Công ty",
            "kế toán trưởng": "Kế toán trưởng",
            "phó trưởng phòng": "Phó phòng",
            "phó phòng": "Phó phòng",
            "trưởng phòng": "Trưởng phòng",
            "phó trưởng ban": "Phó ban",
            "phó ban": "Phó ban",
            "trưởng ban": "Trưởng ban",
            "đội phó": "Đội phó",
            "Công nhân": "Công nhân",
            "đội trưởng": "Đội trưởng"
        }

        st.subheader("1. Bảng thiết lập Hệ số Quy đổi")
        st.write("💡 Các chức danh Đảng/Đoàn/Quân sự đã được tự động loại trừ. Các chức danh không nằm trong danh sách dưới đây sẽ tự động tính hệ số = 1.0.")
        
        he_so_dict = {}
        unique_mapped_titles = sorted(list(set(chuc_danh_mapping.values())))
        
        cols = st.columns(4)
        for i, chuc_danh in enumerate(unique_mapped_titles):
            with cols[i % 4]:
                val = st.text_input(f"{chuc_danh}", value="1.0", key=f"hs_{i}")
                he_so_dict[chuc_danh] = float(val) if val.replace('.','',1).isdigit() else 1.0

        st.divider()
        
        loai_loc_t3 = st.radio("Tùy chọn xuất báo cáo Hệ số:", ("Xuất Tất cả Cán bộ / Nhân viên", "Xuất theo từng Cán bộ / Nhân viên"), key="radio_t3")
        
        df_target = df_employees.copy()
        if loai_loc_t3 == "Xuất theo từng Cán bộ / Nhân viên":
            col_search, _ = st.columns([7, 3])
            with col_search:
                ns_chon_t3 = st.selectbox("Tìm kiếm Cán bộ / Nhân viên:", danh_sach_tim_kiem, key="sb_t3")
            if ns_chon_t3:
                so_hieu_chon = ns_chon_t3.split(']')[0].replace('[', '')
                df_target = df_employees[df_employees['SoHieu'] == so_hieu_chon]
            else:
                df_target = pd.DataFrame()

        col11, col12 = st.columns([7, 3])
        with col11: st.subheader("2. Chi tiết Quy đổi")
        
        all_results_for_excel = [] 
        
        if not df_target.empty:
            for index, row in df_target.iterrows():
                tong_thang_thuc = 0
                tong_thang_quydoi = 0
                person_data = []
                
                for dong in str(row['QuaTrinhCongTac']).split('\n'):
                    dong = dong.strip()
                    dong = re.sub(r'^[-+*•]\s*', '', dong) # Dọn dẹp dấu gạch đầu dòng
                    if not dong: continue
                    
                    # CẢI TIẾN SIÊU MẠNH: Bắt mọi thể loại thời gian (có hai chấm, không hai chấm, gạch ngang, v.v)
                    match = re.search(r'(?i)^(?:từ\s+)?(.*?)\s+(?:đến|-)\s+(.*?)\s*[:,\-]\s*(.*)$', dong)
                    if not match:
                        match = re.search(r'(?i)^(?:từ\s+)?(.*?)\s+(?:đến|-)\s+(nay|hiện tại|\d{1,2}[/-]\d{4}|\d{4}|tháng\s+\d{1,2}[/-]\d{4})\s+(.*)$', dong)
                        
                    if match:
                        start_str = match.group(1).strip()
                        end_str = match.group(2).strip()
                        chuc_danh_day_du = match.group(3).strip()
                        
                        start_date = parse_date(start_str)
                        end_date = parse_date(end_str)
                        
                        if start_date and end_date:
                            title_lower = chuc_danh_day_du.lower()
                            if any(b in title_lower for b in blacklist): continue 
                                
                            so_thang = max(0, (end_date.year - start_date.year) * 12 + (end_date.month - start_date.month))
                            tong_thang_thuc += so_thang
                            
                            nhom_chuc_danh = "Khác (Hệ số 1.0)"
                            hs_apdung = 1.0
                            
                            for key_nhan_dien, ten_nhom in chuc_danh_mapping.items():
                                if key_nhan_dien in title_lower:
                                    nhom_chuc_danh = ten_nhom
                                    hs_apdung = he_so_dict.get(ten_nhom, 1.0)
                                    break
                                    
                            thang_qd = so_thang * hs_apdung
                            tong_thang_quydoi += thang_qd
                            
                            person_data.append({
                                "Giai đoạn": f"{start_str} ➔ {end_str}",
                                "Chức danh (Gốc)": chuc_danh_day_du,
                                "Nhóm áp dụng": nhom_chuc_danh,
                                "Thực (Năm)": so_thang // 12,
                                "Thực (Tháng)": so_thang % 12,
                                "Hệ số": hs_apdung,
                                "Quy đổi (Năm)": int(thang_qd // 12),
                                "Quy đổi (Tháng)": round(thang_qd % 12, 1),
                                "Tổng (Tháng)": round(thang_qd, 1)
                            })
                            
                if person_data:
                    st.markdown(f"#### 👤 {row['HoTen']}")
                    st.caption(f"Số hiệu: {row['SoHieu']} | Đơn vị hiện tại: {row['DonVi']}")
                    
                    df_person = pd.DataFrame(person_data)
                    st.dataframe(df_person, use_container_width=True, hide_index=True)
                    
                    nam_thuc, thang_thuc_le = int(tong_thang_thuc // 12), int(tong_thang_thuc % 12)
                    nam_qd, thang_qd_le = int(tong_thang_quydoi // 12), round(tong_thang_quydoi % 12, 1)
                    st.success(f"**TỔNG CỘNG:** Thực tế làm việc: **{nam_thuc} năm {thang_thuc_le} tháng** ➔ Sau quy đổi: **{nam_qd} năm {thang_qd_le} tháng** (Tổng: {round(tong_thang_quydoi, 1)} tháng)")
                    
                    for p in person_data:
                        p_copy = p.copy()
                        p_copy["Số hiệu"], p_copy["Họ và Tên"] = row['SoHieu'], row['HoTen']
                        all_results_for_excel.append(p_copy)
                    st.divider() 
            
            with col12:
                if all_results_for_excel:
                    df_all_export = pd.DataFrame(all_results_for_excel)
                    cols = ["Số hiệu", "Họ và Tên", "Giai đoạn", "Chức danh (Gốc)", "Nhóm áp dụng", "Thực (Năm)", "Thực (Tháng)", "Hệ số", "Quy đổi (Năm)", "Quy đổi (Tháng)", "Tổng (Tháng)"]
                    df_all_export = df_all_export[cols]
                    
                    if loai_loc_t3 == "Xuất Tất cả Cán bộ / Nhân viên":
                        st.download_button("📥 TẢI EXCEL TẤT CẢ (TỔNG HỢP)", data=convert_df_to_excel(df_all_export, "QuyDoiHeSo"), file_name="TongHop_HeSoQuyDoi_TatCa.xlsx", type="primary", use_container_width=True)
                    else:
                        st.download_button("📥 TẢI EXCEL CÁ NHÂN", data=convert_df_to_excel(df_all_export, "QuyDoiHeSo"), file_name=f"QuyDoiHeSo_{df_target.iloc[0]['SoHieu']}.xlsx", type="primary", use_container_width=True)
