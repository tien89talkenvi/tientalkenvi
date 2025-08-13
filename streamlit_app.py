import streamlit as st  # streamlit=1.47.1
import pandas as pd     # pandas=2.3.1
import os, time

from selenium import webdriver  # selenium=4.34.2
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC

import shutil
from openpyxl import Workbook, load_workbook    # openpyxl=3.1.5
from openpyxl.styles import PatternFill
from io import BytesIO
import xlsxwriter   # xlsxwriter=3.2.5
import tempfile
from datetime import datetime

#from playwright.sync_api import sync_playwright    #playwright==1.54.0
#--------------------------------------------------------------

# Cac Ham Phu --------------------------------------------------
# Ham thu viec doc file txt/cvs dung encoding nao khong gay loi
def check_read_file_txt(filetxt):
    encodings_to_try = ['utf-8', 'utf-8-sig', 'cp1252', 'cp1258', 'utf-16']

    for enc in encodings_to_try:
        try:
            df = pd.read_csv('file.txt', delimiter='\t', encoding=enc)
            print(f"✅ Thành công với encoding: {enc}")
            break
        except Exception as e:
            print(f"❌ {enc}: {e}")

# Hàm kiểm tra giá trị có phải số hoặc ngày không
def is_number_or_date(val):
    if pd.isna(val):  # NaN thì giữ lại
        return False
    # Trường hợp là số thật
    if isinstance(val, (int, float)):
        return True
    # Nếu là chuỗi
    if isinstance(val, str):
        # Nếu chuỗi toàn số hoặc dạng số thập phân
        if val.strip().replace('.', '', 1).isdigit():
            return True
        # Thử parse sang datetime
        try:
            datetime.strptime(val.strip(), '%Y-%m-%d')
            return True
        except ValueError:
            pass
        try:
            datetime.strptime(val.strip(), '%d/%m/%Y')
            return True
        except ValueError:
            pass
        try:
            datetime.strptime(val.strip(), '%m-%d-%y')
            return True
        except ValueError:
            pass
    return False

# Cac Ham Cho Phan III --------------------------------------------------

def Ht_Data_tquat(outputIo):
    st.write('Ht_Data_tquat')
    wb = load_workbook(outputIo)
    ws = wb.active
    # xu li wb, ws o day
    # roi save ra file excel de xem
    # Ghi vào memory (không ghi ra ổ đĩa)
    virtual_workbook = BytesIO()
    wb.save(virtual_workbook)
    
    st.download_button(
        label="Tải file Excel về xem",
        data=virtual_workbook.getvalue(),
        file_name="Data_0.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    return 

def Ht_Data_sxep(outputIo):
    st.write('Ht_Data_sxep')
    wb = load_workbook(outputIo)
    ws = wb.active
    # xu li wb, ws o day
    # roi save ra file excel de xem
    virtual_workbook = BytesIO()
    wb.save(virtual_workbook)
    
    st.download_button(
        label="Tải file Excel về xem",
        data=virtual_workbook.getvalue(),
        file_name="Data_0.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    return 

def Ht_Data_new(outputIo):
    st.write('Ht_Data_new')
    wb = load_workbook(outputIo)
    ws = wb.active
    # xu li wb, ws o day
    # roi save ra file excel de xem
    virtual_workbook = BytesIO()
    wb.save(virtual_workbook)
    
    st.download_button(
        label="Tải file Excel về xem",
        data=virtual_workbook.getvalue(),
        file_name="Data_0.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    return 

def Ht_Data_old(outputIo):
    st.write('Ht_Data_old')
    wb = load_workbook(outputIo)
    ws = wb.active
    # xu li wb, ws o day
    # roi save ra file excel de xem
    virtual_workbook = BytesIO()
    wb.save(virtual_workbook)
    
    st.download_button(
        label="Tải file Excel về xem",
        data=virtual_workbook.getvalue(),
        file_name="Data_0.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    return 

def Ht_Data_max(outputIo):
    st.write('Ht_Data_max')
    wb = load_workbook(outputIo)
    ws = wb.active
    # xu li wb, ws o day
    # roi save ra file excel de xem
    virtual_workbook = BytesIO()
    wb.save(virtual_workbook)
    
    st.download_button(
        label="Tải file Excel về xem",
        data=virtual_workbook.getvalue(),
        file_name="Data_0.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    return 




# Ham tai file txt du lieu dang cvs cua cac mien thuoc bang Cali
@st.cache_data
def download_data_smarts(regions):
    #xoa thu muc downloads va tao lai de chi chua 2 file du lieu
    folder_path_cu = 'downloads'
    # Xóa thư mục nếu tồn tại
    if os.path.exists(folder_path_cu):
        shutil.rmtree(folder_path_cu)  # Xóa toàn bộ thư mục và nội dung bên trong

    download_dir = os.path.abspath("downloads")
    os.makedirs(download_dir, exist_ok=True)

    # ✅ CẤU HÌNH CHROME:
    options = webdriver.ChromeOptions()
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "profile.default_content_setting_values.automatic_downloads": 1   # DÒNG QUAN TRỌNG DE TAT THONG BAO
    }
    options.add_experimental_option("prefs", prefs)
    options.add_argument("--headless")  # chạy ẩn trình duyệt

    # ✅ KHỞI TẠO TRÌNH DUYỆT
    driver = webdriver.Chrome(options=options)

    driver.get("https://smarts.waterboards.ca.gov/smarts/SwPublicUserMenu.xhtml")
    print("✅ Vào trang chính")

    WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.LINK_TEXT, "Download NOI Data By Regional Board"))
    ).click()

    WebDriverWait(driver, 10).until(lambda d: len(d.window_handles) > 1)
    driver.switch_to.window(driver.window_handles[-1])
    print("✅ Đã chuyển sang tab mới")

    links = [
        "Industrial Application Specific Data",
        "Industrial Ad Hoc Reports - Parameter Data"
    ]

    def wait_for_download_and_get_new_file(before_files, timeout=40):
        for _ in range(timeout * 2):
            time.sleep(0.5)
            after_files = set(os.listdir(download_dir))
            new_files = after_files - before_files
            txt_files = [f for f in new_files if f.endswith(".txt")]
            if txt_files:
                return txt_files[0]
        return None
    #---------------------
    region = regions
    print(f"\n🔹 Chọn Region: {region}")
    dropdown = WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.NAME, "intDataFileDowloaddataFileForm:intDataDumpSelectOne"))
    )
    Select(dropdown).select_by_visible_text(region)
    time.sleep(3)  # Đợi dropdown load lại
    
    lfile_datai = []

    for j, name in enumerate(links):
        try:
            print(f"📥 Đang click tải: {name}")
            before = set(os.listdir(download_dir))

            link_elem = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.LINK_TEXT, name))
            )
            driver.execute_script("arguments[0].click();", link_elem)

            fname = wait_for_download_and_get_new_file(before)
            if fname:
                # Tạo tên file chuẩn theo Region + tên file
                src = os.path.join(download_dir, fname)
                dst_name = f"{region} - {name}.txt"
                dst_name = dst_name.replace(" ", "_")  # Nếu muốn
                dst = os.path.join(download_dir, dst_name)
                os.rename(src, dst)
                print(f"File đã lưu: {dst}")
                lfile_datai.append(f"{dst}")
            else:
                print("❌ Không tìm thấy file mới sau khi tải")
        except Exception as e:
            print(f"❌ Lỗi khi tải {name} ở Region {region}: {e}")

    driver.quit()
    print("\n🎉 Hoàn tất tải file cho "+region)
    return lfile_datai
    # CHU Y rang neu ten file dat trung voi file da co thi that bai.

# CAC HAM CHINH-----------------------------------------------------
def ThucThiPhan_4():
    return    


@st.cache_data
def Doc_hthi_data(uploaded_file):
    try:
        # Đọc file Excel thành DataFrame
        df = pd.read_excel(uploaded_file, sheet_name='Data')

        # 1. Sắp xếp dữ liệu theo nhiều cấp độ (multi-level sort):
        df_sorted = df.sort_values(
                by=["OLD/NEW", "PARAMETER", "RESULT"],
                ascending=[True, True, False]
        )
        # 2. Lọc dữ liệu có OLD/NEW == 'New':
        df_new = df_sorted[df_sorted["OLD/NEW"] == "New"]
        # 3. Tô màu (highlight) exceedances thì không thể hiển thị trong DataFrame thông thường nhưng có thể dùng:
        # pandas.ExcelWriter + openpyxl để ghi file Excel có màu.
        # Hoặc đơn giản chỉ đánh dấu bằng cột mới "Exceed" = True/False
        # 4. So sánh kết quả với ngưỡng NAL/NEL/TNAL:
        # tao dic chua nguong
        nal_thresholds = {
            "Ammonia": 4.7,
            "Cadmium": 0.0031,
            "Copper": 0.06749,
            # v.v...
        }
        # Rồi kiểm tra:
        def is_exceed(row):
            param = row["PARAMETER"]
            result = row["RESULT"]
            return result > nal_thresholds.get(param, float('inf'))
        
        df_new["EXCEED"] = df_new.apply(is_exceed, axis=1)
        # 5. Ghi chú các facility cần theo dõi → bạn có thể lọc hoặc thêm cột "Flagged" dựa vào danh sách thủ công.


        st.success(f"Đã tải lên: {uploaded_file.name}")
        st.subheader("📄 Dữ liệu từ file:")

        # Bước 3: Hiển thị DataFrame với cuộn dọc (giả lập 3 dòng)
        st.dataframe(df) #, height=120, use_container_width=True)

    except Exception as e:
        st.error(f"Lỗi khi đọc file: {e}")
    
def Ht_CaiMuonXem_0(tepxlsx):
    # 1.Đọc file Excel thành DataFrame
    df = pd.read_excel(tepxlsx, sheet_name='Data')
    # 2. Sắp xếp dữ liệu theo nhiều cấp độ (multi-level sort):
    df_sorted = df.sort_values(
        by=["OLD/NEW", "PARAMETER", "RESULT"],
        ascending=[True, True, False]
    )
    #Dua vao xlsx de xem kq sx
    
    tepxlsxdexem = "Data_Tracker_3-1.xlsx"
    df_sorted.to_excel(tepxlsxdexem, sheet_name='Datanew', index=False)
    os.startfile(tepxlsxdexem)

#---------------------------------
def Ht_CaiMuonXem_1(uploaded_file):
    uploaded_file = st.file_uploader("Tải lên Data_Tracker_New.xlsx", type=["xlsx"])
    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        # Ghi tạm ra file Excel để xử lý openpyxl
        output = BytesIO()
        df.to_excel(output, index=False)
        output.seek(0)

        # Load và tô màu
        wb = load_workbook(output)
        ws = wb.active

        # Tìm vị trí các cột "OLD/NEW" và "COSO"
        header = [cell.value for cell in ws[1]]

        try:
            old_new_col_idx = header.index("OLD/NEW") + 1
            coso_col_idx = header.index("FACILITY_NAME") + 1
        except ValueError as e:
            raise Exception(f"Không tìm thấy cột: {e}")

        # Tô màu vàng cho dòng 'new' thuộc cơ sở 'CS1'
        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

        # Duyệt từng dòng
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            old_new_val = str(row[old_new_col_idx - 1].value).strip().lower() if row[old_new_col_idx - 1].value else ""
            coso_val = str(row[coso_col_idx - 1].value).strip() if row[coso_col_idx - 1].value else ""

            if old_new_val == "old" and coso_val == 'CS1':
                for cell in row:
                    cell.fill = yellow_fill
                # Dòng này được giữ lại
            else:
                # Ẩn dòng không khớp điều kiện
                ws.row_dimensions[row[0].row].hidden = True

        # Lưu file mới
        # Ghi vào memory (không ghi ra ổ đĩa)
        virtual_workbook = BytesIO()
        wb.save(virtual_workbook)

        st.download_button(
            label="Tải file Excel",
            data=virtual_workbook.getvalue(),
            file_name="data_tracker_1.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        #wb.save("filtered_CS1_new.xlsx")
        #st.write('filtered_CS1_new.xlsx da co')        
        #os.startfile("filtered_CS1_new.xlsx")

        # Lưu lại và cho phép tải xuống
        #final_output = BytesIO()
        #wb.save(final_output)
        #final_output.seek(0)

        #st.download_button("📥 Tải file đã tô màu", final_output, "T_to_mau.xlsx")


# ---- phuc vu cho ThucThiPhan_2()------------------
@st.cache_data
def Combined_to_data_tracker(uploaded_tracker):
    try:
        # Read input files
        sheet1 = pd.read_excel(uploaded_tracker, sheet_name='Sheet1', header=None, dtype=str)
        sheet2 = pd.read_excel(uploaded_tracker, sheet_name='Sheet2', dtype=str)
        df_tracker = pd.read_excel(uploaded_tracker, sheet_name='Data', dtype=str)

        # Clean Sheet1
        cols_to_delete = list(range(15, 28)) + list(range(31, 38)) + [0, 4, 5, 6, 7, 11]
        # Chỉ lấy các chỉ số cột cần xóa, nhưng phải nhỏ hơn num_cols
        num_cols = sheet1.shape[1]
        cols_to_delete = [col for col in cols_to_delete if col < num_cols]
        sheet1_cleaned = sheet1.drop(sheet1.columns[cols_to_delete], axis=1)

        sheet1_cleaned.columns = sheet1_cleaned.iloc[0]
        sheet1_cleaned = sheet1_cleaned[1:]

        # Reorder columns
        cols = list(sheet1_cleaned.columns)
        if 'WDID' in cols and 'APP_ID' in cols:
            cols.remove('WDID')
            cols.insert(cols.index('APP_ID'), 'WDID')
        if 'FACILITY_NAME' in cols and 'OPERATOR_NAME' in cols:
            cols.remove('FACILITY_NAME')
            cols.insert(cols.index('OPERATOR_NAME'), 'FACILITY_NAME')
        sheet1_cleaned = sheet1_cleaned[cols]


        # Clean Sheet2 and merge SIC data
        sheet2_cleaned = sheet2.drop(sheet2.columns[2:8], axis=1)
        sheet2_cleaned = sheet2_cleaned.iloc[:, :5]  # Chỉ lấy 5 cột đầu tiên
        sheet2_cleaned.columns = ['A', 'APP_ID', 'PRIMARY_SIC', 'SECONDARY_SIC', 'TERTIARY_SIC']
        sheet2_cleaned = sheet2_cleaned[['APP_ID', 'PRIMARY_SIC', 'SECONDARY_SIC', 'TERTIARY_SIC']]
        merged = pd.merge(sheet1_cleaned, sheet2_cleaned, on='APP_ID', how='left')
        merged = merged.rename(columns={'PRIMARY_SIC': '1', 'SECONDARY_SIC': '2', 'TERTIARY_SIC': '3'})
        merged['2'] = merged['2'].replace('0', pd.NA)
        merged['3'] = merged['3'].replace('0', pd.NA)
        merged['OLD/NEW'] = 'new'

        # Add "Old" tag to tracker and combine
        df_tracker['OLD/NEW'] = 'old'
        df_combined = pd.concat([df_tracker, merged], ignore_index=True)

        # Export to Excel in-memory
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_combined.to_excel(writer, sheet_name='Data', index=False)
        output.seek(0)
        return output 
    except Exception as e:
        st.error(f"⚠️ An error occurred: {e}")
#----------------

@st.cache_data
def Txt_to_data_tracker(df1, df2, df_data):
    # tao file excel ao xlsx_ao_chua_3df chua df1, df2, df_data
    xlsx_ao_chua_3df = BytesIO()
    with pd.ExcelWriter(xlsx_ao_chua_3df, engine="openpyxl") as writer:
        df1.to_excel(writer, sheet_name="Sheet1", index=False)
        df2.to_excel(writer, sheet_name="Sheet2", index=False)
        df_data.to_excel(writer, sheet_name="Data", index=False)

    xlsx_ao_chua_3df.seek(0)
    return xlsx_ao_chua_3df
    # tao nut download file tren neu can
    #st.download_button("📥 Download Data_Tracker_include_3df.xlsx",
    #    data=xlsx_ao_chua_3df,
    #    file_name="Data_Tracker_include_3df.xlsx",
    #    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    #)

# Cac ham Phan II -----------------------------    
# ham nay de In Sheet2, Tìm các dòng APP_ID trùng, 
# rồi xóa các dòng mà STATUS ≠ Active nhưng giữ lại các dòng có STATUS = Active "
def Xli_P2_1(F_excel_data_ao):
    F_excel_data_ao.seek(0)  # quay lại đầu BytesIO để đọc
    dfSheet2 = pd.read_excel(F_excel_data_ao, sheet_name='Sheet2')
    # Tìm APP_ID bị trùng
    duplicates = dfSheet2[dfSheet2.duplicated(subset='APP_ID', keep=False)]
    # Giữ lại dòng trùng có STATUS khác "Active"
    to_delete = duplicates[duplicates['STATUS'] != 'Active']
    # Xoá các dòng này khỏi dataframe gốc (chu y la file excel van y cu)
    dfSheet2_cleaned = dfSheet2.drop(to_delete.index)
    # ket qua la cac dong trung APP_ID nhung co STATUS la Active duoc giu lai, con
    # cac dong trung APP_ID nhung co STATUS khac Active thi bi xoa
    # cho hien thi df con lai sau khi da xoa cac to_delete
    # phai hieu df_cleaned la df con lai sau khi da lam sach 
    #st.write(df_cleaned)
    #---
    if st.checkbox("View Sheet2_1", key='BCB1'):
        st.write('(rows, cols) = ', len(dfSheet2_cleaned), len(dfSheet2_cleaned.columns))
        st.write(dfSheet2_cleaned)

    # Ghi thêm dfnew vào Sheet2 mà không mất Sheet1
    F_excel_data_ao.seek(0)  # quan trọng: để writer đọc được file hiện tại
    with pd.ExcelWriter(F_excel_data_ao, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        dfSheet2_cleaned.to_excel(writer, sheet_name="Sheet2", index=False)
    # Giờ output có cả Sheet1 và Sheet2
    F_excel_data_ao.seek(0)
    return F_excel_data_ao

# ham nay de Delete, move, re-order columns in Sheet2
def Xli_P2_2(F_excel_data_ao):
    # set dfSheet2 is dataframe of  Sheet2 in uploaded_f3
    dfSheet2 = pd.read_excel(F_excel_data_ao, sheet_name='Sheet2')
    # xoa cac cot AF den AL: 
    # loai bo cac cot trong [] va tra ve kq dfSheet2_cleaned la data con lai
    dfSheet2_cleaned = dfSheet2.drop(['RECEIVING_WATER_NAME','INDIRECTLY',
        'DIRECTLY','CERTIFIER_BY','CERTIFIER_TITLE','CERTIFICATION_DATE',
        'QUESTION_TMDL_ANSWER'],axis=1)
    # tiep tuc loai bo cac cot P den AB:
    dfSheet2_cleaned = dfSheet2_cleaned.drop(['FACILITY_LATITUDE','FACILITY_LONGITUDE',
    	'FACILITY_COUNTY','FACILITY_CONTACT_FIRST_NAME','FACILITY_CONTACT_LAST_NAME',
        'FACILITY_TITLE','FACILITY_PHONE','FACILITY_EMAIL',	'FACILITY_TOTAL_SIZE',
        'FACILITY_TOTAL_SIZE_UNIT',	'FACILITY_AREA_ACTIVITY','FACILITY_AREA_ACTIVITY_UNIT',
        'PERCENT_OF_SITE_IMPERVIOUSNESS'],axis=1)
    # tiep tuc loai bo cac cot A,E-H,L:
    dfSheet2_cleaned = dfSheet2_cleaned.drop(['PERMIT_TYPE','NOI_PROCESSED_DATE',
        'NOT_EFFECTIVE_DATE','REGION_BOARD','COUNTY', 'FACILITY_ADDRESS_2'],axis=1)
    # di chuyen cot WDID sang ben trai cot APP_ID
    cols = list(dfSheet2_cleaned.columns)
    if 'WDID' in cols and 'APP_ID' in cols:
        cols.remove('WDID')
        cols.insert(cols.index('APP_ID'), 'WDID')
    # tiep tuc di chuyen cot FACILITY_NAME sang ben trai cot OPERATOR_NAME
    if 'FACILITY_NAME' in cols and 'OPERATOR_NAME' in cols:
        cols.remove('FACILITY_NAME')
        cols.insert(cols.index('OPERATOR_NAME'), 'FACILITY_NAME')
    dfSheet2_cleaned = dfSheet2_cleaned[cols]
    # cuối cùng phải còn lại là :
    # WDID	APP_ID	STATUS	FACILITY_NAME	OPERATOR_NAME	FACILITY_ADDRESS	FACILITY_CITY	FACILITY_STATE	FACILITY_ZIP	PRIMARY_SIC	SECONDARY_SIC	TERTIARY_SIC

    if st.checkbox("View Sheet2_2", key='BCB2'):
        st.write('(rows, cols) = ', len(dfSheet2_cleaned), len(dfSheet2_cleaned.columns))
        st.write(dfSheet2_cleaned)

    # Ghi cap nhat Sheet2
    F_excel_data_ao.seek(0)  # quan trọng: để writer đọc được file hiện tại
    with pd.ExcelWriter(F_excel_data_ao, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        dfSheet2_cleaned.to_excel(writer, sheet_name="Sheet2", index=False)
    F_excel_data_ao.seek(0)
    return F_excel_data_ao

# ham nay de In Sheet1, delete all rows duplicated and rows showing '4 56' in WDID
def Xli_P2_3(F_excel_data_ao):
    dfSheet1 = pd.read_excel(F_excel_data_ao, sheet_name='Sheet1')
    # drop all duplicated row
    dfSheet1_cleaned = dfSheet1.drop_duplicates()
    # xóa các dòng mà cột 'WDID' chứa chính xác chuỗi "4 56", còn lại các dòng ['WDID'] != '4 56'
    dfSheet1_cleaned = dfSheet1_cleaned[dfSheet1_cleaned['WDID'] != '4 56']
    # Delete all columns not in your tracker (columns A, J, K, U, X, and Y)
    dfSheet1_cleaned = dfSheet1_cleaned.drop(['PERMIT_TYPE', 'MONITORING_LATITUDE', 
        'MONITORING_LONGITUDE', 'ANALYTICAL_METHOD', 'DISCHARGE_END_DATE',	
        'DISCHARGE_END_TIME'],axis=1)
    #After deleting these columns, Sheet1 columns should look like the following: 
    # A – WDID; B – App ID; C – Status; D – Facility Name; E – Operator Name; F –Address; G – City; H – State; I – Zip; J – Primary SIC; K – Secondary SIC; L – Tertiary SIC
    # Sheet1 khong co STATUS, vay phai nhu ben phai day:  WDID	APP_ID	REPORTING_YEAR	REPORT_ID	EVENT_TYPE	MONITORING_LOCATION_NAME	MONITORING_LOCATION_TYPE	MONITOR_LOCATION_DESCRIPTION	SAMPLE_ID	SAMPLE_DATE	SAMPLE_TIME	DISCHARGE_START_DATE	DISCHARGE_START_TIME	PARAMETER	RESULT_QUALIFIER	RESULT	UNITS	MDL	RL	CERTIFIER_NAME	CERTIFIED_DATE
    if st.checkbox("View Sheet1_3", key='BCB3'):
        st.write('(rows, cols) = ', len(dfSheet1_cleaned), len(dfSheet1_cleaned.columns))
        st.write(dfSheet1_cleaned)

    # Ghi cap nhat Sheet2
    F_excel_data_ao.seek(0)  # quan trọng: để writer đọc được file hiện tại
    with pd.ExcelWriter(F_excel_data_ao, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        dfSheet1_cleaned.to_excel(writer, sheet_name="Sheet1", index=False)
    F_excel_data_ao.seek(0)
    return F_excel_data_ao

# Filter Sheet1 for only new sample data
def Xli_P2_4(F_excel_data_ao):
    dfSheet1 = pd.read_excel(F_excel_data_ao, sheet_name="Sheet1")  # chứa cột APP_ID
    dfData = pd.read_excel(F_excel_data_ao, sheet_name="Data")  # chứa cột O và P

    # Giống =VLOOKUP(J2, Data!O:P, 2, FALSE)
    lookup_dict = pd.Series(dfData['PP'].values, index=dfData['OO']).to_dict()
    # Tìm vị trí (chỉ số) của cột 'APP_ID' trong Sheet1
    idx = dfSheet1.columns.get_loc('APP_ID')
    # Chèn cột mới 'VLOOKUP' vào trước 'APP_ID'
    dfSheet1.insert(loc=idx, column='VLOOKUP', value=dfSheet1['SAMPLE_DATE'].map(lookup_dict))
    # cột 'SAMPLE_DATE' là cột 'J' trong công thức =VLOOKUP(J2, Data!O:P, 2, FALSE

    # Xóa (lọc bỏ) tất cả các hàng có số trong cột 'VLOOKUP' 
    dfSheet1 = dfSheet1[~dfSheet1['VLOOKUP'].apply(lambda x: isinstance(x, (int, float)))]
    #dfSheet1 = dfSheet1[~dfSheet1['VLOOKUP'].apply(is_number_or_date)]
    # Giữ lại các dòng mà cột 'VLOOKUP' không chứa số trong Sêt1
    #dfSheet1 = dfSheet1[~dfSheet1['VLOOKUP'].astype(str).str.contains(r'\d', na=False)]
    #dfSheet1 = dfSheet1[~dfSheet1['VLOOKUP'].astype(str).str.contains(r'\d', na=False)]
    # sap xep theo 'VLOOKUP' tăng dần
    #dfSheet1 = dfSheet1.sort_values(by='VLOOKUP', ascending=False)
    dfSheet1 = dfSheet1.sort_values(by='VLOOKUP', ascending=True, na_position='first')

    if st.checkbox("View Sheet1_4", key='BCB4'):
        st.write('(rows, cols) = ', len(dfSheet1), len(dfSheet1.columns))
        st.write(dfSheet1)

    # Ghi cap nhat Sheet1
    F_excel_data_ao.seek(0)  # quan trọng: để writer đọc được file hiện tại
    with pd.ExcelWriter(F_excel_data_ao, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        dfSheet1.to_excel(writer, sheet_name="Sheet1", index=False)
    F_excel_data_ao.seek(0)
    return F_excel_data_ao

# DN Ham Check if facilities in Sheet1 are active
def Xli_P2_5(F_excel_data_ao):
    dfSheet1 = pd.read_excel(F_excel_data_ao, sheet_name="Sheet1")  # chứa cột APP_ID
    dfSheet2 = pd.read_excel(F_excel_data_ao, sheet_name="Sheet2")  # chứa cột O và P
    # Giống =VLOOKUP(C2,Sheet2!B:D,2,FALSE),
    lookup_dict = pd.Series(dfSheet2['STATUS'].values, index=dfSheet2['APP_ID']).to_dict()
    dfSheet1['VLOOKUP'] = dfSheet1['APP_ID'].map(lookup_dict)
    dfSheet1 = dfSheet1.sort_values(by='VLOOKUP', ascending=False)
    dfSheet1 = dfSheet1[dfSheet1['VLOOKUP'] == 'Active']

    if st.checkbox("View Sheet1_5", key='BCB5'):
        st.write('(rows, cols) = ', len(dfSheet1), len(dfSheet1.columns))
        st.write(dfSheet1)

    # Xóa (lọc bỏ) tất cả các hàng có số trong cột 'VLOOKUP' 
    #dfSheet1 = dfSheet1[~dfSheet1['VLOOKUP'].apply(lambda x: isinstance(x, (int, float)))]

    # Ghi cap nhat Sheet1
    F_excel_data_ao.seek(0)  # quan trọng: để writer đọc được file hiện tại
    with pd.ExcelWriter(F_excel_data_ao, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        dfSheet1.to_excel(writer, sheet_name="Sheet1", index=False)
    F_excel_data_ao.seek(0)
    return F_excel_data_ao

# DN Ham Choose the parameters to track in Sheet1
def Xli_P2_6(F_excel_data_ao):
    dfSheet1 = pd.read_excel(F_excel_data_ao, sheet_name="Sheet1")  # chứa cột APP_ID
    #dfSheet2 = pd.read_excel(F_excel_data_ao, sheet_name="Sheet2")  # chứa cột O và P

    # NEW tao menu de chon cac gia tri cua cot PARAMETER de lam viec 
    lpara = dfSheet1["PARAMETER"].dropna().unique().tolist()
    lgiulai = []
    popover = st.popover("Select parameter:")
    for pt in lpara:
        checked = popover.checkbox(pt, True)
        if checked:
            lgiulai.append(pt)
    # tai day list lgiulai da luu cac para chon

    # xem thu cac para da chon
    st.write(lgiulai) 

    # Xóa cột B
    dfSheet1 = dfSheet1.drop(dfSheet1.columns[1], axis=1)

    # Lọc giữ các dong co parameter mong muốn
    dfSheet1 = dfSheet1[dfSheet1["PARAMETER"].isin(lgiulai)]

    # Tách PARAMETER thành 2 cột PARAMETER và QUALIFIER
    split_cols = dfSheet1["PARAMETER"].str.split(",", n=1, expand=True)
    dfSheet1["PARAMETER"] = split_cols[0]
    param_index = dfSheet1.columns.get_loc("PARAMETER")
    dfSheet1.insert(param_index + 1, "QUALIFIER", split_cols[1])

    # xem Sheet1 da cap nhat
    if st.checkbox("View Sheet1_6", key='BCB6'):
        st.write('(rows, cols) = ', len(dfSheet1), len(dfSheet1.columns))
        st.write(dfSheet1)

    # Lưu lại Ghi cap nhat Sheet1
    F_excel_data_ao.seek(0)  # quan trọng: để writer đọc được file hiện tại
    with pd.ExcelWriter(F_excel_data_ao, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        dfSheet1.to_excel(writer, sheet_name="Sheet1", index=False)
    F_excel_data_ao.seek(0)
    return F_excel_data_ao

# DN Ham Make sure all the samples in Sheet1 are in 'mg/L' and not 'ug/L'
def Xli_P2_7(F_excel_data_ao):
    dfSheet1 = pd.read_excel(F_excel_data_ao, sheet_name="Sheet1")  
    # Chuyển đổi từ ug/L sang mg/L
    mask_ug = dfSheet1["UNITS"] == "ug/L"
    dfSheet1.loc[mask_ug, "RESULT"] = dfSheet1.loc[mask_ug, "RESULT"] / 1000
    dfSheet1.loc[mask_ug, "UNITS"] = "mg/L"

    if st.checkbox("View Sheet1_7", key='BCB7'):
        st.write('(rows, cols) = ', len(dfSheet1), len(dfSheet1.columns))
        st.write(dfSheet1)

    # Lưu lại Ghi cap nhat Sheet1
    F_excel_data_ao.seek(0)  # quan trọng: để writer đọc được file hiện tại
    with pd.ExcelWriter(F_excel_data_ao, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        dfSheet1.to_excel(writer, sheet_name="Sheet1", index=False)
    F_excel_data_ao.seek(0)
    return F_excel_data_ao

def Xli_P2_8(F_excel_data_ao):
    dfSheet1 = pd.read_excel(F_excel_data_ao, sheet_name="Sheet1")  
    dfSheet2 = pd.read_excel(F_excel_data_ao, sheet_name="Sheet2")
    
    # Xoa cột C trong Sheet2
    dfSheet2 = dfSheet2.drop(dfSheet2.columns[2], axis=1)

    # Tạo bảng tra cứu giống vùng Sheet2 từ APP_ID
    lookup_cols = ["FACILITY_NAME", "OPERATOR_NAME", "FACILITY_ADDRESS", 
                "FACILITY_CITY", "FACILITY_STATE", "FACILITY_ZIP"]

    lookup_df = (
        dfSheet2
        .drop_duplicates(subset=["APP_ID"], keep="first")  # Giống VLOOKUP lấy bản ghi đầu tiên
        .set_index("APP_ID")[lookup_cols]
    )

    # Thêm 6 cột vào trước cột "Reporting Year" và map dữ liệu
        # Xác định vị trí cột Reporting Year
    pos = dfSheet1.columns.get_loc("REPORTING_YEAR")
        # Map từng cột lookup vào dfSheet1
    for i, col in enumerate(lookup_cols, start=1):
        dfSheet1.insert(pos + i - 1, col, dfSheet1["APP_ID"].map(lookup_df[col]))
        # Dòng tren này tương đương việc viết công thức VLOOKUP() và kéo sang 6 cột trong Excel.

    # Lưu lại Ghi cap nhat Sheet1
    F_excel_data_ao.seek(0)  # quan trọng: để writer đọc được file hiện tại
    with pd.ExcelWriter(F_excel_data_ao, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        dfSheet1.to_excel(writer, sheet_name="Sheet1", index=False)
        dfSheet2.to_excel(writer, sheet_name="Sheet2", index=False)
    F_excel_data_ao.seek(0)
    return F_excel_data_ao  
    
    tambo='''
    # Đặt lại tên cột theo thứ tự mong muốn
    dfSheet2.columns = [
        "WDID", "APP_ID", "FACILITY_NAME", "OPERATOR_NAME",
        "FACILITY_ADDRESS", "FACILITY_CITY", "FACILITY_STATE", "FACILITY_ZIP",
        "PRIMARY_STC", "SECONDARY_SIC", "TERTIARY_SIC"
    ]
	

    # 2. Thêm 6 cột trống vào Sheet1 ở trước cột C (tức index 2 trong pandas)
    for i in range(6):
        dfSheet1.insert(2 + i, f"NewCol{i+1}", "")
    
    # 3. Tạo tra cứu tương đương VLOOKUP từ Sheet2
    # Sử dụng APP_ID làm key để map dữ liệu sang 6 cột mới
    lookup_cols = ["FACILITY_NAME", "OPERATOR_NAME", "FACILITY_ADDRESS", "FACILITY_CITY", "FACILITY_STATE", "FACILITY_ZIP"]
    lookup_df = dfSheet2.set_index("APP_ID")[lookup_cols]

    for i, col in enumerate(lookup_cols):
        dfSheet1[f"NewCol{i+1}"] = dfSheet1["APP_ID"].map(lookup_df[col])
    # doan tambo tren gay loi nen lay doan sau:
    lookup_cols = ["FACILITY_NAME", "OPERATOR_NAME", "FACILITY_ADDRESS", 
                "FACILITY_CITY", "FACILITY_STATE", "FACILITY_ZIP"]

    lookup_df = (
        dfSheet2
        .drop_duplicates(subset=["APP_ID"], keep="first")  # giữ bản ghi đầu tiên cho mỗi APP_ID
        .set_index("APP_ID")[lookup_cols]                  # APP_ID làm index, chỉ giữ cột cần thiết
    )
    st.write('(rows, cols) = ', len(lookup_df), len(lookup_df.columns))
    st.write(dfSheet1)
    return F_excel_data_ao
    
    #------------

    # 4. Ghi kết quả ra file mới
    '''
    
def Xli_P2_9(F_excel_data_ao):
    dfSheet1 = pd.read_excel(F_excel_data_ao, sheet_name="Sheet1")  
    dfSheet2 = pd.read_excel(F_excel_data_ao, sheet_name="Sheet2")
    #trong Sheet2 Xóa cột C → H, gexcePP_ID ngay ben trai PRIMARY_SIC
    # Xóa theo vị trí (C-H = cột 2 → 7 vì Python đếm từ 0)
    dfSheet2 = dfSheet2.drop(dfSheet2.columns[2:8], axis=1)
    # Đảm bảo cột sau khi xóa: WDID, APP_ID, PRIMARY_SIC, SECONDARY_SIC, TERTIARY_SIC

    # Trong Sheet1 – Thêm 3 cột mới ngay sau cột cuối cùng
        # Tạo bảng tra cứu từ Sheet2
    lookup_cols = ["PRIMARY_SIC", "SECONDARY_SIC", "TERTIARY_SIC"]

    lookup_df = (
        dfSheet2
        .drop_duplicates(subset=["APP_ID"], keep="first")  # giống VLOOKUP lấy bản ghi đầu tiên
        .set_index("APP_ID")[lookup_cols]
    )
        # Thêm 3 cột vào Sheet1 bằng map (tương đương viết công thức & kéo sang)
    for col in lookup_cols:
        dfSheet1[col] = dfSheet1["APP_ID"].map(lookup_df[col])

    # Xóa giá trị 0 trong TERTIARY_SIC và SECONDARY_SIC
    #Trong Excel, bước lọc “0” rồi Clear Contents thực chất là xóa tất cả giá trị bằng 0 trong cột.
        # Xóa giá trị 0 ở TERTIARY_SIC
    dfSheet1.loc[dfSheet1["TERTIARY_SIC"] == 0, "TERTIARY_SIC"] = None
        # Xóa giá trị 0 ở SECONDARY_SIC
    dfSheet1.loc[dfSheet1["SECONDARY_SIC"] == 0, "SECONDARY_SIC"] = None

    if st.checkbox("View Sheet1_9", key='BCB9_S1'):
        st.write('(rows, cols) = ', len(dfSheet1), len(dfSheet1.columns))
        st.write(dfSheet1)
    if st.checkbox("View Sheet2_9", key='BCB9_S2'):
        st.write('(rows, cols) = ', len(dfSheet2), len(dfSheet2.columns))
        st.write(dfSheet2)

    # Lưu lại Ghi cap nhat vao excel
    F_excel_data_ao.seek(0)  # quan trọng: để writer đọc được file hiện tại
    with pd.ExcelWriter(F_excel_data_ao, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        dfSheet1.to_excel(writer, sheet_name="Sheet1", index=False)
        dfSheet2.to_excel(writer, sheet_name="Sheet2", index=False)
    F_excel_data_ao.seek(0)

    return F_excel_data_ao  

def Xli_P2_10(F_excel_data_ao):
    dfData = pd.read_excel(F_excel_data_ao, sheet_name="Data")  
    dfNew = pd.read_excel(F_excel_data_ao, sheet_name="Sheet1")
    # Điền "Old" vào cột OLD/NEW cho dữ liệu cũ trong sheet Data
        # Giả sử cột này tên là "OLD/NEW"
    dfData["OLD/NEW"] = "Old"

    # Gắn thêm dữ liệu mới từ Sheet1 vào Data
        # Bỏ hàng header trong Sheet1 (đã loại khi đọc file)
    df_combined = pd.concat([dfData, dfNew], ignore_index=True)

    # Thêm cột OLD/NEW = "New" cho dữ liệu mới
    # Xác định số hàng mới vừa thêm
    num_old = len(dfData)
    df_combined.loc[num_old:, "OLD/NEW"] = "New"

    # Lưu lại Ghi cap nhat vao excel
    F_excel_data_ao.seek(0)  # quan trọng: để writer đọc được file hiện tại
    with pd.ExcelWriter(F_excel_data_ao, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df_combined.to_excel(writer, sheet_name="Data", index=False)
    F_excel_data_ao.seek(0)

    # Xóa Sheet1 và Sheet2 khi xuất lại (chỉ giữ sheet "Data")
    #F_excel_data_ao.seek(0)  # quan trọng: để writer đọc được file hiện tại
    #with pd.ExcelWriter(F_excel_data_ao) as writer:
    #    df_combined.to_excel(writer, sheet_name="Data", index=False)
    #F_excel_data_ao.seek(0)

    return F_excel_data_ao  

#-------------------
def ThucThiPhan_1():
    # cac vung de chon
    regions = st.selectbox("Select a Region:", 
                ("Region 1 - North Coast",
                "Region 2 - San Francisco Bay",
                "Region 3 - Central Coast",
                "Region 4 - Los Angeles",
                "Region 5F - Fresno",
                "Region 5R - Redding",
                "Region 5S - Sacramento",
                "Region 6A - South Lake Tahoe",
                "Region 6B - Victorville",
                "Region 7 - Colorado River Basin",
                "Region 8 - Santa Ana",
                "Region 9 - San Diego"),
                index=None,
                placeholder="No selected Region",
                )
    #neu mot vung duoc chon thi lam
    LOI='OK'
    if regions:
        placeholder_1 = st.empty()
        placeholder_1.write('Wait for downloading 2 files of ' + regions)
        #thuc thi ham download_data_smarts(regions) va tra ve list cac file da tai 
        try :
            lfile_datai = download_data_smarts(regions)
            placeholder_1.write('Downloaded files:')
            st.write(lfile_datai)
        except:
            LOI='LOI'
            placeholder_1.write('Tai file không đạt!')
    if LOI == 'LOI':
        st.write('Nếu không đạt, '+ ':red[ mở trực tiếp trang sau làm theo các bước để tải:]')
        st.markdown("1-[Open Page SMARTS](https://smarts.waterboards.ca.gov/smarts/SwPublicUserMenu.xhtml)", unsafe_allow_html=True)
        st.write('2-Click on “Download NOI Data By Regional Board”')
        st.write('3-Select your region from the dropdown menu')
        st.write('4-Click on both “Industrial Application Specific Data” and “Industrial Ad Hoc Reports - Parameter Data”')
        st.write('5-Data will be downloaded to two separate .txt files, each titled “file”')
        st.write('6-Nên đổi tên 2 file thành Industrial_Application_Specific_Data và Industrial_Ad_Hoc_Reports_-_Parameter_Data rồi chép vào thư mục riêng của bạn để dễ làm việc ở các bước sau.')


#========================= MAIN =====================================================================
# global
DACO_EXCEL_3SHEET=False

# TIEU DE APP
st.header('🏷️Trình hỗ trợ quản lý môi trường nước')

# Phan sider ben trai ---------------------------------------------------------------------------
with st.sidebar:
    st.header('🔎 Documents used as a basis for writing this program')
    # Đọc nội dung file Markdown
    with open("hd-lam-app-cho-thong.md", "r", encoding="utf-8") as f:
        md_content = f.read()
    st.markdown(md_content, unsafe_allow_html=True)


# I. TAI FILES TXT DU LIEU VE TU -----------------------------------------------------------------
st.subheader('✅ I. Download the data', divider=True)
ThucThiPhan_1()

# II Them data moi vao trinh theo doi -------------------------------------------------------------
st.subheader('✅ II. Add the new data to your tracker', divider=True)

laydatafrom = st.radio(
    "GET DATA FROM WHERE?", 
    [":blue[From Local]",":green[From Server]", ":red[Empty]"],
    index=2,horizontal=True , label_visibility="visible") 

if laydatafrom==":red[Empty]":
    DACO_EXCEL_3SHEET=False
    pass  

elif laydatafrom==":blue[From Local]":
    # Add the new data to your tracker 
    # - Upload 3 files
    uploaded_files = st.file_uploader(
        'Upload 1 lần 3 files: "...Industrial_Ad_Hoc...", "...Industrial_Application...", "...Data_Tracker..." ' + ' :red[(nên đặt 3 files này liền nhau trong 1 thư mục)]',
        type=['txt', 'xlsx'],  
        accept_multiple_files=True
    )
    if uploaded_files and len(uploaded_files) == 3:
        # Phân loại file theo đuôi và tên
        uploaded_f1 = next((f for f in uploaded_files if "industrial_ad_hoc" in f.name.lower()), None)
        uploaded_f3 = next((f for f in uploaded_files if f.name.lower().endswith(".xlsx")), None)
        # f2 là file .txt còn lại (không phải f1)
        uploaded_f2 = next((f for f in uploaded_files if f != uploaded_f1 and f.name.lower().endswith(".txt")), None)

        if uploaded_f1 and uploaded_f2 and uploaded_f3:
            try:
                df1 = pd.read_csv(uploaded_f1, sep='\t', encoding='cp1252')
                df2 = pd.read_csv(uploaded_f2, sep='\t', encoding='cp1252')
                #dfData = pd.read_excel(uploaded_f3, sheet_name="Data")  # Chỉ đọc sheet "Data"
            except Exception as e:
                st.error(f"⚠️ Lỗi khi đọc file: {e}")
                st.stop()
            #---
            # Đọc file Excel đã upload
            excel_data = uploaded_f3.read()

            # Ghi DataFrame TXT vào file Excel đã upload
            F_excel_data_ao = BytesIO()
            with pd.ExcelWriter(F_excel_data_ao, engine="openpyxl") as writer:
                # Ghi lại các sheet cũ của file Excel gốc
                original_excel = pd.ExcelFile(BytesIO(excel_data))
                for sheet_name in original_excel.sheet_names:
                    df_old = pd.read_excel(original_excel, sheet_name=sheet_name)
                    df_old.to_excel(writer, sheet_name=sheet_name, index=False)
                # Thêm / Ghi đè sheet "Sheet1" bằng dữ liệu từ file TXT
                df1.to_excel(writer, sheet_name="Sheet1", index=False)
                df2.to_excel(writer, sheet_name="Sheet2", index=False)

            # 3. Tạo nút tải xuống
            st.download_button(
                label="📥 Tải file Excel (Data_tracker_add2sheet.xlsx)",
                data=F_excel_data_ao.getvalue(),
                file_name="Data_tracker_add2sheet.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            DACO_EXCEL_3SHEET=True
else:
    # Đường dẫn thư mục TAM (nằm ngang với streamlit_app.py)
    BASE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Datatest")

    # Tên file mong muốn
    file_names = ["Region_1_-_North_Coast_-_Industrial_Ad_Hoc_Reports_-_Parameter_Data.txt",
            "Region_1_-_North_Coast_-_Industrial_Application_Specific_Data.txt", 
            "Data_Tracker_X.xlsx"]

    # Danh sách đường dẫn file trên server
    server_files = [os.path.join(BASE_DIR, name) for name in file_names]

    # Kiểm tra xem tất cả file có sẵn trên server không
    if all(os.path.exists(path) for path in server_files):
        st.info("📂 Đang dùng file trong thư mục Datatest.")
        #with open(server_files[0], "r", encoding="utf-8") as f1:
        #    txt_content_f1 = f1.read()
        #with open(server_files[1], "r", encoding="utf-8") as f2:
        #    txt_content_f2 = f2.read()
        df1 = pd.read_csv(server_files[0], sep='\t', encoding='cp1252')
        df2 = pd.read_csv(server_files[1], sep='\t', encoding='cp1252')
        #dfData = pd.read_excel(server_files[2])

    else:
        st.warning("📤 File chưa có trong Datatest, vui lòng upload.")
        pass
    excel_data = pd.read_excel(server_files[2])

    # Lưu lại Ghi cap nhat vao excel
    F_excel_data_ao = BytesIO()
    with pd.ExcelWriter(F_excel_data_ao, engine="openpyxl") as writer:
        df1.to_excel(writer, sheet_name="Sheet1", index=False)
        df2.to_excel(writer, sheet_name="Sheet2", index=False)
        excel_data.to_excel(writer, sheet_name="Data", index=False)

    # Tạo nút tải xuống
    st.download_button(
        label="📥 Tải file Excel (Data_tracker_goc.xlsx)",
        data=F_excel_data_ao.getvalue(),
        file_name="Data_tracker_goc.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    DACO_EXCEL_3SHEET=True

if DACO_EXCEL_3SHEET==True :
    #--------------------------------------
    st.write(":red[➡️ Add the new data to your tracker]")

    #checkbox1 = st.checkbox("📌:blue[1. Xóa các dòng mà STATUS ≠ 'Active' trong các dòng có APP_ID trùng lặp in Sheet2]", key='CB1')
    checkbox1 = st.checkbox("📌:blue[1. In your data_tracker.xlsx, create Sheet1, Sheet2 contain 2 file.txt]", key='CB1')
    if checkbox1:
        # tra ve kq la file ao da update cung ten F_excel_data_ao 
        F_excel_data_ao = Xli_P2_1(F_excel_data_ao) 
        if F_excel_data_ao:
            st.write(':green[Xli_P2_1 finished.]')
            # Tạo nút tải xuống
            st.download_button(
                label="📥 Tải file Excel (Data_tracker_add2sheet_1.xlsx)",
                data=F_excel_data_ao.getvalue(),
                file_name="Data_tracker_add2sheet_1.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    #checkbox2 = st.checkbox("📌:blue[2. Delete, move, re-order columns in Sheet2]", key='CB2')
    checkbox2 = st.checkbox("📌:blue[2. Get Sheet2 into the proper format for your tracker]", key='CB2')
    if checkbox1 and checkbox2:
        F_excel_data_ao = Xli_P2_2(F_excel_data_ao)
        if F_excel_data_ao:
            st.write(':green[Xli_P2_2 finished.]')
            # Tạo nút tải xuống
            st.download_button(
                label="📥 Tải file Excel (Data_tracker_add2sheet_2.xlsx)",
                data=F_excel_data_ao.getvalue(),
                file_name="Data_tracker_add2sheet_2.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    #checkbox3 = st.checkbox("📌:blue[3. Delete all rows duplicated and rows showing '4 56' in WDID in Sheet1]", key='CB3')
    checkbox3 = st.checkbox("📌:blue[3. Get Sheet1 into the proper format for your tracker]", key='CB3')
    if checkbox1 and checkbox2 and checkbox3:
        F_excel_data_ao = Xli_P2_3(F_excel_data_ao)
        if F_excel_data_ao:
            st.write(':green[Xli_P2_3 finished.]')
            # Tạo nút tải xuống
            st.download_button(
                label="📥 Tải file Excel (Data_tracker_add2sheet_2.xlsx)",
                data=F_excel_data_ao.getvalue(),
                file_name="Data_tracker_add2sheet_3.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    checkbox4 = st.checkbox("📌:blue[4. Filter Sheet1 for only new sample data]", key='CB4')
    if checkbox1 and checkbox2 and checkbox3 and checkbox4:
        F_excel_data_ao = Xli_P2_4(F_excel_data_ao)
        if F_excel_data_ao:
            st.write(':green[Xli_P2_4 finished.]')
            # Tạo nút tải xuống
            st.download_button(
                label="📥 Tải file Excel (Data_tracker_add2sheet_4.xlsx)",
                data=F_excel_data_ao.getvalue(),
                file_name="Data_tracker_add2sheet_4.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    checkbox5 = st.checkbox("📌:blue[5. Check if facilities in Sheet1 are active]", key='CB5')
    if checkbox1 and checkbox2 and checkbox3 and checkbox4 and checkbox5:
        F_excel_data_ao = Xli_P2_5(F_excel_data_ao)
        if F_excel_data_ao:
            st.write(':green[Xli_P2_5 finished.]')
            # Tạo nút tải xuống
            st.download_button(
                label="📥 Tải file Excel (Data_tracker_add2sheet_5.xlsx)",
                data=F_excel_data_ao.getvalue(),
                file_name="Data_tracker_add2sheet_5.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    checkbox6 = st.checkbox("📌:blue[6. Choose the parameters to track in Sheet1]", key='CB6')
    if checkbox1 and checkbox2 and checkbox3 and checkbox4  and checkbox5 and checkbox6:
        F_excel_data_ao = Xli_P2_6(F_excel_data_ao)

        if F_excel_data_ao:
            st.write(':green[Xli_P2_6 finished.]')
            # Tạo nút tải xuống
            st.download_button(
                label="📥 Tải file Excel (Data_tracker_add2sheet_6.xlsx)",
                data=F_excel_data_ao.getvalue(),
                file_name="Data_tracker_add2sheet_6.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )


    checkbox7 = st.checkbox("📌:blue[7. Make sure all the samples in Sheet1 are in mg/L and not ug/L]", key='CB7')
    if checkbox1 and checkbox2 and checkbox3 and checkbox4  and checkbox5  and checkbox6  and checkbox7:
        F_excel_data_ao = Xli_P2_7(F_excel_data_ao)

        if F_excel_data_ao:
            st.write(':green[Xli_P2_7 finished.]')
            # Tạo nút tải xuống
            st.download_button(
                label="📥 Tải file Excel (Data_tracker_add2sheet_7.xlsx)",
                data=F_excel_data_ao.getvalue(),
                file_name="Data_tracker_add2sheet_7.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )


    checkbox8 = st.checkbox("📌:blue[8. Add facility information from Sheet2 into Sheet1]", key='CB8')
    if checkbox1 and checkbox2 and checkbox3 and checkbox4  and checkbox5  and checkbox6  and checkbox7 and checkbox8:
        F_excel_data_ao = Xli_P2_8(F_excel_data_ao)

        if F_excel_data_ao:
            st.write(':green[Xli_P2_8 finished.]')
            # Tạo nút tải xuống
            st.download_button(
                label="📥 Tải file Excel (Data_tracker_add2sheet_8.xlsx)",
                data=F_excel_data_ao.getvalue(),
                file_name="Data_tracker_add2sheet_8.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    checkbox9 = st.checkbox("📌:blue[9. Add in SIC Codes from Sheet2 into Sheet1]", key='CB9')
    if checkbox1 and checkbox2 and checkbox3 and checkbox4  and checkbox5  and checkbox6  and checkbox7 and checkbox8  and checkbox9:
        F_excel_data_ao = Xli_P2_9(F_excel_data_ao)

        if F_excel_data_ao:
            st.write(':green[Xli_P2_9 finished.]')
            # Tạo nút tải xuống
            st.download_button(
                label="📥 Tải file Excel (Data_tracker_add2sheet_9.xlsx)",
                data=F_excel_data_ao.getvalue(),
                file_name="Data_tracker_add2sheet_9.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    checkbox10 = st.checkbox("📌:blue[10. Combine new data from Sheet1 into existing Data tracker]", key='CB10')
    if checkbox1 and checkbox2 and checkbox3 and checkbox4  and checkbox5  and checkbox6  and checkbox7 and checkbox8  and checkbox9:
        F_excel_data_ao = Xli_P2_10(F_excel_data_ao)

        if F_excel_data_ao:
            st.write(':green[Xli_P2_10 finished.]')
            # Tạo nút tải xuống
            st.download_button(
                label="📥 Tải file Excel (Data_tracker_add2sheet_10.xlsx)",
                data=F_excel_data_ao.getvalue(),
                file_name="Data_tracker_add2sheet_10.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )


# III Them data moi vao trinh theo doi -------------------------------------------------------------
st.subheader('✅ III. Analyze the new data', divider=True)

# IV Them data moi vao trinh theo doi -------------------------------------------------------------
st.subheader('✅ IV. Visualize the data', divider=True)

