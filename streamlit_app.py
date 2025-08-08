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
    fxlsx='Data_Tracker_2Sheet.xlsx'
    with pd.ExcelWriter(fxlsx, engine="openpyxl") as writer:
        df1.to_excel(writer, sheet_name="Sheet1", index=False)
        df2.to_excel(writer, sheet_name="Sheet2", index=False)
        df_data.to_excel(writer, sheet_name="Data", index=False)
    return fxlsx

    #output = BytesIO()
    #with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
    #    df1.to_excel(writer, sheet_name="Sheet1", index=False)
    #    df2.to_excel(writer, sheet_name="Sheet2", index=False)
    #    df_data.to_excel(writer, sheet_name="Data", index=False)
    #output.seek(0)
    #return output
# Cac ham Phan II -----------------------------    
def Xli_P2_0(data_tracker_upload):
    df = pd.read_excel(data_tracker_upload, sheet_name='Sheet2')
    # Tìm APP_ID bị trùng
    duplicates = df[df.duplicated(subset='APP_ID', keep=False)]
    # Giữ lại dòng trùng có STATUS khác "Active"
    to_delete = duplicates[duplicates['STATUS'] != 'Active']
    # Xoá các dòng này khỏi dataframe gốc (chu y la file excel van y cu)
    df_cleaned = df.drop(to_delete.index)
    # ket qua la cac dong trung APP_ID nhung co STATUS la Active duoc giu lai, con
    # cac dong trung APP_ID nhung co STATUS khac Active thi bi xoa
    # cho hien thi df con lai sau khi da xoa cac to_delete
    # phai hieu df_cleaned la df con lai sau khi da lam sach 
    st.write(df_cleaned)
    os.startfile(data_tracker_upload)
    return

def Xli_P2_1(data_tracker_upload):
    return
def Xli_P2_2(data_tracker_upload):
    return
def Xli_P2_3(data_tracker_upload):
    return
def Xli_P2_4(data_tracker_upload):
    return


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
# TIEU DE APP
st.header('🏷️Trình hỗ trợ quản lý môi trường nước')

# Phan sider ben trai ---------------------------------------------------------------------------
with st.sidebar:
    st.header('Lập trình theo tài liệu này:')
    # Đọc nội dung file Markdown
    with open("hd-lam-app-cho-thong.md", "r", encoding="utf-8") as f:
        md_content = f.read()
    st.markdown(md_content, unsafe_allow_html=True)


# I. TAI FILES TXT DU LIEU VE TU -----------------------------------------------------------------
st.subheader('✅ I. Download the data', divider=True)
ThucThiPhan_1()

# II Them data moi vao trinh theo doi -------------------------------------------------------------
st.subheader('✅ II. Add the new data to your tracker', divider=True)
# Upload 3 files
uploaded_files = st.file_uploader(
    'Upload 1 lần 3 files '+':red[(nên đặt 3 files này trước trong 1 thư mục)]',
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
            df_data = pd.read_excel(uploaded_f3, sheet_name="Data")  # Chỉ đọc sheet "Data"
        except Exception as e:
            st.error(f"⚠️ Lỗi khi đọc file: {e}")
            st.stop()
        # Dua 2 txt vao excel Data_Tracker.xlsx tu 3 file tai len
        # va tra ve file ao data_tracker_upload da chua them 2 txt   
        data_tracker_upload = Txt_to_data_tracker(df1, df2,df_data)
        #---
        # lap menu cac ham xu li tung buoc tep  data_tracker_upload
        op_listCacBuocP2 = {
            "0).Tìm các dòng APP_ID trùng, rồi xóa các dòng mà STATUS bằng Active nhưng giữ lại các dòng có STATUS khác Active ": Xli_P2_0, 
            "1). Sắp xếp theo 3 cột": Xli_P2_1,
            "2). Cột OLD/NEW có giá trị new": Xli_P2_2,
            "3). Cột OLD/NEW có giá trị old": Xli_P2_3,
            "4). So sánh giá trị max giữa các cơ sở": Xli_P2_4 
        }
        # menu chon ham/viec
        viec_chon = st.selectbox(
            ":blue[Chọn hàm xử lí Phần II]", 
            (op_listCacBuocP2.keys()),
            index=None,
            placeholder="Chua chon ham nao...",
        )
        # chay ham da chon
        if viec_chon:
            # chay ham tuong ung voi key chon_with_viec, ham nay co ten la gia tri cua key do, 
            # them () de chay ham, tham so la file excel da tai len
            op_listCacBuocP2[viec_chon](data_tracker_upload)   # 👉 Gọi hàm tuong ung

        #--
        # Xu li 2 sheet txt tren data_tracker_upload append vao Data sheet
        # va tra ve tep  Data_Tracker_New de download
        #Data_Tracker_New = Combined_to_data_tracker(data_tracker_upload)    

        #st.success("✅ Đã tạo Data_Tracker_New.xlsx tu 3 file tai len.")
        #st.download_button(
        #    label="📥 Tải xuống file Excel mới: Data_Tracker_New.xlsx",
        #    data=Data_Tracker_New,
        #    file_name="Data_Tracker_New.xlsx",
        #    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        #)
    else:
        st.warning('Hãy đảm bảo đúng file "...Industrial_Application...txt", file "...Industrial_ad_hoc...txt", và file "Data_tracker.xlsx" ')
else:
    st.info('Vui lòng upload đủ 3 file: "...Industrial_Application...txt", "...Industrial_Ad_Hoc...txt", và "Data_tracker.xlsx" ')

#ThucThiPhan_2()

# III Phan tich du lieu------------------------------------------------------------------------
st.subheader('✅ III. Analyze the new data', divider=True)
uploaded_file = st.file_uploader("Tải lên file: " + ":red[Data_Tracker_New.xlsx]", type=["xlsx"])
         
if uploaded_file:
    # Doc file da tai len de ghi du lieu o sheet Data vao df 
    df = pd.read_excel(uploaded_file, sheet_name="Data")

    # dung df ghi tạm ra file Excel dat ten la output để xử lý bằng openpyxl
    output = BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)

    # lap menu cac ham xu li tep output 
    op_listCaiMuonXem = {
        "0). Xem tổng quát tệp  Data_Tracker": Ht_Data_tquat, 
        "1). Sắp xếp theo 3 cột": Ht_Data_sxep,
        "2). Cột OLD/NEW có giá trị new": Ht_Data_new,
        "3). Cột OLD/NEW có giá trị old": Ht_Data_old,
        "4). So sánh giá trị max giữa các cơ sở": Ht_Data_max 
    }
    # menu chon ham/viec
    viec_chon = st.selectbox(
        ":blue[Chọn hàm xử lí Data với kiểu hiển thị]", 
        (op_listCaiMuonXem.keys()),
        index=None,
        placeholder="Chon hien thi...",
    )
    # chay ham da chon
    if viec_chon:
        # chay ham tuong ung voi key chon_with_viec, ham nay co ten la gia tri cua key do, 
        # them () de chay ham, tham so la file excel da tai len
        op_listCaiMuonXem[viec_chon](output)   # 👉 Gọi hàm tuong ung

# IV Do thi hoa du lieu --------------------------------------------------------------------------
st.subheader('✅ IV. Visualize the data', divider=True)
#ThucThiPhan_4()
st.write(':red[Trình đang viết thử để chạy trên Streamlit Cloud.Chưa xong...]')