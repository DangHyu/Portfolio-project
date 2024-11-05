#import các thư viện
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
import pandas as pd
from sqlalchemy import create_engine
import sqlalchemy as sa
from time import sleep
import datetime as dt
import urllib

from selenium.webdriver.support.ui import Select

# import sqlalchemy as sa
import os
import time
# import schedule


# 1.    Date parameter
today           = dt.date.today()                                               # Ngày hôm nay dạng yyyy-mm-dd
yesterday       = today - dt.timedelta(days=1)                                  # Ngày hôm qua dạng yyyy-mm-dd    
bom_date        = yesterday.replace(day=1)                                      # Ngày đầu tháng dạng yyyy-mm-dd
begin_date      = today - dt.timedelta(days=85)
begin_1         = today - dt.timedelta(days=43)
end_1           = begin_1 - dt.timedelta(days=1)
f_date1         = dt.date.strftime(begin_1,'%d/%m/%Y')
f_date2         = dt.date.strftime(begin_date,'%d/%m/%Y')
t_date          = dt.date.strftime(yesterday,'%d/%m/%Y')
t_date2         = dt.date.strftime(end_1,'%d/%m/%Y')
f_date3         = dt.date.strftime(bom_date,'%d/%m/%Y')
f_yesterday     = dt.date.strftime(yesterday,'%d/%m/%Y')

Day_bom         = str(bom_date)[:4] + str(bom_date)[5:7] + str(bom_date)[8:]
Day_T_1         = str(yesterday)[:4] + str(yesterday)[5:7] + str(yesterday)[8:]
Day_T           = str(today)[:4] + str(today)[5:7] + str(today)[8:]
Month_T         = str(today)[:4] + str(today)[5:7]



##3.3   Path
##3.3.1 Login
login_user      = "/html/body/form/div[2]/table/tbody/tr[3]/td/table/tbody/tr/td/div/table/tbody/tr[4]/td[2]/div[2]/div/table/tbody/tr[2]/td[2]/input[1]"
login_pw        = "/html/body/form/div[2]/table/tbody/tr[3]/td/table/tbody/tr/td/div/table/tbody/tr[4]/td[2]/div[2]/div/table/tbody/tr[3]/td[2]/input"
login_button    = "/html/body/form/div[2]/table/tbody/tr[3]/td/table/tbody/tr/td/div/table/tbody/tr[4]/td[2]/div[2]/div/table/tbody/tr[4]/td[2]/input"


##3.3.2 Chọn báo cáo   
chon_BC_quan_tri  = "/html/body/form/div[3]/table/tbody/tr[3]/td/table/tbody/tr/td/div/table/tbody/tr[4]/td[1]/div/table/tbody/tr[1]/td/div/table/tbody/tr/td/a"
chon_link_BC         = 'https://report.shb.com.vn/shbreport/frmreport.aspx?rptname='
 

##3.3.3 Tab Báo cáo
Pham_vi_bao_cao     = "/html/body/form/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[1]/td/table/tbody/tr[1]/td[2]/select"
Ma_chi_nhanh        = "/html/body/form/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[1]/td/table/tbody/tr[2]/td[2]/select" 
Tu_ngay             = "/html/body/form/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[1]/td[2]/input"
Den_ngay            = "/html/body/form/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/input"
Loai_KH             = "/html/body/form/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[5]/td[2]/input"
Loai_KH_2           = "/html/body/form/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[6]/td[2]/input"
Loai_KH_3           = "/html/body/form/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[4]/td[2]/input"
Loai_KH_4           = "/html/body/form/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[3]/td[2]/input"
Xuat_du_lieu_tho    = "/html/body/form/table/tbody/tr/td/table/tbody/tr[3]/td/input[4]"
Xuat_du_lieu_lon    = "/html/body/form/table/tbody/tr/td/table/tbody/tr[3]/td/input[5]"


##3.7   File



##5.1   Connect Chrome
service         = Service(executable_path=driverlink)
options         = webdriver.ChromeOptions()
driver          = webdriver.Chrome(service=service, options=options)



## ĐĂNG NHẬP VÀO SHBREPORT
##Connect URL => Mở shbreport
driver.set_page_load_timeout(10000)
driver.get(url)
sleep(1) 

driver.find_element(By.XPATH,login_button).click() 

def parameter(a):
    if      a == 1:
        driver.find_element(By.XPATH,Tu_ngay).clear()
        sleep(1)
        driver.find_element(By.XPATH,Tu_ngay).send_keys(f_yesterday)
        sleep(1)
        driver.find_element(By.XPATH,Pham_vi_bao_cao).click()
        sleep(1)
        driver.find_element(By.XPATH,Xuat_du_lieu_tho).click()  
    elif    a == 2:
        driver.find_element(By.XPATH,Den_ngay).clear()
        sleep(1)
        driver.find_element(By.XPATH,Den_ngay).send_keys(f_yesterday)
        sleep(1)
        driver.find_element(By.XPATH,Tu_ngay).clear()
        sleep(1)
        driver.find_element(By.XPATH,Tu_ngay).send_keys(f_yesterday)
        sleep(1)
        driver.find_element(By.XPATH,Pham_vi_bao_cao).click()
        sleep(1)
        driver.find_element(By.XPATH,Xuat_du_lieu_tho).click()     
    elif    a == 3:
        driver.find_element(By.XPATH,Den_ngay).clear()
        sleep(1)
        driver.find_element(By.XPATH,Den_ngay).send_keys(f_yesterday)
        sleep(1)
        driver.find_element(By.XPATH,Tu_ngay).clear()
        sleep(1)
        driver.find_element(By.XPATH,Tu_ngay).send_keys(f_yesterday)
        sleep(1)
        driver.find_element(By.XPATH,Loai_KH_2).send_keys("R")
        sleep(1)
        driver.find_element(By.XPATH,Pham_vi_bao_cao).click()
        sleep(1)
        driver.find_element(By.XPATH,Xuat_du_lieu_tho).click() 

    elif    a == 4:            
        driver.find_element(By.XPATH,Tu_ngay).clear()
        sleep(1)
        driver.find_element(By.XPATH,Tu_ngay).send_keys(f_yesterday)
        sleep(1)
        driver.find_element(By.XPATH,Pham_vi_bao_cao).click()
        sleep(1)
        driver.find_element(By.XPATH,Loai_KH).send_keys("R")
        sleep(1)
        driver.find_element(By.XPATH,Xuat_du_lieu_tho).click()                                                           
    elif    a == 5:            
        driver.find_element(By.XPATH,Tu_ngay).clear()
        sleep(1)
        driver.find_element(By.XPATH,Tu_ngay).send_keys("31/12/2024")
        sleep(1)
        driver.find_element(By.XPATH,Pham_vi_bao_cao).click()
        sleep(1)
        driver.find_element(By.XPATH,Loai_KH).send_keys("R")
        sleep(1)
        driver.find_element(By.XPATH,Xuat_du_lieu_tho).click()  
            

        
    elif    a == 6:            
        driver.find_element(By.XPATH,Tu_ngay).clear()
        sleep(1)
        driver.find_element(By.XPATH,Tu_ngay).send_keys(f_yesterday)
        sleep(1)
        driver.find_element(By.XPATH,Pham_vi_bao_cao).click()
        sleep(1)
        driver.find_element(By.XPATH,Loai_KH_3).send_keys("R")
        sleep(1)
        driver.find_element(By.XPATH,Xuat_du_lieu_tho).click()  
        
        
    elif    a == 7:   
        driver.find_element(By.XPATH,Den_ngay).clear()
        sleep(1)
        driver.find_element(By.XPATH,Den_ngay).send_keys(f_yesterday)
        sleep(1)
        driver.find_element(By.XPATH,Tu_ngay).clear()
        sleep(1)
        driver.find_element(By.XPATH,Tu_ngay).send_keys('01/01/2024')
        sleep(1)
        driver.find_element(By.XPATH,Loai_KH_2).send_keys("R")
        sleep(1)
        driver.find_element(By.XPATH,Pham_vi_bao_cao).click()
        sleep(1)
        driver.find_element(By.XPATH,Xuat_du_lieu_tho).click() 

    elif    a == 8: 

        driver.find_element(By.XPATH,Tu_ngay).clear()
        sleep(1)
        driver.find_element(By.XPATH,Tu_ngay).send_keys(f_yesterday)
        sleep(1)
        driver.find_element(By.XPATH,Pham_vi_bao_cao).click()
        sleep(1)
        #Tìm phần tử dropdown 
        Chon_BAO_CAO = Select(driver.find_element(By.NAME,'DDLP_TYPE'))
        Chon_BAO_CAO.select_by_visible_text('POS')
        sleep(1)
        driver.find_element(By.XPATH,Xuat_du_lieu_tho).click()  
        
    elif    a == 9:   
        driver.find_element(By.XPATH,Den_ngay).clear()
        sleep(1)
        driver.find_element(By.XPATH,Den_ngay).send_keys(f_yesterday)
        sleep(1)
        driver.find_element(By.XPATH,Tu_ngay).clear()
        sleep(1)
        driver.find_element(By.XPATH,Tu_ngay).send_keys('01/01/2024')
        sleep(1)
        driver.find_element(By.XPATH,Pham_vi_bao_cao).click()
        sleep(1)
        driver.find_element(By.XPATH,Xuat_du_lieu_tho).click()         
      
    elif    a == 10:
        driver.find_element(By.XPATH,Den_ngay).clear()
        sleep(1)
        driver.find_element(By.XPATH,Den_ngay).send_keys(f_yesterday)
        sleep(1)
        driver.find_element(By.XPATH,Tu_ngay).clear()
        sleep(1)
        driver.find_element(By.XPATH,Tu_ngay).send_keys(f_date3)
        sleep(1)
        driver.find_element(By.XPATH,Loai_KH_4).send_keys("I")
        sleep(1)
        driver.find_element(By.XPATH,Pham_vi_bao_cao).click()
        sleep(1)
        driver.find_element(By.XPATH,Xuat_du_lieu_tho).click()

    elif    a == 11:
        driver.find_element(By.XPATH,Den_ngay).clear()
        sleep(1)
        driver.find_element(By.XPATH,Den_ngay).send_keys(f_yesterday)
        sleep(1)
        driver.find_element(By.XPATH,Tu_ngay).clear()
        sleep(1)
        driver.find_element(By.XPATH,Tu_ngay).send_keys(f_date3)
        sleep(1)
        driver.find_element(By.XPATH,Pham_vi_bao_cao).click()
        sleep(1)
        driver.find_element(By.XPATH,Xuat_du_lieu_tho).click()
        
    elif    a == 12:   
        driver.find_element(By.XPATH,Den_ngay).clear()
        sleep(1)
        driver.find_element(By.XPATH,Den_ngay).send_keys(f_yesterday)
        sleep(1)
        driver.find_element(By.XPATH,Tu_ngay).clear()
        sleep(1)
        driver.find_element(By.XPATH,Tu_ngay).send_keys('17/05/2024')
        sleep(1)
        driver.find_element(By.XPATH,Pham_vi_bao_cao).click()
        sleep(1)
        driver.find_element(By.XPATH,Xuat_du_lieu_tho).click()
      
###################################################################################################

def export_shbreport(file_prefix, a):
    driver.get(url)
    #Chon báo cáo cần xuất
    select = driver.find_element(By.XPATH,chon_BC_quan_tri)
    select.click()
    driver.get(chon_link_BC + file_prefix)
    
    #Xuất báo cáo
    parameter(a)
    # return export_shbreport



###################################################################################################
def import_excel_to_sql_server(excel_file_path,file_prefix,b): 

    # Đọc dữ liệu từ tệp Excel vào DataFrame
    df = pd.read_excel(excel_file_path)
    # df = df.applymap(lambda x:str(x).replace('\n',''))
    print(df)
    
    if file_prefix == 'LN006': 
        df['MA_CIF'] = df['MA_CIF'].map(lambda x: str(x).zfill(10))
    if file_prefix == 'LN009':   
        df['MRS_CUST_ID'] = df['MRS_CUST_ID'].map(lambda x: str(x).zfill(10))
    if file_prefix == 'LN016':   
        df['MLD_CUST_ID'] = df['MLD_CUST_ID'].map(lambda x: str(x).zfill(10))
        # df['NOTE_DESC'] = df['NOTE_DESC'].str.replace('\n','')  
    if file_prefix == 'LN036':   
        df['CIF'] = df['CIF'].map(lambda x: str(x).zfill(10))
    if file_prefix == 'LN056A':   
        df['CUST_ID'] = df['CUST_ID'].map(lambda x: str(x).zfill(10))
        df['NOTE_DESC'] = df['NOTE_DESC'].str.replace('\n','')         
        
    
    #convert unicode
    convert_list = ['TKVAY','RPT_TYPE','RPT_DATE','POS_DESC','CODE_NAME','POS_CD','CODE','TRANGTHAINO','TAX_ID','SUB_PURPOSE','SOHOPDONG','SI','SALARY_DEDUCTION','RPT_NAME','PURPOSE','PSP_SPEC_DESC','PP_PROD_DESC','POS_DESC','PHONGBAN','PHAN_LOAI_NO','PHAM_VI','NORM_SPECIAL','NGAYKYHOPDONG','NGAYKETTHUC','NGAYGIAINGAN','NGAYDAOHAN_SCHD','NGAYDAOHAN','NGAY','MPI_SPRD_CD','MPI_SANC_AUTHORTITY_DESC','MPI_REPMT_CCY','MPI_PROD_CD','MPI_LEGACY_ID','MPI_CRDT_SCH_CD','MATURITY_DATE_FLAG','MAN_FLG','MAKER_HACHTOANGIAINGAN','MAIN_SELL_NAME','MAIN_SELL','LOAN_OFFSET_NEXT_VALUEDATE','LOAI_TIEN','LOAI_LICH_BOOKING','LICH_AUTO_FLAG','LENDINGSTAFF_NAME','LENDINGSTAFF','LENDING_CONTRACT_LIMIT','LAST_REPMT_SCHD','KYHANTT','KY_DCLS_FLA_RT_TYPE','KY_DCLS_FLA_RT_SCHEME','INT_EFFECT_DT','INT_DURATION','INDEX_TENOR_UNIT','INDEX_RESET_DATE','INDEX_MRM_RATE_TYPE','FLOAT_SCHEME','ECO_CODE','ECO_ACTIVE','DONVI_QUANLY','DONVI','DIENTHOAI','DIACHI','CUSTYPE','CUSTNAME','CSL_LAST_EFF_DATE','CROSS_SELL_NAME','CROSS_SELL','CO_SPECIAL','CIF_COMPANY','CHUCDANH','CHECKER_HACHTOANGIAINGAN','BUSINESS_CD','TYPE_DESC','TO_DATE','PRODUCT','POS_DESC','MAIN_POS_DESC','FROM_DATE','POS_DESC','NOTE_FLG','NOTE_DESC','NOTE_CD','NGAYGIAINGAN','MPI_LOAN_TYPE','MPI_LEGACY_ID','MKR_ID','MKR_DT','MAIN_POS_DESC','EOD_DATE','AUTH_ID','AUTH_DT','POS_DESC','NGAYTATTOAN','NGAYGIAINGAN','NGAYBATDAUQUAHANLAI','NGAYBATDAUQUAHANGOC','MAIN_POS_DESC','LOAITIEN','DENNGAY','CUSTOMER_TYPE','CIF_NAME','CBTD','ACCTNO','TRANGTHAINO_TV','TRANGTHAINO','TO_DATE','POS_DESC','NORM_SPECIAL','NGAYGIAINGAN','NGAYDAOHAN','MPI_REPMT_CCY','MPI_LOAN_TYPE','MPI_LEGACY_ID','MAIN_POS_DESC','MA_SAN_PHAM','LOAI_LICH_BOOKING','DUE_DATE','DIENTHOAI','CUSTNAME','CBTD','TRANGTHAINO_TV','TRANGTHAINO_HIENTAI','TO_DATE','TENKH','TEN_SP_VAY','TEN_NGUOI_THUHUONG','SUB_PURPOSE','SOHOPDONG','RPT_NAME','PURPOSE','NORM_SPECIAL','NGAYKYHOPDONG','NGAYKETTHUC','NGAY_GIAI_NGAN','NGAY_GD_GIAI_NGAN','NGAY_DAO_HAN','MPI_SANC_AUTHORTITY_DESC','MAN_FLG','MAIN_SELL','MA_SP_VAY','LOAITIEN','INTEREST_DURATION','INDEX_RESET_DATE','FROM_DATE','ECO_ACTIVE','CUST_TYPE','CROSS_SELL','CORE_COMP_NAME','CO_SPECIAL','CHINHANH','CAN_BO_QLY','CAN_BO_GD','CAN_BO_DUYET','TO_DATE','RPT_NAME','NAT','LV4_DESC','LV3_DESC','LV2_DESC','LV1_DESC','FROM_DATE','DUDAUCO_CONG_PHATSINH','DUCUOICO','CHENHLECH','AC_TYPE','AC_DESC','MA_CIF','DIENTHOAI','MRS_CUST_ID','CIF','ACCTNO','ACCTNO1','CUST_ID','MLD_CUST_ID']
    #convert = {c:sa.Unicode(255) for c in convert_list}
    convert = {c:sa.Unicode if (c in convert_list or (c not in convert_list and df[c].dtype == 'O')) else sa.Float() for c in df.columns}

    # Tên bảng
    if b == 1: 
        table_name = f"{file_prefix}_{Day_T_1}"
        #table_name = f"{file_prefix}_20240922"
        df.to_sql(table_name, engine, if_exists='replace', index=False, dtype = convert)
    
    if b == 2: 
        table_name = f"{file_prefix}_20240101_{Day_T_1}"
        df.to_sql(table_name, engine, if_exists='replace', index=False, dtype = convert)
        
    if b == 3: 
        table_name = f"{file_prefix}_20241231_X{Day_T}"
        df.to_sql(table_name, engine, if_exists='replace', index=False, dtype = convert)
    
    if b == 4: 
        table_name = f"SAOKETINDUNG_{Day_T_1}"
        #table_name2= f"{file_prefix}_{Day_T_1}"  
        #table_name = "SAOKETINDUNG_20240623"
        df.to_sql(table_name, engine, if_exists='replace', index=False, dtype = convert)
        #df.to_sql(table_name2, engine, if_exists='replace', index=False, dtype = convert)
    if b == 5:   
        table_name = f"{file_prefix}"
        df.to_sql(table_name, engine, if_exists='append', index=False, dtype = convert)
    if b == 6:   
        table_name = f"{file_prefix}_20240101_{Day_T_1}"
        df.to_sql(table_name, engine, if_exists='replace', index=False, dtype = convert)
    if b == 7:   
        table_name = f"{file_prefix}_{Month_T}"
        df.to_sql(table_name, engine, if_exists='replace', index=False, dtype = convert) 

    # Đóng kết nối 
    engine.dispose()
    print(f"Imported data from {excel_file_path} to SQL Server successfully.")


###################################################################################################
def check_and_import_excel (file_prefix,a,b):
    folder_path = r'C:\Users\huybd1\Downloads'

    current_date = time.strftime("%Y%m%d")

    if a == 7: file_name = f"{file_prefix}_{current_date}_HUYBD1 (1).xlsx"
    else: file_name = f"{file_prefix}_{current_date}_HUYBD1.xlsx"
    
    excel_file_path = os.path.join(folder_path, file_name)

    if os.path.isfile(excel_file_path):
        import_excel_to_sql_server(excel_file_path, file_prefix,b)
    else:
        export_shbreport(file_prefix,a)
        sleep(60)
        import_excel_to_sql_server(excel_file_path, file_prefix,b)


 
# ################################################ Bắt đầu xuất báo cáo ################################################
#check_and_import_excel('NHS006',2,5)
#check_and_import_excel('GL017B1',2,1) 
# check_and_import_excel('LN009',5,3)
# check_and_import_excel('LN056A',6,1)
# check_and_import_excel('LN016',4,4)
#check_and_import_excel('CU07B',9,6)
check_and_import_excel('LN036',1,1)
check_and_import_excel('LN016',4,4)
#check_and_import_excel('LN009',5,3)
#check_and_import_excel('LN024',3,1)