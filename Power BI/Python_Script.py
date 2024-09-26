from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.service import Service
import time
import zipfile
import os
import re
import shutil

#Part 1

options = Options()
options.set_preference("browser.download.folderList", 2)
options.set_preference("browser.download.dir", "F:\Data_Automation_Project\Temp")
s=Service(r"C:\Users\HPADMIN\Downloads\geckodriver-v0.30.0-win64\geckodriver.exe")
browser = webdriver.Firefox(service=s,options=options)
browser.get('https://hourani.lawsyst.biz/common/login.php')
time.sleep(10)
WebDriverWait(browser, 1000000).until(EC.element_to_be_clickable((By.NAME, 'ms_adbtn'))).click()
time.sleep(10)
username = browser.find_element(by=By.NAME,value='loginfmt')
username.send_keys('itservice@houranipartners.com')
WebDriverWait(browser, 1000000).until(EC.element_to_be_clickable((By.ID, 'idSIButton9'))).click()
time.sleep(10)
password = browser.find_element(by=By.NAME,value='passwd')
password.send_keys('Dav81188')
time.sleep(10)
WebDriverWait(browser, 1000000).until(EC.element_to_be_clickable((By.ID, 'idSIButton9'))).click()
time.sleep(10)
WebDriverWait(browser, 1000000).until(EC.element_to_be_clickable((By.ID, 'idSIButton9'))).click()
time.sleep(10)
browser.get('https://hourani.lawsyst.biz/common/user-listing.php')
time.sleep(10)
WebDriverWait(browser, 1000000).until(EC.element_to_be_clickable((By.ID, 'backupcreatebtn'))).click()
WebDriverWait(browser, 1000000).until(EC.element_to_be_clickable((By.ID, 'backupdownloadlbl'))).click()
time.sleep(60)
browser.close()
filedir = os.listdir(r"F:\Data_Automation_Project\Main_Data\CSV")
if filedir != []:
    for i in filedir:
        fp = os.path.join(r"F:\Data_Automation_Project\Main_Data\CSV",i)
        os.remove(fp)
filedir2 = os.listdir(r"F:\Data_Automation_Project\Main_Data")
if filedir2 != []:
    for i in filedir2:
        if ".xlsx" in i:
            fp = os.path.join(r"F:\Data_Automation_Project\Main_Data",i)
            os.remove(fp)
ls  = os.listdir(r"F:\Data_Automation_Project\Temp")
ls.sort()
fpath = os.path.join(r"F:\Data_Automation_Project\Temp",ls[0])
fh = open(fpath, 'rb')
z = zipfile.ZipFile(fh)
# print(z.namelist())
for name in z.namelist():
    outpath = r"F:\Data_Automation_Project\Main_Data\CSV"
    z.extract(name, outpath)
fh.close()
for i in os.listdir(r"F:\Data_Automation_Project\Main_Data\CSV"):
    src = os.path.join(r"F:\Data_Automation_Project\Main_Data\CSV", i)
    nm = re.sub('\d', '', i)
    nm1 = nm.replace('_','').replace('-','') 
    dst = os.path.join(r"F:\Data_Automation_Project\Main_Data\CSV", str(nm1))
    try:
        os.rename(src,dst)
    except:
        pass
dsc = os.path.join(r"F:\Data_Automation_Project\Archive",ls[0])
shutil.move(fpath,dsc)
filedir2 = os.listdir(r"F:\Data_Automation_Project\Temp")
for i in filedir2:
    fp = os.path.join(r"F:\Data_Automation_Project\Temp",i)
    os.remove(fp)
print("Part 1 Completed")

# #Part 2
# options = Options()
# options.set_preference("browser.download.folderList", 2)
# options.set_preference("browser.download.dir", "F:\Data_Automation_Project\Main_Data")
# s=Service(r"C:\Users\HPADMIN\Downloads\geckodriver-v0.30.0-win64\geckodriver.exe")
# browser = webdriver.Firefox(service=s,options=options)
# browser.get('https://hourani.lawsyst.biz/common/login.php')
# time.sleep(10)
# WebDriverWait(browser, 1000000).until(EC.element_to_be_clickable((By.NAME, 'ms_adbtn'))).click()
# time.sleep(10)
# username = browser.find_element(by=By.NAME,value='loginfmt')
# username.send_keys('itservice@houranipartners.com')
# WebDriverWait(browser, 1000000).until(EC.element_to_be_clickable((By.ID, 'idSIButton9'))).click()
# time.sleep(10)
# password = browser.find_element(by=By.NAME,value='passwd')
# password.send_keys('Dav81177')
# time.sleep(10)
# WebDriverWait(browser, 1000000).until(EC.element_to_be_clickable((By.ID, 'idSIButton9'))).click()
# time.sleep(10)
# WebDriverWait(browser, 1000000).until(EC.element_to_be_clickable((By.ID, 'idSIButton9'))).click()
# time.sleep(10)
# browser.get('https://hourani.lawsyst.biz/common/staff-time-entry-listing.php')
# time.sleep(10)
# lenOfPage = browser.execute_script("window.scrollTo(0, document.body.scrollHeight);var lenOfPage=document.body.scrollHeight;return lenOfPage;")
# match=False
# while(match==False):
#         lastCount = lenOfPage
#         time.sleep(3)
#         lenOfPage = browser.execute_script("window.scrollTo(0, document.body.scrollHeight);var lenOfPage=document.body.scrollHeight;return lenOfPage;")
#         if lastCount==lenOfPage:
#             match=True
# time.sleep(10)
# WebDriverWait(browser, 1000000).until(EC.element_to_be_clickable((By.ID, 'matter_time_entry_report_csv'))).click()
# time.sleep(40)
# browser.close()
# print("Part 2 Completed")

# #Part 3
# options = Options()
# options.set_preference("browser.download.folderList", 2)
# options.set_preference("browser.download.dir", "F:\Data_Automation_Project\Main_Data")
# s=Service(r"C:\Users\HPADMIN\Downloads\geckodriver-v0.30.0-win64\geckodriver.exe")
# browser = webdriver.Firefox(service=s,options=options)
# browser.set_page_load_timeout(20)
# browser.get('https://hourani.lawsyst.biz/common/login.php')
# time.sleep(10)
# WebDriverWait(browser, 1000000).until(EC.element_to_be_clickable((By.NAME, 'ms_adbtn'))).click()
# time.sleep(10)
# username = browser.find_element(by=By.NAME,value='loginfmt')
# username.send_keys('itservice@houranipartners.com')
# WebDriverWait(browser, 1000000).until(EC.element_to_be_clickable((By.ID, 'idSIButton9'))).click()
# time.sleep(10)
# password = browser.find_element(by=By.NAME,value='passwd')
# password.send_keys('Dav81177')
# time.sleep(10)
# WebDriverWait(browser, 1000000).until(EC.element_to_be_clickable((By.ID, 'idSIButton9'))).click()
# time.sleep(10)
# WebDriverWait(browser, 1000000).until(EC.element_to_be_clickable((By.ID, 'idSIButton9'))).click()
# time.sleep(10)
# try:
#     browser.get('https://hourani.lawsyst.biz/common/ajax_total_wip_excel_report.php')
#     time.sleep(5)
#     browser.close()
# except Exception as e:    
#     browser.close()
# print("Part 3 Completed")

#Part 4
options = Options()
options.set_preference("browser.download.folderList", 2)
options.set_preference("browser.download.dir", "F:\Data_Automation_Project\Main_Data")
s=Service(r"C:\Users\HPADMIN\Downloads\geckodriver-v0.30.0-win64\geckodriver.exe")
browser = webdriver.Firefox(service=s,options=options)
browser.set_page_load_timeout(20)
browser.get('https://hourani.lawsyst.biz/common/login.php')
time.sleep(10)
WebDriverWait(browser, 1000000).until(EC.element_to_be_clickable((By.NAME, 'ms_adbtn'))).click()
time.sleep(10)
username = browser.find_element(by=By.NAME,value='loginfmt')
username.send_keys('itservice@houranipartners.com')
WebDriverWait(browser, 1000000).until(EC.element_to_be_clickable((By.ID, 'idSIButton9'))).click()
time.sleep(10)
password = browser.find_element(by=By.NAME,value='passwd')
password.send_keys('Dav81188')
time.sleep(10)
WebDriverWait(browser, 1000000).until(EC.element_to_be_clickable((By.ID, 'idSIButton9'))).click()
time.sleep(10)
WebDriverWait(browser, 1000000).until(EC.element_to_be_clickable((By.ID, 'idSIButton9'))).click()
time.sleep(10)
try:
    browser.get('https://hourani.lawsyst.biz/common/excel-reports/r-aging-excel.php?staff_id=&branch_id=')
    time.sleep(450)
#time.sleep(5)    
    browser.close()
except Exception as e:    
    browser.close()
print("Part 2 Completed")



