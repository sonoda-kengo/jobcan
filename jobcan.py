from selenium import webdriver
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from time import sleep
import xlwings as xw


browser = webdriver.Chrome(ChromeDriverManager().install())
url ="https://ssl.jobcan.jp/login/pc-employee-global?lang_code=ja"
browser.get(url)
elem_client_id = browser.find_element(By.ID, 'client_id')
elem_client_id.send_keys('xxxxxxxxxxxx(company name)')
elem_email = browser.find_element(By.ID, 'email')
elem_email.send_keys('xxxxxxxxxxxx(mail_address)')
elem_password = browser.find_element(By.ID, 'password')
elem_password.send_keys('xxxxxxxxxxxx(your password)')
browser.find_element(By.XPATH, '//button').click()


try:
    browser.find_element(By.XPATH, '//*[@id="sidemenu-closed"]/div/button').click()
    sleep(3)
    browser.find_element(By.XPATH, '//a[text()="出勤簿"]').click()
except:
    browser.find_element(By.XPATH, '//a[text()="出勤簿"]').click()


sleep(3)

# get table
table = browser.find_element(By.XPATH, "//table[@class='table jbc-table text-center jbc-table-bordered jbc-table-hover']")
tbody = table.find_element(By.TAG_NAME, 'tbody')
trs = tbody.find_elements(By.TAG_NAME, 'tr')

# get year & month
data = browser.find_element(By.XPATH, "/html/body/div/div/div[2]/main/div/div/div/h4/div")
data_text = data.text

# get month
index_begin = data_text.find("年") + 1
index_end = data.text.find("度")
getMonth = data_text[index_begin:index_end]

# get working time
tds_dict = {}
for tr in trs:
    tds = tr.find_elements(By.TAG_NAME, 'td')
    if(tds[2].text == ''):
        if('(土)' in tds[0].text or '(日)' in tds[0].text):
            tds_dict[tds[0].text] = ['', '']
        else:
            tds_dict[tds[0].text] = ['9:00', '18:00']
    else:
        tds_dict[tds[0].text] = [tds[2].text, tds[3].text]

tds_dict

# get excel sheet
wb = xw.Book('ExcelFile/job_report.xlsx')
sht = wb.sheets[getMonth]

# insert working time
count = 0;
for i in tds_dict.keys(): 
    if(tds_dict[i][1]!='(勤務中)'):
        if(sht.range(f'O{count+6}').value == '○'):
            sht.range(f'D{count+6}').value = tds_dict[i][0]
            sht.range(f'E{count+6}').value = tds_dict[i][1]
    count+=1

# save sheet
wb.save('job_report.xlsx')