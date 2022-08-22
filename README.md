# jobcan

### 1. Check version
windows 10 <br>
Python 3.10.4 <br>
pip 22.2.2 <br>
selenium 4.4.3 <br>

_if you need update_

    pip install --upgrade pip --user


### 2. Install selenium, webdriver-manager,xlwings
    pip install selenium

    pip install webdriver-manager --user

    pip install xlwings --user


### 3. Insert your info into jobcan.py
(1) company name in line 12

    elem_client_id.send_keys('{your company name})')

(2) e-mail in line 14

    elem_email.send_keys('{your mail_address}')

(3) password in line 16

    elem_password.send_keys('{your password}')


### 4 Put your Excel file
put 'job_report.xlsx' file in the 'ExcelFile' directory
