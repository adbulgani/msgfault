from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException
import os, glob, pandas as pd, pywhatkit as pwt


driver = webdriver.Chrome(executable_path='msgvenv/lib64/python3.8/site-packages/selenium/webdriver/common/chromedriver')
driver.get('https://fms.bsnl.in')
id_box = driver.find_element_by_name('username')
id_box.send_keys('sreedharhmnd_apvsk')
pass_box = driver.find_element_by_name('password')
pass_box.send_keys('01012018')
signin_box = driver.find_element_by_id('submit-form')
signin_box.click()
        
try:
    element = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.ID, "faultOrders")))
except FileNotFoundError:
    FileNotFoundError


element.click()
driver.implicitly_wait(5)
file_faults = driver.find_elements_by_class_name('dt-buttons')

for elements in file_faults:
    elements.click()

list_of_files = glob.glob('/home/gani/Downloads/*')
latest_file = max(list_of_files, key=os.path.getctime)
print(latest_file)

#Read from Excel
xl= pd.ExcelFile(latest_file)

#Parsing Excel Sheet to DataFrame
dfs = xl.parse(xl.sheet_names[0])

df = dfs[dfs["EXCH"]=='VSKMND'].head()
faults = []
#print(df['PHONE'])
for index, row in df.iterrows():
    faults.append('#' + ' ' + str(row['PHONE']) + ' ' + str(row['MDF DETAILS :']) + ' ' + str(row['CUSTOMER']) + ' ')

message = ' '
for i in range(len(faults)):
    message += faults[i]

try:
  # sending message to receiver
  # using pywhatkit
  pwt.sendwhatmsg("+917901290140",
                        message,
                        21, 55
                        )
  print("Successfully Sent!")
except:
  print("An Unexpected Error!")