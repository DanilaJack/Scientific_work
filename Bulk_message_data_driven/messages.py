from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from credentials import Login, Password, Msg
import XLUtils
import time

# set your .exe file location
srv_obj = Service("C:/Users/danil/Desktop/automated_testing/Selenium Web Automation/SeleniumPython_PavanCourse/WebDrivers/chromedriver.exe")
driver = webdriver.Chrome(service=srv_obj)
driver.implicitly_wait(10)

# Launching URL
driver.get("https://edu.vsu.ru/")
driver.maximize_window()

# Authorization
driver.find_element(By.ID, "login_username").send_keys(Login)
time.sleep(1)
driver.find_element(By.ID, "login_password").send_keys(Password)
time.sleep(1)
driver.find_element(By.XPATH, "//input[@value='Вход']").click()

# moving to messages page
driver.find_element(By.XPATH, "//a[@id='action-menu-toggle-0']").click()
driver.find_element(By.XPATH, "//a[@data-title='messages,message']").click()
time.sleep(2)

Private = driver.find_element(By.XPATH, "//div[@data-init='true']//span[@class='font-weight-bold'][normalize-space()='Private']")
Private.click()
time.sleep(2)
roster = driver.find_elements(By.XPATH, "//div[@class='list-group']//a//strong")
print(len(roster))

file = "C:/Users/danil/Desktop/НИИР/Bulk_message_data_driven/testData.xlsx"
rows = XLUtils.getRowCount(file, "Data")

Starred = driver.find_element(By.XPATH, "//span[@class='font-weight-bold'][normalize-space()='Starred']")
k = 1
for r in range(1, rows+1):
    for i in range(k, 5):
        if (roster[i].text == XLUtils.readData(file, "Data", r, 1)):
            roster[i].click()
            time.sleep(3)
            Starred.click()
            time.sleep(3)
            driver.find_element(By.XPATH, "//div[@data-region='content-messages-footer-container']//textarea[@data-region='send-message-txt']").send_keys(Msg)
            time.sleep(1)
            driver.find_element(By.XPATH, "//div[@data-region='content-messages-footer-container']//textarea[@data-region='send-message-txt']").clear()
            time.sleep(1)
            Private.click()
            time.sleep(3)
            k += 1
            break
        else:
            continue
time.sleep(1)
driver.quit()
