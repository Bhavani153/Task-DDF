from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from orange_hrm_locators import Test_Locators
from excel_functions import Test_Excel_Functions

excel_file = r"E:\Automation testing\test_data.xlsx"

sheet_number = "Sheet1"

s = Test_Excel_Functions(excel_file, sheet_number)

chrome_service = ChromeService(r"E:\Automation testing\chromedriver.exe")
driver = webdriver.Chrome(service=chrome_service)

driver.get("https://opensource-demo.orangehrmlive.com/web/index.php/auth/login")
driver.maximize_window()
driver.implicitly_wait(60)

rows = Test_Excel_Functions(excel_file, sheet_number).Row_Count()

for row in range(2, rows+1):
    username = Test_Excel_Functions(excel_file, sheet_number).Read_Data(row, 6)
    password = Test_Excel_Functions(excel_file, sheet_number).Read_Data(row, 7)

    driver.find_element(by=By.XPATH, value=Test_Locators().username_locator).send_keys(username)
    driver.find_element(by=By.XPATH, value=Test_Locators().password_locator).send_keys(password)
    driver.find_element(by=By.XPATH, value=Test_Locators().login_button_locator).click()

     # Wait for the page to load and check if login was successful
    driver.implicitly_wait(10)
    if 'https://opensource-demo.orangehrmlive.com/web/index.php/auth/login' in driver.current_url:
        print("SUCCESS : Login Success with Username {a}".format(a = username))
        Test_Excel_Functions(excel_file, sheet_number).Write_Data(row,8, "TEST PASS")
        driver.back()
    elif('http://www.orangehrm.com/' in driver.current_url):
        print("FAIL : Login Failure with Username {a}".format(a = username))
        Test_Excel_Functions(excel_file, sheet_number).Write_Data(row,8, "TEST FAIL")
        driver.back()

# Close the browser
driver.quit()