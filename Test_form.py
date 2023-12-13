from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium import webdriver
import configparser
from selenium.webdriver import Chrome
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import time
import pytest
import openpyxl

@pytest.fixture()
def test_setup():
    global driver
    path = "D:\\seleniumServer\\chromedriver-win64\\chromedriver.exe"
    driver = Chrome()
    driver.get("D:\\pythonProject1\\txt.html")

def test_popUp(test_setup):

    driver.find_element(By.XPATH, "//button[text()= 'Open Popup']").click()
    # Move to the pop window
    txt = driver.window_handles[0]
    popUp = driver.window_handles[1]
    driver.switch_to.window(popUp)
    print(driver.current_url)
    driver.close()

    # Switch to the main page
    driver.switch_to.window(txt)

def test_frames():
    # Print text from Iframe1 and Iframe2
    iframe1 = driver.find_element(By.ID, 'iframe1')
    driver.switch_to.frame(iframe1)
    print(driver.find_element(By.ID, 'textIframe1').text)

    #Switch to default page
    driver.switch_to.default_content()

    #Switch to frame 2
    iframe2 = driver.find_element(By.ID, 'iframe2')
    driver.switch_to.frame(iframe2)
    print(driver.find_element(By.ID,'textIframe2').text)

    # Switch to default page
    driver.switch_to.default_content()

def test_excelConnection():
    global name,email, fav_Color, age, hf, sh

    #Open excel file
    workbook = openpyxl.load_workbook("pyData.xlsx")
    print(workbook.sheetnames)

    #Selecting desired sheet
    sh = workbook['Sheet1']

    #take values from sheet1
    name = sh['A2'].value
    email= sh['B2'].value
    fav_Color = sh['C2'].value
    age = sh['D2'].value
    hf = sh['E2'].value

def test_Name():
    driver.find_element(By.NAME, 'nameField').send_keys(name)

def test_Email():
    driver.find_element(By.ID, 'emailInput').send_keys(email)

def test_fav_col():
    driver.find_element(By.CSS_SELECTOR, '.colorInput').send_keys(fav_Color)

def test_Age():
    driver.find_element(By.XPATH, "//input[@placeholder='Find by XPath']").send_keys(age)

def test_hidden_field():
    driver.find_element(By.XPATH, "//div[@id='parent']/input[1]").send_keys(hf)

def test_pages():
    time.sleep(2)
    # Open new tab with onclick button
    driver.find_element(By.XPATH, "//button[text()= 'Open New Tab']").click()
    time.sleep(2)
    print(driver.window_handles)

    driver.switch_to.window(driver.window_handles[1])
    print(driver.current_window_handle)

    driver.get("https://google.com")
    time.sleep(2)

    driver.switch_to.window(driver.window_handles[0])
    driver.close()
    time.sleep(1)







