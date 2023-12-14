from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver import Chrome
import pytest
import openpyxl

@pytest.fixture(scope="session")
def test_setup():
    global driver
    path = "D:\\seleniumServer\\chromedriver-win64\\chromedriver.exe"
    driver = Chrome()
    driver.get("D:\\pythonProject1\\txt.html")

    # Teardown code
    yield driver
    driver.quit()

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

def test_Name(test_setup):
    driver.find_element(By.NAME, 'nameField').send_keys(name)

def test_Email():
    driver.find_element(By.ID, 'emailInput').send_keys(email)

def test_fav_col():
    driver.find_element(By.CSS_SELECTOR, '.colorInput').send_keys(fav_Color)

def test_Age():
    driver.find_element(By.XPATH, "//input[@placeholder='Find by XPath']").send_keys(age)

def test_hidden_field():
    driver.find_element(By.XPATH, "//div[@id='parent']/input[1]").send_keys(hf)






