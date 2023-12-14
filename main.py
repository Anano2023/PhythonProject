from selenium import webdriver
from selenium.webdriver import Chrome
from selenium.webdriver.common.by import By
import openpyxl
import pytest
import logging

#Performing data driven testing with pytest

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

@pytest.fixture
def driver_setup():
    path = "D:\\seleniumServer\\chromedriver-win64\\chromedriver.exe"
    driver = Chrome()
    yield driver
    driver.quit()

def extract_data_from_excel(file_path, sheet_name='Sheet1'):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook[sheet_name]
    data = []

    for row in sheet.iter_rows(min_row=2, max_col=5, values_only=True):
        name, email, fav_c, age, hf = row
        data.append((name, email, fav_c, age, hf))

    workbook.close()
    return data


# Usage
file_path = 'pyData.xlsx'
data = extract_data_from_excel(file_path)

@pytest.mark.parametrize("name, email, fav_c, age, hf", data)
def test_process_data(driver_setup, name, email, fav_c, age, hf):
    driver_setup.get("D:\\pythonProject1\\txt.html")

    driver_setup.find_element(By.NAME, 'nameField').send_keys(name)
    logger.info(f"Entering name: {name}")

    driver_setup.find_element(By.ID, 'emailInput').send_keys(email)
    logger.info(f"Entering email: {email}")

    driver_setup.find_element(By.CSS_SELECTOR, '.colorInput').send_keys(fav_c)
    logger.info(f"Entering favorite color: {fav_c}")

    driver_setup.find_element(By.XPATH, "//input[@placeholder='Find by XPath']").send_keys(age)
    logger.info(f"Entering age: {age}")

    driver_setup.find_element(By.XPATH, "//div[@id='parent']/input[1]").send_keys(hf)
    logger.info(f"Entering hidden field: {hf}")
