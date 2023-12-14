from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver import Chrome
import time
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

def test_pages():
    time.sleep(2)
    # Open new tab with onclick button
    driver.find_element(By.XPATH, "//button[text()= 'Open New Tab']").click()
    time.sleep(2)
    print(driver.window_handles)

    #Switch to the new window
    driver.switch_to.window(driver.window_handles[1])
    print(driver.current_window_handle)

    #Open webpage
    driver.get("https://google.com")
    time.sleep(2)

    #Going back to the main window and close it
    driver.switch_to.window(driver.window_handles[0])
    driver.close()
    time.sleep(1)
