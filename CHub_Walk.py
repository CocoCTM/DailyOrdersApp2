"""
Corina Elisii
2022-08-25
"""

"""
This .py file contains all of the functions associated with walking CommerceHub
"""

# IMPORT PYTHON STANDARD LIBRARIES
from selenium import webdriver
import chromedriver_binary
from selenium.webdriver.chrome.options import Options
from fake_useragent import UserAgent
import pandas as pd
import bs4  # note that this is the beautifulsoup4 module
from bs4 import BeautifulSoup
import time

# WEB OPTIONS FUNCTION
# These options are set for Chrome Driver
options = Options()


# OPTION FUNCTION TO SET CHROME DRIVER OPTIONS
def option(comp_user_agent):
    """
    :type comp_user_agent: object
    @comp_user_agent: browser use agent on local device
    """
    options.add_argument(comp_user_agent)
    options.add_experimental_option("useAutomationExtension", False)
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_argument("disable-infobars")
    options.add_argument("--disable-extensions")
    print("Options have been set!")
    # print(options) # UNCOMMENT TO VERIFY OPTIONS


def open_browser(driver_path, login_page):
    """
    :type login_page: object
    @login_page: the CommerceHub Login Page
    :type driver_path: object
    @driver_path: the location of the chromedriver on local machine
    """
    global browser
    browser = webdriver.Chrome(executable_path=driver_path, options=options)
    browser.maximize_window()
    browser.get(login_page)
    time.sleep(3)
    print("CommerceHub Login Page located!")


def enter_user(user, pw, landing_page):
    username = browser.find_element_by_name("username")
    username.send_keys(user)
    browser.find_element_by_name("action").click()
    time.sleep(1)
    browser.find_element_by_name("action").click()
    password = browser.find_element_by_name("password")
    password.send_keys(pw)
    time.sleep(1)
    browser.find_element_by_name("action").click()
    time.sleep(1)
    browser.get(landing_page)
    browser.find_element_by_id("identitySelector").click()
    print("CommerceHub Login Successful!")


def open_ps_page():
    from selenium.common.exceptions import NoSuchElementException
    #pip install easygui
    import easygui
    import sys
    try:
        browser.find_element_by_xpath("//a[contains(@href,'gotoViewPackslips.do')]")
        pass
    except NoSuchElementException:
        easygui.msgbox("There are no new Packing Lists to download at this time!", title="No Packing Lists Found")
        browser.close()
        sys.exit(1)
        pass
    else:
        browser.find_element_by_xpath("//a[contains(@href,'gotoViewPackslips.do')]").click()
        print("Packing Slip Page Open!")


def get_order_count():
    time.sleep(5)
    # order_cnt = pd.read_html(browser.find_elements_by_id("//a[contains(@class='lineData')]").get_attribute('outerHTML'))[0]
    element = browser.find_element_by_id('fileDownloadTable')
    html_text = element.get_attribute('outerHTML')
    # order_cnt = pd.read_html(browser.find_element_by_xpath("//a[@id='fileDownloadTable')]").get_attribute('outerHTML'))[0]
    print("Order Count Created Successfully!")
    return html_text
