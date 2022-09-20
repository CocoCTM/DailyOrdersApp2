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


# THIS FUNCTION AIMS TO LOG INTO THE COMMERCEHUB ORDERSTREAM PORTAL
def enter_user(user, pw, landing_page):
    """
    :type user: string
    @user: CommerceHub User Name
    :type pw: string
    @pw: CommerceHub Password
    :type landing_page
    @landing_page: CommerceHub Order Stream Home Page
    """
    username = browser.find_element_by_name("username")
    username.send_keys(user)
    browser.find_element_by_name("action").click()
    time.sleep(1)
    browser.find_element_by_name("action").click()
    password = browser.find_element_by_name("password")
    password.send_keys(pw)
    time.sleep(1)
    browser.find_element_by_name("action").click()
    time.sleep(5)
    browser.get(landing_page)
    time.sleep(5)
    browser.find_element_by_id("identitySelector").click()
    print("CommerceHub Login Successful!")


# ATTEMPT TO OPEN THE LINK WITH PACKING SLIPS - IF NOT AVAILABLE RETURN NO PACKING SLIPS YET FOR TODAY
def open_ps_page():
    time.sleep(5)
    from selenium.common.exceptions import NoSuchElementException
    # pip install easygui
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

# GIVEN THIS CODE IS REACHED, THERE ARE PACKING SLIPS.
# THIS TAKES THE SUMMARY PAGE OF PACKING SLIPS.
def get_order_count():
    time.sleep(5)
    # order_cnt = pd.read_html(browser.find_elements_by_id("//a[contains(@class='lineData')]").get_attribute('outerHTML'))[0]
    element = browser.find_element_by_id('fileDownloadTable')
    html_text = element.get_attribute('outerHTML')
    # order_cnt = pd.read_html(browser.find_element_by_xpath("//a[@id='fileDownloadTable')]").get_attribute('outerHTML'))[0]
    print("Order Count Created Successfully!")
    return html_text


def ps_handling(fileref, link):
    """
    :type fileref: string
    @fileref: PackingSlip# from CommerceHub in order to like on the correct VIEW link
    @:return Resulting Dataframe of Order Number on the PackingSlip#
    """
    time.sleep(1)
    browser.find_element_by_xpath("//*[@id=\"view-" + fileref + "\"]").click()
    time.sleep(5)
    element = browser.find_element_by_class_name('linedata')
    html_text = element.get_attribute('outerHTML')
    result = pd.read_html(html_text)[0]
    result.columns = result.iloc[0]
    result = result.iloc[1:, :]
    nav_to(link) # Navigate back to Page
   # browser.find_element_by_xpath("//*[@id=\"dl-status-" + fileref + "\"]").click() # Download Packing Slip
    time.sleep(2)
    return result


def get_curr_link():
    return browser.current_url

def nav_to(page):
    browser.get(page)

def dnwld_order_csv(searchName):
    """
    :type searchName: string
    @searchName: CommerceHub Saved Search Name to click for report
    @:return resulting Dataframe of Order Number on the PackingSlip#
    """
    browser.get('https://dsm.commercehub.com/dsm/gotoOrderSearch.do')
    time.sleep(5)
    browser.find_element_by_partial_link_text(searchName).click()
    time.sleep(3)
    all_pre_popup = browser.window_handles
    curr_handle = browser.current_window_handle
    browser.find_element_by_partial_link_text("CSV").click()
    all_post_popup = browser.window_handles
    set1 = set(all_pre_popup)
    set2 = set(all_post_popup)
    diff_handle = list(sorted(set2 - set1))
    diff_handle = diff_handle[0]
    browser.switch_to.window(diff_handle)
    curr_handle2 = browser.current_window_handle
    browser.find_element_by_css_selector('.chub-button').click()
    time.sleep(5)
    browser.switch_to.window(curr_handle)