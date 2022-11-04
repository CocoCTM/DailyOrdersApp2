# IMPORT PYTHON STANDARD LIBRARIES
from selenium import webdriver
import chromedriver_binary
from selenium.webdriver.chrome.options import Options
from fake_useragent import UserAgent
import pandas as pd
import bs4  # note that this is the beautifulsoup4 module
from bs4 import BeautifulSoup
import time
import CHub_Walk as cw
import Variable_Selection as vs
import pandas as pd
import re
import Lowes_SetUp as lsu
import TheHomeDepot_SetUp as thdsu
import BestBuy_SetUp as bbsu
import NetSuite_Imports as nts
from datetime import date
import datetime



comp_user_agent = "user-agent= Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/106.0.0.0 Safari/537.36"
driver_path = '/bin/chromedriver.exe'
options = Options()
options.add_argument(comp_user_agent)
options.add_experimental_option("useAutomationExtension", False)
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_argument("disable-infobars")
options.add_argument("--disable-extensions")

global browser
browser = webdriver.Chrome(executable_path=driver_path, options=options)
browser.maximize_window()
browser.get("https://copilot2.pilotdelivers.com/login.aspx")


user = browser.find_element_by_id("rpLogin_emailTxt")
user.send_keys("michael@ctm-inter.com")
passw = browser.find_element_by_id("rpLogin_passTxt")
passw.send_keys("CTM123")
browser.find_element_by_id("rpLogin_btnLogin").click()

ord_data = pd.read_excel("Y:/User/Order Entry User/CTM Household Appliances Orders Processed/Best Buy/FILES/Daily Order/BB_Order_DTLS_2022-10-31.xlsx", sheet_name='Sheet1')
ns_loc_dir = 'Y:/User/Order Entry User/NETSUITE IMPORTS/LOCATION'
locdf = pd.read_csv(ns_loc_dir + "/NS_Location.csv")
ord_data = ord_data.merge(locdf, on="Location ID", how="left")



#CLICK SHIPMENT
browser.find_element_by_id("Shipment").click()


#SHIPPER
shipperadd = browser.find_element_by_id("ctl00_ctl00_BodyContent_PageCenterContent_ASPxCallbackPanel1_ASPxPageControl1_cbPanelShipperInfo_txtShipperAddress1_I")
shipperadd.clear()
shipperadd.send_keys("123 K Ave")

#Just send in ZIP Code and the rest will update
shipperzip = browser.find_element_by_id("ctl00_ctl00_BodyContent_PageCenterContent_ASPxCallbackPanel1_ASPxPageControl1_cbPanelShipperInfo_txtShipperPostalCode_I")
shipperzip.clear()
shipperzip.send_keys("12919")

#SHIP TO
#Consignee TO
shiptodc = browser.find_element_by_id("ctl00_ctl00_BodyContent_PageCenterContent_ASPxCallbackPanel1_ASPxPageControl1_txtConsigneeName_I")
shiptodc.send_keys("Best Buy DC")

#Customer Name
shiptoname = browser.find_element_by_id("ctl00_ctl00_BodyContent_PageCenterContent_ASPxCallbackPanel1_ASPxPageControl1_txtConsigneeAttention_I")
shiptoname.send_keys(ord_data["ShipTo Name"][0])

#Consignee Phone #
shiptophone = browser.find_element_by_id("ctl00_ctl00_BodyContent_PageCenterContent_ASPxCallbackPanel1_ASPxPageControl1_txtConsigneePhone_I")
shiptophone.send_keys(ord_data["ShipTo Day Phone"][0].astype(str))

#ShipTo Address
shiptoaddress = browser.find_element_by_id("ctl00_ctl00_BodyContent_PageCenterContent_ASPxCallbackPanel1_ASPxPageControl1_txtConsigneeAddress1_I")
shiptoaddress.send_keys(ord_data["ShipTo Address1"][0])
#ShipTo ZIP
shiptozip = browser.find_element_by_id("ctl00_ctl00_BodyContent_PageCenterContent_ASPxCallbackPanel1_ASPxPageControl1_txtConsigneePostalCode_I")
shiptozip.send_keys(ord_data["ShipTo Address1"][0])

#Shipper Reference PO
shipperPO = browser.find_element_by_id("ctl00_ctl00_BodyContent_PageCenterContent_ASPxCallbackPanel1_ASPxPageControl1_memShipperReference_I")
shipperPO.send_keys(ord_data["PO Number"][0])

#Consignee Reference PO
shiptoPO = browser.find_element_by_id("ctl00_ctl00_BodyContent_PageCenterContent_ASPxCallbackPanel1_ASPxPageControl1_memConsigneeReference_I")
shiptoPO.send_keys(ord_data["PO Number"][0])

#Today + 2
dd = datetime.date.strftime((datetime.datetime.today() + datetime.timedelta(days=2)), "%m/%d/%Y")

shipdate = browser.find_element_by_id("ctl00_ctl00_BodyContent_PageCenterContent_ASPxCallbackPanel1_calShipment_I")
browser.execute_script("document.getElementById('ctl00_ctl00_BodyContent_PageCenterContent_ASPxCallbackPanel1_calShipment_I').value = "+dd+";")