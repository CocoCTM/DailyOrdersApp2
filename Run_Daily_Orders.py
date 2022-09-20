"""
Corina Elisii
2022-08-25
"""
import CHub_Walk as cw
import Variable_Selection as vs
import pandas as pd
import re
import Lowes_SetUp as lsu
import TheHomeDepot_SetUp as thdsu
import BestBuy_SetUp as bbsu
import NetSuite_Imports as nts
from datetime import date
import shutil
import os

"""
Setting Up Global Variable
"""
pack_dir = 'C:/Users/corina/Downloads/'

"""
The First Section of the Code will be to set up and use the Options and Parameters used for Selenium Library. 
1. Find your Computer's User Agent: 
2. Download the ChromeDriver and add the Path: 
"""
comp_user_agent = "user-agent= Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.0.0 Safari/537.36"
driver_path = '/bin/chromedriver.exe'

# LOGGING INTO COMMERCEHUB
cw.option(comp_user_agent)
cw.open_browser(driver_path, vs.chub_ostream)
cw.enter_user(vs.chub_user, vs.chub_pass, vs.chub_home)
cw.open_ps_page()

# GET PACKING SLIP OVERVIEW FOR ALL CUSTOMERS
html_order_count = cw.get_order_count()
order_count = pd.read_html(html_order_count)
chub_ord = order_count[0]
chub_ord = chub_ord[chub_ord['Partner'] != 'View']  # Remove items where rows = 'View'

# CREATE DF COLLECTION TO STORE EACH CUSTOMER ORDER NUMBERS
dataframe_collection = {}

# COLLECT EACH ORDER NUMBER PER CUSTOMER
for i in range(len(chub_ord)):
    fileref = chub_ord.iloc[i, 1]  # Change if File Name changed column
    fileref = re.findall(r'\d+', fileref)
    order_dtl = cw.ps_handling(fileref[0], vs.chub_ps)
    dataframe_collection[chub_ord.iloc[i, 0]] = order_dtl
    print(chub_ord.iloc[i, 0]+" Order List Downloaded")
    # Move Downloaded Packing Slip to Customer Folder - RUN 1 TIME WITH THIS COMMENTED
"""
    file_name = fileref[0] + '.pdf'
    new_dir = "Y:/User/Order Entry User/CTM Household Appliances Orders Processed/" + chub_ord.iloc[
        i, 0] + "/NEW - PROCESS/"
    try:
        shutil.move(pack_dir + file_name, new_dir)
    except:
        print("error") + str(IOError)
        pass
"""
"""
COLLECT THE ORDER DETAILS WITH CHUB SAVED SEARCH
Lowe's: B: LOWES Undelivered
THD: A: THD Undelivered
BestBuy: 
THD Special:
"""

# 1. LOWE'S
if "Lowe's" in dataframe_collection:
    lowes_ord_dtl = lsu.download_ord_dtls(dataframe_collection["Lowe's"], pack_dir)
    lowes_ord_dtl["Customer ID"] = nts.get_cust_id("Lowe's")
    lowes_ord_dtl = nts.get_item_id(lowes_ord_dtl)
    lowes_ord_dtl = nts.get_ship_from(lowes_ord_dtl)
    lowes_ord_dtl = lowes_ord_dtl.drop_duplicates()
    lowes_ord_dtl["External ID"] = lowes_ord_dtl["Sending Partner"] + lowes_ord_dtl["PO Number"].astype(str) + lowes_ord_dtl["Location"]
    lowes_ord_dtl = lsu.get_lowes_carrier(lowes_ord_dtl)
    lowes_ord_dtl = lowes_ord_dtl.drop(
        columns=['Total', 'Order Date_y', 'Merchant', 'Merchant SKU', 'Merchant Department',
                 'UPC', 'Unit Cost Currency', 'ShipTo Country', 'ShipTo Customer Number', 'ShipTo Day Phone',
                 'ShipTo First Name', 'Status', 'Substatus', 'Ship From_y', 'ShipTo State_y',
                 'Carrier'])  # DROP UNNECESSARY COLUMNS
    lowes_ord_dtl['ShipTo Postal Code'] = lowes_ord_dtl['ShipTo Postal Code'].astype(str)  # Customer Order # IF NOT 9 DIGITS ADD 0 IN FRONT
    lowes_ord_dtl['ShipTo Postal Code'] = lowes_ord_dtl['ShipTo Postal Code'].str.strip()  # Customer Order # IF NOT 9 DIGITS ADD 0 IN FRONT
    lowes_ord_dtl['ShipTo Postal Code'] = lowes_ord_dtl['ShipTo Postal Code'].str.zfill(5)  # Customer Order # IF NOT 9 DIGITS ADD 0 IN FRONT
    print("Lowe's extract completed")
    lowes_ord_dtl.to_excel(
        r"Y:/User/Order Entry User/CTM Household Appliances Orders Processed/Lowe's/FILES/Daily Order/Lowe's_Order_DTLS_"
        + str(date.today()) + ".xlsx", index=False, header=True)
else:
    print("No Lowe's Orders Today!")

# 2. THD
if "The Home Depot Inc" in dataframe_collection:
    thd_ord_dtl = thdsu.download_ord_dtls(dataframe_collection["The Home Depot Inc"], pack_dir)
    thd_ord_dtl["Customer ID"] = nts.get_cust_id("The Home Depot, Inc.")
    thd_ord_dtl = nts.get_item_id(thd_ord_dtl)
    thd_ord_dtl = nts.get_ship_from(thd_ord_dtl)
    thd_ord_dtl = thd_ord_dtl.drop_duplicates()
    thd_ord_dtl["External ID"] = thd_ord_dtl["Sending Partner"] + thd_ord_dtl["PO Number"].astype(str) + thd_ord_dtl[
        "Location"]
    thd_ord_dtl = thdsu.get_thd_carrier(thd_ord_dtl)
    thd_ord_dtl = thd_ord_dtl.drop(
        columns=['Total', 'Order Date_y', 'Merchant', 'Merchant SKU', 'Merchant Department', 'UPC',
                 'Unit Cost Currency', 'ShipTo Country', 'ShipTo Customer Number',
                 'ShipTo First Name', 'Status', 'Substatus', 'Ship From_y', 'ShipTo State_y', 'Carrier',
                 'ShipTo Address2', 'ShipTo Address3', 'ShipTo County', 'ShipTo Email', 'ShipTo Last Name',
                 'ShipTo Night Phone',
                 'Warehouse ID', 'Line Expected Warehouse ID'])  # DROP UNNECESSARY COLUMNS
    thd_ord_dtl['ShipTo Postal Code'] = thd_ord_dtl['ShipTo Postal Code'].astype(
        str)  # ZIP IF NOT 5 DIGITS ADD 0 IN FRONT
    thd_ord_dtl['ShipTo Postal Code'] = thd_ord_dtl['ShipTo Postal Code'].str.strip()  # ZIP IF NOT 5
    thd_ord_dtl['ShipTo Postal Code'] = thd_ord_dtl['ShipTo Postal Code'].str.zfill(5)  # ZIP IF NOT 5
    thd_ord_dtl['PO Number'] = thd_ord_dtl['PO Number'].astype(str)  # PO# # IF NOT 8 DIGITS ADD 0 IN FRONT
    thd_ord_dtl['PO Number'] = thd_ord_dtl['PO Number'].str.strip()  # PO# # IF NOT 8 DIGITS ADD 0 IN FRONT
    thd_ord_dtl['PO Number'] = thd_ord_dtl['PO Number'].str.zfill(8)  # PO# # IF NOT 8 DIGITS ADD 0 IN FRONT
    print("THD extract completed")
    thd_ord_dtl.to_excel(
        r"Y:/User/Order Entry User/CTM Household Appliances Orders Processed/The Home Depot Inc/FILES/Daily Order/THD_Order_DTLS_"
        + str(date.today()) + ".xlsx", index=False, header=True)
else:
    print("No The Home Depot Orders Today!")

# 3. BESTBUY
if "Best Buy" in dataframe_collection:
    bb_ord_dtl = bbsu.download_ord_dtls(dataframe_collection["Best Buy"], pack_dir)
    bb_ord_dtl["Customer ID"] = nts.get_cust_id("BestBuy")
    bb_ord_dtl = nts.get_item_id(bb_ord_dtl)
    bb_ord_dtl = nts.get_ship_from(bb_ord_dtl)  # Manually put Ship From = TN for now 2022-09-01
    bb_ord_dtl = bb_ord_dtl.drop_duplicates()
    bb_ord_dtl["External ID"] = bb_ord_dtl["Sending Partner"] + bb_ord_dtl["PO Number"].astype(str) + bb_ord_dtl[
        "Location"]
    bb_ord_dtl["Customer Order Number"] = "BB#:" + bb_ord_dtl["Customer Order Number"]
    bb_ord_dtl['ShipTo Postal Code'] = bb_ord_dtl['ShipTo Postal Code'].astype(
        str)  # ZIP IF NOT 5 DIGITS ADD 0 IN FRONT
    bb_ord_dtl['ShipTo Postal Code'] = bb_ord_dtl['ShipTo Postal Code'].str.strip()  # ZIP IF NOT 5
    bb_ord_dtl['ShipTo Postal Code'] = bb_ord_dtl['ShipTo Postal Code'].str.zfill(9)  # ZIP IF NOT 5
    bb_ord_dtl = bb_ord_dtl.drop(
        columns=['Total', 'Order Date_y', 'Merchant', 'Merchant SKU', 'Merchant Department', 'UPC',
                 'Unit Cost Currency', 'ShipTo Address Rate Class', 'ShipTo Company Name', 'ShipTo Address2',
                 'ShipTo Address3', 'ShipTo County', 'ShipTo Postal Code Ext', 'ShipTo Email', 'ShipTo Last Name',
                 'ShipTo Night Phone', 'ShipTo Country', 'ShipTo Customer Number',
                 'ShipTo First Name', 'Status', 'Substatus',
                 'Warehouse ID', 'Line Expected Warehouse ID'])  # DROP UNNECESSARY COLUMNS
    print("BB extract completed")
    bb_ord_dtl.to_excel(
        r"Y:/User/Order Entry User/CTM Household Appliances Orders Processed/Best Buy/FILES/Daily Order/BB_Order_DTLS_"
        + str(date.today()) + ".xlsx", index=False, header=True)
else:
    print("No Best Buy Orders Today!")

print("Done")
