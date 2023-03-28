"""
Corina Elisii
2022-08-25
"""
import CHub_Walk as cw
import Variable_Selection as vs
import pandas as pd
import numpy as np
import re
import Lowes_SetUp as lsu
import TheHomeDepot_SetUp as thdsu
import BestBuy_SetUp as bbsu
import NetSuite_Imports as nts
import os
import win32com.client as win32
from pathlib import Path

from datetime import date
import shutil
import os

"""
Setting Up Global Variable
"""

pack_dir = 'C:/Users/corina/Downloads/'
gift_item = ['COB10002', 'COB10501', 'COB10502','COP30001','COP30002','COP30003','COP30501','COP30502','COP30503','COP30901','COP30902'] # GIFTWARE ITEMS
me_item = ['CF00260-WH1', 'CF00260-BL1', 'CF00260-TTH', 'CF00260-CPH', 'CF00266-TTR', 'CF00266-ORR', 'CF00272-TTR', 'CF00272-ORR', 'CF01360-BNP', 'CF01566-WH1', 'CF01566-BL1', 'CF01566-TT1', 'CF01566-CP1', 'CF01666-TTR',
'CF01666-ORR', 'CF01672-TTR', 'CF01672-AGR', 'CF02118-BB1', 'CF02118-BN1']
#me_item = ['']
"""
The First Section of the Code will be to set up and use the Options and Parameters used for Selenium Library. 
1. Find your Computer's User Agent: 
2. Download the ChromeDriver and add the Path: 
"""
comp_user_agent = "user-agent= Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36"
driver_path = '/bin/chromedriver.exe'

# LOGGING INTO COMMERCEHUB
cw.option(comp_user_agent)
cw.open_browser(driver_path, vs.chub_ostream)
cw.enter_user(vs.chub_user, vs.chub_pass, vs.chub_home)
cw.open_ps_page()


print('Final Excel sheet now generated at the same location:')
outputxlsx.to_excel("Y:/User/Order Entry User/CTM Household Appliances Orders Processed/Lowe's/FILES/NetSuiteImport_2023-02-23.xlsx", index=False)