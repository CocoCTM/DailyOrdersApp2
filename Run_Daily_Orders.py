"""
Corina Elisii
2022-08-25
"""
import CHub_Walk as cw
import Variable_Selection as vs
import pandas as pd

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

html_order_count = cw.get_order_count()
order_count = pd.read_html(html_order_count)
print(order_count)
print(type(order_count))