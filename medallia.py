from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import os.path
import codecs
import glob
import os
import shutil
import time
import datetime
from datetime import date, timedelta
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from dateutil.relativedelta import *
import os
from openpyxl import Workbook
from openpyxl import load_workbook
import win32com.client as win32

def download():
	driver.find_element_by_xpath("(//i[@class='icon icon-export icon-size-18'])[position()=1]").click()
	driver.find_element_by_xpath("//a[@data-selenium-id='obb_EXPORT']").click()

today = datetime.date.today()
################################################################Previous Week#############################################################
start_delta = datetime.timedelta(days=today.weekday(), weeks=1)
start_of_week = today - start_delta - timedelta(days=1)
End_of_week = start_of_week + timedelta(days=6)
sow=start_of_week.strftime('%m/%d/%y')
eow=End_of_week.strftime('%m/%d/%y')
################################################################Previous Month of Previous Year#############################################################
start_of_week = datetime.datetime.today() - relativedelta(months=1,years=1,days=today.day-1)
End_of_week = datetime.datetime.today() - relativedelta(years=1,days=today.day)
pys=start_of_week.strftime('%m/%d/%y')
print(pys)
pye=End_of_week.strftime('%m/%d/%y')
print(pye)
start=[sow,pys]
end=[eow,pye]

#############################################################################################################################################################
chrome_options = Options()

chrome_options.add_experimental_option("prefs", {
  "download.default_directory": "C:/Users/purushv/Downloads/Medallia",
  "download.prompt_for_download": False,
})

#chrome_options.add_argument("--headless")
driver = webdriver.Chrome(chrome_options=chrome_options)
#driver.minimize_window()
#driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
#params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': 'C:/Users/purushv/Downloads/Medallia'}}
#command_result = driver.execute("send_command", params)
#driver.minimize_window()
import base64
c=base64.standard_b64decode('VmlraXZpZ29AMDA3')
f=codecs.decode(c, encoding='utf-8', errors='strict')
url="https://express.medallia.com/reflections/homepage.do?v=b5YOFM1IaneA&id=529"

driver.get(url)
driver.find_element_by_id("username").send_keys("cking")
driver.find_element_by_id ('password').send_keys("1theateama6kige")
driver.find_element_by_id("logBtn").click()

driver.find_element_by_xpath("//a[@class='nav__action dropdown-trigger']").click()
driver.find_element_by_xpath("//a[@data-selenium-id='sections-link:Ranker']").click()

#driver.find_element_by_id("menu-link-Ranker").click()

############################################## Selecting North America and past 12 months to Date Data !##############################################################################

driver.find_element_by_xpath("(//span[@class='action-indicator multi'])[position()=6]").click()
driver.find_element_by_xpath("//span[@data-selenium-id='select-item-0']").click()
download()
print('DONE')
#url2="https://express.medallia.com/reflections/propertyranker.do?id=561&v=boVI3OJ5ahh9vLM4k4R*bEcSXOO9hM9hKJ6SF6qDnlT9UYYL10KKbJFhL-6-OkUd9.g6*FCWYRgYl6EjA5JTSfQ76kLI4xyOBOYvEp-7HB4-&exportFn=RankerExport.xls&sss=config.viewstate%3DbW3yPoZ84YPi04-Z7TY92mpcqR.ouh6AcLlL0A3P3Te8dDd7EH9u-buQvrZYE8kgOppsnYbvxbRBc.oyElhEYc3ZLv5Mu7j6iKgbCxI3f3sk3sRX2u795p2cR*f5uxE0UCwTs4u1yBo9Kk13VF_1ncVbmVJHKVQo"
#driver.get(url2)

################################################## custom Time period Data ##########################################################################
for i in range(2):
	driver.find_element_by_xpath("(//span[@class='name ng-binding'])[position()=1]").click()
	driver.find_element_by_xpath("//span[@data-selenium-id='select-item-35']").click()

	div_element=driver.find_element_by_xpath("(//span[@class='action-indicator single'])[position()=1]")
	hover = ActionChains(driver).move_to_element(div_element)
	hover.perform()
	try:
		driver.find_element_by_xpath("(//span[@class='edit ng-scope'])[position()=1]").click()
	except:
		driver.find_element_by_xpath("//span[@data-selenium-id='select-item-36']").click()
		div_element=driver.find_element_by_xpath("(//span[@class='name ng-binding'])[position()=1]")
		hover = ActionChains(driver).move_to_element(div_element)
		hover.perform()
		driver.find_element_by_xpath("(//span[@class='edit ng-scope'])[position()=1]").click()

	driver.find_element_by_xpath("(//input[@class='input-mini active'])[position()=1]").send_keys(Keys.CONTROL + "a")
	driver.find_element_by_xpath("(//input[@class='input-mini active'])[position()=1]").send_keys(Keys.DELETE)
	

	driver.find_element_by_xpath("(//input[@class='input-mini active'])[position()=1]").send_keys(start[i])
	driver.find_element_by_xpath("(//input[@class='input-mini'])[position()=1]").send_keys(Keys.CONTROL + "a")
	driver.find_element_by_xpath("(//input[@class='input-mini'])[position()=1]").send_keys(Keys.DELETE)

	driver.find_element_by_xpath("(//input[@class='input-mini'])[position()=1]").send_keys(end[i])
	driver.find_element_by_xpath("(//button[@class='applyBtn btn btn-sm btn-primary'])[position()=1]").click()
	driver.find_element_by_xpath("(//span[@class='action-indicator multi'])[position()=6]").click()
	driver.find_element_by_xpath("//span[@data-selenium-id='select-item-0']").click()
	download()
	#url2="https://express.medallia.com/reflections/propertyranker.do?id=561&v=boVI3OJ5ahh9vLM4k4R*bEcSXOO9hM9hKJ6SF6qDnlT9UYYL10KKbJFhL-6-OkUd9.g6*FCWYRgYl6EjA5JTSfQ76kLI4xyOBOYvEp-7HB4-&exportFn=RankerExport.xls&sss=config.viewstate%3DbW3yPoZ84YPi04-Z7TY92mpcqR.ouh6AcLlL0A3P3Te8dDd7EH9u-buQvrZYE8kgOppsnYbvxbRBc.oyElhEYc3ZLv5Mu7j6iKgbCxI3f3sk3sRX2u795p2cR*f5uxE0UCwTs4u1yBo9Kk13VF_1ncVbmVJHKVQo"
	#driver.get(url2)

############################ Previous Month and Past 3 Months #####################################################################
a=[21,19]
for i in range(2):
	driver.find_element_by_xpath("(//span[@class='name ng-binding'])[position()=1]").click()
	driver.find_element_by_xpath("//span[@data-selenium-id='select-item-"+str(a[i])+"']").click()
	driver.find_element_by_xpath("(//span[@class='action-indicator multi'])[position()=6]").click()
	try:
		driver.find_element_by_xpath("//span[@data-selenium-id='select-item-0']").click()
	except:
		pass
	download()
	#url2="https://express.medallia.com/reflections/propertyranker.do?id=561&v=boVI3OJ5ahh9vLM4k4R*bEcSXOO9hM9hKJ6SF6qDnlT9UYYL10KKbJFhL-6-OkUd9.g6*FCWYRgYl6EjA5JTSfQ76kLI4xyOBOYvEp-7HB4-&exportFn=RankerExport.xls&sss=config.viewstate%3DbW3yPoZ84YPi04-Z7TY92mpcqR.ouh6AcLlL0A3P3Te8dDd7EH9u-buQvrZYE8kgOppsnYbvxbRBc.oyElhEYc3ZLv5Mu7j6iKgbCxI3f3sk3sRX2u795p2cR*f5uxE0UCwTs4u1yBo9Kk13VF_1ncVbmVJHKVQo"
	#driver.get(url2)

driver.find_element_by_xpath("//a[@data-selenium-id='subsections:subsection.District/Affiliate']").click()
driver.find_element_by_xpath("(//span[@class='name ng-binding'])[position()=1]").click()
driver.find_element_by_xpath("//span[@data-selenium-id='select-item-"+str(a[0])+"']").click()
driver.find_element_by_xpath("(//span[@class='action-indicator multi'])[position()=6]").click()
download()
#url2="https://express.medallia.com/reflections/propertyranker.do?id=561&v=boVI3OJ5ahh9vLM4k4R*bEcSXOO9hM9hKJ6SF6qDnlT9UYYL10KKbJFhL-6-OkUd9.g6*FCWYRgYl6EjA5JTSfQ76kLI4xyOBOYvEp-7HB4-&exportFn=RankerExport.xls&sss=config.viewstate%3DbW3yPoZ84YPi04-Z7TY92mpcqR.ouh6AcLlL0A3P3Te8dDd7EH9u-buQvrZYE8kgOppsnYbvxbRBc.oyElhEYc3ZLv5Mu7j6iKgbCxI3f3sk3sRX2u795p2cR*f5uxE0UCwTs4u1yBo9Kk13VF_1ncVbmVJHKVQo"
#driver.get(url2)




time.sleep(20)
driver.quit()
names=['12_Months_Ranker','Previous_Week','Previous_Month_of_PY','Previous Month','Past_3_months','District_affl_previous_Month']
import os
os.chdir('C:/Users/purushv/Downloads/Medallia')

for file in os.listdir():
    src=file
    if src=='RankerExport.xls':
               dst=names[0]+".xls"
               os.rename(src,dst)
    else:
        for i in range(5):
            if src=='RankerExport ('+str(i+1)+').xls':
                   dst=names[i+1]+".xls"
                   os.rename(src,dst)
	######################################################################################################################################################################################################
######################################################################################################################################################################################################
######################################################################################################################################################################################################
######################################################################################################################################################################################################
######################################################################################################################################################################################################

for i in names:
	fname = 'C:\\Users\\purushv\\Downloads\\Medallia\\'+i+'.xls'
	excel = win32.gencache.EnsureDispatch('Excel.Application')
	wb = excel.Workbooks.Open(fname)

	wb.SaveAs(fname+"x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
	wb.Close()                               #FileFormat = 56 is for .xls extension
	excel.Application.Quit()
	
	os.remove(fname)


	# Instantiating a Workbook object by excel file path
	wb = load_workbook('C://Users//purushv//Downloads//Medallia//'+i+'.xlsx')
	ws = wb.active
	ws.delete_rows(1,7)

	# Saving the modified Excel file in default (that is Excel 2003) format
	wb.save('C://Users//purushv//Downloads//Medallia//'+i+'.xlsx')