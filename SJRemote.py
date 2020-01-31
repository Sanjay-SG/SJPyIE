from selenium import webdriver
from selenium.webdriver.common.keys import Keys

import time

import xlsxwriter   

# driver = webdriver.Ie()

driver = webdriver.Ie(executable_path = 'C:\Sanjay\IEDriver\IEDriverServer.exe')


driver.get("http://hsean196-uwfr.uesc.com/utils/security/securityhome.jsp")

driver.switch_to.frame("main")
# driver.switch_to.default_content()
uname_element = driver.find_element_by_xpath("//input[@id='username']")
uname_element.send_keys("ssg")

password_element = driver.find_element_by_xpath("//input[@id='password']")
password_element.send_keys("admin")

airline_element = driver.find_element_by_xpath("//input[@id='airline']")
airline_element.send_keys("UW")

# driver.find_element_by_name("btnLogin").click()

# driver.find_element_by_xpath("/html/body/form/table/tbody/tr/td/table/tbody/tr[6]/td/input").click()

driver.find_element_by_xpath("/html/body/form/table/tbody/tr/td/table/tbody/tr[6]/td/input").send_keys(Keys.ENTER);

# driver.findElement(By.id("searchMovielanding")).sendKeys(KEYS.ENTER);
time.sleep(2)

driver.switch_to.default_content()


driver.switch_to.frame("main")


# driver.find_element_by_link_text("Back Office Agent").click()
# driver.find_element_by_xpath("//html/body/table/tbody/tr/td/table[1]/tbody/tr/td[2]/table/tbody/tr[4]/td[1]/a").click()
time.sleep(4)

driver.find_element_by_xpath("/html/body/table/tbody/tr/td/table[1]/tbody/tr/td[2]/table/tbody/tr[4]/td[1]/a").send_keys(Keys.ENTER)

time.sleep(2)

driver.switch_to.default_content()

driver.switch_to.frame("banner")

time.sleep(2)

# driver.find_element_by_partial_link_text("Utilities").click()

driver.find_element_by_partial_link_text("Utilities").send_keys(Keys.ENTER)

driver.switch_to.default_content()

driver.switch_to.frame("contents")

time.sleep(2)

driver.find_element_by_partial_link_text("Administrate").send_keys(Keys.ENTER)

driver.switch_to.default_content()

driver.switch_to.frame("main")

time.sleep(2)

driver.find_element_by_partial_link_text("CFDSDSR").send_keys(Keys.ENTER)

popupXpath ="/html[1]/body[1]/div[2]/table[1]/tbody[1]/tr[3]/td[1]/a[1]"

# print(dir(driver))

parentWindowHandler = driver.current_window_handle

time.sleep(2)
# parentWindowHandler = driver.getWindowHandle
handles = driver.window_handles;
size = len(handles);

for x in range(size):
# 	driver.switch_to.window(handles[x]);
	subWindowHandler = handles[x]
# 	print (driver.title);

# driver.switch_to_window(subWindowHandler)

time.sleep(2)

driver.switch_to.window(subWindowHandler)

time.sleep(2)

driver.find_element_by_partial_link_text("Start Job").send_keys(Keys.ENTER)

driver.switch_to.window(parentWindowHandler)

driver.switch_to.default_content()

driver.switch_to.frame("main")

filename = driver.find_element_by_name("txtInputParam0")

time.sleep(2)

filename.send_keys("FILE")

airlineInput = driver.find_element_by_name("txtInputParam1")

time.sleep(2)

airlineInput.send_keys("DL")

fileType = driver.find_element_by_name("txtInputParam2")

time.sleep(2)

fileType.send_keys("H")