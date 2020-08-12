from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys

driver = webdriver.Chrome("C:/chromedriver_win32/chromedriver.exe")
driver.get("http://wwmfg.analog.com/wwmfg/apps/TRS/SpecIndex.cfm")
usernameInp = driver.find_element_by_name("username")
usernameInp.clear()
usernameInp.send_keys("Mlokman")

passInp = driver.find_element_by_name("password")
passInp.clear()
passInp.send_keys("413C@Batuuban")
passInp.send_keys(Keys.RETURN)

select = Select(driver.find_element_by_name('FieldName'))

# select by visible text
select.select_by_visible_text('SpecNumber')

trsnumberInp = driver.find_element_by_name("FieldNameString")
trsnumberInp.clear()
trsnumberInp.send_keys("024332")
trsnumberInp.send_keys(Keys.RETURN)

driver.find_element_by_xpath("/html/body/div[2]/form/table/tbody/tr/td/table/tbody/tr[4]/td[1]/a").click()
driver.find_element_by_xpath("//*[@id=\"adi-sse-wwmfg-header-app-bar\"]/span/span[15]/a").click()

mrkup = driver.find_element_by_xpath("//*[@id=\"webkit-xml-viewer-source-xml\"]/TRSForm")
print(mrkup.get_attribute('innerHTML'))