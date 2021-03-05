from selenium.webdriver import Firefox
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select


opts = Options()

driver = Firefox(options=opts)

driver.get(
    "http://maximoui.corp.oncor.com:8080/maximo/webclient/login/login.jsp?appservauth=true"
)

driver.find_element_by_id("j_username").send_keys("u6zn")
driver.find_element_by_id("j_password").send_keys("q$zxc1kp")
driver.find_element_by_id("loginbutton").click()

WebDriverWait(driver, 60).until(
    EC.element_to_be_clickable((By.ID, "titlebar_hyperlink_6-lbshowmenu_reportsmenu"))
).click()


WebDriverWait(driver, 15).until(
    EC.element_to_be_clickable((By.LINK_TEXT, "Inventory"))
).click()
WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "//*[@id='menu0_INVENTOR_MODULE_a_tnode']"))
).click()
WebDriverWait(driver, 25).until(
    EC.element_to_be_clickable(
        (By.XPATH, "/html/body/form/table[1]/tbody/tr/td/ul/li[2]/ul/li[2]/a/span")
    )
).click()

hoursTable = driver.find_elements_by_css_selector("table.m48482eee_tbod-tbd tr")
for tr in driver.find_elements_by_id("m48482eee_tbod-tbd"):
    print(tr)
    tds = tr.find_elements_by_tag_name("td")
    print([td.text for td in tds])
# FTS!!!