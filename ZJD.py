from selenium import webdriver


driver = webdriver.Edge("H:\software\edgedriver_win64\msedgedriver.exe")

driver.get("https://www.baidu.com/")


# input = driver.find_element_by_css_selector("#kv")
input = driver.find_element_by_id("kw")
input.send_keys("abc")

button = driver.find_element_by_id("su")
button.click()
