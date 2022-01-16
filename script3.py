import xlrd
import time
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from datetime import datetime
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException        
from time import sleep
import selenium.webdriver.support.expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import sys

# xlrd.xlsx.ensure_elementtree_imported(False, None)
# xlrd.xlsx.Element_has_iter = True

driver = webdriver.Chrome(executable_path="./chromedriver")
driver.maximize_window()
driver.get('https://www.google.com')

file_path = "data.xlsx"
wb = xlrd.open_workbook(file_path)
sheet = wb.sheet_by_index(0)


def check_exists_by_xpath():
    try:
        button = driver.find_element(By.CSS_SELECTOR, 'button[elname="submit"]')
        if button.text.lower().strip() == "submit":
        	return True
        return False
    except NoSuchElementException:
        return False
    return True


def updateFunc(row, tab_id):
	try:
		tab_id = "tab" + str(tab_id)
		driver.execute_script("window.open('https://zfrmz.in/JZ0uRBYOZ8QOc9ynFSFy', '{tab_id}');".format(tab_id=tab_id))
		driver.switch_to.window(tab_id)	
		# time.sleep(1)

		enrollment_no = int(row[1])
		validation_status = row[2]
		entry_date = row[3]
		unique_no = int(row[4])
		officer_name = row[5].strip()
		tl_name = row[6].strip()
		location = row[7].strip()
		capturing = row[8].strip()
		check_list = row[9].strip()
		caused_list = row[10].strip()
		system_id = row[11]
		cafe_id = row[12]

		date_obj = xlrd.xldate.xldate_as_datetime(entry_date, wb.datemode)
		date_str = date_obj.strftime("%d-%b-%Y")


		driver.find_element(By.NAME, 'Number').send_keys(enrollment_no)
		button = driver.find_element(By.CSS_SELECTOR, 'button[elname="next"]')
		ActionChains(driver).move_to_element(button).click(button).perform()

		WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, 'MultipleChoice-multiple-id')))
		driver.find_element(By.ID, 'MultipleChoice-multiple-id').send_keys(validation_status)
		button = driver.find_element(By.XPATH, '/html/body/div[2]/div[1]/div[2]/div/form/div[2]/ul[3]/li/div[1]/div[2]/div/button')
		ActionChains(driver).move_to_element(button).click(button).perform()

		driver.find_element(By.NAME, 'Date').send_keys(date_str)
		driver.find_element(By.XPATH, '/html/body/div[2]/div[1]/div[2]/div/form/div[2]/ul[4]/li/div[1]/div[2]/div/button').click()

		driver.find_element(By.NAME, 'Number1').send_keys(unique_no)
		driver.find_element(By.NAME, 'Dropdown').send_keys(officer_name)
		driver.find_element(By.NAME, 'Dropdown1').send_keys(tl_name)
		button = driver.find_element(By.XPATH, '/html/body/div[2]/div[1]/div[2]/div/form/div[2]/ul[5]/li/div[1]/div[2]/div/button')
		ActionChains(driver).move_to_element(button).click(button).perform()

		driver.find_element(By.NAME, 'Dropdown2').send_keys(location)
		driver.find_element(By.NAME, 'Dropdown3').send_keys(capturing)
		driver.find_element(By.NAME, 'Dropdown4').send_keys(check_list)
		driver.find_element(By.NAME, 'Dropdown5').send_keys(caused_list)


		while check_exists_by_xpath():
			time.sleep(1)


	except Exception as e:
		exc_type, exc_obj, exc_tb = sys.exc_info()
		print(e, exc_tb.tb_lineno)
		raise e


first_idx = int(input("Enter first line index -> "))
last_idx = int(input("Enter last line index -> "))

row_cnt = 0
for i in range(first_idx, last_idx+1):
	row_cnt += 1
	try:
		row = []
		for j in range(0, 13):
			row.append(sheet.cell(i, j).value)
		updateFunc(row, row_cnt)
	except Exception as e :
		print(e)
		break


#to refresh the browser
# driver.refresh()
#to close the browser
# driver.close()
