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
# from PIL import Image
# import urllib.request
import sys
# import pytesseract

xlrd.xlsx.ensure_elementtree_imported(False, None)
xlrd.xlsx.Element_has_iter = True

driver = webdriver.Chrome(executable_path="./chromedriver")
driver.maximize_window()
driver.get('https://www.google.com')

file_path = "data.xlsx"
wb = xlrd.open_workbook(file_path)
sheet = wb.sheet_by_index(0)


def updateFunc(row, tab_id):
	try:
		tab_id = "tab" + str(tab_id)
		driver.execute_script("window.open('https://zfrmz.in/xnfA3LNxIw1g6qO8eyAC', '{tab_id}');".format(tab_id=tab_id))
		driver.switch_to.window(tab_id)	
		# time.sleep(1)

		enrollment_no = int(row[1])
		id_no = int(row[2])
		status = row[3]
		digital_key_no = int(row[4])
		key_id = int(row[5])
		date = row[6]
		invalidation_req_id = int(row[7])
		name = row[8]

		first_name = name.split()[0]
		second_name = name.split()[-1]

		date_obj = xlrd.xldate.xldate_as_datetime(date, wb.datemode)
		date_str = date_obj.strftime("%d-%b-%Y")


		driver.find_element(By.XPATH, '//*[@id="Checkbox_2"]').click()
		# time.sleep(1)

		driver.find_element(By.XPATH, '//*[@id="formAccess"]/div[1]/div/div[2]/button').click()
		# time.sleep(1)

		driver.find_element(By.XPATH, '//*[@id="formAccess"]/div[1]/div[2]/div[2]/button').click()
		# time.sleep(1)

		driver.find_element(By.XPATH, '//*[@id="Number-li"]/div[1]/span[1]/input').send_keys(enrollment_no)
		button = driver.find_element(By.XPATH, '/html/body/div[2]/div[1]/div[2]/div/form/div[2]/ul[4]/li/div[1]/div[2]/div[2]/button')
		# driver.implicitly_wait(5)
		ActionChains(driver).move_to_element(button).click(button).perform()
		# time.sleep(1)

		driver.find_element(By.XPATH, '//*[@id="Number1-li"]/div[1]/span[1]/input').send_keys(id_no)
		driver.find_element(By.XPATH, '//*[@id="Dropdown-li"]/div[1]/div[1]/select').send_keys(status)
		button = driver.find_element(By.XPATH, '/html/body/div[2]/div[1]/div[2]/div/form/div[2]/ul[5]/li/div[1]/div[2]/div[2]/button')
		# driver.implicitly_wait(5)
		ActionChains(driver).move_to_element(button).click(button).perform()
		# time.sleep(1)

		driver.find_element(By.XPATH, '//*[@id="Number2-li"]/div[1]/span[1]/input').send_keys(digital_key_no)
		button = driver.find_element(By.XPATH, '/html/body/div[2]/div[1]/div[2]/div/form/div[2]/ul[6]/li/div[1]/div[2]/div[2]/button')
		# driver.implicitly_wait(5)
		ActionChains(driver).move_to_element(button).click(button).perform()
		# time.sleep(1)

		driver.find_element(By.XPATH, '//*[@id="Number3-li"]/div[1]/span[1]/input').send_keys(key_id)
		button = driver.find_element(By.XPATH, '/html/body/div[2]/div[1]/div[2]/div/form/div[2]/ul[7]/li/div[1]/div[2]/div[2]/button')
		# driver.implicitly_wait(5)
		ActionChains(driver).move_to_element(button).click(button).perform()
		# time.sleep(1)

		driver.find_element(By.XPATH, '//*[@id="Date-date"]').send_keys(date_str)
		driver.find_element(By.XPATH, '//*[@id="Name-li"]/div[1]/div/span[1]/input').send_keys(first_name)
		driver.find_element(By.XPATH, '//*[@id="Name-li"]/div[1]/div/span[2]/input').send_keys(second_name)
		driver.find_element(By.XPATH, '//*[@id="Number4-li"]/div[1]/span[1]/input').send_keys(invalidation_req_id)

		# img = driver.find_element(By.XPATH, '/html/body/div[2]/div[1]/div[2]/div/form/div[2]/div[1]/ul[7]/li[4]/div[1]/div/div[2]/img')
		# src = img.get_attribute('src')

		# image_name = "captcha_" + tab_id + ".png"
		# urllib.request.urlretrieve(src, image_name)

		# image_str = pytesseract.image_to_string(Image.open(image_name))
		# print(image_str)

		previous_url = driver.current_url
		while driver.current_url == previous_url:
			time.sleep(1)

		span_element = driver.find_element(By.XPATH, '/html/body/div[2]/div[1]/div[2]/div/div/div[1]/span/label/div[2]/span')
		file_id = span_element.text.split(' ')[-1]

		write_file = open("id_file.txt", 'a')
		write_file.write('"' + file_id + '"\n')
		write_file.close()
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
		for j in range(0, 9):
			row.append(sheet.cell(i, j).value)
		updateFunc(row, row_cnt)
	except Exception as e :
		print(e)
		break


#to refresh the browser
# driver.refresh()
#to close the browser
# driver.close()
