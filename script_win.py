import xlrd
import time
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from datetime import datetime
import sys

# xlrd.xlsx.ensure_elementtree_imported(False, None)
# xlrd.xlsx.Element_has_iter = True

driver = webdriver.Chrome(executable_path="./chromedriver.exe")
driver.maximize_window()
driver.get('https://www.google.com')

file_path = "data.xlsx"
wb = xlrd.open_workbook(file_path)
sheet = wb.sheet_by_index(0)

def updateFunc(row, tab_id, submitted_by, sheet_index):
	try:
		deo_login = row[3]

		submitted_by = deo_login.lower()
		if submitted_by != submitted_by:
			return 

		tab_id = "tab" + str(tab_id)
		driver.execute_script("window.open('https://zfrmz.in/zijoajAUizo4csAiGHo5', '{tab_id}');".format(tab_id=tab_id))
		driver.switch_to.window(tab_id)	
		time.sleep(2)

		digital_ekyc_no = int(row[1])
		name = row[2]
		deo_login = row[3]
		obligations = row[4]
		utr_no = int(row[5])
		captured = row[6]
		date = row[7]
		first_name = name.split()[0]
		second_name = name.split()[-1]

		if obligations == 0:
			obligations = "False"
		elif obligations == 1:
			obligations = "True"
		else:
			obligations = "Hold"

		# date_obj = xlrd.xldate.xldate_as_datetime(float(date), wb.datemode)
		date_obj = datetime.strptime(date, "%d/%m/%Y")
		date_str = date_obj.strftime("%d-%b-%Y")


		driver.find_element(By.XPATH, '//*[@id="Number-li"]/div[1]/span[1]/input').send_keys(digital_ekyc_no)
		driver.find_element(By.XPATH, '//*[@id="Name-li"]/div[1]/div/span[1]/input').send_keys(first_name)
		driver.find_element(By.XPATH, '//*[@id="Name-li"]/div[1]/div/span[2]/input').send_keys(second_name)
		driver.find_element(By.XPATH, '//*[@id="Dropdown-li"]/div[1]/div[1]/select').send_keys(deo_login)

		driver.find_element(By.XPATH, '//*[@id="Dropdown1-li"]/div[1]/div[1]/select').send_keys(obligations)
		driver.find_element(By.XPATH, '//*[@id="Number1-li"]/div[1]/span[1]/input').send_keys(utr_no)

		driver.find_element(By.XPATH, '//*[@id="Dropdown2-li"]/div[1]/div[1]/select').send_keys(captured)

		driver.find_element(By.XPATH, '//*[@id="Date-date"]').send_keys(date_str)

		previous_url = driver.current_url
		while driver.current_url == previous_url:
			time.sleep(1)

		print("Entry for index -> {index} has submitted successfully".format(
			index=sheet_index))
	except Exception as e:
		exc_type, exc_obj, exc_tb = sys.exc_info()
		print(e, exc_tb.tb_lineno)
		raise e


first_idx = int(input("Enter first index -> "))
last_idx = int(input("Enter last index -> "))
submitted_by = input("Enter your name -> ")

row_cnt = 0
for i in range(first_idx, last_idx+1):
	row_cnt += 1
	try:
		row = []
		for j in range(0, 8):
			row.append(sheet.cell(i, j).value)
		updateFunc(row, row_cnt, submitted_by, i)
	except Exception as e :
		print(e)
		break

#to refresh the browser
# driver.refresh()
#to close the browser
# driver.close()
