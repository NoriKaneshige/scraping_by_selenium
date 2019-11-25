# coding: utf-8

import pandas as pd  
from pandas import read_excel
import csv
from bs4 import BeautifulSoup
from time import sleep
import re
from selenium import webdriver
import time


"""PubChem first, then Combi, Signa etc"""

driver = webdriver.Chrome('/Users/Koitaro/Desktop/Selenium/chromedriver')


my_sheet = 'Chemical'
file_name = 'Wang_Lab_Inventory_Haibin.xlsx'
CSV_FILE = 'Reagent_Name_CAS.csv'

df = read_excel(file_name, sheet_name = my_sheet)

data = df.loc[:, ['Item Name *','CAS #','Vendor','Catalog #','Serial Number']]


def Save_Reagent_Name_CAS():
	with open(CSV_FILE, 'a') as csvFile:
		writer = csv.writer(csvFile)
		writer.writerows(csvData_list)
		print("Saved CSV data: " + str(list(cas_num)[0]))
	csvFile.close()


combi_blocks_url = 'https://www.combi-blocks.com/'
sigma_url = 'https://www.sigmaaldrich.com/united-states.html'

for cas_num in data.iterrows():
	named_tuple = time.localtime() # get struct_time
	time_string = time.strftime("%m/%d/%Y, %H:%M:%S", named_tuple)
	print(time_string)
	print('For {}'.format(list(cas_num)[1].values[:][0]))
	if type(list(cas_num)[1].values[:][1]) == str:
		csvData_list = [[list(cas_num)[1].values[:][0],list(cas_num)[1].values[:][1],list(cas_num)[1].values[:][4]]]
		Save_Reagent_Name_CAS()
		print('Already_1')

	else:
		unique_url = 'https://pubchem.ncbi.nlm.nih.gov/compound/{}#section=CAS&fullscreen=true'.format(str(re.split(r'[,\s][(\s*9R>=≥]',list(cas_num)[1].values[:][0])[0]))
		try:
			sleep(1)
			driver.get(unique_url)
			sleep(1)
			cas = driver.find_element_by_class_name('section-content-item')
			print(cas.text)
			csvData_list = [[list(cas_num)[1].values[:][0],cas.text,list(cas_num)[1].values[:][4]]]
			Save_Reagent_Name_CAS()
			print('PubChem_2')
			sleep(1)

		except Exception as inst:
			print('checking if Combi-Blocks/Sigma-Aldrich or any others')
			print('PubChem_3')
			
			try:
				if list(cas_num)[1].values[:][2] == 'Combi-Blocks':
					print('Combi here!')
					try:
						driver.get(combi_blocks_url)
						sleep(1)
						driver.find_element_by_id("search-input").send_keys(list(cas_num)[1].values[:][3])
						driver.find_element_by_xpath('//span[@class="input-group-btn"]').click()
						sleep(1)
						cas_info_table_split = driver.find_element_by_id("item").text.split('\n')
						sleep(1)
						try:
							for el_cas_info_table_split in cas_info_table_split:
								if 'CAS number' in el_cas_info_table_split:
									Combi_CAS_obtained = el_cas_info_table_split.split(' ')[2]
									# print(Combi_CAS_obtained)
									csvData_list = [[list(cas_num)[1].values[:][0],Combi_CAS_obtained,list(cas_num)[1].values[:][4]]]
									Save_Reagent_Name_CAS()
									sleep(1)
									print('Combi_4')
									driver.quit()
									driver = webdriver.Chrome('/Users/Koitaro/Desktop/Selenium/chromedriver')
									sleep(1)
								else:
									print('Failed to get CAS from Combi!')
									continue
						except Exception as inst:
							print('Failed to get CAS from Combi')
							print('Combi_5-1')
							sleep(1)

					except Exception as inst:
						print('Failed to get CAS from Combi')
						print('Combi_5-3')
						sleep(1)
					

				elif list(cas_num)[1].values[:][2] == 'Sigma-Aldrich':
					print('Sigma here!')
					try:
						driver.get(sigma_url)
						sleep(1)
						driver.find_element_by_id("Query").send_keys(list(cas_num)[1].values[:][3])
						driver.find_element_by_xpath('//input[@id="submitSearch"]').click()
						sleep(1)
						html = driver.page_source.encode('utf-8')
						soup = BeautifulSoup(html, "html.parser")
						spans = soup.find_all('span', {'class' : 'info'})
						for el_spans in spans:
							if el_spans.find('a'):
								# print(el_spans.text)
								csvData_list = [[list(cas_num)[1].values[:][0],el_spans.text,list(cas_num)[1].values[:][4]]]
								Save_Reagent_Name_CAS()
								print('Sigma_6')
								driver.quit()
								driver = webdriver.Chrome('/Users/Koitaro/Desktop/Selenium/chromedriver')
								sleep(1)
							else:
								print('Failed to get CAS from Sigma')
								continue

					except Exception as inst:
						print('Failed to get CAS from Sigma')
						print('Sigma_7-2')
						sleep(1)

				else:
					print('Failed to get CAS from any source')
					print('8')
					sleep(1)

			except Exception as inst:
					print('Failed to get CAS from any source(2)')
					print('9')
					sleep(1)
			

