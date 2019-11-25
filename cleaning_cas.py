# coding: utf-8

import pandas as pd  
from pandas import read_excel
import csv


my_sheet = 'Chemical'
file_name = 'Wang_Lab_Inventory_Haibin.xlsx'
df = read_excel(file_name, sheet_name = my_sheet)
data = df.loc[:, ['Item Name *','CAS #','Serial Number']]


my_sheet_obtained = 'Reagent_Name_CAS'
file_name_obtained = 'Reagent_Name_CAS.csv'
CSV_FILE_cleaned = 'Reagent_Name_CAS_cleaned.csv'
colnames=['Item Name *','CAS #','Serial Number'] 
df_obtained = pd.read_csv(file_name_obtained, names=colnames, header=None)
data_obtained = df_obtained[:]




def Save_Reagent_Name_CAS():
	with open(CSV_FILE_cleaned, 'a') as csvFile:
		writer = csv.writer(csvFile)
		writer.writerows(csvData_list)
		print("Saved CSV data into cleaned file!: " + str(list(cas_num)[0]))
	csvFile.close()


for cas_num in data.iterrows():
	if type(list(cas_num)[1].values[:][1]) == str:
		csvData_list = [[list(cas_num)[1].values[:][0],list(cas_num)[1].values[:][1],list(cas_num)[1].values[:][2]]]
		Save_Reagent_Name_CAS()
		print('Already_1')
		continue
	else:
		csvData_list = [[list(cas_num)[1].values[:][0],''.join(data_obtained.loc[data_obtained['Serial Number']==list(cas_num)[1].values[:][2],'CAS #'].values[:]),list(cas_num)[1].values[:][2]]]
		Save_Reagent_Name_CAS()
		print('CAS_newly_added!')
		