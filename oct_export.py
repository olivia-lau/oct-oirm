import csv
import datetime
import os
import shutil
import subprocess
import time as t
from datetime import *
from datetime import datetime as dt
import numpy as np
import pandas as pd
import pyautogui
# from openpyxl import *
# import xlrd
from tkinter import *
import tkinter as tk
from tkinter import filedialog

df_main = pd.read_excel('oct_var_match.xlsx')
key_match = pd.read_excel('key_match.xlsx')

def count_dir(dir_path, file_type):
	n_files = 0
	for files in os.listdir(dir_path):
		len_file_type = len(file_type) + 1
		if files[ -len_file_type: ] == "." + file_type:
			n_files += 1
	return n_files

def xml2xlsx(filesource):
	total_xml_files = count_dir(filesource, "xml")
	xml_index_completion = 0
	original_xml_files = []
	for original_xml_file in os.listdir(filesource):
		original_xml_files.append(original_xml_file)
	converted_xml_files = []
	if len(converted_xml_files) < len(original_xml_files):
		for  xml_index, xml_file in enumerate (os.listdir(filesource)):
				converted_file_name =  "Copy of " + xml_file[:-3] + "xlsx"
				converted_file_path = filesource +"/" + converted_file_name
				if xml_file[ -3: ] == 'xml' and not os.path.exists(converted_file_path):
						xml_index_completion += 1
						print(xml_file)
						print("XML ---> XLSX: " + str(xml_index_completion) + "/ " + str(total_xml_files) + " files" )
						subprocess.call(['open', filesource + "/" + xml_file])
						t.sleep(12)
						pyautogui.press('down')
						t.sleep(3)
						while not os.path.exists(converted_file_path):
							pyautogui.hotkey('command' , 'shift' , 's')
							t.sleep(3)
							pyautogui.press('enter')
							t.sleep(3)
						pyautogui.hotkey('command' , 'q')
						t.sleep(3)
	try:
		os.mkdir(filesource +  "/2. Converted XLSX/")
	except FileExistsError:
		pass
	for file in os.listdir(filesource):
		if "Copy of" in file:
			shutil.move(filesource + '/' + file , filesource + '/2. Converted XLSX/' + file)
	try:
		os.mkdir(filesource +  "/1. Original XML/")
	except FileExistsError:
		pass
	for file in os.listdir(filesource):
		if ".xml" in file:
			shutil.move(filesource + '/' + file , filesource + '/1. Original XML/' + file)
			
			

def xlsx2csv(filesource):
	dir_path = filesource + '/2. Converted XLSX'
	dir_count_xlsx = count_dir(dir_path, "xlsx")
	file_num = 0
	for files in os.listdir(dir_path):
		if files[ -5: ] == ".xlsx":
			file_name = files[ :-5 ]
			print(dir_path + "/" + files)
			df_source = pd.read_excel(dir_path + "/" + files)
			df_source = df_source.transpose()
			left_join = pd.merge(df_main , df_source , left_on='var_lookup' , right_on=0 , how='left')
			new = left_join.transpose().to_numpy()
			b = pd.DataFrame(new).drop_duplicates()
			od_a = [ [ 8 , 14 ] , [ 20 , 34 ] , [ 48 , 57 ] ]
			od_val = [ ]
			for n in od_a:
				for x in range(n[ 0 ] , n[ 1 ]):
					od_val.append(x)
			os_val = [ ]
			for y in range(8 , 66):
				if y not in od_val:
					os_val.append(y)
			new_1 = b.to_numpy()
			for m in new_1:
				if m[ 7 ] == "OS":
					for n in od_val:
						m[ n ] = pd.NA
				if m[ 7 ] == "OD":
					for n in os_val:
						m[ n ] = pd.NA
				if str(m[ 8 ]) + "--" + str(m[ 9 ]) + "--" + str(m[ 10 ]) == "nan--" + str(m[ 9 ]) + "--nan":
					m[ 9 ] = pd.NA
				if str(m[ 14 ]) + "--" + str(m[ 15 ]) + "--" + str(m[ 16 ]) == "nan--" + str(m[ 15 ]) + "--nan":
					m[ 15 ] = pd.NA
			
			df = b.groupby([ 0 , 6 ]).first().reset_index()
			new_2 = df.to_numpy()
			
			for y in new_2:
				y[ 21 ] = str(y[ 21 ]) + "%% " + str(y[ 22 ])
				y[ 35 ] = str(y[ 35 ]) + "%% " + str(y[ 36 ])
				# print(type(y[4]))
				# if type(y[ 4 ]) is dt.datetime:
				# 	y[ 4 ] = y[ 4 ].strftime("%Y-%m-%d")
				# 	# print(y[4])
				# 	y[ 6 ] = y[ 6 ].strftime("%Y-%m-%d")
				#
			new_df_a = np.delete(new_2 , 36 , 1)
			new_df_b = np.delete(new_df_a , 22 , 1)
			new_df_c = np.delete(new_df_b , 7 , 1)
			new_3 = np.delete(new_df_c , 1 , 1)
			df1 = pd.DataFrame(new_3)
			df1 = df1.transpose()
			df1 = df1.rename(columns={'Unnamed: 0': 'nom' , '0': 0})
			# print(df1)
			# print(df1.keys())
			df1.transpose().to_csv(dir_path + "/" + file_name + '.csv')
			file_num += 1
			print("XLSX ---> CSV: " + str(file_num) + "/" + str(dir_count_xlsx) + " files")
			# shutil.move(dir_path + '/' + file_name + '.xlsx' , dir_path + file_name + '.xlsx')
	try:
		os.mkdir(filesource+ "/3. Converted CSV/")
	except FileExistsError:
		pass
	for file in os.listdir(dir_path):
		if ".csv" in file:
			shutil.move(dir_path + "/" + file , filesource + '/3. Converted CSV/' + file)


# xlsx_to_csv()

col_title = ['0', '/PATIENT/PERSON_NAME/ALPHABETIC_COMPONENT/LAST_NAME', '/PATIENT/PERSON_NAME/ALPHABETIC_COMPONENT/FIRST_NAME', '/PATIENT/PATIENT_ID', '/PATIENT/BIRTH_DATE', '/PATIENT/GENDER', '/PATIENT/VISITS/STUDY/VISIT_DATE', '/PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/OPTIC_DISC/AVERAGETHICKNESS', '/PATIENT/VISITS/STUDY/SERIES/SCAN/CZMOCTSETTINGS/SIGNALSTRENGTH', '/PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/OPTIC_DISC/QUADRANT_S', '/PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/OPTIC_DISC/QUADRANT_N', '/PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/OPTIC_DISC/QUADRANT_I', '/PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/OPTIC_DISC/QUADRANT_T', '/PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/OPTIC_DISC/AVERAGETHICKNESS', '/PATIENT/VISITS/STUDY/SERIES/SCAN/CZMOCTSETTINGS/SIGNALSTRENGTH', '/PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/OPTIC_DISC/QUADRANT_S', '/PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/OPTIC_DISC/QUADRANT_N', '/PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/OPTIC_DISC/QUADRANT_I', '/PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/OPTIC_DISC/QUADRANT_T', '/PATIENT/VISITS/STUDY/SERIES/SCAN/CZMOCTSETTINGS/SIGNALSTRENGTH', '/PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/GCA/FOVEA_LOCATION/X, /PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/GCA/FOVEA_LOCATION/Y', '/PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/MTA/Z_OUTERSUPERIOR', '/PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/MTA/Z_OUTERRIGHT', '/PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/MTA/Z_OUTERINFERIOR', '/PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/MTA/Z_OUTERLEFT', '/PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/MTA/Z_INNERSUPERIOR', '/PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/MTA/Z_INNERRIGHT', '/PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/MTA/Z_INNERINFERIOR', '/PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/MTA/Z_INNERLEFT', '/PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/MTA/CENTRALSUBFIELDTHICKNESS/ILMRPE', '/PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/MTA/CUBEVOLUME/ILMRPE', '/PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/MTA/CUBEAVGTHICKNESS/ILMRPE', '/PATIENT/VISITS/STUDY/SERIES/SCAN/CZMOCTSETTINGS/SIGNALSTRENGTH', '/PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/GCA/FOVEA_LOCATION/X, /PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/GCA/FOVEA_LOCATION/Y', '/PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/MTA/Z_OUTERSUPERIOR', '/PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/MTA/Z_OUTERLEFT', '/PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/MTA/Z_OUTERINFERIOR', '/PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/MTA/Z_OUTERRIGHT', '/PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/MTA/Z_INNERSUPERIOR', '/PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/MTA/Z_INNERLEFT', '/PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/MTA/Z_INNERINFERIOR', '/PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/MTA/Z_INNERRIGHT', '/PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/MTA/CENTRALSUBFIELDTHICKNESS/ILMRPE', '/PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/MTA/CUBEVOLUME/ILMRPE', '/PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/MTA/CUBEAVGTHICKNESS/ILMRPE', '/PATIENT/VISITS/STUDY/SERIES/SCAN/CZMOCTSETTINGS/SIGNALSTRENGTH', '/PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/GCA/GC_SUP', '/PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/GCA/GC_TEMPSUP', '/PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/GCA/GC_TEMPINF', '/PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/GCA/GC_INF', '/PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/GCA/GC_NASINF', '/PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/GCA/GC_NASSUP', '/PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/GCA/GC_AVERAGE', '/PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/GCA/GC_MINIMUM', '/PATIENT/VISITS/STUDY/SERIES/SCAN/CZMOCTSETTINGS/SIGNALSTRENGTH', '/PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/GCA/GC_SUP', '/PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/GCA/GC_TEMPSUP', '/PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/GCA/GC_TEMPINF', '/PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/GCA/GC_INF', '/PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/GCA/GC_NASINF', '/PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/GCA/GC_NASSUP', '/PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/GCA/GC_AVERAGE', '/PATIENT/VISITS/STUDY/SERIES/SCAN/ANALYSIS/GCA/GC_MINIMUM']

def combine_csv_files(filesource):
	fol_path = filesource + "/3. Converted CSV"
	##-----------------------------------------------------------------------------------------------##
	## READING  NEW CSV FILES & COMBINING INTO FORMATTED CSV##
	new_list = [ ]
	for file in os.listdir(fol_path):
		if fol_path[ -1: ] == "/" and file[-4:] ==".csv":
			csv_file_path = fol_path + file
		elif  file[-4:] ==".csv":
			csv_file_path = fol_path + "/" + file
		with open(csv_file_path , newline='') as csvfile:
			spamreader = csv.reader(csvfile, delimiter=',')
			m = 0
			for row in spamreader:
				if m > 1:
					new_list.append(row)
				# print(m)
				# print(', '.join(row))
				m += 1
	df = pd.DataFrame(new_list , columns=col_title)
	try:
		os.mkdir(filesource +  "/4. Output/")
	except FileExistsError:
		pass
	combine_file_path = filesource + "/4. Output/Output File.csv"
	combine_file_path_formatted = filesource + "/4. Output/Output File Formatted.csv"
	df.to_csv(combine_file_path)
	output_file = combine_file_path
	new_csv_list = [ ]
	with open(output_file) as csvfile: ## FORMATS CSV
		output_csv_original = csv.reader(csvfile , delimiter=',')
		for row in output_csv_original:
			new_row = [ ]
			cell_index = 0
			for x in row:
				cell_index += 1
				if cell_index <= 2:
					pass
				elif x == "" or x == "None%% None":
					new_row.append("6666")
				elif "%%" in x:
					new_x = x.replace("%%" , ",")
					# new_row.append(x)
					new_row.append(new_x)
				elif "00:00:00" in  x:
					new_x = x.replace(" 00:00:00" , "")
					new_row.append(new_x)
				else:
					new_row.append(x)
			new_csv_list.append(new_row)
	
	
	with open(combine_file_path_formatted , mode='w') as f: ##WRITES FORMATTED CSV
		writer = csv.writer(f)
		for row in new_csv_list:
			writer.writerow(row)
	
	##-----------------------------------------------------------------------------------------------##

	##ADDING LAST OCT DATASET TO DICT W/O DUPLICATES##

	dict_oct_data = {}

	output_path = "/Users/olivia/oct-export/Combined Outputs/"
	latest_output_file_num = ""
	for file in os.listdir(output_path):
		file_number = file.split("___")[0]
		if latest_output_file_num == "":
			latest_output_file_num += file_number
		elif file_number > latest_output_file_num:
			latest_output_file_num = file_number
			
	existing_file_path = output_path + str(latest_output_file_num) + "___Last_Combined_Output.csv"
	
	with open(existing_file_path) as csvfile:
		existing_csv = csv.reader(csvfile , delimiter=',')
		for row in existing_csv:
			if row[2] in dict_oct_data:
				dict_oct_data[row[2]].append(row)
			else:
				dict_oct_data[row[2]] = [row]

	n = dict_oct_data.keys()
	# for x in n:
	# 	print(x)
	num_existing_mrn = len(n)
	##-- ---------------------------------------------------------------------------------------------##
	## FORMATTED CSV TO DICT##
	#
	
	for row in new_csv_list:
		# print(row)
		if row[ 2 ] in dict_oct_data and row not in dict_oct_data[ row[ 2 ] ]:
			# print(row[2])
			dict_oct_data[ row[ 2 ] ].append(row)
		else:
			dict_oct_data[ row[ 2 ] ] = [ row ]
	
	n = dict_oct_data.keys()
	# for x in n:
	# 	print(x)
	num_total_mrn = len(n)
	new_mrn_added = num_total_mrn - num_existing_mrn - 1
	print("New MRNs added = " + str(new_mrn_added))
	##-----------------------------------------------------------------------------------------------##

	##-----------------------------------------------------------------------------------------------##
	##PRINT DICT##
	final_csv = []
	csv_header = []
	mrn_count = 0
	for mrn, rows in dict_oct_data.items():
		# print(mrn)
		# print(len(rows))
		for row in rows:
			if row[ 0 ] == "Last Name":
				csv_header.append(row)
			elif row[ 0] != "/PATIENT/PERSON_NAME/ALPHABETIC_COMPONENT/LAST_NAME" and  row[ 0] != "":
				final_csv.append(row)
		# print('---------------')
		mrn_count += 1
	print("Total Unique MRNs  = " + str(mrn_count))
	newest_output_file = output_path + str(int(latest_output_file_num) + 1) + "___Last_Combined_Output.csv"
	with open(newest_output_file , mode='w') as f: ##WRITES FORMATTED CSV
		writer = csv.writer(f)
		# writer.writerow(csv_header[0])
		for row in final_csv:
			# print(row)
			writer.writerow(row)
		
		##-- ---------------------------------------------------------------------------------------------##
		## FORMATTED CSV TO DICT##
		#
		
		for row in new_csv_list:
			# print(row)
			if row[ 2 ] in dict_oct_data and row not in dict_oct_data[ row[ 2 ] ]:
				# print(row[2])
				dict_oct_data[ row[ 2 ] ].append(row)
			else:
				dict_oct_data[ row[ 2 ] ] = [ row ]
		
		n = dict_oct_data.keys()
		# for x in n:
		# 	print(x)
		num_total_mrn = len(n)
	#-----------------------------------------------------------------------------------------------##
	
	newest_output_file_sorted = output_path + str(int(latest_output_file_num) + 1) + "___Last_Combined_Output_SORTED.csv"
	newest_output_file_df = pd.read_csv(newest_output_file)
	new_df = newest_output_file_df.sort_values(by=['MRN'])
	new_df.to_csv(newest_output_file_sorted)

# combine_csv_files("/Users/olivia/Downloads/Working OCT XML/Working OCT CSV")

root = tk.Tk()  ## attach process
root.withdraw()  ## required for file selection
xml_directory_path= filedialog.askdirectory()
xml2xlsx(xml_directory_path)
print("----------------------------------------------------------------------")
xlsx2csv(xml_directory_path)
print("----------------------------------------------------------------------")
combine_csv_files(xml_directory_path)
print("----------------------------------------------------------------------")
