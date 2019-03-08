import openpyxl
import json

cage_wb = openpyxl.load_workbook('CageCodelist.xlsx')
cage_ws = cage_wb.active
cage_max_row = cage_ws.max_row

cage_dict = {}

for i in range (1, cage_max_row+1):
	if cage_ws['D' + str(i)].value and cage_ws['B' + str(i)].value:
		dict_entry = {'email': cage_ws['D'+str(i)].value}
		if cage_ws['A' + str(i)].value is None:
			dict_entry['options'] = ""
		else:
			dict_entry['options'] = cage_ws['A' + str(i)].value
			
		cage_dict[cage_ws['B' + str(i)].value] = dict_entry
		
json = json.dumps(cage_dict)
f = open("cage_dict.json","w")
f.write(json)
f.close()