import openpyxl
import json

excluded_cols = ["International", "Event_Link", "Round_Code"]

wb = openpyxl.load_workbook('MotionCorpus.xlsx')
sheet = wb.active
motions = []
key_map = {}
count = 0
for row in sheet.rows:
	motion = {}
	CAs = []
	topics = []
	for index, cell in enumerate(row):
		if count == 0:
			key_map[index] = str(cell.value)
		else:
			if "CA" in key_map[index]:
				if cell.value != None:
					CAs.append(str(cell.value))
			elif "Topic" in key_map[index]:
				if cell.value != None:
					topics.append(str(cell.value))
			elif key_map[index] in excluded_cols:
				pass
			else:
				motion[key_map[index]] = str(cell.value)
	if count > 0:
		motion["CAs"] = CAs
		motion["Topics"] = topics
		motions.append(motion)
	count += 1
print(json.dumps(motions))
