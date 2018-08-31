import gspread
from oauth2client.service_account import ServiceAccountCredentials
#TODO
#add x when list taken
#add more print stements as it runs
#add a counter to print out stats
#inster better fromater
 
print("Welcome to the J.O.S.E., or the SHPE AI Secretary! Please sit back and let me do all of Jose's work!\n")

# use creds to create a client to interact with the Google Drive API
scope = ['https://spreadsheets.google.com/feeds']
creds = ServiceAccountCredentials.from_json_keyfile_name('shpe-officer-7f041aaa1b91.json', scope)
client = gspread.authorize(creds)
 
# Find a workbook by name and open the first sheet
# Make sure you use the right name here.
jose = client.open("The J.O.S.E. (17-18)")
sheetEA = jose.worksheet("Event Attend. F17-S18")
sheetPM = jose.worksheet("Paid Members List")
sheetPS = jose.worksheet("Point System")

print("Opening the sheets!\n")

print("Going through the Attendance Sheet\n")

listOfList = []

for col in range(3,sheetEA.col_count):
	list1 = []
	if(sheetEA.cell(4,col).value == "o"):
		print("Got list at: %d" % col)
		list1 = sheetEA.col_values(col)
		del list1[0:3]
		list1 = list(set(list1))
		list1 = [eid.lower() for eid in list1]
		list1.remove('')
		list1.insert(0,sheetEA.cell(2,col).value)
		list1.insert(1, sheetEA.cell(3,col).value)
		listOfList.append(list1)
		sheetEA.update_cell(4,col,"x")
print("Printing Event Attendence")
print(listOfList) 
print("*******************************************************")

print("Creating EID to Name Dictionary!")
SHPEmap = dict()

for row in range(2,sheetPM.row_count+1):
	eid = sheetPM.cell(row,4).value 
	SHPEmap[eid] = sheetPM.cell(row,2).value
	print("Matched EID: %s to the name: %s" %(eid,SHPEmap[eid]))

print("")
print("Printing paid members dict")
print(SHPEmap)

print("Adding up the points, let the magic begin...")

for listPos in range(len(listOfList)):
	print("For event %s" % listOfList[listPos][0])
	colChange = sheetPS.find(listOfList[listPos][1]).col
	for index in range(2,len(listOfList[listPos])):
		if listOfList[listPos][index] in SHPEmap:
			cellFound = sheetPS.find(SHPEmap[listOfList[listPos][index]])
			rowChange = cellFound.row
			val = sheetPS.cell(rowChange,colChange).value
			if(val != ""):
				sheetPS.update_cell(rowChange,colChange, (int(val) + 1))
			else:
				sheetPS.update_cell(rowChange,colChange, 1)
	print("Pts added for event %s" % listOfList[listPos][0])
print("Updated the point system!!!")
