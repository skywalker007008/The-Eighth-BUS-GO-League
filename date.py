import xlrd

def getPath() :
	pydir = sys.argv[0]
	rootdir = pydir[:(len(pydir) - 11)]
	return rootdir

def readSheetList(path) :
	workbook = xlrd.open_workbook(path)
	sheet_1 = workbook.sheet_by_index(1)
	sheet_2 = workbook.sheet_by_index(2)
	sheet_3 = workbook.sheet_by_index(3)
	sheet = []
	sheet.append(sheet_1)
	sheet.append(sheet_2)
	sheet.append(sheet_3)
	return sheet

def readOneMatch(sheet, matchOrder, pathout) :
	startLine = (matchOrder - 1) * 2
	readLineNum = 2
	firstRow = 2
	lastRow = 6

	for i in range(firstRow - 1, lastRow) :
		row = sheet.row_values(i)
		pathout.write("+ ")
		team1 = row[startLine]
		pathout.write(team1 + "&emsp;vs&emsp;")
		team2 = row[startLine + 1]
		pathout.write(team2 + "\n")

def readMatchList(sheet, pathout) :
	for i in range(0, 9) :
		pathout.write("**" + sheet.cell_value(0, i * 2) + "**\n")
		readOneMatch(sheet, i, pathout)
		pathout.write("\n")

def main() :
	path = "联赛赛程.xls"
	sheetList = readSheetList(path)
	pathout1 = open("联赛赛程-甲组.md", "w+", encoding = 'UTF-8')
	pathout2 = open("联赛赛程-乙组.md", "w+", encoding = 'UTF-8')
	pathout3 = open("联赛赛程-丙组.md", "w+", encoding = 'UTF-8')
	pathout1.write("# **甲组**\n")
	readMatchList(sheetList[0], pathout1)
	pathout2.write("# **乙组**\n")
	readMatchList(sheetList[1], pathout2)
	pathout3.write("# **丙组**\n")
	readMatchList(sheetList[2], pathout3)
	pathout1.close()
	pathout2.close()
	pathout3.close()

if __name__ == '__main__':
	main()