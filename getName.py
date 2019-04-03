import xlrd
import os
import sys

def getPath() :
	pydir = sys.argv[0]
	rootdir = pydir[:(len(pydir) - 11)]
	finaldir = rootdir + '.\\队员信息'
	return finaldir

def groupbyBehind(rootdir, behind) :
	list = os.listdir(rootdir)
	tempList = []
	for file in list :
		if file.endswith(behind) :
			tempList.append(file)
	return tempList

def isEnd(content) :
	if str(content) == "队伍" :
		return True
	else :
		return False

def genTeamData(sheet, startRow, pathout) :
	#pre-data
	teamName = sheet.cell_value(startRow, 1)
	print(teamName)
	leaderName = sheet.cell_value(startRow + 2, 1)
	leaderCell = sheet.cell_value(startRow + 2, 3)
	leaderQQ = sheet.cell_value(startRow + 2, 8)
	leaderEmail = sheet.cell_value(startRow + 4, 1)
	leaderWeChat = sheet.cell_value(startRow + 4, 6)
	pathout.write("### **" + teamName + "**" + "\n\n")
	pathout.write("+ 领队：&emsp;" + leaderName + "\n")
	if sheet.cell_type(startRow + 4, 6) == 1 :
		pathout.write("+ 微信：&emsp;" + str(leaderWeChat) + "\n")
	elif sheet.cell_type(startRow + 4, 6) == 2 :
		pathout.write("+ 微信：&emsp;" + str(int(leaderWeChat)) + "\n")
	pathout.write("+ 电话：&emsp;" + str(int(leaderCell)) + "\n")
	pathout.write("+ 邮箱：&emsp;" + str(leaderEmail) + "\n")
	pathout.write("+ QQ&ensp;：&emsp;" + str(leaderQQ) + "\n")
	#MemberData
	pathout.write("#### **队员名单——**\n\n")
	pathout.write("姓名|性别|年级|棋力\n")
	pathout.write(":--:|:-:|:--:|:--:\n")
	for i in range(startRow + 7, sheet.nrows) :
		rows = sheet.row_values(i)
		if len(rows[0]) == 0 :
			continue
		if isEnd(rows[0]) :
			break
		pathout.write(str(rows[0]) + "|" + str(rows[2]) + "|" + str(rows[5]) + "|" + str(rows[8]) + "\n")

	pathout.write("\n")
	return i

def genGroupDataInOneSheet(pathin, pathout) :
	workbook = xlrd.open_workbook(pathin)
	sheet = workbook.sheet_by_index(0)
	row_start_num = 0
	while(1) :
		row_start_num = genTeamData(sheet, row_start_num, pathout)
		print(row_start_num)
		if row_start_num == sheet.nrows - 1 :
			return


def main() :
	filePath = ".\\队员信息\\"
	path = getPath()
	elsList = groupbyBehind(path, ".xlsx")
	out = open(".\\队员信息\\nameList.md", "w+", encoding = "utf-8")
	out1 = open(filePath + "甲组名单.md", "w+", encoding = "utf-8")
	# out2 = open(filePath + "乙组名单.md")
	# out3 = open(filePath + "丙组名单.md")
	for file in elsList :
		if file.startswith("甲组") :
			out1.write("# **甲组**\n")
			genGroupDataInOneSheet(filePath + file, out1)
			out1.close()
		# elif file.startswith("乙组") :
		# 	genGroupDataInOneSheet(file, out2)
		# 	close(out2)
		# elif file.startswith("丙组") :
		# 	genGroupDataInSheets(file, out3)
		# 	close(out3)


if __name__ == '__main__' :
	main()
