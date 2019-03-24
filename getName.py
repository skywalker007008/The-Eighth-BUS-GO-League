import xlrd
import os
import sys

def getPath() :
	pydir = sys.argv[0]
	rootdir = pydir[:(len(pydir) - 11)]
	return rootdir

def getList(rootdir) :
	list = os.listdir(rootdir)
	print(list)
	list1 = []
	list2 = []
	list3 = []
	list4 = []
	listList = []
	for file in list :
		if file.endswith(".xls") :
			list1.append(file)
		elif file.endswith(".xlsx") :
			list2.append(file)
		elif file.endswith(".py") :
			list3.append(file)
		else :
			list4.append(file)
	listList.append(list1)
	listList.append(list2)
	listList.append(list3)
	listList.append(list4)
	return listList

def printList(listList, number) :
	for i in range(0, number) :
		print(listList[i])

def genTeamData(pathin, pathout) :
	#Get sheet
	workbook = xlrd.open_workbook(pathin)
	sheet = workbook.sheet_by_index(0)
	#pre-data
	teamName = sheet.cell_value(0, 1)
	leaderName = sheet.cell_value(2, 1)
	leaderCell = sheet.cell_value(2, 3)
	leaderQQ = sheet.cell_value(2, 8)
	leaderEmail = sheet.cell_value(4, 1)
	leaderWeChat = sheet.cell_value(4, 6)
	pathout.write("### **" + teamName + "**" + "\n\n")
	pathout.write("+ 领队：&emsp;" + leaderName + "\n")
	pathout.write("+ 微信：&emsp;" + leaderWeChat + "\n")
	pathout.write("+ 电话：&emsp;" + str(int(leaderCell)) + "\n")
	pathout.write("+ 邮箱：&emsp;" + str(leaderEmail) + "\n")
	pathout.write("+ QQ&ensp;：&emsp;" + str(leaderQQ) + "\n")
	#MemberData
	pathout.write("#### **队员名单——**\n\n")
	pathout.write("姓名|性别|年级|棋力\n")
	pathout.write(":--:|:-:|:--:|:--:\n")
	for i in range(7, sheet.nrows) :
		rows = sheet.row_values(i)
		if len(rows[0]) == 0 :
			continue
		pathout.write(str(rows[0]) + "|" + str(rows[2]) + "|" + str(rows[5]) + "|" + str(rows[8]) + "\n")

	pathout.write("\n")

def main() :
	path = getPath()
	listList = getList(path)
	printList(listList, 4)
	out = open(".\\队员信息\\nameList.md", "w+", encoding = "utf-8")
	genTeamData((listList[1])[0], out)

if __name__ == '__main__' :
	main()
