import csv
import xlsxwriter 
# 2.csv is original data
file_name="2.csv"
csvFile=open(file_name,'rt')
csvReader=csv.reader(csvFile,delimiter=",")
list_=list()
for row in csvReader:
        list_.append(row)
#1.csv is students data
file_name2="1.csv"
csvFile2=open(file_name2,'rt')
csvReader2=csv.reader(csvFile2,delimiter=",")
list1=list()
for row in csvReader2:
        list1.append(row)
list2=list()
del list1[0]
for i in list1:
        l=[]
        if(int(i[-1])>60 and int(i[-1])<80):
                k=i[0].split("-")
                l.append(k[0])
                l.append(k[1])
                l.append(k[2])
                l.append(i[-3])
                l.append(i[-2])
                l.append(i[-1])
                list2.append(l)
list3=[]
m=['Roll Number','Name','College','Join Time','Leave Time','Session Time']
list3.append(m)
for i in list2:
                if i[0] not in list3:
                                list3.append(i)
print(list3)
row = 0
col = 0
#60 to 80.xlsx is excel file name
workbook = xlsxwriter.Workbook('60to80.xlsx')
worksheet = workbook.add_worksheet("My sheet")
for roll,name,college,jointime,leavetime,sessiontime in list3:
                worksheet.write(row, col, roll)
                worksheet.write(row, col + 1, name) 
                worksheet.write(row, col + 2,college)
                worksheet.write(row, col + 3, jointime)
                worksheet.write(row, col + 4, leavetime)
                worksheet.write(row, col + 5, sessiontime)
                row=row+1
workbook.close()
