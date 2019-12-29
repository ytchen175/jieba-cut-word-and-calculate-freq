import jieba.posseg as pseg
from openpyxl import load_workbook
import xlsxwriter
txt=open("new.txt",encoding="utf-8").read()
words = pseg.cut(txt)
items=[]

workbook=xlsxwriter.Workbook('cutwordsfreq.xlsx')
worksheet = workbook.add_worksheet("sheet1")
worksheet.write(0, 0, "字詞")
worksheet.write(0, 1, "屬性")
worksheet.write(0, 2, "次數")

turn=0
for word, flag in words:
 print('目前正在跑第'+str(turn)+'次')
 if len(items) == 0:
  items.append([word, flag, 1])
 else:
  location = -1
  for i in range(len(items)):
   if(word == items[i][0] and flag == items[i][1]):
    location = i
    break
  if(location == -1):
   items.append([word, flag, 1])
  else:
   items[location][2] = items[location][2] + 1
 turn=turn+1
 
 
for i in range(len(items)):
 worksheet.write((i+1), 0, items[i][0])
 worksheet.write((i+1), 1, items[i][1])
 worksheet.write((i+1), 2, items[i][2])
workbook.close()