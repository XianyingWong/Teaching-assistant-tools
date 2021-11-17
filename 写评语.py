import openpyxl

workbook = openpyxl.load_workbook(r'C:\Users\86157\Desktop\大学计算机－实验三周四-空白.xlsx')
sheet = workbook.worksheets[0]
for i in range(2, len(sheet['B']) + 1):
    if sheet['R' + str(i)].value == 1:
        if sheet['S' + str(i)].value >= 94:
            sheet['P' + str(i)] = '很认真地完成了实验的各项要求，继续努力！'
        elif sheet['S' + str(i)].value >= 80:
            sheet['P' + str(i)] ='认真按照要求完成，可惜实验：'+str(sheet['D' + str(i)].value)+'简历：'+str(sheet['J' + str(i)].value)+'论文：'+str(sheet['L' + str(i)].value)+'继续加油！'
        elif sheet['S' + str(i)].value >= 70:
            sheet['P' + str(i)] ='基本按照要求完成，可惜实验：'+str(sheet['D' + str(i)].value)+'简历：'+str(sheet['J' + str(i)].value)+'论文：'+str(sheet['L' + str(i)].value)+'继续加油！'
        else:
            sheet['P' + str(i)] ='各部分有待提高，实验：'+str(sheet['D' + str(i)].value)+'简历：'+str(sheet['J' + str(i)].value)+'论文：'+str(sheet['L' + str(i)].value)+'继续加油！'
workbook.save(r'C:\Users\86157\Desktop\大学计算机－实验三周四-空白.xlsx')

# 很认真地完成了实验的各项要求，继续努力！
# 认真按照要求完成，可惜继续加油！
# 基本按照要求完成，可惜已经很不错了，继续加油！
# 各部分有待提高，继续加油！