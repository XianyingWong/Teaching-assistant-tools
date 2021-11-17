import os
import tkinter as tk
import tkinter.font as tkFont
import docx
import pdfplumber
from shutil import move
import openpyxl
import win32clipboard as wc
import win32con

#-----------------------------------图形界面---------------------------------------------
window = tk.Tk()
window.title('改简历')
window.geometry('180x340')

fontStyle = tkFont.Font(family="Lucida Grande", size=20)
上一个 = tk.Label(window,text='p_aname',fg='red',font=fontStyle)
上一个.pack()

tk.Label(window,text="简历",bg='lightblue').pack()
comment=tk.IntVar()
comment1=tk.IntVar()
缺少个人基本信息=tk.Checkbutton(window,text='缺少个人基本信息',variable=comment, onvalue=1,offvalue=0)
缺少个人基本信息.pack()
排版有待提高=tk.Checkbutton(window,text='排版有待提高',variable=comment1, onvalue=1,offvalue=0)
排版有待提高.pack()

tk.Label(window,text="总分",bg='lightblue').pack()
expScore=tk.IntVar()
expScore1=tk.IntVar()
expScore2=tk.IntVar()
expScore3=tk.IntVar()
expScore4=tk.IntVar()
expScore5=tk.IntVar()
sub=tk.Checkbutton(window,text='-5',variable=expScore, onvalue=1,offvalue=0)
sub.pack()
sub1=tk.Checkbutton(window,text='-5',variable=expScore1, onvalue=1,offvalue=0)
sub1.pack()
sub2=tk.Checkbutton(window,text='-3',variable=expScore2, onvalue=1,offvalue=0)
sub2.pack()
sub3=tk.Checkbutton(window,text='-2',variable=expScore3, onvalue=1,offvalue=0)
sub3.pack()
sub4=tk.Checkbutton(window,text='-1',variable=expScore4, onvalue=1,offvalue=0)
sub4.pack()
sub5=tk.Checkbutton(window,text='-15',variable=expScore5, onvalue=1,offvalue=0)
sub5.pack()

#-----------------------------------读取文件---------------------------------------------

set = os.listdir(r'C:\Users\86157\Desktop\个人简历')
if 'desktop.ini' in set:
    set.remove('desktop.ini')
for i in set:
    if i.startswith('~$'):
        set.remove(i)
setIter = iter(set)
count = 16 # stop
for i in range(0,count):
    next(setIter)
def open():
    docStr=next(setIter)
    os.startfile("C:\\Users\\86157\\Desktop\\个人简历\\" + docStr)
    global count
    print(count)
    count+=1

def getCopyText():
    wc.OpenClipboard()
    copy_text = wc.GetClipboardData(win32con.CF_TEXT)
    wc.CloseClipboard()
    return copy_text

workbook = openpyxl.load_workbook(r'C:\Users\86157\Desktop\大学计算机－实验三周四-空白.xlsx')
sheet = workbook.worksheets[0]
anameT=''
def nextone():
    aname=getCopyText().decode('GB2312')
    aname.strip()
    aname = aname.replace(' ', '')
    print(aname)
    global anameT
    if anameT==aname:
        button['bg']='red'
    else:
        button['bg'] = 'SystemButtonFace'
    anameT=aname
    上一个['text']='上个：'+anameT
    for i in range(2,len(sheet['B'])+1):
        if aname==sheet['B'+str(i)].value:
            sum = 15
            if expScore.get()==1:
                sum-=5
            if expScore1.get()==1:
                sum-=5
            if expScore2.get()==1:
                sum-=3
            if expScore3.get()==1:
                sum-=2
            if expScore4.get()==1:
                sum-=1
            if expScore5.get()==1:
                sum-=15
            sheet['I' + str(i)] = sum

            comStr=''
            if comment.get()==1:
                comStr+='缺少个人基本信息，'
            if comment1.get()==1:
                comStr+='排版有待提高，'
            sheet['J' + str(i)] =comStr

            workbook.save(r'C:\Users\86157\Desktop\大学计算机－实验三周四-空白.xlsx')
            break
    open()
    缺少个人基本信息.deselect()
    排版有待提高.deselect()
    sub.deselect()
    sub1.deselect()
    sub2.deselect()
    sub3.deselect()
    sub4.deselect()
    sub5.deselect()

button = tk.Button(window, text='下一个', width=15, height=2,command=nextone)
button.pack()

open()
window.mainloop()