import xlwings as xw
from requests_html import HTMLSession
import tkinter as tk
from tkinter import messagebox

window = tk.Tk()
window.title('sabrina dataprocess')
window.geometry('500x300')
tk.Label(window, text='待处理起始行A：').place(x=100, y=50)
tk.Label(window, text='待处理结束行A： ').place(x=100, y=100)
tk.Label(window, text='数据库结束行B： ').place(x=100, y=150)
tk.Label(window, text='powered by zls', font=('Arial', 8)).place(x=400, y=270)
e1 = tk.Entry(window, show=None, font=('Arial', 14))
e2 = tk.Entry(window, show=None, font=('Arial', 14))
e3 = tk.Entry(window, show=None, font=('Arial', 14))
e1.place(x=200, y=50)
e2.place(x=200, y=100)
e3.place(x=200, y=150)


def inputError():
    tk.messagebox.showinfo(title='Error', message='请检查输入')


#查询过程
def clickButton1():
    #连接到excel
    book = xw.Book(r'myexcel.xlsx')
    workbook = book.sheets('Sheet1')  #连接excel文件
    workbook2 = book.sheets('Sheet2')  #连接excel文件
    #创建字典
    dbend = e3.get()
    dbname = workbook2.range('B1:B' + str(dbend)).value
    dbarea = workbook2.range('C1:C' + str(dbend)).value
    dic = dict(zip(dbname, dbarea))
    #模糊匹配
    inputrangebegin = e1.get()
    inputrangeend = e2.get()
    leng = str(int(inputrangeend) - int(inputrangebegin) + 1)
    #检查输入合法
    if (inputrangebegin > inputrangeend or inputrangebegin == 0
            or inputrangeend == 0):
        inputError()
        return
    #创建次级进度窗口
    windowProcess = tk.Toplevel()
    windowProcess.geometry('300x200')
    windowProcess.title('进度')
    l1 = tk.Label(windowProcess,
                    text='处理进度0/' + leng,
                    font=('Microsoft YaHei', 14))
    l1.place(x=100, y=50)
    count = 0
    #输入数据范围
    dataInput = workbook.range('A' + str(inputrangebegin) + ':' + 'A' +
                                str(inputrangeend)).value
    aList = []
    session = HTMLSession()

    #在百科搜索名字，在字典中匹配
    for school in dataInput:
        url = 'https://baike.baidu.com/search/word?word=' + school
        try:
            r = session.get(url)
        except:
            tk.messagebox.showinfo(title='Error', message='无连接，请检查网络')
            return
        #查找sel123的数据
        sel0 = 'body > div.body-wrapper.feature.feature_small.collegeSmall > div.feature_poster > div > div.poster-left > div.poster-top > dl > dd > h1'
        sel1 = 'body > div.body-wrapper > div.content-wrapper > div > div.main-content > dl.lemmaWgt-lemmaTitle.lemmaWgt-lemmaTitle- > dd > h1'
        sel2 = '#body_wrapper > div.searchResult > dl > dd:nth-child(2) > a'

        results = r.html.find(sel0)
        key = ''
        for result in results:
            key = result.text
        results = r.html.find(sel1)
        for result in results:
            key = result.text

        results = r.html.find(sel2)
        for result in results:
            mylink = list(result.absolute_links)[0]
            r2 = session.get(mylink)
            results2 = r2.html.find(sel0)
            for result in results2:
                key = result.text
            results2 = r2.html.find(sel1)
            for result in results2:
                key = result.text
        #更新处理进度
        count += 1
        l1["text"] = '处理进度' + str(count) + '/' + leng
        l1.update()
        #将结果加入list
        if dic.__contains__(key):
            aList.append(dic[key])
        elif dic.__contains__(school):
            aList.append(dic[school])
        else:
            aList.append('null')
    #输出结果
    workbook.range('B' +
                    str(inputrangebegin)).options(transpose=True).value = aList
    windowProcess.destroy()
    tk.messagebox.showinfo(title='Success', message='输入完成')


#学习过程
def clickButton2():
    i = 0
    #连接到excel
    book = xw.Book(r'myexcel.xlsx')
    workbook = book.sheets('Sheet1')  #连接excel文件
    workbook2 = book.sheets('Sheet2')  #连接excel文件
    dbend = e3.get()
    #模糊匹配
    inputrangebegin = e1.get()
    inputrangeend = e2.get()
    leng = str(int(inputrangeend) - int(inputrangebegin) + 1)
    #检查输入合法
    if (inputrangebegin > inputrangeend or inputrangebegin == 0
            or inputrangeend == 0):
        inputError()
        return
    dataInput = workbook.range('A' + str(inputrangebegin) + ':' + 'A' +
                                str(inputrangeend)).value
    windowProcess = tk.Toplevel()
    windowProcess.geometry('300x200')
    windowProcess.title('进度')
    l1 = tk.Label(windowProcess,
                    text='处理进度0/' + leng,
                    font=('Microsoft YaHei', 14))
    l1.place(x=100, y=50)

    aList = []
    session = HTMLSession()
    #在百科搜索名字，在字典中匹配
    for school in dataInput:
        url = 'https://baike.baidu.com/search/word?word=' + school
        try:
            r = session.get(url)
        except:
            tk.messagebox.showinfo(title='Error', message='无连接，请检查网络')
            return
        sel0 = 'body > div.body-wrapper.feature.feature_small.collegeSmall > div.feature_poster > div > div.poster-left > div.poster-top > dl > dd > h1'
        sel1 = 'body > div.body-wrapper > div.content-wrapper > div > div.main-content > dl.lemmaWgt-lemmaTitle.lemmaWgt-lemmaTitle- > dd > h1'
        sel2 = '#body_wrapper > div.searchResult > dl > dd:nth-child(2) > a'

        results = r.html.find(sel0)
        key = ''
        for result in results:
            key = result.text
        results = r.html.find(sel1)
        for result in results:
            key = result.text

        results = r.html.find(sel2)
        for result in results:
            mylink = list(result.absolute_links)[0]
            r2 = session.get(mylink)
            results2 = r2.html.find(sel0)
            for result in results2:
                key = result.text
            results2 = r2.html.find(sel1)
            for result in results2:
                key = result.text

        l1["text"] = '处理进度' + str(i + 1) + '/' + leng
        l1.update()
        if key:
            aList.append(key)
        else:
            aList.append(
                workbook.range('A' + str(int(inputrangebegin) + i)).value)
        i += 1
    #输出结果
    workbook2.range('B' +
                    str(int(dbend) + 1)).options(transpose=True).value = aList
    workbook2.range('C' + str(int(dbend) + 1)).options(
        transpose=True).value = workbook.range('B' + str(inputrangebegin) +
                                                ':' + 'B' +
                                                str(inputrangeend)).value
    windowProcess.destroy()
    tk.messagebox.showinfo(title='Success', message='学习完成')

#两个按钮
b1 = tk.Button(window,
                text='开始处理',
                font=('Microsoft YaHei', 12),
                width=10,
                height=1,
                command=clickButton1).place(x=100, y=200)
b2 = tk.Button(window,
                text='开始学习',
                font=('Microsoft YaHei', 12),
                width=10,
                height=1,
                command=clickButton2).place(x=300, y=200)

window.mainloop()
