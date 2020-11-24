import xlwings as xw
from requests_html import HTMLSession
#连接到excel
book = xw.Book(r'myexcel.xlsx')
workbook = book.sheets('Sheet1')#连接excel文件
workbook2 = book.sheets('Sheet2')#连接excel文件
#创建字典
dbname = workbook2.range('B1:B2002').value
dbarea = workbook2.range('C1:C2002').value
dic = dict(zip(dbname, dbarea))
#模糊匹配
inputrangebegin = 63
inputrangeend = 68
dataInput = workbook.range('A' + inputrangebegin + ':' + 'A' + inputrangeend).value
session = HTMLSession()

for school in dataInput:
    url = 'https://baike.baidu.com/search/word?word='+school
    r = session.get(url)
    sel0 = 'body > div.body-wrapper.feature.feature_small.collegeSmall > div.feature_poster > div > div.poster-left > div.poster-top > dl > dd > h1'
    sel1 = 'body > div.body-wrapper > div.content-wrapper > div > div.main-content > dl.lemmaWgt-lemmaTitle.lemmaWgt-lemmaTitle- > dd > h1'
    results = r.html.find(sel0)
    for result in results:
        key = result.text
    results = r.html.find(sel1)
    for result in results:
        key = result.text
    print(key)
    if dic.__contains__(key):
        print(dic[key])
    else:
        print('null')

