import openpyxl
from bs4 import BeautifulSoup
from openpyxl import Workbook
import requests

header={
    'user-agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/101.0.4951.64 Safari/537.36 Edg/101.0.1210.53',
    'referer':'https://forum.gamer.com.tw/B.php?page=1&bsn=60076',
}

#輸入要蒐集的網頁數
start=input('start:')
end=input('end:')
pageNumStart=int(start)
pageNumEnd=int(end)

#填入excel基建
new=Workbook()
sheet=new.create_sheet('head_and_id',0)
row=1

#開始尋找，加入表單
for num in range(pageNumStart,pageNumEnd):

    url='https://forum.gamer.com.tw/C.php?bsn=60076&snA='+str(num)+'&tnum=11'
    html=requests.get(url,headers=header)
    bs=BeautifulSoup(html.text,'lxml')
    main=bs.find('h1',{'class':'title'}).text
    id=bs.find('div',{'class':'hint'})
    if not id:
        id=bs.find('a',{'class':'userid'}).text
    else:
        id=bs.find('div',{'class':'hint'}).text
        id=id[9:-3]


    sheet['A'+str(row)]=main
    sheet['B'+str(row)]=id
    row+=1
input('enter')

new.save('bahamut.xlsx')
new.close()