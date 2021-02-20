#utf-8
import requests

from zlib import decompress 

from bs4 import BeautifulSoup 

import re

import xlsxwriter

def connet():
    
    try:
        
        headers={"User-Agent":"Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.26 Safari/537.36 Core/1.63.5558.400 QQBrowser/10.1.1695.400"}
        
        url="https://cmts.iqiyi.com/bullet/36/00/4337533600_300_9.z?&business=danmu&is_iqiyi=true&is_video_page=true&tvid=4337533600&albumid=225086801&categoryid=2&qypid=01010021010000000000"
        
        html = requests.get(url,headers = headers)
        
        bullets_data = html.content
        
        print("连接成功")
        
        print("正在解析文本")

        return bullets_data
    
    except:
        
        print("连接网页失败")

def getcomments(bullets_data):
    
    try:

        bullets = decompress(bullets_data)
    
        danmu =bullets.decode("utf-8","ignore")

        soup=BeautifulSoup(danmu,"lxml")
    
        workbook = xlsxwriter.Workbook(file_excel)

        sheet = workbook.add_worksheet('弹幕文件')
 
        bold = workbook.add_format({'bold': True})

        sheet.write('A1', u'正文', bold)

        sheet.write('I1', u'时间', bold)

        sheet.write('L1', u'集数', bold)

        bulletInfo = soup.select('bulletInfo')

        row = 2

        for x in bulletInfo:

            content = x.content.text
     
            name="烈火军校第十五集"

            times = x.showtime.text

            time = str(int(eval(times)/60))+"分"+str(eval(times)-int(eval(times)/60)*60)+"秒"#分进制   eval函数转化数据格式变为可计算

            sheet.write('A%d' % row,66666)
        
            sheet.write('I%d' % row, time)

            sheet.write('L%d' % row, name)
 
            

        workbook.close()

        print("正在写入excel")

        print("写入成功")
    
        soups=soup.find_all("content")

        for i in soups:

            print(i)

        print("读取文件成功")

        return 0

    except:

        print("解析文件失败")


if __name__ == '__main__':

    file_excel="弹幕文件11.xlsx"
    
    bullets_data=connet()

    getcomments(bullets_data)

   
    
    
