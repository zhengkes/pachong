# -*- coding: UTF-8 -*-
import sys
reload(sys)
#eclipse这个插件问题，可以忽略
sys.setdefaultencoding('utf8')
from bs4 import BeautifulSoup
import re
import urllib2
import xlwt
import MySQLdb

#得到页面全部内容
def askURL(url):
    request = urllib2.Request(url)#发送请求
    try:
        response = urllib2.urlopen(request)#取得响应
        html= response.read()#获取网页内容
        #print html
    except urllib2.URLError, e:
        if hasattr(e,"code"):
            print e.code
        if hasattr(e,"reason"):
            print e.reason
    return html

#获取相关内容
def getData(baseurl):
    findImgSrc=re.compile(r'<img.*src="(.*jpg)"',re.S)#找到影片图片
    findTitle=re.compile(r'<span class="title">(.*)</span>')#找到片名
    #找到评分
    findRating=re.compile(r'<span class="rating_num" property="v:average">(.*)</span>')
    #找到描述
    findInq=re.compile(r'<span class="inq">(.*)</span>')
    datalist=[]
    for i in range(0,10):
        url=baseurl+str(i*25)
        html=askURL(url)
        soup = BeautifulSoup(html)
        for item in soup.find_all('div',class_='item'):#找到每一个影片项
            data=[]
            item=str(item)#转换成字符串
            imgSrc=re.findall(findImgSrc,item)[0]
            data.append(imgSrc)#添加图片链接
            titles=re.findall(findTitle,item)
            data.append(titles[0])#添加中文片名
            rating=re.findall(findRating,item)[0]
            data.append(rating)#添加评分
            inq=re.findall(findInq,item)
            #可能没有概况
            if len(inq)!=0:
                inq=inq[0].replace("。","")#去掉句号
                data.append(inq)#添加概况
            else:
                data.append(' ')#留空
            datalist.append(data)
    return datalist

#将相关数据写入数据库中
def saveToDatabase(datalist,db):
    cursor=db.cursor()#获取数据库操作游标
    cursor.execute('SET NAMES UTF8')
    for i in range(0,250):
        data=datalist[i]#['pic','name','score','info']
        sql="insert into movielist(itempic,itemname,itemscore,itemdescrp)\
             values('%s','%s','%f','%s')" % \
             (data[0],data[1],float(data[2]),data[3])
        try:
            cursor.execute(sql)
            db.commit()
        except:
            db.rollback()
    db.close()
        
#将相关数据写入excel中
def saveDataAsExcell(datalist,savepath):
    book=xlwt.Workbook(encoding='utf-8',style_compression=0)
    sheet=book.add_sheet('豆瓣电影Top250',cell_overwrite_ok=True)
    col=('图片链接','影片中文名',
                '评分','概况')
    for i in range(0,4):
        sheet.write(0,i,col[i])#列名
    for i in range(0,250):
        data=datalist[i]
        for j in range(0,4):
            sheet.write(i+1,j,data[j])#数据
    book.save(savepath)#保存

def main():
    #豆瓣top250链接
    baseurl='https://movie.douban.com/top250?start='
    datalist=getData(baseurl)
    print '成功爬取到数据'
    #获取Mysql连接
    db=MySQLdb.connect(user="root",passwd="",host="127.0.0.1",db="data")
    saveToDatabase(datalist, db)
    print '导入数据库保存正常'
    savepath=u'movieTop250.xlsx'
    saveDataAsExcell(datalist, savepath)
    print '导出到excel正常'
main()