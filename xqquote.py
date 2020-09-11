import datetime
import time
import requests
import openpyxl
#from  selenium import webdriver
#import re
#import json
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
import smtplib

sCode='SH511380'
statuscode=5  # 默认5，调试可按不同状态

def sendml(file_name):
    # smtp setup
    smtp_server = 'smtp.163.com'
    usr = 'carter317'
    ps = 'LIEJBKYFFZXYADPV'
    from_addr = 'carter317@163.com'
    to_addr = 'wmz_317@sina.com'
    #生成一个空的带附件的邮件实例
    message = MIMEMultipart()
    title = 'Py xls Mail Test' 
    message['Subject'] = title
    content = '本文发送时间：'+ datetime.datetime.now().strftime('%y%m%d%H%M')
    #将正文以text的形式插入邮件中
    message.attach(MIMEText(content, 'plain', 'utf-8'))
    #读取附件的内容
    att = MIMEText(open(file_name, 'rb').read(), 'base64', 'utf-8')
    #att["Content-Type"] = 'application/octet-stream' #不明所以
    #生成附件的名称
    att.add_header('Content-Disposition', 'attachment', filename=Header(file_name,'utf-8').encode())
    #将附件内容插入邮件中
    message.attach(att)
    
    try:
        server = smtplib.SMTP_SSL(smtp_server, 465) # 启用SSL发信, 端口一般是465
        server.login(usr, ps)
        server.sendmail(from_addr, [to_addr], message.as_string())
        print("邮件发送成功")
        server.quit()  # 关闭连接
    except Exception as e:
        print('出现错误：',e)
        print("邮件发送失败") 

# 建立excel表格
#dirname = "D:\\staf\\QNT\\" # 本地存储
#dirname = "C:\\Users\\SR\\Desktop\\QntA\\" # 本地存储
nowTime = datetime.datetime.now().strftime('_%m%d_%H%M')
file_name1 = sCode+'Quote'+nowTime+'.xlsx'
wb = openpyxl.Workbook() #新建工作簿（也自动生成1个工作表'Sheet'）
ws = wb["Sheet"]
ws.title= sCode+nowTime

url= 'https://stock.xueqiu.com/v5/stock/quote.json?symbol='+sCode+'&extend=detail'
url2='https://stock.xueqiu.com/v5/stock/realtime/pankou.json?symbol='+sCode

headers ={"User-Agent": "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.108 Safari/537.36"}

session = requests.Session()
# 第一步,向xueqiu网首页发送一条请求,获取cookie
session.get(url="https://xueqiu.com",headers=headers)

# 第二步,获取动态加载的数据

time.sleep(1);

# 工作表第一行名称title表头写入
a1 = session.get(url=url,headers=headers).json()
a2=a1['data']['quote'] #主要数据
# Title_quote=list(a2.keys()) #a2.keys是dict_keys格式，需要转化
Title_quote = ['timestamp', 'current', 'premium_rate', 'percent','symbol',  'high','low','avg_price', 'acc_unit_nav','limit_up','limit_down', 'amplitude', 'float_market_capital', 'market_capital', 'iopv', 'amount', 'chg', 'last_close', 'open','volume', 'unit_nav', 'total_shares', 'time','status']

for c in range(len(Title_quote)):
    ws.cell(1,c+1,Title_quote[c])

columns=ws.max_column
ws.cell(1,columns+1,'status')
ws.cell(1,columns+2,'status_id')
ws.cell(1,columns+3,'pankou_ratio')
ws.cell(1,columns+4,'融')
ws.cell(1,columns+5,'空')
ws.cell(1,columns+6,'error_code')
ws.cell(1,columns+7,'实时时间')

title_5d= ['timestamp','bp1','bc1', 'bp2', 'bc2','bp3', 'bc3', 'bp4','bc4', 'bp5', 'bc5', 'current', 'sp1', 'sc1', 'sp2', 'sc2', 'sp3', 'sc3', 'sp4', 'sc4', 'sp5', 'sc5','buypct','sellpct','diff','ratio']
for d in range(len(title_5d)):
    ws.cell(1,columns+d+8,title_5d[d])
print('已创建表头')

# 循环获取数据并写入xls

for i in range(0,6000):
    # get HQ quote data
    try: 
        a1 = session.get(url=url,headers=headers).json()
        a2 = a1['data']['quote'] #主要数据[字典]
    except Exception as e:
        print('报错1：',e)
        a1['error_code']=e #有故障时使用前一次数据，并把故障备注

    # 更改时间为易读格式
    timestr = time.strftime('%H:%M:%S ',time.localtime(a2['timestamp']/1000))
    a2['timestamp'] = timestr

    # get 5dang data  
    try:
        a51 = session.get(url=url2,headers=headers).json()
        a52=a51['data'] #主要数据[字典]
    except Exception as e:
        print('报错2：',e)
        a1['error_code']=e #有故障时使用前一次数据，并把故障备注
        
    #另+1网络实时时间
    #realtime= time.strftime('%H:%M:%S ',time.localtime(time.time()+28800))
    realtime= datetime.datetime.now().strftime('%H:%M:%S')
    
    rows = ws.max_row # 获得行数
    # 当前页数据写入xlsx
    HQstatus =a1['data']['market']['status_id']
    if HQstatus ==statuscode: 
        n=0
        for v in Title_quote:
            n+=1
            ws.cell(rows+1, n, a2[v])
        ws.cell(rows+1,columns+1,a1['data']['market']['status'])  
        ws.cell(rows+1,columns+2,a1['data']['market']['status_id'])
        ws.cell(rows+1,columns+3,a1['data']['others']['pankou_ratio'])
        ws.cell(rows+1,columns+4,a1['data']['tags'][0]['value'])
        ws.cell(rows+1,columns+5,a1['data']['tags'][1]['value'])
        ws.cell(rows+1,columns+6,a1['error_code'])
        ws.cell(rows+1,columns+7,realtime)
       
        # 5dang数据写入
        for x in title_5d:
            ws.cell(rows+1, columns+n+8, a52[x])
        if i%60==0:
            print(realtime+'_测试-写入数据完成'+str(i))
    time.sleep(3)
    #时间到则退出

    if HQstatus == 7:
        print('数据时间已结束')
        break

# 保存数据到指定位置
#wb.save(dirname+file_name1) # 本地存储
wb.save(file_name1) # romote存储
print('HQ file saved!')

sendml(file_name1)

print('任务完成！')
            
'''
https://stock.xueqiu.com/v5/stock/quote.json?symbol=SH511380&extend=detail
# 5dang https://stock.xueqiu.com/v5/stock/realtime/pankou.json?symbol=SH511380
# 明细 https://stock.xueqiu.com/v5/stock/history/trade.json?symbol=SH511380&coun
https://xueqiu.com/S/SH511380

quote data:

{"data":{"market":{"status_id":5,"region":"CN","status":"交易中","time_zone":"Asia/Shanghai","time_zone_desc":null},"quote":{"symbol":"SH511380","code":"511380","acc_unit_nav":1.034,"high52w":10.626,"nav_date":1599494400000,"avg_price":10.256,"delayed":0,"type":13,"expiration_date":null,"percent":-1.03,"tick_size":0.001,"float_shares":null,"limit_down":9.307,"amplitude":0.92,"current":10.235,"high":10.307,"current_year_percent":3.42,"float_market_capital":null,"issue_date":1586188800000,"low":10.212,"sub_type":"EBS","market_capital":8.473218745E8,"currency":"CNY","lot_size":100,"lock_set":null,"iopv":10.238,"timestamp":1599620776400,"found_date":1583424000000,"amount":5.9096886E7,"chg":-0.106, "last_close":10.341,"volume":5762400,"volume_ratio":null,"limit_up":11.375,"turnover_rate":null,"low52w":9.555,"name":"可转债ETF","premium_rate":-0.03,"exchange":"SH","unit_nav":10.342,"time":1599620776400,"total_shares":82786700,"open":10.307,"status":1},"others":{"pankou_ratio":-4.16,"cyb_switch":true},"tags":[{"description":"融","value":6},{"description":"空","value":7}]},"error_code":0,"error_description":""}

status_id: 1 未开  3集竞 4午休 5trading 7已收
'''
