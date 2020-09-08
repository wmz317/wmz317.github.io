import requests
import json
import openpyxl
import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
import smtplib
'''
1 定时-获取数据
2 存入xlsx
3 email send out
'''
# 打开xls文件准备写入
#dirname = "C:\\Users\\SR\\desktop\\QntA\\" # 本地存储
nowTime = datetime.datetime.now().strftime('_%m%d%H%M')
file_name = '511380Quote'+nowTime+'.xlsx'
# workbook = openpyxl.load_workbook(dirname + filename)
wb = openpyxl.Workbook() #新建工作簿（也自动生成1个工作表'Sheet'）
#ws = wb.create_sheet() # 新建工作表依序'Sheet1'
ws = wb["Sheet"]
ws.title= '511380'+nowTime

# 工作表第一行名称写入
dataTitle=['time','current价格','premium_Rate','volume成交量','total_shares总份额','market_capital总净值','iopv']
for c in range(len(dataTitle)):
    ws.cell(1,c+1,dataTitle[c])
 
# 保存数据到指定位置
#wb.save(dirname+file_name) # 本地存储
wb.save(file_name) #远端存储
sendmail(fime_name)

def sendmail(file_name):
    # smtp setup
    smtp_server = 'smtp.163.com'
    usr = 'carter317'
    ps = 'LIEJBKYFFZXYADPV'
    from_addr = 'carter317@163.com'
    to_addr = 'carter317@163.com'
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
    except: Exception:
        print(traceback.print_exc())
        print("邮件发送失败")
        
    


print('任务完成')
