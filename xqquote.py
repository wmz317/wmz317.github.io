import requests
import json
import openpyxl
import datetime
'''
1 定时-获取数据
2 存入xlsx
3 email send out
'''
# 打开xls文件准备写入
#dirname = "C:\\Users\\SR\\desktop\\QntA\\" # 本地存储
filename = '511380Quote'
nowTime = datetime.datetime.now().strftime('_%m%d%H%M')
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
#wb.save(dirname+filename+nowTime+'.xlsx') # 本地存储
wb.save(filename+nowTime+'.xlsx') #远端存储
print('任务完成')
