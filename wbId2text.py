from requests import get
import json

def wbId2Text1(id='4529065535736168'):
    id=str(id)
    url='https://m.weibo.cn/statuses/extend?id='
    url+=id
    jsn = get(url).text
    try:
        text= jsn.encode('latin-1').decode('unicode_escape')
        txt=text.split('"')[9].split("<br \/><br \/>")
    except UnicodeEncodeError:
        #print("编码异常")
        txt=["此条编码异常..."]
    except:
        print("编码或未知异常")
        
    
    for i in txt: print('>>'+i)
    print('**************分隔符******************')


#wbId2Text1()
# 特例故障u1='https://m.weibo.cn/statuses/extend?id=4529178461872473'
# 正常例子 url='https://m.weibo.cn/statuses/extend?id=4529065535736168'

