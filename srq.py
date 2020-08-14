from requests import get
# import requests
import time

def ssrq():
    a500='sh510500'
    a300='sh510300'
    cy50='sz159949'
    ic = 'sz159995'
    lsts=[a300,a500,cy50,ic]
    # nw= time.strftime('%m/%d %H:%M:%S ',time.localtime(time.time()))
    # 远程执行显示的是UTC时间，需要加8
    nw= time.strftime('%m/%d %H:%M:%S ',time.localtime(time.time()+28800))
    print("300+500+C50+IC"+" | "+nw)
    for i in lsts: srq(i)


def srq(s):
#预留,判断SH SZ
   
    ht='http://hq.sinajs.cn/list='
    ht+=s
    r= get(ht).text
    #r= requests.get(ht).text
    lst= r.split(',')
    open=lst[1]
    yest=lst[2]
    now= lst[3]
    max= lst[4]
    min= lst[5]
    ratio = int((float(now)/float(yest)-1)*10000)/100
    t=s+'== now:'+now+'; max:'+max+';min:'+min+ "; r:"+str(ratio)
    return print(t)

ssrq()
print('**************************************************')
print("Digit + 科蓝 | 华阳 + 康德 +atl")
srq("sz300579");
srq("sz300663");
srq("sz002906")
srq("sh603987"); 
srq("sz300750")
