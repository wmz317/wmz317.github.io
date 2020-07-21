from requests import get

url='https://m.weibo.cn/statuses/extend?id=4529065535736168'
jsn = get(url).text
print(jsn)
