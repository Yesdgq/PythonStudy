#!/usr/bin/python3

import requests
import concurrent.futures 

# GET
# r = requests.get('https://api.github.com/user', auth=('user', 'pass'))
# print(r.status_code)


# print(b'Python is a programming language' in r.content)
# print(r.headers['content-type'])
# print(r.encoding)
# print(r.json())


# url = 'http://driver_53.api.tyidian.nucarf.tech/v3/memberCoupon/receive?sign=b8cd690e47006cc3bdb551d9adddb73e_1693390271_2399&appid=wxw&source=uniapp'  
  
# def make_request(url):  
#     response = requests.get(url)  
#     print(response.text)  
  
# with concurrent.futures.ThreadPoolExecutor(max_workers=10) as executor:  
#     while True:  
#         executor.submit(make_request, url)  
#         # time.sleep(1)

url = "http://driver_53.api.tyidian.nucarf.tech/v3/memberCoupon/receive?sign=b8cd690e47006cc3bdb551d9adddb73e_1693390271_2399&appid=wxw&source=uniapp"

sign = "b8cd690e47006cc3bdb551d9adddb73e_1693390271_2399"
appid = "wxw"
source = "uniapp"

num_requests = 10
for i in range(num_requests):
    response = requests.get(url, params={
        "sign": sign,
        "appid": appid,
        "source": source
    })
    print(f"{i+1}. Request: {response.status_code}")