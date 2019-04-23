# -*- coding: utf-8 -*-
"""
Created on Sat Aug  4 12:40:15 2018

@author: Yang Cao
"""
from IPython import get_ipython
get_ipython().magic('reset -sf')

import pandas as pd
import json
import requests
from urllib.request import urlretrieve,ProxyHandler,build_opener,install_opener
import glob, os
from urllib.request import URLError, HTTPError
from http.client import RemoteDisconnected
import freeproxy


def import_stock_list(path):
    stock = pd.read_excel(path,dtype=str)
    stocklist = [stock.iloc[:,0][i] for i in range(len(stock.index))]
    return stocklist

def download(stock,start,end):
    base_url = 'http://static.cninfo.com.cn'
    request_url = 'http://www.cninfo.com.cn/new/fulltextSearch/full?searchkey='+stock+'&sdate=&edate=&isfulltext=false&sortName=nothing&sortType=desc&pageNum=1'
    HEADER = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/68.0.3440.75 Safari/537.36',
                'X-Requested-With': 'XMLHttpRequest'
                }
    r = requests.post(request_url, HEADER)
    text = json.loads(r.text)
    num = len(text['announcements'])
    if num == 0:
        print('股票%s没有找到任何公告' %stock)
        return False
    else:
        print('股票%s找到公告文件，正在下载...' %stock)
        for i in range(num):
            download_links = text['announcements'][i]['adjunctUrl']
            if download_links.endswith('.PDF'):
                try:
                    titles = text['announcements'][i]['announcementTitle'].split('：')[1].replace('<em>','').replace('</em>','')
                except IndexError: 
                    titles = text['announcements'][i]['announcementTitle'].replace('<em>','').replace('</em>','')
                time = text['announcements'][i]['adjunctUrl'].split('/')[1]
                time1 = pd.to_datetime(time, format='%Y-%m-%d')
                start = pd.to_datetime(start, format='%Y-%m-%d')
                end = pd.to_datetime(end, format='%Y-%m-%d')
                if time1 < start or time1 > end:
                    continue
                secCode = stock
                secName = text['announcements'][0]['secName']
                filename = secCode+'_'+secName+'_'+time+'_'+titles+'.PDF'
                if filename not in currentFiles:
                    try:
                        urlretrieve(base_url+'/'+download_links,filename=filename)
                    except (URLError,HTTPError,RemoteDisconnected,ConnectionResetError):
                       ipn = 0
                       tbd = 0
                       while tbd == 0:
                          proxies = {"http": proxylist[ipn]}
                          proxy_support = ProxyHandler(proxies)
                          opener = build_opener(proxy_support)
                          install_opener(opener)
                          try:
                              response = opener.open(base_url+'/'+download_links)
                          except (URLError,HTTPError):
                              ipn += 1
                              continue
                          with open(filename,'wb') as f:
                              f.write(response.read())
                              f.close()
                          tbd = 1          
                else:
                    continue
        p = 2
        while text['hasMore'] == True:
            request_url = 'http://www.cninfo.com.cn/new/fulltextSearch/full?searchkey='+stock+'&sdate=&edate=&isfulltext=false&sortName=nothing&sortType=desc&pageNum='+str(p)
            r = requests.post(request_url, HEADER)
            text = json.loads(r.text)
            num = len(text['announcements'])
            for i in range(num):
                download_links = text['announcements'][i]['adjunctUrl']
                if download_links.endswith('.PDF'):
                    try:
                        titles = text['announcements'][i]['announcementTitle'].split('：')[1].replace('<em>','').replace('</em>','')
                    except IndexError: 
                        titles = text['announcements'][i]['announcementTitle'].replace('<em>','').replace('</em>','')
                    time = text['announcements'][i]['adjunctUrl'].split('/')[1]
                    time1 = pd.to_datetime(time, format='%Y-%m-%d')
                    if time1 < start or time1 > end:
                        continue
                    secCode = stock
                    secName = text['announcements'][0]['secName']
                    filename = secCode+'_'+secName+'_'+time+'_'+titles+'.PDF'
                    if filename not in currentFiles:
                        try:
                            urlretrieve(base_url+'/'+download_links,filename=filename)
                        except (URLError,HTTPError,RemoteDisconnected,ConnectionResetError):
                           ipn = 0
                           tbd = 0
                           while tbd == 0:
                              proxies = {"http": proxylist[ipn]}
                              proxy_support = ProxyHandler(proxies)
                              opener = build_opener(proxy_support)
                              install_opener(opener)
                              try:
                                  response = opener.open(base_url+'/'+download_links)
                              except (URLError,HTTPError):
                                  ipn += 1
                                  continue
                              with open(filename,'wb') as f:
                                  f.write(response.read())
                                  f.close()
                              tbd = 1  
                    else:
                        continue
            p = p+1
        print('下载完毕')
        return True
#-------------------------------
#Directory that you want to place your pdfs
final_dir = 'C:/UMD/hezhen' #下载文件存放地址
stock_list_dir = 'C:/UMD/hezhen/股票代码.xlsx' #股票代码列表存放地址
start = '2011-05-30'
end = '2019-06-01'

os.chdir(final_dir)
stocklist = import_stock_list(stock_list_dir)
nonFound = []
found = []
pdfFiles = glob.glob('*.pdf')
currentFiles = set(pdfFiles)
proxylist = freeproxy.from_xici_daili()

for stock in stocklist:
    if download(stock,start,end) == True:
        found.append(stock)
    else:
        nonFound.append(stock)
        
print('列表下载完毕')
print('找到%d家公司公告\n%s\n' %(len(found),found))
print('没有找到%d家公司公告\n%s\n' %(len(nonFound),nonFound))
