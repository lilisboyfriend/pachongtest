import concurrent
import os
from concurrent.futures import ThreadPoolExecutor
import requests
from bs4 import BeautifulSoup
import time

def header(referer):

    headers = {
        'Host': 'i.meizitu.net',
        'Pragma': 'no-cache',
        'Accept-Encoding': 'gzip, deflate',
        'Accept-Language': 'zh-CN,zh;q=0.8,en;q=0.6',
        'Cache-Control': 'no-cache',
        'Connection': 'keep-alive',
        'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/72.0.3626.121 Safari/537.36',
        'Accept': 'image/webp,image/apng,image/*,*/*;q=0.8',
        'Referer': '{}'.format(referer),
    }

    return headers


def request_page(url):
    try:
        response = requests.get(url)
        response.encoding='utf-8'
        if response.status_code == 200:
            return response.text
    except requests.RequestException:
        return None


def get_page_urls():
    urls = []
    for i in range(1, 2):
        baseurl = 'https://www.tuaoo.cc/category-5_{}.html'.format(i)
        html = request_page(baseurl)
        soup = BeautifulSoup(html, 'lxml')
        list = soup.find(id='container').find('main').find_all('article')

        for item in list:
            url = item.find('a').get('href')
            print('页面链接：%s' % url)
            url2 = url
            urls.append(url2)


    return urls


def download_Pic(title, image_list):
    # 新建文件夹
    mk = "picmt/"+title
    # os.mkdir(mk)
    os.makedirs(mk)
    j = 1
    # 下载图片
    for item in image_list:
        # item2 = "https:"+item
        filename = 'picmt/%s/%s.jpg' % (title, str(j))
        print('downloading....%s : NO.%s' % (title, str(j)))
        img = requests.get(item)
        with open(filename, 'wb') as f:
            f.write(img.content)
            time.sleep(1)
        j += 1

def download(url):
    html = request_page(url)
    soup = BeautifulSoup(html, 'lxml')
    total = soup.find(id='container').find('main').find('ul').find_all('li')[-2].find('a').string
    # total = soup.find(id='container').find('main').find('article').find(class_='center').find('ul').find_all('li')[-2].find('a').string
    title = soup.find('h1').string
    image_list = []

    # soup = BeautifulSoup(html, 'lxml')
    # img_url = soup.find(id='container').find('main').find('img').get('src')
    # image_list.append(img_url)

    for i in range(1,int(total)):
        # ta = "_{}.html".format(i)
        # # url2 = url+'_'+'%s'%(i + 1)
        # url2 = url.replace(".html",ta)
        url2 = url+"?page=%s"%i
        html = request_page(url2)
        soup = BeautifulSoup(html, 'lxml')
        img_url = soup.find(id='container').find('main').find('img').get('src')
        image_list.append(img_url)
    download_Pic(title, image_list)


def download_all_images(list_page_urls):
    # 获取每一个详情妹纸
    # works = len(list_page_urls)
    with concurrent.futures.ProcessPoolExecutor(max_workers=8) as exector:
        for url in list_page_urls:
            exector.submit(download, url)


if __name__ == '__main__':
    # 获取每一页的链接和名称
    list_page_urls = get_page_urls()

    # download_all_images(list_page_urls)

    for url in list_page_urls:
        download(url)
