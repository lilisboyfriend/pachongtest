import concurrent
import os
from concurrent.futures import ThreadPoolExecutor
import requests
from bs4 import BeautifulSoup


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
    for i in range(1, 6):
        baseurl = 'https://www.bbmeitu.com/mote/list_2_{}.html'.format(i)
        # baseurl = 'https://www.bbmeitu.com/fuliji/list_1_{}.html'.format(i)
        html = request_page(baseurl)
        soup = BeautifulSoup(html, 'lxml')
        list = soup.find(class_='content g-panel').find(class_='main').find(class_='list').find_all("article")

        for item in list:
            url = item.find('a').get('href')
            print('页面链接：%s' % url)
            url2 = "https://www.bbmeitu.com/"+url
            urls.append(url2)

    return urls


def download_Pic(title, image_list):
    # 新建文件夹
    mk = "picfl/"+title
    # os.mkdir(mk)
    os.makedirs(mk)
    j = 1
    # 下载图片
    for item in image_list:
        item2 = "https:"+item
        filename = 'picfl/%s/%s.jpg' % (title, str(j))
        print('downloading....%s : NO.%s' % (title, str(j)))
        img = requests.get(item2)
        with open(filename, 'wb') as f:
            f.write(img.content)
        j += 1

def download(url):
    html = request_page(url)
    soup = BeautifulSoup(html, 'lxml')
    total = soup.find(class_='content g-panel').find(class_='main').find(class_='pager').find_all('a')[-2].string
    title = soup.find('h2').string
    image_list = []

    soup = BeautifulSoup(html, 'lxml')
    img_url = soup.find('img').get('src')
    image_list.append(img_url)

    for i in range(2,int(total)+1):
        ta = "_{}.html".format(i)
        # url2 = url+'_'+'%s'%(i + 1)
        url2 = url.replace(".html",ta)
        html = request_page(url2)
        soup = BeautifulSoup(html, 'lxml')
        img_url = soup.find('img').get('src')
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

    download_all_images(list_page_urls)

    # for url in list_page_urls:
    #     download(url)
