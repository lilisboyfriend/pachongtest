import requests
import os
import time
import threading
from bs4 import BeautifulSoup
import random

def download_page(url):
   '''
   用于下载页面
   '''
   headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:61.0) Gecko/20100101 Firefox/61.0"}
   r = requests.get(url, headers=headers)
   r.encoding = 'gb2312'
   return r.text

def get_pic_list(html):
   '''
   获取每个页面的套图列表,之后循环调用get_pic函数获取图片
   '''
   soup = BeautifulSoup(html, 'html.parser')
   #pic_list = soup.find_all('li')  # class id
   pic_list = soup.find('main').find('div',id='thumbs').find('section').find('ul').find_all('li')
   for i in pic_list:
       a_tag = i.find('a')
       link = a_tag.get('href')  # 套图链接
       text = random.randint(0,1000)
      # text = a_tag.get_text()   # 套图名字
       get_pic(link, text)

def get_pic(link, text):
   '''
   获取当前页面的图片,并保存
   '''
   link2 = '{}.html'.format(link)
   html = download_page(link)  # 下载界面
   soup = BeautifulSoup(html, 'html.parser')
   # pic_list = soup.find('div', class_="scrollbox").find_all('img')  # 找到界面所有图片
   # pptt1 = soup.find('main')
   # pptt2 = pptt1.find('section',id = 'showcase')
   # pptt3 = pptt2.find('div',class_ = 'scrollbox')
   # pptt4 = pptt3.find('img')
   pic_list = soup.find('main').find('section').find('div',class_='scrollbox').find_all('img')
   headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:61.0) Gecko/20100101 Firefox/61.0"}
   # create_dir('pic/{}'.format(text))
   for i in pic_list:
       pic_link = i.get('src')  # 拿到图片的具体 url
       r = requests.get(pic_link, headers=headers)  # 下载图片，之后保存到文件
       #with open('pic/{}/{}'.format(text, link.split('/')[-1]), 'wb') as f:
       with open('pic/{}.jpg'.format(link.split('/')[-1]), 'wb') as f:
           f.write(r.content)
           time.sleep(1)

def create_dir(name):
   if not os.path.exists(name):
       os.makedirs(name)


def execute(url):
   page_html = download_page(url)
   get_pic_list(page_html)


def main():
   # create_dir('pic2')

   for i in range(24,30):
       url = 'https://wallhaven.cc/toplist?page={}'.format(i)
       page_html = download_page(url)
       get_pic_list(page_html)
   # url = 'https://wallhaven.cc/toplist?page=3'
   # page_html = download_page(url)
   # get_pic_list(page_html)

   # queue = [i for i in range(11, 30)]   # 构造 url 链接 页码。
   # threads = []
   # while len(queue) > 0:
   #     for thread in threads:
   #         if not thread.is_alive():
   #             threads.remove(thread)
   #     while len(threads) < 5 and len(queue) > 0:   # 最大线程数设置为 5
   #         cur_page = queue.pop(0)
   #         url = 'https://wallhaven.cc/toplist?page={}.html'.format(cur_page)
   #         thread = threading.Thread(target=execute, args=(url,))
   #         thread.setDaemon(True)
   #         thread.start()
   #         print('{}正在下载{}页'.format(threading.current_thread().name, cur_page))
   #         threads.append(thread)

if __name__ == '__main__':
   main()