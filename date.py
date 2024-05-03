# -*- coding=utf-8 -*-

from selenium import webdriver
from selenium.webdriver.common.by import By
import requests
import json
import time
# import chardet
import re
from urllib import parse
# from get_gtk import getGTK
# from save_to_execl import write_excel
from openpyxl import Workbook
import datetime

login_url = 'https://i.qq.com/'

UserName = 'un'
PassWd = 'pwd'

index_url = "https://user.qzone.qq.com/" + UserName


def getGTK(cookie):
    """ 根据cookie得到GTK """
    hashes = 5381
    for letter in cookie['p_skey']:
        hashes += (hashes << 5) + ord(letter)

    return hashes & 0x7fffffff


def Login(UserName, PassWd):
    '''
    return cookies
    '''
    browser = webdriver.Chrome()
    browser.get(login_url)
    browser.switch_to.frame(browser.find_element(By.XPATH, '//iframe[@id="login_frame"]'))
    time.sleep(1)
    l_btn01 = browser.find_element(By.XPATH, '//*[@id="switcher_plogin"]').click()
    time.sleep(1)
    browser.find_element(By.XPATH, "//input[@id='u']").send_keys(UserName)
    time.sleep(1)
    browser.find_element(By.XPATH, '//input[@id="p"]').send_keys(PassWd)
    time.sleep(1)
    browser.find_element(By.XPATH, '//input[@class="btn"]').click()
    time.sleep(1)
    browser.switch_to.default_content()  # 切出
    time.sleep(1)
    return browser


def get_cookies(browser):
    cookie_items = browser.get_cookies()

    post = {}
    for cookie_item in cookie_items:
        post[cookie_item['name']] = cookie_item['value']
    cookies_dict_str = json.dumps(post)
    cookies_dict = eval(cookies_dict_str)
    cookie_str = ""
    cookie_dict = {}
    for key in cookies_dict:
        value = cookies_dict[key]
        str1 = key + "=" + value
        cookie_str = cookie_str + str1 + ";"
        cookie_dict[key] = value
    return cookie_str, cookie_dict


def get_tocken(cookies):
    url = index_url
    headers = {
        "cookie": cookies,
        'referer': "https://i.qq.com/",
        "upgrade-insecure-requests": "1",
        "user-agent": "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/66.0.3359.139 Safari/537.36",
    }
    # 这里注意混合编码，网页中charset有两种，utf-8,gbk
    res = requests.get(url, headers=headers)
    contents = res.content
    # encode = chardet.detect(contents)   # 获取网页编码格式字典信息，字典encode中键encoding的值为编码格式
    html = contents.decode('gbk', 'ignore')  # 用分析出的网页编码格式，解码
    try:
        pattern = re.compile(r'{ try{return\s?"(.*?)";}')
        token = re.findall(pattern, html)[0]
    except IndexError:
        token = ''
    return token


def get_all_friends(uin, cookies, gtk, token):
    # 获得所有好友QQ号list
    base_url = 'https://user.qzone.qq.com/proxy/domain/r.qzone.qq.com/cgi-bin/tfriend/friend_ship_manager.cgi?'
    headers = {
        "cookie": cookies,
        'referer': "https://i.qq.com/",
        "upgrade-insecure-requests": "1",
        "user-agent": "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/66.0.3359.139 Safari/537.36",
    }

    # 通过urlencode拼接
    data = {

        "uin": uin,
        "do": "1",
        "fupdate": "1",
        "clean": "1",
        "g_tk": gtk,  #
        "qzonetoken": get_tocken(cookies),
        "g_tk": gtk,
    }
    data = parse.urlencode(data)
    url = base_url + data
    res = requests.get(url, headers=headers)
    html = res.text
    html = html.replace('_Callback(', '').replace(');', '')

    json_obj = json.loads(html)
    items_list = json_obj['data']['items_list']
    return items_list


def subtime(time1, time2):
    time1 = datetime.datetime.strptime(time1, "%Y-%m-%d %H:%M:%S")
    time2 = datetime.datetime.strptime(time2, "%Y-%m-%d %H:%M:%S")
    return time1 - time2


def get_time_info(uin, frienduin, g_tk, token, cookies, meta):
    url = 'https://user.qzone.qq.com/proxy/domain/r.qzone.qq.com/cgi-bin/friendship/cgi_friendship?'
    headers = {
        "cookie": cookies,
        "referer": index_url,
        "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/68.0.3440.106 Safari/537.36",
    }
    params = {
        "activeuin": uin,
        "passiveuin": frienduin,
        "situation": "1",
        "isCalendar": "1",
        "g_tk": g_tk,
        "qzonetoken": token,
        "g_tk": g_tk,
    }

    response = requests.get(url, headers=headers, params=params)
    html = response.text
    html = html.replace('_Callback(', '').replace(');', '')
    json_obj = json.loads(html)
    if 'data' not in json_obj:
        return {}
    data = json_obj['data']

    # 对方星座
    info = {}
    info['好友'] = meta['好友']
    info['开始时间'] = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(data.get('addFriendTime', 0)))
    if 'constellation' in data:
        info['好友星座'] = data['constellation']['parCauTitle']
        info['我的星座'] = data['constellation']['title']
    else:
        info['好友星座'] = ' '
        info['我的星座'] = ' '
    info['现在时间'] = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
    info['成为好友时间'] = subtime(info['现在时间'], info['开始时间'])
    # 写入Excel需要把时间类型处理成str
    info['开始时间'], info['现在时间'], info['成为好友时间'] = map(str, [info['开始时间'], info['现在时间'],
                                                                         info['成为好友时间']])
    return info


def main():
    browser = Login(UserName, PassWd)
    cookies, cookie_map = get_cookies(browser)
    print(" => cookies:", cookies)
    uin = UserName
    token = get_tocken(cookies)
    print("Test => token:", token)
    g_tk = getGTK(cookie_map)
    print("Test => g_tk:", g_tk)
    items_list = get_all_friends(uin, cookies, g_tk, token)
    workbook = Workbook()

    # 默认sheet
    sheet = workbook.active  # 激活sheet
    sheet.title = "friends"  # 设置sheet名字
    sheet.append(["好友", "开始时间", "好友星座", "我的星座", "现在时间", "成为好友时间"])  # 插入标题

    for key in items_list:
        friend_name = key['name']
        frienduin = key['uin']

        meta = {
            "好友": friend_name,
        }
        info = get_time_info(uin, frienduin, g_tk, token, cookies, meta)
        try:
            print(info)
        except Exception as e:
            continue

        sheet.append(list(info.values()))
        # write_excel(info, 'QQ_friends')
    browser.close()
    workbook.save("friends.xlsx")


if __name__ == '__main__':
    main()
