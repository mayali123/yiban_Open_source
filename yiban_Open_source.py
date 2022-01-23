# -*- coding: UTF-8 -*-
# @Time :  13:26
# @Author :mayali123
# @File : yiban.py
# @Software : PyCharm
import requests     # 访问网页
import json         # json库
import xlrd         # 读取 表格 文件 用到
import time         # 时间库
import re           # 正则表达式库
import hashlib      # md5
# 登入的 网页
login_url = 'http://xggl.hnie.edu.cn/website/login'
# 获取上一次打卡data的网站
last_data_url = 'http://xggl.hnie.edu.cn/content/student/temp/zzdk/lastone'
# 访问 登入网页 要用的 请求头
headers = {
    'Host': "xggl.hnie.edu.cn",
    'Accept-Language': "zh-CN,zh;q=0.9,en;q=0.8",
    'Accept-Encoding': "gzip, deflate",
    'Content-Type': "application/x-www-form-urlencoded; charset=UTF-8",
    'Connection': "keep-alive",
    'Referer': "http://xggl.hnie.edu.cn/index",
    "user-agent":"Mozilla / 5.0(Linux;Android 4.0.4;Galaxy Nexus Build / IMM76B) AppleWebKit / 535.19(KHTML, like Gecko)"+\
    "Chrome / 18.0.1025.133 Mobile Safari / 535.19"
    }

# 用于转成学校需要的md5码
def getPW(pwd):
    encode_pwd = hashlib.md5(bytes(pwd, encoding='utf-8')).hexdigest()
    if len(encode_pwd) > 5:
        encode_pwd = encode_pwd[0:5] + "a" + encode_pwd[5:]

    if len(encode_pwd) > 10:
        encode_pwd = encode_pwd[0:10] + "b" + encode_pwd[10:]
    encode_pwd = encode_pwd[:len(encode_pwd)-2]
    return encode_pwd

# 这个函数 用于读取保存在本地的 用户信息
def read_data():
    # 打开保存信息的文件
    workbook = xlrd.open_workbook('test.xls')
    # 打开 sheet1
    sheet1 = workbook.sheet_by_index(0)
    # 行数
    row = sheet1.nrows
    # 保存 学号
    ID = []
    # 保存 密码
    password = []
    # 保存邮箱
    mail = []
    # 读取 文件 并 将对应数据 添加到 ID password mail
    for i in range(row - 1):
        ID.append(sheet1.cell_value(i + 1, 0))
        password.append(getPW(sheet1.cell_value(i + 1, 1)))
        mail.append(sheet1.cell_value(i + 1, 2))
    return ID, password, mail  # 返回 ID password mail

# text 用来发送的文字  to_addr 是 收信方邮箱  这个是在网上找的
def to_meg(text, to_addr):
    # smtplib 用于邮件的发信动作
    import smtplib
    from email.mime.text import MIMEText
    # email 用于构建邮件内容
    from email.header import Header
    # 用于构建邮件头

    # 发信方的信息：发信邮箱，QQ 邮箱授权码
    from_addr = '这里要填你的邮箱'
    password = '邮箱授权码 这里要填你的邮箱的授权码'

    # 收信方邮箱
    # to_addr = to_addr

    # 发信服务器
    smtp_server = 'smtp.qq.com'

    # 邮箱正文内容，第一个参数为内容，第二个参数为格式(plain 为纯文本)，第三个参数为编码
    msg = MIMEText(text, 'plain', 'utf-8')

    # 邮件头信息
    msg['From'] = Header(from_addr)
    msg['To'] = Header(to_addr)
    msg['Subject'] = Header('易班打卡')

    # 开启发信服务，这里使用的是加密传输
    server = smtplib.SMTP_SSL(host=smtp_server)
    server.connect(smtp_server, 465)
    # 登录发信邮箱
    server.login(from_addr, password)
    # 发送邮件
    server.sendmail(from_addr, to_addr, msg.as_string())
    # 关闭服务器
    server.quit()

def main():
    # 接受 read_data()函数返回的数据
    ID, password, mailID = read_data()
    for i in range(len(ID)):
        data = {    # 这个data 用来访问 登入网页  要用到的
            'uname': ID[i],
            'pd_mm': password[i]
        }
        req = requests.session()
        # 向 登入网页 发起post 请求
        recv = req.post(login_url, headers=headers, data=data).text
        # 将 str 转为 json 类型 json也就是字典
        recv = json.loads(recv)

        # 向 有上一次打卡data 的网站 发起get 请求
        last_data = req.get(url=last_data_url, headers=headers).text
        print(last_data)
        # 转为字典
        dic = json.loads(last_data)
        # 最后打卡要用到的 数据
        data = {
        'operationType': 'Create',
        'sfzx': dic['sfzx'],
        'jzdSheng.dm' : dic['jzdSheng']['dm'],
        'jzdShi.dm': dic['jzdShi']['dm'],
        'jzdXian.dm':dic['jzdXian']['dm'],
        'jzdDz': dic['jzdDz'],
        'jzdDz2': dic['jzdDz2'],
        'lxdh': dic['lxdh'],
        'tw': dic['tw'],
        'bz': dic['bz'],
        'dm': None,
        'brJccry.dm':dic['brJccry']['dm'],
        'brJccry1':dic['brJccry']['mc'],
        'brStzk.dm':dic['brStzk']['dm'],
        'brStzk1':dic['brStzk']['mc'],
        'dkd':dic['dkd'],
        'dkdz':dic['dkdz'],
        'dkly':dic['dkly'],
        'hsjc':dic['hsjc'],
        'jkm':dic['jkm'],
        'jrJccry.dm':dic['jrJccry']['dm'],
        'jrJccry1':dic['jrJccry']['mc'],
        'jrStzk1':dic['jrStzk']['mc'],
        'jrStzk.dm': dic['jrStzk']['dm'],
        'jzInd':dic['jzInd'],
        'tw1':dic['twM']['mc'],
        'twM.dm': dic['twM']['dm'],
        'xcm':dic['xcm'],
        'xgym':dic['xgym'],
        'yczk.dm':dic['yczk']['dm'],
        'yczk1': dic['yczk']['mc']
        }

        # 处理登入网页返回的数据   可以拿到一个 wap/main/welcome?_t_s_=1602909162783
        # _t_s_=1602909162783这个是 我们需要的
        add_url_data = recv['goto2'].replace('wap/main/welcome', '')
        print(add_url_data)
        # 提交打卡信息要用到的 请求头
        headers1 = {
            'Accept': '*/*',
            'Accept-Encoding': 'gzip, deflate',
            'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6',
            'Connection': 'keep-alive',
            'Host': 'xggl.hnie.edu.cn',
            "user-agent": "Mozilla / 5.0(Linux;Android 4.0.4;Galaxy Nexus Build / IMM76B) AppleWebKit / 535.19(KHTML, like Gecko)" + \
                          "Chrome / 18.0.1025.133 Mobile Safari / 535.19",
            'Referer': 'http://xggl.hnie.edu.cn/content/menu/student/temp/zzdk' + add_url_data,
            'Origin': 'http://xggl.hnie.edu.cn',
        }
        # 提交打卡信息的网站  后面要加上_t_s_=1602909162783
        clock_url = 'http://xggl.hnie.edu.cn/content/student/temp/zzdk' + add_url_data
        # 获取打卡后的信息
        clock_over_data = req.post(url=clock_url, headers=headers1, data=data).text
        # 将打卡后的信息 转为 json类型
        clock_over_data1 = json.loads(clock_over_data)
        print(clock_over_data1)
        # 这个text 是用来保存 发邮件的信息的
        text = ''
        # 获取当前时间
        now_time = time.localtime()
        # 如果打卡成功
        if clock_over_data1['result']:
            # 通知 用户 打卡成功
            text = '易班打卡' + '\nhello ' + str(ID[i]) + '\n' + '你的易班已经打卡成功' + '\n' + '{:d}年{:d}月{:d}日'.format(
                now_time.tm_year, now_time.tm_mon, now_time.tm_mday)
            print('ID' + str(ID[0]) + '已经打卡！！')
        else:
            # 如果打卡不成功
            error_message = re.compile('"message":"(.*?)"')
            # 获取打卡失败的信息
            message = re.findall(error_message, clock_over_data)[0]
            # 通知 用户 打卡失败 及失败原因
            text = '易班打卡' + '\nhello ' + str(ID[i]) + '\n' + '你的易班打卡失败' +'\n原因：'+ message + '\n' + '{:d}年{:d}月{:d}日'.format(
                now_time.tm_year, now_time.tm_mon, now_time.tm_mday)
            print('ID' + str(ID[i]) + '打卡失败')
        # 发送邮件
        to_meg(text, mailID[i])


if __name__ == "__main__":
    main()




