# -*- encoding=utf8 -*-
"""
1、表格以邮件形式呈现
2、加入pid信息

"""
import os, time, re, xlsxwriter
import smtplib
from email.mime.text import MIMEText
from email.header import Header
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from support.password import mail_pass
# adb等待时间(s)
adb_wait_time = 3
#测试间隔
test_inteval = 1
#每个间隔时间(s)
test_inteval_time = 5 - adb_wait_time
#测试总时长 = test_time_inteval * test_time_num
test_time_num = 1
# 如需要早上7点自动测试则 07，默认21:00
test_clock = "19"
"""
在设置的时间
1、盒子固定ip为28.1.88.212(/support/sup_mes.ini中有同样设置)
2、需要airtest支持，需要先安装airtest（本地路径：F:\AirtestIDE_2018-10-11_py3_win64\AirtestIDE）
"""
def have_a_meeting():
    with open("../support/sup_mes.ini") as fp:
        mes = fp.readline()
        mes = r'"' + mes + '"'

        os.system(mes)
        time.sleep(5)

#用于邮件发送时查找excel
global file_name_global

if __name__ == '__main__':
    time_global = str(time.strftime("%Y-%m-%d-%H-%M-%S", time.localtime()))
    file_name_global = '../temp/' + time_global + '.xlsx'

    workbook = xlsxwriter.Workbook(file_name_global)
    worksheet = workbook.add_worksheet()
    worksheet.set_column('A:A', 20)
    worksheet.set_column('B:B', 20)
    worksheet.set_column('C:C', 20)
    worksheet.set_column('D:D', 20)
    worksheet.set_column('E:E', 20)
    worksheet.write('A1', '时间')
    worksheet.write('B1', 'Unknown Pss')
    worksheet.write('C1', 'Unknown Private')
    worksheet.write('D1', 'Unknown Private')
    worksheet.write('E1', 'Unknown Swapped')
    while(1):
        time_now = time.strftime("%H", time.localtime())
        if  "21"== str(time_now) or test_clock == str(time_now):
            have_a_meeting()
            for i in range(0, test_time_num):
                os.system("adb kill-server")
                time.sleep(int(adb_wait_time))
                mem_mes_get = os.popen("adb connect 20.1.88.212 && adb shell dumpsys meminfo com.homedoor2.tvlauncher").read()
                time_now = time.strftime("%Y-%m-%d-%H-%M-%S", time.localtime())
                file_name ='../temp/' + str(time_now) + '.txt'

                with open(file_name, mode='w', encoding='utf-8') as fw:
                    fw.write(mem_mes_get)

                with open(file_name, mode='r', encoding='utf-8') as fr:
                    lines = fr.readlines()
                    for line in lines:
                        search_result = re.search(r'(.*) Unknown (.*?) .*', line, re.M | re.I)
                        if search_result:
                            unknow_mem = search_result.group().split()
                            worksheet.write('A' + str(i + 2), str(time_now))
                            worksheet.write('B' + str(i + 2), unknow_mem[1])
                            worksheet.write('C' + str(i + 2), unknow_mem[2])
                            worksheet.write('D' + str(i + 2), unknow_mem[3])
                            worksheet.write('E' + str(i + 2), unknow_mem[4])
                time.sleep(int(test_inteval) * (test_inteval_time))
            workbook.close()
            os.system("adb reboot")
        break

    # 第三方 SMTP 服务 smtp.mxhichina.com同一个账号一天只能发送3次
    mail_host = "smtp.mxhichina.com"  # 设置服务器
    mail_user = "wangwenpeng@meetsoon.cn"  # 用户名
    sender = 'wangwenpeng@meetsoon.cn'
    receivers = ['593469560@qq.com']  # 接收邮件，可设置为你的QQ邮箱或者其他邮箱

    message = MIMEMultipart()
    message['From'] = Header("wwp", 'utf-8')
    message['To'] = Header("zz", 'utf-8')

    subject = '内存测试结果'
    message['Subject'] = Header(subject, 'utf-8')

    message.attach(MIMEText('内存测试结果' + time_global, 'plain', 'utf-8'))

    att1 = MIMEApplication(open(file_name_global, 'rb').read())
    att1["Content-Type"] = 'application/octet-stream'
    # 这里的filename可以任意写，写什么名字，邮件中显示什么名字

    att1["Content-Disposition"] = 'attachment; filename="test_record.xlsx"'
    message.attach(att1)

    try:
        smtpObj = smtplib.SMTP()
        smtpObj.connect(mail_host, 25)  # 25 为 SMTP 端口号
        smtpObj.login(mail_user, mail_pass)
        smtpObj.sendmail(sender, receivers, message.as_string())
        print("send sucess")

    except smtplib.SMTPException:
        print("send fail")


