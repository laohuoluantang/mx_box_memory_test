# -*- encoding=utf8 -*-

import os, time, re, xlsxwriter
#测试间隔（min）
test_time_inteval = 5
#测试总时长 = test_time_inteval * test_time_num
test_time_num = 24
# 如需要早上7点自动测试则 07，默认21:00
test_clock = "10"
"""
在设置的时间
1、盒子固定ip为28.1.88.212(/support/sup_mes.ini中有同样设置)
2、需要airtest支持，需要先安装airtest（本地路径：F:\AirtestIDE_2018-10-11_py3_win64\AirtestIDE）
3、邮件发送被防火墙拦截，暂时不弄了
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
    file_name_global = '../temp/' + str(time.strftime("%Y-%m-%d-%H-%M-%S", time.localtime())) + '.xlsx'
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
                #adb等待时间
                adb_wait_time = 3
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
                time.sleep(int(test_time_inteval) * (60 - adb_wait_time))
            workbook.close()
            os.system("adb reboot")
        break



