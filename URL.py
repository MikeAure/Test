from urllib import request, parse
import requests
from bs4 import BeautifulSoup
from selenium import webdriver
import xlsxwriter
import time
import datetime
import DATA
from selenium.webdriver.remote.webelement import WebElement


def _to_chinese4(num):
    assert (0 <= num and num < DATA._S4)
    if num < 20:
        return DATA._MAPPING[num]
    else:
        lst = []
        while num >= 10:
            lst.append(num % 10)
            num = num // 10
        lst.append(num)
        c = len(lst)  # 位数
        result = u''

        for idx, val in enumerate(lst):
            val = int(val)
            if val != 0:
                result += DATA._P0[idx] + DATA._MAPPING[val]
                if idx < c - 1 and lst[idx + 1] == 0:
                    result += u'零'
        return result[::-1]


def login1():
    # user={"username":"成都理工大学","password":12345678,"ip":"103.27.25.64---%E5%9B%BD%E5%86%85%E6%9C%AA%E8%83%BD%E8%AF%86%E5%88%AB%E7%9A%84%E5%9C%B0%E5%8C%BA"}
    # response=requests.post("http://dxx.scyol.com/dxxBackend/manage/login/checkLogin.html",headers=headers,data=user)
    # print(response.content.decode("utf-8"))
    driver = webdriver.Chrome()
    url = "http://dxx.scyol.com/dxxBackend/manage/login/logout"
    driver.get(url)
    username = driver.find_element_by_id('userName')
    password = driver.find_element_by_id('password')
    logbutton = driver.find_element_by_xpath("//a[contains(@class,'btn btn-primary block full-width m-b')]")
    # loginbutton=driver.find_element_by_xpath()
    username.send_keys("成都理工大学")
    password.send_keys(12345678)
    logbutton.click()


def login2():  # 登录获取cookie
    Login_information = {"username": "成都理工大学", "password": "cdlgdx12345678", "ip": "103.27.25.64---国内未能识别的地区"}
    Login_url = "http://dxx.scyol.com/dxxBackend/manage/login/checkLogin.html"
    sender = requests.session()
    resourse = sender.post(Login_url, headers=DATA.headers, data=Login_information)
    print(resourse.content.decode('utf-8'))
    Indexpage_url = "http://dxx.scyol.com/dxxBackend/manage/Index/index.html"
    sender.get(Indexpage_url) #获取index页
    Resource_url = "http://dxx.scyol.com/dxxBackend/manage/count/index.html?pid=6438" #理工后台数据
    sender.get(Resource_url) #利用从index页获取的headers进入资源页
    return sender


def get_stage_default(sender, title_default):  # 获取最新一期数据
    # title_default
    url = "http://dxx.scyol.com/dxxBackend/manage/count/getstages.html"
    stage_data_default = {'title': title_default}
    stagenum = sender.post(url, headers=DATA.get_stage_headers, data=stage_data_default)
    stagenum = stagenum.json()[0]['id']
    return stagenum


def get_stage_input(sender, title_input, num):  # 获取之前的数据
    url = "http://dxx.scyol.com/dxxBackend/manage/count/getstages.html"
    get_stage_data = {'title': title_input}
    stagenum = sender.post(url, headers=DATA.get_stage_headers, data=get_stage_data)
    stagenum=stagenum.json()
    stageresult=str()
    for x in stagenum:
        if x['snum']==int(num):
            stageresult=x['id']
    return stageresult


# 从网页源码中获取最新一期的title
def get_newest_title():
    default_title = str()
    pass
    return default_title  # 返回最新的title


def query_stage(sender, title_default):  # 查询页面的stage
    stage_num_result = str()
    question = input('是否需要最新一期？ y/n')
    if question == 'y':
        stage_num_result = get_stage_default(sender, title_default=title_default)
        return stage_num_result
    if question == 'n':
        title_inputs = input("请输入想要获取数据的具体时间（输入格式：第x季第y期）：")
        title_upload = title_inputs[0:3]
        need_num = title_inputs[4]
        stage_num_result = get_stage_input(sender, title_upload, need_num)
        return stage_num_result


def get_data(preheadres,sender, stage):  # 获取需要的数据
    url = "http://dxx.scyol.com/dxxBackend/manage/Count/getData.html"
    data = {"pid": "6438", "stage": stage, 'debug': '1', 'page': '1', 'limit': '30'}
    result = sender.post(url, data=data,headers=preheadres)

    if result.status_code==502:
        raise Exception("服务器内部错误！")

    # with open('dxxdata(pre).txt', 'w', encoding='utf-8') as fp:
    #     fp.write(str(result.json()))
    else:
        return result.json()


def get_data_ee(preheader, sender, stage):  # 获取需要的数据
    url = "http://dxx.scyol.com/dxxBackend/manage/Count/getData.html"
    for x in DATA.idlist.keys():
        for i in range(1,20):
            data = {"pid": DATA.idlist[x], "stage": stage, 'debug': '1', 'page': str(i), 'limit': '10'}
            result = sender.post(url, data=data, headers=preheader)
            with open(x+str(i)+'.txt', 'w', encoding='utf-8') as fp:
                fp.write(str(result.json()))


def get_tomorrow_date():
    tomorrow_date = (datetime.date.today() + datetime.timedelta(days=+1)).strftime('%m%d')
    temp = list(tomorrow_date)
    del (temp[0])
    temp.insert(1, '.')
    tomorrow_date = ''.join(temp)
    return tomorrow_date


def data_process(pro_data, sumlst):
    result_data = {}
    for i in pro_data['data']:
        if (i['org_name'] != '沉积地质研究院'):
            temp1 = dict(学习人数=str(i['current_count']))
            result_data.setdefault(i['org_name'][0:-2], temp1)
    for x in sumlst.keys():
        if (x in result_data.keys()):
            result_data[x].setdefault('团员总数', DATA.sumlist[x])
            percent = int(result_data[x]['学习人数']) / int(result_data[x]['团员总数'])
            percent = '{:.2%}'.format(percent)
            result_data[x].setdefault('比例', percent)
    with open('dxxdata.txt', 'w', encoding='utf-8') as f:
        f.write(str(result_data))

    return result_data


def export_to_excel(result_data, excel_title, tomorrow):
    head = ['学院', '学习人数', '团员总人数', '比例', '排名']
    reportbook = xlsxwriter.Workbook('青年大学习' + excel_title + '各院学习情况统计表' + ' ' + '(' + tomorrow + ')' + '.xlsx')
    report = reportbook.add_worksheet("sheet1")
    row = 0;
    col = 0
    row_num = 2
    col_num = 0
    title_format = reportbook.add_format({'align': 'center', 'bold': 'true'})
    text_format = reportbook.add_format({'align': 'center'})
    report.merge_range('A1:E1', '青年大学习' + excel_title + '各院学习情况统计表' + ' ' + '(' + tomorrow + ')', title_format)
    report.write_row('A2', head, text_format)
    for x in result_data.keys():
        report.write(row_num, col_num, x, text_format)
        col_num += 1
        for y in result_data[x].values():
            report.write(row_num, col_num, y, text_format)
            col_num += 1
        col_num = 0
        row_num += 1
    reportbook.close()


def get_excel_weekend(sender, stagenum, title, date):
    for x in DATA.idlist.keys():
        url5 = 'http://dxx.scyol.com/dxxBackend/manage/count/export_excel?pid=' + DATA.idlist[x] + '&stage=' + str(
            stagenum) + '&debug=1'

        File = sender.get(url5)
        with open('[细]' + x[0:-2] + '第二期' + '.xlsx', 'wb') as fq:
            fq.write(File.content)


def get_excel_weekday(sender, stagenum, title, date):
    for x in DATA.idlist.keys():
        url5 = 'http://dxx.scyol.com/dxxBackend/manage/count/export_excel?pid=' + DATA.idlist[x] + '&stage=' + str(
            stagenum) + '&debug=1'
        File = sender.get(url5)
        with open(title+' ' + x[0:-2] +' '+date+ '.xlsx', 'wb') as fq:
            fq.write(File.content)


def main():
    tomorrow_date = get_tomorrow_date()
    sender = login2()
    stagenum = query_stage(sender, title_default='第八季')
    try:
        get_data(DATA.headers,sender, stagenum)
    except:
        print("请重启该脚本")
    else:
        prodata = get_data(DATA.headers,sender, stagenum)
        result = data_process(prodata, DATA.sumlist)
        title = input('请输入第几季第几期')
        export_to_excel(result, title, tomorrow_date)
        que = input('是否需要表2？: y/n ')
        if que == 'y':
            get_excel_weekday(sender, stagenum, title, tomorrow_date)



if __name__ == "__main__":
    print("Hello,worlds!")
    print("Vim is so powerful!")
