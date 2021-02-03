import requests
import datetime
import msvcrt
import xlrd

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
                  "AppleWebKit/537.36(KHTML, like Gecko) Chrome/83.0.4103.9"
                  "7 Safari/537.36 Edg/83.0.478.45"
}
time = datetime.date.today()
username = ""   #数字石大工号
password = ""   #数字石大密码
group_id = ""   #学院id


def login(username, password):
    data = {
        "username": username,
        "password": password
    }
    session = requests.Session()
    login_url = "https://app.upc.edu.cn/uc/wap/login/check"
    response = session.post(url=login_url, headers=headers, data=data, timeout=20)
    response.encoding = "UTF-8"
    print("\n登录:", eval(response.text)["m"])
    return session


def get_num(session):
    data = {
        "type": "weishangbao",
        "group_type": "1",
        "date": str(time),
        "group_id": group_id
    }
    num_url = "https://app.upc.edu.cn/ncov/wap/upc/diff"
    response = session.post(url=num_url, headers=headers, data=data, timeout=20)
    response.encoding = "UTF-8"
    print("\n今日未填报：" + str(eval(response.text)["d"]["today"]) + "人\n")
    return eval(response.text)["d"]["today"]


def get_list(session):
    type = "weishangbao"
    list_url = "https://app.upc.edu.cn/ncov/wap/upc/ulists?date=" + str(
        time) + "&type=" + type + "&page=1&page_size=390&group_id=" + group_id + "&keywords=&group_type=1"
    response = session.get(url=list_url, headers=headers, timeout=20)
    response.encoding = "UTF-8"
    lis = eval(response.text)["d"]["lists"]
    l = [str(i["xgh"]) + " " + i["realname"] for i in lis]
    print("今日未填报：\n" + "\n".join(l))


def where(session):
    url = "https://app.upc.edu.cn/ncov/wap/upc/export-download?group_id="+group_id+"&group_type=1&type=tianbao&date="+str(time)
    response = session.get(url=url, headers=headers, timeout=20)
    with open("temp.xlsx", "wb") as code:
        code.write(response.content)
    xl = xlrd.open_workbook("temp.xlsx")
    sheet = xl.sheets()[0]
    lis = sheet.col_values(25)
    lis.pop(0)
    for i in range(len(lis)):
        lis[i] = lis[i].split(" ")[0]
    l = [[i,lis.count(i)] for i in set(lis)]
    print("="*30+"\n所有地区\n"+"-" * 30)
    for i in l:
        print(format(i[0],"<12"),"\t",str(i[1])+"个"+"\n"+"-"*30)
    print("\n")
    return l


def main_area(session):
    dic = {}
    l = where(session)
    #names列表中存放重点地区名称
    names = ["北京市","河北省","黑龙江省","吉林省","辽宁省","天津市","山西省","内蒙古自治区","陕西省","河南省"]
    for i in l:
        if i[0] in names:
            dic.update({i[0]:i[1]})
    s = str(sum(dic.values()))
    print("="*30,format("\n重点地区","<12"),"\t",s+"个\n"+"-"*30)
    for name in names:
        print(format(name,"<12"),"\t",str(dic[name])+"个\n"+"-"*30)


s = login(username, password)
if get_num(s) != 0:
    get_list(s)
else:
    main_area(s)

print("\n按任意键退出...")
q = ord(msvcrt.getch())