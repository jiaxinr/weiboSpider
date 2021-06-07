#-*- codeing=utf-8 -*-
# author:           夹心
# pc_type           HUAWEI
# create_time:      2021/6/3 2:50
# file_name:        weiboUserSearch.py
# github:           https://github.com/jiaxinr

from bs4 import BeautifulSoup
import urllib.parse,urllib.request,urllib.error
import xlwt
import re
import xlrd

namelist=[]
file_location = r'C:\Users\Jiaxinr\Desktop\name.xls'#todo 在单引号内输入自己希望导入的表格地址
data = xlrd.open_workbook(file_location)
sheet = data.sheet_by_index(0)

#print(sheet.nrows)

for cnt in range(1,132):#todo 输入希望选取的范围
    #print(sheet.cell_value(cnt+1, 0))
    namelist.append(sheet.cell_value(cnt+1,0))#todo 括号内的参数为(行,列)，可根据需要调整

#print(namelist)
#exit()


#namelist = ["怪你過分美麓","pixie_媛"]  #todo 手动输入搜索用户名合集，如使用将会覆盖上面从表格中导入的数据，使用时请将此行最开始的“#”删去
nameNum = len(namelist)
book = xlwt.Workbook(encoding="utf- 8",style_compression=0)  # 创建workbook对象
seet = book.add_sheet('sex',cell_overwrite_ok=True)  # 创建工作表

def weiboSearch():
    # 爬取网页 逐一解析数据 保存数据
    baseurl1="https://s.weibo.com/user?q="
    baseurl2="&Refer=SUer_box"
    savepath = ".\\新浪搜索try3.xls"#todo 结果保存路径，可自定义

    for i in range(0, nameNum):
        datalist = getData(baseurl1,baseurl2,i)
        saveData(datalist,i)
        if(i%10==0):
            print("本次已累计爬取用户数量为："+str(i))

    book.save(savepath)

#正则表达式匹配
findSex=re.compile(r'icon-sex-(.*?)"></i>')#性别正则
findFollowings=re.compile(r'click:user_friends">(.*?)</a>')#关注正则
findFollowers=re.compile(r'click:user_fans">(.*?)</a>')#粉丝正则


#爬取网页
def getData(baseurl1,baseurl2,i):

    url=baseurl1+urllib.request.quote(namelist[i])+baseurl2
    html=askURL(url)
    soup=BeautifulSoup(html,"html.parser")

    datalist = []
    data=[]
    # print(re.findall(r'card-no-result',html))
    # exit()
    if len(re.findall(r'card-no-result',html))!=0:#如果该用户已改名，将在第二栏显示“找不到该用户”
        datalist.append("找不到该用户")
        datalist.append("")
        datalist.append("")
        return datalist
    sex=re.findall(findSex,html)[0]
    #print("sex:"+str(sex))
    datalist.append(sex)
    followings=re.findall(findFollowings,html)[0]
    datalist.append(followings)
    followers=re.findall(findFollowers,html)[0]
    datalist.append(followers)


    #print("datalist:"+str(datalist))
    return datalist


#得到指定一个url的网页内容
def askURL(url):
    head={
        "User-Agent": "Mozilla / 5.0(Windows NT 10.0;Win64;x64) AppleWebKit / 537.36(KHTML, like Gecko) Chrome / 87.0.4280.67 Safari / 537.36 Edg / 87.0.664.47"
    } #用户代理 伪装浏览器

    request=urllib.request.Request(url,headers=head)
    html=""
    try:
        response=urllib.request.urlopen(request)
        html=response.read().decode("utf-8")
        #print("html"+str(html))
    except urllib.error.URLError as e:
        if hasattr(e,"code"):
            print(e.code)
        if hasattr(e,"reason"):
            print(e.reason)
    return html


def saveData(datalist,i):
    seet.write(0,0,"用户名")
    seet.write(0,1,"性别") #行 列 内容
    seet.write(0,2,"关注数")
    seet.write(0,3,"粉丝数")

    seet.write(1+i,0,namelist[i])
    num = len(datalist)
    for j in range(0, num):
        data = datalist[j]
        seet.write(1+i, j+1, data)


if __name__ == '__main__':
    weiboSearch()