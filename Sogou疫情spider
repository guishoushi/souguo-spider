import requests, re, json
import xlsxwriter


def get_yq_info():
    url = "http://sa.sogou.com/new-weball/page/sgs/epidemic?type_page=pcpop"
    # 请求数据，看是否能够获得
    response = requests.get(url)
    response.encoding = "utf-8"
    # 匹配规则
    pattern_obj = re.compile('window.__INITIAL_STATE__ =(.*?)</script>')
    # 返回的数据是列表，所以自取第一项
    # findall() 返回字符串中所有和pattern_obj向匹配的全部字符串，返回的形式是列表
    res = re.findall(pattern_obj, response.text)[0]
    res_dict = json.loads(res)  # 将字符串转化为字典
    data = res_dict['data']['mapStats']['provinceDetail']  # 先取出data数据
    return data  # 返回数据


def get_guow_info():
    '''
    爬取国外的疫情数据
    :return: data 字典类型
    '''
    url = "http://sa.sogou.com/new-weball/page/sgs/epidemic?type_page=pcpop"
    # 请求数据，看是否能够获得
    response = requests.get(url)
    response.encoding = "utf-8"
    # 匹配规则
    pattern_obj = re.compile('window.__INITIAL_STATE__ =(.*?)</script>')
    # 返回的数据是列表，所以自取第一项
    # findall() 返回字符串中所有和pattern_obj向匹配的全部字符串，返回的形式是列表
    res = re.findall(pattern_obj, response.text)[0]
    res_dict = json.loads(res)  # 将字符串转化为字典
    data = res_dict['data']['overseas']  # 先取出data数据
    return data  # 返回数据


for i in get_guow_info():
    print(i['continents'], i['provinceName'], "当前存在确诊数量：", i['currentConfirmedCount'], "总确诊人数：", i['confirmedCount'],
          "治愈数量：", i['curedCount'], "死亡数量：", i['deadCount'])

# 保存数据 --- Excel
# 1，新建一个Excel文件
workbook = xlsxwriter.Workbook('疫情数据.xlsx')
# 2，创建一张表
sheet1 = workbook.add_worksheet("国内疫情数据")
sheet2 = workbook.add_worksheet("全球疫情数据")
# 3，写入数据,按行写入

sheet1.write_row("A1", ["地区", "确诊数量", "治愈数量", "死亡数量"])
sheet2.write_row("A1", ["地区", '国家', "当前确诊数量", "总确诊数量", "治愈数量", "死亡数量"])
# 4，写入每行的具体数据
number = 2  # 构造写入的第几行，控制写入的行数
for city in get_yq_info():  # 遍历数据
    city_list = city.split(' ')  # 把字符串分割为列表
    print(city_list)
    try:
        death_number = int(city_list[6])  # 如果没有死亡病例，捕获异常
    except Exception as e:
        death_number = 0  # 如果没有死亡病例，则默认为0
    # 依次写入每行的数据
    sheet1.write_row("A" + str(number), [city_list[0], int(city_list[2]), int(city_list[4]), death_number])
    number += 1  # 下一次写入数据的位置，应该是下一行，让number+1 到下一行去写入数据

# 写入国外数据
number1 = 2
for i in get_guow_info():
    sheet2.write_row("A" + str(number1),
                     [i['continents'], i['provinceName'], i['currentConfirmedCount'], i['confirmedCount'],
                      i['curedCount'], i['deadCount']])
    number1 += 1
# 最后一步,关闭文件才会保存文件
workbook.close()
