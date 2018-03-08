# coding:utf-8
import requests
from lxml import etree
from xlwt import Workbook

book = Workbook()
sheet1 = book.add_sheet("网贷平台数据")
# 写表头
for k, j in enumerate(["序号", "平台名称", "详情页链接", "评级", "发展指数","标签", "参考利率", "待还余额", "注册地",
                       "上线时间", "网友印象", "综合评分", "点评人数"]):
    sheet1.write(0, k, j)

num = 25
# for page in range(1, 77):
for page in range(1, 243):
    # url = "https://www.wdzj.com/dangan/search?filter=e1&currentPage={}".format(page)
    url = "https://www.wdzj.com/dangan/search?filter&currentPage={}".format(page)
    headers = {
        "Host": "www.wdzj.com",
        "Referer": "https://www.wdzj.com/dangan/",
        "Upgrade-Insecure-Requests": "1",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36",
    }
    response = requests.get(url, headers=headers)
    # print(response.text)
    tree = etree.HTML(response.text)
    terraceList = tree.xpath('//*[@id="showTable"]/ul/li')

    result = []
    for index, terrace in enumerate(terraceList):
        itemTitle = "".join(terrace.xpath('div[1]/h2/a/text()'))  # 平台名称
        detail_url = "https://www.wdzj.com" + "".join(terrace.xpath('div[1]/h2/a/@href'))  # 详情页链接
        itemTitleTag = terrace.xpath('div[1]/div/em/text()|\
                                       div[1]/div/em/strong/text()|\
                                       div[1]/div/ul/li/text()')  # 标签
        if len(itemTitleTag) > 4:
            if "评级：" == itemTitleTag[0]:
                pingji = itemTitleTag[1]
                fazhan = itemTitleTag[2]
                itemTitleTag = itemTitleTag[3:]
            else:
                pingji = "没有评级"
                fazhan = "没有发展指数"
        else:
            pingji = "没有评级"
            fazhan = "没有发展指数"
        itemTitleTag = "|".join(itemTitleTag)
        biaotag = "".join(terrace.xpath('div[2]/a/div[1]/label/em/text()'))  # 参考利率
        daihuan = "".join(terrace.xpath('div[2]/a/div[2]/text()'))  # 待还余额
        zhuce = "".join(terrace.xpath('div[2]/a/div[3]/text()'))  # 注册地
        shangxian = "".join(terrace.xpath('div[2]/a/div[4]/text()'))  # 上线时间
        comment = "|".join(terrace.xpath('div[2]/a/div[5]/span/text()'))  # 网友印象
        grade = "".join(terrace.xpath('div[2]/a/div[5]/strong/text()'))  # 综合评分
        com_num = "".join(terrace.xpath('div[2]/a/div[5]/em/text()'))  # 点评人数
        print(itemTitle, detail_url, pingji, fazhan, itemTitleTag, biaotag, daihuan, zhuce, shangxian, comment, grade,
              com_num)
        result.append(
            [itemTitle, detail_url, pingji, fazhan, itemTitleTag, biaotag, daihuan, zhuce, shangxian, comment, grade,
             com_num])

    for index, info in enumerate(result):
        for k, j in enumerate(info):
            # 添加序号 在第一列
            if k == 0:
                sheet1.write((page - 1) * num + index + 1, k, (page - 1) * num + index + 1)
            sheet1.write((page - 1) * num + index + 1, k + 1, j)
    book.save('网贷数据(包括停业及转型).xls')
