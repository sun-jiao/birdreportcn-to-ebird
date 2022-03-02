import csv
from datetime import datetime

import xlrd


def convert(infile, locationfile, outfile, allow_record_id=True, allow_software_info=True):
    # 因为记录中心导出的记录没有地点，因此我们需要手动从网页上复制到一个txt文件中，供转换器读取

    f = open(locationfile, encoding='utf-8')  # 读取地点文件
    line = f.readline()  # 逐行读取
    index = 0
    record_id = 0  # 记录编号
    locations = {}  # 地点字典
    while line:
        line = line.replace('\n', '').replace('\r', '')
        if (index % 2) == 0:
            record_id = line
        else:
            locations[record_id] = line

        index = index + 1
        line = f.readline()

    f.close()

    inexcel = xlrd.open_workbook(infile)  # 读取excel文件
    table = inexcel.sheets()[0]  # 通过索引顺序获取sheet
    # 文件内容检查
    if table.ncols < 12:
        print('文件内容错误，请检查输入文件。')
        return
    nrows = table.nrows
    if not (table.cell_value(0, 0) == '活动编号' and table.cell_value(0, 1) == '观测开始时间' and
            table.cell_value(0, 2) == '观测结束时间' and table.cell_value(0, 3) == '中文名' and
            table.cell_value(0, 4) == '学名' and table.cell_value(0, 5) == '省' and
            table.cell_value(0, 6) == '州/市' and table.cell_value(0, 7) == '区/县' and
            table.cell_value(0, 8) == 'IUCN受胁级别' and table.cell_value(0, 9) == '国家保护等级' and
            table.cell_value(0, 10) == 'CITES保护等级' and table.cell_value(0, 11) == '鸟种数量'):
        print('文件内容错误，请检查输入文件。')
        return

    # ebird 文件格式：
    # 0 俗名，(1 属名，2 种名),3 数量，(4 鸟种备注)，5 地点，(6 纬度，7 经度)，8 日期，9 开始时间，
    # 10 省，11 国家，(12 定点/行进，13 人数)，14 持续时间，(15 记录是否完整，16 行进距离，17 Effort area acres，18 记录备注)
    new_csv_file = open(outfile, 'a', encoding='utf-8', newline='')
    writer = csv.writer(new_csv_file)
    items = ['' for _ in range(19)]

    for row in range(1, nrows):
        items[0] = table.cell_value(row, 3)  # 0
        # items[1] = table.cell_value(row, 4).split(' ')[0]  # 1
        items[2] = table.cell_value(row, 4)  # 2
        amount = int(table.cell_value(row, 11))
        if amount == 0:
            amount = 1
            items[4] = 'heard'
        else:
            items[4] = ''

        items[3] = ("%.0f" % amount)  # 3

        recordid = table.cell_value(row, 0)
        items[5] = locations[recordid]  # 5
        start = datetime.strptime(table.cell_value(row, 1), '%Y-%m-%d %H:%M:%S')
        end = datetime.strptime(table.cell_value(row, 2), '%Y-%m-%d %H:%M:%S')
        items[8] = start.strftime('%m/%d/%Y')  # 8
        items[9] = start.strftime('%H:%M')  # 9
        delta = end - start
        minutes = delta.seconds // 60
        if minutes < 1:
            minutes = 1
        items[14] = str(minutes)  # 14
        items[10] = province_convert(table.cell_value(row, 5))  # 10
        items[11] = 'CN'  # 11

        items[12] = 'stationary'
        items[13] = '1'
        items[15] = 'Y'
        if allow_record_id:
            items[18] = 'Originally uploaded to birdreport.cn , record id = ' + recordid + '. '
            if allow_software_info:
                items[18] = items[18] + ' Transfered using birdreportcn-to-ebird developed by Sun Jiao. ' \
                                        'Source code: https://github.com/sun-jiao/birdreportcn-to-ebird . '

        writer.writerow(items)


def province_convert(province):
    provdict = {
        '北京市': 'BJ',
        '上海市': 'SH',
        '天津市': 'TJ',
        '重庆市': 'CQ',
        '河北省': 'HE',
        '山西省': 'SX',
        '内蒙古自治区': 'NM',
        '辽宁省': 'LN',
        '吉林省': 'JL',
        '黑龙江省': 'HL',
        '江苏省': 'JS',
        '浙江省': 'ZJ',
        '安徽省': 'AH',
        '福建省': 'FJ',
        '江西省': 'JX',
        '山东省': 'SD',
        '河南省': 'HA',
        '湖北省': 'HB',
        '湖南省': 'HN',
        '广东省': 'GD',
        '广西壮族自治区': 'GX',
        '海南省': 'HI',
        '四川省': 'SC',
        '贵州省': 'GZ',
        '云南省': 'YN',
        '西藏自治区': 'XZ',
        '陕西省': 'SN',
        '甘肃省': 'GS',
        '青海省': 'QH',
        '宁夏回族自治区': 'NX',
        '新疆维吾尔族自治区': 'XJ',
        '台湾省': 'TW',
        '香港特别行政区': 'HK',
        '澳门特别行政区': 'MO'
    }
    return provdict.get(province)


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    while True:
        try:
            infile = input('记录中心导出文件（回车结束输入）：')
            locationfile = input('观测点文件：')
            bool1 = False
            bool2 = False
            print('您可以选择将记录中心原始记录编号写入ebird的记录备注中，我们也希望显示其是使用本软件转换而成，示例如下：\n'
                  'Originally uploaded to birdreport.cn , record id = 2020010100001.  '
                  'Transfered using birdreportcn-to-ebird developed by Sun Jiao. '
                  'Source code: https://github.com/sun-jiao/birdreportcn-to-ebird .')
            allow1 = input('是否写入记录中心的记录编号（Y：写入，N：不写入）：')
            while True:
                if allow1 == 'Y':
                    bool1 = True
                    break
                elif allow1 == 'N':
                    bool1 = False
                    break
                else:
                    allow1 = input('输入内容不符，请重新输入（Y：写入，N：不写入）：')
                    continue
            if bool1:
                allow2 = input('是否写入软件信息（Y：写入，N：不写入）：')
                while True:
                    if allow2 == 'Y':
                        bool2 = True
                        break
                    elif allow2 == 'N':
                        bool2 = False
                        break
                    else:
                        allow2 = input('输入内容不符，请重新输入（Y：写入，N：不写入）：')
                        continue

            outfile = infile.split('.')[0] + '_out.csv'
            convert(infile, locationfile, outfile, bool1, bool2)
        except Exception as e:
            print(e)
        finally:
            test = input('回车继续转换新文件，输入任意内容后回车退出程序\n')
            if test != '':
                break
            else:
                continue
# See PyCharm help at https://www.jetbrains.com/help/pycharm/
