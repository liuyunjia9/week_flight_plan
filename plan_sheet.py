import datetime
import xlwings


def is_off_day(date_info):
    # 休息日
    off_date = ['20220101', '20220102', '20220103', '20220131', '20220201', '20220202', '20220203', '20220204',
                '20220205', '20220206', '20220403', '20220404', '20220405', '20220430', '20220501', '20220502',
                '20220503', '20220504', '20220603', '20220604', '20220605', '20220910', '20220911', '20220912',
                '20221001', '20221002', '20221003', '20221004', '20221005', '20221006', '20221007']
    # 工作日
    work_date = ['20220129', '20220130', '20220402', '20220424', '20220507', '20221008', '20221009']
    # 转换格式
    off_date_dateformat = []
    for i in off_date:
        off_date_dateformat.append(datetime.datetime.strptime(i, '%Y%m%d').date())

    work_date_dateformat = []
    for i in work_date:
        work_date_dateformat.append(datetime.datetime.strptime(i, '%Y%m%d').date())

    if date_info in off_date_dateformat:
        return True
    elif date_info in work_date_dateformat:
        return False
    elif date_info.isoweekday() == 6 or date_info.isoweekday() == 7:
        return True
    else:
        return False


content = """
1.开始日期
0.结束
"""
content1 = """
默认今天吗？
1.是
2.自定义日期
"""
date_list = []
date = datetime.datetime.today().date()
while True:
    print(content)
    number = input("请输入代码：")
    if number == '0':
        break
    elif number == '1':
        a = input(content1)
        if a == '1':
            pass
        elif a == '2':
            date_str = input("请输入日期：")
            date = datetime.datetime.strptime(date_str, '%Y%m%d').date()
        else:
            print('输入有误')
        date_list = [(date + datetime.timedelta(days=i)) for i in range(0, 7)]
        break
    else:
        print('输入有误')

xw = xlwings.App(visible=False, add_book=False)
wb = xw.books.add()
wb_sht = wb.sheets.add('生产运行室')
# 第一行
wb_sht.range('a1').value = '客舱部空勤干部及借调空勤人员航班安排'
wb_sht.range('a1').font.size = 14
wb_sht.range('a1').font.bold = True
wb_sht.range('a1:j1').merge()
# 第二行
wb_sht.range('a2').value = '时间'
wb_sht.range('b2').value = min(date_list).strftime('%Y年%m月%d日')+"-"+max(date_list).strftime('%Y年%m月%d日')
wb_sht.range('b2:j2').merge()
# 第三行
wb_sht.range('a3:c3').value = ['单位', '姓名', '工作号']
wb_sht.range('D3:J3').value = date_list
wb_sht.range('D3:J3').number_format = 'AAAA'
for i in range(0, 7):
    if is_off_day(date_list[i]):
        wb_sht.range('D3:J3')[i].font.color = '#ff0000'
# 第四行
wb_sht.range('a4:c4').value = ['运行室', '樊雪', '32157']
# 第五行
wb_sht.range('a5:c5').value = ['运行室', '刘运嘉', '32418']
wb_sht.range('D5:J5').value = ['行政', '行政', '休息', '休息', '飞行', '飞行', '飞行', ]
wb_sht.range('D4:J5').api.Validation.Add(3, 1, 3, '飞行,休息,行政,值班,复训,休假,接机,隔离,培训')
# 第六行
wb_sht.range('a6').value = '备注：此表每周二11:00点前，以BQQ的方式发至黎晓处。填写内容：飞行、行政、培训、开会、航休'
wb_sht.range('a6:j7').merge()
wb_sht.range('a6').font.size = 10

wb_sht.range('a1:j7').font.name = '微软雅黑'
for i in range(7, 12):
    wb_sht.range('a2:j7').api.Borders(i).LineStyle = 1
wb_sht.range('a2:j7').api.Borders(12).LineStyle = 1
wb_sht.range('a1:j7').api.HorizontalAlignment = -4108
wb_sht.range('a1:j7').api.VerticalAlignment = -4130
wb_sht.autofit()

wb.save('c:/Users/liuyunjia9/Desktop/'+'生产运行室('+wb_sht.range('b2').value+').xlsx')
wb.close()
xw.quit()
print('文件已生成，程序结束')
