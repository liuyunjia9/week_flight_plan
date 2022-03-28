import datetime


def is_off_day(date):
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

    if date in off_date_dateformat:
        return '0'
    elif date in work_date_dateformat:
        return '1'
    elif date.isoweekday() == 6 or date.isoweekday() == 7:
        return '0'
    else:
        return '1'


while True:
    number = input("请输入代码：")
    if number == '0':
        break
    elif number == '1':
        date = input("请输入日期：")
        date1 = datetime.datetime.strptime(date, '%Y%m%d').date()
        print(is_off_day(date1))
    else:
        print('输入有误')





