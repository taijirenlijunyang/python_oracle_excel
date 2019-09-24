import cx_Oracle
import datetime
import openpyxl

def user_info():
    print('请输入用户名和密码')
    user = input('请输入用户名:')
    pwd = input('请输入密码:')
    return user, pwd


def data_sql(user, pwd, name_list, sql_list,num):
    """
    导入确定的sql文件，去查询数据
    :param user: 用户名
    :param pwd: 用户密码
    :param name_list: 查询数据的名称列表
    :param sql_list: 查询sql的列表
    :return:
    """
    # finally_data = []
    queray_data = {}
    db = cx_Oracle.connect(user, pwd, 'ip:port/name')
    print('连接数据库完成')
    cursor = db.cursor()
    number = 1
    for i_name_sql in zip(name_list, sql_list):
        cursor.execute(i_name_sql[1] % (num,))
        result = cursor.fetchall()
        # for data in result:
        #     for clear_date in data[1:]:
        #         finally_data.append(clear_date)
        queray_data[i_name_sql[0]] = result
        print('查询%s---第%d完成' % (i_name_sql[0], number))
        number += 1



    cursor.close()
    db.close()
    print('数据查询完成')
    return queray_data


def select_time(num):
    """
    :param num: 当天前休息了几天
    :return: 直接print时间范围
    """
    today = datetime.date.today()
    if num == 0:
        time_slot = datetime.timedelta(days=num)
        yesterday = today - datetime.timedelta(days=1)
        print('查询时间为{}'.format(yesterday))

    else:
        time_slot = datetime.timedelta(days=num)
        queray_time_start = today - time_slot
        yesterday = today - datetime.timedelta(days=1)
        print('查询时间范围为从{}到{}'.format(queray_time_start, yesterday))



class Write_excel(object):
    '''修改excel数据'''
    def __init__(self, filename):
        self.filename = filename
        self.wb = openpyxl.load_workbook(self.filename)
        self.ws = self.wb.active  # 激活sheet

    def write(self, row_n, col_n, value):
        self.ws.cell(row_n, col_n, value)

    def save(self):
        self.wb.save(self.filename)

def main(num):
    """
    :param num: 查询时间开始距离今天的天数
    :return:
    """
    name_list = []
    sql_list = []
    with open('sql0.txt') as file:
        data = file.readlines()
        for i_data in data:
            name_list.append(i_data.split(':')[0])
            sql_list.append(i_data.split(':')[1].replace('\n', ''))
    print('查询sql一共%d条'%(len(name_list),))
    # if num == 0:
    #     with open('sql0.txt') as file:
    #         data = file.readlines()
    #         for i_data in data:
    #             # print(i_data)
    #             name_list.append(i_data.split(':')[0])
    #             sql_list.append(i_data.split(':')[1].replace('\n', ''))
    # else:
    #     with open('sql0.txt') as file:
    #         data = file.readlines()
    #         for i_data in data:
    #             name_list.append(i_data.split(':')[0])
    #             sql_list.append(i_data.split(':')[1].replace('\n', ''))
    user, pwd = user_info()
    queray_data= data_sql(user, pwd, name_list, sql_list,num)
    wb = Write_excel('data_integrity.xlsx')
    col_num = 5
    len_data = 0
    start_row_num =1
    for i_name in name_list:  # 遍历sql对应的名称
        name_data = queray_data[i_name]  # 得到数据源名称对应的所有数据
        # print(name_data)
        for i_data in name_data:  # 遍历数据中各个元组
            row_num = start_row_num
            len_data = len(i_data[1:])
            for write_data in i_data[1:]:  # 取出元组中数据
                wb.write(row_num, col_num, write_data)
                # print(row_num, col_num, write_data)
                row_num += 1
            col_num += 1
        start_row_num = start_row_num + len_data
        col_num = 5

    wb.save()
    print('数据写入完成')


if __name__ == '__main__':
    start_timie = datetime.datetime.now()
    num = int(input('查询几天前到昨天的数据'))
    select_time(num)
    main(num)
    #print(datetime.datetime.strftime(result[0], '%Y-%m-%d %H-%M-%S'))
    end_time = datetime.datetime.now()
    print(end_time-start_timie)
