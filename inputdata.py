# coding:utf-8

import xlrd
import os
import pymysql
from datetime import datetime
import time
from xlrd import xldate_as_tuple


def find(url):
    # data = xlrd.open_workbook("/Users/nihui/Documents/鲁文环---肃州区村干部信息采集表（果园镇镇丁家闸村）.xls")  ## 读取D盘中名为sales_data的excel表格
    data = xlrd.open_workbook(url)
    table_one = data.sheet_by_index(0)  ## 根据sheet索引获取sheet的内容
    table_two = data.sheet_by_index(2)
    table_three = data.sheet_by_index(3)
    table_four = data.sheet_by_index(4)
    table_five = data.sheet_by_index(1)

    name = table_one.cell_value(1, 1)

    find_table_test(table_one)
    # find_table_one(table_one)
    # find_table_tow(name, table_two)
    # find_table_three(name, table_three)
    # find_table_four(name,table_four)
    # find_table_five(name,table_five)




def find_table_five(user_name, table_five):
    # # 连接mysql
    # db = pymysql.connect("localhost", "root", "root", "cgb-layui", use_unicode=True, charset="utf8")
    db = pymysql.connect("139.199.111.126", "root", "ZhiKu0713@2020", "cgb-layui", use_unicode=True, charset="utf8")
    # village = table_two.cell_value(5, 3)
    sql = """SELECT id,name FROM tb_user WHERE name = '%s'""" % (user_name)
    cursor = db.cursor();
    cursor.execute(sql)
    results = cursor.fetchall()

    if not results:
        user_id = 0
        user_name = 0
    else:
        user_id = results[0][0]
        user_name = results[0][1]

        print(results[0][0], results[0][1])

    rows = table_five.nrows

    if rows > 2:
        if str(table_five.cell_value(2, 2)).strip() != '':
            for nrows_one in range(2, int(table_five.nrows)):


                start_time = table_five.cell_value(nrows_one, 1)
                print("开始时间",start_time)

                end_time_s = table_five.cell_value(nrows_one, 2)
                if end_time_s == "":
                    end_time = time.strftime('%Y%m%d%H%M%S', time.localtime(time.time()))
                else:
                    end_time = end_time_s
                print("结束时间" , end_time)
                typed = table_five.cell_value(nrows_one, 3)
                ## 职位：0-第一书记，1-村支部书记，2-村主任，3-副村支部书记，4-副村主任，5-妇女主任，6-文书，7-监委会主任，8-村书记兼主任 ，9-村总支书记，10-村总支副书记',
                if typed == "村党总支书记、主任":
                    type = 0
                elif typed == "村党支部书记、主任":
                    type = 1
                elif typed == "村党总支副书记":
                    type = 2
                elif typed == "村党支部副书记":
                    type = 3
                elif typed == "村委会副主任":
                    type = 4
                elif typed == "文书":
                    type = 5
                else:
                    type = 11

                print("职位",type)
                remak = "无"
                create_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))

                sql = "insert into tb_education (user_id,user_name,begin_time,end_time,education,remark,create_time) " \
                      " values (%d,'%s',%s,%s,%s,'%s','%s')" % \
                      (user_id, user_name,start_time,end_time,type,typed,create_time)
                print(sql)
                try:
                    # 使用 cursor() 方法创建一个游标对象 cursor
                    cursor = db.cursor()
                    cursor.execute(sql)
                except Exception as e:
                    # 发生错误时回滚
                    db.rollback()
                    print(str(e))
                else:
                    db.commit()  # 事务提交
                    print('事务处理成功')
                    print("\n")

                print(start_time, end_time,typed )





def find_table_four(user_name, table_four):
    # # 连接mysql
    db = pymysql.connect("localhost", "root", "root", "cgb-layui", use_unicode=True, charset="utf8")
    # village = table_two.cell_value(5, 3)
    sql = """SELECT id,name FROM tb_user WHERE name = '%s'""" % (user_name)
    cursor = db.cursor();
    cursor.execute(sql)
    results = cursor.fetchall()

    if not results:
        user_id = 0
        user_name = 0
    else:
        user_id = results[0][0]
        user_name = results[0][1]

        print(results[0][0], results[0][1])

    rows = table_four.nrows

    # print(rows)
    # print(str(table_three.cell_value(1, 1)).strip() != '')
    if rows > 2:

        if str(table_four.cell_value(2, 2)).strip() != '':
            for nrows_one in range(2, int(table_four.nrows)):

                # if str(table_four.cell_value(2, 0)).strip() != '':
                #     # print("索引",nrows_one,table_three.cell_value(nrows_one, 0))
                #     index = int(table_four.cell_value(nrows_one, 0))
                # else:
                #     index = 0;

                start_time = table_four.cell_value(nrows_one, 1)

                title = table_four.cell_value(nrows_one, 2)
                party_representative = 0
                npc_deputies = 0
                cppcc = 0
                if "党" in title:
                    party_representative = 1
                elif "人大" in title:
                    npc_deputies = 1
                elif "政协" in title:
                    cppcc = 1
                else:
                    party_representative = 0
                    npc_deputies = 0
                    cppcc = 0

                content = table_four.cell_value(nrows_one, 3)
                create_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))


                sql = "insert into tb_politic (user_id,user_name,party_representative,npc_deputies,cppcc,remark,create_time) " \
                      " values (%d,'%s',%d,%d,%d,'%s','%s')" % \
                      (user_id, user_name, party_representative, npc_deputies,cppcc, content, create_time)
                print(sql)
                try:
                    # 使用 cursor() 方法创建一个游标对象 cursor
                    cursor = db.cursor()
                    cursor.execute(sql)
                except Exception as e:
                    # 发生错误时回滚
                    db.rollback()
                    print(str(e))
                else:
                    db.commit()  # 事务提交
                    print('事务处理成功')
                    print("\n")

                print(start_time, title, content)





def find_table_three(user_name, table_three):
    # # 连接mysql
    db = pymysql.connect("139.199.111.126", "root", "ZhiKu0713@2020", "cgb-layui", use_unicode=True, charset="utf8")
    # village = table_two.cell_value(5, 3)
    sql = """SELECT id,name FROM tb_user WHERE name = '%s'""" % (user_name)
    cursor = db.cursor();
    cursor.execute(sql)
    results = cursor.fetchall()

    if not results:
        user_id = 0
        user_name = 0
    else:
        user_id = results[0][0]
        user_name = results[0][1]

        print(results[0][0], results[0][1])

    rows = table_three.nrows

    # print(rows)
    # print(str(table_three.cell_value(1, 1)).strip() != '')
    if rows > 2:

        if str(table_three.cell_value(2, 2)).strip() != '':
            for nrows_one in range(2, int(table_three.nrows)):

                if str(table_three.cell_value(2, 0)).strip() != '':
                    # print("索引",nrows_one,table_three.cell_value(nrows_one, 0))
                    index = int(table_three.cell_value(nrows_one, 0))
                else:
                    index = 0;

                start_time = table_three.cell_value(nrows_one, 1)
                start_time = str(start_time);
                start_time.strip();
                print(len(start_time))
                if (len(start_time) == 8):
                    begin_time = start_time[0:4] + '-' + start_time[4:6] + '-' + start_time[6:8] + " 00:00:00"
                elif len(start_time) == 7:
                    begin_time = start_time[0:4] + '-' + start_time[4:6] + '-' + start_time[6:7] + " 00:00:00"
                elif len(start_time) == 6:
                    begin_time = start_time[0:4] + '-' + start_time[4:6] + '-01' + " 00:00:00"
                else:
                    begin_time = "1971-01-01 00:00:00"

                title = table_three.cell_value(nrows_one, 1)
                content = table_three.cell_value(nrows_one, 2)
                resultstr = table_three.cell_value(nrows_one, 3)

                # 培训结果 0-不及格，1-及格，2-良好，3-优秀
                if "不及格" in resultstr:
                    result = 0
                elif "及格" in resultstr:
                    result = 1
                elif "良好" in resultstr:
                    result = 2
                elif "优秀" in resultstr:
                    result = 3
                else:
                    result = 4;

                danwei = table_three.cell_value(nrows_one, 4)
                remark = table_three.cell_value(nrows_one, 5)
                # remark = "无"
                year = 0
                create_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))

                sql = "insert into tb_reward (user_id,user_name,`year`,begin_time,tittle,content,part_name,result,remark,create_time) " \
                      " values (%d,'%s',%d,'%s','%s','%s','%s',%d,'%s','%s')" % \
                      (user_id, user_name, year, begin_time,title, content,danwei, result, remark, create_time)
                print(sql)
                try:
                    # 使用 cursor() 方法创建一个游标对象 cursor
                    cursor = db.cursor()
                    cursor.execute(sql)
                except Exception as e:
                    # 发生错误时回滚
                    db.rollback()
                    print(str(e))
                else:
                    db.commit()  # 事务提交
                    print('事务处理成功')
                    print("\n")

                print(index, start_time, begin_time, title, content, result, danwei, remark)


def find_table_tow(user_name, table_two):
    # # 连接mysql
    db = pymysql.connect("localhost", "root", "root", "cgb-layui", use_unicode=True, charset="utf8")
    # village = table_two.cell_value(5, 3)
    sql = """SELECT id,name FROM tb_user WHERE name = '%s'""" % (user_name)
    cursor = db.cursor();
    cursor.execute(sql)
    results = cursor.fetchall()

    if not results:
        user_id = 0
        user_name = 0
    else:
        user_id = results[0][0]
        user_name = results[0][1]

        print(results[0][0], results[0][1])

    rows = table_two.nrows

    print(rows)
    print(str(table_two.cell_value(1, 1)).strip() != '')
    for nrows_one in range(2, int(table_two.nrows)):
        print("进入第"+str(nrows_one)+"行记录")
        index = int(table_two.cell_value(nrows_one, 0))

        start_time = table_two.cell_value(nrows_one, 1)
        start_time = str(start_time);
        begin_time = start_time[0:4] + '-' + start_time[4:6] + '-' + start_time[6:8] + " 00:00:00"

        days = table_two.cell_value(nrows_one, 2)

        title = table_two.cell_value(nrows_one, 3)
        content = table_two.cell_value(nrows_one, 4)
        resultstr = table_two.cell_value(nrows_one, 5)

        # 培训结果 0-不及格，1-及格，2-良好，3-优秀
        if "不及格" in resultstr:
            result = 0
        elif "及格" in resultstr:
            result = 1
        elif "良好" in resultstr:
            result = 2
        elif "优秀" in resultstr:
            result = 3
        else:
            result = 4;

        org_name = table_two.cell_value(nrows_one, 6)
        fanwei = table_two.cell_value(nrows_one, 7)

        if "省级" in fanwei:
            ranges = 0
        elif "市州" in fanwei:
            ranges = 1
        elif "县(市.区)" in fanwei:
            ranges = 2
        elif "乡镇" in fanwei:
            ranges = 3
        else:
            ranges = 4;
        remark = "无"
        year = 0
        create_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))

        sql = "insert into tb_practice (user_id,user_name,`year`,begin_time,days,tittle,content,result,org_name,remark,create_time,`range`) " \
              " values (%d,'%s',%d,'%s','%s','%s','%s','%s','%s','%s','%s','%s')" % \
              (
                  user_id, user_name, year, begin_time, days, title, content, result, org_name, remark, create_time,
                  ranges)
        print(sql)
        try:
            # 使用 cursor() 方法创建一个游标对象 cursor
            cursor = db.cursor()
            cursor.execute(sql)
        except Exception as e:
            # 发生错误时回滚
            db.rollback()
            print(str(e))
        else:
            db.commit()  # 事务提交
            print('事务处理成功')
            print("\n")

        print(index, start_time, ranges, title, content, result, org_name, fanwei, remark)


def find_table_one(table_one):
    # # 连接mysql
    db = pymysql.connect("localhost", "root", "root", "cgb-layui", use_unicode=True, charset="utf8")
    village = table_one.cell_value(5, 3)
    print(village)
    sql = """SELECT id FROM sys_org WHERE `name` = '%s'""" % (village)
    cursor = db.cursor();
    cursor.execute(sql)
    results = cursor.fetchall()

    if not results:
        org_id = 0
    else:
        org_id = results[0][0]
    print("部门ID",org_id)
    name = table_one.cell_value(1, 1)
    id_card = table_one.cell_value(1, 3)
    print(id_card)
    if int(id_card[16]) % 2 == 0:
        sex = 2;
    else:
        sex = 1;

    nation = table_one.cell_value(2, 3)

    typed = table_one.cell_value(3, 3)
    ## 职位：0-第一书记，1-村支部书记，2-村主任，3-副村支部书记，4-副村主任，5-妇女主任，6-文书，7-监委会主任，8-村书记兼主任 ，9-村总支书记，10-村总支副书记',
    if typed == "村党总支书记、主任":
        type = 0
    elif typed == "村党支部书记、主任":
        type = 1
    elif typed == "村党总支副书记":
        type = 2
    elif typed == "村党支部副书记":
        type = 3
    elif typed == "村委会副主任":
        type = 4
    elif typed == "文书":
        type = 5
    else:
        type = 11
    endtime = time.strftime('%Y', time.localtime(time.time()))
    age = int(endtime) - int(id_card[6:10])

    ## 政治面貌,1-群众，2-团员，3-党员，4-其他
    politicstr = str(table_one.cell_value(2, 1))
    if "党员" in politicstr:
        politic = 3
    elif "团员" in politicstr:
        politic = 2
    elif "群众" in politicstr:
        politic = 1
    else:
        politic = 4;

    educationstr = table_one.cell_value(4, 1)
    ## 学历，0-小学，1-初中，2-高中，3-大专，4-本科，5-硕士，6-博士及博士及以上'
    if educationstr == "初中及以下":
        education = 0
    elif educationstr == "高中及中专":
        education = 1
    elif educationstr == "大专":
        education = 2
    elif educationstr == "本科及以上":
        education = 3
    else:
        education = 6

    degree = table_one.cell_value(4, 3)
    birthday = str(id_card[6:10]) + "-" + str(id_card[10:12]) + "-" + str(id_card[12:14]) + " 00:00:00"
    timestr = str(table_one.cell_value(3, 1))
    if timestr.strip() != '':
        join_time = timestr[0:4] + '-' + timestr[4:6] + '-' + timestr[6:8] + " 00:00:00"
    else:
        join_time = "1971-01-01 00:00:00"

    tel = int(table_one.cell_value(7, 3))

    vcadresstr = table_one.cell_value(15, 1)
    if vcadresstr == "否":
        vcadres = 0
    elif vcadresstr == "是":
        vcadres = 1
    else:
        vcadres = 2

    first_people = 0;
    address = table_one.cell_value(15, 3)

    statusstr = table_one.cell_value(6, 3)
    # 状态1-在职。2-离休，3-后备
    if "后备" in statusstr:
        status = 3
    elif "离休" in statusstr:
        status = 2
    elif "在职" in statusstr:
        status = 1
    else:
        status = 4;

    begin_timestr = str(table_one.cell_value(9, 1))
    begin_time = begin_timestr[0:4] + '-' + begin_timestr[4:6] + '-' + begin_timestr[6:8] + " 00:00:00"

    serving_timestr = str(table_one.cell_value(9, 3))
    serving_time = serving_timestr[0:4] + '-' + serving_timestr[4:6] + '-' + serving_timestr[6:8] + " 00:00:00"

    lw_statusstr = table_one.cell_value(8, 1)
    if "是" in lw_statusstr:
        lw_status = 1
    elif "否" in lw_statusstr:
        lw_status = 0
    else:
        lw_status = 2;

    part_statusstr = table_one.cell_value(13, 1)
    if "是" in part_statusstr:
        part_status = 1
    elif "否" in part_statusstr:
        part_status = 0
    else:
        part_status = 2;

    #
    last_status = table_one.cell_value(6, 1)

    #

    is_nowstr = table_one.cell_value(7, 1)
    if "是" in is_nowstr:
        is_now = 1
    elif "否" in is_nowstr:
        is_now = 0
    else:
        is_now = 2;

    #
    country = table_one.cell_value(5, 1)
    #
    village = table_one.cell_value(5, 3)

    # 是否是两委员一代表0-无，1-市级，2-区级，3-乡级'
    lwstr = table_one.cell_value(12, 1)

    if "市级" in lwstr:
        lw = 1
    elif "区级" in lwstr:
        lw = 2
    elif "乡级" in lwstr:
        lw = 3
    else:
        lw = 0;

    # 1-党代表，2-人大代表，3-政协委员
    lw_remarkstr = table_one.cell_value(12, 3)

    if "党代表" in lw_remarkstr:
        lw_remark = 1
    elif "人大代表" in lw_remarkstr:
        lw_remark = 2
    elif "政协委员" in lw_remarkstr:
        lw_remark = 3
    else:
        lw_remark = 0;
    # 1-经济合作组织,2-专业合作社,3-社会组织,4-其他
    part_remarkstr = table_one.cell_value(13, 3)

    if "经济合作组织" in part_remarkstr:
        part_remark = 1
    elif "专业合作社" in part_remarkstr:
        part_remark = 2
    elif "社会组织" in part_remarkstr:
        part_remark = 3
    else:
        part_remark = 4;

    # 是否为专职村党组织书记,0-1否，1是'
    zz_statusstr = table_one.cell_value(8, 3)
    if "是" in zz_statusstr:
        zz_status = 1
    elif "否" in zz_statusstr:
        zz_status = 0
    else:
        zz_status = 2;

    #
    practice = table_one.cell_value(10, 1)

    #
    work = table_one.cell_value(10, 3)

    #

    salary = table_one.cell_value(14, 1)
    remark = '无'

    create_time = time.strftime('%Y-%m-%d  %H:%M:%S', time.localtime(time.time()))

    print(org_id, name, sex, nation, type, id_card, age, politic, education, degree, birthday, join_time, tel, vcadres,
          first_people, address, status, begin_time, serving_time, lw_status, part_status, last_status,
          is_now, country, village, lw, lw_remark, part_remark, zz_status, practice, work, salary)

    #     # 将数据存入数据库
    # insert into `cgb-layui`.`tb_user` ( `update_time`, `vcadres`, `remark`, `education`, `is_now`, `country`, `zz_status`, `address`,
    # `part_remark`, `salary`, `last_status`, `tel`, `education_file`, `is_deleted`, `type`, `practice`, `join_time`, `sex`, `create_time`,
    # `create_by`, `id_card`, `birthday`, `politic`, `begin_file`, `name`, `serving_file`, `org_id`, `status`, `remarks`, `degree_file`,
    # `lw_status`, `lw`, `part_status`, `serving_time`, `age`, `first_people`, `degree`, `nation`, `work`, `lw_remark`, `update_by`, `begin_time`, `village`)
    # values ( null, '0', 'http://148.70.13.237:89/upload/accab912-df32-4d32-bf3d-c48cbb5ce093.jpg', '1', '0', '东洞镇', '0', '东洞镇新沟村', '2', '24600', '称职',
    # '17718661326', null, '0', '6', '无', '2010-06-11 00:00:00', '1', '2020-05-05 11:25:00', '0', '62210219680125731X', '1968-01-25 00:00:00', '3', null, '贾光忠',
    # null, 'ff8080817022df970170253ba6f3062c', '1', '无', null, '0', '3', '1', '2018-03-10 00:00:00', '51', '0', '无', '汉族', '东洞镇新沟村村委会', '2', '0', '2018-03-10 00:00:00',
    # '新沟村');

    # insert into `cgb-layui`.`tb_user` ( `update_time`, `vcadres`, `remark`, `education`, `is_now`, `country`, `zz_status`, `address`,
    # `part_remark`, `salary`, `last_status`, `tel`, `education_file`, `is_deleted`, `type`, `practice`, `join_time`, `sex`, `create_time`,
    # `create_by`, `id_card`, `birthday`, `politic`, `begin_file`, `name`, `serving_file`, `org_id`, `status`, `remarks`, `degree_file`,
    # `lw_status`, `lw`, `part_status`, `serving_time`, `age`, `first_people`, `degree`, `nation`, `work`, `lw_remark`, `update_by`, `begin_time`, `village`)

    # sql = "insert into sales_data(SITE, PAYDAY, SALES, QUANTITY_TICKET, RATE_ELECTRONIC, SALES_THANLASTWEEK," \
    #           "SALES_THANLASTYEAR, SESSION, SESSION_THANLASTWEEK, RATE_CONVERSION, RATE_PAYSUCCESS)" \
    #           " values ('%s','%s', %f, %d, %f, %f, %f, %d, %f, %f, %f)" %\
    #           (site, payday, sales, quantity_ticket, rate_electronic, sales_thanlastweek, sales_thanlastyear,
    #            session, session_thanlastweek, rate_conversion, rate_paysuccess)

    sql = "insert into tb_user(org_id,name,sex,nation,type,id_card,age,politic,education,degree,birthday,join_time,tel,vcadres,first_people,address,status,begin_time,serving_time,lw_status,part_status,last_status,is_now,country,village,lw,lw_remark,part_remark,zz_status,practice,work,remarks,salary,create_time)" \
          " values ('%s','%s',%d,'%s',%d,'%s',%d, %d, %d,'%s','%s','%s',%d,%d, %d,'%s',%d,'%s','%s',%d,%d,'%s',%d,'%s','%s',%d,%d,%d,%d,'%s','%s','%s','%s','%s')" % \
          (org_id, name, sex, nation, type, id_card, age, politic, education, degree, birthday, join_time, tel, vcadres,
           first_people, address, status, begin_time, serving_time, lw_status, part_status, last_status, is_now,
           country, village, lw, lw_remark, part_remark, zz_status, practice, work, remark, salary, create_time)
    print(sql)
    try:
        # 使用 cursor() 方法创建一个游标对象 cursor
        cursor = db.cursor()
        cursor.execute(sql)
    except Exception as e:
        # 发生错误时回滚
        db.rollback()
        print(str(e))
    else:
        db.commit()  # 事务提交
        print('事务处理成功')
        print("\n")


def find_table_test(table_one):
    # # 连接mysql
    db = pymysql.connect("localhost", "root", "root", "cgb-layui", use_unicode=True, charset="utf8")
    village = table_one.cell_value(5, 3)
    print(village)
    sql = """SELECT id FROM sys_org WHERE `name` = '%s'""" % (village)
    cursor = db.cursor();
    cursor.execute(sql)
    results = cursor.fetchall()

    if not results:
        org_id = 0
    else:
        org_id = results[0][0]
    print("部门ID",org_id)
    name = table_one.cell_value(1, 1)
    id_card = table_one.cell_value(1, 3)
    print(id_card)
    if int(id_card[16]) % 2 == 0:
        sex = 2;
    else:
        sex = 1;

    nation = table_one.cell_value(2, 3)

    typed = table_one.cell_value(3, 3)
    ## 职位：0-第一书记，1-村支部书记，2-村主任，3-副村支部书记，4-副村主任，5-妇女主任，6-文书，7-监委会主任，8-村书记兼主任 ，9-村总支书记，10-村总支副书记',
    if typed == "村党总支书记、主任":
        type = 0
    elif typed == "村党支部书记、主任":
        type = 1
    elif typed == "村党总支副书记":
        type = 2
    elif typed == "村党支部副书记":
        type = 3
    elif typed == "村委会副主任":
        type = 4
    elif typed == "文书":
        type = 5
    else:
        type = 11
    endtime = time.strftime('%Y', time.localtime(time.time()))
    age = int(endtime) - int(id_card[6:10])

    ## 政治面貌,1-群众，2-团员，3-党员，4-其他
    politicstr = str(table_one.cell_value(2, 1))
    if "党员" in politicstr:
        politic = 3
    elif "团员" in politicstr:
        politic = 2
    elif "群众" in politicstr:
        politic = 1
    else:
        politic = 4;

    educationstr = table_one.cell_value(4, 1)
    ## 学历，0-小学，1-初中，2-高中，3-大专，4-本科，5-硕士，6-博士及博士及以上'
    if educationstr == "初中及以下":
        education = 0
    elif educationstr == "高中及中专":
        education = 1
    elif educationstr == "大专":
        education = 2
    elif educationstr == "本科及以上":
        education = 3
    else:
        education = 6

    degree = table_one.cell_value(4, 3)
    birthday = str(id_card[6:10]) + "-" + str(id_card[10:12]) + "-" + str(id_card[12:14]) + " 00:00:00"
    timestr = str(table_one.cell_value(3, 1))
    if timestr.strip() != '':
        join_time = timestr[0:4] + '-' + timestr[4:6] + '-' + timestr[6:8] + " 00:00:00"
    else:
        join_time = "1971-01-01 00:00:00"

    tel = int(table_one.cell_value(7, 3))

    vcadresstr = table_one.cell_value(15, 1)
    if vcadresstr == "否":
        vcadres = 0
    elif vcadresstr == "是":
        vcadres = 1
    else:
        vcadres = 2

    first_people = 0;
    address = table_one.cell_value(15, 3)

    statusstr = table_one.cell_value(6, 3)
    # 状态1-在职。2-离休，3-后备
    if "后备" in statusstr:
        status = 3
    elif "离休" in statusstr:
        status = 2
    elif "在职" in statusstr:
        status = 1
    else:
        status = 4;

    begin_timestr = str(table_one.cell_value(9, 1))
    begin_time = begin_timestr[0:4] + '-' + begin_timestr[4:6] + '-' + begin_timestr[6:8] + " 00:00:00"

    serving_timestr = str(table_one.cell_value(9, 3))
    serving_time = serving_timestr[0:4] + '-' + serving_timestr[4:6] + '-' + serving_timestr[6:8] + " 00:00:00"

    lw_statusstr = table_one.cell_value(8, 1)
    if "是" in lw_statusstr:
        lw_status = 1
    elif "否" in lw_statusstr:
        lw_status = 0
    else:
        lw_status = 2;

    part_statusstr = table_one.cell_value(13, 1)
    if "是" in part_statusstr:
        part_status = 1
    elif "否" in part_statusstr:
        part_status = 0
    else:
        part_status = 2;

    #
    last_status = table_one.cell_value(6, 1)

    #

    is_nowstr = table_one.cell_value(7, 1)
    if "是" in is_nowstr:
        is_now = 1
    elif "否" in is_nowstr:
        is_now = 0
    else:
        is_now = 2;

    #
    country = table_one.cell_value(5, 1)
    #
    village = table_one.cell_value(5, 3)

    # 是否是两委员一代表0-无，1-市级，2-区级，3-乡级'
    lwstr = table_one.cell_value(12, 1)

    if "市级" in lwstr:
        lw = 1
    elif "区级" in lwstr:
        lw = 2
    elif "乡级" in lwstr:
        lw = 3
    else:
        lw = 0;

    # 1-党代表，2-人大代表，3-政协委员
    lw_remarkstr = table_one.cell_value(12, 3)

    if "党代表" in lw_remarkstr:
        lw_remark = 1
    elif "人大代表" in lw_remarkstr:
        lw_remark = 2
    elif "政协委员" in lw_remarkstr:
        lw_remark = 3
    else:
        lw_remark = 0;
    # 1-经济合作组织,2-专业合作社,3-社会组织,4-其他
    part_remarkstr = table_one.cell_value(13, 3)

    if "经济合作组织" in part_remarkstr:
        part_remark = 1
    elif "专业合作社" in part_remarkstr:
        part_remark = 2
    elif "社会组织" in part_remarkstr:
        part_remark = 3
    else:
        part_remark = 4;

    # 是否为专职村党组织书记,0-1否，1是'
    zz_statusstr = table_one.cell_value(8, 3)
    if "是" in zz_statusstr:
        zz_status = 1
    elif "否" in zz_statusstr:
        zz_status = 0
    else:
        zz_status = 2;

    #
    practice = table_one.cell_value(10, 1)

    #
    work = table_one.cell_value(10, 3)

    #

    salary = table_one.cell_value(14, 1)
    remark = '无'

    create_time = time.strftime('%Y-%m-%d  %H:%M:%S', time.localtime(time.time()))

    print(org_id, name, sex, nation, type, id_card, age, 99999+politic, education, degree, birthday, join_time, tel, vcadres,
          first_people, address, status, begin_time, serving_time, lw_status, part_status, last_status,
          9999+is_now, country, village, lw, lw_remark, part_remark, zz_status, practice, work, salary)


path = r'F:\村干部信息系统收资料（2020.9.20）(1)\村干部信息系统收资料（2020.9.20）\A'
filenames = os.listdir(path)
filepath = 'F:\村干部信息系统收资料（2020.9.20）(1)\村干部信息系统收资料（2020.9.20）\A\村干部信息导入模板 - 王莎莉.xls'
# find(filepath)
for filename in filenames:
    print(filenames.__len__())
    find(path + '\\' + filename)
    # if filename.endswith(".xls"):
    #     # print(filename)
    #     print(filename)
    #     find(path + '\\' + filename)
