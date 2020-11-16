#!/usr/bin/env python
# -*- coding:utf-8 -*-
import pymysql


def conn_db():  # 连接数据库函数
    conn = pymysql.connect(
        host='cdb-68vlr0so.bj.tencentcdb.com',
        port=10231,
        user='root',
        passwd='Nihui@123456',
        db='test',
        charset='utf8')
    cur = conn.cursor()
    return conn, cur


def exe_update(cur, sql):  # 更新语句，可执行update,insert语句
    sta = cur.execute(sql)
    return sta


def exe_delete(cur, ids):  # 删除语句，可批量删除
    for eachID in ids.split(' '):
        sta = cur.execute('delete from cms where id =%d' % int(eachID))
    return sta


def exe_query(cur, sql):  # 查询语句
    cur.execute(sql)
    return cur


def exe_commit(cur):
    cur.connection.commit()  # 执行commit操作，插入语句才能生效


def conn_close(conn, cur):  # 关闭所有连接
    cur.close()
    conn.close()

if __name__ == '__main__':
    conn_db()