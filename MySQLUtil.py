# -*- coding: utf-8 -*-

from pymysql import connect
from pymysql.cursors import DictCursor


class MySQLUtil:
    """MySQL 基本功能封装 """

    def __init__(self,
                 host: str,
                 port: int,
                 user: str,
                 passwd: str,
                 database: str):
        self.conn = connect(
            host=host,
            port=port,
            user=user,
            passwd=passwd,
            db=database,
            charset='utf8',  # utf8mb4 是utf8的超集
            cursorclass=DictCursor  # 返回类字典类型游标
            )  
        self.cur = self.conn.cursor()

    def execute(self, sql, args=None):
        self.cur.execute(sql, args)

    def executemany(self, sql, args):
        self.cur.executemany(sql, args)
        self.commit()

    def fetchall(self, sql, args=None):
        self.execute(sql, args)
        return self.cur.fetchall()

    def fetchone(self, sql, args=None):
        self.execute(sql, args)
        return self.cur.fetchone()

    def has_table(self, table_name: str) -> bool:
        """该用户下是否存在表table_name"""
        
        sql = "SELECT count(*) total FROM information_schema.TABLES WHERE table_name =%(table_name)s"
        arg = {'table_name': table_name}
        return self.fetchone(sql, arg).get('total') == 1

    @property
    def rowcount(self):
        """受影响的行数"""
        return self.cur.rowcount

    def commit(self):
        self.conn.commit()

    def rollback(self):
        self.conn.rollback()

    def close(self):
        self.conn.close()

    #def __del__(self):
    #    self.close()
        

    