import sqlite3
from os import environ
from dotenv import load_dotenv

class Ping():
    def __init__(self) -> None:
        load_dotenv()
        #For "function name is ping"
        self.host_name = environ['url']
        self.port = 80
        self.address = (self.host_name, self.port)

    def ping(self) -> bool:
        import socket
        try:
            socket.create_connection(self.address)
            return True
        except OSError:
            return False

class DB_operation():
    def __init__(self) -> None:
        #環境変数を読み込む
        load_dotenv()
        self.dbname = environ['db_path']
        self.conn = sqlite3.connect(self.dbname)
        self.cur = self.conn.cursor()
        self.cur.execute("create table if not exists equipment(id integer primary key autoincrement,goods string,goods_detail string,use_group string,name string,purchase_date integer,control_num string unique,note string,waste string)")

    def db_register(self, get_list: list) -> None:
        try:
            all_list = []
            self.cur.execute('INSERT INTO equipment(goods,goods_detail,use_group,name,purchase_date,control_num,note) values(?,?,?,?,?,?,?)',
                             (get_list[0], get_list[1], get_list[2], get_list[3], get_list[4], get_list[5], get_list[6]))
        finally:
            self.conn.commit()
            self.cur.close()
            self.conn.close()

    def db_waste(self, control_num: str) -> None:
        try:
            self.cur.execute(
                "update equipment set waste=? where control_num=?", ('済', control_num))
        finally:
            self.conn.commit()
            self.cur.close()
            self.conn.close()

    def db_delete(self, control_num: str) -> None:
        try:
            self.cur.execute(
                "delete from equipment where control_num=?", (control_num,))
        finally:
            self.conn.commit()
            self.cur.close()
            self.conn.close()
    def db_extract(self) -> list:
        all_list = []
        try:
            self.cur.execute(
                "select goods, goods_detail, use_group, name, purchase_date, control_num, note, waste from equipment")
            for i in self.cur:
                all_list.append(i)
        finally:
            self.conn.commit()


        return all_list

    def shutdown(self):
        self.cur.close()
        self.conn.close()