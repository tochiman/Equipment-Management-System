import sqlite3

#create table equipment(id integer primary key autoincrement,goods string,goods_detail string,use_group string,name string,purchase_date integer,control_num string unique,note string,waste string)

class Ping():
    def __init__(self) -> None:
        #For "function name is ping"
        self.host_name = "www.google.com"
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
        self.dbname = r"C:\Users\yuuto\OneDrive\VS_code\equipment-management-system\main.sqlite3"
        self.conn = sqlite3.connect(self.dbname)
        self.cur = self.conn.cursor()

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
                "select * from equipment")
            for i in self.cur:
                all_list.append(i)
        finally:
            self.conn.commit()
            self.cur.close()
            self.conn.close()

        return all_list
