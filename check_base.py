import sqlite3
import os.path

from PyQt5 import QtCore, QtGui, QtWidgets
from QT import Ui_Form


class DataBase:
    """
    Класс базы данных.
    Создает базу данных в корневой папке
    
    connectDB - функция подключения к БД
    createDB - функция создания БД
    dataDB - возвращает все данные из БД
    updDB(upd_data) - позволяет обновить БД в соответствии с поступаемой информацией
    clearDB - функция обнуления БД
    """
    def __init__(self):
        if not os.path.exists('check.db'):
            self.createDB()
    
    
    def connectDB(self):
        try:
            self.conn = sqlite3.connect('check.db')
            self.cursor = self.conn.cursor()
        
        except sqlite3.Error as error:
                print("Ошибка при работе с SQLite", error)
                if self.conn:
                    self.conn.close()
    
    
    def createDB(self):
        try:
            self.connectDB()
            # Создание таблицы
            self.cursor.execute("""CREATE TABLE SchoolCheck
                          (School, id1, id2, id3, id4,
                           id5, id6, id7, id8, id9, id10)""")
            
            school = """1 2(всош) 3 4 5 6 7 8 10 11 13 14 15 17 18 19 20
                     21 22 23 24 25 26 27 28 29 30 31 32 33 34 35 36 37
                     38 41 44 Г1""".split()
            
            start_data = [(i, '-', '-', '-', '-', '-', '-', '-', '-', '-', '-') for i in school]
            
            self.cursor.executemany("INSERT INTO SchoolCheck VALUES (?,?,?,?,?,?,?,?,?,?,?)", start_data)
            self.conn.commit()
            self.cursor.close()
        
        except sqlite3.Error as error:
            print("Ошибка при работе с SQLite", error)
        finally:
            if self.conn:
                self.conn.close()
            

    def dataDB(self):
        try:
            self.connectDB()
            self.cursor.execute("""SELECT * from SchoolCheck""")
            self.data = self.cursor.fetchall()
            
            return self.data
        
        except sqlite3.Error as error:
            print("Ошибка при работе с SQLite", error)
        
        finally:
            if self.conn:
                self.conn.close()
    
    
    def updDB(self, upd_data):
        try:
            self.connectDB()
            self.cursor.execute("""UPDATE SchoolCheck SET id1=?, id2=?, id3=?, id4=?,
                                id5=?, id6=?, id7=?, id8=?, id9=?, id10=?
                                WHERE School=?;""", upd_data)
            self.conn.commit()
            self.cursor.close()   
        
        except sqlite3.Error as error:
            print("Ошибка при работе с SQLite", error)
        
        finally:
            if self.conn:
                self.conn.close()   
            
            
    def clearDB(self):
        try:
            self.connectDB()
            self.cursor.execute("DROP TABLE SchoolCheck")
            self.conn.commit()
            self.cursor.close()
            self.conn.close()
            
            self.createDB()
        
        finally:
            if self.conn:
                self.conn.close()
            

class DataBaseTable(Ui_Form):
    """
    Класс подключения интерфейса таблицы QT к базе данных
    """
    def __init__(self, form):
        self.db = DataBase()
        self.data = self.db.dataDB()
        super().__init__()
        
        self.form = form
        self.setupUi(self.form)
        
        self.upd()
        
        self.form.show()
    
    
    def clear(self):
        self.db.clearDB()
        self.data = self.db.dataDB()
        self.upd()
    
    
    def upd(self):
        row = 0
        for school in self.data:
            col = 0
            for val in school[1:]:
                cellinfo = QtWidgets.QTableWidgetItem(val)         
                self.tableWidget.setItem(row, col, cellinfo)
                col += 1
            row += 1
        
    
if __name__ == '__main__':
    pass