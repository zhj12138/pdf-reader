import sqlite3
import time


class recentFile:
    def __init__(self, path, opentime):
        self.path = path
        self.opentime = opentime


class MyDb:
    def __init__(self):
        self.name = "info.db"
        self.recenTable = 'recentfile'
        self.kMailTable = 'kindlemail'
        conn = sqlite3.connect(self.name)
        cursor = conn.cursor()
        try:
            cursor.execute("create table %s(path text primary key , opentime float)" % self.recenTable)
            cursor.execute("create table %s(mail text primary key )" % self.kMailTable)
        except:
            pass
        conn.commit()
        conn.close()

    def addRecentFile(self, file_path):
        conn = sqlite3.connect(self.name)
        cursor = conn.cursor()
        cursor.execute("insert into %s values ('%s', %s)" % (self.recenTable, file_path, time.time()))
        conn.commit()
        conn.close()

    def updateRecentFile(self, file_path):
        conn = sqlite3.connect(self.name)
        cursor = conn.cursor()
        cursor.execute("update %s set opentime=%s where path='%s'" % (self.recenTable, time.time(), file_path))
        conn.commit()
        conn.close()

    def deleteRecentFile(self, file_path):
        conn = sqlite3.connect(self.name)
        cursor = conn.cursor()
        cursor.execute("delete from %s where path='%s'" % (self.recenTable, file_path))
        conn.commit()
        conn.close()

    def getAllRencentFile(self):
        conn = sqlite3.connect(self.name)
        cursor = conn.cursor()
        ret = cursor.execute("select * from %s" % self.recenTable)
        fileList = []
        for row in ret:
            path, opentime = row[0], row[1]
            fileList.append(recentFile(path, opentime))
        conn.commit()
        conn.close()
        return fileList

    def fileInDB(self, file_path):
        conn = sqlite3.connect(self.name)
        cursor = conn.cursor()
        ret = cursor.execute("select * from %s where path='%s'" % (self.recenTable, file_path))
        tem = []
        for row in ret:
            tem.append(row)
        conn.commit()
        conn.close()
        if not tem:
            return False
        return True

    def getAllKindleMail(self):
        conn = sqlite3.connect(self.name)
        cursor = conn.cursor()
        ret = cursor.execute("select * from %s" % self.kMailTable)
        emailList = []
        for row in ret:
            email = row[0]
            emailList.append(email)
        conn.commit()
        conn.close()
        return emailList

    def mailInDB(self, mail):
        conn = sqlite3.connect(self.name)
        cursor = conn.cursor()
        ret = cursor.execute("select * from %s where mail='%s'" % (self.kMailTable, mail))
        tem = []
        for row in ret:
            tem.append(row)
        conn.commit()
        conn.close()
        if not tem:
            return False
        return True

    def addKindleMail(self, email):
        conn = sqlite3.connect(self.name)
        cursor = conn.cursor()
        cursor.execute("insert into %s values ('%s')" % (self.kMailTable, email))
        conn.commit()
        conn.close()

