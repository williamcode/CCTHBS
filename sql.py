# encoding=utf-8 
import MySQLdb
import pymssql

class MYSQL():
    def __init__(self):
        self.host=''
        self.port=
        self.user=''
        self.passwd=''
        self.db=''


    def __connection(self):
        try:
            self.cnn=MySQLdb.connect(host=self.host,port=self.port,user=self.user,passwd=self.passwd,db=self.db,charset='utf8',use_unicode=True)
           
        except:
            pass
    def Sqlexe(self,sql):
        
        self.__connection()
        if self.cnn:
            cur=self.cnn.cursor()
            cur.execute(sql)
            res=cur.fetchall()
            cur.close()
            self.cnn.close()
            return res
        else:
            return None
    def exe_noqueryset(self,sql):
        self.__connection()
        if self.cnn:
            self.cnn.cursor().execute(sql)
            self.cnn.commit()
            self.cnn.close()



class MSSQL():
    def __init__(self):
        self.host='10.9.1.52'
        #self.port=3
        self.user='sa'
        self.passwd='sa'
        self.db='CCTHBS'
        
    def __connection(self):
        try:
            self.cnn=pymssql.connect(host = self.host,user=self.user,password=self.passwd,database=self.db,charset='utf8')
        except:
            pass

    def Sqlexe(self,sql):
        
        self.__connection()
        if self.cnn:
            cur=self.cnn.cursor()
            cur.execute(sql)
            res=cur.fetchall()
            cur.close()
            self.cnn.close()
            return res
        else:
            return None
    def exe_noqueryset(self,sql):
        self.__connection()
        if self.cnn:
            self.cnn.cursor().execute(sql)
            self.cnn.commit()
            self.cnn.close()
            
   
