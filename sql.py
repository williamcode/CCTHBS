# encoding=utf-8
import MySQLdb
import pymssql

class MYSQL():
    def __init__(self):
        self.host='10.9.1.52'
        self.port=3316
        self.user='root'
        self.passwd='htoa8000'
        self.db='oa8000'


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

class Voucher():
    #借方发生额：日期，摘要，科目，币种，记账币金额，外币金额，汇率，客商项目（代码，如有），对方科目
    #list类型，日期，外币币种，外币金额，本位币金额，借方科目，借方单位，借方科目币种，贷方科目，贷方单位，贷方科目币种，备注，ref)
    def __init__(self):
        account_kbo = {
        ''：['',''], #科目名称：【科目代码，项目代码】
        ''：['',''],
        ''：['',''],
        ''：['',''],
        ''：['',''],
        ''：['',''],
        }
        account_unit = {
        '':['',''],
        '':['',''],
        '':['',''],
        }

    def bko(self,sql_re,unit):
        if sql_re[6]==unit or sql_re[9]==unit:
            result=[]
            if sql_re[6]==unit:
                temp= sql_re[1],sql_re[11],account_kbo[sql_rec[5]][0],sql_rec[7],sql_rec[4]，sql_rec[3]，if sql_rec[4] then round(sql_rec[3]/sql_rec[4],4),
                                                                                        account_kbo[sql_rec[5]][1],account_kbo[sql_rec[5]][0]
                result.append(temp)
            else:
                temp= sql_re[1],sql_re[11],account_unit[sql_rec[5]][0],sql_rec[7],sql_rec[4]，sql_rec[3]，if sql_rec[4] then round(sql_rec[3]/sql_rec[4],4),
                                                                                        account_unit[sql_rec[5]][1],account_kbo[sql_rec[5]][0]
                result.append(temp)
