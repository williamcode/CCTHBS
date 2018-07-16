# encoding=utf-8
import MySQLdb
import pymssql

class MYSQL():
    def __init__(self):
        self.host=''
        self.port=''
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

class Voucher():
    #借方发生额：日期，摘要，科目，币种，记账币金额，外币金额，对方科目
    #list类型，日期，外币币种，外币金额，本位币金额，借方科目，借方单位，借方科目币种，贷方科目，贷方单位，贷方科目币种，备注，ref)
    #    0    ,  1,    2   ,    3   ,    4           5        6       7            8       9          10       11    12
    def bko(self,sql_re,unit):
        if sql_re[6]==unit or sql_re[9]==unit:
            result=[]
            temp= sql_re[1],sql_re[11],sql_re[5] if sql_re[6]==unit else sql_re[6],sql_rec[7],sql_rec[4]，sql_rec[3]，sql_re[8] if sql_re[9]==unit else sql_re[9],'dr'
            result.append(temp)
            temp= sql_re[1],sql_re[11],sql_re[8] if sql_re[9]==unit else sql_re[9],sql_rec[10],sql_rec[4]，sql_rec[3]，sql_re[5] if sql_re[6]==unit else sql_re[6],'cr'
            result.append(temp)

        return result
    # type,日期，币种，外币金额，本位币金额，‘’，账套，‘’，‘’，‘’，‘’，合同号，供应商
    def ap(sel,sql_re,unit):
        if sql_re[6]==unit:
            result=[]
            temp=sql_re[1],sql_re[11],'1405',sql_re[2],sql_re[4],sql_re[3],sql_re[12],'dr'
            result.append(temp)
            temp=sql_re[1],sql_re[11]+'_采购',sql_re[12],sql_re[2],sql_re[4],sql_re[3],'1405','cr'
            result.append(temp)
        return result

    # type,日期，币种，外币金额，本位币金额，‘’，账套，‘’，‘’，‘’，‘’，合同号，客户
    def ar(self,sql_re,unit):
        if sql_re[6]==unit:
            result=[]
            temp=sql_re[1],sql_re[11]+'_销售',sql_re[12],sql_re[2],sql_re[4],sql_re[3],sql_re[6],'dr'
            result.append(temp)
            temp=sql_re[1],sql_re[11],sql_re[6],sql_re[2],sql_re[4],sql_re[3],sql_re[12],'cr'
            result.append(temp)
        return result
    # type,日期，币种，手续费，银行金额，收款银行，账套，‘’，金额（字符型），客户,收款币种，合同号，b.excelserverrcid
    def bank_receive(self,sql_re,unit):
        if sql_re[6]==unit:
            result=[]
            temp=sql_re[1],sql_re[11]+'_收款',sql_re[5],sql_re[2],sql_re[4],sql_re[4],sql_re[9],'dr'
            result.append(temp)
            temp=sql_re[1],'续费',sql_re[5],sql_re[2],sql_re[3],sql_re[3],sql_re[9],'dr'
            result.append(temp)
            temp=sql_re[1],sql_re[11],sql_re[9],sql_re[2],sql_re[8],sql_re[8],sql_re[5],'cr'
            result.append(temp)
        return result

    # type,日期，币种，外币金额，本位币金额，供应商，账套，‘’，支付银行，‘’,‘’，合同号，‘’
    def bank_pay(self,sql_re,unit):
        if sql_re[6]==unit:
            result=[]
            temp=sql_re[1],sql_re[11]+'_支付',sql_re[5],sql_re[2],sql_re[4],sql_re[3],sql_re[8],'dr'
            result.append(temp)
            temp=sql_re[1],sql_re[11],sql_re[8],sql_re[2],sql_re[4],sql_re[3],sql_re[5],'cr'
            result.append(temp)
        return result
