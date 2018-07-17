# -*- coding:utf-8 -*-
#############################################################################################
#导入所需信息                                                                                #
#借方发生额：日期，摘要，科目，币种，记账币金额，外币金额，汇率，客商项目（代码，如有），对方科目    #
#贷方发生额：日期，摘要，科目，币种，记账币金额，外币金额，汇率，客商项目（代码，如有），对方科目    #
#作者：谢尚佑                                                                                #
#时间：2018.7.8                                                                              #
#程序摘要：获取excel系统及OA系统业务数据，形成用友T3可导入的数据格式，可以excel文档，也可以xml     #
#############################################################################################1

from sql import MSSQL , Voucher
#import sqlite3
from operator import itemgetter
#from __builtin__ import str

def main():
    y=2018
    m=7
    sql=[]
    #list类型，日期，外币币种，比外金额，本位币金额，借方科目，借方单位，借方科目币种，贷方科目，贷方单位，贷方科目币种，备注，ref)

    msquery=MSSQL()

    bank_other_sql ="""select 'bko', a.日期,isnull(a.外币币种,''),isnull(a.外币金额,0),isnull(a.本位币金额,0),a.借方科目,b.账套 as 借方账套,b.币种 as 借方币种,a.贷方科目,c.账套 as 贷方账套,c.币种 as 贷方币种,
    a.备注,''  from 银行其他收支 a inner join 科目表 b on a.借方科目=b.科目名称  inner join 科目表 c on a.贷方科目=c.科目名称   where month(a.日期)=%d and year(a.日期)=%d
        """
    bank_other_re=msquery.Sqlexe(bank_other_sql%(m,y))

    if bank_other_re:
        #print _bank_other_re
        sql.extend( bank_other_re)


    ap_sql = """select 'ap',日期,case when 币种='CNY' then '' else 币种 end as cur ,case when 币种='CNY' then 0 else 不含税金额 end as famt,case when 币种='CNY' then  不含税金额  else 0 end as amt,
                '',账套,'','','','',合同号, 供应商     from 采购收货_主表    where month(日期)=%d and year(日期)=%d    """

    ap_sql_re = msquery.Sqlexe(ap_sql%(m,y))

    if ap_sql_re:
        sql.extend(ap_sql_re)
        # type,日期，币种，外币金额，本位币金额，‘’，账套，‘’，‘’，‘’，‘’，合同号，供应商

    ar_sql = """select 'ar',提单日期 ,case when 币种='CNY' then '' else 币种 end as cur, case when 币种='CNY' then 0 else 不含税金额 end as famt,case when 币种='CNY' then  不含税金额  else 0 end as amt ,
                '',账套 ,'','','','',合同号, 客户    from 销售发货_主表    where month(提单日期)=%d and year(提单日期)=%d    """

    ar_sql_re=msquery.Sqlexe(ar_sql%(m,y))

    if ar_sql_re:
        sql.extend(ar_sql_re)
        # type,日期，币种，外币金额，本位币金额，‘’，账套，‘’，‘’，‘’，‘’，合同号，客户


    bank_rec_sql = """select 'br',a.日期,a.币种,a.手续费,a.实收金额,a.收款银行, a.账套,'',b.金额 ,b.客户,b.币种,b.合同号 ,b.excelserverrcid
                from 资金收取_主表 a inner join 资金收取_明细 b  on a.excelserverrcid=b.excelserverrcid where month(a.日期)=%d and year(a.日期)=%d
            """
    bank_rec_sql_re = msquery.Sqlexe(bank_rec_sql%(m,y))

    if bank_rec_sql_re:
        sql.extend(bank_rec_sql_re)
        # type,日期，币种，手续费，银行金额，收款银行，账套，‘’，金额（字符型），客户,收款币种，合同号，b.excelserverrcid


    bank_pay_sql = """select 'bp',支付时间,case when 币种='CNY' then '' else 币种 end as cur,case when 币种='CNY' then 0 else 付款金额 end as famt,case when 币种='CNY' then  付款金额 else 0 end as amt,
                收款人,payer,'', 支付银行,'','',cno,''   from 付款明细OA_excel where month(支付时间)=%d and year(支付时间)=%d """

    bank_pay_sql_re = msquery.Sqlexe(bank_pay_sql%(m,y))

    if bank_pay_sql_re:
        sql.extend(bank_pay_sql_re)
        # type,日期，币种，外币金额，本位币金额，供应商，账套，‘’，支付银行，‘’,‘’，合同号，‘’

    result=sorted(sql,key=itemgetter(1,12))
    #print result[0]

    #formate the result 日期，摘要，科目，币种，记账币金额，外币金额，对方科目
    unit={1:'深圳市华讯方舟企业服务有限公司'.decode('utf-8'),
          2:'华讯方舟企业服务有限公司'.decode('utf-8')

          }
    sql_km='''select 科目名称,科目代码 from 科目表
            '''
    sql_kh='''select ccusname,ccuscode from customer
            '''
    sql_gys='''select cvenname,cvencode from vendor
                '''
    dist_km=dist(msquery.Sqlexe(sql_km))
    msquery.db='003_2018'
    dist_kh_sz=dist(msquery.Sqlexe(sql_kh))
    dist_gys_sz=dist(msquery.Sqlexe(sql_gys))
    


    voucher_make=Voucher()
    for _re in result:

        re=voucher_make.voucher_make(_re,unit[1])
        if re:
            for detail in re:




if __name__=='__main__':
    main()
