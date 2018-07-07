# -*- coding:utf-8 -*-  

from sql import MSSQL,MYSQL
import sqlite3
import sys

def main():
    con=sqlite3.connect(":memory:")
    cur=con.cursor()
    y=2018
    m=6
    table_sql="""create table voucher (ctype  varchar(20),
                                        ddate date,
                                        cur varchar(5),
                                        famt numeric(10,2),
                                        amt numeric(10,2),
                                        dr varchar(20),
                                        dr_unit varchar(100),
                                        dr_cur varchar(5),
                                        cr varchar(20),
                                        cr_unit varchar(100),
                                        cr_cur varchar(5),
                                        mark  varchar(30),
                                        ref varchar(50) )
                """   #类型，日期，外币币种，比外金额，本位币金额，借方科目，借方单位，借方科目币种，贷方科目，贷方单位，贷方科目币种，备注，ref

    cur.execute(table_sql)
    msquer=MSSQL()
    
    bank_other_sql ="""select  a.日期,isnull(a.外币币种,''),isnull(a.外币金额,0),isnull(a.本位币金额,0),a.借方科目,b.账套 as 借方账套,b.币种 as 借方币种,a.贷方科目,c.账套 as 贷方账套,c.币种 as 贷方币种,a.备注 
            from 银行其他收支 a inner join 科目表 b on a.借方科目=b.科目名称  inner join 科目表 c on a.贷方科目=c.科目名称   where month(a.日期)=%d and year(a.日期)=%d 
        """
    bank_other_re=msquer.Sqlexe(bank_other_sql%(m,y))
    sql=[]
    for _bank_other_re in bank_other_re:
        #print _bank_other_re
        sql.append( """insert into voucher values('bko','%s','%s',%8.2f,%8.2f,'%s','%s','%s','%s','%s','%s','%s','')"""%_bank_other_re
            )
        
    
    ap_sql = """select 账套 ,日期, 供应商, case when 币种='CNY' then '' else 币种 end as cur ,case when 币种='CNY' then 0 else 不含税金额 end as famt,case when 币种='CNY' then  不含税金额  else 0 end as amt,合同号     from 采购收货_主表    where month(日期)=%d and year(日期)=%d    """
    
    ap_sql_re=msquer.Sqlexe(ap_sql%(m,y))
    
    for _ap_sql_re in ap_sql_re:
        sql.append("""insert into voucher values('ap','%s','%s',%8.2f,%8.2f,'','%s','','','','','%s','%s')"""%(_qp_sql_re[1],_qp_sql_re[3],_qp_sql_re[4],_qp_sql_re[5],_qp_sql_re[6],_qp_sql_re[2])) # type,日期，币种，外币金额，本位币金额，‘’，账套，‘’，‘’，‘’，‘’，合同号，供应商
                    
    
    ar_sql = """select 账套 ,提单日期, 客户, case when 币种='CNY' then '' else 币种 end as cur ,case when 币种='CNY' then 0 else 不含税金额 end as famt,case when 币种='CNY' then  不含税金额  else 0 end as amt ,合同号    from 销售发货_主表    where month(提单日期)=%d and year(提单日期)=%d    """
    
    ar_sql_re=msquery.Sqlexe(ar_sql%(m,y))
    
    for _ar_sql_re in ar_sql_re:
        sql.append("""insert into voucher values('ar','%s','%s',%8.2f,%8.2f,'','%s','','','','','%s','%s')"""%(_qp_sql_re[1],_qp_sql_re[3],_qp_sql_re[4],_qp_sql_re[5],_qp_sql_re[6],_qp_sql_re[2])) # type,日期，币种，外币金额，本位币金额，‘’，账套，‘’，‘’，‘’，‘’，合同号，客户
      
        
    bank_rec_sql = """select a.账套 ,a.日期,a.收款银行,a.实收金额,a.手续费,a.币种,b.合同号,b.客户,b.金额,b.币种 ,b.备注,b.excelserverrcid
                from 资金收取_主表 a inner join 资金收取_明细 b  on a.excelserverrcid=b.excelserverrcid where month(a.日期)=%d and year(a.日期)=%d
            """
            
    bank_pay_sql = """select payer,cno,收款人,付款金额,币种,支付时间,支付银行   from 付款明细OA_excel where month(支付时间)=%d and year(支付时间)=%d """
    
    for s in sql:
        print s
        cur.execute(s)
    con.commit()
    
    
if __name__=='__main__':
    main()
