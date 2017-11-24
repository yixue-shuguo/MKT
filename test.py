# -*- coding: utf-8 -*-
"""
Created on Fri Nov 24 14:24:52 2017

@author: Administrator
"""


print('开始')
import pymysql 
import pandas as pd 
from datetime import datetime
now = datetime.now().strftime("%Y-%m-%d  %H-%M-%S")

conn = pymysql.connect('rr-uf647q511648367cso.mysql.rds.aliyuncs.com','crm_root',
                       'Hello2017','ecustomer', charset='utf8')
cursor = conn.cursor()
df = pd.read_excel(r'c:/repeat.xlsx')

L = []
err = []
i = 0
for tel in df['tel']:
    i = i+1 
    if i%100 ==0 :
        print ('运行了%d行'%i)
    else :
        pass
    try:
        sql = '''
select distinct c.keymobile ,c.leadsource ,IF(s.shitingsid is null ,"无","有") ,max_track 
from 
(
select c.clueregsid,c.keymobile,c.leadsource  from ec_clueregs c where c.keymobile ='%d'  and c.deleted = 0 
) c 
left join ec_shitings s 
on c.clueregsid = s.clueregsid 
left join 
(
select t.clueregsid , max(t.modifiedtime)  max_track 
from ec_trackrecords t 
where t.clueregsid = (select c.clueregsid  from ec_clueregs c where c.keymobile ='%d' and c.deleted = 0 )
)t 
on c.clueregsid = t.clueregsid 

             '''%(int(tel),int(tel))
        cursor.execute(sql)
        data = cursor.fetchone()
        L.append(data)
    except Exception as e :
        err.append([tel,str(e)])

result = pd.DataFrame(L,columns = ['电话','渠道','试听','最近联系日期'])
err_result = pd.DataFrame(err,columns = ['电话','错误类型'])
writer = pd.ExcelWriter(r'c:/repeat_result %s.xlsx'%(now))       
result.to_excel(writer,'Sheet1',index = False)
err_result.to_excel(writer,'Sheet2',index = False)


writer.save()
cursor.close()
conn.close()
print('结束')


