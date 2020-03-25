import sqlite3
#连接到SQlite数据库，不存在则自动创建
conn = sqlite3.connect('sqltest.db')
#创建一个cursor，sqlite在cursor中执行命令：
cursor = conn.cursor()

#构建基础表，来自小米的订单表，每行对应一只主轴，但是不一定有发货，需结合营销唐小姐的表才能确定主轴去向
cursor.execute('DROP TABLE milist')
cursor.execute('''CREATE TABLE milist
      (id int primary key, 
      typeid int,
      typename text,
      productid text,
      orderid int,
      ordertime int,
      storetime int,
      saletime int, 
      bearingid int,
      bearing text)''')

#导入excel表格存储到sqlite
import xlrd
data = xlrd.open_workbook('milist.xlsx')
table = data.sheet_by_index(0)

nrows = table.nrows
ncols = table.ncols

for i in range(nrows):
    cursor.executemany('INSERT INTO milist VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)', [table.row_values(i)])

#查询指定型号的主轴的时间数据
date=[]
spindles=[]
typeid=15002744
#use gruop by get date for one type of spindle
cursor.execute('SELECT milist.storetime FROM milist WHERE milist.typeid=? GROUP BY milist.storetime',[typeid])
values=cursor.fetchall() 
#use where get spndles sotred each day
for i in values:
    cursor.execute('SELECT * FROM milist WHERE milist.storetime=?',i)
    counts=len(cursor.fetchall())
    date.append(xlrd.xldate.xldate_as_datetime(i[0],0))
    spindles.append(counts)

plt.figure(figsize=(20,3))    
plt.bar(date,spindles,label='typeid ='+str(typeid))
plt.legend(loc="upper right") 
plt.xlabel('date') 
plt.ylabel('spindles pc') 
plt.title('store distribution') 
plt.show()    
    
for i,row in enumerate(cursor.execute('SELECT * FROM milist')):
    producttime = xlrd.xldate.xldate_as_datetime((table.cell(i,6).value),0)-xlrd.xldate.xldate_as_datetime((table.cell(i,5).value),0)
    stocktime = xlrd.xldate.xldate_as_datetime((table.cell(i,7).value),0)-xlrd.xldate.xldate_as_datetime((table.cell(i,6).value),0)
    #print(row)
    #print(producttime.days,stocktime.days)
    
#构建发货表，来自营销唐小姐
#退货和取回的时间是退回后检测开单的时间
#其他是发货的时间
cursor.execute('DROP TABLE tanglist')
cursor.execute('''CREATE TABLE tanglist
      (id int primary key, 
      typeid int,
      typename text,
      saletime int,      
      orderform text,
      ordertype text,
      customerid int,
      customer text,
      productid text)''')

#导入excel表格存储到sqlite
data = xlrd.open_workbook('tanglist.xlsx')
table = data.sheet_by_index(0)

nrows = table.nrows
ncols = table.ncols

for i in range(nrows):
    cursor.executemany('INSERT INTO tanglist VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)', [table.row_values(i)])

cursor.execute('SELECT * FROM tanglist ')
values=cursor.fetchall() 
print(len(values))

#模糊搜索，证明米小姐根据营销数据查询结果未见异常，序号与时间都对的上
cursor.execute('SELECT milist.productid,milist.saletime FROM milist WHERE milist.typeid = ? ',[tid])
values=cursor.fetchall() 
productid_list=[]
saletime_list=[]
pickid_list=[]
null_count=[]
for i,v in enumerate(values):
    pid=v[0]
    stime=v[1]
    productid_list.append(pid)
    saletime_list.append(stime)
    cursor.execute('SELECT tanglist.id FROM tanglist WHERE tanglist.productid LIKE ? AND tanglist.saletime=? ',['%'+pid+'%',stime])
    values=cursor.fetchall() 
    if len(values)<1:
        pickid_list.append('null')   
        null_count.append(1)
    else:
        pickid_list.append(values[0][0])
print('id like and time match ',len(pickid_list)-len(null_count))

#构建退货表，手动删除转返修，“/” 替换为数字0
cursor.execute('DROP TABLE tuilist')
cursor.execute('''CREATE TABLE tuilist
      (id int primary key, 
      returntime int,
      typename text,
      productid text,
      saletime int,
      typeid int,
      bearing text,
      bnum int)''')

#导入excel表格存储到sqlite
import xlrd
data = xlrd.open_workbook('2018tui.xlsx')
table = data.sheet_by_index(0)

nrows = table.nrows
ncols = table.ncols

for i in range(nrows):
    cursor.executemany('INSERT INTO tuilist VALUES (?, ?, ?, ?, ?, ?, ?, ?)', [table.row_values(i)])
    
data = xlrd.open_workbook('2019tui.xlsx')
table = data.sheet_by_index(0)

nrows = table.nrows
ncols = table.ncols

for i in range(nrows):
    cursor.executemany('INSERT INTO tuilist VALUES (?, ?, ?, ?, ?, ?, ?, ?)', [table.row_values(i)])    
    
cursor.execute('SELECT * FROM tuilist ')
values=cursor.fetchall() 
print(len(values))

#构建返修表，过保需根据时间长度判定，
cursor.execute('DROP TABLE fanlist')
cursor.execute('''CREATE TABLE fanlist
      (id int primary key, 
      typename text,
      productid text,
      typeid int,
      saletime1 int,
      saletime2 int,
      returntime int, 
      type text)''')

#导入excel表格存储到sqlite
import xlrd
data = xlrd.open_workbook('2018fan.xlsx')
table = data.sheet_by_index(0)

nrows = table.nrows
ncols = table.ncols

for i in range(nrows):
    cursor.executemany('INSERT INTO fanlist VALUES (?, ?, ?, ?, ?, ?, ?, ?)', [table.row_values(i)])
    
data = xlrd.open_workbook('2019fan.xlsx')
table = data.sheet_by_index(0)

nrows = table.nrows
ncols = table.ncols

for i in range(nrows):
    cursor.executemany('INSERT INTO fanlist VALUES (?, ?, ?, ?, ?, ?, ?, ?)', [table.row_values(i)])    
    
cursor.execute('SELECT * FROM fanlist WHERE fanlist.id = ? ',[1])
values=cursor.fetchall() 
print(values)

def tuiinspect():
    
    #查询milist包含主轴型号
    cursor.execute('SELECT typeid FROM milist group by typeid') 
    row=cursor.fetchall()
    
    #对查询到的每个型号进行进一步查询
    
    for i,tid in enumerate(row):
        
        #按型号提取发货主轴的名称与总数
        cursor.execute('SELECT typename FROM milist WHERE typeid = ? ',[tid[0]])
        values=cursor.fetchall()
        total=len(values)     
        cursor.execute('SELECT typename FROM milist WHERE typeid = ? group by typename',[tid[0]])
        values=cursor.fetchall()
        name=values[0]    
        
        #按型号提取序号
        cursor.execute('SELECT milist.productid FROM milist WHERE milist.typeid = ?',[tid[0]]) 
        row=cursor.fetchall()
        
        #按上述型号序号查询有无退货返修
        
        #有退货无拆检(手动删除转返修)
        tuiwob=[]
        for i,sid in enumerate(row):
            cursor.execute('SELECT tuilist.productid FROM tuilist WHERE tuilist.productid LIKE ? AND tuilist.bnum = ? ',['%'+sid[0]+'%',0])
            values=cursor.fetchall() 
            if len(values)> 0:   
                #对查询到的主轴序号备份到数组以便后续快速查询
                tuiwob.append(values[0])

            
        #有退货有拆检
        tuiwb=[]
        for i,sid in enumerate(row):
            cursor.execute('SELECT tuilist.productid FROM tuilist WHERE tuilist.productid LIKE ? AND tuilist.bnum > ? ',['%'+sid[0]+'%',0])
            values=cursor.fetchall() 
            if len(values)> 0:   
                #对查询到的主轴序号备份到数组以便后续快速查询
                tuiwb.append(values[0])
        
        #轴承异常返修
        for i,sid in enumerate(row):
            cursor.execute('SELECT fanlist.productid FROM fanlist WHERE fanlist.productid LIKE ? AND fanlist.type LIKE ? ',['%'+sid[0]+'%','%'+'轴承异常'+'%'])
            values=cursor.fetchall() 
            if len(values)> 0:        
                tuiwb.append(values[0])

        #过载返修
        for i,sid in enumerate(row):
            cursor.execute('SELECT fanlist.productid FROM fanlist WHERE fanlist.productid LIKE ? AND fanlist.type LIKE ? ',['%'+sid[0]+'%','%'+'过载'+'%'])
            values=cursor.fetchall() 
            if len(values)> 0:        
                tuiwb.append(values[0])
        
        #返修（含轴承异常与过载）
        fan=[]
        for i,sid in enumerate(row):
            cursor.execute('SELECT fanlist.productid FROM fanlist WHERE fanlist.productid LIKE ? ',['%'+sid[0]+'%'])
            values=cursor.fetchall() 
            if len(values)> 0:        
                fan.append(values[0]) 
                   
        #从返修中删除轴承异常与过载
        for i,a in enumerate(fan):
            if len(tuiwb) > 0:
                for j,b in enumerate(tuiwb):
                    if a == b:
                        fan.pop(i)

        tuiwob.extend(fan)      

        print(name,total,'pc','wob '+str(round(100*len(tuiwob)/total,2))+'%','wb '+str(round(100*len(tuiwb)/total,2))+'%')
        
        #查询指定型号的主轴的入库时间数据
        date=[]
        spindles=[]
        typeid=tid[0]
        #use gruop by get date for one type of spindle
        cursor.execute('SELECT milist.storetime FROM milist WHERE milist.typeid=? GROUP BY milist.storetime',[typeid])
        values=cursor.fetchall() 
        #use where get spndles sotred each day
        for i in values:
            cursor.execute('SELECT * FROM milist WHERE milist.storetime=?',i)
            counts=len(cursor.fetchall())
            date.append(xlrd.xldate.xldate_as_datetime(i[0],0))
            spindles.append(counts)

        plt.figure(figsize=(20,3))    
        plt.bar(date,spindles,label='typeid ='+str(typeid))
        plt.legend(loc="upper right") 
        plt.xlabel('date') 
        plt.ylabel('spindles pc') 
        plt.title('store distribution') 
        plt.show()  
        
        #查询指定型号的主轴的销售时间数据
        date=[]
        spindles=[]
        typeid=tid[0]
        #use gruop by get date for one type of spindle
        cursor.execute('SELECT milist.saletime FROM milist WHERE milist.typeid=? GROUP BY milist.saletime',[typeid])
        values=cursor.fetchall() 
        #use where get spndles sotred each day
        for i in values:
            cursor.execute('SELECT * FROM milist WHERE milist.saletime=?',i)
            counts=len(cursor.fetchall())
            date.append(xlrd.xldate.xldate_as_datetime(i[0],0))
            spindles.append(counts)

        plt.figure(figsize=(20,3))    
        plt.bar(date,spindles,label='typeid ='+str(typeid))
        plt.legend(loc="upper right") 
        plt.xlabel('date') 
        plt.ylabel('spindles pc') 
        plt.title('sale distribution') 
        plt.show()         

tuiinspect()

#关闭cursor
#关闭conn
cursor.close()
#提交事务：
conn.commit()
conn.close()
