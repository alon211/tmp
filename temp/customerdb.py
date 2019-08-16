import pyodbc
import os
import pysnooper
import traceback
import sys
cur_path,file=os.path.split(os.path.realpath(sys.argv[0])) 
print(cur_path)
@pysnooper.snoop(os.path.join(cur_path,'db.log'),depth=2)
def test_db():
    try:
        # cnxn=pyodbc.connect('DRIVER={SQL Server\};SERVER=172.30.125.221;DATABASE=SMT_Manager;UID=sa;PWD=0123456789',timeout=5)
        cnxn=pyodbc.connect('DRIVER={SQL Server\};SERVER=DESKTOP-AR24T88\SAK;DATABASE=SMT_Manage;UID=sa;PWD=111',timeout=5)        
        cursor=cnxn.cursor()
        sql = "select * from dbo.orderlist where 订单号='3111809906' and 行号='08'"
        cursor.execute(sql)
        rs= cursor.fetchall()
        print(rs[0].PCBA版本)
    except Exception as e:
        # a=traceback.format_exc(limit=1)
        pass

test_db()