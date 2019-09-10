import openpyxl
import os,time
import pymysql
from pymysql.cursors import DictCursor


def mysqlquery(query, update=False):
    connection = pymysql.connect(
        host='localhost',
        user='stalker',
        password='1',
        db='stalker_db',
        charset='utf8mb4',
        cursorclass=DictCursor)
    try: 
        cursor = connection.cursor()
        cursor.execute('use stalker_db')
        if update:
            cursor.execute(query)
            connection.commit()
        else:
            cursor.execute(query)
            return cursor.fetchall()
    finally:
        connection.close()
    

if __name__ == '__main__':

    out = []
    outnew = []
    name = []
    i = 0                                       # iterator for name
    wb = openpyxl.load_workbook('test.xlsx')
    sheet = wb.active
    print('begin')
    for rowobject in sheet['F2':'F304']:
        for cell in rowobject:
            out.append('udp://'+cell.value.replace(' ', ''))
    for rowobject in sheet['G2':'G304']:
        for mell in rowobject:
            outnew.append('udp://'+mell.value.replace(' ', ''))
    for rowobject in sheet['H2':'H304']:
        for tell in rowobject:
            name.append(tell.value)    
    for i in range(0,len(out)):
        if out[i] != outnew[i]:
            time.sleep(0.1)
            query = [] 
            query = mysqlquery('select id, cmd, mc_cmd from test where name ="'+str(name[i])+'"')
            print('query = ', query)
            if query[0]['mc_cmd']:
                print('update with replace')
                mysqlquery('update test set cmd = "'+outnew[i]+'", mc_cmd = "'+outnew[i]+'" where cmd = "'+out[i]+ '"', True)
            else:
                mysqlquery('update test set cmd = "'+outnew[i]+'" where cmd = "'+out[i]+ '"', True)
                print('update without replace')
