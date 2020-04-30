import pymysql

# conn = pymysql.connect(host='localhost', port=3307, user='aplanmis_local', passwd='aplanmis_local', db='aplanmis_foshan')
conn = pymysql.connect()
cursor = conn.cursor()
sql = 'select mat_name ,mat_code from aea_item_mat where is_deleted=%s'
cursor.execute(sql, ('0',))
fetchall = cursor.fetchall()
print(fetchall)
print(len(fetchall))
cursor.close()

conn.close()
d = dict(fetchall)
print(d)
