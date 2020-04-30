import common

conn = common.get_mysql_conn()

cursor = conn.cursor()
sql = 'select mat_name ,mat_code from aea_item_mat where is_deleted=%s'
cursor.execute(sql, ('0',))
fetchall = cursor.fetchall()
print(fetchall)
print(len(fetchall))
cursor.close()

conn.close()