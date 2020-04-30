import mysql.connector

conn = mysql.connector.connect(user='aplanmis_local', password='aplanmis_local',
                               database='aplanmis_foshan', port='3307', charset='utf8'
                               , host='localhost')
cursor = conn.cursor()
sql = 'select mat_name ,mat_code from aea_item_mat where is_deleted=%s'
cursor.execute(sql, ('0',))
fetchall = cursor.fetchall()
print(fetchall)
print(len(fetchall)) 
cursor.close()

conn.close()
