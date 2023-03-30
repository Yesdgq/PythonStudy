
import mysql.connector

mydb = mysql.connector.connect(
  host="localhost",       # 数据库主机地址
  user="root",    # 数据库用户名
  passwd="1234abcd"   # 数据库密码
)
 
print(mydb)

mycursor = mydb.cursor()
print(mycursor)
# mycursor.execute("CREATE DATABASE runoob_db")

mycursor.execute("SHOW DATABASES")
 
for x in mycursor:
  print(x)



mydb = mysql.connector.connect(
  host="localhost",
  user="root",
  passwd="1234abcd",
  database="runoob_db"
)

mycursor = mydb.cursor()
# mycursor.execute("CREATE TABLE sites (name VARCHAR(255), url VARCHAR(255))")

# 没有主键 追加个主键
# mycursor.execute("ALTER TABLE sites ADD COLUMN id INT AUTO_INCREMENT PRIMARY KEY")



# mycursor.execute("CREATE TABLE students (id INT AUTO_INCREMENT PRIMARY KEY, name VARCHAR(255), sex VARCHAR(255), age VARCHAR(255), phone VARCHAR(255), address VARCHAR(255), url VARCHAR(255))")

# 添加column
# mycursor.execute("ALTER TABLE students ADD COLUMN age VARCHAR(255)")

# table插入数据
sql = "INSERT INTO students (name, sex, age, phone) VALUES (%s, %s, %s, %s)"
val = [
  ('张三', '男', '19', '13901234123'),
  ('李四', '男', '29', '13901234123'),
  ('王武', '男', '39', '13561234123'),
  ('赵麻子', '男', '9', '1870123410'),
  ('张三', '男', '19', '13901234123'),
  ('李四', '男', '29', '13901234123'),
  ('王武', '男', '39', '13561234123'),
  ('小沙', '女', '9', '1870123410'),
  ('大鑫子', '男', '19', '13901234123'),
  ('阿彪', '男', '29', '13901234123'),
  ('佳雯', '女', '39', '13561234123'),
  ('赵麻子', '男', '9', '1870123410'),
]

# mycursor.executemany(sql, val)
# mydb.commit()  

# for i in range(0, 100):
#   print(i)
#   mycursor.executemany(sql, val)
# mydb.commit()    # 数据表内容有更新，必须使用到该语句
 
# print(mycursor.rowcount, "记录插入成功。")



mydb = mysql.connector.connect(
  host="localhost",
  user="root",
  passwd="1234abcd",
  database="runoob_db"
)

mycursor = mydb.cursor()
mycursor.execute("SELECT * FROM students WHERE name = '李四'")
 
myresult = mycursor.fetchall()     # fetchall() 获取所有记录
 

# for x in myresult:
#   print(x)