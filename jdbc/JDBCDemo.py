import mysql.connector
mydb = mysql.connector.connect(
    host="localhost",
    user="root",
    passwd="root"
)
mycursor = mydb.cursor()
mycursor.execute("SHOW DATABASES")

for item in mycursor:
    print(item)