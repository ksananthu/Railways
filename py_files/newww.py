import sqlite3

conn = sqlite3.connect("E:\\pycharm_files\\AccountKFC\\Database\\May_2019\\May_2019.db")
c = conn.cursor()

Date = '2019-05-29'
c.execute("SELECT LOP, FOP, LLT, FLT, LL, WC, DFC, GST, DC, VD, EB, CC, UC, OsCld, MISC AUCTION FROM cashbook"
          " WHERE {} = '{}' ".format('Date', Date))

n_update = c.fetchall()
print(n_update)

sum_new = 0
for i in n_update:
    for j in i:
        print(j)
        sum_new = sum_new + int(j)


print(sum_new)