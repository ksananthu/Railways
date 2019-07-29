import sqlite3
import os
import json


dire = os.getcwd()


def data_add():
    # # ....................................Select month and year....................................................
    mon_list = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October',
                'November', 'December']
    print('\n')
    print('Months should be given like this "January, February, March, April, May, June, July, August, September,'
          ' October, November,'
          ' December "')
    print('\n')
    m = input('Input month  :  ')
    while m not in mon_list:
        print('\n')
        print('.................Warning !!!!!.................')
        print('Months should be given like this "January, February, March, April, May, June, July, August, September,'
              ' October, November,'
              ' December "')
        print('\n')
        m = input('Input month  :  ')

    y = input('Input year   :  ')
    m_y = str(m) + '_' + str(y)
    mon_yr = [str(m_y)]
    y_choice = [str(y)]

    with open(str(dire) + '\\py_files\\m_y_choice.json', 'w') as outfile:
        json.dump(mon_yr, outfile)

    with open(str(dire) + '\\py_files\\y_choice.json', 'w') as outfile:
        json.dump(y_choice, outfile)

    # # ...........................creating databse dir and opening database.......................................
    if not os.path.exists(str(dire) + '\\Database\\' + m_y):
        os.makedirs(str(dire) + '\\Database\\' + m_y)

    conn = sqlite3.connect(str(dire) + '\\Database\\' + m_y + '\\' + m_y + '.db')
    c = conn.cursor()

    # # ...................................cashbook table.............................................................
    def create_table_cashbook():
        c.execute("CREATE TABLE IF NOT EXISTS CashBook(Date TEXT UNIQUE, LOP INTEGER, FOP INTEGER, LLT INTEGER,"
                  "	FLT INTEGER,"
                  "LL INTEGER, WC INTEGER, KFC INTEGER, DFC INTEGER, GST INTEGER, DC INTEGER, VD INTEGER, EB INTEGER,"
                  " CC INTEGER,"
                  "UC INTEGER, OsCld INTEGER,	MISC INTEGER, AUCTION INTEGER, TOTAL INTEGER, OS INTEGER,	"
                  "POS INTEGER,	vR INTEGER,	CASH INTEGER, Remittance INTEGER)")

    # #...........................lop table......................................................................
    def create_table_lop():
        c.execute("CREATE TABLE IF NOT EXISTS lop(DATE TEXT UNIQUE, PLS INTEGER, WT REAL, Fre INTEGER, KFC INTEGER,"
                  " DFC INTEGER, "
                  "ST INTEGER, TOTAL INTEGER)")

    # #...........................fop table......................................................................
    def create_table_fop():
        c.execute(
            "CREATE TABLE IF NOT EXISTS fop(DATE TEXT UNIQUE, PLS INTEGER, WT REAL, Fre INTEGER, DFC INTEGER, "
            "ST INTEGER, TOTAL INTEGER)")

    # #...........................Pos table......................................................................
    def create_table_pos():
        c.execute("CREATE TABLE IF NOT EXISTS Pos(Date TEXT, amt INTEGER, ID TEXT UNIQUE, Cmt TEXT)")

    # #...........................WC table......................................................................
    def create_table_wc():
        c.execute(
            "CREATE TABLE IF NOT EXISTS wc(Date TEXT UNIQUE, WHA INTEGER, ST INTEGER, TOTAL INTEGER)")

    # #...........................llt table......................................................................
    def create_table_llt():
        c.execute(
            "CREATE TABLE IF NOT EXISTS llt(DATE TEXT UNIQUE, NL INTEGER, WT REAL, Fre INTEGER, KFC INTEGER, DFC INTEGER, "
            "ST INTEGER, TOTAL INTEGER)")

    # #...........................flt table......................................................................
    def create_table_flt():
        c.execute(
            "CREATE TABLE IF NOT EXISTS flt(DATE TEXT UNIQUE, NL INTEGER, WT REAL, Fre INTEGER, DFC INTEGER, "
            "ST INTEGER, TOTAL INTEGER)")

    # #...........................cc table......................................................................
    def create_table_cc():
        c.execute(
            "CREATE TABLE IF NOT EXISTS cc(DATE TEXT UNIQUE, LT INTEGER, AMT INTEGER)")

    create_table_cashbook()
    create_table_lop()
    create_table_fop()
    create_table_pos()
    create_table_wc()
    create_table_llt()
    create_table_flt()
    create_table_cc()


if __name__ == "__main__":
    data_add()




