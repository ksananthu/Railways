import sqlite3
import json
import os
import pandas as pd
from py_files.pos_xlsx_fun import pos_xlsx_writer


dire = os.getcwd()


def clear():
    # for windows
    if os.name == 'nt':
        _ = os.system('cls')
    else:
        _ = os.system('clear')


def pos_table_create():

    # # ...............................setting default month and year.............................................
    with open(str(dire) + '\\py_files\\m_y_choice.json') as json_file:
        data = json.load(json_file)

    y = str(data).replace("['", '').replace("']", '')

    # # ...........selecting default year........................
    with open(str(dire) + '\\py_files\\y_choice.json') as json_file2:
        data_y = json.load(json_file2)

    year = str(data_y).replace("['", '').replace("']", '')

    # year = '2019'  # # ++++++++++++> year replacement

    # # ---------------------------month slice begin ------------------------
    month = 0
    if y.find('Jan') != -1:
        month = '01'
    elif y.find('Feb') != -1:
        month = '02'
    elif y.find('Mar') != -1:
        month = '03'
    elif y.find('April') != -1:
        month = '04'

    elif y.find('May') != -1:
        month = '05'

    elif y.find('June') != -1:
        month = '06'

    elif y.find('July') != -1:
        month = '07'

    elif y.find('August') != -1:
        month = '08'

    elif y.find('Sep') != -1:
        month = '09'

    elif y.find('Oct') != -1:
        month = '10'

    elif y.find('Nov') != -1:
        month = '11'

    elif y.find('Dec') != -1:
        month = '12'
    # # ---------------------------month slice end ------------------------

    # # ...........................opening Database................................................................
    conn = sqlite3.connect(str(dire) + '\\Database\\' + str(y) + '\\' + str(y) + '.db')
    c = conn.cursor()

    # # ..................option add or view.......................................................................
    while True:
        add_view = input('Add(a) or View(v)...........: ')
        if str(add_view) == 'a':
            while True:
                print('\n')
                Date_si = input('Date.....: ')
                if int(Date_si) < 10:
                    Date = str(year) + '-' + str(month) + '-0' + str(Date_si)
                else:
                    Date = str(year) + '-' + str(month) + '-' + str(Date_si)
                # print(Date)
                amt = input('Amt......: ')
                if amt.isdigit() == False:
                    amt = 0
                ID = input('ID.......: ')
                Cmt = input('Cmt......: ')

                c.execute("INSERT INTO Pos (Date, amt, ID, Cmt) VALUES (?, ?, ?, ?)", (Date, amt, ID, Cmt))
                conn.commit()

                print('\n')
                again_op = input('Add more(any key) or stop(s)...........: ')
                if str(again_op) == 's':
                    break
                else:
                    continue
        elif str(add_view) == 'v':
            clear()
            print('............POS Table.................')
            print('\n')
            pd.set_option('display.width', 100)
            df = pd.read_sql_query(
                "SELECT Date, Amt, ID, Cmt FROM Pos  ORDER BY Date ASC", conn)
            print(df)
            break
        else:
            continue

    # # ......................option print or exist...........................................................
    while True:
        print('\n')
        p_or_exit = input('Print the table(p) or go back to the main menu(e) : ')
        print('\n')
        if str(p_or_exit) == 'p':
            pos_xlsx_writer()
            break
        elif str(p_or_exit) == 'e':
            break
        else:
            continue



