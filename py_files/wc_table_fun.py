import sqlite3
import json
import os
import pandas as pd
from py_files.wc_xlsx_fun import wc_xlsx_writer


dire = os.getcwd()


def clear():
    # for windows
    if os.name == 'nt':
        _ = os.system('cls')
    else:
        _ = os.system('clear')


def wc_table_create():
    # # ...............................setting default month and year.............................................
    with open(str(dire) + '\\py_files\\m_y_choice.json') as json_file:
        data = json.load(json_file)
    y = str(data).replace("['", '').replace("']", '')

    # # ...........selecting default year........................
    with open(str(dire) + '\\py_files\\y_choice.json') as json_file2:
        data_y = json.load(json_file2)
    year = str(data_y).replace("['", '').replace("']", '')
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

    while True:
        add_view = input('Add(a) or View(v)...........: ')
        if str(add_view) == 'a':
            while True:
                print('\n')

                # #........................Input.................................
                Date_si = input('Date..............: ')
                if int(Date_si) < 10:
                    Date = str(year) + '-' + str(month) + '-0' + str(Date_si)
                else:
                    Date = str(year) + '-' + str(month) + '-' + str(Date_si)

                c.execute("select TOTAL from wc WHERE {} LIKE '%{}%'".format('Date', Date))
                date_check = c.fetchall()

                if len(date_check) != 0:
                    print('\n')
                    print('...............Warning.................')
                    print('Entry already exists on ' + Date + '\n')
                    break

                c.execute("select WC from CashBook WHERE {} LIKE '%{}%'".format('Date', Date))
                v_cash = c.fetchall()

                if len(v_cash) == 0:
                    print('\n')
                    print('...............Warning.................')
                    print('Add entry to Cashbook on ' + Date + '\n')
                    break
                else:
                    v_cash = str(v_cash).replace('[(', '').replace(',)]', '')
                    WHA = int(v_cash)
                    print('Wharfage...........: ' + v_cash)

                ST = input('ST................: ')
                if ST.isdigit() == False:
                    ST = 0

                TOTAL = int(WHA) + int(ST)
                print('TOTAL.............: ' + str(TOTAL))

                c.execute("INSERT INTO wc (Date, WHA, ST, TOTAL) VALUES (?, ?,?, ?)",
                          (Date, WHA, ST, TOTAL))

                conn.commit()

                print('\n')

                again_op = input('Add more(any key) or Stop(s) or Edit any entry(e)...........: ')

                if str(again_op) == 's':
                    break

                # #...............Edit function........................................................
                elif str(again_op) == 'e':
                    print('\n')
                    print('............Options.............')
                    print('ST, exit')
                    print('\n')

                    while True:
                        option_list = ['ST', 'exit']
                        edit_in = input('Select entry to edit :')

                        if edit_in not in option_list:
                            print('\n')
                            print('.........Select from below options...........')
                            print('ST, exit')
                            print('\n')
                            continue

                        elif edit_in == 'exit':
                            print('\n')
                            break

                        else:
                            # >>>>>>>> update >>>>>>>>>>>
                            new_edit = input(str(edit_in) + '......: ')
                            c.execute("UPDATE wc SET {} = '{}' WHERE Date = '{}'".format(edit_in, new_edit, Date))
                            conn.commit()

                            # >>>>>>> total >>>>>>>>>>>>>>>
                            c.execute("SELECT WHA, ST FROM wc WHERE {} = '{}' ".format('Date', Date))

                            n_update = c.fetchall()

                            n_total = 0
                            for i in n_update:
                                for j in i:
                                    n_total = n_total + int(j)

                            c.execute("UPDATE wc SET Total = '{}' WHERE Date = '{}'".format(n_total, Date))
                            conn.commit()

                            print('\n')
                            print('..........entry updated...........')
                            print('\n')
                            continue

                    break
                # #...............Edit function ends........................................................

                else:
                    continue

        elif str(add_view) == 'v':
            clear()
            print('............WC Table.................')
            print('\n')
            pd.set_option('display.max_columns', 500)
            df = pd.read_sql_query(
                "SELECT * FROM wc ORDER BY Date ASC", conn)
            print(df)
            print('\n')

            # #.......................Delete function..............................
            while True:
                dlt_opt = input('Dlt a entry(d) or Continue to print menu(c): ')
                if dlt_opt == 'd':
                    print('\n')

                    dlt_date1 = input('Date of a entry to delete : ')
                    if int(dlt_date1) < 10:
                        dlt_date = str(year) + '-' + str(month) + '-0' + str(dlt_date1)
                    else:
                        dlt_date = str(year) + '-' + str(month) + '-' + str(dlt_date1)

                    c.execute("DELETE FROM wc WHERE {} = '{}'".format('Date', dlt_date))
                    conn.commit()

                    print('\n')
                    print('...........entry deleted.........')
                    print('\n')

                    pd.set_option('display.max_columns', 500)
                    df = pd.read_sql_query(
                        "SELECT * FROM wc  ORDER BY Date ASC", conn)
                    print(df)
                    print('\n')
                    continue
                elif dlt_opt == 'c':
                    break
                else:
                    continue
            # #...................Delete function ends...........................

            break
        else:
            continue

    # # ......................option print or exist...........................................................
    while True:
        print('\n')
        p_or_exit = input('Print the table(p) or go back to the main menu(e) : ')
        print('\n')
        if str(p_or_exit) == 'p':
            wc_xlsx_writer()
            break
        elif str(p_or_exit) == 'e':
            break
        else:
            continue



# wc_table_create()



