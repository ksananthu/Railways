import sqlite3
import json
import os
import pandas as pd
from py_files.flt_xlsx_fun import flt_xlsx_writer


dire = os.getcwd()


def clear():
    # for windows
    if os.name == 'nt':
        _ = os.system('cls')
    else:
        _ = os.system('clear')


def flt_table_create():
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

                c.execute("select TOTAL from flt WHERE {} LIKE '%{}%'".format('Date', Date))
                date_check = c.fetchall()

                if len(date_check) != 0:
                    print('\n')
                    print('...............Warning.................')
                    print('Entry already exists on ' + Date + '\n')
                    break

                c.execute("select flt from CashBook WHERE {} LIKE '%{}%'".format('Date', Date))
                v_cash = c.fetchall()

                if len(v_cash) == 0:
                    print('\n')
                    print('...............Warning.................')
                    print('Add entry to Cashbook on ' + Date + '\n')
                    break
                else:
                    v_cash = str(v_cash).replace('[(', '').replace(',)]', '')
                    FRE = int(v_cash)
                    print('Freight...........: ' + v_cash)

                NL = input('No of Lug.........: ')
                if NL.isdigit() == False:
                    NL = 0

                WT = input('Wt in Qtls........: ')
                try:
                    WT = float(WT)
                except ValueError:
                    WT = 0

                DFC = input('DFC...............: ')
                if DFC.isdigit() == False:
                    DFC = 0

                ST = input('ST................: ')
                if ST.isdigit() == False:
                    ST = 0

                TOTAL = int(FRE) + int(DFC) + int(ST)
                print('TOTAL.............: ' + str(TOTAL))

                c.execute("INSERT INTO flt (Date, NL, WT, FRE, DFC, ST, TOTAL) VALUES (?, ?,?, ?, ?, ?, ?)",
                          (Date, NL, WT, FRE, DFC, ST, TOTAL))

                conn.commit()

                print('\n')
                again_op = input('Add more(any key) or Stop(s) or Edit any entry(e)...........: ')

                if str(again_op) == 's':
                    break

                # #...............Edit function........................................................
                elif str(again_op) == 'e':
                    print('\n')
                    print('............Options.............')
                    print('No of Lug(NL), Wt in Qtls(WT), DFC, ST, exit')
                    print('\n')

                    while True:
                        option_list = ['NL', 'WT', 'ST', 'DFC', 'exit']
                        edit_in = input('Select entry to edit :')

                        if edit_in not in option_list:
                            print('\n')
                            print('.........Select below options...........')
                            print('No of Lug(NL), Wt in Qtls(WT), DFC, ST, exit')
                            print('\n')
                            continue

                        elif edit_in == 'exit':
                            print('\n')
                            break

                        elif edit_in == 'WT':
                            new_edit = input(str(edit_in) + '......: ')
                            try:
                                new_edit = float(new_edit)
                            except ValueError:
                                new_edit = 0

                            c.execute("UPDATE flt SET {} = '{}' WHERE Date = '{}'".format(edit_in, new_edit, Date))
                            conn.commit()
                            continue

                        else:
                            #>>>>>>>>>>>> update >>>>>>>>>>>>>>>>>>>
                            new_edit = input(str(edit_in) + '......: ')
                            if new_edit.isdigit() == False:
                                new_edit = 0

                            c.execute("UPDATE flt SET {} = '{}' WHERE Date = '{}'".format(edit_in, new_edit, Date))
                            conn.commit()

                            # >>>>>>> total >>>>>>>>>>>>>>>
                            c.execute("SELECT FRE, DFC, ST FROM flt WHERE {} = '{}' ".format('Date', Date))

                            n_update = c.fetchall()

                            n_total = 0
                            for i in n_update:
                                for j in i:
                                    n_total = n_total + int(j)

                            c.execute("UPDATE flt SET Total = '{}' WHERE Date = '{}'".format(n_total, Date))
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
            print('............flt Table.................')
            print('\n')
            pd.set_option('display.max_columns', 500)
            df = pd.read_sql_query(
                "SELECT * FROM flt  ORDER BY Date ASC", conn)
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

                    c.execute("DELETE FROM flt WHERE {} = '{}'".format('Date', dlt_date))
                    conn.commit()

                    print('\n')
                    print('...........entry deleted.........')
                    print('\n')

                    pd.set_option('display.max_columns', 500)
                    df = pd.read_sql_query(
                        "SELECT * FROM flt  ORDER BY Date ASC", conn)
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
            flt_xlsx_writer()
            break
        elif str(p_or_exit) == 'e':
            break
        else:
            continue


# flt_table_create()




