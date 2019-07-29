import sqlite3
import json
import os
import pandas as pd
from py_files.cashbook_xlsx_fun import cashbook_xlsx_writer


def clear():
    # for windows
    if os.name == 'nt':
        _ = os.system('cls')
    else:
        _ = os.system('clear')


dire = os.getcwd()


def cashbook_table_create():
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

    # # ..................option add or view.......................................................................
    while True:
        add_view = input('Add(a) or View(v)...........: ')
        if str(add_view) == 'a':
            while True:
                print('\n')

                # #........................Input.................................
                Date_si = input('Date.......: ')
                if int(Date_si) < 10:
                    Date = str(year) + '-' + str(month) + '-0' + str(Date_si)
                else:
                    Date = str(year) + '-' + str(month) + '-' + str(Date_si)

                c.execute("select TOTAL from CashBook WHERE {} LIKE '%{}%'".format('Date', Date))
                date_check = c.fetchall()

                if len(date_check) != 0:
                    print('\n')
                    print('...............Warning.................')
                    print('Entry already exists on ' + Date + '\n')
                    break

                LOP = input('LOP........: ')
                if LOP.isdigit() == False:
                    LOP = 0

                FOP = input('FOP........: ')
                if FOP.isdigit() == False:
                    FOP = 0

                LLT = input('LLT........: ')
                if LLT.isdigit() == False:
                    LLT = 0

                FLT = input('FLT........: ')
                if FLT.isdigit() == False:
                    FLT = 0

                LL = input('LL.........: ')
                if LL.isdigit() == False:
                    LL = 0

                WC = input('WC.........: ')
                if WC.isdigit() == False:
                    WC = 0

                KFC = input('KFC........: ')
                if KFC.isdigit() == False:
                    KFC = 0

                DFC = input('DFC........: ')
                if DFC.isdigit() == False:
                    DFC = 0

                GST = input('GST........: ')
                if GST.isdigit() == False:
                    GST = 0

                DC = input('DC.........: ')
                if DC.isdigit() == False:
                    DC = 0

                VD = input('VD.........: ')
                if VD.isdigit() == False:
                    VD = 0

                EB = input('EB.........: ')
                if EB.isdigit() == False:
                    EB = 0

                CC = input('CC.........: ')
                if CC.isdigit() == False:
                    CC = 0

                UC = input('UC.........: ')
                if UC.isdigit() == False:
                    UC = 0

                OsCld = input('OsCld......: ')
                if OsCld.isdigit() == False:
                    OsCld = 0

                MISC = input('MISC.......: ')
                if MISC.isdigit() == False:
                    MISC = 0

                AUCTION = input('AUCTION....: ')
                if AUCTION.isdigit() == False:
                    AUCTION = 0

                TOTAL2 = int(DC) + int(VD) + int(EB) + int(CC) + int(UC) + int(OsCld) + int(MISC) + int(AUCTION) \
                         + int(KFC)
                TOTAL = int(LOP) + int(FOP) + int(LLT) + int(FLT) + int(LL) + int(WC) + int(DFC) + int(GST) + TOTAL2

                print('Total......: ' + str(TOTAL))

                OS = input('OS.........: ')
                if OS.isdigit() == False:
                    OS = 0

                POS = input('POS........: ')
                if POS.isdigit() == False:
                    POS = 0

                vR = input('vR.........: ')
                if vR.isdigit() == False:
                    vR = 0

                CASH = int(TOTAL) - (int(OS) + int(POS) + int(vR))
                print('Cash.......: ' + str(CASH))

                Remittance = int(POS) + int(vR) + int(CASH)
                print('Remittance.: ' + str(Remittance))

                c.execute(
                    "INSERT INTO CashBook (Date, LOP, FOP, LLT, FLT, LL, WC, KFC, DFC, GST, DC, VD, EB, CC, UC, OsCld, MISC,"
                    " AUCTION,TOTAL, OS, POS, vR, CASH, Remittance) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,"
                    "?, ?, ?, ?, ?, ?, ?)",
                    (Date, LOP, FOP, LLT, FLT, LL, WC, KFC, DFC, GST, DC, VD, EB, CC, UC, OsCld, MISC, AUCTION, TOTAL, OS,
                     POS,
                     vR,
                     CASH, Remittance))
                conn.commit()

                print('\n')
                again_op = input('Add more(any key) or Stop(s) or Edit any entry(e)...........: ')
                if str(again_op) == 's':
                    break

                # #...............Edit function........................................................
                elif str(again_op) == 'e':
                    print('\n')
                    print('............Options.............')
                    print('LOP, FOP, LLT, FLT, LL, WC, KFC, DFC, GST, DC, VD, EB, CC, UC, OsCld, MISC, AUCTION,'
                          'OS, POS, vR, exit')
                    print('\n')

                    while True:
                        option_list = ['LOP', 'FOP', 'LLT', 'FLT', 'LL', 'WC', 'KFC', 'DFC', 'GST',
                                       'DC', 'VD', 'EB', 'CC', 'UC', 'OsCld', 'MISC', 'AUCTION', 'OS', 'POS', 'vR',
                                       'exit']
                        edit_in = input('Select entry to edit :')

                        if edit_in not in option_list:
                            print('\n')
                            print('.........Select below options...........')
                            print('LOP, FOP, LLT, FLT, LL, WC, KFC, DFC, GST, DC, VD, EB, CC, UC, OsCld, MISC, AUCTION, '
                                  'OS, POS, vR, exit')
                            print('\n')
                            continue

                        elif edit_in == 'exit':
                            print('\n')
                            break

                        else:
                            new_edit = input(str(edit_in) + '......: ')
                            if new_edit.isdigit() == False:
                                new_edit = 0

                            #>>>>>>>> update >>>>>>>>>>>>>>>
                            c.execute("UPDATE cashbook SET {} = '{}' WHERE Date = '{}'".format(edit_in, new_edit, Date))
                            conn.commit()

                            # >>>>>>> total >>>>>>>>>>>>>>>
                            c.execute("SELECT LOP, FOP, LLT, FLT, LL, WC, KFC, DFC, GST, DC, VD, EB, CC, UC, OsCld, MISC,"
                                      " AUCTION FROM cashbook WHERE {} = '{}' ".format('Date', Date))
                            n_update = c.fetchall()

                            n_total = 0
                            for i in n_update:
                                for j in i:
                                    n_total = n_total + int(j)

                            c.execute("UPDATE cashbook SET Total = '{}' WHERE Date = '{}'".format(n_total, Date))
                            conn.commit()

                            # >>>>>>> Cash >>>>>>>>>>>>>>>
                            c.execute("SELECT OS, POS, vR FROM cashbook WHERE {} = '{}' ".format('Date', Date))
                            n_update = c.fetchall()

                            sum_new = 0
                            for i in n_update:
                                for j in i:
                                    sum_new = sum_new + int(j)
                            n_cash = n_total - sum_new
                            c.execute("UPDATE cashbook SET CASH = '{}' WHERE Date = '{}'".format(n_cash, Date))
                            conn.commit()

                            # >>>>>>> remi >>>>>>>>>>>>>>>
                            c.execute("SELECT CASH, POS, vR FROM cashbook WHERE {} = '{}' ".format('Date', Date))
                            n_update = c.fetchall()

                            n_remi = 0
                            for i in n_update:
                                for j in i:
                                    n_remi = n_remi + int(j)
                            c.execute("UPDATE cashbook SET Remittance = '{}' WHERE Date = '{}'".format(n_remi, Date))
                            conn.commit()

                            print('\n')
                            print('..........Entry updated...........')
                            print('\n')
                            continue

                    break
                # #...............Edit function ends........................................................

                else:
                    continue

        elif str(add_view) == 'v':
            clear()
            print('............Cashbook Table.................')
            print('\n')
            pd.set_option('display.max_columns', 500)
            df = pd.read_sql_query(
                "SELECT * FROM CashBook  ORDER BY Date ASC", conn)
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

                    c.execute("DELETE FROM cashbook WHERE {} = '{}'".format('Date', dlt_date))
                    conn.commit()

                    print('\n')
                    print('...........Entry deleted.........')
                    print('\n')

                    pd.set_option('display.max_columns', 500)
                    df = pd.read_sql_query(
                        "SELECT * FROM cashbook  ORDER BY Date ASC", conn)
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
        p_or_exit = input('Print the table(p) or Go back to the main menu(e) : ')
        print('\n')
        if str(p_or_exit) == 'p':
            cashbook_xlsx_writer()
            break
        elif str(p_or_exit) == 'e':
            break
        else:
            continue


if __name__ == "__main__":
    cashbook_table_create()










