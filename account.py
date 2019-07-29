import sqlite3
import json
import os
from py_files.database_creator_fun import data_add
from py_files.pos_table_fun import pos_table_create
from py_files.cashbook_table_fun import cashbook_table_create
from py_files.wc_table_fun import wc_table_create
from py_files.lop_table_fun import lop_table_create
from py_files.fop_table_fun import fop_table_create
from py_files.llt_table_fun import llt_table_create
from py_files.flt_table_fun import flt_table_create
from py_files.other_fun import other_create

import pyfiglet


def clear():
    # for windows
    if os.name == 'nt':
        _ = os.system('cls')
    else:
        _ = os.system('clear')


dire = os.getcwd()


# #..............................anscii text art..............................................................
print('###########################################################################################################')
print(pyfiglet.figlet_format('..Accounts..', font='larry3d'))
print(pyfiglet.figlet_format('...................................................................by ks', font='pepper'))
print('###########################################################################################################')
print('\n')


# # ...................opening default month and year..........................................................
with open(str(dire) + '\\py_files\\m_y_choice.json') as json_file:
    data = json.load(json_file)

x = str(data).replace("['", '').replace("']", '')
print('###...............' + str(x) + ' is selected..............###')
print('\n')

# # ........................Option continue or edit...............................................................
while True:
    to_change = input('To continue(c) and to edit the month and year (e) :  ')
    if str(to_change) == 'e':
        data_add()
        break
    elif str(to_change) == 'c':
        break
    else:
        continue

clear()

# #......................Table selection.........................................................................
with open(str(dire) + '\\py_files\\m_y_choice.json') as json_file:
    data1 = json.load(json_file)

y = str(data1).replace("['", '').replace("']", '')
print('###...............' + str(y) + ' is selected..............###')

print('\n')

while True:
    print('...........................Options..............................')
    print(' cashbook, lop, fop, wc, llt, flt, other(DFC, LL, KFC and GST)')
    print('\n')
    select_table = input('Select table or statement ...............: ')

    if select_table == 'pos':
        clear()
        print('.................................POS is selected....................................')
        print('\n')
        pos_table_create()

    elif select_table == 'cashbook':
        clear()
        print('.................................Cashbook is selected....................................')
        print('\n')
        cashbook_table_create()

    elif select_table == 'wc':
        clear()
        print('.................................WC is selected....................................')
        print('\n')
        wc_table_create()

    elif select_table == 'lop':
        clear()
        print('.................................LOP is selected....................................')
        print('\n')
        lop_table_create()

    elif select_table == 'fop':
        clear()
        print('.................................FOP is selected....................................')
        print('\n')
        fop_table_create()

    elif select_table == 'llt':
        clear()
        print('.................................LLT is selected....................................')
        print('\n')
        llt_table_create()

    elif select_table == 'flt':
        clear()
        print('.................................FLT is selected....................................')
        print('\n')
        flt_table_create()

    elif select_table == 'other':
        clear()
        print('.................................other is selected....................................')
        print('\n')
        other_create()






