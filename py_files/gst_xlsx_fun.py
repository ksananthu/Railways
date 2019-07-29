import xlsxwriter
import sqlite3
import json
import os
import webbrowser
from num2words import num2words


dire = os.getcwd()


def clear():
    # for windows
    if os.name == 'nt':
        _ = os.system('cls')
    else:
        _ = os.system('clear')


def change_gst():
    while True:
        new_cgst = input('Input CGST rate(%)  :  ')
        try:
            new_cgst = float(new_cgst)
            break
        except ValueError:
            continue

    new_cgst_st = [str(new_cgst)]
    # print(new_cgst_st)

    with open(str(dire) + '\\py_files\\cgst_rate.json', 'w') as outfile:
        json.dump(new_cgst_st, outfile)



def gst_xlsx_writer():
    # # ...................loading default month and year........................................................
    with open(str(dire) + '\\py_files\\m_y_choice.json') as json_file:
        data = json.load(json_file)

    y = str(data).replace("['", '').replace("']", '')

    # # ...........selecting default year........................
    with open(str(dire) + '\\py_files\\y_choice.json') as json_file2:
        data_y = json.load(json_file2)
    year = str(data_y).replace("['", '').replace("']", '')

    # # ..................creating directory if not exists.......................................................
    if not os.path.exists(str(dire) + '\\xlsx\\' + y):
        os.makedirs(str(dire) + '\\xlsx\\' + y)

    # # ...................opening default cgst_rate..........................................................
    with open(str(dire) + '\\py_files\\cgst_rate.json') as json_file:
        cgst_rate = json.load(json_file)

    x = str(cgst_rate).replace("['", '').replace("']", '')

    clear()
    print('...............CGST rate selected is ' + str(x) + '%..............')
    print('\n')

    # # ........................Option continue or edit...............................................................
    while True:
        to_change = input('Continue(c) or Edit gst rate(e) :  ')
        if str(to_change) == 'e':
            change_gst()
            continue
        elif str(to_change) == 'c':
            break
        else:
            continue

    clear()

    with open(str(dire) + '\\py_files\\cgst_rate.json') as json_file:
        cgst_rate_load = json.load(json_file)

    x = str(cgst_rate_load).replace("['", '').replace("']", '')
    print('...............CGST rate selected is ' + str(x) + '%..............')
    print('\n')
    cgst_rate = float(x)
    sgst_rate = round(100 - cgst_rate, 1)
    print('CGST rate is ........: ' + str(cgst_rate))
    print('SGST rate is ........: ' + str(sgst_rate))

    # # .....................creating xlsx file................................................................
    workbook = xlsxwriter.Workbook(str(dire) + '\\xlsx\\' + y + '\\gst.xlsx')
    worksheet2 = workbook.add_worksheet('gst')

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

    # # ........................fetching table from database and writing first.......................................
    conn = sqlite3.connect(str(dire) + '\\Database\\' + y + '\\' + y + '.db')
    c = conn.cursor()

    # c.execute("select date, ll from cashbook ORDER BY Date ASC")
    # data1 = c.fetchall()

    # # Caption writing
    cap_m_y = str(y).replace('_', ' ').upper()
    caption = 'GST STATEMENT OF CLT PARCELS FOR THE MONTH OF ' + cap_m_y

    merge_format = workbook.add_format({
        'bold': 4,
        'border': 1,
        'align': 'center',
        'font_size': 14,
        'valign': 'vcenter'})

    # default cell format to size 11
    workbook.formats[0].set_font_size(11)

    worksheet2.set_column('A:A', 14)
    worksheet2.set_column('B:G', 14)
    worksheet2.set_row(0, 40)
    worksheet2.set_default_row(16)

    worksheet2.merge_range('A1:F1', caption, merge_format)

    # Add a table to the worksheet.
    worksheet2.add_table('B5:E12', {'style': 'Table Style Light 15',
                                   'columns': [{'header': 'Item'}, {'header': 'CGST'},
                                               {'header': 'SGST'}, {'header': 'Amount'}
                                               ]})

    # ...............lop....................................
    c.execute("select st from lop ORDER BY Date ASC")
    amount1 = c.fetchall()
    total_amt1 = 0
    # print(amount)
    for j in amount1:
        j = str(j).replace('(', '').replace(',)', '')
        total_amt1 = int(j) + total_amt1
    # print(total_amt)
    worksheet2.write('B6', 'LOP')
    worksheet2.write('C6', (total_amt1 * (cgst_rate)/100))
    worksheet2.write('D6', (total_amt1 * (sgst_rate) / 100))
    worksheet2.write('E6', total_amt1)

    # ...............fop....................................
    c.execute("select st from fop ORDER BY Date ASC")
    amount2 = c.fetchall()
    total_amt2 = 0
    # print(amount)
    for j in amount2:
        j = str(j).replace('(', '').replace(',)', '')
        total_amt2 = int(j) + total_amt2
    # print(total_amt)
    worksheet2.write('B7', 'FOP')
    worksheet2.write('C7', (total_amt2 * (cgst_rate) / 100))
    worksheet2.write('D7', (total_amt2 * (sgst_rate) / 100))
    worksheet2.write('E7', total_amt2)

    # ...............llt....................................
    c.execute("select st from llt ORDER BY Date ASC")
    amount3 = c.fetchall()
    total_amt3 = 0
    # print(amount)
    for j in amount3:
        j = str(j).replace('(', '').replace(',)', '')
        total_amt3 = int(j) + total_amt3
    # print(total_amt)
    worksheet2.write('B8', 'LLT')
    worksheet2.write('C8', (total_amt3 * (cgst_rate) / 100))
    worksheet2.write('D8', (total_amt3 * (sgst_rate) / 100))
    worksheet2.write('E8', total_amt3)

    # ...............flt....................................
    c.execute("select st from flt ORDER BY Date ASC")
    amount4 = c.fetchall()
    total_amt4 = 0
    # print(amount)
    for j in amount4:
        j = str(j).replace('(', '').replace(',)', '')
        total_amt4 = int(j) + total_amt4
    # print(total_amt)
    worksheet2.write('B9', 'FLT')
    worksheet2.write('C9', (total_amt4 * (cgst_rate) / 100))
    worksheet2.write('D9', (total_amt4 * (sgst_rate) / 100))
    worksheet2.write('E9', total_amt4)

    # ...............WC....................................
    c.execute("select st from wc ORDER BY Date ASC")
    amount5 = c.fetchall()
    total_amt5 = 0
    # print(amount)
    for j in amount5:
        j = str(j).replace('(', '').replace(',)', '')
        total_amt5 = int(j) + total_amt5
    # print(total_amt)
    worksheet2.write('B10', 'WC')
    worksheet2.write('C10', (total_amt5 * (cgst_rate) / 100))
    worksheet2.write('D10', (total_amt5 * (sgst_rate) / 100))
    worksheet2.write('E10', total_amt5)

    # ...............DC....................................
    print('\n')
    print('.............Warning............')
    print('If there is no entry add zero')
    print('\n')
    while True:
        dc = input('DC...........: ')
        dc_op = input('Edit(e) or Continue(c)...: ')
        if dc_op == 'e':
            continue
        elif dc_op == 'c':
            try:
                dc = int(dc)
                break
            except ValueError:
                continue
        else:
            continue
    # print(dc)
    worksheet2.write('B11', 'DC')
    worksheet2.write('C11', (dc * (cgst_rate) / 100))
    worksheet2.write('D11', (dc * (sgst_rate) / 100))
    worksheet2.write('E11', dc)

    # ...............AUCTION....................................
    print('\n')
    print('.............Warning............')
    print('If there is no entry add zero')
    print('\n')
    while True:
        auc = input('AUCTION...........: ')
        auc_op = input('Edit(e) or Continue(c)...: ')
        if auc_op == 'e':
            continue
        elif auc_op == 'c':
            try:
                auc = int(auc)
                break
            except ValueError:
                continue

        else:
            continue
    # print(dc)
    worksheet2.write('B12', 'AUCTION')
    worksheet2.write('C12', (auc * (cgst_rate) / 100))
    worksheet2.write('D12', (auc * (sgst_rate) / 100))
    worksheet2.write('E12', auc)

    # ..................total...........................
    cell_format = workbook.add_format({'bold': True, 'font_color': 'black', 'font_size': '11', 'align': 'center'})
    worksheet2.write('B13', 'Total', cell_format)
    worksheet2.write_formula('C13', '{=SUM(C6:C12)}', cell_format)
    worksheet2.write_formula('D13', '{=SUM(D6:D12)}', cell_format)
    worksheet2.write_formula('E13', '{=SUM(E6:E12)}', cell_format)

    # # ...........................writing total in words....................................................
    total_amt = total_amt1 + total_amt2 + total_amt3 + total_amt4 + total_amt5 + dc + auc
    p = num2words(total_amt, lang='en_IN')
    in_words1 = str(p).capitalize()
    in_words = 'Rupees. ' + in_words1 + ' only.'
    # print(in_words)

    # #........................Adding date of submission........................................................
    # # ---------------------------month slice begin ------------------------
    month1 = 0
    if y.find('Jan') != -1:
        month1 = '01-02-2019'

    elif y.find('Feb') != -1:
        month1 = '01-03-2019'

    elif y.find('Mar') != -1:
        month1 = '01-04-2019'

    elif y.find('April') != -1:
        month1 = '01-05-2019'

    elif y.find('May') != -1:
        month1 = '01-06-2019'

    elif y.find('June') != -1:
        month1 = '01-07-2019'

    elif y.find('July') != -1:
        month1 = '01-08-2019'

    elif y.find('August') != -1:
        month1 = '01-09-2019'

    elif y.find('Sep') != -1:
        month1 = '01-10-2019'

    elif y.find('Oct') != -1:
        month1 = '01-11-2019'

    elif y.find('Nov') != -1:
        month1 = '01-12-2019'

    elif y.find('Dec') != -1:
        month1 = '01-01-2020'
    # # ---------------------------month slice end ------------------------
    # # .............................Writing the last lines......................................................
    # cell_format = workbook.add_format({'bold': True, 'font_color': 'black', 'font_size': '13', 'align': 'center'})

    worksheet2.write('A15', in_words)
    worksheet2.write('A18', 'CLT PO', cell_format)
    worksheet2.write('A19', month1, cell_format)
    worksheet2.write('F18', 'CPS/CLT', cell_format)

    # .....................................................................
    workbook.close()

    # # opening the file
    webbrowser.open(str(dire) + '\\xlsx\\' + y + '\\gst.xlsx')
    print('\n')




if __name__ == "__main__":
    gst_xlsx_writer()


# gst_xlsx_writer()