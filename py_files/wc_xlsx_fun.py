import xlsxwriter
import sqlite3
import json
import os
import webbrowser
from num2words import num2words


dire = os.getcwd()


def wc_xlsx_writer():
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

    # # .....................creating xlsx file................................................................
    workbook = xlsxwriter.Workbook(str(dire) + '\\xlsx\\' + y + '\\wc.xlsx')
    worksheet2 = workbook.add_worksheet('WC')

    # # ........................fetching table from database and writing first.......................................
    conn = sqlite3.connect(str(dire) + '\\Database\\' + y + '\\' + y + '.db')
    c = conn.cursor()

    c.execute("select * from wc ORDER BY Date ASC")
    data1 = c.fetchall()

    # # Caption writing
    cap_m_y = str(y).replace('_', ' ').upper()
    caption = 'Wharfage return of CLT parcel for the month of  ' + cap_m_y

    merge_format = workbook.add_format({
        'bold': 4,
        'border': 1,
        'align': 'center',
        'font_size': 15,
        'valign': 'vcenter'})

    # default cell format to size 11
    workbook.formats[0].set_font_size(11)

    worksheet2.set_column('A:E', 19)
    worksheet2.set_row(0, 35)
    worksheet2.set_default_row(17)

    worksheet2.merge_range('A1:D1', caption, merge_format)

    # Add a table to the worksheet.
    worksheet2.add_table('A2:D33', {'data': data1,
                                    'style': 'Table Style Light 15',
                                    'columns': [{'header': 'Date'},
                                                {'header': 'Wharfage'},
                                                {'header': 'ST'},
                                                {'header': 'TOTAL'}
                                                ]})

    cell_format = workbook.add_format({'bold': True, 'font_color': 'black', 'font_size': '12', 'align': 'center'})
    worksheet2.write('A34', 'Total', cell_format)
    worksheet2.write_formula('B34', '{=SUM(B3:B33)}', cell_format)
    worksheet2.write_formula('C34', '{=SUM(C3:C33)}', cell_format)
    worksheet2.write_formula('D34', '{=SUM(D3:D33)}', cell_format)

    # # ...........................writing total in words....................................................
    c.execute("select TOTAL from wc ORDER BY Date ASC")
    amount = c.fetchall()
    total_amt = 0
    # print(amount)
    for j in amount:
        j = str(j).replace('(', '').replace(',)', '')
        total_amt = int(j) + total_amt
    # print(total_amt)
    p = num2words(total_amt, lang='en_IN')
    in_words1 = str(p).capitalize()
    in_words = 'Rupees. ' + in_words1 + ' only.'
    # print(in_words)

    # #........................Adding date of submission........................................................
    # # ---------------------------month slice begin ------------------------
    month = 0
    if y.find('Jan') != -1:
        month = '01-02-2019'

    elif y.find('Feb') != -1:
        month = '01-03-2019'

    elif y.find('Mar') != -1:
        month = '01-04-2019'

    elif y.find('April') != -1:
        month = '01-05-2019'

    elif y.find('May') != -1:
        month = '01-06-2019'

    elif y.find('June') != -1:
        month = '01-07-2019'

    elif y.find('July') != -1:
        month = '01-08-2019'

    elif y.find('August') != -1:
        month = '01-09-2019'

    elif y.find('Sep') != -1:
        month = '01-10-2019'

    elif y.find('Oct') != -1:
        month = '01-11-2019'

    elif y.find('Nov') != -1:
        month = '01-12-2019'

    elif y.find('Dec') != -1:
        month = '01-01-2020'
    # # ---------------------------month slice end ------------------------

    # # .............................Writing the last lines......................................................
    # cell_format = workbook.add_format({'bold': True, 'font_color': 'black', 'font_size': '13', 'align': 'center'})

    worksheet2.write('A36', in_words)
    worksheet2.write('A38', 'CLT PO', cell_format)
    worksheet2.write('A39', month, cell_format)
    worksheet2.write('D38', 'CPS/CLT', cell_format)

    workbook.close()

    # # opening the file
    webbrowser.open(str(dire) + '\\xlsx\\' + y + '\\wc.xlsx')
    print('\n')


# wc_xlsx_writer()



