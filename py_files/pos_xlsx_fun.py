import xlsxwriter
import sqlite3
import json
import os
import webbrowser
from num2words import num2words


dire = os.getcwd()


def pos_xlsx_writer():
    # # ...................loading default month and year........................................................
    with open(str(dire) + '\\py_files\\m_y_choice.json') as json_file:
        data = json.load(json_file)

    y = str(data).replace("['", '').replace("']", '')

    # # ..................creating directory if not exists.......................................................
    if not os.path.exists(str(dire) + '\\xlsx\\' + y):
        os.makedirs(str(dire) + '\\xlsx\\' + y)

    # # .....................creating xlsx file................................................................
    workbook = xlsxwriter.Workbook(str(dire) + '\\xlsx\\' + y + '\\pos.xlsx')
    worksheet2 = workbook.add_worksheet()

    # # ........................fetching table from database and writing.......................................
    conn = sqlite3.connect(str(dire) + '\\Database\\' + y + '\\' + y + '.db')
    c = conn.cursor()

    c.execute("select * from Pos ORDER BY Date ASC")

    data = c.fetchall()

    # # Caption writing
    cap_m_y = str(y).replace('_', ' ').upper()
    caption = 'SUMMARY OF POS TRANSACTIONS FOR ' + cap_m_y + ' CALICUT PARCELS'

    # default cell format to size 13
    # workbook.formats[0].set_font_size(10)

    merge_format = workbook.add_format({
        'bold': 4,
        'border': 1,
        'align': 'center',
        'font_size': 15,
        'valign': 'vcenter'})

    worksheet2.set_row(0, 20)
    worksheet2.merge_range('A1:D2', caption, merge_format)

    # Set the columns widths.
    worksheet2.set_column('A:C', 15.5)
    worksheet2.set_column('D:D', 45)

    # Add a table to the worksheet.
    worksheet2.add_table('A3:D45', {'data': data,
                                    'style': 'Table Style Medium 11',
                                    'columns': [{'header': 'Date'},
                                                {'header': 'Amount'},
                                                {'header': 'ID'},
                                                {'header': 'Cmt'}]})

    # # ...........................writing total in words....................................................
    c.execute("select amt from Pos ORDER BY Date ASC")
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
    cell_format = workbook.add_format({'bold': True, 'font_color': 'black', 'font_size': '13', 'align': 'center'})
    worksheet2.write('A45', 'Total', cell_format)
    worksheet2.write_formula('B45', '{=SUM(B4:B44)}', cell_format)
    worksheet2.write('A47', in_words)
    worksheet2.write('A48', 'CLT PO', cell_format)
    worksheet2.write('A49', month, cell_format)
    worksheet2.write('D49', 'CPS/CLT', cell_format)
    workbook.close()

    # # opening the file
    webbrowser.open(str(dire) + '\\xlsx\\' + y + '\\pos.xlsx')
    print('\n')


