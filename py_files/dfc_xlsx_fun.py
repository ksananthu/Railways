import xlsxwriter
import sqlite3
import json
import os
import webbrowser
from num2words import num2words


dire = os.getcwd()


def dfc_xlsx_writer():
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
    workbook = xlsxwriter.Workbook(str(dire) + '\\xlsx\\' + y + '\\dfc.xlsx')
    worksheet2 = workbook.add_worksheet('dfc')

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
    caption = 'DFC STATEMENT OF CLT PARCELS FOR THE MONTH OF ' + cap_m_y

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
    worksheet2.add_table('C5:D9', {'style': 'Table Style Light 15',
                                   'columns': [{'header': 'Item'}, {'header': 'Amount'}]})

    # ...............lop....................................
    c.execute("select dfc from lop ORDER BY Date ASC")
    amount1 = c.fetchall()
    total_amt1 = 0
    # print(amount)
    for j in amount1:
        j = str(j).replace('(', '').replace(',)', '')
        total_amt1 = int(j) + total_amt1
    # print(total_amt)
    worksheet2.write('C6', 'LOP')
    worksheet2.write('D6', total_amt1)

    # ...............fop....................................
    c.execute("select dfc from fop ORDER BY Date ASC")
    amount2 = c.fetchall()
    total_amt2 = 0
    # print(amount)
    for j in amount2:
        j = str(j).replace('(', '').replace(',)', '')
        total_amt2 = int(j) + total_amt2
    # print(total_amt)
    worksheet2.write('C7', 'FOP')
    worksheet2.write('D7', total_amt2)

    # ...............llt....................................
    c.execute("select dfc from llt ORDER BY Date ASC")
    amount3 = c.fetchall()
    total_amt3 = 0
    # print(amount)
    for j in amount3:
        j = str(j).replace('(', '').replace(',)', '')
        total_amt3 = int(j) + total_amt3
    # print(total_amt)
    worksheet2.write('C8', 'LLT')
    worksheet2.write('D8', total_amt3)

    # ...............flt....................................
    c.execute("select dfc from flt ORDER BY Date ASC")
    amount4 = c.fetchall()
    total_amt4 = 0
    # print(amount)
    for j in amount4:
        j = str(j).replace('(', '').replace(',)', '')
        total_amt4 = int(j) + total_amt4
    # print(total_amt)
    worksheet2.write('C9', 'FLT')
    worksheet2.write('D9', total_amt4)

    # ..................total...........................
    cell_format = workbook.add_format({'bold': True, 'font_color': 'black', 'font_size': '11', 'align': 'center'})
    worksheet2.write('C10', 'Total', cell_format)
    worksheet2.write_formula('D10', '{=SUM(D6:D9)}', cell_format)

    # # ...........................writing total in words....................................................
    total_amt = total_amt1 + total_amt2 + total_amt3 + total_amt4
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
    webbrowser.open(str(dire) + '\\xlsx\\' + y + '\\dfc.xlsx')
    print('\n')


if __name__ == "__main__":
    dfc_xlsx_writer()

# dfc_xlsx_writer()


