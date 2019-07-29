import xlsxwriter
import sqlite3
import json
import os
import webbrowser
from num2words import num2words


dire = os.getcwd()


def flt_xlsx_writer():
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
    workbook = xlsxwriter.Workbook(str(dire) + '\\xlsx\\' + y + '\\flt.xlsx')
    worksheet2 = workbook.add_worksheet('flt')

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

    c.execute("select * from flt ORDER BY Date ASC")
    data1 = c.fetchall()

    # # Caption writing
    cap_m_y = str(y).replace('_', ' ').upper()
    caption = 'FOREIGN LUGGAGE RETURN OF CALICUT PARCELS  FOR THE MONTH OF ' + cap_m_y

    merge_format = workbook.add_format({
        'bold': 4,
        'border': 1,
        'align': 'center',
        'font_size': 15,
        'valign': 'vcenter'})

    # default cell format to size 11
    workbook.formats[0].set_font_size(10)

    worksheet2.set_column('A:G', 14.6)
    worksheet2.set_row(0, 30)
    worksheet2.set_default_row(16)

    worksheet2.merge_range('A1:G1', caption, merge_format)

    # Add a table to the worksheet.
    worksheet2.add_table('A2:G33', {'data': data1,
                                    'style': 'Table Style Light 15',
                                    'columns': [{'header': 'Date'},
                                                {'header': 'No of Lug'},
                                                {'header': 'Wt in Qtls'},
                                                {'header': 'Freight'},
                                                {'header': 'DFC'},
                                                {'header': 'ST'},
                                                {'header': 'TOTAL'}
                                                ]})

    cell_format = workbook.add_format({'bold': True, 'font_color': 'black', 'font_size': '11', 'align': 'center'})
    worksheet2.write('A34', 'Total', cell_format)
    worksheet2.write_formula('B34', '{=SUM(B3:B33)}', cell_format)
    worksheet2.write_formula('C34', '{=SUM(C3:C33)}', cell_format)
    worksheet2.write_formula('D34', '{=SUM(D3:D33)}', cell_format)
    worksheet2.write_formula('E34', '{=SUM(E3:E33)}', cell_format)
    worksheet2.write_formula('F34', '{=SUM(F3:F33)}', cell_format)
    worksheet2.write_formula('G34', '{=SUM(G3:G33)}', cell_format)

    # # .............................Writing extra lines ......................................................
    # # ...........................writing total in words....................................................
    c.execute("select TOTAL from flt ORDER BY Date ASC")
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

    worksheet2.write('A36', in_words)
    worksheet2.write('A39', 'CLT PO', cell_format)
    worksheet2.write('A40', month1, cell_format)
    worksheet2.write('G39', 'CPS/CLT', cell_format)

    # .....................................................................
    workbook.close()

    # # opening the file
    webbrowser.open(str(dire) + '\\xlsx\\' + y + '\\flt.xlsx')
    print('\n')




# flt_xlsx_writer()




