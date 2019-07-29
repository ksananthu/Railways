import xlsxwriter
import sqlite3
import json
import os
import webbrowser
from num2words import num2words


dire = os.getcwd()


def fop_xlsx_writer():
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
    workbook = xlsxwriter.Workbook(str(dire) + '\\xlsx\\' + y + '\\fop.xlsx')
    worksheet2 = workbook.add_worksheet('fop')

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

    sort_date1 = str(year) + '-' + str(month) + '-' + '01'
    sort_date2 = str(year) + '-' + str(month) + '-' + '10'

    c.execute("select * from fop WHERE Date BETWEEN '{}' AND '{}' ORDER BY Date ASC".format(sort_date1, sort_date2))
    data1 = c.fetchall()

    # # Caption writing
    cap_m_y = str(y).replace('_', ' ').upper()
    caption = 'FOP return of CLT parcel for the month of  ' + cap_m_y

    merge_format = workbook.add_format({
        'bold': 4,
        'border': 1,
        'align': 'center',
        'font_size': 15,
        'valign': 'vcenter'})

    # default cell format to size 11
    workbook.formats[0].set_font_size(10)

    worksheet2.set_column('A:G', 13)
    worksheet2.set_row(0, 30)
    worksheet2.set_default_row(16)

    worksheet2.merge_range('A1:G1', caption, merge_format)

    # Add a table to the worksheet.
    worksheet2.add_table('A2:G12', {'data': data1,
                                    'style': 'Table Style Light 15',
                                    'columns': [{'header': 'Date'},
                                                {'header': 'No of pls'},
                                                {'header': 'Wt in Qtls'},
                                                {'header': 'Freight'},
                                                {'header': 'DFC'},
                                                {'header': 'ST'},
                                                {'header': 'TOTAL'}
                                                ]})

    cell_format = workbook.add_format({'bold': True, 'font_color': 'black', 'font_size': '11', 'align': 'center'})
    worksheet2.write('A13', 'Total 1', cell_format)
    worksheet2.write_formula('B13', '{=SUM(B3:B12)}', cell_format)
    worksheet2.write_formula('C13', '{=SUM(C3:C12)}', cell_format)
    worksheet2.write_formula('D13', '{=SUM(D3:D12)}', cell_format)
    worksheet2.write_formula('E13', '{=SUM(E3:E12)}', cell_format)
    worksheet2.write_formula('F13', '{=SUM(F3:F12)}', cell_format)
    worksheet2.write_formula('G13', '{=SUM(G3:G12)}', cell_format)

    # # .............................Writing middle lines......................................................
    sort_date10 = str(year) + '-' + str(month) + '-' + '11'
    sort_date20 = str(year) + '-' + str(month) + '-' + '20'

    c.execute("select * from fop WHERE Date BETWEEN '{}' AND '{}' ORDER BY Date ASC".format(sort_date10, sort_date20))
    data2 = c.fetchall()

    # Add a table to the worksheet.
    worksheet2.add_table('A14:G23', {'data': data2, 'header_row': False, 'style': 'Table Style Light 15'})

    worksheet2.write('A24', 'Total 2', cell_format)
    worksheet2.write_formula('B24', '{=SUM(B14:B23)}', cell_format)
    worksheet2.write_formula('C24', '{=SUM(C14:C23)}', cell_format)
    worksheet2.write_formula('D24', '{=SUM(D14:D23)}', cell_format)
    worksheet2.write_formula('E24', '{=SUM(E14:E23)}', cell_format)
    worksheet2.write_formula('F24', '{=SUM(F14:F23)}', cell_format)
    worksheet2.write_formula('G24', '{=SUM(G14:G23)}', cell_format)

    # # .............................Writing last lines......................................................
    sort_date100 = str(year) + '-' + str(month) + '-' + '21'
    sort_date200 = str(year) + '-' + str(month) + '-' + '31'

    c.execute("select * from fop WHERE Date BETWEEN '{}' AND '{}' ORDER BY Date ASC".format(sort_date100, sort_date200))
    data3 = c.fetchall()

    # Add a table to the worksheet.
    worksheet2.add_table('A25:G35', {'data': data3, 'header_row': False, 'style': 'Table Style Light 15'})

    worksheet2.write('A36', 'Total 3', cell_format)
    worksheet2.write_formula('B36', '{=SUM(B25:B35)}', cell_format)
    worksheet2.write_formula('C36', '{=SUM(C25:C35)}', cell_format)
    worksheet2.write_formula('D36', '{=SUM(D25:D35)}', cell_format)
    worksheet2.write_formula('E36', '{=SUM(E25:E35)}', cell_format)
    worksheet2.write_formula('F36', '{=SUM(F25:F35)}', cell_format)
    worksheet2.write_formula('G36', '{=SUM(G25:G35)}', cell_format)

    # # .............................Writing Total ......................................................

    worksheet2.write('A37', 'Sum Total', cell_format)
    worksheet2.write_formula('B37', '{=SUM(B13+B24+B36)}', cell_format)
    worksheet2.write_formula('C37', '{=SUM(C13+C24+C36)}', cell_format)
    worksheet2.write_formula('D37', '{=SUM(D13+D24+D36)}', cell_format)
    worksheet2.write_formula('E37', '{=SUM(E13+E24+E36)}', cell_format)
    worksheet2.write_formula('F37', '{=SUM(F13+F24+F36)}', cell_format)
    worksheet2.write_formula('G37', '{=SUM(G13+G24+G36)}', cell_format)

    # # .............................Writing extra lines ......................................................
    # # ...........................writing total in words....................................................
    c.execute("select TOTAL from fop ORDER BY Date ASC")
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

    worksheet2.write('A39', in_words)
    worksheet2.write('A41', 'CLT PO', cell_format)
    worksheet2.write('A42', month1, cell_format)
    worksheet2.write('G41', 'CPS/CLT', cell_format)

    workbook.close()

    # # opening the file
    webbrowser.open(str(dire) + '\\xlsx\\' + y + '\\fop.xlsx')
    print('\n')

# fop_xlsx_writer()
