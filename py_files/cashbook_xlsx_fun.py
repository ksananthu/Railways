import xlsxwriter
import sqlite3
import json
import os
import webbrowser


dire = os.getcwd()


def cashbook_xlsx_writer():
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
    workbook = xlsxwriter.Workbook(str(dire) + '\\xlsx\\' + y + '\\cashbook.xlsx')
    worksheet2 = workbook.add_worksheet()

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

    c.execute(
        "select * from CashBook WHERE Date BETWEEN '{}' AND '{}' ORDER BY Date ASC".format(sort_date1, sort_date2))
    data1 = c.fetchall()

    # # Caption writing
    cap_m_y = str(y).replace('_', ' ').upper()
    caption = 'CASH BOOK SUMMARY FOR  ' + cap_m_y + ' CALICUT PARCELS'

    merge_format = workbook.add_format({
        'bold': 4,
        'border': 1,
        'align': 'center',
        'font_size': 15,
        'valign': 'vcenter'})

    # default cell format to size 9
    workbook.formats[0].set_font_size(9)

    worksheet2.set_column('A:A', 11.5)
    worksheet2.set_row(0, 17)
    worksheet2.set_default_row(12.2)

    worksheet2.merge_range('A1:R1', caption, merge_format)

    # Add a table to the worksheet.
    worksheet2.add_table('A2:X12', {'data': data1,
                                    'style': 'Table Style Light 15',
                                    'columns': [{'header': 'Date'},
                                                {'header': 'LOP'},
                                                {'header': 'FOP'},
                                                {'header': 'LLT'},
                                                {'header': 'FLT'},
                                                {'header': 'LL'},
                                                {'header': 'WC'},
                                                {'header': 'KFC'},
                                                {'header': 'DFC'},
                                                {'header': 'GST'},
                                                {'header': 'DC'},
                                                {'header': 'VD'},
                                                {'header': 'EB'},
                                                {'header': 'CC'},
                                                {'header': 'UC'},
                                                {'header': 'OsCld'},
                                                {'header': 'MISC'},
                                                {'header': 'Auction'},
                                                {'header': 'Total'},
                                                {'header': 'OS'},
                                                {'header': 'POS'},
                                                {'header': 'vR'},
                                                {'header': 'Cash'},
                                                {'header': 'Remittance'},
                                                ]})

    cell_format = workbook.add_format({'bold': True, 'font_color': 'black', 'font_size': '10', 'align': 'center'})
    worksheet2.write('A13', 'Total 1', cell_format)
    worksheet2.write_formula('B13', '{=SUM(B3:B12)}', cell_format)
    worksheet2.write_formula('C13', '{=SUM(C3:C12)}', cell_format)
    worksheet2.write_formula('D13', '{=SUM(D3:D12)}', cell_format)
    worksheet2.write_formula('E13', '{=SUM(E3:E12)}', cell_format)
    worksheet2.write_formula('F13', '{=SUM(F3:F12)}', cell_format)
    worksheet2.write_formula('G13', '{=SUM(G3:G12)}', cell_format)
    worksheet2.write_formula('H13', '{=SUM(H3:H12)}', cell_format)
    worksheet2.write_formula('I13', '{=SUM(I3:I12)}', cell_format)
    worksheet2.write_formula('J13', '{=SUM(J3:J12)}', cell_format)
    worksheet2.write_formula('K13', '{=SUM(K3:K12)}', cell_format)
    worksheet2.write_formula('L13', '{=SUM(L3:L12)}', cell_format)
    worksheet2.write_formula('M13', '{=SUM(M3:M12)}', cell_format)
    worksheet2.write_formula('N13', '{=SUM(N3:N12)}', cell_format)
    worksheet2.write_formula('O13', '{=SUM(O3:O12)}', cell_format)
    worksheet2.write_formula('P13', '{=SUM(P3:P12)}', cell_format)
    worksheet2.write_formula('Q13', '{=SUM(Q3:Q12)}', cell_format)
    worksheet2.write_formula('R13', '{=SUM(R3:R12)}', cell_format)
    worksheet2.write_formula('S13', '{=SUM(S3:S12)}', cell_format)
    worksheet2.write_formula('T13', '{=SUM(T3:T12)}', cell_format)
    worksheet2.write_formula('U13', '{=SUM(U3:U12)}', cell_format)
    worksheet2.write_formula('V13', '{=SUM(V3:V12)}', cell_format)
    worksheet2.write_formula('W13', '{=SUM(W3:W12)}', cell_format)
    worksheet2.write_formula('X13', '{=SUM(X3:X12)}', cell_format)

    # # .............................Writing middle lines......................................................
    sort_date10 = str(year) + '-' + str(month) + '-' + '11'
    sort_date20 = str(year) + '-' + str(month) + '-' + '20'

    c.execute(
        "select * from CashBook WHERE Date BETWEEN '{}' AND '{}' ORDER BY Date ASC".format(sort_date10, sort_date20))
    data2 = c.fetchall()

    # Add a table to the worksheet.
    worksheet2.add_table('A14:X23', {'data': data2, 'header_row': False, 'style': 'Table Style Light 15'})

    worksheet2.write('A24', 'Total 2', cell_format)
    worksheet2.write_formula('B24', '{=SUM(B14:B23)}', cell_format)
    worksheet2.write_formula('C24', '{=SUM(C14:C23)}', cell_format)
    worksheet2.write_formula('D24', '{=SUM(D14:D23)}', cell_format)
    worksheet2.write_formula('E24', '{=SUM(E14:E23)}', cell_format)
    worksheet2.write_formula('F24', '{=SUM(F14:F23)}', cell_format)
    worksheet2.write_formula('G24', '{=SUM(G14:G23)}', cell_format)
    worksheet2.write_formula('H24', '{=SUM(H14:H23)}', cell_format)
    worksheet2.write_formula('I24', '{=SUM(I14:I23)}', cell_format)
    worksheet2.write_formula('J24', '{=SUM(J14:J23)}', cell_format)
    worksheet2.write_formula('K24', '{=SUM(K14:K23)}', cell_format)
    worksheet2.write_formula('L24', '{=SUM(L14:L23)}', cell_format)
    worksheet2.write_formula('M24', '{=SUM(M14:M23)}', cell_format)
    worksheet2.write_formula('N24', '{=SUM(N14:N23)}', cell_format)
    worksheet2.write_formula('O24', '{=SUM(O14:O23)}', cell_format)
    worksheet2.write_formula('P24', '{=SUM(P14:P23)}', cell_format)
    worksheet2.write_formula('Q24', '{=SUM(Q14:Q23)}', cell_format)
    worksheet2.write_formula('R24', '{=SUM(R14:R23)}', cell_format)
    worksheet2.write_formula('S24', '{=SUM(S14:S23)}', cell_format)
    worksheet2.write_formula('T24', '{=SUM(T14:T23)}', cell_format)
    worksheet2.write_formula('U24', '{=SUM(U14:U23)}', cell_format)
    worksheet2.write_formula('V24', '{=SUM(V14:V23)}', cell_format)
    worksheet2.write_formula('W24', '{=SUM(W14:W23)}', cell_format)
    worksheet2.write_formula('X24', '{=SUM(X14:X23)}', cell_format)

    # # .............................Writing last lines......................................................
    sort_date100 = str(year) + '-' + str(month) + '-' + '21'
    sort_date200 = str(year) + '-' + str(month) + '-' + '31'

    c.execute(
        "select * from CashBook WHERE Date BETWEEN '{}' AND '{}' ORDER BY Date ASC".format(sort_date100, sort_date200))
    data3 = c.fetchall()

    # Add a table to the worksheet.
    worksheet2.add_table('A25:X35', {'data': data3, 'header_row': False, 'style': 'Table Style Light 15'})

    worksheet2.write('A36', 'Total 3', cell_format)
    worksheet2.write_formula('B36', '{=SUM(B25:B35)}', cell_format)
    worksheet2.write_formula('C36', '{=SUM(C25:C35)}', cell_format)
    worksheet2.write_formula('D36', '{=SUM(D25:D35)}', cell_format)
    worksheet2.write_formula('E36', '{=SUM(E25:E35)}', cell_format)
    worksheet2.write_formula('F36', '{=SUM(F25:F35)}', cell_format)
    worksheet2.write_formula('G36', '{=SUM(G25:G35)}', cell_format)
    worksheet2.write_formula('H36', '{=SUM(H25:H35)}', cell_format)
    worksheet2.write_formula('I36', '{=SUM(I25:I35)}', cell_format)
    worksheet2.write_formula('J36', '{=SUM(J25:J35)}', cell_format)
    worksheet2.write_formula('K36', '{=SUM(K25:K35)}', cell_format)
    worksheet2.write_formula('L36', '{=SUM(L25:L35)}', cell_format)
    worksheet2.write_formula('M36', '{=SUM(M25:M35)}', cell_format)
    worksheet2.write_formula('N36', '{=SUM(N25:N35)}', cell_format)
    worksheet2.write_formula('O36', '{=SUM(O25:O35)}', cell_format)
    worksheet2.write_formula('P36', '{=SUM(P25:P35)}', cell_format)
    worksheet2.write_formula('Q36', '{=SUM(Q25:Q35)}', cell_format)
    worksheet2.write_formula('R36', '{=SUM(R25:R35)}', cell_format)
    worksheet2.write_formula('S36', '{=SUM(S25:S35)}', cell_format)
    worksheet2.write_formula('T36', '{=SUM(T25:T35)}', cell_format)
    worksheet2.write_formula('U36', '{=SUM(U25:U35)}', cell_format)
    worksheet2.write_formula('V36', '{=SUM(V25:V35)}', cell_format)
    worksheet2.write_formula('W36', '{=SUM(W25:W35)}', cell_format)
    worksheet2.write_formula('X36', '{=SUM(X25:X35)}', cell_format)

    # # .............................Writing Total ......................................................

    worksheet2.write('A37', 'Sum Total', cell_format)
    worksheet2.write_formula('B37', '{=SUM(B13+B24+B36)}', cell_format)
    worksheet2.write_formula('C37', '{=SUM(C13+C24+C36)}', cell_format)
    worksheet2.write_formula('D37', '{=SUM(D13+D24+D36)}', cell_format)
    worksheet2.write_formula('E37', '{=SUM(E13+E24+E36)}', cell_format)
    worksheet2.write_formula('F37', '{=SUM(F13+F24+F36)}', cell_format)
    worksheet2.write_formula('G37', '{=SUM(G13+G24+G36)}', cell_format)
    worksheet2.write_formula('H37', '{=SUM(H13+H24+H36)}', cell_format)
    worksheet2.write_formula('I37', '{=SUM(I13+I24+I36)}', cell_format)
    worksheet2.write_formula('J37', '{=SUM(J13+J24+J36)}', cell_format)
    worksheet2.write_formula('K37', '{=SUM(K13+K24+K36)}', cell_format)
    worksheet2.write_formula('L37', '{=SUM(L13+L24+L36)}', cell_format)
    worksheet2.write_formula('M37', '{=SUM(M13+M24+M36)}', cell_format)
    worksheet2.write_formula('N37', '{=SUM(N13+N24+N36)}', cell_format)
    worksheet2.write_formula('O37', '{=SUM(O13+O24+O36)}', cell_format)
    worksheet2.write_formula('P37', '{=SUM(P13+P24+P36)}', cell_format)
    worksheet2.write_formula('Q37', '{=SUM(Q13+Q24+Q36)}', cell_format)
    worksheet2.write_formula('R37', '{=SUM(R13+R24+R36)}', cell_format)
    worksheet2.write_formula('S37', '{=SUM(S13+S24+S36)}', cell_format)
    worksheet2.write_formula('T37', '{=SUM(T13+T24+T36)}', cell_format)
    worksheet2.write_formula('U37', '{=SUM(U13+U24+U36)}', cell_format)
    worksheet2.write_formula('V37', '{=SUM(V13+V24+V36)}', cell_format)
    worksheet2.write_formula('W37', '{=SUM(W13+W24+W36)}', cell_format)
    worksheet2.write_formula('X37', '{=SUM(X13+X24+X36)}', cell_format)

    workbook.close()

    # # opening the file
    webbrowser.open(str(dire) + '\\xlsx\\' + y + '\\cashbook.xlsx')


# cashbook_xlsx_writer()

if __name__ == "__main__":
    cashbook_xlsx_writer()





