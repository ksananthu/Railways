from py_files.dfc_xlsx_fun import dfc_xlsx_writer
from py_files.ll_xlsx_fun import ll_xlsx_writer
from py_files.gst_xlsx_fun import gst_xlsx_writer
from py_files.kfc_xlsx_fun import kfc_xlsx_writer


def other_create():
    while True:
        print('...............Select from below options.................')
        print('Other(LL ,DFC and KFC) or GST')
        print('\n')
        oth_op = input('Others(o) or GST(g)...: ')

        if oth_op == 'o':
            dfc_xlsx_writer()
            kfc_xlsx_writer()
            ll_xlsx_writer()
            break

        elif oth_op == 'g':
            gst_xlsx_writer()
            break

        else:
            continue










