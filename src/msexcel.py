import xlwings as xw


class Msexcel():

    def __init__(self, book=None):
        if book:
            self.xls = xw.Book(book)
            self.books = xw.books
        else:
            print("Specify the name of the excel file")
            exit()

    def sheet(self, sname):
         return self.xls.sheets(sname)

    def get_col_dat(self, sht_obj: object, col_beg: str, col_end: str, row_beg: int, rows: int):
        rng = col_beg + str(row_beg) + ':' + col_end + str(row_beg)
        col = sht_obj.range(rng).value
        if col_beg == col_end:
            col = [col]
        if rows > 0:
            rng = col_beg + str(row_beg+1) + ':' + \
                col_end + str(row_beg+1+rows)
            dat = sht_obj.range(rng).value
        else:
            dat = None
        return col, dat

    def get_lst_fm_rng(self, sht_obj: object, col_beg: str, col_end: str, row_id: int):
        lst = sht_obj.range(col_beg + str(row_id) + ":" + col_end + str(row_id)).value
        if lst:
            if isinstance(lst,list):
                if any(item is None for item in lst):
                    return False, lst
                else:
                    return True, lst
            else:
                return True, [lst]
        return False, []

