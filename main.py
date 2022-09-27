import pymongo
from pymongo import MongoClient
import certifi
import datetime
import xlwings as xw
import pandas as pd
import os
import glob
import openpyxl

def covert_row_col_number_to_letter(row, col):  # starts from 1
    col_letter = []
    while col:
        col_letter.append(col % 26)
        col //= 26
    col_letter = "".join([chr(ord('A') + x - 1) for x in col_letter])[::-1]
    return col_letter + str(row)


class Order():
    _client_string = 'mongodb+srv://yan:56358335@downton.huob4ua.mongodb.net/?retryWrites=true&w=majority'

    def __init__(self):
        pass

    @classmethod
    def recreate_table(cls):
        client = MongoClient(cls._client_string, tlsCAFile=certifi.where())
        db = client.Downton
        db.drop_collection('Orders')
        col = db["Orders"]
        client.close()

    @classmethod
    def insert_documents(cls, docs):
        client = MongoClient(cls._client_string, tlsCAFile=certifi.where())
        db = client.Downton
        col = db.Orders
        _ids=[]
        for doc in docs:
            if None in doc:
                del doc[None]
            if doc.get('_id') is None or str(doc.get('_id')).strip() == '':
                _id = cls.get_new_id()
                while col.find_one({'_id': _id}) is not None:
                    _id = cls.get_new_id()
                doc['_id'] = cls.get_new_id()
            doc['_id'] = str(doc['_id'])
            col.update_one({'_id': doc['_id']}, {'$set': doc}, upsert=True)
            _ids.append(_id)
        client.close()
        return _ids

    # @classmethod
    # def find_documents(cls,filter={},projection={},sort):
    #     client = MongoClient(cls._client_string,tlsCAFile=certifi.where())
    #     db = client.Downton
    #     col = db.Orders
    #     docs=list(col.find())
    #     return docs

    @staticmethod
    def get_new_id():
        now = datetime.datetime.now()
        res = now.strftime("%Y%m%d-%H%M%S%f")
        return res

    @staticmethod
    def convert_sheet_to_docs():
        wb = xw.Book.caller()
        ws = wb.sheets.active
        rownum = ws.range('A1').current_region.last_cell.row
        column = ws.range('A1').current_region.last_cell.column
        df = ws.range((2, 1), (rownum, column)).options(pd.DataFrame,
                                                        header=1,
                                                        index=False,
                                                        ).value
        df.dropna(how='all', inplace=True)
        return df.to_dict('records')



def recreate_table():
    client = MongoClient(Order._client_string, tlsCAFile=certifi.where())
    db = client.Downton
    db.drop_collection('Orders')
    col = db["Orders"]
    client.close()
    wb = xw.Book.caller()
    wb.sheets['hidden'].range('B1').value = 1

def upsert_sheet():
    docs = Order.convert_sheet_to_docs()
    Order.insert_documents(docs)
    wb = xw.Book.caller()
    wb.sheets['hidden'].range('B1').value = 1


def upsert_selection():
    wb = xw.Book.caller()
    cellRange = wb.selection
    if cellRange.row == 2 and cellRange.column == 1:
        df = xw.load(index=False, header=1)
        docs = df.to_dict('records')
        Order.insert_documents(docs)
    else:
        raise Exception('Wrong Selected Range')
    wb.sheets['hidden'].range('B1').value = 1


def download_sheet():
    wb = xw.Book.caller()
    ws = wb.sheets.active
    max_row = ws.range('A1').current_region.last_cell.row
    max_column = ws.range('B1').current_region.last_cell.column
    title = ws.range((2, 1), (2, max_column)).value

    ws.range((3, 1), (max(max_row, 3), max_column)).clear_contents()

    client = MongoClient(Order._client_string, tlsCAFile=certifi.where())
    db = client.Downton
    col = db.Orders
    docs = list(col.find().sort([("_id", 1)]))
    row_no = 3
    # print(title)
    for doc in docs:
        ws.range('A' + str(row_no)).value = [doc.get(t) for t in title]
        row_no += 1
    client.close()
    wb.sheets['hidden'].range('B1').value = 1

def create_orders_from_booking_forms():
    wb = xw.Book.caller()
    ws=wb.sheets.active
    ws.range("H:H").api.Clear()
    cwd=os.path.dirname(wb.fullname)
    wb=openpyxl.open(os.path.join(cwd,"config.xlsx"))
    ws=wb["Booking form"]
    cell_to_column_name={}
    for row in ws.rows:
        cell_to_column_name[row[0].value]=row[1].value

    docs=[]
    for fn in glob.glob(os.path.join(cwd,"booking forms","*.xlsx")):
        wb=openpyxl.open(fn)
        ws=wb.active
        doc={}
        for cell,column_name in cell_to_column_name.items():
            doc[column_name]=ws[cell].value
        docs.append(doc)

    wb = xw.Book.caller()
    ws=wb.sheets.active
    _ids=Order.insert_documents(docs)
    ws.range("H1").value=[[_id] for _id in _ids]
    wb.sheets['hidden'].range('B1').value = 1

if __name__ == '__main__':
    pass
    # Expects the Excel file next to this source file, adjust accordingly.
    xw.Book('mainbook.xlsm').set_mock_caller()
    # create_orders_from_booking_forms()

#
#
# # Order.create_table()
#
# print(covert_row_col_number_to_letter(99,16384))
