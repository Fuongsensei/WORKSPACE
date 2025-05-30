#pylint:disable = all

import xlwings as xw
import psutil
import time as t 
import os
import pandas as pd
import datetime as dt 
from xlsx2csv import Xlsx2csv

def clear_sheet_data(wb) -> None:
    sheet = wb.sheets["Verify data"]
    last_row = sheet.range("A" + str((sheet.cells.last_cell).row)).end('up').row
    if last_row > 2:
        sheet.api.Rows(f"3:{last_row}").Delete()
    sheet.range('A2:N2').clear()


def write_df_to_excel(data, des_path:str, callback_1,callback_2) -> None:
    try:
        wb = xw.Book(des_path,update_links=False)
        callback_1(wb)
        sheet = wb.sheets['Verify data']
        sheet.range('A2:N2').value = data.values
        wb.save()
    except Exception:
        print(f'\nLỗi xung đột: File đang mở — sẽ tự đóng.')
        callback_2("EXCEL.EXE")


def close_excel(app_name: str) -> None:
    for proc in psutil.process_iter():
        if proc.name() == app_name:
            proc.kill()


def call_macro(des_path,user_input):
    wb = xw.Book(des_path)
    macros : dict = {
        '1': wb.macro("Ngay"),
        '2': wb.macro('DEM')
    }
    macro_delete = wb.macro("Xoa_Data")
    macro_run_data = wb.macro('chay_Data')

    macros[user_input]()
    t.sleep(1)
    macro_delete()
    t.sleep(3)
    macro_run_data()

def check_state_file(path:str) -> bool:
        is_open : bool = True
        file_name :str = os.path.basename(path)
        for app in xw.apps:
            for book in app.books:
                        if book.name != file_name:
                                is_open = False
        return is_open


def open_file(path:str) -> None:
    if check_state_file(path):
            os.startfile(path)
            return
    print("FILE ĐANG ĐƯỢC MỞ ")
    return



des_path:str =r"C:\Users\3601183\Desktop\Report Scan Verify Shiftly (RCV).xlsm"


file_name :str = os.path.basename(des_path)

def delete_blank(path):
        wb= xw.Book(path)
        sheet = wb.sheets['GRN (10 so)']
        last_rng : int = (sheet.cells.last_cell).row 
        last_row : int = sheet.range(last_rng,1).end('up').row
        sheet.api.Range(f'A{last_row+1}:A{last_rng}').EntireRow.Delete()
        wb.save()
        print('Done')


def get_criteria(path:str ):
        import  getpass
        csv_path : str =rf"C:\Users\{getpass.getuser()}\Documents\entered_date.csv"
        Xlsx2csv(path,outputencoding='utf-8').convert(csv_path,sheetid=3)
        entered_str_date : pd.DataFrame = pd.read_csv(csv_path,usecols=[2],dtype=str)
        entered_date = pd.to_datetime(entered_str_date.iloc[:,0],format='%m-%d-%y')
        mask : pd.Series[bool]  = entered_date == entered_date.iloc[0]
        is_false : int = mask.idxmin()
        criteria : str = dt.datetime.strftime(entered_date[is_false],format='%m-%d-%Y')
        return criteria

def delete_entered_on_date(path:str)-> None:
        
        wb = xw.Book(path,update_links=False)
        sheet = wb.sheets['GRN (10 so)']
        last_row : int = sheet.range((sheet.cells.last_cell).row,3).end('up').row
        sheet.api.Range('C1').AutoFilter(Field=3, Criteria1='06-01-2025')
        
        try:
                visible_row = sheet.api.Range(f'C2:C{last_row}').SpecialCells(12)
                visible_row.EntireRow.Delete()

        except Exception as er:
                print('KHÔNG TÌM THẤY ')
                
        sheet.api.ShowAllData()





# delete_blank(des_path)
# get_criteria(des_path)
delete_entered_on_date(des_path)