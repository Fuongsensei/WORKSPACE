#pylint:disable = all

import xlwings as xw
import psutil

def clear_sheet_data(file_path: str) -> None:
    wb = xw.Book(file_path, update_links=False)
    sheet = wb.sheets["Verify data"]
    last_row = sheet.range("A" + str((sheet.cells.last_cell).row)).end('up').row
    if last_row > 2:
        sheet.api.Rows(f"3:{last_row}").Delete()
    sheet.range('A2:N2').clear()

def write_df_to_excel(data, callback, des_path:str) -> None:
    try:
        wb = xw.Book(des_path)
        sheet = wb.sheets['Verify data']
        sheet.range('A2:N2').value = data.values
    except Exception:
        print(f'\nLỗi xung đột: File đang mở — sẽ tự đóng.')
        callback("EXCEL.EXE")

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
    macros[user_input]()
    macro_delete = wb.macro("Xoa_Data")
    macro_run_data = wb.macro('chay_Data')
    macro_delete()
    macro_run_data()