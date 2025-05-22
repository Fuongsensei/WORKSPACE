#pylint:disable=all
import time
from traceback import print_tb
from openpyxl import load_workbook
import pandas as pd
import psutil 
from datetime import date, timedelta
import io
from datetime import datetime 
import xlwings as xw
import msoffcrypto as ms
import os 
import getpass 
import shutil
import threading as th
import re
start = time.time()
df_list : list  = []
day : datetime = datetime.today().replace(hour=6,minute=0,second=0)
in_day : datetime = day.replace(hour=23,minute=59,second=59)
night : datetime = (day - timedelta(days=1)).replace(hour=18,minute=0,second=0)
password : str = 'J@bil2022'
current_user : str = getpass.getuser()


def get_des_path(callback) -> str:
        contain_path : str = r'C:\Users\nguye\Documents\path.txt'
        is_exists : bool  = os.path.exists(contain_path)
        if is_exists:
                with open(contain_path,mode='r') as file:
                        path: str = file.read()
                user_input:str = input(f"Đường dẫn hiện tại là {path} bạn có muốn thay đổi  Y/N ? :...")
                if user_input.upper() == 'Y': return callback()
                elif user_input.upper() == 'N':return path
                else : 
                        print("Lựa chọn không hợp lý")
                        return

        else:
                des_path : str = input("Vui lòng copy đường dẫn file Report Scan Verify Shiftly (RCV) dán vào đây..: ")
                with open(contain_path,mode='w+') as file:
                        file.write(des_path)
                return des_path
                
                
def change_des_path() -> str:
        changed_path : str = input("Vui lòng nhập đường dẫn mới :")
        return  changed_path


def get_list_sap()-> list :
        user_input : str = input("Vui lòng nhập số SAP cách nhau bằng ký tự trừ số nguyên:....")
        user_sap : list = [user.strip() for user in re.split('[\D]+',user_input)]
        return user_sap

def copy_file_from_net(list_sap,callback)->None:
        list_path : list = []
        try:
                for sap in list_sap:
                        path : str =f"D:\\Program Files\\Download\\GR Verification {sap}.xlsx"
                        path_copy = f'C:\\Users\\nguye\\Documents\\WORKSPACE\\GR Verification {sap}.xlsx'
                        shutil.copy2(path, path_copy)
                        list_path.append(path_copy) 
        except Exception as error:
                print(f"Lỗi :  {type(error).__name__}")
                
        return list_path
        

def  load_data_with_key(path,password):
        with open(path,mode='rb') as file:
                file_encrytp = ms.OfficeFile(file)
                file_encrytp.load_key(password)
                file_decrypted = io.BytesIO()
                file_encrytp.decrypt(file_decrypted)
                data : pd.DataFrame = pd.read_excel(file_decrypted)
                df_list.append(data)
        



                
# def concat_df(list_data:list[pd.DataFrame],callback) -> pd.DataFrame:
#         data_concat = pd.concat(list_data,axis=0)
#         return callback(data_concat) 


# def resize_dataframe(data_frame:pd.DataFrame) -> pd.DataFrame:
#         data : pd.DataFrame = data_frame.drop(columns=data_frame.columns[[14,15]])
#         return data


# def filter_df(data:pd.DataFrame,start_time:datetime,end_time:datetime,callback)-> pd.DataFrame:
#         user_and_time_col : pd.Series = data.iloc[:,5].str.split("|").str[1]
#         date_col: pd.Series = pd.to_datetime(user_and_time_col,format='%m/%d/%Y %I:%M:%S %p')
#         bool_col : pd.DataFrame = data.drop(columns= data.columns[11])
#         mask : pd.Series = ((date_col >= start_time) & (date_col <= end_time) & (bool_col.iloc[:,9:].all(axis=1)))
#         return callback(data[mask])


# def unique_data(data: pd.DataFrame)-> pd.DataFrame:
#         data_unique : pd.Series = data.duplicated(subset=data.columns[0],keep='first')
#         return data[~data_unique]


# def clear_sheet_data(file_path:str)-> None:
#         app = xw.App(visible=True, add_book=False)
#         wb = app.books.open(file_path, update_links=False)
#         sheet = wb.sheets["Verify data"]
#         last_row = sheet.range("A" + str((sheet.cells.last_cell).row)).end('up').row
#         if last_row > 2:
#                 sheet.api.Rows(f"3:{last_row}").Delete()
#         sheet.range('A2:N2').clear()
        

# def write_df_to_excel(data: pd.DataFrame,callback)-> None:
#         try:
#                 wb = xw.Book(des_path,update_links=False)
#                 sheet = wb.sheets['Verify data']
#                 sheet.range('A2:N2').value = data.values
#         except Exception as error:
#                 print('Lỗi : File Report Scan Verify Shiftly (RCV) đang được mở trước đó nên sau đây sẽ đóng file vui lòng chạy lại chương trình')
#                 callback("EXCEL.EXE")
        
        
        
# def close_excel(app_name : str):
#         for proc in psutil.process_iter():
#                 if proc.name() == app_name:
#                         proc.kill()
        
        
# def main() -> None:
#         print(
#                                 "######################################################################\n"
#                                 "##                                                                  ##\n"
#                                 "##                      Nhập 1 để chọn day                          ##\n"
#                                 "##                                                                  ##\n"
#                                 "##                                                                  ##\n"
#                                 "##                                                                  ##\n"
#                                 "##                      Nhập 2 để chọn night                        ##\n"
#                                 "##                                                                  ##\n"
#                                 "##                                                                  ##\n"
#                                 "##                                                                  ##\n"
#                                 "##                      Nhập close để thoát                         ##\n"
#                                 "##                                                                  ##\n"
#                                 "######################################################################")

#         user_input :str = input()
#         start = time.time()
#         if user_input != '' and isinstance(user_input,str):
#                 if user_input == '1':
#                         load_data_with_key(df_list,password)
#                         clear_sheet_data(des_path)
#                         write_df_to_excel(filter_df(concat_df(df_list,resize_dataframe),day,in_day,unique_data),close_excel)
#                 elif user_input == '2':
#                         load_data_with_key(df_list,password)
#                         clear_sheet_data(des_path)
#                         write_df_to_excel(filter_df(concat_df(df_list,resize_dataframe),night,day,unique_data),close_excel)
#                 elif user_input =='close':
#                         return
#                 else:
#                         print("Vui lòng nhập đúng lựa chọn")
#                         main()
#                         return
                        
#         print(f"Tổng thời gian bằng {time.time()-start}s")


# load_data_with_key(copy_file_from_net(get_list_sap()))
def main() ->None:
        
        user_sap : list [str] = get_list_sap()
        start = time.time()
        list_file_path : list[str] = copy_file_from_net(user_sap)
        threads : list[th.Thread] = []
        for path in list_file_path:
                task = th.Thread(target=load_data_with_key,args=(path,password))
                task.start()
                threads.append(task)
        for i in threads:
                i.join()
        df = pd.concat(df_list,axis=0)
        print(df)
        print(f'{time.time()-start:4f}s')
main()


input("Nhấn Enter để thoát...")

