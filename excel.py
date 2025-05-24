#pylint:disable=all
import time
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
from blessed import Terminal
import sys
from colorama import init, Fore, Back, Style
done : th.Event = th.Event()
is_event : th.Event = th.Event()
progress : float = 0
df_list : list  = []
term = Terminal()
day : datetime = datetime.today().replace(hour=6,minute=0,second=0)
in_day : datetime = day.replace(hour=23,minute=59,second=59)
night : datetime = (day - timedelta(days=1)).replace(hour=18,minute=0,second=0)
password : str = 'J@bil2022'
current_user : str = getpass.getuser()
init()


def apply_color(char : str | list  ) -> str:
        if isinstance(char,str):
                apply : str  = Fore.GREEN + char + Fore.RESET
                return apply
        else :
                apply : str = ''
                for i in char:
                        apply += '|'+ Fore.GREEN + i + Fore.RESET + '|'
                return apply


def print_authors() -> None:
        authors = 'Developed by phuongnguyen1183 using Python, compiled to C'
        name = authors.upper()
        char = list(name)
        for i in char :
                sys.stdout.write(f'{apply_color(i)}')
                sys.stdout.flush()
                time.sleep(0.03)
        print('\n'*2)        


print_authors()


def get_des_path(callback) -> str:
        contain_path : str = r'C:\Users\nguye\Documents\path.txt'
        is_exists : bool  = os.path.exists(contain_path)
        if is_exists:
                with open(contain_path,mode='r') as file:
                        path: str = file.read()
                user_input:str = input(f"Đường dẫn hiện tại là {path} bạn có muốn thay đổi   '[{apply_color('Y')}]/{apply_color("N")}'? :...")
                print(f'{'\n'*2}')
                if user_input.upper() == 'Y': return callback()
                else  :return path

        else:
                des_path : str = input(f"Vui lòng copy đường dẫn file {apply_color('Report Scan Verify Shiftly (RCV)')} dán vào đây..: ")
                with open(contain_path,mode='w+') as file:
                        file.write(des_path)
                return des_path
                
                
def change_des_path() -> str:
        changed_path : str = input("Vui lòng nhập đường dẫn mới :")
        return changed_path




def get_list_sap()-> list :
        user_input : str = input("Vui lòng nhập số SAP cách nhau bằng ký tự trừ số nguyên:....")
        print(f'{'\n'*2}')
        user_sap : list = [user.strip() for user in re.split(r'[\D]+',user_input)]
        check_input = input(f"Đây là danh sách SAP ---> {apply_color(user_sap)} <--- vui lòng kiểm tra nếu đúng nhấn enter còn nếu sai nhập N để nhập lại...:")
        print(f'{'\n'*2}')
        if check_input.upper() != "N":   return user_sap
                
        else:
                return get_list_sap()


def copy_file_from_net(list_sap)->list:
        list_path : list = []
        try:
                for sap in list_sap:
                        path : str =rf"D:\Program Files\Download\GR Verification {sap}.xlsx"
                        path_copy = rf'C:\Users\nguye\Documents\WORKSPACE\GR Verification {sap}.xlsx'
                        shutil.copy(path, path_copy)
                        list_path.append(path_copy) 
        except Exception as error:
                print(f"Lỗi :  {type(error).__name__}")
                
        return list_path
        

def  load_data_with_key(path,password) -> None:
        with open(path,mode='rb') as file:
                file_encrytp = ms.OfficeFile(file)
                file_encrytp.load_key(password)
                file_decrypted = io.BytesIO()
                file_encrytp.decrypt(file_decrypted)
                data : pd.DataFrame = pd.read_excel(file_decrypted)
                df_list.append(data)
        



                
def concat_df(list_data:list[pd.DataFrame],callback) -> pd.DataFrame:
        data_concat = pd.concat(list_data,axis=0)
        return callback(data_concat) 


def resize_dataframe(data_frame:pd.DataFrame) -> pd.DataFrame:
        data : pd.DataFrame = data_frame.drop(columns=data_frame.columns[[14,15]])
        return data


def filter_df(data:pd.DataFrame,start_time:datetime,end_time:datetime,callback)-> pd.DataFrame:
        user_and_time_col : pd.Series = data.iloc[:,5].str.split("|").str[1]
        date_col: pd.Series = pd.to_datetime(user_and_time_col,format='%m/%d/%Y %I:%M:%S %p')
        bool_col : pd.DataFrame = data.drop(columns= data.columns[11])
        mask : pd.Series = ((date_col >= start_time) & (date_col <= end_time) & (bool_col.iloc[:,9:].all(axis=1)))
        return callback(data[mask])


def unique_data(data: pd.DataFrame)-> pd.DataFrame:
        data_unique : pd.Series = data.duplicated(subset=data.columns[0],keep='first')
        return data[~data_unique]


def clear_sheet_data(file_path:str)-> None:
        app = xw.App(visible=True, add_book=False)
        wb = app.books.open(file_path, update_links=False)
        sheet = wb.sheets["Verify data"]
        last_row = sheet.range("A" + str((sheet.cells.last_cell).row)).end('up').row
        if last_row > 2:
                sheet.api.Rows(f"3:{last_row}").Delete()
        sheet.range('A2:N2').clear()
        

def write_df_to_excel(data: pd.DataFrame,callback,des_path)-> None:
        try:
                
                wb = xw.Book(des_path)
                sheet = wb.sheets['Verify data']
                sheet.range('A2:N2').value = data.values
        except Exception as error:
                print(f'\nLỗi xung đột tiến trình : File Report Scan Verify Shiftly (RCV) đang được mở sau đây sẽ đóng file vui lòng chạy lại')
                callback("EXCEL.EXE")
        
        
        
def close_excel(app_name : str)->None:
        for proc in psutil.process_iter():
                if proc.name() == app_name:
                        proc.kill()
        
        
def create_dataframe(dercrypt,list_file_path)-> None:
        threads : list[th.Thread] = []
        for path in list_file_path:
                task = th.Thread(target=dercrypt,args=(path,password))
                task.start()
                threads.append(task)
        for i in threads:
                i.join()
        

def print_loading()->None:
        global progress
        
        while not is_event.is_set():
                start = progress
                done.wait()
                for i in range(start,progress+1):
                        bar = '█'*i
                        sys.stdout.write(f'\rLoading: {bar:░<100} {i}%')
                        sys.stdout.flush()
                        time.sleep(0.03)
                done.clear()


def process() ->None:
        des_path = get_des_path(change_des_path)
        des_path = des_path.strip('"')
        global progress
        print(term.center("----->   Nhập 1 để chọn ca ngày   <-----"))
        print(term.center("----->   Nhập 2 để chọn ca đêm    <-----"))
        user_input = input()
        if user_input !='':
                if user_input =='1':
                        sap_list : list = get_list_sap()
                        progress = progress + 10
                        done.set()
                        
                        path_list : list = copy_file_from_net(sap_list)
                        progress = progress + 10
                        done.set()
                        
                        create_dataframe(load_data_with_key,path_list)
                        progress = progress + 20
                        done.set()
                        
                        data_frame : pd.DataFrame = concat_df(df_list,resize_dataframe)
                        progress = progress + 20
                        done.set()
                        
                        data_filter : pd.DataFrame = filter_df(data_frame,day,in_day,unique_data)
                        progress = progress + 20
                        done.set()
                        clear_sheet_data(des_path)
                        write_df_to_excel(data_filter,close_excel,des_path)
                        progress = progress + 20
                        done.set()
                        is_event.set()
                        
                        
                        
                elif user_input =='2':
                        sap_list : list = get_list_sap()
                        progress = progress + 10
                        done.set()
                        
                        path_list : list = copy_file_from_net(sap_list)
                        progress = progress + 10
                        done.set()
                        
                        create_dataframe(load_data_with_key,path_list)
                        progress = progress + 20
                        done.set()
                        
                        data_frame : pd.DataFrame = concat_df(df_list,resize_dataframe)
                        progress = progress + 20
                        done.set()
                        
                        data_filter : pd.DataFrame = filter_df(data_frame,night,day,unique_data)
                        progress = progress + 20
                        done.set()
                        clear_sheet_data(des_path)
                        write_df_to_excel(data_filter,close_excel,des_path)
                        progress = progress + 20
                        done.set()
                        is_event.set()
                
        else : return


def main()->None:

        task_1= th.Thread(target=print_loading)
        task_2 = th.Thread(target=process)
        task_1.start()
        task_2.start()
        task_2.join()
        task_1.join()
        

while True :
        print()
        print("NHẬP 'EXIT' ĐỂ THOÁT")
        print()
        print("NHẤN ENTER ĐỂ TIẾP TỤC")
        ip = input()
        if ip.upper() == 'EXIT': break
        else:
                os.system('cls')
                main()


