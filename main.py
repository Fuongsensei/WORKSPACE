#pylint:disable = all
from ui_console import print_authors, get_des_path, change_des_path, get_list_sap, print_loading,ask_user
from excel_handler import clear_sheet_data, write_df_to_excel, close_excel,call_macro,open_file
from data_utils import create_dataframe, concat_df, resize_dataframe, filter_df, unique_data,load_data_with_key, df_list, password, day, in_day, night
import threading as th
import time
import os
from blessed import Terminal
from file_utils import copy_file_from_net
import constains

global_des_path :str =''
global_user_input = ''

def process()->None:
    global global_des_path 
    global global_user_input

    
    des_path = get_des_path(change_des_path).strip('"')
    global_des_path = des_path
    
    open_file(des_path)
    print('\n'*4)
    print(constains.term.center("----->   Nhập 1 để chọn ca ngày   <-----"))
    print(constains.term.center("----->   Nhập 2 để chọn ca đêm    <-----"))

    user_input = input()
    global_user_input = user_input
    os.system('cls')


    constains.get_user_and_path.set()
    constains.is_run_macro.wait()

    if user_input in ('1', '2'):
        sap_list = get_list_sap()
        constains.progress += 10; constains.done.set()

        path_list = copy_file_from_net(sap_list)
        constains.progress += 10; constains.done.set()

        create_dataframe(load_data_with_key,path_list)
        constains.progress += 20; constains.done.set()

        df = concat_df(df_list, resize_dataframe)
        constains.progress += 20; constains.done.set()

        shift_df = filter_df(df, day if user_input == '1' else night,in_day if user_input == '1' else day, unique_data)
        constains.progress += 20; constains.done.set()

        constains.macro_done.wait()
        write_df_to_excel(shift_df,des_path,clear_sheet_data,close_excel)
        constains.progress += 20; constains.done.set()

        constains.is_event.set()


def run_macro()->None:
    constains.get_user_and_path.wait() 
    global global_des_path , global_user_input
    if ask_user("BẠN CÓ MUỐN CHẠY DATA"):
        os.system('cls')
        constains.is_run_macro.set()
        time.sleep(3)
        call_macro(global_des_path,global_user_input)
        constains.macro_done.set()
        
    else: 
        constains.is_run_macro.set()
        constains.macro_done.set()


print_authors()


def main():
    
    global global_des_path 
    global global_user_input
    task_1 : th.Thread = th.Thread(target=run_macro)
    task_2 : th.Thread = th.Thread(target=process)
    task_3 : th.Thread = th.Thread(target=print_loading)
    task_1.start()
    task_3.start()
    task_2.start()
    task_1.join()
    task_3.join()
    task_2.join()
    constains.macro_done.clear()
    constains.get_user_and_path.clear()
    constains.is_run_macro.clear()
    constains.is_event.clear()
    constains.progress = 0



if __name__ == "__main__":
    while True:
        print(constains.term.center("NHẬP 'EXIT' ĐỂ THOÁT"))
        print('\n' * 5)
        print(constains.term.center("NHẤN ENTER ĐỂ TIẾP TỤC"))
        if input().upper() == 'EXIT': break
        os.system('cls')
        main()
