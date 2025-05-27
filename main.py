#pylint:disable = all
from ui_console import print_authors, get_des_path, change_des_path, get_list_sap, print_loading,ask_user
from excel_handler import clear_sheet_data, write_df_to_excel, close_excel,call_macro
from data_utils import create_dataframe, concat_df, resize_dataframe, filter_df, unique_data,load_data_with_key, df_list, password, day, in_day, night
import threading as th
import time
import os
from blessed import Terminal
from file_utils import copy_file_from_net
import constains

global_des_path :str =''
global_user_input = ''
def process():
    global global_des_path 
    global global_user_input

    des_path = get_des_path(change_des_path).strip('"')
    
    global_des_path = des_path
    print('\n'*4)
    print(constains.term.center("----->   Nhập 1 để chọn ca ngày   <-----"))
    print(constains.term.center("----->   Nhập 2 để chọn ca đêm    <-----"))
    user_input = input()
    global_user_input = user_input
    os.system('cls')
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

        clear_sheet_data(des_path)
        write_df_to_excel(shift_df, close_excel, des_path)
        constains.progress += 20; constains.done.set()

        constains.is_event.set()




print_authors()


def main():
    global global_des_path 
    global global_user_input
    task_1 = th.Thread(target=print_loading)
    task_1.start()
    start = time.time()
    process()
    task_1.join()
    print('\n' * 5)
    print(f'\nTotal time: {time.time() - start:.2f}s')
    constains.is_event.clear(); constains.progress = 0

    os.system('cls')
    is_run_macro = ask_user()
    if is_run_macro:
        call_macro(global_des_path,global_user_input)
    else :return

if __name__ == "__main__":
    while True:
        print("\nNHẬP 'EXIT' ĐỂ THOÁT\nNHẤN ENTER ĐỂ TIẾP TỤC")
        if input().upper() == 'EXIT': break
        os.system('cls')
        main()
