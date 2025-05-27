#pylint:disable = all
from ui_console import print_authors, get_des_path, change_des_path, get_list_sap, print_loading
from excel_handler import clear_sheet_data, write_df_to_excel, close_excel
from data_utils import create_dataframe, concat_df, resize_dataframe, filter_df, unique_data, df_list, password, day, in_day, night
import threading as th
import time
import os
from blessed import Terminal
from file_utils import copy_file_from_net
import constains

def process():
    
    des_path = get_des_path(change_des_path).strip('"')
    print('\n'*4)
    print(constains.term.center("----->   Nhập 1 để chọn ca ngày   <-----"))
    print(constains.term.center("----->   Nhập 2 để chọn ca đêm    <-----"))
    user_input = input()
    os.system('cls')
    if user_input in ('1', '2'):
        sap_list = get_list_sap()
        constains.progress += 10; constains.done.set()

        path_list = copy_file_from_net(sap_list)
        constains.progress += 10; constains.done.set()
        create_dataframe(path_list, password)
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
    
    task_1 = th.Thread(target=print_loading)
    task_1.start()
    start = time.time()
    process()
    task_1.join()
    print('\n' * 5)
    print(f'\nTotal time: {time.time() - start:.2f}s')
    constains.is_event.clear(); constains.progress = 0

if __name__ == "__main__":
    while True:
        print("\nNHẬP 'EXIT' ĐỂ THOÁT\nNHẤN ENTER ĐỂ TIẾP TỤC")
        if input().upper() == 'EXIT': break
        os.system('cls')
        main()
