from ui_console import print_authors, get_des_path, change_des_path, get_list_sap, print_loading
from excel_handler import clear_sheet_data, write_df_to_excel, close_excel
from data_utils import create_dataframe, concat_df, resize_dataframe, filter_df, unique_data, df_list, password, day, in_day, night
import threading as th
import time
import os
from blessed import Terminal
from file_utils import copy_file_from_net

done = th.Event()
is_event = th.Event()
progress = 0
term = Terminal()

def process():
    global progress
    des_path = get_des_path(change_des_path).strip('"')
    print(term.center("----->   Nhập 1 để chọn ca ngày   <-----"))
    print(term.center("----->   Nhập 2 để chọn ca đêm    <-----"))
    user_input = input()
    os.system('cls')
    if user_input in ('1', '2'):
        sap_list = get_list_sap()
        progress += 10; done.set()

        path_list = copy_file_from_net(sap_list)
        progress += 10; done.set()
        create_dataframe(path_list, password)
        progress += 20; done.set()
        df = concat_df(df_list, resize_dataframe)
        progress += 20; done.set()
        shift_df = filter_df(df, day if user_input == '1' else night,in_day if user_input == '1' else day, unique_data)
        progress += 20; done.set()
        clear_sheet_data(des_path)
        write_df_to_excel(shift_df, close_excel, des_path)
        progress += 20; done.set()
        is_event.set()


print_authors()


def main():
    global progress
    task_1 = th.Thread(target=print_loading, args=(done, is_event, lambda: progress))
    task_1.start()
    start = time.time()
    process()
    task_1.join()
    print('\n' * 5)
    print(f'\nTotal time: {time.time() - start:.2f}s')
    is_event.clear(); progress = 0

if __name__ == "__main__":
    while True:
        print("\nNHẬP 'EXIT' ĐỂ THOÁT\nNHẤN ENTER ĐỂ TIẾP TỤC")
        if input().upper() == 'EXIT': break
        os.system('cls')
        main()
