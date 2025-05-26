from datetime import datetime, timedelta
import pandas as pd
import threading as th
import msoffcrypto
import io

df_list = []
password = 'J@bil2022'

day = datetime.today().replace(hour=6, minute=0, second=0)
in_day = day.replace(hour=23, minute=59, second=59)
night = (day - timedelta(days=1)).replace(hour=18, minute=0, second=0)

def load_data_with_key(path, password):
    try:
        with open(path, 'rb') as file:
            office_file = msoffcrypto.OfficeFile(file)
            office_file.load_key(password)
            decrypted = io.BytesIO()
            office_file.decrypt(decrypted)
            df = pd.read_excel(decrypted)
            df_list.append(df)
    except Exception:
        df = pd.read_excel(path)
        df_list.append(df)

def create_dataframe(path_list, password):
    import threading
    threads = [threading.Thread(target=load_data_with_key, args=(p, password)) for p in path_list]
    [t.start() for t in threads]
    [t.join() for t in threads]

def concat_df(list_data, callback):
    return callback(pd.concat(list_data, axis=0))

def resize_dataframe(df):
    return df.drop(columns=df.columns[[14, 15]])

def filter_df(df, start, end, callback):
    dt_col = pd.to_datetime(df.iloc[:, 5].str.split("|").str[1], format='%m/%d/%Y %I:%M:%S %p')
    valid = (dt_col >= start) & (dt_col <= end) & df.iloc[:, 12:].all(axis=1)
    return callback(df[valid])

def unique_data(df):
    return df[~df.duplicated(subset=df.columns[0], keep='first')]
