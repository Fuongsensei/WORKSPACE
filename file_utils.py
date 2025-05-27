#pylint:disable = all
import os, shutil, getpass
import datetime as dt 
def copy_file_from_net(sap_list):
    current_user : str  = getpass.getuser()
    current_time : dt  = dt.date.today()
    short_month : str  = current_time.strftime('%b')
    year : dt = current_time.year
    list_path : list[str] = []
    for sap in sap_list:
        src : str =rf"\\AWASE1HCMICAP01\AppsData\GR Ver Report\{short_month} {year}\GR Verification {sap}.xlsx"
        dst = rf"C:\Users\{current_user}\Documents\GR Verification {sap}.xlsx"
        shutil.copy(src, dst)
        list_path.append(dst)
    return list_path
