#pylint:disable = all
import os, shutil, getpass

def copy_file_from_net(sap_list):
    current_user = getpass.getuser()
    list_path : list = []
    for sap in sap_list:
        src : str =rf"D:\Program Files\Download\GR Verification {sap}.xlsx"
        dst = rf'C:\Users\nguye\Documents\WORKSPACE\GR Verification {sap}.xlsx'
        shutil.copy(src, dst)
        list_path.append(dst)
    return list_path
