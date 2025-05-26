import os, shutil, getpass

def copy_file_from_net(sap_list):
    current_user = getpass.getuser()
    list_path : list = []
    for sap in sap_list:
        src : str = rf"\\AWASE1HCMICAP01\AppsData\GR Ver Report\May 2025\GR Verification {sap}.xlsx"
        dst : str = rf"C:\Users\{current_user}\Documents\GR Verification {sap}.xlsx"
        shutil.copy(src, dst)
        list_path.append(dst)
    return list_path
