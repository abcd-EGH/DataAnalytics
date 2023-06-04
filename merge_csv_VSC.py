import pandas as pd
import os

# DO NOT CHANGE OR DAMAGE THIS PROGRAM WITHOUT PERMISSION
print("#############################################################")
print("# This program merges several excel files(.xlsx).           #")
print("# Before run, check your excel files are in Desktop folder. #")
print("# DO NOT CHANGE OR DAMAGE THIS PROGRAM WITHOUT PERMISSION.  #")
print("#                                      - made by Lee, J. H. #")
print("#############################################################")
print()

windows_user_name = os.path.expanduser('~')

df_base_name = input("ENTER the first file name:")
if '.xlsx' not in df_base_name:
    df_base_name = df_base_name + '.xlsx'
try:
    windows_user_name = os.path.expanduser('~')
    df_base_name = f"{windows_user_name}\Desktop\{df_base_name}"
    df_base = pd.read_excel(df_base_name, engine = 'openpyxl')
except:
    print("The file does not exist or is an invalid file name.")
    print("Please run again.")
    os.system("pause")
    
else:
    while True:
        print()
        print("------- If you have entered all file names -> Enter 0 -------")
        df_name = input('ENTER the file name(or 0):')
        if df_name == '0':
            break
        if '.xlsx' not in df_name:
            df_name = df_name + '.xlsx'
        try:
            windows_user_name = os.path.expanduser('~')
            df_name = f"{windows_user_name}\Desktop\{df_name}"
            df_add = pd.read_excel(df_name, engine = 'openpyxl')
        except:
            print("A file with that name does not exist. Enter the file name again.")
            continue
        else:
            df_base = pd.concat([df_base,df_add],ignore_index=True)
            print("Merged successfully.")

    print("The length of the result excel file: ", len(df_base))

    df_final_name = input("ENTER the a file name to save: ")
    if '.xlsx' not in df_final_name:
        df_final_name = df_final_name + '.xlsx'
    try:
        windows_user_name = os.path.expanduser('~')
        df_final_name = f"{windows_user_name}\Desktop\{df_final_name}"
        df_base.to_excel(df_final_name, index = False)
    except:
        print("Sorry, Unable to save file.")
        print("Please set the file name different from the other file,")
        print("Or run this program again.")
        os.system("pause")
    else:
        print("File saved successfully. Check the folder.")
        print("Thank you for using this program.")
        os.system("pause")