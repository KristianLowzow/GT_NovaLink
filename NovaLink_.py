from win32com.client import Dispatch
import pandas as pd
from openpyxl import load_workbook
import os
import time
import pandas as pd
import shutil
import glob
import datetime
import numpy as np
import ctypes  # An included library with Python install.
pd.options.display.width = 0 # Viser alle colloner ved print
pd.options.display.float_format = '{:,.2f}'.format

#Setter opp mapper og leser definisjonsfiler
folder = os.path.dirname(os.path.abspath(__file__))
folder_Inn = folder + "\Inndata\\"
folder_Resultat = folder + "/Resultat/"
folder_Definisjon = folder + "/Definisjon/"
NovaLinkDef = folder_Definisjon+"NovaLinkDef.xlsx"
#Definisjonsfiler
df = pd.read_excel(NovaLinkDef)
Inndata, Lenkefil_Start = df['Input'].iloc[0], df['Input'].iloc[1]

#Fuksjoner
def find_File_And_Rename(tellepunkt):
    shutil.copy(os.path.join(folder_Definisjon, "TAnalyse-Mal.xlsx"), folder_Resultat)
    shutil.copy(os.path.join(folder_Inn, tellepunkt), folder_Resultat)
    rename_file = folder_Resultat + "TAnalyse-Mal.xlsx"
    newName = folder_Resultat + tellepunkt
    newName = newName.replace("TData", "TAnalyse")
    os.rename(str(rename_file),str(newName))
    TAnalyse = newName.replace("\\", "/")
    TData = os.path.join(folder_Resultat, tellepunkt).replace("\\", "/")
    print (TAnalyse, TData)
    return str(TAnalyse), str(TData)

def delete_folderFiles(folder):
    dir = folder
    for f in os.listdir(dir):
        os.remove(os.path.join(dir, f))

def run_macro(TAnalyse, TData, com_instance, tellepunkt):
    print(TData)
    wb2 = com_instance.workbooks.open(folder_Definisjon + str(Lenkefil_Start))
    wb3 = com_instance.workbooks.open(TData)
    wb = com_instance.workbooks.open(TAnalyse)
    com_instance.AskToUpdateLinks = False
    print("Workbook LinkSources: ", wb.LinkSources())
    try:
       wb.ChangeLink(folder_Definisjon + str(Lenkefil_Start), TData)
       #wb.UpdateLink(Name=wb.LinkSources())
       print("DID it work!?!?!!!!!!!!!!!!!")

    except Exception as e:
       print(e)

    finally:
       try:
           wb.Close(True)
       except:
        print("Klatte ikke lukke WB")
        try:
            wb2.Close(True)
        except:
            print("Klatte ikke lukke WB2")
       try:
           wb3.Close(True)
       except:
        print("Klatte ikke lukke WB3")
       wb = None
       wb2 = None
       wb3 = None
    shutil.move(TData, folder + "/Koblet/" + tellepunkt)
    shutil.move(TAnalyse, folder + "/Koblet/" + tellepunkt.replace("TData", "TAnalyse"))
    delete_folderFiles(folder_Inn)
    delete_folderFiles(folder_Resultat)
    return True

def main(folder_Resultat, TAnalyse, TData, tellepunkt):
    dir_root  = (folder_Resultat)
    xl_app = Dispatch("Excel.Application")
    xl_app.Visible = False
    xl_app.DisplayAlerts = False
    run_macro(TAnalyse, TData, xl_app, tellepunkt)
    xl_app.Quit()
    xl_app = None

def fixLenkeinfo():
    print(("Fix LenkeInfo: " + folder_Definisjon + str(Lenkefil_Start)).replace("\\", "/"))
    xl_app = Dispatch("Excel.Application")
    xl_app.Visible = False
    xl_app.DisplayAlerts = False
    wb2 = xl_app.workbooks.open(folder_Definisjon + str(Lenkefil_Start))
    wb = xl_app.workbooks.open(folder_Definisjon + "TAnalyse-Mal.xlsx")
    wb.Close(True)
    wb2.Close(True)
    wb2 = None
    wb= None
    xl_app.Quit()
    xl_app = None

def file_in_folder(Inndata):
    files = os.listdir(Inndata)
    newlist = []
    for names in files:
        if names.endswith(".xlsx"):
            newlist.append(names)
    return newlist

#Hovedkjøring
def run(tellepunkt):
    TAnalyse, TData = find_File_And_Rename(tellepunkt)
    print("TAnalyse: ", str(TAnalyse), " TData: ", str(TData))
    main(folder_Resultat, TAnalyse, TData, tellepunkt)


#Sletter alle filer og henter filer fra resultatkatalog til NovaTData til analysearket. Kopierer filer i def filen til input fil i programmet.
delete_folderFiles(folder_Inn)
delete_folderFiles(folder_Resultat)
file_names = file_in_folder(Inndata)
fixLenkeinfo()


for file_name in file_names:
    print("Fil_navn: ", file_name)
    shutil.copy(os.path.join(Inndata, file_name), folder_Inn)
    run(file_name)

now = datetime.datetime.now()
dt_string = now.strftime("%H:%M:%S")
ctypes.windll.user32.MessageBoxW(0, "Ferdig. Data ligger i koblet katalogen \nTid: " + dt_string , "NovaLink", 1)
