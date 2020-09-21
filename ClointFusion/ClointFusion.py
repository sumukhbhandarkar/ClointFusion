# Project Name: ClointFusion
# Project Description: A Python based RPA Automation Framework for Desktop GUI, Citrix, Web and basic Excel operations.
import subprocess
import os
import time
import sys
import platform

current_environment = 0
os_path=os.environ['USERPROFILE']
win_venv_scripts_folder_path = (r"{}\Envs\ClointFusion\Scripts".format(os_path))
linux_mac_venv_scripts_folder_path = r"\home\users"
win_venv_python_path = os.path.join(win_venv_scripts_folder_path, "python.exe")
linux_mac_venv_python_path = ""
env_pip_path = os.path.join(win_venv_scripts_folder_path,"pip")

print("Hi {} !".format(str(os.getlogin()).title()))

if os.path.exists("{}\\Envs\\ClointFusion\\cf_venv_activated.txt".format(os_path)) == False:
    print("Its our recommendation to dedicate a separate Python virtual environment on your system for ClointFusion. Please wait, while we create one for you...")
    subprocess.call("powershell Start-Process cmd.exe -ArgumentList '/c pip install wheel virtualenv virtualenvwrapper-win & mkvirtualenv -p 3 ClointFusion & workon ClointFusion & pip install --upgrade ClointFusion & deactivate & type nul > {}\\Envs\\ClointFusion\\cf_venv_activated.txt'".format(os_path))

    while True:
        if os.path.exists("{}\\Envs\\ClointFusion\\cf_venv_activated.txt".format(os_path)):
            break

if sys.executable.lower() == win_venv_python_path.lower():
    current_environment = 1
else:
    activate_venv = r"{}\Envs\ClointFusion\Scripts\activate_this.py".format(os_path)
    exec(open(activate_venv).read(), {'__file__': activate_venv})

list_of_required_packages = ["howdoi","seaborn","texthero","emoji","helium","kaleido", "folium", "zipcodes", "plotly", "PyAutoGUI", "PyGetWindow", "XlsxWriter" ,"PySimpleGUI", "chromedriver-autoinstaller", "gspread", "imutils", "keyboard", "joblib", "opencv-python", "python-imageseach-drov0", "openpyxl", "pandas", "pif", "pytesseract", "scikit-image", "selenium", "xlrd", "clipboard"]

#decorator to push a function to background using asyncio
def background(f):
    """
    Decorator function to push a function to background using asyncio
    """
    import asyncio
    try:
        from functools import wraps
        @wraps(f)
        def wrapped(*args, **kwargs):
            loop = asyncio.get_event_loop()
            if callable(f):
                return loop.run_in_executor(None, f, *args, **kwargs)
            else:
                raise TypeError('Task must be a callable')    
        return wrapped
    except Exception as ex:
        print("Task pushed to background = "+str(f) + str(ex))

import emoji
from pandas.core.algorithms import mode

os_name = platform.system()

def show_emoji(strInput=""):
    """
    Function which prints Emojis

    Usage: 
    print(show_emoji('thumbsup'))
    print("OK",show_emoji('thumbsup'))
    """
    if not strInput:
        return(emoji.emojize(":{}:".format(str('thumbsup').lower()),use_aliases=True,variant="emoji_type"))
    else:
        return(emoji.emojize(":{}:".format(str(strInput).lower()),use_aliases=True,variant="emoji_type"))

def load_missing_python_packages():
    """
    Installs missing python packages
    """       

    subprocess.call("powershell Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser")
    
    print("Welcome to ClointFusion, Made in India with " + show_emoji('red_heart'))

    if current_environment:
        print("Already inside ClointFusion Virtual Environment, type deactive in VSCode terminal to use Default Python Interpreter")
    else:
        print("Entering 'ClointFusion' Virtual Environment at {}".format(win_venv_python_path))

    print("Checking the required dependencies for {} OS".format(win_venv_python_path))
        
    try:
        import PySimpleGUI
    except:
        os.system("{} install --upgrade {}".format(sys.executable,'PySimpleGUI'))

    #install missing packages
    try:
        
        reqs = subprocess.check_output([sys.executable, '-m', 'pip', 'list'])
        installed_packages = [r.decode().split('==')[0] for r in reqs.split()]

        missing_packages = ' '.join(list(set(list_of_required_packages)-set(installed_packages)))

        if missing_packages:
            print("{} package(s) are missing".format(missing_packages)) 
            
            os.system("{} -m pip install --upgrade pip".format(sys.executable))
            
            os.system("{} install --upgrade {}".format(env_pip_path,missing_packages)) 
            print("Missing Dependencies Installed")
        else:
            print("All required packages are already available " + show_emoji('smile'))

    except Exception as ex:
        print("Error in load_missing_python_packages="+str(ex))

#upgrade existing packages
@background
def update_all_packages_in_cloint_fusion_virtual_environment():
    """
    Function to UPGRADE all packages related to ClointFusion. This function runs in background and is silent.
    """
    try:
        updating_required_packages= ' '.join(list(set(list_of_required_packages)))
        print("Updating existing packages in 'ClointFusion' ") 
        _ = subprocess.run("{} install --upgrade {}".format(env_pip_path,updating_required_packages),capture_output=True)
    except Exception as ex:
        print("Error in update_all_packages_cloint_fusion_virtual_environment="+str(ex))

load_missing_python_packages()
update_all_packages_in_cloint_fusion_virtual_environment()

from unicodedata import name
import pyautogui as pg
import pygetwindow as gw
import json
import time
import pandas as pd
import keyboard as kb
import logging
import PySimpleGUI as sg
import os
import xlrd
import openpyxl as op
from openpyxl import Workbook
from openpyxl import load_workbook
import datetime
import subprocess
import sys
from functools import lru_cache
import threading
from threading import Timer
import traceback
import gspread
import socket
from cv2 import cv2
import base64
import pytesseract
import signal
from skimage.measure import compare_ssim
from skimage.metrics import structural_similarity
import imutils
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
import chromedriver_autoinstaller
from selenium.webdriver.support.expected_conditions import _find_elements
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options
from python_imagesearch.imagesearch import imagesearch
from python_imagesearch.imagesearch import region_grabber
from python_imagesearch.imagesearch import imagesearcharea
import shutil
from shutil import which
from joblib import Parallel, delayed
from concurrent.futures import ThreadPoolExecutor, as_completed
import clipboard
import re
from openpyxl import load_workbook
from openpyxl.styles import Font
from pydrive2.drive import GoogleDrive
import plotly.express as px
from kaleido.scopes.plotly import PlotlyScope
import plotly.graph_objects as go
import zipcodes
import folium
from json import (load as jsonload, dump as jsondump)
from helium import *
from os import link
from selenium.webdriver import ChromeOptions
import dis
import texthero as hero
from texthero import preprocessing
import seaborn as sb
import matplotlib.pyplot as plt
from urllib.request import urlopen 

sg.theme('Dark') 

if 'Windows' in os_name:
    c_drive_base_dir = r"C:\ClointFusion\My_Bot"  
else: # to be tested in Linux Environemnt
    c_drive_base_dir = r"\home\ClointFusion\My_Bot"

img_path =  "" 
batch_file_path = ""
config_folder_path = ""
output_folder_path = ""
error_screen_shots_path = ""
folderPathToStatusLogFile = ""
current_working_dir = os.path.dirname(os.path.realpath(__file__)) 
Cloint_PNG_Logo_Path = ""
bot_name = ""
excel_operations_excel_file_1 = ""

print("ClointFusion module running at " + str(current_working_dir) + " " + show_emoji())

#Web Browser Automation Global Variables
service = ""
driver = ""
URL = ""
Chrome_Service_Started = False
cloint_logo_base64 = b'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'

ss_path_b = os.path.join(config_folder_path,"my_screen_shot_before.png") #before search
ss_path_a = os.path.join(config_folder_path,"my_screen_shot_after.png") #after search

fullPathToStatusLogFile = ""
global ai_screenshot #AiLocateImageOnScreen
ai_processes = []

global HLaunched
HLaunched=False

def timeit(method):
    """
    Decorator for computing time taken

    parameters:
        Method() name, by using @timeit just above the def: - defination of the function.

    returns:
        prints time take by the function 
    """
    def timed(*args, **kw):
        ts = time.time()
        result = method(*args, **kw)
        te = time.time()
        print('%r  %2.2f ms' % (method.__name__, (te - ts) * 1000))
        return result
    return timed

def pysimplegui_save_settings(settings_file, values):
    """
    Save Settings of GUI functions. This is an internal function called by all 4 GUI functions
    """
    try:
        if values:    
            with open(settings_file, 'w') as f:
                json.dump(eval(str(values)), f)
    except Exception as ex:
        print("Error in pysimplegui_save_settings="+str(ex))

def pysimplegui_load_settings(settings_file):
    """
    Load Settings of GUI functions. This is an internal function called by all 4 GUI functions
    """
    with open(settings_file, 'r') as f:
        return jsonload(f)

def gui_get_any_file_from_user(Extension_Without_Dot="*"):    
    """
    Generic function to accept file path from user using GUI. Returns the filepath value in string format.Default allows all files i.e *
    """
    try:

        GFFU_SETTINGS = {'-FILE-': ""}
        SETTINGS_FILE = os.path.join(config_folder_path, r'settings_gui_get_any_file_from_user.cfg')
        
        try:
            SETTINGS_KEYS_TO_ELEMENT_KEYS = pysimplegui_load_settings(SETTINGS_FILE)
        except:
            SETTINGS_KEYS_TO_ELEMENT_KEYS = GFFU_SETTINGS

        oldFilePath = SETTINGS_KEYS_TO_ELEMENT_KEYS['-FILE-']

        layout = [[sg.Text("ClointFusion - Set Yourself Free for Better Work", font='Courier 16',text_color='orange')],
                [sg.Text('Please choose the file having the extension:',text_color='yellow'),sg.Text(text="." + Extension_Without_Dot,font=('Courier 16'),text_color='red')],[sg.Input(default_text=oldFilePath, key='-FILE-', visible=True, enable_events=True), sg.FileBrowse(file_types=((".{} File".format(Extension_Without_Dot), "*.{}".format(Extension_Without_Dot)),))],
                [sg.Submit('Done',button_color=('white','green')),sg.CloseButton('Close',button_color=('white','firebrick'))]]

        window = sg.Window('ClointFusion',layout, return_keyboard_events=True,use_default_focus=True,disable_close=False,element_justification='c',keep_on_top=True, finalize=True)

        while True:
                event, values = window.read()
                if event == sg.WIN_CLOSED or event == 'Close':
                    break
                if event == 'Done':
                    if values['-FILE-']:
                        break
                    else:
                        message_pop_up("Please enter the required values")
                        print("Please enter the values")
        window.close()
        pysimplegui_save_settings(SETTINGS_FILE,values)
        return str(values['-FILE-'])

    except Exception as ex:
        print("Error in gui_get_any_file_from_user="+str(ex))
    
def excel_get_all_sheet_names(excelFilePath=""):
    """
    Gives you all names of the sheets in the given excel sheet.

    Parameters:
        excelFilePath  (str) : Full path to the excel file with slashes.
    
    returns : 
        all the names of the excelsheets as a LIST.
    """
    try:
        if not excelFilePath:
            excelFilePath = gui_get_any_file_from_user("xlsx")

        xls = xlrd.open_workbook(excelFilePath, on_demand=True)
        return xls.sheet_names()
    except Exception as ex:
        print("Error in excel_get_all_sheet_names="+str(ex))
    
def excel_get_all_header_columns(excelFilePath,sheet_name="Sheet1",header=0):
    """
    Gives you all column header names of the given excel sheet.
    """
    col_lst = []
    try:
        col_lst = pd.read_excel(excelFilePath,sheet_name=sheet_name,header=header,nrows=1,dtype=str).columns.tolist()
        return col_lst
    except Exception as ex:
        print("Error in excel_get_all_header_columns="+str(ex))
    
def gui_get_excel_sheet_header_from_user(): 
    """
    Generic function to accept excel path, sheet name and header from user using GUI. Returns all these values in disctionary format.
    """
    global excel_operations_excel_file_1
    try:
        GESHFU_SETTINGS = {'-FILEPATH-': "", '-SHEET-': "Sheet1" , '-HEADER-': '0', '-USE_THIS_EXCEL-': True}
        SETTINGS_FILE = os.path.join(config_folder_path, r'settings_gui_get_excel_sheet_header_from_user.cfg')
        sheet_namesLst = []

        try:
            SETTINGS_KEYS_TO_ELEMENT_KEYS = pysimplegui_load_settings(SETTINGS_FILE)
        except:
            SETTINGS_KEYS_TO_ELEMENT_KEYS = GESHFU_SETTINGS

        oldFilePath = SETTINGS_KEYS_TO_ELEMENT_KEYS['-FILEPATH-']
        oldSheet = SETTINGS_KEYS_TO_ELEMENT_KEYS['-SHEET-']
        oldHeader = SETTINGS_KEYS_TO_ELEMENT_KEYS['-HEADER-']
        old_Use_This_excel = SETTINGS_KEYS_TO_ELEMENT_KEYS['-USE_THIS_EXCEL-']

        layout = [[sg.Text("ClointFusion - Set Yourself Free for Better Work", font='Courier 16', text_color='orange')],
                    [sg.Text('Please choose your excel file',auto_size_text=True), sg.Input(default_text=oldFilePath,key="-FILEPATH-",enable_events=True,change_submits=True), sg.FileBrowse(file_types=(("Excel File", "*.xls"),("Excel File", "*.xlsx")))], 
                    [sg.Text('Sheet Name'), sg.Combo(sheet_namesLst,default_value=oldSheet,size=(20, 0),key="-SHEET-",enable_events=True)], 
                    [sg.Text('Choose the header row'),sg.Spin(values=('0', '1', '2', '3', '4', '5'),initial_value=oldHeader,key="-HEADER-",enable_events=True,change_submits=True)],
                    [sg.Checkbox('Use this excel file for all the excel related operations of this BOT', key='-USE_THIS_EXCEL-',default=old_Use_This_excel, text_color='yellow')],
                    [sg.Submit('Done',button_color=('white','green')),sg.CloseButton('Close',button_color=('white','firebrick'))]]
        values = []

        window = sg.Window('ClointFusion',layout, return_keyboard_events=True,use_default_focus=False,disable_close=False,element_justification='c',keep_on_top=True, finalize=True)

        while True:
            
            if oldFilePath: 
                sheet_namesLst = excel_get_all_sheet_names(oldFilePath)
                window['-SHEET-'].update(values=sheet_namesLst)   
            
            event, values = window.read()
            
            if event is None or event == 'Cancel' or event == "Escape:27":
                values = []
                break

            if event == 'Done':
                if values and values['-FILEPATH-'] and values['-SHEET-']:
                    break
                else:
                    message_pop_up("Please enter all the values")

            if values['-FILEPATH-']: 
                sheet_namesLst = excel_get_all_sheet_names(values['-FILEPATH-'])
                window['-SHEET-'].update(values=sheet_namesLst)   

                if event == values["-USE_THIS_EXCEL-"]:
                    excel_operations_excel_file_1 = values['-FILEPATH-']
                
            if values['-SHEET-'] != 'Sheet1':
                window['-SHEET-'].update(value=values['-SHEET-']) 
            elif len(sheet_namesLst) > 1:
                window['-SHEET-'].update(value=sheet_namesLst[0]) 

        window.close()
        pysimplegui_save_settings(SETTINGS_FILE,values)
        return values['-FILEPATH-'], values ['-SHEET-'], int(values['-HEADER-'])
    
    except Exception as ex:
        print("Error in gui_get_excel_sheet_header_from_user="+str(ex))
    
def gui_get_folder_path_from_user():    
    """
    Generic function to accept folder path from user using GUI. Returns the folderpath value in string format.
    """
    try:

        GFFU_SETTINGS = {'-FOLDER-': ""}
        SETTINGS_FILE = os.path.join(config_folder_path, r'settings_gui_get_folder_path_from_user.cfg')
        
        try:
            SETTINGS_KEYS_TO_ELEMENT_KEYS = pysimplegui_load_settings(SETTINGS_FILE)
        except:
            SETTINGS_KEYS_TO_ELEMENT_KEYS = GFFU_SETTINGS

        oldFolderPath = SETTINGS_KEYS_TO_ELEMENT_KEYS['-FOLDER-']

        layout = [[sg.Text("ClointFusion - Set Yourself Free for Better Work", font='Courier 16',text_color='orange')],
                [sg.Text('Please choose the folder',text_color='yellow')],[sg.Input(default_text=oldFolderPath ,key='-FOLDER-', visible=True, enable_events=True), sg.FolderBrowse()],
                [sg.Submit('Done',button_color=('white','green')),sg.CloseButton('Close',button_color=('white','firebrick'))]]

        window = sg.Window('ClointFusion',layout, return_keyboard_events=True,use_default_focus=True,disable_close=False,element_justification='c',keep_on_top=True,finalize=True)

        while True:
            event, values = window.read()
            if event == sg.WIN_CLOSED or event == 'Close':
                break
            if event == 'Done':
                if values and values['-FOLDER-']:
                    break
                else:
                    message_pop_up("Please enter the required values")
        
        window.close()
        pysimplegui_save_settings(SETTINGS_FILE,values)
        return str(values['-FOLDER-'])

    except Exception as ex:
        print("Error in gui_get_folder_path_from_user="+str(ex))
    
def gui_get_any_input_from_user(msgForUser=""):    
    """
    Generic function to accept any input (text / numeric) from user using GUI. Returns the value in string format.
    """
    
    try:
        GIFU_SETTINGS = {'-KEY-':"",'-VALUE-': ""}
        SETTINGS_FILE = os.path.join(config_folder_path, r'settings_gui_get_any_input_from_user.cfg')
        
        try:
            SETTINGS_KEYS_TO_ELEMENT_KEYS = pysimplegui_load_settings(SETTINGS_FILE)
        except:
            SETTINGS_KEYS_TO_ELEMENT_KEYS = GIFU_SETTINGS

        oldKey = SETTINGS_KEYS_TO_ELEMENT_KEYS['-KEY-']
        oldValue = SETTINGS_KEYS_TO_ELEMENT_KEYS['-VALUE-']

        # if not oldKey :
        #     oldKey = msgForUser
        if msgForUser:
            oldKey = msgForUser
            oldValue = ""
        layout = [[sg.Text("ClointFusion - Set Yourself Free for Better Work", font='Courier 16',text_color='orange')],
                [sg.Text('Please enter the '),sg.Text(text=oldKey,font=('Courier 12'),text_color='yellow'),sg.Input(default_text=oldValue,key='-VALUE-', justification='c')],
                [sg.Text('This field is mandatory',text_color='red')],
                [sg.Submit('Done',button_color=('white','green')),sg.CloseButton('Close',button_color=('white','firebrick'))]]

        window = sg.Window('ClointFusion',layout, return_keyboard_events=True,use_default_focus=True,disable_close=True,element_justification='c',keep_on_top=True,finalize=True)

        while True:
            
            event, values = window.read()
            if event == sg.WIN_CLOSED or event == 'Close':
                print(values['-VALUE-'])

                if oldValue or (values and values['-VALUE-']):
                    break

                else:
                    message_pop_up("Its a mandatory field !.. Cannot proceed, exiting now..")
                    print("Exiting ClointFusion, as Mandatory field is missing")
                    sys.exit(0)
            
            if event == 'Done':
                if values['-VALUE-']:
                    break
                else:
                    message_pop_up("This value is required. Please enter the value..")
        
        window.close()
        values['-KEY-'] = msgForUser
        pysimplegui_save_settings(SETTINGS_FILE,values)

        return str(values['-VALUE-']).strip()
    except Exception as ex:
        print("Error in gui_get_any_input_from_user="+str(ex))
    
def extract_filename_from_filepath(strFilePath=""):
    """
    Function which extracts file name from the given filepath
    """
    try:
        if strFilePath:
            strFileName = strFilePath[strFilePath.rindex("\\") + 1 : ]
            strFileName = strFileName.split(".")[0]
            return strFileName
    except Exception as ex:
        print("Error in extract_filename_from_filepath="+str(ex))
    
# @timeit
def folder_create(strFolderPath):
    """
    while making leaf directory if any intermediate-level directory is missing,
    folder_create() method will create them all.

    Parameters:
        folderPath (str) : path to the folder where the folder is to be created.

    For example consider the following path:
        
        Suppose we want to create directory ‘krishna’ but Directory ‘project’ and ‘python’ are unavailable in the path.
        Then folder_create() method will create all unavailable/missing directory in the specified path.
        ‘project’ and ‘python’ will be created first then ‘krishna’ directory will be created.
    """
    try:
        if not os.path.exists(strFolderPath):
            os.makedirs(strFolderPath)
    except Exception as ex:
        print("Error in folder_create="+str(ex))
    
def _create_status_log_file(xtLogFilePath):
    """
    Internal Function to create Status Log File
    """
    if not os.path.exists(xtLogFilePath):
        df = pd.DataFrame({'Timestamp': [], 'Status':[]})
        writer = pd.ExcelWriter(xtLogFilePath, engine='xlsxwriter')
        df.to_excel(writer, sheet_name='Sheet1', index=False)
        writer.save()

# @timeit
def _init_status_log_file():
    """
    Generates the log and saves it to the file in the given base directory. Internal function
    """
    global fullPathToStatusLogFile
    global bot_name

    try:        
        if bot_name:
            excelFileName = str(bot_name) + "-StatusLog.xlsx"
        else:
            excelFileName = "StatusLog.xlsx"

        fullPathToStatusLogFile = os.path.join(folderPathToStatusLogFile,excelFileName)
        
        folder_create(folderPathToStatusLogFile)
        
        _create_status_log_file(fullPathToStatusLogFile)

    except Exception as ex:
        print("ERROR in _init_status_log_file="+str(ex))
    
# @background
def excel_create_excel_file_in_given_folder(fullPathToTheFolder="",excelFileName="",sheet_name="Sheet1"):
    """
    Creates an excel file in the desired folder with desired filename

    Internally this uses folder_create() method to create folders if the folder/s does not exist.

    Parameters:
        fullPathToTheFolder (str) : Complete path to the folder with double slashes.
        excelFileName       (str) : File Name of the excel to be created (.xlsx extension will be added automatically.
        sheet_name           (str) : By default it will be "Sheet1".
    
    Returns:
        returns boolean TRUE if the excel file is created
    """
    
    try:
        wb = Workbook()
        ws = wb.active
        ws.title =sheet_name

        if fullPathToTheFolder:
            folder_create(fullPathToTheFolder)
        else:
            fullPathToTheFolder = gui_get_folder_path_from_user()

        if not excelFileName:
            excelFileName = gui_get_any_input_from_user("excel file name (without extension)")
        
        if ".xlsx" in excelFileName:
            excel_path = os.path.join(fullPathToTheFolder,excelFileName)
        else:
            excel_path = os.path.join(fullPathToTheFolder,excelFileName + ".xlsx")
        
        # print("Excel path="+str(excel_path))

        wb.save(filename = excel_path)
        
        return True
    except Exception as ex:
        print("Error in excel_create_excel_file_in_given_folder="+str(ex))
    
def _set_bot_name(strBotName=""):
    """
    Internal function
    If a botname is given, it will be used in the log file and in Task Scheduler
    we can also access the botname variable globally.

    Parameters :
        strBotName (str) : Name of the bot
    """
    global c_drive_base_dir
    global bot_name

    if not strBotName: #if user has not given bot_name
        bot_name = current_working_dir[current_working_dir.rindex("\\") + 1 : ] #Assumption that user has given proper folder name and so taking it as BOT name
    else:
        strBotName = string_remove_special_characters(strBotName) 
        bot_name = strBotName

    c_drive_base_dir = c_drive_base_dir + "_" + bot_name
    
def folder_create_text_file(textFolderPath="",txtFileName=""):
    """
    Creates Text file in the given path.
    Internally this uses folder_create() method to create folders if the folder/s does not exist.
    automatically adds txt extension if not given in textFilePath.

    Parameters:
        textFilePath (str) : Complete path to the folder with double slashes.
    """
    try:

        if not textFolderPath:
            textFolderPath = gui_get_folder_path_from_user()
        
        if not txtFileName:
            txtFileName = gui_get_any_input_from_user("Text File Name")
            txtFileName = txtFileName 

        if ".txt" not in txtFileName:
            txtFileName = txtFileName + ".txt"
            
        f = open(os.path.join(textFolderPath, txtFileName), 'w')
        f.close()
        print("Text file created")

    except Exception as ex:
        print("Error in folder_create_text_file="+str(ex))
    

def get_image_from_base64(imgFileName,imgBase64Str):
    """
    Coverts the given Base64 string to an image and saves in given path

    Parameters:
        imgFileName  (str) : image file name with png extension and optional path.
        imgBase64Str (str) : Base64 string for conversion.
    """
    if not os.path.exists(imgFileName) :
        try:
            img_binary = base64.decodebytes(imgBase64Str)
            with open(imgFileName,"wb") as f:
                f.write(img_binary)
        except Exception as ex:
            print("Error in get_image_from_base64="+str(ex))
        
def string_remove_special_characters(inputStr):
    """
    Removes all the special character.

    Parameters:
        inputStr  (str) : string for removing all the special character in it.

    Returns :
        outputStr (str) : returns the alphanumeric string.
    """
    outputStr = ''.join(e for e in inputStr if e.isalnum())
    return outputStr    

# ########################
_set_bot_name()
folder_create(c_drive_base_dir) 

img_path =  os.path.join(c_drive_base_dir, "Images") 
batch_file_path = os.path.join(c_drive_base_dir, "Batch_File") 
config_folder_path = os.path.join(c_drive_base_dir, "Config_Files") 
output_folder_path = os.path.join(c_drive_base_dir, "Output") 
error_screen_shots_path = os.path.join(c_drive_base_dir, "Error_Screenshots")
folderPathToStatusLogFile = os.path.join(c_drive_base_dir,"StatusLogExcel")
Cloint_PNG_Logo_Path = os.path.join(img_path,"Cloint_Logo.PNG")

folder_create(img_path)
folder_create(batch_file_path)
folder_create(config_folder_path)
folder_create(error_screen_shots_path)
folder_create(output_folder_path)
_init_status_log_file()
get_image_from_base64(Cloint_PNG_Logo_Path,cloint_logo_base64)

# ########################

def create_batch_file(application_exe_pyw_file_path):
    """
    Creates .bat file for the given application / exe or even .pyw BOT developed by you. This is required in Task Scheduler.
    """

    global batch_file_path
    try:
        application_name = application_exe_pyw_file_path[application_exe_pyw_file_path.rindex("\\")+1:]

        cmd = ""

        if "exe" in application_name:
            application_name = str(application_name).replace("exe","bat")
            cmd = "start \"\" " + '"' + application_exe_pyw_file_path + '" /popup\n'

        elif "pyw" in application_name: 
            application_name = str(application_name).replace("pyw","bat")
            cmd = "start \"\" " + '"' + sys.executable + '" ' + '"' + application_exe_pyw_file_path + '" /popup\n'

        batch_file_path = os.path.join(batch_file_path,application_name)
        if not os.path.exists(batch_file_path):
            
            f = open(batch_file_path, 'w')
            f.write("@ECHO OFF\n")
            f.write("timeout 5 > nul\n")
            f.write(cmd) 
            f.write("exit")    
            f.close()

    except Exception as ex:
        print("Error in create_batch_file="+str(ex))
    

def excel_create_file(fullPathToTheFile="",fileName="",sheet_name="Sheet1"):
    try:
        if not fullPathToTheFile:
            fullPathToTheFile = gui_get_folder_path_from_user()

        if not fileName:
            fileName = gui_get_any_input_from_user("Excel File Name (without extension)")

        if not os.path.exists(fullPathToTheFile):
            os.makedirs(fullPathToTheFile)
        if ".xlsx" not in fileName:
            fileName = fileName + ".xlsx"

        wb = Workbook()
        ws = wb.active
        ws.title =sheet_name
        fileName = os.path.join(fullPathToTheFile,fileName)
        wb.save(filename = fileName)
        print("Excel file created in the path " + fullPathToTheFile)
        return True
    except Exception as ex:
        print("Error in excel_create_file="+str(ex))
    
def folder_get_all_filenames_as_list(strFolderPath="",extension='all'):
    """
    Get all the files of the given folder in a list.

    Parameters:
        strFolderPath  (str) : Location of the folder.
        extension      (str) : extention of the file. by default all the files will be listed regarless of the extension.
    
    returns:
        allFilesOfaFolderAsLst (List) : all the file names as a list.
    """
    try:
        if not strFolderPath:
            strFolderPath = gui_get_folder_path_from_user()

        if extension == "all":
            allFilesOfaFolderAsLst = [ f for f in os.listdir(strFolderPath)]
        else:
            allFilesOfaFolderAsLst = [ f for f in os.listdir(strFolderPath) if f.endswith(extension) ]

        return allFilesOfaFolderAsLst
    except Exception as ex:
        print("Error in folder_get_all_filenames_as_list="+str(ex))
    
def folder_delete_all_files(fullPathOfTheFolder="",file_extension_without_dot="all"):  
    """
    Deletes all the files of the given folder

    Parameters:
        fullPathOfTheFolder  (str) : Location of the folder.
        extension            (str) : extention of the file. by default all the files will be deleted inside the given folder 
                                    regarless of the extension.
    returns:
        count (int) : number of files deleted.
    """ 
    file_extension_with_dot = ''
    try:
        if not fullPathOfTheFolder:
            fullPathOfTheFolder = gui_get_folder_path_from_user()

        count = 0 
        if "." not in file_extension_without_dot :
            file_extension_with_dot = "." + file_extension_without_dot

        if file_extension_with_dot.lower() == ".all":
            filelist = [ f for f in os.listdir(fullPathOfTheFolder) ]
        else:
            filelist = [ f for f in os.listdir(fullPathOfTheFolder) if f.endswith(file_extension_with_dot) ]
        print(filelist)
        for f in filelist:
            try:
                os.remove(os.path.join(fullPathOfTheFolder, f))
                count +=1 
            except:
                pass
        print(str(count) + " files deleted")
        return count
    except Exception as ex:
        print("Error in folder_delete_all_files="+str(ex)) 
        return -1
    
def message_pop_up(strMsg="",delay=3):
    """
    Specified message will popup on the screen for a specified duration of time.

    Parameters:
        strMsg  (str) : message to popup.
        delay   (int) : duration of the popup.
    """
    try:
        if not strMsg:
            strMsg = gui_get_any_input_from_user("Message")
        sg.popup_no_wait(strMsg,title='ClointFusion',auto_close_duration=delay, auto_close=True, keep_on_top=True,background_color="white",text_color="black")#,icon=cloint_ico_logo_base64)
    except Exception as ex:
        print("Error in message_pop_up="+str(ex))
    
def key_hit_enter():
    """
    Enter key will be pressed once.
    """
    time.sleep(0.5)
    kb.press_and_release('enter')
    time.sleep(0.5)

def message_flash(msg="",delay=3):
    """
    specified msg will popup for a specified duration of time with OK button.

    Parameters:
        msg     (str) : message to popup.
        delay   (int) : duration of the popup.
    """
    try:
        if not msg:
            msg = gui_get_any_input_from_user("Message")

        r = Timer(int(delay), key_hit_enter)
        r.start()
        pg.alert(text=msg, title='ClointFusion', button='OK')
    except Exception as ex:
        print("ERROR in message_flash="+str(ex))
    
def launch_any_exe_bat_application(pathOfExeFile):
    """
    Launches any exe or batch file.

    Parameters:
        pathOfExeFile  (str) : location of the file with extension.
    """
    try: 
        subprocess.Popen(pathOfExeFile)
        time.sleep(1) 
    except Exception as ex:
        print("ERROR in launch_any_exe_bat_application="+str(ex))
    
class myThread1 (threading.Thread):
    def __init__(self,err_str):
        threading.Thread.__init__(self)
        self.err_str = err_str

    def run(self):
        message_flash(self.err_str)

class myThread2 (threading.Thread):
    def __init__(self,strFilePath):
        threading.Thread.__init__(self)
        self.strFilePath = strFilePath

    def run(self):
        time.sleep(1)
        img = pg.screenshot()
        time.sleep(1)

        dt_tm= str(datetime.datetime.now())    
    
        dt_tm = dt_tm.replace(" ","_")
        dt_tm = dt_tm.replace(":","-")
        dt_tm = dt_tm.split(".")[0]
        filePath = self.strFilePath + str(dt_tm)  + ".PNG"

        img.save(str(filePath))
        
def take_error_screenshot(err_str):
    """
    Takes screenshot of an error popup parallely without waiting for the flow of the program.
    The screenshot will be saved in the log folder for reference.

    Parameters:
        err_str  (str) : exception.
    """
    global error_screen_shots_path
    try:
        thread1 = myThread1(err_str)
        thread2 = myThread2(error_screen_shots_path)

        thread1.start()
        thread2.start()

        thread1.join()
        thread2.join()
    except Exception as ex:
        print("Error in take_error_screenshot="+str(ex))
    
def update_log_excel_file(message=""):
    """
    Given message will be updated in the excel log file.

    Parameters:
        message  (str) : message to update.

    Retursn:
        returns a boolean true if updated sucessfully
    """
    global fullPathToStatusLogFile
    try:
        if not message:
            message = gui_get_any_input_from_user("Message")

        df = pd.DataFrame({'Timestamp': [datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")], 'Status':[message]})
        writer = pd.ExcelWriter(fullPathToStatusLogFile, engine='openpyxl')
        writer.book = load_workbook(fullPathToStatusLogFile)
        writer.sheets = dict((ws.title, ws) for ws in writer.book.worksheets)
    
        reader = pd.read_excel(fullPathToStatusLogFile)
        df.to_excel(writer,index=False,header=False,startrow=len(reader)+1)
        writer.close()

        print("Message updated at the row ")
        return True
    except Exception as ex:
        print("Error in update_log_excel_file="+str(ex))
        return False
    
def string_extract_only_alphabets(inputString=""):
    """
    Returns only alphabets from given input string
    """
    if not inputString:
        inputString = gui_get_any_input_from_user("Input String")

    outputStr = ''.join(e for e in inputString if e.isalpha())
    return outputStr 

def string_extract_only_numbers(inputString=""):
    """
    Returns only numbers from given input string
    """
    if not inputString:
        inputString = gui_get_any_input_from_user("Input String")

    outputStr = ''.join(e for e in inputString if e.isnumeric())
    return outputStr   

def window_show_desktop():
    """
    Minimizes all the applications and shows Desktop.
    """
    try:
        time.sleep(0.5)
        kb.press_and_release('win+d')
        time.sleep(0.5)
    except Exception as ex:
        print("Error in window_show_desktop="+str(ex))
    
def window_get_all_opened_titles():
    """
    Gives the title of all the existing (open) windows.

    Returns:
        allTitles_lst  (list) : returns all the titles of the window as list.
    """
    try:
        allTitles_lst = []
        lst = gw.getAllTitles()
        for item in lst:
            if str(item).strip() != "" and str(item).strip() not in allTitles_lst:
                allTitles_lst.append(str(item).strip())
        return allTitles_lst
    except Exception as ex:
        print("Error in window_get_all_opened_titles="+str(ex))
    
def _window_find_exact_name(windowName=""):
    """
    Gives you the exact window name you are looking for.

    Parameters:
        windowName  (str) : Name of the window to find.

    Returns:
        win (str)              : Exact window name.
        window_found (boolean) : A boolean TRUE if the window is found
    """
    win = ""
    window_found = False

    if not windowName:
        windowName = gui_get_any_input_from_user("Partial Window Name")

    try:
        lst = gw.getAllTitles()
        
        for item in lst:
            if str(item).strip():
                if str(windowName).lower() in str(item).lower():
                    win = item
                    window_found = True
                    break
        return win, window_found
    except Exception as ex:
        print("Error in _window_find_exact_name="+str(ex))
    
def window_activate_and_maximize(windowName=""):
    """
    Activates and maximizes the desired window.

    Parameters:
        windowName  (str) : Name of the window to maximize.
    """
    try:
        if not windowName:
            windowName = gui_get_any_input_from_user("Window Name")

        item,window_found = _window_find_exact_name(windowName)
        if window_found:
            windw = gw.getWindowsWithTitle(item)[0]
            windw.activate()
            time.sleep(2)
            windw.maximize()
            time.sleep(2)
        else:
            print("No window OPEN by name="+str(windowName))
    except Exception as ex:
        print("Error in window_activate_and_maximize="+str(ex))
    
def window_minimize(windowName=""):
    """
    Activates and minimizes the desired window.

    Parameters:
        windowName  (str) : Name of the window to miniimize.
    """
    try:
        if not windowName:
            windowName = gui_get_any_input_from_user("Window Name")

        item,window_found = _window_find_exact_name(windowName)
        if window_found:
            windw = gw.getWindowsWithTitle(item)[0]
            windw.minimize()
            time.sleep(1)
        else:
            print("No window available to minimize by name="+str(windowName))
    except Exception as ex:
        print("Error in window_minimize="+str(ex))
    
def window_close(windowName=""):
    """
    Close the desired window.

    Parameters:
        windowName  (str) : Name of the window to close.
    """
    try:
        if not windowName:
            windowName = gui_get_any_input_from_user("Window Name")

        item,window_found = _window_find_exact_name(windowName)
        if window_found:
            windw = gw.getWindowsWithTitle(item)[0]
            windw.close()
            time.sleep(1)
        else:
            print("No window available to close, by name="+str(windowName))
    except Exception as ex:
        print("Error in window_close="+str(ex))
    
@lru_cache(None)
def call_otsu_threshold(img_title, is_reduce_noise=False):
    """
    OpenCV internal function for OCR
    """
    # Read the image in a greyscale mode
    image = cv2.imread(img_title, 0)

    # Apply GaussianBlur to reduce image noise if it is required
    if is_reduce_noise:
        image = cv2.GaussianBlur(image, (5, 5), 0)

    # Optimal threshold value is determined automatically.otsu_threshold
    _ , image_result = cv2.threshold(
        image, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU,
    )
    
    cv2.imwrite(img_title, image_result)
    cv2.destroyAllWindows()

@lru_cache(None)
def read_image_cv2(img_path):
    """
    Saves the image in cv2 format.

    Parameters:
        img_path  (str) : location of the image.
    
    returns:
        image (cv2) : image in cv2 format will be returned.
    """
    if img_path and os.path.exists(img_path):
        try:
            image = cv2.imread(img_path)
            return image
        except Exception as ex:
            print("read_image_cv2 = "+str(ex))
        
    else:
        print("File not found="+str(img_path))

def ocr_get_coordinates(img_path=""):
    """
    Gets the coordinates for performing OCR on that specific region
    """
    try:
        if not img_path:
            img_path = gui_get_any_file_from_user(Extension_Without_Dot="PNG")

        if img_path and os.path.exists(img_path):
            frame=read_image_cv2(img_path)

            cv2.namedWindow('ClointFusion', cv2.WINDOW_FREERATIO) 
            cv2.setWindowProperty('ClointFusion', cv2.WND_PROP_FULLSCREEN, cv2.WINDOW_FULLSCREEN)
            cv2.imshow("ClointFusion", frame)

            while True:
                key=cv2.waitKey(1) & 0xFF
                if key == ord("s"):
                    initBB = cv2.selectROI("ClointFusion", frame, fromCenter=False,showCrosshair=True)
                    print(initBB)
                    
                elif key == ord("q"):
                    break
        else:
            print("File not found="+str(img_path))

    except Exception as ex:
        print("Error in ocr_get_coordinates="+str(ex))
    
# @timeit
def ocr_magic(img_path,x,y,w,h): #CROP specific part
    """
    Coverts the given co-ordinates /bounds of an image to OCR text. Capture the bounds using ocr_get_coordinates 

    Parameters:
        img_path  (str) : Location of the image.
        x,y,w,h   (int) : the bounds value of the croped area for ocr.
    
    Returns:
        data      (str) : the OCR processed string.
    """
    try:
        r = (x, y, w, h) 
        image = read_image_cv2(img_path)
        img_cropped = image[int(r[1]):int(r[1]+r[3]), int(r[0]):int(r[0]+r[2])]

        crop_path = str(x) + "_" + str(y) + "_" + str(w) + "_" + str(h)
        img_path_cropped = str(img_path).replace(".","_" + crop_path + ".")
        
        cv2.imwrite(img_path_cropped, img_cropped)

        call_otsu_threshold(img_path_cropped)

        image = read_image_cv2(img_path_cropped)
        
        data = pytesseract.image_to_string(image, lang='eng',config='--oem 3 --psm 7')
        return data
    except Exception as ex:
        print("Error in ocr_magic="+str(ex))
    
def excel_get_row_column_count(excel_path="", sheet_name="Sheet1", header=0):
    """
    Gets the row and coloumn count of the provided excel sheet.

    Parameters:
        excel_path  (str) : Full path to the excel file with slashes.
        sheet_name           (str) : by default it is Sheet1.

    Returns:
        row (int) : number of rows
        col (int) : number of coloumns
    """
    try:
        if not excel_path:
            excel_path, sheet_name, header = gui_get_excel_sheet_header_from_user()
            
        df = pd.read_excel(excel_path, sheet_name=sheet_name, header=header)
        row, col = df.shape
        row = row + 1
        return row, col
    except Exception as ex:
        print("Error in excel_get_row_column_count="+str(ex))
    
def excel_copy_range_from_sheet(excel_path="",*, sheet_name='Sheet1', startCol=1, startRow=1, endCol=1, endRow=1):
    """
    Copies the specific range from the provided excel sheet and returns copied data as a list
    Parameters:
        excel_path :"Full path of the excel file with double slashes"
        sheet_name     :"Source sheet name from where contents are to be copied"
        startCol          :"Starting column number (index starts from 1) from where copying starts"
        startRow          :"Starting row number (index starts from 1) from where copying starts"
        endCol            :"Ending column number ex:4 upto where cells to be copied"
        endRow            :"Ending column number ex:5 upto where cells to be copied"

    Returns:
    rangeSelected        : the copied range data
    """
    try:
        if not excel_path:
            excel_path, sheet_name, header = gui_get_excel_sheet_header_from_user()
            
        from_wb = load_workbook(filename = excel_path)
        try:
            fromSheet = from_wb[sheet_name]
        except:
            fromSheet = from_wb.worksheets[0]
        rangeSelected = []

        if endRow < startRow:
            endRow = startRow

        #Loops through selected Rows
        for i in range(startRow,endRow + 1,1):
            #Appends the row to a RowSelected list
            rowSelected = []
            for j in range(startCol,endCol+1,1):
                rowSelected.append(fromSheet.cell(row = i, column = j).value)
            #Adds the RowSelected List and nests inside the rangeSelected
            rangeSelected.append(rowSelected)
    
        return rangeSelected
    except Exception as ex:
        print("Error in copy_range_from_excel_sheet="+str(ex))
    
def excel_paste_range_to_sheet(excel_path="",*, sheet_name='Sheet1', startCol=1, startRow=1, endCol=1, endRow=1, copiedData):
    """
    Pastes the copied data in specific range of the given excel sheet.
    """
    try:
        try:
            if not excel_path:
                excel_path, sheet_name, header = gui_get_excel_sheet_header_from_user()
                
            to_wb = load_workbook(filename = excel_path)
            toSheet = to_wb[sheet_name]
        except:
            excel_create_excel_file_in_given_folder((str(excel_path[:(str(excel_path).rindex("\\"))])),(str(excel_path[str(excel_path).rindex("\\")+1:excel_path.find(".")])),sheet_name)
            to_wb = load_workbook(filename = excel_path)
            toSheet = to_wb[sheet_name]

        if endRow < startRow:
            endRow = startRow

        countRow = 0
        for i in range(startRow,endRow+1,1):
            countCol = 0
            for j in range(startCol,endCol+1,1):
                
                toSheet.cell(row = i, column = j).value = copiedData[countRow][countCol]
                countCol += 1
            countRow += 1
        to_wb.save(excel_path)
        return countRow-1
    except Exception as ex:
        print("Error in excel_paste_range_to_sheet="+str(ex))
    
def _excel_copy_range(startCol=1, startRow=1, endCol=1, endRow=1, sheet='Sheet1'):
    """
    Copies the specific range from the given excel sheet.
    """
    try:
        rangeSelected = []
        #Loops through selected Rows
        for k in range(startRow,endRow + 1,1):
            #Appends the row to a RowSelected list
            rowSelected = []
            for l in range(startCol,endCol+1,1):
                rowSelected.append(sheet.cell(row = k, column = l).value)
            #Adds the RowSelected List and nests inside the rangeSelected
            rangeSelected.append(rowSelected)

        return rangeSelected

    except Exception as ex:
        print("Error in _excel_copy_range="+str(ex))
    
def _excel_paste_range(startCol=1, startRow=1, endCol=1, endRow=1, sheetReceiving='Sheet1',copiedData=[]):
    """
    Pastes the specific range to the given excel sheet.
    """
    try:
        countRow = 0
        for k in range(startRow,endRow+1,1):
            countCol = 0
            for l in range(startCol,endCol+1,1):
                sheetReceiving.cell(row = k, column = l).value = copiedData[countRow][countCol]
                countCol += 1
            countRow += 1
        return countRow

    except Exception as ex:
        print("Error in _excel_paste_range="+str(ex))
    
def excel_split_by_column(excel_path="",*,sheet_name='Sheet1',header=0,columnName=""):
    """
    Splits the excel file by Column Name
    """
    try:
        if not excel_path:
            excel_path, sheet_name, header = gui_get_excel_sheet_header_from_user()
            
        if not columnName:
            columnName = gui_get_any_input_from_user("Column Name")

        data_df = pd.read_excel(excel_path,sheet_name=sheet_name,header=header,dtype=str)
        grouped_df = data_df.groupby(columnName)

        for data in  grouped_df:  
            grouped_df.get_group(data[0]).to_excel(excel_path[:excel_path.rindex("\\")] + "\\" + str(data[0]) + ".xlsx", index=False)

    except Exception as ex:
        print("Error in excel_split_by_column="+str(ex))
    
def excel_split_the_file_on_row_count(excel_path="",*, sheet_name = 'Sheet1', rowSplitLimit="", outputFolderPath="", outputTemplateFileName ="Split"):
    """
    Splits the excel file as per given row limit
    """
    try:
        if not excel_path:
            excel_path, sheet_name, _ = gui_get_excel_sheet_header_from_user()
            
        if not rowSplitLimit:
            rowSplitLimit = int(gui_get_any_input_from_user("Row Split Count/Limit Ex: 20"))

        if not outputFolderPath:
            outputFolderPath = gui_get_folder_path_from_user()

        src_wb = op.load_workbook(excel_path)
        src_ws = src_wb.worksheets[0] 

        src_ws_max_rows = src_ws.max_row
        src_ws_max_cols= src_ws.max_column 

        i = 1
        start_row = 2

        while start_row <= src_ws_max_rows:
            
            dest_wb = Workbook()
            dest_ws = dest_wb.active
            dest_ws.title = sheet_name

            #Copy ROW-1 (Header) from SOURCE to Each DESTINATION file
            selectedRange = _excel_copy_range(1,1,src_ws_max_cols,1,src_ws) #startCol, startRow, endCol, endRow, sheet
            _ =_excel_paste_range(1,1,src_ws_max_cols,1,dest_ws,selectedRange) #startCol, startRow, endCol, endRow, sheetReceiving,copiedData
            
            selectedRange = ""
            selectedRange = _excel_copy_range(1,start_row,src_ws_max_cols,start_row + rowSplitLimit - 1,src_ws) #startCol, startRow, endCol, endRow, sheet   
            _ =_excel_paste_range(1,2,src_ws_max_cols,rowSplitLimit + 1,dest_ws,selectedRange) #startCol, startRow, endCol, endRow, sheetReceiving,copiedData

            start_row = start_row + rowSplitLimit
            dest_file_name = outputFolderPath + "\\" + outputTemplateFileName + "-" + str(i) + ".xlsx"
            dest_wb.save(dest_file_name)
            print("Created " +  dest_file_name)
            i = i + 1
        return True
    except Exception as ex:
        print("Error in excel_split_the_file_on_row_count="+str(ex))
    
def excel_merge_all_files(fullPathOfTheFolder="",outputFolderPath=""):
    """
    Merges all the excel files in the given folder
    """
    try:
        if not fullPathOfTheFolder:
            fullPathOfTheFolder = gui_get_folder_path_from_user()

        if not outputFolderPath:
            outputFolderPath = gui_get_folder_path_from_user()
        
        filelist = [ f for f in os.listdir(fullPathOfTheFolder) if f.endswith(".xlsx") ]
        all_excel_file_lst = []
        for file1 in filelist:
            file_path = os.path.join(fullPathOfTheFolder,file1)
            print(file_path)
            all_excel_file = pd.read_excel(file_path,dtype=str)
            all_excel_file_lst.append(all_excel_file)

        appended_df = pd.concat(all_excel_file_lst)
        time_stamp_now=datetime.datetime.now().strftime("%m-%d-%Y")
        final_path= outputFolderPath + "\\Final-" + time_stamp_now + ".xlsx"
        appended_df.to_excel(final_path, index=False)
        print("Files Merged....")
        return True
    except Exception as ex:
        print("Error in excel_merge_all_files="+str(ex))
    
def excel_drop_columns(txtFilePath,columnsToBeDropped):
    """
    Drops the desired column from the given excel file
    """
    try:
        df = pd.read_excel(txtFilePath) 

        if isinstance(columnsToBeDropped, list):
            df.drop(columnsToBeDropped, axis = 1, inplace = True) 
        else:
            df.drop([columnsToBeDropped], axis = 1, inplace = True) 

        df.to_excel(txtFilePath,index=False)
    except Exception as ex:
        print("Error in excel_drop_columns="+str(ex))
    
def excel_sort_columns(excel_path="",*,sheet_name='Sheet1',header=0,firstColumnToBeSorted=None,secondColumnToBeSorted=None,thirdColumnToBeSorted=None,firstColumnSortType=True,secondColumnSortType=True,thirdColumnSortType=True):
    """
    A function which takes excel full path to excel and column names on which sort is to be performed

    """
    try:
        if not excel_path:
            excel_path, sheet_name, header = gui_get_excel_sheet_header_from_user()
            
        df=pd.read_excel(excel_path,sheet_name=sheet_name, header=header)
        if thirdColumnToBeSorted is not None and secondColumnToBeSorted is not None and firstColumnToBeSorted is not None:
            df=df.sort_values([firstColumnToBeSorted,secondColumnToBeSorted,thirdColumnToBeSorted],ascending=[firstColumnSortType,secondColumnSortType,thirdColumnSortType])
        
        elif secondColumnToBeSorted is not None and firstColumnToBeSorted is not None:
            df=df.sort_values([firstColumnToBeSorted,secondColumnToBeSorted],ascending=[firstColumnSortType,secondColumnSortType])
        
        elif firstColumnToBeSorted is not None:
            df=df.sort_values([firstColumnToBeSorted],ascending=[firstColumnSortType])

        df.to_excel(excel_path,index=False)
        print("Sorted")
        # return True
    except Exception as ex:
        print("Error in excel_sort_columns="+str(ex))
    
def excel_clear_sheet(excel_path="",sheet_name="Sheet1", header=0):
    """
    Clears the contents of given excel files keeping header row intact
    """
    try:
        if not excel_path:
            excel_path, sheet_name, header = gui_get_excel_sheet_header_from_user()

        df = pd.read_excel(excel_path,sheet_name=sheet_name,header=header) 
        df = df.head(0)

        with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a') as writer:
            df.to_excel(writer, index=False, sheet_name=sheet_name)

        # writer = pd.ExcelWriter(excel_path, engine='openpyxl')
        # writer.book = load_workbook(excel_path)
        # writer.sheets = dict((ws.title, ws) for ws in writer.book.worksheets)
    
        # # reader = pd.read_excel(excel_path)
        # df.to_excel(writer, index=False,startrow=0)
        writer.close()

    except Exception as ex:
        print("Error in excel_clear_sheet="+str(ex))
    
def excel_set_single_cell(excel_path="", *, sheet_name="Sheet1", header=0, columnName="", cellNumber=0, setText=""): 
    """
    Writes the given text to the desired column/cell number for the given excel file
    """
    try:
        if not excel_path:
            excel_path, sheet_name, header = gui_get_excel_sheet_header_from_user()
            
        if not columnName:
            columnName = gui_get_any_input_from_user("Column Name")

        if not setText:
            setText = gui_get_any_input_from_user("Text Value")

        df = pd.read_excel(excel_path,sheet_name=sheet_name,header=header)
        Filepath_writer = pd.ExcelWriter(excel_path)
        df.at[cellNumber,columnName] = setText
        df.to_excel(Filepath_writer, sheet_name=sheet_name ,index=False)    
        Filepath_writer.save()
        return True

    except Exception as ex:
        print("Error in excel_set_single_cell="+str(ex))
    
def excel_get_single_cell(excel_path="",*,sheet_name="Sheet1",header=0, columnNames="",cellNumber=0): 
    """
    Gets the text from the desired column/cell number of the given excel file
    """
    try:
        if not excel_path:
            excel_path, sheet_name, header = gui_get_excel_sheet_header_from_user()
            
        if not columnNames:
            columnNames = gui_get_any_input_from_user("Column Name")

        df = pd.read_excel(excel_path,sheet_name=sheet_name,header=header,usecols={columnNames})
        cellValue = df.at[cellNumber,columnNames]
        return cellValue
    except Exception as ex:
        print("Error in excel_get_single_cell="+str(ex))
    
def excel_remove_duplicates(excel_path="", *, sheet_name="Sheet1", header=0, columnName="", saveResultsInSameExcel=True, which_one_to_keep="first"): 
    """
    Drops the duplicates from the desired Column of the given excel file
    """
    try:
        if not excel_path:
            excel_path, sheet_name, header = gui_get_excel_sheet_header_from_user()
            
        if not columnName:
            columnName = gui_get_any_input_from_user("Column Name")

        df = pd.read_excel(excel_path, sheet_name=sheet_name,header=header) 
        count = 0 
        if saveResultsInSameExcel:
            df.drop_duplicates(subset=columnName, keep=which_one_to_keep, inplace=True)
            df.to_excel(excel_path, index=False)
            count = df.shape[0]
        else:
            df1 = df.drop_duplicates(subset=columnName, keep=which_one_to_keep, inplace=False)
            excel_path = str(excel_path).replace(".","_DupDropped.")
            df1.to_excel(excel_path, index=False)
            count = df1.shape[0]

        print(str(count) + " rows affected")
        return count
    except Exception as ex:
        print("Error in excel_remove_duplicates="+str(ex))
    
def excel_vlook_up(filepath_1="", *, sheet_name_1 = 'Sheet1', header_1 = 0, filepath_2="", sheet_name_2 = 'Sheet1', header_2 = 0, Output_path="", OutputExcelFileName="", match_column_name="",how='left'):
    """
    Performs excel_vlook_up on the given excel files for the desired columns. Possible values for how are "inner","left", "right", "outer"
    """
    try:
        if not filepath_1:
            filepath_1, sheet_name_1, header_1 = gui_get_excel_sheet_header_from_user()
             
        if not filepath_2:
            filepath_2, sheet_name_2, header_2 = gui_get_excel_sheet_header_from_user()
            
        if not match_column_name:
            match_column_name = gui_get_any_input_from_user("Column Name To Be Matched")

        df1 = pd.read_excel(filepath_1, sheet_name = sheet_name_1, header = header_1)
        df2 = pd.read_excel(filepath_2, sheet_name = sheet_name_2, header = header_2)

        df = pd.merge(df1, df2, on= match_column_name, how = how)

        output_file_path = ""
        if str(OutputExcelFileName).endswith(".*"):
            OutputExcelFileName = OutputExcelFileName.split(".")[0]
        
        if Output_path and OutputExcelFileName:
            if ".xlsx" in OutputExcelFileName:
                output_file_path = os.path.join(Output_path, OutputExcelFileName)
            else:
                output_file_path = os.path.join(Output_path, OutputExcelFileName  + ".xlsx")

        else:
            output_file_path = filepath_1

        # with pd.ExcelWriter(output_file_path, engine='openpyxl', mode='a') as writer:
        #     df.to_excel(writer, index=False)

        df.to_excel(output_file_path, index = False)
        print("excel_vlook_up Done")
        return True
    
    except Exception as ex:
        print("Error in excel_vlook_up="+str(ex))
    
def screen_clear_search(delay=0.2):
    """
    Clears previously found text (crtl+f highlight)
    """
    try:
        kb.press_and_release("ctrl+f")
        time.sleep(delay)
        pg.typewrite("^%#")
        time.sleep(delay)
        kb.press_and_release("esc")
        time.sleep(delay)
    except Exception as ex:
        print("Error in screen_clear_search="+str(ex))
    
def scrape_save_contents_to_notepad(folderPathToSaveTheNotepad=""): #"Full path to the folder (with double slashes) where notepad is to be stored"
    """
    Copy pastes all the available text on the screen to notepad and saves it.
    """
    try:
        if not folderPathToSaveTheNotepad:
            folderPathToSaveTheNotepad = gui_get_folder_path_from_user()

        time.sleep(1)
        screen_size=pg.size()
        pg.click(screen_size[0]/2,screen_size[1]/2)
        time.sleep(0.5)

        kb.press_and_release("ctrl+a")
        time.sleep(1)
        kb.press_and_release("ctrl+c")
        time.sleep(1)
        
        clipboard_data = clipboard.paste()
        time.sleep(2)
        
        screen_clear_search()

        if str(folderPathToSaveTheNotepad).endswith("\\"):
            notepad_file_path = folderPathToSaveTheNotepad + "notepad-contents.txt"
        else:
            notepad_file_path = folderPathToSaveTheNotepad + "\\notepad-contents.txt"

        f = open(notepad_file_path, "w")
        f.write(clipboard_data)
        time.sleep(10)
        f.close()

        return "Saved the contents at " + notepad_file_path
    except Exception as ex:
        print("Error in scrape_SaveContentsToNotepad = "+str(ex))
    
def scrape_get_contents_by_search_copy_paste(highlightText=""):
    """
    Gets the focus on the screen by searching given text using crtl+f and performs copy/paste of all data. Useful in Citrix applications
    """
    output_lst_newline_removed = []
    try:
        if not highlightText:
            highlightText = gui_get_any_input_from_user("Text to be Searched")

        time.sleep(1)
        kb.press_and_release("ctrl+f")
        time.sleep(1)
        pg.typewrite(highlightText)
        time.sleep(1)
        kb.press_and_release("enter")
        time.sleep(1)
        kb.press_and_release("esc")
        time.sleep(2)

        pg.PAUSE = 2
        kb.press_and_release("ctrl+a")
        time.sleep(2)
        kb.press_and_release("ctrl+c")
        time.sleep(2)
        
        clipboard_data = clipboard.paste()
        time.sleep(2)
        
        screen_clear_search()

        entire_data_as_list= clipboard_data.splitlines()
        for line in entire_data_as_list:
            if line.strip():
                output_lst_newline_removed.append(line.strip())

        return output_lst_newline_removed
    except Exception as ex:
        print("Error in scrape_get_contents_by_search_copy_paste="+str(ex))
    
def mouse_move(x="",y=""):
    """
    Moves the cursor to the given X Y Co-ordinates.
    """
    try:
        if not x and not y:
            x_y = str(gui_get_any_input_from_user("Values ex: 200,215"))
            if "," in x_y:
                x, y = x_y.split(",")
                x = int(x)
                y = int(y)
            else:
                message_pop_up("Please enter like 200,300")

        if x and y:
            time.sleep(0.2)
            pg.moveTo(x,y)
            time.sleep(0.2)
    except Exception as ex:
        print("Error in mouse_move="+str(ex))
    
def mouse_get_color_by_position(pos=[]):
    """
    Gets the color by X Y co-ordinates of the screen.
    """
    try:
        if not pos:
            pos = gui_get_any_input_from_user("Values ex: 200,215")

        im = pg.screenshot()
        return im.getpixel((pos[0],pos[1]))
    except Exception as ex:
        print("Error in mouse_get_color_by_position = "+str(ex))
    
def mouse_click(x="", y="", left_or_right="left", single_double_triple="single", copyToClipBoard_Yes_No="no"):
    """
    Clicks at the given X Y Co-ordinates on the screen using ingle / double / tripple click(s).
    Optionally copies selected data to clipboard (works for double / triple clicks)
    """
    try:
        if not x and not y:
            x_y = str(gui_get_any_input_from_user("Values ex: 200,215"))
            if "," in x_y:
                x, y = x_y.split(",")
                x = int(x)
                y = int(y)
            else:
                message_pop_up("Please enter like 200,300")

        copiedText = ""
        time.sleep(1)

        if x and y:
            if single_double_triple.lower() == "single" and left_or_right.lower() == "left":
                pg.click(x,y)
            elif single_double_triple.lower() == "double" and left_or_right.lower() == "left":
                pg.doubleClick(x,y)
            elif single_double_triple.lower() == "triple" and left_or_right.lower() == "left":
                pg.tripleClick(x,y)
            elif single_double_triple.lower() == "single" and left_or_right.lower() == "right":
                pg.rightClick(x,y)
            time.sleep(1)    

            if copyToClipBoard_Yes_No.lower() == "yes":
                kb.press_and_release("ctrl+c")
                time.sleep(1)
                copiedText = clipboard.paste().strip()
                time.sleep(1)
                
            time.sleep(1)    
            return copiedText
    except Exception as ex:
        print("Error in mouseClick="+str(ex))
    
def mouse_drag_from_to(X1="",Y1="",X2="",Y2="",delay=0.5):
    """
    Clicks and drags from X1 Y1 co-ordinates to X2 Y2 Co-ordinates on the screen
    """
    try:

        x_y = str(gui_get_any_input_from_user("FROM Values ex: 200,215"))
        if "," in x_y:
            X1, Y1 = x_y.split(",")
            X1 = int(X1)
            Y1 = int(Y1)

        x_y = str(gui_get_any_input_from_user("TO Values ex: 200,215"))
        if "," in x_y:
            X2, Y2 = x_y.split(",")
            X2 = int(X2)
            Y2 = int(Y2)

        time.sleep(0.2)
        pg.moveTo(X1,Y1,duration=delay)
        pg.dragTo(X2,Y2,duration=delay,button='left')
        time.sleep(0.2)
    except Exception as ex:
        print("Error in mouse_drag_from_to="+str(ex))
    
def search_highlight_tab_enter_open(searchText="",hitEnterKey="Yes"):
    """
    Searches for a text on screen using crtl+f and hits enter.
    """
    try:
        if not searchText:
            searchText = gui_get_any_input_from_user("Search Text")

        time.sleep(0.5)
        kb.press_and_release("ctrl+f")
        time.sleep(0.5)
        kb.write(searchText)
        time.sleep(0.5)
        kb.press_and_release("enter")
        time.sleep(0.5)
        kb.press_and_release("esc")
        time.sleep(0.2)
        if hitEnterKey.lower() == "yes":
            kb.press_and_release("tab")
            time.sleep(0.3)
            kb.press_and_release("shift+tab")
            time.sleep(0.3)
            kb.press_and_release("enter")
            time.sleep(2)
        return True

    except Exception as ex:
        print("Error in search_highlight_tab_enter_open="+str(ex))
    
def key_press(strKeys=""):
    """
    Emulates the given keystrokes.
    """
    try:
        if not strKeys:
            strKeys = gui_get_any_input_from_user("Keys")

        strKeys = strKeys.lower()
        if "shift" in strKeys:
            strKeys = strKeys.replace("shift","left shift+right shift")

        time.sleep(0.5)
        kb.press_and_release(strKeys)
        time.sleep(0.5)
    except Exception as ex:
        print("Error in key_press="+str(ex))
    
def key_write_enter(strMsg="",delay=1,key="e"):
    """
    Writes/Types the given text and press enter (by default) or tab key.
    """
    try:
        if not strMsg:
            strMsg = gui_get_any_input_from_user("Message")

        time.sleep(0.2)
        kb.write(strMsg)
        time.sleep(delay)
        if key.lower() == "e":
            key_press('enter')
        elif key.lower() == "t":
            key_press('tab')
        time.sleep(1)
    except Exception as ex:
        print("Error in key_write_enter="+str(ex))
    
def date_convert_to_US_format(input_str):
    """
    Converts the given date to US date format.
    """
    try:
        match = re.search(r'\d{4}-\d{2}-\d{2}', input_str) #1
        if match == None:
            match = re.search(r'\d{2}-\d{2}-\d{4}', input_str) #2
            if match == None:
                match = re.search(r'\d{2}/\d{2}/\d{4}', input_str) #3
                if match == None:
                    match = re.search(r'\d{4}/\d{2}/\d{2}', input_str) #4
                    if match == None:
                        match = re.findall(r'(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s\d\d,\s\d{4}',input_str) #5
                        dt=datetime.datetime.strptime(match[0], '%b %d, %Y').date() #5 Jan 01, 2020
                    else:    
                        dt=datetime.datetime.strptime(match.group(), '%Y/%m/%d').date() #4
                else:
                    try:
                        dt=datetime.datetime.strptime(match.group(),'%d/%m/%Y').date() #3
                    except:
                        dt=datetime.datetime.strptime(match.group(),'%m/%d/%Y').date() #3
            else:
                try:
                    dt=datetime.datetime.strptime(match.group(), '%d-%m-%Y').date()#2
                except:
                    dt=datetime.datetime.strptime(match.group(), '%m-%d-%Y').date()#2
        else:
            dt=datetime.datetime.strptime(match.group(), '%Y-%m-%d').date() #1
        return dt.strftime('%m/%d/%Y')    
    except Exception as ex:
        print("Error in date_convert_to_US_format="+str(ex))
    
def mouse_search_snip_return_coordinates_x_y(img="", conf=0.9, wait=180,region=(0,0,1366,768)): #180
    """
    Searches the given image on the screen and returns its center of X Y co-ordinates.
    """
    try:
        if not img:
            img = gui_get_any_file_from_user(".PNG")

        time.sleep(1)
        pos = pg.locateOnScreen(img,confidence=conf,region=region) 
        i = 0
        while pos == None and i < int(wait):
            pos = ()
            pos = pg.locateOnScreen(img, confidence=conf,region=region)   
            time.sleep(1)
            i = i + 1

        time.sleep(1)

        if pos:
            x,y = pos.left + int(pos.width / 2), pos.top + int(pos.height / 2)
            pos = ()
            pos=(x,y)
            
            return pos
        return pos
    except Exception as ex:
        print("Error in mouse_search_snip_return_coordinates_x_y="+str(ex))
    
def is_text_found_on_screen(searchText="",successImg="Full path to image",conf=0.8, delay=15):
    """
    Finds if a text is available on screen.
    """
    try:
        if not searchText:
            searchText = gui_get_any_input_from_user("Search Text")

        time.sleep(1)
        kb.press_and_release("ctrl+f")
        time.sleep(1)
        kb.write(searchText)
        time.sleep(1)
        kb.press_and_release("enter")
        time.sleep(1)

        pos = mouse_search_snip_return_coordinates_x_y(successImg,conf,delay) 
        found = False

        if pos != None:
            found = True
            
        kb.press_and_release("esc")
        time.sleep(1)
        return found
    except Exception as ex:
        print("Error in is_text_found_on_screen="+str(ex))
    
def find_text_on_screen(searchText="",delay=0.1, occurance=1,isSearchToBeCleared=False):
    """
    Clears previous search and finds the provided text on screen.
    """
    screen_clear_search() #default

    if not searchText:
        searchText = gui_get_any_input_from_user("Search Text")

    time.sleep(delay)
    kb.press_and_release("ctrl+f")
    time.sleep(delay)
    pg.typewrite(searchText)
    time.sleep(delay)

    for i in range(occurance-1):
        kb.press_and_release("enter")
        time.sleep(delay)

    kb.press_and_release("esc")
    time.sleep(delay)

    if isSearchToBeCleared:
        screen_clear_search()

def mouse_search_snip_return_coordinates_box(img, conf=0.9, wait=180,region=(0,0,pg.size()[0],pg.size()[1])):
    """
    Searches the given image on the screen and returns the 4 bounds co-ordinates (x,y,w,h)
    """
    try:
        time.sleep(1)
        
        pos = pg.locateOnScreen(img,confidence=conf,region=region) 
        i = 0
        while pos == None and i < int(wait):
            pos = ()
            pos = pg.locateOnScreen(img, confidence=conf,region=region)   
            time.sleep(1)
            i = i + 1
        time.sleep(1)
        return pos

    except Exception as ex:
        print("Error in mouse_search_snip_return_coordinates_box="+str(ex))
    
def mouse_find_highlight_click(searchText,delay=0.1,occurance=1):
    """
    Searches the given text on the screen, highlights and clicks it.
    """  
    try:
        if not searchText:
            searchText = gui_get_any_input_from_user("Search Text")

        time.sleep(0.2)

        find_text_on_screen(searchText,delay=delay,occurance=occurance,isSearchToBeCleared = True) #clear the search

        img = pg.screenshot()
        img.save(ss_path_b)
        time.sleep(0.2)
        imageA = cv2.imread(ss_path_b)
        time.sleep(0.2)

        find_text_on_screen(searchText,delay=delay,occurance=occurance,isSearchToBeCleared = False) #dont clear the searched text

        img = pg.screenshot()
        img.save(ss_path_a)
        time.sleep(0.2)
        imageB = cv2.imread(ss_path_a)
        time.sleep(0.2)

        # convert both images to grayscale
        grayA = cv2.cvtColor(imageA, cv2.COLOR_BGR2GRAY)
        grayB = cv2.cvtColor(imageB, cv2.COLOR_BGR2GRAY)

        # compute the Structural Similarity Index (SSIM) between the two
        (_, diff) = structural_similarity(grayA, grayB, full=True)
        diff = (diff * 255).astype("uint8")

        thresh = cv2.threshold(diff, 0, 255,
            cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU)[1]
        cnts = cv2.findContours(thresh.copy(), cv2.RETR_EXTERNAL,
            cv2.CHAIN_APPROX_SIMPLE)
        cnts = imutils.grab_contours(cnts)

        # loop over the contours
        for c in cnts:
            (x, y, w, h) = cv2.boundingRect(c)
            
            X = int(x + (w/2))
            Y = int(y + (h/2))
            
            pg.click(X, Y)
            time.sleep(0.5)
            break

    except Exception as ex:
        print("Error in mouse_find_highlight_click="+str(ex))
    
#Web Browser Automation Functions
def _start_chrome_service(dp=False,dn=True,igc=True,smcp=True,i=False,h=False,whtsp=False, profile_path=""):        #pass boolean True or False
    """
    Starts the chrome driver service and launches the chrome browser with specified arguments.
    """
    global driver
    global service
    try:
        options = webdriver.ChromeOptions()
        chromedriver_autoinstaller.install()
        chrome_driver_path = which(chromedriver_autoinstaller.utils.get_chromedriver_filename())
        service = Service(chrome_driver_path)
        service.start()
        
        if dp:
            options.add_argument("--disable-popup-blocking")                #False
        if dn:  
            options.add_argument("--disable-notifications")                 #True
        if igc:
            options.add_argument("--ignore-certificate-errors")             #True
        if smcp:
            options.add_argument("--suppress-message-center-popups")        #True
        if i:
            options.add_argument("--incognito")                             #False
        if h:
            options.add_argument("--headless")                              #False
            options.add_argument("--disable-gpu")
        
        if whtsp :
            options.add_argument("--user-data-dir={}".format(profile_path))
            options.add_argument('--profile-directory=Profile 108')        
        options.add_argument("--disable-translate")
        options.add_argument("--start-maximized")                           #True
        options.add_argument("--ignore-autocomplete-off-autofill")          #True
        options.add_argument("--no-first-run")                              #True
        #options.add_argument("--window-size=1920,1080")
        
        driver = webdriver.Remote(service.service_url,options=options)
    except Exception as ex:
        print("Error in _start_chrome_service = "+str(ex))
        driver.quit
    
def browser_title_s(bd="t"):
    """
    Gives you the current browser tilte or handle.
    """
    try:
        if bd.lower() == "t" :              #Get browser Title
            return driver.title
        if bd.lower() == "wh" :             #Get Browser Handles
            return driver.window_handles
    except Exception as ex:
        print("Error in browser_title_s = "+str(ex))
        return None
    
def browser_page_source_html_s():
    """
    Gets the complete html page source of the given page as string.
    """
    try:
        return driver.page_source           #Get whole HTML Source in the page as List
    except Exception as ex:
        print("Error in browser_page_source_html_s = "+str(ex))
    
def browser_page_source_text_s():
    """
    Gets complete text of the given page as list.
    """
    try:                                     #Get whole text in the page as List
        text_items = []
        data = driver.find_elements_by_tag_name('html')
        for item in data:
            text_items.append(item.text)
        return text_items
    except Exception as ex:
        print("Error in browser_page_source_text_s = "+str(ex))
    
def browser_navigate_s(nav):
    """
    Navigates to given url or goes forward, backward or refresh according to the specified argument.
    """
    try:
        if len(nav) >5 :
            driver.get(nav)                 #Navigate to URL
            return
        if nav.lower() == "b" :             #Navigate Back
            driver.back()
            return
        if nav.lower() == "f" :             #Navigate Forward 
            driver.forward()
            return
        if nav.lower() == "r" :             #Refresh Page
            driver.refresh()
            return
    except Exception as ex:
        print("Error in browser_navigate_s = "+str(ex))
        driver.quit()
    
def browser_locate_element_s(element,xpath="xpath"):   
    """
    Locates the XPATH element on screen and returns the element.
    """
    try:                                                                 #returns a single element
        if xpath.lower() == "xpath":                                            
            element1 = driver.find_element_by_xpath(element)                #find by Xpath
            return element1 
        if xpath.lower() == "href":
            element1 = driver.find_element_by_link_text(element)            #find by link text href
            return element1
        if xpath.lower() == "css":                                          #find by css selector
            element1 = driver.find_element_by_css_selector(element)
            return element1
    except Exception as ex:
        print("Error in browser_locate_element_s = "+str(ex))
    
def browser_locate_elements_s(element,xpath="xpath"):                            #returns multiple elements as a list
    """
    Locates the XPATH elements on screen and return the elements.
    """
    try:
        if xpath.lower() == "xpath":
            elements = driver.find_elements_by_xpath(element)               
            return elements
        if xpath.lower() == "href":
            elements = driver.find_element_by_link_text(element)            #find by link text href
            return elements
        if xpath.lower() == "css":                                          #find by css selector
            elements = driver.find_element_by_css_selector(element)
            return elements
    except Exception as ex:
        print("Error in browser_locate_elements_s = "+str(ex))
    
def browser_mouse_click_s(element,i="c"):   
    """
    Clicks on given web element using XPath
    """
    try:                                                                #Click or Send enter key to the element
        if i == "c":
            element.click()                                                 
            return
        if i == "e":
            element.send_keys(Keys.ENTER)                                   
            return
    except Exception as ex:
        print("Error in browser_mouse_click_s = "+str(ex))
    
def browser_write_s(element,write):
    try:
        element.send_keys(write)                                            #write text in text feilds
        return
    except Exception as ex:
        print("Error in browser_write_s = "+str(ex))
    
def browser_wait_s():
    try:
        driver.implicitly_wait(60)
        return
    except Exception as ex:
        print("Error in browser_wait_s = "+str(ex))
    
def launch_website_s(URL1="",dp=False,dn=True,igc=True,smcp=True,i=False,h=False,whtsp=False, profile_path=""):
    """
    Starts the chrome service, opens the browser and launches your website URL. This is the first function to be called for any web-related BOT Development
    """
    global Chrome_Service_Started

    try:
        if not URL1:
            URL1 = gui_get_any_input_from_user("Website URL like https://www.google.com")

        if not Chrome_Service_Started: 
            _start_chrome_service(dp=dp,dn=dn,igc=igc,smcp=smcp,i=i,h=h,whtsp=whtsp, profile_path=profile_path)
            browser_wait_s()
            Chrome_Service_Started = True

        browser_navigate_s(URL1)
        
        browser_wait_s()
        print("Launched " + browser_title_s())
        
    except Exception as ex:
        print("Error in launch_website_s="+str(ex))
        driver.quit()
    
def _search_image(img,confidence):  
    """
    Internal Function
    """
    try:
        w,h = pg.size()  
        im = region_grabber((0, 0, w, h))
        pos = imagesearcharea(img, 0, 0, w, h, confidence, im)
        if pos[0] > 0 and pos[1] > 0 :
            return pos
    except Exception as ex:
        print("Errror in _search_image="+str(ex))
    
#searches multiple images and returns list of tuple for all images found
def search_multiple_images_in_parallel(img_lst, confidence=0.9):
    """
    Returns the postion of all the images passed as list
    """
    try:
        if len(img_lst) > 0 :
            results = Parallel(n_jobs=10)(delayed(_search_image)(img,confidence) for img in img_lst)
            return results
    except Exception as ex:
        print("Errror in search_multiple_images_in_parallel="+str(ex))
    
def _locate_image(snip_url, ai_screenshot,confidence):
    """
    Internal function
    """
    try:
        cordinates = pg.locate(snip_url, ai_screenshot, confidence=confidence)
        if cordinates is None:
            return None
        return cordinates
    except Exception as ex:
        print("Error in _locate_image="+str(ex))
    
def _predict_ai_coordinates():
    """
    Internal function
    """
    try:
        x,y,c,m,n = 0,0,0,0,0
        _,_,w,h = 0,0,0,0
        global ai_processes
        for task in as_completed(ai_processes):
            if task.result() is not None:
                if m == 0:
                    m,n = pg.center(task.result())
                    l,t,w,h = task.result()
                a,b = pg.center(task.result())
                if ((m+(w//4)+1) >= a and (m-(w//4)-1) <= a) and ((n+(h//4)+1) >= b and (n-(h//4)-1) <= b):
                    pass
                else:
                    return "Multiple images detected, use AiLocateMultipleImageOnScreen() function","Multiple images detected, use AiLocateMultipleImageOnScreen() function"
                c += 1
                x += a
                y += b
        if (x>0) and (y>0):
            ai_x = x//c
            ai_y = y//c
            return ai_x,ai_y
        else:
            return None,None

    except Exception as ex:
        print("Error in _predict_ai_coordinates="+str(ex))
    
def ai_screenshot():
    try:
        global ai_screenshot
        ai_screenshot = pg.screenshot()
        return
    except Exception as ex:
        print("Error in ai_screenshot="+str(ex))
    
def _multitreading_locateimage(ai_snip_list,confidence=.8):    
    """
    internal function
    """
    try:
        with ThreadPoolExecutor(max_workers=50) as executor:
            for snip in ai_snip_list:
                snip_url = snip
                ai_processes.append(executor.submit(_locate_image,snip_url,ai_screenshot,confidence=confidence))
    except Exception as ex:
        print("Error in _multitreading_locateimage="+str(ex))
    
def mouse_ai_locate_snip_on_screen(ai_snip_list, confidence=.8):
    try:
        ai_screenshot()
        _multitreading_locateimage(ai_snip_list=ai_snip_list,confidence=confidence)
        ai_x,ai_y = _predict_ai_coordinates() 
        return ai_x,ai_y
    except Exception as ex:
        print("Error in mouse_ai_locate_snip_on_screen="+str(ex))
    
def mouse_ai_locate_multiple_images_on_screen(ai_snip_list,confidence=.8,click=False):
    try:
        ai_multiple_x_y = []
        ai_screenshot()
        _multitreading_locateimage(ai_snip_list=ai_snip_list,confidence=confidence)
        for task in as_completed(ai_processes):
            if task.result() is not None:
                ai_multiple_x_y.append(pg.center(task.result()))
        if ai_multiple_x_y == []:
            ai_multiple_x_y = None
        print(ai_multiple_x_y)
        if ai_multiple_x_y and click==True:
            for click in ai_multiple_x_y:
                time.sleep(0.5)
                pg.click(click)
        return ai_multiple_x_y
    except Exception as ex:
        print("Error in mouse_ai_locate_multiple_images_on_screen="+str(ex))
    
def schedule_create_task(Weekly_Daily="D",*,week_day="Sun",start_time_hh_mm_24_hr_frmt="11:00"):
    """
    Schedules (weekly & daily options as of now) the current BOT (.bat) using Windows Task Scheduler. Please call create_batch_file() function before using this function to convert .pyw file to .bat
    """
    try:
        global bot_name

        bot_name = string_remove_special_characters(bot_name)

        str_cmd = ""

        if Weekly_Daily == "D":
            str_cmd = r"powershell.exe Start-Process schtasks '/create  /SC DAILY /tn ClointFusion\{} /tr {} /st {}' ".format(bot_name,batch_file_path,start_time_hh_mm_24_hr_frmt)
        elif Weekly_Daily == "W":
            str_cmd = r"powershell.exe Start-Process schtasks '/create  /SC WEEKLY /D {} /tn ClointFusion\{} /tr {} /st {}' ".format(week_day,bot_name,batch_file_path,start_time_hh_mm_24_hr_frmt)

        subprocess.call(str_cmd)
        print("Task Scheduled")
    except Exception as ex:
        print("Error in schedule_create_task="+str(ex))

def schedule_delete_task():
    """
    Deletes already scheduled task. Asks user to supply task_name used during scheduling the task. You can also perform this action from Windows Task Scheduler.
    """
    try:
        str_cmd = r"powershell.exe Start-Process schtasks '/delete /tn ClointFusion\{} ' ".format(bot_name)
        
        subprocess.call(str_cmd)
        print("Task {} Deleted".format(bot_name))

    except Exception as ex:
        print("Error in schedule_delete_task="+str(ex))

@lru_cache(None)
def _get_tabular_data_from_website(Website_URL):
    """
    internal function
    """
    all_tables = ""
    try:
        all_tables = pd.read_html(Website_URL)
    except Exception as ex:
        print("Error in _get_tabular_data_from_website="+str(ex))
    
def browser_get_html_tabular_data_from_website(Website_URL="",table_index=-1,drop_first_row=False,drop_first_few_rows=[0],drop_last_row=False):
    """
    Web Scrape HTML Tables : Gets Website Table Data Easily as an Excel using Pandas. Just pass the URL of Website having HTML Tables.
    If there are 5 tables on that HTML page and you want 4th table, pass table_index as 3

    Ex: browser_get_html_tabular_data_from_website(Website_URL=URL)
    """
    try:
        if not Website_URL:
            Website_URL= gui_get_any_input_from_user("Website URL ex: https://www.google.com")

        all_tables = _get_tabular_data_from_website(Website_URL)

        if all_tables:
            
            # if no table_index is specified, then get all tables in output
            if table_index == -1:
                strFileName = Website_URL[Website_URL.rindex("/")+1:] + "_All_Tables" +  ".xlsx"
                excel_create_excel_file_in_given_folder(output_folder_path,strFileName)
            else:
                strFileName = Website_URL[Website_URL.rindex("/")+1:] + "_" + str(table_index) +  ".xlsx"

            strFileName = os.path.join(output_folder_path,strFileName)
            
            if table_index == -1:
                for i in range(len(all_tables)):
                    table = all_tables[i] #lool thru table_index values

                    table = table.reset_index(drop=True) #Avoid multi index error in our dataframes

                    with pd.ExcelWriter(strFileName, engine='openpyxl', mode='a') as writer:
                        table.to_excel(writer, sheet_name=str(i)) #index=False
            else:
                table = all_tables[table_index] #get required table_index
                
                if drop_first_row:
                    table = table.drop(drop_first_few_rows) # Drop first few rows (passed as list)

                if drop_last_row:
                    table = table.drop(len(table)-1) # Drop last row

            # table.columns = list(table.iloc[0])
            # table = table.drop(len(drop_first_few_rows)) 

                table = table.reset_index(drop=True) 

                table.to_excel(strFileName, index=False)

            print("Table saved as Excel at {} ".format(strFileName))

        else:
            print("No tables found in given website " + str(Website_URL))

    except Exception as ex:
        print("Error in browser_get_html_tabular_data_from_website="+str(ex))

def excel_charts(excel_path="",sheet_name='Sheet1', header=0, x_col="", y_col="", color="", chart_type='bar', title='ClointFusion', show_chart=False):

    """
    Interactive data visualization function, which accepts excel file, X & Y column. 
    Chart types accepted are bar , scatter , pie , sun , histogram , box  , strip. 
    You can pass color column as well, having a boolean value.
    Image gets saved as .PNG in the same path as excel file.

    Usage: excel_charts(<excel path>,x_col='Name',y_col='Age', chart_type='bar',show_chart=True)
    """
    try:
        if not excel_path:
            excel_path, sheet_name, header = gui_get_excel_sheet_header_from_user()
            
        if not x_col:
            x_col = gui_get_any_input_from_user("X Axis Column")

        if not y_col:
            y_col = gui_get_any_input_from_user("Y Axis Column")

        if x_col and y_col:
            if color:
                df = pd.read_excel(excel_path,sheet_name=sheet_name,header=header,usecols={x_col,y_col,color})
            else:
                df = pd.read_excel(excel_path,sheet_name=sheet_name,header=header,usecols={x_col,y_col})

            fig = go.Figure()

            if chart_type == 'bar':

                fig.add_trace(go.Bar(x=df[x_col].values.tolist()))
                fig.add_trace(go.Bar(y=df[y_col].values.tolist()))

                if color:
                    fig = px.bar(df, x=x_col, y=y_col, barmode="group",color=color)
                else:
                    fig = px.bar(df, x=x_col, y=y_col, barmode="group")
                    
            elif chart_type == 'scatter':

                fig.add_trace(go.Scatter(x=df[x_col].values.tolist()))
                fig.add_trace(go.Scatter(y=df[x_col].values.tolist()))

            elif chart_type =='pie':

                if color:
                    fig = px.pie(df, names=x_col, values=y_col, title=title,color=color)#,hover_data=df.columns)
                else:
                    fig = px.pie(df, names=x_col, values=y_col, title=title)#,hover_data=df.columns)

            elif chart_type =='sun':

                if color:
                    fig = px.sunburst(df, path=[x_col], values=y_col,hover_data=df.columns,color=color)
                else:
                    fig = px.sunburst(df, path=[x_col], values=y_col,hover_data=df.columns)

            elif chart_type == 'histogram':

                if color:
                    fig = px.histogram(df, x=x_col, y=y_col, marginal="rug",color=color, hover_data=df.columns)
                else:
                    fig = px.histogram(df, x=x_col, y=y_col, marginal="rug",hover_data=df.columns)

            elif chart_type == 'box':

                if color:
                    fig = px.box(df, x=x_col, y=y_col, notched=True,color=color)
                else:
                    fig = px.box(df, x=x_col, y=y_col, notched=True)

            elif chart_type == 'strip':

                if color:
                    fig = px.strip(df, x=x_col, y=y_col, orientation="h",color=color)
                else:
                    fig = px.strip(df, x=x_col, y=y_col, orientation="h")

            fig.update_layout(title = title)
            
            if show_chart:
                fig.show()
            
            # strFileName = (excel_path[excel_path.rindex("\\")+1:]).split(".")[0] + ".PNG"
            strFileName = excel_path.replace(".xlsx",".PNG")
            
            scope = PlotlyScope()
            with open(strFileName, "wb") as f:
                f.write(scope.transform(fig, format="png"))
            print("Chart saved at " + strFileName)
        else:
            print("Please supply all the required values")

    except Exception as ex:
        print("Error in excel_charts=" + str(ex))
    
def get_long_lat(strZipCode=""):
    """
    Function takes zip_code as input and returns longitude, latitude, state, city, county. 
    """

    try:
        if not strZipCode:
            strZipCode = gui_get_any_input_from_user("USA Zip Code ex: 77429")

        all_data_dict=zipcodes.matching(strZipCode)
        all_data_dict = all_data_dict[0]

        long = all_data_dict['long']
        lat = all_data_dict['lat']
        state = all_data_dict['state']
        city = all_data_dict['city']
        county = all_data_dict['county']
        return long, lat, state, city, county    
    except Exception as ex:
        print("Error in get_long_lat="+str(ex))

def excel_geotag_using_zipcodes(excel_path="",sheet_name='Sheet1',header=0,zoom_start=5,zip_code_column="ZIP CODE",data_columns_as_list=[],color_boolean_column=""):
    """
    Function takes Excel file having ZipCode column as input. Takes one data column at present. 
    Creates .html file having geo-tagged markers/baloons on the page.

    Ex: excel_geotag_using_zipcodes(excel_path)
    """

    try:
        if not excel_path:
            excel_path, sheet_name, header = gui_get_excel_sheet_header_from_user()

        m = folium.Map(location=[40.178877,-100.914253 ], zoom_start=zoom_start)

        if len(data_columns_as_list) == 1:
            data_columns_as_str = str(data_columns_as_list).replace("[","").replace("]","").replace("'","")
        else:
            data_columns_as_str = str(data_columns_as_list).replace("[","").replace("]","")
            data_columns_as_str = data_columns_as_str[1:-1]
            
        use_cols = data_columns_as_list
        use_cols.append(zip_code_column)

        if color_boolean_column:
            use_cols.append(color_boolean_column)

        df = pd.read_excel(excel_path,sheet_name=sheet_name,header=header,usecols=use_cols)
        
        for _, row in df.iterrows():
            if not pd.isna(row[zip_code_column]) and str(row[zip_code_column]).isnumeric():
                
                long, lat, state, city, county = get_long_lat(str(row[zip_code_column]))
                county = str(county).replace("County","")
                
                if color_boolean_column and data_columns_as_str and row[color_boolean_column] == True:
                    folium.Marker(location=[lat, long], popup='State: ' + state + ',\nCity:' + city + ',\nCounty:' + county + ',\nDevice:' + row[data_columns_as_str], icon=folium.Icon(color='green', icon='info-sign')).add_to(m)
                elif data_columns_as_str:
                    folium.Marker(location=[lat, long], popup='State: ' + state + ',\nCity:' + city + ',\nCounty:' + county + ',\nDevice:' + row[data_columns_as_str], icon=folium.Icon(color='red', icon='info-sign')).add_to(m)
                else:
                    folium.Marker(location=[lat, long], popup='State: ' + state + ',\nCity:' + city + ',\nCounty:' + county, icon=folium.Icon(color='blue', icon='info-sign')).add_to(m)

        graphFileName = excel_path.replace(".xlsx",".html")
        print("GeoTagged Graph saved at "+ graphFileName)
        m.save(graphFileName)
    
    except Exception as ex:
        print("Error in excel_geotag_using_zipcodes="+str(ex))
    
def _accept_cookies_h():
    """
    Internal function to accept cookies.
    """
    try:
        if Text('Accept cookies?').exists():
            click('I accept')
    except Exception as ex:
        print("Error in _accept_cookies_h="+str(ex))
    
def launch_website_h(URL,dp=False,dn=True,igc=True,smcp=True,i=False,headless=False):
    try:
        """
        Internal function to launch browser.
        """
        global HLaunched
        HLaunched=True
        options = ChromeOptions()
        if dp:
            options.add_argument("--disable-popup-blocking")                
        if dn:  
            options.add_argument("--disable-notifications")                
        if igc:
            options.add_argument("--ignore-certificate-errors")             
        if smcp:
            options.add_argument("--suppress-message-center-popups")       
        if i:
            options.add_argument("--incognito")                             
        
        options.add_argument("--disable-translate")
        options.add_argument("--start-maximized")                          
        options.add_argument("--ignore-autocomplete-off-autofill")          
        options.add_argument("--no-first-run")                             
        #options.add_argument("--window-size=1920,1080")
        start_chrome(url=URL,options=options,headless=headless)
        Config.implicit_wait_secs = 30
        _accept_cookies_h()
    except Exception as ex:
        print("Error in launch_website_h = "+str(ex))
        kill_browser()
    
def browser_navigate_h(url="",dp=False,dn=True,igc=True,smcp=True,i=False,headless=False):
    try:
        """
        Navigates to Specified URL.
        """
        if not url:
            url = gui_get_any_file_from_user("Website URL ex: https://www.google.com")

        global HLaunched
        if not HLaunched:
            launch_website_h(URL=url,dp=dp,dn=dn,igc=igc,smcp=smcp,i=i,headless=headless)
            return
        go_to(url.lower())
        _accept_cookies_h()
    except Exception as ex:
        print("Error in browser_navigate_h = "+str(ex))
    
def browser_write_h(Value="",User_Visible_Text_Element="",alert=False):
    """
    Write a string on the given element.
    """
    try:
        if not alert:
            if Value and User_Visible_Text_Element:
                write(Value, into=User_Visible_Text_Element)
        if alert:
            if Value and User_Visible_Text_Element:
                write(Value, into=Alert(User_Visible_Text_Element))
    except Exception as ex:
        print("Error in browser_write_h = "+str(ex))
    
def browser_key_press_h(key):
    """
    keyboard simulation.
    """
    try:
        press(key)
    except Exception as ex:
        print("Error in browser_key_press_h="+str(ex))
    
def browser_mouse_click_h(User_Visible_Text_Element="",element="d"):
    """
    click on the given element.
    """
    try:
        if not User_Visible_Text_Element:
            User_Visible_Text_Element = gui_get_any_input_from_user("Visible Text Element")

        if User_Visible_Text_Element and element.lower()=="d":      #default
            click(User_Visible_Text_Element)
        elif User_Visible_Text_Element and element.lower()=="l":    #link
            click(link(User_Visible_Text_Element))
        elif User_Visible_Text_Element and element.lower()=="b":    #button
            click(Button(User_Visible_Text_Element))
        elif User_Visible_Text_Element and element.lower()=="t":    #textfeild
            click(TextField(User_Visible_Text_Element))
        elif User_Visible_Text_Element and element.lower()=="c":    #checkbox
            click(CheckBox(User_Visible_Text_Element))
        elif User_Visible_Text_Element and element.lower()=="r":    #radiobutton
            click(RadioButton(User_Visible_Text_Element))
        elif User_Visible_Text_Element and element.lower()=="i":    #image ALT Text
            click(Image(alt=User_Visible_Text_Element))
    except Exception as ex:
        print("Error in browser_mouse_click_h = "+str(ex))
    
        logging.info("browser_mouse_click_h")

def browser_mouse_double_click_h(User_Visible_Text_Element=""):
    """
    Doubleclick on the given element.
    """
    try:
        if not User_Visible_Text_Element:
            User_Visible_Text_Element = gui_get_any_input_from_user("Visible Text Element")

        if User_Visible_Text_Element:
            doubleclick(User_Visible_Text_Element)
    except Exception as ex:
        print("Error in browser_mouse_double_click_h = "+str(ex))
    
        logging.info("browser_mouse_double_click_h")

def browser_locate_element_h(element,value=True):
    """
    Find the element by Xpath, id or css selection.
    """
    try:
        if value:
            return S(element).value
        return S(element)
    except Exception as ex:
        print("Error in browser_locate_element_h = "+str(ex))
    
def browser_locate_elements_h(element,value=True):
    """
    Find the elements by Xpath, id or css selection.
    """
    try:
        if value:
            return find_all(S(element).value)
        return find_all(S(element))
    except Exception as ex:
        print("Error in browser_locate_elements_h = "+str(ex))
    
def browser_wait_until_h(text="",element="t"):
    """
    Wait until a specific element is found.
    """
    try:
        if not text:
            text = gui_get_any_input_from_user("Text Element To Search & Wait For")

        if element.lower()=="t":
            wait_until(Text(text).exists,10)        #text
        elif element.lower()=="b":
            wait_until(Button(text).exists,10)      #button
    except Exception as ex:
        print("Error in browser_wait_until_h = "+str(ex))
    
def browser_mouse_click_xy_h(XYTuple):
    """
    Click on the given X Y Co-ordinates.
    """
    try:
        click(XYTuple)
    except Exception as ex:
        print("Error in browser_mouse_click_xy_h = "+str(ex))
    
def browser_refresh_page_h():
    """
    Refresh the page.
    """
    try:
        refresh()
    except Exception as ex:
        print("Error in browser_refresh_page_h = "+str(ex))
    
def browser_quit_h():
    """
    Close the browser.
    """
    try:
        kill_browser()
    except Exception as ex:
        print("Error in browser_quit_h = "+str(ex))
    

#Utility Functions
def dismantle_code(strFunctionName):
    """
    This functions dis-assembles given function and shows you column-by-column summary to explain the output of disassembled bytecode.
    Ex: dismantle_code(show_emoji)
    """
    try:
        if strFunctionName:
            print("Code dismantling {}".format(strFunctionName))
            return dis.dis(strFunctionName) 
    except Exception as ex:
       print("Error in dismantle_code="+str(ex)) 

def excel_clean_data(excel_path="",sheet_name='Sheet1',header=0,column_to_be_cleaned="",cleaning_pipe_line="Default"):
    """
    fillna(s) Replace not assigned values with empty spaces.
    lowercase(s) Lowercase all text.
    remove_digits() Remove all blocks of digits.
    remove_punctuation() Remove all string.punctuation (!"#$%&'()*+,-./:;<=>?@[\]^_`{|}~).
    remove_diacritics() Remove all accents from strings.
    remove_stopwords() Remove all stop words.
    remove_whitespace() Remove all white space between words.
    """
    try:
        if not excel_path:
            excel_path, sheet_name, header = gui_get_excel_sheet_header_from_user()
            
        if not column_to_be_cleaned:
            column_to_be_cleaned = gui_get_any_input_from_user("Column Name")

        if column_to_be_cleaned:
            df = pd.read_excel(excel_path,sheet_name=sheet_name,header=header)

            new_column_name = "Clean_" + column_to_be_cleaned

            if 'Default' in cleaning_pipe_line:
                df[new_column_name] = df[column_to_be_cleaned].pipe(hero.clean)
            else:
                custom_pipeline = [preprocessing.fillna, preprocessing.lowercase]
                df[new_column_name] = df[column_to_be_cleaned].pipe(hero.clean,custom_pipeline)    

            with pd.ExcelWriter(path=excel_path, mode='a',engine='openpyxl') as writer:
                df.to_excel(writer,index=False)

            print("Data Cleaned. Please see the output in {}".format(new_column_name))
    except Exception as ex:
        print("Error in excel_clean_data="+str(ex))
    
def excel_charts_numerical_pair_plot(excel_path="", sheet_name="", header=0, usecols=""):
    """
    Funtion for statistical data visualization.
    """
    try:
        if not excel_path:
            excel_path, sheet_name, header = gui_get_excel_sheet_header_from_user()

        if not usecols:
            usecols = gui_get_any_input_from_user("Columns To Be Used Ex: A:K")

        df = pd.read_excel(excel_path, sheet_name=sheet_name, header=header, usecols=usecols)
        
        # df = sb.load_dataset('iris')
        g = sb.PairGrid(df)
        g.map(plt.scatter)
        g.map_diag(plt.hist)
        g.map_offdiag(plt.scatter)
        
        mng = plt.get_current_fig_manager()
        mng.full_screen_toggle()
        plt.show()

        strFileName = excel_path.replace(".xlsx",".PNG")
        plt.savefig(strFileName)
    
        print("Chart saved at " + strFileName)

    except Exception as ex:
        print("Error in excel_charts_numerical_pair_plot="+str(ex))

def excel_charts_correlation_heatmap(excel_path="", sheet_name="", header=0, usecols=""):
    """
    Function for co-relation between columns of the given excel
    """
    try:
        if not excel_path:
            excel_path, sheet_name, header = gui_get_excel_sheet_header_from_user()

        if not usecols:
            usecols = gui_get_any_input_from_user("Columns To Be Used Ex: A:K")

        df = pd.read_excel(excel_path, sheet_name=sheet_name, header=header, usecols=usecols)

        # print(df.corr())
        sb.heatmap(df.corr(), annot = True, cmap = 'viridis')
        
        strFileName = excel_path.replace(".xlsx",".PNG")
        
        plt.savefig(strFileName)
        print("Chart saved at " + strFileName)

        mng = plt.get_current_fig_manager()
        mng.full_screen_toggle()
        plt.show()
    except Exception as ex:
        print("Error in excel_charts_correlation_heatmap="+str(ex))
    
def help_me(user_query=""):
    """
    Function which gives instant coding answers

    Usage:

    help_me("convert mp4 to animated gif")
    help_me("create tar archive")
    help_me("howdoi")
    help_me("formatting date")
    """
    try:
        if not user_query:
            user_query = gui_get_any_input_from_user("Query Ex: formatting date")

        print("Output from instant coding answers:")
        os.system("howdoi {}".format(user_query))
        print(show_emoji('point_up'))
    except Exception as ex:
        print("Error in help_me="+str(ex))

def browser_get_header_source_code(URL=""):
    """
    Function to get Header and Source Code for the given URL
    """
    try:
        if not URL:
            URL = gui_get_any_input_from_user("Website URL Ex: https://www.cloint.com")

        page = urlopen(URL) 
        print("\nPage Headers\n")
        print(page.headers) 
        
        content = page.read() 

        file_name = string_remove_special_characters(URL) + ".html"
        file_path = os.path.join(output_folder_path,file_name)

        with open(file_path,"wb") as f:
            f.write(content)
            
        print("Source code saved at "+ str(file_path))
    except Exception as ex:
        print("Error in browser_get_header_source_code="+str(ex))