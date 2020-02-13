#!/usr/bin/env python3
# coding: utf-8

import gspread
from oauth2client.service_account import ServiceAccountCredentials
import openpyxl
import subprocess
import os

os.environ['http_proxy'] = "http://*****:*****@proxy.tsc.ts:8080"
os.environ['https_proxy'] = "https://*****:*****@proxy.tsc.ts:8080"
os.environ['ftp_proxy'] = "ftp://*****:*****@proxy.tsc.ts:8080"

google_table_name = "Секретный гуглодок УСБС"
google_sheet_name = "Реестр сервисов проверка скрипта"

TO_UPDATE = False
scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]

try:
    p = subprocess.check_output(['svn status -u -v Реестр\ ТСК/Реестр\ сервисов\ УСБС\ ТСК.xlsx'], universal_newlines=True, shell=True)
    print(p)
    if '*' in p.split():
        TO_UPDATE = True
except Exception as e:
    print(e)
    #subprocess.call(['svn checkout "https://bcvm370.tsc.ts/repos/stuff/OracleDepartment/USBS/ProjectLibrary/13.AD. Административные документы/Реестр сервисов/Реестр ТСК/" --depth "files"'], shell=True)
    #print("Making svn checkout")
    #TO_UPDATE = True

if TO_UPDATE:
    creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)
    client = gspread.authorize(creds)
    sh = client.open(google_table_name)
    sheets = sh.worksheets()
    for s in sheets:
        if s.title == "Реестр сервисов проверка скрипта":
            sheet = s
    try:
        data = sheet.get_all_values()
    except Exception as e:
        print(e)


    subprocess.call(['svn update "Реестр ТСК"'], shell=True)
    print("Updating googlesheet")
    wb = openpyxl.load_workbook(filename = 'Реестр ТСК/Реестр сервисов УСБС ТСК.xlsx')
    local_sheet = wb.worksheets[0]

    columns = ['F','G','H','I','J','K','L','M','N','AO','AQ','AR']
    loc_nrows = local_sheet.max_row
    loc_ncols = len(columns)
    glob_nrows = len(data)
    if (glob_nrows == 0):
        glob_ncols = 0
    else:
        glob_ncols = len(data[0])

    range_build = 'A1:' + chr((ord('A') + loc_ncols-1)) + str(loc_nrows)
    cell_list = sheet.range(range_build)
    for rownum in range(loc_nrows):
        for colnum, col in enumerate(columns):
            cell_list[rownum*(loc_ncols)+colnum].value = local_sheet[col+str(rownum+1)].value

    if (glob_ncols > loc_ncols and glob_nrows > loc_nrows):
        range_build = chr((ord('A') + loc_ncols)) + '1:' + chr((ord('A') + glob_ncols)) + str(glob_nrows)
        cell_list_vert = sheet.range(range_build)
        for i in range(len(cell_list_vert)):
            cell_list_vert[i].value = ''
    else:
        cell_list_vert = []

    if (glob_ncols > loc_ncols and glob_nrows > loc_nrows):
        range_build = 'A' + str(loc_nrows) + ':' + chr((ord('A') + loc_ncols)) + str(glob_nrows)
        cell_list_hor = sheet.range(range_build)
        for i in range(len(cell_list_hor)):
            cell_list_hor[i].value = ''
    else:
        cell_list_hor = []

    sheet.update_cells(cell_list+cell_list_hor+cell_list_vert)
else:
    print("Already up-to-date")


