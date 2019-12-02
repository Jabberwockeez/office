# encoding: utf-8

import settings
import xlwings as xw
import win32com
from win32com.client import constants as c
import itertools as itr
import os
import warnings



# fileName = settings.GetRecentFile(settings.sourcePath)

app = xw.App(add_book=False, visible=settings.isVisible)
app.display_alerts = False

for file in os.listdir(settings.sourcePath):

    if not file.endswith('.xls'):
        continue
    
    print(file)
    fileName = os.path.join(settings.sourcePath, file)

    wb = app.books.open(fileName)

    length = 12  # 密码长度

    # words_before = [['A', 'B']]
    # words_after = [[chr(i) for i in range(32, 127)]]
    # lst = words_before * (length - 1) + words_after

    # r = itr.product(*lst)
    # with open(os.path.join(settings.mainPath, 'pws_12.txt'), 'w') as f:
    #     for i in r:
    #         f.write(''.join(i))

    pwsList = []

    """ 检查工作簿保护 """

    guess = None

    if not wb.api.ProtectStructure:
        with open(os.path.join(settings.path, 'log.txt'), 'a') as f:
            f.write('{} has no protection.\n'.format(wb.name))
        
    with open(os.path.join(settings.mainPath, 'pws_12.txt'), 'r') as f:
        while wb.api.ProtectStructure:
            guess = f.read(length)
            if not guess:
                break
            try:
                wb.api.Unprotect(guess)
            except:
                pass

        with open(os.path.join(settings.path, 'log.txt'), 'a') as f:
            if guess and not wb.api.ProtectStructure:
                f.write('{} Password is: {}\n'.format(wb.name, guess))
                pwsList.append(guess)
            elif wb.api.ProtectStructure:
                f.write('{} Crack failed.\n'.format(wb.name))
        

    """ 检查工作表保护 """

    for ws in wb.sheets:

        ws.api.Visible = 1

        guess = None

        if not (ws.api.ProtectScenarios or ws.api.ProtectContents):
            with open(os.path.join(settings.path, 'log.txt'), 'a') as f:
                f.write('{} has no protection.\n'.format(ws.name))
            continue

        with open(os.path.join(settings.mainPath, 'pws_12.txt'), 'r') as f:
            for pw in pwsList:
                guess = pw
                try:
                    ws.api.Unprotect(guess)
                except:
                    pass
                if not ws.api.ProtectScenarios:
                    break

            while ws.api.ProtectScenarios or ws.api.ProtectContents:
                guess = f.read(length)
                if not guess:
                    break
                try:
                    ws.api.Unprotect(guess)
                except:
                    pass


        with open(os.path.join(settings.path, 'log.txt'), 'a') as f:
            if guess and not ws.api.ProtectScenarios and not ws.api.ProtectContents:
                print('{} Password is: {}'.format(ws.name, guess))
                f.write('{} Password is: {}\n'.format(ws.name, guess))
                pwsList.append(guess)
                ws.cells.rows(1).api.EntireColumn.Hidden = 0
                ws.cells.columns(1).api.EntireRow.Hidden = 0
            else:
                print('{} Crack failed.'.format(ws.name))
                f.write('{} Crack failed.\n'.format(ws.name))

    wb.api.SaveAs(os.path.join(settings.reportPath, os.path.split(fileName)[-1].replace('.xls', '_破解.xlsx')), c.xlOpenXMLWorkbook)
    wb.close()
    
app.kill()
