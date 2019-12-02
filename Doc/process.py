# coding: utf-8
import settings
import json
# from pandas.io.json import json_normalize
from itertools import product
import os


colorRed = (255, 0, 0)

def RGB(red, green, blue): 
    assert 0 <= red <=255
    assert 0 <= green <=255
    assert 0 <= blue <=255
    return red + (green << 8) + (blue << 16)


def GetItem():

    with open(os.path.join(settings.sourcePath, settings.jsonFileName), 'r', encoding='utf_8') as f:
        json_data = json.load(f)
        
    for idx, item in enumerate(json_data, 1):
        # print(idx, item['type'])
        res = []

        if item['type'] == 'open':
            res = [[subitem['text'], subitem['weight'], '%'] for subitem in item['data']['pageA']]

        elif item['type'] == 'scene':
            for i, subitem in enumerate(item['data'], 1):
                res.append([f"{i}. {subitem['name']}", '', ''])
                res.extend([[f"\t{options['name']}", options['value'], '%'] for options in subitem['options']])

        else:
            res = [[subitem['name'], subitem['value'], '%'] for subitem in item['data']]

        yield {'idx':idx, 'name':item['name'], 'res':res}


def UpdateTable(table, data):

    table.AllowAutoFit = False
    
    if table.Rows.Count - len(data) < 2:
        for _ in range(len(data) - table.Rows.Count + 2):
            table.Rows.Add(BeforeRow=table.Rows(2))
    elif table.Rows.Count - len(data) > 2:
        for _ in range(table.Rows.Count - len(data) - 2):
            table.Rows(table.Rows.Count - 1).Delete()

    for row, col in product(range(len(data)), range(len(data[0]))):
        table.Cell(row + 2, col + 1).Range.Text = data[row][col]
        
    table.Cell(row + 3, 2).Range.Text = sum(map(int, filter(None, [x[1] for x in data])))