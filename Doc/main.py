# coding: utf-8
import settings
import process
# import docx
from win32com import client
import os
import re
from sys import exit


App = client.gencache.EnsureDispatch('Word.Application')
App.Visible = settings.isVisible

doc = App.Documents.Open(os.path.join(settings.sourcePath, settings.sourceFileName))

tmplate_start, tmplate_end = doc.Paragraphs(2).Range.Start, doc.Range().End

for item in process.GetItem():
    # print(item)

    doc.Paragraphs.Add()
    subject = f"第{item['idx']}题 {item['name']}"
    pattern = re.compile('(<.*?>)(.*?)(<.*?>)')    # 去除格式标签
    doc.Range(doc.Range().End - 1, doc.Range().End).Text = pattern.sub('\g<2>', subject)
    if pattern.search(subject):  # 存在格式标签,字体标红色
        start = doc.Paragraphs(doc.Paragraphs.Count).Range.Start + pattern.search(subject).regs[1][0]
        end = start + pattern.search(subject).regs[2][1] - pattern.search(subject).regs[2][0]
        doc.Range(start, end).Font.Color = process.RGB(*process.colorRed)
    doc.Paragraphs.Add()
    
    doc.Paragraphs.Add()
    doc.Tables(1).Select()
    App.Selection.Copy()
    doc.Range(doc.Range().End - 1, doc.Range().End).Paste()
    table = doc.Tables(doc.Tables.Count)
    process.UpdateTable(table, item['res'])
    doc.Paragraphs.Add()

doc.Range(tmplate_start, tmplate_end).Delete()

doc.SaveAs(os.path.join(settings.reportPath, settings.sourceFileName))
doc.Close()
App.Quit()