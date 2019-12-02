import settings
# import process
# from process import sales_data, price_data, last_data, curr_data
# import pandas as pd
import win32com.client as com
from win32com.client import constants as c
import os
import re



reportFileName = settings.GetRecentFile(settings.reportPath, '价格分析报告')
# TODO: 自动获取上期报告
App = com.gencache.EnsureDispatch('Powerpoint.Application')
ppt = App.Presentations.Open(reportFileName, WithWindow=1)

for slide in ppt.Slides:

    if not 'Title_' in (sp.Name for sp in slide.Shapes):
        continue
    
    if not slide.SlideIndex in range(13,24):
        continue
    # if not 'Title 9' in (sp.Name for sp in slide.Shapes):
    #     continue
    
    # print(slide.SlideIndex, tuple(shape.Name for shape in slide.Shapes if 'Object' in shape.Name))
    shape_list = tuple(shape for shape in slide.Shapes if 'Object' in shape.Name)
    if len(shape_list):
        shape_list[0].Name = '细分市场加权均价'


    # # 检查是否存在
    # if slide.Shapes('Title_').TextFrame.TextRange.Text.startswith('本品价格走势分析'):
    #     shape_list = tuple(sp.Name for sp in slide.Shapes if sp.Name == '内容占位符 2')
    #     print('{:<3} - {}'.format(slide.SlideIndex, shape_list))

    # # 检查存在几个
    # if slide.Shapes('Title_').TextFrame.TextRange.Text.startswith('本品价格走势分析'):
    #     if '对角圆角矩形_车型名字' in (sp.Name for sp in slide.Shapes):
    #         print('{}{} 对角圆角矩形_车型名字'.format(slide.SlideIndex, '对角圆角矩形_车型名字' in (sp.Name for sp in slide.Shapes)))
    #         # print('{} - {}'.format(slide.SlideIndex, tuple(sp.Name for sp in slide.Shapes).count('对角圆角矩形_车型名字')))

    # 修改窗格名称
    # if '综合情况分析' in slide.Shapes('Title_').TextFrame.TextRange.Text:
    #     # if not '内容占位符 2' in (sp.Name for sp in slide.Shapes):
    #     #     print(slide.SlideIndex)
    #     #     continue
    #     # slide.Shapes('内容占位符 2').Name = '成交价描述'
    #     print(slide.SlideIndex, slide.Shapes('Title_').TextFrame.TextRange.Text)
    #     for shape in (shape for shape in slide.Shapes if shape.Type == 3):
    #         print(shape.Name)
    #         shape.Name = '加权成交价汇总'
    #     for shape in (shape for shape in slide.Shapes if shape.Type == 19):
    #         print(shape.Name)
    #         shape.Name = '月度数据汇总'
            

    # # 修改窗格名称
    # if slide.Shapes('Title_').TextFrame.TextRange.Text.startswith('本品价格走势分析'):
    #     # shape_list = list(filter(lambda x: int(re.search('\d+', x).group()) > 26, tuple(sp.Name for sp in slide.Shapes if sp.Name.startswith('Chart '))))
    #     shape_list = tuple(sp.Name for sp in slide.Shapes if sp.Name.startswith('Chart '))
    #     if shape_list:
    #         print('{} - {}'.format(slide.SlideIndex, shape_list))
    #         # slide.Shapes(shape_list[0]).Name = '城市成交价对比_详细'
    #     else:
    #         print(slide.SlideIndex)
        
    
# resultFileName = re.sub('(\d+月)(.*-)(\d+)', '{}月\g<2>{}'.format(settings.last_month.month, pd.datetime.now().strftime('%Y%m%d%H%M')), reportFileName)
# print(resultFileName)
# ppt.Save()
# ppt.Close()
# App.Quit()