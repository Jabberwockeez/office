import settings
import process
from process import UpdateShapeChart, UpdateShapeTable
from process import GetData, GetSummaryData
import pandas as pd
import xlwings as xw
import win32com.client as com
from win32com.client import constants as c
import os
import re



def main(ppt):

    ppt.Slides(1).Shapes('标题 1').TextFrame.TextRange.Text = f'价格分析报告补充——{settings.price_date%10000//100}月'

    for slide in ppt.Slides:
        
        # if slide.SlideIndex < 13:
        #     continue

        if not 'Title_' in (shape.Name for shape in slide.Shapes):
            continue

        if '细分市场价格带' in slide.Shapes('Title_').TextFrame.TextRange.Text:
            chart = slide.Shapes('细分市场价格带')
            segment = re.search('(?<=\W)[^\W].*\D', slide.Shapes('Title_').TextFrame.TextRange.Text).group()
            columns_order = [series.Name for series in chart.Chart.SeriesCollection()]
            data = GetData(segment, columns_order)
            UpdateShapeChart(data, chart)
            del chart

        else:
            if '加权均价' in slide.Shapes('Title_').TextFrame.TextRange.Text:
                chart_name = '细分市场加权均价'
                sheet_name = '加权成交价汇总'
            elif '加权折扣率' in slide.Shapes('Title_').TextFrame.TextRange.Text:
                chart_name = '细分市场加权折扣率'
                sheet_name = '加权折扣率汇总'

            chart = slide.Shapes(chart_name)
            if slide.Shapes('Title_').TextFrame.TextRange.Text.startswith('细分市场及整体'):
                segment = '整体'
                index_order = [series.Name for series in chart.Chart.SeriesCollection()]
            elif slide.Shapes('Title_').TextFrame.TextRange.Text.startswith('各竞争对手整体'):
                segment = '对手'
                index_order = [series.Name for series in chart.Chart.SeriesCollection()]
            else:
                segment = re.search('(?<=\W)[^\W].*\D', slide.Shapes('Title_').TextFrame.TextRange.Text).group()
                index_order = [series.Name for series in chart.Chart.SeriesCollection() if series.Name]
            data = GetSummaryData(segment, sheet_name, index_order)
            UpdateShapeChart(data, chart)

        logger.info(f"Done - {slide.SlideIndex:02} {slide.Shapes('Title_').TextFrame.TextRange.Text}")
        del slide


if __name__ == "__main__":

    logger = settings.logger
    logger.info(f"\n开始更新...")

    reportFileName = settings.GetRecentFile(settings.reportPath, '价格分析报告补充')
    App = com.gencache.EnsureDispatch('Powerpoint.Application')
    ppt = App.Presentations.Open(reportFileName, WithWindow=settings.isVisible)
    App.WindowState = 2    # 窗口最小化

    process.OpenExcel()

    try:
        main(ppt)
    except Exception as e: 
        logger.info(e)
        logger.info(f"中断更新.")
    finally:    
        resultFileName = re.sub('(\d+年\d+月)(.*?)(\d+.)', '{}\g<2>{}.'.format(settings.report_date.strftime('%Y年%#m月'), pd.datetime.now().strftime('%Y%m%d%H%M')), reportFileName)
        print('本期报告:  ', resultFileName)
        ppt.SaveAs(os.path.join(settings.reportPath, resultFileName))
        ppt.Close()
        App.Quit()

        process.CloseExcel()
        logger.info(f"完成更新.")
