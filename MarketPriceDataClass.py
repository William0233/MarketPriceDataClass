import os
import pandas as pd
from datetime import datetime, timedelta
from pandas_datareader._utils import RemoteDataError
import xlwings as xw

class MarketPriceData(object):

    def __init__(self, selectedTicker, input_date_string, numData):
        self.numData = numData
        numData = numData + (numData/2)
        self.selectedTicker = selectedTicker
        self.input_date = datetime.strptime(input_date_string, '%m/%d/%Y')
        self.start_date = self.input_date - timedelta(days=numData)
        self.end_date = self.input_date.strftime("%m/%d/%Y")
        self.start_date = self.start_date.strftime("%m/%d/%Y")
        self.folder_date = self.input_date.strftime('%Y-%m-%d')
        dirname = os.path.dirname(__file__)
        self.folderName = os.path.join(dirname,'Data\\RandomSecurity\\{}_Imports'.format(self.folder_date))
    
    def get_workbook(self, workbook_name):
        xw.Book(os.path.join(os.path.dirname(__file__),workbook_name)).set_mock_caller()
        wb = xw.Book.caller()
        return wb

    def get_sheet(self,workbook_name, sheet_name):
        wb = self.get_workbook(workbook_name)
        return wb.sheets[sheet_name]
    
    @xw.sub
    def loadSelected(self):
        #This function dowloads security data from finacance yahoo
        selected_ticker = self.selectedTicker
        input_date = self.start_date    
        end_date = self.end_date
        folderName = self.folderName
        os.makedirs(folderName, exist_ok=True)
        csvFilename = '{}\\{}.csv'.format(folderName,selected_ticker)
        panel_data = pd.DataFrame()
        
        try:
            from pandas_datareader import data
            panel_data = data.DataReader(
                            selected_ticker, 'yahoo', input_date, end_date)
            panel_data = panel_data[['Open', 'High', 'Low', 'Close', 'Adj Close',
                                        'Volume']].tail(self.numData)
            print("Writing CSV File {}.csv".format(selected_ticker))
            panel_data.to_csv(csvFilename)
            #panel_data.reset_index(inplace=True)
        except RemoteDataError:
            print("No symbol called {}".format(selected_ticker))
            exit()    
        return panel_data




tickerName = 'AAPL'
todayDate ='8/25/2022'
numData = 252 # one year of trading

data = MarketPriceData(tickerName, todayDate, numData)
price_panel = data.loadSelected()
print(price_panel)

## if you have xlwings installed in excel, you may run the next few lines
## to have the data prints directely in excel
## if not use the # to comment the next 7 lines
workbookName ='Workbook1.xlsb'
sheetName = 'Sheet1'
clearCols = 'A1:G500000'
initialCols = 'A1'
sheet = data.get_sheet(workbookName, sheetName) 
sheet.range(clearCols).clear_contents()
sheet.range(initialCols).value = price_panel
