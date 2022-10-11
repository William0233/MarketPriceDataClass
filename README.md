# MarketPriceDataClass
HOW TO DOWNLOAD HISTORICAL STOCK PRICE DATA DIRECTLY INTO EXCEL USING PYTHON
How would you like to have a program that downloads the latest price data for any company you want to invest in? 
Maybe you are a day trader and looking for a more efficient way to get data into the same excel sheet where you have formulas that automatically evaluate security. 
Perhaps you are a financial modeler who’s looking for a way to start your model. This program named “MarketPriceDataClass “ will help with downloading market prices 
into your preferred workbook, it will find the designated sheet and the row columns you want your data to download.  
MarketPriceDataClass is a program that fetches current and historical data from Finance Yahoo and returns a Panda DataFrame; then it exports into an excel sheet 
where it may be used for all sorts of applications. This script also can create folders (arranged by dates) to store the downloaded ticker data in a CSV file for 
future uses.

First thing is to know that this is easier than you think, especially if you’re not a programmer. Simply follow the instructions laid out, and you shall have an automated system in no time.  

SETTING UP YOUR ENVIRONMENT:
Be sure to have Visual Code installed to run your python file.
Create an excel workbook in the same folder as your Visual Code.

Install Xlwings as an Add-in excel, which will help python to communicate with excel: https://docs.xlwings.org/en/stable/addin.htmlCopy and paste MarketPriceDataClass python scrip into Visual Code

RUNNING MarketPriceDataClass PYTHON SCRIP:
Add your desired ticker in “tickerName”
Enter the current date of interest in “todayDate”
Enter the number of data desired in “numData”. (Key in 252 to get data for year-to-date)

Like magic, your data should be downloaded and printed. If you have Xlwings installed, then you may enter the followings:
The name of your excel workbook in “workbookName”
Select the excel sheet to allocate the data “sheetName”
Clear previous data of a range with “clearCols”
The row-col you want your initial data point to be allocated in excel.
I wish you the best in your endeavor.
