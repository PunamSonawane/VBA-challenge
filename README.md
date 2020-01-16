# VBA Challenge
A VBA script used to analyze real stock market data in a Microsoft Excel workbook.

## Background

For this project, I created a VBA (Visual Basic) script to analyze some stock market data. The data is inside a Microsoft Excel workbook and includes stock data for three years (2014, 2015, and 2016). Each year is a different tab/sheet inside the workbook. 

## Testing

I ran this script on both the testing Excel workbook (alphabetical_testing.xlsx) and on the final multiple year stock workbook (multiple_year_stock_data.xlsx).

My environment is Windows 10, Microsoft Excel 365. So, the script should work if you are using this environment. Not sure about Macs though, as I didn't test the script on a Mac.

## About the Script

You can find the script inside the **VBAStocks** folder of this repository. The script file is called **CalculateStockStats.bas**.

After you download and open up the multiple year stock data Excel workbook, you can run the script by doing the following:

1. Click the **Developer** tab.

2. Click **Visual Basic** to open the Visual Basic editor.

3. Inside the Visual Basic editor, click **File** > **Import File** and import the **VBA_WallStreet_Stock_Market.vbs** file in this repository.

4. Open up the **VBA_WallStreet_Stock_Market.vbs** file in the Visual Basic editor and then click the Run Macro button (green play icon) in the toolbar to run the script.

The script does take some time to run because it is running on every sheet. So, no need to run it more than once.

As the script runs, it is doing the following:

* It loops through all the stocks for one year for each run and takes the following information:

  * The ticker symbol

  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
  
  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
  
  * The total stock volume of the stock.

* It applies conditional formatting by highlighting positive yearly change values in green and negative yearly change values in red.

* Finally, it return the stock with the greatest percent increase, greatest percent decrease, and greatest total volume

## Sample Output

After the script has completed, go to the Excel workbook, and you should see the results of the script.

Here are screenshots of what the output looks like when I ran the scripts on my computer. These screenshots are also available in the **VBAStocks/** folder of this repository.

### 2014 Stock Data

![Image of 2014 Stock Data](./VBAStocks/2014Expected_result.PNG)

### 2015 Stock Data

![Image of 2015 Stock Data](./VBAStocks/2015Expected_result.PNG)

### 2016 Stock Data

![Image of 2016 Stock Data](./VBAStocks/2016Expected_result.PNG)
