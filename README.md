## Stock Market Analysis using VBA

In this project I created a VBA script that will loop through all the stocks for several sheets and outputed:

  * The ticker symbol.

  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The total stock volume of the stock.

I then applied conditional formatting that will highlight positive change in green and negative change in red.


### Note
* The files "Alphabetical_Testing" and "Multiple_Year_Stock_Data" are base files with original data kept as back up in case the "xlsm" files with the VBA scripts are corrupted. The "xlsm" files are in the analysis folder

* Due to the large nature of the "Multiple_Year_Stock_Data" file, running the VBA code will take a long time so the code is run and tested on the "Alphabetical_Testing" file first. Once everything works, the code is then applied to the "Multiple_Year_Stock_Data".

* The "Alphabetical_Testing" file contains six worksheets "A","B","C","D","E","P". Each sheet contains information for Ticker symbols that start with that alaphabet in the sheet name. 

* The "Multiple_Year_Stock_Data" files contains three sheets "2014","2015" and "2016". Each sheet contains information for all ticker symbols for the stated year in the sheet name.

 * The VBA code is meant to produce the same outcome irrespective of which file. It loops through each sheet and outpuths information for all ticker symbols on the sheet.

 * The VBA script is also saved in a markdown file "VBA _Script_Stocks".


