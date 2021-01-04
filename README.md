# VBA-challenge

Create a script that will loop through all the stocks for one year and output the following information.

The ticker symbol. --this is a string variable and will be 1st column, this is the what we are tracking
Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
    yearly_chg is the diff between first_openpr - last_clospr for each ticker, use first_openpr for the denomitator and
    yearly_chg as the numerator.
The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
    yearly_chg / first_openpr - this needs to be checked for 0 value for error trapping
The total stock volume of the stock.
    
You should also have conditional formatting that will highlight positive change in green and negative change in red.
