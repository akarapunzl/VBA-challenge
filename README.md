# VBA-challenge
Code written for the Stock Market analysis to:
- Print the unique Tickers for each year
- Print the Yearly Change for each unique ticker
- Print the Percent change for each unique ticker
- Print the total volume for each unique ticker
- Find the Greatest Increase for each year and corresponding ticker
- Find the Greatest Decrease for each year and corresponding ticker
- Find the Max Total Volume for each year and corresponding ticker

This section of the code was discussed with classmate Rocky M. in slack
--
Sub cycle()

'set the variable
Dim ws As Worksheet

For Each ws In ActiveWorkbook.Worksheets
    Ticker ws
    
Next

End Sub
Sub Ticker(ws As Worksheet)

ws.Activate

--
I used activity 06-stu_creditcardchecker from class to structure the basis of the code - to find the individual ticker and the total volume per ticker. I also worked with a tutor and the ASKBCS to develop the rest of the code, specifically the percent change, the max increase/decrease/total volume. 

