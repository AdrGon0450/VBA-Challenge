# VBA-Challenge

This was my attempt to create a VBA script, It was not fully successful at the time of turn in.

For my references I used various google searches which I will list below. I did alter the code examples given from these sources to match my usecase:
https://stackoverflow.com/questions/45023938/how-do-activesheet-activeworkbook-activesheet-and-application-activesheet-behav
https://excelchamps.com/vba/count-sheets/#Count_Sheets_from_the_Active_Workbook
https://support.microsoft.com/en-au/topic/macro-to-loop-through-all-worksheets-in-a-workbook-feef14e3-97cf-00e2-538b-5da40186e2b0 - Code modified from here
https://techcommunity.microsoft.com/t5/excel/understand-cells-rows-count-a-end-xlup-row/m-p/292728 - cells row count from here
https://www.wallstreetmojo.com/vba-row-count/
https://excelchamps.com/vba/cell-value/

I also referenced a portion of script I will list below from fellow student Ryan MacFarlane, Who was tutored on this portion as well.
I did make modifications to the sample script to match my use case

-------------------------------------------------------------------------------------------------------------------------------------------
 If Cells(i + 1, "A").Value <> Cells(i, "A").Value Then
          'Get ClosePrice
              ClosePrice = Cells(i, "F").Value
              YearlyChange = ClosePrice - OpenPrice
              PercentChange = YearlyChange / OpenPrice * 100
              Cells(SummaryRow, "I").Value = Cells(i, "A").Value
              Cells(SummaryRow, "J").Value = YearlyChange
              Cells(SummaryRow, "K").Value = "%" & PercentChange
              Cells(SummaryRow, "L").Value = Totalvol
              'Assign green and red
              If YearlyChange > 0 Then
                    Cells(SummaryRow, "J").Interior.ColorIndex = 4
                ElseIf YearlyChange < 0 Then
                    Cells(SummaryRow, "J").Interior.ColorIndex = 3
                Else
                    Cells(SummaryRow, "J").Interior.ColorIndex = 2
              End If
              SummaryRow = SummaryRow + 1
              Totalvol = 0
              OpenPrice = Cells(i + 1, "C").Value
-------------------------------------------------------------------------------------------------------------------------------------------

The script included does not provide accurate calculations, and returns incorrect results.
I tried.
