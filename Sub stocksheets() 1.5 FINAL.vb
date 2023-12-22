Sub stocksheets()
Dim wscount As Integer
Dim ws As Integer

'set wscount equal to the number of worksheets in the active workbook.
wscount = ThisWorkbook.Worksheets.Count

'begin the loop. 
For ws = 1 To wscount
'set variables
    Dim Summary As Integer
    Dim Openvalue As Double
    Dim Closevalue As Double
    Dim Totalyearchange As Double
    Dim Totalpercentchange As Double
    Dim Greatestincrease As Double
    Dim Greatestdecrease As Double
    Dim Greatestvolume As Double
    Dim Totalvolume As Double
'set values
        Greatestincrease = 0
        Greatestdecrease = 0
        Greatestvolume = 0
        Totalvolume = 0
        Summary = 2

'set column and row titles
    Cells(1,9).Value = "Ticker"
    Cells(1,10).Value = "Yearly Change"
    Cells(1,11).Value = "Percent Change"
    Cells(1,12).Value = "Total Stock Volume"
    Cells(1,16).Value = "Ticker"
    Cells(1,17).Value = "Value"
    Cells(2,15).Value = "Greatest % Increase"
    Cells(3,15).Value = "Greatest % Decrease"
    Cells(4,15).Value = "Greatest Total Volume"

'set row count
    rowcount = Cells(Rows.Count, "A").End(xlUp).Row
'for loop through rows and getting total
    For i = 2 To rowcount
        Totalvolume = Totalvolume + Cells(i,7).Value
        Openvalue = Cells(2,3).Value
    If Cells(i + 1,1).Value <> Cells(i,1).Value Then
        Closevalue = Cells(i,6).Value
        Totalyearchange = Closevalue - Openvalue
        Totalpercentchange = Totalyearchange / Openvalue * 100
        Cells(Summary,9).Value = Cells(i,1).Value
        Cells(Summary,10).Value = Totalyearchange
        Cells(Summary,11).Value = "%" & Totalpercentchange
        Cells(Summary,12).Value = Totalvolume
'change yearly color
            If Totalyearchange > 0 Then
                Cells(Summary, 10).Interior.ColorIndex = 4
                ElseIf Totalyearchange < 0 Then
                    Cells(Summary, 10).Interior.ColorIndex = 3
                Else
                    Cells(Summary, 10).Interior.ColorIndex = 2
            End If
        Summary = Summary + 1
        Totalvolume = 0
        Openvalue = Cells(i + 1, 3).Value
    End If
    Next i
'set row to zero and count again
    rowcount = 0
    rowcount = Cells(Rows.Count, "A").End(xlUp).Row
'loop through for greatests increase and decrease
    For i = 2 To rowcount
    
    If Cells(i,11).Value > 0 And Cells(i,11).Value > GreatestIncrease Then
        GreatestIncrease = Cells(i,11).Value
        Cells(2, 16).Value = Cells(i,9).Value
    ElseIf Cells(i,11).Value < 0 And Cells(i,11).Value < GreatestDecrease Then
        GreatestDecrease = Cells(i,11).Value
        Cells(3,16).Value = Cells(i,9).Value
    End If
'set greatest increase and decrease
        Cells(2, "Q").Value = "%" & GreatestIncrease * 100
        Cells(3, "Q").Value = "%" & GreatestDecrease * 100
'set greatest volume
    If Cells(i,12).Value > GreatestVolume Then
        GreatestVolume = Cells(i,12).Value
        Cells(4,16).Value = Cells(i,9).Value
    End If
        Cells(4, "Q").Value = GreatestVolume
    Next i
      
Next ws
End Sub