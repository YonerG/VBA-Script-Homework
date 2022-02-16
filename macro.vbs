Sub Stocks()
' Declare Dims
   Dim total As Double
   Dim i As Long
   Dim change As Double
   Dim k As Integer
   Dim start As Long
   Dim rowCount As Long
   Dim percent As Double
  
   ' Create columns
   Range("J1").Value = "Ticker"
   Range("K1").Value = "Yearly Change"
   Range("L1").Value = "Percent Change"
   Range("M1").Value = "Total Stock Volume"
   
   
   ' Initialize
   k = 0
   total = 0
   change = 0
   start = 2
   ' get last row #
   rowCount = Cells(Rows.Count, "A").End(xlUp).Row
   For i = 2 To rowCount
       ' Update ticker when it changes
       If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
           ' Update variables
           total = total + Cells(i, 7).Value
           ' zero values
           If total = 0 Then
               ' populate summary table
               Range("J" & 2 + k).Value = Cells(i, 1).Value
               Range("K" & 2 + k).Value = 0
               Range("L" & 2 + k).Value = "%" & 0
               Range("M" & 2 + k).Value = 0
           Else
               ' nonzero value updates
               If Cells(start, 3) = 0 Then
                   For find_value = start To i
                       If Cells(find_value, 3).Value <> 0 Then
                           start = find_value
                           Exit For
                       End If
                    Next find_value
               End If
               ' Calculate Change
               change = (Cells(i, 6) - Cells(start, 3))
               percent = change / Cells(start, 3)
               ' start of the next stock ticker
               start = i + 1
               ' populate summary table
               Range("J" & 2 + k).Value = Cells(i, 1).Value
               Range("K" & 2 + k).Value = change
               Range("K" & 2 + k).NumberFormat = "0.00"
               Range("L" & 2 + k).Value = percent
               Range("L" & 2 + k).NumberFormat = "0.00%"
               Range("M" & 2 + k).Value = total
               ' color changes from the positive and negative results
               Select Case change
                   Case Is > 0
                       Range("K" & 2 + j).Interior.ColorIndex = 4
                   Case Is < 0
                       Range("K" & 2 + j).Interior.ColorIndex = 3
                   Case Else
                       Range("K" & 2 + j).Interior.ColorIndex = 0
               End Select
           End If
           ' reset vars for next totals
           total = 0
           change = 0
           k = k + 1
           
       ' if ticker the same enter total
       Else
           total = total + Cells(i, 7).Value
       End If
   Next i


End Sub

Sub WorksheetLoop()


 
 
Dim ws As Worksheet
Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet 'remember which worksheet is active in the beginning

For Each ws In ThisWorkbook.Worksheets
    ws.Activate
    Call Stocks
    
    
Next

starting_ws.Activate 'activate the worksheet that was originally active

End Sub
