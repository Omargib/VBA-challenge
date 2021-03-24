sub stockanalysis ()
'Assigning Headers
range("i1").value = "Ticker"
range("j1").value = "Yearly Change"
range("k1").value = "Percent Change"
range("l1").value = "Total Stock Volume"
range("P1").value = "Ticker"
range("Q1").value = "Value"
range("o2").value = "Greatest % Increase"
range("o3").value = "Greatest % Decrease"
range("o4").value = "Greatest Total Volume"
'Formating Headers
Range("i1:l1, p1:q1, o2:o4").Font.Bold = True
Range("i1:l1, p1:q1, o2:o4").ColumnWidth = 20


'Determine Last Row
Dim LastRow as long
LastRow = Range("A" & Rows.Count).End(xlUp).Row
'Year Open
Dim Year_open as Double
Year_Open = Cells (2,3).value
'Year Close
Dim Year_Close as Double 
'Set Ticker Variable
Dim Ticker as String
'Set Total Stock Volume Variable 
Dim Total_Stock_Volume As Double
Total_Stock_Volume = 0
  ' Set Yearly Change
dim Yearly_Change as Double
' Set Percent Change
Dim Percent_Change as Double
' Set Greatest Increase
Dim Greatest_Increase as Double
' track of the location for each ticker in the ticker column
Dim Summary_Table_Row as Integer
Summary_Table_Row = 2
' Lastrow of the summarytable
Dim Lstr_summ as Double
' Set Max Valume
Dim Max_Volume as Double
'Set  Percentage of Greatest Increase
Dim PrcGreatest_Increase as Double
'Set Percentage of Greatest Decrease
Dim PrcGreatest_Decrease as Double

' Loop through all Ticker volume
For i = 2 To LastRow
 ' Check if we are still within the same Ticker, if it is not...
  If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
  'Set the Ticker
      Ticker = Cells(i, 1).Value
      ' Add to the Ticker Volume Total
      Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
       'Print the Tickers in the Summary Table
        Range("I" & Summary_Table_Row).Value = Ticker
        ' Print the Ticker Volme to the Summary Table
      Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
        ' Reset the Total_Stock_Volume
      Total_Stock_Volume = 0
      'Set Year Clos
      Year_Close = Cells (i, 6).Value
      ' Yearly change calculation
      Yearly_Change = Year_Close - Year_Open
      ' Print Yearly Change
      Range("J" & Summary_Table_Row).Value = Yearly_Change
' Validating 0
      If Year_Open = 0 Then
        Percent_Change = 0
      Else
        ' Percentage Change Calculation
        Percent_Change = (Year_Close - Year_Open)/Year_Open
      End If

      ' Pirnt Percent Change 
      Range("K" & Summary_Table_Row).Value = Percent_Change
      Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
      ' Next first open
      Year_Open = Cells (i + 1, 3).Value
      'Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
' If the cell immediately following a row is the same Ticker...
    Else

      ' Add to the Total_Stock_Volume
      Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

    End If

  Next i


' Determine the last row of the summary table
Lstr_summ = Range("I" & Rows.Count).End(xlUp).Row

' Determine the Max_Volume
Max_Volume = WorksheetFunction.Max(Range("L1:L" & Lstr_summ))

' Assign Greatest Total Voleme of the final Summary Table
Range ("Q4").Value = Max_Volume

' Detrmine the Ticker with the greatest volume
Range ("p4").Value = Range ("I" & WorksheetFunction.Match(Max_Volume, Range("L:L"), 0))

' Determine the Greatest % Increase
PrcGreatest_Increase = WorksheetFunction.Max(Range("K1:K" & Lstr_summ))

'Assign Greatest % Increase of the final Summary Table
Range ("Q2").Value = PrcGreatest_Increase
Range("Q2").NumberFormat = "0.00%"

'Determine the Ticker of the Greatest % Increase
Range ("P2").Value = Range ("I" & WorksheetFunction.Match(PrcGreatest_Increase, Range("K:K"), 0))

' Determine The Greatest % Decrease
PrcGreatest_Decrease = WorksheetFunction.Min(Range("K1:K" & Lstr_summ))

'Assign Greatest % Decrease of the final Summary Table
Range ("Q3").Value = PrcGreatest_Decrease
Range("Q3").NumberFormat = "0.00%"

'Determine the Ticker of the Greatest % Decrease
Range ("P3").Value = Range ("I" & WorksheetFunction.Match(PrcGreatest_Decrease, Range("K:K"), 0))

End Sub
