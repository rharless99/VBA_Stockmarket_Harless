Attribute VB_Name = "HarlessCode"
Sub SummaryTablePractice()

For Each ws In Worksheets
ws.Activate

'Define Variables
    Dim Ticker_Name As String
    Dim Volume_Total As Variant
    Dim Volume_Amount As Long
    Dim Summary_Table_Row As Integer
    Dim LastRow As Long
    Dim Open_Value As Double
    Dim Close_Value As Double

'Initialize Variables
    Open_Value = Cells(2, 3).Value
    Volume_Total = 0
    Summary_Table_Row = 2
    LastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row

'Format Summary Table
    Cells(1, 9).Value = "Ticker Name"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"

    Cells(1, 9).Font.Bold = True
    Cells(1, 10).Font.Bold = True
    Cells(1, 11).Font.Bold = True
    Cells(1, 12).Font.Bold = True

    Cells(1, 9).EntireColumn.AutoFit
    Cells(1, 10).EntireColumn.AutoFit
    Cells(1, 11).EntireColumn.AutoFit
    Cells(1, 12).EntireColumn.AutoFit
    
'Sum the volume total and store corresponding ticker name in Column I and L
        For i = 2 To LastRow
            Volume_Amount = Cells(i, 7).Value
            If Cells(i, 1).Value = Cells(i + 1, 1).Value Then
                Volume_Total = Volume_Total + Volume_Amount
            Else
                Ticker_Name = Cells(i, 1).Value
                Volume_Total = Volume_Total + Volume_Amount
        
                Range("I" & Summary_Table_Row).Value = Ticker_Name
                Range("L" & Summary_Table_Row).Value = Volume_Total
     
                Summary_Table_Row = Summary_Table_Row + 1
        
                 Volume_Total = 0
            End If
        Next i

'Calculate the change value and percent chance values and store them in summary table
    Dim Change As Variant
    Dim Percent_Change As Variant
    
    Change = Close_Value - Open_Value
    Summary_Table_Row = 2


        For a = 2 To LastRow
  
            If Cells(a, 1).Value <> Cells(a + 1, 1).Value Then
                Close_Value = Cells(a, 6).Value
                Change = Close_Value - Open_Value
                    If Open_Value = 0 Then
                        Percent_Change = 0#
                    Else
                        Percent_Change = Change / Open_Value
                    End If
                Range("J" & Summary_Table_Row).Value = Change
                Range("K" & Summary_Table_Row).Value = Percent_Change
                Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                Summary_Table_Row = Summary_Table_Row + 1
                Open_Value = Cells(a + 1, 3).Value
            End If
        Next a

'Color code Percent Change with green positive and red negative
     Dim SumTableTickerCounter As Integer
    Dim new_last_row As Integer
    new_last_row = ActiveSheet.Cells(Rows.Count, 9).End(xlUp).Row
        For b = 2 To new_last_row
            If Cells(b, 10).Value < 0 Then
                Cells(b, 10).Interior.ColorIndex = 3
            ElseIf Cells(b, 10).Value > 0 Then
                Cells(b, 10).Interior.ColorIndex = 4
            Else
                Cells(b, 10).Interior.ColorIndex = 2
            End If
        Next b

'Create new summary table for Greatest % increase and decrease and total volume
'Use worksheet functions and index matching

    Dim Greatest_Percent_Increase As Double
    Dim Greatest_Percent_Decrease As Double
    Dim Greatest_Total_Volume As Variant
    
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    
    Columns("O").ColumnWidth = 20
    Cells(1, 16).EntireColumn.AutoFit
    Cells(1, 17).EntireColumn.AutoFit
    
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    
    Greatest_Percent_Increase = WorksheetFunction.Max(Range("K:K"))
    Cells(2, 17).Value = Greatest_Percent_Increase
    Range("Q2").NumberFormat = "0.00%"
    Range("P2") = "=Index(I:I, match(Q2, K:K, 0))"
    
    Greatest_Percent_Decrease = WorksheetFunction.Min(Range("K:K"))
    Cells(3, 17).Value = Greatest_Percent_Decrease
    Range("Q3").NumberFormat = "0.00%"
    Range("P3") = "=Index(I:I, match(Q3, K:K, 0))"
    
    Greatest_Total_Volume = WorksheetFunction.Max(Range("L:L"))
    Cells(4, 17).Value = Greatest_Total_Volume
    Range("P4") = "=Index(I:I, match(Q4, L:L, 0))"
    
 Next ws
    


End Sub

