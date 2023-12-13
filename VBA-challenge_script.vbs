Attribute VB_Name = "Module1"
Sub alphabetical_testing()
    Dim Ticker As String
    Dim Yearly_Change As Double
    Dim Percentage_Change As Double
    Dim Total_Volume As LongLong
    Dim Max As Double
    Dim ws As Worksheet
    
For Each ws In Worksheets

    Yearly_Change = 0
    Percentage_Change = 0
    Total_Volume = 0
    
    Dim Ticker_Summary As Integer
    Ticker_Summary = 2
    
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
    
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Ticker = Cells(i, 1).Value
        Yearly_Change = Yearly_Change + Cells(i, 3).Value - Cells(i, 6).Value
        Percentage_Change = Percentage_Change + (Cells(i, 3).Value - Cells(i, 6).Value) / Cells(i, 6).Value
        Total_Volume = Total_Volume + Cells(i, 7).Value
        Range("I" & Ticker_Summary).Value = Ticker
        Range("J" & Ticker_Summary).Value = Yearly_Change
        Range("K" & Ticker_Summary).Value = Percentage_Change
        Range("L" & Ticker_Summary).Value = Total_Volume
        Ticker_Summary = Ticker_Summary + 1
        Yearly_Change = 0
        Percentage_Change = 0
        Total_Volume = 0
        
    Else
        Yearly_Change = Yearly_Change + (Cells(i, 3).Value - Cells(i, 6).Value)
        Percentage_Change = Percentage_Change + (Cells(i, 3).Value - Cells(i, 6).Value) / Cells(i, 6).Value
        Total_Volume = Total_Volume + Cells(i, 7).Value
    End If

    Next i
    
ws.Activate

    Next
    
End Sub

Sub Format()
    Dim ws As Worksheet
    
For Each ws In Worksheets

    For b = 2 To Cells(Rows.Count, 1).End(xlUp).Row
    If Cells(b, 10).Value > 0 Then
        Cells(b, 10).Interior.ColorIndex = 4
    ElseIf Cells(b, 10).Value < 0 Then
        Cells(b, 10).Interior.ColorIndex = 3
    End If
    
    Next b
    
    For c = 2 To Cells(Rows.Count, 1).End(xlUp).Row
    If Cells(c, 11).Value > 0 Then
        Cells(c, 11).Interior.ColorIndex = 4
    ElseIf Cells(c, 11).Value < 0 Then
        Cells(c, 11).Interior.ColorIndex = 3
    End If
    
    Next c
    
ws.Activate
Next
    
End Sub
Sub Bonus()
    Dim ws As Worksheet
    
For Each ws In Worksheets
    
    Cells(2, 17).Value = Application.WorksheetFunction.Max(Range("K2:K91"))
    Cells(3, 17).Value = Application.WorksheetFunction.Min(Range("K2:K91"))
    Cells(4, 17).Value = Application.WorksheetFunction.Max(Range("L2:L91"))

    Cells(2, 16).Value = WorksheetFunction.Index(Range("I2:I91"), WorksheetFunction.Match(Cells(2, 17).Value, Range("K2:K91"), 0))
    Cells(3, 16).Value = WorksheetFunction.Index(Range("I2:I91"), WorksheetFunction.Match(Cells(3, 17).Value, Range("K2:K91"), 0))
    Cells(4, 16).Value = WorksheetFunction.Index(Range("I2:I91"), WorksheetFunction.Match(Cells(4, 17).Value, Range("L2:L91"), 0))

ws.Activate
Next

End Sub
