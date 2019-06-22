VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub homework_2()

   
        
            Dim WS As Worksheet
        
            
            For Each WS In ActiveWorkbook.worksheets
            WS.Activate
        
        ' Determine the Last Row
            lastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

 
        ' Add Heading for summary
            Cells(1, "I").Value = "Ticker"
            Cells(1, "J").Value = "Yearly Change"
            Cells(1, "K").Value = "Percent Change"
            Cells(1, "L").Value = "Total Stock Volume"
        
        'Create Variable to hold Value
            Dim Open_Price As Double
            Dim Close_Price As Double
            Dim Yearly_Change As Double
            Dim Ticker_Name As String
            Dim Percent_Change As Double
            Dim Vol As Double
            Vol = 0
            Dim Row As Double
            Row = 2
            Dim Column As Integer
            Column = 1
            Dim i As Long
        
        'Set Initial Open Price
            Open_Price = Cells(2, Column + 2).Value
         ' Loop through all ticker symbol
        
                For i = 2 To lastRow
                ' Check if we are still within the same ticker symbol, if it is not...
                    If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
                        ' Set Ticker name
                            Ticker_Name = Cells(i, Column).Value
                            Cells(Row, Column + 8).Value = Ticker_Name
                        ' Set Close Price
                             Close_Price = Cells(i, Column + 5).Value
                        'Add Yearly Change
                            Yearly_Change = Close_Price - Open_Price
                            Cells(Row, Column + 9).Value = Yearly_Change
                        ' Add Percent Change
                            If (Open_Price = 0 And Close_Price = 0) Then
                                Percent_Change = 0
                            ElseIf (Open_Price = 0 And Close_Price <> 0) Then
                                Percent_Change = 1
                            Else
                                Percent_Change = Yearly_Change / Open_Price
                                Cells(Row, Column + 10).Value = Percent_Change
                                Cells(Row, Column + 10).NumberFormat = "0.00%"
                             End If
                
                    ' Add Total Volumn
                        Vol = Vol + Cells(i, Column + 6).Value
                        Cells(Row, Column + 11).Value = Vol
                    ' Add one to the summary table row
                        Row = Row + 1
                    ' reset the Open Price
                        Open_Price = Cells(i + 1, Column + 2)
                 ' reset the Volumn Total
                        Vol = 0
                'if cells are the same ticker
                    Else
                        Vol = Vol + Cells(i, Column + 6).Value
                    End If
                 
                 Next i
        
        ' Determine the Last Row of Yearly Change per WS
        YCLastRow = WS.Cells(Rows.Count, Column + 8).End(xlUp).Row
        ' Set the Cell Colors
        For j = 2 To YCLastRow
            If (Cells(j, Column + 9).Value > 0 Or Cells(j, Column + 9).Value = 0) Then
                Cells(j, Column + 9).Interior.ColorIndex = 10
            ElseIf Cells(j, Column + 9).Value < 0 Then
                Cells(j, Column + 9).Interior.ColorIndex = 3
            End If
        Next j
        
        ' Set Greatest % Increase, % Decrease, and Total Volume
                Cells(2, Column + 14).Value = "Greatest % Increase"
                Cells(3, Column + 14).Value = "Greatest % Decrease"
                Cells(4, Column + 14).Value = "Greatest Total Volume"
                Cells(1, Column + 15).Value = "Ticker"
                Cells(1, Column + 16).Value = "Value"
        
        
        ' Look through each rows to find the greatest value and its associate ticker
        For Z = 2 To YCLastRow
                If Cells(Z, Column + 10).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & YCLastRow)) Then
                    Cells(2, Column + 15).Value = Cells(Z, Column + 8).Value
                    Cells(2, Column + 16).Value = Cells(Z, Column + 10).Value
                    Cells(2, Column + 16).NumberFormat = "0.00%"
                ElseIf Cells(Z, Column + 10).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & YCLastRow)) Then
                    Cells(3, Column + 15).Value = Cells(Z, Column + 8).Value
                    Cells(3, Column + 16).Value = Cells(Z, Column + 10).Value
                    Cells(3, Column + 16).NumberFormat = "0.00%"
                ElseIf Cells(Z, Column + 11).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & YCLastRow)) Then
                    Cells(4, Column + 15).Value = Cells(Z, Column + 8).Value
                    Cells(4, Column + 16).Value = Cells(Z, Column + 11).Value
                End If
        
        Next Z
        
    Next WS
        
End Sub


