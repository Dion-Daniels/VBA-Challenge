VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub VBA_Challenge_Dion_Daniels()

For Each ws In Worksheets

Dim Ticker As String
Dim Volume, GreatestVol As LongLong
Dim Yearly_Change, Percentage_Change, GreatestInc, GreatestDec, Closing, Opening As Double

Volume = 0
GreatestVol = 0
Opening = ws.Range("C2").Value
Closing = 0
Yearly_Change = 0
Percentage_Change = 0
GreatestInc = 0
GreatestDec = 0
Table = 2

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percentage Change"
ws.Range("L1").Value = "Total Stock Volume"

ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To LastRow

    If ws.Range("C" & i).Value = 0 And ws.Range("F" & i).Value = 0 And ws.Range("G" & i).Value = 0 Then
    
            Opening = ws.Cells(i + 1, 3).Value

    ElseIf ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
    
            Ticker = ws.Cells(i, 1).Value
    
            Closing = ws.Range("F" & i).Value
    
            Yearly_Change = Closing - Opening
    
            Percentage_Change = Yearly_Change / Opening
    
            Volume = Volume + ws.Range("G" & i).Value
    
            ws.Range("I" & Table).Value = Ticker
            
            ws.Range("J" & Table).NumberFormat = "$0.00"
    
            ws.Range("J" & Table).Value = Yearly_Change
    
            ws.Range("K" & Table).NumberFormat = "0.00%"
    
            ws.Range("K" & Table).Value = Percentage_Change
    
            ws.Range("L" & Table).Value = Volume
    
            Opening = ws.Cells(i + 1, 3).Value
    
            Table = Table + 1
    
            Volume = 0
    
        Else
    
            Volume = Volume + ws.Range("G" & i).Value
    
    End If
Next i
For j = 2 To LastRow
    If ws.Range("J" & j).Value > 0 Then
    
            ws.Range("J" & j).Interior.ColorIndex = 4
            ws.Range("K" & j).Interior.ColorIndex = 4
    
        ElseIf ws.Range("J" & j).Value < 0 Then
    
            ws.Range("J" & j).Interior.ColorIndex = 3
            ws.Range("K" & j).Interior.ColorIndex = 3
            
        ElseIf ws.Range("J" & j).Value = 0 And ws.Range("L" & j).Value > 0 Then
        
            ws.Range("J" & j).Interior.ColorIndex = 27
            ws.Range("K" & j).Interior.ColorIndex = 27
        Else
         Exit For
    
    End If
Next j
For k = 2 To LastRow
    If ws.Range("K" & k).Value > GreatestInc Then
    
        GreatestInc = ws.Range("K" & k).Value
        ws.Range("P2").Value = ws.Range("I" & k).Value
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q2").Value = GreatestInc
        
        
    ElseIf ws.Range("K" & k).Value < GreatestDec Then
    
        GreatestDec = ws.Range("K" & k).Value
        ws.Range("P3").Value = ws.Range("I" & k).Value
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("Q3").Value = GreatestDec
    
    
    End If
Next k
For m = 2 To LastRow
    If ws.Range("L" & m).Value > GreatestVol Then
    
        GreatestVol = ws.Range("L" & m).Value
        ws.Range("P4").Value = ws.Range("I" & m).Value
        ws.Range("Q4").Value = GreatestVol
    
    End If


Next m

Next ws


End Sub
