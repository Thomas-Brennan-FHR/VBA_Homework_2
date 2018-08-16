Attribute VB_Name = "Module1"
Sub Worksheet_Loop()

Dim Sheet As Worksheet

    For Each Sheet In Worksheets
        
        Sheet.Select
        Call stock_analysis
    
    Next Sheet
    
End Sub



Sub stock_analysis():

'Declare Variables
Dim Stockticker(0 To 5000) As String
Dim Stockyearopen(0 To 5000) As Double
Dim Stockyearclose(0 To 5000) As Double
Dim stockvolume(0 To 5000) As Double
Dim ArrayIndex As Integer


'Assign inital variable values
ArrayIndex = 0

'Write Output Headers
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

'Store Values in Array
For I = 2 To Application.WorksheetFunction.CountA(Range("A:A"))

    stockvolume(ArrayIndex) = stockvolume(ArrayIndex) + Cells(I, 7).Value
    
    If (I = 2 Or Cells(I, 2).Value < Cells(I - 1, 2).Value) Then
        Stockticker(ArrayIndex) = Cells(I, 1).Value
        Stockyearopen(ArrayIndex) = Cells(I, 3).Value
    
    ElseIf Cells(I, 2).Value > Cells(I + 1, 2).Value Then
        Stockyearclose(ArrayIndex) = Cells(I, 6).Value
        ArrayIndex = ArrayIndex + 1
    End If

Next I

'Output Loop
j = 0
While Stockticker(j) <> ""

    'Ticker Output
    Cells(j + 2, 9).Value = Stockticker(j)
    
    'Yearly Change Output & Formating
    Cells(j + 2, 10).Value = Stockyearclose(j) - Stockyearopen(j)
        If Cells(j + 2, 10).Value > 0 Then
         Cells(j + 2, 10).Interior.ColorIndex = 4
        ElseIf Cells(j + 2, 10).Value < 0 Then
         Cells(j + 2, 10).Interior.ColorIndex = 3
        End If
    
    
    'Percent Change Output
    If Stockyearopen(j) = 0 Then
        Cells(j + 2, 11).Value = Format(0, "Percent")
    Else
        Cells(j + 2, 11).Value = Format(Stockyearclose(j) / Stockyearopen(j) - 1, "Percent")
    End If
    
    'Total Volume Output
    Cells(j + 2, 12).Value = stockvolume(j)
    
j = j + 1
Wend



'Hard Mode Analysis
'-------------------------------------------------------------------------

Dim GreatestIncrease As Double
Dim GreatestDecrease As Double
Dim GreatestVol As Double
Dim GreatestIncreaseTicker As String
Dim GreatestDecreaseTicker As String
Dim GreatestVolTicker As String
k = 0

GreatestIncrease = 0
GreatestDecrease = 0
GreatestVol = 0

'Loop to store Greatest Values
While Stockticker(k) <> ""

    If Stockyearopen(k) > 0 Then
        If Stockyearclose(k) / Stockyearopen(k) - 1 > GreatestIncrease Then
           GreatestIncrease = Stockyearclose(k) / Stockyearopen(k) - 1
             GreatestIncreaseTicker = Stockticker(k)
        End If
    
        If Stockyearclose(k) / Stockyearopen(k) - 1 < GreatestDecrease Then
             GreatestDecrease = Stockyearclose(k) / Stockyearopen(k) - 1
             GreatestDecreaseTicker = Stockticker(k)
         End If

    End If


    If stockvolume(k) > GreatestVol Then
        GreatestVol = stockvolume(k)
        GreatestVolTicker = Stockticker(k)
    End If

    k = k + 1

Wend

'Write Outputs
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"
Range("P2").Value = GreatestIncreaseTicker
Range("P3").Value = GreatestDecreaseTicker
Range("P4").Value = GreatestVolTicker
Range("Q2").Value = Format(GreatestIncrease, "Percent")
Range("Q3").Value = Format(GreatestDecrease, "Percent")
Range("Q4").Value = GreatestVol

End Sub
