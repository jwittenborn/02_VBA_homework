Sub Stocks()

Dim Ticker As String
Dim dateValue As Integer
Dim OpenValue As Double
Dim CloseValue As Double
Dim HiValue As Double
Dim LoValue As Double
Dim Volume As Double
Volume = 0

Dim Change As Double
Change = 0
Dim PctChange As Double

Dim OutputCounter As Integer
Dim TickerCounter As Integer

OutputCounter = 2
TickerCounter = 1

Ticker = Cells(2, 1).Value
OpenValue = Cells(2, 3).Value


Dim WS As Worksheet
 
 For Each WS In Worksheets
 

    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    'Cells(2, 15).Value = lastRow
    
    Volume = 0
    OutputCounter = 2


    For I = 3 To lastRow
    
        Volume = Volume + Cells(I, 7)
    
        'find last ticker
      If Cells(I, 1).Value <> Cells(I + 1, 1).Value Then 'last value of ticker
        CloseValue = Cells(I, 6).Value
        Change = CloseValue - OpenValue
        
        If OpenValue <> 0 Then
        PctChange = Change / OpenValue
        Else
        PctChange = 0
        End If
        
        Cells(OutputCounter, 9).Value = Ticker
        Cells(OutputCounter, 10).Value = Change
        Cells(OutputCounter, 11).Value = PctChange
        Cells(OutputCounter, 12).Value = Volume
        
        OutputCounter = OutputCounter + 1
        
        Ticker = Cells(I, 1)
        OpenValue = Cells(I, 3).Value
        Volume = 0
        TickerCounter = TickerCounter + 1
    
      End If
    
    Next I
    
    
    
    'output max increase, decrease and volume
    Dim MaxChange As Double
    Dim MinChange As Double
    Dim MaxVolume As Double

    MaxChange = Application.WorksheetFunction.Max(Range("J2:J1000"))
    Cells(2, 16).Value = MaxChange

    MinChange = Application.WorksheetFunction.Min(Range("J2:J1000"))
    Cells(3, 16).Value = MinChange

    MaxVolume = Application.WorksheetFunction.Max(Range("L2:L1000"))
    Cells(4, 16).Value = MaxVolume


 Next 'WS
 



End Sub

