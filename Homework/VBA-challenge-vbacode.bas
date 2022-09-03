Attribute VB_Name = "Module1"
Sub annual_summary()

Dim ticker_name As String
Dim i, next_entry As Integer

Cells(1, 9) = "Ticker"
Cells(1, 10) = "Yearly Change"
Cells(1, 11) = "Percent Change"
Cells(1, 12) = "Total Stock Volume"

' start of  Ticker list
ticker_name = Cells(2, 1)
Cells(2, 9) = ticker_name
next_entry = 3
i = 2
' If Ticker doesn't match previous ticker, then add it to the list
While Cells(i, 1) <> ""
    If Cells(i, 1) <> ticker_name Then
        Cells(next_entry, 9) = Cells(i, 1)
        next_entry = next_entry + 1
        ticker_name = Cells(i, 1)
    End If
    i = i + 1
Wend

' reset i and next_entry
i = 2
next_entry = 2
ticker_name = Cells(i, 9)

' new variables
Dim closing, opening As Double
opening = Cells(i, 3)

' Loops down ticker list and calculates yearly change and percent change
While ticker_name <> ""
    If Cells(i + 1, 1) <> ticker_name Then
        closing = Cells(i, 6)
        Cells(next_entry, 10) = closing - opening ' Yearly Change
        Cells(next_entry, 11) = (closing - opening) / opening 'Percent Change
        next_entry = next_entry + 1
        ticker_name = Cells(i + 1, 1)
        opening = Cells(i + 1, 3)
    End If
    i = i + 1
Wend

' reset variables
i = 2
next_entry = 2
ticker_name = Cells(i, 9)

' new variables
Dim volume As Double
volume = 0

' Calculates Volume by Ticker
While ticker_name <> ""
    volume = volume + Cells(i, 7)
    If Cells(i + 1, 1) <> ticker_name Then
        Cells(next_entry, 12) = volume
        next_entry = next_entry + 1
        ticker_name = Cells(i + 1, 1)
        volume = 0
    End If
    i = i + 1
Wend


' Bonus

Cells(2, 14) = "Greatest % Increase"
Cells(3, 14) = "Greatest % Decrease"
Cells(4, 14) = "Greatest Total Volume"
Cells(1, 15) = "Ticker"
Cells(1, 16) = "Value"


Cells(2, 16) = Application.WorksheetFunction.max(Range("k:k"))
Cells(3, 16) = Application.WorksheetFunction.min(Range("k:k"))
Cells(4, 16) = Application.WorksheetFunction.max(Range("l:l"))

Dim end_row As Double
ticker_name = Cells(2, 1)
i = 2
end_row = 1

' index of last row
While ticker_name <> "":
 end_row = end_row + 1
    If Cells(i + 1, 9) = "" Then
        ticker_name = ""
    End If
    i = i + 1
Wend

Dim max, min, tot_vol As Double
max = 0
min = 0
tot_vol = 0

' max increase
For i = 2 To end_row
    If Cells(i, 11) > max Then
       max = Cells(i, 11)
       ticker_name = Cells(i, 9)
    End If
Next i

Cells(2, 15) = ticker_name

' max decrease
For i = 2 To end_row
    If Cells(i, 11) < min Then
       min = Cells(i, 11)
       ticker_name = Cells(i, 9)
    End If
Next i

Cells(3, 15) = ticker_name

' max total volume
For i = 2 To end_row
    If Cells(i, 12) > tot_vol Then
       tot_vol = Cells(i, 12)
       ticker_name = Cells(i, 9)
    End If
Next i

Cells(4, 15) = ticker_name

Dim yearly_change As Range
Set yearly_change = Range("j:j")
yearly_change.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
yearly_change.FormatConditions(1).Interior.Color = RGB(255, 0, 0)
yearly_change.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
yearly_change.FormatConditions(2).Interior.Color = RGB(0, 255, 0)
Range("J1:J1").FormatConditions.Delete

' number formatting
Range("k:k").NumberFormat = "0.00%"
Range("P2:P3").NumberFormat = "0.00%"



End Sub

