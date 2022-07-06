Attribute VB_Name = "Module2"
Sub Tickers2018():


'Set Variables
Dim Ticker_Symbol As String
Dim Ticker_Map As String
Dim Trade_Date As Date
Dim Open_Price As Double
Dim High_Price As Double
Dim Low_Price As Double
Dim Close_Price As Double
Dim Volume As Double

'Set initial Variables
Dim YOY_CHG As Double
YOY_CHG = 0
Dim PER_CHG As Long
PER_CHG = 0
Dim Tot_Vol As Double
Tot_Vol = 0

             
'Summary Table Labels
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

'Track location for each Ticker Symbol
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

'Track location for each Ticker Symbol
Dim Map_Table_Row As Integer
Map_Table_Row = 2

'Track location for each Ticker Symbol
Dim Pct_Table_Row As Integer
Pct_Table_Row = 2

FinalRow = Cells(Rows.Count, 1).End(xlUp).Row

'Loop Through Tickers
For i = 2 To FinalRow

'Remove Duplicates
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

'Set Ticker Symbol
Ticker_Symbol = Cells(i, 1).Value

'Add Total Volume
Tot_Vol = Tot_Vol + Cells(i, 7).Value


'Print Symbol
Range("I" & Summary_Table_Row).Value = Ticker_Symbol

'Print Total Volume
Range("L" & Summary_Table_Row).Value = Tot_Vol

'Add one to summary table row
Summary_Table_Row = Summary_Table_Row + 1

'Reset Tot Vol
Tot_Vol = 0
 
 'If same ticker
 Else
 
 'Add to Tot Vol
 Tot_Vol = Tot_Vol + Cells(i, 7).Value
 
 End If
 
 
 Next i
 

Range("A1:G1").Copy Destination:=Range("Z1:AF1")
' Find the last row of data
    FinalRow = Cells(Rows.Count, 1).End(xlUp).Row
    ' Loop through each row
    For x = 2 To FinalRow
        ' Decide if to copy based on column D
        ThisValue = Cells(x, 2).Value
        If ThisValue = "20180102" Then
            Cells(x, 1).Resize(1, 7).Copy
            Sheets("2018").Select
            NextRow = Cells(Rows.Count, 26).End(xlUp).Row + 1
            Cells(NextRow, 26).Select
            ActiveSheet.Paste
            Sheets("2018").Select
        ElseIf ThisValue = "20181231" Then
            Cells(x, 1).Resize(1, 7).Copy
            Sheets("2018").Select
            NextRow = Cells(Rows.Count, 26).End(xlUp).Row + 1
            Cells(NextRow, 26).Select
            ActiveSheet.Paste
            Sheets("2018").Select
        End If
                     
        Next x
        
        
        
'Loop Through Tickers
For j = 2 To FinalRow

'Remove Duplicates
If Cells(j + 1, 26).Value <> Cells(j, 26).Value Then

'Set Ticker Symbol
Ticker_Map = Cells(j, 26).Value

'Difference Year over Year
YOY_CHG = YOY_CHG + Cells(j, 31).Value

'Print Yearly Change
Range("J" & Map_Table_Row).Value = YOY_CHG
Range("AX" & Map_Table_Row).Value = YOY_CHG

'Add one to summary table row
Map_Table_Row = Map_Table_Row + 1

'Reset Tot Vol
YOY_CHG = 0

 'If same ticker
 Else

 'Add To Yearly Change
 YOY_CHG = YOY_CHG - Cells(j, 31).Value

 End If
 

Next j

For j = 2 To FinalRow

If Cells(j, 10).Value > 0 Then

Cells(j, 10).Interior.ColorIndex = 4

ElseIf Cells(j, 10).Value < 0 Then

Cells(j, 10).Interior.ColorIndex = 3

End If
 
 Next j

Range("Z1:AF1").Copy Destination:=Range("AH1:AN1")
' Find the last row of data
    FinalRow = Cells(Rows.Count, 26).End(xlUp).Row
    ' Loop through each row
    For m = 2 To FinalRow
        ' Decide if to copy based on column D
        ThisValue = Cells(m, 27).Value
        If ThisValue = 20180102 Then
            Cells(m, 26).Resize(1, 7).Copy
            Sheets("2018").Select
            NextRow = Cells(Rows.Count, 34).End(xlUp).Row + 1
            Cells(NextRow, 34).Select
            ActiveSheet.Paste
            Sheets("2018").Select
        End If
                     
        Next m
        
  
Range("Z1:AF1").Copy Destination:=Range("AP1:AV1")
' Find the last row of data
    FinalRow = Cells(Rows.Count, 26).End(xlUp).Row
    ' Loop through each row
    For n = 2 To FinalRow
        ' Decide if to copy based on column D
        ThisValue = Cells(n, 27).Value
        If ThisValue = 20181231 Then
            Cells(n, 26).Resize(1, 7).Copy
            Sheets("2018").Select
            NextRow = Cells(Rows.Count, 42).End(xlUp).Row + 1
            Cells(NextRow, 42).Select
            ActiveSheet.Paste
            Sheets("2018").Select
        End If
                  Cells(n, 39).Copy
                  Sheets("2018").Select
            NextRow = Cells(Rows.Count, 51).End(xlUp).Row + 1
            Cells(NextRow, 51).Select
            ActiveSheet.Paste
            Sheets("2018").Select
        Next n
 
'End Sub
'
'Sub Perchange()



Dim subvar As Double
Dim oldval As Double
Dim Perchange As String

FinalRow = Cells(Rows.Count, 47).End(xlUp).Row

'Track location for each Ticker Symbol
Dim Temp_Table_Row As Integer
Temp_Table_Row = 2


'Loop Through Tickers
For o = 2 To FinalRow


'Difference Year over Year
Perchange = (Cells(o, 47).Value - Cells(o, 39).Value) / (Cells(o, 39).Value)

'Print Yearly Change
Range("K" & Temp_Table_Row).Value = Perchange



'Add one to summary table row
Temp_Table_Row = Temp_Table_Row + 1

'Reset Tot Vol
Perchange = 0


 'Add To Yearly Change



Next o

For o = 2 To FinalRow
Range("K" & o).Value = FormatPercent(Range("K" & o))
Next o




Cells(ActiveWindow.SplitRow + 1, ActiveWindow.SplitColumn + 1).Select
 
 Range("Z:AY").Clear
 
 
 End Sub
 
 Sub Tickers2019():


'Set Variables
Dim Ticker_Symbol As String
Dim Ticker_Map As String
Dim Trade_Date As Date
Dim Open_Price As Double
Dim High_Price As Double
Dim Low_Price As Double
Dim Close_Price As Double
Dim Volume As Double

'Set initial Variables
Dim YOY_CHG As Double
YOY_CHG = 0
Dim PER_CHG As Long
PER_CHG = 0
Dim Tot_Vol As Double
Tot_Vol = 0

             
'Summary Table Labels
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

'Track location for each Ticker Symbol
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

'Track location for each Ticker Symbol
Dim Map_Table_Row As Integer
Map_Table_Row = 2

'Track location for each Ticker Symbol
Dim Pct_Table_Row As Integer
Pct_Table_Row = 2

FinalRow = Cells(Rows.Count, 1).End(xlUp).Row

'Loop Through Tickers
For i = 2 To FinalRow

'Remove Duplicates
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

'Set Ticker Symbol
Ticker_Symbol = Cells(i, 1).Value

'Add Total Volume
Tot_Vol = Tot_Vol + Cells(i, 7).Value


'Print Symbol
Range("I" & Summary_Table_Row).Value = Ticker_Symbol

'Print Total Volume
Range("L" & Summary_Table_Row).Value = Tot_Vol

'Add one to summary table row
Summary_Table_Row = Summary_Table_Row + 1

'Reset Tot Vol
Tot_Vol = 0
 
 'If same ticker
 Else
 
 'Add to Tot Vol
 Tot_Vol = Tot_Vol + Cells(i, 7).Value
 
 End If
 
 
 Next i
 

Range("A1:G1").Copy Destination:=Range("Z1:AF1")
' Find the last row of data
    FinalRow = Cells(Rows.Count, 1).End(xlUp).Row
    ' Loop through each row
    For x = 2 To FinalRow
        ' Decide if to copy based on column D
        ThisValue = Cells(x, 2).Value
        If ThisValue = "20190102" Then
            Cells(x, 1).Resize(1, 7).Copy
            Sheets("2019").Select
            NextRow = Cells(Rows.Count, 26).End(xlUp).Row + 1
            Cells(NextRow, 26).Select
            ActiveSheet.Paste
            Sheets("2019").Select
        ElseIf ThisValue = "20191231" Then
            Cells(x, 1).Resize(1, 7).Copy
            Sheets("2019").Select
            NextRow = Cells(Rows.Count, 26).End(xlUp).Row + 1
            Cells(NextRow, 26).Select
            ActiveSheet.Paste
            Sheets("2019").Select
        End If
                     
        Next x
        
        
        
'Loop Through Tickers
For j = 2 To FinalRow

'Remove Duplicates
If Cells(j + 1, 26).Value <> Cells(j, 26).Value Then

'Set Ticker Symbol
Ticker_Map = Cells(j, 26).Value

'Difference Year over Year
YOY_CHG = YOY_CHG + Cells(j, 31).Value

'Print Yearly Change
Range("J" & Map_Table_Row).Value = YOY_CHG
Range("AX" & Map_Table_Row).Value = YOY_CHG

'Add one to summary table row
Map_Table_Row = Map_Table_Row + 1

'Reset Tot Vol
YOY_CHG = 0

 'If same ticker
 Else

 'Add To Yearly Change
 YOY_CHG = YOY_CHG - Cells(j, 31).Value

 End If
 

Next j

For j = 2 To FinalRow

If Cells(j, 10).Value > 0 Then

Cells(j, 10).Interior.ColorIndex = 4

ElseIf Cells(j, 10).Value < 0 Then

Cells(j, 10).Interior.ColorIndex = 3

End If
 
 Next j

Range("Z1:AF1").Copy Destination:=Range("AH1:AN1")
' Find the last row of data
    FinalRow = Cells(Rows.Count, 26).End(xlUp).Row
    ' Loop through each row
    For m = 2 To FinalRow
        ' Decide if to copy based on column D
        ThisValue = Cells(m, 27).Value
        If ThisValue = 20190102 Then
            Cells(m, 26).Resize(1, 7).Copy
            Sheets("2019").Select
            NextRow = Cells(Rows.Count, 34).End(xlUp).Row + 1
            Cells(NextRow, 34).Select
            ActiveSheet.Paste
            Sheets("2019").Select
        End If
                     
        Next m
        
  
Range("Z1:AF1").Copy Destination:=Range("AP1:AV1")
' Find the last row of data
    FinalRow = Cells(Rows.Count, 26).End(xlUp).Row
    ' Loop through each row
    For n = 2 To FinalRow
        ' Decide if to copy based on column D
        ThisValue = Cells(n, 27).Value
        If ThisValue = 20191231 Then
            Cells(n, 26).Resize(1, 7).Copy
            Sheets("2019").Select
            NextRow = Cells(Rows.Count, 42).End(xlUp).Row + 1
            Cells(NextRow, 42).Select
            ActiveSheet.Paste
            Sheets("2019").Select
        End If
                  Cells(n, 39).Copy
                  Sheets("2019").Select
            NextRow = Cells(Rows.Count, 51).End(xlUp).Row + 1
            Cells(NextRow, 51).Select
            ActiveSheet.Paste
            Sheets("2019").Select
        Next n
 
'End Sub
'
'Sub Perchange()



Dim subvar As Double
Dim oldval As Double
Dim Perchange As String

FinalRow = Cells(Rows.Count, 47).End(xlUp).Row

'Track location for each Ticker Symbol
Dim Temp_Table_Row As Integer
Temp_Table_Row = 2


'Loop Through Tickers
For o = 2 To FinalRow


'Difference Year over Year
Perchange = (Cells(o, 47).Value - Cells(o, 39).Value) / (Cells(o, 39).Value)

'Print Yearly Change
Range("K" & Temp_Table_Row).Value = Perchange



'Add one to summary table row
Temp_Table_Row = Temp_Table_Row + 1

'Reset Tot Vol
Perchange = 0


 'Add To Yearly Change



Next o

For o = 2 To FinalRow
Range("K" & o).Value = FormatPercent(Range("K" & o))
Next o




Cells(ActiveWindow.SplitRow + 1, ActiveWindow.SplitColumn + 1).Select
 
 Range("Z:AY").Clear
 
 
 End Sub
 
 Sub Tickers2020():


'Set Variables
Dim Ticker_Symbol As String
Dim Ticker_Map As String
Dim Trade_Date As Date
Dim Open_Price As Double
Dim High_Price As Double
Dim Low_Price As Double
Dim Close_Price As Double
Dim Volume As Double

'Set initial Variables
Dim YOY_CHG As Double
YOY_CHG = 0
Dim PER_CHG As Long
PER_CHG = 0
Dim Tot_Vol As Double
Tot_Vol = 0

             
'Summary Table Labels
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

'Track location for each Ticker Symbol
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

'Track location for each Ticker Symbol
Dim Map_Table_Row As Integer
Map_Table_Row = 2

'Track location for each Ticker Symbol
Dim Pct_Table_Row As Integer
Pct_Table_Row = 2

FinalRow = Cells(Rows.Count, 1).End(xlUp).Row

'Loop Through Tickers
For i = 2 To FinalRow

'Remove Duplicates
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

'Set Ticker Symbol
Ticker_Symbol = Cells(i, 1).Value

'Add Total Volume
Tot_Vol = Tot_Vol + Cells(i, 7).Value


'Print Symbol
Range("I" & Summary_Table_Row).Value = Ticker_Symbol

'Print Total Volume
Range("L" & Summary_Table_Row).Value = Tot_Vol

'Add one to summary table row
Summary_Table_Row = Summary_Table_Row + 1

'Reset Tot Vol
Tot_Vol = 0
 
 'If same ticker
 Else
 
 'Add to Tot Vol
 Tot_Vol = Tot_Vol + Cells(i, 7).Value
 
 End If
 
 
 Next i
 

Range("A1:G1").Copy Destination:=Range("Z1:AF1")
' Find the last row of data
    FinalRow = Cells(Rows.Count, 1).End(xlUp).Row
    ' Loop through each row
    For x = 2 To FinalRow
        ' Decide if to copy based on column D
        ThisValue = Cells(x, 2).Value
        If ThisValue = "20200102" Then
            Cells(x, 1).Resize(1, 7).Copy
            Sheets("2020").Select
            NextRow = Cells(Rows.Count, 26).End(xlUp).Row + 1
            Cells(NextRow, 26).Select
            ActiveSheet.Paste
            Sheets("2020").Select
        ElseIf ThisValue = "20201231" Then
            Cells(x, 1).Resize(1, 7).Copy
            Sheets("2020").Select
            NextRow = Cells(Rows.Count, 26).End(xlUp).Row + 1
            Cells(NextRow, 26).Select
            ActiveSheet.Paste
            Sheets("2020").Select
        End If
                     
        Next x
        
        
        
'Loop Through Tickers
For j = 2 To FinalRow

'Remove Duplicates
If Cells(j + 1, 26).Value <> Cells(j, 26).Value Then

'Set Ticker Symbol
Ticker_Map = Cells(j, 26).Value

'Difference Year over Year
YOY_CHG = YOY_CHG + Cells(j, 31).Value

'Print Yearly Change
Range("J" & Map_Table_Row).Value = YOY_CHG
Range("AX" & Map_Table_Row).Value = YOY_CHG

'Add one to summary table row
Map_Table_Row = Map_Table_Row + 1

'Reset Tot Vol
YOY_CHG = 0

 'If same ticker
 Else

 'Add To Yearly Change
 YOY_CHG = YOY_CHG - Cells(j, 31).Value

 End If
 

Next j

For j = 2 To FinalRow

If Cells(j, 10).Value > 0 Then

Cells(j, 10).Interior.ColorIndex = 4

ElseIf Cells(j, 10).Value < 0 Then

Cells(j, 10).Interior.ColorIndex = 3

End If
 
 Next j

Range("Z1:AF1").Copy Destination:=Range("AH1:AN1")
' Find the last row of data
    FinalRow = Cells(Rows.Count, 26).End(xlUp).Row
    ' Loop through each row
    For m = 2 To FinalRow
        ' Decide if to copy based on column D
        ThisValue = Cells(m, 27).Value
        If ThisValue = 20200102 Then
            Cells(m, 26).Resize(1, 7).Copy
            Sheets("2020").Select
            NextRow = Cells(Rows.Count, 34).End(xlUp).Row + 1
            Cells(NextRow, 34).Select
            ActiveSheet.Paste
            Sheets("2020").Select
        End If
                     
        Next m
        
  
Range("Z1:AF1").Copy Destination:=Range("AP1:AV1")
' Find the last row of data
    FinalRow = Cells(Rows.Count, 26).End(xlUp).Row
    ' Loop through each row
    For n = 2 To FinalRow
        ' Decide if to copy based on column D
        ThisValue = Cells(n, 27).Value
        If ThisValue = 20201231 Then
            Cells(n, 26).Resize(1, 7).Copy
            Sheets("2020").Select
            NextRow = Cells(Rows.Count, 42).End(xlUp).Row + 1
            Cells(NextRow, 42).Select
            ActiveSheet.Paste
            Sheets("2020").Select
        End If
                  Cells(n, 39).Copy
                  Sheets("2020").Select
            NextRow = Cells(Rows.Count, 51).End(xlUp).Row + 1
            Cells(NextRow, 51).Select
            ActiveSheet.Paste
            Sheets("2020").Select
        Next n
 
'End Sub
'
'Sub Perchange()



Dim subvar As Double
Dim oldval As Double
Dim Perchange As String

FinalRow = Cells(Rows.Count, 47).End(xlUp).Row

'Track location for each Ticker Symbol
Dim Temp_Table_Row As Integer
Temp_Table_Row = 2


'Loop Through Tickers
For o = 2 To FinalRow


'Difference Year over Year
Perchange = (Cells(o, 47).Value - Cells(o, 39).Value) / (Cells(o, 39).Value)

'Print Yearly Change
Range("K" & Temp_Table_Row).Value = Perchange



'Add one to summary table row
Temp_Table_Row = Temp_Table_Row + 1

'Reset Tot Vol
Perchange = 0


 'Add To Yearly Change



Next o

For o = 2 To FinalRow
Range("K" & o).Value = FormatPercent(Range("K" & o))
Next o




Cells(ActiveWindow.SplitRow + 1, ActiveWindow.SplitColumn + 1).Select
 
 Range("Z:AY").Clear
 
 
 End Sub
 


