Attribute VB_Name = "Module1"
Sub Main()
    Dim wbk As Workbook, wksht As Worksheet, sheetCount As Integer
    
    ' Set the name of the workbook the script will run on
    Set wbk = Application.Workbooks("Multiple_year_stock_data.xlsm")
    
    ' Run the script for every sheet in wbk
    sheetCount = wbk.Worksheets.Count
    For i = 1 To sheetCount
        Set wksht = wbk.Worksheets(i)
        Call ClearCells(wksht)
        Call SortColumns(wksht)
        Call SetHeader(wksht)
        Call Run(wksht)
    Next i
End Sub

Sub SortColumns(wksht As Worksheet)
    'Sort the data By ticker and then by date
    Dim data As Range
    With wksht.Range("A:G")
    .Cells.Sort Key1:=.Columns("A"), Order1:=xlAscending, _
                Key2:=.Columns("B"), Order2:=xlAscending, _
                Orientation:=xlTopToBottom, Header:=xlYes
    End With
End Sub

Sub ClearCells(wksht As Worksheet)
    ' Clears the cells from any previous run
    wksht.Range("I1:P10000").Value = ""
    wksht.Range("I1:P10000").Interior.ColorIndex = 0
End Sub

Sub SetHeader(wksht As Worksheet)
    ' Set the headers for the analysis
    wksht.Range("I1").Value = "Ticker"
    wksht.Range("J1").Value = "Yearly Change"
    wksht.Range("K1").Value = "Percent Change"
    wksht.Range("L1").Value = "Total Stock Volume"
End Sub

Sub Run(wksht As Worksheet)
    ' Declare varibles
    Dim currTicker As String, initialClose As Double, endClose As Double, volume As Double, currRow As Long, writeRow As Long, percentChange As Double
    Dim change As Double
    
    ' Arrays to track value and name for Leaders. [Greatest increase, Greatest Decrease, Greatest Volume]
    Dim greatestValues(2) As Double, greatestNames(2) As String
    
    'Set inital values
    currRow = 3
    writeRow = 2
    currTicker = wksht.Range("A" & 2).Value
    initialClose = wksht.Range("F" & 2).Value
    volume = wksht.Range("G" & 2).Value
    Debug.Print ("Very initial: " & initialClose)

    While Not IsEmpty(wksht.Cells(currRow, 1))
        If wksht.Cells(currRow, 1).Value = currTicker Then
            volume = volume + wksht.Cells(currRow, 7)
            currRow = currRow + 1
        Else:
            ' Write the values
            'MsgBox ("Curr Row: " & currRow & " Inital close: " & initialClose)
            'Debug.Print ("Curr Row: " & currRow & " Inital close: " & initialClose)
            endClose = wksht.Range("F" & currRow - 1).Value
            'Call WriteRows(wksht, currTicker, currWriteRow, initialClose, wksht.Cells(currRow - 1, 6).Value, volume)
            change = endClose - initialClose
            wksht.Range("I" & writeRow).Value = currTicker
            ' Apply conditional formatting
            wksht.Range("J" & writeRow).Value = change
            If change > 0 Then
                wksht.Range("J" & writeRow).Interior.ColorIndex = 4
            ElseIf change < 0 Then
                wksht.Range("J" & writeRow).Interior.ColorIndex = 3
            End If
            
            ' Calculate and forma the percentage rounded to the neares two dec (for the %)
            ' Check to see if the initalClose = 0 to avoid dividing by 0
            If initialClose = 0 Then
                percentChange = 0
            Else:
                percentChange = Round(change / initialClose, 4)
            End If
            wksht.Range("K" & writeRow).Value = percentChange
            wksht.Range("K" & writeRow).NumberFormat = "0.00%"
            'Write the volume
            wksht.Range("L" & writeRow).Value = Round(volume, 0)
            
            ' Update the leaderboard
            ' Check if it's the first time through
            If writeRow = 2 Then
                greatestValues(0) = percentChange
                greatestValues(1) = percentChange
                greatestValues(2) = percentChange
                For i = 0 To 2
                    greatestNames(i) = currTicker
                Next i
            Else:
                ' Is it the greatest increase in percent change?
                If percentChange > greatestValues(0) Then
                    greatestValues(0) = percentChange
                    greatestNames(0) = currTicker
                ' Is it the greatest decrease in percent change?
                ElseIf percentChange < greatestValues(1) Then
                    greatestValues(1) = percentChange
                    greatestNames(1) = currTicker
                End If
                
                'Is it the greatest volume
                If volume > greatestValues(2) Then
                    greatestValues(2) = volume
                    greatestNames(2) = currTicker
                End If
            End If
            
            ' Reset values for new Ticker
            currTicker = wksht.Range("A" & currRow)
            initialClose = wksht.Range("F" & currRow)
            volume = wksht.Range("G" & currRow)
            writeRow = writeRow + 1
            currRow = currRow + 1
        End If
    Wend
    
    ' Write the leaderboard
    wksht.Range("O1").Value = "Ticker"
    wksht.Range("P1").Value = "Value"
    wksht.Range("N2").Value = "Greatest % Increase"
    wksht.Range("N3").Value = "Greatest % Decrease"
    wksht.Range("N4").Value = "Greatest Total Volume"
    
    ' The format for the percentages is different
    For i = 0 To 1
        wksht.Range("O" & i + 2).Value = greatestNames(i)
        wksht.Range("P" & i + 2).Value = greatestValues(i)
        wksht.Range("P" & i + 2).NumberFormat = "0.00%"
    Next i
    
    ' Write the greatest volume
    wksht.Range("O4").Value = greatestNames(i)
    wksht.Range("P4").Value = greatestValues(i)
    
    ' Autofit the analysis columns
    wksht.Columns("I:P").AutoFit
    
    
End Sub


