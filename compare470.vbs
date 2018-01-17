Sub compare_470()
'put together by twcoplien Jan 2018

    Dim masterWB, dailyWB As Workbook 'setting workbook variables
    Dim mWS, dWS As Worksheet 'setting worksheet variables
    Dim dWBName As String 'for storing full path of daily report as string
    Dim lRow, lRow2 As Long
    Dim c As Range
    
    'storing workbook and worksheet variables and opening Master
    Set dailyWB = ActiveWorkbook
    dWBName = dailyWB.Path & "\470_Detail_Report.xlsx"
    Set masterWB = Workbooks.Open("C:\Users\tcoplien\Desktop\eRate 2018\470_Detail_Report --MASTER--.xlsx")
    Set mWS = masterWB.Sheets(1)
    Set dWS = dailyWB.Sheets(1)
    
    'Turn off the annoying stuff
    With Application
        .ScreenUpdating = False 'turn screen refreshing off
        .DisplayAlerts = False  'turn system alerts off
        .EnableEvents = False   'turn other macros off
    End With
    
    'remove useless columns first
    dWS.Columns("C:D").Delete Shift:=xlToLeft
    dWS.Columns("D:F").Delete Shift:=xlToLeft
    dWS.Columns("G:K").Delete Shift:=xlToLeft
    dWS.Columns("H:I").Delete Shift:=xlToLeft
    dWS.Columns("I:K").Delete Shift:=xlToLeft
    dWS.Columns("O:AR").Delete Shift:=xlToLeft
    dWS.Columns("P:W").Delete Shift:=xlToLeft
    dWS.Columns("Q").Delete Shift:=xlToLeft
    dWS.Columns("T").Delete Shift:=xlToLeft
    dWS.Columns("AB:AI").Delete Shift:=xlToLeft

    'sort Certified Timestamp Column TopToBottom
    dWS.ListObjects("Table1").Sort. _
        SortFields.Clear
    dWS.ListObjects("Table1").Sort. _
        SortFields.add Key:=Range("Table1[[#All],[Certified Timestamp]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
    With dWS.ListObjects("Table1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    lRow = mWS.Cells(Rows.Count, "A").End(xlUp).Row 'finds last row of Master
    lRow1 = dWS.Cells(Rows.Count, "A").End(xlUp).Row 'finds last row of Daily

    'remove any highlights from previous update to Master
    'With mWS.Range("A2:T" & lRow)
    '    .Cells.Interior.Color = xlColorIndexNone
    'End With
    
    'highlighting new records and copying them to Master
    For Each c In dWS.Range("A2:A" & lRow1)
        If Not c.Value = mWS.Cells(c.Row, c.Column).Value Then
            c.EntireRow.Interior.Color = vbYellow 'highlights entire new row in Daily
            c.EntireRow.Copy mWS.Range("A" & lRow + 1) 'copys entire new row to Master
                lRow = lRow + 1 'looks for next new row and repeats last two steps
        End If
    Next c
    
    'close and delete the daily downloaded workbook
    dailyWB.Close SaveChanges:=False
    Kill dWBName
    
    'turn annoying stuff back on
    With Application
        .DisplayAlerts = True   'turn system alerts back on
        .EnableEvents = True    'turn other macros back on
        .ScreenUpdating = True  'refreshes the screen
    End With
    
    mWS.Activate

End Sub
