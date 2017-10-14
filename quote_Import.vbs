Sub quote_Import()

        Dim wrkngBk As Workbook
        Dim wrkngSht As Worksheet
        Dim sellDisc As String 'for storing sell discount
        Dim costDisc As String 'for storing cost discount
        Dim lRow As Long 'for finding last row
        Const sRow As Long = 1 'What's the first row with data?
        
        'Turn off the annoying stuff
        With Application
            .ScreenUpdating = False 'turn screen refreshing off
            .DisplayAlerts = False  'turn system alerts off
            .EnableEvents = False   'turn other macros off
        End With
        
        Set wrkngBk = ActiveWorkbook
        Set wrkngSht = wrkngBk.ActiveSheet
        
        'ask user what sell discount is
        sellDisc = Application.InputBox(Prompt:="What discount(off of list) are we selling at?", Type:=1 + 8)
        
        'ask user what cost discount is
        costDisc = Application.InputBox(Prompt:="What discount(off of list) are we getting from disty?", Type:=1 + 8)
        
        'delete rows up to header row and then 2 rows after
        With wrkngSht
            .Shapes.Range(Array("Picture 2")).Delete
            .Rows("1:10").Delete
            .Rows("2:3").Delete
        End With
        
        With wrkngSht
            lRow = .Cells(.Rows.Count, "B").End(xlUp).Row
            For i = lRow To sRow Step -1
                If .Cells(i, "C").Value = "" Then
                    .Cells(i, "C").EntireRow.Delete
                End If
            Next i
        End With
            
        With wrkngSht
            lRow = .Cells(.Rows.Count, "H").End(xlUp).Row
            For i = lRow To sRow Step -1
                If .Cells(i, "B").Value = "" Then
                    .Cells(i, "B").EntireRow.Delete
                End If
            Next i
        End With
        
        With wrkngSht
            lRow = .Cells(.Rows.Count, "H").End(xlUp).Row
            .Rows(lRow).Offset(-1).Resize(3).EntireRow.Delete
        End With
        
        'find strings with brackets and delete brackets and whatevers inbetween
        With wrkngSht
            Cells.Replace What:="[*]", Replacement:="", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False
        End With
        
        'enter sell discount into column e
        For Each c In Range("E14:E" & lRow)
            If Range("D14:D" & lRow) Is Nothing Then
                c.Value = "0"
            Else
                If c.Offset(, -1) > "1" Then
                    c.Value = sellDisc
                End If
            End If
            If c.Offset(, -1) = "Included" Then
                c.Value = "0.00"
            End If
        Next c
        
        'add column i for our cost from disty
        Range("C1").Value = "Our Cost"
        If costDisc > 0 Then
            Range("C2:C" & lRow).Formula = "=if(isnumber(d2),d2-(d2*(" & costDisc & "/100)),0)"
        End If
        
        'turn back on the annoying stuff
        With Application
            .DisplayAlerts = True   'turn system alerts back on
            .EnableEvents = True    'turn other macros back on
            .ScreenUpdating = True  'refreshes the screen
        End With
End Sub
