Sub IRIS_Edit_NEW()
'
' IRIS_Edit_NEW Macro
'
    Dim theWB As Workbook
    Dim theWS As Worksheet
    Dim lRow As Long 'for storing the last row variable
    Dim theDN, theST As String 'for storing Discount and SubTotal variable
    Dim c As Range 'for adding the 1 and 2 in column m for total
    Dim sectionTitle As Range 'for moving the section title over a cell to the right
    Dim theMS, foundMS As Range 'for doing the subtotals
    Dim fAdd, msAdd As String 'same as above
    Dim thePIC As Variant 'for putting in logo

    'Turn off screen updating & alerts
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    'set worksheet variable and rename the tab
    Set theWB = ActiveWorkbook
    Set theWS = theWB.ActiveSheet
    theWS.Name = "Report"
    
    'remove all tabs besides Report
    Do While Sheets.Count > Sheets("Report").Index
        Sheets(Sheets("Report").Index + 1).Delete
    Loop
        
    'Set Window Size and remove grid lines
    With ActiveWindow
        .Zoom = 86
        .DisplayGridlines = False
    End With
        
    'make all font uniform
    With theWS.Cells
       .Interior.ColorIndex = xlColorIndexNone
       .Font.ColorIndex = xlColorIndexAutomatic
       .Borders.LineStyle = xlLineStyleNone
       .Font.Name = "Calibri"
       .Font.Size = 10
    End With
        
    'find last row; offset(row column)
    lRow = Cells(Rows.Count, "B").End(xlUp).Row
    
    'move columns around
    theDN = "Discount"
    For Each c In Range("A1", Range("A1").End(xlToRight))
    If c.Value = theDN Then
        c.EntireColumn.Cut
        Range("E1").Insert Shift:=xlToRight
    End If
    Next c
    
    theST = "Sub Total"
    For Each c In Range("A1", Range("A1").End(xlToRight))
    If c.Value = theST Then
        c.EntireColumn.Cut
        Range("F1").Insert Shift:=xlToRight
    End If
    Next c
        
    'move section title over one column
    For Each sectionTitle In Range("A2:A" & lRow)
        If sectionTitle Like "*. *" Then
            sectionTitle.Copy sectionTitle.Offset(0, 1)
            sectionTitle.Offset(0, 1).Font.Bold = True
            sectionTitle.Offset(0, 1).Font.Size = 11
            sectionTitle.Offset(0, 1).Font.Color = RGB(0, 32, 96)
        End If
    Next sectionTitle
        
    'try and add blank lines after each section
    For i = lRow To 3 Step -1
        If Cells(i - 1, "C").Value = "" Then
            Cells(i, "C").Resize(1).EntireRow.Insert
            Cells(i - 1, "C").Resize(2).EntireRow.Insert
            Cells(i + 2, "B").Value = "Materials"
            Cells(i + 2, "B").Font.Bold = True
            Cells(i + 2, "B").Font.Color = RGB(51, 102, 205)
        End If
    Next i
        
    'clear column A and set the column widths
    With Columns("A")
        .ClearContents
        .ColumnWidth = 1.29
    End With
    Columns("B").ColumnWidth = 25
    Columns("C").ColumnWidth = 56.29
    Columns("D").ColumnWidth = 13.57
    Columns("E").ColumnWidth = 12.29
    Columns("F").ColumnWidth = 15.57
    Columns("G").ColumnWidth = 8.14
    Columns("H").ColumnWidth = 17.29
    Columns("I").ColumnWidth = 4.43
        
    'change headers
    Range("E1").Value = "Discount %"
    Range("F1").Value = "Discounted Price"
    Range("G1").Value = "Qty"
    Range("H1").Value = "Extended Price"
        
    'adjust format of Discount%, Discounted Price, & ExtPrice Columns
    lRow = Cells(Rows.Count, "B").End(xlUp).Row 'recalculate
    Range("C6:C" & lRow).WrapText = True
    With Range("D6:D" & lRow)
        .HorizontalAlignment = xlRight
        .NumberFormat = "$#,##0.00"
    End With
    With Range("E6:E" & lRow)
        .HorizontalAlignment = xlCenter
        .Value = Value / 100
        .NumberFormat = "0.00"
        .Font.Color = RGB(0, 0, 0)
    End With
    With Range("F6:F" & lRow)
        .Value = "=iferror(d6-(d6*(e6/100)),0)"
        .NumberFormat = "$#,##0.00"
        .HorizontalAlignment = xlRight
    End With
    Range("G6:G" & lRow).HorizontalAlignment = xlCenter
    With Range("H6:H" & lRow)
        .HorizontalAlignment = xlRight
        .NumberFormat = "$#,##0.00"
        .Value = "=f6*g6"
    End With
    
    'now get rid of all the extra 0.00's
    For Each c In Range("C6:C" & lRow)
        If c.Value = "" Then
            c.Offset(, 2).ClearContents
            c.Offset(, 3).ClearContents
            c.Offset(, 5).ClearContents
        End If
    Next c
    
    'delete some unneeded rows that got added above
    Rows("2:3").Delete Shift:=xlUp
        
    'insert rows 1 through 10
    theWS.Rows(1).Resize(10).EntireRow.Insert
    
    'hide row 9
    Rows("9").EntireRow.Hidden = True
        
    'Define the Customer Contact Variable and format cell
    With theWS.Range("B7")
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .Font.Bold = True
        .Value = "Customer Contact"
    End With
    
    'Define the Date Variable and format cell
    With theWS.Range("B8")
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .Font.Bold = True
        .Value = "=TODAY()"
    End With
    
    'for placing the pricing expires after 30 days line
    With Range("C7")
        .Value = "**Pricing Expires After 30 Days**"
        .Font.Size = 10
    End With
    
    'For placing the Customer Name
    With theWS.Range("C3:D3")
        .HorizontalAlignment = xlCenter
        .MergeCells = True
        .Value = "Customer Name"
        .Font.Size = 16
        .Font.Bold = True
    End With

    'For placing the BoM Description
    With theWS.Range("C4:D4")
        .HorizontalAlignment = xlCenter
        .MergeCells = True
        .Font.Size = 11
        .Value = "BoM Description"
    End With
    
    'Add the borders around the Header Row and Change the Color
    Set theR11 = Range("B11", Range("B11").End(xlToRight))
    With theR11.Interior
        .Color = RGB(191, 191, 191)
    End With
    With theR11.Font
        .Size = 11
        .Bold = True
    End With
    With theR11.Borders
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    theR11.Borders(xlEdgeBottom).LineStyle = xlNone
        
    'Add the borders around each section in the BoM
    lRow = Cells(Rows.Count, "B").End(xlUp).Row 'recalculate
    For Each Cell In Range("C14:C" & lRow)
    If Cell.Value > "" Then
        With Range(Cell.Offset(0, -1), Cell.Offset(-1, 5))
             With .Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
        End With
    End If
    Next
        
    'delete the Quote Total row left over from preMacro
    Rows(lRow + 1).EntireRow.Delete
    
    'going to try and meld the findMaterials submodule and the below together
    Set theMS = Columns(2).Find(What:="Materials", LookIn:=xlValues, _
        Lookat:=xlWhole, SearchOrder:=xlByRows, SearchDirection _
        :=xlNext, MatchCase:=False, searchFormat:=False)
    If Not theMS Is Nothing Then
        fAdd = theMS.Address
        Do
            Set foundMS = theMS.End(xlDown).Offset(1, 1)
            With Range(foundMS, foundMS.Offset(0, 5))
                .Interior.Color = RGB(149, 179, 215)
                .Font.Color = vbWhite
                .Font.Bold = True
                .Font.Size = 11
            End With
            With Range(foundMS, foundMS.Offset(0, 4))
                .MergeCells = True
                .HorizontalAlignment = xlRight
                .Value = theMS.Offset(-1) & " Subtotal:"
            End With
            Set theMS = Columns(2).FindNext(theMS)
        Loop While theMS.Address <> fAdd
    End If
    
    'section to try and add identification in col M for subtotaling sections etc
    lRow = Cells(Rows.Count, "C").End(xlUp).Row
    For Each c In Range("H14:H" & lRow)
    If c.Offset(, -5).Value Like "*Subtotal*" Then
        c.Offset(, 5).Value = "2"
    ElseIf c.Value > "" Then
        c.Offset(, 5).Value = "1"
    End If
    Next c
    
    For Each rFind In Range("H13:H" & lRow).SpecialCells(xlCellTypeFormulas, Value:=xlNumbers).Areas
        rFind(rFind.Count + 1).Value = Application.Sum(rFind)
    Next rFind
    
    'correct number format of last subtotal
    With Range("H" & lRow)
        .NumberFormat = "$#,##0.00"
    End With
    
    'hide column m
    theWS.Range("M:M").EntireColumn.Hidden = True
    
    'Create the Project Total Row
    lRow = Cells(Rows.Count, "C").End(xlUp).Row 'recalculate
    Set thePTRC = Range("C" & lRow).Offset(2)
    Set thePTRCR = Range(thePTRC, thePTRC.Offset(0, 4))
    With thePTRCR.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = RGB(0, 51, 204)
    End With
    With thePTRCR.Font
        .Color = vbWhite
        .Size = 14
        .Bold = True
    End With
    With thePTRCR
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .WrapText = False
        .MergeCells = True
        .Value = "Project Total(USD):"
    End With
    
    'This is for the Project Total Total Cell
    Set thePTRH = Range("H" & lRow).Offset(2)
    With thePTRH.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = RGB(0, 51, 204)
    End With
    With thePTRH.Font
        .Color = vbWhite
        .Size = 14
        .Bold = True
    End With
    With thePTRH
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        .Value = Application.SumIf(Range("M14:M" & lRow), "=2", Range("H14:H" & lRow))
    End With
    
    'Add Border Above Project Total Row
    lRow = Cells(Rows.Count, "C").End(xlUp).Row
    With Range("B" & lRow, "H" & lRow)
        .Borders(xlEdgeTop).Weight = xlThin
    End With
    
    'Insert PDS Logo
    Set thePIC = theWS.Shapes.AddPicture("I:\Inside PreSales\Hewlett Packard\Customer  Quotes\pdsLogo2018.png", _
                 msoFalse, msoTrue, 15, 10, -1, -1)

    'Set the print area
    With theWS.PageSetup
        .PrintArea = "$B:$H"
        .PrintTitleRows = ""
            .PrintTitleColumns = ""
            .LeftMargin = Application.InchesToPoints(0.25)
            .RightMargin = Application.InchesToPoints(0.25)
            .TopMargin = Application.InchesToPoints(0.25)
            .BottomMargin = Application.InchesToPoints(0.25)
            .HeaderMargin = Application.InchesToPoints(0.25)
            .FooterMargin = Application.InchesToPoints(0.25)
            .PaperSize = xlPaperLetter
            .Orientation = xlPortrait 'xlLandscape
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False
            .CenterHorizontally = True
            .ScaleWithDocHeaderFooter = True
        End With
        
    'Turn on screen updating & alerts
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub
