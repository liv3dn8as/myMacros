Sub IRIS_Edit()
'
' IRIS_Edit Macro
'

'
    Dim theWB As Workbook 'workbook variable
    Dim theWS As Worksheet 'worksheet variable
    Dim fName As String 'for storing the xml file name
    Dim listObj As ListObject ' For determining if table1 exist
    Dim rList As Range ' For removing Table and converting to Range
    Dim lRow, lRow1 As Long ' For storing the last row when filling in Discounts and other columns
    Dim i As Long 'for inserting blank rows after each change in Column A
    Dim c, sRange1, sRange2 As Range 'for adding actual subtotals to each section
    'Dim theRng As Range
    Dim Cell As Range
    Dim NR As Long
    Dim theR11 As Range 'for formatting the header row, row 11
    Dim thePTRC, thePTRCR, thePTRH As Range 'project total variables
    Dim n, q As Long, x As String
    Dim thePic As Shape 'for inserting PDS logo
    'n = 1
    
    Set theWB = ActiveWorkbook
    Set theWS = theWB.Sheets("Sheet1")
    
    'Turn off screen updating
    Application.ScreenUpdating = False
    
    'Set Window Size and remove grid lines
    With ActiveWindow
        .Zoom = 86
        .DisplayGridlines = False
    End With
    
    ' Get rid of the table looking stuff if it exists
    'If listObj Is Nothing Then
    '    Rows("1").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    '    Else
        Set listObj = theWS.ListObjects("Table1")
            With listObj
                Set rList = .Range
                .Unlist
            End With
    'End If
    
    With rList
       .Interior.ColorIndex = xlColorIndexNone
       .Font.ColorIndex = xlColorIndexAutomatic
       .Borders.LineStyle = xlLineStyleNone
       .Font.Name = "Calibri"
       .Font.Size = 10
    End With
    'End If
    
    ' Delete a group of unneeded columns
    Range("B:B, E:E, I:K").Delete

    ' Start Changing Headers
    Range("B1").Value = "Part Number"
    Range("C1").Value = "Description"
    Range("D1").Value = "Unit Price"
    Columns("E:F").Insert Shift:=xlToRight
    Range("E1").Value = "Discount %"
    Range("F1").Value = "Discounted Price"
    Range("G1").Value = "Qty"
    Range("H1").Value = "Extended Price"

    lRow = Range("D" & Rows.Count).End(xlUp).Row
    
    'Fill in Discount%, Discounted Price, & ExtPrice Columns
    Range("C2:C" & lRow).WrapText = True
    With Range("D2:D" & lRow)
        .HorizontalAlignment = xlRight
        .NumberFormat = "0.00"
    End With
    Range("E2") = "0.00": Range("E2:E" & lRow).FillDown
    Range("E2:E" & lRow).HorizontalAlignment = xlCenter
    Range("F2:F" & lRow).NumberFormat = "0.00"
Range("F2") = "=iferror(d2-(d2*(e2/100)),0)": Range("F2:F" & lRow).FillDown
    Range("H2:H" & lRow).NumberFormat = "$#,##0.00"
    Range("H2") = "=f2*g2": Range("H2:H" & lRow).FillDown
    Range("H2:H" & lRow).HorizontalAlignment = xlRight
    Range("G1:G" & lRow).HorizontalAlignment = xlCenter
    
    'Insert Blank Rows After Each Change In Column A
    For i = lRow To 2 Step -1
        If Cells(i - 1, "A").Value <> Cells(i, "A").Value Then
            Cells(i, "A").Resize(4).EntireRow.Insert
            'Range("B" & i + 2).Value2 = Range("A" & i + 4).Value2
                 With Range("B" & i + 2)
                    .Font.Bold = True
                    .Font.Size = 11
                    .Font.Color = RGB(0, 32, 96)
                    .HorizontalAlignment = xlLeft
                    .Value2 = Range("A" & i + 4).Value2
                End With
                With Range("B" & i + 3)
                    .Font.Bold = True
                    .Font.Size = 11
                    .Font.Color = RGB(51, 102, 255)
                    .HorizontalAlignment = xlLeft
                    .Value2 = "Materials"
                End With
        End If
    Next i
    
    'section to try and add identification in col M for subtotaling sections etc
    lRow1 = Cells(Rows.Count, "B").End(xlUp).Row
    Set sRange1 = Range("M2:M" & lRow1)
    Set sRange2 = Range("H2:H" & lRow1)
    
    For Each c In Range("H2:H" & lRow1)
    If Not c.Value <> "" Then
        c.Offset(, 5).Value = "0"
    Else
        c.Offset(, 5).Value = "1"
    lRow1 = lRow1 + 1
    End If
    Next c

    'setting project total variable
    myVal = Application.WorksheetFunction.SumIf(sRange1, "=1", sRange2)

    'Delete some unneeded rows
    Rows("2:3").Delete Shift:=xlUP

    'insert rows 1 through 10
    theWS.Rows(1).Resize(10).EntireRow.Insert
    
    'hide row 9    
    Rows("9").EntireRow.Hidden = True
      
    'Set the column widths
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
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
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
    For Each Cell In Range("C14", Range("C65536").End(xlUp))
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
    
    findMaterials_IRIS 'Move to Module for formatting the subtotal row

    'Create the Project Total Row
    Set thePTRC = Range("C65536").End(xlUp).Offset(2)
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
    Set thePTRH = Range("H65536").End(xlUp).Offset(3)
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
        .Value = myVal
    End With
    
    ' Add Border Above Project Total Row
    NR = Range("B" & Rows.Count).End(xlUp).Row + 3
    With Range("B" & NR, "H" & NR)
        .Borders(xlEdgeTop).Weight = xlThin
    End With
    
    'Insert PDS Logo
    Set thePic = theWS.Shapes.AddPicture("C:\Users\tcoplien\Desktop\SMARTnet\pdsLogo.png", _
                 msoFalse, msoTrue, 15, 10, 74, 80)
    
    'Rename the tab
    theWS.Name = "Report"

    'hide column m
    theWS.Range("M:M").EntireColumn.Hidden = True
    
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
        
    'Turn on screen updating
    Application.ScreenUpdating = True
    
End Sub
