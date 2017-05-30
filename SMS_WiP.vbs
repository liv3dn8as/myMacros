' The current working macro as of 5/30/2017

Sub SMS_WIP()

    Dim wrkngBk As Workbook
    Dim wrkngSht As Worksheet 'this is the sheet in the downloaded file from CSCC
    Dim orderSht As Worksheet 'this is the sheet that contains the SMS SKU Order info
    Dim nameSht As Worksheet 'this is the temp sheet that will hold the future file name
    Dim theSO As Range, thePN As Range 'for find the word SO, Product Number, etc
    Dim theSS As Range, theBG As Range, thePR As Range
    Dim soRng, pnRng, ssRng, bgRng, prRng As Range 'for selecting the ranges of the above sections
    Dim bBox As Range, bBox2 As Range 'Order Sheet variables
    Dim x As Long 'for setting row heights
    Dim str As String
    Dim str1 As String
    Dim fName As String
    Dim printRange As String 'for setting the print range
    Const startRow As Long = 1 'What's the first row with data?
        
    'Turn off the annoying stuff
    With Application
        .ScreenUpdating = False 'turn screen refreshing off
        .DisplayAlerts = False  'turn system alerts off
        .EnableEvents = False   'turn other macros off
    End With
        
    'Set Window Size and remove Gridlines
    With ActiveWindow
        .Zoom = 86
        .DisplayGridlines = False
    End With
        
    'Set downloaded worksheet variable and rename the tab
    Set wrkngBk = ActiveWorkbook
    Set wrkngSht = wrkngBk.ActiveSheet
    wrkngSht.Name = "Report"
        
    'Set the font and font size
    With wrkngSht
        .Cells.Interior.ColorIndex = xlColorIndexNone
        .Cells.Font.ColorIndex = xlColorIndexAutomatic
        .Cells.Borders.LineStyle = xlLineStyleNone
        .Cells.Font.Name = "Calibri"
        .Cells.Font.Size = 10
    End With
        
    'Find the SMS SKU Section and then copy
    Set theSO = Cells.Find(What:="SO", LookAt:=xlWhole)
    Set soRng = Range(theSO, theSO.End(xlDown).End(xlToRight))
    soRng.Copy
        
    'Create new tab for SMS SKU Section and paste
    With wrkngSht
        Set orderSht = Sheets.Add(After:=wrkngSht)
        With orderSht
            .Name = "SMS Order Info"
            .Tab.ColorIndex = 3 'Changing the Color of the Tab to Red
            .Range("A1").PasteSpecial Paste:=xlPasteValues
        End With
    End With
        
    'Copy cells that will be used for future file name, create new sheet, etc
    With wrkngSht
        .Range("A1:B2").Copy
        Set nameSht = Sheets.Add(After:=orderSht)
        With nameSht
            .Range("A1").PasteSpecial Paste:=xlPasteValues
            .Name = "File Name"
        End With
    End With
    
    wrkngSht.Activate
            
    'Delete rows that are blank in column AP to reduce working data
    With wrkngSht
        lastRow = .Cells(.Rows.Count, "AP").End(xlUp).Row
        For i = lastRow To startRow Step -1
            If .Cells(i, "AP").Value = "" Then
                .Cells(i, "AP").EntireRow.Delete
            End If
        Next i
    End With
        
    'Remove un-needed Data
    With wrkngSht
        .Range("A:G").Delete
        .Range("B:B").Delete
        .Range("C:D").Delete
        .Range("D:J").Delete
        .Range("H:K").Delete
        .Range("J:T").Delete
        .Range("K:AK").Delete
        .Range("L:M").Delete
    End With
        
    'Rearange Columns for better filtering
    Set thePN = wrkngSht.Cells.Find(What:="PRODUCT NUMBER", LookAt:=xlWhole, MatchCase:=True)
    Set pnRng = wrkngSht.Range(thePN, thePN.End(xlDown))
    pnRng.Cut
    wrkngSht.Range("A1").Insert Shift:=xlToRight

    Set theSK = wrkngSht.Cells.Find(What:="SERVICE SKU", LookAt:=xlWhole, MatchCase:=True)
    Set skRng = wrkngSht.Range(theSK, theSK.End(xlDown))
    skRng.Cut

    wrkngSht.Range("C:C").Insert Shift:=xlToRight

    Set theBD = wrkngSht.Cells.Find(What:="BEGIN DATE(DD-MON-YYYY)", LookAt:=xlWhole, MatchCase:=True)
    Set bdRng = wrkngSht.Range(theBD.Offset(0, 1), theBD.End(xlDown))
    bdRng.Cut

    wrkngSht.Range("D1").Insert Shift:=xlToRight

    Set thePR = wrkngSht.Cells.Find(What:="PRO RATED SERVICE NET PRICE", LookAt:=xlWhole, MatchCase:=True)
    Set prRng = wrkngSht.Range(thePR, thePR.End(xlDown))
    prRng.Cut

    wrkngSht.Range("F1").Insert Shift:=xlToRight
    
    wrkngSht.Activate
    
    Call SMSSorting 'Call the sorting module
    'SMSSortingNew 'Call the hopefully new sorting module
    
    'Insert new column A for left border and set column width
    With wrkngSht.Columns("A")
        .Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    End With
    With wrkngSht.Columns("A")
        .ColumnWidth = 1.29
    End With
        
    'Change Heading Row and start setting column properties
    With wrkngSht
    .Range("B1").Value = "Product Name"
        .Columns("B").ColumnWidth = 17.86
    .Range("C1").Value = "Serial Number"
        .Columns("C").ColumnWidth = 13.43
    .Range("D1").Value = "Service SKU"
        .Columns("D").ColumnWidth = 18
    .Range("E1").Value = "Start Date"
    .Range("F1").Value = "End Date"
        .Columns("E:F").ColumnWidth = 10.71
    .Range("G1").Value = "List Price"
    End With
    
    'Bring all qty's to 1
    With wrkngSht.Columns("H")
        .Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    End With
    
    lr = Range("B" & Rows.Count).End(xlUp).Row
    With wrkngSht.Range("H2:H" & lr)
        .Formula = "=G2/I2"
        .Copy
    End With
    
    Range("G2").PasteSpecial Paste:=xlPasteValues
    Range("H:H").Delete
    
    With wrkngSht.Columns("G")
        .ColumnWidth = 10.43
        .NumberFormat = "0.00"
    End With
    
    'Insert Discount and Discounted Price Column
    wrkngSht.Columns("H:I").Insert Shift:=xlToRight, _
        CopyOrigin:=xlFormatFromLeftOrAbove
    
    wrkngSht.Range("H1").Value = "Discount %"
    With wrkngSht.Columns("H")
        .ColumnWidth = 12.29
        .HorizontalAlignment = xlCenter
        .NumberFormat = "0.00"
    End With
    wrkngSht.Range("I1").Value = "Discounted Price"
    With wrkngSht.Columns("I")
        .ColumnWidth = 15.57
        .NumberFormat = "0.00"
        .HorizontalAlignment = xlCenter
    End With
    wrkngSht.Range("J1").Value = "Qty"
    wrkngSht.Columns("J").ColumnWidth = 5.57
    
    'Insert Extended Price Column
    wrkngSht.Columns("K").Insert Shift:=xlToRight, _
        CopyOrigin:=xlFormatFromLeftOrAbove
    wrkngSht.Range("K1").Value = "Extended Price"
    With wrkngSht.Columns("K")
        .ColumnWidth = 15.57
        .NumberFormat = "0.00"
        .HorizontalAlignment = xlRight
    End With
    
    With wrkngSht
        lr = Range("B" & Rows.Count).End(xlUp).Row
        .Range("H2:H" & lr).Value = "0.00"
        .Range("I2:I" & lr).Formula = "=G2-(G2*(H2/100))"
        .Range("K2:K" & lr).Formula = "=I2*J2"
    End With
        
    'Insert Blank Rows After Each Change In SMARTnet(M) Level but now changed it to Contract(L)
    With wrkngSht
    For i = lr To 3 Step -1
        If Cells(i - 1, "M").Value <> Cells(i, "M").Value Then
            Cells(i, "M").Resize(3).EntireRow.Insert
            Range("B" & i + 2).Value2 = "Contract# " & Range("L" & i + 3).Value2 & _
                 " -- SiteID# " & Range("N" & i + 3).Value2
                 Range("B" & i + 2).Font.FontStyle = "Bold"
        End If
    Next i
    End With
        
    'Add rows to start making the sheet look pretty
    wrkngSht.Rows(1).Resize(10).EntireRow.Insert
    wrkngSht.Rows(12).Resize(2).EntireRow.Insert
        
    'Set row heights
    wrkngSht.Rows(1).RowHeight = 12
    For x = 2 To 8
        Rows(x).RowHeight = 15
    Next
    wrkngSht.Rows(9).RowHeight = 0.75
    For x = 10 To 255
        Rows(x).RowHeight = 15
    Next
        
    wrkngSht.Range("B13").Value2 = "Contract# " & Range("L14").Value2 & _
        " -- SiteID# " & Range("N14").Value2
        Range("B13").Font.FontStyle = "Bold"
                 
    'Set report title and title properties
    wrkngSht.Range("D3").Value = WorksheetFunction.Proper(wrkngSht.Range("O14").Value)
    With wrkngSht.Range("D3:H3")
            .MergeCells = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = True
            .Font.Size = 16
        End With
    With wrkngSht.Range("D4:H4")
            .MergeCells = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Value = "Report Description"
            .Font.Bold = True
            .Font.Size = 10
        End With
        
    With wrkngSht.Range("B7")
            .Value = "Customer Contact"
            .Font.Bold = True
            .Font.Size = 11
        End With
        
    With wrkngSht.Range("B8")
            .Value = "=TODAY()"
            .HorizontalAlignment = xlLeft
        End With
        
    wrkngSht.Range("D7").Value = "**Pricing Expires After 30 Days**"
                 
    'Delete the last bit of un-needed columns
    wrkngSht.Columns("L:T").EntireColumn.Delete
    
    'Set header row font size and weight
    With wrkngSht.Range("B11:K11")
            .HorizontalAlignment = xlCenter
            .Font.Size = 11
            .Font.Bold = True
        End With
     With wrkngSht.Range("B11:K11").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = -0.249977111117893
            .PatternTintAndShade = 0
        End With
    With wrkngSht.Range("B11:K11").Borders
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
    wrkngSht.Range("B11:K11").Borders(xlEdgeBottom).LineStyle = xlNone
        
    wrkngSht.Columns("L").ColumnWidth = 4.43
        
    'Insert PDS Logo into upper-left corner
    wrkngSht.Range("B1").Activate
    wrkngSht.Pictures.Insert("C:\Users\tcoplien\Desktop\SMARTnet\pdsLogo.png") _
    .Select
        With Selection.ShapeRange
            .Height = 64.8
            .ScaleHeight 1.25, msoFalse, msoScaleFromTopLeft
            .IncrementLeft 6
            .IncrementTop 4
        End With

    With wrkngSht.Range("I7")
            .Value = "CSCC Q:"
            .HorizontalAlignment = xlRight
            .Font.Bold = True
            .Font.Size = 8
        End With
        
    'Copy CSCC Q# from File Name sheet to J7
    nameSht.Range("B1").Copy
    
    wrkngSht.Range("J7").PasteSpecial Paste:=xlPasteValues
    With wrkngSht.Range("J7:K7")
            .MergeCells = True
            .HorizontalAlignment = xlLeft
            .NumberFormat = "General"
            .Font.Bold = True
            .Font.Size = 8
        End With
        
    'Set horizontal alignment of table from row 13 and down
    With wrkngSht
        lr = .Range("B" & Rows.Count).End(xlUp).Row
        .Range("B13:B" & lr).HorizontalAlignment = xlLeft
        .Range("E13:E" & lr).HorizontalAlignment = xlCenter
        .Range("G13:G" & lr).HorizontalAlignment = xlRight
        .Range("H13:H" & lr).HorizontalAlignment = xlCenter
        .Range("I13:I" & lr).HorizontalAlignment = xlRight
        .Range("J13:J" & lr).HorizontalAlignment = xlCenter
        With .Range("K13:K" & lr)
            .HorizontalAlignment = xlRight
            .NumberFormat = "0.00"
        End With
    End With
        
    'Now go to SMS Order Info Tab and clean that up
    orderSht.Activate
    With ActiveWindow
        .DisplayGridlines = False
        .Zoom = 86
    End With
    With orderSht.Cells
        .Interior.ColorIndex = xlColorIndexNone
        .Font.ColorIndex = xlColorIndexAutomatic
        .Borders.LineStyle = xlLineStyleNone
        .Font.Name = "Calibri"
        .Font.Size = 12
    End With
    
    With orderSht
        .Range("A:A, D:M").Delete
        .Range("B:B, E:F").Delete
    End With
    
    orderSht.Columns("A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    orderSht.Rows(1).Resize(3).EntireRow.Insert
        
    With orderSht.Range("B2")
            .Value = "To order coverage below, PDS will order the following SMS SKU:"
            .Font.Bold = True
            .Font.Size = 9
        End With
    With orderSht
        .Range("B4").Value = "Line Item"
        .Range("C4").Value = "Part Number"
        .Range("D4").Value = "Quantity"
        .Range("B4:D4").Font.Bold = True
    End With
    
    Set bBox = orderSht.Range("B2:D2")
    Set bBox2 = Range(bBox, bBox.End(xlDown).End(xlDown))
    With bBox2.Borders
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With bBox2
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
        
    bBox.HorizontalAlignment = xlLeft

    With orderSht
        .Columns("B:B").ColumnWidth = 10.14
        .Columns("C:C").ColumnWidth = 30
        .Columns("D:D").ColumnWidth = 17.71
    End With
    
    For x = 2 To 13
        orderSht.Rows(x).RowHeight = 15
    Next
    
    'Move back to wrkngSheet and add the borders around each section in the BoM
    wrkngSht.Activate
    For Each Cell In Range("D14", Range("D65536").End(xlUp))
    If Cell.Value > "" Then
        With Range(Cell.Offset(0, -2), Cell.Offset(0, 7))
            With .Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
        End With
    End If
    Next
    
    findMaterials_SMS 'Module for formatting the subtotal row under each section
        
    'Create the Project Total Row
    wrkngSht.Range("D65536").End(xlUp).Offset(2).Select
    wrkngSht.Range(ActiveCell, ActiveCell.Offset(0, 6)).Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = RGB(0, 51, 204)
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With Selection
            .HorizontalAlignment = xlRight
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = True
            .Value = "xYr Renewal Total(USD):"
        End With
        With Selection.Font
            .TintAndShade = 0
            .Color = vbWhite
            .Size = 14
            .Bold = True
        End With
    
    wrkngSht.Range("K65536").End(xlUp).Offset(3).Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = RGB(0, 51, 204)
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With Selection.Font
            .Color = vbWhite
            .TintAndShade = 0
            .Size = 14
            .Bold = True
        End With
        With Selection
            .HorizontalAlignment = xlRight
            .VerticalAlignment = xlCenter
            .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            .Formula = "=SUM(K14:K" & lr & ")"
        End With
        
    'Add Border Above Project Total Row
    lr = lr + 3
        With wrkngSht
            .Range("B" & lr, "K" & lr).Borders(xlEdgeTop).Weight = xlThin
        End With
    
    str = nameSht.Range("B2")
    strMid = Mid(str, 10)
    
    str1 = nameSht.Range("B1")
    str1Left = Left(str1, 8)
    
    fName = Format(Now(), "yyyymmdd") & " " & strMid & " CSCC Q" & str1 & " wPRICING" & ".xlsx"
    
    nameSht.Delete
    
    With wrkngSht.PageSetup
        .PrintArea = "$B:$K"
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
        
    wrkngSht.SaveAs fileName:=ActiveWorkbook.Path & "\" & fName, _
        FileFormat:=xlOpenXMLWorkbook, Password:="", WriteResPassword:="", _
        ReadOnlyRecommended:=False, CreateBackup:=False
        
    With Application
        .DisplayAlerts = True   'turn system alerts back on
        .EnableEvents = True    'turn other macros back on
        .ScreenUpdating = True  'refreshes the screen
    End With
    
    End Sub

