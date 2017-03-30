Sub IRIS_Edit()
'
' IRIS_Edit Macro
'

'
    Dim theWB As Workbook
    Dim theWS As Worksheet
    Dim listObj As ListObject ' For determining if table1 exist
    Dim rList As Range ' For removing Table and converting to Range
    Dim lRow As Long ' For storing the last row when filling in Discounts and other columns
    Dim lr As Long ' For inserting blank rows after each change in Column A
    Dim i As Long
    Dim theRng As Range
    Dim theDate As Range ' For placing the TODAY formula and formatting that cell
    Dim theContact As Range, theCustomer As Range
    Dim theDesc As Range
    Dim thePic As String
    Dim Cell As Range
    Dim NR As Long
    Dim theB11 As Range
    Dim thePTRC As Range
    Dim thePTRCR As Range
    Dim thePTRH As Range
    Dim n As Long, q As Long, x As String
    n = 1
    
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
    
    'Fill in Discount%, Discounted Price, & ExtPrice Columns
    lRow = Range("D" & Rows.Count).End(xlUp).Row
    Range("C2:C" & lRow).WrapText = True
    Range("D2:D" & lRow).HorizontalAlignment = xlRight
    Range("E2") = "0.00": Range("E2:E" & lRow).FillDown
    Range("E2:E" & lRow).HorizontalAlignment = xlCenter
    Range("F2:F" & lRow).NumberFormat = "0.00"
    Range("F2") = "=if(isnumber(d2),d2-(d2*(e2/100)),0)": Range("F2:F" & lRow).FillDown
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
                    .Font.Color = RGB(149, 179, 215)
                    .HorizontalAlignment = xlLeft
                    .Value2 = "Materials"
                End With
        End If
    Next i
    
    'Delete some unneeded rows
    Rows("2:3").Delete Shift:=xlUp
    
    With Columns("A")
        .ClearContents
        .ColumnWidth = 1.29
    End With
    
    theWS.Rows(1).Resize(10).EntireRow.Insert
    
    Rows("9").EntireRow.Hidden = True
      
    'Set the column widths
    Columns("B").ColumnWidth = 25
    Columns("C").ColumnWidth = 56.29
    Columns("D").ColumnWidth = 13.57
    Columns("E").ColumnWidth = 12.29
    Columns("F").ColumnWidth = 15.57
    Columns("G").ColumnWidth = 8.14
    Columns("H").ColumnWidth = 17.29
    Columns("I").ColumnWidth = 4.43
    
    Dim rw As Long, srw As Long, col As Long

    With ActiveSheet
        For col = 1 To Cells(1, Columns.Count).End(xlToLeft).Column
            If .Cells(11, col) = "Extended Price" Then
                srw = 2
                For rw = 2 To .Cells(Rows.Count, col).End(xlUp).Row + 1
                    If IsEmpty(.Cells(rw, col)) And rw > srw Then
                        '.Cells(rw, col).Value = Application.Sum(.Range(.Cells(srw, col), .Cells(rw - 1, col)))
                        .Cells(rw, col).Formula = "=SUM(" & .Cells(srw, col).Address(0, 0) & _
                                    Chr(58) & .Cells(rw - 1, col).Address(0, 0) & ")"
                        .Cells(rw, col).NumberFormat = _
                          "[color5]_($* #,##0.00_);[color9]_($* (#,##0.00);[color15]_("" - ""_);[color10]_(@_)"
                        srw = rw + 1
                    End If
                Next rw
            End If
        Next col
    End With

    'Define the Customer Contact Variable and format cell
    Set theContact = theWS.Range("B7")
    With theContact
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .Font.Bold = True
        .Value = "Customer Contact"
    End With
    
    'Define the Date Variable and format cell
    Set theDate = theWS.Range("B8")
    With theDate
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .Font.Bold = True
        .Value = "=TODAY()"
    End With
    
    
    With Range("C7")
        .Value = "**Pricing Expires After 30 Days**"
        .Font.Size = 10
    End With
    
    'For placing the Customer Name
    Set theCustomer = theWS.Range("C3:D3")
    With theCustomer
        .HorizontalAlignment = xlCenter
        .MergeCells = True
        .Value = "Customer Name"
        .Font.Size = 16
        .Font.Bold = True
    End With

    'For placing the BoM Description
    Set theDesc = theWS.Range("C4:D4")
    With theDesc
        .HorizontalAlignment = xlCenter
        .MergeCells = True
        .Font.Size = 11
        .Value = "BoM Description"
    End With
    
    'Add the borders around the Header Row and Change the Color
    Set theB11 = Range("B11", Range("B11").End(xlToRight))
    With theB11.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    With theB11.Font
        .Size = 11
        .Bold = True
    End With
    With theB11.Borders
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    theB11.Borders(xlEdgeBottom).LineStyle = xlNone

    
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
    End With
    
    ' Add Border Above Project Total Row
    NR = Range("B" & Rows.Count).End(xlUp).Row + 3
    With Range("B" & NR, "H" & NR)
        .Borders(xlEdgeTop).Weight = xlThin
    End With
    
    'Insert PDS Logo
    Range("B1").Select
    theWS.Pictures.Insert("C:\Users\tcoplien\Desktop\SMARTnet\pdsLogo.png") _
         .Select
    With Selection.ShapeRange
        .Height = 64.8
        .ScaleHeight 1.2, msoFalse, msoScaleFromTopLeft
        .IncrementLeft 4
        .IncrementTop 4
    End With
    
    'Rename the tab
    theWS.Name = "Report"
    
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
