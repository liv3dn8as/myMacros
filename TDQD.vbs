    Public downSell As String 'getting the discount off list %
    Public isTrade As String 'asking if there is trade involved
    Public lRow As Long 'for getting last row
    Public sMarginPercent As Long 'for getting desired margin
    Public sMargin As String 'for converting Margin to decimal

    


Sub TDQ_Discount()
'
'TD Quote Discounting Macro
'
    Dim thisWB As Workbook 'this workbook
    Dim thisWS As Worksheet 'this worksheet
    'Dim sMarginPercent As Long 'for getting desired margin
    'Dim sMargin As Long 'for converting Margin to decimal
    'Dim markUp As Long 'for getting markup from margin
    Const oneHundred As Long = 100 'for blah blah blah
    Const sRow As Long = 14 'What's the first row with data?
    
    Set thisWB = ActiveWorkbook
    Set thisWS = thisWB.Sheets("Quote")
    
    'turn off screen updating
    Application.ScreenUpdating = False
    
    'remove a couple of un-needed objects
    With thisWS
        '.ListObjects("keepComments").Delete
        '.ListObjects("keeplogo").Delete
        .Range("I5:I7").Cut Range("I5:I7").Offset(0, -2)
        .Range("H:J").Delete
        .Range("B:B").Delete
    End With
    
    'find SMARTnet SKU's and delete that row
    With thisWS
        For i = lRow To sRow Step -1
        'If Not IsError(.Value) Then
            If i Like "CON-*" Then
                i.EntireRow.Delete
            End If
        Next i
    End With
    
    'get the last row
    lRow = Range("B" & Rows.Count).End(xlUp).Row
    
    'ask what you would like downsell to be
    'downSell = Application.InputBox(Prompt:="What will Discount off List be?", Type:=1 + 8)
    
    'ask what you would like margin to be at
    sMarginPercent = Application.InputBox(Prompt:="What is the desired margin?", Type:=1 + 8)
    
  '  sMarginPercent1 = FormatPercent(sMarginPercent, 2)
    
    'for converting the Margin Percent to Decimal
    sMargin = sMarginPercent / oneHundred
    
    'for converting Margin to Mark-Up
    'markUp = sMargin / (1 - sMargin)
        
    'ask if trade is involved
    isTrade = MsgBox("Is there trade involved?", vbYesNo + vbQuestion, "IsTradeInvolved")
        If isTrade = vbYes Then
            tradeInYes 'trade-in module here if trade exist
        Else
            tradeInNo 'module if trade doesn't exist
        End If
    
End Sub

Private Sub tradeInYes()
'
'if trade exist Private Sub Module
'
    
    'insert some column headers
    Range("G13").Value = "% Blended"
    Range("H13").Value = "w/o Credit %"
    Range("I13").Value = "Trade-In Credit"
    Range("J13").Value = "Extended Sell Price"
    Range("K13").Value = "Margin"
    
    'start filling in formulas
    Range("G14:G" & lRow).Formula = "=SUM(1-(D14/F14))*100" 'our discount
    'Range("H14:H" & lRow).Formula = "=SUM(1-(D14+(I14/A14))/F14)*100" 'our discount before trade
        If downSell > 0 Then
            Range("J14:J" & lRow).Formula = "=(f14-(f14*(" & sMargin & "/100)))*a14" 'sell price
        End If
    Range("k14:k" & lRow).Formula = "=((J14-E14)/J14)*100" 'margin
    
    'set numberformat for columns H and I
    Range("H14:K" & lRow).NumberFormat = "0.00"
    
    With Range("G14", Range("G14").End(xlToRight).End(xlDown)).Borders
        .LineStyle = xlContinuous
        .Color = Black
        .Weight = xlThin
    End With
            
End Sub

Private Sub tradeInNo()
'
'if trade doesn't exist then do this private module
'

    'insert some column headers
    Range("G13").Value = "% Off List"
    Range("H13").Value = "Sell Price"
    Range("I13").Value = "Margin"
    Range("J13").Value = "Customer %"
    
    'start filling in formulas
    Range("G14:G" & lRow).Formula = "=SUM(1-(D14/F14))*100"
        If sMarginPercent > 0 Then
            Range("H14:H" & lRow).Formula = "=d14/(1-" & sMargin & ")"
        End If
    Range("I14:I" & lRow).Formula = "=((H14-D14)/H14)*100"
    
    'adds the sell discount
    Range("J14:J" & lRow).Formula = "=SUM(1-(H14/F14))*100"
    
    'set numberformat for columns H and I
    Range("H14:J" & lRow).NumberFormat = "0.00"
    
    With Range("G14", Range("G14").End(xlToRight).End(xlDown)).Borders
        .LineStyle = xlContinuous
        .Color = Black
        .Weight = xlThin
    End With
End Sub
