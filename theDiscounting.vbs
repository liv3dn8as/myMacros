    'for entering discounting in appropriate column
    'created by liv3dn8as
    'current as of 5/30/2017
    
    Sub theDiscount()

        Dim dType As Variant
        Dim bDisc As String 'for getting the HW Discount on NetformX BoM
        Dim iDisc As String 'for getting the HW Discount on IRIS Exported BoM
        Dim sDisc As String 'for getting the Discount on SMARTnet Quote
        Dim snDiscount As String 'for getting the SMARTnet Discount
        
        'Find out if this is a NetformX BoM, IRIS Exported BoM, or SMARTnet Quotes
        dType = Application.InputBox(Prompt:="Type NetformX, IRIS, or SMS depending on BoM/Quote type?")
        
        If dType = "NetformX" Then
            bomDisc
        ElseIf dType = "IRIS" Then
            irisDisc
        ElseIf dType = "SMS" Then
            smsDisc
        End If
    
    End Sub
    
    '--------------------------------------------------
    'Section for NetformX BoM
    '--------------------------------------------------
    
    Private Sub bomDisc()
    
        Dim lRow As Long 'for finding the last row
        
        'get the last row
        lRow = Range("B" & Rows.Count).End(xlUp).Row
    
        'Get user input for HW Discount
        bDisc = Application.InputBox(Prompt:="What would you like to Discount the Hardware at?", _
        Type:=1 + 8)
                
        'Get user input for SMARTnet Discount
        snDiscount = Application.InputBox(Prompt:="What would you like to Discount the SMARTnet at?", _
        Type:=1 + 8)
    
        'Change the HW Discount
        For Each C In Range("E14:E" & lRow)
            If Range("D14:D" & lRow) Is Nothing Then
                C.Value = "0"
            Else
                If C.Offset(, -1) > "1" Then
                    C.Value = bDisc
                End If
            End If
            If C.Offset(, -1) = "Included" Then
                C.Value = "0.00"
            End If
        Next C
        
        'Change the SMARTnet Discount
        y = "CON-"
        For Each Cell In Range("B14:B" & lRow)
        If Cell Like y & "*" And snDiscount > 0 Then
            Cell.Offset(, 3).Value = snDiscount
        End If
        Next
        
    End Sub
    
    '----------------------------------------------
    'Section for IRIS Exported BoMs
    '----------------------------------------------
    
    Private Sub irisDisc()
    
        Dim dCol As Range, bCol As Range 'putting it into a variable
        Dim lRow As Long 'for finding the last row
        
        'get the last row
        lRow = Range("B" & Rows.Count).End(xlUp).Row

        'Get user input
        iDisc = Application.InputBox(Prompt:="What would you like to Discount this at?", _
            Type:=1 + 8)
    
        For Each C In Range("E14:E" & lRow)
            If C.Offset(, -1) = "incl." Then
                C.Value = "0.00"
            Else
                If C.Offset(, -1) > "" Then
                    C.Value = iDisc
                    C.NumberFormat = "0.00"
                End If
            End If
        Next C
    
    End Sub
    
    '----------------------------------------------
    'Section for SMARTnet Quotes
    '----------------------------------------------

    Private Sub smsDisc()
    
        Dim lRow As Long 'for finding the last row
        
        'get the last row
        lRow = Range("B" & Rows.Count).End(xlUp).Row
        
    
        'Get user input
        sDisc = Application.InputBox(Prompt:="What would you like to Discount this at?", _
            Type:=1 + 8)
    
        For Each C In Range("H14:H" & lRow)
            If Range("G14:G" & lRow) Is Nothing Then
                C.Value = "0"
            Else
                If C.Offset(, -1) > "1" Then
                    C.Value = sDisc
                End If
            End If
        Next C
        
     End Sub
