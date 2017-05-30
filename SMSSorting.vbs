Option Private Module
'created by liv3dn8as
'current as of 5/30/2017

Sub SMSSorting()
 
    'Define variables for the data
    Dim MyDataFirstCell, MyDataLastCell
    Dim COLUMNICellStart, COLUMNICellEnd
    Dim COLUMNHCellStart, COLUMNHCellEnd
    Dim COLUMNCCellStart, COLUMNCCellEnd
    Dim COLUMNECellStart, COLUMNECellEnd
 
    'Establish the Data Area
    MyDataFirstCell = Range("A1").Address
    MyDataLastCell = Range("A1").End(xlToRight).End(xlDown).Address
 
    'COLUMN I
    COLUMNICellStart = Range("I2").Address
    COLUMNICellEnd = Range("I2").End(xlDown).Address
 
    'COLUMN H
    COLUMNHCellStart = Range("H2").Address
    COLUMNHCellEnd = Range("H2").End(xlDown).Address
 
    'COLUMN C
    COLUMNCCellStart = Range("C2").Address
    COLUMNCCellEnd = Range("C2").End(xlDown).Address
 
    'COLUMN E
    COLUMNECellStart = Range("E2").Address
    COLUMNECellEnd = Range("E2").End(xlDown).Address
 
    'Start the sort by specifying sort area and columns and then sort full sheet.
    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Sort.SortFields.Add Key:=Range(COLUMNICellStart & ":" & COLUMNICellEnd), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveSheet.Sort.SortFields.Add Key:=Range(COLUMNHCellStart & ":" & COLUMNHCellEnd), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveSheet.Sort.SortFields.Add Key:=Range(COLUMNCCellStart & ":" & COLUMNCCellEnd), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveSheet.Sort.SortFields.Add Key:=Range(COLUMNECellStart & ":" & COLUMNECellEnd), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortOnValues
 
    With ActiveSheet.Sort
        .SetRange Range(MyDataFirstCell & ":" & MyDataLastCell)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub
