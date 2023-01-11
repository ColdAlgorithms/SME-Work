Sub Giris_Cikis_Saat()


'Declare Variables
Dim WB As Workbook
Dim PSheet As Worksheet
Dim DSheet As Worksheet
Dim ESheet As Worksheet
Dim ASheet As Worksheet
Dim PCache As PivotCache
Dim ECache As PivotCache
Dim ACache As PivotCache
Dim PTable As PivotTable
Dim ETable As PivotTable
Dim ATable As PivotTable
Dim DRange As Range
Dim LastRow As Long
Dim LastCol As Long
Set DSheet = Worksheets("Geçiş Listesi")
Set WB = ActiveWorkbook

'Define Data Range
LastRow = DSheet.Cells(Rows.Count, 1).End(xlUp).Row
LastCol = DSheet.Cells(1, Columns.Count).End(xlToLeft).Column
Set DRange = DSheet.Cells(1, 1).Resize(LastRow, LastCol)

'Add "Ad Soyad" Column
Columns("E:E").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Range("E1").Select
ActiveCell.FormulaR1C1 = "=RC[-2]&"" ""&RC[-1]"
Range("E1").Select
Selection.AutoFill Destination:=Range("E1:E" & LastRow)
Range("E1:E" & LastRow).Select
Columns("E:E").EntireColumn.AutoFit

'Set Pivot Table with Entrance Hours
On Error Resume Next
Application.DisplayAlerts = False
Worksheets("PivotTable").Delete
Sheets.Add Before:=ActiveSheet
ActiveSheet.Name = "PivotTable"
Application.DisplayAlerts = True
Set PSheet = Worksheets("PivotTable")

'Define Pivot Cache
Set PCache = ActiveWorkbook.PivotCaches.Create _
(SourceType:=xlDatabase, SourceData:=DRange). _
CreatePivotTable(TableDestination:=PSheet.Cells(2, 2), _
TableName:="PivotTable")

'Pivot Modifications
    With ActiveSheet.PivotTables("PivotTable").PivotFields("Ad Soyad")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable").PivotFields("Giriş Saat")
        .Orientation = xlRowField
        .Position = 2
    End With
    
'Counter for Data on the PivotTable Page
Dim PivotCount As Integer
PivotCount = PSheet.Cells(Rows.Count, "B").End(xlUp).Row

'Modify Pivot Table
    Sheets("PivotTable").Select
    Range("C3").Select
    ActiveCell.FormulaR1C1 = 1
    Range("C4").Select
    ActiveCell.FormulaR1C1 = 2
    Range("C3:C4").Select
    Selection.AutoFill Destination:=Range("C3:C" & PivotCount)
    Range("A3").Select
    ActiveCell.FormulaR1C1 = 1
    Range("A4").Select
    ActiveCell.FormulaR1C1 = 2
    Range("A3:A4").Select
    Selection.AutoFill Destination:=Range("A3:A" & PivotCount)
    
'Set Pivot Table for Employee List
On Error Resume Next
Application.DisplayAlerts = False
Worksheets("EmployeeList").Delete
Sheets.Add Before:=ActiveSheet
ActiveSheet.Name = "EmployeeList"
Application.DisplayAlerts = True
Set ESheet = Worksheets("EmployeeList")
Set DRange = DSheet.Cells(1, 1).Resize(LastRow, LastCol)

'Set Employee List
Set ECache = ActiveWorkbook.PivotCaches.Create _
(SourceType:=xlDatabase, SourceData:=DRange). _
CreatePivotTable(TableDestination:=ESheet.Cells(2, 2), _
TableName:="EmployeeList")
With ESheet.PivotTables("EmployeeList").PivotFields("Ad Soyad")
    .Orientation = xlRowField
    .Position = 1
End With

'Employee Counter
    Dim EmployeeCount As Integer
    EmployeeCount = ESheet.Cells(Rows.Count, "B").End(xlUp).Row
    
'Employee List Modif
    Sheets("EmployeeList").Select
    Range("C3").Select
    ActiveCell.FormulaR1C1 = 1
    Range("C4").Select
    ActiveCell.FormulaR1C1 = 2
    Range("C3:C4").Select
    Selection.AutoFill Destination:=Range("C3:C" & EmployeeCount)
    Range("A3").Select
    ActiveCell.FormulaR1C1 = 1
    Range("A4").Select
    ActiveCell.FormulaR1C1 = 2
    Range("A3:A4").Select
    Selection.AutoFill Destination:=Range("A3:A" & EmployeeCount)

'Get Column D & E Functions
    Sheets("PivotTable").Select
    Range("D3").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2],'Geçiş Listesi'!C[1],1,)"
    
    Selection.AutoFill Destination:=Range("D3:D" & PivotCount)
    
    Range("E3").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]=RC[-3],R[1]C[-3],"""")"
    Range("E3").Select
    Selection.AutoFill Destination:=Range("E3:E" & PivotCount)

'Get Column F Function
    Range("F3").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(VLOOKUP(VLOOKUP(VLOOKUP(RC[-2],EmployeeList!R3C2:R200C3,2,FALSE)+1,EmployeeList!R3C1:R" & PivotCount & "C2,2,FALSE),PivotTable!R3C2:R" & PivotCount & "C3,2,FALSE)-1,PivotTable!RC[-5]:R[" & PivotCount & "]C[-4],2,FALSE)"

'
    Range("D3:F3").Select
    Range("F3").Activate
    Selection.AutoFill Destination:=Range("D3:F" & PivotCount)
    Range("D3:F" & PivotCount).Select
    Range("D3").Select
    Columns("D:D").EntireColumn.AutoFit

    Sheets("PivotTable").Select
    Range("G3").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(RC[-3],"""")"
    Range("G3").Select
    Selection.AutoFill Destination:=Range("G3:I3"), Type:=xlFillDefault
    Range("G3:I3").Select
    Selection.AutoFill Destination:=Range("G3:I" & PivotCount)
    Range("G3:I" & PivotCount).Select
    Selection.Copy
    Range("J3").Select
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveWorkbook.Worksheets("PivotTable").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("PivotTable").Sort.SortFields.Add Key:=Range("J3") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    
    With ActiveWorkbook.Worksheets("PivotTable").Sort
        .SetRange Range("J3:L1000")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    Range("J3:L3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("J3:L" & PivotCount).Select
    Selection.End(xlUp).Select
    Range("J3:J4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("J4").Select
    Selection.End(xlDown).Select
    Range("J" & PivotCount).Select
    Selection.End(xlUp).Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Range("L3").Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Range("J3").Select
    Columns("J:J").EntireColumn.AutoFit

'Copy Analysed Data
    Sheets("PivotTable").Select
    Sheets("PivotTable").Copy
    Columns("A:I").Select
    Range("I1").Activate
    Selection.Delete Shift:=xlToLeft
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "Ad Soyad"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "Giriş Saati"
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "Çıkış Saati"
    Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$2:$C$" & EmployeeCount), , xlYes).Name = _
        "Tablo1"

'Range Alphabetically
 ActiveWorkbook.Worksheets("PivotTable").ListObjects("Tablo1").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("PivotTable").ListObjects("Tablo1").Sort.SortFields. _
        Add Key:=Range("Tablo1[[#All],[Ad Soyad]]"), SortOn:=xlSortOnValues, Order _
        :=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("PivotTable").ListObjects("Tablo1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

'Head for "Gelmeyenler" Column
    Range("F2").Select
    ActiveCell.FormulaR1C1 = "GELMEYENLER"
    Columns("F:F").ColumnWidth = 23.53
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "#"
    
    Range("E2:F2").Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$E$2:$F$2"), , xlYes).Name = _
        "Tablo2"
    Range("Tablo2['#]").Select
    ActiveCell.FormulaR1C1 = 1
    Range("E4").Select
    ActiveCell.FormulaR1C1 = 2
    
'Pivot Table For Absents
On Error Resume Next
Application.DisplayAlerts = False
Worksheets("Absents").Delete
Sheets.Add Before:=ActiveSheet
ActiveSheet.Name = "Absents"
Application.DisplayAlerts = True
Set ASheet = Worksheets("Absents")
Set ARange = DSheet.Cells(1, 1).Resize(LastRow, LastCol)

Set ACache = ActiveWorkbook.PivotCaches.Create _
(SourceType:=xlDatabase, SourceData:=ARange). _
CreatePivotTable(TableDestination:=ASheet.Cells(2, 2), _
TableName:="Absents")

'Pivot Modifications
 With ActiveSheet.PivotTables("Absents").PivotFields("Ad Soyad")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveWindow.SmallScroll Down:=12
    With ActiveSheet.PivotTables("Absents").PivotFields("Çıkış Durum")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveWindow.SmallScroll Down:=-27
    With ActiveSheet.PivotTables("Absents").PivotFields("Departman")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("Absents").PivotFields("Çıkış Durum"). _
        ClearAllFilters
    ActiveSheet.PivotTables("Absents").PivotFields("Çıkış Durum"). _
        CurrentPage = "Tatil"
    ActiveSheet.PivotTables("Absents").PivotFields("Departman"). _
        CurrentPage = "(All)"
    With ActiveSheet.PivotTables("Absents").PivotFields("Departman")
        .PivotItems("MİSAFİR").Visible = False
        .PivotItems("STAJYER").Visible = False
        .PivotItems("YÖNETİM").Visible = False
    End With
    ActiveSheet.PivotTables("Absents").PivotFields("Departman"). _
        EnableMultiplePageItems = True

    ActiveSheet.PivotTables("Absents").PivotSelect "'Ad Soyad'[All]", xlLabelOnly _
        + xlFirstRow, True
    Selection.Copy
    Sheets("PivotTable").Select
    Range("F3").Select
    ActiveSheet.Paste

    Dim AbsentsCount As Integer
    AbsentsCount = Cells(Rows.Count, "F").End(xlUp).Row
    Range("E3:E4").Select
    Selection.AutoFill Destination:=Range("E3:E" & AbsentsCount)
    
    
'Save the Final Report
ActiveSheet.Name = DSheet.Range("G2")

End Sub
