Attribute VB_Name = "AutoGetData"
'autoGetData - Add word filter and relax
Sub Autocheck()
Application.ScreenUpdating = False
    If BookOpenClosed(ThisWorkbook) Then
        MsgBox "ÓÌ€≥…›ªÒ¡ —≥Á·’·ı√€≥Ÿµ ˝«…ÔÒÌª… ª›"
    Else
        MsgBox "ËªÂ« Î≥Ò˘≥Õ Excel ˝≥€…¡ „ªÎ µ≥Ûª…, µ≥Û«"
    End If
Application.ScreenUpdating = True
End Sub

' check in open WB if find dif sheet
Function Sh_Exist(wb As Workbook, sName As String) As Boolean
    Dim wsSh As Worksheet
    On Error Resume Next
    Set wsSh = wb.Sheets(sName)
    Sh_Exist = Not wsSh Is Nothing
End Function

' check all workbooks
Function BookOpenClosed(thisBook As Workbook) As Boolean
    Dim myBook As Workbook
    For Each myBook In Workbooks
        If Sh_Exist(myBook, "Kaskaceli1") Then
            myBook.Activate
            Call AutoGet(thisBook, myBook)
            BookOpenClosed = True
            Exit For
        End If
    Next
    thisBook.Activate
    Sheets("TempDataBase").Select
    Call clearFormation
    Range("A1").Select
End Function

Sub AutoGet(thisBook As Workbook, currentBook As Workbook)
    Dim sheetName As String
'loop REMEMBERME: In 2025 we have from 1 to 19 Kaskaceli sheets
    For i = 1 To 19
        sheetName = ("Kaskaceli" & i)
        Sheets(sheetName).Select
        ' clear all filtres and unhide all cells
        Cells.Select
        Selection.EntireColumn.Hidden = False
        Selection.AutoFilter
        'autohide
        ActiveSheet.Range("$A$1:$BK$13006").AutoFilter Field:=2, Criteria1:=Array( _
        "0", "1", "10", "12", "13", "15", "16", "17", "19", "2", "20", "21", "22", "24", "25", "26", _
        "28", "29", "3", "30", "31", "32", "33", "34", "35", "36", "37", "38", "39", "4", "40", "43", _
        "44", "45", "46", "47", "48", "5", "51", "52", "53", "56", "58", "6", "7", "9", "ÿ/◊", "="), _
        Operator:=xlFilterValues
    ' start loop cell hide
        Columns("A:B").Select
        Selection.EntireColumn.Hidden = True
        Columns("F:P").Select
        Selection.EntireColumn.Hidden = True
        Columns("R:R").Select
        Selection.EntireColumn.Hidden = True
        Range("Q6").Select
        Selection.AutoFilter
        ActiveSheet.Range("$A$1:$BK$13006").AutoFilter Field:=17, Criteria1:="0"
    ' filter by searching data values
    ' FIXME: in 2025 need add NKT in filter too
        ActiveSheet.Range("$A$1:$BK$13006").AutoFilter Field:=19, Criteria1:=Array( _
            "∏≤–Œ", "∏≤–Œ —≥ﬂÔ·ı√€·ı›", "≤Èœ≥ ø µ≥≈›ªŸ≥Î", "‰– ", "¥≥≈›ªŸ≥Î« µÈ›≥∑≥›”·ıŸ", "‹ªÒœ≥€≥ÛÌª… ø œ≥Ô≥Ò·’≥œ≥› √ªÒ√«"), Operator:=xlFilterValues
    Call DataTransfer(thisBook, currentBook)
    Next
End Sub

Sub DataTransfer(thisBook As Workbook, currentBook As Workbook)

    Range("C1:S13007").Select
    Selection.Copy
    thisBook.Activate
    Sheets("TempSh").Select
    Range("A1").Select
    ActiveSheet.Paste
    If Range("A1").Offset(1, 0).Value <> "" Then
        Range("A1").Offset(1, 0).Range("A1").Select
        If ActiveCell.Offset(1, 0).Value <> "" Then
            Range(Selection, Selection.End(xlDown)).Select
        End If
        Range(Selection, Selection.End(xlToRight)).Select
        Selection.Copy
        Sheets("TempDataBase").Select
        Range("A1").Select
        Selection.End(xlDown).Offset(1, 0).Range("A1").Select
        ActiveSheet.Paste
    End If
    Sheets("TempSh").Select
    Cells.Select
    Selection.ClearContents
    currentBook.Activate
    
End Sub

Sub clearFormation()
    
    Cells.Select
    Cells.FormatConditions.Delete
    
    Columns("A:A").Select
    Selection.FormatConditions.AddUniqueValues
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).DupeUnique = xlDuplicate
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Columns("C:C").Select
    Selection.FormatConditions.AddUniqueValues
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).DupeUnique = xlDuplicate
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
End Sub


