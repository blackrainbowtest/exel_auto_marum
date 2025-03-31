Attribute VB_Name = "DataMoveModule"
'Option Explicit

Sub MoveData()
Application.ScreenUpdating = False
' init ranges
    Dim currentRange As Range
    Dim AutoCreateRange As Range
    Dim LoopRange As Range
    Dim LoopStep As Integer
' init addres
    Set currentRange = Selection
    If currentRange.Column <> 1 Then
        MsgBox "Choose first (A) column"
        Exit Sub
    End If
    Set AutoCreateRange = Sheets("AutoCreate").Range("A3")
    Set LoopRange = Sheets("BASE").Range("O1")
' init temp data
    LoopStep = 0
' clear autoCreate programm area
    Sheets("AutoCreate").Select
    If Range("A3").Value <> "" Then
        Range("A3").Select
        If Range("A4").Value <> "" Then
            Range(Selection, Selection.End(xlDown)).Select
        End If
        ActiveCell.Range("A1:U50").Select
        'Range(Selection, Selection.End(xlToRight)).Select
        Selection.ClearContents
    End If

    Sheets("TempDataBase").Select
    currentRange.Select
' monkey code style (sry not enought time for normal coding(
Do While currentRange.Offset(LoopStep, 0).Value <> ""
    
    '{USERLSCODE}
    AutoCreateRange.Offset(LoopStep, 0).Value = currentRange.Offset(LoopStep, 0).Value
    '{USERENG}
    AutoCreateRange.Offset(LoopStep, 1).Value = GHEAToEnglish(currentRange.Offset(LoopStep, 1).Value)
    '{CODEAUTO}
    AutoCreateRange.Offset(LoopStep, 2).Value = currentRange.Offset(LoopStep, 2).Value & "_AUTO"
    '{DAY}
    AutoCreateRange.Offset(LoopStep, 3).Value = Day(Get_Date)
    If AutoCreateRange.Offset(LoopStep, 3).Value < 10 Then
        AutoCreateRange.Offset(LoopStep, 3).Value = 0 & AutoCreateRange.Offset(LoopStep, 3).Value
    End If
    '{MONTH}
    AutoCreateRange.Offset(LoopStep, 4).Value = Month(Get_Date)
    If AutoCreateRange.Offset(LoopStep, 4).Value < 10 Then
        AutoCreateRange.Offset(LoopStep, 4).Value = 0 & AutoCreateRange.Offset(LoopStep, 4).Value
    End If
    '{YEAR}
    AutoCreateRange.Offset(LoopStep, 5).Value = Mid(Year(Get_Date), 3, 2)
    '{CODE}
    AutoCreateRange.Offset(LoopStep, 6).Value = currentRange.Offset(LoopStep, 12).Value
    '{MONEYBYNUM}
    AutoCreateRange.Offset(LoopStep, 12).Value = currentRange.Offset(LoopStep, 6).Value
    '{STOCKBYNUM}
    AutoCreateRange.Offset(LoopStep, 14).Value = currentRange.Offset(LoopStep, 9).Value
    '{TOWHOM} 7
    For i = 0 To 13
        If LoopRange.Offset(0, i).Value = currentRange.Offset(LoopStep, 7).Value Then
            '{USERLSCODE}
            AutoCreateRange.Offset(LoopStep, 0).Value = LoopRange.Offset(7, i).Value & "_" & AutoCreateRange.Offset(LoopStep, 0).Value
            '{TOWHOM}
            AutoCreateRange.Offset(LoopStep, 7).Value = LoopRange.Offset(1, i).Value
            '{TOWHERE}
            AutoCreateRange.Offset(LoopStep, 8).Value = LoopRange.Offset(2, i).Value
            '---{TOWHICH} done---
            If InStr(currentRange.Offset(LoopStep, 5).Value, "/17/") = 0 And InStr(currentRange.Offset(LoopStep, 5).Value, "-") = 0 Then
                AutoCreateRange.Offset(LoopStep, 9).Value = LoopRange.Offset(3, i).Value
            Else
                AutoCreateRange.Offset(LoopStep, 9).Value = LoopRange.Offset(3, 14).Value
            End If
            If InStr(currentRange.Offset(LoopStep, 5).Value, Sheets("BASE").Range("A19")) <> 0 Then
                AutoCreateRange.Offset(LoopStep, 9).Value = LoopRange.Offset(3, 15).Value
            End If
            'LoopRange.Offset(3, 14).Value
            
            If currentRange.Offset(LoopStep, 6).Value <> "" Then
                '{TEMP01}
                AutoCreateRange.Offset(LoopStep, 16).Value = LoopRange.Offset(4, i).Value
            '{AMD-USD}
                '{TEMP02}
                If currentRange.Offset(LoopStep, 15).Value = "USD" Then
                    AutoCreateRange.Offset(LoopStep, 17).Value = LoopRange.Offset(10, i).Value
                Else
                    AutoCreateRange.Offset(LoopStep, 17).Value = LoopRange.Offset(5, i).Value
                End If
                If AutoCreateRange.Offset(LoopStep, 14).Value <> "" Then
                    '{TEMP03}
                    AutoCreateRange.Offset(LoopStep, 18).Value = LoopRange.Offset(6, i).Value
                    '{TEMP04}
                    AutoCreateRange.Offset(LoopStep, 19).Value = LoopRange.Offset(8, i).Value
                End If
                '{TEMP05}
                AutoCreateRange.Offset(LoopStep, 20).Value = LoopRange.Offset(9, i).Value
            Else
                '{TEMP02}
                AutoCreateRange.Offset(LoopStep, 17).Value = vbCrLf ' add enter
                '{TEMP03}
                AutoCreateRange.Offset(LoopStep, 18).Value = vbCrLf
            End If

            Exit For
        Else
            AutoCreateRange.Offset(LoopStep, 9).Value = currentRange.Offset(LoopStep, 7).Value
        End If
    Next
    '{DOC}
    AutoCreateRange.Offset(LoopStep, 10).Value = currentRange.Offset(LoopStep, 5).Value
    '{USER}
    AutoCreateRange.Offset(LoopStep, 11).Value = currentRange.Offset(LoopStep, 1).Value
    '{MONEYBYTXT}
    AutoCreateRange.Offset(LoopStep, 13).Value = currentRange.Offset(LoopStep, 8).Value
    '{STOCKBYTXT}
    AutoCreateRange.Offset(LoopStep, 15).Value = currentRange.Offset(LoopStep, 10).Value
    'change init address range
    LoopStep = LoopStep + 1
    ' check
    'MsgBox LoopStep
Loop
Application.ScreenUpdating = True
MsgBox "ÓÌ€≥…›ªÒ¡ —≥Á·’·ı√€≥Ÿµ ˜·À≥›ÛÌª… ª›", vbInformation, "ºªœ·ı€Û"
End Sub



Function Get_Date() As String: Get_Date = Replace(Replace(DateValue(Now), "/", "-"), ".", "-"): End Function

