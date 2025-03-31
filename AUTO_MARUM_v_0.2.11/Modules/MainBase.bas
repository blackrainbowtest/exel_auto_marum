Attribute VB_Name = "MainBase"
Sub Check_v02()
'
' Check_v02 Macro
'
' Keyboard Shortcut: Ctrl+o
'
' checking and replacing all courts

    Dim currentRange As Range
    Dim currentMoney As Range
    Dim tempMoneyVal As String
    Dim finalMoneyVal As String
    Dim tempSecMoneyVal As String
    Dim tempSendMoneyVal As String
    
    Set currentRange = Selection
    Set currentRange = currentRange.Offset(0, 5)
    
    If currentRange.Value = "" Then
        Exit Sub
    End If
    
    Set currentMoney = currentRange.Offset(0, 1)

Question:
'change TimesArmenian to GHEAgrapalat
    currentRange.Value = TimesToGHEA(currentRange.Value)
    currentRange.Offset(0, -1).Value = TimesToGHEA(currentRange.Offset(0, -1).Value)
    currentRange.Offset(0, -4).Value = TimesToGHEA(currentRange.Offset(0, -4).Value)
' change GHEA to english
    'MsgBox GHEAToEnglish(currentRange.Offset(0, -4).Value)
    'REMEMBERME: converting qax gotrsi hamary
    For i = 2 To 20
        If InStr(currentRange.Value, "/17/") = 0 Then
            If InStr(currentRange.Value, Sheets("BASE").Range("E" & i).Value) > 0 Then
                currentRange.Offset(0, 2).Value = Sheets("BASE").Range("F" & i).Value
                
                GoTo Continue
            End If
        End If
Continue:
    Next
    ' swaping number style
    tempMoney = currentRange.Offset(0, 1).Value
    finalMoneyVal = tempMoney
    If InStr(tempMoney, " ") Then
        tempMoney = Application.Trim(tempMoney)
        tempMoney = Replace(tempMoney, ",", ".")
        finalMoneyVal = Replace(tempMoney, " ", "")
        tempMoney = Replace(tempMoney, " ", ",")
        currentRange.Offset(0, 1).NumberFormat = "@"
        tempMoney = Replace(tempMoney, ".00", "")
        currentRange.Offset(0, 1) = tempMoney
    End If
    ' write nums by names
    currentRange.Offset(0, 3).Value = BasicFormula.AramTivTar(finalMoneyVal)
    ' write nums by names 2
    tempSecMoneyVal = currentRange.Offset(0, 4).Value
    tempSendMoneyVal = Replace(tempSecMoneyVal, ",", "")
    currentRange.Offset(0, 5).Value = BasicFormula.AramTivTar(tempSendMoneyVal)
    ' swaping name
    Call SwapFullName(currentRange.Offset(0, -4).Address, currentRange.Offset(0, 12).Address)
    ' delete "  " to " " data
    currentRange.Offset(0, 1).Value = Replace(currentRange.Offset(0, 1).Value, "  ", " ")
    ' check if "YAN " is end of string replace "YAN " to "YAN"
    If Right(currentRange.Offset(0, -4).Value, 4) = Sheets("BASE").Range("A14").Value & " " Then
        currentRange.Offset(0, -4).Value = Replace(currentRange.Offset(0, -4).Value, Sheets("BASE").Range("A14").Value & " ", Sheets("BASE").Range("A14").Value)
    End If
    ' check if "OV " is end of string replace "OV " to "OV"
    If Right(currentRange.Offset(0, -4).Value, 3) = Sheets("BASE").Range("A15").Value & " " Then
        currentRange.Offset(0, -4).Value = Replace(currentRange.Offset(0, -4).Value, Sheets("BASE").Range("A15").Value & " ", Sheets("BASE").Range("A15").Value)
    End If
    '
        currentRange.Offset(0, -4).Value = Replace(currentRange.Offset(0, -4).Value, "  ", " ")
    '
    If currentRange.Offset(1, 0).Value <> "" Then
        Set currentRange = currentRange.Offset(1, 0)
        GoTo Question
    End If
End Sub

' autocreate sheet
Sub SwapFullName(TargetRange, TempRange)
    If Right(Sheets("TempDataBase").Range(TargetRange).Value, 4) <> Sheets("BASE").Range("A14").Value & " " _
    And Right(Sheets("TempDataBase").Range(TargetRange).Value, 3) <> Sheets("BASE").Range("A14").Value And _
    Right(Sheets("TempDataBase").Range(TargetRange).Value, 3) <> Sheets("BASE").Range("A15").Value & " " And _
    Right(Sheets("TempDataBase").Range(TargetRange).Value, 2) <> Sheets("BASE").Range("A15").Value Then
'need to change this path of code __ important
        Sheets("TempDataBase").Range(TempRange).FormulaR1C1 = _
            "=MID(RC[-16],SEARCH("" "",RC[-16])+1,300)&"" ""&MID(RC[-16],1,SEARCH("" "",RC[-16]))"
        Sheets("TempDataBase").Range(TargetRange) = Sheets("TempDataBase").Range(TempRange).Value
        Sheets("TempDataBase").Range(TempRange).Value = ""
' End change path
    End If
End Sub


'change GHEAgrapalat to English letters
Function GHEAToEnglish(TEXT As String) As String
    'Sheets("BASE").Range("A15").Value
    For i = 2 To 80
        TEXT = Replace(TEXT, Sheets("BASE").Range("I" & i).Value, Sheets("BASE").Range("H" & i).Value)
    Next
    GHEAToEnglish = TEXT
End Function

'change TimesArmenian to GHEAgrapalat
Function TimesToGHEA(TEXT As String) As String
    'Sheets("BASE").Range("A15").Value
    For i = 2 To 39
        TEXT = Replace(TEXT, Sheets("BASE").Range("K" & i).Value, Sheets("BASE").Range("J" & i).Value)
    Next
    For i = 2 To 40
        TEXT = Replace(TEXT, Sheets("BASE").Range("M" & i).Value, Sheets("BASE").Range("L" & i).Value)
    Next
    TimesToGHEA = TEXT
End Function


