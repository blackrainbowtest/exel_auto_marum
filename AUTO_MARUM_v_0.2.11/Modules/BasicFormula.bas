Attribute VB_Name = "BasicFormula"
'Main Function
Function AramTivTar(ByVal MyNumber)
'if a non-numeric value is entered
If IsNumeric(MyNumber) = False Then
    Exit Function
End If
    Dim Drams, Cents, Temp
    Dim DecimalPlace, Count
    ReDim Place(9) As String
    Place(2) = Sheets("BASE").Range("D3").Value
    Place(3) = Sheets("BASE").Range("D4").Value
    Place(4) = Sheets("BASE").Range("D5").Value
    Place(5) = Sheets("BASE").Range("D6").Value
    ' String representation of amount.
    MyNumber = Trim(str(MyNumber))
    ' Position of decimal place 0 if none.
    DecimalPlace = InStr(MyNumber, ".")
    ' Convert cents and set MyNumber to dollar amount.
    If DecimalPlace > 0 Then
        Cents = GetTens(Left(Mid(MyNumber, DecimalPlace + 1) & _
                  "00", 2))
        MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))
    End If
    Count = 1
    Do While MyNumber <> ""
        Temp = GetHundreds(Right(MyNumber, 3))
        If Temp <> "" Then Drams = Temp & Place(Count) & Drams
        If Len(MyNumber) > 3 Then
            MyNumber = Left(MyNumber, Len(MyNumber) - 3)
        Else
            MyNumber = ""
        End If
        Count = Count + 1
    Loop
    Select Case Drams
        Case ""
            Drams = Sheets("BASE").Range("D8").Value
        Case "One"
            Drams = "One Dollar"
         Case Else
            Drams = Drams
    End Select
    Select Case Cents
        Case ""
            Cents = ""
        Case "One"
            Cents = " and One Cent"
              Case Else
            Cents = Sheets("BASE").Range("D7").Value & Cents
    End Select
    AramTivTar = "/" & Drams & Cents & "/"
    AramTivTar = Replace(AramTivTar, " /", "/")
    AramTivTar = Replace(AramTivTar, "  ", " ")
End Function

' Converts a number from 100-999 into text
Function GetHundreds(ByVal MyNumber)
    Dim Result As String
    If Val(MyNumber) = 0 Then Exit Function
    MyNumber = Right("000" & MyNumber, 3)
    ' Convert the hundreds place.
    If Mid(MyNumber, 1, 1) <> "0" Then
        Result = GetDigit(Mid(MyNumber, 1, 1)) & Sheets("BASE").Range("D2").Value
    End If
    ' Convert the tens and ones place.
    If Mid(MyNumber, 2, 1) <> "0" Then
        Result = Result & GetTens(Mid(MyNumber, 2))
    Else
        Result = Result & GetDigit(Mid(MyNumber, 3))
    End If
    GetHundreds = Result
End Function

' Converts a number from 10 to 99 into text.
Function GetTens(TensText)
    Dim Result As String
    Result = ""           ' Null out the temporary function value.
    If Val(Left(TensText, 1)) = 1 Then   ' If value between 10-19...
        Select Case Val(TensText)
            Case 10: Result = Sheets("BASE").Range("B2").Value
            Case 11: Result = Sheets("BASE").Range("B3").Value
            Case 12: Result = Sheets("BASE").Range("B4").Value
            Case 13: Result = Sheets("BASE").Range("B5").Value
            Case 14: Result = Sheets("BASE").Range("B6").Value
            Case 15: Result = Sheets("BASE").Range("B7").Value
            Case 16: Result = Sheets("BASE").Range("B8").Value
            Case 17: Result = Sheets("BASE").Range("B9").Value
            Case 18: Result = Sheets("BASE").Range("B10").Value
            Case 19: Result = Sheets("BASE").Range("B11").Value
            Case Else
        End Select
    Else                                 ' If value between 20-99...
        Select Case Val(Left(TensText, 1))
            Case 2: Result = Sheets("BASE").Range("C2").Value
            Case 3: Result = Sheets("BASE").Range("C3").Value
            Case 4: Result = Sheets("BASE").Range("C4").Value
            Case 5: Result = Sheets("BASE").Range("C5").Value
            Case 6: Result = Sheets("BASE").Range("C6").Value
            Case 7: Result = Sheets("BASE").Range("C7").Value
            Case 8: Result = Sheets("BASE").Range("C8").Value
            Case 9: Result = Sheets("BASE").Range("C9").Value
            Case Else
        End Select
        Result = Result & GetDigit _
            (Right(TensText, 1))  ' Retrieve ones place.
    End If
    GetTens = Result
End Function

' Converts a number from 1 to 9 into text.
Function GetDigit(Digit)
    Select Case Val(Digit)
        Case 1: GetDigit = Sheets("BASE").Range("A2").Value
        Case 2: GetDigit = Sheets("BASE").Range("A3").Value
        Case 3: GetDigit = Sheets("BASE").Range("A4").Value
        Case 4: GetDigit = Sheets("BASE").Range("A5").Value
        Case 5: GetDigit = Sheets("BASE").Range("A6").Value
        Case 6: GetDigit = Sheets("BASE").Range("A7").Value
        Case 7: GetDigit = Sheets("BASE").Range("A8").Value
        Case 8: GetDigit = Sheets("BASE").Range("A9").Value
        Case 9: GetDigit = Sheets("BASE").Range("A10").Value
        'while get unknown
        Case Else: GetDigit = ""
    End Select
End Function '

Function MoveFIO(sVal As String, Optional sDelim As String = "/")
    Dim asp, asf
    Dim sBefore As String, sAfter As String, sres As String, s As String
    Dim lp As Long
    lp = InStr(1, sVal, sDelim, 1)
    If lp > 1 Then
        sBefore = Trim(Mid(sVal, 1, lp - 1))
        sAfter = Trim(Mid(sVal, lp + 1, Len(sVal) - lp))
        If sAfter <> "" Then
            asp = Split(sAfter, ", ")
            For lp = LBound(asp) To UBound(asp)
                s = asp(lp)
                asf = Split(s, " ")
                s = asf(1) & " " & asf(0)
                If sres = "" Then
                    sres = s
                Else
                    sres = sres & ", " & s
                End If
            Next
        End If
    End If
    MoveFIO = sBefore & " " & sDelim & " " & sres
End Function


