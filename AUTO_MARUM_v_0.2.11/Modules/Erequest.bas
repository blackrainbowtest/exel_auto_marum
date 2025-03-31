Attribute VB_Name = "Erequest"
Sub OpenEdgeWithTrackingCodes()
    Dim cell As Range
    Dim trackingCode As String
    Dim url As String
    Dim edge As Object

    ' Check if a range is selected
    If Selection Is Nothing Then
        MsgBox "Please select a range of cells!", vbExclamation
        Exit Sub
    End If

    ' Create a Shell object for Edge
    Set edge = CreateObject("Shell.Application")

    ' Loop through each selected cell
    For Each cell In Selection
        ' Get the value from the current cell
        trackingCode = Trim(cell.Value)
        
        ' Check if the cell is not empty
        If trackingCode <> "" Then
            ' Construct the URL
            url = "https://e-request.am/hy/e-letter/check?trackingC%D6%85de=" & trackingCode
            
            ' Open Microsoft Edge with the constructed URL
            edge.ShellExecute "microsoft-edge:" & url
            
            ' Pause for 2 seconds before opening the next link (adjust if necessary)
            Application.Wait Now + TimeValue("00:00:02")
        End If
    Next cell

    ' Release the object
    Set edge = Nothing

    ' MsgBox "All tracking codes have been processed!", vbInformation
End Sub
