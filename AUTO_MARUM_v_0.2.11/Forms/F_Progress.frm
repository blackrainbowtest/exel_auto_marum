VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F_Progress 
   Caption         =   "∂·ÒÕ¡›√≥Û¡ ¡›√≥Û˘« ŸªÁ ø..."
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   6240
   OleObjectBlob   =   "F_Progress.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "F_Progress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    If Button = 2 Then Unload Me
End Sub

Private Sub UserForm_Initialize()
    Me.PrBar.Caption = ""
End Sub



