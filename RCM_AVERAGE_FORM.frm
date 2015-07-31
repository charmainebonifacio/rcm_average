VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RCM_AVERAGE_FORM 
   Caption         =   "UserForm1"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8640
   OleObjectBlob   =   "RCM_AVERAGE_FORM.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RCM_AVERAGE_FORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Date Created : January 9, 2015
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : January 10, 2015
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
'---------------------------------------------------------------------------------------

Private Sub CommandButton1_Click()
' Download Environment Data

    If Val(Application.Version) < 12 Then
        MsgBox "You are using Microsoft Excel 2003 and older."
    Else
        UserForm1.Hide
        Debug.Print "Microsoft Excel 2007 or higher."
        Call RCM_MAIN
    End If
    Unload UserForm1

End Sub
'---------------------------------------------------------------------------------------
' Date Created : July 18, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : July 18, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Description  : Placed section for processing Help Tips.
'---------------------------------------------------------------------------------------
Private Sub UserForm_Activate()
    UserForm1.Label6.Visible = False
End Sub

Private Sub CommandButton1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Dim strLabel6 As String
    UserForm1.Label6.Visible = True
    UserForm1.Label6.BackColor = RGB(255, 255, 153)
    strLabel6 = HELPFILE(1)
    UserForm1.Label6 = strLabel6
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    UserForm1.Label6.Visible = False
End Sub

Private Sub VarBox_Change()

End Sub
