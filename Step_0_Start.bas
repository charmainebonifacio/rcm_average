Attribute VB_Name = "Step_0_Start"
Public refIDArray() As String
'---------------------------------------------------------------------
' Date Created : January 9, 2015
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : January 9, 2015
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : Start_Here
' Description  : The purpose of function is to initialize the userform.
'---------------------------------------------------------------------
Sub Start_Here()
   
    Dim myForm As UserForm1
    Dim button1 As String, button2 As String, button3 As String
    Dim button4 As String, button5 As String, button6 As String
    Dim strLabel1 As String, strLabel2 As String
    Dim strLabel3 As String, strLabel4 As String
    Dim strLabel5 As String, strLabel6 As String
    Dim strLabel7 As String, strLabel8 As String
    Dim frameLabel1 As String, frameLabel2 As String, frameLabel3 As String
    Dim userFormCaption As String
    
    ' Initialize Drop Down Menu
    ReDim refIDArray(0 To 2)
    refIDArray(0) = "PPT"
    refIDArray(1) = "TMX"
    refIDArray(2) = "TMN"
    
    ' Disable all the pop-up menus
    Application.ScreenUpdating = False
    Set myForm = UserForm1
    
    ' Label Strings
    userFormCaption = "KIENZLE LAB TOOLS"
    button1 = "CREATE SUMMARY WORKSHEET"
    button2 = "ALBERTA 10K FILE"
    button3 = "HRU FILE"
    frameLabel2 = "TOOL GUIDE"
    frameLabel3 = "HELP SECTION"
    
    strLabel1 = "THE RCM MONTHLY AVERAGE SUMMARY MACRO"
    strLabel2 = "STEP 1."
    strLabel3 = "Export table from Database as .CSV file" & vbLf
    strLabel4 = "STEP 2."
    strLabel5 = "For more information, hover mouse over button."
    strLabel7 = "STEP 3."
    strLabel8 = "Select variable type: "
    
    ' UserForm Initialize
    myForm.Caption = userFormCaption
    myForm.Frame2.Caption = frameLabel2
    myForm.Frame5.Caption = frameLabel3
    myForm.Frame2.Font.Bold = True
    myForm.Frame5.Font.Bold = True
    myForm.Label1.Caption = strLabel1
    myForm.Label1.Font.Size = 21
    myForm.Label1.Font.Bold = True
    myForm.Label1.TextAlign = fmTextAlignCenter
    
    myForm.Label2 = strLabel2
    myForm.Label2.Font.Size = 13
    myForm.Label2.Font.Bold = True
    myForm.Label3 = strLabel3
    myForm.Label3.Font.Size = 11
    myForm.Label4 = strLabel4
    myForm.Label4.Font.Size = 13
    myForm.Label4.Font.Bold = True
    myForm.Label7 = strLabel7
    myForm.Label7.Font.Size = 13
    myForm.Label7.Font.Bold = True
    myForm.Label8 = strLabel8
    myForm.Label8.Font.Size = 13
    myForm.CommandButton1.Caption = button1
    myForm.CommandButton1.Font.Size = 11
    
    ' Add Each Item to the Drop Down List
    For i = LBound(refIDArray) To UBound(refIDArray)
        myForm.VarBox.AddItem refIDArray(i), i
    Next i

    ' Help File
    myForm.Label5 = strLabel5
    myForm.Label5.Font.Size = 8
    myForm.Label5.Font.Italic = True
    
    Application.StatusBar = "Macro has been initiated."
    myForm.Show

End Sub
'---------------------------------------------------------------------------------------
' Date Created : January 9, 2015
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : January 10, 2015
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : HELPFILE
' Description  : This function will feed the help tip section depending on the button
'                that has been activated.
' Parameters   : String
' Returns      : String
'---------------------------------------------------------------------------------------
Function HELPFILE(ByVal Notification As Integer) As String

    Dim NotifyUser As String
    
    Select Case Notification
        Case 1
            NotifyUser = "TITLE: THE GRID FILE MONTHLY AVERAGE SUMMARY MACRO" & vbLf
            NotifyUser = NotifyUser & "DESCRIPTION: This macro will convert the exported monthly average " & _
                "summary table in a proper format that can be used with a .DBF file. " & vbLf
            NotifyUser = NotifyUser & "INPUT: Find the location of the exported table in .CSV format" & vbLf
            NotifyUser = NotifyUser & "OUTPUT: .CSV and .XLSX files per file" & vbLf
    End Select
    
    HELPFILE = NotifyUser
    
End Function
'---------------------------------------------------------------------
' Date Created : January 9, 2015
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : January 10, 2015
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : SaveLogFile
' Description  : This function saves file as .TXT.
'                When new file is named after an existing file, the
'                same name is used with an number attached to it.
' Parameters   : String, String
' Returns      : -
'---------------------------------------------------------------------
Function SaveLogFile(ByVal fileDir As String, _
ByVal fileName As String, ByVal fileExt As String) As String

    Dim saveFile As String
    Dim formatDate As String
    Dim saveDate As String
    Dim saveName As String
    Dim sPath As String

    ' Date
    formatDate = Format(Date, "yyyy/mm/dd")
    saveDate = Replace(formatDate, "/", "")
    
    ' Save information as Temp, which can then be renamed later..
    sPath = fileDir
    If Right(fileDir, 1) <> "\" Then sPath = fileDir & "\"
    saveName = fileName & "_" & saveDate & fileExt
    
    ' Rename existing file
    i = 1
    If CheckFileExists(sPath, saveName) = True Then
        If Dir(sPath & fileName & "_" & saveDate & "_" & i & fileExt) <> "" Then
            Do Until Dir(sPath & fileName & "_" & saveDate & "_" & i & fileExt) = ""
                i = i + 1
            Loop
            saveFile = sPath & fileName & "_" & saveDate & "_" & i & fileExt
        Else: saveFile = sPath & fileName & "_" & saveDate & "_" & i & fileExt
        End If
    Else: saveFile = sPath & fileName & "_" & saveDate & fileExt
    End If
    
    SaveLogFile = saveFile
    
End Function

