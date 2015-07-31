Attribute VB_Name = "Step_1_Main"
Public objFSOlog As Object
Public logfile As TextStream
Public logtxt As String
Public appSTATUS As String
'---------------------------------------------------------------------------------------
' Date Created : January 9, 2015
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Date Edited  : April 8, 2015
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : RCM_MAIN
' Description  : This is the main function process the .CSV files from database and
'                convert the data into proper format.
'---------------------------------------------------------------------------------------
Function RCM_MAIN()

    Dim start_time As Date, end_time As Date
    Dim ProcessingTime As Long
    Dim MessageSummary As String, SummaryTitle As String
    
    Dim UserSelectedFolder As String, DBFDIR As String
    Dim MAINFolder As String, compareIndex As Integer
    Dim PROGDIR As String, outDir As String, AB10KDIR As String
    Dim CopiedFiles As Long
    
    Dim MainOUT As String, AB10KOUT As String
    Dim TmpOUT As String, NewOUT As String
    Dim refIDArray() As String, refIndex As Integer
    Dim VarSelected As String
    
    ' Initialize Variables
    SummaryTitle = "Zonal Statistics Macro Diagnostic Summary"
    outDir = "Output"
    VarSelected = UserForm1.VarBox.Text
    
    ' Disable all the pop-up menus
    Application.ScreenUpdating = False
        
    '---------------------------------------------------------------------
    ' I. USER INPUT
    '---------------------------------------------------------------------
    UserSelectedFolder = GetFolder
    Debug.Print UserSelectedFolder
    If Len(UserSelectedFolder) = 0 Then GoTo Cancel
    MAINFolder = ReturnFolderName(UserSelectedFolder)
    Debug.Print MAINFolder
    
    ' Start Time
    start_time = Now()
  
    '---------------------------------------------------------------------
    ' II. Create output folder for processing grid files within the
    ' GridFile.
    '---------------------------------------------------------------------
    TmpOUT = ReturnSubFolder(UserSelectedFolder, outDir)
    CheckOUTFolder = CheckFolderExists(TmpOUT)
    Debug.Print CheckOUTFolder
    If CheckOUTFolder = False Then NewOUT = CreateFolder(TmpOUT)
    
    ' Setup Log File
    Dim logfilename As String, logtextfile As String, logext As String
    logext = ".txt"
    logfilename = "corrzs_log_" & VarSelected
    logtextfile = SaveLogFile(TmpOUT, logfilename, logext)
    
    Set objFSOlog = CreateObject("Scripting.FileSystemObject")
    Set logfile = objFSOlog.CreateTextFile(logtextfile, True)
    
    ' Maintain log starting from here
    logfile.WriteLine " [ Start of Program. ] "
    logfile.WriteLine "Selected directory: " & UserSelectedFolder
    logfile.WriteLine "Main directory: " & MAINFolder
    logfile.WriteLine "Output directory: " & TmpOUT
    
    '---------------------------------------------------------------------
    ' III. PROCESS ALL .CSV FILES
    '---------------------------------------------------------------------
    Call WarningMessage
    Call PROCESSFILES(UserSelectedFolder, TmpOUT)

    ' End Time
    logfile.WriteLine " [ End of Program. ] "
    end_time = Now()
    
    ' Time Elapsed
    ProcessingTime = DateDiff("n", CDate(start_time), CDate(end_time))
    MessageSummary = MacroTimer(ProcessingTime)
    MsgBox MessageSummary, vbOKOnly, SummaryTitle

    ' Close Log File
    logfile.Close
    Set logfile = Nothing
    Set objFSOlog = Nothing
    
Cancel:
    If Len(UserSelectedFolder) = 0 Then
        MsgBox "No folder selected.", vbOKOnly, SummaryTitle
    End If
End Function
