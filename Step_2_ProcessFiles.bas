Attribute VB_Name = "Step_2_ProcessFiles"
Public FinalHeader() As String
'---------------------------------------------------------------------
' Date Created : January 9, 2015
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : January 10, 2015
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : PROCESSDBFFILES
' Description  : This function processes all the .dbf files within a
'                file directory. The .dbf files must be numbered
'                chronologically. User must select the first file that
'                denotes the first month of the year (ie. January).
'                The function returns the file directory which will
'                be used in later functions.
' Parameters   : -
' Returns      : String
'---------------------------------------------------------------------
Function PROCESSFILES(ByVal fileDir As String, ByVal outDir As String) As String

    Dim objFolder As Object, objFSO As Object
    Dim wbSource As Workbook, SourceSheet As Worksheet
    Dim wbDest As Workbook, DestSheet As Worksheet
    Dim FileCounter As Long
    Dim sThisFilePath As String, sFile As String
    Dim GridName As String
    Dim VarType As String
    
    ' Disable all the pop-up menus
    Application.ScreenUpdating = False

    ' Status Bar Update
    appSTATUS = "Processing folder for each file...."
    Application.StatusBar = appSTATUS
    logtxt = appSTATUS
    logfile.WriteLine logtxt
    
    ' Initialize Variables
    FileCounter = 0
    VarType = UserForm1.VarBox.Text
    
    '-------------------------------------------------------------
    ' Check the files... which should be obvious at this point.
    '-------------------------------------------------------------
    sThisFilePath = fileDir
    If (Right(sThisFilePath, 1) <> "\") Then sThisFilePath = sThisFilePath & "\"
    sFile = Dir(sThisFilePath & "*.csv")

    '-------------------------------------------------------------
    ' Loop through all the .txt files
    '-------------------------------------------------------------
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.GetFolder(sThisFilePath).Files
    
    For Each objFILE In objFolder
        logtxt = objFILE
        Debug.Print objFILE
        logfile.WriteLine objFILE
        
        If UCase(Right(objFILE.Path, (Len(objFILE.Path) - InStrRev(objFILE.Path, ".")))) = UCase("csv") Then
            FileCounter = FileCounter + 1
            logtxt = FileCounter & " of files processed."
            Debug.Print logtxt
            logfile.WriteLine logtxt
            
            ' Add New WorkBook
            Set wbDest = Workbooks.Add(1)
            Set DestSheet = wbDest.Worksheets(1)
            DestSheet.Name = "ORIG_DATA"
            
            ' Open file and set it as source worksheet
            Set DestSheet = wbDest.Worksheets(1)
            Set wbSource = Workbooks.Open(objFILE.Path)
            Set SourceSheet = wbSource.Worksheets(1)
            SourceSheet.Activate
            GridName = SourceSheet.Name
            
            ' Copy / Paste Data
            Call DatabaseFile(SourceSheet, DestSheet)
            Call CheckHeaders(DestSheet)
            wbSource.Close SaveChanges:=False

            ' Process New Files
            Call CreatePivotTable(wbSource, SourceSheet, VarType)
            Call CopyPivotSummary(wbSource, SourceSheet, VarType)
            
            ' Ignore Clipboard Alerts
            Application.CutCopyMode = True
            
            ' Save Changes to the Processed Files
            Call SaveFileAs(wbDest, outDir, 1, GridName & "_" & VarType)
            Call SaveFileAs(wbDest, outDir, 2, GridName & "_" & VarType)
            wbDest.Close SaveChanges:=False

        Else:
            logtxt = objFILE & " is not a valid file to process."
            Debug.Print logtxt
            logfile.WriteLine logtxt
        End If
    Next
    
    appSTATUS = "Processed .txt files for " & VarSelected & "."
    Application.StatusBar = appSTATUS
    logtxt = appSTATUS
    logfile.WriteLine logtxt
    
    Application.StatusBar = False
    
Cancel:
    Set wbSource = Nothing
    Set SourceSheet = Nothing
    Set wbDest = Nothing
    Set DestSheet = Nothing
    Set objFSO = Nothing
    Set objFolder = Nothing
End Function
'---------------------------------------------------------------------
' Date Created : June 3, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : June 6, 2013
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : DatabaseFile
' Description  : This function copies the data from the zonal stats
'                .dbf file.
' Parameters   : Worksheet, Worksheet
' Returns      : -
'---------------------------------------------------------------------
Function DatabaseFile(SourceSht As Worksheet, DestSht As Worksheet)
    
    Dim RngSelect
    Dim PasteSelect

    ' Disable all the pop-up menus
    Application.ScreenUpdating = False

    ' Activate Source Worksheet.
    SourceSht.Activate
      
    '-------------------------------------------------------------
    ' Call FindRange function to select the current used data
    ' within the Source Worksheet. Only copy the selected data.
    '-------------------------------------------------------------
    Call FindRange(SourceSht)
    RngSelect = Selection.Address
    Range(RngSelect).Copy

    ' Activate Destination Worksheet.
    DestSht.Activate
    
    '-------------------------------------------------------------
    ' Call RowCheck function to check the last row.
    ' Then append the copied data into the Destination Worksheet.
    '-------------------------------------------------------------
    Call RowCheck(DestSht)
    PasteSelect = Selection.Address
    Range(PasteSelect).Select
    DestSht.Paste
    
    ' Clear Clipboard of any copied data.
    Application.CutCopyMode = False
    
End Function
'---------------------------------------------------------------------
' Date Created : June 3, 2013
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : March 31, 2015
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : DatabaseFile
' Description  : This function changes the headers
' Parameters   : Worksheet
' Returns      : -
'---------------------------------------------------------------------
Function CheckHeaders(DestSht As Worksheet)

    Dim D As String
    Dim E As String
    Dim F As String
    
    DestSht.Activate
    D = "SUM(PPT)"
    E = "AVG(Tmax)"
    F = "AVG(Tmin)"
    If Range("D1").Value = D Then
        Range("D1").Value = "SUM_PPT"
    End If
    If Range("E1").Value = E Then
        Range("E1").Value = "AVG_TMX"
    End If
    If Range("F1").Value = F Then
        Range("F1").Value = "AVG_TMN"
    End If
    
End Function
    
