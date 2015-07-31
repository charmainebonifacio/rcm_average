Attribute VB_Name = "Step_3_Pivot"
Public headerArray() As String
Public headerArray2() As String
'---------------------------------------------------------------------
' Date Created : January 9, 2015
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : April 07, 2015
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : CreatePivotTableTemperature
' Description  : This function creates a pivot table for every data
'                cited on the currently activeworkbook. This will be
'                used in temperature data.
' Parameters   : Workbook
' Returns      : -
'---------------------------------------------------------------------
Function CreatePivotTable(ByRef wbSource As Workbook, _
ByRef SourceSheet As Worksheet, ByVal VarType As String)

    Dim PivotSheet As Worksheet
    Dim pivotTableName As String
    Dim pivotSheetName As String
    Dim sumOfElement As String
    Dim aveOfElement As String
    Dim items As Range
    Dim VARTOPIVOT As String
    
    ' Status Bar Update
    appSTATUS = "Creating Pivot Table of Monthly Total per Year."
    Application.StatusBar = appSTATUS
    logtxt = appSTATUS
    logfile.WriteLine logtxt
    
    ' This is for the pivot table
    pivotTableName = "PTable"
    pivotSheetName = "PT_"
    
    ' This is for data field
    sumOfElement = "Sum of "
    aveOfElement = "Average of "
    
    ' Initialize Variable to be Pivoted
    VARTOPIVOT = "AVG_" & VarType
    If VarType = "PPT" Then VARTOPIVOT = "SUM_" & VarType
    
    ' Handle Alerts
    Application.DisplayAlerts = False
    
    ' Workbook Settings
    Set wbSource = ActiveWorkbook
    Set SourceSheet = wbSource.Worksheets(1)
    
    ' Set Header
    SourceSheet.Activate
    ReDim headerArray(0 To 1)
    headerArray(0) = SourceSheet.Range("A1").Value
    headerArray(1) = "MONTH"
    
    ' Data Source
    SourceSheet.Range("A1").Select
    SourceSheet.Range(Selection, Selection.End(xlToRight)).Select
    SourceSheet.Range(Selection, Selection.End(xlDown)).Select
    
    ' Set selected items as current selection
    Set items = Selection
            
    ' Creates the pivot table in the worksheet
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
                items, Version:=xlPivotTableVersion12).CreatePivotTable _
                TableDestination:="", TableName:=pivotTableName, DefaultVersion _
                :=xlPivotTableVersion12
            
    ' Always move new pivot table worksheet to the end to keep it organized
    ActiveSheet.Move After:=Sheets(Sheets.Count)
    ActiveSheet.Name = pivotSheetName & "DATA"
    ActiveSheet.PivotTables(pivotTableName).Location = "$A$3"
    Cells(3, 1).Select
    
    'ActiveSheet.PivotTables(pivotTableName).AddDataField
    With ActiveSheet.PivotTables(pivotTableName).PivotFields(headerArray(0))
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables(pivotTableName).PivotFields(headerArray(1))
        .Orientation = xlColumnField
        .Position = 1
    End With
    ActiveSheet.PivotTables(pivotTableName).AddDataField ActiveSheet.PivotTables( _
        pivotTableName).PivotFields(VARTOPIVOT), aveOfElement & VARTOPIVOT, _
        xlAverage

    Dim LR As Long, LC As Long
    Call FindLastRowColumn(LR, LC)
    Debug.Print LR & "-" & LC

    ' Handle Alerts
    Application.DisplayAlerts = True

Cancel:
    Set WKBook = Nothing
    Set DestSheet = Nothing
    Set SourceSheet = Nothing
End Function
