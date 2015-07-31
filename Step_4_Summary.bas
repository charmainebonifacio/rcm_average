Attribute VB_Name = "Step_4_Summary"
'---------------------------------------------------------------------
' Date Created : January 9, 2015
' Created By   : Charmaine Bonifacio
'---------------------------------------------------------------------
' Date Edited  : March 15, 2015
' Edited By    : Charmaine Bonifacio
' Comments By  : Charmaine Bonifacio
'---------------------------------------------------------------------
' Organization : Department of Geography, University of Lethbridge
' Title        : CopyPivotSummary
' Description  : This function creates a pivot table for every data
'                cited on the currently activeworkbook.
' Parameters   : Workbook, Worksheet
' Returns      : -
'---------------------------------------------------------------------
Function CopyPivotSummary(ByRef wbSource As Workbook, _
ByRef SourceSheet As Worksheet, ByVal VarType As String)
    
    Dim TmpSheet As Worksheet
    Dim pivotTableName As String
    Dim pivotSheetName As String
    Dim sumOfElement As String
    Dim aveOfElement As String
    Dim items As Range
    Dim LR As Long, LC As Long
    
    ' Handle Alerts
    Application.DisplayAlerts = False

    Set wbSource = ActiveWorkbook
    Set SourceSheet = wbSource.Worksheets(2)
    Set TmpSheet = Sheets.Add(After:=Sheets(Sheets.Count))
    TmpSheet.Name = "SUMMARY_" & VarType
    
    ' Data Source
    SourceSheet.Activate
    Call FindLastRowColumn(LR, LC)
    Debug.Print LR & "-" & LC
    Range(Cells(4, 1), Cells(LR - 1, LC - 1)).Select
    Selection.Copy
    TmpSheet.Activate
    Range("A1").Select
    ActiveSheet.Paste
    
    ' Rename Headers
    TmpSheet.Activate
    Dim VARS As String
    VARS = UserForm1.VarBox.Text
    ReDim headerArray2(0 To 12)
    headerArray2(0) = "RCM_ID"
    headerArray2(1) = VARS & "_01"
    headerArray2(2) = VARS & "_02"
    headerArray2(3) = VARS & "_03"
    headerArray2(4) = VARS & "_04"
    headerArray2(5) = VARS & "_05"
    headerArray2(6) = VARS & "_06"
    headerArray2(7) = VARS & "_07"
    headerArray2(8) = VARS & "_08"
    headerArray2(9) = VARS & "_09"
    headerArray2(10) = VARS & "_10"
    headerArray2(11) = VARS & "_11"
    headerArray2(12) = VARS & "_12"
    For i = LBound(headerArray2) To UBound(headerArray2)
        Range("A1").Offset(0, i).Value = headerArray2(i)
    Next i
    Selection.NumberFormat = "00.000"
    
    'Create a separate Column for the new RCM ID with only number - String Format
    TmpSheet.Activate
    Call FindLastRowColumn(LR, LC)
    Debug.Print LR & "-" & LC
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "TID"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "=CLEAN(RIGHT(RC[-1],4))"
    Range("B2").Select
    Selection.AutoFill Destination:=Range("B2:B" & LR)
    
    ' Handle Alerts
    Application.DisplayAlerts = True

Cancel:
    Set TmpSheet = Nothing
End Function


