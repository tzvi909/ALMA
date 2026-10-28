Attribute VB_Name = "ALMA_Main"
Option Explicit
'==============================================================================
' ALMA for Excel - macros behind the buttons on the ALMA sheet.
' Every button calls RunAction "<action>".  RunAction can also be called from
' other code with a target sheet name (the automated tests do this).
'
' Lists sheet columns (data from row 2 down, a list ends at the first blank):
'   A List 1 (rows)              B List 2 (columns)
'   C Comparison strings (var)   D Constant strings added to every query
'   E Text generator 1st range   F Text generator 2nd range
'   G Filter: hypotheses         H Filter: values (years, countries)
'   J/K parameters (label / value)
'==============================================================================

Public Const LISTS_SHEET As String = "Lists"

'------------------------------------------------------------- button macros
Public Sub ALMA_FullPipeline()
    RunAction "pipeline"
End Sub


Public Sub ALMA_NewSearchMatrix()
    RunAction "search"
End Sub


Public Sub ALMA_AddConstantString()
    RunAction "addstring"
End Sub


Public Sub ALMA_ComparisonMatrix()
    RunAction "comparison"
End Sub


Public Sub ALMA_ComparisonList()
    RunAction "comparisonlist"
End Sub


Public Sub ALMA_NRFRanking()
    RunAction "nrf"
End Sub


Public Sub ALMA_NormalizeRows()
    RunAction "normrows"
End Sub


Public Sub ALMA_NormalizeCols()
    RunAction "normcols"
End Sub


Public Sub ALMA_SortBySum()
    RunAction "sortsum"
End Sub


Public Sub ALMA_SortByMax()
    RunAction "sortmax"
End Sub


Public Sub ALMA_SortByColumn()
    RunAction "sortcol"
End Sub


Public Sub ALMA_ClusterSort()
    RunAction "cluster"
End Sub


Public Sub ALMA_Colorize()
    RunAction "colorize"
End Sub


Public Sub ALMA_TextGenerator()
    RunAction "textgen"
End Sub


Public Sub ALMA_FilterMatrix()
    RunAction "filter"
End Sub


Public Sub ALMA_ClearCache()
    RunAction "clearcache"
End Sub


Public Sub ALMA_SelfTest()
    RunAction "selftest"
End Sub


Public Sub ALMA_ShowLists()
    ThisWorkbook.Worksheets(LISTS_SHEET).Activate
End Sub


Public Sub ALMA_ShowTree()
    ThisWorkbook.Worksheets(TREE_SHEET).Activate
End Sub


'------------------------------------------------------------------ dispatch
Public Function RunAction(ByVal action As String, Optional ByVal target As String = "") As String
    Dim result As String, msg As String
    On Error GoTo failed
    BeginWork
    gFetched = 0
    gFromCache = 0
    Select Case LCase$(action)
        Case "pipeline"
            result = DoPipeline()
        Case "search"
            result = DoSearch()
        Case "addstring"
            result = DoAddString(target)
        Case "comparison"
            result = DoComparison(target, False)
        Case "comparisonlist"
            result = DoComparison(target, True)
        Case "nrf"
            result = DoNrf(target)
        Case "normrows"
            result = DoNormalize(target, "rows")
        Case "normcols"
            result = DoNormalize(target, "cols")
        Case "sortsum"
            result = DoSort(target, "sum")
        Case "sortmax"
            result = DoSort(target, "max")
        Case "sortcol"
            result = DoSortColumn(target)
        Case "cluster"
            result = DoCluster(target)
        Case "colorize"
            result = DoColorize(target)
        Case "textgen"
            result = DoTextGen()
        Case "filter"
            result = DoFilter()
        Case "clearcache"
            result = DoClearCache()
        Case "selftest"
            result = DoSelfTest()
        Case Else
            Err.Raise ALMA_ERR + 799, "ALMA", "Unknown action: " & action
    End Select
    EndWork
    result = result & "   [PubMed: " & gFetched & " fetched, " & gFromCache & " from cache]"
    SetSetting "laststatus", "OK " & action & ": " & result
    If Not Quiet() Then MsgBox result, vbInformation, "ALMA"
    RunAction = result
    Exit Function
failed:
    msg = "ERROR in " & action & " (at " & gStep & "): " & Err.Description
    EndWork
    SetSetting "laststatus", msg
    If Not Quiet() Then MsgBox msg, vbExclamation, "ALMA"
    RunAction = msg
End Function


Private Sub BeginWork()
    On Error Resume Next
    Application.ScreenUpdating = False
    Application.EnableCancelKey = 2          ' xlErrorHandler: Esc stops cleanly
    Application.Calculation = -4135          ' xlCalculationManual
    On Error GoTo 0
End Sub


Private Sub EndWork()
    On Error Resume Next
    Application.Calculation = -4105          ' xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
    On Error GoTo 0
End Sub


'--------------------------------------------------------------------- lists
Public Function ListItems(ByVal col As Long, ByRef n As Long) As Variant
    Dim ws As Worksheet, r As Long, items() As String, v As String
    Set ws = ThisWorkbook.Worksheets(LISTS_SHEET)
    n = 0
    ReDim items(1 To 1)
    r = 2
    Do
        v = Trim$(CStr(ws.Cells(r, col).Value))
        If Len(v) = 0 Then Exit Do
        n = n + 1
        ReDim Preserve items(1 To n)
        items(n) = v
        r = r + 1
    Loop
    ListItems = items
End Function


Public Function OneItem(ByVal s As String) As Variant
    Dim items() As String
    ReDim items(1 To 1)
    items(1) = s
    OneItem = items
End Function


Public Function JoinList(ByVal col As Long) As String
    Dim items As Variant, n As Long, i As Long, s As String
    items = ListItems(col, n)
    s = ""
    For i = 1 To n
        If Len(s) > 0 Then s = s & ALMA_SEP
        s = s & items(i)
    Next i
    JoinList = s
End Function


Public Function ListSetting(ByVal label As String, ByVal dflt As String) As String
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets(LISTS_SHEET)
    ListSetting = dflt
    For r = 1 To 40
        If LCase$(Trim$(CStr(ws.Cells(r, 10).Value))) = LCase$(label) Then
            If Len(Trim$(CStr(ws.Cells(r, 11).Value))) > 0 Then ListSetting = Trim$(CStr(ws.Cells(r, 11).Value))
            Exit Function
        End If
    Next r
End Function
