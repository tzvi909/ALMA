Attribute VB_Name = "ALMA_Core"
Option Explicit
'==============================================================================
' ALMA for Excel - core: matrices on worksheets (layout, reading and writing),
' search and comparison matrices, sorting, normalisation, cluster sorting,
' NRF (novelty / reasonability / feasibility) scoring, colours, tree manager
' and the actions behind the buttons (RunAction in ALMA_Main calls Do...).
' Everything that uses the AlmaMatrix / AlmaComparison types lives in this
' module (some VBA hosts do not share user-defined types between modules).
' Same rules and sheet layout as the Python version (alma/core.py, excel.py):
'   A1 ALMA_MATRIX | B1 kind (search, normalized, comparison, nrf, text)
'   A2 Name, A3 Context, A4 Column tag, A5 Parent, A6 Var / Note, A7 PubMed
'   row 9 = header row of the matrix or table, data from row 10
'==============================================================================

Public Const HDR_ROW As Long = 9
Public Const TREE_SHEET As String = "Tree"
Public Const TREE_FIRST As Long = 6
Public Const NRF_COLS As Long = 18

Public Type AlmaMatrix
    Kind As String
    Title As String
    Context As String
    ColTag As String
    ParentName As String
    Note As String
    Corner As String
    nR As Long
    nC As Long
    RowLab() As String
    ColLab() As String
    V() As Double
End Type

' A comparison matrix keeps its base search matrix inline (no nested type):
' RowLab / ColLab / LR = rows, columns and NOPs of the base matrix.
Public Type AlmaComparison
    Title As String
    VarName As String
    VarTotal As Double
    BaseTitle As String
    Context As String
    ColTag As String
    nR As Long
    nC As Long
    RowLab() As String
    ColLab() As String
    LR() As Double
    RowVar() As Double
    ColVar() As Double
    Tri() As Double
End Type

Private mNrfRec() As Variant
Private mNrfKeys() As Double
Private mNrfN As Long

'==================================================================== sheets
Public Function SheetExists(ByVal nm As String) As Boolean
    Dim i As Long
    SheetExists = False
    For i = 1 To ThisWorkbook.Sheets.Count
        If LCase$(ThisWorkbook.Sheets(i).Name) = LCase$(nm) Then
            SheetExists = True
            Exit Function
        End If
    Next i
End Function


Public Function SafeSheetName(ByVal s As String) As String
    Dim bad As Variant, i As Long
    bad = Array("[", "]", ":", "*", "?", "/", "\", "'")
    For i = LBound(bad) To UBound(bad)
        s = Replace(s, CStr(bad(i)), "_")
    Next i
    s = Trim$(s)
    If Len(s) = 0 Then s = "Sheet"
    SafeSheetName = Left$(s, 31)
End Function


Public Function NewSheet(ByVal title As String) As Worksheet
    Dim nm As String, baseName As String, i As Long, ws As Worksheet
    gStep = "NewSheet"
    baseName = SafeSheetName(title)
    nm = baseName
    i = 1
    Do While SheetExists(nm)
        i = i + 1
        nm = Left$(baseName, 31 - Len("~" & i)) & "~" & i
    Loop
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = nm
    Set NewSheet = ws
End Function


Public Function IsAlmaSheet(ByVal ws As Worksheet) As Boolean
    IsAlmaSheet = (CStr(ws.Cells(1, 1).Value) = "ALMA_MATRIX")
End Function


Public Function SheetKind(ByVal ws As Worksheet) As String
    SheetKind = ""
    If IsAlmaSheet(ws) Then SheetKind = CStr(ws.Cells(1, 2).Value)
End Function


Public Function NumOr0(ByVal v As Variant) As Double
    NumOr0 = 0
    If IsEmpty(v) Then Exit Function
    If IsError(v) Then Exit Function
    If IsNumeric(v) Then NumOr0 = CDbl(v)
End Function


Private Sub WriteHeader(ByVal ws As Worksheet, ByVal kind As String, ByVal title As String, _
                        ByVal context As String, ByVal tag As String, ByVal parentName As String, ByVal note As String)
    Dim lab As Variant, vals As Variant, i As Long
    gStep = "WriteHeader"
    lab = Array("ALMA_MATRIX", "Name", "Context", "Column tag", "Parent", "Var / Note", "PubMed")
    vals = Array(kind, title, context, tag, parentName, note, PubMedNote())
    For i = 0 To 6
        ws.Cells(i + 1, 1).Value = lab(i)
        ws.Cells(i + 1, 1).Font.Color = RGB(128, 128, 128)
        ws.Cells(i + 1, 2).NumberFormat = "@"
        ws.Cells(i + 1, 2).Value = vals(i)
    Next i
    ws.Cells(2, 2).Font.Bold = True
End Sub


Private Sub StyleSheet(ByVal ws As Worksheet, ByVal lastCol As Long)
    gStep = "StyleSheet"
    On Error Resume Next
    ws.Rows(HDR_ROW).Font.Bold = True
    ws.Rows(HDR_ROW).WrapText = True
    ws.Rows(HDR_ROW).RowHeight = 45
    ws.Range(ws.Cells(HDR_ROW, 1), ws.Cells(HDR_ROW + 5000, 1)).Font.Bold = True
    ws.Columns(1).ColumnWidth = 30
    If lastCol >= 2 Then ws.Range(ws.Columns(2), ws.Columns(lastCol)).ColumnWidth = 11
    On Error GoTo 0
End Sub


'============================================================= read / write
Public Function ReadMatrix(ByVal ws As Worksheet) As AlmaMatrix
    Dim m As AlmaMatrix, i As Long, j As Long, data As Variant, nR As Long, nC As Long
    Dim rl() As String, cl() As String, vv() As Double
    gStep = "ReadMatrix"
    If Not IsAlmaSheet(ws) Then
        Err.Raise ALMA_ERR + 600, "ALMA", "Sheet '" & ws.Name & "' is not an ALMA matrix sheet."
    End If
    m.Kind = CStr(ws.Cells(1, 2).Value)
    m.Title = ws.Name
    m.Context = CStr(ws.Cells(3, 2).Value)
    m.ColTag = CStr(ws.Cells(4, 2).Value)
    m.ParentName = CStr(ws.Cells(5, 2).Value)
    m.Note = CStr(ws.Cells(6, 2).Value)
    m.Corner = CStr(ws.Cells(HDR_ROW, 1).Value)
    nC = 0
    Do While Len(CStr(ws.Cells(HDR_ROW, nC + 2).Value)) > 0
        nC = nC + 1
    Loop
    nR = 0
    Do While Len(CStr(ws.Cells(HDR_ROW + nR + 1, 1).Value)) > 0
        nR = nR + 1
    Loop
    If nR = 0 Or nC = 0 Then
        Err.Raise ALMA_ERR + 601, "ALMA", "No matrix found on sheet '" & ws.Name & "'."
    End If
    ReDim rl(1 To nR)
    ReDim cl(1 To nC)
    ReDim vv(1 To nR, 1 To nC)
    For j = 1 To nC
        cl(j) = CStr(ws.Cells(HDR_ROW, j + 1).Value)
    Next j
    For i = 1 To nR
        rl(i) = CStr(ws.Cells(HDR_ROW + i, 1).Value)
    Next i
    If nR = 1 And nC = 1 Then
        vv(1, 1) = NumOr0(ws.Cells(HDR_ROW + 1, 2).Value)
    Else
        data = ws.Range(ws.Cells(HDR_ROW + 1, 2), ws.Cells(HDR_ROW + nR, nC + 1)).Value
        For i = 1 To nR
            For j = 1 To nC
                vv(i, j) = NumOr0(data(i, j))
            Next j
        Next i
    End If
    m.nR = nR
    m.nC = nC
    m.RowLab = rl
    m.ColLab = cl
    m.V = vv
    ReadMatrix = m
End Function


Public Function WriteMatrix(m As AlmaMatrix, ByVal title As String) As Worksheet
    Dim ws As Worksheet, i As Long, j As Long, lab() As Variant, arr() As Variant, rng As Range
    gStep = "WriteMatrix"
    Set ws = NewSheet(title)
    WriteHeader ws, m.Kind, title, m.Context, m.ColTag, m.ParentName, m.Note
    ws.Cells(HDR_ROW, 1).NumberFormat = "@"
    ws.Cells(HDR_ROW, 1).Value = m.Corner
    ReDim lab(1 To 1, 1 To m.nC)
    For j = 1 To m.nC
        lab(1, j) = m.ColLab(j)
    Next j
    Set rng = ws.Range(ws.Cells(HDR_ROW, 2), ws.Cells(HDR_ROW, m.nC + 1))
    rng.NumberFormat = "@"
    rng.Value = lab
    ReDim lab(1 To m.nR, 1 To 1)
    For i = 1 To m.nR
        lab(i, 1) = m.RowLab(i)
    Next i
    Set rng = ws.Range(ws.Cells(HDR_ROW + 1, 1), ws.Cells(HDR_ROW + m.nR, 1))
    rng.NumberFormat = "@"
    rng.Value = lab
    ReDim arr(1 To m.nR, 1 To m.nC)
    For i = 1 To m.nR
        For j = 1 To m.nC
            arr(i, j) = m.V(i, j)
        Next j
    Next i
    Set rng = ws.Range(ws.Cells(HDR_ROW + 1, 2), ws.Cells(HDR_ROW + m.nR, m.nC + 1))
    rng.Value = arr
    If m.Kind = "normalized" Then
        rng.NumberFormat = "0.00"
        ColorRange rng, 1#
    Else
        ColorRange rng, HighThreshold()
    End If
    StyleSheet ws, m.nC + 1
    AddTree ws.Name, m.Kind, title, m.ParentName, m.Note, m.nR & " x " & m.nC
    Set WriteMatrix = ws
End Function


'========================================================== search matrices
Public Function SearchMatrix(ByVal rowItems As Variant, ByVal nR As Long, ByVal colItems As Variant, ByVal nC As Long, _
                             ByVal ctx As String, ByVal tag As String, _
                             ByVal rowTitle As String, ByVal colTitle As String) As AlmaMatrix
    Dim m As AlmaMatrix, i As Long, j As Long, k As Long
    Dim rl() As String, cl() As String, vv() As Double
    gStep = "SearchMatrix"
    m.Kind = "search"
    m.Context = ctx
    m.ColTag = tag
    m.Corner = rowTitle & " \ " & colTitle
    m.nR = nR
    m.nC = nC
    ReDim rl(1 To nR)
    ReDim cl(1 To nC)
    ReDim vv(1 To nR, 1 To nC)
    For i = 1 To nR
        rl(i) = CStr(rowItems(i))
    Next i
    For j = 1 To nC
        cl(j) = CStr(colItems(j))
    Next j
    k = 0
    For i = 1 To nR
        For j = 1 To nC
            vv(i, j) = PubMedCount(AndQuery(rl(i), Tagged(cl(j), tag), ctx))
            k = k + 1
            Progress "search matrix", k, nR * nC
        Next j
    Next i
    m.RowLab = rl
    m.ColLab = cl
    m.V = vv
    SearchMatrix = m
End Function


'====================================================== comparison matrices
' Copy the base search matrix into a comparison (through local arrays).
Private Sub CopyBase(c As AlmaComparison, m As AlmaMatrix)
    Dim rl() As String, cl() As String, vv() As Double
    rl = m.RowLab
    cl = m.ColLab
    vv = m.V
    c.BaseTitle = m.Title
    c.Context = m.Context
    c.ColTag = m.ColTag
    c.nR = m.nR
    c.nC = m.nC
    c.RowLab = rl
    c.ColLab = cl
    c.LR = vv
End Sub

' Add a new string (var) to every hypothesis of a search matrix.
Public Function BuildComparison(m As AlmaMatrix, ByVal newVar As String) As AlmaComparison
    Dim c As AlmaComparison, i As Long, j As Long, k As Long, total As Long
    Dim rv() As Double, cv() As Double, tr() As Double
    gStep = "BuildComparison"
    CopyBase c, m
    c.VarName = newVar
    ReDim rv(1 To m.nR)
    ReDim cv(1 To m.nC)
    ReDim tr(1 To m.nR, 1 To m.nC)
    total = 1 + m.nR + m.nC + m.nR * m.nC
    c.VarTotal = PubMedCount(AndQuery(newVar, m.Context))
    k = 1
    For i = 1 To m.nR
        rv(i) = PubMedCount(AndQuery(m.RowLab(i), newVar, m.Context))
        k = k + 1
        Progress "comparison '" & newVar & "'", k, total
    Next i
    For j = 1 To m.nC
        cv(j) = PubMedCount(AndQuery(Tagged(m.ColLab(j), m.ColTag), newVar, m.Context))
        k = k + 1
        Progress "comparison '" & newVar & "'", k, total
    Next j
    For i = 1 To m.nR
        For j = 1 To m.nC
            tr(i, j) = PubMedCount(AndQuery(m.RowLab(i), Tagged(m.ColLab(j), m.ColTag), newVar, m.Context))
            k = k + 1
            Progress "comparison '" & newVar & "'", k, total
        Next j
    Next i
    c.RowVar = rv
    c.ColVar = cv
    c.Tri = tr
    BuildComparison = c
End Function


' Layout of the paper (Fig 3b/3d, 4b, 5b): first column = row + var, header
' row = column + var, then for every column the pair [row+col, row+col+var].
Public Function WriteComparison(c As AlmaComparison, ByVal title As String) As Worksheet
    Dim ws As Worksheet, i As Long, j As Long, arr() As Variant, nCols As Long, nR As Long, nC As Long
    gStep = "WriteComparison"
    nR = c.nR
    nC = c.nC
    nCols = 2 + 2 * nC
    Set ws = NewSheet(title)
    WriteHeader ws, "comparison", title, c.Context, c.ColTag, c.BaseTitle, c.VarName
    ReDim arr(1 To nR + 2, 1 To nCols)
    arr(1, 1) = "var=" & c.VarName
    arr(1, 2) = c.VarName
    arr(2, 1) = c.VarName
    arr(2, 2) = c.VarTotal
    For j = 1 To nC
        arr(1, 1 + 2 * j) = c.ColLab(j)
        arr(1, 2 + 2 * j) = "& Var"
        arr(2, 1 + 2 * j) = c.ColVar(j)
        arr(2, 2 + 2 * j) = Empty
    Next j
    For i = 1 To nR
        arr(i + 2, 1) = c.RowLab(i)
        arr(i + 2, 2) = c.RowVar(i)
        For j = 1 To nC
            arr(i + 2, 1 + 2 * j) = c.LR(i, j)
            arr(i + 2, 2 + 2 * j) = c.Tri(i, j)
        Next j
    Next i
    ws.Range(ws.Cells(HDR_ROW, 1), ws.Cells(HDR_ROW + nR + 1, 1)).NumberFormat = "@"
    ws.Range(ws.Cells(HDR_ROW, 1), ws.Cells(HDR_ROW, nCols)).NumberFormat = "@"
    ws.Range(ws.Cells(HDR_ROW, 1), ws.Cells(HDR_ROW + nR + 1, nCols)).Value = arr
    ColorRange ws.Range(ws.Cells(HDR_ROW + 1, 2), ws.Cells(HDR_ROW + nR + 1, nCols)), HighThreshold()
    StyleSheet ws, nCols
    AddTree ws.Name, "comparison", title, c.BaseTitle, c.VarName, nR & " x " & nC
    Set WriteComparison = ws
End Function


Public Function ReadComparison(ByVal ws As Worksheet) As AlmaComparison
    Dim c As AlmaComparison, i As Long, j As Long, nR As Long, nC As Long, data As Variant
    Dim rl() As String, cl() As String, vv() As Double, rv() As Double, cv() As Double, tr() As Double
    gStep = "ReadComparison"
    If SheetKind(ws) <> "comparison" Then
        Err.Raise ALMA_ERR + 602, "ALMA", "Sheet '" & ws.Name & "' is not an ALMA comparison matrix."
    End If
    c.Title = ws.Name
    c.VarName = CStr(ws.Cells(6, 2).Value)
    c.BaseTitle = CStr(ws.Cells(5, 2).Value)
    c.Context = CStr(ws.Cells(3, 2).Value)
    c.ColTag = CStr(ws.Cells(4, 2).Value)
    nC = 0
    Do While Len(CStr(ws.Cells(HDR_ROW, 3 + 2 * nC).Value)) > 0
        nC = nC + 1
    Loop
    nR = 0
    Do While Len(CStr(ws.Cells(HDR_ROW + 2 + nR, 1).Value)) > 0
        nR = nR + 1
    Loop
    If nR = 0 Or nC = 0 Then
        Err.Raise ALMA_ERR + 603, "ALMA", "No comparison matrix found on sheet '" & ws.Name & "'."
    End If
    ReDim rl(1 To nR)
    ReDim cl(1 To nC)
    ReDim vv(1 To nR, 1 To nC)
    ReDim rv(1 To nR)
    ReDim cv(1 To nC)
    ReDim tr(1 To nR, 1 To nC)
    data = ws.Range(ws.Cells(HDR_ROW, 1), ws.Cells(HDR_ROW + 1 + nR, 2 + 2 * nC)).Value
    c.VarTotal = NumOr0(data(2, 2))
    For j = 1 To nC
        cl(j) = CStr(data(1, 1 + 2 * j))
        cv(j) = NumOr0(data(2, 1 + 2 * j))
    Next j
    For i = 1 To nR
        rl(i) = CStr(data(i + 2, 1))
        rv(i) = NumOr0(data(i + 2, 2))
        For j = 1 To nC
            vv(i, j) = NumOr0(data(i + 2, 1 + 2 * j))
            tr(i, j) = NumOr0(data(i + 2, 2 + 2 * j))
        Next j
    Next i
    c.nR = nR
    c.nC = nC
    c.RowLab = rl
    c.ColLab = cl
    c.LR = vv
    c.RowVar = rv
    c.ColVar = cv
    c.Tri = tr
    ReadComparison = c
End Function


'============================================================== NRF scoring
Public Function ReasonScore(ByVal n As Double, ByVal high As Double, ByVal med As Double) As Long
    If n >= high Then
        ReasonScore = 2
    ElseIf n >= med Then
        ReasonScore = 1
    Else
        ReasonScore = 0
    End If
End Function


Public Function NoveltyScore(ByVal n As Double, ByVal high As Double, ByVal med As Double) As Long
    If n >= high Then
        NoveltyScore = 0
    ElseIf n >= med Then
        NoveltyScore = 1
    Else
        NoveltyScore = 2
    End If
End Function


' NRF records are accumulated one comparison matrix at a time (NrfReset, NrfAdd)
' and NrfTable returns the ranked table (header row + one row per hypothesis).
'   LR  = row + col        THR = row + var        TVR = col + var
'   HR  = row + var with the other columns (sum of triples)
'   VR  = col + var with the other rows (sum of triples)
'   N   = row + col + var  (inverse score)
' Ranking: N, LR, TVR, THR, VR, HR (high first); ties: fewer publications of the
' hypothesis first, then the stronger weakest pair, then input order.
Public Sub NrfReset()
    mNrfN = 0
    Erase mNrfRec
    Erase mNrfKeys
End Sub

Public Sub NrfAdd(c As AlmaComparison, ByVal high As Double, ByVal med As Double, _
                  ByVal novelMax As Double, ByVal pairMin As Double)
    Dim i As Long, j As Long, p As Long, t As Double, lr As Double, thr As Double, tvr As Double
    Dim hr As Double, vr As Double, sN As Long, sLR As Long, sTVR As Long, sTHR As Long, sVR As Long, sHR As Long
    Dim minPair As Double, rowTot() As Double, colTot() As Double
    gStep = "NrfAdd"
    ReDim rowTot(1 To c.nR)
    ReDim colTot(1 To c.nC)
    For i = 1 To c.nR
        For j = 1 To c.nC
            rowTot(i) = rowTot(i) + c.Tri(i, j)
            colTot(j) = colTot(j) + c.Tri(i, j)
        Next j
    Next i
    If mNrfN = 0 Then
        ReDim mNrfRec(1 To NRF_COLS, 1 To c.nR * c.nC)
        ReDim mNrfKeys(1 To 9, 1 To c.nR * c.nC)
    Else
        ReDim Preserve mNrfRec(1 To NRF_COLS, 1 To mNrfN + c.nR * c.nC)
        ReDim Preserve mNrfKeys(1 To 9, 1 To mNrfN + c.nR * c.nC)
    End If
    p = mNrfN
    For i = 1 To c.nR
        For j = 1 To c.nC
            p = p + 1
            t = c.Tri(i, j)
            lr = c.LR(i, j)
            thr = c.RowVar(i)
            tvr = c.ColVar(j)
            hr = rowTot(i) - t
            vr = colTot(j) - t
            sN = NoveltyScore(t, high, med)
            sLR = ReasonScore(lr, high, med)
            sTVR = ReasonScore(tvr, high, med)
            sTHR = ReasonScore(thr, high, med)
            sVR = ReasonScore(vr, high, med)
            sHR = ReasonScore(hr, high, med)
            mNrfRec(2, p) = c.RowLab(i) & ALMA_SEP & c.ColLab(j) & ALMA_SEP & c.VarName
            mNrfRec(3, p) = c.RowLab(i)
            mNrfRec(4, p) = c.ColLab(j)
            mNrfRec(5, p) = c.VarName
            mNrfRec(6, p) = t
            mNrfRec(7, p) = lr
            mNrfRec(8, p) = thr
            mNrfRec(9, p) = tvr
            mNrfRec(10, p) = hr
            mNrfRec(11, p) = vr
            mNrfRec(12, p) = sN
            mNrfRec(13, p) = sLR
            mNrfRec(14, p) = sTVR
            mNrfRec(15, p) = sTHR
            mNrfRec(16, p) = sVR
            mNrfRec(17, p) = sHR
            minPair = lr
            If thr < minPair Then minPair = thr
            If tvr < minPair Then minPair = tvr
            mNrfRec(18, p) = (t <= novelMax And minPair >= pairMin)
            mNrfKeys(1, p) = -sN
            mNrfKeys(2, p) = -sLR
            mNrfKeys(3, p) = -sTVR
            mNrfKeys(4, p) = -sTHR
            mNrfKeys(5, p) = -sVR
            mNrfKeys(6, p) = -sHR
            mNrfKeys(7, p) = t
            mNrfKeys(8, p) = -minPair
            mNrfKeys(9, p) = p
        Next j
    Next i
    mNrfN = p
End Sub

Public Function NrfTable() As Variant
    Dim n As Long, i As Long, q As Long, keys() As Double, ord As Variant, tbl() As Variant, hdr As Variant
    gStep = "NrfTable"
    n = mNrfN
    hdr = Array("rank", "hypothesis", "row", "col", "var", "NOP", "NOP_LR", "NOP_THR", "NOP_TVR", _
                "NOP_HR", "NOP_VR", "N", "LR", "TVR", "THR", "VR", "HR", "novel_reasonable")
    ReDim tbl(1 To n + 1, 1 To NRF_COLS)
    For q = 1 To NRF_COLS
        tbl(1, q) = hdr(q - 1)
    Next q
    If n = 0 Then
        NrfTable = tbl
        Exit Function
    End If
    ReDim keys(1 To n, 1 To 9)
    For i = 1 To n
        For q = 1 To 9
            keys(i, q) = mNrfKeys(q, i)
        Next q
    Next i
    ord = SortIndexLex(keys, n, 9)
    For i = 1 To n
        tbl(i + 1, 1) = i
        For q = 2 To NRF_COLS
            tbl(i + 1, q) = mNrfRec(q, ord(i))
        Next q
    Next i
    NrfTable = tbl
End Function


'=========================================================== sorting helpers
' Stable merge sort of 1..n by the rows of keys(1..n, 1..nk), ascending, lexicographic.
Public Function SortIndexLex(keys() As Double, ByVal n As Long, ByVal nk As Long) As Variant
    Dim idx() As Long, tmp() As Long, i As Long, runLen As Long, lo As Long, md As Long, hi As Long
    ReDim idx(1 To n)
    ReDim tmp(1 To n)
    For i = 1 To n
        idx(i) = i
    Next i
    runLen = 1
    Do While runLen < n
        lo = 1
        Do While lo <= n
            md = lo + runLen - 1
            If md > n Then md = n
            hi = lo + 2 * runLen - 1
            If hi > n Then hi = n
            MergeRuns idx, tmp, lo, md, hi, keys, nk
            lo = lo + 2 * runLen
        Loop
        For i = 1 To n
            idx(i) = tmp(i)
        Next i
        runLen = runLen * 2
    Loop
    SortIndexLex = idx
End Function


Private Sub MergeRuns(idx() As Long, tmp() As Long, ByVal lo As Long, ByVal md As Long, ByVal hi As Long, _
                      keys() As Double, ByVal nk As Long)
    Dim a As Long, b As Long, k As Long
    a = lo
    b = md + 1
    k = lo
    Do While a <= md And b <= hi
        If LexLE(keys, idx(a), idx(b), nk) Then
            tmp(k) = idx(a)
            a = a + 1
        Else
            tmp(k) = idx(b)
            b = b + 1
        End If
        k = k + 1
    Loop
    Do While a <= md
        tmp(k) = idx(a)
        a = a + 1
        k = k + 1
    Loop
    Do While b <= hi
        tmp(k) = idx(b)
        b = b + 1
        k = k + 1
    Loop
End Sub


Private Function LexLE(keys() As Double, ByVal x As Long, ByVal y As Long, ByVal nk As Long) As Boolean
    Dim q As Long
    LexLE = True
    For q = 1 To nk
        If keys(x, q) < keys(y, q) Then
            LexLE = True
            Exit Function
        ElseIf keys(x, q) > keys(y, q) Then
            LexLE = False
            Exit Function
        End If
    Next q
End Function


'========================================================= matrix operations
' Copy of the scalar fields of a matrix (arrays are always built fresh: some VBA
' hosts share arrays between copies of a user-defined type).
Private Sub CopyFields(r As AlmaMatrix, m As AlmaMatrix)
    r.Kind = m.Kind
    r.Title = m.Title
    r.Context = m.Context
    r.ColTag = m.ColTag
    r.ParentName = m.ParentName
    r.Note = m.Note
    r.Corner = m.Corner
    r.nR = m.nR
    r.nC = m.nC
End Sub

Private Function Reorder(m As AlmaMatrix, ByVal ro As Variant, ByVal co As Variant) As AlmaMatrix
    Dim r As AlmaMatrix, i As Long, j As Long, rl() As String, cl() As String, vv() As Double
    ReDim rl(1 To m.nR)
    ReDim cl(1 To m.nC)
    ReDim vv(1 To m.nR, 1 To m.nC)
    For i = 1 To m.nR
        rl(i) = m.RowLab(ro(i))
        For j = 1 To m.nC
            vv(i, j) = m.V(ro(i), co(j))
        Next j
    Next i
    For j = 1 To m.nC
        cl(j) = m.ColLab(co(j))
    Next j
    CopyFields r, m
    r.RowLab = rl
    r.ColLab = cl
    r.V = vv
    r.ParentName = m.Title
    Reorder = r
End Function


Private Function Identity(ByVal n As Long) As Variant
    Dim a() As Long, i As Long
    ReDim a(1 To n)
    For i = 1 To n
        a(i) = i
    Next i
    Identity = a
End Function


Private Function TransposeMatrix(m As AlmaMatrix) As AlmaMatrix
    Dim r As AlmaMatrix, i As Long, j As Long, rl() As String, cl() As String, vv() As Double
    CopyFields r, m
    r.nR = m.nC
    r.nC = m.nR
    ReDim rl(1 To r.nR)
    ReDim cl(1 To r.nC)
    ReDim vv(1 To r.nR, 1 To r.nC)
    For i = 1 To r.nR
        rl(i) = m.ColLab(i)
    Next i
    For j = 1 To r.nC
        cl(j) = m.RowLab(j)
    Next j
    For i = 1 To r.nR
        For j = 1 To r.nC
            vv(i, j) = m.V(j, i)
        Next j
    Next i
    r.RowLab = rl
    r.ColLab = cl
    r.V = vv
    TransposeMatrix = r
End Function


Private Function AggRow(m As AlmaMatrix, ByVal i As Long, ByVal by As String) As Double
    Dim j As Long, s As Double, mx As Double
    s = 0
    mx = m.V(i, 1)
    For j = 1 To m.nC
        s = s + m.V(i, j)
        If m.V(i, j) > mx Then mx = m.V(i, j)
    Next j
    Select Case by
        Case "max"
            AggRow = mx
        Case "mean"
            AggRow = s / m.nC
        Case Else
            AggRow = s
    End Select
End Function


Private Function AggCol(m As AlmaMatrix, ByVal j As Long, ByVal by As String) As Double
    Dim i As Long, s As Double, mx As Double
    s = 0
    mx = m.V(1, j)
    For i = 1 To m.nR
        s = s + m.V(i, j)
        If m.V(i, j) > mx Then mx = m.V(i, j)
    Next i
    Select Case by
        Case "max"
            AggCol = mx
        Case "mean"
            AggCol = s / m.nR
        Case Else
            AggCol = s
    End Select
End Function


' Sort rows and/or columns by their sum / max / mean, highest first (stable).
Public Function SortMatrix(m As AlmaMatrix, ByVal by As String, ByVal axis As String) As AlmaMatrix
    Dim keys() As Double, ro As Variant, co As Variant, i As Long, j As Long
    gStep = "SortMatrix"
    ro = Identity(m.nR)
    co = Identity(m.nC)
    If axis = "both" Or axis = "rows" Then
        ReDim keys(1 To m.nR, 1 To 2)
        For i = 1 To m.nR
            keys(i, 1) = -AggRow(m, i, by)
            keys(i, 2) = i
        Next i
        ro = SortIndexLex(keys, m.nR, 2)
    End If
    If axis = "both" Or axis = "cols" Then
        ReDim keys(1 To m.nC, 1 To 2)
        For j = 1 To m.nC
            keys(j, 1) = -AggCol(m, j, by)
            keys(j, 2) = j
        Next j
        co = SortIndexLex(keys, m.nC, 2)
    End If
    SortMatrix = Reorder(m, ro, co)
End Function


' Sort the rows by one column (highest first, stable).
Public Function SortByColumn(m As AlmaMatrix, ByVal colIndex As Long) As AlmaMatrix
    Dim keys() As Double, ro As Variant, co As Variant, i As Long
    ReDim keys(1 To m.nR, 1 To 2)
    For i = 1 To m.nR
        keys(i, 1) = -m.V(i, colIndex)
        keys(i, 2) = i
    Next i
    ro = SortIndexLex(keys, m.nR, 2)
    co = Identity(m.nC)
    SortByColumn = Reorder(m, ro, co)
End Function


' Divide every row (horizontal) or column (vertical) by its maximum.
Public Function NormalizeMatrix(m As AlmaMatrix, ByVal axis As String) As AlmaMatrix
    Dim r As AlmaMatrix, i As Long, j As Long, mx As Double, rl() As String, cl() As String, vv() As Double
    gStep = "NormalizeMatrix"
    ReDim vv(1 To m.nR, 1 To m.nC)
    rl = m.RowLab
    cl = m.ColLab
    CopyFields r, m
    r.Kind = "normalized"
    r.ParentName = m.Title
    If axis = "rows" Then
        r.Note = "normalized by row maximum"
        For i = 1 To m.nR
            mx = 0
            For j = 1 To m.nC
                If m.V(i, j) > mx Then mx = m.V(i, j)
            Next j
            For j = 1 To m.nC
                If mx > 0 Then vv(i, j) = m.V(i, j) / mx
            Next j
        Next i
    Else
        r.Note = "normalized by col maximum"
        For j = 1 To m.nC
            mx = 0
            For i = 1 To m.nR
                If m.V(i, j) > mx Then mx = m.V(i, j)
            Next i
            For i = 1 To m.nR
                If mx > 0 Then vv(i, j) = m.V(i, j) / mx
            Next i
        Next j
    End If
    r.RowLab = rl
    r.ColLab = cl
    r.V = vv
    NormalizeMatrix = r
End Function


' Cluster sort of a normalized matrix: every row joins the cluster of a column
' where nNOP >= thr; the column capturing most unassigned rows goes first.
Public Function ClusterSortMatrix(m As AlmaMatrix, ByVal thr As Double, ByVal axis As String) As AlmaMatrix
    Dim a As AlmaMatrix, r As AlmaMatrix
    Dim unassigned() As Boolean, remaining() As Boolean, orderR() As Long, orderC() As Long
    Dim nOR As Long, nOC As Long, i As Long, j As Long, best As Long, bestCount As Long, cnt As Long
    Dim keys() As Double, members() As Long, nm As Long, srt As Variant, k As Long
    gStep = "ClusterSortMatrix"
    If axis = "cols" Then
        a = TransposeMatrix(m)
    Else
        a = m
    End If
    ReDim unassigned(1 To a.nR)
    ReDim remaining(1 To a.nC)
    ReDim orderR(1 To a.nR)
    ReDim orderC(1 To a.nC)
    For i = 1 To a.nR
        unassigned(i) = True
    Next i
    For j = 1 To a.nC
        remaining(j) = True
    Next j
    nOR = 0
    nOC = 0
    Do
        best = 0
        bestCount = 0
        For j = 1 To a.nC
            If remaining(j) Then
                cnt = 0
                For i = 1 To a.nR
                    If unassigned(i) Then
                        If a.V(i, j) >= thr Then cnt = cnt + 1
                    End If
                Next i
                If cnt > bestCount Then
                    best = j
                    bestCount = cnt
                End If
            End If
        Next j
        If best = 0 Then Exit Do
        ReDim keys(1 To bestCount, 1 To 2)
        ReDim members(1 To bestCount)
        nm = 0
        For i = 1 To a.nR
            If unassigned(i) Then
                If a.V(i, best) >= thr Then
                    nm = nm + 1
                    members(nm) = i
                    keys(nm, 1) = -a.V(i, best)
                    keys(nm, 2) = i
                End If
            End If
        Next i
        srt = SortIndexLex(keys, nm, 2)
        For k = 1 To nm
            nOR = nOR + 1
            orderR(nOR) = members(srt(k))
            unassigned(members(srt(k))) = False
        Next k
        nOC = nOC + 1
        orderC(nOC) = best
        remaining(best) = False
    Loop
    For j = 1 To a.nC
        If remaining(j) Then
            nOC = nOC + 1
            orderC(nOC) = j
        End If
    Next j
    For i = 1 To a.nR
        If unassigned(i) Then
            nOR = nOR + 1
            orderR(nOR) = i
        End If
    Next i
    r = Reorder(a, orderR, orderC)
    If axis = "cols" Then r = TransposeMatrix(r)
    r.ParentName = m.Title
    r.Note = "cluster sorted, threshold " & Format$(thr, "0.00")
    ClusterSortMatrix = r
End Function


'================================================================== colours
' Red (0) -> yellow (hi/2) -> green (>= hi), like the paper's colour bars.
Public Function ScaleColor(ByVal v As Double, ByVal hi As Double) As Long
    Dim t As Double
    If hi <= 0 Then hi = 1
    t = v / hi
    If t < 0 Then t = 0
    If t > 1 Then t = 1
    If t < 0.5 Then
        ScaleColor = MixColor(248, 105, 107, 255, 235, 132, t / 0.5)
    Else
        ScaleColor = MixColor(255, 235, 132, 99, 190, 123, (t - 0.5) / 0.5)
    End If
End Function


Private Function MixColor(ByVal r1 As Double, ByVal g1 As Double, ByVal b1 As Double, _
                          ByVal r2 As Double, ByVal g2 As Double, ByVal b2 As Double, ByVal f As Double) As Long
    MixColor = RGB(CInt(r1 + (r2 - r1) * f), CInt(g1 + (g2 - g1) * f), CInt(b1 + (b2 - b1) * f))
End Function


Public Sub ColorRange(ByVal rng As Range, ByVal hi As Double)
    Dim i As Long, j As Long, v As Variant, data As Variant, nR As Long, nC As Long
    gStep = "ColorRange"
    nR = rng.Rows.Count
    nC = rng.Columns.Count
    data = rng.Value
    For i = 1 To nR
        For j = 1 To nC
            If nR = 1 And nC = 1 Then
                v = data
            Else
                v = data(i, j)
            End If
            If Not IsEmpty(v) Then
                If VarType(v) <> vbString And VarType(v) <> vbBoolean And IsNumeric(v) Then
                    rng.Cells(i, j).Interior.Color = ScaleColor(CDbl(v), hi)
                End If
            End If
        Next j
    Next i
End Sub


'==================================================================== tables
Private Function IsTextColumn(ByVal h As String) As Boolean
    Select Case h
        Case "hypothesis", "row", "col", "var", "first", "second"
            IsTextColumn = True
        Case Else
            IsTextColumn = False
    End Select
End Function


' tbl: 2-D array, first row = column headers.
Public Function WriteTable(tbl As Variant, ByVal title As String, ByVal kind As String, _
                           ByVal parentName As String, ByVal note As String) As Worksheet
    Dim ws As Worksheet, nR As Long, nC As Long, j As Long, h As String
    gStep = "WriteTable"
    nR = UBound(tbl, 1) - LBound(tbl, 1) + 1
    nC = UBound(tbl, 2) - LBound(tbl, 2) + 1
    Set ws = NewSheet(title)
    WriteHeader ws, kind, title, "", "", parentName, note
    For j = 1 To nC
        If IsTextColumn(CStr(tbl(1, j))) Then
            ws.Range(ws.Cells(HDR_ROW, j), ws.Cells(HDR_ROW + nR - 1, j)).NumberFormat = "@"
        End If
    Next j
    ws.Range(ws.Cells(HDR_ROW, 1), ws.Cells(HDR_ROW + nR - 1, nC)).Value = tbl
    If nR > 1 Then
        For j = 1 To nC
            h = CStr(tbl(1, j))
            Select Case h
                Case "N", "LR", "TVR", "THR", "VR", "HR"
                    ColorRange ws.Range(ws.Cells(HDR_ROW + 1, j), ws.Cells(HDR_ROW + nR - 1, j)), 2#
                Case "NOP_LR", "NOP_THR", "NOP_TVR"
                    ColorRange ws.Range(ws.Cells(HDR_ROW + 1, j), ws.Cells(HDR_ROW + nR - 1, j)), HighThreshold()
                Case "NOP"
                    If kind = "text" Then
                        ColorRange ws.Range(ws.Cells(HDR_ROW + 1, j), ws.Cells(HDR_ROW + nR - 1, j)), HighThreshold()
                    End If
            End Select
        Next j
    End If
    On Error Resume Next
    ws.Rows(HDR_ROW).Font.Bold = True
    ws.Columns(1).ColumnWidth = 8
    For j = 2 To nC
        If IsTextColumn(CStr(tbl(1, j))) Then
            ws.Columns(j).ColumnWidth = 36
        Else
            ws.Columns(j).ColumnWidth = 10
        End If
    Next j
    On Error GoTo 0
    AddTree ws.Name, kind, title, parentName, note, (nR - 1) & " rows"
    Set WriteTable = ws
End Function


'============================================================= tree manager
Public Sub AddTree(ByVal sheetName As String, ByVal kind As String, ByVal title As String, _
                   ByVal parentName As String, ByVal added As String, ByVal size As String)
    Dim ws As Worksheet, r As Long
    gStep = "AddTree"
    Set ws = ThisWorkbook.Worksheets(TREE_SHEET)
    r = TREE_FIRST
    Do While Len(CStr(ws.Cells(r, 1).Value)) > 0
        r = r + 1
    Loop
    ws.Cells(r, 1).NumberFormat = "@"
    ws.Cells(r, 1).Value = sheetName
    On Error Resume Next
    ws.Hyperlinks.Add Anchor:=ws.Cells(r, 1), Address:="", SubAddress:="'" & sheetName & "'!A1", _
                      TextToDisplay:=sheetName
    On Error GoTo 0
    ws.Cells(r, 2).Value = kind
    ws.Cells(r, 3).NumberFormat = "@"
    ws.Cells(r, 3).Value = title
    ws.Cells(r, 4).NumberFormat = "@"
    ws.Cells(r, 4).Value = parentName
    ws.Cells(r, 5).NumberFormat = "@"
    ws.Cells(r, 5).Value = added
    ws.Cells(r, 6).NumberFormat = "@"
    ws.Cells(r, 6).Value = size
    ws.Cells(r, 7).Value = Format$(Now, "yyyy-mm-dd hh:nn")
End Sub


Public Sub Progress(ByVal what As String, ByVal k As Long, ByVal n As Long)
    If (k Mod 5 = 0) Or (k = n) Then
        On Error Resume Next
        Application.StatusBar = "ALMA: " & what & " " & k & "/" & n & _
                                "   (PubMed: " & gFetched & " fetched, " & gFromCache & " from cache)"
        On Error GoTo 0
        DoEvents
    End If
End Sub

'================================================================== actions

'------------------------------------------------------------ target sheets
Private Function KindMatches(ByVal kind As String, ByVal want As String) As Boolean
    Select Case want
        Case "matrix"
            KindMatches = (kind = "search" Or kind = "normalized")
        Case "nrfsource"
            KindMatches = (kind = "search" Or kind = "comparison")
        Case "any"
            KindMatches = (Len(kind) > 0)
        Case Else
            KindMatches = (kind = want)
    End Select
End Function


' The named sheet, else the active sheet if suitable, else the newest suitable sheet in the Tree.
Private Function ResolveTarget(ByVal target As String, ByVal want As String) As Worksheet
    Dim ws As Worksheet, t As Worksheet, r As Long, last As Long, nm As String
    If Len(target) > 0 Then
        Set ws = ThisWorkbook.Worksheets(target)
        If Not KindMatches(SheetKind(ws), want) Then
            Err.Raise ALMA_ERR + 701, "ALMA", "Sheet '" & target & "' is not a suitable ALMA sheet (" & want & ")."
        End If
        Set ResolveTarget = ws
        Exit Function
    End If
    Set ws = Nothing
    On Error Resume Next
    Set ws = ActiveSheet
    On Error GoTo 0
    If Not ws Is Nothing Then
        If KindMatches(SheetKind(ws), want) Then
            Set ResolveTarget = ws
            Exit Function
        End If
    End If
    Set t = ThisWorkbook.Worksheets(TREE_SHEET)
    last = TREE_FIRST
    Do While Len(CStr(t.Cells(last, 1).Value)) > 0
        last = last + 1
    Loop
    For r = last - 1 To TREE_FIRST Step -1
        nm = CStr(t.Cells(r, 1).Value)
        If SheetExists(nm) Then
            If KindMatches(SheetKind(ThisWorkbook.Worksheets(nm)), want) Then
                Set ResolveTarget = ThisWorkbook.Worksheets(nm)
                Exit Function
            End If
        End If
    Next r
    Err.Raise ALMA_ERR + 702, "ALMA", "No suitable ALMA sheet found (" & want & ") - create or select one first."
End Function


'------------------------------------------------------------------- actions
Private Function NewSearchFromLists() As Worksheet
    Dim rowItems As Variant, colItems As Variant, nR As Long, nC As Long, m As AlmaMatrix
    rowItems = ListItems(1, nR)
    colItems = ListItems(2, nC)
    If nR = 0 Or nC = 0 Then
        Err.Raise ALMA_ERR + 703, "ALMA", "Fill List 1 (column A) and List 2 (column B) on the Lists sheet."
    End If
    m = SearchMatrix(rowItems, nR, colItems, nC, JoinList(4), "", ListSetting("Row title", "rowItems"), _
                     ListSetting("Column title", "columns"))
    Set NewSearchFromLists = WriteMatrix(m, ListSetting("Matrix name", "SM"))
End Function


Public Function DoSearch() As String
    Dim ws As Worksheet
    Set ws = NewSearchFromLists()
    ws.Activate
    DoSearch = "search matrix '" & ws.Name & "' created"
End Function


' Search matrix -> sorted copy -> comparison matrices for every string in column C -> NRF ranking.
Public Function DoPipeline() As String
    Dim ws As Worksheet, m As AlmaMatrix, sm As AlmaMatrix, res As String
    gStep = "DoPipeline"
    Set ws = NewSearchFromLists()
    m = ReadMatrix(ws)
    sm = SortMatrix(m, "sum", "both")
    WriteMatrix sm, ws.Name & "_sorted"
    res = ComparisonsAndNrf(ws, True, "")
    DoPipeline = "search matrix '" & ws.Name & "'; " & res
End Function


Private Function ComparisonsAndNrf(ByVal ws As Worksheet, ByVal allVars As Boolean, ByVal oneVar As String) As String
    Dim m As AlmaMatrix, vars As Variant, nV As Long, c As AlmaComparison, k As Long
    Dim wsC As Worksheet, wsN As Worksheet, tbl As Variant, nNovel As Long
    gStep = "ComparisonsAndNrf"
    m = ReadMatrix(ws)
    If m.Kind <> "search" Then
        Err.Raise ALMA_ERR + 704, "ALMA", "Comparison matrices need a search matrix (sheet '" & ws.Name & "' is " & m.Kind & ")."
    End If
    If allVars Then
        vars = ListItems(3, nV)
    Else
        nV = 1
        vars = OneItem(oneVar)
    End If
    If nV = 0 Then
        Err.Raise ALMA_ERR + 705, "ALMA", "Fill the comparison strings (column C) on the Lists sheet."
    End If
    NrfReset
    For k = 1 To nV
        c = BuildComparison(m, vars(k))
        Set wsC = WriteComparison(c, ws.Name & "_cmp_" & vars(k))
        NrfAdd c, HighThreshold(), MedThreshold(), SettingNum("novelmax", 1), SettingNum("pairmin", 10)
    Next k
    tbl = NrfTable()
    Set wsN = WriteTable(tbl, ws.Name & "_NRF", "nrf", ws.Name, "high>=" & HighThreshold() & " med>=" & MedThreshold())
    nNovel = WriteNovel(tbl, ws.Name & "_novel", wsN.Name)
    wsN.Activate
    ComparisonsAndNrf = nV & " comparison matrices, NRF ranking '" & wsN.Name & "', " & _
                        nNovel & " novel & reasonable hypotheses"
End Function


' Rows of an NRF table that pass the paper's filter (last column TRUE).
Private Function WriteNovel(tbl As Variant, ByVal title As String, ByVal parentName As String) As Long
    Dim n As Long, i As Long, q As Long, k As Long, out() As Variant
    gStep = "WriteNovel"
    n = 0
    For i = 2 To UBound(tbl, 1)
        If tbl(i, NRF_COLS) = True Then n = n + 1
    Next i
    ReDim out(1 To n + 1, 1 To NRF_COLS)
    For q = 1 To NRF_COLS
        out(1, q) = tbl(1, q)
    Next q
    k = 1
    For i = 2 To UBound(tbl, 1)
        If tbl(i, NRF_COLS) = True Then
            k = k + 1
            For q = 1 To NRF_COLS
                out(k, q) = tbl(i, q)
            Next q
        End If
    Next i
    WriteTable out, title, "nrf", parentName, "N<=" & SettingNum("novelmax", 1) & " publications and every pair >=" & _
               SettingNum("pairmin", 10)
    WriteNovel = n
End Function


Public Function DoComparison(ByVal target As String, ByVal allVars As Boolean) As String
    Dim ws As Worksheet, v As String, n As Long, vars As Variant
    Set ws = ResolveTarget(target, "search")
    If allVars Then
        DoComparison = ComparisonsAndNrf(ws, True, "")
    Else
        vars = ListItems(3, n)
        v = ""
        If n > 0 Then v = vars(1)
        If Not Quiet() Then
            v = InputBox("String to add to every hypothesis of '" & ws.Name & "':", "ALMA comparison matrix", v)
        End If
        If Len(Trim$(v)) = 0 Then Err.Raise ALMA_ERR + 706, "ALMA", "No comparison string given."
        DoComparison = ComparisonsAndNrf(ws, False, Trim$(v))
    End If
End Function


Public Function DoNrf(ByVal target As String) As String
    Dim ws As Worksheet, wsC As Worksheet, wsN As Worksheet, c As AlmaComparison, n As Long, i As Long
    Dim tbl As Variant, nNovel As Long, par As String
    gStep = "DoNrf"
    Set ws = ResolveTarget(target, "nrfsource")
    NrfReset
    If SheetKind(ws) = "comparison" Then
        n = 1
        c = ReadComparison(ws)
        NrfAdd c, HighThreshold(), MedThreshold(), SettingNum("novelmax", 1), SettingNum("pairmin", 10)
    Else
        n = 0
        For i = 1 To ThisWorkbook.Worksheets.Count
            Set wsC = ThisWorkbook.Worksheets(i)
            If SheetKind(wsC) = "comparison" Then
                par = CStr(wsC.Cells(5, 2).Value)
                If par = ws.Name Or par = CStr(ws.Cells(2, 2).Value) Then
                    n = n + 1
                    c = ReadComparison(wsC)
                    NrfAdd c, HighThreshold(), MedThreshold(), SettingNum("novelmax", 1), SettingNum("pairmin", 10)
                End If
            End If
        Next i
        If n = 0 Then
            Err.Raise ALMA_ERR + 707, "ALMA", "No comparison matrices of '" & ws.Name & "' found - create them first."
        End If
    End If
    tbl = NrfTable()
    Set wsN = WriteTable(tbl, ws.Name & "_NRF", "nrf", ws.Name, "high>=" & HighThreshold() & " med>=" & MedThreshold())
    nNovel = WriteNovel(tbl, ws.Name & "_novel", wsN.Name)
    wsN.Activate
    DoNrf = "NRF ranking '" & wsN.Name & "' of " & n & " comparison matrices, " & nNovel & " novel & reasonable hypotheses"
End Function


Public Function DoAddString(ByVal target As String) As String
    Dim ws As Worksheet, wsO As Worksheet, m As AlmaMatrix, r As AlmaMatrix, s As String, ctx As String
    Set ws = ResolveTarget(target, "search")
    m = ReadMatrix(ws)
    s = ListSetting("String to add", "")
    If Not Quiet() Then
        s = InputBox("Constant string to add to every query of '" & ws.Name & "':", "ALMA", s)
    End If
    s = Trim$(s)
    If Len(s) = 0 Then Err.Raise ALMA_ERR + 709, "ALMA", "No string given."
    ctx = m.Context
    If Len(ctx) > 0 Then ctx = ctx & ALMA_SEP
    ctx = ctx & s
    r = SearchMatrix(m.RowLab, m.nR, m.ColLab, m.nC, ctx, m.ColTag, "rows", "columns")
    r.Corner = m.Corner
    r.ParentName = ws.Name
    r.Note = "added string: " & s
    Set wsO = WriteMatrix(r, ws.Name & "+" & s)
    wsO.Activate
    DoAddString = "search matrix '" & wsO.Name & "' (+ " & s & ")"
End Function


Public Function DoNormalize(ByVal target As String, ByVal axis As String) As String
    Dim ws As Worksheet, wsO As Worksheet, m As AlmaMatrix, r As AlmaMatrix
    gStep = "DoNormalize"
    Set ws = ResolveTarget(target, "matrix")
    m = ReadMatrix(ws)
    r = NormalizeMatrix(m, axis)
    Set wsO = WriteMatrix(r, ws.Name & "_norm_" & axis)
    wsO.Activate
    DoNormalize = "normalized matrix '" & wsO.Name & "'"
End Function


Public Function DoSort(ByVal target As String, ByVal by As String) As String
    Dim ws As Worksheet, wsO As Worksheet, m As AlmaMatrix, r As AlmaMatrix, title As String
    Set ws = ResolveTarget(target, "matrix")
    m = ReadMatrix(ws)
    r = SortMatrix(m, by, "both")
    title = ws.Name & "_sorted"
    If by <> "sum" Then title = title & "_" & by
    Set wsO = WriteMatrix(r, title)
    wsO.Activate
    DoSort = "sorted matrix '" & wsO.Name & "' (by " & by & ")"
End Function


Public Function DoSortColumn(ByVal target As String) As String
    Dim ws As Worksheet, wsO As Worksheet, m As AlmaMatrix, r As AlmaMatrix, lab As String, j As Long, idx As Long
    Set ws = ResolveTarget(target, "matrix")
    m = ReadMatrix(ws)
    lab = ListSetting("Sort column", m.ColLab(m.nC))
    If Not Quiet() Then lab = InputBox("Sort the rows of '" & ws.Name & "' by which column?", "ALMA", lab)
    idx = 0
    For j = 1 To m.nC
        If LCase$(m.ColLab(j)) = LCase$(Trim$(lab)) Then idx = j
    Next j
    If idx = 0 Then Err.Raise ALMA_ERR + 710, "ALMA", "Column '" & lab & "' not found."
    r = SortByColumn(m, idx)
    Set wsO = WriteMatrix(r, ws.Name & "_by_" & m.ColLab(idx))
    wsO.Activate
    DoSortColumn = "matrix '" & wsO.Name & "' sorted by " & m.ColLab(idx)
End Function


Public Function DoCluster(ByVal target As String) As String
    Dim ws As Worksheet, wsO As Worksheet, m As AlmaMatrix, r As AlmaMatrix, axis As String, thr As Double
    gStep = "DoCluster"
    Set ws = ResolveTarget(target, "matrix")
    m = ReadMatrix(ws)
    thr = SettingNum("cluster", 0.9)
    axis = "rows"
    If m.Kind <> "normalized" Then
        m = NormalizeMatrix(m, "rows")
        Set ws = WriteMatrix(m, ws.Name & "_norm_rows")
        m = ReadMatrix(ws)
    ElseIf InStr(1, m.Note, "col", vbTextCompare) > 0 Then
        axis = "cols"
    End If
    r = ClusterSortMatrix(m, thr, axis)
    Set wsO = WriteMatrix(r, ws.Name & "_clusters")
    wsO.Activate
    DoCluster = "cluster-sorted matrix '" & wsO.Name & "' (threshold " & thr & ")"
End Function


Public Function DoColorize(ByVal target As String) As String
    Dim ws As Worksheet, kind As String, m As AlmaMatrix, c As AlmaComparison, j As Long, last As Long, h As String
    Set ws = ResolveTarget(target, "any")
    kind = SheetKind(ws)
    Select Case kind
        Case "search", "normalized"
            m = ReadMatrix(ws)
            If kind = "normalized" Then
                ColorRange ws.Range(ws.Cells(HDR_ROW + 1, 2), ws.Cells(HDR_ROW + m.nR, m.nC + 1)), 1#
            Else
                ColorRange ws.Range(ws.Cells(HDR_ROW + 1, 2), ws.Cells(HDR_ROW + m.nR, m.nC + 1)), HighThreshold()
            End If
        Case "comparison"
            c = ReadComparison(ws)
            ColorRange ws.Range(ws.Cells(HDR_ROW + 1, 2), ws.Cells(HDR_ROW + 1 + c.nR, 2 + 2 * c.nC)), HighThreshold()
        Case Else
            last = HDR_ROW
            Do While Len(CStr(ws.Cells(last + 1, 1).Value)) > 0
                last = last + 1
            Loop
            If last > HDR_ROW Then
                For j = 1 To 40
                    h = CStr(ws.Cells(HDR_ROW, j).Value)
                    Select Case h
                        Case "N", "LR", "TVR", "THR", "VR", "HR"
                            ColorRange ws.Range(ws.Cells(HDR_ROW + 1, j), ws.Cells(last, j)), 2#
                        Case "NOP_LR", "NOP_THR", "NOP_TVR", "NOP"
                            ColorRange ws.Range(ws.Cells(HDR_ROW + 1, j), ws.Cells(last, j)), HighThreshold()
                    End Select
                Next j
            End If
    End Select
    DoColorize = "colours updated on '" & ws.Name & "' (high threshold " & HighThreshold() & ")"
End Function


Public Function DoTextGen() As String
    Dim a As Variant, b As Variant, nA As Long, nB As Long, i As Long, j As Long, n As Long, k As Long
    Dim prefix As String, buffer As String, suffix As String, mode As String, ws As Worksheet
    Dim hyp() As String, fa() As String, fb() As String, nop() As Double, keys() As Double, ord As Variant, tbl() As Variant
    gStep = "DoTextGen"
    a = ListItems(5, nA)
    b = ListItems(6, nB)
    If nA = 0 Or nB = 0 Then
        Err.Raise ALMA_ERR + 708, "ALMA", "Fill the text generator ranges (columns E and F) on the Lists sheet."
    End If
    prefix = ListSetting("Prefix", "")
    buffer = ListSetting("Buffer", "")
    suffix = ListSetting("Suffix", "")
    mode = LCase$(ListSetting("Mode (all/single)", "all"))
    If mode = "single" Then
        n = nA
        If nB < n Then n = nB
    Else
        n = nA * nB
    End If
    ReDim hyp(1 To n)
    ReDim fa(1 To n)
    ReDim fb(1 To n)
    ReDim nop(1 To n)
    ReDim keys(1 To n, 1 To 2)
    k = 0
    If mode = "single" Then
        For i = 1 To n
            k = k + 1
            fa(k) = a(i)
            fb(k) = b(i)
        Next i
    Else
        For i = 1 To nA
            For j = 1 To nB
                k = k + 1
                fa(k) = a(i)
                fb(k) = b(j)
            Next j
        Next i
    End If
    For k = 1 To n
        hyp(k) = TextQuery(prefix, fa(k), buffer, fb(k), suffix)
        nop(k) = PubMedCount(hyp(k))
        keys(k, 1) = -nop(k)
        keys(k, 2) = k
        Progress "text generator", k, n
    Next k
    ord = SortIndexLex(keys, n, 2)
    ReDim tbl(1 To n + 1, 1 To 4)
    tbl(1, 1) = "hypothesis"
    tbl(1, 2) = "first"
    tbl(1, 3) = "second"
    tbl(1, 4) = "NOP"
    For k = 1 To n
        tbl(k + 1, 1) = hyp(ord(k))
        tbl(k + 1, 2) = fa(ord(k))
        tbl(k + 1, 3) = fb(ord(k))
        tbl(k + 1, 4) = nop(ord(k))
    Next k
    Set ws = WriteTable(tbl, "TextGen", "text", "", TextQuery(prefix, "<first>", buffer, "<second>", suffix))
    ws.Activate
    DoTextGen = n & " generated hypotheses in '" & ws.Name & "'"
End Function


' Temporal / geographic analysis: hypotheses x values with a PubMed field tag.
Public Function DoFilter() As String
    Dim h As Variant, v As Variant, nH As Long, nV As Long, tag As String
    Dim m As AlmaMatrix, nm As AlmaMatrix, ws As Worksheet, ws2 As Worksheet
    gStep = "DoFilter"
    h = ListItems(7, nH)
    v = ListItems(8, nV)
    If nH = 0 Or nV = 0 Then
        Err.Raise ALMA_ERR + 711, "ALMA", "Fill the filter hypotheses (column G) and values (column H) on the Lists sheet."
    End If
    tag = ListSetting("Filter tag", "[dp]")
    m = SearchMatrix(h, nH, v, nV, "", tag, "hypothesis", tag)
    Set ws = WriteMatrix(m, "FM")
    m = ReadMatrix(ws)
    nm = NormalizeMatrix(m, "rows")
    Set ws2 = WriteMatrix(nm, ws.Name & "_norm_rows")
    ws2.Activate
    DoFilter = "filter matrix '" & ws.Name & "' (" & nH & " x " & nV & ", tag " & tag & ") and '" & ws2.Name & "'"
End Function


Public Function DoClearCache() As String
    Dim n As Long
    n = CacheSize()
    If Not Quiet() Then
        If MsgBox("Delete all " & n & " cached PubMed counts?", vbYesNo + vbQuestion, "ALMA") <> vbYes Then
            DoClearCache = "cache kept (" & n & " counts)"
            Exit Function
        End If
    End If
    ClearCacheSheet
    DoClearCache = n & " cached counts deleted"
End Function


'------------------------------------------------------------------ self-test
Private Sub TestEq(ByVal ws As Worksheet, ByRef r As Long, ByRef nPass As Long, ByRef nFail As Long, _
                   ByVal label As String, ByVal got As String, ByVal expected As String)
    gStep = "TestEq"
    ws.Cells(r, 1).Value = label
    ws.Cells(r, 2).NumberFormat = "@"
    ws.Cells(r, 2).Value = got
    ws.Cells(r, 3).NumberFormat = "@"
    ws.Cells(r, 3).Value = expected
    If got = expected Then
        ws.Cells(r, 4).Value = "PASS"
        ws.Cells(r, 4).Interior.Color = RGB(99, 190, 123)
        nPass = nPass + 1
    Else
        ws.Cells(r, 4).Value = "FAIL"
        ws.Cells(r, 4).Interior.Color = RGB(248, 105, 107)
        nFail = nFail + 1
    End If
    r = r + 1
End Sub


Private Function ToyMatrix() As AlmaMatrix
    Dim m As AlmaMatrix, rl() As String, cl() As String, vv() As Double
    m.Kind = "normalized"
    m.Title = "toy"
    m.nR = 3
    m.nC = 3
    ReDim rl(1 To 3)
    ReDim cl(1 To 3)
    ReDim vv(1 To 3, 1 To 3)
    rl(1) = "a"
    rl(2) = "b"
    rl(3) = "c"
    cl(1) = "X"
    cl(2) = "Y"
    cl(3) = "Z"
    vv(1, 1) = 1
    vv(1, 2) = 0.1
    vv(1, 3) = 0
    vv(2, 1) = 0.92
    vv(2, 2) = 1
    vv(2, 3) = 0
    vv(3, 1) = 0
    vv(3, 2) = 0.3
    vv(3, 3) = 1
    m.RowLab = rl
    m.ColLab = cl
    m.V = vv
    ToyMatrix = m
End Function


' Fig 3b example: palbociclib / gefitinib for head and neck cancer + radiotherapy.
Private Function ToyComparison() As AlmaComparison
    Dim c As AlmaComparison, m As AlmaMatrix, rl() As String, cl() As String, vv() As Double
    Dim rv() As Double, cv() As Double, tr() As Double
    m.Kind = "search"
    m.Title = "toy"
    m.nR = 2
    m.nC = 1
    ReDim rl(1 To 2)
    ReDim cl(1 To 1)
    ReDim vv(1 To 2, 1 To 1)
    rl(1) = "Gefitinib"
    rl(2) = "Palbociclib"
    cl(1) = "Head and neck cancer"
    vv(1, 1) = 420
    vv(2, 1) = 16
    m.RowLab = rl
    m.ColLab = cl
    m.V = vv
    CopyBase c, m
    ReDim rv(1 To 2)
    ReDim cv(1 To 1)
    ReDim tr(1 To 2, 1 To 1)
    rv(1) = 462
    rv(2) = 24
    cv(1) = 58292
    tr(1, 1) = 92
    tr(2, 1) = 0
    c.VarName = "radiotherapy"
    c.VarTotal = 354364
    c.RowVar = rv
    c.ColVar = cv
    c.Tri = tr
    ToyComparison = c
End Function


Public Function DoSelfTest() As String
    Dim ws As Worksheet, r As Long, nPass As Long, nFail As Long, n As Double
    Dim m As AlmaMatrix, s As AlmaMatrix, c As AlmaComparison, tbl As Variant
    gStep = "DoSelfTest"
    Set ws = NewSheet("SelfTest")
    ws.Cells(1, 1).Value = "ALMA self-test"
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(2, 1).Value = "Test"
    ws.Cells(2, 2).Value = "Result"
    ws.Cells(2, 3).Value = "Expected"
    ws.Cells(2, 4).Value = "Status"
    ws.Rows(2).Font.Bold = True
    r = 3
    TestEq ws, r, nPass, nFail, "AndQuery", AndQuery("palbociclib", "Head and neck cancer + radiotherapy", ""), _
           "(palbociclib) AND (Head and neck cancer) AND (radiotherapy)"
    TestEq ws, r, nPass, nFail, "TextQuery", TextQuery("treating", "Lung cancer", "with", "EGFR", _
           "inhibitor combined with radiotherapy"), "treating Lung cancer with EGFR inhibitor combined with radiotherapy"
    TestEq ws, r, nPass, nFail, "UrlEncode", UrlEncode("head and neck [tiab]"), "head%20and%20neck%20%5Btiab%5D"
    TestEq ws, r, nPass, nFail, "UrlEncode UTF-8", UrlEncode("Sj" & ChrW(246) & "gren"), "Sj%C3%B6gren"
    TestEq ws, r, nPass, nFail, "ParseCount", CStr(ParseCount( _
           "{""header"":{""type"":""esearch""},""esearchresult"":{""count"":""9""}}")), "9"
    TestEq ws, r, nPass, nFail, "Scores", ReasonScore(16, 20, 5) & "," & ReasonScore(24, 20, 5) & "," & _
           NoveltyScore(0, 20, 5) & "," & NoveltyScore(92, 20, 5), "1,2,2,0"
    m = ToyMatrix()
    s = SortMatrix(m, "sum", "both")
    TestEq ws, r, nPass, nFail, "Sort by sum", Join(s.RowLab, ",") & "|" & Join(s.ColLab, ","), "b,c,a|X,Y,Z"
    s = ClusterSortMatrix(m, 0.9, "rows")
    TestEq ws, r, nPass, nFail, "Cluster sort", Join(s.RowLab, ",") & "|" & Join(s.ColLab, ","), "a,b,c|X,Z,Y"
    c = ToyComparison()
    NrfReset
    NrfAdd c, 20, 5, 1, 10
    tbl = NrfTable()
    TestEq ws, r, nPass, nFail, "NRF ranking (Fig 3b example)", tbl(2, 3) & ":" & tbl(2, 12) & tbl(2, 13) & _
           tbl(2, 14) & tbl(2, 15), "Palbociclib:2122"
    If SettingBool("offline") Then
        ws.Cells(r, 1).Value = "PubMed online test"
        ws.Cells(r, 4).Value = "skipped (offline mode)"
        r = r + 1
    Else
        n = FetchCount("(gefitinib) AND (head and neck cancer)", "edat", "2020/02/10")
        If n >= 380 And n <= 480 Then
            TestEq ws, r, nPass, nFail, "PubMed online: gefitinib + HNC to 2020/02/10 (paper 420)", "in range", "in range"
        Else
            TestEq ws, r, nPass, nFail, "PubMed online: gefitinib + HNC to 2020/02/10 (paper 420)", CStr(n), "380-480"
        End If
    End If
    ws.Columns(1).ColumnWidth = 48
    ws.Columns(2).ColumnWidth = 60
    ws.Columns(3).ColumnWidth = 60
    ws.Activate
    DoSelfTest = nPass & " passed, " & nFail & " failed (sheet '" & ws.Name & "')"
End Function
