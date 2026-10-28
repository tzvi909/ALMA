Attribute VB_Name = "ALMA_PubMed"
Option Explicit
'==============================================================================
' ALMA for Excel - settings, PubMed query building, NOP retrieval and cache.
' Re-implementation of ALMA (Harif et al., Technion).  The query convention is
' identical to the Python version (alma/core.py) so both share one cache:
'   every element in parentheses, joined with " AND "
'   a label may contain several elements separated by " + "
'==============================================================================

Public Const ALMA_SEP As String = " + "
Public Const SETTINGS_SHEET As String = "ALMA"
Public Const CACHE_SHEET As String = "Cache"
Private Const ESEARCH_URL As String = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi"
Private Const MIN_DATE As String = "1000/01/01"

Private mCache As Collection
Private mCacheNext As Long
Private mLastCall As Double
Public gFetched As Long
Public gStep As String
' = vbObjectError (a constant some VBA hosts do not define)
Public Const ALMA_ERR As Long = -2147221504
Public gFromCache As Long

'------------------------------------------------------------------ settings
' Settings live on the ALMA sheet: key in column A (hidden), label in B, value in C.
Public Function Setting(ByVal key As String) As Variant
    Dim ws As Worksheet, r As Long
    Setting = Empty
    Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET)
    For r = 1 To 100
        If LCase$(Trim$(CStr(ws.Cells(r, 1).Value))) = LCase$(key) Then
            Setting = ws.Cells(r, 3).Value
            Exit Function
        End If
    Next r
End Function

Public Sub SetSetting(ByVal key As String, ByVal v As Variant)
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET)
    For r = 1 To 100
        If LCase$(Trim$(CStr(ws.Cells(r, 1).Value))) = LCase$(key) Then
            ws.Cells(r, 3).Value = v
            Exit Sub
        End If
    Next r
End Sub

Public Function SettingNum(ByVal key As String, ByVal dflt As Double) As Double
    Dim v As Variant
    v = Setting(key)
    SettingNum = dflt
    If Not IsEmpty(v) Then
        If IsNumeric(v) Then SettingNum = CDbl(v)
    End If
End Function

Public Function SettingBool(ByVal key As String) As Boolean
    Dim v As Variant
    v = Setting(key)
    SettingBool = False
    If VarType(v) = vbBoolean Then
        SettingBool = v
    Else
        Select Case LCase$(Trim$(CStr(v)))
            Case "true", "yes", "y", "1", "wahr", "vrai"
                SettingBool = True
        End Select
    End If
End Function

Public Function HighThreshold() As Double
    HighThreshold = SettingNum("high", 20)
End Function

Public Function MedThreshold() As Double
    MedThreshold = SettingNum("med", 5)
End Function

Public Function Quiet() As Boolean
    Quiet = SettingBool("quiet")
End Function

' Date limit as YYYY/MM/DD text, whatever Excel turned the cell into.
Public Function MaxDateSetting() As String
    Dim v As Variant
    v = Setting("maxdate")
    MaxDateSetting = ""
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbDate Then
        MaxDateSetting = YMD(CDate(v))
    ElseIf VarType(v) = vbDouble Or VarType(v) = vbLong Or VarType(v) = vbInteger Then
        If v >= 1000 And v <= 9999 Then
            MaxDateSetting = CStr(CLng(v))
        ElseIf v > 9999 Then
            MaxDateSetting = YMD(CDate(v))
        End If
    Else
        MaxDateSetting = Replace(Replace(Trim$(CStr(v)), "-", "/"), ".", "/")
    End If
End Function

Public Function DateTypeSetting() As String
    DateTypeSetting = ""
    If Len(MaxDateSetting()) > 0 Then
        DateTypeSetting = LCase$(Trim$(CStr(Setting("datetype"))))
        If Len(DateTypeSetting) = 0 Then DateTypeSetting = "edat"
    End If
End Function

Public Function PubMedNote() As String
    If Len(MaxDateSetting()) > 0 Then
        PubMedNote = "maxdate " & MaxDateSetting() & " (" & DateTypeSetting() & ")"
    Else
        PubMedNote = "no date limit (searched " & YMD(Date) & ")"
    End If
End Function

Private Function YMD(ByVal d As Date) As String
    YMD = Year(d) & "/" & Right$("0" & Month(d), 2) & "/" & Right$("0" & Day(d), 2)
End Function

'------------------------------------------------------------------- queries
Public Function AndQuery(ParamArray parts() As Variant) As String
    Dim i As Long, q As String
    q = ""
    For i = LBound(parts) To UBound(parts)
        q = AppendTerms(q, CStr(parts(i)))
    Next i
    AndQuery = q
End Function

Private Function AppendTerms(ByVal q As String, ByVal label As String) As String
    Dim bits() As String, k As Long, t As String
    If Len(Trim$(label)) > 0 Then
        bits = Split(label, ALMA_SEP)
        For k = LBound(bits) To UBound(bits)
            t = Trim$(bits(k))
            If Len(t) > 0 Then
                If Len(q) > 0 Then q = q & " AND "
                q = q & "(" & t & ")"
            End If
        Next k
    End If
    AppendTerms = q
End Function

' Hypothesis text generator: prefix + first + buffer + second + suffix
Public Function TextQuery(ByVal prefix As String, ByVal first As String, ByVal buffer As String, _
                          ByVal second As String, ByVal suffix As String) As String
    Dim parts(1 To 5) As String, i As Long, q As String
    parts(1) = prefix
    parts(2) = first
    parts(3) = buffer
    parts(4) = second
    parts(5) = suffix
    q = ""
    For i = 1 To 5
        If Len(Trim$(parts(i))) > 0 Then
            If Len(q) > 0 Then q = q & " "
            q = q & Trim$(parts(i))
        End If
    Next i
    TextQuery = q
End Function

Public Function Tagged(ByVal label As String, ByVal tag As String) As String
    Tagged = label & tag
End Function

Public Function UrlEncode(ByVal s As String) As String
    Dim i As Long, c As Long, ch As String, out As String
    out = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        c = AscW(ch)
        If c < 0 Then c = c + 65536
        If (c >= 48 And c <= 57) Or (c >= 65 And c <= 90) Or (c >= 97 And c <= 122) _
           Or c = 45 Or c = 46 Or c = 95 Or c = 126 Then
            out = out & ch
        ElseIf c < 128 Then
            out = out & "%" & Right$("0" & Hex$(c), 2)
        ElseIf c < 2048 Then
            out = out & "%" & Hex$(192 + (c \ 64)) & "%" & Hex$(128 + (c Mod 64))
        Else
            out = out & "%" & Hex$(224 + (c \ 4096)) & "%" & Hex$(128 + ((c \ 64) Mod 64)) _
                  & "%" & Hex$(128 + (c Mod 64))
        End If
    Next i
    UrlEncode = out
End Function

' Extract the hit count from an E-utilities esearch JSON reply; -1 if absent.
Public Function ParseCount(ByVal s As String) As Double
    Dim p As Long, i As Long, ch As String, digits As String
    ParseCount = -1
    digits = ""
    p = InStr(1, s, """count""", vbBinaryCompare)
    If p = 0 Then Exit Function
    For i = p + 7 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then
            digits = digits & ch
        ElseIf Len(digits) > 0 Then
            Exit For
        ElseIf ch <> ":" And ch <> """" And ch <> " " Then
            Exit For
        End If
    Next i
    If Len(digits) > 0 Then ParseCount = CDbl(digits)
End Function

'--------------------------------------------------------------------- cache
Private Function CacheKey(ByVal term As String, ByVal dtype As String, ByVal maxd As String) As String
    CacheKey = term & vbTab & dtype & vbTab & maxd
End Function

Public Sub ResetCache()
    Set mCache = Nothing
End Sub

Private Sub EnsureCache()
    If mCache Is Nothing Then LoadCache
End Sub

Private Sub LoadCache()
    Dim ws As Worksheet, last As Long, r As Long, data As Variant
    Set mCache = New Collection
    Set ws = ThisWorkbook.Worksheets(CACHE_SHEET)
    gStep = "LoadCache"
    last = LastUsedRow(ws, 1)
    If last < 2 Then
        mCacheNext = 2
        Exit Sub
    End If
    mCacheNext = last + 1
    If last = 2 Then
        AddCacheItem CStr(ws.Cells(2, 1).Value), CStr(ws.Cells(2, 2).Value), CStr(ws.Cells(2, 3).Value), ws.Cells(2, 4).Value
        Exit Sub
    End If
    data = ws.Range(ws.Cells(2, 1), ws.Cells(last, 4)).Value
    For r = 1 To UBound(data, 1)
        AddCacheItem CStr(data(r, 1)), CStr(data(r, 2)), CStr(data(r, 3)), data(r, 4)
    Next r
End Sub

Private Sub AddCacheItem(ByVal term As String, ByVal dtype As String, ByVal maxd As String, ByVal n As Variant)
    If Len(term) = 0 Then Exit Sub
    If IsEmpty(n) Then Exit Sub
    If Not IsNumeric(n) Then Exit Sub
    On Error Resume Next
    mCache.Add CDbl(n), CacheKey(term, dtype, maxd)
    On Error GoTo 0
End Sub

Private Function CacheLookup(ByVal key As String, ByRef n As Double) As Boolean
    On Error GoTo miss
    n = mCache.Item(key)
    CacheLookup = True
    Exit Function
miss:
    CacheLookup = False
End Function

Private Sub CachePut(ByVal term As String, ByVal dtype As String, ByVal maxd As String, ByVal n As Double)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(CACHE_SHEET)
    gStep = "CachePut"
    ws.Cells(mCacheNext, 1).NumberFormat = "@"
    ws.Cells(mCacheNext, 1).Value = term
    ws.Cells(mCacheNext, 2).NumberFormat = "@"
    ws.Cells(mCacheNext, 2).Value = dtype
    ws.Cells(mCacheNext, 3).NumberFormat = "@"
    ws.Cells(mCacheNext, 3).Value = maxd
    ws.Cells(mCacheNext, 4).Value = n
    ws.Cells(mCacheNext, 5).Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    mCacheNext = mCacheNext + 1
    AddCacheItem term, dtype, maxd, n
End Sub

' Last row before the first empty cell of a column (scans in blocks; works without a window).
Public Function LastUsedRow(ByVal ws As Worksheet, ByVal col As Long) As Long
    Dim r As Long, i As Long, blk As Long, data As Variant
    blk = 2000
    r = 1
    LastUsedRow = 0
    Do
        data = ws.Range(ws.Cells(r, col), ws.Cells(r + blk - 1, col)).Value
        For i = 1 To blk
            If IsEmpty(data(i, 1)) Then
                LastUsedRow = r + i - 2
                Exit Function
            End If
            If Len(CStr(data(i, 1))) = 0 Then
                LastUsedRow = r + i - 2
                Exit Function
            End If
        Next i
        r = r + blk
    Loop
End Function

Public Function CacheSize() As Long
    EnsureCache
    CacheSize = mCache.Count
End Function

Public Sub ClearCacheSheet()
    Dim ws As Worksheet, last As Long
    Set ws = ThisWorkbook.Worksheets(CACHE_SHEET)
    last = LastUsedRow(ws, 1)
    If last >= 2 Then ws.Range(ws.Cells(2, 1), ws.Cells(last, 5)).ClearContents
    Set mCache = Nothing
End Sub

'-------------------------------------------------------------------- PubMed
' Number of publications (NOP) of a PubMed query, honouring the date limit,
' the cache and offline mode set on the ALMA sheet.
Public Function PubMedCount(ByVal term As String) As Double
    Dim maxd As String, dtype As String, n As Double
    maxd = MaxDateSetting()
    dtype = DateTypeSetting()
    EnsureCache
    If SettingBool("usecache") Or SettingBool("offline") Then
        If CacheLookup(CacheKey(term, dtype, maxd), n) Then
            PubMedCount = n
            gFromCache = gFromCache + 1
            Exit Function
        End If
    End If
    If SettingBool("offline") Then
        Err.Raise ALMA_ERR + 513, "ALMA", "Offline mode: query not in the Cache sheet: " & term
    End If
    n = FetchCount(term, dtype, maxd)
    CachePut term, dtype, maxd, n
    PubMedCount = n
End Function

' E-utilities esearch URL that returns the hit count of a query.
Public Function EsearchUrl(ByVal term As String, ByVal dtype As String, ByVal maxd As String) As String
    Dim url As String, apiKey As String
    url = ESEARCH_URL & "?db=pubmed&rettype=count&retmode=json&tool=ALMA-Excel&term=" & UrlEncode(term)
    If Len(maxd) > 0 Then
        If Len(dtype) = 0 Then dtype = "edat"
        url = url & "&datetype=" & dtype & "&mindate=" & UrlEncode(MIN_DATE) & "&maxdate=" & UrlEncode(maxd)
    End If
    apiKey = Trim$(CStr(Setting("apikey")))
    If Len(apiKey) > 0 Then url = url & "&api_key=" & UrlEncode(apiKey)
    EsearchUrl = url
End Function

Public Function FetchCount(ByVal term As String, ByVal dtype As String, ByVal maxd As String) As Double
    Dim url As String, resp As String, status As Long, attempt As Long, n As Double
    url = EsearchUrl(term, dtype, maxd)
    status = 0
    For attempt = 1 To 8
        Throttle
        resp = HttpGet(url, status)
        If status = 200 Then
            n = ParseCount(resp)
            If n >= 0 Then
                FetchCount = n
                gFetched = gFetched + 1
                Exit Function
            End If
        End If
        WaitMs 1500 * attempt
    Next attempt
    Err.Raise ALMA_ERR + 514, "ALMA", "PubMed request failed (HTTP " & status & ") for: " & term
End Function

Private Function HttpGet(ByVal url As String, ByRef status As Long) As String
    status = 0
    If IsMacOS() Then
        HttpGet = HttpGetMac(url, status)
    Else
        HttpGet = HttpGetWin(url, status)
    End If
End Function

Public Function IsMacOS() As Boolean
    IsMacOS = (InStr(1, Application.OperatingSystem, "Mac", vbTextCompare) > 0)
End Function

' Keep under NCBI's limit (3 requests/s without an API key, 10/s with one).
Private Sub Throttle()
    Dim gap As Double, waited As Double
    gap = SettingNum("delay", 350) / 1000#
    If Len(Trim$(CStr(Setting("apikey")))) = 0 And gap < 0.34 Then gap = 0.34
    If gap < 0.1 Then gap = 0.1
    If mLastCall > 0 Then
        waited = Timer - mLastCall
        If waited >= 0 And waited < gap Then WaitMs CLng((gap - waited) * 1000)
    End If
    mLastCall = Timer
End Sub

Public Sub WaitMs(ByVal ms As Long)
    Dim t0 As Double
    t0 = Timer
    Do While (Timer - t0) * 1000# < ms
        If Timer < t0 Then Exit Do
        DoEvents
    Loop
End Sub
