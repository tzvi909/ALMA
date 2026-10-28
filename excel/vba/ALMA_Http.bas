Attribute VB_Name = "ALMA_Http"
Option Explicit
'==============================================================================
' ALMA for Excel - platform specific HTTP GET used by ALMA_PubMed.
' Windows: MSXML2.ServerXMLHTTP 6.0 (falls back to WinHTTP / XMLHTTP).
' macOS:   curl through MacScript.
'==============================================================================

Public Function HttpGetWin(ByVal url As String, ByRef status As Long) As String
    Dim http As Object
    status = 0
    HttpGetWin = ""
    On Error Resume Next
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    If http Is Nothing Then Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    If http Is Nothing Then Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    If http Is Nothing Then Exit Function
    http.setTimeouts 15000, 15000, 30000, 60000
    Err.Clear
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "ALMA-Excel/1.0"
    http.send
    If Err.Number <> 0 Then Exit Function
    status = http.Status
    HttpGetWin = http.responseText
End Function

Public Function HttpGetMac(ByVal url As String, ByRef status As Long) As String
    Dim res As String
    status = 0
    HttpGetMac = ""
    On Error Resume Next
    res = MacScript("do shell script ""curl -s -m 60 '" & url & "'""")
    If Err.Number = 0 And Len(res) > 0 Then status = 200
    HttpGetMac = res
End Function
