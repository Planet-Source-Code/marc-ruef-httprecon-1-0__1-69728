Attribute VB_Name = "modHttpHandler"
Option Explicit

Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Long, ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" (ByVal hHttpSession As Long, ByVal sVerb As String, ByVal sObjectName As String, ByVal sVersion As String, ByVal sReferer As String, ByVal something As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function InternetSetOption Lib "wininet.dll" Alias "InternetSetOptionA" (ByVal hInternet As Long, ByVal lOption As Long, ByRef sBuffer As Any, ByVal lBufferLength As Long) As Integer
Private Declare Function HttpAddRequestHeaders Lib "wininet.dll" Alias "HttpAddRequestHeadersA" (ByVal hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal lModifiers As Long) As Integer
Private Declare Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestA" (ByVal hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal sOptional As String, ByVal lOptionalLength As Long) As Integer
Private Declare Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoA" (ByVal hHttpRequest As Long, ByVal lInfoLevel As Long, ByRef sBuffer As Any, ByRef lBufferLength As Long, ByRef lIndex As Long) As Integer
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer

'These functions are for debugging purposes only. Leave them commented during run-time.
'Private Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer

Private Const HTTP_QUERY_RAW_HEADERS_CRLF As Integer = 22
Private Const INTERNET_SERVICE_HTTP As Integer = 3
Private Const INTERNET_OPEN_TYPE_PRECONFIG As Integer = 0
Private Const INTERNET_OPTION_CONNECT_TIMEOUT = 2
Private Const INTERNET_FLAG_RELOAD = &H80000000
Private Const INTERNET_FLAG_KEEP_CONNECTION = &H400000
Private Const HTTP_ADDREQ_FLAG_ADD = &H20000000
Private Const HTTP_ADDREQ_FLAG_REPLACE = &H80000000

Private Const HTTP_LEGITIMATE_PROTOCOL As String = "HTTP/1.1"
Private Const HTTP_AVAILABLE_RESSOURCE As String = "/"
Private Const HTTP_NOT_AVAILABLE_RESSOURCE As String = "/404test.html"
Private Const LONG_REQUEST_LENGTH As Integer = 1024
Private Const LONG_REQUEST_CHAR As String = "a"
Private Const HTTP_ATTACK_REQUEST As String = "/etc/passwd?format=%%%%&xss=""><script>alert('xss');</script>&traversal=../../&sql=' OR 1;"

Private Const HTTP_MAGIC_ANSWER As Integer = 3

Public response_attackrequest As String
Public response_delete As String
Public response_getexist As String
Public response_getlongrequest As String
Public response_get_nonexistent As String
Public response_head As String
Public response_options As String
Public response_testmethod As String
Public response_protocolversion As String

Public Function SendHttpRequest(ByRef sHost As String, ByRef lPort As Long, sMethod As String, ByRef sURL As String, ByRef sProtocol As String) As String
    Dim sBuffer As String * 1024
    Dim lBufferLength As Long
    Dim hInternetSession As Long
    Dim hInternetConnect As Long
    Dim hHttpOpenRequest As Long
    Dim hHttpSendRequest As Integer
    Dim iNullCharPosition As Integer
    
    lBufferLength = 1024

    sHost = SanitizeHostInput(sHost)
    
    hInternetSession = InternetOpen(APP_NAME, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
    If CBool(hInternetSession) = False Then
        SendHttpRequest = 0
        Exit Function
    End If
    
    hInternetConnect = InternetConnect(hInternetSession, sHost, lPort, "", "", INTERNET_SERVICE_HTTP, 0, 0)
    hHttpOpenRequest = HttpOpenRequest(hInternetConnect, sMethod, sURL, sProtocol, vbNullString, 0, INTERNET_FLAG_RELOAD Or INTERNET_FLAG_KEEP_CONNECTION, 0)
    Call InternetSetOption(hHttpOpenRequest, INTERNET_OPTION_CONNECT_TIMEOUT, 5000, 4)
    hHttpSendRequest = HttpSendRequest(hHttpOpenRequest, vbNullString, 0, vbNullString, 0)
    
    If (hHttpSendRequest) Then
        Call HttpQueryInfo(hHttpOpenRequest, HTTP_QUERY_RAW_HEADERS_CRLF, ByVal sBuffer, lBufferLength, 0)
        
        iNullCharPosition = InStr(1, sBuffer, Chr(0), vbBinaryCompare)
        If (iNullCharPosition > 1) Then
            SendHttpRequest = Mid$(sBuffer, 1, iNullCharPosition - 1)
        Else
            SendHttpRequest = sBuffer
        End If
    End If

    InternetCloseHandle (hHttpOpenRequest)
    InternetCloseHandle (hInternetSession)
    InternetCloseHandle (hInternetConnect)
    DoEvents
End Function

Public Function PostFingerprinToWebsite(ByRef sImplementation As String, ByRef sRemarks As String, ByRef sFingerprint As String) As Integer
    Dim hInternetSession As Long
    Dim hInternetConnect As Long
    Dim hHttpOpenRequest As Long
    Dim sHeader As String
    Dim sPostData As String
  
'    Dim sReadBuffer As String * 2048
'    Dim bDoLoop As Boolean
'    Dim ptrResult As String
'    Dim lNumberOfBytesRead As Long
    
    hInternetSession = InternetOpen(APP_NAME, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
    If CBool(hInternetSession) = False Then
        PostFingerprinToWebsite = 0
        Exit Function
    End If
    
    hInternetConnect = InternetConnect(hInternetSession, PROJECT_WEBSERVER, PROJECT_WEBPORT, "", "", INTERNET_SERVICE_HTTP, 0, 0)
    hHttpOpenRequest = HttpOpenRequest(hInternetConnect, "POST", PROJECT_WEBUPLOAD_FILE, "HTTP/1.1", vbNullString, 0, INTERNET_FLAG_RELOAD Or INTERNET_FLAG_KEEP_CONNECTION, 0)
    
    sHeader = "Content-Type: multipart/form-data; boundary=AaB03x" & vbCrLf
    Call HttpAddRequestHeaders(hHttpOpenRequest, sHeader, Len(sHeader), HTTP_ADDREQ_FLAG_REPLACE Or HTTP_ADDREQ_FLAG_ADD)
    
    sPostData = _
    "--AaB03x" & vbCrLf & _
    "Content-Disposition: form-data; name=""implementation""" & vbCrLf & _
    "Content-Type: text/plain" & vbCrLf & vbCrLf & sImplementation & vbCrLf & "--AaB03x--" & vbCrLf & _
    "--AaB03x" & vbCrLf & _
    "Content-Disposition: form-data; name=""remarks""" & vbCrLf & _
    "Content-Type: text/plain" & vbCrLf & vbCrLf & sRemarks & vbCrLf & "--AaB03x--" & vbCrLf & _
    "--AaB03x" & vbCrLf & _
    "Content-Disposition: form-data; name=""question""" & vbCrLf & _
    "Content-Type: text/plain" & vbCrLf & vbCrLf & HTTP_MAGIC_ANSWER & vbCrLf & "--AaB03x--" & vbCrLf & _
    "--AaB03x" & vbCrLf & _
    "Content-Disposition: form-data; name=""fingerprint""" & vbCrLf & _
    "Content-Type: text/plain" & vbCrLf & vbCrLf & sFingerprint & vbCrLf & "--AaB03x--" & vbCrLf
    
    Call InternetSetOption(hHttpOpenRequest, INTERNET_OPTION_CONNECT_TIMEOUT, 10000, 4)
    Call HttpSendRequest(hHttpOpenRequest, vbNullString, 0, sPostData, Len(sPostData))
    
'    ptrResult = ""
'    Do
'        sReadBuffer = vbNullString
'        bDoLoop = InternetReadFile(hHttpOpenRequest, sReadBuffer, Len(sReadBuffer), lNumberOfBytesRead)
'        ptrResult = ptrResult & Left(sReadBuffer, lNumberOfBytesRead)
'        If Not CBool(lNumberOfBytesRead) Or Not bDoLoop Then Exit Do
'    Loop
    
    InternetCloseHandle (hHttpOpenRequest)
    InternetCloseHandle (hInternetSession)
    InternetCloseHandle (hInternetConnect)
End Function

Public Function RunTestRequests(ByRef sTargetHost As String, ByRef lTargetPort As Long) As Boolean
    response_getexist = SendHttpRequest(sTargetHost, lTargetPort, "GET", HTTP_AVAILABLE_RESSOURCE, HTTP_LEGITIMATE_PROTOCOL)
    
    If (LenB(response_getexist)) Then
        response_getlongrequest = SendHttpRequest(sTargetHost, lTargetPort, "GET", "/" & String$(LONG_REQUEST_LENGTH, LONG_REQUEST_CHAR), HTTP_LEGITIMATE_PROTOCOL)
        response_get_nonexistent = SendHttpRequest(sTargetHost, lTargetPort, "GET", HTTP_NOT_AVAILABLE_RESSOURCE, HTTP_LEGITIMATE_PROTOCOL)
        response_protocolversion = SendHttpRequest(sTargetHost, lTargetPort, "GET", HTTP_AVAILABLE_RESSOURCE, "HTTP/9.8")
        response_head = SendHttpRequest(sTargetHost, lTargetPort, "HEAD", HTTP_AVAILABLE_RESSOURCE, HTTP_LEGITIMATE_PROTOCOL)
        response_options = SendHttpRequest(sTargetHost, lTargetPort, "OPTIONS", "/", HTTP_LEGITIMATE_PROTOCOL)
        response_delete = SendHttpRequest(sTargetHost, lTargetPort, "DELETE", HTTP_AVAILABLE_RESSOURCE, HTTP_LEGITIMATE_PROTOCOL)
        response_testmethod = SendHttpRequest(sTargetHost, lTargetPort, "TEST", HTTP_AVAILABLE_RESSOURCE, HTTP_LEGITIMATE_PROTOCOL)
        response_attackrequest = SendHttpRequest(sTargetHost, lTargetPort, "GET", HTTP_ATTACK_REQUEST, HTTP_LEGITIMATE_PROTOCOL)
        
        RunTestRequests = True
    Else
        response_getlongrequest = vbNullString
        response_get_nonexistent = vbNullString
        response_protocolversion = vbNullString
        response_head = vbNullString
        response_options = vbNullString
        response_delete = vbNullString
        response_testmethod = vbNullString
        response_attackrequest = vbNullString
        
        RunTestRequests = False
    End If
End Function

Public Sub FillResponses()
    Dim iIndex As Integer
    
    iIndex = frmMain.tbsViews.SelectedItem.Index

    If (iIndex = 1) Then
        frmMain.txtResponses.Text = response_getexist
    ElseIf (iIndex = 2) Then
        frmMain.txtResponses.Text = response_getlongrequest
    ElseIf (iIndex = 3) Then
        frmMain.txtResponses.Text = response_get_nonexistent
    ElseIf (iIndex = 4) Then
        frmMain.txtResponses.Text = response_protocolversion
    ElseIf (iIndex = 5) Then
        frmMain.txtResponses.Text = response_head
    ElseIf (iIndex = 6) Then
        frmMain.txtResponses.Text = response_options
    ElseIf (iIndex = 7) Then
        frmMain.txtResponses.Text = response_delete
    ElseIf (iIndex = 8) Then
        frmMain.txtResponses.Text = response_testmethod
    ElseIf (iIndex = 9) Then
        frmMain.txtResponses.Text = response_attackrequest
    End If
End Sub

Public Function SanitizeHostInput(ByRef sHost As String) As String
    sHost = Trim$(sHost)
    sHost = LCase(sHost)

    If (Left$(sHost, 7) = "http://") Then
        sHost = Right$(sHost, Len(sHost) - 7)
    ElseIf (Left$(sHost, 8) = "https://") Then
        sHost = Right$(sHost, Len(sHost) - 8)
    ElseIf (Left$(sHost, 6) = "ftp://") Then
        sHost = Right$(sHost, Len(sHost) - 6)
    ElseIf (Left$(sHost, 2) = "\\") Then
        sHost = Right$(sHost, Len(sHost) - 2)
    End If

    sHost = TrimSymbol(sHost, ":")
    sHost = TrimSymbol(sHost, ";")
    sHost = TrimSymbol(sHost, "/")
    sHost = TrimSymbol(sHost, "\")
    sHost = TrimSymbol(sHost, "*")
    sHost = TrimSymbol(sHost, "@")
    sHost = TrimSymbol(sHost, "%")
    sHost = TrimSymbol(sHost, " ")
    
    SanitizeHostInput = sHost
End Function

Private Function TrimSymbol(ByRef sInput As String, ByRef sSymbol As String) As String
    Dim iPosition As Integer
    
    iPosition = InStr(1, sInput, sSymbol, vbBinaryCompare)
    
    If (iPosition) Then
        TrimSymbol = Mid$(sInput, 1, iPosition - 1)
    Else
        TrimSymbol = sInput
    End If
End Function
