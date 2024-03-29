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

Public req_timeout_connect As Integer
Public req_timeout_send As Integer
Public req_timeout_receive As Integer

Public req_protocol_legitimate As String
Public req_protocol_wrong As String
Public req_resource_available As String
Public req_resource_notavailable As String
Public req_resource_attack As String
Public req_longrequest_length As Integer
Public req_longrequest_char As String
Public req_method_notallowed As String
Public req_method_notexisting As String
Public req_agent_name As String

Public tests_count As Integer

Public response_attackrequest As String
Public response_delete As String
Public response_getexist As String
Public response_getlongrequest As String
Public response_get_nonexistent As String
Public response_head As String
Public response_options As String
Public response_testmethod As String
Public response_protocolversion As String

Private Const INTERNET_SERVICE_HTTP As Integer = 3
Private Const INTERNET_OPEN_TYPE_PRECONFIG As Integer = 0
Private Const INTERNET_FLAG_RELOAD = &H80000000
Private Const INTERNET_FLAG_KEEP_CONNECTION = &H400000
Private Const INTERNET_OPTION_CONNECT_TIMEOUT = 2
Private Const INTERNET_OPTION_RECEIVE_TIMEOUT = 6
Private Const INTERNET_OPTION_SEND_TIMEOUT = 5

Private Const HTTP_QUERY_RAW_HEADERS_CRLF As Integer = 22
Private Const HTTP_ADDREQ_FLAG_ADD = &H20000000
Private Const HTTP_ADDREQ_FLAG_REPLACE = &H80000000

Private Const INTERNET_OPTION_SECURITY_FLAGS = 31
Private Const INTERNET_FLAG_IGNORE_UNKNOWN_CA = &H100
Private Const INTERNET_FLAG_IGNORE_CERT_DATE_INVALID = &H2000
Private Const INTERNET_FLAG_IGNORE_CERT_CN_INVALID = &H1000
Private Const INTERNET_FLAG_SECURE = &H800000

Private Const HTTP_MAGIC_ANSWER As Integer = 3

Public Function SendHttpRequest(ByRef sHost As String, ByRef lPort As Long, sMethod As String, ByRef sURL As String, ByRef sProtocol As String, ByRef iSecure As Integer) As String
    Dim sBuffer As String * 1024
    Dim lBufferLength As Long
    Dim hInternetSession As Long
    Dim hInternetConnect As Long
    Dim hHttpOpenRequest As Long
    Dim hHttpSendRequest As Integer
    Dim iNullCharPosition As Integer
    Dim lSecFlag As Long
    
    lBufferLength = 1024

    sHost = SanitizeHostInput(sHost)
    
    If (iSecure = 1) Then
        lSecFlag = INTERNET_FLAG_SECURE Or _
            INTERNET_FLAG_IGNORE_CERT_CN_INVALID Or _
            INTERNET_FLAG_IGNORE_CERT_DATE_INVALID
    Else
        lSecFlag = 0
    End If
    
    Call ChangeStatusBar("Sending request " & Chr(34) & sMethod & " " & sURL & " " & sProtocol & Chr(34) & "...")
    
    hInternetSession = InternetOpen(req_agent_name, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
    If CBool(hInternetSession) = False Then
        SendHttpRequest = 0
        Exit Function
    End If
    
    hInternetConnect = InternetConnect(hInternetSession, sHost, lPort, "", "", INTERNET_SERVICE_HTTP, 0, 0)
    hHttpOpenRequest = HttpOpenRequest(hInternetConnect, sMethod, sURL, sProtocol, vbNullString, 0, INTERNET_FLAG_RELOAD Or INTERNET_FLAG_KEEP_CONNECTION Or lSecFlag, 0)
    Call InternetSetOption(hHttpOpenRequest, INTERNET_OPTION_CONNECT_TIMEOUT, req_timeout_connect, 4)
    Call InternetSetOption(hHttpOpenRequest, INTERNET_OPTION_SEND_TIMEOUT, req_timeout_send, 4)
    Call InternetSetOption(hHttpOpenRequest, INTERNET_OPTION_RECEIVE_TIMEOUT, req_timeout_receive, 4)
    
    If (iSecure = 1) Then
        Call InternetSetOption(hHttpOpenRequest, INTERNET_OPTION_SECURITY_FLAGS, INTERNET_FLAG_IGNORE_UNKNOWN_CA, 4)
    End If
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

    Call InternetCloseHandle(hHttpOpenRequest)
    Call InternetCloseHandle(hInternetSession)
    Call InternetCloseHandle(hInternetConnect)
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
    
    Call ChangeStatusBar("Uploading new fingerprint...")
    
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
    
    Call InternetCloseHandle(hHttpOpenRequest)
    Call InternetCloseHandle(hInternetSession)
    Call InternetCloseHandle(hInternetConnect)
    
    Call ChangeStatusBarDone
End Function

Public Function RunTestRequests(ByRef sTargetHost As String, ByRef lTargetPort As Long, ByRef iSecure As Integer) As Boolean
    response_getlongrequest = vbNullString
    response_get_nonexistent = vbNullString
    response_protocolversion = vbNullString
    response_head = vbNullString
    response_options = vbNullString
    response_delete = vbNullString
    response_testmethod = vbNullString
    response_attackrequest = vbNullString
    
    If (scan_test_getexisting = 1) Then
        response_getexist = SendHttpRequest(sTargetHost, lTargetPort, "GET", req_resource_available, req_protocol_legitimate, iSecure)
        
        If (LenB(response_getexist)) Then
            frmMain.fraTarget.Caption = "Target (" & PreFetchBanner(response_getexist) & ")"
            
            If (scan_test_getlong <> 0) Then
                response_getlongrequest = SendHttpRequest(sTargetHost, lTargetPort, "GET", "/" & String$(req_longrequest_length, req_longrequest_char), req_protocol_legitimate, iSecure)
            End If
            
            If (scan_test_getnonexisting <> 0) Then
                response_get_nonexistent = SendHttpRequest(sTargetHost, lTargetPort, "GET", req_resource_notavailable, req_protocol_legitimate, iSecure)
            End If
                
            If (scan_test_wrongprotocol <> 0) Then
                response_protocolversion = SendHttpRequest(sTargetHost, lTargetPort, "GET", req_resource_available, req_protocol_wrong, iSecure)
            End If
                            
            If (scan_test_head <> 0) Then
                response_head = SendHttpRequest(sTargetHost, lTargetPort, "HEAD", req_resource_available, req_protocol_legitimate, iSecure)
            End If
            
            If (scan_test_options <> 0) Then
                response_options = SendHttpRequest(sTargetHost, lTargetPort, "OPTIONS", "/", req_protocol_legitimate, iSecure)
            End If
                
            If (scan_test_wrongmethod <> 0) Then
                response_delete = SendHttpRequest(sTargetHost, lTargetPort, req_method_notallowed, req_resource_available, req_protocol_legitimate, iSecure)
            End If
                
            If (scan_test_nonexistingmethod <> 0) Then
                response_testmethod = SendHttpRequest(sTargetHost, lTargetPort, req_method_notexisting, req_resource_available, req_protocol_legitimate, iSecure)
            End If
                
            If (scan_test_attack <> 0) Then
                response_attackrequest = SendHttpRequest(sTargetHost, lTargetPort, "GET", req_resource_attack, req_protocol_legitimate, iSecure)
            End If
            
            RunTestRequests = True
        Else
            RunTestRequests = False
        End If
    End If
End Function

Public Sub AddTestCount(ByRef sTestname As String)
    If (LenB(sTestname)) Then
        tests_count = tests_count + 1
    End If
End Sub

Public Function SanitizeHostInput(ByRef sHost As String) As String
    sHost = Trim$(sHost)
    sHost = LCase(sHost)

    Call TrimPrefix(sHost, "http://")
    Call TrimPrefix(sHost, "https://")
    Call TrimPrefix(sHost, "ftp://")
    Call TrimPrefix(sHost, "\\")

    Call TrimSuffix(sHost, ":")
    Call TrimSuffix(sHost, ";")
    Call TrimSuffix(sHost, "/")
    Call TrimSuffix(sHost, "\")
    Call TrimSuffix(sHost, "*")
    Call TrimSuffix(sHost, "@")
    Call TrimSuffix(sHost, "%")
    Call TrimSuffix(sHost, " ")
    
    SanitizeHostInput = sHost
End Function

Private Sub TrimPrefix(ByRef sInput As String, ByRef sSymbol As String)
    Dim iLength As Integer
    
    iLength = Len(sSymbol)

    If (Left$(sInput, iLength) = sSymbol) Then
        sInput = Mid$(sInput, iLength + 1, Len(sInput) - iLength)
    End If
End Sub

Private Sub TrimSuffix(ByRef sInput As String, ByRef sSymbol As String)
    Dim iPosition As Integer
    
    iPosition = InStr(1, sInput, sSymbol, vbBinaryCompare)
    
    If (iPosition) Then
        sInput = Mid$(sInput, 1, iPosition - 1)
    End If
End Sub

Public Function ExtractTargetPort(ByRef sInput As String) As Integer
    Dim iPositionPortStart As Integer
    Dim iPositionHostStart As Integer
    Dim iPotentialPort As Integer
    
    iPositionHostStart = InStr(1, sInput, "://", vbBinaryCompare)
    If (iPositionHostStart) Then
        iPositionPortStart = InStr(iPositionHostStart + 1, sInput, ":", vbBinaryCompare)
    Else
        iPositionPortStart = InStr(1, sInput, ":", vbBinaryCompare)
    End If
    
    If (iPositionPortStart) Then
        iPotentialPort = CInt(Val(Mid$(sInput, iPositionPortStart + 1, Len(sInput) - iPositionPortStart)))
        
        If (iPotentialPort = 0) Then
            ExtractTargetPort = 80
        ElseIf (iPotentialPort > 65535) Then
            ExtractTargetPort = 80
        Else
            ExtractTargetPort = iPotentialPort
        End If
    ElseIf (Left$(sInput, 8) = "https://") Then
        ExtractTargetPort = 443
    Else
        ExtractTargetPort = 80
    End If
End Function
