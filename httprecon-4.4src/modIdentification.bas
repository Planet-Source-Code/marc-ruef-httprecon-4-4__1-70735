Attribute VB_Name = "modIdentification"
Option Explicit

Private Const APP_HITPOINTS_DELIMITER As String = ":"

Public app_hitpoints_minimum As Integer
Public app_hitpoints_maximum As Integer

Public scan_besthitcount As Integer
Public scan_besthitname As String

Public scan_time As String
Public scan_date As String
Public scan_targethost As String
Public scan_targetport As Long
Public scan_targetsecure As Integer

Public scan_test_getexisting As Integer
Public scan_test_getnonexisting As Integer
Public scan_test_getlong As Integer
Public scan_test_head As Integer
Public scan_test_options As Integer
Public scan_test_wrongmethod As Integer
Public scan_test_nonexistingmethod As Integer
Public scan_test_wrongprotocol As Integer
Public scan_test_attack As Integer

Public Function FindMatchInDatabase(ByRef sDatabase As String, ByRef sFingerprint As String) As String
    Dim sDatabaseContent() As String
    Dim sFingerprintInDatabase As String
    Dim iDatabaseEntries As Integer
    Dim iDelimiterPosition As Integer
    Dim i As Integer
    Dim cMatches As Concat
    
    Set cMatches = New Concat
    
    sDatabaseContent = Split(ReadFile(sDatabase), vbCrLf, , vbBinaryCompare)
    iDatabaseEntries = UBound(sDatabaseContent)
    
    For i = 0 To iDatabaseEntries
        If LenB(sFingerprint) Then
            If LenB(sDatabaseContent(i)) Then
                iDelimiterPosition = InStr(1, sDatabaseContent(i), ";", vbBinaryCompare)
                sFingerprintInDatabase = Mid$(sDatabaseContent(i), iDelimiterPosition + 1, Len(sDatabaseContent(i)) - iDelimiterPosition)
                
                If (sFingerprintInDatabase = sFingerprint) Then
                    cMatches.Concat Mid$(sDatabaseContent(i), 1, InStr(1, sDatabaseContent(i), ";", vbBinaryCompare) - 1)
                    
                    If (i < iDatabaseEntries) Then
                        cMatches.Concat ";"
                    End If
                End If
            End If
        End If
    Next i
    
    FindMatchInDatabase = cMatches.Value
End Function

Public Function GenerateMatchStatistics(ByRef sMatchList As String) As String
    Dim sMatches() As String
    Dim iMatchCount As Integer
    Dim i As Integer
    Dim j As Integer
    Dim bDuplicate As Boolean
    Dim cMatchStatistic As Concat
    
    Set cMatchStatistic = New Concat
    
    sMatches = Split(sMatchList, ";", , vbBinaryCompare)
    Call RemoveDuplicatesFromArray(sMatches)
    iMatchCount = UBound(sMatches)
    
    For i = 0 To iMatchCount
        If (LenB(sMatches(i))) Then
            cMatchStatistic.Concat sMatches(i) & APP_HITPOINTS_DELIMITER & ArrayCountIf(sMatchList, sMatches(i)) & vbCrLf
        End If
    Next i
    
    GenerateMatchStatistics = cMatchStatistic.Value
End Function

Public Sub RemoveDuplicatesFromArray(ByRef sArray() As String)
    Dim lLowBound As Long
    Dim lUpBound As Long
    Dim sTempArray() As String
    Dim lCurrent As Long
    Dim i As Long
    Dim j As Long
    
    lUpBound = UBound(sArray)
    
    If (lUpBound > 0) Then
        lLowBound = LBound(sArray)
        
        ReDim sTempArray(lLowBound To lUpBound)
        
        lCurrent = lLowBound
        sTempArray(lCurrent) = sArray(lLowBound)
        
        For i = lLowBound + 1 To lUpBound
            For j = lLowBound To lCurrent
                If LenB(sTempArray(j)) = LenB(sArray(i)) Then
                    If InStrB(1, sArray(i), sTempArray(j), vbBinaryCompare) = 1 Then
                        Exit For
                    End If
                End If
            Next j
            
            If j > lCurrent Then
                lCurrent = j
                sTempArray(lCurrent) = sArray(i)
            End If
        Next i
        
        ReDim Preserve sTempArray(lLowBound To lCurrent)
        sArray = sTempArray
    End If
End Sub

Public Function ArrayCountIf(ByRef sInput As String, ByRef sSearch As String) As Integer
    Dim sArray() As String
    Dim iArrayCount As Integer
    Dim i As Integer
    Dim iSum As Integer
    
    sArray = Split(sInput, ";", , vbBinaryCompare)
    iArrayCount = UBound(sArray)
    
    For i = 0 To iArrayCount
        If (sArray(i) = sSearch) Then
            iSum = iSum + 1
        End If
    Next i
    
    ArrayCountIf = iSum
End Function

Public Function GenerateHttpdIcon(ByVal sImplementation As String) As Integer
    sImplementation = LCase(sImplementation)

    If (InStrB(1, sImplementation, "aol", vbBinaryCompare)) Then
        GenerateHttpdIcon = 1
    ElseIf (InStrB(1, sImplementation, "abyss", vbBinaryCompare)) Then
        GenerateHttpdIcon = 40
    ElseIf (InStrB(1, sImplementation, "allegro", vbBinaryCompare)) Then
        GenerateHttpdIcon = 91
    ElseIf (InStrB(1, sImplementation, "and-http", vbBinaryCompare)) Then
        GenerateHttpdIcon = 41
    ElseIf (InStrB(1, sImplementation, "anti-web", vbBinaryCompare)) Then
        GenerateHttpdIcon = 51
    ElseIf (InStrB(1, sImplementation, "apache", vbBinaryCompare)) Then
        GenerateHttpdIcon = 2
    ElseIf (InStrB(1, sImplementation, "araneida", vbBinaryCompare)) Then
        GenerateHttpdIcon = 92
    ElseIf (InStrB(1, sImplementation, "axis", vbBinaryCompare)) Then
        GenerateHttpdIcon = 59
    ElseIf (InStrB(1, sImplementation, "badblue", vbBinaryCompare)) Then
        GenerateHttpdIcon = 62
    ElseIf (InStrB(1, sImplementation, "barracuda", vbBinaryCompare)) Then
        GenerateHttpdIcon = 80
    ElseIf (InStrB(1, sImplementation, "basehttp", vbBinaryCompare)) Then
        GenerateHttpdIcon = 92
    ElseIf (InStrB(1, sImplementation, "boa", vbBinaryCompare)) Then
        GenerateHttpdIcon = 82
    ElseIf (InStrB(1, sImplementation, "bea", vbBinaryCompare)) Then
        GenerateHttpdIcon = 3
    ElseIf (InStrB(1, sImplementation, "belkin", vbBinaryCompare)) Then
        GenerateHttpdIcon = 81
    ElseIf (InStrB(1, sImplementation, "bozo", vbBinaryCompare)) Then
        GenerateHttpdIcon = 90
    ElseIf (InStrB(1, sImplementation, "caudium", vbBinaryCompare)) Then
        GenerateHttpdIcon = 31
    ElseIf (InStrB(1, sImplementation, "cherokee", vbBinaryCompare)) Then
        GenerateHttpdIcon = 33
    ElseIf (InStrB(1, sImplementation, "cisco", vbBinaryCompare)) Then
        GenerateHttpdIcon = 4
    ElseIf (InStrB(1, sImplementation, "cl-http", vbBinaryCompare)) Then
        GenerateHttpdIcon = 93
    ElseIf (InStrB(1, sImplementation, "compaq", vbBinaryCompare)) Then
        GenerateHttpdIcon = 5
    ElseIf (InStrB(1, sImplementation, "cougar", vbBinaryCompare)) Then
        GenerateHttpdIcon = 10
    ElseIf (InStrB(1, sImplementation, "dell", vbBinaryCompare)) Then
        GenerateHttpdIcon = 77
    ElseIf (InStrB(1, sImplementation, "divar", vbBinaryCompare)) Then
        GenerateHttpdIcon = 76
    ElseIf (InStrB(1, sImplementation, "dwhttpd", vbBinaryCompare)) Then
        GenerateHttpdIcon = 75
    ElseIf (InStrB(1, sImplementation, "emule", vbBinaryCompare)) Then
        GenerateHttpdIcon = 27
    ElseIf (InStrB(1, sImplementation, "firecat", vbBinaryCompare)) Then
        GenerateHttpdIcon = 42
    ElseIf (InStrB(1, sImplementation, "flexwatch", vbBinaryCompare)) Then
        GenerateHttpdIcon = 82
    ElseIf (InStrB(1, sImplementation, "fnord", vbBinaryCompare)) Then
        GenerateHttpdIcon = 84
    ElseIf (InStrB(1, sImplementation, "gatling", vbBinaryCompare)) Then
        GenerateHttpdIcon = 43
    ElseIf (InStrB(1, sImplementation, "globalscape", vbBinaryCompare)) Then
        GenerateHttpdIcon = 100
    ElseIf (InStrB(1, sImplementation, "google", vbBinaryCompare)) Then
        GenerateHttpdIcon = 34
    ElseIf (InStrB(1, sImplementation, "hp", vbBinaryCompare)) Then
        GenerateHttpdIcon = 7
    ElseIf (InStrB(1, sImplementation, "hiawatha", vbBinaryCompare)) Then
        GenerateHttpdIcon = 44
    ElseIf (InStrB(1, sImplementation, "httpi", vbBinaryCompare)) Then
        GenerateHttpdIcon = 94
    ElseIf (InStrB(1, sImplementation, "ibm", vbBinaryCompare)) Then
        GenerateHttpdIcon = 8
    ElseIf (InStrB(1, sImplementation, "icewarp", vbBinaryCompare)) Then
        GenerateHttpdIcon = 50
    ElseIf (InStrB(1, sImplementation, "indy", vbBinaryCompare)) Then
        GenerateHttpdIcon = 85
    ElseIf (InStrB(1, sImplementation, "iis 4", vbBinaryCompare)) Then
        GenerateHttpdIcon = 9
    ElseIf (InStrB(1, sImplementation, "iis 5", vbBinaryCompare)) Then
        GenerateHttpdIcon = 9
    ElseIf (InStrB(1, sImplementation, "iis ", vbBinaryCompare)) Then
        GenerateHttpdIcon = 10
    ElseIf (InStrB(1, sImplementation, "jana", vbBinaryCompare)) Then
        GenerateHttpdIcon = 11
    ElseIf (InStrB(1, sImplementation, "jetty", vbBinaryCompare)) Then
        GenerateHttpdIcon = 37
    ElseIf (InStrB(1, sImplementation, "jigsaw", vbBinaryCompare)) Then
        GenerateHttpdIcon = 55
    ElseIf (InStrB(1, sImplementation, "lancom", vbBinaryCompare)) Then
        GenerateHttpdIcon = 65
    ElseIf (InStrB(1, sImplementation, "konica", vbBinaryCompare)) Then
        GenerateHttpdIcon = 66
    ElseIf (InStrB(1, sImplementation, "lexmark", vbBinaryCompare)) Then
        GenerateHttpdIcon = 79
    ElseIf (InStrB(1, sImplementation, "lighttpd", vbBinaryCompare)) Then
        GenerateHttpdIcon = 29
    ElseIf (InStrB(1, sImplementation, "linksys", vbBinaryCompare)) Then
        GenerateHttpdIcon = 12
    ElseIf (InStrB(1, sImplementation, "listmanager", vbBinaryCompare)) Then
        GenerateHttpdIcon = 58
    ElseIf (InStrB(1, sImplementation, "litespeed", vbBinaryCompare)) Then
        GenerateHttpdIcon = 49
    ElseIf (InStrB(1, sImplementation, "lotus", vbBinaryCompare)) Then
        GenerateHttpdIcon = 6
    ElseIf (InStrB(1, sImplementation, "mikrotik", vbBinaryCompare)) Then
        GenerateHttpdIcon = 13
    ElseIf (InStrB(1, sImplementation, "mongrel", vbBinaryCompare)) Then
        GenerateHttpdIcon = 86
    ElseIf (InStrB(1, sImplementation, "net2phone", vbBinaryCompare)) Then
        GenerateHttpdIcon = 64
    ElseIf (InStrB(1, sImplementation, "netgear", vbBinaryCompare)) Then
        GenerateHttpdIcon = 35
    ElseIf (InStrB(1, sImplementation, "netopia", vbBinaryCompare)) Then
        GenerateHttpdIcon = 63
    ElseIf (InStrB(1, sImplementation, "netscape", vbBinaryCompare)) Then
        GenerateHttpdIcon = 14
    ElseIf (InStrB(1, sImplementation, "nginx", vbBinaryCompare)) Then
        GenerateHttpdIcon = 38
    ElseIf (InStrB(1, sImplementation, "novell", vbBinaryCompare)) Then
        GenerateHttpdIcon = 15
    ElseIf (InStrB(1, sImplementation, "omnihttpd", vbBinaryCompare)) Then
        GenerateHttpdIcon = 73
    ElseIf (InStrB(1, sImplementation, "oracle", vbBinaryCompare)) Then
        GenerateHttpdIcon = 39
    ElseIf (InStrB(1, sImplementation, "orion", vbBinaryCompare)) Then
        GenerateHttpdIcon = 96
    ElseIf (InStrB(1, sImplementation, "osu", vbBinaryCompare)) Then
        GenerateHttpdIcon = 95
    ElseIf (InStrB(1, sImplementation, "packetshaper", vbBinaryCompare)) Then
        GenerateHttpdIcon = 87
    ElseIf (InStrB(1, sImplementation, "philips", vbBinaryCompare)) Then
        GenerateHttpdIcon = 78
    ElseIf (InStrB(1, sImplementation, "publicfile", vbBinaryCompare)) Then
        GenerateHttpdIcon = 88
    ElseIf (InStrB(1, sImplementation, "qnap", vbBinaryCompare)) Then
        GenerateHttpdIcon = 71
    ElseIf (InStrB(1, sImplementation, "resin", vbBinaryCompare)) Then
        GenerateHttpdIcon = 56
    ElseIf (InStrB(1, sImplementation, "ricoh", vbBinaryCompare)) Then
        GenerateHttpdIcon = 72
    ElseIf (InStrB(1, sImplementation, "roxen", vbBinaryCompare)) Then
        GenerateHttpdIcon = 45
    ElseIf (InStrB(1, sImplementation, "smc", vbBinaryCompare)) Then
        GenerateHttpdIcon = 16
    ElseIf (InStrB(1, sImplementation, "snap", vbBinaryCompare)) Then
        GenerateHttpdIcon = 17
    ElseIf (InStrB(1, sImplementation, "sonicwall", vbBinaryCompare)) Then
        GenerateHttpdIcon = 52
    ElseIf (InStrB(1, sImplementation, "sony", vbBinaryCompare)) Then
        GenerateHttpdIcon = 61
    ElseIf (InStrB(1, sImplementation, "squid", vbBinaryCompare)) Then
        GenerateHttpdIcon = 30
    ElseIf (InStrB(1, sImplementation, "stweb", vbBinaryCompare)) Then
        GenerateHttpdIcon = 97
    ElseIf (InStrB(1, sImplementation, "sun", vbBinaryCompare)) Then
        GenerateHttpdIcon = 18
    ElseIf (InStrB(1, sImplementation, "swat", vbBinaryCompare)) Then
        GenerateHttpdIcon = 28
    ElseIf (InStrB(1, sImplementation, "symantec", vbBinaryCompare)) Then
        GenerateHttpdIcon = 74
    ElseIf (InStrB(1, sImplementation, "tandberg", vbBinaryCompare)) Then
        GenerateHttpdIcon = 89
    ElseIf (InStrB(1, sImplementation, "tcl", vbBinaryCompare)) Then
        GenerateHttpdIcon = 58
    ElseIf (InStrB(1, sImplementation, "thttpd", vbBinaryCompare)) Then
        GenerateHttpdIcon = 19
    ElseIf (InStrB(1, sImplementation, "tomcat", vbBinaryCompare)) Then
        GenerateHttpdIcon = 20
    ElseIf (InStrB(1, sImplementation, "ubicom", vbBinaryCompare)) Then
        GenerateHttpdIcon = 22
    ElseIf (InStrB(1, sImplementation, "userland", vbBinaryCompare)) Then
        GenerateHttpdIcon = 54
    ElseIf (InStrB(1, sImplementation, "virtuoso", vbBinaryCompare)) Then
        GenerateHttpdIcon = 46
    ElseIf (InStrB(1, sImplementation, "vnc", vbBinaryCompare)) Then
        GenerateHttpdIcon = 23
    ElseIf (InStrB(1, sImplementation, "vqserver", vbBinaryCompare)) Then
        GenerateHttpdIcon = 98
    ElseIf (InStrB(1, sImplementation, "vs", vbBinaryCompare)) Then
        GenerateHttpdIcon = 99
    ElseIf (InStrB(1, sImplementation, "wdaemon", vbBinaryCompare)) Then
        GenerateHttpdIcon = 53
    ElseIf (InStrB(1, sImplementation, "webcamxp", vbBinaryCompare)) Then
        GenerateHttpdIcon = 69
    ElseIf (InStrB(1, sImplementation, "wn", vbBinaryCompare)) Then
        GenerateHttpdIcon = 57
    ElseIf (InStrB(1, sImplementation, "webrick", vbBinaryCompare)) Then
        GenerateHttpdIcon = 47
    ElseIf (InStrB(1, sImplementation, "xitami", vbBinaryCompare)) Then
        GenerateHttpdIcon = 32
    ElseIf (InStrB(1, sImplementation, "xserver", vbBinaryCompare)) Then
        GenerateHttpdIcon = 24
    ElseIf (InStrB(1, sImplementation, "yaws", vbBinaryCompare)) Then
        GenerateHttpdIcon = 48
    ElseIf (InStrB(1, sImplementation, "zeus", vbBinaryCompare)) Then
        GenerateHttpdIcon = 25
    ElseIf (InStrB(1, sImplementation, "zope", vbBinaryCompare)) Then
        GenerateHttpdIcon = 26
    ElseIf (InStrB(1, sImplementation, "zyxel", vbBinaryCompare)) Then
        GenerateHttpdIcon = 60
    ElseIf (InStrB(1, sImplementation, "4d", vbBinaryCompare)) Then
        GenerateHttpdIcon = 36
    
' Operating systems collector
    ElseIf (InStrB(1, sImplementation, "bsd", vbBinaryCompare)) Then
        GenerateHttpdIcon = 67
    ElseIf (InStrB(1, sImplementation, "debian", vbBinaryCompare)) Then
        GenerateHttpdIcon = 68
    ElseIf (InStrB(1, sImplementation, "suse", vbBinaryCompare)) Then
        GenerateHttpdIcon = 70
    ElseIf (InStrB(1, sImplementation, "linux", vbBinaryCompare)) Then
        GenerateHttpdIcon = 21
    ElseIf (InStrB(1, sImplementation, "windows", vbBinaryCompare)) Then
        GenerateHttpdIcon = 10
    Else
        GenerateHttpdIcon = 101
    End If
End Function

Public Sub AnnounceFingerprintMatches(ByRef sFullMatchList As String)
    Dim sResultList As String
    Dim sResultArray() As String
    Dim i As Integer
    Dim iResultCount As Integer
    Dim lList As ListItem
    Dim sEntry() As String
    Dim iBestHitter As Integer
    Dim dBestMatch As Double
    
    Call ChangeStatusBar("Preparing Results...")

    sResultList = GenerateMatchStatistics(sFullMatchList)
    sResultArray = Split(sResultList, vbCrLf, , vbBinaryCompare)
    iResultCount = UBound(sResultArray)
    
    For i = 0 To iResultCount
        If (LenB(sResultArray(i))) Then
            sEntry = Split(sResultArray(i), APP_HITPOINTS_DELIMITER, , vbBinaryCompare)
            
            If (scan_besthitcount < sEntry(1)) Then
                scan_besthitname = sEntry(0)
                scan_besthitcount = sEntry(1)
            End If
        End If
    Next i
    If (scan_besthitcount < (app_hitpoints_minimum * tests_count)) Then
        iBestHitter = (app_hitpoints_minimum * tests_count)
    ElseIf (scan_besthitcount > (app_hitpoints_maximum * tests_count)) Then
        iBestHitter = (app_hitpoints_maximum * tests_count)
    Else
        iBestHitter = scan_besthitcount
    End If
    
    frmMain.lsvResults.Visible = False
    frmMain.lsvResults.ListItems.Clear
    For i = 0 To iResultCount
        If (LenB(sResultArray(i))) Then
            sEntry = Split(sResultArray(i), APP_HITPOINTS_DELIMITER, , vbBinaryCompare)
            
            dBestMatch = (100 / iBestHitter * sEntry(1))
            If (dBestMatch > 100) Then
                dBestMatch = 100
            End If
            
            Set lList = frmMain.lsvResults.ListItems.Add(, , vbNullString, , GenerateHttpdIcon(sEntry(0)))
                lList.SubItems(1) = sEntry(0)
                lList.SubItems(2) = sEntry(1)
                lList.SubItems(3) = dBestMatch
        End If
    Next i
    frmMain.lsvResults.Visible = True
    
    Call ListViewSort(frmMain.lsvResults, frmMain.lsvResults.ColumnHeaders(3), 1)
    Call ChangeStatusBarReady
End Sub

Public Function IdentifyGlobalFingerprint(ByRef sFingerprintDirectory As String, ByRef sOriginalResponse As String) As String
    If (LenB(sOriginalResponse)) Then
        Dim cFullMatchList As Concat
    
        Set cFullMatchList = New Concat
        
        Call AddTestCount(sOriginalResponse)
        
        cFullMatchList.Concat FindMatchInDatabase(sFingerprintDirectory & app_file_banner, GetBanner(sOriginalResponse))
        cFullMatchList.Concat FindMatchInDatabase(sFingerprintDirectory & app_file_xpoweredby, GetXPoweredBy(sOriginalResponse))
        cFullMatchList.Concat FindMatchInDatabase(sFingerprintDirectory & app_file_protocolname, GetProtocolName(sOriginalResponse))
        cFullMatchList.Concat FindMatchInDatabase(sFingerprintDirectory & app_file_protocolversion, GetProtocolVersion(sOriginalResponse))
        cFullMatchList.Concat FindMatchInDatabase(sFingerprintDirectory & app_file_statuscode, GetStatusCode(sOriginalResponse))
        cFullMatchList.Concat FindMatchInDatabase(sFingerprintDirectory & app_file_statustext, GetStatusText(sOriginalResponse))
        cFullMatchList.Concat FindMatchInDatabase(sFingerprintDirectory & app_file_headerspace, GetHeaderSpace(sOriginalResponse))
        cFullMatchList.Concat FindMatchInDatabase(sFingerprintDirectory & app_file_headercapitalafterdash, GetHeaderCapitalAfterDash(sOriginalResponse))
        cFullMatchList.Concat FindMatchInDatabase(sFingerprintDirectory & app_file_headerorder, GetHeaderOrder(sOriginalResponse, vbNullString))
        cFullMatchList.Concat FindMatchInDatabase(sFingerprintDirectory & app_file_headerorder, GetHeaderOrder(sOriginalResponse, "X-|Set-Cookie"))
        cFullMatchList.Concat FindMatchInDatabase(sFingerprintDirectory & app_file_optionsallowed, GetOptionsAllowed(sOriginalResponse))
        cFullMatchList.Concat FindMatchInDatabase(sFingerprintDirectory & app_file_optionspublic, GetOptionsPublic(sOriginalResponse))
        cFullMatchList.Concat FindMatchInDatabase(sFingerprintDirectory & app_file_optionsdelimiter, GetOptionsDelimiter(sOriginalResponse))
        cFullMatchList.Concat FindMatchInDatabase(sFingerprintDirectory & app_file_etaglength, GetEtagLength(sOriginalResponse))
        cFullMatchList.Concat FindMatchInDatabase(sFingerprintDirectory & app_file_etagquotes, GetEtagQuotes(sOriginalResponse))
        cFullMatchList.Concat FindMatchInDatabase(sFingerprintDirectory & app_file_contenttype, GetContentType(sOriginalResponse))
        cFullMatchList.Concat FindMatchInDatabase(sFingerprintDirectory & app_file_acceptrange, GetAcceptRange(sOriginalResponse))
        cFullMatchList.Concat FindMatchInDatabase(sFingerprintDirectory & app_file_connection, GetConnection(sOriginalResponse))
        cFullMatchList.Concat FindMatchInDatabase(sFingerprintDirectory & app_file_cachecontrol, GetCacheControl(sOriginalResponse))
        cFullMatchList.Concat FindMatchInDatabase(sFingerprintDirectory & app_file_pragma, GetPragma(sOriginalResponse))
        cFullMatchList.Concat FindMatchInDatabase(sFingerprintDirectory & app_file_varyorder, GetVaryOrder(sOriginalResponse))
        cFullMatchList.Concat FindMatchInDatabase(sFingerprintDirectory & app_file_varycapitalize, GetVaryCapitalized(sOriginalResponse))
        cFullMatchList.Concat FindMatchInDatabase(sFingerprintDirectory & app_file_varydelimiter, GetVaryDelimiter(sOriginalResponse))
        cFullMatchList.Concat FindMatchInDatabase(sFingerprintDirectory & app_file_htaccessrealm, GetHtaccessRealm(sOriginalResponse))
        
        IdentifyGlobalFingerprint = cFullMatchList.Value
    End If
End Function

Public Sub ServerAnalysis()
    Dim bSecure As Boolean

    Call DisableElements
    Call ChangeStatusBar("Starting Analysis...")

    scan_time = Time
    scan_date = Date
    scan_besthitcount = 0
    scan_besthitname = vbNullString

    scan_targethost = frmMain.txtTargetHost.Text
    scan_targetport = frmMain.cboTargetPort.Text
    frmMain.Caption = APP_NAME & " - " & scan_targethost & ":" & scan_targetport
    frmMain.fraTarget.Caption = "Target (unknown)"
    
    Call WriteConfigurationToFile(app_configuration_filename)
    DoEvents

    If (RunTestRequests(scan_targethost, scan_targetport, scan_targetsecure) = True) Then
        Call AnalyzeFingerprintsAndShowResult
    Else
        Call ChangeStatusBar("Target " & scan_targethost & ":" & scan_targetport & " is not a web server. Aborting.")
'        MsgBox "Target " & scan_targethost & ":" & scan_targetport & " is not a web server." & vbCrLf & _
'            "Please check your settings.", vbExclamation + vbOKOnly, "No web server found"
    End If

    Call EnableElements
End Sub

Public Sub AnalyzeFingerprintsAndShowResult()
    Dim cFullIdentifyList As Concat

    Set cFullIdentifyList = New Concat
    
    tests_count = 0
    
    cFullIdentifyList.Concat IdentifyGlobalFingerprint(app_dir_attackrequest, response_attackrequest)
    cFullIdentifyList.Concat IdentifyGlobalFingerprint(app_dir_deleteexisting, response_delete)
    cFullIdentifyList.Concat IdentifyGlobalFingerprint(app_dir_getexisting, response_getexist)
    cFullIdentifyList.Concat IdentifyGlobalFingerprint(app_dir_getlong, response_getlongrequest)
    cFullIdentifyList.Concat IdentifyGlobalFingerprint(app_dir_getnonexisting, response_get_nonexistent)
    cFullIdentifyList.Concat IdentifyGlobalFingerprint(app_dir_headexisting, response_head)
    cFullIdentifyList.Concat IdentifyGlobalFingerprint(app_dir_options, response_options)
    cFullIdentifyList.Concat IdentifyGlobalFingerprint(app_dir_wrongmethod, response_testmethod)
    cFullIdentifyList.Concat IdentifyGlobalFingerprint(app_dir_wrongversion, response_protocolversion)
    
    Call FillResponses
    Call AnnounceFingerprintMatches(cFullIdentifyList.Value)
    frmMain.fraTarget.Caption = "Target (" & scan_besthitname & ")"
    frmMain.tbsResults.Tabs(1).Caption = "Matchlist (" & frmMain.lsvResults.ListItems.Count & " implementations)"
    frmMain.mnuFileSaveAsScanItem.Enabled = True
    frmMain.mnuFingerprintingSaveFingerprintItem.Enabled = True
    frmMain.mnuFingerprintingReanalyzeItem.Enabled = True
    frmMain.mnuReportingGenerateReportItem.Enabled = True
End Sub

Public Sub FillResponses()
    Dim iIndex As Integer
    
    iIndex = frmMain.tbsViews.SelectedItem.Index

    If (iIndex = 1) Then
        Call ShowResponseData(response_getexist)
    ElseIf (iIndex = 2) Then
        Call ShowResponseData(response_getlongrequest)
    ElseIf (iIndex = 3) Then
        Call ShowResponseData(response_get_nonexistent)
    ElseIf (iIndex = 4) Then
        Call ShowResponseData(response_protocolversion)
    ElseIf (iIndex = 5) Then
        Call ShowResponseData(response_head)
    ElseIf (iIndex = 6) Then
        Call ShowResponseData(response_options)
    ElseIf (iIndex = 7) Then
        Call ShowResponseData(response_delete)
    ElseIf (iIndex = 8) Then
        Call ShowResponseData(response_testmethod)
    ElseIf (iIndex = 9) Then
        Call ShowResponseData(response_attackrequest)
    End If
End Sub

Public Sub ShowResponseData(ByRef sResponse As String)
    Dim iBannerStart As Integer

    frmMain.txtResponses.Text = sResponse
    frmMain.txtResponses.ToolTipText = Len(sResponse) & " bytes"
    
    iBannerStart = InStr(1, sResponse, GetBanner(sResponse), vbBinaryCompare)
    If (iBannerStart > 1) Then
        frmMain.txtResponses.SelStart = InStr(1, sResponse, GetBanner(sResponse), vbBinaryCompare) - 1
        frmMain.txtResponses.SelLength = Len(GetBanner(sResponse))
    End If
    frmMain.txtFingerprint.Text = GenerateFingerprintDetails(sResponse)
End Sub

Public Function GenerateFingerprintDetails(ByRef sOriginalResponse As String) As String
    Dim cFingerprintDetails As Concat

    If (LenB(sOriginalResponse)) Then
        Set cFingerprintDetails = New Concat
        
        cFingerprintDetails.Concat "Protocol Name" & vbTab & vbTab & GetProtocolName(sOriginalResponse) & vbCrLf
        cFingerprintDetails.Concat "Protocol Version" & vbTab & GetProtocolVersion(sOriginalResponse) & vbCrLf
        cFingerprintDetails.Concat "Statuscode" & vbTab & vbTab & GetStatusCode(sOriginalResponse) & vbCrLf
        cFingerprintDetails.Concat "Statustext" & vbTab & vbTab & GetStatusText(sOriginalResponse) & vbCrLf
        cFingerprintDetails.Concat "Banner" & vbTab & vbTab & vbTab & GetBanner(sOriginalResponse) & vbCrLf
        cFingerprintDetails.Concat "X-Powered-By" & vbTab & vbTab & GetXPoweredBy(sOriginalResponse) & vbCrLf
        cFingerprintDetails.Concat "Header Spaces" & vbTab & vbTab & GetHeaderSpace(sOriginalResponse) & vbCrLf
        cFingerprintDetails.Concat "Capital after Dash" & vbTab & GetHeaderCapitalAfterDash(sOriginalResponse) & vbCrLf
        cFingerprintDetails.Concat "Header-Order Full" & vbTab & GetHeaderOrder(sOriginalResponse, "") & vbCrLf
        cFingerprintDetails.Concat "Header-Order Limit" & vbTab & GetHeaderOrder(sOriginalResponse, "X-|Set-Cookie") & vbCrLf
        cFingerprintDetails.Concat "Options-Allowed" & vbTab & vbTab & GetOptionsAllowed(sOriginalResponse) & vbCrLf
        cFingerprintDetails.Concat "Options-Public" & vbTab & vbTab & GetOptionsPublic(sOriginalResponse) & vbCrLf
        cFingerprintDetails.Concat "Options-Delimiter" & vbTab & GetOptionsDelimiter(sOriginalResponse) & vbCrLf
        cFingerprintDetails.Concat "ETag" & vbTab & vbTab & vbTab & GetEtag(sOriginalResponse) & vbCrLf
        cFingerprintDetails.Concat "ETag-Length" & vbTab & vbTab & GetEtagLength(sOriginalResponse) & vbCrLf
        cFingerprintDetails.Concat "ETag-Quotes" & vbTab & vbTab & GetEtagQuotes(sOriginalResponse) & vbCrLf
        cFingerprintDetails.Concat "Content-Type" & vbTab & vbTab & GetContentType(sOriginalResponse) & vbCrLf
        cFingerprintDetails.Concat "Accept-Range" & vbTab & vbTab & GetAcceptRange(sOriginalResponse) & vbCrLf
        cFingerprintDetails.Concat "Connection" & vbTab & vbTab & GetConnection(sOriginalResponse) & vbCrLf
        cFingerprintDetails.Concat "Cache-Control" & vbTab & vbTab & GetCacheControl(sOriginalResponse) & vbCrLf
        cFingerprintDetails.Concat "Pragma" & vbTab & vbTab & vbTab & GetPragma(sOriginalResponse) & vbCrLf
        cFingerprintDetails.Concat "Vary-Order" & vbTab & vbTab & GetVaryOrder(sOriginalResponse) & vbCrLf
        cFingerprintDetails.Concat "Vary-Capitalized" & vbTab & GetVaryCapitalized(sOriginalResponse) & vbCrLf
        cFingerprintDetails.Concat "Vary-Delimiter" & vbTab & vbTab & GetVaryDelimiter(sOriginalResponse) & vbCrLf
        cFingerprintDetails.Concat "htaccess-Realm" & vbTab & vbTab & GetHtaccessRealm(sOriginalResponse) & vbCrLf
       
        GenerateFingerprintDetails = cFingerprintDetails.Value
    End If
End Function

Public Function PreFetchBanner(ByRef sRequest As String) As String
    Dim sBanner As String
    
    sBanner = GetHeaderValue(response_getexist, "Server", True)
    
    If (LenB(sBanner)) Then
        PreFetchBanner = sBanner
    Else
        PreFetchBanner = "no banner available"
    End If
End Function

Public Sub ResetAll()
    frmMain.Caption = APP_NAME
    
    scan_besthitcount = 0
    scan_besthitname = vbNullString
    
    scan_time = vbNullString
    scan_date = vbNullString
    scan_targethost = "127.0.0.1"
    scan_targetport = 80
    Call ChangeSSLMode(False)
    
    frmMain.fraTarget.Caption = "Target"
    
    frmMain.txtTargetHost = scan_targethost
    frmMain.cboTargetPort = scan_targetport
    frmMain.cboScheme.ListIndex = 0
    
    response_attackrequest = vbNullString
    response_delete = vbNullString
    response_getexist = vbNullString
    response_getlongrequest = vbNullString
    response_get_nonexistent = vbNullString
    response_head = vbNullString
    response_options = vbNullString
    response_testmethod = vbNullString
    response_protocolversion = vbNullString
    
    frmMain.lsvResults.ListItems.Clear
    frmMain.txtResponses.Text = vbNullString
    frmMain.txtResponses.ToolTipText = vbNullString
    frmMain.txtFingerprint.Text = vbNullString
    
    frmMain.mnuFileSaveAsScanItem.Enabled = False
    frmMain.mnuFingerprintingReanalyzeItem.Enabled = False
    frmMain.mnuFingerprintingSaveFingerprintItem.Enabled = False
    frmMain.mnuReportingGenerateReportItem.Enabled = False
End Sub

Public Sub ChangeSSLMode(ByRef bSecure As Boolean)
    If (bSecure = False) Then
        frmMain.cboScheme.ListIndex = 0
        scan_targetsecure = 0
    Else
        frmMain.cboScheme.ListIndex = 1
        scan_targetsecure = 1
    End If
End Sub

Public Sub DisableElements()
    frmMain.txtResponses.SetFocus
    frmMain.txtResponses.Text = vbNullString
    frmMain.txtFingerprint.Text = vbNullString
    frmMain.lsvResults.ListItems.Clear
    frmMain.cmdAnalyze.Enabled = False
    frmMain.mnuFileNewItem.Enabled = False
    frmMain.mnuFileOpenScanlistItem.Enabled = False
    frmMain.mnuFileOpenScanItem.Enabled = False
    frmMain.mnuFileSaveAsScanItem.Enabled = False
    frmMain.mnuConfigurationEditItem.Enabled = False
    frmMain.mnuFingerprintingAnalyzeItem.Enabled = False
    frmMain.mnuFingerprintingReanalyzeItem.Enabled = False
    frmMain.mnuFingerprintingSaveFingerprintItem.Enabled = False
    frmMain.txtTargetHost.Enabled = False
    frmMain.cboTargetPort.Enabled = False
    frmMain.cboScheme.Enabled = False
    Screen.MousePointer = vbArrowHourglass
End Sub

Public Sub EnableElements()
    frmMain.cmdAnalyze.Enabled = True
    frmMain.mnuFileNewItem.Enabled = True
    frmMain.mnuFileOpenScanlistItem.Enabled = True
    frmMain.mnuFileOpenScanItem.Enabled = True
    frmMain.mnuFileSaveAsScanItem.Enabled = True
    frmMain.mnuConfigurationEditItem.Enabled = True
    frmMain.mnuFingerprintingAnalyzeItem.Enabled = True
    frmMain.mnuFingerprintingReanalyzeItem.Enabled = True
    frmMain.mnuFingerprintingSaveFingerprintItem.Enabled = True
    frmMain.txtTargetHost.Enabled = True
    frmMain.cboTargetPort.Enabled = True
    frmMain.cboScheme.Enabled = True
    Screen.MousePointer = vbNormal
End Sub

