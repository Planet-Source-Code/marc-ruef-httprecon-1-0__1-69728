Attribute VB_Name = "modIdentification"
Option Explicit

Private Const APP_HITPOINTS_MINIMUM As Integer = 70
Private Const APP_HITPOINTS_MAXIMUM As Integer = 120
Private Const APP_HITPOINTS_DELIMITER As String = ":"

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
    ElseIf (InStrB(1, sImplementation, "and-http", vbBinaryCompare)) Then
        GenerateHttpdIcon = 41
    ElseIf (InStrB(1, sImplementation, "anti-web", vbBinaryCompare)) Then
        GenerateHttpdIcon = 51
    ElseIf (InStrB(1, sImplementation, "apache", vbBinaryCompare)) Then
        GenerateHttpdIcon = 2
    ElseIf (InStrB(1, sImplementation, "axis", vbBinaryCompare)) Then
        GenerateHttpdIcon = 59
    ElseIf (InStrB(1, sImplementation, "badblue", vbBinaryCompare)) Then
        GenerateHttpdIcon = 62
    ElseIf (InStrB(1, sImplementation, "bea", vbBinaryCompare)) Then
        GenerateHttpdIcon = 3
    ElseIf (InStrB(1, sImplementation, "caudium", vbBinaryCompare)) Then
        GenerateHttpdIcon = 31
    ElseIf (InStrB(1, sImplementation, "cherokee", vbBinaryCompare)) Then
        GenerateHttpdIcon = 33
    ElseIf (InStrB(1, sImplementation, "cisco", vbBinaryCompare)) Then
        GenerateHttpdIcon = 4
    ElseIf (InStrB(1, sImplementation, "compaq", vbBinaryCompare)) Then
        GenerateHttpdIcon = 5
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
    ElseIf (InStrB(1, sImplementation, "gatling", vbBinaryCompare)) Then
        GenerateHttpdIcon = 43
    ElseIf (InStrB(1, sImplementation, "google", vbBinaryCompare)) Then
        GenerateHttpdIcon = 34
    ElseIf (InStrB(1, sImplementation, "hp", vbBinaryCompare)) Then
        GenerateHttpdIcon = 7
    ElseIf (InStrB(1, sImplementation, "hiawatha", vbBinaryCompare)) Then
        GenerateHttpdIcon = 44
    ElseIf (InStrB(1, sImplementation, "ibm", vbBinaryCompare)) Then
        GenerateHttpdIcon = 8
    ElseIf (InStrB(1, sImplementation, "icewarp", vbBinaryCompare)) Then
        GenerateHttpdIcon = 50
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
    ElseIf (InStrB(1, sImplementation, "litespeed", vbBinaryCompare)) Then
        GenerateHttpdIcon = 49
    ElseIf (InStrB(1, sImplementation, "lotus", vbBinaryCompare)) Then
        GenerateHttpdIcon = 6
    ElseIf (InStrB(1, sImplementation, "mikrotik", vbBinaryCompare)) Then
        GenerateHttpdIcon = 13
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
    ElseIf (InStrB(1, sImplementation, "philips", vbBinaryCompare)) Then
        GenerateHttpdIcon = 78
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
    ElseIf (InStrB(1, sImplementation, "sun", vbBinaryCompare)) Then
        GenerateHttpdIcon = 18
    ElseIf (InStrB(1, sImplementation, "swat", vbBinaryCompare)) Then
        GenerateHttpdIcon = 28
    ElseIf (InStrB(1, sImplementation, "symantec", vbBinaryCompare)) Then
        GenerateHttpdIcon = 74
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
    ElseIf (InStrB(1, sImplementation, "wdaemon", vbBinaryCompare)) Then
        GenerateHttpdIcon = 53
    ElseIf (InStrB(1, sImplementation, "webcamxp", vbBinaryCompare)) Then
        GenerateHttpdIcon = 69
    ElseIf (InStrB(1, sImplementation, "wn", vbBinaryCompare)) Then
        GenerateHttpdIcon = 57
    ElseIf (InStrB(1, sImplementation, "webrick", vbBinaryCompare)) Then
        GenerateHttpdIcon = 47
    ElseIf (InStrB(1, sImplementation, "virtuoso", vbBinaryCompare)) Then
        GenerateHttpdIcon = 46
    ElseIf (InStrB(1, sImplementation, "vnc", vbBinaryCompare)) Then
        GenerateHttpdIcon = 23
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
    Else
        GenerateHttpdIcon = 80
    End If
End Function

Public Sub AnnounceFingerprintMatches(ByRef sFullMatchList As String)
    Dim sResultList As String
    Dim sResultArray() As String
    Dim i As Integer
    Dim iResultCount As Integer
    Dim lList As ListItem
    Dim sEntry() As String
    Dim sPercentage As Double
    Dim iHighestHit As Integer

    sResultList = GenerateMatchStatistics(sFullMatchList)
    sResultArray = Split(sResultList, vbCrLf, , vbBinaryCompare)
    iResultCount = UBound(sResultArray)
    
    frmMain.lsvResults.ListItems.Clear
    
    For i = 0 To iResultCount
        If (LenB(sResultArray(i))) Then
            sEntry = Split(sResultArray(i), APP_HITPOINTS_DELIMITER, , vbBinaryCompare)
            
            If (iHighestHit < sEntry(1)) Then
                iHighestHit = sEntry(1)
            End If
        End If
    Next i
    If (iHighestHit < APP_HITPOINTS_MINIMUM) Then
        iHighestHit = APP_HITPOINTS_MINIMUM
    ElseIf (iHighestHit > APP_HITPOINTS_MAXIMUM) Then
        iHighestHit = APP_HITPOINTS_MAXIMUM
    End If
    
    For i = 0 To iResultCount
        If (LenB(sResultArray(i))) Then
            sEntry = Split(sResultArray(i), APP_HITPOINTS_DELIMITER, , vbBinaryCompare)
            
            sPercentage = (100 / iHighestHit * sEntry(1))
            If (sPercentage > 100) Then
                sPercentage = 100
            End If
            
            Set lList = frmMain.lsvResults.ListItems.Add(, , vbNullString, , GenerateHttpdIcon(sEntry(0)))
                lList.SubItems(1) = sEntry(0)
                lList.SubItems(2) = sEntry(1)
                lList.SubItems(3) = sPercentage
        End If
    Next i

    Call ListViewSort(frmMain.lsvResults, frmMain.lsvResults.ColumnHeaders(3), 1)
End Sub

Public Function IdentifyGlobalFingerprint(ByRef sFingerprintDirectory As String, ByRef sOriginalResponse As String) As String
    If (LenB(sOriginalResponse)) Then
        Dim cFullMatchList As Concat
    
        Set cFullMatchList = New Concat
        
        cFullMatchList.Concat FindMatchInDatabase(sFingerprintDirectory & app_file_banner, GetBanner(sOriginalResponse))
        cFullMatchList.Concat FindMatchInDatabase(sFingerprintDirectory & app_file_xpoweredby, GetXPoweredBy(sOriginalResponse))
        cFullMatchList.Concat FindMatchInDatabase(sFingerprintDirectory & app_file_protocolname, GetProtocolName(sOriginalResponse))
        cFullMatchList.Concat FindMatchInDatabase(sFingerprintDirectory & app_file_protocolversion, GetProtocolVersion(sOriginalResponse))
        cFullMatchList.Concat FindMatchInDatabase(sFingerprintDirectory & app_file_statuscode, GetStatusCode(sOriginalResponse))
        cFullMatchList.Concat FindMatchInDatabase(sFingerprintDirectory & app_file_statustext, GetStatusText(sOriginalResponse))
        cFullMatchList.Concat FindMatchInDatabase(sFingerprintDirectory & app_file_headerspace, GetHeaderSpace(sOriginalResponse))
        cFullMatchList.Concat FindMatchInDatabase(sFingerprintDirectory & app_file_headercapitalafterdash, GetHeaderCapitalAfterDash(sOriginalResponse))
        cFullMatchList.Concat FindMatchInDatabase(sFingerprintDirectory & app_file_headerorder, GetHeaderOrder(sOriginalResponse))
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
    Dim sTargetHost As String
    Dim lTargetPort As Long

    frmMain.txtResponses.SetFocus
    frmMain.txtResponses.Text = ""
    frmMain.lsvResults.ListItems.Clear
    frmMain.cmdAnalyze.Enabled = False
    frmMain.mnuFingerprinting.Enabled = False
    frmMain.txtTargetHost.Enabled = False
    frmMain.cboTargetPort.Enabled = False
    DoEvents

    sTargetHost = frmMain.txtTargetHost.Text
    lTargetPort = frmMain.cboTargetPort.Text
    frmMain.Caption = APP_NAME & " - " & sTargetHost & ":" & lTargetPort

    If (RunTestRequests(sTargetHost, lTargetPort) = True) Then
        Call AnalyzeFingerprintsAndShowResult
    Else
        MsgBox "Target " & sTargetHost & ":" & lTargetPort & " is not a web server." & vbCrLf & _
            "Please check your settings.", vbExclamation + vbOKOnly, "No web server found"
    End If

    frmMain.cmdAnalyze.Enabled = True
    frmMain.mnuFingerprinting.Enabled = True
    frmMain.txtTargetHost.Enabled = True
    frmMain.cboTargetPort.Enabled = True
End Sub

Public Sub AnalyzeFingerprintsAndShowResult()
    Dim cFullIdentifyList As Concat

    Set cFullIdentifyList = New Concat
    
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
    frmMain.fraResults.Caption = "Results (" & frmMain.lsvResults.ListItems.Count & " implementations known)"
    frmMain.mnuFileSaveAsItem.Enabled = True
    frmMain.mnuFingerprintingSaveFingerprintItem.Enabled = True
    frmMain.mnuReportingGenerateReportItem.Enabled = True
End Sub

Public Sub ResetAll()
    frmMain.Caption = APP_NAME
    
    frmMain.txtTargetHost = "unknown"
    frmMain.cboTargetPort = 80
    
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
    
    frmMain.mnuFileSaveAsItem.Enabled = False
    frmMain.mnuFingerprintingSaveFingerprintItem.Enabled = False
    frmMain.mnuReportingGenerateReportItem.Enabled = False
End Sub
