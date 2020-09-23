Attribute VB_Name = "modFileHandling"
Option Explicit

Public Function GenerateFingerprintXML() As String
    Dim sFullFingerprint As Concat
        
    Set sFullFingerprint = New Concat

    sFullFingerprint.Concat "<" & APP_TESTNAME_GETEXISTING & ">" & vbCrLf & response_getexist & "</" & APP_TESTNAME_GETEXISTING & ">" & vbCrLf
    sFullFingerprint.Concat "<" & APP_TESTNAME_GETLONG & ">" & vbCrLf & response_getlongrequest & "</" & APP_TESTNAME_GETLONG & ">" & vbCrLf
    sFullFingerprint.Concat "<" & APP_TESTNAME_GETNONEXISTING & ">" & vbCrLf & response_get_nonexistent & "</" & APP_TESTNAME_GETNONEXISTING & ">" & vbCrLf
    sFullFingerprint.Concat "<" & APP_TESTNAME_WRONGVERSION & ">" & vbCrLf & response_protocolversion & "</" & APP_TESTNAME_WRONGVERSION & ">" & vbCrLf
    sFullFingerprint.Concat "<" & APP_TESTNAME_HEADEXISTING & ">" & vbCrLf & response_head & "</" & APP_TESTNAME_HEADEXISTING & ">" & vbCrLf
    sFullFingerprint.Concat "<" & APP_TESTNAME_OPTIONS & ">" & vbCrLf & response_options & "</" & APP_TESTNAME_OPTIONS & ">" & vbCrLf
    sFullFingerprint.Concat "<" & APP_TESTNAME_DELETEEXISTING & ">" & vbCrLf & response_delete & "</" & APP_TESTNAME_DELETEEXISTING & ">" & vbCrLf
    sFullFingerprint.Concat "<" & APP_TESTNAME_WRONGMETHOD & ">" & vbCrLf & response_testmethod & "</" & APP_TESTNAME_WRONGMETHOD & ">" & vbCrLf
    sFullFingerprint.Concat "<" & APP_TESTNAME_ATTACKREQUEST & ">" & vbCrLf & response_attackrequest & "</" & APP_TESTNAME_ATTACKREQUEST & ">" & vbCrLf
    
    GenerateFingerprintXML = sFullFingerprint.Value
End Function

Public Sub ReadFingerprintXML(ByRef sFingerprints As String)
    response_attackrequest = ExtractFingerprintXML(sFingerprints, APP_TESTNAME_ATTACKREQUEST)
    response_delete = ExtractFingerprintXML(sFingerprints, APP_TESTNAME_DELETEEXISTING)
    response_getexist = ExtractFingerprintXML(sFingerprints, APP_TESTNAME_GETEXISTING)
    response_getlongrequest = ExtractFingerprintXML(sFingerprints, APP_TESTNAME_GETLONG)
    response_get_nonexistent = ExtractFingerprintXML(sFingerprints, APP_TESTNAME_GETNONEXISTING)
    response_head = ExtractFingerprintXML(sFingerprints, APP_TESTNAME_HEADEXISTING)
    response_options = ExtractFingerprintXML(sFingerprints, APP_TESTNAME_OPTIONS)
    response_testmethod = ExtractFingerprintXML(sFingerprints, APP_TESTNAME_WRONGMETHOD)
    response_protocolversion = ExtractFingerprintXML(sFingerprints, APP_TESTNAME_WRONGVERSION)
End Sub

Public Function ExtractFingerprintXML(ByRef sFingerprints As String, ByRef sTag As String) As String
    Dim iStart As Integer
    Dim iEnd As Integer
    Dim iContentStart As Integer
    Dim iContentEnd As Integer
    Dim iContentLength As Integer
    
    iStart = InStr(1, sFingerprints, "<" & sTag & ">", vbBinaryCompare)
    
    If (iStart > 0) Then
        iContentStart = iStart + (Len(sTag) + 4)
        
        If (iContentStart > 0) Then
            iEnd = InStr(1, sFingerprints, "</" & sTag & ">", vbBinaryCompare)
            
            If (iEnd > iStart) Then
                iContentLength = iEnd - iContentStart
                If (iContentLength > 0) Then
                    ExtractFingerprintXML = Mid$(sFingerprints, iContentStart, iContentLength)
                End If
            End If
        End If
    End If
End Function

