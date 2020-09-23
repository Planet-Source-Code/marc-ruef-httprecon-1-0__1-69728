Attribute VB_Name = "modReporting"
Option Explicit

Public Function GenerateHtmlReport() As String
    Dim cReport As Concat

    Set cReport = New Concat

    cReport.Concat "<?xml version='1.0' encoding='iso-8859-1' ?><!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.1//EN"" ""http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd""> " & vbCrLf
    cReport.Concat "<html xmlns=""http://www.w3.org/1999/xhtml"">"
    cReport.Concat "<head>"
    cReport.Concat "<title>" & APP_NAME & " Report</title>"
    cReport.Concat GetEmbeddedCSS()
    cReport.Concat "</head>"
    cReport.Concat "<body>"
    
    cReport.Concat "<h3 id='top'>" & APP_NAME & " Report </h3>"
    cReport.Concat "Target of Scan: " & HtmlEncode(frmMain.txtTargetHost.Text) & ":" & frmMain.cboTargetPort.Text & "<br />"
    cReport.Concat "Date of Export: " & Date & "<br />"
    
    cReport.Concat "<h4 id='contents'>Contents</h4>"
    cReport.Concat "<ol style'list-style-type:decimal'>"
    cReport.Concat "<li><a href='#matches'>Matches</a></li>"
    cReport.Concat "<li><a href='#responses'>Responses</a></li>"
    cReport.Concat "</ol>"
    
    cReport.Concat "<h4 id='matches'>List of Matches <a href='#top'>&uarr;</a></h4>"
    cReport.Concat GenerateHitList(20)
    
    cReport.Concat "<h4 id='responses'>HTTP Response Header <a href='#top'>&uarr;</a></h4>"
    cReport.Concat ShowTestCase(APP_TESTNAME_GETEXISTING, response_getexist)
    cReport.Concat ShowTestCase(APP_TESTNAME_GETLONG, response_getlongrequest)
    cReport.Concat ShowTestCase(APP_TESTNAME_GETNONEXISTING, response_get_nonexistent)
    cReport.Concat ShowTestCase(APP_TESTNAME_HEADEXISTING, response_head)
    cReport.Concat ShowTestCase(APP_TESTNAME_OPTIONS, response_options)
    cReport.Concat ShowTestCase(APP_TESTNAME_DELETEEXISTING, response_delete)
    cReport.Concat ShowTestCase(APP_TESTNAME_WRONGMETHOD, response_testmethod)
    cReport.Concat ShowTestCase(APP_TESTNAME_WRONGVERSION, response_protocolversion)
    cReport.Concat ShowTestCase(APP_TESTNAME_ATTACKREQUEST, response_attackrequest)

    cReport.Concat "</body>"
    cReport.Concat "</html>"

    GenerateHtmlReport = cReport.Value
End Function

Public Function ShowTestCase(ByRef sName As String, ByRef sResponse As String) As String
    Dim cTestcase As Concat
    Dim iLength As Integer
    
    Set cTestcase = New Concat
    
    iLength = Len(sResponse)
    
    cTestcase.Concat "<table class='table'>"
    cTestcase.Concat "<tr class='title'><td>" & HtmlEncode(sName) & "</td><tr>"
    If (iLength) Then
        cTestcase.Concat "<tr><td class='response' title='Length: " & iLength & " bytes'>" & HtmlEncode(sResponse) & "</td><tr>"
    Else
        cTestcase.Concat "<tr class='databaseline'><td class='databaseline'>no response available</td><tr>"
    End If
    cTestcase.Concat "</table><br />"
    
    ShowTestCase = cTestcase.Value
End Function

Public Function GenerateHitList(ByRef iCount As Integer) As String
    Dim cResults As Concat
    Dim iListItemCount As Integer
    Dim i As Integer
    
    Set cResults = New Concat
    
    iListItemCount = frmMain.lsvResults.ListItems.Count
    
    If (iListItemCount > iCount) Then
        iListItemCount = iCount
    End If
    
    cResults.Concat "<table class='table'><tr class='title'><td style='width:20px'>&nbsp;</td><td>Name</td><td>Hits</td><td>Match</td></tr>"
    For i = 1 To iListItemCount
         cResults.Concat "<tr class='databaseline'><td style='text-align:right' class='databaseline'>" & i & ".</td><td class='databaseline'>" & HtmlEncode(frmMain.lsvResults.ListItems(i).ListSubItems(1).Text) & "</td><td class='databaseline'>" & HtmlEncode(frmMain.lsvResults.ListItems(i).ListSubItems(2).Text) & "</td><td class='databaseline'>" & Round(frmMain.lsvResults.ListItems(i).ListSubItems(3).Text, 2) & "% </td></tr>"
    Next i
    cResults.Concat "</table><br />"
    
    GenerateHitList = cResults.Value
End Function

Public Function HtmlEncode(ByRef sInput As String) As String
    Dim sOutput As String
    
    sOutput = Replace(sInput, "<", "&gt;", 1, , vbBinaryCompare)
    sOutput = Replace(sOutput, ">", "&lt;", 1, , vbBinaryCompare)
    sOutput = Replace(sOutput, Chr(34), "&quot;", 1, , vbBinaryCompare)
    sOutput = Replace(sOutput, "&", "&amp;", 1, , vbBinaryCompare)
    
    sOutput = Replace(sOutput, vbCrLf, "<br />", 1, , vbBinaryCompare)
    
    HtmlEncode = sOutput
End Function

Public Function GetEmbeddedCSS() As String
    Dim cCSS As Concat
    
    Set cCSS = New Concat

    cCSS.Concat "<style type=""text/css"">"
    cCSS.Concat "<!-- "
    
    cCSS.Concat "body{"
    cCSS.Concat "font-family:verdana;"
    cCSS.Concat "font-size:11px;"
    cCSS.Concat "color:black;"
    cCSS.Concat "}"
    
    cCSS.Concat "a{"
    cCSS.Concat "color:darkred;"
    cCSS.Concat "text-decoration:none;"
    cCSS.Concat "}"

    cCSS.Concat "a:hover{"
    cCSS.Concat "color:red;"
    cCSS.Concat "}"

    cCSS.Concat "table.table{"
    cCSS.Concat "border:1px solid darkgrey;"
    cCSS.Concat "width:640px;"
    cCSS.Concat "}"

    cCSS.Concat "tr.title{"
    cCSS.Concat "font-weight:bold;"
    cCSS.Concat "background:lightgrey;"
    cCSS.Concat "}"
        
    cCSS.Concat "tr.databaseline:hover{"
    cCSS.Concat "background-color:lightgrey;"
    cCSS.Concat "}"

    cCSS.Concat "td.databaseline{"
    cCSS.Concat "border:1px solid lightgrey;"
    cCSS.Concat "}"

    cCSS.Concat "td.response{"
    cCSS.Concat "font-family:'courier new';"
    cCSS.Concat "color:lightgreen;"
    cCSS.Concat "background:black;"
    cCSS.Concat "}"

    cCSS.Concat " -->"
    cCSS.Concat "</style>"
    
    GetEmbeddedCSS = cCSS.Value
End Function
