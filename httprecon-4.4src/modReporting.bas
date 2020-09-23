Attribute VB_Name = "modReporting"
Option Explicit

Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public report_directory_reports As String

Public Function GenerateHtmlReport() As String
    Dim cReport As Concat

    Set cReport = New Concat

    Call ChangeStatusBar("Generate HTML Report...")

    cReport.Concat "<?xml version='1.0' encoding='iso-8859-1' ?><!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.1//EN"" ""http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd""> " & vbCrLf
    cReport.Concat "<html xmlns=""http://www.w3.org/1999/xhtml"">" & vbCrLf
    cReport.Concat "<head>" & vbCrLf
    cReport.Concat "<title>" & APP_NAME & " Report (" & HtmlEncode(scan_targethost) & ":" & scan_targetport & ")</title>" & vbCrLf
    cReport.Concat GetEmbeddedCSS()
    cReport.Concat "<meta name='keywords' content='httprecon, Webserver, web server, HTTPD, http, Fingerprinting, Report' />" & vbCrLf
    cReport.Concat "</head>" & vbCrLf
    cReport.Concat "<body>" & vbCrLf
    
    cReport.Concat "<h3>" & APP_NAME & " Report</h3>" & vbCrLf
    cReport.Concat "Target: <a href='http://" & HtmlEncode(scan_targethost) & ":" & scan_targetport & "'>" & HtmlEncode(scan_targethost) & ":" & scan_targetport & "</a> (" & tests_count & " test cases)<br />" & vbCrLf
    cReport.Concat "Auditor: " & GetLocalUsername & "<br />" & vbCrLf
    cReport.Concat "Scan: " & scan_date & " - " & scan_time & "<br />" & vbCrLf
    cReport.Concat "Export: " & Date & " - " & Time & vbCrLf
    
    cReport.Concat "<h4 id='contents'>Contents</h4>" & vbCrLf
    cReport.Concat "<ol style='list-style-type:decimal'>" & vbCrLf
    cReport.Concat "<li><a href='#summary'>Summary</a></li>" & vbCrLf
    cReport.Concat "<li><a href='#matches'>Matches</a></li>" & vbCrLf
    cReport.Concat "<li><a href='#responses'>Responses</a></li>" & vbCrLf
    cReport.Concat "<li><a href='#details'>Details</a></li>" & vbCrLf
    cReport.Concat "</ol>" & vbCrLf
    
    cReport.Concat "<h4 id='summary'>Summary <a href='#'>&uarr;</a></h4>" & vbCrLf
    cReport.Concat "An advanced web server fingerprinting for the host " & HtmlEncode(scan_targethost) & " and port tcp/" & scan_targetport & " was done with " & tests_count & " test cases at " & scan_date & " " & scan_time & ".<br /><br />" & vbCrLf
    cReport.Concat "This analysis was able to determine the target httpd service as " & HtmlEncode(scan_besthitname) & " with " & scan_besthitcount & " fingerprint hits in the database." & vbCrLf
    
    cReport.Concat "<h4 id='matches'>List of Matches <a href='#'>&uarr;</a></h4>" & vbCrLf
    cReport.Concat GenerateHitList(20)
    
    cReport.Concat "<h4 id='responses'>HTTP Response Header <a href='#'>&uarr;</a></h4>" & vbCrLf
    cReport.Concat ShowTestCase(APP_TESTNAME_GETEXISTING, response_getexist) & vbCrLf
    cReport.Concat ShowTestCase(APP_TESTNAME_GETLONG, response_getlongrequest) & vbCrLf
    cReport.Concat ShowTestCase(APP_TESTNAME_GETNONEXISTING, response_get_nonexistent) & vbCrLf
    cReport.Concat ShowTestCase(APP_TESTNAME_HEADEXISTING, response_head) & vbCrLf
    cReport.Concat ShowTestCase(APP_TESTNAME_OPTIONS, response_options) & vbCrLf
    cReport.Concat ShowTestCase(APP_TESTNAME_DELETEEXISTING, response_delete) & vbCrLf
    cReport.Concat ShowTestCase(APP_TESTNAME_WRONGMETHOD, response_testmethod) & vbCrLf
    cReport.Concat ShowTestCase(APP_TESTNAME_WRONGVERSION, response_protocolversion) & vbCrLf
    cReport.Concat ShowTestCase(APP_TESTNAME_ATTACKREQUEST, response_attackrequest) & vbCrLf

    cReport.Concat "<h4 id='details'>Fingerprint Details <a href='#'>&uarr;</a></h4>" & vbCrLf
    cReport.Concat ShowTestCase(APP_TESTNAME_GETEXISTING, GenerateFingerprintDetails(response_getexist)) & vbCrLf
    cReport.Concat ShowTestCase(APP_TESTNAME_GETLONG, GenerateFingerprintDetails(response_getlongrequest)) & vbCrLf
    cReport.Concat ShowTestCase(APP_TESTNAME_GETNONEXISTING, GenerateFingerprintDetails(response_get_nonexistent)) & vbCrLf
    cReport.Concat ShowTestCase(APP_TESTNAME_HEADEXISTING, GenerateFingerprintDetails(response_head)) & vbCrLf
    cReport.Concat ShowTestCase(APP_TESTNAME_OPTIONS, GenerateFingerprintDetails(response_options)) & vbCrLf
    cReport.Concat ShowTestCase(APP_TESTNAME_DELETEEXISTING, GenerateFingerprintDetails(response_delete)) & vbCrLf
    cReport.Concat ShowTestCase(APP_TESTNAME_WRONGMETHOD, GenerateFingerprintDetails(response_testmethod)) & vbCrLf
    cReport.Concat ShowTestCase(APP_TESTNAME_WRONGVERSION, GenerateFingerprintDetails(response_protocolversion)) & vbCrLf
    cReport.Concat ShowTestCase(APP_TESTNAME_ATTACKREQUEST, GenerateFingerprintDetails(response_attackrequest)) & vbCrLf

    cReport.Concat "<div id='bottom' class='copyright'>&copy; 2007-" & Year(Now) & " by <a href='" & APP_WEBSITE_URL & "'>" & APP_COPYRIGHT_OWNER & "</a></div>" & vbCrLf

    cReport.Concat "</body>" & vbCrLf
    cReport.Concat "</html>" & vbCrLf

    GenerateHtmlReport = cReport.Value
    
    Call ChangeStatusBarDone
End Function

Public Function ShowTestCase(ByRef sName As String, ByRef sResponse As String) As String
    Dim cTestcase As Concat
    Dim iLength As Integer
    
    Set cTestcase = New Concat
    
    iLength = Len(sResponse)
    
    cTestcase.Concat "<table class='table'>" & vbCrLf
    cTestcase.Concat "<tr class='databaseheader'><td>" & HtmlEncode(sName) & "</td><tr>" & vbCrLf
    If (iLength) Then
        cTestcase.Concat "<tr><td class='response' title='Length: " & iLength & " bytes'>" & HtmlEncode(sResponse) & "</td><tr>" & vbCrLf
    Else
        cTestcase.Concat "<tr class='databaseline'><td class='databaseline'>no response available</td><tr>" & vbCrLf
    End If
    cTestcase.Concat "</table><br />" & vbCrLf
    
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
    
    cResults.Concat "<table class='table'><tr class='databaseheader'><td style='width:20px'>&nbsp;</td><td>Name</td><td>Hits</td><td>Match</td></tr>" & vbCrLf
    For i = 1 To iListItemCount
         cResults.Concat "<tr class='databaseline'><td style='text-align:right' class='databaseline'>" & i & ".</td><td class='databaseline'>" & HtmlEncode(frmMain.lsvResults.ListItems(i).ListSubItems(1).Text) & "</td><td class='databaseline'>" & HtmlEncode(frmMain.lsvResults.ListItems(i).ListSubItems(2).Text) & "</td><td class='databaseline'>" & Round(frmMain.lsvResults.ListItems(i).ListSubItems(3).Text, 2) & "% </td></tr>" & vbCrLf
    Next i
    cResults.Concat "</table>" & vbCrLf
    
    GenerateHitList = cResults.Value
End Function

Public Function HtmlEncode(ByRef sInput As String) As String
    Dim sOutput As String
    
    sOutput = Replace$(sOutput, "&", "&amp;", 1, , vbBinaryCompare)
    
    sOutput = Replace$(sInput, "<", "&gt;", 1, , vbBinaryCompare)
    sOutput = Replace$(sOutput, ">", "&lt;", 1, , vbBinaryCompare)
    sOutput = Replace$(sOutput, Chr(34), "&quot;", 1, , vbBinaryCompare)
    
    sOutput = Replace$(sOutput, vbTab, " ", 1, , vbBinaryCompare)
    sOutput = Replace$(sOutput, vbCrLf, "<br />" & vbCrLf, 1, , vbBinaryCompare)
    
    HtmlEncode = sOutput
End Function

Public Function GetEmbeddedCSS() As String
    Dim cCSS As Concat
    
    Set cCSS = New Concat

    cCSS.Concat "<style type=""text/css"">" & vbCrLf
    cCSS.Concat "<!-- " & vbCrLf
    
    cCSS.Concat "body{" & vbCrLf
    cCSS.Concat "font-family:verdana;" & vbCrLf
    cCSS.Concat "font-size:11px;" & vbCrLf
    cCSS.Concat "color:black;" & vbCrLf
    cCSS.Concat "}" & vbCrLf
    
    cCSS.Concat "a{" & vbCrLf
    cCSS.Concat "color:darkred;" & vbCrLf
    cCSS.Concat "text-decoration:none;" & vbCrLf
    cCSS.Concat "}" & vbCrLf

    cCSS.Concat "a:hover{" & vbCrLf
    cCSS.Concat "color:red;" & vbCrLf
    cCSS.Concat "}" & vbCrLf

    cCSS.Concat "table.table{" & vbCrLf
    cCSS.Concat "border:1px solid darkgrey;" & vbCrLf
    cCSS.Concat "width:640px;" & vbCrLf
    cCSS.Concat "}" & vbCrLf

    cCSS.Concat "tr.databaseheader{" & vbCrLf
    cCSS.Concat "font-weight:bold;" & vbCrLf
    cCSS.Concat "background-color:darkgrey;" & vbCrLf
    cCSS.Concat "color:white;" & vbCrLf
    cCSS.Concat "}" & vbCrLf
        
    cCSS.Concat "tr.databaseline:hover{" & vbCrLf
    cCSS.Concat "background-color:lightgrey;" & vbCrLf
    cCSS.Concat "}" & vbCrLf

    cCSS.Concat "td.databaseline{" & vbCrLf
    cCSS.Concat "border:1px solid lightgrey;" & vbCrLf
    cCSS.Concat "}" & vbCrLf

    cCSS.Concat "td.response{" & vbCrLf
    cCSS.Concat "font-family:'courier new';" & vbCrLf
    cCSS.Concat "color:lightgreen;" & vbCrLf
    cCSS.Concat "background:black;" & vbCrLf
    cCSS.Concat "}" & vbCrLf

'    cCSS.Concat "td.fingerprint{"
'    cCSS.Concat "font-family:'courier new';"
'    cCSS.Concat "color:lightgrey;"
'    cCSS.Concat "background:black;"
'    cCSS.Concat "}"

    cCSS.Concat ".copyright{" & vbCrLf
    cCSS.Concat "font-size:10px;" & vbCrLf
    cCSS.Concat "}" & vbCrLf

    cCSS.Concat " -->" & vbCrLf
    cCSS.Concat "</style>" & vbCrLf
    
    GetEmbeddedCSS = cCSS.Value
End Function

Public Function WrapLine(ByRef sLine As String, Optional ByRef iLength As Integer = "72", Optional ByRef sBreak As String = vbCrLf) As String
    Dim iWrapCount As Integer
    Dim i As Integer
    Dim sTemp As String
    
    iWrapCount = Len(sLine) / iLength
    
    If (iWrapCount) Then
        sTemp = Mid$(sLine, 1, iLength)
        For i = 1 To iWrapCount
            sTemp = sTemp & vbCrLf & Mid$(sLine, (i * iLength) + 1, iLength)
        Next i
        WrapLine = sTemp
    Else
        WrapLine = sLine
    End If
End Function

Public Function GetLocalUsername() As String
    Dim sTemp As String
    
    sTemp = String(255, 0)
    GetUserName sTemp, 255
    GetLocalUsername = Left$(sTemp, InStr(sTemp, Chr$(0)) - 1)
End Function


