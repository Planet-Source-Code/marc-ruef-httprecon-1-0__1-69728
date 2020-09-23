Attribute VB_Name = "modAnalysisBanner"
Option Explicit

Public Function GetBanner(ByRef sInput As String) As String
    GetBanner = GetHeaderValue(sInput, "Server")
End Function

Public Function GetXPoweredBy(ByRef sInput As String) As String
    GetXPoweredBy = GetHeaderValue(sInput, "X-Powered-By")
End Function
