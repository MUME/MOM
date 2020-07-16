Attribute VB_Name = "Tools"
Option Explicit
Option Compare Binary
Public lastedTime
Public oldTime

Public Function GetRegExp(strPattern As String, strBody As Variant, blnGlobal As Boolean)
  Dim objRegExp, Match, Matches   ' Create variable.
  Set objRegExp = CreateObject("VBScript.RegExp")   ' Create a regular expression.
  objRegExp.Pattern = strPattern   ' Set pattern.
  objRegExp.IgnoreCase = True   ' Set case insensitivity.
  objRegExp.Global = blnGlobal   ' Set global applicability.
  Set GetRegExp = objRegExp.Execute(strBody)   ' Execute search.
  Set objRegExp = Nothing
End Function

Public Sub Log(strLogFile As String, strLogText As String)
  On Error Resume Next
  Open App.Path & "\logs\" & strLogFile & ".log" For Append As #1
  Print #1, strLogText
  Dim n
   For n = 1 To Len(strLogText)
      Print #1, "(" & Mid(strLogText, n, 1) & ")" & Asc(Mid(strLogText, n, 1))
   Next
  Print #1, "__________________________________________________________"
  Close #1
End Sub

Public Function GetLogDate()
  GetLogDate = "_" & Day(Date) & "-" & Month(Date) & "-" & Year(Date)
End Function


