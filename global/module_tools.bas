Attribute VB_Name = "tools"
Option Explicit

Public Sub informClient(ByRef s As String, Optional force As Boolean)
On Error Resume Next
   If force = True Or frmMap.mnuFeedback.Checked = True Then
      frmMap.tcpMUD.SendData lookHeader & "0" & lookFooter & "# " & s & vbCrLf
   End If
On Error GoTo 0
End Sub

Public Function GetRegExp(strPattern As String, strBody As Variant, blnGlobal As Boolean)
  Dim objRegExp, Match, Matches   ' Create variable.
  Set objRegExp = CreateObject("VBScript.RegExp")   ' Create a regular expression.
  objRegExp.Pattern = strPattern   ' Set pattern.
  objRegExp.IgnoreCase = True   ' Set case insensitivity.
  objRegExp.Global = blnGlobal   ' Set global applicability.
  Set GetRegExp = objRegExp.Execute(strBody)   ' Execute search.
  Set objRegExp = Nothing
End Function

Public Function IIF(expression As Boolean, valueDefault, valueElse) As String
   If (expression = True) Then
      IIF = valueDefault
   Else
      IIF = valueElse
   End If
End Function
