Attribute VB_Name = "modErrorHandler"
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
'       This module was written in, and formatted for, Courier New font.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
'
'  Author:  Clayton Parman
'
'    Date:  04-18-01
'
' Purpose:  "Central Error Handling"
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'

Option Explicit

Option Compare Text

Const ModuleName           As String = "modErrorHandler"

Public gbFormModeChange    As Boolean

Const NotHandled           As Integer = 99
Public Outcome             As Integer

Public Const RS            As Integer = 1       'ReSume
Public Const RN            As Integer = 2       'Resume Next
Public Const RL            As Integer = 3       'Resume Label:
Public Const GZ            As Integer = 4       'On Error GoTo 0 (Goto Zero)
Public Const EP            As Integer = 5       'Exit Project

'Note: Enclose all Error numbers (in strings) on both sides with "|" char.

Public Const GlobalExceptionRN  As String = "|9|"

'Only ONE OCCURRENCE of an Error Number can be listed in the following group!
 
Const CodeErrorsRS   As String = "|18|58|"
Const CodeErrorsRN   As String = "|3|5|6|9|10|11|13|14|17|35|49|59|63|335|"

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
'                    This is the Main Error handling procedure.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'

Public Sub HandleError(ErrNum As Long, ErrLine As Long, _
                       CodeModule As String, ProcName As String, _
                       Outcome As Integer)
                                             
   Dim sErrDesc  As String
   Dim sErrWork  As String
    
   sErrWork = "|" & Trim(CStr(Err.Number)) & "|"
    
  'Create the Error Description for the Message box.
    
   sErrDesc = "Error [ " & Err.Number & " ]:  " & Err.Description & vbCrLf _
   & vbCrLf & "Location:  Procedure ( " & ProcName & " )" & _
   " - Module ( " & CodeModule & " )" & IIf(ErrLine <> 0, _
   vbCrLf & vbCrLf & "Line number:  " & CStr(ErrLine), "") & vbCrLf
   
   Outcome = NotHandled      'NotHandled = "ExitProject"
      
   If InStr(1, GlobalExceptionRN, sErrWork) <> 0 Then Outcome = RN
   
  'If the Outcome still = NotHandled, then the Error is NOT an Exception.
   
   If Outcome = NotHandled Then
      
      If InStr(1, CodeErrorsRN, sErrWork) Then
         MsgBox sErrDesc, vbExclamation, "Attention!"
         Outcome = RN
      End If

      If Outcome = NotHandled Then    'Final catch-all for unhandled Errors!
         MsgBox sErrDesc, vbCritical, "Unexpected Error!"
         Outcome = EP
      End If
      
   End If
   
End Sub
