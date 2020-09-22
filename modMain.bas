Attribute VB_Name = "modMain"
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
'       This project was written in, and formatted for, Courier New font.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
'
'  Author:  Clayton Parman
'
'    Date:  April 01, 2001
'
'    Desc:  Global variables and misc procedures used by "EventTracer"
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'

Option Explicit

Const ModuleName            As String = "modMain"

' The following variable makes sure the "Call SetFocus(frmAddIn.hwnd)
' does not get executed when the program is first loaded.

Public gbOkToSetFocus       As Boolean
Public gbOpenedAllPanes     As Boolean
                     
Public Const vbQ            As String = """"

Public FF                   As Integer       'FreeFile number for Modules

'The following API is used with the [Move] "CursorToFocus" Function.

Public Declare Function SetCursorPos Lib "user32" _
                       (ByVal x As Long, ByVal y As Long) As Long

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'

Public Function IsCompiled() As Boolean

 'If Debug.Print divide by zero triggers Error, then IsCompiled stays False.

  On Error GoTo NotCompiled

   Debug.Print 1 / 0
   IsCompiled = True

NotCompiled:

End Function
