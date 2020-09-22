Attribute VB_Name = "modTracer"
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
'       This module was written in, and formatted for, Courier New font.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
'
'    Author: Clayton Parman
'
'   Purpose: Custom version of modEventTracer to debug EventTracer application.
'
'      Date: 05-26-02
'
'     Notes: Uncomment this module and all TrT & TrV commands if frmAddIn in
'            order to run this module. Comment them out again before shipping.
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'



'Option Explicit
'Option Compare Text
'
'Const ModuleName      As String = "modTracer"
'
'Public gFN            As Integer               'Holds a Freefile Number.
'
'
'Public Sub TrT(Text As String, _
'       Optional ContinueOnSameLine As Boolean, _
'       Optional HasConsecutiveCalls As Boolean, _
'       Optional Indent As Integer, _
'       Optional BlankLinesBefore As Integer, _
'       Optional BlankLinesAfter As Integer)
'
'   On Error GoTo ErrorHandler
'
'   Dim intCnt                  As Integer
'
'   Static strLastText          As String
'   Static intCallCount         As Integer
'   Static LastWasConsecutive   As Boolean
'   Static PrevBlankLinesAfter  As Integer
'
'   BlankLinesBefore = BlankLinesBefore - PrevBlankLinesAfter
'   If BlankLinesBefore < 0 Then BlankLinesBefore = 0
'
'   If InStr(1, Text, "DebugForOutput") Then
'      If InStr(1, Text, "Open") Then
'         OpenDebugForOutput   'Send Trace to file instead of Immediate Window
'      Else
'         CloseDebugForOutput
'      End If
'      Exit Sub
'   End If
'
'   If InStr(Text, "_Resize") <> 0 Or _
'      InStr(Text, "_MouseMove") <> 0 Or _
'      InStr(Text, "_Timer") <> 0 Then       'These types of Events typically
'      HasConsecutiveCalls = True            'have multiple consecutive calls.
'   End If
'
'   If HasConsecutiveCalls Then ContinueOnSameLine = False
'
'   If HasConsecutiveCalls Or LastWasConsecutive Then
'
'      If Text = strLastText Then
'         intCallCount = intCallCount + 1
'         Exit Sub
'      End If
'
'      If intCallCount > 0 Then
'
'        'Finish printing the Call Counts for the previous Trc command.
'
'         If gFN = 0 Then
'            If intCallCount > 1 Then
'               Debug.Print "  * " & intCallCount
'            Else
'               Debug.Print
'            End If
'
'            For intCnt = 1 To PrevBlankLinesAfter
'               Debug.Print
'            Next intCnt
'         Else
'            If intCallCount > 1 Then
'               Print #gFN, "  * " & intCallCount
'            Else
'               Print #gFN,
'            End If
'
'            For intCnt = 1 To PrevBlankLinesAfter
'               Print #gFN,
'            Next intCnt
'         End If
'
'         intCallCount = 0
'         LastWasConsecutive = False
'      End If
'
'     'Check to see if the new command is Consecutive too.
'
'      If HasConsecutiveCalls Then
'         If intCallCount = 0 Then
'
'            If gFN = 0 Then
'               For intCnt = 1 To BlankLinesBefore
'                  Debug.Print
'               Next intCnt
'
'               Debug.Print Space(Indent) & Text;
'            Else
'               For intCnt = 1 To BlankLinesBefore
'                  Print #gFN,
'               Next intCnt
'
'               Print #gFN, Space(Indent) & Text;
'            End If
'
'            intCallCount = 1
'            LastWasConsecutive = True
'         End If
'      End If
'   End If
'
'   If Not HasConsecutiveCalls Then
'      If gFN = 0 Then
'         For intCnt = 1 To BlankLinesBefore
'            Debug.Print
'         Next intCnt
'
'         If ContinueOnSameLine Then
'            Debug.Print Space(Indent) & Text & ", ";
'         Else
'            Debug.Print Space(Indent) & Text
'         End If
'
'         For intCnt = 1 To BlankLinesAfter
'            Debug.Print
'         Next intCnt
'      Else
'         For intCnt = 1 To BlankLinesBefore
'            Print #gFN,
'         Next intCnt
'
'         If ContinueOnSameLine Then
'            Print #gFN, Space(Indent) & Text & ", ";
'         Else
'            Print #gFN, Space(Indent) & Text
'         End If
'
'         For intCnt = 1 To BlankLinesAfter
'            Print #gFN,
'         Next intCnt
'      End If
'   End If
'
'   strLastText = Text
'   PrevBlankLinesAfter = BlankLinesAfter      'Hold for multiple calls.
'
'   Exit Sub
'
'ErrorHandler:
'
'   MsgBox Err & ": Error in TrT." & vbCrLf & _
'            "Error Message: " & Err.Description, vbCritical, "Warning"
'
'   Resume Next
'
'End Sub
'
'
'Public Sub TrV(VarNameInQuotes As String, _
'       Optional VarName As Variant, _
'       Optional ContinueOnSameLine As Boolean, _
'       Optional HasConsecutiveCalls As Boolean, _
'       Optional Indent As Integer, _
'       Optional BlankLinesBefore As Integer, _
'       Optional BlankLinesAfter As Integer)
'
'   On Error GoTo ErrorHandler
'
'   Dim strText   As String
'   Dim strValue  As String
'   Dim intType   As Integer
'
'   intType = VarType(VarName)
'
'   Select Case intType
'      Case 0 To 1       'Empty, Null
'         strValue = ""
'      Case 2 To 7       'Integer, Long, Single, Double, Currency, Date
'         strValue = CStr(VarName)
'      Case 8            'String
'         strValue = VarName
'      Case 9            'Object
'         strValue = "Unable to debug.print TrV VarName"
'      Case 10 To 11     'Error, Byte
'         strValue = CStr(VarName)
'      Case Else
'         strValue = "Unable to debug.print TrV VarName"
'   End Select
'
'   If strValue = "True" Then strValue = "T"   'Shorten up
'   If strValue = "False" Then strValue = "F"
'
'   strText = VarNameInQuotes & "=" & strValue
'
'CallTraceText:
'
'   TrT strText, ContinueOnSameLine, HasConsecutiveCalls, _
'                Indent, BlankLinesBefore, BlankLinesAfter
'
'   Exit Sub
'
'ErrorHandler:
'
'   strText = "Unable to debug.print " & VarNameInQuotes & " in TrV." & _
'              vbCrLf & vbCrLf & _
'             "Unprintable VarName"
'
'   MsgBox strText, vbExclamation + vbOKOnly, "TrV Error"
'
'   strText = strText & " Error in TrcV. Unable to debug.print"
'
'   Resume CallTraceText
'
'End Sub
'
'
'Public Sub OpenDebugForOutput()
'
'   Dim intCnt   As Integer
'   Dim strFile  As String
'
'   If gFN > 0 Then Exit Sub          'Already Opened. Do not Open again.
'
'   intCnt = 1
'
'   strFile = Dir(".\Trace*.Log")
'
'   Do While strFile <> ""
'      intCnt = intCnt + 1
'      strFile = Dir                  'Count additional Trace?.Log files.
'   Loop
'
'   gFN = FreeFile                    'Open Tracing to a Log File.
'
'   Open ".\Trace" & intCnt & ".Log" For Output As #gFN
'End Sub
'
'
'Public Sub CloseDebugForOutput()
'   If gFN <> 0 Then Close #gFN       'Close Tracing Write to File.
'      gFN = 0                        'Revert back to the Immediate Window.
'End Sub
