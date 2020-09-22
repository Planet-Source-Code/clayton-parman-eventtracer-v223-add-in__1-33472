Attribute VB_Name = "modTmpTrace"
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
'       This project was written in, and formatted for, Courier New font.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
'
'  Author:  Clayton Parman
'
'    Date:  April 08, 2001
'
'    Desc:  The procedure in this module creates the "tmpModule.txt"
'
'   Notes:  A "temporary text file is first written to the drive, then
'           added to the users Project as "modEventTracer.bas"
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'

Option Explicit

Const ModuleName  As String = "modTmpTrace"


Public Sub CreateTraceTempTxtModule()
     
Print #FF, "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'"
Print #FF, "'       This module was written in, and formatted for, Courier New font."
Print #FF, "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'"
Print #FF, "'    Author: Clayton Parman"
Print #FF, "'"
Print #FF, "'   Purpose: Adds control and text formatting of, " & vbQ & "debug.print" & vbQ & " output"
Print #FF, "'            to the Immediate Window, or to a Text file."
Print #FF, "'"
Print #FF, "'      Date: 02-27-02"
Print #FF, "'"
Print #FF, "'     Usage: The " & vbQ & "EventTracer" & vbQ & "Add-In automates the management of the commands"
Print #FF, "'            you can create manually from this module. Turning OFF the Trc?"
Print #FF, "'            commands, comments them out so they will not be compiled into"
Print #FF, "'            the final release, (if you want to keep them)."
Print #FF, "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'"
Print #FF, "'"
Print #FF, "' The " & vbQ & "TrcT" & vbQ & " command can be used for most any type of Text message you want to"
Print #FF, "' track in the Immediate window. This includes Procedures, Event Procedures,"
Print #FF, "' Functions, manual comments, and the value of most any variable."
Print #FF, "'"
Print #FF, "' Globally Comment out / Uncomment all " & vbQ & "Trc?" & vbQ & " commands to turn Event Tracing"
Print #FF, "' OFF and ON. (Where " & vbQ & "?" & vbQ & " is replaced by either " & vbQ & "T" & vbQ & " or " & vbQ & "V" & "). This can be easily"
Print #FF, "' done through the " & vbQ & "EventTracer" & vbQ & " Add-In."
Print #FF, "'"
Print #FF, "'"
Print #FF, "' NOTES: If there are Sub procedures which you NEVER want to follow in the"
Print #FF, "'        Immediate window, insert the word " & vbQ & "No" & vbQ & " in front of the Trc? command"
Print #FF, "'        (e.g. " & vbQ & "NoTrc?" & vbQ & ") and leave the command commented out in your code."
Print #FF, "'        This prevents " & vbQ & "EventTracer" & vbQ & " from accidentally adding it back in later."
Print #FF, "'"
Print #FF, "'        Multiple Trc? commands can be manually added anywhere in a procedure."
Print #FF, "'"
Print #FF, "'        Use " & vbQ & "TrcV" & vbQ & " to debug.print Variables. The format is similar"
Print #FF, "'         to " & vbQ & "TrcT" & vbQ & " except the 2nd parameter is the Name of the Variable."
Print #FF, "'"
Print #FF, "'"
Print #FF, "' 1st parameter (Text):     is a string value to print in the Immediate Window."
Print #FF, "'"
Print #FF, "'~~~~~~~~~~~~~~~ All of the remaining parameters are (optional): ~~~~~~~~~~~~~~'"
Print #FF, "'"
Print #FF, "' 2nd parameter (ContinueOnSameLine):      True-False."
Print #FF, "'      If True:  The " & vbQ & "next" & vbQ & " Trc? command prints on the current line."
Print #FF, "'"
Print #FF, "' 3rd parameter (HasConsecutiveCalls):     True-False."
Print #FF, "'      If True:  Consecutive, sequential calls to a procedure print 1-line only."
Print #FF, "'"
Print #FF, "' 4th parameter (Indent):            Inserts spaces in front of (Text)"
Print #FF, "'"
Print #FF, "' 5th parameter (BlankLinesBefore):  Prints one or more blank lines " & vbQ & "before" & vbQ
Print #FF, "'                                    outputting the (Text) line."
Print #FF, "'"
Print #FF, "' 6th parameter (BlankLinesAfter):   Prints one or more blank lines " & vbQ & "after" & vbQ
Print #FF, "'                                    outputting the (Text) line."
Print #FF, "'"
Print #FF, "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'"
Print #FF,
Print #FF, "Option Compare Text"
Print #FF,
Print #FF, "Const ModuleName      As String = " & vbQ & "modEventTracer" & vbQ
Print #FF,
Print #FF, "Public gFN            As Integer               'Holds a Freefile Number."
Print #FF,
Print #FF,
Print #FF, "Public Sub TrcT(Text As String, _"
Print #FF, "       Optional ContinueOnSameLine As Boolean, _"
Print #FF, "       Optional HasConsecutiveCalls As Boolean, _"
Print #FF, "       Optional Indent As Integer, _"
Print #FF, "       Optional BlankLinesBefore As Integer, _"
Print #FF, "       Optional BlankLinesAfter As Integer)"
Print #FF,
Print #FF, "   On Error    GoTo ErrorHandler"
Print #FF,
Print #FF, "   Dim intCnt                  As Integer"
Print #FF,
Print #FF, "   Static strLastText          As String"
Print #FF, "   Static intCallCount         As Integer"
Print #FF, "   Static LastWasConsecutive   As Boolean"
Print #FF, "   Static PrevBlankLinesAfter  As Integer"
Print #FF,
Print #FF, "   BlankLinesBefore = BlankLinesBefore - PrevBlankLinesAfter"
Print #FF, "   If BlankLinesBefore < 0 Then BlankLinesBefore = 0"
Print #FF,
Print #FF, "   If Instr(1,Text, " & vbQ & "DebugForOutput" & vbQ & ") Then"
Print #FF, "      If Instr(1,Text, " & vbQ & "Open" & vbQ & ") Then"
Print #FF, "         OpenDebugForOutput   'Send Trace to file instead of Immediate Window"
Print #FF, "      Else"
Print #FF, "         CloseDebugForOutput"
Print #FF, "      End IF"
Print #FF, "      Exit Sub"
Print #FF, "   End If"
Print #FF,
Print #FF, "   If InStr(Text," & vbQ & "_Resize" & vbQ & ") <> 0 Or _"
Print #FF, "      InStr(Text," & vbQ & "_MouseMove" & vbQ & ") <> 0 Or _"
Print #FF, "      InStr(Text," & vbQ & "_Timer" & vbQ & ") <> 0 Then        'These types of Events typically"
Print #FF, "      HasConsecutiveCalls = True            'have multiple consecutive calls."
Print #FF, "   End If"
Print #FF,
Print #FF, "   If HasConsecutiveCalls Then ContinueOnSameLine = False"
Print #FF,
Print #FF, "   If HasConsecutiveCalls Or LastWasConsecutive Then"
Print #FF,
Print #FF, "      If Text = strLastText Then"
Print #FF, "         intCallCount = intCallCount + 1"
Print #FF, "         Exit Sub"
Print #FF, "      End If"
Print #FF,
Print #FF, "      If intCallCount > 0 Then"
Print #FF,
Print #FF, "        'Finish printing the Call Counts for the previous Trc command."
Print #FF,
Print #FF, "         If gFN = 0 Then"
Print #FF, "            If intCallCount > 1 Then"
Print #FF, "               Debug.Print " & vbQ & "  * " & vbQ & "& intCallCount"
Print #FF, "            Else"
Print #FF, "               Debug.Print"
Print #FF, "            End If"
Print #FF,
Print #FF, "            For intCnt = 1 To PrevBlankLinesAfter"
Print #FF, "               Debug.Print "
Print #FF, "            Next intCnt"
Print #FF, "         Else"
Print #FF, "            If intCallCount > 1 Then"
Print #FF, "               Print #gFN, " & vbQ & "  * " & vbQ & "& intCallCount"
Print #FF, "            Else"
Print #FF, "               Print #gFN,"
Print #FF, "            End If"
Print #FF,
Print #FF, "            For intCnt = 1 To PrevBlankLinesAfter"
Print #FF, "               Print #gFN,"
Print #FF, "            Next intCnt"
Print #FF, "         End If"
Print #FF,
Print #FF, "         intCallCount = 0"
Print #FF, "         LastWasConsecutive = False"
Print #FF, "      End If"
Print #FF,
Print #FF, "     'Check to see if the new command is Consecutive too."
Print #FF,
Print #FF, "      If HasConsecutiveCalls Then"
Print #FF, "         If intCallCount = 0 Then"
Print #FF,
Print #FF, "            If gFN = 0 Then"
Print #FF, "               For intCnt = 1 To BlankLinesBefore"
Print #FF, "                  Debug.Print"
Print #FF, "               Next intCnt"
Print #FF,
Print #FF, "               Debug.Print Space(Indent) & Text;"
Print #FF, "            Else"
Print #FF, "               For intCnt = 1 To BlankLinesBefore"
Print #FF, "                  Print #gFN,"
Print #FF, "               Next intCnt"
Print #FF,
Print #FF, "               Print #gFN, Space(Indent) & Text;"
Print #FF, "            End If"
Print #FF,
Print #FF, "            intCallCount = 1"
Print #FF, "            LastWasConsecutive = True"
Print #FF, "         End If"
Print #FF, "      End If"
Print #FF, "   End If"
Print #FF,
Print #FF, "   If Not HasConsecutiveCalls Then"
Print #FF, "      If gFN = 0 Then"
Print #FF, "         For intCnt = 1 To BlankLinesBefore"
Print #FF, "            Debug.Print"
Print #FF, "         Next intCnt"
Print #FF,
Print #FF, "         If ContinueOnSameLine Then"
Print #FF, "            Debug.Print Space(Indent) & Text & " & vbQ & ", " & vbQ & ";"
Print #FF, "         Else"
Print #FF, "            Debug.Print Space(Indent) & Text"
Print #FF, "         End If"
Print #FF,
Print #FF, "         For intCnt = 1 To BlankLinesAfter"
Print #FF, "            Debug.Print"
Print #FF, "         Next intCnt"
Print #FF, "      Else"
Print #FF, "         For intCnt = 1 To BlankLinesBefore"
Print #FF, "            Print #gFN,"
Print #FF, "         Next intCnt"
Print #FF,
Print #FF, "         If ContinueOnSameLine Then"
Print #FF, "            Print #gFN, Space(Indent) & Text & " & vbQ & ", " & vbQ & ";"
Print #FF, "         Else"
Print #FF, "            Print #gFN, Space(Indent) & Text"
Print #FF, "         End If"
Print #FF,
Print #FF, "         For intCnt = 1 To BlankLinesAfter"
Print #FF, "            Print #gFN,"
Print #FF, "         Next intCnt"
Print #FF, "      End If"
Print #FF, "   End If"
Print #FF,
Print #FF, "   strLastText = Text"
Print #FF, "   PrevBlankLinesAfter = BlankLinesAfter      'Hold for multiple calls."
Print #FF,
Print #FF, "   Exit Sub"
Print #FF,
Print #FF, "ErrorHandler:"
Print #FF, ""
Print #FF, "   MsgBox Err & " & vbQ & ": Error in TrcT." & vbQ & " & vbCrLf & _"
Print #FF, "            " & vbQ & "Error Message: " & vbQ & " & Err.Description, vbCritical, " & vbQ & "Warning" & vbQ
Print #FF, ""
Print #FF, "   Resume Next"
Print #FF,
Print #FF, "End Sub"
Print #FF,
Print #FF,
Print #FF, "Public Sub TrcV(VarNameInQuotes As String, _"
Print #FF, "       Optional VarName As Variant, _"
Print #FF, "       Optional ContinueOnSameLine As Boolean, _"
Print #FF, "       Optional HasConsecutiveCalls As Boolean, _"
Print #FF, "       Optional Indent As Integer, _"
Print #FF, "       Optional BlankLinesBefore As Integer, _"
Print #FF, "       Optional BlankLinesAfter As Integer)"
Print #FF,
Print #FF, "   On Error    GoTo ErrorHandler"
Print #FF,
Print #FF, "   Dim strText   As String"
Print #FF, "   Dim strValue  As String"
Print #FF, "   Dim intType   As Integer"
Print #FF,
Print #FF, "   intType = VarType(VarName)"
Print #FF,
Print #FF, "   Select Case intType"
Print #FF, "      Case 0 To 1       'Empty, Null"
Print #FF, "         strValue = " & vbQ & vbQ
Print #FF, "      Case 2 To 7       'Integer, Long, Single, Double, Currency, Date"
Print #FF, "         strValue = CStr(VarName)"
Print #FF, "      Case 8            'String"
Print #FF, "         strValue = VarName"
Print #FF, "      Case 9            'Object"
Print #FF, "         strValue = " & vbQ & "Unable to debug.print TrcV VarName" & vbQ
Print #FF, "      Case 10 To 11     'Error, Byte"
Print #FF, "         strValue = CStr(VarName)"
Print #FF, "      Case Else"
Print #FF, "         strValue = " & vbQ & "Unable to debug.print TrcV VarName" & vbQ
Print #FF, "   End Select"
Print #FF,
Print #FF, "   strText = VarNameInQuotes & " & vbQ & "=" & vbQ & " & strValue"
Print #FF,
Print #FF, "CallTraceText:"
Print #FF,
Print #FF, "   TrcT strText, ContinueOnSameLine, HasConsecutiveCalls, _"
Print #FF, "                 Indent, BlankLinesBefore, BlankLinesAfter"
Print #FF,
Print #FF, "   Exit Sub"
Print #FF,
Print #FF, "ErrorHandler:"
Print #FF,
Print #FF, "   strText = " & vbQ & "Unable to debug.print " & vbQ & " & VarNameInQuotes &" & vbQ & " in TrcV." & vbQ & " & _"
Print #FF, "           vbCrLf & vbCrLf & _ "
Print #FF, "          " & vbQ & "Unprintable VarName" & vbQ
Print #FF,
Print #FF, "   MsgBox strText, vbOKOnly, " & vbQ & "TrcV Error" & vbQ
Print #FF,
Print #FF, "   strText = strText & " & vbQ & " Error in TrcV. Unable to debug.print" & vbQ
Print #FF,
Print #FF, "   Resume CallTraceText"
Print #FF,
Print #FF, "End Sub"
Print #FF,
Print #FF,
Print #FF, "Public Sub OpenDebugForOutput()"
Print #FF,
Print #FF, "   Dim intCnt   As Integer"
Print #FF, "   Dim strFile  As String"
Print #FF,
Print #FF, "   If gFN > 0 Then Exit Sub          'Already Opened. Do not Open again."
Print #FF,
Print #FF, "   intCnt = 1"
Print #FF,
Print #FF, "   strFile = Dir(" & vbQ & ".\Trace*.Log" & vbQ & ")"
Print #FF,
Print #FF, "   Do While strFile <> " & vbQ & vbQ
Print #FF, "      intCnt = intCnt + 1"
Print #FF, "      strFile = Dir                  'Count additional Trace?.Log files."
Print #FF, "   Loop"
Print #FF,
Print #FF, "   gFN = Freefile                    'Open Tracing to a Log File."
Print #FF,
Print #FF, "   Open " & vbQ & ".\Trace" & vbQ & "& intCnt &" & vbQ & ".Log" & vbQ & " For Output as #gFN"
Print #FF, "End Sub"
Print #FF,
Print #FF,
Print #FF, "Public Sub CloseDebugForOutput()"
Print #FF, "   If gFN <> 0 Then Close #gFN       'Close Tracing Write to File."
Print #FF, "      gFN = 0                        'Revert back to the Immediate Window."
Print #FF, "End Sub"
Print #FF,
Print #FF, "'Leave the following as the very last procedure in this module (for Speed.)"
Print #FF,
Print #FF, "Private Sub OnOffTestForTrace()"
Print #FF, "   TrcT " & vbQ & "OnOffTest" & vbQ & "                  'A No Op sub-procedure."
Print #FF, "End Sub"
   
End Sub
