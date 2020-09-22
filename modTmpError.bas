Attribute VB_Name = "modTmpError"
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
'       This project was written in, and formatted for, Courier New font.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
'
'  Author:  Clayton Parman
'
'    Date:  April 22, 2001
'
'    Desc:  The procedure in this module creates the "tmpModule.txt"
'
'   Notes:  A "temporary text file is first written to the drive, then
'           added to the users Project as "modErrorHandler.bas"'
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'

Option Explicit

Const ModuleName As String = "modTmpError"


Public Sub CreateErrorTempTxtModule()

Print #FF, "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'"
Print #FF, "'       This module was written in, and formatted for, Courier New font."
Print #FF, "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'"
Print #FF, "'"
Print #FF, "'  Author:  Clayton Parman"
Print #FF, "'"
Print #FF, "'    Date:  04-22-01"
Print #FF, "'"
Print #FF, "' Purpose:  " & vbQ & "Centralized Error Handling" & vbQ & " (with optional Error logging)."
Print #FF, "'"
Print #FF, "'   Notes:  This module has a lot of code, but the logic is simple and"
Print #FF, "'           highly repetitive. It should be very easy to understand with"
Print #FF, "'           little studying. The design is very simple and can be easily"
Print #FF, "'           modified and/or expanded as your project requires."
Print #FF, "'"
Print #FF, "'           It has been tested to make sure the logic works. However, not"
Print #FF, "'           all of the Errors have been tested under " & vbQ & "real" & vbQ & " conditions. In"
Print #FF, "'           many instances a " & vbQ & "best guess" & vbQ & " was made for the proper Outcome."
Print #FF, "'"
Print #FF, "'           This module was written using the article for " & vbQ & "Centralized"
Print #FF, "'           Error Handling" & vbQ & " in the VB Online Documentation as a guide."
Print #FF, "'           One of the design philosophies was to build in an excess of options,"
Print #FF, "'           based on the consideration that it is usually much faster to remove"
Print #FF, "'           extra code than it is to add it."
Print #FF, "'"
Print #FF, "'           All errors below 325 (and a few above 325) are handled."
Print #FF, "'           See " & vbQ & "Trappable Errors" & vbQ & " in the VB Online Documentation for a"
Print #FF, "'           list of all the errors that " & vbQ & "can" & vbQ & " be handled in this module."
Print #FF, "'"
Print #FF, "' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~[ HELP ]~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
Print #FF, "'"
Print #FF, "'     A. The two character identifiers at the end of the Variable Names"
Print #FF, "'        specify the expected " & vbQ & "Outcome" & vbQ & " for an error. For example:"
Print #FF, "'"
Print #FF, "'          Select Case Outcome               'Variable Name examples"
Print #FF, "'             Case RS:   ReSume              'GlobalExceptionRS,  ExceptionRS"
Print #FF, "'             Case RN:   Resume Next         'GlobalExceptionRN,  ExceptionRN"
Print #FF, "'             Case RL:   Resume Label:       'GlobalExceptionRL,  ExceptionRL"
Print #FF, "'             Case GZ:   On Error GoTo 0     'GlobalExceptionGZ,  ExceptionGZ"
Print #FF, "'             Case EP:   ExitProject (call)  'GlobalExceptionEP,  ExceptionEP"
Print #FF, "'          End Select"
Print #FF, "'"
Print #FF, "'     B. This module has a number of Optional Error Exceptions and Optional"
Print #FF, "'        Procedures, (with pre-defined entry points), built into it to give"
Print #FF, "'        the Error Handling a lot of flexibility."
Print #FF, "'"
Print #FF, "'           1. Define your own " & vbQ & "Global Error Exceptions" & vbQ & " (None by default.)"
Print #FF, "'           2. Define your own " & vbQ & "Local Error Exceptions" & vbQ & "  (None by default.)"
Print #FF, "'           3. Customizable procedures for any routines that need to be"
Print #FF, "'              done " & vbQ & "before" & vbQ & " or " & vbQ & "after" & vbQ & " the error message has displayed."
Print #FF, "'           4. Customizable procedures, for custom Outcomes, for any error."
Print #FF, "'"
Print #FF, "'     C. Delimit all Error numbers in the " & vbQ & "Exception strings" & vbQ & " using"
Print #FF, "'        vertical bar (pipe) characters.   e.g. " & vbQ & "|9|11|16|" & vbQ & "  or  " & vbQ & "|11|" & vbQ
Print #FF, "'"
Print #FF, "'     D. An Error number that is set up as a " & vbQ & "Global Exception" & vbQ & " can be"
Print #FF, "'        temporarily shut off by sending the Error Number as a " & vbQ & "Passed"
Print #FF, "'        Exception" & vbQ & ".   i.e. Exceptions which are " & vbQ & "both" & vbQ & " Global and Passed"
Print #FF, "'        are handled " & vbQ & "normally" & vbQ & ", not as an exception."
Print #FF, "'"
Print #FF, "'     E. Logging Errors to a file is OFF by default. Set"
Print #FF, "'        " & vbQ & "mblnLogErrorToFile = True" & vbQ & " to turn this feature ON (or define"
Print #FF, "'        a condition, or setup option, which will turn it ON.) Separate"
Print #FF, "'        Log files are maintained for each month, with a maximum of"
Print #FF, "'        three months of log files. Older logs are deleted automatically."
Print #FF, "'        See " & vbQ & "WriteErrorToLog" & vbQ & " for the names of the Log files."
Print #FF, "'"
Print #FF, "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'"
Print #FF,
Print #FF, "Option Compare Text"
Print #FF,
Print #FF, "Const ModuleName            As String = " & vbQ & "modErrorHandler" & vbQ
Print #FF,
Print #FF, "Const mblnLogErrorsToFile   As Boolean = True    'Turn ON as needed."
Print #FF,
Print #FF, "Const NotHandled            As Integer = 99"
Print #FF, "Public Outcome              As Integer"
Print #FF,
Print #FF, "Public Const RS             As Integer = 1       'ReSume"
Print #FF, "Public Const RN             As Integer = 2       'Resume Next"
Print #FF, "Public Const RL             As Integer = 3       'Resume Label:"
Print #FF, "Public Const GZ             As Integer = 4       'On Error GoTo 0 (Goto Zero)"
Print #FF, "Public Const EP             As Integer = 5       'Exit Project"
Print #FF,
Print #FF, "'Note: Enclose all Error numbers (in strings) on both sides with " & vbQ & "|" & vbQ & " char."
Print #FF,
Print #FF, "Public Const GlobalExceptionRS As String = " & vbQ & vbQ
Print #FF, "Public Const GlobalExceptionRN As String = " & vbQ & vbQ
Print #FF, "Public Const GlobalExceptionRL As String = " & vbQ & vbQ
Print #FF, "Public Const GlobalExceptionGZ As String = " & vbQ & vbQ
Print #FF, "Public Const GlobalExceptionEP As String = " & vbQ & vbQ
Print #FF,
Print #FF, "'Only ONE OCCURRENCE of an Error Number can be listed in the following group!"
Print #FF,
Print #FF, "Const CodeErrorsRS   As String = " & vbQ & "|18|58|" & vbQ
Print #FF, "Const CodeErrorsRN   As String = " & vbQ & "|3|5|6|9|10|11|13|14|17|35|49|59|63|335|" & vbQ
Print #FF, "Const CodeErrorsRL   As String = " & vbQ & vbQ
Print #FF, "Const CodeErrorsGZ   As String = " & vbQ & "|20|" & vbQ
Print #FF, "Const CodeErrorsEP   As String = " & vbQ & "|7|16|28|47|48|51|67|70|92|93|94|" & vbQ
Print #FF, "Const ObjectErrors   As String = " & vbQ & "|91|97|98|" & vbQ
Print #FF, "Const PrinterErrors  As String = " & vbQ & "|396|482|483|484|486|" & vbQ
Print #FF, "Const DatabaseErrors As String = " & vbQ & "|3004|3021|3022|3024|3044|3058|3315|" & vbQ
Print #FF, "Const FileErrors     As String = " & vbQ & "|52|53|54|55|57|61|62|64|68|71|74|75|76" & vbQ & " & _"
Print #FF, "                                 " & vbQ & "|298|320|321|322|325|735|31036|31037|" & vbQ
Print #FF,
Print #FF, "'              WinAPI network Functions for login names."
Print #FF,
Print #FF, "Private Declare Function GetWorkStation Lib " & vbQ & "kernel32" & vbQ & " _"
Print #FF, "Alias " & vbQ & "GetComputerNameA" & vbQ & " (ByVal lpBuffer As String, nSize As Long) As Long"
Print #FF, ""
Print #FF, "Private Declare Function WNetGetUser Lib " & vbQ & "mpr.dll" & vbQ & " Alias " & vbQ & "WNetGetUserA" & vbQ & " _"
Print #FF, "(ByVal lpName As String, ByVal lpUserName As String, lpnLength As Long) As Long"
Print #FF,
Print #FF, "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'"
Print #FF, "'    You can code your own Custom Outcomes ... Otherwise leave them empty."
Print #FF, "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'"
Print #FF,
Print #FF, "Private Sub CustomRS(ErrNum As Long, ErrDesc As String, _"
Print #FF, "                     CodeModule As String, ProcName As String)"
Print #FF, "   'Optional - Your Custom Code for " & vbQ & "ReSume" & vbQ
Print #FF, "End Sub"
Print #FF,
Print #FF, "Private Sub CustomRN(ErrNum As Long, ErrDesc As String, _"
Print #FF, "                     CodeModule As String, ProcName As String)"
Print #FF, "   'Optional - Your Custom Code for " & vbQ & "Resume Next" & vbQ
Print #FF, "End Sub"
Print #FF,
Print #FF, "Private Sub CustomRL(ErrNum As Long, ErrDesc As String, _"
Print #FF, "                     CodeModule As String, ProcName As String)"
Print #FF, "   'Optional - Your Custom Code for " & vbQ & "Resume Label:" & vbQ
Print #FF, "End Sub"
Print #FF,
Print #FF, "Private Sub CustomGZ(ErrNum As Long, ErrDesc As String, _"
Print #FF, "                     CodeModule As String, ProcName As String)"
Print #FF, "   'Optional - Your Custom Code for " & vbQ & "Goto Zero" & vbQ
Print #FF, "End Sub"
Print #FF,
Print #FF, "Private Sub CustomEP(ErrNum As Long, ErrDesc As String, _"
Print #FF, "                     CodeModule As String, ProcName As String)"
Print #FF, "   'Optional - Your Custom Code for " & vbQ & "Exit Project" & vbQ
Print #FF, "End Sub"
Print #FF,
Print #FF, "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'"
Print #FF, "'            You can code " & vbQ & "before" & vbQ & " and " & vbQ & "after" & vbQ & " error routines here."
Print #FF, "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'"
Print #FF,
Print #FF, "Private Sub DoBeforeErrorMessage(ErrNum As Long, ErrDesc As String, _"
Print #FF, "                                 CodeModule As String, ProcName As String)"
Print #FF, "   'Optional - Code which executes " & vbQ & "before" & vbQ & " error is handled."
Print #FF, "   '(Such as setting a form that is TOPMOST to be NOTOPMOST.)"
Print #FF, "End Sub"
Print #FF,
Print #FF,
Print #FF, "Private Sub DoAfterErrorMessage(ErrNum As Long, ErrDesc As String, _"
Print #FF, "                                CodeModule As String, ProcName As String)"
Print #FF, "   'Optional - Code which executes " & vbQ & "after" & vbQ & " error has been handled."
Print #FF, "   '(Such as re-setting a form to always be TOPMOST.)"
Print #FF, "End Sub"
Print #FF,
Print #FF, "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'"
Print #FF, "'                    This is the Main Error handling procedure."
Print #FF, "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'"
Print #FF,
Print #FF, "Public Sub HandleError(ErrNum As Long, ErrLine As Long, _"
Print #FF, "                       CodeModule As String, ProcName As String, _"
Print #FF, "                       Outcome As Integer, _"
Print #FF, "                       Optional ExceptionRS As String, _ "
Print #FF, "                       Optional ExceptionRN As String, _"
Print #FF, "                       Optional ExceptionRL As String, _"
Print #FF, "                       Optional ExceptionGZ As String, _"
Print #FF, "                       Optional ExceptionEP As String)"
Print #FF,
Print #FF, "   Dim sErrDesc  As String"
Print #FF, "   Dim sErrWork  As String"
Print #FF,
Print #FF, "   sErrWork = " & vbQ & "|" & vbQ & " & Trim(CStr(Err.Number)) & " & vbQ & "|" & vbQ
Print #FF,
Print #FF, "  'Create the Error Description for the Message box."
Print #FF,
Print #FF, "   sErrDesc = " & vbQ & "Error [ " & vbQ & " & Err.Number & " & vbQ & " ]:  " & vbQ & " & Err.Description & vbCrLf _"
Print #FF, "   & vbCrLf & " & vbQ & "Location:  Procedure ( " & vbQ & " & ProcName & " & vbQ & " )" & vbQ & " & _"
Print #FF, "   " & vbQ & " - Module ( " & vbQ & " & CodeModule & " & vbQ & " )" & vbQ & " & IIf(ErrLine <> 0, _"
Print #FF, "   vbCrLf & vbCrLf & " & vbQ & "Line number:  " & vbQ & " & CStr(ErrLine), " & vbQ & vbQ & ") & vbCrLf"
Print #FF,
Print #FF, "   On Error Resume Next      'In case of an Error within Error Module."
Print #FF,
Print #FF, "  'A customizable procedure which Runs " & vbQ & "before" & vbQ & " the message is displayed."
Print #FF,
Print #FF, "   DoBeforeErrorMessage ErrNum, sErrDesc, CodeModule, ProcName"
Print #FF,
Print #FF, "   Outcome = NotHandled      'NotHandled = " & vbQ & "ExitProject" & vbQ
Print #FF,
Print #FF, "  'Check for any Passed Exceptions."
Print #FF,
Print #FF, "   If InStr(1, ExceptionRS, sErrWork) <> 0 Then"
Print #FF, "      If Not InStr(1, GlobalExceptionRS, sErrWork) <> 0 Then Outcome = RS"
Print #FF, "   End If"
Print #FF,
Print #FF, "   If InStr(1, ExceptionRN, sErrWork) <> 0 Then"
Print #FF, "      If Not InStr(1, GlobalExceptionRN, sErrWork) <> 0 Then Outcome = RN"
Print #FF, "   End If"
Print #FF,
Print #FF, "   If InStr(1, ExceptionRL, sErrWork) <> 0 Then"
Print #FF, "      If Not InStr(1, GlobalExceptionRL, sErrWork) <> 0 Then Outcome = RL"
Print #FF, "   End If"
Print #FF,
Print #FF, "   If InStr(1, ExceptionGZ, sErrWork) <> 0 Then"
Print #FF, "      If Not InStr(1, GlobalExceptionGZ, sErrWork) <> 0 Then Outcome = GZ"
Print #FF, "   End If"
Print #FF,
Print #FF, "   If InStr(1, ExceptionEP, sErrWork) <> 0 Then"
Print #FF, "      If Not InStr(1, GlobalExceptionEP, sErrWork) <> 0 Then Outcome = EP"
Print #FF, "   End If"
Print #FF,
Print #FF, "  'If not a Passed Exception, then check to see if it is a Global Exception."
Print #FF,
Print #FF, "   If Outcome = NotHandled Then"
Print #FF, "      If InStr(1, GlobalExceptionRS, sErrWork) <> 0 Then"
Print #FF, "         If Not InStr(1, ExceptionRS, sErrWork) <> 0 Then Outcome = RS"
Print #FF, "      End If"
Print #FF,
Print #FF, "      If InStr(1, GlobalExceptionRN, sErrWork) <> 0 Then"
Print #FF, "         If Not InStr(1, ExceptionRN, sErrWork) <> 0 Then Outcome = RN"
Print #FF, "      End If"
Print #FF,
Print #FF, "      If InStr(1, GlobalExceptionRL, sErrWork) <> 0 Then"
Print #FF, "         If Not InStr(1, ExceptionRL, sErrWork) <> 0 Then Outcome = RL"
Print #FF, "      End If"
Print #FF,
Print #FF, "      If InStr(1, GlobalExceptionGZ, sErrWork) <> 0 Then"
Print #FF, "         If Not InStr(1, ExceptionGZ, sErrWork) <> 0 Then Outcome = GZ"
Print #FF, "      End If"
Print #FF,
Print #FF, "      If InStr(1, GlobalExceptionEP, sErrWork) <> 0 Then"
Print #FF, "         If Not InStr(1, ExceptionEP, sErrWork) <> 0 Then Outcome = EP"
Print #FF, "      End If"
Print #FF, "   End If"
Print #FF,
Print #FF, "  'If the Outcome still = NotHandled, then the Error is NOT an Exception."
Print #FF,
Print #FF, "   If Outcome = NotHandled Then"
Print #FF,
Print #FF, "      If InStr(1, CodeErrorsRS, sErrWork) Then"
Print #FF, "         MsgBox sErrDesc, vbInformation, " & vbQ & "Attention!" & vbQ
Print #FF, "         Outcome = RS"
Print #FF, "      End If"
Print #FF,
Print #FF, "      If InStr(1, CodeErrorsRN, sErrWork) Then"
Print #FF, "         MsgBox sErrDesc, vbExclamation, " & vbQ & "Attention!" & vbQ
Print #FF, "         Outcome = RN"
Print #FF, "      End If"
Print #FF,
Print #FF, "      If InStr(1, CodeErrorsRL, sErrWork) Then"
Print #FF, "         MsgBox sErrDesc, vbInformation, " & vbQ & "Attention!" & vbQ
Print #FF, "         Outcome = RL"
Print #FF, "      End If"
Print #FF,
Print #FF, "      If InStr(1, CodeErrorsGZ, sErrWork) Then"
Print #FF, "         MsgBox sErrDesc, vbExclamation, " & vbQ & "Attention!" & vbQ
Print #FF, "         Outcome = GZ"
Print #FF, "      End If"
Print #FF,
Print #FF, "      If InStr(1, CodeErrorsEP, sErrWork) Then"
Print #FF, "         MsgBox sErrDesc, vbCritical, " & vbQ & "Warning!" & vbQ
Print #FF, "         Outcome = EP"
Print #FF, "      End If"
Print #FF,
Print #FF, "      If InStr(1, FileErrors, sErrWork) Then"
Print #FF, "         Outcome = HandleFileErrors(ErrNum)"
Print #FF, "      End If"
Print #FF,
Print #FF, "      If InStr(1, DatabaseErrors, sErrWork) Then"
Print #FF, "         Outcome = HandleDatabaseErrors(ErrNum)"
Print #FF, "      End If"
Print #FF,
Print #FF, "      If InStr(1, PrinterErrors, sErrWork) Then"
Print #FF, "         Outcome = HandlePrinterErrors(ErrNum)"
Print #FF, "      End If"
Print #FF,
Print #FF, "      If InStr(1, ObjectErrors, sErrWork) Then"
Print #FF, "         Outcome = HandleObjectErrors(ErrNum)"
Print #FF, "      End If"
Print #FF,
Print #FF, "      If Outcome = NotHandled Then    'Final catch-all for unhandled Errors!"
Print #FF, "         MsgBox sErrDesc, vbCritical, " & vbQ & "Unexpected Error!" & vbQ
Print #FF, "         Outcome = EP"
Print #FF, "      End If"
Print #FF,
Print #FF, "   End If"
Print #FF,
Print #FF, "  'CustomRS, RN, RL, etc are customizable procedures which you can use for"
Print #FF, "  'special conditions.  As is, they have absolutely no effect on " & vbQ & "Outcome" & vbQ & "."
Print #FF,
Print #FF, "   Select Case Outcome"
Print #FF, "     Case RS:   CustomRS ErrNum, sErrDesc, CodeModule, ProcName  'ReSume"
Print #FF, "     Case RN:   CustomRN ErrNum, sErrDesc, CodeModule, ProcName  'Resume Next"
Print #FF, "     Case RL:   CustomRL ErrNum, sErrDesc, CodeModule, ProcName  'Resume Label"
Print #FF, "     Case GZ:   CustomGZ ErrNum, sErrDesc, CodeModule, ProcName  'Goto Zero"
Print #FF, "     Case EP:   CustomEP ErrNum, sErrDesc, CodeModule, ProcName  'Exit Project"
Print #FF, "     Case Else: Outcome = EP"
Print #FF, "   End Select"
Print #FF,
Print #FF, "   If mblnLogErrorsToFile = True Then WriteErrorToLog sErrDesc, Outcome"
Print #FF,
Print #FF, "  'A customizable procedure which Runs " & vbQ & "after" & vbQ & " the message has displayed."
Print #FF,
Print #FF, "   DoAfterErrorMessage ErrNum, sErrDesc, CodeModule, ProcName"
Print #FF,
Print #FF, "End Sub"
Print #FF,
Print #FF, "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'"
Print #FF, "'  This procedure handles " & vbQ & "Database" & vbQ & " errors and is called by " & vbQ & "HandleError()" & vbQ
Print #FF, "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'"
Print #FF,
Print #FF, "Private Function HandleDatabaseErrors(ErrNum As Long) As Integer"
Print #FF,
Print #FF, "   Select Case ErrNum"
Print #FF,
Print #FF, "      Case 3004, 3024, 3044          'Databse, File, Path not found"
Print #FF, "         HandleDatabaseErrors = RS   'ReSume"
Print #FF,
Print #FF, "      Case 3021                      'No current record"
Print #FF, "         HandleDatabaseErrors = RN   'Resume Next"
Print #FF,
Print #FF, "      Case 3022                      'Duplicate key field"
Print #FF, "         HandleDatabaseErrors = GZ   'Shut off error trap to allow change."
Print #FF,
Print #FF, "      Case 3058, 3315                'No entry in key field."
Print #FF, "         HandleDatabaseErrors = GZ   'Shut off error trap to allow completion."
Print #FF,
Print #FF, "      Case Else"
Print #FF, "         HandleDatabaseErrors = EP   'Exit Project"
Print #FF,
Print #FF, "   End Select"
Print #FF,
Print #FF, "End Function"
Print #FF,
Print #FF, "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'"
Print #FF, "'   This procedure handles " & vbQ & "Printer" & vbQ & " errors and is called by " & vbQ & "HandleError()" & vbQ
Print #FF, "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'"
Print #FF,
Print #FF, "Private Function HandlePrinterErrors(ErrNum As Long) As Integer"
Print #FF,
Print #FF, "   Dim strMsg      As String"
Print #FF,
Print #FF, "   Dim intButtons  As Integer"
Print #FF, "   Dim intResponse As Integer"
Print #FF,
Print #FF, "   Select Case ErrNum"
Print #FF,
Print #FF, "      Case 396"
Print #FF, "         strMsg = " & vbQ & "Property cannot be set within a page" & vbQ
Print #FF, "         intButtons = vbExclamation + vbOKCancel"
Print #FF,
Print #FF, "      Case 482"
Print #FF, "         strMsg = " & vbQ & "Printer Error" & vbQ
Print #FF, "         intButtons = vbExclamation + vbAbortRetryIgnore"
Print #FF,
Print #FF, "      Case 483"
Print #FF, "         strMsg = " & vbQ & "Printer driver does not support the property" & vbQ
Print #FF, "         intButtons = vbExclamation + vbOKCancel"
Print #FF,
Print #FF, "      Case 484"
Print #FF, "         strMsg = " & vbQ & "Printer driver unavailable" & vbQ
Print #FF, "         intButtons = vbExclamation + vbOKCancel"
Print #FF,
Print #FF, "      Case 486"
Print #FF, "         strMsg = " & vbQ & "Can't print form image to this type of printer" & vbQ
Print #FF, "         intButtons = vbExclamation + vbOKCancel"
Print #FF,
Print #FF, "   End Select"
Print #FF,
Print #FF, "   intResponse = MsgBox(strMsg, intButtons, " & vbQ & "Printer Error" & vbQ & ")"
Print #FF,
Print #FF, "   Select Case intResponse"
Print #FF, "      Case vbOK:      HandlePrinterErrors = RS    'ReSume"
Print #FF, "      Case vbAbort:   HandlePrinterErrors = EP    'Exit Project"
Print #FF, "      Case vbRetry:   HandlePrinterErrors = RS    'ReSume"
Print #FF, "      Case vbIgnore:  HandlePrinterErrors = RN    'Resume Next"
Print #FF, "      Case vbCancel:  HandlePrinterErrors = EP    'Exit Project"
Print #FF, "      Case Else:      HandlePrinterErrors = RS    'ReSume"
Print #FF, "   End Select"
Print #FF,
Print #FF, "End Function"
Print #FF,
Print #FF, "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'"
Print #FF, "'   This procedure handles " & vbQ & "File" & vbQ & " errors and is called by " & vbQ & "HandleError()" & vbQ
Print #FF, "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'"
Print #FF,
Print #FF, "Private Function HandleFileErrors(ErrNum As Long) As Integer"
Print #FF,
Print #FF, "   Dim strMsg      As String"
Print #FF,
Print #FF, "   Dim intButtons  As Integer"
Print #FF, "   Dim intResponse As Integer"
Print #FF,
Print #FF, "   Select Case ErrNum"
Print #FF,
Print #FF, "      Case 52, 64                                 '52-Bad file name or number"
Print #FF, "         strMsg = " & vbQ & "That filename is illegal." & vbQ
Print #FF, "         intButtons = vbExclamation + vbOKCancel"
Print #FF,
Print #FF, "      Case 53                                     '53-File not found"
Print #FF, "         strMsg = " & vbQ & "File not found!" & vbQ
Print #FF, "         intButtons = vbExclamation + vbOKCancel"
Print #FF,
Print #FF, "      Case 54                                     '54-Bad file mode"
Print #FF, "         strMsg = " & vbQ & "Can't open your file for that type of access." & vbQ
Print #FF, "         intButtons = vbExclamation + vbOKCancel"
Print #FF,
Print #FF, "      Case 55                                     '55-File already open"
Print #FF, "         strMsg = " & vbQ & "This file is already open." & vbQ
Print #FF, "         intButtons = vbExclamation + vbOKOnly"
Print #FF,
Print #FF, "      Case 57                                     '57-Device I/O error"
Print #FF, "         strMsg = " & vbQ & "Internal disk error." & vbQ
Print #FF, "         intButtons = vbExclamation + vbOKOnly"
Print #FF,
Print #FF, "      Case 61                                     '61-Disk full"
Print #FF, "         strMsg = " & vbQ & "Disk is full. Continue?" & vbQ
Print #FF, "         intButtons = vbExclamation + vbAbortRetryIgnore"
Print #FF,
Print #FF, "      Case 62                                     '62-Input past end of file"
Print #FF, "         strMsg = " & vbQ & "This file has a nonstandard end-of-file marker," & vbQ & " & vbCrLf"
Print #FF, "         strMsg = strMsg & " & vbQ & "or an attempt was made " & vbQ
Print #FF, "         strMsg = strMsg & " & vbQ & "to read beyond the end-of-file marker." & vbQ
Print #FF, "         intButtons = vbExclamation + vbAbortRetryIgnore"
Print #FF,
Print #FF, "      Case 68                                     '68-Device unavailable"
Print #FF, "         strMsg = " & vbQ & "That device appears unavailable." & vbQ
Print #FF, "         intButtons = vbExclamation + vbOKCancel"
Print #FF,
Print #FF, "      Case 71                                     '71-Disk not ready"
Print #FF, "         strMsg = " & vbQ & "Insert a disk in the drive and close the door." & vbQ
Print #FF, "         intButtons = vbExclamation + vbOKCancel"
Print #FF,
Print #FF, "      Case 74                          '74-Can't rename with different drive"
Print #FF, "         strMsg = " & vbQ & "Cannot rename with a different drive" & vbQ
Print #FF, "         intButtons = vbExclamation + vbOKCancel"
Print #FF,
Print #FF, "      Case 76                                     '75-Path/File access error"
Print #FF, "         strMsg = " & vbQ & "Cannot access either Path or File." & vbQ
Print #FF, "         intButtons = vbExclamation + vbOKCancel"
Print #FF,
Print #FF, "      Case 76                                     '76-Path not found"
Print #FF, "         strMsg = " & vbQ & "That path doesn't exist" & vbQ
Print #FF, "         intButtons = vbExclamation + vbOKCancel"
Print #FF,
Print #FF, "      Case 298"
Print #FF, "         strMsg = " & vbQ & "System DLL could not be loaded" & vbQ
Print #FF, "         intButtons = vbExclamation + vbOKCancel"
Print #FF,
Print #FF, "      Case 320"
Print #FF, "         strMsg = " & vbQ & "Can't use character device names in specified file names" & vbQ
Print #FF, "         intButtons = vbExclamation + vbOKCancel"
Print #FF,
Print #FF, "      Case 321"
Print #FF, "         strMsg = " & vbQ & "Invalid file format" & vbQ
Print #FF, "         intButtons = vbExclamation + vbOKCancel"
Print #FF,
Print #FF, "      Case 322"
Print #FF, "         strMsg = " & vbQ & "Can't create necessary temporary file" & vbQ
Print #FF, "         intButtons = vbExclamation + vbAbortRetryIgnore"
Print #FF,
Print #FF, "      Case 325"
Print #FF, "         strMsg = " & vbQ & "Invalid format in resource file" & vbQ
Print #FF, "         intButtons = vbExclamation + vbAbortRetryIgnore"
Print #FF,
Print #FF, "      Case 735"
Print #FF, "         strMsg = " & vbQ & "Can't save file to TEMP directory" & vbQ
Print #FF, "         intButtons = vbExclamation + vbAbortRetryIgnore"
Print #FF,
Print #FF, "      Case 31036"
Print #FF, "         strMsg = " & vbQ & "Error saving to file" & vbQ
Print #FF, "         intButtons = vbExclamation + vbOKCancel"
Print #FF,
Print #FF, "      Case 31037"
Print #FF, "         strMsg = " & vbQ & "Error loading from file" & vbQ
Print #FF, "         intButtons = vbExclamation + vbOKCancel"
Print #FF,
Print #FF, "   End Select"
Print #FF,
Print #FF, "   intResponse = MsgBox(strMsg, intButtons, " & vbQ & "Disk Error" & vbQ & ")"
Print #FF,
Print #FF, "   Select Case intResponse"
Print #FF, "      Case vbOK:      HandleFileErrors = RS    'ReSume"
Print #FF, "      Case vbAbort:   HandleFileErrors = EP    'Exit Project"
Print #FF, "      Case vbRetry:   HandleFileErrors = RS    'ReSume"
Print #FF, "      Case vbIgnore:  HandleFileErrors = RN    'Resume Next"
Print #FF, "      Case vbCancel:  HandleFileErrors = EP    'Exit Project"
Print #FF, "      Case Else:      HandleFileErrors = RS    'ReSume"
Print #FF, "   End Select"
Print #FF,
Print #FF, "End Function"
Print #FF,
Print #FF, "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'"
Print #FF, "'  This procedure handles " & vbQ & "Object" & vbQ & " errors and is called by " & vbQ & "HandleError()" & vbQ
Print #FF, "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'"
Print #FF,
Print #FF, "Private Function HandleObjectErrors(ErrNum As Long) As Integer"
Print #FF,
Print #FF, "   Select Case ErrNum"
Print #FF,
Print #FF, "      Case 91                      'Object variable or With block variable not set"
Print #FF, "         HandleObjectErrors = EP   'Exit Project"
Print #FF,
Print #FF, "      Case 97                      'Can't call Friend procedure on an object that is ..."
Print #FF, "         HandleObjectErrors = EP   'Exit Project"
Print #FF,
Print #FF, "      Case 98                      'Property or method call cannot reference a private object ..."
Print #FF, "         HandleObjectErrors = EP   'Exit Project"
Print #FF,
Print #FF, "      Case Else"
Print #FF, "         HandleObjectErrors = EP   'Exit Project"
Print #FF,
Print #FF, "   End Select"
Print #FF,
Print #FF, "End Function"
Print #FF,
Print #FF, "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'"
Print #FF, "' This procedure " & vbQ & "can" & vbQ & " be used with " & vbQ & "HandleFileErrors()" & vbQ & " to decide a course"
Print #FF, "' of action for sequential file access modes.  It calls HandleFileErrors()."
Print #FF, "' You can Delete this procedure if it is not used. It is here for convenience."
Print #FF, "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'"
Print #FF,
Print #FF, "Function ConfirmFile(FName As String, Operation As Integer) As Integer"
Print #FF,
Print #FF, "   ' Parameters:"
Print #FF, "   '      Fname:  File to be checked for and confirmed."
Print #FF, "   '  Operation:  Code for sequential file access mode (Output, Input, etc.)"
Print #FF, "   '"
Print #FF, "   ' Note that the procedure works for binary & random access because messages"
Print #FF, "   ' are conditioned on Operation being <> to certain sequential modes."
Print #FF, "   '"
Print #FF, "   ' Return values:"
Print #FF, "   '    1   Confirms operation will not cause a problem."
Print #FF, "   '    0   User decided not to go through with operation."
Print #FF,
Print #FF, "   Const conSaveFile = 1"
Print #FF, "   Const conLoadFile = 2"
Print #FF, "   Const conReplaceFile = 1"
Print #FF, "   Const conReadFile = 2"
Print #FF, "   Const conAddToFile = 3"
Print #FF, "   Const conRandomFile = 4"
Print #FF, "   Const conBinaryFile = 5"
Print #FF,
Print #FF, "   Dim intConfirmation As Integer"
Print #FF, "   Dim intAction       As Integer"
Print #FF, "   Dim intErrNum       As Integer"
Print #FF, "   Dim varMsg          As Variant"
Print #FF,
Print #FF, "   On Error GoTo ConfirmFileError   ' Turn error trap ON."
Print #FF,
Print #FF, "   FName = Dir(FName)               ' See if the file exists."
Print #FF,
Print #FF, "   On Error GoTo 0                  ' Turn error trap OFF."
Print #FF,
Print #FF, "  'If user is saving text to a file that already exists..."
Print #FF,
Print #FF, "   If FName <> " & vbQ & vbQ & " And Operation = conReplaceFile Then"
Print #FF,
Print #FF, "      varMsg = " & vbQ & "The file " & vbQ & " & FName & " & vbQ & "already exists on " & vbQ & " & vbCrLf _"
Print #FF, "      & " & vbQ & "disk. Saving the text box " & vbQ & " & vbCrLf _"
Print #FF, "      & " & vbQ & "contents to that file will destroy the file's current " & vbQ & " _"
Print #FF, "      & " & vbQ & "contents, " & vbQ & " & vbCrLf _"
Print #FF, "      & " & vbQ & "replacing them with the text from the text box." & vbQ & " _"
Print #FF, "      & vbCrLf & vbCrLf _"
Print #FF, "      & " & vbQ & "Choose OK to replace file, Cancel to stop." & vbQ
Print #FF,
Print #FF, "      intConfirmation = MsgBox(varMsg, 65, " & vbQ & "File Message" & vbQ & ")"
Print #FF,
Print #FF, "     'If user wants to load text from a file that doesn't exist."
Print #FF,
Print #FF, "   ElseIf FName = " & vbQ & vbQ & " And Operation = conReadFile Then"
Print #FF,
Print #FF, "      varMsg = " & vbQ & "The file " & vbQ & " & FName & " & vbQ & " doesn't exist." & vbQ & " & vbCrLf _"
Print #FF, "      & " & vbQ & "Would you like to create and then edit it?" & vbQ & " & vbCrLf _"
Print #FF, "      & vbCrLf & " & vbQ & "Choose OK to create file, Cancel to stop." & vbQ
Print #FF,
Print #FF, "      intConfirmation = MsgBox(varMsg, 65, " & vbQ & "File Message" & vbQ & ")"
Print #FF,
Print #FF, "     'If FName doesn't exist:"
Print #FF, "     'force procedure to return 0 by setting intConfirmation = 2."
Print #FF,
Print #FF, "   ElseIf FName = " & vbQ & vbQ & " Then"
Print #FF,
Print #FF, "      If Operation = conRandomFile Or Operation = conBinaryFile Then"
Print #FF, "         intConfirmation = 2"
Print #FF, "      End If"
Print #FF,
Print #FF, "     'If the file exists and operation isn't successful,"
Print #FF, "     'intConfirmation = 0 and procedure returns 1."
Print #FF,
Print #FF, "   End If"
Print #FF,
Print #FF, "  'If no box was displayed, intConfirmation = 0."
Print #FF,
Print #FF, "  'In either case, intConfirmation = 1 and ConfirmFile should"
Print #FF, "  'return 1 to confirm that the intended operation is OK."
Print #FF, "  '"
Print #FF, "  'If intConfirmation > 1, ConfirmFile should return 0,"
Print #FF, "  'because user doesn't want to go through with the operation..."
Print #FF,
Print #FF, "   If intConfirmation > 1 Then"
Print #FF, "      ConfirmFile = 0"
Print #FF, "   Else"
Print #FF, "      ConfirmFile = 1"
Print #FF,
Print #FF, "      If intConfirmation = 1 Then          'User wants to create file."
Print #FF,
Print #FF, "         If Operation = conLoadFile Then   'Assign conReplaceFile so"
Print #FF, "            Operation = conReplaceFile     'caller will understand the"
Print #FF, "         End If                            'action that will be taken."
Print #FF,
Print #FF, "        'Return code confirming action to either"
Print #FF, "        'replace existing file or create new one."
Print #FF, "      End If"
Print #FF, "   End If"
Print #FF,
Print #FF, "   Exit Function"
Print #FF,
Print #FF, "ConfirmFileError:"
Print #FF,
Print #FF, "   intAction = HandleFileErrors(Err.Number)"
Print #FF,
Print #FF, "   Select Case intAction"
Print #FF, "      Case 0:  Resume"
Print #FF, "      Case 1:  Resume Next"
Print #FF, "      Case 2:  Exit Function"
Print #FF, "      Case Else"
Print #FF, "         intErrNum = Err.Number"
Print #FF, "         Err.Raise Number:=intErrNum"
Print #FF, "         Err.Clear"
Print #FF, "   End Select"
Print #FF,
Print #FF, "End Function"
Print #FF,
Print #FF, "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'"
Print #FF, "'             Procedures for writing errors to a log file."
Print #FF, "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'"
Print #FF,
Print #FF, "Private Sub WriteErrorToLog(ErrDesc As String, Outcome As Integer)"
Print #FF,
Print #FF, "   On Error GoTo ErrorHandler"
Print #FF,
Print #FF, "   Dim Today            As Date"
Print #FF,
Print #FF, "   Dim intFF            As Integer"
Print #FF, "   Dim intMM            As Integer"
Print #FF, "   Dim intYY            As Integer"
Print #FF, "   Dim intCnt           As Integer"
Print #FF,
Print #FF, "   Dim strMM            As String"
Print #FF, "   Dim strYY            As String"
Print #FF, "   Dim sTrcurrLogFile   As String"
Print #FF, "   Dim strPrev1LogFile  As String"
Print #FF, "   Dim strPrev2LogFile  As String"
Print #FF, "   Dim strAnyLogFile    As String"
Print #FF, "   Dim strOutcome       As String"
Print #FF,
Print #FF, "   Dim lngSizeOfFile    As Long"
Print #FF,
Print #FF, "   ErrDesc = Replace(ErrDesc, vbCrLf & vbCrLf, vbCrLf, 1)  'Replace double line"
Print #FF, "   ErrDesc = Replace(ErrDesc, vbCrLf & vbCrLf, vbCrLf, 1)  'feeds with singles."
Print #FF, "   ErrDesc = Left(ErrDesc, Len(ErrDesc) - 2)               'Remove last CR/LF."
Print #FF,
Print #FF, "   Select Case Outcome"
Print #FF, "      Case 1: strOutcome = " & vbQ & "Resume" & vbQ
Print #FF, "      Case 2: strOutcome = " & vbQ & "Resume Next" & vbQ
Print #FF, "      Case 3: strOutcome = " & vbQ & "Resume at label" & vbQ
Print #FF, "      Case 4: strOutcome = " & vbQ & "On Error GoTo 0" & vbQ
Print #FF, "      Case 5: strOutcome = " & vbQ & "Exit Project" & vbQ
Print #FF, "   End Select"
Print #FF,
Print #FF, "   Today = Now"
Print #FF,
Print #FF, "   For intCnt = 0 To 2"
Print #FF,
Print #FF, "      intYY = Year(Today)"
Print #FF, "      intMM = Month(Today)"
Print #FF,
Print #FF, "      intMM = intMM - intCnt"
Print #FF,
Print #FF, "      If intMM < 1 Then"
Print #FF, "         intMM = intMM + 12"
Print #FF, "         intYY = intYY - 1"
Print #FF, "      End If"
Print #FF,
Print #FF, "      strMM = Right(0 & CInt(intMM), 2)"
Print #FF, "      strYY = Right(0 & CInt(intYY), 2)"
Print #FF,
Print #FF, "     'Create the name of current month's log file and previous two months."
Print #FF,
Print #FF, "      Select Case intCnt"
Print #FF, "         Case 0:  sTrcurrLogFile = App.Title & strYY & strMM & " & vbQ & "Error.Log" & vbQ
Print #FF, "         Case 1:  strPrev1LogFile = App.Title & strYY & strMM & " & vbQ & "Error.Log" & vbQ
Print #FF, "         Case 2:  strPrev2LogFile = App.Title & strYY & strMM & " & vbQ & "Error.Log" & vbQ
Print #FF, "      End Select"
Print #FF,
Print #FF, "   Next intCnt"
Print #FF,
Print #FF, "   strAnyLogFile = Dir(App.Path & " & vbQ & "\" & vbQ & " & App.Title & " & vbQ & "????Error.Log" & vbQ & ")"
Print #FF,
Print #FF, "  'Delete any existing Log files that are more than two months old."
Print #FF,
Print #FF, "   Do While strAnyLogFile <> """
Print #FF, "      If strAnyLogFile <> sTrcurrLogFile Then"
Print #FF, "         If strAnyLogFile <> strPrev1LogFile Then"
Print #FF, "            If strAnyLogFile <> strPrev2LogFile Then"
Print #FF, "               Kill strAnyLogFile                     'Delete old Log file."
Print #FF, "            End If"
Print #FF, "         End If"
Print #FF, "      Else"
Print #FF, "         lngSizeOfFile = FileLen(sTrcurrLogFile)    'Length of current Log."
Print #FF, "      End If"
Print #FF, "      strAnyLogFile = Dir"
Print #FF, "   Loop"
Print #FF,
Print #FF, "   sTrcurrLogFile = App.Path & " & vbQ & "\" & vbQ & " & sTrcurrLogFile"
Print #FF,
Print #FF, "   intFF = FreeFile                     'Get the next available File number."
Print #FF,
Print #FF, "   Open sTrcurrLogFile For Append Shared As #intFF"
Print #FF,
Print #FF, "   Seek #intFF, lngSizeOfFile + 1          'Bug work around for end-of-file."
Print #FF, "                             "
Print #FF, "  'Write the Error to the log file."
Print #FF,
Print #FF, "   Print #intFF, ErrDesc _"
Print #FF, "        & Space(20) & " & vbQ & "Outcome: " & vbQ & " & strOutcome & vbCrLf & _"
Print #FF, "         " & vbQ & "Date-Time: " & vbQ & " & Format(Now, " & vbQ & "mm/dd/yy hh:nn:ss" & vbQ & ") _"
Print #FF, "        & Space(5) & " & vbQ & "Workstation: " & vbQ & " & WorkStation _"
Print #FF, "        & Space(5) & " & vbQ & "User: " & vbQ & " & UserName & vbCrLf & _"
Print #FF, "  " & vbQ & "----------------------------------------------------------------------------" & vbQ
Print #FF,
Print #FF, "CloseLog:"
Print #FF,
Print #FF, "   Close #intFF"
Print #FF,
Print #FF, "   Exit Sub"
Print #FF,
Print #FF, "ErrorHandler:"
Print #FF,
Print #FF, "   HandleError Err.Number, Erl, ModuleName, " & vbQ & "WriteErrorToLog" & vbQ & ", Outcome"
Print #FF,
Print #FF, "   Resume CloseLog 'Do not Exit the project for Errors in the log file."
Print #FF,
Print #FF, "End Sub"
Print #FF,
Print #FF,
Print #FF, "Private Function WorkStation() As String"
Print #FF,
Print #FF, "   On Error Resume Next                    'Returns the Computers Name."
Print #FF,
Print #FF, "   Dim strWork As String * 255"
Print #FF,
Print #FF, "   strWork = Space(255)"
Print #FF,
Print #FF, "   If GetWorkStation(strWork, 255&) <> 0 Then"
Print #FF, "      WorkStation = Left$(strWork, InStr(strWork, vbNullChar) - 1)"
Print #FF, "   Else"
Print #FF, "      WorkStation = " & vbQ & "(Unknown)" & vbQ
Print #FF, "   End If"
Print #FF,
Print #FF, "End Function"
Print #FF,
Print #FF,
Print #FF, "Private Function UserName() As String"
Print #FF,
Print #FF, "   On Error Resume Next                    'Returns the Users log-in Name."
Print #FF,
Print #FF, "   Dim strWork As String * 255"
Print #FF,
Print #FF, "   strWork = Space(255)"
Print #FF,
Print #FF, "   Call WNetGetUser(vbNullString, strWork, 255&)"
Print #FF,
Print #FF, "   UserName = Left$(strWork, InStr(strWork, vbNullChar) - 1)"
Print #FF,
Print #FF, "End Function"
Print #FF,
Print #FF, "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'"
Print #FF, "'   This procedure does a general Close Project. Customize it as needed."
Print #FF, "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'"
Print #FF,
Print #FF, "Public Sub ExitProject()"
Print #FF,
Print #FF, "  'A general " & vbQ & "End/Close Project" & vbQ & " procedure."
Print #FF,
Print #FF, "   Dim Frm  As Form"
Print #FF,
Print #FF, "   For Each Frm In Forms"
Print #FF, "       Unload Frm"
Print #FF, "       Set Frm = Nothing"
Print #FF, "   Next Frm"
Print #FF,
Print #FF, "   End 'Quit VB Project."
Print #FF,
Print #FF, "End Sub"
Print #FF,
Print #FF, "'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'"

End Sub




