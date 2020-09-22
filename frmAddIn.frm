VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Begin VB.Form frmAddIn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Event Tracer"
   ClientHeight    =   6075
   ClientLeft      =   2175
   ClientTop       =   1935
   ClientWidth     =   2955
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddIn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   2955
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab 
      Height          =   4875
      Left            =   75
      TabIndex        =   5
      Top             =   315
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   8599
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Event Tracing"
      TabPicture(0)   =   "frmAddIn.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblFiller(2)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblFiller(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblFiller(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblFiller(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdTrace(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdTrace(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdTrace(2)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdTrace(6)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdTrace(7)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdTrace(8)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdTrace(4)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdTrace(5)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "fraScope(0)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdTrace(3)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "ErrorHandling"
      TabPicture(1)   =   "frmAddIn.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraScope(1)"
      Tab(1).Control(1)=   "cmdErrors(2)"
      Tab(1).Control(2)=   "cmdErrors(1)"
      Tab(1).Control(3)=   "cmdErrors(0)"
      Tab(1).ControlCount=   4
      Begin VB.CommandButton cmdTrace 
         Caption         =   "INSERT Trace BELOW Variable"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   3
         Left            =   150
         TabIndex        =   29
         Top             =   2660
         Width           =   2500
      End
      Begin VB.Frame fraScope 
         Caption         =   " Scope "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1150
         Index           =   1
         Left            =   -74850
         TabIndex        =   25
         Top             =   375
         Width           =   2500
         Begin VB.OptionButton optScopeError 
            Caption         =   " Project"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   750
            TabIndex        =   28
            Top             =   225
            Width           =   1440
         End
         Begin VB.OptionButton optScopeError 
            Caption         =   " Module"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   750
            TabIndex        =   27
            Top             =   525
            Width           =   1440
         End
         Begin VB.OptionButton optScopeError 
            Caption         =   " Procedure"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   750
            TabIndex        =   26
            Top             =   825
            Width           =   1440
         End
      End
      Begin VB.Frame fraScope 
         Caption         =   " Scope "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1150
         Index           =   0
         Left            =   150
         TabIndex        =   21
         Top             =   375
         Width           =   2500
         Begin VB.OptionButton OptScopeTrace 
            Caption         =   " Project"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   750
            TabIndex        =   24
            Top             =   225
            Width           =   1440
         End
         Begin VB.OptionButton OptScopeTrace 
            Caption         =   " Module"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   750
            TabIndex        =   23
            Top             =   525
            Width           =   1440
         End
         Begin VB.OptionButton OptScopeTrace 
            Caption         =   " Procedure"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   750
            TabIndex        =   22
            Top             =   825
            Width           =   1440
         End
      End
      Begin VB.CommandButton cmdTrace 
         Caption         =   "INSERT  Trace to File -  CLOSE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   5
         Left            =   150
         TabIndex        =   20
         Top             =   3345
         Width           =   2500
      End
      Begin VB.CommandButton cmdTrace 
         Caption         =   "INSERT  Trace  to  File -  OPEN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   4
         Left            =   150
         TabIndex        =   19
         Top             =   3030
         Width           =   2500
      End
      Begin VB.CommandButton cmdTrace 
         Caption         =   "DELETE Trace Commands..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   8
         Left            =   150
         TabIndex        =   17
         Top             =   4410
         Width           =   2500
      End
      Begin VB.CommandButton cmdTrace 
         Caption         =   "Turn Trace Commands   ON"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   7
         Left            =   150
         TabIndex        =   16
         Top             =   4040
         Width           =   2500
      End
      Begin VB.CommandButton cmdTrace 
         Caption         =   "Turn Trace Commands  OFF"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   6
         Left            =   150
         TabIndex        =   15
         Top             =   3725
         Width           =   2500
      End
      Begin VB.CommandButton cmdTrace 
         Caption         =   "INSERT  Trace ABOVE Variable"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   2
         Left            =   150
         TabIndex        =   14
         Top             =   2345
         Width           =   2500
      End
      Begin VB.CommandButton cmdTrace 
         Caption         =   "INSERT Trace  Procedure  EXIT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   1
         Left            =   150
         TabIndex        =   13
         Top             =   1965
         Width           =   2500
      End
      Begin VB.CommandButton cmdTrace 
         Caption         =   "INSERT Trace for PROCEDURE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   0
         Left            =   150
         TabIndex        =   12
         Top             =   1650
         Width           =   2500
      End
      Begin VB.CommandButton cmdErrors 
         Caption         =   "REMOVE  Numbers from Lines"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   2
         Left            =   -74850
         TabIndex        =   11
         Top             =   2460
         Width           =   2500
      End
      Begin VB.CommandButton cmdErrors 
         Caption         =   "ADD Procedure Line Numbers"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   1
         Left            =   -74850
         TabIndex        =   10
         Top             =   2055
         Width           =   2500
      End
      Begin VB.CommandButton cmdErrors 
         Caption         =   "INSERT Error Handling Routine"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   0
         Left            =   -74850
         TabIndex        =   9
         Top             =   1650
         Width           =   2500
      End
      Begin VB.Label lblFiller 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   160
         TabIndex        =   30
         Top             =   4275
         Width           =   2485
      End
      Begin VB.Label lblFiller 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   160
         TabIndex        =   18
         Top             =   2160
         Width           =   2485
      End
      Begin VB.Label lblFiller 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   160
         TabIndex        =   8
         Top             =   2885
         Width           =   2485
      End
      Begin VB.Label lblFiller 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   160
         TabIndex        =   7
         Top             =   3555
         Width           =   2485
      End
   End
   Begin VB.CommandButton cmdMisc 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   75
      TabIndex        =   4
      Top             =   5640
      Width           =   930
   End
   Begin VB.CommandButton cmdMisc 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   1005
      TabIndex        =   3
      Top             =   5640
      Width           =   940
   End
   Begin VB.CommandButton cmdMisc 
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   1950
      TabIndex        =   2
      Top             =   5640
      Width           =   930
   End
   Begin VB.CommandButton cmdMisc 
      Caption         =   "Close All VB Code Windows"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   75
      TabIndex        =   1
      Top             =   5340
      Width           =   2800
   End
   Begin VB.PictureBox picBox 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   75
      Left            =   255
      ScaleHeight     =   75
      ScaleWidth      =   2340
      TabIndex        =   0
      Top             =   5235
      Width           =   2340
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   75
      TabIndex        =   6
      Top             =   15
      Width           =   2805
   End
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
'       This project was written in, and formatted for, Courier New font.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
'
'  Author:  Clayton Parman
'
'    Date:  April 01, 2001
' Revised:  February 27, 2002 - April 1, 2002 - May 26, 2002
'
'    Desc:  Adds a formatted "debug.print" routine to Sub Procedures and
'           Functions so their sequence of execution can be traced in the
'           Immediate Window (or to "Trace?.Log" ... See Notes below).
'
'           Also, has a capability for adding a "Centralized Error Handler"
'           to a Project and its Procedures. Help for programming with the
'           Error Handler are included in the "modErrorHandler" that is
'           created when you add Error Handling to a Project or a Procedure.
'
' Effects:  This Add-In inserts a Trace Text (TrcT) command in Sub Procedures
'           and Functions and adds the module "modEventTracer.bas" to your
'           Project (This module has the "TrcT" and "TrcV" procedures).
'
'           Each procedures is checked for pre-existing "TrcT" commands before
'           inserting a new one. Subsequent uses of this program on a Project
'           or Module only updates those Functions and Procedures which have
'           not already been processed.
'
'           "TrcT" commands can be manually added to any section of code that
'           you wish to verify the execution of (example:  TrcT "Hit here")
'
'   Notes:  The "INSERT Trace for VARIABLE" (TrcV) Procedure can be used
'           to print variables. Use it anywhere and as often as needed.
'           The variable value can be printed either before or after it
'           changes.
'
'           If you "INSERT Trace Procedure EXIT" and the procedure you are
'           working in has an "Exit Sub" (or "Exit Function") command(s): An
'           "ExitSub:" label will be added to the end of the procedure and
'           the "Exit Sub" will be changed to either a "GoTo ExitSub" or
'           "GoTo ExitFunction".
'
'           Instructions for how to use "TrcT" and "TrcV" are included
'           in the "modEventTracer.bas" which is created by this program.
'
'           The "Insert Trace to File - OPEN (and CLOSE)" will redirect the
'           output to a Text file named "Trace?.Log" (where "?" is a sequen-
'           tial number). If you are not sure which procedure is the first
'           one to be executed, you can place multiple "Trace OPENS" within
'           your program. Only one of them will execute ... until a CLOSE
'           command has executed. To start at "Trace1.Log", be sure that all
'           "Trace*.Log" files have been deleted before running your program.
'
'Potential: Add automatic indenting for each level. Would require that each
'Future     procedure contain a "TrcT xxx [Exit]". An Indexed array in the
'Enhance-   "modEventTracer.bas" could be used to keep track of the indent-
'ments      ing and outdenting of levels. Each new level would be indented
'           by 3. This would make complicated Traces much easier to read.
'           NOTE: Add a button for "Insert Trace for BOTH" which would add
'           both "Trace for procedure" and "Trace for procedure Exit" in one
'           shot. An new parameter added to the end of the current parameters
'           would be used to tell the TrcT command to Indent or Outdent.
'
'           Make "Trace to file" an automatic default that does not require
'           adding "Open/CloseDebugForOutput" commands. "Possibly" eliminate
'           the "debug.print" command as the Immediate window is not nearly
'           as useful for Tracing events as output to a file is.
'
'           Add a button to automatically remove ErrorHandling (by each scope)
'
'           Add a "remove" option that would change "GoTo ExitSub" back to
'           plain "Exit Sub" and remove the "ExitSub:" label.
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
'
' This program (mostly) follows the programming conventions recommended in
' the Microsoft Visual Basic Programmer's Guide  (with the noted exception
' of using three characters for "variable" names).   Examples:
'
' 1. Standard prefixes for control names:
'
'    Control name              Prefix        Example
'    --------------------     ---------   ------------
'    Check box                  chk         chkReadOnly
'    Command button             cmd         cmdOK
'    Frame                      fra         fraContainer
'    Form                       frm         frmEntry
'    Horizontal scroll bar      hsb         hsbVolume
'    Image                      img         imgIcon
'    Label                      lbl         lblHelpMessage
'    List box                   lst         lstPolicyCodes
'    MAPI message               mpm         mpmMessages
'    MAPI session               mps         mpsSession
'    Menu                       mnu         mnuFileOpen
'    Option (radio) button      opt         optFirstPage
'    Picture                    pic         picVGA
'    Text box                   txt         txtLastName
'    Timer                      tmr         tmrAlarm
'    Vertical scroll bar        vsb         vsbRate
'
' 2. Not quite so Standard prefixes for variable names:
'
'    Variable Type             Prefix        Example
'    --------------------     --------    -------------
'    Boolean                   b (bln)     bFound
'    Currency                  c (cur)     cRevenue
'    Date (Time)               d (dat)     dStartDate
'    Double                       dbl      dblTolerance
'    Error                        err      errOrderNum
'    Integer                   i (int)     iQuantity
'    Long                      l (lng)     lDistance
'    Object                    o (obj)     oCurrent
'    Single                       sng      sngAverage
'    String                    s (str)     sFName
'    Variant                   v (var)     vCheckSum
'
'    Note:  Unless it is a reserved word, "passed
'           variables" will not use "any" Prefix.
'
' 3. Standard prefixes to describe variable scope:
'
'    Scope                     Prefix        Example
'    ----------               --------    --------------
'    Global                      g          giPageCount
'    Module  or  form            m          miPageCount
'    Procedure only             none        iLoopCount
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'

Option Explicit
Option Compare Text

Private Enum TypeOfOperation
   InsertTraceForProcedure = 0
   InsertTraceForProcedureExit = 1
   InsertTheErrorHandler = 2
   AddErrorLineNumbers = 3
   RemoveTheErrorLineNumbers = 4
   TurnTraceCommandsOFF = 5
   TurnTraceCommandsON = 6
End Enum

Private Enum ScopeofOperation
   ProjectLevel = 0
   ModuleLevel = 1
   ProcedureLevel = 2
End Enum
 
Public VBI             As VBIDE.VBE        'VBI=VBInstance
Public VBACP           As VBIDE.CodePane   'VBACP=VBI.ActiveCodePane
Public Connect         As Connect

Dim mbCompleted        As Boolean
Dim mbSetTheFocus      As Boolean
Dim mbDoAutoFocus      As Boolean
Dim mbHelpHasFocus     As Boolean
Dim mbAboutHasFocus    As Boolean
Dim mbRemoveHasFocus   As Boolean
Dim mbSettingTabInit   As Boolean
Dim mbHasTraceModule   As Boolean

Dim miTopLn            As Integer
Dim miScopeOfError     As Integer
Dim miScopeOfTrace     As Integer
Dim miSetFocusTo       As Integer
Dim miResetCount       As Integer
Dim miRemoveCount      As Integer
Dim miAlreadyOpenPanes As Integer

Dim msWindow(0 To 99)  As String
Dim msModuleName       As String

Const ModuleName       As String = "frmAddIn"

Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Sub Form_Initialize()
      'TrT "Form_Initialize"

  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
  '  WARNING! Removing this command puts EventTracer into an Endless Loop!  '
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
  
   gbOkToSetFocus = False     'Does not trigger everytime EventTracer runs.
    
End Sub


Private Sub Form_Activate()
      'TrT "OpenDebugForOutput"
      'TrT "Form_Activate"

   gbOkToSetFocus = True      'Triggers everytime EventTracer is run.
                                
   mbDoAutoFocus = False
   gbOpenedAllPanes = False
   gbFormModeChange = False
   miAlreadyOpenPanes = 0

   If Not mbAboutHasFocus Or _
          mbRemoveHasFocus Then mbSetTheFocus = True
   
   ResetInfoWindow
   SetModeFormOnTop
   
End Sub


Private Sub Form_Paint()
      'TrT "Form_Paint", , , , 1

   If gbFormModeChange Then Exit Sub   ' See "SetDefaultFocus" about trigger.
   
   If mbHelpHasFocus Then
      cmdMisc(2).SetFocus              ' Set Focus back on Help/Okay button.
      mbHelpHasFocus = False
      mbSetTheFocus = False
      Exit Sub
   End If
   
   If mbAboutHasFocus Then
      cmdMisc(3).SetFocus              ' Set Focus back on About button.
      mbAboutHasFocus = False
      mbSetTheFocus = False
      Exit Sub
   End If
   
   If mbRemoveHasFocus Then
      cmdTrace(8).SetFocus             ' Set Focus back on Delete button.
      mbRemoveHasFocus = False
      mbSetTheFocus = False
      Exit Sub
   End If

   If mbSetTheFocus Then SetDefaultFocus
   
   If mbDoAutoFocus = True Then

      mbDoAutoFocus = False
      
      If SSTab.Tab = 0 Then            'Event Tracing TAB has Focus
      
         Select Case miSetFocusTo
            Case 1
               cmdTrace(0).SetFocus    'Insert Trace for Procedure
            Case 2
               cmdTrace(1).SetFocus    'Insert Trace for [EXIT]
            Case 3
               cmdTrace(3).SetFocus    'Insert Trace for Variable "Below"
            Case 4
               cmdTrace(4).SetFocus    'Insert OPEN
            Case 5
               cmdTrace(5).SetFocus    'Insert CLOSE
            Case 6
               cmdTrace(6).SetFocus    'Turn OFF
            Case 7
               cmdTrace(7).SetFocus    'Turn ON
            Case 8
               cmdTrace(8).SetFocus    'Delete Traces
            Case 9
               cmdTrace(2).SetFocus    'Insert Trace for Variable "Above"
            Case Else
               cmdTrace(3).SetFocus    'Insert Trace for Var "Below" (default)
         End Select
         
      Else                             'Error Handling TAB has Focus
      
         Select Case miSetFocusTo
            Case 1
               cmdErrors(0).SetFocus   'Add Error Handling
            Case 2
               cmdErrors(1).SetFocus   'Add Line Numbers
            Case 3
               cmdErrors(2).SetFocus   'Remove Line Numbers
         End Select
         
      End If
   
      Exit Sub

   End If

   If Not gbFormModeChange Then
     
      If gbOpenedAllPanes Then
         cmdMisc(1).SetFocus        ' Set Focus on Cancel/Close button.
      Else
         cmdTrace(0).SetFocus       ' Set Focus on ADD for Project.
      End If
      
      gbFormModeChange = False
      
   End If
   
End Sub


Private Sub cmdTrace_Click(Index As Integer)
      'TrT "cmdTrace_Click", True
      'TrV "Index", Index
 
   Dim iCnt          As Integer
   
   If EventTracerIsBeingRunOnItsSelf Then Exit Sub
   
  'EventTracer will operate when the "Active module" is "modEventTracer."
  'However, whenever you switch to another module, the "method of the
  'object" may fail. I think? it's a VB bug, but the solution is to just
  'not allow operations to be "started" from the "modEventTracer" module.
  
   If Not VBI.ActiveCodePane Is Nothing Then
      If VBI.ActiveCodePane.CodeModule.Parent.Name = "modEventTracer" Then
         ReleaseFormOnTopMode
         MsgBox "Cannot perform operations from the " & vbQ & _
                "modEvenTracer" & vbQ & " module itself!", _
                    vbExclamation + vbOKOnly, "Illegal Operation!"
         SetModeFormOnTop
         Exit Sub
      End If
   End If
   
   ResetInfoWindow
   
   Select Case miScopeOfTrace
   
      Case ProjectLevel
      
         Select Case Index
         
            Case 0
            
               OpenAllCodePanes
               ForEntireProject InsertTraceForProcedure
               CloseOpenedPanes
               
            Case 1
            
               OpenAllCodePanes
               ForEntireProject InsertTraceForProcedureExit
               CloseOpenedPanes
               
            Case 6
                       
               OpenAllCodePanes
               ForEntireProject TurnTraceCommandsOFF
               CloseOpenedPanes
               CloseEventTracer
               
            Case 7
            
               OpenAllCodePanes
               ForEntireProject TurnTraceCommandsON
               CloseOpenedPanes
               CloseEventTracer
               
            Case 8
            
               dlgRemove.Top = Me.Top + 600
               dlgRemove.Left = Me.Left + 60
               
               ReleaseFormOnTopMode
               
               For iCnt = 0 To 3
                  cmdMisc(iCnt).Enabled = False
               Next
               
               mbRemoveHasFocus = True
         
               dlgRemove.Show vbModal
      
               For iCnt = 0 To 4
                  If dlgRemove.chkRemove(iCnt).Value = vbChecked Then
                     iCnt = 9
                     Exit For
                  End If
               Next
                      
               If iCnt = 9 Then       'Operator marked an item for removal.
                  OpenAllCodePanes
                  DeleteEventTracerCodeFor "PROJECT"
                  CloseOpenedPanes
               End If
               
               Unload dlgRemove       'Done with this form.
               Set dlgRemove = Nothing
               
               For iCnt = 0 To 3
                  cmdMisc(iCnt).Enabled = True
               Next
               
         End Select
         
      Case ModuleLevel
      
         Select Case Index
      
            Case 0: ForModuleOnly InsertTraceForProcedure
               
            Case 1: ForModuleOnly InsertTraceForProcedureExit
            
            Case 6: ForModuleOnly TurnTraceCommandsOFF
               
            Case 7: ForModuleOnly TurnTraceCommandsON
            
            Case 8
            
               Dim CodePane As VBIDE.CodePane
               Set CodePane = VBI.ActiveCodePane
            
               If CodePane Is Nothing Then          'Bug Fix
                  ReleaseFormOnTopMode
      MsgBox "You must click the mouse in the Module you want to be    " & _
                        vbCrLf & _
             "processed before making your selections to be deleted!   ", _
                        vbExclamation + vbOKOnly, "No module selected"
                  SetModeFormOnTop
                  Exit Sub
               End If
            
               dlgRemove.Top = Me.Top + 600
               dlgRemove.Left = Me.Left + 60
               
               ReleaseFormOnTopMode
               
               For iCnt = 0 To 3
                  cmdMisc(iCnt).Enabled = False
               Next
               
               mbRemoveHasFocus = True
         
               dlgRemove.Show vbModal
      
               For iCnt = 0 To 4
                  If dlgRemove.chkRemove(iCnt).Value = vbChecked Then
                     iCnt = 9
                     Exit For
                  End If
               Next
                      
               If iCnt = 9 Then       'Operator marked an item for removal.
                  DeleteEventTracerCodeFor "MODULE"
               End If
               
               Unload dlgRemove       'Done with this form.
               Set dlgRemove = Nothing
               
               For iCnt = 0 To 3
                  cmdMisc(iCnt).Enabled = True
               Next
         
         End Select
   
      Case ProcedureLevel
      
         Select Case Index
   
            Case 0: InsertTraceProcedure
            
            Case 1: InsertTraceProcedureExit
                     
            Case 2: InsertTraceForVariable "Above"
            
            Case 3: InsertTraceForVariable "Below"
               
            Case 4: InsertTraceOutputOpen
               
            Case 5: InsertTraceOutputClose
                  
         End Select
         
         If Not gbOpenedAllPanes Then If mbCompleted Then CloseEventTracer
         
   End Select
   
End Sub


Private Sub cmdErrors_Click(Index As Integer)
      'TrT "cmdErrors_Click", True
      'TrV "Index", Index
        
   ResetInfoWindow
   
   Select Case miScopeOfError
   
      Case ProjectLevel
      
         Select Case Index
         
            Case 0
               OpenAllCodePanes
               ForEntireProject InsertTheErrorHandler
               CloseOpenedPanes
      
            Case 1
               OpenAllCodePanes
               ForEntireProject AddErrorLineNumbers
               CloseOpenedPanes
      
            Case 2
               OpenAllCodePanes
               ForEntireProject RemoveTheErrorLineNumbers
               CloseOpenedPanes
               
         End Select
         
      Case ModuleLevel
      
         Select Case Index
         
            Case 0: ForModuleOnly InsertTheErrorHandler
      
            Case 1: ForModuleOnly AddErrorLineNumbers
      
            Case 2: ForModuleOnly RemoveTheErrorLineNumbers
               
         End Select
         
      Case ProcedureLevel
      
         Select Case Index
         
            Case 0: InsertErrorHandling
               
            Case 1: InsertLineNumbers
               
            Case 2: RemoveLineNumbers
            
         End Select
     
   End Select
   
   If Not gbOpenedAllPanes Then              ' If the user is only doing one
      If mbCompleted Then CloseEventTracer   ' operation, then Hide the Form
   End If                                    ' if it successfully completed.

End Sub


Private Sub cmdMisc_Click(Index As Integer)
      'TrT "cmdMisc_Click", True
      'TrV "Index", Index
        
   Select Case Index
               
      Case 0                  'Close All Code Windows
      
         CloseAllCodeWindows
         CloseEventTracer

      Case 1
      
         CloseEventTracer     'Close EventTracer
         
      Case 2                  'Help button
      
         If cmdMisc(2).Caption = "Help" Then
            ShowHelp
         Else
            picBox.Visible = False
            
            Me.Caption = "EventTracer"
            Me.Top = Me.Top + 500
            Me.Height = Me.Height - 500
            Me.Left = Me.Left + 1365
            Me.Width = Me.Width - 2730
            cmdMisc(2).Left = cmdMisc(2).Left - 1365
            cmdMisc(2).Top = cmdMisc(2).Top - 500
            Me.Refresh
   
            SSTab.Visible = True
            cmdMisc(0).Visible = True
            cmdMisc(1).Visible = True
            cmdMisc(2).Caption = "Help"
            cmdMisc(3).Visible = True
            mbHelpHasFocus = False
            
            If lblStatus.Caption > "" Then lblStatus.Visible = True
         End If
      
      Case 3                        'About button
      
         FormWinRegPos Me, True     'Saves current window position.
         ReleaseFormOnTopMode
         
         frmAbout.Show vbModal, frmAddIn
         
         SetModeFormOnTop
         mbAboutHasFocus = True
         
   End Select
 
End Sub


Private Sub CloseEventTracer()
      'TrT "CloseEventTracer", , , , 1
 
   On Error Resume Next
   
   miAlreadyOpenPanes = 0    'Variables must be initialized "before" exiting!
   gbOpenedAllPanes = False
   gbFormModeChange = False
   
   ResetInfoWindow
   
   cmdMisc(1).Caption = "Cancel"
   cmdMisc(1).SetFocus        'Keeps Mouse Cursor from going Off to side.
   
   FormWinRegPos Me, True     'Save current window position for next Startup.
   
   Connect.Hide
   
   'TrT "CloseDebugForOutput"
   
End Sub


Private Sub ForEntireProject(TypeOfOperation As Integer)
      'TrT "ForEntireProject", True
      'TrV "TypeOfOperation", TypeOfOperation
        
  'Performs Operations for the entire Project.

   On Error GoTo ErrorHandler

   Dim CodePane     As VBIDE.CodePane
   Dim CodeMod      As CodeModule
   Dim VBP          As VBProject
   Dim VBC          As VBComponent
   
   Dim ProcName     As String
   Dim sCodeLine    As String
   Dim sCaption     As String
   
   Dim iCount       As Integer

   Dim CurrentLn    As Long
   
   Select Case TypeOfOperation
   
      Case InsertTraceForProcedure
         Me.Caption = "Adding commands ..."
         sCaption = "Added "
      
      Case InsertTraceForProcedureExit
         Me.Caption = "Adding " & vbQ & "[Exit]'s" & vbQ & " ..."
         sCaption = "Added "

      Case InsertTheErrorHandler
         Me.Caption = "Adding Error Handler"
         sCaption = "Added "
 
      Case AddErrorLineNumbers
         Me.Caption = "Adding Line Numbers"
         sCaption = "Added Numbers to "
      
      Case RemoveTheErrorLineNumbers
         Me.Caption = "Removing Line Numbers"
         sCaption = "Removed Numbers from "
         
      Case TurnTraceCommandsOFF
         Me.Caption = "Turning Trace OFF"
         sCaption = "Turning Traces OFF"
         
      Case TurnTraceCommandsON
         Me.Caption = "Turning Trace ON"
         sCaption = "Turning Traces ON"
      
   End Select

   Screen.MousePointer = vbHourglass
   iCount = 0
   Me.Refresh
   
   For Each CodePane In VBI.CodePanes
   
      DoEvents
      
      CodePane.Show

      Set CodeMod = CodePane.CodeModule
      
      If CodePane.CodeModule <> "modEventTracer" And _
         CodePane.CodeModule <> "modErrorHandler" Then
         
        'Work backwards from the End because of "Insert" lines.
   
         For CurrentLn = CodePane.CodeModule.CountOfLines To 1 Step -1
   
            sCodeLine = CodePane.CodeModule.Lines(CurrentLn, 1)
            
            If IsFirstLineOfProcedure(sCodeLine) Then
   
               ProcName = CodeMod.ProcOfLine(CurrentLn, vbext_pk_Proc)
               
               If ProcName > "" Then
               
                 'Set the Focus in the current CodePane
                  CodePane.SetSelection CurrentLn, 1, CurrentLn, 1
                  
                  Select Case TypeOfOperation
      
                     Case InsertTraceForProcedure
                          InsertTraceProcedure True
         
                     Case InsertTraceForProcedureExit
                          InsertTraceProcedureExit True
            
                     Case InsertTheErrorHandler
                          InsertErrorHandling True
                          
                     Case AddErrorLineNumbers
                          InsertLineNumbers True
         
                     Case RemoveTheErrorLineNumbers
                          RemoveLineNumbers True
                          
                     Case TurnTraceCommandsOFF
                          TurnTraceCommands "Off"
                          iCount = iCount + miResetCount
                          Exit For
                          
                     Case TurnTraceCommandsON
                          TurnTraceCommands "On"
                          iCount = iCount + miResetCount
                          Exit For
         
                  End Select
                  
                  If mbCompleted Then iCount = iCount + 1
                     
               End If
               
            End If
   
         Next CurrentLn
            
         
         If TypeOfOperation = InsertTheErrorHandler Then
            InsertDeclarations CodePane, CodeMod
         End If
         
      End If
      
   Next CodePane
     
     
   Select Case TypeOfOperation

      Case InsertTraceForProcedure
         sCaption = sCaption & iCount & " Commands"
         AddEventTracerCodeModule sCaption

      Case InsertTraceForProcedureExit
         sCaption = sCaption & iCount & " Commands"
         AddEventTracerCodeModule sCaption
         
      Case InsertTheErrorHandler
         sCaption = sCaption & iCount & " Routines"
         AddErrorHandlerCodeModule sCaption
         
      Case AddErrorLineNumbers
         sCaption = sCaption & iCount & " Procedures"

      Case RemoveTheErrorLineNumbers
         sCaption = sCaption & iCount & " Procedures"
         
      Case TurnTraceCommandsOFF
         TurnTraceModule "Off"
         sCaption = "Turned OFF " & iCount & " Traces"
         
      Case TurnTraceCommandsON
         TurnTraceModule "On"
         sCaption = "Turned ON " & iCount & " Traces"
               
   End Select
   
   If msModuleName > "" Then
      Set VBP = VBI.ActiveVBProject
      Set VBC = VBP.VBComponents(msModuleName)  'Restore original CodePane
      Set CodeMod = VBC.CodeModule              'in case a module was added.
      Set CodePane = CodeMod.CodePane
   
      CodePane.TopLine = miTopLn
      CodePane.Show
   End If
 
   lblStatus.Caption = sCaption
   
   EnableTraceButtonsForScope
   
   Screen.MousePointer = vbDefault
   
   Me.Caption = "EventTracer"
   Me.Refresh
   
   cmdMisc(1).SetFocus

   Exit Sub
 
ErrorHandler:

   ReleaseFormOnTopMode
   HandleError Err.Number, Erl, ModuleName, "ForEntireProject", Outcome
   SetModeFormOnTop
  
   Select Case Outcome
      Case RS:    Resume
      Case RN:    Resume Next
      Case Else:  CloseEventTracer
   End Select
 
End Sub


Private Sub ForModuleOnly(TypeOfOperation As Integer)
      'TrT "ForModuleOnly", True
      'TrV "TypeOfOperation", TypeOfOperation
        
  'Performs Operations for the currently Active Code Pane.

   On Error GoTo ErrorHandler

   Dim CodePane     As VBIDE.CodePane
   Dim CodeMod      As CodeModule
   Dim VBP          As VBProject
   Dim VBC          As VBComponent
   
   Dim ProcName     As String
   Dim sModuleName  As String
   Dim sCodeLine    As String
   Dim sCaption     As String
   
   Dim iCount       As Integer
   
   Dim TopLn        As Long
   Dim CurrentLn    As Long

 
   Select Case TypeOfOperation
   
      Case InsertTraceForProcedure
         Me.Caption = "Adding commands ..."
         sCaption = "Added "
      
      Case InsertTraceForProcedureExit
         Me.Caption = "Adding " & vbQ & "[Exit]'s" & vbQ & " ..."
         sCaption = "Added "
         
      Case InsertTheErrorHandler
         Me.Caption = "Adding Error Handler"
         sCaption = "Added "
      
      Case AddErrorLineNumbers
         Me.Caption = "Adding Line Numbers"
         sCaption = "Added Numbers to "
      
      Case RemoveTheErrorLineNumbers
         Me.Caption = "Removing Line Numbers"
         sCaption = "Removed Numbers from "
             
   End Select
   
   lblStatus.Caption = "Working"
   lblStatus.Visible = True

   Screen.MousePointer = vbHourglass
   iCount = 0
   Me.Refresh
   
   Set CodePane = VBI.ActiveCodePane
      
   If CodePane Is Nothing Then
      ReleaseFormOnTopMode

      MsgBox "Click the mouse in the " & vbQ & "Module" & vbQ & _
             " you want to work on.", _
             vbOKOnly, "No " & vbQ & "Active" & vbQ & " Code Module!"
      
      ResetInfoWindow
      
      SetModeFormOnTop
   
      Exit Sub
   End If
   
   TopLn = CodePane.TopLine
   sModuleName = CodePane.CodeModule.Parent.Name

   Set CodeMod = CodePane.CodeModule

   If CodePane.CodeModule <> "modEventTracer" And _
      CodePane.CodeModule <> "modErrorHandler" Then
   
     'Work backwards from the End because some routines "Insert" lines.

      For CurrentLn = CodePane.CodeModule.CountOfLines To 1 Step -1

         sCodeLine = CodePane.CodeModule.Lines(CurrentLn, 1)
         
         If IsFirstLineOfProcedure(sCodeLine) Then

            ProcName = CodeMod.ProcOfLine(CurrentLn, vbext_pk_Proc)
            
            If ProcName > "" Then
            
               CodePane.SetSelection CurrentLn, 1, CurrentLn, 1 'Sets Focus!
                                         
               Select Case TypeOfOperation
   
                  Case InsertTraceForProcedure
                       InsertTraceProcedure True
      
                  Case InsertTraceForProcedureExit
                       InsertTraceProcedureExit True
         
                  Case InsertTheErrorHandler
                       InsertErrorHandling True
                     
                  Case AddErrorLineNumbers
                       InsertLineNumbers True
      
                  Case RemoveTheErrorLineNumbers
                       RemoveLineNumbers True
                       
                  Case TurnTraceCommandsOFF
                       TurnTraceCommands "Off"
                       Exit For
                       
                  Case TurnTraceCommandsON
                       TurnTraceCommands "On"
                       Exit For
                                   
               End Select
                              
               If mbCompleted Then iCount = iCount + 1
                  
            End If
            
         End If

      Next CurrentLn
      
      If TypeOfOperation = InsertTheErrorHandler Then
         InsertDeclarations CodePane, CodeMod
      End If
      
   End If
      
     
   Select Case TypeOfOperation

      Case InsertTraceForProcedure
         sCaption = sCaption & iCount & " Commands"
         AddEventTracerCodeModule sCaption

      Case InsertTraceForProcedureExit
         sCaption = sCaption & iCount & " Commands"
         AddEventTracerCodeModule sCaption
         
      Case InsertTheErrorHandler
         sCaption = sCaption & iCount & " Routines"
         AddErrorHandlerCodeModule sCaption

      Case AddErrorLineNumbers
         sCaption = sCaption & iCount & " Procedures"

      Case RemoveTheErrorLineNumbers
         sCaption = sCaption & iCount & " Procedures"
         
      Case TurnTraceCommandsOFF
         TurnTraceModule "Off"
         sCaption = lblStatus.Caption

      Case TurnTraceCommandsON
         TurnTraceModule "On"
         sCaption = lblStatus.Caption
              
   End Select
   
   Set VBP = VBI.ActiveVBProject
   Set VBC = VBP.VBComponents(sModuleName)  'Restore original CodePane
   Set CodeMod = VBC.CodeModule             'in case a module was added.
   Set CodePane = CodeMod.CodePane
   
   CodePane.TopLine = TopLn
   CodePane.Show
   
   lblStatus.Caption = sCaption
   
   Screen.MousePointer = vbDefault
   
   Me.Caption = "EventTracer"
   Me.Refresh
   
   cmdMisc(1).SetFocus

   Exit Sub
 
ErrorHandler:

   ReleaseFormOnTopMode
   HandleError Err.Number, Erl, ModuleName, "ForModuleOnly", Outcome
   SetModeFormOnTop
  
   Select Case Outcome
      Case RS:    Resume
      Case RN:    Resume Next
      Case Else:  CloseEventTracer
   End Select
 
End Sub


Private Sub OpenAllCodePanes()
      'TrT "OpenAllCodePanes"
 
  'Open all CodePanes so other procedures can operate on the source code.

   On Error GoTo ErrorHandler

   Dim CodePane      As VBIDE.CodePane
   Dim VBW           As VBIDE.Window
   Dim VBP           As VBProject
   Dim VBC           As VBComponent

   Set VBP = VBI.ActiveVBProject
   Set VBACP = VBI.ActiveCodePane    'Save .ActiveCodePane to be restored.

   miAlreadyOpenPanes = 0
   gbOpenedAllPanes = True
   
   Screen.MousePointer = vbHourglass

   Me.Caption = "Opening CodePanes"
   cmdMisc(1).Caption = "Close"
   lblStatus.Caption = "Working"
   lblStatus.Visible = True

   Me.Refresh
      
   If VBACP Is Nothing Then
      msModuleName = ""
   Else
      miTopLn = VBACP.TopLine
      msModuleName = VBACP.CodeModule.Parent.Name
   End If
   
  'Save the names of all currently visible Windows.

   For Each VBW In VBI.Windows
      If VBW.Type = vbext_wt_CodeWindow Then
         msWindow(miAlreadyOpenPanes) = VBW.Caption
         miAlreadyOpenPanes = miAlreadyOpenPanes + 1
      End If
   Next VBW

   For Each VBC In VBP.VBComponents
   
     'Do not activate the CodePanes for .Res, ActiveX, or .Doc Modules.
     'The code for these modules is not accessible within the VB Project.
   
      If VBP.VBComponents(VBC.Name).Type <> vbext_ct_ResFile Then
         If VBP.VBComponents(VBC.Name).Type <> vbext_ct_ActiveXDesigner Then
            If VBP.VBComponents(VBC.Name).Type <> vbext_ct_RelatedDocument Then

               If VBP.VBComponents(VBC.Name).Type = vbext_ct_VBForm Or _
                  VBP.VBComponents(VBC.Name).Type = vbext_ct_VBMDIForm Then
            
                 'Use the .Show method for Form modules.
                  VBP.VBComponents(VBC.Name).CodeModule.CodePane.Show
               Else
                  VBP.VBComponents(VBC.Name).Activate   ' ... non-Form modules.
               End If
            End If
         End If
      End If
      
   Next VBC

   Me.Refresh

   Exit Sub
 
ErrorHandler:
 
   ReleaseFormOnTopMode
   HandleError Err.Number, Erl, ModuleName, "OpenAllCodePanes", Outcome
   SetModeFormOnTop
 
   Select Case Outcome
      Case RS:    Resume
      Case RN:    Resume Next
      Case Else:  CloseEventTracer
   End Select
 
End Sub


Private Sub CloseOpenedPanes()
      'TrT "CloseOpenedPanes"
        
  'Close the Panes that were opened by EventTracer, so that the IDE
  'goes back to the way it was before the programmer ran EventTracer.
  
   On Error GoTo ErrorHandler

   Dim VBW          As VBIDE.Window
   Dim bCloseIt     As Boolean
   Dim iCnt         As Integer

   If miAlreadyOpenPanes > 0 Then

      For Each VBW In VBI.Windows
         If VBW.Type = vbext_wt_CodeWindow Then
         
            bCloseIt = True
         
            For iCnt = 0 To miAlreadyOpenPanes - 1
               If VBW.Caption = msWindow(iCnt) Then bCloseIt = False
            Next iCnt
            
            If bCloseIt Then VBW.Close
         End If
      Next VBW
      
   End If
   
   If Not VBACP Is Nothing Then
      ReleaseFormOnTopMode
         VBACP.Show       'Returns focus to .ActiveCodePane
         Me.Show
      SetModeFormOnTop
   End If

   miAlreadyOpenPanes = 0
   
   Screen.MousePointer = vbDefault
   
   Exit Sub
 
ErrorHandler:

   ReleaseFormOnTopMode
   HandleError Err.Number, Erl, ModuleName, "CloseOpenedPanes", Outcome
   SetModeFormOnTop
  
   Select Case Outcome
      Case RS:    Resume
      Case RN:    Resume Next
      Case Else:  CloseEventTracer
   End Select
 
End Sub


Private Sub TurnTraceCommands(OnOrOff As String)
      'TrT "TurnTraceCommands", True
      'TrV "", OnOrOff
        
   On Error GoTo ErrorHandler

   Dim CodePane     As VBIDE.CodePane
   Dim CodeMod      As CodeModule

   Dim ProcName     As String
   Dim sCodeLine    As String
   
   Dim iLine        As Integer

   Dim FirstLn      As Long
   Dim LastLn       As Long
   
   If OnOrOff = "Off" Then
      Me.Caption = "Turning Trace OFF"
   Else
      Me.Caption = "Turning Trace ON"
   End If
      
   Me.Refresh
   
   miResetCount = 0
   mbCompleted = False

   DoEvents

   Set CodePane = VBI.ActiveCodePane      'Allows doing "by MODULE"
   Set CodeMod = CodePane.CodeModule


   For iLine = 1 To CodePane.CodeModule.CountOfLines

      sCodeLine = CodePane.CodeModule.Lines(iLine, 1)

      If IsFirstLineOfProcedure(sCodeLine) Then

         ProcName = CodeMod.ProcOfLine(iLine, vbext_pk_Proc)

         If ProcName > "" Then
         
            GetFirstLastLinesOf ProcName, _
                                  CodePane, CodeMod, FirstLn, LastLn
                        
            TurnComments OnOrOff, CodePane, CodeMod, FirstLn, LastLn
            
            iLine = LastLn   'Speeds things up. Otherwise repeats lines
         End If
      End If
      
   Next iLine
      
      
   Select Case OnOrOff
      Case "On":   lblStatus.Caption = miResetCount & " Traces Turned ON"
      Case "Off":  lblStatus.Caption = miResetCount & " Traces Turned OFF"
   End Select
   
   cmdMisc(1).SetFocus
   
   mbCompleted = True

   Exit Sub
 
ErrorHandler:
 
   ReleaseFormOnTopMode
   HandleError Err.Number, Erl, ModuleName, "TurnTraceCommands (On-Off)", _
                           Outcome
   SetModeFormOnTop
 
   Select Case Outcome
      Case RS:    Resume
      Case RN:    Resume Next
      Case Else:  CloseEventTracer
   End Select
 
End Sub


Private Sub TurnComments(OnOrOff As String, _
                         CodePane As Object, CodeMod As Object, _
                         ByVal FirstLn As Long, LastLn As Long)
                         
      'TrT "TurnComments", True
      'TrV "", OnOrOff, True
      'TrV "FirstLn", FirstLn, True
      'TrV "LastLn", LastLn
                 
   On Error GoTo ErrorHandler
   
   Dim Begin        As Long
   Dim EEnd         As Long

   Dim bHit         As Boolean
   
   Dim iNoPos       As Integer
   Dim iTracePos    As Integer
   Dim iQuotePos    As Integer
   
   Dim sCodeLine    As String
   
   Begin = FirstLn    'Do not use "FirstLn" or "LastLn" in CodeMod.Find!
   
   Do
      EEnd = LastLn
      bHit = CodeMod.Find("Trc", Begin, 1, EEnd, -1, False, True)
      
      If Begin < FirstLn Then bHit = False   'Bug fix. Do Not Remove!

      If bHit Then

         sCodeLine = CodePane.CodeModule.Lines(Begin, 1)
                      
         iNoPos = InStr(1, sCodeLine, "No")
         iTracePos = InStr(1, sCodeLine, "Trc")
         iQuotePos = InStr(1, sCodeLine, "'")

         If OnOrOff = "Off" Then             'Comment out the Trc commands.
         
            If iQuotePos = 0 Or iQuotePos > iTracePos Then
            
               If UCase(Left(sCodeLine, 3)) = "Trc" Then
                  sCodeLine = " " & sCodeLine       'Insert a leading space.
               End If

               sCodeLine = Replace(sCodeLine, " Trc", "'Trc", 1, 1)
               CodeMod.ReplaceLine Begin, sCodeLine
               miResetCount = miResetCount + 1
            End If
         Else                                'Un-comment the Trc commands.
         
            If iQuotePos > 0 And iQuotePos < iTracePos Then
         
               If iNoPos = 0 Or iNoPos > iTracePos Then
                  sCodeLine = Replace(sCodeLine, "'", " ", 1, 1)
                  CodeMod.ReplaceLine Begin, sCodeLine
                  miResetCount = miResetCount + 1
               Else
                 'Leave 'NoTrcT  and  'NoTrcV  lines alone.
               End If
            End If
         End If
      End If
      
      Begin = Begin + 1
      
      If Begin >= LastLn Then bHit = False   'Bug Fix. Do not Remove!!
   
   Loop Until Not bHit
   
   Exit Sub
 
ErrorHandler:
 
   ReleaseFormOnTopMode
   HandleError Err.Number, Erl, ModuleName, "TurnComments (On-Off)", Outcome
   SetModeFormOnTop
  
   Select Case Outcome
      Case RS:    Resume
      Case RN:    Resume Next
      Case Else:  CloseEventTracer
   End Select
 
End Sub


Private Sub TurnTraceModule(OnOrOff As String)
      'TrT "TurnTraceModule", True
      'TrV "", OnOrOff

  'This procedure Turns 1 line in the "modEventTracer" ON/OFF. It is called
  'when turning the commands ON/OFF at the Module or Project level. Testing
  'the status of this single line controls the logic for the On/Off "Default".
  
   Dim CodePane     As VBIDE.CodePane
   Dim CodeMod      As CodeModule

   Dim VBP          As VBProject
   Dim VBCTraceMod  As VBComponent
   
   Dim ProcName     As String
   Dim sCodeLine    As String
   
   Dim iLine        As Integer

   Dim FirstLn      As Long
   Dim LastLn       As Long
   
   Set VBP = VBI.ActiveVBProject
   
   Set VBCTraceMod = VBP.VBComponents("modEventTracer")
   
   Set CodeMod = VBCTraceMod.CodeModule
   Set CodePane = CodeMod.CodePane
   
   For iLine = 1 To CodePane.CodeModule.CountOfLines
   
      sCodeLine = CodePane.CodeModule.Lines(iLine, 1)
   
      If IsFirstLineOfProcedure(sCodeLine) Then
   
         ProcName = CodeMod.ProcOfLine(iLine, vbext_pk_Proc)
   
         If ProcName > "" Then
         
            GetFirstLastLinesOf ProcName, _
                                  CodePane, CodeMod, FirstLn, LastLn
                        
            TurnComments OnOrOff, CodePane, CodeMod, FirstLn, LastLn
            
            iLine = LastLn  'Speeds things up. Otherwise repeats lines
                          
         End If
      End If
      
   Next iLine

   Screen.MousePointer = vbDefault
   
   Set CodePane = VBI.ActiveCodePane
   
End Sub


Private Sub DeleteEventTracerCodeFor(ProjectOrModule As String)
      'TrT "DeleteEventTracerCodeFor", True
      'TrV "", ProjectOrModule

   On Error GoTo ErrorHandler
   
   Dim CodePane     As VBIDE.CodePane
   Dim CodeMod      As CodeModule
   Dim ActiveMod    As CodeModule

   Dim ProcName     As String
   Dim sCodeLine    As String
   Dim sCaption     As String

   Dim iCnt         As Integer
   Dim CurrentLn    As Integer
   
   Dim FirstLn      As Long
   Dim LastLn       As Long
   Dim VeryFirstLn  As Long
   
   Dim bRmvNo       As Boolean    '= Remove NoTrcT & NoTrcV
   Dim bRmvVar      As Boolean    '= Remove TrcV
   Dim bRmvExit     As Boolean    '= Remove TrcT [Exit] Procedure
   Dim bRmvTrace    As Boolean    '= Remove TrcT
   Dim bRmvModule   As Boolean    '= Remove modEventTracer.bas
   
   Dim bWaxModule   As Boolean

   If dlgRemove.chkRemove(0).Value = vbChecked Then bRmvTrace = True
   If dlgRemove.chkRemove(1).Value = vbChecked Then bRmvExit = True
   If dlgRemove.chkRemove(2).Value = vbChecked Then bRmvNo = True
   If dlgRemove.chkRemove(3).Value = vbChecked Then bRmvVar = True
   If dlgRemove.chkRemove(4).Value = vbChecked Then bRmvModule = True
   
   miRemoveCount = 0
   
   sCaption = "Removed "
   
   Screen.MousePointer = vbHourglass
    
   Set ActiveMod = VBI.ActiveCodePane.CodeModule
   
   SetModeFormOnTop  'Somewhere this gets unset. Don't have time to fix right.
   
   If bRmvTrace Or bRmvExit Or bRmvNo Or bRmvVar Then
   
      For Each CodePane In VBI.CodePanes
      
         DoEvents
         
         CodePane.Show
         
        '~~~~~~~~~ A quick patch for a cosmetic bug. Fix it right later ~~~~~~
         Me.Caption = "Removing commands..."
         lblStatus.Caption = "Working"     ' Somewhere in this loop
         lblStatus.Visible = True          ' "ResetInfoWindow" must get hit.
         Me.Refresh
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        
         If ProjectOrModule = "MODULE" Then
            If CodePane.CodeModule = ActiveMod Then
               bWaxModule = True
            Else
               bWaxModule = False
            End If
         Else
            bWaxModule = True
         End If
               
         If bWaxModule Then
               
            Set CodeMod = CodePane.CodeModule
             
            If CodeMod.Parent.Name <> "modEventTracer" Then
      
               For CurrentLn = CodePane.CodeModule.CountOfLines To 1 Step -1
      
                  sCodeLine = CodePane.CodeModule.Lines(CurrentLn, 1)
                  
                  If IsFirstLineOfProcedure(sCodeLine) Then
         
                     ProcName = CodeMod.ProcOfLine(CurrentLn, vbext_pk_Proc)
         
                     If ProcName > "" Then
                     
                        GetFirstLastLinesOf ProcName, _
                        CodePane, CodeMod, FirstLn, LastLn, VeryFirstLn
                                    
                        If bRmvTrace Then DeleteTraceCommands _
                        CodePane, CodeMod, FirstLn, LastLn
     
                        If bRmvExit Then DeleteTraceExitCommands _
                        CodePane, CodeMod, FirstLn, LastLn
   
                        If bRmvNo Then DeleteNoTraceCommands _
                        CodePane, CodeMod, FirstLn, LastLn
   
                        If bRmvVar Then DeleteTraceVarCommands _
                        CodePane, CodeMod, FirstLn, LastLn
                        
                       'Speed up processing by not repeating lines.
                        CurrentLn = VeryFirstLn
   
                     End If
                  End If
         
               Next CurrentLn
               
            End If
         End If
         
      Next CodePane
         
      sCaption = sCaption & " " & miRemoveCount & " Commands"

   End If
   
   If bRmvModule Then
      DeleteEventTracerCodeModule sCaption
   End If
   
   If ProjectOrModule = "MODULE" Then
      Set CodeMod = ActiveMod.CodePane.CodeModule
      Set CodePane = CodeMod.CodePane
      CodePane.Show
   End If
   
   Screen.MousePointer = vbDefault
   
   lblStatus.Caption = sCaption
   lblStatus.Visible = True
   
   For iCnt = 0 To 3
      cmdMisc(iCnt).Enabled = True
   Next
   
   Me.Caption = "EventTracer"
   Me.Refresh
   
   cmdMisc(1).SetFocus
 
   Exit Sub
 
ErrorHandler:
 
   ReleaseFormOnTopMode
   HandleError Err.Number, Erl, ModuleName, "DeleteEventTracerCodeFor", Outcome
   SetModeFormOnTop
  
   Select Case Outcome
      Case RS:    Resume
      Case RN:    Resume Next
      Case Else:  CloseEventTracer
   End Select
 
End Sub


Private Sub DeleteEventTracerCodeModule(sCaption As String)
      'TrT "DeleteEventTracerCodeModule"
        
  'Delete the EventTracer module.
   
   On Error GoTo ErrorHandler

   Dim VBC    As VBComponent
   Dim VBAP   As VBProject
   
   Set VBAP = VBI.ActiveVBProject
   
   For Each VBC In VBAP.VBComponents

       If VBAP.VBComponents(VBC.Name).Name = "modEventTracer" Then
       
          VBAP.VBComponents.Remove VBC

          If InStr(1, sCaption, "Commands") Then
             sCaption = sCaption & ", "
          End If
             sCaption = sCaption & "1 Module"
          Exit For

       End If

   Next VBC
   
   Exit Sub
 
ErrorHandler:
 
   ReleaseFormOnTopMode
   HandleError Err.Number, Erl, ModuleName, "DeleteEventTracerCodeModule", _
                           Outcome
   SetModeFormOnTop
  
   Select Case Outcome
      Case RS:    Resume
      Case RN:    Resume Next
      Case Else:  CloseEventTracer
   End Select
 
End Sub


Private Sub DeleteTraceCommands _
           (CodePane As Object, CodeMod As Object, _
            ByVal FirstLn As Long, ByVal LastLn As Long)
            
      'TrT "DeleteTraceCommands", True
      'TrV "FirstLn", FirstLn, True
      'TrV "lastln", LastLn
                             
  'Delete lines containing "Trc" commands.
  
   On Error GoTo ErrorHandler

   Dim sCodeLine    As String

   Dim iNoPos       As Integer
   Dim iVarPos      As Integer
   Dim iTracePos    As Integer
   Dim iExitPos     As Integer
   
   Dim Begin        As Long
   Dim EEnd         As Long

   Dim bHit         As Boolean
   Dim bEraseIt     As Boolean
               
   Begin = FirstLn    'Do not use "FirstLn" or "LastLn" in CodeMod.Find!
   
   Do
      EEnd = LastLn
      bHit = CodeMod.Find("Trc", Begin, 1, EEnd, -1, False, True)
      
      If Begin < FirstLn Then bHit = False  'VB bug fix. Do not Remove!!!

      If bHit Then

         sCodeLine = CodePane.CodeModule.Lines(Begin, 1)

         iNoPos = InStr(1, sCodeLine, "No")
         iVarPos = InStr(1, sCodeLine, "TrcV")
         iTracePos = InStr(1, sCodeLine, "Trc ")
         iExitPos = InStr(1, sCodeLine, "[Exit]")

         bEraseIt = False            'Default to DO NOT ERASE.
         
         If iVarPos = 0 Then         'TrcV and NoTrc? and [Exit]
            If iExitPos = 0 Then     'are not removed in this procedure.
               If iNoPos = 0 Or iNoPos > iTracePos Then
                  bEraseIt = True
               End If
            End If
         End If
         
         If bEraseIt Then       'Has met all the criteria for deletion.
            RemoveLine CodePane, CodeMod, Begin, LastLn
         Else
            Begin = Begin + 1   'Check the next line in the Procedure.
         End If
      End If
      
      If Begin >= LastLn Then bHit = False   'Bug Fix. Do not Remove!!

   Loop Until Not bHit
 
   Exit Sub
 
ErrorHandler:
 
   ReleaseFormOnTopMode
   HandleError Err.Number, Erl, ModuleName, "DeleteTraceCommands", Outcome
   SetModeFormOnTop
  
   Select Case Outcome
      Case RS:    Resume
      Case RN:    Resume Next
      Case Else:  CloseEventTracer
   End Select
 
End Sub


Private Sub DeleteTraceExitCommands _
           (CodePane As Object, CodeMod As Object, _
            ByVal FirstLn As Long, ByVal LastLn As Long)
            
      'TrT "DeleteTraceExitCommands", True
      'TrV "FirstLn", FirstLn, True
      'TrV "LastLn", LastLn
                             
  'Delete lines containing "TrcT [Exit] Procedure" commands.
  
   On Error GoTo ErrorHandler

   Dim sCodeLine    As String

   Dim iNoPos       As Integer
   Dim iTracePos    As Integer
   Dim iExitPos     As Integer
   
   Dim Begin        As Long
   Dim EEnd         As Long

   Dim bHit         As Boolean
   Dim bEraseIt     As Boolean
                 
   Begin = FirstLn    'Do not use "FirstLn" or "LastLn" in CodeMod.Find!
   
   Do
      EEnd = LastLn
      bHit = CodeMod.Find("Trc", Begin, 1, EEnd, -1, False, True)
      
      If Begin < FirstLn Then bHit = False  'VB bug fix. Do not Remove!!!

      If bHit Then

         sCodeLine = CodePane.CodeModule.Lines(Begin, 1)

         iNoPos = InStr(1, sCodeLine, "No")
         iTracePos = InStr(1, sCodeLine, "Trc ")
         iExitPos = InStr(1, sCodeLine, "[Exit]")

         bEraseIt = False
         
         If iNoPos = 0 Or iNoPos > iTracePos Then  'No TrcT and TrcV are
            If iExitPos <> 0 Then bEraseIt = True  'not done in this proc.
         End If
         
         If bEraseIt Then       'Has met all the criteria for deletion.
            If Left(sCodeLine, 4) = "Exit" Then
               sCodeLine = Left(sCodeLine, InStr(5, sCodeLine, ":"))
               CodeMod.ReplaceLine Begin, sCodeLine
            Else
               RemoveLine CodePane, CodeMod, Begin, LastLn
            End If
         Else
            Begin = Begin + 1   'Check the next line in the Procedure.
         End If
      End If
      
      If Begin >= LastLn Then bHit = False   'Bug Fix. Do not Remove!!

   Loop Until Not bHit
 
   Exit Sub
 
ErrorHandler:
 
   ReleaseFormOnTopMode
   HandleError Err.Number, Erl, ModuleName, "DeleteTraceExitCommands", Outcome
   SetModeFormOnTop
  
   Select Case Outcome
      Case RS:    Resume
      Case RN:    Resume Next
      Case Else:  CloseEventTracer
   End Select
 
End Sub


Private Sub DeleteNoTraceCommands _
           (CodePane As Object, CodeMod As Object, _
            ByVal FirstLn As Long, ByVal LastLn As Long)
            
      'TrT "DeleteNoTraceCommands", True
      'TrV "FirstLn", FirstLn, True
      'TrV "LastLn", LastLn
                           
  'Delete lines containing "NoTrc?" or "No Trc?" commands.
  
   On Error GoTo ErrorHandler

   Dim sCodeLine    As String
   
   Dim iNoPos       As Integer
   Dim iTracePos    As Integer
   
   Dim Begin        As Long
   Dim EEnd         As Long

   Dim bHit         As Boolean
   Dim bEraseIt     As Boolean
   
   Begin = FirstLn    'Do not use "FirstLn" or "LastLn" in CodeMod.Find!

   Do
      EEnd = LastLn
      bHit = CodeMod.Find("Trc", Begin, 1, EEnd, -1, False, True)
      
      If Begin < FirstLn Then bHit = False  'VB bug fix. Do not Remove!!!
      
      If bHit Then

         sCodeLine = CodePane.CodeModule.Lines(Begin, 1)
      
         iNoPos = InStr(1, sCodeLine, "No")
         iTracePos = InStr(1, sCodeLine, "Trc")
         
         bEraseIt = True
   
         If iNoPos = 0 Or iNoPos > iTracePos Then
            bEraseIt = False
         End If
   
         If bEraseIt Then
            RemoveLine CodePane, CodeMod, Begin, LastLn
         Else
            Begin = Begin + 1   'Check the next line in the Procedure.
         End If
         
      End If
      
      If Begin >= LastLn Then bHit = False   'Bug Fix. Do not Remove!!

   Loop Until Not bHit
 
   Exit Sub
 
ErrorHandler:
 
   ReleaseFormOnTopMode
   HandleError Err.Number, Erl, ModuleName, "DeleteNoTraceCommands", Outcome
   SetModeFormOnTop
  
   Select Case Outcome
      Case RS:    Resume
      Case RN:    Resume Next
      Case Else:  CloseEventTracer
   End Select
 
End Sub


Private Sub DeleteTraceVarCommands _
           (CodePane As Object, CodeMod As Object, _
            ByVal FirstLn As Long, ByVal LastLn As Long)
            
      'TrT "DeleteTraceVarCommands", True
      'TrV "FirstLn", FirstLn, True
      'TrV "LastLn", LastLn
                                
  'Delete lines containing "TrcV" or "NoTrcV" commands.
  
   On Error GoTo ErrorHandler
   
   Dim Begin       As Long
   Dim EEnd        As Long

   Dim bHit        As Boolean
   
   Begin = FirstLn   'Do not use "FirstLn" in CodeMod.Find as changes it.

   Do
      EEnd = LastLn
      bHit = CodeMod.Find("TrcV", Begin, 1, EEnd, -1, False, True)
      
      If Begin < FirstLn Then bHit = False  'VB bug fix. Do not Remove!!!

      If bHit Then
         RemoveLine CodePane, CodeMod, Begin, LastLn
      Else
         Begin = Begin + 1         'Check the next line in the Procedure.
      End If
      
      If Begin >= LastLn Then bHit = False   'Bug Fix. Do not Remove!!

   Loop Until Not bHit
 
   Exit Sub
 
ErrorHandler:
 
   ReleaseFormOnTopMode
   HandleError Err.Number, Erl, ModuleName, "DeleteTraceVarCommands", Outcome
   SetModeFormOnTop
  
   Select Case Outcome
      Case RS:    Resume
      Case RN:    Resume Next
      Case Else:  CloseEventTracer
   End Select
 
End Sub


Private Sub InsertTraceProcedure(Optional InBatchMode As Boolean)
      'TrT "InsertTraceProcedure", True
      'TrV "InBatchMode", InBatchMode
 
   On Error GoTo ErrorHandler

   Dim CodePane     As VBIDE.CodePane
   Dim CodeMod      As CodeModule
   
   Dim ProcName     As String
   Dim sTraceName   As String
   Dim sCodeLine    As String
   
   Dim bHit         As Boolean
   
   Dim Indent       As Integer

   Dim FirstLn      As Long
   Dim LastLn       As Long
   Dim StartLn      As Long
   Dim StartCol     As Long
   Dim WorkLn       As Long
   
   mbCompleted = False

   Determine ProcName, CodePane, CodeMod, StartLn, StartCol
   
   If ProcName = "" Then
      ShowSetFocusMessage "EventTracer - INSERT Trace for Procedure"
      Exit Sub            'The Cursor/Focus is not in a Procedure!
   End If
   
   If CodePane.CodeModule = "modEventTracer" Then Exit Sub 'Don't Run on self!
    
   GetFirstLastLinesOf ProcName, CodePane, CodeMod, FirstLn, LastLn
   
   sTraceName = "TrcT " & vbQ & ProcName
   
  'Check to see if the procedure already has this TrcT command.
   
   bHit = AlreadyHas(sTraceName, CodePane, CodeMod, FirstLn, LastLn, WorkLn)
   
   If bHit = True Then
      sCodeLine = CodePane.CodeModule.Lines(WorkLn, 1)
      If InStr(1, sCodeLine, " [Exit]") <> 0 Then bHit = False
   End If

   If bHit = False Then
   
      FirstLn = CodeMod.ProcBodyLine(ProcName, vbext_pk_Proc)
      
     'The sCodeLine must be obtained "before" AdjustForLineContinuation
     
      sCodeLine = CodePane.CodeModule.Lines(FirstLn, 1)
      
      AdjustForLineContinuation ProcName, CodeMod, FirstLn

      RemoveBlankLines "After", FirstLn, CodePane, CodeMod, Indent
       
     'Insert the "TrcT" command for the Procedure.
     'Insert a blank line above and below the TrcT command and
     'insert spaces to match indentation of the next line of code.

      CodeMod.InsertLines FirstLn + 1, " "
      
      Indent = InStr(4, sCodeLine, ProcName) - 7 'Align directly under ProcName
      
      If Indent < 0 Then Indent = 0      'No Optionals (Public, Private, etc.)
      
      If InStr(1, ProcName, "_Click") <> 0 Then
            CodeMod.InsertLines FirstLn + 1, Space(Indent) & _
            "TrcT " & vbQ & ProcName & "()" & vbQ & ", , , , 1"
      Else
         If InStr(1, ProcName, "Form_") <> 0 Then
            CodeMod.InsertLines FirstLn + 1, Space(Indent) & _
            "TrcT " & vbQ & ProcName & " (" & _
            CodeMod.Parent.Name & ")" & vbQ & ", , , , 1"
         Else
            CodeMod.InsertLines FirstLn + 1, Space(Indent) & _
            "TrcT " & vbQ & ProcName & vbQ
         End If
      End If
             
     'CodeMod.InsertLines FirstLn + 1, " " 'If you want a blank line "before"
      
      If Not InBatchMode Then    'Place the Cursor at the start of line.
         AddEventTracerCodeModule    'Add module if not already present.
         If Indent = 0 Then Indent = 1     'A bug fix for the Next line.
         CodePane.SetSelection FirstLn + 2, Indent, FirstLn + 2, Indent
         CodePane.Show
      End If
      
      mbCompleted = True
      
   End If

   Exit Sub
 
ErrorHandler:

   ReleaseFormOnTopMode
   HandleError Err.Number, Erl, ModuleName, "InsertTraceProcedure", Outcome
   SetModeFormOnTop
  
   Select Case Outcome
      Case RS:    Resume
      Case RN:    Resume Next
      Case Else:  CloseEventTracer
   End Select
 
End Sub


Private Sub InsertTraceProcedureExit(Optional InBatchMode As Boolean)
      'TrT "InsertTraceProcedureExit", True
      'TrV "InBatchMode", InBatchMode
        
   On Error GoTo ErrorHandler

   Dim CodePane     As VBIDE.CodePane
   Dim CodeMod      As CodeModule
   
   Dim ProcName     As String
   Dim sCodeLine    As String
   
   Dim Indent       As Integer
   
   Dim bHit         As Boolean
   Dim bAddSubLabel As Boolean
   Dim bAddFunLabel As Boolean
   Dim bAddBlankLn  As Boolean

   Dim FirstLn      As Long
   Dim LastLn       As Long
   Dim StartLn      As Long
   Dim EndLn        As Long
   Dim StartCol     As Long
   Dim StartLn1     As Long
   Dim StartLn2     As Long
   Dim EndLn1       As Long
   Dim EndLn2       As Long
   Dim WorkLn       As Long
   
   mbCompleted = False

   Determine ProcName, CodePane, CodeMod, StartLn, StartCol
   
   If ProcName = "" Then
      ShowSetFocusMessage "EventTracer - INSERT Trace for Procedure EXIT"
      Exit Sub            'The Cursor/Focus is not in a Procedure!
   End If
   
   If CodePane.CodeModule = "modEventTracer" Then Exit Sub 'Don't Run on self!
   
   GetFirstLastLinesOf ProcName, CodePane, CodeMod, FirstLn, LastLn
   
   StartLn = FirstLn      'Preserve values
   StartLn1 = FirstLn
   StartLn2 = FirstLn
   EndLn = LastLn
   EndLn1 = LastLn
   EndLn2 = LastLn
   
  'Check to see if the procedure already has this Trace command.
   
   bHit = CodeMod.Find("[Exit]", FirstLn, 1, LastLn, -1, False, True)
   
   If bHit = False Then
   
     'Test to see if an "ExitProcedureType:" label needs to be added.

      bAddFunLabel = CodeMod.Find("Exit Function", _
                                   StartLn1, 1, EndLn1, -1, False, True)
      bAddSubLabel = CodeMod.Find("Exit Sub", _
                                   StartLn2, 1, EndLn2, -1, False, True)
                                   
      If bAddFunLabel = False Then
         If AlreadyHas("ExitFunction:", CodePane, CodeMod, FirstLn, LastLn) Then
            bAddFunLabel = True        'Label gets Removed and Added back.
         End If
      End If
      
      If bAddSubLabel = False Then
         If AlreadyHas("ExitSub:", CodePane, CodeMod, FirstLn, LastLn) Then
            bAddSubLabel = True        'Label gets Removed and Added back.
         End If
      End If
      
      AdjustForErrorHandling ProcName, CodePane, CodeMod, FirstLn, LastLn
            
     'Remove all of the blank lines at the End of the Procedure.
     
      RemoveBlankLines "Before", LastLn, CodePane, CodeMod, Indent
      
      bAddBlankLn = True
      
     'Insert the "TrcT" command at the End of the Procedure with blank lines
     'and insert spaces to match indentation of the next to Last line of code.
   
      sCodeLine = CodePane.CodeModule.Lines(LastLn, 1)
      
      If InStr(1, sCodeLine, "End") <> 0 Then
         Indent = 6
      Else
         CodeMod.InsertLines LastLn, " "
      End If
      
      If InStr(1, sCodeLine, " Function") Then
      
        'Use Trace for Variable for Functions to show the return value.
         
         sCodeLine = "ExitFunction:    TrcV " & vbQ & ProcName & _
                   " [Exit]" & vbQ & "," & ProcName & ", , , , , 1"
         
         If AlreadyHas("ExitFunction:", _
                       CodePane, CodeMod, StartLn, EndLn, WorkLn) Then
            CodeMod.ReplaceLine WorkLn, sCodeLine
            bAddBlankLn = False
            EndLn = WorkLn - 1 'Don't Fix "Exit Function" past this line.
         Else
            CodeMod.InsertLines LastLn, sCodeLine
            EndLn = LastLn - 1 'Don't Fix "Exit Function" past this line.
         End If

         ReplaceExitsWithGoTos ProcName, _
                     CodePane, CodeMod, StartLn, EndLn, "Exit Function"
                     
      Else
      
         sCodeLine = "ExitSub:    TrcT " & vbQ & ProcName & _
                   " [Exit]" & vbQ & ", , , , , 1"
         
         If AlreadyHas("ExitSub:", _
                       CodePane, CodeMod, StartLn, EndLn, WorkLn) Then
            CodeMod.ReplaceLine WorkLn, sCodeLine
            bAddBlankLn = False
            EndLn = WorkLn - 1 'Don't fix "Exit Sub" past this line.
         Else
            CodeMod.InsertLines LastLn, sCodeLine
            EndLn = LastLn - 1 'Don't fix "Exit Sub" past this line.
         End If
         
         ReplaceExitsWithGoTos ProcName, _
                     CodePane, CodeMod, StartLn, EndLn, "Exit Sub"
      End If
      
      
      If bAddBlankLn Then CodeMod.InsertLines LastLn, " "
      
      If Not InBatchMode Then   'Place the Cursor at the start of line.
         AddEventTracerCodeModule   'Add module if not already present.
         If Indent = 0 Then Indent = 1    'A bug fix for the Next line.
         CodePane.SetSelection LastLn + 1, Indent, LastLn + 1, Indent
         CodePane.Show
      End If
      
      mbCompleted = True
      
   End If
   
   Exit Sub
 
ErrorHandler:
 
   ReleaseFormOnTopMode
   HandleError Err.Number, Erl, ModuleName, "InsertTraceProcedureExit", Outcome
   SetModeFormOnTop
  
   Select Case Outcome
      Case RS:    Resume
      Case RN:    Resume Next
      Case Else:  CloseEventTracer
   End Select
 
End Sub


Private Sub ReplaceExitsWithGoTos( _
            ProcName As String, CodePane As Object, CodeMod As Object, _
            FirstLn As Long, LastLn As Long, ExitType As String)
            
      'TrT "ReplaceExitWithGoTos", True
      'TrV "ProcName", ProcName, True
      'TrV "FirstLn", FirstLn, True
      'TrV "LastLn", LastLn
            
  'Replaces "Exit Sub/Function" with "GoTo ExitSub" or "GoTo ExitFunction"
            
   On Error GoTo ErrorHandler
   
   Dim WorkLn       As Long
   Dim StartLn      As Long
   Dim EndLn        As Long
   
   Dim sCodeLine    As String
   
   EndLn = LastLn                'Don't work with the passed values!
   StartLn = FirstLn
   
   For WorkLn = StartLn + 1 To EndLn - 1
    
      sCodeLine = CodePane.CodeModule.Lines(WorkLn, 1)
         
      If InStr(1, sCodeLine, ExitType) <> 0 Then
         
         If ExitType = "Exit Sub" Then
            sCodeLine = Replace( _
            sCodeLine, "Exit Sub", "GoTo ExitSub", 1, 1)
         Else
            sCodeLine = Replace( _
            sCodeLine, "Exit Function", "GoTo ExitFunction", 1, 1)
         End If
         
         CodeMod.ReplaceLine WorkLn, sCodeLine
                
      End If

   Next WorkLn

   Exit Sub
 
ErrorHandler:
 
   ReleaseFormOnTopMode
   HandleError Err.Number, Erl, ModuleName, "ReplaceExitsWithGoTos", Outcome
   SetModeFormOnTop
  
   Select Case Outcome
      Case RS:    Resume
      Case RN:    Resume Next
      Case Else:  CloseEventTracer
   End Select

End Sub


Private Sub InsertTraceForVariable(ByRef InsertWhere As String)
      'TrT "InsertTraceForVariable", True
      'TrV "InsertWhere", InsertWhere

   On Error GoTo ErrorHandler

   Dim CodePane     As VBIDE.CodePane
   Dim CodeMod      As CodeModule
   
   Dim ProcName     As String
   Dim sCodeLine    As String
   Dim sVarName     As String
   Dim sIndentWrk   As String
   
   Dim Indent       As Integer
   Dim NoIndent     As Integer
   Dim iDigits      As Integer
   
   Dim bExit        As Boolean
   Dim bDoOffset    As Boolean
   
   Dim FirstLn      As Long
   Dim LastLn       As Long
   Dim StartLn      As Long
   Dim StartCol     As Long

   mbCompleted = False
  
   Determine ProcName, CodePane, CodeMod, StartLn, StartCol
   
   If ProcName = "" Then
      ShowSetFocusMessage "EventTracer - INSERT Trace for VARIABLE", _
           "Click on the Variable you want to be Traced (to set Focus)"
      Exit Sub                'The Cursor/Focus is not in a Procedure!
   End If
   
   sCodeLine = CodePane.CodeModule.Lines(StartLn, 1)
   
   sIndentWrk = sCodeLine  'See if we need to compensate for line numbers.
   
   If InStr(1, "123456789", Left(sIndentWrk, 1)) <> 0 Then
      iDigits = InStr(1, sIndentWrk, " ")
      sIndentWrk = Right(sCodeLine, Len(sCodeLine) - iDigits)
   End If
   
   Indent = Len(sCodeLine) - Len(LTrim(sIndentWrk))

   sVarName = ExtractVarName(sCodeLine, StartCol)
   
   If sVarName = "" Then
      ShowSetFocusMessage "EventTracer - Insert Trace for Variable", _
           "Click on the Variable you want to be Traced (to set Focus)"
      Exit Sub                           'The Cursor is not on a Word!
   End If
   
   bDoOffset = NotOnAContinuedLine(StartLn, ProcName, CodePane, CodeMod)
   
   If InsertWhere = "Above" Then         'Go up 1 line
      RemoveBlankLines "Before", StartLn, CodePane, CodeMod, NoIndent
      StartLn = StartLn - 1
   Else
      AdjustForLineContinuation ProcName, CodeMod, StartLn
      RemoveBlankLines "After", StartLn, CodePane, CodeMod, NoIndent
   End If
   

   If sVarName = "Exit" Then
      If InStr(1, sCodeLine, " Sub") Then
         bExit = True
      ElseIf InStr(1, sCodeLine, " Function") Then
         bExit = True
      End If
   End If
   
   If bExit Then         'Not inserting for Variable, but rather an [Exit]
   
      CodeMod.InsertLines StartLn + 1, Space(Indent) & "TrcT " & vbQ & _
      ProcName & " [Exit]" & vbQ & ", , , , , 1"     'Add Blank line.
      
   Else
   
     'Insert the "TrcV" command for the selected Variable.
     'Insert a blank line above and below the TrcV command and
     'insert spaces to match indentation of the next line of code.
   
      CodeMod.InsertLines StartLn + 1, " "

      If bDoOffset Then
         CodeMod.InsertLines StartLn + 1, Space(Indent) & _
                "TrcV " & vbQ & sVarName & vbQ & ", " & sVarName & _
                ", , , 3"
      Else
         CodeMod.InsertLines StartLn + 1, Space(Indent) & _
                "TrcV " & vbQ & sVarName & vbQ & ", " & sVarName
      End If

   End If
   
   CodeMod.InsertLines StartLn + 1, " "   'Insert blank line just "Above"

   AddEventTracerCodeModule        'Add module if not already present.
   
  'Place the Cursor at the start of the new line.
  
   If Indent = 0 Then Indent = 1   'A Bug fix for the next line.
   
   CodePane.SetSelection StartLn + 2, Indent, StartLn + 2, Indent
   CodePane.Show
          
   mbCompleted = True
 
   Exit Sub
 
ErrorHandler:
 
   ReleaseFormOnTopMode
   HandleError Err.Number, Erl, ModuleName, "InsertTraceForVariable", Outcome
   SetModeFormOnTop
  
   Select Case Outcome
      Case RS:    Resume
      Case RN:    Resume Next
      Case Else:  CloseEventTracer
   End Select
 
End Sub


Private Sub InsertTraceOutputOpen()
      'TrT "InsertTraceOutputOpen"
   
   On Error GoTo ErrorHandler
 
   Dim CodePane     As VBIDE.CodePane
   Dim CodeMod      As CodeModule
   
   Dim ProcName     As String
   Dim sCodeLine    As String
   Dim sVarName     As String
      
   Dim Indent       As Integer
   
   Dim FirstLn      As Long
   Dim LastLn       As Long
   Dim StartLn      As Long
   Dim StartCol     As Long
   
   mbCompleted = False
   
   Determine ProcName, CodePane, CodeMod, StartLn, StartCol
   
   If ProcName = "" Then  'The Cursor/Focus is not in a Procedure!
      ShowSetFocusMessage "EventTracer - INSERT Trace to File - OPEN", _
        "Click on the Line where you want the command to be Inserted."
      Exit Sub
   End If
   
   If AlreadyHas("OpenDebugForOutput", _
                  CodePane, CodeMod, FirstLn, LastLn) Then
      Exit Sub
   End If
   
   sCodeLine = CodePane.CodeModule.Lines(StartLn, 1)
   
   AdjustForLineContinuation ProcName, CodeMod, StartLn

   RemoveBlankLines "After", StartLn, CodePane, CodeMod, Indent

  'Insert the "TrcT OpenDebugForOutput" command at the cursor.
  'Insert a blank line below the TrcT command and insert
  'spaces to match indentation of the next line of code.
   
   CodeMod.InsertLines StartLn + 1, " "
   
   CodeMod.InsertLines StartLn + 1, Space(Indent) & _
          "TrcT " & vbQ & "OpenDebugForOutput" & vbQ
          
   AddEventTracerCodeModule    'Add module if not already present.
   
  'Place the Cursor at the start of the new line.
  
   If Indent = 0 Then Indent = 1     'A bug fix for the Next line.

   CodePane.SetSelection StartLn + 2, Indent, StartLn + 2, Indent
   CodePane.Show
   
   mbCompleted = True
 
   Exit Sub
 
ErrorHandler:
 
   ReleaseFormOnTopMode
   HandleError Err.Number, Erl, ModuleName, "InsertTraceOutputOpen", Outcome
   SetModeFormOnTop
  
   Select Case Outcome
      Case RS:    Resume
      Case RN:    Resume Next
      Case Else:  CloseEventTracer
   End Select
 
End Sub


Private Sub InsertTraceOutputClose()
      'TrT "InsertTraceOutputClose"
   
   On Error GoTo ErrorHandler
 
   Dim CodePane     As VBIDE.CodePane
   Dim CodeMod      As CodeModule
   
   Dim ProcName     As String
   Dim sCodeLine    As String
   Dim sVarName     As String
      
   Dim Indent       As Integer
   
   Dim FirstLn      As Long
   Dim LastLn       As Long
   Dim StartLn      As Long
   Dim StartCol     As Long
   
   mbCompleted = False
   
   Determine ProcName, CodePane, CodeMod, StartLn, StartCol

   If ProcName = "" Then  'The Cursor/Focus is not in a Procedure!
      ShowSetFocusMessage "EventTracer - INSERT Trace to File - CLOSE", _
         "Click on the Line where you want the command to be Inserted."
      Exit Sub
   End If
   
   If AlreadyHas("CloseDebugForOutput", _
                  CodePane, CodeMod, FirstLn, LastLn) Then
      Exit Sub
   End If
   
   sCodeLine = CodePane.CodeModule.Lines(StartLn, 1)
   
   AdjustForLineContinuation ProcName, CodeMod, StartLn
 
   RemoveBlankLines "After", StartLn, CodePane, CodeMod, Indent

  'Insert the "TrcT CloseDebugForOutput" command at the cursor.
  'Insert a blank line below the TrcT command and insert
  'spaces to match indentation of the next line of code.
   
   CodeMod.InsertLines StartLn + 1, " "
   
   CodeMod.InsertLines StartLn + 1, Space(Indent) & _
          "TrcT " & vbQ & "CloseDebugForOutput" & vbQ
          
   AddEventTracerCodeModule    'Add module if not already present.
   
  'Place the Cursor at the start of the new line.
  
   If Indent = 0 Then Indent = 1     'A bug fix for the Next line.

   CodePane.SetSelection StartLn + 2, Indent, StartLn + 2, Indent
   CodePane.Show
    
   mbCompleted = True
 
   Exit Sub
 
ErrorHandler:
 
   ReleaseFormOnTopMode
   HandleError Err.Number, Erl, ModuleName, "InsertTraceOutputClose", Outcome
   SetModeFormOnTop
  
   Select Case Outcome
      Case RS:    Resume
      Case RN:    Resume Next
      Case Else:  CloseEventTracer
   End Select
 
End Sub


Private Sub InsertErrorHandling(Optional InBatchMode As Boolean)
      'TrT "InsertErrorHandling", True
      'TrV "InBatchMode", InBatchMode
        
   On Error GoTo ErrorHandler

   Dim VBC            As VBComponent
   Dim CodePane       As VBIDE.CodePane
   Dim CodeMod        As CodeModule
   
   Dim ProcName       As String
   Dim sCaption       As String
   Dim sExitType      As String
   Dim sTraceName     As String
   
   Dim bExists        As Boolean
   Dim bExitSubLabel  As Boolean
   Dim bExitFunLabel  As Boolean
   
   Dim Ndent          As Integer
   Dim Ndent3         As Integer
   
   Dim lCnt           As Long
   Dim lOptLine       As Long
   Dim lLastComment   As Long
   
   Dim FirstLn        As Long
   Dim LastLn         As Long
   Dim StartLn        As Long
   Dim StartCol       As Long
   Dim BeginLn        As Long
   Dim WorkLn         As Long
   
   mbCompleted = False

   Determine ProcName, CodePane, CodeMod, StartLn, StartCol
   
   If ProcName = "" Then  'The Cursor/Focus is not in a Procedure!
      ShowSetFocusMessage "EventTracer - Insert ERROR Handling Routine"
      Exit Sub
   End If
   
   If CodePane.CodeModule = "modErrorHandler" Then Exit Sub 'Don't Run on self!
   
   GetFirstLastLinesOf ProcName, CodePane, CodeMod, FirstLn, LastLn
   
   sTraceName = "TrcT " & vbQ & ProcName
   
   BeginLn = FirstLn   'Preserving values
   WorkLn = LastLn
   
   bExitSubLabel = CodeMod.Find("ExitSub:", _
                                 BeginLn, 1, WorkLn, -1, False, True)
   BeginLn = FirstLn
   WorkLn = LastLn
   
   bExitFunLabel = CodeMod.Find("ExitFunction:", _
                                 BeginLn, 1, WorkLn, -1, False, True)
   
  'Set Exit Type for either "Exit Sub" or "Exit Function" based on last line.
  
   sExitType = _
   Replace((CodePane.CodeModule.Lines(LastLn, 1)), "End ", "Exit ")
   
   If AlreadyHas("On Error", CodePane, CodeMod, FirstLn, LastLn) Then
      
      If Not InBatchMode Then
         ShowSetFocusMessage "EventTracer - Insert ERROR Handling Routine", _
         "Procedure or Function ALREADY HAS an Error Handling Routine!"
      End If
      
      Exit Sub
   End If

  'Check to see if the procedure has a "TrcT ProcName" command.
   
   bExists = AlreadyHas(sTraceName, CodePane, CodeMod, FirstLn, LastLn, WorkLn)
    
   If bExists = True Then
      sTraceName = CodePane.CodeModule.Lines(WorkLn, 1)
      If InStr(1, sTraceName, " [Exit]") <> 0 Then bExists = False
   End If
   

   If bExists = True Then
   
     'Insert the "On Error GoTo" below the "Trc" command(s) after FirstLn
   
      Do
         WorkLn = WorkLn + 1
         sTraceName = CodePane.CodeModule.Lines(WorkLn, 1)
      Loop Until InStr(1, sTraceName, "Trc") = 0
      
      FirstLn = WorkLn - 1
      
   Else
      FirstLn = CodeMod.ProcBodyLine(ProcName, vbext_pk_Proc)
    
      AdjustForLineContinuation ProcName, CodeMod, FirstLn
   End If
   
   RemoveBlankLines "Before", LastLn, CodePane, CodeMod, Ndent
   
   Ndent = 3                     'Hardcoded at my own personal preference.
   Ndent3 = Ndent + 3
   
  'Insert the Error Routine at the End of the Procedure with blank lines
  'and insert spaces to match Ndentation of the next to Last line of code.
  
  'Code is inserted in "Reverse order", but assembles itself in forward order.

   CodeMod.InsertLines LastLn, " "
   CodeMod.InsertLines LastLn, Space(Ndent) & "End Select"
   CodeMod.InsertLines LastLn, Space(Ndent3) & "Case Else:  ExitProject"
   CodeMod.InsertLines LastLn, Space(Ndent3) & _
                              "Case GZ:    On Error GoTo 0:  'GoTo Label:"
   CodeMod.InsertLines LastLn, Space(Ndent3) & "Case RL:   'Resume Label:"
   CodeMod.InsertLines LastLn, Space(Ndent3) & "Case RN:    Resume Next"
   CodeMod.InsertLines LastLn, Space(Ndent3) & "Case RS:    Resume"
   CodeMod.InsertLines LastLn, Space(Ndent) & "Select Case Outcome"
   CodeMod.InsertLines LastLn, " "
   CodeMod.InsertLines LastLn, Space(Ndent) & "HandleError Err.number, " _
                               & "Erl, ModuleName , " _
                               & vbQ & ProcName & vbQ & ", Outcome"
   CodeMod.InsertLines LastLn, " "
   CodeMod.InsertLines LastLn, "ErrorHandler:"
   CodeMod.InsertLines LastLn, " "
   CodeMod.InsertLines LastLn, Space(Ndent) & sExitType
   CodeMod.InsertLines LastLn, " "
     
  'Insert the "On Error GoTo" command for the Procedure.
  'Insert a blank line above and below the command and
  'insert spaces to match Ndentation of the next line of code.
  
   RemoveBlankLines "After", FirstLn, CodePane, CodeMod, Ndent

   CodeMod.InsertLines FirstLn + 1, " "
    
   CodeMod.InsertLines FirstLn + 1, Space(Ndent) & _
           "On Error GoTo ErrorHandler"
           
   CodeMod.InsertLines FirstLn + 1, " "
   
   If Not InBatchMode Then
      InsertDeclarations CodePane, CodeMod    'Add if not already present.
      AddErrorHandlerCodeModule               'Add if not already present.
      
     'Return the cursor to the "ErrorHandler:" label
     
      CodePane.SetSelection LastLn + 7, 14, LastLn + 7, 14
      CodePane.Show
   End If
   
   mbCompleted = True
   
   Exit Sub
   
ErrorHandler:
   
   ReleaseFormOnTopMode
   HandleError Err.Number, Erl, ModuleName, "InsertErrorHandling", Outcome
   SetModeFormOnTop
  
   Select Case Outcome
      Case RS:    Resume
      Case RN:    Resume Next
      Case Else:  CloseEventTracer
   End Select
    
End Sub


Private Sub InsertDeclarations(CodePane As Object, CodeMod As Object)
      'TrT "InsertDeclarations"
                              
   Dim FirstLn   As Long
   Dim LastLn    As Long
   Dim Indent    As Integer
   
   If Not AlreadyHas("Const ModuleName ", CodePane, CodeMod, 1, LastLn) Then
                      
      FirstLn = InsertionPointFor("Const", CodePane, CodeMod)
                                                            
      RemoveBlankLines "After", FirstLn, CodePane, CodeMod, Indent

      CodeMod.InsertLines FirstLn + 1, " "
      CodeMod.InsertLines FirstLn + 1, Space(Indent) & _
                                      "Const ModuleName As String = " _
                                       & vbQ & CodeMod.Parent.Name & vbQ
      CodeMod.InsertLines FirstLn + 1, " "
   End If
   
   If Not AlreadyHas("Option Explicit", CodePane, CodeMod, 1, LastLn) Then
                      
      FirstLn = InsertionPointFor("Option Explicit", CodePane, CodeMod)
      
      RemoveBlankLines "After", FirstLn, CodePane, CodeMod, Indent
      
      CodeMod.InsertLines FirstLn + 1, " "
      CodeMod.InsertLines FirstLn + 1, "Option Explicit"
      CodeMod.InsertLines FirstLn + 1, " "
   End If

End Sub


Private Sub InsertLineNumbers(Optional InBatchMode As Boolean)
      'TrT "InsertLineNumbers", True
      'TrV "InBatchMode", InBatchMode
        
   On Error GoTo ErrorHandler

   Dim CodePane         As VBIDE.CodePane
   Dim CodeMod          As CodeModule
   
   Dim ProcName         As String
   Dim sCodeLine        As String
   Dim sTestLine        As String
    
   Dim SkipCnt          As Integer
   Dim StartNumber      As Integer
   
   Dim bInConstruct     As Boolean
   Dim bInsertNumber    As Boolean
   Dim bMoreToLine      As Boolean
   Dim bSkipNextCase    As Boolean
   Dim bStarted         As Boolean
   Dim bDoIndent        As Boolean
   
   Dim FirstLn          As Long
   Dim LastLn           As Long
   Dim WorkLn           As Long
   Dim StartLn          As Long
   Dim StartCol         As Long
   Dim EndLn            As Long
   
   Screen.MousePointer = vbHourglass
   
   mbCompleted = False

   Determine ProcName, CodePane, CodeMod, StartLn, StartCol
                               
   If ProcName = "" Then  'The cursor/focus is not in a Procedure.
      ShowSetFocusMessage "EventTracer - Add Procedure Line Numbers"
      Exit Sub
   End If
   
   If CodePane.CodeModule = "modErrorHandler" Then Exit Sub 'Don't Run on self!

   GetFirstLastLinesOf ProcName, CodePane, CodeMod, FirstLn, LastLn
      
   AdjustForErrorHandling ProcName, CodePane, CodeMod, FirstLn, LastLn
   
   If Not AlreadyHas("On Error GoTo ", _
                      CodePane, CodeMod, FirstLn, LastLn) Then
                      
     'Operator may want to add line numbers "without" having an error
     'handler in the procedure. If so, must adjust so as to not number
     'the very first and last lines of the procedure.
                      
      LastLn = LastLn - 1
      FirstLn = FirstLn + 1
      
     'The next line is a bug fix for a procedure with 1 line of code.
     
      If FirstLn > LastLn Then FirstLn = LastLn
      
   End If
   
  'Get a count of the lines which DO NOT get line numbering.
  
   SkipCnt = 0
   
   For WorkLn = FirstLn To LastLn Step 1
   
      sCodeLine = CodePane.CodeModule.Lines(WorkLn, 1)
      
      sTestLine = Trim(sCodeLine)
      
      If sTestLine > "" Then
      
        'Test for existing line numbers
        
         If InStr(1, "1", Left(sTestLine, 1)) <> 0 Then Exit Sub
         
         If sTestLine = "ErrorHandler:" Then Exit For
           
        'Test for lines that DO NOT get numbered.
        
         If UCase(Left(sTestLine, 4)) = "#IF " Then
            bInConstruct = True
            SkipCnt = SkipCnt + 1
         End If
         
         If bInConstruct = False Then
      
            If Left(sTestLine, 1) = "'" Then SkipCnt = SkipCnt + 1
            If Left(sTestLine, 4) = "Rem " Then SkipCnt = SkipCnt + 1
            If Left(sTestLine, 4) = "Dim " Then SkipCnt = SkipCnt + 1
            If Left(sTestLine, 7) = "Static " Then SkipCnt = SkipCnt + 1
            If Left(sTestLine, 8) = "On Error" Then SkipCnt = SkipCnt + 1
            If Left(sTestLine, 11) = "Select Case" Then SkipCnt = SkipCnt + 1
            
            If InStr(2, sTestLine, ":") Then      'Tests for a LabelLine:
               If Left(sCodeLine, 1) <> " " Then
                  If InStr(2, sTestLine, " ") Then
                     If InStr(2, sTestLine, " ") > _
                        InStr(2, sTestLine, ":") _
                        Then SkipCnt = SkipCnt + 1
                  Else
                    If Right(sTestLine, 1) = ":" Then SkipCnt = SkipCnt + 1
                  End If
               End If
            End If
                
            If InStr(1, sCodeLine, " _") > 0 Then SkipCnt = SkipCnt + 1
         Else
            If UCase(Left(sTestLine, 7)) = "#END IF" Then bInConstruct = False
         End If
         
      End If

   Next WorkLn
   
   If (LastLn - FirstLn) - SkipCnt > 90 Then
      StartNumber = 1000
   Else
      StartNumber = 100
   End If
   
   bStarted = False
   bSkipNextCase = False             'Set to catch "Select Case" blocks.
   
   EndLn = LastLn   'Bug fix for "empty procedures" without any code in them.
   
   For WorkLn = FirstLn To LastLn
    
      sCodeLine = CodePane.CodeModule.Lines(WorkLn, 1)
      
      If Not bMoreToLine Then
      
         sTestLine = Trim(sCodeLine)
      
         If sTestLine > "" Then
         
            If sTestLine = "ErrorHandler:" Then Exit For
         
            bInsertNumber = True
            
           'Cannot put a line number on first case statement in Select Case.
           
            If Left(sTestLine, 12) = "Select Case " Then bSkipNextCase = True
            
            If Left(sTestLine, 5) = "Case " And bSkipNextCase = True Then
               bInsertNumber = False
               bSkipNextCase = False     'Reset for next "Select Case" block
               
              'Indent the line to match the other lines being numbered.
               If StartNumber < 1000 Then
                  CodeMod.ReplaceLine WorkLn, Space(4) & sCodeLine
               Else
                  CodeMod.ReplaceLine WorkLn, Space(5) & sCodeLine
               End If
            End If
            
            If UCase(Left(sTestLine, 4)) = "#IF " Then bInConstruct = True
         
            If bInConstruct = True Then bInsertNumber = False
            
            If Left(sTestLine, 1) = "'" Then bInsertNumber = False
            If Left(sTestLine, 4) = "Rem " Then bInsertNumber = False
            If Left(sTestLine, 4) = "Dim " Then bInsertNumber = False
            If Left(sTestLine, 7) = "Static " Then bInsertNumber = False
            If Left(sTestLine, 8) = "On Error" Then bInsertNumber = False
            
            If InStr(2, sTestLine, ":") Then        'Tests for a LabelLine:
               If Left(sCodeLine, 1) <> " " Then
                  If InStr(2, sTestLine, " ") Then
                     If InStr(2, sTestLine, " ") > _
                        InStr(2, sTestLine, ":") _
                        Then bInsertNumber = False
                  Else
                    If Right(sTestLine, 1) = ":" Then bInsertNumber = False
                  End If
               End If
            End If
            
            If bInsertNumber Then
               bStarted = True
               sCodeLine = StartNumber & " " & sCodeLine
               CodeMod.ReplaceLine WorkLn, sCodeLine
               
               StartNumber = StartNumber + 10
               EndLn = WorkLn
            ElseIf bStarted = True Then
            
              'Indent comments too (if numbering has already started).
              
               bDoIndent = False
            
               If Left(sTestLine, 1) = "'" Then bDoIndent = True
               If Left(sTestLine, 1) = "#" Then bDoIndent = True
               If Left(sTestLine, 4) = "Rem " Then bDoIndent = True
               If Left(sTestLine, 4) = "Dim " Then bDoIndent = True
               If Left(sTestLine, 7) = "Static " Then bDoIndent = True
              
               If bDoIndent = True Then
                  If StartNumber < 1000 Then
                     CodeMod.ReplaceLine WorkLn, Space(4) & sCodeLine
                  Else
                     CodeMod.ReplaceLine WorkLn, Space(5) & sCodeLine
                  End If
               End If
            End If
            
            If UCase(Left(sTestLine, 7)) = "#END IF" Then bInConstruct = False
            
            If InStr(1, sCodeLine, " _") > 0 Then bMoreToLine = True
         End If
      Else
      
         If EndLn = WorkLn - 1 Then       'Indent to match number
            If StartNumber < 1000 Then
               CodeMod.ReplaceLine WorkLn, Space(4) & sCodeLine
            Else
               CodeMod.ReplaceLine WorkLn, Space(5) & sCodeLine
            End If
            EndLn = WorkLn
         End If

         If InStr(1, sCodeLine, " _") > 0 Then
            bMoreToLine = True
         Else
            bMoreToLine = False
         End If
      End If

   Next WorkLn
   
   CodePane.SetSelection EndLn, 1, EndLn, 1   'Put cursor on the last line.
   
   mbCompleted = True
   
   Screen.MousePointer = vbDefault
   
   Exit Sub
   
ErrorHandler:
   
   ReleaseFormOnTopMode
   HandleError Err.Number, Erl, ModuleName, "InsertLineNumbers", Outcome
   SetModeFormOnTop
  
   Select Case Outcome
      Case RS:    Resume
      Case RN:    Resume Next
      Case Else:  CloseEventTracer
   End Select
 
End Sub


Private Sub RemoveLineNumbers(Optional InBatchMode As Boolean)
      'TrT "RemoveLineNumbers", True
      'TrV "InBatchMode", InBatchMode
        
   On Error GoTo ErrorHandler

   Dim CodePane         As VBIDE.CodePane
   Dim CodeMod          As CodeModule
   
   Dim ProcName         As String
   Dim sCodeLine        As String
   Dim sChar            As String
   Dim sDigits          As String
   Dim sTestLine        As String
   
   Dim bFixCase         As Boolean
   Dim bMoreToLine      As Boolean
   Dim bStarted         As Boolean
   Dim bRemoveIndent    As Boolean

   Dim iCnt             As Integer
   Dim iMakeNulls       As Integer
   Dim iLineLen         As Integer
   
   Dim lSavedTopLine    As Long
   Dim lSpaces          As Long
   Dim WorkLn           As Long
   Dim StartLn          As Long
   Dim StartCol         As Long
   Dim EndLn            As Long
   
   Screen.MousePointer = vbHourglass
   
   iMakeNulls = 0
   
   bFixCase = False
   bStarted = False
   mbCompleted = False

   Determine ProcName, CodePane, CodeMod, StartLn, StartCol
   
   If ProcName = "" Then  'The Cursor/Focus is not in a Procedure!
      ShowSetFocusMessage "EventTracer - REMOVE Numbers from Lines"
      Exit Sub
   End If
   
   lSavedTopLine = CodePane.TopLine
     
   GetFirstLastLinesOf ProcName, CodePane, CodeMod, StartLn, EndLn

   For WorkLn = StartLn + 1 To EndLn - 1
    
      sCodeLine = CodePane.CodeModule.Lines(WorkLn, 1)
      
      If Not bMoreToLine Then
            
         If Trim(sCodeLine) > "" Then
         
            If Trim(sCodeLine) = "ErrorHandler:" Then Exit For

            If InStr(1, "123456789", Left(sCodeLine, 1)) <> 0 Then
            
               bStarted = True
               
               sDigits = ""
               
               For iCnt = 1 To 10
               
                   sChar = Mid(sCodeLine, iCnt, 1)
                  
                   If InStr(1, "01234 56789", sChar) Then
                      sDigits = sDigits & sChar
                   End If
                  
                   If InStr(1, "0123456789", sChar) = 0 Then Exit For
                  
               Next iCnt
               
               CodeMod.ReplaceLine WorkLn, Replace(sCodeLine, sDigits, "")
               
               mbCompleted = True
               
               If iMakeNulls = 0 Then iMakeNulls = Len(sDigits)
               
               If InStr(1, sCodeLine, " _") > 0 Then bMoreToLine = True

               If InStr(1, sCodeLine, "Select Case ") > 0 Then bFixCase = True
               
            ElseIf bFixCase = True Then
            
              'Adjust indentation of the first "Case" following "Select Case"
              
               If InStr(1, sCodeLine, "Case ") > 0 Then
                  bFixCase = False
                 
                 'CodeMod.ReplaceLine WorkLn, _
                          Replace(sCodeLine, Space(iMakeNulls), "") 'No workie.
                          
                  iLineLen = Len(RTrim(sCodeLine))      'Doing it the ugly way.
                  sCodeLine = Trim(sCodeLine)
                  lSpaces = (iLineLen - Len(sCodeLine)) - iMakeNulls
                  
                  If lSpaces < 0 Then lSpaces = 0
                  
                  CodeMod.ReplaceLine WorkLn, Space(lSpaces) & sCodeLine
                              
               End If
               
            ElseIf bStarted = True Then
            
              'Remove Indents for comments (if numbering has already started).
              
               sTestLine = Trim(sCodeLine)

               bRemoveIndent = False
              
               If Left(sTestLine, 1) = "#" Then bRemoveIndent = True
               If Left(sTestLine, 1) = "'" Then bRemoveIndent = True
               If Left(sTestLine, 4) = "Rem " Then bRemoveIndent = True
               If Left(sTestLine, 4) = "Dim " Then bRemoveIndent = True
               If Left(sTestLine, 7) = "Static " Then bRemoveIndent = True
              
               If bRemoveIndent = True Then
                 'CodeMod.ReplaceLine WorkLn, _
                          Replace(sCodeLine, Space(iMakeNulls), "") 'No Workie!
                          
                  iLineLen = Len(RTrim(sCodeLine))      'Doing it the ugly way.
                  sCodeLine = Trim(sCodeLine)
                  lSpaces = (iLineLen - Len(sCodeLine)) - iMakeNulls
                  
                  If lSpaces < 0 Then lSpaces = 0
                  
                  CodeMod.ReplaceLine WorkLn, Space(lSpaces) & sCodeLine
               End If
            
            End If
         End If
      Else

        'CodeMod.ReplaceLine WorkLn, _
                 Replace(sCodeLine, Space(iMakeNulls), "") 'Doesn't work!

         iLineLen = Len(RTrim(sCodeLine))         'Doing it the ugly way.
         sCodeLine = Trim(sCodeLine)
         lSpaces = (iLineLen - Len(sCodeLine)) - iMakeNulls
         
         If lSpaces < 0 Then lSpaces = 0
         
         CodeMod.ReplaceLine WorkLn, Space(lSpaces) & sCodeLine
                 
         If InStr(1, sCodeLine, " _") > 0 Then
            bMoreToLine = True
         Else
            bMoreToLine = False
         End If
      
      End If

   Next WorkLn
   
   If Not InBatchMode Then
      CodePane.TopLine = lSavedTopLine     'Reposition text to where it was.
      CodePane.SetSelection lSavedTopLine, 1, lSavedTopLine, 1
   End If
   
   Screen.MousePointer = vbDefault
 
   Exit Sub
 
ErrorHandler:
 
   ReleaseFormOnTopMode
   HandleError Err.Number, Erl, ModuleName, "RemoveLineNumbers", Outcome
   SetModeFormOnTop
  
   Select Case Outcome
      Case RS:    Resume
      Case RN:    Resume Next
      Case Else:  CloseEventTracer
   End Select
 
End Sub


Private Sub AddErrorHandlerCodeModule(Optional ByRef sCaption As String)
      'TrT "AddErrorHandlerCodeModule"
        
   On Error GoTo ErrorHandler
 
  'Adds "modErrorHandler.bas" if it does not already exist.
  
   Dim VBP         As VBProject
   Dim VBC         As VBComponent
   Dim CodePane    As VBIDE.CodePane
   Dim CodeMod     As CodeModule

   Set VBC = VBI.ActiveVBProject.VBComponents("modErrorHandler")
   
   Screen.MousePointer = vbDefault
   
   If VBC Is Nothing Then
   
      CreateModuleFor "modErrorHandler"

      If InStr(1, sCaption, "Routine") <> 0 Then
         sCaption = sCaption & ", 1 Module"
      Else
         sCaption = "1 Module"
      End If

      Me.Refresh
   
      cmdMisc(1).SetFocus
      
     'Add "Option Explicit" separately in case the VB IDE did it automatically.
   
      Set VBP = VBI.ActiveVBProject
      Set VBC = VBP.VBComponents("modErrorHandler")
      Set CodeMod = VBC.CodeModule
      Set CodePane = CodeMod.CodePane
      
      If Not CodePane Is Nothing Then
         Set CodeMod = CodePane.CodeModule
         InsertDeclarations CodePane, CodeMod   'This will check before adding.
      End If

   End If
 
   Exit Sub
 
ErrorHandler:
 
   ReleaseFormOnTopMode
   HandleError Err.Number, Erl, ModuleName, "AddErrorHandlerCodeModule", _
                           Outcome
   SetModeFormOnTop
 
   Select Case Outcome
      Case RS:    Resume
      Case RN:    Resume Next
      Case Else:  CloseEventTracer
   End Select
 
End Sub


Private Sub AddEventTracerCodeModule(Optional sCaption As String)
      'TrT "AddEventTracerCodeModule"
        
   On Error GoTo ErrorHandler
 
  'Adds "modEventTracer.bas" if it does not already exist.
  
   Dim VBP         As VBProject
   Dim VBC         As VBComponent
   Dim CodePane    As VBIDE.CodePane
   Dim CodeMod     As CodeModule

   Set VBC = VBI.ActiveVBProject.VBComponents("modEventTracer")
   
   Screen.MousePointer = vbDefault
   
   If VBC Is Nothing Then
   
      CreateModuleFor "modEventTracer"
      
      sCaption = sCaption & ", 1 Module"
   
      Me.Refresh
   
      cmdMisc(1).SetFocus
      
     'Add "Option Explicit" separately in case the VB IDE did it automatically.
  
      Set VBP = VBI.ActiveVBProject
      Set VBC = VBP.VBComponents("modEventTracer")
      Set CodeMod = VBC.CodeModule
      Set CodePane = CodeMod.CodePane
      
      If Not CodePane Is Nothing Then
         Set CodeMod = CodePane.CodeModule
         InsertDeclarations CodePane, CodeMod   'This will check before adding.
      End If
      
   Else
      TurnTraceModule "ON"
   End If

   Exit Sub
 
ErrorHandler:
 
   ReleaseFormOnTopMode
   HandleError Err.Number, Erl, ModuleName, "AddEventTracerCodeModule", Outcome
   SetModeFormOnTop
 
   Select Case Outcome
      Case RS:    Resume
      Case RN:    Resume Next
      Case Else:  CloseEventTracer
   End Select
 
End Sub


Private Sub CreateModuleFor(ModulesName As String)
      'TrT "CreateModuleFor", True
      'TrV "", ModulesName
        
   On Error GoTo ErrorHandler
  
   Dim lLeft      As Long
   Dim lWidth     As Long
   Dim lTop       As Long
   Dim lHeight    As Long
      
   Dim sTempFile  As String
   
  'If EventTracer is not compiled, then it is being debugged in the VB IDE.
   
   If IsCompiled Then
      sTempFile = ".\tmpModule.txt"
   Else
      sTempFile = App.Path & "\tmpModule.txt"
   End If
   
   
   FF = FreeFile                      'Get the next available File number.
   
   Open sTempFile For Output As #FF
   
   If ModulesName = "modEventTracer" Then
      CreateTraceTempTxtModule
      mbHasTraceModule = True
   Else
      CreateErrorTempTxtModule
   End If
   
   Close #FF
   
  'Add the above text file to the Active project as a new code module.
   
   If Dir(sTempFile) > "" Then           'Do not execute this block
      Dim VBP As VBProject               'of code if the creation of
      Dim VBC As VBComponent             'the file was not successful.
      Dim VBW As VBIDE.Window

      Set VBP = VBI.ActiveVBProject
      Set VBC = VBP.VBComponents(ModulesName)
      
      If VBC Is Nothing Then
         Set VBC = VBP.VBComponents.Add(vbext_ct_StdModule)
         VBC.Name = ModulesName
         VBC.InsertFile sTempFile
         
        'Delete the temporary text file for module and Close it's CodePane.
        'It gets saved as "modEventTracer.bas" when the project gets saved.
        
         Kill sTempFile
         
         For Each VBW In VBI.Windows
            If InStr(VBW.Caption, ModulesName) Then
              'Do nothing. For some reason "Not InStr" would not work here.
            Else
              If VBW.Type = vbext_wt_CodeWindow Then
                 If VBW.Visible = True Then
                    lLeft = VBW.Left              'Get the dimensions of
                    lWidth = VBW.Width            'a "visible" Code Window.
                    lTop = VBW.Top
                    lHeight = VBW.Height
                    Exit For
                  End If
               End If
             End If
         Next VBW
 
         For Each VBW In VBI.Windows
            If InStr(VBW.Caption, ModulesName) Then
               VBW.Left = lLeft              'Set the dimensions to
               VBW.Width = lWidth            'match a current Code Window.
               VBW.Top = lTop
               VBW.Height = lHeight          'and Hide it so the
               VBW.Close                     'programmer doesn't have to.
               Exit For
            End If
         Next VBW
      End If
   End If
   
   Exit Sub
   
ErrorHandler:

   ReleaseFormOnTopMode
   
   If Err.Number = 50132 Then
      
      MsgBox Err & ":Error in CreateModuleFor." & vbCrLf & _
            "Error Message: " & Err.Description & vbCrLf & vbCrLf & _
                   "Unable to add " & vbQ & ModulesName & ".bas" & vbQ & _
                   " (module) to the Project", _
                   vbCritical, "Warning"
      
      SetModeFormOnTop
   Else
   
      HandleError Err.Number, Erl, ModuleName, "CreateModuleFor", Outcome
      SetModeFormOnTop
  
      Select Case Outcome
         Case RS:    Resume
         Case RN:    Resume Next
         Case Else:  CloseEventTracer
      End Select

   End If

End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
'         Miscellaneous "work" routines used by the above Procedures.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'

Private Function AlreadyHas(CodeSnippet As String, _
                 CodePane As Object, CodeMod As Object, _
                 FirstLn As Long, LastLn As Long, _
                 Optional WorkLn As Long) As Boolean
                 
           'TrT "AlreadyHas", True
           'TrV "", CodeSnippet, True
           'TrV "FirstLn", FirstLn, True
           'TrV "LastLn", LastLn, True
           'TrV "WorkLn", WorkLn, True
                 
  'Tests to see if a Code Snippet already exists in a Procedure.
         
   On Error GoTo ErrorHandler
   
   Dim sError       As String
   
   Dim StartLn      As Long
   Dim EndLn        As Long
   
   EndLn = LastLn           'Do not work with the passed values!
   StartLn = FirstLn
   
   AlreadyHas = _
   CodeMod.Find(CodeSnippet, StartLn, 1, EndLn, -1, False, True)
   
   If IsMissing(WorkLn) Then
   Else
      WorkLn = StartLn      'Passes back the line number text was found on.
   End If
   
 'TrV "", AlreadyHas
  
   Exit Function
 
ErrorHandler:
 
   ReleaseFormOnTopMode
   HandleError Err.Number, Erl, ModuleName, "AlreadyHas", Outcome
   SetModeFormOnTop
  
   Select Case Outcome
      Case RS:    Resume
      Case RN:    Resume Next
      Case Else:  CloseEventTracer
   End Select
 
End Function


Private Function ExtractVarName _
                (ByVal CodeLine As String, ByVal StartCol As Long)
           'TrT "ExtraceVarName", True
           'TrV "", CodeLine, True
           'TrV "StartCol", StartCol, True
                
  'Extracts the name of the Variable that the user has Clicked on.
  
   On Error GoTo ErrorHandler

   Dim sChar      As String * 1
   Dim sVarName   As String
   
   Dim iCnt       As Integer
   Dim iEnd       As Integer
                                                                                                                   
   sVarName = ""
   
  'Search from current column to first character of the variable name.
  
   For iCnt = StartCol To 1 Step -1
   
      sChar = Mid(CodeLine, iCnt, 1)

      If InStr(1, " '(", sChar) Then Exit For
      
      sVarName = sChar & sVarName
   
   Next iCnt
   
  'Search from current column to End character of the variable name.
  
   If Trim(sVarName) > "" Then StartCol = StartCol + 1
   iEnd = 255 - Len(sVarName)
  
   For iCnt = StartCol To iEnd
   
      sChar = Mid(CodeLine, iCnt, 1)
      
      If sChar = ":" Then
         sVarName = ""       'It is a label, not a variable.
         Exit For
      End If

      If InStr(1, " ()", sChar) Then Exit For
      
      sVarName = sVarName & sChar
   
   Next iCnt
                              
   ExtractVarName = Trim(sVarName)
   
 'TrV "", ExtractVarName
   
   Exit Function
 
ErrorHandler:
 
   ReleaseFormOnTopMode
   HandleError Err.Number, Erl, ModuleName, "ExtractVarName", Outcome
   SetModeFormOnTop
  
   Select Case Outcome
      Case RS:    Resume
      Case RN:    Resume Next
      Case Else:  CloseEventTracer
   End Select
 
End Function


Private Sub AdjustForLineContinuation _
           (ProcName As String, CodeMod As Object, StartLn As Long)
      'TrT "AdjustForLineContinuation", True, , 3
      'TrV "ProcName", ProcName, True
      'TrV "StartLn Before", StartLn, True
                   
  'Checks (and adjusts) for a line continuation character on the current line.
   
   Do While InStr(1, CodeMod.Lines(StartLn, 1), " _") > 0
      StartLn = StartLn + 1
   Loop
            
      'TrV "StartLn After", StartLn
End Sub


Private Function NotOnAContinuedLine _
                (ByVal CurrentLn As Long, ProcName As String, _
                 CodePane As Object, CodeMod As Object) As Boolean
           'TrT "NotOnAContinuedLine", True
           'TrV "CurrentLn", CurrentLn, True
           'TrV "ProcName", ProcName, True
 
  'Tests to see if previous Trc? command continues printing on same line.
  
   On Error GoTo ErrorHandler
                 
   Dim FirstLn      As Long
   Dim LastLn       As Long
   Dim lCnt         As Long

   Dim iCommaPos    As Integer
   
   Dim sCodeLine    As String
   
   NotOnAContinuedLine = True
   
   GetFirstLastLinesOf ProcName, CodePane, CodeMod, FirstLn, LastLn
   
   For lCnt = CurrentLn - 1 To FirstLn + 1 Step -1
   
      sCodeLine = CodePane.CodeModule.Lines(lCnt, 1)

      If InStr(1, sCodeLine, "TrcT ") <> 0 Then
         iCommaPos = InStr(8, sCodeLine, ",")
         If iCommaPos = 0 Then Exit For
      End If
      
      If InStr(1, sCodeLine, "TrcV ") <> 0 Then
         iCommaPos = InStr(8, sCodeLine, ",")
         iCommaPos = InStr(iCommaPos + 2, sCodeLine, ",")
         If iCommaPos = 0 Then Exit For
      End If

      If iCommaPos <> 0 Then
         If Mid(sCodeLine, iCommaPos + 2, 4) = "True" Then
            NotOnAContinuedLine = False
         End If
         Exit For
      End If
   
   Next lCnt
   
 'TrV "NotOnAContinuedLine", NotOnAContinuedLine
   
   Exit Function
 
ErrorHandler:
 
   ReleaseFormOnTopMode
   HandleError Err.Number, Erl, ModuleName, "NotOnAContinuedLine", Outcome
   SetModeFormOnTop
  
   Select Case Outcome
      Case RS:    Resume
      Case RN:    Resume Next
      Case Else:  CloseEventTracer
   End Select
 
End Function


Private Sub GetFirstLastLinesOf(ProcName As String, _
            CodePane As Object, CodeMod As Object, _
            FirstLn As Long, LastLn As Long, _
            Optional VeryFirstLn As Long)
            
      'TrT "GetFirstLastLinesOf", True, , , 1
      'TrV "ProcName", ProcName
            
  'Gets the First and Last Line numbers of the current Procedure.
            
   On Error GoTo ErrorHandler
            
   Dim sTestLine   As String

  'Note:  The FirstLn & LastLn "returned" by these Functions are
  '       the lines just "Before" and just "After" the procedure.
  '       Compensate by (1) to search only inside the procedure.
            
   FirstLn = CodeMod.ProcStartLine(ProcName, vbext_pk_Proc)
   
   VeryFirstLn = FirstLn       'Before any adjustments are made.
   
  'On occassion the VBIDE will return a StartLine that is off by more than
  'one line. Loop thru code until we can guarantee a hit on the first line.
  'This can happen when there are many blank lines or a comment line has
  'been inserted between procedures.
  
   Do
      sTestLine = Trim(CodePane.CodeModule.Lines(FirstLn, 1))
      
      If Left(Trim(sTestLine), 1) = "'" Then sTestLine = ""  'Is a comment!
   
      If sTestLine = "" Then FirstLn = FirstLn + 1
   Loop Until sTestLine > ""
   
   LastLn = CodeMod.ProcCountLines(ProcName, vbext_pk_Proc) + FirstLn - 1
   
  'Extra lines at the end of the Code Module can cause the VBIDE
  'to mis-count the the number of lines in the last Procedure.
  'Double check for the last line by looking for the word "End"
  
   Do
      sTestLine = Left(Trim(CodePane.CodeModule.Lines(LastLn, 1)), 3)
      
      If sTestLine <> "End" Then LastLn = LastLn - 1
   Loop Until sTestLine = "End"
   
  'Now, "After" getting the Last line, check the FirstLn to see
  'if it needs to compensate for a Line Continuation character.
   
   AdjustForLineContinuation ProcName, CodeMod, FirstLn
   
 'TrV "FirstLn", FirstLn, True, , 3
 'TrV "LastLn", LastLn
 
 'TrT "GetFirstLastLinesOf [Exit]"
   
   Exit Sub
 
ErrorHandler:

   ReleaseFormOnTopMode
   HandleError Err.Number, Erl, ModuleName, "GetFirstLastLinesOf", Outcome
   SetModeFormOnTop
  
   Select Case Outcome
      Case RS:    Resume
      Case RN:    Resume Next
      Case Else:  CloseEventTracer
   End Select
 
End Sub


Private Sub Determine(ProcName As String, _
            CodePane As Object, CodeMod As Object, _
            StartLn As Long, StartCol As Long)
      'TrT "Determine", True
      'TrV "ProcName", ProcName, True
      'TrV "StartLn", StartLn, True
      'TrV "StartCol", StartCol
 
  'Determines whether or not the cursor/focus is inside of a Procedure.
  
   On Error GoTo ErrorHandler
  
   Dim EndLn  As Long
   Dim EndCol As Long
            
   ProcName = ""
            
   If Not VBI.ActiveCodePane Is Nothing Then
   
      Set CodePane = VBI.ActiveCodePane
          ProcName = GetProcedureName(CodePane, CodeMod, _
                                      StartLn, StartCol, EndLn, EndCol)
   End If
 
   Exit Sub
 
ErrorHandler:
 
   ReleaseFormOnTopMode
   HandleError Err.Number, Erl, ModuleName, "Determine", Outcome
   SetModeFormOnTop
  
   Select Case Outcome
      Case RS:    Resume
      Case RN:    Resume Next
      Case Else:  CloseEventTracer
   End Select
 
End Sub


Private Function GetProcedureName _
                (CodePane As Object, CodeMod As Object, _
                 StartLn As Long, StartCol As Long, _
                 EndLn As Long, EndCol As Long) As String
           'TrT "GetProcedureName", True
 
  'Finds the location of the cursor in the CodeMod and returns the name
  'of the Function or Procedure it is in.   Also sets it's line position.
  
   On Error GoTo ErrorHandler
                 
   Dim ProcName  As String
   Dim sCodeLine As String
   Dim LineCount As Integer
                 
   GetProcedureName = ""

   If CodePane Is Nothing Then Exit Function

   Set CodeMod = CodePane.CodeModule
   
  'The GetSelection method will return the cursor line number to "iLine"
  
   CodePane.GetSelection StartLn, StartCol, EndLn, EndCol
   
  'Note: StartLn can be 1-2 lines "outside" a procedure and still work.
  
   ProcName = CodeMod.ProcOfLine(StartLn, vbext_pk_Proc)
   
  'VBIDE bug. The above line will still return "ProcName" even if it
  'is in a Property statement. Therefore have to check the LineCount of
  'the ProcName. If it is zero, then the ProcName is a Property Statement.
  
   On Error Resume Next
  
   LineCount = CodeMod.ProcCountLines(ProcName, vbext_pk_Proc)
  
   If LineCount = 0 Then ProcName = ""
   
   GetProcedureName = ProcName
   
 'TrV "", GetProcedureName

   Exit Function
 
ErrorHandler:
 
   ReleaseFormOnTopMode
   HandleError Err.Number, Erl, ModuleName, "GetProcedureName", Outcome
   SetModeFormOnTop
  
   Select Case Outcome
      Case RS:    Resume
      Case RN:    Resume Next
      Case Else:  CloseEventTracer
   End Select
 
End Function


Private Function IsFirstLineOfProcedure(CodeLine As String) As Boolean
           'TrT "IsFirstLineOfProcedure", True, True

  'Tests a line of code to see if it is the First Line of a Procedure.

   On Error GoTo ErrorHandler
 
   Dim bIsInProcedure   As Boolean
   
   If Left(Trim(CodeLine), 1) <> "'" Then
      If InStr(1, CodeLine, "Sub ") <> 0 Or _
         InStr(1, CodeLine, "Function ") <> 0 Then
        'InStr(1, CodeLine, "Property ")  'Not worth the effort at this time.
        
         If InStr(1, CodeLine, "Declare ") = 0 Then
         
            If InStr(1, CodeLine, "(") <> 0 Then bIsInProcedure = True
            
            If InStr(1, CodeLine, "Private") = 0 Then
            If InStr(1, CodeLine, "Public") = 0 Then
            If InStr(1, CodeLine, "Static") = 0 Then
            If InStr(1, CodeLine, "Friend") = 0 Then
            If InStr(1, "SubFun", Left(CodeLine, 3)) = 0 Then
               bIsInProcedure = False   'It is NOT a Procedure or Function.
            End If: End If:  End If: End If: End If
            
         End If
      End If
   End If
   
  'The following test is for the "TrcT" procecures in the modEventTracer.bas.
  'These procedures should not be messed with, so set bIsInProcedure = False.
   
   If InStr(1, CodeLine, "Sub TrcT(") <> 0 Or _
      InStr(1, CodeLine, "Sub TrcV(") <> 0 Or _
      InStr(1, CodeLine, "DebugForOutput(") <> 0 Then bIsInProcedure = False

   IsFirstLineOfProcedure = bIsInProcedure
   
   Exit Function
 
ErrorHandler:
 
   ReleaseFormOnTopMode
   HandleError Err.Number, Erl, ModuleName, "IsFirstLineOfProcedure", Outcome
   SetModeFormOnTop
  
   Select Case Outcome
      Case RS:    Resume
      Case RN:    Resume Next
      Case Else:  CloseEventTracer
   End Select
 
End Function


Private Function InsertionPointFor(DeclarationName As String, _
                 CodePane As Object, CodeMod As Object) As Long
           'TrT "InsertionPointFor", True
           'TrV "DeclarationName", True
 
  'Determines an aesthetic and logical place to insert Declaration variables.
            
   On Error GoTo ErrorHandler
   
   Dim lCnt                As Long
   Dim lConstLine          As Long
   Dim lOptionLine         As Long
   Dim lInsertPoint        As Long
   Dim lFirstOtherLine     As Long
   Dim lLastHeaderComment  As Long
   
   Dim sCodeLine           As String

   For lCnt = 1 To CodeMod.CountOfDeclarationLines
   
      sCodeLine = Trim(CodePane.CodeModule.Lines(lCnt, 1))
       
      If sCodeLine > "" Then
         
         If Left(sCodeLine, 1) <> "'" Then
            If sCodeLine = "Option Explicit" Then lOptionLine = lCnt
            If Left(sCodeLine, 6) = "Const " Then lConstLine = lCnt
            
            If lFirstOtherLine = 0 Then
               If InStr(1, sCodeLine, "Const ") = 0 Then
                  If InStr(1, sCodeLine, "Option ") = 0 Then
                     lFirstOtherLine = lCnt
                  End If
               End If
            End If
            
         Else
            If lFirstOtherLine = 0 Then
               lLastHeaderComment = lCnt      'Count the comment lines.
            End If
         End If
         
      End If
   Next lCnt
   
   Select Case DeclarationName
     
      Case Is = "Const"
      
         If lConstLine > 0 Then
            lInsertPoint = lConstLine
         Else
            lInsertPoint = lOptionLine
         End If
 
      Case Is = "Option Explicit"
      
         lInsertPoint = lLastHeaderComment
   
   End Select
   
   If lInsertPoint = 0 Then
      If lLastHeaderComment > 0 Then
         lInsertPoint = lLastHeaderComment
      Else
         If lFirstOtherLine > 0 Then
            lInsertPoint = lFirstOtherLine
         Else
            lInsertPoint = CodeMod.CountOfDeclarationLines
         End If
      End If
   End If
     
   InsertionPointFor = lInsertPoint
   
 'TrV "", InsertionPointFor
 
   Exit Function
 
ErrorHandler:
 
   ReleaseFormOnTopMode
   HandleError Err.Number, Erl, ModuleName, "InsertionPointFor", Outcome
   SetModeFormOnTop
  
   Select Case Outcome
      Case RS:    Resume
      Case RN:    Resume Next
      Case Else:  CloseEventTracer
   End Select
 
End Function


Private Sub AdjustForErrorHandling( _
            ProcName As String, CodePane As Object, CodeMod As Object, _
            FirstLn As Long, LastLn As Long)
      'TrT "AdjustForErrorHandling", True
      'TrV "ProcName", ProcName
            
  'Adjusts First and Last Line numbers to accommodate Error handler at End.
            
   On Error GoTo ErrorHandler
   
   Dim StartLn      As Long
   Dim EndLn        As Long
            
   Dim sErrLine     As String
   Dim sErrLabel    As String
   Dim sExitLine    As String
   
   Dim bExists      As Boolean
   
   EndLn = LastLn                  'Don't work with the passed values!
   StartLn = FirstLn
   
   bExists = _
   CodeMod.Find("On Error GoTo ", StartLn, 1, EndLn, -1, False, True)
   
  'If "On Error GoTo" exists, then adjust the LastLn of the procedure
  'to be the "Exit" command above the Error handlers Label.

   If bExists = True Then
   
      FirstLn = StartLn         'Adjust StartLn to be "On Error GoTo"
      
      sErrLine = (CodePane.CodeModule.Lines(StartLn, 1))
      
      sErrLabel = _
      Trim(Right(sErrLine, (InStr(1, StrReverse(sErrLine), " ")) - 1))
      
      If sErrLabel <> "0" Then
      
         sErrLabel = sErrLabel & ":"               'Look for a label.
         
         bExists = _
         CodeMod.Find(sErrLabel, StartLn, EndLn, -1, False, True)
         
        'The "adjusted for" LastLn gets passed back to the procedure.
      
         If bExists Then
         
            LastLn = StartLn - 1     'Start adjustment at the "label:"
            
            For EndLn = LastLn To FirstLn Step -1
            
               sExitLine = Trim(CodePane.CodeModule.Lines(EndLn, 1))
               
               If sExitLine > "" Then
                  If Left(sExitLine, 1) <> "'" Or _
                     Left(sExitLine, 4) <> "Rem " Then
                     
                    'Use InStr to test. There may be leading line numbers.

                     If InStr(1, sExitLine, "End") > 0 Or _
                        InStr(1, sExitLine, "Exit Sub") > 0 Or _
                        InStr(1, sExitLine, "Exit Function") > 0 Or _
                        InStr(1, sExitLine, "Exit Property") > 0 Then
                           
                        LastLn = EndLn
                     End If
                     
                     Exit For        'Stop! Don't go any further.
                        
                  End If
               End If
               
            Next EndLn
            
         End If
      End If
   End If
   
   Exit Sub
   
ErrorHandler:
 
   ReleaseFormOnTopMode
   HandleError Err.Number, Erl, ModuleName, "AdjustForErrorHandling", Outcome
   SetModeFormOnTop
  
   Select Case Outcome
      Case RS:    Resume
      Case RN:    Resume Next
      Case Else:  CloseEventTracer
   End Select
 
End Sub


Private Sub RemoveBlankLines _
           (BeforeOrAfter As String, LinePosition As Long, _
            CodePane As Object, CodeMod As Object, Indent As Integer)
      'TrT "RemoveBlankLines", True
      'TrV "", BeforeOrAfter, True
      'TrV "LinePosition", LinePosition
            
  'Removes all blank lines either "Before or "After" the current line.
 
   On Error GoTo ErrorHandler
                                        
   Dim sTestLine   As String
     
   If UCase(BeforeOrAfter) = "BEFORE" Then    'Pass "LastLn" as 2nd Param.
   
     'Find the Indent of the next line "before" removing the blank spaces!

      Indent = FindIndent(LinePosition - 1, CodePane, CodeMod)
      
      Do
         sTestLine = CodePane.CodeModule.Lines(LinePosition - 1, 1)
         
         If Trim(sTestLine) = "" Then       'Trim fixes "     " lines.
            CodeMod.DeleteLines LinePosition - 1, 1
            LinePosition = LinePosition - 1 'The "Last" line has moved up (1).
         End If
      Loop Until Trim(sTestLine) > ""
   Else                                     'Pass "FirstLn" as 2nd Param.
   
     'Find the Indent of the next line "before" removing the blank spaces!

      Indent = FindIndent(LinePosition + 1, CodePane, CodeMod)

      Do
         sTestLine = CodePane.CodeModule.Lines(LinePosition + 1, 1)
         
         If Trim(sTestLine) = "" Then
            If LinePosition + 1 < CodeMod.CountOfLines Then
               CodeMod.DeleteLines LinePosition + 1, 1
              'Don't have to adjust "Last" when deleting lines from the top.
            Else
               Exit Do             'Bug fix for modules with NO procedures.
            End If
         End If
      Loop Until Trim(sTestLine) > ""
   End If
 
   Exit Sub
 
ErrorHandler:
 
   ReleaseFormOnTopMode
   HandleError Err.Number, Erl, ModuleName, _
                          "RemoveBlankLines (Before-After)", Outcome
   SetModeFormOnTop
  
   Select Case Outcome
      Case RS:    Resume
      Case RN:    Resume Next
      Case Else:  CloseEventTracer
   End Select
 
End Sub


Private Sub RemoveLine _
           (CodePane As Object, CodeMod As Object, _
            ByVal Begin As Long, ByVal LastLn)
      'TrT "RemoveLine", True
      'TrV "Begin", Begin, True
      'TrV "LastLn", LastLn
            
  'Deletes a line of code and any extra blank lines (if needed.)
 
   On Error GoTo ErrorHandler
                                  
   Dim sBeforeLn   As String
   Dim sAfterLn    As String
              
   sBeforeLn = Trim(CodePane.CodeModule.Lines(Begin - 1, 1))
   sAfterLn = Trim(CodePane.CodeModule.Lines(Begin + 1, 1))
                    
   If sBeforeLn = "" And sAfterLn = "" Then
      LastLn = LastLn - 2
      CodePane.CodeModule.DeleteLines Begin, 2
   Else
      LastLn = LastLn - 1
      CodePane.CodeModule.DeleteLines Begin
   End If

   miRemoveCount = miRemoveCount + 1
 
   Exit Sub
 
ErrorHandler:
 
   ReleaseFormOnTopMode
   HandleError Err.Number, Erl, ModuleName, "RemoveLine", Outcome
   SetModeFormOnTop
  
   Select Case Outcome
      Case RS:    Resume
      Case RN:    Resume Next
      Case Else:  CloseEventTracer
   End Select
 
End Sub


Private Function FindIndent _
                (ByVal LinePosition As Long, _
                 CodePane As Object, CodeMod As Object) As Integer
           'TrT "FindIndent", True
                 
  'Finds the Indent of the First line of code close to the cursor position.

   On Error GoTo ErrorHandler
                                        
   Dim sCodeLine    As String
   Dim sIndentWrk   As String
   
   Dim wrkPosition  As Long
   
   Dim intCnt       As Integer
   Dim iDigits      As Integer
   
   wrkPosition = LinePosition
   
   Do
      sCodeLine = CodePane.CodeModule.Lines(wrkPosition, 1)
      wrkPosition = wrkPosition + 1
      intCnt = intCnt + 1
      
      If intCnt > 20 Then Exit Do     'Bug fix for NO procedures in module.
      
      If Left(Trim(sCodeLine), 1) = "'" Then sCodeLine = "" 'Skip comments.
      
   Loop Until Trim(sCodeLine) > ""
   
   If Left(sCodeLine, 4) = "End " Then        ' Hit the [End] of procedure.
                                              ' Go the other direction.
      wrkPosition = LinePosition
      
      Do
         sCodeLine = CodePane.CodeModule.Lines(wrkPosition, 1)
         wrkPosition = wrkPosition - 1
         If Left(Trim(sCodeLine), 1) = "'" Then sCodeLine = ""  'A comment.
      Loop Until Trim(sCodeLine) > ""
      
   End If
   
   sIndentWrk = sCodeLine   'See if we need to compensate for line numbers.
   
   If InStr(1, "123456789", Left(sIndentWrk, 1)) <> 0 Then
      iDigits = InStr(1, sIndentWrk, " ")
      sIndentWrk = Right(sCodeLine, Len(sCodeLine) - iDigits)
   End If

   FindIndent = Len(sCodeLine) - Len(LTrim(sIndentWrk))
   
  'If the user has clicked on a blank line, column 1, there will be no indent.
  
   If FindIndent = 0 Then FindIndent = 3   '(To match my own indenting style).
   
 'TrV "", FindIndent
   
   Exit Function
 
ErrorHandler:
 
   ReleaseFormOnTopMode
   HandleError Err.Number, Erl, ModuleName, "FindIndent", Outcome
   SetModeFormOnTop
  
   Select Case Outcome
      Case RS:    Resume
      Case RN:    Resume Next
      Case Else:  CloseEventTracer
   End Select
 
End Function


Private Sub ShowSetFocusMessage(Title As String, Optional Text As String)
      'TrT "ShowSetFocusMessage", True
      'TrV "Title", Title, True
      'TrV "Text", Text

  'Displays an error message when the user has not provided the proper Focus.

   ReleaseFormOnTopMode
   
   If Text > "" Then
      MsgBox Text, vbOKOnly, Title
   Else
      
      MsgBox "   Click the Mouse in a  Function or Procedure," & vbCrLf & _
             "   in order to set the FOCUS for this Operation.", _
             vbOKOnly, Title
   End If
   
   ResetInfoWindow
      
   SetModeFormOnTop
   
End Sub


Private Sub CloseAllCodeWindows()
      'TrT "CloseAllCodeWindows"

  'Close All Code Windows/Panes  "and"  Immediate and Watch windows.
 
   On Error GoTo ErrorHandler

   Dim VBW      As VBIDE.Window

   For Each VBW In VBI.Windows
      If VBW.Type = vbext_wt_CodeWindow Or _
         VBW.Type = vbext_wt_Immediate Or _
         VBW.Type = vbext_wt_Watch Then
                    VBW.Close
      End If
   Next VBW

   miAlreadyOpenPanes = 0
   
   cmdMisc(1).SetFocus
   
   Exit Sub
 
ErrorHandler:
 
   ReleaseFormOnTopMode
   HandleError Err.Number, Erl, ModuleName, "CloseAllCodeWindows", Outcome
   SetModeFormOnTop
  
   Select Case Outcome
      Case RS:    Resume
      Case RN:    Resume Next
      Case Else:  CloseEventTracer
   End Select
 
End Sub


Private Function EventTracerIsBeingRunOnItsSelf() As Boolean
           'TrT "EventTracerIsBeingRunOnItsSelf", True
             
  '"EventTracer" cannot run some commands on itself.
  
   If VBI.ActiveVBProject.Name = "EventTracer" Then
   
      EventTracerIsBeingRunOnItsSelf = True
      
      MsgBox "AnCaOtNa  UnRa  EYaEntVaAcerTra  OnYa  ItsYa  ElfSa." _
              & vbCrLf & vbCrLf & _
             "ItYa  IsYa  aYa  EryVa  EryVa  adBa  ingTha  oTa  oDa!!", _
              vbExclamation, "Cannot Run EventTracer On Its Self."
   Else
      EventTracerIsBeingRunOnItsSelf = False
   End If
   
  'TrV "", EventTracerIsBeingRunOnItsSelf
   
End Function


Private Sub ShowHelp()
      'TrT "ShowHelp"
 
   On Error GoTo ErrorHandler

   mbHelpHasFocus = True

   lblStatus.Visible = False
 
   SSTab.Visible = False
   cmdMisc(0).Visible = False
   cmdMisc(1).Visible = False
   cmdMisc(2).Caption = "Okay"
   cmdMisc(3).Visible = False

   Me.Caption = "EventTracer Help"
   Me.Left = Me.Left - 1365
   Me.Width = Me.Width + 2730
   Me.Top = Me.Top - 500
   Me.Height = Me.Height + 500
   cmdMisc(2).Left = cmdMisc(2).Left + 1365
   cmdMisc(2).Top = cmdMisc(2).Top + 500
   Me.Refresh
   
   picBox.Top = 150
   picBox.Left = 150
   picBox.Width = Me.Width - 370
   picBox.Height = cmdMisc(2).Top - 300
   
   picBox.FontSize = 9
   picBox.BackColor = vbWhite
   picBox.ForeColor = vbBlack
   picBox.Font = "Courier New"
   picBox.BorderStyle = vbFixedSingle
    
   picBox.Cls
   
   picBox.Print
   picBox.Print "  The functions of most buttons are self-evident."
   picBox.Print
   picBox.Print "  Click on a button and observe it's effects. You"
   picBox.Print "  can do this, then quit the Project without sav-"
   picBox.Print "  ing any of the changes made by " & vbQ;
   picBox.Print "EventTracer." & vbQ
   
                '  The "INSERT Trace for VARIABLE" buttons require
                '  that you click on a variable "before" you click
                '  the button. (This sets the Focus.)
   picBox.Print
   picBox.Print "  The " & vbQ;
   picBox.ForeColor = vbBlue
   picBox.Print "INSERT Trace for VARIABLE";
   picBox.ForeColor = vbBlack
   picBox.Print vbQ & " buttons require"
   picBox.Print "  that you click on a variable " & vbQ;
   picBox.ForeColor = vbBlue
   picBox.Print "before";
   picBox.ForeColor = vbBlack
   picBox.Print vbQ & " you click"
   picBox.Print "  the button. (This sets the Focus.)"
   
               '  The "INSERT Trace to File (OPEN-CLOSE)" buttons
               '  redirect the output to an ASCII Text file named
               '  "Trace?.Log" (where "?" is a sequential number).
               '  These are usually in the "FORM INITIALIZE" and
               '  "UNLOAD/END" events, but can be used almost any-
               '  where in your code. Delete all the "Trace*.Log"
               '  files to restart the numbering at (1).
   
   picBox.Print
   picBox.Print "  The " & vbQ;
   picBox.ForeColor = vbBlue
   picBox.Print "INSERT Trace to File (OPEN-CLOSE)";
   picBox.ForeColor = vbBlack
   picBox.Print vbQ & " buttons"
   picBox.Print "  redirect the output to an ASCII Text file named"
   picBox.Print "  " & vbQ;
   picBox.ForeColor = vbBlue
   picBox.Print "Trace";
   picBox.ForeColor = vbRed
   picBox.FontBold = True
   picBox.Print "?";
   picBox.FontBold = False
   picBox.ForeColor = vbBlue
   picBox.Print ".Log";
   picBox.ForeColor = vbBlack
   picBox.Print vbQ & " (where " & vbQ;
   picBox.ForeColor = vbRed
   picBox.FontBold = True
   picBox.Print "?";
   picBox.FontBold = False
   picBox.ForeColor = vbBlack
   picBox.Print vbQ & " is a sequential number)."
   picBox.Print "  These are usually in the " & vbQ;
   picBox.Print "FORM INITIALIZE" & vbQ & " and"
   picBox.Print "  UNLOAD/END" & vbQ;
   picBox.Print " events, but can be used almost any-"
   picBox.Print "  where in your code. Delete all the " & vbQ;
   picBox.ForeColor = vbBlue
   picBox.Print "Trace";
   picBox.ForeColor = vbRed
   picBox.Print "*";
   picBox.ForeColor = vbBlue
   picBox.Print ".Log";
   picBox.ForeColor = vbBlack
   picBox.Print vbQ
   picBox.Print "  files to restart the numbering at (1)."
      
                '  When you add Event Tracing or Error Handling,
                '  "EventTracer" adds a module(s) to your project.
                '  Refer to "modEventTracer" or "modErrorHandler"
                '  for programming Help and suggestions.
   picBox.Print
   picBox.Print "  When you add Event Tracing, or Error Handling,"
   picBox.Print "  " & vbQ & "EventTracer" & vbQ; " adds a module(s)";
   picBox.Print " to your project."
   picBox.Print "  Refer to " & vbQ;
   picBox.ForeColor = vbBlue
   picBox.Print "modEventTracer";
   picBox.ForeColor = vbBlack
   picBox.Print vbQ & " or " & vbQ;
   picBox.ForeColor = vbBlue
   picBox.Print "modErrorHandler";
   picBox.ForeColor = vbBlack
   picBox.Print vbQ
   picBox.Print "  for programming Help and suggestions."
      
   picBox.Visible = True

   Exit Sub
 
ErrorHandler:
 
   ReleaseFormOnTopMode
   HandleError Err.Number, Erl, ModuleName, "ShowHelp", Outcome
   SetModeFormOnTop
  
   Select Case Outcome
      Case RS:    Resume
      Case RN:    Resume Next
      Case Else:  CloseEventTracer
   End Select
 
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
'             General Procedures for handling Interface display.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'

Private Sub SetModeFormOnTop()
      'TrT "SetModeFormOnTop"
        
  'Lock in frmAddin to be the Topmost window
   SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
   gbFormModeChange = True
End Sub


Private Sub ReleaseFormOnTopMode()
      'TrT "ReleaseFormOnTopMode"
        
  'Release frmAddin as Topmost window (usually so a message box can be on top).
   Screen.MousePointer = vbDefault
   SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
   gbFormModeChange = True
   Me.Refresh
   Me.Show       'Keeps message boxes from making Form disappear.
End Sub


Private Sub ResetInfoWindow()
      'TrT "ResetInfoWindow"
        
   Me.Caption = "EventTracer"
   lblStatus.Caption = ""
   lblStatus.Visible = False
   picBox.Visible = False
   Me.Refresh
   
End Sub


Private Sub SetErrorScopeOption()
      'TrT "SetErrorScopeOption", True
      'TrV "ScopeOfOption", miScopeOfError
        
  'Sets the option value and Font Bold for the Error Handling option box.
   
   Dim iCnt As Integer
   
   For iCnt = 0 To 2
      optScopeError(iCnt).Value = False
      optScopeError(iCnt).FontBold = False
   Next
   
   optScopeError(miScopeOfError).Value = True
   optScopeError(miScopeOfError).FontBold = True

End Sub


Private Sub SetTraceScopeOption()
      'TrT "SetTraceScopeOption", True
      'TrV "ScopeOfOption", miScopeOfTrace

  'Sets the option value and Font Bold for the EventTracer option box.
   
   Dim iCnt As Integer
   
   For iCnt = 0 To 2
      OptScopeTrace(iCnt).Value = False
      OptScopeTrace(iCnt).FontBold = False
   Next
   
   OptScopeTrace(miScopeOfTrace).Value = True
   OptScopeTrace(miScopeOfTrace).FontBold = True
   
   EnableTraceButtonsForScope
   
End Sub


Private Sub EnableTraceButtonsForScope()
      'TrT "EnableTraceButtonsForScope", True
      'TrV "ScopeOfTrace", miScopeOfTrace

   Dim iCnt          As Integer

   For iCnt = 2 To 8
      cmdTrace(iCnt).Enabled = True
   Next
   
   If miScopeOfTrace <> ProcedureLevel Then
      cmdTrace(2).Enabled = False            'Trace for Variable "Above"
      cmdTrace(3).Enabled = False            'Trace for Variable "Below"
      cmdTrace(4).Enabled = False            'Trace OPEN to file
      cmdTrace(5).Enabled = False            'Trace CLOSE to file
   End If
   
   If miScopeOfTrace = ProcedureLevel Then
      cmdTrace(6).Enabled = False            'Turn Traces OFF
      cmdTrace(7).Enabled = False            'Turn Traces ON
      cmdTrace(8).Enabled = False            'Delete Trc commands
   End If

   If mbHasTraceModule = False Then
      cmdTrace(6).Enabled = False            'Turn Traces OFF
      cmdTrace(7).Enabled = False            'Turn Traces ON
      cmdTrace(8).Enabled = False            'Delete Trace commands
   End If
   
End Sub


Private Sub cmdTrace_GotFocus(Index As Integer)
      'TrT "cmdTrace_GotFocus", True
      'TrV "Index", Index
        
   CursorToFocus 25, 80
End Sub


Private Sub cmdErrors_GotFocus(Index As Integer)
      'TrT "cmdErrors_GotFocus", True
      'TrV "Index", Index
        
   CursorToFocus 25, 80
End Sub


Private Sub cmdMisc_LostFocus(Index As Integer)
      'TrT "cmdMisc_LostFocus", True
      'TrV "Index", Index
        
   If Not gbOpenedAllPanes Then
      If Index = 1 Then cmdMisc(1).Caption = "Cancel"   'Change from "Close"
   End If
End Sub


Private Sub cmdMisc_GotFocus(Index As Integer)
      'TrT "cmdMisc_GotFocus", True
      'TrV "Index", Index

   Dim Over As Long
   
   Select Case Index  'Anal adjustments.
      Case 0
         Over = 80    'Anal.
      Case 1
         cmdMisc(1).Caption = "Close"
         Over = 21    'Very anal.
      Case 2
         Over = 18    'Extremely anal.
      Case 3
         Over = 19    'Need professional Help.
   End Select
      
   CursorToFocus 5, Over
End Sub


Private Sub CursorToFocus(Down As Long, Over As Long)

  'Moves the Mouse cursor to the current control with the Focus.

   Dim x As Long, y As Long, xy As Long

   If Me.BorderStyle = 0 Then
     x = Me.ActiveControl.Left / Screen.TwipsPerPixelX + _
       ((Me.ActiveControl.Width / 2) / Screen.TwipsPerPixelX) + _
        (Me.Left / Screen.TwipsPerPixelX)
     y = Me.ActiveControl.Top / Screen.TwipsPerPixelY + _
       ((Me.ActiveControl.Height / 2) / Screen.TwipsPerPixelY) + _
        (Me.Top / Screen.TwipsPerPixelY)
   Else
     x = Me.ActiveControl.Left / Screen.TwipsPerPixelX + _
       ((Me.ActiveControl.Width / 2 + 60) / Screen.TwipsPerPixelX) + _
        (Me.Left / Screen.TwipsPerPixelX)
     y = Me.ActiveControl.Top / Screen.TwipsPerPixelY + _
       ((Me.ActiveControl.Height / 2 + 360) / Screen.TwipsPerPixelY) + _
        (Me.Top / Screen.TwipsPerPixelY)
   End If

    xy = SetCursorPos((x + Over), (y + Down))

End Sub


Private Sub optScopeTrace_Click(Index As Integer)
      'TrT "optScopeTrace_Click", True
      'TrV "Index", Index
        
   Dim iCnt               As Integer
   Dim bWasProcedure      As Boolean
   
   Dim CodePane           As VBIDE.CodePane
   Dim CodeMod            As CodeModule
   
   Dim ProcName           As String

   Dim StartLn            As Long
   Dim StartCol           As Long
   Dim EndLn              As Long

   miScopeOfTrace = Index        'Scope of Operations for Event Tracing.
   
   If OptScopeTrace(2).FontBold = True Then bWasProcedure = True
   
   For iCnt = 0 To 2
      OptScopeTrace(iCnt).FontBold = False
   Next iCnt
   
   OptScopeTrace(Index).FontBold = True
   
   EnableTraceButtonsForScope
   
  'When the User switches the scope to the Module Level, it is usually to
  'turn the Trace ON or OFF. Go ahead and snap the cursor to these buttons.
   
   If bWasProcedure And cmdTrace(6).Enabled = True Then
   
         If VBI.ActiveCodePane Is Nothing Then Exit Sub
   
         Set CodePane = VBI.ActiveCodePane
         Set CodeMod = CodePane.CodeModule
      
         Determine ProcName, CodePane, CodeMod, StartLn, StartCol
         
         If ProcName = "" Then Exit Sub
      
         GetFirstLastLinesOf _
         ProcName, CodePane, CodeMod, StartLn, EndLn
      
      If CodeMod.Find("'Trc", StartLn, 1, EndLn, -1, False, True) Then
         cmdTrace(7).SetFocus    'Was Turn OFF. Now Turn ON
      Else
         cmdTrace(6).SetFocus    'Turn OFF
      End If
   End If
   
End Sub


Private Sub optScopeError_Click(Index As Integer)
      'TrT "optScopeError_Click", True
      'TrV "Index", Index

   Dim iCnt As Integer
   
   miScopeOfError = Index       'Scope of Operations for Error Handling.
   
   For iCnt = 0 To 2
      optScopeError(iCnt).FontBold = False
   Next iCnt
   
   optScopeError(Index).FontBold = True
   
End Sub


Private Sub SSTab_Click(PreviousTab As Integer)
      'TrT "SSTab_Click", True
      'TrV "PreviousTab", PreviousTab
        
  'Save which ever Tab was last clicked on, so "EventTracer"
  'can default back to the same Tab the next time it is run.
  
   Dim iCnt As Integer
  
   SaveSetting App.EXEName, Me.Name, "LastTab", SSTab.Tab
   
   If mbSettingTabInit Then Exit Sub
   
  'Adjust the height of the Form for the Tab being displayed.
   
   If SSTab.Tab = 1 Then
      SSTab.Height = SSTab.Height - 1750
      For iCnt = 0 To 3
         cmdMisc(iCnt).Top = cmdMisc(iCnt).Top - 1750
      Next iCnt
      Me.Height = Me.Height - 1750
      cmdMisc(2).Enabled = False      'Don't want to mess with scaling Help.
      
      miScopeOfError = ProcedureLevel 'Default to Procedure when switching

      SetErrorScopeOption
   Else
      SSTab.Height = SSTab.Height + 1750
      For iCnt = 0 To 3
         cmdMisc(iCnt).Top = cmdMisc(iCnt).Top + 1750
      Next iCnt
      Me.Height = Me.Height + 1750
      cmdMisc(2).Enabled = True
      
      miScopeOfTrace = ProcedureLevel 'Default to Procedure when switching
      
      SetTraceScopeOption
   End If
   
   Me.Refresh
   
End Sub


Private Sub FormWinRegPos(MyForm As Form, Optional Save As Boolean)
      'TrT "FormWinRegPos"
 
   On Error GoTo ErrorHandler
    
  ' Purpose: Gets or Saves a Forms last Position.
  '      By: Brad Skidmore - Off of "Planet-Source-Code".
   
   If MyForm Is Nothing Then Exit Sub

   With MyForm
    
      If Save Then
      
        'If Form was Minimized or Maximized then Save the current Windowstate
        'and set Back to Normal Or previous non Max or Min State.
        'Then Save Positioning Parameters SaveSetting App.EXEName ..."
        
         SaveSetting App.EXEName, .Name, "WindowState", .WindowState
         
         If .WindowState = vbMinimized Or .WindowState = vbMaximized Then
            .WindowState = vbNormal
         End If
   
         SaveSetting App.EXEName, .Name, "Top", .Top
         SaveSetting App.EXEName, .Name, "Left", .Left

      Else
        'If Not Saving, then must Be Getting.
        'Need to ref the:   AppName, FrmName, KeyName
         
         .Top = GetSetting(App.EXEName, .Name, "Top", .Top)
         .Left = GetSetting(App.EXEName, .Name, "Left", .Left)
         
        'Be sure to set the WindowState last
        '(Can't Change POSN if vbMinimized Or Maximized
       
         .WindowState = _
          GetSetting(App.EXEName, .Name, "WindowState", .WindowState)
         
      End If
   End With
   
   Exit Sub
 
ErrorHandler:
 
   ReleaseFormOnTopMode
   HandleError Err.Number, Erl, ModuleName, "FormWinRegPos", Outcome
   SetModeFormOnTop
  
   Select Case Outcome
      Case RS:    Resume
      Case RN:    Resume Next
      Case Else:  CloseEventTracer
   End Select
 
End Sub


Private Sub SetDefaultFocus()
      'TrT "SetDefaultFocus"
  
  'Sets the Focus to the "most probable" button when the Form is activated.
  
   On Error GoTo ErrorHandler

   Dim CodePane           As VBIDE.CodePane
   Dim CodeMod            As CodeModule
   
   Dim ProcName           As String
   Dim sCodeLine          As String
   Dim sTestLine          As String
   Dim sVarName           As String
   
   Dim iCnt               As Integer
   Dim iTraceMustFocus    As Integer
   Dim iErrorMustFocus    As Integer
   Dim iTraceOption       As Integer
   Dim iErrorOption       As Integer
   
   Dim bShowingTracePane  As Boolean
   Dim bTraceIsOff        As Boolean
   Dim bOkayForTrace      As Boolean
   Dim bOkayForTraceExit  As Boolean
   Dim bOkayForOpen       As Boolean
   Dim bOkayForClose      As Boolean
   Dim bOkayForVariable   As Boolean
   Dim bDeclarationsFocus As Boolean

   Dim FirstLn            As Long
   Dim LastLn             As Long
   Dim StartLn            As Long
   Dim StartCol           As Long
   Dim EndLn              As Long
   Dim EndCol             As Long
   Dim CurrentLn          As Long
    
   Dim VBW                As VBIDE.Window
   Dim VBP                As VBProject
   Dim VBC                As VBComponent
   Dim VBCTraceMod        As VBComponent
   Dim VBCErrorMod        As VBComponent
   
   miSetFocusTo = 0
   mbDoAutoFocus = True
   mbSetTheFocus = False
   mbHasTraceModule = False
   
   miScopeOfTrace = ProcedureLevel
   miScopeOfError = ProcedureLevel
   
  'Restore which ever Tab was "last clicked on" by the user.
  'A change in SSTab can trigger the "Form_Paint" event.
  
   gbFormModeChange = True     'This temporarily disables "Form_Paint"
   mbSettingTabInit = True
  
   SSTab.Tab = GetSetting(App.EXEName, frmAddIn.Name, "LastTab", SSTab.Tab)
   
   mbSettingTabInit = False

   If VBI.ActiveVBProject Is Nothing Then Exit Sub
   
   Set VBP = VBI.ActiveVBProject
   
   Set VBCTraceMod = VBP.VBComponents("modEventTracer")
   Set VBCErrorMod = VBP.VBComponents("modErrorHandler")
   
   Set CodePane = VBI.ActiveCodePane
      
   If Not CodePane Is Nothing Then
    
      If Not VBCTraceMod Is Nothing Then
         If VBI.ActiveCodePane.CodeModule = VBCTraceMod.Name Then
            bShowingTracePane = True  'May need to Opened for ON-OFF test.
         End If
      End If
      
      Set CodeMod = CodePane.CodeModule

      ProcName = GetProcedureName(CodePane, CodeMod, _
                                  StartLn, StartCol, EndLn, EndCol)
                                                              
      If ProcName > "" Then
         GetFirstLastLinesOf ProcName, CodePane, CodeMod, FirstLn, LastLn
         
         sCodeLine = CodePane.CodeModule.Lines(StartLn, 1)
         sVarName = ExtractVarName(sCodeLine, StartCol)
      Else
         sCodeLine = CodePane.CodeModule.Lines(StartLn, 1)
         If StartLn <= CodeMod.CountOfDeclarationLines Then
            bDeclarationsFocus = True
         End If
      End If
   End If
   
  'See what possibilities are available for setting the default Focus.
   
  'Resolve the default Focus for Event Tracing.
   
   If VBCTraceMod Is Nothing Then
      bOkayForTrace = True
      iTraceMustFocus = 1              'Add TrcT to Procedure
      miScopeOfTrace = ModuleLevel
   Else
   
      mbHasTraceModule = True
   
      Set CodeMod = VBCTraceMod.CodeModule
      Set CodePane = CodeMod.CodePane
      
      For CurrentLn = CodePane.CodeModule.CountOfLines To 1 Step -1
         sCodeLine = CodePane.CodeModule.Lines(CurrentLn, 1)
         If IsFirstLineOfProcedure(sCodeLine) Then
            ProcName = CodeMod.ProcOfLine(CurrentLn, vbext_pk_Proc)
            
            If ProcName = "OnOffTestForTrace" Then
            
               GetFirstLastLinesOf _
               ProcName, CodePane, CodeMod, StartLn, EndLn
               
               If CodeMod.Find("'TrcT ", StartLn, 1, EndLn, -1, _
                                         False, True) Then
                  bTraceIsOff = True    'Trc is turned OFF.
               End If
               
               Exit For
            End If
         End If
      Next CurrentLn
      
      If Not bShowingTracePane Then
         For Each VBW In VBI.Windows     'Close the Window we opened.
            If InStr(VBW.Caption, "modEventTracer") Then VBW.Close
         Next VBW
      End If
      
   End If
      
      
   Set CodePane = VBI.ActiveCodePane
   
   If Not CodePane Is Nothing Then
   
      Set CodeMod = CodePane.CodeModule
   
      Determine ProcName, CodePane, CodeMod, StartLn, StartCol

      If ProcName > "" Then

         sCodeLine = CodePane.CodeModule.Lines(StartLn, 1)
         sVarName = ExtractVarName(sCodeLine, StartCol)
         
         If IsFirstLineOfProcedure(sCodeLine) Then
            sVarName = ""
            If Not AlreadyHas("TrcT " & vbQ & ProcName, _
                               CodePane, CodeMod, FirstLn, LastLn) Then
               iTraceMustFocus = 1      'Trc for Procedure
            End If
         End If
         
         If sVarName > "" Then
            If Left(Trim(sVarName), 3) = "Trc" Then sVarName = ""
            If Left(Trim(sVarName), 4) = "'Trc" Then sVarName = ""
         End If
         
         If Trim(sVarName) > "" Then
            If Not AlreadyHas("TrcV " & vbQ & sVarName & vbQ, _
                              CodePane, CodeMod, FirstLn, LastLn) Then
               bOkayForVariable = True
            End If
         End If
         
         GetFirstLastLinesOf ProcName, CodePane, CodeMod, StartLn, LastLn
         
        'If Trace was Turned ON in another module and the user has switched
        'to a module where the Trace is OFF, this will detect it.
         
         If bTraceIsOff = False Then
            If CodeMod.Find("'Trc", StartLn, 1, EndLn, -1, False, True) Then
               bTraceIsOff = True
            End If
         End If
         
         If InStr(1, ProcName, "Form_", vbTextCompare) Then
            If InStr(1, ProcName, "Unload") Then
               If Not AlreadyHas("CloseDebug", _
                                  CodePane, CodeMod, FirstLn, LastLn) Then
                  bOkayForClose = True
               End If
            Else
               If InStr(1, ProcName, "_Load") Or _
                  InStr(1, ProcName, "_Initialize") Then
                  If Not AlreadyHas("OpenDebug", _
                         CodePane, CodeMod, FirstLn, LastLn) Then
                     bOkayForOpen = True
                  End If
               End If
            End If
         End If
         
         If Not AlreadyHas("TrcT " & vbQ & ProcName, _
                            CodePane, CodeMod, FirstLn, LastLn) Then
            bOkayForTrace = True
         End If
         
         If Not AlreadyHas("[Exit]", _
                            CodePane, CodeMod, FirstLn, LastLn) Then
            bOkayForTraceExit = True
         End If
         
      End If
   End If

   If iTraceMustFocus = 0 Then
      If bTraceIsOff Then
         iTraceMustFocus = 7            'Turn Trace ON (priority)
      End If

      If bOkayForVariable Then
         iTraceOption = 3               'Trace for Variable "Below"
         
         If sVarName = "Exit" Then      'Trace for Variable "Above"
            iTraceOption = 9
         End If
         
      ElseIf bOkayForClose Then
         iTraceOption = 5               'Trace to File CLOSE
      ElseIf bOkayForTrace Then
         iTraceOption = 1               'Trace for Procedure
      ElseIf bOkayForOpen Then
         iTraceOption = 4               'Trace to File OPEN
      ElseIf bOkayForTraceExit Then
         iTraceOption = 2               'Trace for [Exit]
      ElseIf bDeclarationsFocus Then
         iTraceOption = 1               'Just Focus on the Top Button.
      Else
         iTraceOption = 6               'Turn Trace OFF (last option)
      End If
   End If
   
   If iTraceOption = 6 Or iTraceMustFocus = 7 Then
      miScopeOfTrace = ModuleLevel
      If bOkayForTrace Then iTraceMustFocus = 1  'Override at Module level.
   Else
      miScopeOfTrace = ProcedureLevel
   End If
                  
  'Resolve the default Focus for Error Handling
  
   If VBCErrorMod Is Nothing Then
      iErrorMustFocus = 1               'Set to INSERT ERROR HANDLING.
      miScopeOfError = ProcedureLevel
   Else
   
      Set CodePane = VBI.ActiveCodePane
   
      If Not CodePane Is Nothing Then
   
         Set CodeMod = CodePane.CodeModule
   
         If Not AlreadyHas("Const ModuleName ", _
                            CodePane, CodeMod, 1, LastLn) Then
                            
            iErrorOption = 1            'Set to INSERT ERROR HANDLING.
         Else
            If ProcName > "" Then
            
               If AlreadyHas("On Error", _
                              CodePane, CodeMod, FirstLn, LastLn) Then
         
                  For CurrentLn = FirstLn To LastLn Step 1
                  
                     sCodeLine = CodePane.CodeModule.Lines(CurrentLn, 1)
                     
                     sTestLine = Trim(sCodeLine)
                     
                     If sTestLine > "" Then
                     
                       'Test for existing line numbers
                       
                        If InStr(1, "1", Left(sTestLine, 1)) <> 0 Then
                           iErrorOption = 3   'Set to REMOVE LINE NUMBERS
                           Exit For
                        End If
                     
                     End If
               
                  Next CurrentLn
      
                  If iErrorOption <> 3 Then iErrorOption = 2 'ADD NUMBERS
            
               Else
                  iErrorOption = 1          'Set to INSERT ERROR HANDLING
               End If
            Else
               iErrorOption = 2                  'Set to ADD LINE NUMBERS
            End If
         End If
      Else
         iErrorOption = 1                  'Last resort to INSERT HANDLER
      End If
   End If
   
   
   If SSTab.Tab = 0 Then          'Event Tracing TAB has Focus.
      
     'Set the Focus Option
   
      If iTraceMustFocus <> 0 Then
         miSetFocusTo = iTraceMustFocus
      Else
         miSetFocusTo = iTraceOption
         If bDeclarationsFocus Then miScopeOfTrace = ModuleLevel
      End If
      
     'The following is a bug fix when adding EventTracing to a "new" Project.
     'If Scope is at Procedure level and the programmer clicks on Module
     'level without first clicking in the module, you get a runtimer err 5.
     'This only happens with compiled code, not in the VBIDE. This is just
     'a down and dirty fix.
      
      If mbHasTraceModule = False Then miScopeOfTrace = ModuleLevel
      
   Else                           'Error Handling TAB has Focus
   
     'Resize the Form for the smaller Error Handler screen.
   
      SSTab.Height = SSTab.Height - 1750
      For iCnt = 0 To 3
         cmdMisc(iCnt).Top = cmdMisc(iCnt).Top - 1750
      Next iCnt
      Me.Height = Me.Height - 1750
      cmdMisc(2).Enabled = False      'Don't want to mess with scaling Help.
      
     'Set the Focus Option
     
      miScopeOfError = ProcedureLevel 'Always default to Procedure Level.

      If iErrorMustFocus <> 0 Then
         miSetFocusTo = iErrorMustFocus
      Else
         miSetFocusTo = iErrorOption
      End If
      
   End If
   
   Me.Refresh
   
   SetErrorScopeOption
   SetTraceScopeOption
   
   gbFormModeChange = False       'Allow "Form_Paint" to set the Focus.
      
   Exit Sub

ErrorHandler:
 
   ReleaseFormOnTopMode
   HandleError Err.Number, Erl, ModuleName, "SetDefaultFocus", Outcome
   SetModeFormOnTop
  
   Select Case Outcome
      Case RS:    Resume
      Case RN:    Resume Next
      Case Else:  CloseEventTracer
   End Select
 
End Sub
