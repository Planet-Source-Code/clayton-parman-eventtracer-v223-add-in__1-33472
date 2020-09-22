VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Event Tracer"
   ClientHeight    =   2580
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   3015
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1780.762
   ScaleMode       =   0  'User
   ScaleWidth      =   2831.241
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "Okay"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   750
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1875
      Width           =   1515
   End
   Begin VB.Image imgAbout 
      Height          =   1020
      Left            =   975
      Picture         =   "frmAbout.frx":030A
      Stretch         =   -1  'True
      Top             =   675
      Width           =   1065
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Application Title"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   345
      Left            =   240
      TabIndex        =   1
      Top             =   300
      Width           =   2535
   End
End
Attribute VB_Name = "frmAbout"
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
'    Date:  April 03, 2001
'
'    Desc:  Simple "About" Form.
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'

Option Explicit

Const ModuleName As String = "frmAbout"


Private Sub Form_Activate()
   cmdOK.SetFocus
End Sub


Private Sub Form_Load()
   Me.Caption = "About " & App.Title
                       
   lblTitle.Caption = "EventTracer v" & App.Major & "." & App.Minor
End Sub


Private Sub cmdOK_Click()
   Unload Me
End Sub


Private Sub cmdOK_GotFocus()
   CursorToFocus 50, 32
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
