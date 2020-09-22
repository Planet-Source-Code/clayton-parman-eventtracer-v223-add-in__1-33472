VERSION 5.00
Begin VB.Form dlgRemove 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Delete Trc Commands"
   ClientHeight    =   4230
   ClientLeft      =   2760
   ClientTop       =   3615
   ClientWidth     =   2835
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "dlgRemove.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   2835
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CancelButton 
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
      Height          =   320
      Left            =   487
      TabIndex        =   7
      Top             =   3615
      Width           =   1860
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Okay"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   487
      TabIndex        =   6
      Top             =   3300
      Width           =   1860
   End
   Begin VB.Frame fraRemoveOptions 
      Caption         =   " Select Items to Delete "
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
      Height          =   2640
      Left            =   75
      TabIndex        =   0
      Top             =   300
      Width           =   2680
      Begin VB.CheckBox chkRemove 
         Caption         =   "modEventTracer  (module)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   75
         TabIndex        =   5
         Top             =   2100
         Width           =   2490
      End
      Begin VB.CheckBox chkRemove 
         Caption         =   "All   ""NoTrc""  commands"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   75
         TabIndex        =   4
         Top             =   1200
         Width           =   2490
      End
      Begin VB.CheckBox chkRemove 
         Caption         =   "All   ""Trc [Exit ]   procedure""  "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   3
         Top             =   750
         Width           =   2520
      End
      Begin VB.CheckBox chkRemove 
         Caption         =   "All   ""Trc""   commands"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   75
         TabIndex        =   2
         Top             =   300
         Width           =   2490
      End
      Begin VB.CheckBox chkRemove 
         Caption         =   "All   ""TrcV""   commands"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   75
         TabIndex        =   1
         Top             =   1650
         Width           =   2490
      End
   End
End
Attribute VB_Name = "dlgRemove"
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
'    Desc:  Options for removing various Trace commands from Project.
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'

Option Explicit

Const ModuleName    As String = "dlgRemove"


Private Sub Form_Activate()

   chkRemove(0).Value = vbChecked
   chkRemove(1).Value = vbChecked
   chkRemove(2).Value = vbUnchecked
   chkRemove(3).Value = vbUnchecked
   chkRemove(4).Value = vbUnchecked
   CancelButton.SetFocus
   
End Sub


Private Sub OKButton_Click()
   Me.Hide
End Sub


Private Sub CancelButton_Click()

   Dim iCnt As Integer
   
   For iCnt = 0 To 4
      chkRemove(iCnt).Value = vbUnchecked
   Next
   
   Me.Hide
End Sub
   
   
Private Sub OKButton_GotFocus()
   CursorToFocus -3, 20
End Sub

   
Private Sub CancelButton_GotFocus()
   CursorToFocus -3, 30
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
