VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3960
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7065
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   3960
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   770
      Left            =   480
      Top             =   0
   End
   Begin MSComctlLib.ProgressBar P1 
      Height          =   135
      Left            =   5520
      TabIndex        =   0
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   238
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   550
      Scrolling       =   1
   End
   Begin VB.Label lblCopyright 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Global Torrent Searcher"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   6855
   End
   Begin VB.Label lblLicenseTo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Auther : Ashraf Magdi"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2004 GTorrent. All Right Reserved."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   2
      Top             =   3700
      Width           =   4815
   End
   Begin VB.Label lblLicenseTo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Powered by: iGate Solution"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   4080
      TabIndex        =   1
      Top             =   3360
      Width           =   2775
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a As Integer
Dim B As Integer
Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    Main.Show
    Main.Hide
    'lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    'lblProductName.Caption = App.Title
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub


Private Sub imgBG_Click()

End Sub

Private Sub Timer1_Timer()
a = a + 1
If a = 7 Then
Main.Show
Unload Me
End If
End Sub

Private Sub Timer2_Timer()
B = B + 1
P1.Value = B
End Sub


