VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form Main 
   BorderStyle     =   0  'None
   Caption         =   ".: GTorrent :."
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab3 
      Height          =   9495
      Left            =   120
      TabIndex        =   52
      Top             =   1920
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   16748
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Introduction"
      TabPicture(0)   =   "Main.frx":27C92
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label12"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label13(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label13(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Line1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Line2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label14"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Line3"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label15"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label16"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label17"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Picture1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Picture2"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "About The Program"
      TabPicture(1)   =   "Main.frx":27CAE
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Line4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Line5"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Line6"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Line7"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label18"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label19"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Line8(0)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Line8(2)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Line9"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Line10"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label20"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Line11"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Line12"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Line13"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Line14"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Label22"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Label23"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Label24"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Label25"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Line15"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Label26"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Label27"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Label28"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Label29"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Line16"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "Label30"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "Label31"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "Line17"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "Label32"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "Label33"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "Line18"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).ControlCount=   31
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   6135
         Left            =   -67320
         Picture         =   "Main.frx":27CCA
         ScaleHeight     =   6105
         ScaleWidth      =   7305
         TabIndex        =   59
         Top             =   2640
         Width           =   7335
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   6135
         Left            =   -74880
         Picture         =   "Main.frx":2CEA2
         ScaleHeight     =   6105
         ScaleWidth      =   7425
         TabIndex        =   58
         Top             =   2640
         Width           =   7455
      End
      Begin VB.Line Line18 
         X1              =   7320
         X2              =   0
         Y1              =   5280
         Y2              =   5280
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Caption         =   "Owner a Company For WebHosting && Web Developing "
         Height          =   255
         Left            =   3120
         TabIndex        =   77
         Top             =   4920
         Width           =   4095
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         Caption         =   "Job"
         Height          =   255
         Left            =   120
         TabIndex        =   76
         Top             =   4920
         Width           =   2775
      End
      Begin VB.Line Line17 
         X1              =   7320
         X2              =   0
         Y1              =   4800
         Y2              =   4800
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         Caption         =   "18 Years Old"
         Height          =   255
         Left            =   3120
         TabIndex        =   75
         Top             =   4440
         Width           =   4095
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         Caption         =   "Age"
         Height          =   375
         Left            =   120
         TabIndex        =   74
         Top             =   4440
         Width           =   2775
      End
      Begin VB.Line Line16 
         X1              =   7320
         X2              =   7320
         Y1              =   5280
         Y2              =   2400
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         Caption         =   "Tunisie"
         Height          =   255
         Left            =   3120
         TabIndex        =   73
         Top             =   3960
         Width           =   4095
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         Caption         =   "Egypt"
         Height          =   255
         Left            =   3120
         TabIndex        =   72
         Top             =   3480
         Width           =   4095
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         Caption         =   "United Arab Emirates"
         Height          =   255
         Left            =   3120
         TabIndex        =   71
         Top             =   3000
         Width           =   4095
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         Caption         =   "Ashraf Magdi Kaabi"
         Height          =   255
         Left            =   3120
         TabIndex        =   70
         Top             =   2520
         Width           =   4095
      End
      Begin VB.Line Line15 
         X1              =   7320
         X2              =   0
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         Caption         =   "From"
         Height          =   255
         Left            =   120
         TabIndex        =   69
         Top             =   3960
         Width           =   2775
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         Caption         =   "Live in"
         Height          =   255
         Left            =   120
         TabIndex        =   68
         Top             =   3480
         Width           =   2775
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         Caption         =   "Born in "
         Height          =   255
         Left            =   120
         TabIndex        =   67
         Top             =   3000
         Width           =   2775
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Caption         =   "Full Name"
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   2520
         Width           =   2775
      End
      Begin VB.Line Line14 
         X1              =   7320
         X2              =   0
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Line Line13 
         X1              =   7320
         X2              =   0
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line12 
         X1              =   7320
         X2              =   0
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line11 
         X1              =   3000
         X2              =   3000
         Y1              =   2400
         Y2              =   5280
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Caption         =   "About The Programmer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   2040
         Width           =   14775
      End
      Begin VB.Line Line10 
         X1              =   15000
         X2              =   15000
         Y1              =   2400
         Y2              =   1920
      End
      Begin VB.Line Line9 
         X1              =   0
         X2              =   15000
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line8 
         Index           =   2
         X1              =   0
         X2              =   15000
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line8 
         Index           =   0
         X1              =   0
         X2              =   15000
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label Label19 
         Caption         =   $"Main.frx":30E15
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   63
         Top             =   1080
         Width           =   14895
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Caption         =   "About The Program"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   600
         Width           =   14775
      End
      Begin VB.Line Line7 
         X1              =   15000
         X2              =   0
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line6 
         X1              =   15000
         X2              =   15000
         Y1              =   960
         Y2              =   480
      End
      Begin VB.Line Line5 
         X1              =   0
         X2              =   15000
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   6240
         Y1              =   300
         Y2              =   300
      End
      Begin VB.Label Label17 
         Caption         =   $"Main.frx":30EDA
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74760
         TabIndex        =   61
         Top             =   8880
         Width           =   14775
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "The BitTorrent Solution: customers help distribute content"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -67320
         TabIndex        =   60
         Top             =   2400
         Width           =   7455
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "Problem: more customers require more bandwidth"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   57
         Top             =   2400
         Width           =   7455
      End
      Begin VB.Line Line3 
         X1              =   -60000
         X2              =   -75000
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "Introduction About Bittorrent"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   56
         Top             =   480
         Width           =   14775
      End
      Begin VB.Line Line2 
         X1              =   -60000
         X2              =   -60000
         Y1              =   840
         Y2              =   360
      End
      Begin VB.Line Line1 
         X1              =   -75000
         X2              =   -60000
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label13 
         Caption         =   $"Main.frx":30FA6
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   -74880
         TabIndex        =   55
         Top             =   1920
         Width           =   10935
      End
      Begin VB.Label Label13 
         Caption         =   "There is a solution. BitTorrent is a simple software product which addresses all of these problems."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   -74880
         TabIndex        =   54
         Top             =   1680
         Width           =   10935
      End
      Begin VB.Label Label12 
         Caption         =   $"Main.frx":3107F
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -74880
         TabIndex        =   53
         Top             =   960
         Width           =   14775
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   120
      Top             =   360
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   9495
      Left            =   120
      TabIndex        =   38
      Top             =   1920
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   16748
      _Version        =   393216
      Tabs            =   8
      TabsPerRow      =   8
      TabHeight       =   520
      TabCaption(0)   =   "Forum Index"
      TabPicture(0)   =   "Main.frx":311B7
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "WebBrowser11"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Private Message"
      TabPicture(1)   =   "Main.frx":311D3
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "WebBrowser12"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Profile"
      TabPicture(2)   =   "Main.frx":311EF
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "WebBrowser13"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "UserGroup"
      TabPicture(3)   =   "Main.frx":3120B
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "WebBrowser14"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Memberlist"
      TabPicture(4)   =   "Main.frx":31227
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "WebBrowser15"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Search"
      TabPicture(5)   =   "Main.frx":31243
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "WebBrowser16"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "FAQ"
      TabPicture(6)   =   "Main.frx":3125F
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "WebBrowser17"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "Register"
      TabPicture(7)   =   "Main.frx":3127B
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "WebBrowser18"
      Tab(7).ControlCount=   1
      Begin SHDocVwCtl.WebBrowser WebBrowser18 
         Height          =   8895
         Left            =   -74880
         TabIndex        =   39
         Top             =   480
         Width           =   14895
         ExtentX         =   26273
         ExtentY         =   15690
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser17 
         Height          =   8895
         Left            =   -74880
         TabIndex        =   40
         Top             =   480
         Width           =   14895
         ExtentX         =   26273
         ExtentY         =   15690
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser16 
         Height          =   8895
         Left            =   -74880
         TabIndex        =   41
         Top             =   480
         Width           =   14895
         ExtentX         =   26273
         ExtentY         =   15690
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser15 
         Height          =   8895
         Left            =   -74880
         TabIndex        =   42
         Top             =   480
         Width           =   14895
         ExtentX         =   26273
         ExtentY         =   15690
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser14 
         Height          =   8895
         Left            =   -74880
         TabIndex        =   43
         Top             =   480
         Width           =   14895
         ExtentX         =   26273
         ExtentY         =   15690
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser13 
         Height          =   8895
         Left            =   -74880
         TabIndex        =   44
         Top             =   480
         Width           =   14895
         ExtentX         =   26273
         ExtentY         =   15690
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser12 
         Height          =   8895
         Left            =   -74880
         TabIndex        =   45
         Top             =   480
         Width           =   14895
         ExtentX         =   26273
         ExtentY         =   15690
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser11 
         Height          =   8895
         Left            =   120
         TabIndex        =   46
         Top             =   480
         Width           =   14895
         ExtentX         =   26273
         ExtentY         =   15690
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Communiction"
      Height          =   735
      Left            =   2400
      TabIndex        =   37
      Tag             =   "1008"
      Top             =   1080
      Width           =   5085
      Begin VB.CommandButton Command3 
         Caption         =   "Chat"
         Height          =   345
         Left            =   1800
         TabIndex        =   49
         Tag             =   "1006"
         Top             =   240
         Width           =   1500
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser19 
         Height          =   135
         Left            =   1920
         TabIndex        =   51
         Top             =   360
         Width           =   135
         ExtentX         =   238
         ExtentY         =   238
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin VB.CommandButton Command5 
         Caption         =   "About / Help"
         Height          =   345
         Left            =   3360
         TabIndex        =   50
         Tag             =   "1007"
         Top             =   240
         Width           =   1605
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Forum"
         Height          =   345
         Left            =   120
         TabIndex        =   48
         Tag             =   "1005"
         Top             =   250
         Width           =   1620
      End
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser10 
      Height          =   1455
      Left            =   7560
      TabIndex        =   36
      Top             =   360
      Width           =   7695
      ExtentX         =   13573
      ExtentY         =   2566
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.ComboBox lstLanguages 
      Height          =   315
      Left            =   240
      TabIndex        =   34
      Text            =   "English"
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Frame Frame4 
      Caption         =   "Select Your Language"
      Height          =   735
      Left            =   120
      TabIndex        =   35
      Tag             =   "1003"
      Top             =   1080
      Width           =   2175
   End
   Begin VB.CommandButton Bgo 
      Caption         =   "Go"
      Default         =   -1  'True
      Height          =   345
      Left            =   4560
      TabIndex        =   30
      Tag             =   "1001"
      Top             =   550
      Width           =   900
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9465
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   15150
      _ExtentX        =   26723
      _ExtentY        =   16695
      _Version        =   393216
      Tabs            =   9
      TabsPerRow      =   9
      TabHeight       =   520
      TabCaption(0)   =   "SuprNova"
      TabPicture(0)   =   "Main.frx":31297
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "WebBrowser1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Status"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Torrent Search"
      TabPicture(1)   =   "Main.frx":312B3
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2(0)"
      Tab(1).Control(1)=   "WebBrowser2"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "BT Etree"
      TabPicture(2)   =   "Main.frx":312CF
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(1)=   "WebBrowser3"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Fresh Torrents"
      TabPicture(3)   =   "Main.frx":312EB
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "WebBrowser4"
      Tab(3).Control(1)=   "Frame2(1)"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Search Torrent"
      TabPicture(4)   =   "Main.frx":31307
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "WebBrowser5"
      Tab(4).Control(1)=   "Frame2(2)"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "isoHunt"
      TabPicture(5)   =   "Main.frx":31323
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "WebBrowser6"
      Tab(5).Control(1)=   "Frame2(3)"
      Tab(5).ControlCount=   2
      TabCaption(6)   =   "The Vault"
      TabPicture(6)   =   "Main.frx":3133F
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "WebBrowser7"
      Tab(6).Control(1)=   "Frame2(4)"
      Tab(6).ControlCount=   2
      TabCaption(7)   =   "SLO Torrent"
      TabPicture(7)   =   "Main.frx":3135B
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "WebBrowser8"
      Tab(7).Control(1)=   "Frame2(5)"
      Tab(7).ControlCount=   2
      TabCaption(8)   =   "Torrent Box"
      TabPicture(8)   =   "Main.frx":31377
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Frame2(6)"
      Tab(8).Control(1)=   "WebBrowser9"
      Tab(8).ControlCount=   2
      Begin VB.Frame Frame2 
         Caption         =   "Status"
         Height          =   615
         Index           =   6
         Left            =   -74880
         TabIndex        =   21
         Tag             =   "1002"
         Top             =   360
         Width           =   14775
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   14535
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Status"
         Height          =   615
         Index           =   5
         Left            =   -74880
         TabIndex        =   20
         Tag             =   "1002"
         Top             =   360
         Width           =   14775
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   14535
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Status"
         Height          =   615
         Index           =   4
         Left            =   -74880
         TabIndex        =   19
         Tag             =   "1002"
         Top             =   360
         Width           =   14775
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   14535
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Status"
         Height          =   615
         Index           =   3
         Left            =   -74880
         TabIndex        =   18
         Tag             =   "1002"
         Top             =   360
         Width           =   14775
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   14535
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Status"
         Height          =   615
         Index           =   2
         Left            =   -74880
         TabIndex        =   17
         Tag             =   "1002"
         Top             =   360
         Width           =   14775
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   14535
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Status"
         Height          =   615
         Index           =   1
         Left            =   -74880
         TabIndex        =   16
         Tag             =   "1002"
         Top             =   360
         Width           =   14775
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   14535
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Status"
         Height          =   615
         Left            =   -74880
         TabIndex        =   15
         Tag             =   "1002"
         Top             =   360
         Width           =   14775
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   14535
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Status"
         Height          =   615
         Index           =   0
         Left            =   -74880
         TabIndex        =   14
         Tag             =   "1002"
         Top             =   360
         Width           =   14775
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   14535
         End
      End
      Begin VB.Frame Status 
         Caption         =   "Status"
         Height          =   615
         Left            =   120
         TabIndex        =   12
         Tag             =   "1002"
         Top             =   360
         Width           =   14775
         Begin VB.Label lblStatus 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   14535
         End
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser9 
         Height          =   8295
         Left            =   -74880
         TabIndex        =   11
         Top             =   1080
         Width           =   14775
         ExtentX         =   26061
         ExtentY         =   14631
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser8 
         Height          =   8295
         Left            =   -74880
         TabIndex        =   10
         Top             =   1080
         Width           =   14775
         ExtentX         =   26061
         ExtentY         =   14631
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser7 
         Height          =   8295
         Left            =   -74880
         TabIndex        =   9
         Top             =   1080
         Width           =   14775
         ExtentX         =   26061
         ExtentY         =   14631
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser6 
         Height          =   8295
         Left            =   -74880
         TabIndex        =   8
         Top             =   1080
         Width           =   14775
         ExtentX         =   26061
         ExtentY         =   14631
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser5 
         Height          =   8295
         Left            =   -74880
         TabIndex        =   7
         Top             =   1080
         Width           =   14775
         ExtentX         =   26061
         ExtentY         =   14631
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser4 
         Height          =   8295
         Left            =   -74880
         TabIndex        =   6
         Top             =   1080
         Width           =   14775
         ExtentX         =   26061
         ExtentY         =   14631
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser3 
         Height          =   8295
         Left            =   -74880
         TabIndex        =   5
         Top             =   1080
         Width           =   14775
         ExtentX         =   26061
         ExtentY         =   14631
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser2 
         Height          =   8295
         Left            =   -74880
         TabIndex        =   4
         Top             =   1080
         Width           =   14775
         ExtentX         =   26061
         ExtentY         =   14631
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   8295
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   14775
         ExtentX         =   26061
         ExtentY         =   14631
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
   End
   Begin VB.TextBox SearchText 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   4215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Global Torrents Searcher"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Tag             =   "1000"
      Top             =   360
      Width           =   7380
      Begin VB.CommandButton Command1 
         Caption         =   "Last Search"
         Height          =   345
         Left            =   5400
         TabIndex        =   47
         Tag             =   "1004"
         Top             =   200
         Width           =   1875
      End
   End
   Begin VB.Label Label21 
      Caption         =   "Born in      : United Arab Emirates"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   65
      Top             =   4800
      Width           =   4695
   End
   Begin VB.Line Line8 
      Index           =   1
      X1              =   0
      X2              =   15000
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   ".: GTorrent :."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   33
      Top             =   30
      Width           =   14775
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Height          =   135
      Left            =   14880
      TabIndex        =   32
      Top             =   120
      Width           =   135
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   15000
      TabIndex        =   31
      ToolTipText     =   "End"
      Top             =   0
      Width           =   255
   End
   Begin VB.Image Image9 
      Height          =   300
      Left            =   14790
      Picture         =   "Main.frx":31393
      Top             =   0
      Width           =   510
   End
   Begin VB.Image Image8 
      Height          =   300
      Left            =   14640
      Picture         =   "Main.frx":317AD
      Top             =   0
      Width           =   180
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   0
      Picture         =   "Main.frx":31ABF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14700
   End
   Begin VB.Image Image2 
      Height          =   645
      Left            =   0
      Picture         =   "Main.frx":31B07
      Top             =   0
      Width           =   60
   End
   Begin VB.Image Image4 
      Height          =   120
      Left            =   0
      Picture         =   "Main.frx":31D4D
      Top             =   11400
      Width           =   345
   End
   Begin VB.Image Image3 
      Height          =   11490
      Left            =   0
      Picture         =   "Main.frx":31FCF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   60
   End
   Begin VB.Image Image7 
      Height          =   120
      Left            =   240
      Picture         =   "Main.frx":32065
      Stretch         =   -1  'True
      Top             =   11400
      Width           =   14775
   End
   Begin VB.Image Image6 
      Height          =   11490
      Left            =   15300
      Picture         =   "Main.frx":32167
      Stretch         =   -1  'True
      Top             =   0
      Width           =   60
   End
   Begin VB.Image Image5 
      Height          =   120
      Left            =   15000
      Picture         =   "Main.frx":321CD
      Top             =   11400
      Width           =   330
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Starting in the Program in 27/09/2004
'Add Status Label in 29/09/2004 06:00 PM
'Add PopupBlocker in 30/09/2004 12:52 AM
'Add KDE Skin in 30/09/2004 9:45PM
'Add Multi Language in 1/10/2004 12:36 AM
'Add Forum in 2/10/2004 12:02 AM

Private Sub SuprNovasearch()
Dim SuprNova As String
SuprNova = "http://search.suprnova.org//search.php?data%5Bsessionid%5D=be8dacd01b675ad07d4eff66df832206&data%5Baction%5D=searchTorrent&data%5BsearchTerm%5D=" + SearchText + "&data%5BsearchType%5D=and&data%5BsearchDomain%5D=torrents&data%5Btype_id%5D=0&data%5Bcategory_id%5D=0&data%5Borderby%5D=Date&data%5BorderDir%5D=desc"
WebBrowser1.Navigate SuprNova
End Sub
Private Sub TorrentSearchsearch()
Dim torrentsearch As String
torrentsearch = "http://www.torrentsearch.us/search.php?query=" + SearchText + "&submit=Search"
WebBrowser2.Navigate torrentsearch
End Sub

Private Sub Btetreesearch()
Dim torrentindex As String
Btetree = "http://bt.etree.org/?search=" + SearchText + "&cat=0"
WebBrowser3.Navigate Btetree
End Sub
Private Sub FreshTorrentssearch()
Dim FreshTorrents As String
FreshTorrents = "http://freshtorrents.net/?search=" + SearchText + "&withdead=1"
WebBrowser4.Navigate FreshTorrents
End Sub
Private Sub searchtorrentsearch()
Dim searchtorrent As String
searchtorrent = "http://www2.searchtorrent.com/s?q=" + SearchText + "&s=0&t=AND"
WebBrowser5.Navigate searchtorrent
End Sub
Private Sub isohuntsearch()
Dim isohunt As String
isohunt = "http://isohunt.com/torrents.php?ihq=" + SearchText + "&ext=&op=and"
WebBrowser6.Navigate isohunt
End Sub

Private Sub thevaultsearch()
Dim thevault As String
thevault = "http://thevault.gotdns.com/modules.php?name=Bittorrent&file=index&search=" + SearchText + "&cat=0&orderby=0&ordertype=0"
WebBrowser7.Navigate thevault
End Sub

Private Sub slotorrentsearch()
Dim slotorrent As String
slotorrent = "http://slotorrent.bounceme.net:4545/index.html?search=" + SearchText + "&filter="
WebBrowser8.Navigate slotorrent
End Sub
Private Sub torrentboxsearch()
Dim torrentbox As String
torrentbox = "http://www.torrentbox.com/torrents-search.php?search=" + SearchText + "&cat=0&submit=Search"
WebBrowser9.Navigate torrentbox
End Sub



Private Sub Bgo_Click()
WebBrowser10.Refresh
SSTab1.Visible = True
SSTab2.Visible = False
SSTab3.Visible = False
On Error GoTo ender

Dim a As String
Dim B As String
Dim c As String

Do Until B = "0"
a = InStr(1, SearchText2.Text, " ")
B = a
SearchText.Text = Mid(SearchText.Text, 1, a - 1) + "+" + Mid(SearchText.Text, a + 1, Len(SearchText.Text) - a)
Loop

ender:

Call SuprNovasearch
Call TorrentSearchsearch
Call Btetreesearch
Call FreshTorrentssearch
Call searchtorrentsearch
Call isohuntsearch
Call thevaultsearch
Call slotorrentsearch
Call torrentboxsearch
End Sub

Private Sub Command1_Click()
WebBrowser10.Navigate "http://igatehosting.com/ts"
SSTab2.Visible = False
SSTab3.Visible = False
SSTab1.Visible = True
End Sub

Private Sub Command2_Click()
WebBrowser10.Refresh
SSTab3.Visible = False
SSTab2.Visible = True
WebBrowser11.Navigate "http://igatehosting.com/board/index.php"
WebBrowser12.Navigate "http://igatehosting.com/board/privmsg.php?folder=inbox"
WebBrowser13.Navigate "http://igatehosting.com/board/profile.php?mode=editprofile"
WebBrowser14.Navigate "http://igatehosting.com/board/groupcp.php"
WebBrowser15.Navigate "http://igatehosting.com/board/memberlist.php"
WebBrowser16.Navigate "http://igatehosting.com/board/search.php"
WebBrowser17.Navigate "http://igatehosting.com/board/faq.php"
WebBrowser18.Navigate "http://igatehosting.com/board/profile.php?mode=register&sid=ea5ce6203167049746e706c9251d69be"
End Sub

Private Sub Command3_Click()
WebBrowser10.Refresh
WebBrowser19.Navigate "irc://mesra.dal.net/torrent"
End Sub

Private Sub Command4_Click()
SSTab1.Visible = False
SSTab2.Visible = False
SSTab3.Visible = True
End Sub

Private Sub Command5_Click()
WebBrowser10.Refresh
SSTab1.Visible = False
SSTab2.Visible = False
SSTab3.Visible = True
End Sub

Private Sub Form_Load()
WebBrowser1.Navigate "http://suprnova.org"
WebBrowser2.Navigate "http://TorrentSearch.US"
WebBrowser3.Navigate "http://BT.etree.org"
WebBrowser4.Navigate "http://freshtorrents.net"
WebBrowser5.Navigate "http://SearchTorrent.com"
WebBrowser6.Navigate "http://isohunt.com"
WebBrowser7.Navigate "http://thevault.gotdns.com"
WebBrowser8.Navigate "http://slotorrent.bounceme.net:4545/"
WebBrowser9.Navigate "http://www.torrentbox.com"
Timer1.Enabled = True
SSTab3.Visible = False
SSTab2.Visible = False
SSTab1.Visible = True
WebBrowser10.Navigate "http://www.igatehosting.com/ts"
WebBrowser11.Navigate "http://www.igatehosting.com/ts/Load/index.htm"
WebBrowser12.Navigate "http://www.igatehosting.com/ts/Load/index.htm"
WebBrowser13.Navigate "http://www.igatehosting.com/ts/Load/index.htm"
WebBrowser14.Navigate "http://www.igatehosting.com/ts/Load/index.htm"
WebBrowser15.Navigate "http://www.igatehosting.com/ts/Load/index.htm"
WebBrowser16.Navigate "http://www.igatehosting.com/ts/Load/index.htm"
WebBrowser17.Navigate "http://www.igatehosting.com/ts/Load/index.htm"
WebBrowser18.Navigate "http://www.igatehosting.com/ts/Load/index.htm"
   AllowPopup = flase
    ''Search for ALL the Language Files
    SetLanguageFile "Lang\*.lan"
    txtSelected = "english"
    ''Load the default Language
    LoadResStrings Main, txtSelected & ".lan"
    LoadResStrings Main, txtSelected & ".lan"
End Sub

Private Sub Label10_Click()
Me.WindowState = 1
End Sub

Private Sub Label9_Click()
End
End Sub
Private Sub lstLanguages_Click()
    ''Set the Selected Language
    txtSelected = lstLanguages.Text
    ''Load the Selected Language
    LoadResStrings Main, txtSelected & ".lan"
    ''Load also in the other Form
    LoadResStrings Main, txtSelected & ".lan"
End Sub

Private Sub Timer1_Timer()
WebBrowser1.Silent = True
WebBrowser2.Silent = True
WebBrowser3.Silent = True
WebBrowser4.Silent = True
WebBrowser5.Silent = True
WebBrowser6.Silent = True
WebBrowser7.Silent = True
WebBrowser8.Silent = True
WebBrowser9.Silent = True
Timer1.Enabled = False
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
'shows done in the status bar
lblStatus.Caption = "Done"
End Sub

Private Sub WebBrowser1_DownloadBegin()
'Starting download
lblStatus.Caption = "Starting Download"
End Sub

Private Sub WebBrowser1_DownloadComplete()
'Done downloading
lblStatus.Caption = "Download Done!"
End Sub

Private Sub WebBrowser1_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
'Loaded page
lblStatus.Caption = "Done Loading!"
End Sub


Private Sub WebBrowser1_OnToolBar(ByVal toolbar As Boolean)
    On Error Resume Next


    If toolbar = False And mnuOptionsAllow.Checked = False Then
        Unload Me
    End If
End Sub

Private Sub WebBrowser1_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
'Shows progress in status bar
lblStatus.Caption = "Reading " & Progress & "  of  " & ProgressMax
End Sub

Private Sub webBrowser1_StatusTextChange(ByVal Text As String)
'shows new text in status bar
lblStatus.Caption = Text
End Sub
Private Sub WebBrowser2_DocumentComplete(ByVal pDisp As Object, URL As Variant)
'shows done in the status bar
Label1.Caption = "Done"
End Sub

Private Sub WebBrowser2_DownloadBegin()
'Starting download
Label1.Caption = "Starting Download"
End Sub

Private Sub WebBrowser2_DownloadComplete()
'Done downloading
Label1.Caption = "Download Done!"
End Sub

Private Sub WebBrowser2_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
'Loaded page
Label1.Caption = "Done Loading!"
End Sub


Private Sub WebBrowser2_OnToolBar(ByVal toolbar As Boolean)
    On Error Resume Next


    If toolbar = False And mnuOptionsAllow.Checked = False Then
        Unload Me
    End If
End Sub

Private Sub WebBrowser2_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
'Shows progress in status bar
Label1.Caption = "Reading " & Progress & "  of  " & ProgressMax
End Sub

Private Sub webBrowser2_StatusTextChange(ByVal Text As String)
'shows new text in status bar
Label1.Caption = Text
End Sub
Private Sub WebBrowser3_DocumentComplete(ByVal pDisp As Object, URL As Variant)
'shows done in the status bar
Label2.Caption = "Done"
End Sub

Private Sub WebBrowser3_DownloadBegin()
'Starting download
Label2.Caption = "Starting Download"
End Sub

Private Sub WebBrowser3_DownloadComplete()
'Done downloading
Label2.Caption = "Download Done!"
End Sub

Private Sub WebBrowser3_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
'Loaded page
Label2.Caption = "Done Loading!"
End Sub


Private Sub WebBrowser3_OnToolBar(ByVal toolbar As Boolean)
    On Error Resume Next


    If toolbar = False And mnuOptionsAllow.Checked = False Then
        Unload Me
    End If
End Sub

Private Sub WebBrowser3_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
'Shows progress in status bar
Label2.Caption = "Reading " & Progress & "  of  " & ProgressMax
End Sub

Private Sub webBrowser3_StatusTextChange(ByVal Text As String)
'shows new text in status bar
Label2.Caption = Text
End Sub
Private Sub WebBrowser4_DocumentComplete(ByVal pDisp As Object, URL As Variant)
'shows done in the status bar
Label3.Caption = "Done"
End Sub

Private Sub WebBrowser4_DownloadBegin()
'Starting download
Label3.Caption = "Starting Download"
End Sub

Private Sub WebBrowser4_DownloadComplete()
'Done downloading
Label3.Caption = "Download Done!"
End Sub

Private Sub WebBrowser4_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
'Loaded page
Label3.Caption = "Done Loading!"
End Sub


Private Sub WebBrowser4_OnToolBar(ByVal toolbar As Boolean)
    On Error Resume Next


    If toolbar = False And mnuOptionsAllow.Checked = False Then
        Unload Me
    End If
End Sub

Private Sub WebBrowser4_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
'Shows progress in status bar
Label3.Caption = "Reading " & Progress & "  of  " & ProgressMax
End Sub

Private Sub webBrowser4_StatusTextChange(ByVal Text As String)
'shows new text in status bar
Label3.Caption = Text
End Sub
Private Sub WebBrowser5_DocumentComplete(ByVal pDisp As Object, URL As Variant)
'shows done in the status bar
Label4.Caption = "Done"
End Sub

Private Sub WebBrowser5_DownloadBegin()
'Starting download
Label4.Caption = "Starting Download"
End Sub

Private Sub WebBrowser5_DownloadComplete()
'Done downloading
Label4.Caption = "Download Done!"
End Sub

Private Sub WebBrowser5_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
'Loaded page
Label4.Caption = "Done Loading!"
End Sub


Private Sub WebBrowser5_OnToolBar(ByVal toolbar As Boolean)
    On Error Resume Next


    If toolbar = False And mnuOptionsAllow.Checked = False Then
        Unload Me
    End If
End Sub

Private Sub WebBrowser5_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
'Shows progress in status bar
Label4.Caption = "Reading " & Progress & "  of  " & ProgressMax
End Sub

Private Sub webBrowser5_StatusTextChange(ByVal Text As String)
'shows new text in status bar
Label4.Caption = Text
End Sub
Private Sub WebBrowser6_DocumentComplete(ByVal pDisp As Object, URL As Variant)
Label5.Caption = "Done"
End Sub

Private Sub WebBrowser6_DownloadBegin()
Label5.Caption = "Starting Download"
End Sub

Private Sub WebBrowser6_DownloadComplete()
Label5.Caption = "Download Done!"
End Sub

Private Sub WebBrowser6_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
Label5.Caption = "Done Loading!"
End Sub


Private Sub WebBrowser6_OnToolBar(ByVal toolbar As Boolean)
    On Error Resume Next


    If toolbar = False And mnuOptionsAllow.Checked = False Then
        Unload Me
    End If
End Sub

Private Sub WebBrowser6_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
Label5.Caption = "Reading " & Progress & "  of  " & ProgressMax
End Sub

Private Sub webBrowser6_StatusTextChange(ByVal Text As String)
Label5.Caption = Text
End Sub
Private Sub WebBrowser7_DocumentComplete(ByVal pDisp As Object, URL As Variant)
Label6.Caption = "Done"
End Sub

Private Sub WebBrowser7_DownloadBegin()
Label6.Caption = "Starting Download"
End Sub

Private Sub WebBrowser7_DownloadComplete()
Label6.Caption = "Download Done!"
End Sub

Private Sub WebBrowser7_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
Label6.Caption = "Done Loading!"
End Sub


Private Sub WebBrowser7_OnToolBar(ByVal toolbar As Boolean)
    On Error Resume Next


    If toolbar = False And mnuOptionsAllow.Checked = False Then
        Unload Me
    End If
End Sub

Private Sub WebBrowser7_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
Label6.Caption = "Reading " & Progress & "  of  " & ProgressMax
End Sub

Private Sub webBrowser7_StatusTextChange(ByVal Text As String)
Label6.Caption = Text
End Sub
Private Sub WebBrowser8_DocumentComplete(ByVal pDisp As Object, URL As Variant)
Label7.Caption = "Done"
End Sub

Private Sub WebBrowser8_DownloadBegin()
Label7.Caption = "Starting Download"
End Sub

Private Sub WebBrowser8_DownloadComplete()
Label7.Caption = "Download Done!"
End Sub

Private Sub WebBrowser8_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
Label7.Caption = "Done Loading!"
End Sub


Private Sub WebBrowser8_OnToolBar(ByVal toolbar As Boolean)
    On Error Resume Next


    If toolbar = False And mnuOptionsAllow.Checked = False Then
        Unload Me
    End If
End Sub

Private Sub WebBrowser8_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
Label7.Caption = "Reading " & Progress & "  of  " & ProgressMax
End Sub

Private Sub webBrowser8_StatusTextChange(ByVal Text As String)
Label7.Caption = Text
End Sub
Private Sub WebBrowser9_DocumentComplete(ByVal pDisp As Object, URL As Variant)
Label8.Caption = "Done"
End Sub

Private Sub WebBrowser9_DownloadBegin()
Label8.Caption = "Starting Download"
End Sub

Private Sub WebBrowser9_DownloadComplete()
Label8.Caption = "Download Done!"
End Sub

Private Sub WebBrowser9_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
Label8.Caption = "Done Loading!"
End Sub


Private Sub WebBrowser9_OnToolBar(ByVal toolbar As Boolean)
    On Error Resume Next


    If toolbar = False And mnuOptionsAllow.Checked = False Then
        Unload Me
    End If
End Sub

Private Sub WebBrowser9_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
Label8.Caption = "Reading " & Progress & "  of  " & ProgressMax
End Sub

Private Sub webBrowser9_StatusTextChange(ByVal Text As String)
Label8.Caption = Text
End Sub
Private Sub WebBrowser1_NewWindow2(ppDisp As Object, Cancel As Boolean)
Set ppDisp = Nothing
Cancel = True
End Sub
Private Sub WebBrowser2_NewWindow2(ppDisp As Object, Cancel As Boolean)
Set ppDisp = Nothing
Cancel = True
End Sub
Private Sub WebBrowser3_NewWindow2(ppDisp As Object, Cancel As Boolean)
Set ppDisp = Nothing
Cancel = True
End Sub
Private Sub WebBrowser4_NewWindow2(ppDisp As Object, Cancel As Boolean)
Set ppDisp = Nothing
Cancel = True
End Sub
Private Sub WebBrowser5_NewWindow2(ppDisp As Object, Cancel As Boolean)
Set ppDisp = Nothing
Cancel = True
End Sub
Private Sub WebBrowser6_NewWindow2(ppDisp As Object, Cancel As Boolean)
Set ppDisp = Nothing
Cancel = True
End Sub
Private Sub WebBrowser7_NewWindow2(ppDisp As Object, Cancel As Boolean)
Set ppDisp = Nothing
Cancel = True
End Sub
Private Sub WebBrowser8_NewWindow2(ppDisp As Object, Cancel As Boolean)
Set ppDisp = Nothing
Cancel = True
End Sub
Private Sub WebBrowser9_NewWindow2(ppDisp As Object, Cancel As Boolean)
Set ppDisp = Nothing
Cancel = True
End Sub
