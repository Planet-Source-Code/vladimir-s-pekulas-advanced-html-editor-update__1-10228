VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmDocument 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   5700
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5085
   ScaleWidth      =   5700
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab EV 
      Height          =   5100
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   5580
      _ExtentX        =   9843
      _ExtentY        =   8996
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Edit"
      TabPicture(0)   =   "frmDocument.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "rtfText2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Picture9(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "AutoSyntaxPic"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "BrowserTest"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Picture3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdTable"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Picture9(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdExitDoc"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdFind"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdFullView"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdReplace"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdOpenDoc"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdColorEdit"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdSepartate"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "rtfText"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lsMain"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "rtfText3"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "View"
      TabPicture(1)   =   "frmDocument.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "P2"
      Tab(1).Control(1)=   "pResizeWeb"
      Tab(1).Control(2)=   "cmdFor"
      Tab(1).Control(3)=   "cmdBack"
      Tab(1).Control(4)=   "cmdRef"
      Tab(1).ControlCount=   5
      Begin RichTextLib.RichTextBox rtfText3 
         Height          =   630
         Left            =   4680
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   2520
         Visible         =   0   'False
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1111
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         TextRTF         =   $"frmDocument.frx":0038
      End
      Begin Project1.ListSearch lsMain 
         Height          =   1935
         Left            =   480
         TabIndex        =   22
         Top             =   600
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   3413
      End
      Begin Project1.CodeHighlight rtfText 
         Height          =   2415
         Left            =   480
         TabIndex        =   24
         Top             =   480
         Width           =   4935
         _ExtentX        =   8281
         _ExtentY        =   4260
         Language        =   3
         KeywordColor    =   12582912
         OperatorColor   =   12582912
         DelimiterColor  =   32896
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.PictureBox cmdRef 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   -74903
         Picture         =   "frmDocument.frx":00F9
         ScaleHeight     =   195
         ScaleWidth      =   150
         TabIndex        =   20
         ToolTipText     =   "Refresh"
         Top             =   1455
         Width           =   150
      End
      Begin VB.PictureBox cmdBack 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   165
         Left            =   -74910
         Picture         =   "frmDocument.frx":0155
         ScaleHeight     =   165
         ScaleWidth      =   135
         TabIndex        =   19
         ToolTipText     =   "Back"
         Top             =   960
         Width           =   135
      End
      Begin VB.PictureBox cmdFor 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   165
         Left            =   -74910
         Picture         =   "frmDocument.frx":01A8
         ScaleHeight     =   165
         ScaleWidth      =   135
         TabIndex        =   18
         ToolTipText     =   "Forward"
         Top             =   1185
         Width           =   135
      End
      Begin VB.PictureBox cmdSepartate 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   165
         Left            =   120
         Picture         =   "frmDocument.frx":01FD
         ScaleHeight     =   165
         ScaleWidth      =   135
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   945
         Width           =   135
      End
      Begin VB.PictureBox pResizeWeb 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   -74925
         Picture         =   "frmDocument.frx":0254
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   15
         ToolTipText     =   "Size test"
         Top             =   645
         Width           =   195
      End
      Begin VB.PictureBox P2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   3075
         Left            =   -74535
         ScaleHeight     =   3075
         ScaleWidth      =   6015
         TabIndex        =   13
         Top             =   555
         Width           =   6015
         Begin VB.PictureBox Pruler 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   360
            Left            =   0
            Picture         =   "frmDocument.frx":02CD
            ScaleHeight     =   330
            ScaleWidth      =   18555
            TabIndex        =   23
            ToolTipText     =   "Different Screen Width"
            Top             =   0
            Visible         =   0   'False
            Width           =   18585
         End
         Begin SHDocVwCtl.WebBrowser Web 
            Height          =   1935
            Left            =   0
            TabIndex        =   14
            Top             =   240
            Visible         =   0   'False
            Width           =   3015
            ExtentX         =   5318
            ExtentY         =   3413
            ViewMode        =   0
            Offline         =   0
            Silent          =   0
            RegisterAsBrowser=   0
            RegisterAsDropTarget=   1
            AutoArrange     =   0   'False
            NoClientEdge    =   0   'False
            AlignLeft       =   0   'False
            ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
            Location        =   "res://C:\WINDOWS\SYSTEM\SHDOCLC.DLL/dnserror.htm#http:///"
         End
         Begin VB.Shape Shape 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00C0C0C0&
            FillStyle       =   5  'Downward Diagonal
            Height          =   2985
            Left            =   0
            Top             =   240
            Width           =   5940
         End
      End
      Begin VB.PictureBox cmdColorEdit 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   60
         Picture         =   "frmDocument.frx":1420F
         ScaleHeight     =   240
         ScaleWidth      =   270
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Color Editor"
         Top             =   2610
         Width           =   270
      End
      Begin VB.PictureBox cmdOpenDoc 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   60
         Picture         =   "frmDocument.frx":145D1
         ScaleHeight     =   195
         ScaleWidth      =   270
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "List of Open Documents"
         Top             =   2070
         Width           =   270
      End
      Begin VB.PictureBox cmdReplace 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   60
         Picture         =   "frmDocument.frx":148EB
         ScaleHeight     =   240
         ScaleWidth      =   270
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Replace in Text"
         Top             =   1530
         Width           =   270
      End
      Begin VB.PictureBox cmdFullView 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   75
         Picture         =   "frmDocument.frx":14CAD
         ScaleHeight     =   225
         ScaleWidth      =   255
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Full Screen View"
         Top             =   2355
         Width           =   255
      End
      Begin VB.PictureBox cmdFind 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   60
         Picture         =   "frmDocument.frx":14FFB
         ScaleHeight     =   195
         ScaleWidth      =   270
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Find in Text"
         Top             =   1245
         Width           =   270
      End
      Begin VB.PictureBox cmdExitDoc 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   135
         Left            =   105
         Picture         =   "frmDocument.frx":15315
         ScaleHeight     =   135
         ScaleWidth      =   180
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Close Current Document"
         Top             =   600
         Width           =   180
      End
      Begin VB.PictureBox Picture9 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   60
         Index           =   0
         Left            =   60
         Picture         =   "frmDocument.frx":1549B
         ScaleHeight     =   60
         ScaleWidth      =   270
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1890
         Width           =   270
      End
      Begin VB.PictureBox cmdTable 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   90
         Picture         =   "frmDocument.frx":155BD
         ScaleHeight     =   225
         ScaleWidth      =   210
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Table Wizard"
         Top             =   2970
         Width           =   210
      End
      Begin VB.PictureBox Picture3 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   60
         Left            =   60
         Picture         =   "frmDocument.frx":15893
         ScaleHeight     =   60
         ScaleWidth      =   270
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   840
         Width           =   270
      End
      Begin VB.PictureBox BrowserTest 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   75
         Picture         =   "frmDocument.frx":159B5
         ScaleHeight     =   255
         ScaleWidth      =   240
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Test With Favorite Browser"
         Top             =   3315
         Width           =   240
      End
      Begin VB.PictureBox AutoSyntaxPic 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   75
         Picture         =   "frmDocument.frx":15D27
         ScaleHeight     =   210
         ScaleWidth      =   240
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Auto Syntax OFF"
         Top             =   3930
         Width           =   240
      End
      Begin VB.PictureBox Picture9 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   60
         Index           =   1
         Left            =   60
         Picture         =   "frmDocument.frx":16009
         ScaleHeight     =   60
         ScaleWidth      =   270
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   3750
         Width           =   270
      End
      Begin RichTextLib.RichTextBox rtfText2 
         Height          =   1755
         Left            =   450
         TabIndex        =   17
         Top             =   2925
         Visible         =   0   'False
         Width           =   4920
         _ExtentX        =   8678
         _ExtentY        =   3096
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         ScrollBars      =   3
         RightMargin     =   3
         TextRTF         =   $"frmDocument.frx":1612B
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList IMG 
      Left            =   9360
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":161EC
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":162FE
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":16410
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":16522
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":16634
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":16746
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":16858
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":1696A
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":16A7C
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":16B8E
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":16CA0
            Key             =   "Align Left"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":16DB2
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":16EC4
            Key             =   "Align Right"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":16FD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":170EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":171FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":17312
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":17426
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":1753A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   9480
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   28
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":1764E
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":17760
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":17872
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":17984
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":17A96
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":17BA8
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":17CBA
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":17DCC
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":17EDE
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":17FF0
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":18102
            Key             =   "Align Left"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":18214
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":18326
            Key             =   "Align Right"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":18438
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":18B6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":18FC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":19414
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":19868
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":19CBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":1A110
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":1A564
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":1A9B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":1AE10
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":1B268
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":1B6BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":1BB10
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":1C0A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocument.frx":1C390
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************************
' Project:  Casper HTML   v.2.0                              *
' Filename: n/a                                              *
' Author:   Vladimir S. Pekulas Jr.                          *
' Date:     7/22/2000                                        *
' Copyright © 2000 Vladimir S. Pekulas Jr.                   *
'                                                            *
' Use this program as you wish, but please let me know       *
' if you like it. Anyway, you can do whatever you want       *
' with it. I'm not responsible for any demage tough :)       *
'*************************************************************

Private Declare Function GetCaretPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Const vbKeyLessThan = 60

Dim pt          As POINTAPI
Dim lngStart    As Long

Public WebSize As Integer '# What Size is the Web Window in ?
Public Separate As Integer '# Do we have two desktops ?

Private Sub eXIT_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub AutoSyntaxPic_Click()
On Error Resume Next
 fMainForm.mnuOptionsComplete.Checked = Not fMainForm.mnuOptionsComplete.Checked
 rtfText.SetFocus
 If fMainForm.mnuOptionsComplete.Checked = True Then AutoSyntaxPic.ToolTipText = "Syntaxing On"
 If fMainForm.mnuOptionsComplete.Checked = False Then AutoSyntaxPic.ToolTipText = "Syntaxing Off"
End Sub

Private Sub BrowserTest_Click() '# test edited file in default Browser
On Error Resume Next
 Dim strView As String
 Dim intFile As Integer
 intFile = FreeFile
 Open "c:\Casper~temp.html" For Output As #intFile
 Print #intFile, fMainForm.ActiveForm.rtfText.Text
 Close #intFile
 Shell ("start c:\Casper~temp.html")
End Sub

Private Sub cmdBack_Click()
On Error Resume Next
 Web.GoBack
End Sub

Private Sub cmdColorEdit_Click()
On Error Resume Next
 frmColor.Show 1, fMainForm
End Sub

Private Sub cmdFind_Click()
On Error Resume Next
 frmFind.Show 1, fMainForm
End Sub

Private Sub cmdFor_Click()
On Error Resume Next
 Web.GoForward
End Sub

Private Sub cmdFullView_Click()
On Error Resume Next
 fMainForm.FileManager.Checked = Not fMainForm.FileManager.Checked
 fMainForm.PTAB.Visible = fMainForm.FileManager.Checked

 fMainForm.mnuViewStatusBar.Checked = Not fMainForm.mnuViewStatusBar.Checked
 fMainForm.sbStatusBar.Visible = fMainForm.mnuViewStatusBar.Checked

 fMainForm.mnuViewToolbar.Checked = Not fMainForm.mnuViewToolbar.Checked
 fMainForm.tbToolBar.Visible = fMainForm.mnuViewToolbar.Checked
End Sub

Private Sub cmdOpenDoc_Click()
On Error Resume Next
MsgBox "Function not yet available." & vbCrLf & "Currently " & fMainForm.lDocumentCount & " document(s) opened.", vbInformation, "Unavailable"
End Sub

Private Sub cmdRef_Click()
On Error Resume Next
 Web.Refresh
End Sub

Private Sub cmdReplace_Click()
On Error Resume Next
 frmFind.Show 1, fMainForm
End Sub

Private Sub cmdSepartate_Click()
On Error Resume Next
 If Separate = 1 Then
  rtfText2.Text = rtfText.Text
  rtfText2.Visible = True
  rtfText.Height = (rtfText.Height / 2) - 50
  rtfText2.Left = rtfText.Left
  rtfText2.Width = rtfText.Width
  rtfText2.Top = (Me.ScaleHeight - 330) - rtfText.Height + 200
  rtfText2.Height = rtfText.Height
  Separate = 0
 Else
  rtfText2.Visible = False
  rtfText.Move 370, 300, Me.ScaleWidth - 470, Me.ScaleHeight - 330
  Separate = 1
 End If
End Sub

Private Sub cmdExitDoc_Click()
On Error Resume Next
  If MsgBox("Are you sure to close current document ?", vbQuestion + vbYesNo, "Close document ?") = vbYes Then
        fMainForm.lDocumentCount = fMainForm.lDocumentCount - 1
        Unload Me
   Else
        Exit Sub
   End If
End Sub

Private Sub cmdTable_Click()
On Error Resume Next
frmTables.Show 1, fMainForm
End Sub



Private Sub EV_Click(PreviousTab As Integer)
On Error Resume Next
'#################
'# Save for View #
'#################
Dim strView As String
Dim intFile As Integer
intFile = FreeFile
Open App.Path & "\Casper~temp.html" For Output As #intFile
Print #intFile, rtfText.Text
Close #intFile

If EV.Caption = "Edit" Then
 rtfText.Visible = True
 P2.Visible = False
 Web.Visible = False
 fMainForm.sbStatusBar.Panels(1).Text = "Status: Ready to Edit"
Else
 rtfText.Visible = False
 P2.Visible = True
 Web.Visible = True
 Web.Navigate (App.Path & "\Casper~temp.html")
 fMainForm.sbStatusBar.Panels(1).Text = "Status: Visual Viewing"
End If
End Sub



Private Sub Form_Load()
On Error Resume Next
 Separate = 1
 WebSize = 1 '# Set WebWindow Resizer to it's full size
 rtfText.Visible = True
 P2.Visible = False
 Web.Visible = False
 'fMainForm.CoDocs.AddItem "Untitled"
fMainForm.mnuOptionsUpper.Checked = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    EV.Top = 0
    EV.Left = 0
    EV.Height = Me.Height
    EV.Width = Me.Width
    rtfText.Move 370, 300, Me.ScaleWidth - 470, Me.ScaleHeight - 330
    Web.Move 370, 300, Me.ScaleWidth - 470, Me.ScaleHeight - 330
    P2.Move 370, 300, Me.ScaleWidth - 470, Me.ScaleHeight - 330
    Web.Top = 0
    Web.Left = 0
    PProp.Width = Me.Width - 350
    PProp.Left = 305
    PProp.Top = rtfText.Height + 75
    Shape.Height = P2.Height
    Shape.Width = P2.Width
    rtfText2.Width = rtfText.Width
End Sub



Private Sub pResizeWeb_Click() '# Determine and/or change size of WebWindow
On Error Resume Next
 If WebSize = 0 Then '# If WebSize=0 then resize to original shape
   'Call Form_Resize
   Web.Top = 0
   Web.Width = Me.ScaleWidth - 470
   Web.Height = Me.ScaleHeight - 330
   Pruler.Visible = False
   WebSize = 1
 Else  '# If WebSize = 1 then resize to smaller WebWindow
   Pruler.Visible = True
   Pruler.Top = 25
   Pruler.Left = 0
   Web.Top = 360
   Web.Width = ((Me.Width - Web.Left) / 100) * 75
   Web.Height = ((Me.Height - Web.Top) / 100) * 75
   WebSize = 0
 End If
End Sub


Private Sub rtfText_Change()
On Error Resume Next
 rtfText2.Text = rtfText.Text
 With rtfText
  'Count Lines
   'LineCount = SendMessageLong(.hwnd, EM_GETLINECOUNT, 0&, 0&)
   fMainForm.sbStatusBar.Panels(2).Text = "Lines: " & LineCount
 End With
End Sub

Private Sub rtfText_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
 If Button = 2 Then
  fMainForm.Fileopt.Enabled = True
  fMainForm.CloseDoc.Enabled = False 'Temporarly: For some reason my VB6 is crashing while closing document from pop-up menu ?@#$$!!$
  fMainForm.Paste1.Enabled = True

  If rtfText.SelLength > 0 Then
   fMainForm.Cut1.Enabled = True
  Else
   fMainForm.Cut1.Enabled = False
  End If
   
  If rtfText.SelLength > 0 Then
   fMainForm.Copy1.Enabled = True
  Else
   fMainForm.Copy1.Enabled = False
  End If

  fMainForm.InsertTag.Enabled = True
  fMainForm.Edittag.Enabled = True
  fMainForm.Date.Enabled = True
  
  fMainForm.PopupMenu fMainForm.RTFMenu
 End If
End Sub

Private Sub rtfText2_KeyPress(KeyAscii As Integer)
On Error Resume Next
 rtfText.Text = rtfText2.Text
End Sub

Private Sub lsMain_Done(ByVal Text As String)
On Error Resume Next
    ' Hide the popup window and add the text
    If fMainForm.mnuOptionsComplete.Checked = True Then
        ' Add the tag and close it
        rtfText.SelText = Text & "></" & Text & ">"
        ' Move the caret in between the two tags
        rtfText.SelStart = rtfText.SelStart - Len("</" & Text & ">")
    Else
        ' Add the tag without closing it
        rtfText.SelText = Text & ">"
    End If
    lsMain.Visible = False
    rtfText.SetFocus
End Sub

Private Sub lsMain_Escape()
On Error Resume Next
    ' Hide the popup window and dont add the text
    lsMain.Visible = False
    rtfText.SetFocus
End Sub

Private Sub rtfText_KeyPress(KeyAscii As Integer)
On Error Resume Next
 If fMainForm.mnuOptionsComplete.Checked = True Then
    If KeyAscii = vbKeyLessThan Then
        ' Get the position of the caret
        GetCaretPos pt
        ' Get the selstart
        lngStart = rtfText.SelStart
        ' Move the popup window to the caret
        'lsMain.Move pt.x + rtfText.Font.Size, pt.y + (2 * rtfText.Font.Size)
        ' Check if the popup window is within the form
        'If lsMain.Left + lsMain.Width > ScaleWidth Then lsMain.Move pt.x - lsMain.Width
        'If lsMain.Top + lsMain.Height > ScaleHeight Then lsMain.Move lsMain.Left, pt.y - lsMain.Height
        ' Fill the popup window with tags (only if there are no errors!)
        If lsMain.FillWithTags(App.Path & "\tags.lst", fMainForm.mnuOptionsUpper.Checked) = 0 Then Exit Sub
        ' Fill the popup window with fonts 'lsMain.FillWithFonts ' Fill the popup window with available drives
        'lsMain.FillWithDrives mnuOptionsUpper.Checked ' Show the popup window
        lsMain.Visible = True
        ' Give the window focus
        lsMain.SetFocus
    End If
   Else
 End If
End Sub


Public Sub Undo(ByVal bUndo As Boolean)
On Error Resume Next
    On Error Resume Next
    Dim OK As Long
    
    OK = SendMessageLong(Screen.ActiveForm.ActiveControl.hwnd, EM_UNDO, 0&, 0&)
    If (bUndo) Then
        'mnuRightUndo.Enabled = False
        'mnuUndo.Enabled = False
        'mnuRightRedo.Enabled = True
        'mnuRedo.Enabled = True
        'Toolbar1.Buttons(16).Enabled = False
        'Toolbar1.Buttons(17).Enabled = True
    Else
        'mnuRightUndo.Enabled = True
        'mnuUndo.Enabled = True
        'mnuRightRedo.Enabled = False
        'mnuRedo.Enabled = False
        'Toolbar1.Buttons(16).Enabled = True
        'Toolbar1.Buttons(17).Enabled = False
    End If
    
    Exit Sub
End Sub


