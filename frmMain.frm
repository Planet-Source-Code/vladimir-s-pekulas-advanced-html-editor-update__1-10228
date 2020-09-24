VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Project1"
   ClientHeight    =   7845
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10290
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9630
      Top             =   2475
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox PTAB 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   7215
      Left            =   0
      ScaleHeight     =   7215
      ScaleWidth      =   2340
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   360
      Width           =   2340
      Begin VB.PictureBox PicCloseTab 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   2055
         Picture         =   "frmMain.frx":0E42
         ScaleHeight     =   135
         ScaleWidth      =   180
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Hide File Manager"
         Top             =   15
         Width           =   210
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   135
         Index           =   1
         Left            =   15
         Picture         =   "frmMain.frx":0FC8
         ScaleHeight     =   135
         ScaleWidth      =   2175
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   15
         Width           =   2175
      End
      Begin TabDlg.SSTab XTAB 
         Height          =   6960
         Left            =   45
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   210
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   12277
         _Version        =   393216
         Style           =   1
         TabHeight       =   520
         WordWrap        =   0   'False
         ShowFocusRect   =   0   'False
         TabCaption(0)   =   "Files"
         TabPicture(0)   =   "frmMain.frx":222A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "DRV"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "DIR"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "FILE"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "ListFiles"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Sizer"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).ControlCount=   5
         TabCaption(1)   =   "Java"
         TabPicture(1)   =   "frmMain.frx":2246
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "JavaList"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Snip"
         TabPicture(2)   =   "frmMain.frx":2262
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "SnippList"
         Tab(2).ControlCount=   1
         Begin VB.PictureBox Sizer 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   585
            Left            =   720
            ScaleHeight     =   585
            ScaleWidth      =   630
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   5715
            Visible         =   0   'False
            Width           =   630
         End
         Begin ComctlLib.ListView ListFiles 
            Height          =   3735
            Left            =   135
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   2970
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   6588
            View            =   2
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   327682
            SmallIcons      =   "ImageList1"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   0
         End
         Begin MSComctlLib.ListView JavaList 
            Height          =   6285
            Left            =   -74865
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   435
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   11086
            View            =   3
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            SmallIcons      =   "imlToolbarIcons"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Java Snippets"
               Object.Width           =   3528
            EndProperty
         End
         Begin VB.FileListBox FILE 
            Appearance      =   0  'Flat
            Height          =   3735
            Left            =   120
            ReadOnly        =   0   'False
            System          =   -1  'True
            TabIndex        =   6
            TabStop         =   0   'False
            ToolTipText     =   "Available Files"
            Top             =   2955
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.DirListBox DIR 
            Appearance      =   0  'Flat
            Height          =   2115
            Left            =   105
            TabIndex        =   5
            TabStop         =   0   'False
            ToolTipText     =   "Available Directories"
            Top             =   750
            Width           =   1935
         End
         Begin VB.DriveListBox DRV 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   120
            TabIndex        =   4
            TabStop         =   0   'False
            ToolTipText     =   "Available Drives"
            Top             =   390
            Width           =   1935
         End
         Begin MSComctlLib.ListView SnippList 
            Height          =   6285
            Left            =   -74865
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   435
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   11086
            View            =   3
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            SmallIcons      =   "imlToolbarIcons"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "HTML Snippets"
               Object.Width           =   3528
            EndProperty
         End
      End
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10290
      _ExtentX        =   18150
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   26
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print"
            ImageKey        =   "Print"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bold"
            Object.ToolTipText     =   "Bold"
            ImageKey        =   "Bold"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Italic"
            Object.ToolTipText     =   "Italic"
            ImageKey        =   "Italic"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Align Left"
            Object.ToolTipText     =   "Align Left"
            ImageKey        =   "Align Left"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Center"
            Object.ToolTipText     =   "Center"
            ImageKey        =   "Center"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Align Right"
            Object.ToolTipText     =   "Align Right"
            ImageKey        =   "Align Right"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Image Map"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Spell Check"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Char"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Tags"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "TagsIns"
            ImageIndex      =   23
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Colors"
            ImageIndex      =   24
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "html"
                  Text            =   "HTML Coloring"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "vb"
                  Text            =   "VB Coloring"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "java"
                  Text            =   "Java Coloring"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
      Begin VB.ComboBox CoFonts 
         Height          =   315
         Left            =   7440
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   0
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   1
      Top             =   7575
      Width           =   10290
      _ExtentX        =   18150
      _ExtentY        =   476
      ShowTips        =   0   'False
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12515
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "7:24 PM"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   9645
      Top             =   465
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   9600
      Top             =   1110
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   24
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":227E
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2390
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":24A2
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":25B4
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":26C6
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":27D8
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":28EA
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":29FC
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2B0E
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2C20
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2D32
            Key             =   "Align Left"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E44
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F56
            Key             =   "Align Right"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3068
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":34BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3910
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3F50
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":43A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":44B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":490C
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4A20
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4E74
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":52C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":571C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   9585
      Top             =   1785
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":5B70
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":5E8A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFILE 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu Template 
         Caption         =   "&Template"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnusep 
         Caption         =   "-"
      End
      Begin VB.Menu OpenWeb 
         Caption         =   "&Open from The Web ..."
      End
      Begin VB.Menu ConvertTextFile 
         Caption         =   "&Convert text file"
      End
      Begin VB.Menu mnusepp 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileSaveAll 
         Caption         =   "Save A&ll"
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePageSetup 
         Caption         =   "Page Set&up..."
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSend 
         Caption         =   "Sen&d..."
      End
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar5 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
      End
      Begin VB.Menu mnuEditBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu SelectAll 
         Caption         =   "&Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu MNUQ 
         Caption         =   "-"
      End
      Begin VB.Menu SelectWrod 
         Caption         =   "&Select Word"
         Shortcut        =   ^L
      End
      Begin VB.Menu SentenceNow 
         Caption         =   "&Select Sentence"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuerer 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditWordWrap 
         Caption         =   "Word Wrap"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Status &Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu FileManager 
         Caption         =   "&File Manager"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Open Doc"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowNewWindow 
         Caption         =   "&New Doc"
      End
   End
   Begin VB.Menu Syn 
      Caption         =   "&Syntaxing"
      Begin VB.Menu mnuOptionsLower 
         Caption         =   "&Lower Case"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuOptionsUpper 
         Caption         =   "&Upper Case"
      End
      Begin VB.Menu mnuOptionsComplete 
         Caption         =   "&Syntaxing ON/OFF"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuHelpSearchForHelpOn 
         Caption         =   "&Search For Help On..."
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About "
      End
   End
   Begin VB.Menu Files 
      Caption         =   "Files"
      Visible         =   0   'False
      Begin VB.Menu Filter 
         Caption         =   "&Filter"
         Begin VB.Menu HTM 
            Caption         =   "&HTML Files"
         End
         Begin VB.Menu Pictures 
            Caption         =   "&Picture Files"
         End
         Begin VB.Menu TXT 
            Caption         =   "&Cgi &Pl &Txt Files"
         End
         Begin VB.Menu m 
            Caption         =   "-"
         End
         Begin VB.Menu Custom 
            Caption         =   "&Custom ..."
         End
      End
      Begin VB.Menu Sort 
         Caption         =   "&Sort Files"
         Begin VB.Menu ABC 
            Caption         =   "&ABCDE"
            Enabled         =   0   'False
         End
         Begin VB.Menu EDC 
            Caption         =   "&EDCBA"
         End
      End
      Begin VB.Menu Style 
         Caption         =   "&Style"
         Begin VB.Menu Longa 
            Caption         =   "&Long Style"
         End
         Begin VB.Menu Shorta 
            Caption         =   "&Short Style"
            Enabled         =   0   'False
         End
      End
   End
   Begin VB.Menu Java 
      Caption         =   "Java"
      Visible         =   0   'False
      Begin VB.Menu Add 
         Caption         =   "&Add New Java Script"
      End
      Begin VB.Menu Refresh 
         Caption         =   "&Refresh List"
      End
      Begin VB.Menu mnus 
         Caption         =   "-"
      End
      Begin VB.Menu sABC 
         Caption         =   "&Sorting ABC"
      End
      Begin VB.Menu sCBA 
         Caption         =   "&Sorting CBA"
      End
   End
   Begin VB.Menu Snippets 
      Caption         =   "Snippets"
      Visible         =   0   'False
      Begin VB.Menu Addsnipp 
         Caption         =   "&Add Snippet"
      End
      Begin VB.Menu RefreshSnipp 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnuss 
         Caption         =   "-"
      End
      Begin VB.Menu ABCSnip 
         Caption         =   "&Sorting ABC"
      End
      Begin VB.Menu CBAsnip 
         Caption         =   "&Sorting CBA"
      End
   End
   Begin VB.Menu RTFMenu 
      Caption         =   "RTFMenu"
      Visible         =   0   'False
      Begin VB.Menu Edittag 
         Caption         =   "&Edit Current Tag"
         Shortcut        =   ^E
      End
      Begin VB.Menu InsertTag 
         Caption         =   "&Insert Tag"
         Shortcut        =   ^I
      End
      Begin VB.Menu MNUUUUU 
         Caption         =   "-"
      End
      Begin VB.Menu Fileopt 
         Caption         =   "&File"
         Begin VB.Menu NewDoc1 
            Caption         =   "&New"
         End
         Begin VB.Menu OpenDoc 
            Caption         =   "&Open"
         End
         Begin VB.Menu SaveDoc 
            Caption         =   "&Save"
         End
         Begin VB.Menu SaveAsDoc 
            Caption         =   "&Save As"
         End
      End
      Begin VB.Menu mnunn2 
         Caption         =   "-"
      End
      Begin VB.Menu Copy1 
         Caption         =   "&Copy"
      End
      Begin VB.Menu Paste1 
         Caption         =   "&Paste"
      End
      Begin VB.Menu Cut1 
         Caption         =   "&Cut"
      End
      Begin VB.Menu Date 
         Caption         =   "&Time/Date"
         Shortcut        =   +^{F1}
      End
      Begin VB.Menu mnuuu1 
         Caption         =   "-"
      End
      Begin VB.Menu NewDoc 
         Caption         =   "&New Document"
      End
      Begin VB.Menu CloseDoc 
         Caption         =   "&Close Current Document"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
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
      '**  SEE frmAbout FOR FULL CREDITS ! **
Public CustomExtention As String
Public FileType As String
Private Type ViewSnipps
    intID As Integer
    strTitle As String * 99
    strArtist As String * 100
End Type
Public lDocumentCount As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Const EM_UNDO = &HC7

Private Sub ABC_Click()
 ListFiles.SortOrder = lvwAscending
 ABC.Enabled = False
 EDC.Enabled = True
End Sub

Private Sub ABCSnip_Click()
 SnippList.SortOrder = lvwAscending
End Sub

Private Sub Add_Click()
 frmJava.Show 1, fMainForm
End Sub

Private Sub Addsnipp_Click()
 frmSnipp.Show 1, fMainForm
End Sub

Private Sub CBAsnip_Click()
 SnippList.SortOrder = lvwDescending
End Sub


Private Sub CloseDoc_Click()
On Error Resume Next
  If MsgBox("Are you sure to close current document ?", vbQuestion + vbYesNo, "Close document ?") = vbYes Then
        lDocumentCount = lDocumentCount - 1
        Unload fMainForm.ActiveForm
   Else
        Exit Sub
   End If
End Sub

Private Sub CoFonts_Click()
 On Error Resume Next
 ActiveForm.rtfText.SelText = "<font face=" & Chr(34) & CoFonts.Text & Chr(34) & ">"
End Sub

Private Sub ConvertTextFile_Click()
On Error Resume Next
Dim intFileNum As Integer
Dim strFilename As String
 
 With dlgCommonDialog
  .ShowOpen
  sFile = .Filename
 End With
  
  LoadNewDoc
  'ActiveForm.rtfText3.Language = hlNOHighLight
  ActiveForm.rtfText3.Text = "<html>" & vbCrLf & "<head>" & vbCrLf & "<title>Untitled</title>" & vbCrLf & "</head>" & vbCrLf & vbCrLf & "<body bgcolor=" & Chr(34) & "white" & Chr(34) & ">" & vbCrLf & vbCrLf
intFileNum = FreeFile
Open sFile For Input As #intFileNum
Do While Not EOF(intFileNum)
 Line Input #intFileNum, strTextLine
 ActiveForm.rtfText3.Text = ActiveForm.rtfText3.Text & strTextLine & "<br>" & vbCrLf
Loop
ActiveForm.rtfText3.Text = ActiveForm.rtfText3.Text & vbCrLf & "<br>"
Close #intFileNum
 ActiveForm.rtfText.Text = ActiveForm.rtfText3.Text
End Sub

Private Sub Copy1_Click()
 On Error Resume Next
 Clipboard.SetText ActiveForm.rtfText.SelText
End Sub

Private Sub Cut1_Click()
 On Error Resume Next
 Clipboard.SetText ActiveForm.rtfText.SelText
 ActiveForm.rtfText.SelText = vbNullString
End Sub

Private Sub Date_Click()
 ActiveForm.rtfText.SelText = Now
End Sub

Private Sub DIR_Change()
FILE.Path = DIR.Path
ListFilesWicons
End Sub

Private Sub DRV_Change()
DIR.Path = DRV.Drive
End Sub


Private Sub EDC_Click()
 ListFiles.SortOrder = lvwDescending
 ABC.Enabled = True
 EDC.Enabled = False
End Sub

Private Sub Edittag_Click()
On Error Resume Next
Dim TagRec As String
Dim FullTag As String
  Me.ActiveForm.rtfText.Span "<", False, True        ' Select Full Tag
  Me.ActiveForm.rtfText.Span ">", True, True         ' Select Full Tag
  TagRec = "<" & Me.ActiveForm.rtfText.SelText & ">" ' Add <> to the tag

  TagRec = Mid(Me.ActiveForm.rtfText.SelText, 1, 2)
    If Mid(TagRec, 1, 1) = "/" Then
      MsgBox "Can't edit the end of TAG", vbInformation
      Exit Sub
    End If
'body
    If LCase(TagRec) = "bo" Then
     TabNumber = 1
     frmTagEdit.Show 1, fMainForm
     Exit Sub
    End If
'div
    If LCase(TagRec) = "di" Then
     TabNumber = 3
     frmTagEdit.Show 1, fMainForm
     Exit Sub
    End If
'anchor
    If LCase(TagRec) = "a " Then
     TabNumber = 0
     frmTagEdit.Show 1, fMainForm
     Exit Sub
    End If
'font
    If LCase(TagRec) = "fo" Then
     TabNumber = 4
     frmTagEdit.Show 1, fMainForm
     Exit Sub
    End If
'img
    If LCase(TagRec) = "im" Then
     TabNumber = 6
     frmTagEdit.Show 1, fMainForm
     Exit Sub
    End If
'select
    If LCase(TagRec) = "se" Then
     TabNumber = 8
     frmTagEdit.Show 1, fMainForm
     Exit Sub
    End If
'hr
    If LCase(TagRec) = "hr" Then
     TabNumber = 5
     frmTagEdit.Show 1, fMainForm
     Exit Sub
    End If
'textarea
    If LCase(TagRec) = "te" Then
     TabNumber = 10
     frmTagEdit.Show 1, fMainForm
     Exit Sub
    End If
'table
    If LCase(TagRec) = "ta" Then
     TabNumber = 13
     frmTagEdit.Show 1, fMainForm
     Exit Sub
    End If
'Td
    If LCase(TagRec) = "td" Then
     TabNumber = 14
     frmTagEdit.Show 1, fMainForm
     Exit Sub
    End If
'tr
    If LCase(TagRec) = "tr" Then
     TabNumber = 15
     frmTagEdit.Show 1, fMainForm
     Exit Sub
    End If
  
  If LCase(TagRec) = "in" Then
    TagRecNew = LCase(Mid(Me.ActiveForm.rtfText.SelText, 13, 3))

'Submit
     If TagRecNew = Chr(34) & "sub" Then
      TabNumber = 9
      frmTagEdit.Show 1, fMainForm
      Exit Sub
     End If
     If TagRecNew = "sub" Then
      TabNumber = 9
      frmTagEdit.Show 1, fMainForm
      Exit Sub
     End If
'Radio
     If TagRecNew = Chr(34) & "rad" Then
      TabNumber = 7
      frmTagEdit.Show 1, fMainForm
      Exit Sub
     End If
     If TagRecNew = "rad" Then
      TabNumber = 7
      frmTagEdit.Show 1, fMainForm
      Exit Sub
     End If
'Checkbox
     If TagRecNew = Chr(34) & "che" Then
      TabNumber = 2
      frmTagEdit.Show 1, fMainForm
      Exit Sub
     End If
     If TagRecNew = "che" Then
      TabNumber = 2
      frmTagEdit.Show 1, fMainForm
      Exit Sub
     End If
'text
     If TagRecNew = Chr(34) & "tex" Then
      TabNumber = 12
      frmTagEdit.Show 1, fMainForm
      Exit Sub
     End If
     If TagRecNew = "tex" Then
      TabNumber = 12
      frmTagEdit.Show 1, fMainForm
      Exit Sub
     End If
'hidden
     If TagRecNew = Chr(34) & "hid" Then
      TabNumber = 12
      frmTagEdit.Show 1, fMainForm
      Exit Sub
     End If
     If TagRecNew = "hid" Then
      TabNumber = 12
      frmTagEdit.Show 1, fMainForm
      Exit Sub
     End If
  End If
If TabNumber = 0 Then MsgBox "Tag not supported", vbInformation
' End If
End Sub

Private Sub FILE_DblClick()
 OpenAFile
End Sub

Private Sub FILE_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = 2 Then
   Me.PopupMenu Files
  End If
End Sub

Private Sub FileManager_Click()
FileManager.Checked = Not FileManager.Checked
PTAB.Visible = FileManager.Checked
End Sub

Private Sub HTM_Click()
 Dim All As Integer
 All = FILE.ListCount - 1
  ListFiles.ListItems.Clear
 For I = 0 To All
   
   SplitName = FILE.List(I)
   Extension = vbNullString
   intPos = Len(SplitName)
  Do While intPos > 0
   Select Case Mid$(SplitName, intPos, 1)
   Case "."
   Extension = Mid$(SplitName, intPos + 1)
   Exit Do
   Case Else
   End Select
   intPos = intPos - 1
  Loop
           If Extension = "html" Then ListFiles.ListItems.Add , , FILE.List(I), , 2
           If Extension = "htm" Then ListFiles.ListItems.Add , , FILE.List(I), , 2
           If Extension = "HTML" Then ListFiles.ListItems.Add , , FILE.List(I), , 2
           If Extension = "Html" Then ListFiles.ListItems.Add , , FILE.List(I), , 2
           If Extension = "HTM" Then ListFiles.ListItems.Add , , FILE.List(I), , 2
           If Extension = "Htm" Then ListFiles.ListItems.Add , , FILE.List(I), , 2
 Next I
End Sub

Private Sub InsertTag_Click()
 frmTags.Show 1, fMainForm
End Sub

Private Sub JavaList_DblClick()
 Dim OurPath As String
If JavaList.SelectedItem.Text = "News Ticker Script" Then
    OurPath = App.Path & "\Java\news.html"
    LoadNewDoc
    fMainForm.ActiveForm.rtfText.LoadFile OurPath
    Exit Sub
End If


If JavaList.SelectedItem.Text = "Email Form Script" Then
    OurPath = App.Path & "\Java\email.html"
    LoadNewDoc
    fMainForm.ActiveForm.rtfText.LoadFile OurPath
    Exit Sub
End If

If JavaList.SelectedItem.Text = "Status Bar Script" Then
    OurPath = App.Path & "\Java\statusbar.html"
    LoadNewDoc
    fMainForm.ActiveForm.rtfText.LoadFile OurPath
    Exit Sub
End If

If JavaList.SelectedItem.Text = "Counter Script" Then
    OurPath = App.Path & "\Java\counter.html"
    LoadNewDoc
    fMainForm.ActiveForm.rtfText.LoadFile OurPath
    Exit Sub
End If

If JavaList.SelectedItem.Text = "Gallery Script" Then
    OurPath = App.Path & "\Java\gallery.html"
    LoadNewDoc
    fMainForm.ActiveForm.rtfText.LoadFile OurPath
    Exit Sub
End If

If JavaList.SelectedItem.Text = "IP Address Script" Then
    OurPath = App.Path & "\Java\ip.html"
    LoadNewDoc
    fMainForm.ActiveForm.rtfText.LoadFile OurPath
    Exit Sub
End If

If JavaList.SelectedItem.Text = "Text Effect Script" Then
    OurPath = App.Path & "\Java\textfx.html"
    LoadNewDoc
    fMainForm.ActiveForm.rtfText.LoadFile OurPath
    Exit Sub
End If

If JavaList.SelectedItem.Text = "Redirection Script" Then
    OurPath = App.Path & "\Java\redirect.html"
    LoadNewDoc
    fMainForm.ActiveForm.rtfText.LoadFile OurPath
    Exit Sub
End If

If JavaList.SelectedItem.Text = "Resolution Script" Then
    OurPath = App.Path & "\Java\resolution.html"
    LoadNewDoc
    fMainForm.ActiveForm.rtfText.LoadFile OurPath
    Exit Sub
End If

If JavaList.SelectedItem.Text = "Scroller Script" Then
    OurPath = App.Path & "\Java\scroller.html"
    LoadNewDoc
    fMainForm.ActiveForm.rtfText.LoadFile OurPath
    Exit Sub
End If

' Now we are going to extract the path of snippet file
' that's going to be open. That's for Users added snipps.
 Dim udtCDToView As ViewSnipps
 Dim intCDFile As Integer, lngRecLength As Long
 Dim lngTotalRecords As Long, lngCDID As Long
 Dim intFileNum As Integer
 Dim NumRecords As Long
 '
 Dim Text As String
 intFileNum = FreeFile
 'Open File
 intCDFile = FreeFile
 lngRecLength = LenB(udtCDToView)
 Open App.Path & "\JavaIndex.dat" For Random As #intCDFile Len = lngRecLength

 '# of Rec.
 If LOF(intFileNum) Mod lngRecLength = 0 Then
 NumRecords = (LOF(intCDFile) \ lngRecLength)
 Else
 NumRecords = (LOF(intCDFile) \ lngRecLength) + 1
 End If
 lngTotalRecords = NumRecords

 'View Rec if Valid
     If lngTotalRecords = 0 Then
 MsgBox "Error 001 - Can not read the record"
 End If
 lngCDID = 0
 Do
 lngCDID = lngCDID + 1
 Get #intCDFile, lngCDID, udtCDToView
 If udtCDToView.strTitle = JavaList.SelectedItem.Text Then
 Trim udtCDToView.strArtist
 LoadNewDoc
 fMainForm.ActiveForm.rtfText.LoadFile udtCDToView.strArtist
 Exit Sub
 End If
 ' udtCDToView.strTitle  ' That's what we search for
 ' udtCDToView.strArtist ' and this is the path to the file
 Loop
 Close #intCDFile
End Sub

Private Sub JavaList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
 Me.PopupMenu Java
End If
End Sub

Function OpenAFile()
Dim intPos As Integer
 'Get Path First
 If DIR.Path = "C:\" Then
  MyPath = Mid(DIR.Path, 1, 2)
 Else
  MyPath = DIR.Path
 End If
 '# Is it a Picture ?
 
 If FileType = "List" Then SplitName = ListFiles.SelectedItem.Text
 If FileType = "FILE" Then SplitName = FILE.Filename
 
 'Get The Extention
 Extension = vbNullString
 intPos = Len(SplitName)
 Do While intPos > 0
      Select Case Mid$(SplitName, intPos, 1)
          Case "."
            Extension = Mid$(SplitName, intPos + 1)
            Exit Do
          Case Else
      End Select
  intPos = intPos - 1
 Loop
If Extension = "GIF" Then
 Sizer.Picture = LoadPicture(MyPath & "\" & SplitName)
 PicRatioW = Int(Sizer.Width * (0.064367816091954) - 1.33)
 PicRatioH = Int(Sizer.Height * (0.064367816091954) - 1.33)
 ActiveForm.rtfText.SelText = "<img src=" & Chr(34) & MyPath & "\" & SplitName & Chr(34) & " width=" & Chr(34) & PicRatioW & Chr(34) & " height=" & Chr(34) & PicRatioH & Chr(34) & " alt=" & Chr(34) & Chr(34) & " Border=" & Chr(34) & "0" & Chr(34) & ">"
 Exit Function
End If
If Extension = "gif" Then
 Sizer.Picture = LoadPicture(MyPath & "\" & SplitName)
 PicRatioW = Int(Sizer.Width * (0.064367816091954) - 1.33)
 PicRatioH = Int(Sizer.Height * (0.064367816091954) - 1.33)
 ActiveForm.rtfText.SelText = "<img src=" & Chr(34) & MyPath & "\" & SplitName & Chr(34) & " width=" & Chr(34) & PicRatioW & Chr(34) & " height=" & Chr(34) & PicRatioH & Chr(34) & " alt=" & Chr(34) & Chr(34) & " Border=" & Chr(34) & "0" & Chr(34) & ">"
 Exit Function
End If
If Extension = "jpg" Then
 Sizer.Picture = LoadPicture(MyPath & "\" & SplitName)
 PicRatioW = Int(Sizer.Width * (0.064367816091954) - 1.33)
 PicRatioH = Int(Sizer.Height * (0.064367816091954) - 1.33)
 ActiveForm.rtfText.SelText = "<img src=" & Chr(34) & MyPath & "\" & SplitName & Chr(34) & " width=" & Chr(34) & PicRatioW & Chr(34) & " height=" & Chr(34) & PicRatioH & Chr(34) & " alt=" & Chr(34) & Chr(34) & " Border=" & Chr(34) & "0" & Chr(34) & ">"
 Exit Function
End If
If Extension = "JPG" Then
 Sizer.Picture = LoadPicture(MyPath & "\" & SplitName)
 PicRatioW = Int(Sizer.Width * (0.064367816091954) - 1.33)
 PicRatioH = Int(Sizer.Height * (0.064367816091954) - 1.33)
 ActiveForm.rtfText.SelText = "<img src=" & Chr(34) & MyPath & "\" & SplitName & Chr(34) & " width=" & Chr(34) & PicRatioW & Chr(34) & " height=" & Chr(34) & PicRatioH & Chr(34) & " alt=" & Chr(34) & Chr(34) & " Border=" & Chr(34) & "0" & Chr(34) & ">"
 Exit Function
End If
 frmDocument.rtfText.LoadFile MyPath & "\" & SplitName
 lDocumentCount = lDocumentCount + 1
End Function

Private Sub ListFiles_DblClick()
 OpenAFile
End Sub

Private Sub ListFiles_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 If Button = 2 Then Me.PopupMenu Files
End Sub

Private Sub Longa_Click()
 FileType = "FILE"
 FILE.Visible = True
 ListFiles.Visible = False
 Filter.Enabled = False
 Sort.Enabled = False
 Shorta.Enabled = True
 Longa.Enabled = False
End Sub


Private Sub MDIForm_Load()
On Error GoTo ErrorHand:
'################################################################################
'# The Following statemnet will basicly run only he first time you run this app #
'# it will create a file containing the names of all the font on your PC        #
'# and load it to the invisible combo box from which the font names will be     #
'# transfred wherever necessary.                                                #
'#                                                                              #
'# It is way (I mean WAY!) faster this way then loading the fonts each time     #
'# you need them by screen.font ...                                             #
'################################################################################
        '# Check if we have created the file
        Dim intFileNum As Integer, strFilename As String
        strFilename = App.Path & "\fontsQ.txt"
        intFileNum = FreeFile
        Open strFilename For Input As #intFileNum
         Do While Not EOF(intFileNum)
              Line Input #intFileNum, FontQ
              If FontQ = "" Then GoTo Continue:
         Loop
Continue:
        Close #intFileNum
        '##
             If Trim(FontQ) = "" Then
        '# Create a file with all the fonts
        strFilename = App.Path & "\fonts.txt"
        intFileNum = FreeFile
        Open strFilename For Output As #intFileNum
        For I = 1 To Screen.FontCount
        Print #intFileNum, Screen.Fonts(I)
         Next I
        Close #intFileNum
        'Close it
        Open App.Path & "\fontsQ.txt" For Output As #intFileNum
            Print #intFileNum, "1"
        GoTo LoadFonts:
        '##
            Else
        '# Addd fonts to the combo box
LoadFonts:
        Close #intFileNum
        strFilename = App.Path & "\fonts.txt"
        intFileNum = FreeFile
        Open strFilename For Input As #intFileNum
        Do While Not EOF(intFileNum)
            Line Input #intFileNum, FontNameA
            If Trim(FontNameA) = "" Then GoTo Contin:
            CoFonts.AddItem FontNameA
Contin:
        Loop
        Close #intFileNum
         CoFonts.ListIndex = 0
                End If
'###########################################################################
    TabNumber = 0
    FileType = "List"
    LoadNewDoc
    Refresh_Click
    RefreshSnipp_Click
    ListFilesWicons
Exit Sub
ErrorHand:
    MsgBox "An Error has occured: " & Err.Description & " = " & Err.Number
End Sub

Function ListFilesWicons()
On Error Resume Next
Dim All As Integer
 All = FILE.ListCount - 1
 ListFiles.ListItems.Clear
 For I = 0 To All
  ListFiles.ListItems.Add , , FILE.List(I), , 2
 Next I
End Function

Public Sub LoadNewDoc()
On Error Resume Next
    Dim frmD As frmDocument
    lDocumentCount = lDocumentCount + 1
    Set frmD = New frmDocument
    frmD.Show
    ActiveForm.rtfText.Text = "<!DOCTYPE HTML PUBLIC " & Chr(34) & "-//W3C//CASPER HTML Editor//EN" & Chr(34) & ">" & vbCrLf
    ActiveForm.rtfText.Text = ActiveForm.rtfText.Text & vbCrLf & "<html>" & vbCrLf & "<head>" & vbCrLf & "         <title>Untitled</title>" & vbCrLf & "</head>" & vbCrLf & vbCrLf & "<body>" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "</body>" & vbCrLf & "</html>" & vbCrLf
End Sub

Function LoadNewDocFunction()
On Error Resume Next
    Dim frmD As frmDocument
    fMainForm.lDocumentCount = fMainForm.lDocumentCount + 1
    Set frmD = New frmDocument
    frmD.Show
    ActiveForm.rtfText.Text = "<!DOCTYPE HTML PUBLIC " & Chr(34) & "-//W3C//CASPER HTML Editor//EN" & Chr(34) & ">" & vbCrLf
    ActiveForm.rtfText.Text = ActiveForm.rtfText.Text & "<html>" & vbCrLf & "<head>" & vbCrLf & "         <title>Untitled</title>" & vbCrLf & "</head>" & vbCrLf & vbCrLf & "<body>" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "</body>" & vbCrLf & "</html>" & vbCrLf
End Function

Private Sub MDIForm_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 On Error Resume Next
 If Button = 2 Then
     CloseDoc.Enabled = False
     Paste1.Enabled = False
     Cut1.Enabled = False
     Copy1.Enabled = False
     InsertTag.Enabled = False
     Edittag.Enabled = False
     Fileopt.Enabled = False
     Date.Enabled = False
     Me.PopupMenu RTFMenu
 End If
End Sub

Private Sub MDIForm_Resize()
On Error Resume Next
XTAB.Height = Me.Height - 1625
FILE.Height = XTAB.Height - 3000
ListFiles.Left = FILE.Left
ListFiles.Top = FILE.Top
ListFiles.Height = FILE.Height
JavaList.Height = XTAB.Height - 600
SnippList.Height = XTAB.Height - 600
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
 End
End Sub


Private Sub mnuHelpAbout_Click()
frmAbout.Show 1, fMainForm
End Sub

Private Sub mnuOptionsComplete_Click()
 mnuOptionsComplete.Checked = Not mnuOptionsComplete.Checked
End Sub

Private Sub mnuOptionsLower_Click()
    mnuOptionsLower.Checked = Not mnuOptionsLower.Checked
    mnuOptionsUpper.Checked = Not mnuOptionsUpper.Checked
End Sub

Private Sub mnuOptionsUpper_Click()
 mnuOptionsLower_Click
End Sub

Private Sub NewDoc_Click()
    LoadNewDoc
End Sub

Private Sub NewDoc1_Click()
 LoadNewDoc
End Sub

Private Sub OpenDoc_Click()
 mnuFileOpen_Click
End Sub

Private Sub OpenWeb_Click()
 frmOpenWWW.Show 1, fMainForm
End Sub

Private Sub Paste1_Click()
 On Error Resume Next
 ActiveForm.rtfText.SelText = Clipboard.GetText
End Sub

Private Sub PicCloseTab_Click()
' PTAB.Visible = False
' FileManager.Checked = False
GetSrcProperty
End Sub

Private Sub Pictures_Click()
On Error Resume Next
 Dim All As Integer
 All = FILE.ListCount - 1
  ListFiles.ListItems.Clear
 For I = 0 To All
   
   SplitName = FILE.List(I)
   Extension = vbNullString
   intPos = Len(SplitName)
  Do While intPos > 0
   Select Case Mid$(SplitName, intPos, 1)
   Case "."
   Extension = Mid$(SplitName, intPos + 1)
   Exit Do
   Case Else
   End Select
   intPos = intPos - 1
  Loop
           If Extension = "GIF" Then ListFiles.ListItems.Add , , FILE.List(I), , 2
           If Extension = "Gif" Then ListFiles.ListItems.Add , , FILE.List(I), , 2
           If Extension = "gif" Then ListFiles.ListItems.Add , , FILE.List(I), , 2
           If Extension = "jpg" Then ListFiles.ListItems.Add , , FILE.List(I), , 2
           If Extension = "JPG" Then ListFiles.ListItems.Add , , FILE.List(I), , 2
           If Extension = "Jpg" Then ListFiles.ListItems.Add , , FILE.List(I), , 2
 Next I
End Sub

Private Sub Refresh_Click()
On Error Resume Next
 Dim IRef As Integer
 'Delete it first !
 For IRef = 1 To JavaList.ListItems.Count
  JavaList.ListItems.Remove (1)
 Next IRef
 ' Load it again ! (Our Own)
 JavaList.ListItems.Add , , "Counter Script", , 14
 JavaList.ListItems.Add , , "Gallery Script", , 14
 JavaList.ListItems.Add , , "IP Address Script", , 14
 JavaList.ListItems.Add , , "Text Effect Script", , 14
 JavaList.ListItems.Add , , "Redirection Script", , 14
 JavaList.ListItems.Add , , "Resolution Script", , 14
 JavaList.ListItems.Add , , "Scroller Script", , 14
 JavaList.ListItems.Add , , "Status Bar Script", , 14
 JavaList.ListItems.Add , , "Email Form Script", , 14
 JavaList.ListItems.Add , , "News Ticker Script", , 14
 ' Load it again ! (Users)
 Dim udtJavaToView As ViewSnipps
 Dim intJavaFile As Integer, lngRecLengthJava As Long
 Dim lngTotalRecordsJava As Long, lngJavaID As Long
 Dim intFileNumJava As Integer
 Dim NumRecordsJava As Long
 intFileNumJava = FreeFile
 'Open File
 intJavaFile = FreeFile
 lngRecLengthJava = LenB(udtJavaToView)
 Open App.Path & "\JavaIndex.dat" For Random As #intJavaFile Len = lngRecLengthJava

 
 If LOF(intFileNumJava) Mod lngRecLengthJava = 0 Then
  NumRecordsJava = (LOF(intJavaFile) \ lngRecLengthJava)
 Else
  NumRecordsJava = (LOF(intJavaFile) \ lngRecLengthJava) + 1
 End If
 lngTotalRecordsJava = NumRecordsJava

 'View Rec if Valid
 If lngTotalRecordsJava = 0 Then
 Exit Sub
 End If
 lngJavaID = 0
 Do
     If lngJavaID = lngTotalRecordsJava Then
     Close #intJavaFile
     Exit Sub
     Else
 lngJavaID = lngJavaID + 1
  If lngJavaID > 0 And lngJavaID <= lngTotalRecordsJava Then
 Get #intJavaFile, lngJavaID, udtJavaToView
 JavaList.ListItems.Add , , udtJavaToView.strTitle, , 14
  End If
     End If
 Loop
 Close #intJavaFile
End Sub

Private Sub RefreshSnipp_Click()
On Error Resume Next
 Dim IRef As Integer
 'Delete it first !
 For IRef = 1 To SnippList.ListItems.Count
  SnippList.ListItems.Remove (1)
 Next IRef
 ' Load it again ! (Our Own)
 SnippList.ListItems.Add , , "Bohemia Gift Finder", , 14
 SnippList.ListItems.Add , , "GoTo.com Search Engine", , 14
 SnippList.ListItems.Add , , "InfoSeek.com Search Engine", , 14
 ' Load it again ! (Users)
 Dim udtCDToView As ViewSnipps
 Dim intCDFile As Integer, lngRecLength As Long
 Dim lngTotalRecords As Long, lngCDID As Long
 Dim intFileNum As Integer
 Dim NumRecords As Long
 intFileNum = FreeFile
 'Open File
 intCDFile = FreeFile
 lngRecLength = LenB(udtCDToView)
 Open App.Path & "SnippetIndex.dat" For Random As #intCDFile Len = lngRecLength
 '# of Rec.
 If LOF(intFileNum) Mod lngRecLength = 0 Then
  NumRecords = (LOF(intCDFile) \ lngRecLength)
 Else
  NumRecords = (LOF(intCDFile) \ lngRecLength) + 1
 End If
 lngTotalRecords = NumRecords

 'View Rec if Valid
     If lngTotalRecords = 0 Then
     Exit Sub
     Close #intCDFile
     End If
 lngCDID = 0
 Do
 lngCDID = lngCDID + 1
     If lngCDID > lngTotalRecords Then
 Exit Sub
     Close #intCDFile
     Else
  If lngCDID > 0 And lngCDID <= lngTotalRecords Then
 Get #intCDFile, lngCDID, udtCDToView
 SnippList.ListItems.Add , , udtCDToView.strTitle, , 14
  End If
     End If
 Loop
 Close #intCDFile
End Sub

Private Sub sABC_Click()
 JavaList.SortOrder = lvwAscending
End Sub

Private Sub SaveAsDoc_Click()
 mnuFileSaveAs_Click
End Sub

Private Sub SaveDoc_Click()
 mnuFileSave_Click
End Sub

Private Sub sCBA_Click()
 JavaList.SortOrder = lvwDescending
End Sub

Private Sub SelectAll_Click()
   With fMainForm.ActiveForm.rtfText
        .SetFocus
        .SelStart = 0
        .SelLength = Len(fMainForm.ActiveForm.rtfText.Text)
    End With
End Sub

Private Sub SelectWrod_Click()
 ActiveForm.rtfText.Span " ,;:.?!", True, True
End Sub

Private Sub SentenceNow_Click()
    With ActiveForm.rtfText
        .Span ".?!:", True, True
        .SelLength = .SelLength + 1
    End With
End Sub

Private Sub Shorta_Click()
On Error Resume Next
 FILE.Visible = False
 ListFiles.Visible = True
 Filter.Enabled = True
 Sort.Enabled = True
 Shorta.Enabled = False
 Longa.Enabled = True
End Sub

Private Sub SnippList_DblClick()
On Error Resume Next
' First load up snippets that comes with Casper HTML
If SnippList.SelectedItem.Text = "InfoSeek.com Search Engine" Then
    OurPath = App.Path & "\Snipps\infoseek.html"
    LoadNewDoc
    fMainForm.ActiveForm.rtfText.LoadFile OurPath

    Exit Sub
End If

If SnippList.SelectedItem.Text = "GoTo.com Search Engine" Then
    OurPath = App.Path & "\Snipps\goto.html"
    LoadNewDoc
    fMainForm.ActiveForm.rtfText.LoadFile OurPath
    Exit Sub
End If

If SnippList.SelectedItem.Text = "Bohemia Gift Finder" Then
    OurPath = App.Path & "\Snipps\bohemia.html"
    LoadNewDoc
    fMainForm.ActiveForm.rtfText.LoadFile OurPath
    Exit Sub
End If

' Now we are going to extract the path of snippet file
' that's going to be open.
 Dim udtCDToView As ViewSnipps
 Dim intCDFile As Integer, lngRecLength As Long
 Dim lngTotalRecords As Long, lngCDID As Long
 Dim intFileNum As Integer
 Dim NumRecords As Long
 Dim Text As String
 intFileNum = FreeFile
 'Open File
 intCDFile = FreeFile
 lngRecLength = LenB(udtCDToView)
 Open App.Path & "\SnippetIndex.dat" For Random As #intCDFile Len = lngRecLength

'# of Rec.
 If LOF(intFileNum) Mod lngRecLength = 0 Then
  NumRecords = (LOF(intCDFile) \ lngRecLength)
 Else
  NumRecords = (LOF(intCDFile) \ lngRecLength) + 1
 End If
 lngTotalRecords = NumRecords

 'View Rec if Valid
     If lngTotalRecords = 0 Then
 MsgBox "Error 001 - Can not read the record"
 End If
 lngCDID = 0
 Do
 lngCDID = lngCDID + 1
 Get #intCDFile, lngCDID, udtCDToView
 If udtCDToView.strTitle = SnippList.SelectedItem.Text Then
 Trim udtCDToView.strArtist
 LoadNewDoc
 fMainForm.ActiveForm.rtfText.LoadFile udtCDToView.strArtist
  Exit Sub
 End If
 Loop
 Close #intCDFile
End Sub

Private Sub SnippList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then Me.PopupMenu Snippets
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComCtlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "New"
            LoadNewDoc
        Case "Open"
            mnuFileOpen_Click
        Case "Save"
            mnuFileSave_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Cut"
            mnuEditCut_Click
        Case "Copy"
            mnuEditCopy_Click
        Case "Paste"
            mnuEditPaste_Click
        Case "Bold"
            ActiveForm.rtfText.SelText = "<B> </B>"
        Case "Italic"
            ActiveForm.rtfText.SelText = "<I> </I>"
        Case "Align Left"
            ActiveForm.rtfText.SelText = "<Div align=" & Chr(34) & "Left" & Chr(34) & "> </Div>"
        Case "Center"
            ActiveForm.rtfText.SelText = "<Center> </Center>"
        Case "Align Right"
            ActiveForm.rtfText.SelText = "<Div align=" & Chr(34) & "Right" & Chr(34) & "> </Div>"
        Case "Image Map"
            frmImageMap.Show 1, fMainForm
        Case "Spell Check"
        
        Case "Char"
            frmChar.Show 1, fMainForm
        Case "Tags"
            frmTagEdit.Show 1, fMainForm
        Case "TagsIns"
            frmTags.Show 1, fMainForm
        Case "Colors"
'# Sh...t I forgot how to work with submenus ..... >:)
    End Select
End Sub

Private Sub mnuWindowNewWindow_Click()
    LoadNewDoc
End Sub

Private Sub mnuViewWebBrowser_Click()
    'ToDo: Add 'mnuViewWebBrowser_Click' code.
    MsgBox "Add 'mnuViewWebBrowser_Click' code."
End Sub

Private Sub mnuViewOptions_Click()
 frmSettings.Show
End Sub

Private Sub mnuViewRefresh_Click()
 MsgBox "See readme file"
End Sub

Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    sbStatusBar.Visible = mnuViewStatusBar.Checked
End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    tbToolBar.Visible = mnuViewToolbar.Checked
End Sub

Private Sub mnuEditPasteSpecial_Click()
    'ToDo: Add 'mnuEditPasteSpecial_Click' code.
    MsgBox "Add 'mnuEditPasteSpecial_Click' code."
End Sub

Private Sub mnuEditPaste_Click()
    On Error Resume Next
    ActiveForm.rtfText.SelText = Clipboard.GetText
End Sub

Private Sub mnuEditCopy_Click()
    On Error Resume Next
    Clipboard.SetText ActiveForm.rtfText.SelText
End Sub

Private Sub mnuEditCut_Click()
 On Error Resume Next
 Clipboard.SetText ActiveForm.rtfText.SelText
 ActiveForm.rtfText.SelText = vbNullString
End Sub

Private Sub mnuEditUndo_Click()
Call frmDocument.Undo(True)
End Sub


Private Sub mnuFileExit_Click()
    'unload the form
    Unload Me

End Sub

Private Sub mnuFileSend_Click()
    MsgBox "Huh ? Oh, sorry I felt asleep ...."
End Sub

Private Sub mnuFilePrint_Click()
    On Error Resume Next
    If ActiveForm Is Nothing Then Exit Sub
    

    With dlgCommonDialog
        .DialogTitle = "Print"
        .CancelError = True
        .Flags = cdlPDReturnDC + cdlPDNoPageNums
        If ActiveForm.rtfText.SelLength = 0 Then
            .Flags = .Flags + cdlPDAllPages
        Else
            .Flags = .Flags + cdlPDSelection
        End If
        .ShowPrinter
        If Err <> MSComDlg.cdlCancel Then
            ActiveForm.rtfText.SelPrint .hDC
        End If
    End With

End Sub

Private Sub mnuFilePrintPreview_Click()
    'ToDo: Add 'mnuFilePrintPreview_Click' code.
    MsgBox "Add 'mnuFilePrintPreview_Click' code."
End Sub

Private Sub mnuFilePageSetup_Click()
    On Error Resume Next
    With dlgCommonDialog
        .DialogTitle = "Page Setup"
        .CancelError = True
        .ShowPrinter
    End With

End Sub

Private Sub mnuFileProperties_Click()
    'ToDo: Add 'mnuFileProperties_Click' code.
    MsgBox "Add 'mnuFileProperties_Click' code."
End Sub


Private Sub mnuFileSaveAs_Click()
 SaveDocument
End Sub

Private Sub mnuFileSave_Click()
 SaveDocument
End Sub

Private Sub mnuFileClose_Click()
    'ToDo: Add 'mnuFileClose_Click' code.
    MsgBox "Add 'mnuFileClose_Click' code."
End Sub

Private Sub mnuFileOpen_Click()
    Dim sFile As String
    If ActiveForm Is Nothing Then LoadNewDoc
    With dlgCommonDialog
        .DialogTitle = "Open"
        .CancelError = False
        .Filter = "All Files (*.*)|*.*"
        .ShowOpen
        If Len(.Filename) = 0 Then
            Exit Sub
        End If
        sFile = .Filename
    
'# Recent.dat
 Dim intFileNum As Integer
 Dim strTextLine As String, strFilename As String
 intFileNum = FreeFile
 Open App.Path & "\recent.dat" For Append As #intFileNum
 Print #intFileNum, sFile
 Close #intFileNum
    
    End With
    ActiveForm.rtfText.LoadFile sFile
End Sub

Private Sub mnuFileNew_Click()
    LoadNewDoc
End Sub

Private Sub Template_Click()
 Dim intFileNum As Integer
 Dim strTextLine As String, strFilename As String
 intFileNum = FreeFile
 sDat = "0"
 Open App.Path & "\exitfront.txt" For Output As #intFileNum
 Print #intFileNum, sDat
 Close #intFileNum
 frmFront.Show 1, fMainForm
End Sub

Private Sub TXT_Click()
On Error Resume Next
 Dim All As Integer
 All = FILE.ListCount - 1
  ListFiles.ListItems.Clear
 For I = 0 To All
   
   SplitName = FILE.List(I)
   Extension = vbNullString
   intPos = Len(SplitName)
  Do While intPos > 0
   Select Case Mid$(SplitName, intPos, 1)
   Case "."
   Extension = Mid$(SplitName, intPos + 1)
   Exit Do
   Case Else
   End Select
   intPos = intPos - 1
  Loop
           If Extension = "txt" Then ListFiles.ListItems.Add , , FILE.List(I), , 2
           If Extension = "TXT" Then ListFiles.ListItems.Add , , FILE.List(I), , 2
           If Extension = "pl" Then ListFiles.ListItems.Add , , FILE.List(I), , 2
           If Extension = "PL" Then ListFiles.ListItems.Add , , FILE.List(I), , 2
           If Extension = "Pl" Then ListFiles.ListItems.Add , , FILE.List(I), , 2
           If Extension = "Cgi" Then ListFiles.ListItems.Add , , FILE.List(I), , 2
           If Extension = "CGI" Then ListFiles.ListItems.Add , , FILE.List(I), , 2
           If Extension = "cgi" Then ListFiles.ListItems.Add , , FILE.List(I), , 2
 Next I
End Sub



Private Sub Custom_Click()
On Error Resume Next
Dim All As Integer
'Ask first? Duh ...
ExtensionOwn = InputBox("Enter your own extension such as 'txt'." & vbCrLf & "Case sensitive!", "Custom Extension")
If ExtensionOwn = "" Then Exit Sub
  
  All = FILE.ListCount - 1
  ListFiles.ListItems.Clear
 For I = 0 To All
   SplitName = FILE.List(I)
   Extension = vbNullString
   intPos = Len(SplitName)
  Do While intPos > 0
   Select Case Mid$(SplitName, intPos, 1)
   Case "."
   Extension = Mid$(SplitName, intPos + 1)
   Exit Do
   Case Else
   End Select
   intPos = intPos - 1
  Loop
          If Extension = ExtensionOwn Then ListFiles.ListItems.Add , , FILE.List(I), , 2
 Next I
End Sub

