VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSettings 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Settings ...."
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   480
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":0454
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   5040
      TabIndex        =   2
      Top             =   4080
      Width           =   1485
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   345
      Left            =   6600
      TabIndex        =   1
      Top             =   4080
      Width           =   1485
   End
   Begin VB.Frame Frame1 
      Height          =   3975
      Left            =   105
      TabIndex        =   0
      Top             =   0
      Width           =   7995
      Begin VB.ComboBox CoSize 
         Height          =   315
         Left            =   4920
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   330
         Width           =   975
      End
      Begin MSComDlg.CommonDialog cdc 
         Left            =   840
         Top             =   2760
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.PictureBox Fore 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1920
         ScaleHeight     =   165
         ScaleWidth      =   1425
         TabIndex        =   18
         Top             =   720
         Width           =   1455
      End
      Begin VB.PictureBox fun 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1920
         ScaleHeight     =   165
         ScaleWidth      =   1425
         TabIndex        =   17
         Top             =   1080
         Width           =   1455
      End
      Begin VB.PictureBox Key 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1920
         ScaleHeight     =   165
         ScaleWidth      =   1425
         TabIndex        =   16
         Top             =   1440
         Width           =   1455
      End
      Begin VB.PictureBox oper 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1920
         ScaleHeight     =   165
         ScaleWidth      =   1425
         TabIndex        =   15
         Top             =   1800
         Width           =   1455
      End
      Begin VB.PictureBox Delim 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1920
         ScaleHeight     =   165
         ScaleWidth      =   1425
         TabIndex        =   14
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmdFunction 
         Height          =   195
         Left            =   1440
         TabIndex        =   13
         Top             =   1080
         Width           =   255
      End
      Begin VB.CommandButton cmdKey 
         Height          =   195
         Left            =   1440
         TabIndex        =   12
         Top             =   1440
         Width           =   255
      End
      Begin VB.CommandButton cmdOper 
         Height          =   195
         Left            =   1440
         TabIndex        =   11
         Top             =   1800
         Width           =   255
      End
      Begin VB.CommandButton cmdFore 
         Height          =   195
         Left            =   1440
         TabIndex        =   10
         Top             =   720
         Width           =   255
      End
      Begin VB.CommandButton cmdDelim 
         Height          =   195
         Left            =   1440
         TabIndex        =   9
         Top             =   360
         Width           =   255
      End
      Begin Project1.CodeHighlight TestArea 
         Height          =   1455
         Left            =   120
         TabIndex        =   3
         Top             =   2400
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   2566
         Language        =   1
         KeywordColor    =   0
         OperatorColor   =   0
         DelimiterColor  =   0
         ForeColor       =   0
         FunctionColor   =   0
         HighlightCode   =   0
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
      Begin VB.Label Label1 
         Caption         =   "Text Size"
         Height          =   255
         Index           =   5
         Left            =   4080
         TabIndex        =   19
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Operator Color:"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   8
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Function Color:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Keyword Color:"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Fore Color:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Delimeter Color:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#
'# I was too lazy to work on this, but I probably will .....
'#

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdDelim_Click()
 cdc.ShowColor
 Delim.BackColor = cdc.Color
 TestArea.DelimiterColor = cdc.Color
 TestArea.Text = TestArea.Text & vbCrLf
End Sub

Private Sub cmdFore_Click()
 cdc.ShowColor
 Fore.BackColor = cdc.Color
 TestArea.ForeColor = cdc.Color
 TestArea.Text = TestArea.Text & vbCrLf
End Sub

Private Sub cmdFunction_Click()
 cdc.ShowColor
 fun.BackColor = cdc.Color
 TestArea.FunctionColor = cdc.Color
 TestArea.Text = TestArea.Text & vbCrLf
End Sub

Private Sub cmdKey_Click()
 cdc.ShowColor
 Key.BackColor = cdc.Color
 TestArea.KeywordColor = cdc.Color
 TestArea.Text = TestArea.Text & vbCrLf
End Sub

Private Sub cmdOK_Click()
 Unload Me
End Sub

Private Sub cmdOper_Click()
 cdc.ShowColor
 oper.BackColor = cdc.Color
 TestArea.OperatorColor = cdc.Color
 TestArea.Text = TestArea.Text & vbCrLf
End Sub

Private Sub Form_Load()
TestArea.Text = TestArea.Text & "<!--- This is a test Area --->" & vbCrLf & "<html>" & vbCrLf & "<body bgcolor=" & Chr(34) & "White" & Chr(34) & ">" & vbCrLf & vbCrLf & "</html>"
For I = 8 To 24 Step 1
 CoSize.AddItem I
Next I
CoSize.ListIndex = 1
End Sub
