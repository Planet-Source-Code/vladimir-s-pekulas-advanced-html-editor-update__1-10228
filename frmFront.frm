VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmFront 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Start Up ..."
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox ChExit 
      Caption         =   "Don't show this dialog in the future."
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   5040
      Width           =   2895
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   8493
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Teplates"
      TabPicture(0)   =   "frmFront.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Templates"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Command6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Picture1(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdSecurity"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdBlank"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdFeedBack"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdGlossary"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdContents"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdContact"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdFAQ"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "CmdSite"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Picture3(0)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "Existing"
      TabPicture(1)   =   "frmFront.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(1)=   "Label2"
      Tab(1).Control(2)=   "drvDrive"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "DirDirectory"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "filFileName"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtFileName"
      Tab(1).Control(6)=   "coFileType"
      Tab(1).Control(7)=   "Command1"
      Tab(1).Control(8)=   "Command2"
      Tab(1).Control(9)=   "Picture1(1)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Recent"
      TabPicture(2)   =   "frmFront.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "RecentFiles"
      Tab(2).Control(1)=   "img"
      Tab(2).Control(2)=   "Picture3(1)"
      Tab(2).Control(3)=   "Picture1(2)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Command4"
      Tab(2).Control(5)=   "Command3"
      Tab(2).Control(6)=   "Label3"
      Tab(2).ControlCount=   7
      Begin MSComctlLib.TreeView RecentFiles 
         Height          =   3255
         Left            =   -73440
         TabIndex        =   30
         Top             =   600
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   5741
         _Version        =   393217
         Style           =   5
         ImageList       =   "img"
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin MSComctlLib.ImageList img 
         Left            =   -74400
         Top             =   3000
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFront.frx":0054
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.PictureBox Picture3 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   1
         Left            =   -74760
         Picture         =   "frmFront.frx":04A8
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   29
         Top             =   4080
         Width           =   480
      End
      Begin VB.PictureBox Picture3 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   0
         Left            =   240
         Picture         =   "frmFront.frx":05FA
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   28
         Top             =   4080
         Width           =   480
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         ClipControls    =   0   'False
         Height          =   1260
         Index           =   2
         Left            =   -74760
         Picture         =   "frmFront.frx":074C
         ScaleHeight     =   1200
         ScaleMode       =   0  'User
         ScaleWidth      =   870
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "Casper Productions"
         Top             =   600
         Width           =   930
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         ClipControls    =   0   'False
         Height          =   1260
         Index           =   1
         Left            =   -74760
         Picture         =   "frmFront.frx":11B2
         ScaleHeight     =   1200
         ScaleMode       =   0  'User
         ScaleWidth      =   870
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "Casper Productions"
         Top             =   600
         Width           =   930
      End
      Begin VB.CommandButton CmdSite 
         BackColor       =   &H80000009&
         Caption         =   "SiteMap"
         Height          =   855
         Left            =   5400
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmFront.frx":1C18
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Site Map Template"
         Top             =   2400
         Width           =   975
      End
      Begin VB.CommandButton cmdFAQ 
         BackColor       =   &H80000009&
         Caption         =   "F. A. Q."
         Height          =   855
         Left            =   4200
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmFront.frx":205A
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Frequent Asked Questions Template"
         Top             =   2400
         Width           =   975
      End
      Begin VB.CommandButton cmdContact 
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
         Caption         =   "Contact"
         Height          =   855
         Left            =   3000
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmFront.frx":279C
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Contact Template"
         Top             =   2400
         Width           =   975
      End
      Begin VB.CommandButton cmdContents 
         BackColor       =   &H80000009&
         Caption         =   "Contents"
         Height          =   855
         Left            =   5400
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmFront.frx":2BDE
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Table of Contents Template"
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton cmdGlossary 
         BackColor       =   &H80000009&
         Caption         =   "Glossary"
         Height          =   855
         Left            =   4200
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmFront.frx":3020
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Glossary Template"
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton cmdFeedBack 
         BackColor       =   &H80000009&
         Caption         =   "Feed Back"
         Height          =   855
         Left            =   3000
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmFront.frx":3462
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Feed Back Template"
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton cmdBlank 
         BackColor       =   &H80000009&
         Caption         =   "Blank Page"
         Height          =   855
         Left            =   1800
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmFront.frx":3BA4
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Blank Page"
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton cmdSecurity 
         BackColor       =   &H80000009&
         Caption         =   "Security"
         Height          =   855
         Left            =   1800
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmFront.frx":3FE6
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Security Statment Template"
         Top             =   2400
         Width           =   975
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         ClipControls    =   0   'False
         Height          =   1260
         Index           =   0
         Left            =   240
         Picture         =   "frmFront.frx":4428
         ScaleHeight     =   1200
         ScaleMode       =   0  'User
         ScaleWidth      =   870
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Casper Productions"
         Top             =   600
         Width           =   930
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   5280
         TabIndex        =   14
         Top             =   4200
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   -71520
         TabIndex        =   12
         Top             =   4200
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Open"
         Height          =   375
         Left            =   -69720
         TabIndex        =   11
         Top             =   4200
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   -69960
         TabIndex        =   10
         Top             =   4200
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Open"
         Height          =   375
         Left            =   -69960
         TabIndex        =   9
         Top             =   3660
         Width           =   1575
      End
      Begin VB.ComboBox coFileType 
         Height          =   315
         Left            =   -73560
         TabIndex        =   8
         Text            =   "All Files   (*.*)"
         Top             =   4230
         Width           =   3375
      End
      Begin VB.TextBox txtFileName 
         Height          =   285
         Left            =   -73560
         TabIndex        =   6
         Top             =   3705
         Width           =   3375
      End
      Begin VB.FileListBox filFileName 
         Height          =   2820
         Left            =   -71160
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   600
         Width           =   3015
      End
      Begin VB.DirListBox DirDirectory 
         Height          =   2340
         Left            =   -73560
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1080
         Width           =   2295
      End
      Begin VB.DriveListBox drvDrive 
         Height          =   315
         Left            =   -73560
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   600
         Width           =   2295
      End
      Begin MSComctlLib.ListView Templates 
         Height          =   3135
         Left            =   1440
         TabIndex        =   1
         Top             =   600
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   5530
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label4 
         Caption         =   "Slect one of the above listed templates and click the icon."
         Height          =   495
         Left            =   840
         TabIndex        =   15
         Top             =   4200
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "Double Click recent file or select and click ""Open""."
         Height          =   495
         Left            =   -74160
         TabIndex        =   13
         Top             =   4200
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Files of &type:"
         Height          =   255
         Left            =   -74640
         TabIndex        =   7
         Top             =   4260
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "File &Name:"
         Height          =   255
         Left            =   -74640
         TabIndex        =   5
         Top             =   3720
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmFront"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ChExit_Click()
On Error Resume Next
Dim intFileNum As Integer
Dim strTextLine As String, strFilename As String

 intFileNum = FreeFile
If ChExit.Value = 1 Then
 sDat = "1"
 Open App.Path & "\exitfront.txt" For Output As #intFileNum
 Print #intFileNum, sDat
 Close #intFileNum
Else
 sDat = "0"
 Open App.Path & "\exitfront.txt" For Output As #intFileNum
 Print #intFileNum, sDat
 Close #intFileNum
End If
End Sub

Private Sub cmdBlank_Click()
  Unload Me
End Sub

Private Sub cmdContact_Click()
On Error Resume Next
  fMainForm.ActiveForm.rtfText.LoadFile App.Path & "\Templates\contact.html"
 Unload Me
End Sub

Private Sub cmdContents_Click()
On Error Resume Next
 fMainForm.ActiveForm.rtfText.LoadFile App.Path & "\Templates\table.html"
 Unload Me
End Sub

Private Sub cmdFAQ_Click()
On Error Resume Next
 fMainForm.ActiveForm.rtfText.LoadFile App.Path & "\Templates\table.html"
 Unload Me
End Sub

Private Sub cmdFeedBack_Click()
On Error Resume Next
 fMainForm.ActiveForm.rtfText.LoadFile App.Path & "\Templates\feedback.html"
 Unload Me
End Sub

Private Sub cmdGlossary_Click()
On Error Resume Next
 fMainForm.ActiveForm.rtfText.LoadFile App.Path & "\Templates\glossary.html"
 Unload Me
End Sub

Private Sub cmdSecurity_Click()
 On Error Resume Next
 fMainForm.ActiveForm.rtfText.LoadFile App.Path & "\Templates\security.html"
 Unload Me
End Sub

Private Sub CmdSite_Click()
On Error Resume Next
 fMainForm.ActiveForm.rtfText.LoadFile App.Path & "\Templates\sitemap.html"
 Unload Me
End Sub

Private Sub Command1_Click()
On Error Resume Next
 If txtFileName.Text = "" Then
  MsgBox "Please Select File to open"
  Exit Sub
 End If
fMainForm.ActiveForm.rtfText.LoadFile txtFileName.Text

' Don't forgot to use recent.dat
If LCase(filFileName.Path) = "c:\" Then
 If LCase(filFileName.Path) = "c:\" Then
  sFile = filFileName.Path & filFileName.Filename
 Else
  sFile = filFileName.Path & "\" & filFileName.Filename
 End If
Else
  sFile = filFileName.Path & "\" & filFileName.Filename
End If
 Dim intFileNum As Integer
 Dim strTextLine As String, strFilename As String
 intFileNum = FreeFile
 Open App.Path & "\recent.dat" For Append As #intFileNum
 Print #intFileNum, sFile
 Close #intFileNum
 Unload Me
End Sub

Private Sub Command2_Click()
On Error Resume Next
 Unload Me
End Sub

Private Sub Command3_Click()
On Error Resume Next
 fMainForm.ActiveForm.rtfText.LoadFile RecentFiles.SelectedItem.Text
 Unload Me
End Sub

Private Sub Command4_Click()
 Unload Me
End Sub
Private Sub Command6_Click()
 Unload Me
End Sub

Private Sub DirDirectory_Change()
On Error Resume Next
 filFileName.Path = DirDirectory.Path
End Sub
Private Sub drvDrive_Change()
On Error Resume Next
 DirDirectory.Path = drvDrive.Drive
End Sub

Private Sub filFileName_Click()
On Error Resume Next
Dim intFileNum As Integer
Dim strTextLine As String, strFilename As String
 If Right(DirDirectory.Path, 1) = "\" Then
  strFilename = filFileName.Path & filFileName.Filename
 Else
  strFilename = filFileName.Path & "\" & filFileName.Filename
 End If
  txtFileName.Text = ""
  txtFileName.Text = strFilename
End Sub

Private Sub filFileName_DblClick()
On Error Resume Next
 Call Command1_Click
End Sub

Private Sub Form_Load()
On Error Resume Next
'Make sure this form is suppoused to show up
Dim intFileNum As Integer
Dim strTextLine As String, strFilename As String
 
 intFileNum = FreeFile
 Open App.Path & "\exitfront.txt" For Input As #intFileNum
  Do While Not EOF(intFileNum)
   Line Input #intFileNum, strTextLine
  Loop
 Close #intFileNum
  If strTextLine = "1" Then
    Unload Me 'If the value is 1 then don't show up
    Exit Sub
  End If

' Let's load up the recent.dat file for recent files
intFileNum = FreeFile
Open App.Path & "\recent.dat" For Input As #intFileNum
Do While Not EOF(intFileNum)
Line Input #intFileNum, strTextLine
 RecentFiles.Nodes.Add , , , strTextLine, 1, 1
Loop
Close #intFileNum
End Sub

Private Sub RecentFiles_DblClick()
On Error Resume Next
 fMainForm.ActiveForm.rtfText3.LoadFile RecentFiles.SelectedItem.Text
 fMainForm.ActiveForm.rtfText.Text = fMainForm.ActiveForm.rtfText3.Text
 Unload Me
End Sub
