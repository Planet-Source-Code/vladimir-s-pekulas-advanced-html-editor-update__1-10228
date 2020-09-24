VERSION 5.00
Begin VB.Form frmSnipp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "HTML Snippets Library ..."
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   165
      TabIndex        =   9
      Top             =   120
      Width           =   5415
      Begin VB.DriveListBox drvDrive 
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   2055
      End
      Begin VB.DirListBox DirDirectory 
         Height          =   1890
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   2055
      End
      Begin VB.FileListBox filFileName 
         Height          =   2235
         Left            =   2280
         TabIndex        =   10
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1815
      Left            =   150
      TabIndex        =   0
      Top             =   2940
      Width           =   5415
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   2160
         Picture         =   "frmSnipp.frx":0000
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   13
         Top             =   240
         Width           =   480
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add Script"
         Height          =   375
         Left            =   3840
         TabIndex        =   5
         Top             =   650
         Width           =   1455
      End
      Begin VB.TextBox txtArtist 
         Height          =   285
         Left            =   840
         TabIndex        =   4
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox txtTitle 
         Height          =   285
         Left            =   840
         TabIndex        =   3
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox txtCDID 
         Height          =   285
         Left            =   840
         TabIndex        =   2
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "OK"
         Height          =   375
         Left            =   3840
         TabIndex        =   1
         Top             =   1110
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "File:"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   8
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Title:"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   7
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Snipp ID:"
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   6
         Top             =   480
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmSnipp"
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

Option Explicit
Dim NumRecords As Integer
Dim intFileNum As Integer
Dim lngRecLength As Long
Private Type ViewSnipps
    intID As Integer
    strTitle As String * 99
    strArtist As String * 100
End Type
Private Type AddCD
    intID As Integer
    strTitle As String * 99
    strArtist As String * 100
End Type

Private Sub cmdDone_Click()
 Unload Me
End Sub

Private Sub Command1_Click()
 'Refresh List
 Dim IRef As Integer
 'Delete it first !
 For IRef = 1 To fMainForm.SnippList.ListItems.Count
  fMainForm.SnippList.ListItems.Remove (1)
 Next IRef
 ' Load it again ! (Our Own)
 fMainForm.SnippList.ListItems.Add , , "Bohemia Gift Finder", , 14
 fMainForm.SnippList.ListItems.Add , , "GoTo.com Search Engine", , 14
 fMainForm.SnippList.ListItems.Add , , "InfoSeek.com Search Engine", , 14
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
 Unload Me
     Exit Sub
     Close #intCDFile
     End If
 lngCDID = 0
 Do
 lngCDID = lngCDID + 1
     If lngCDID > lngTotalRecords Then
 Unload Me
 Exit Sub
     Close #intCDFile
     Else
  If lngCDID > 0 And lngCDID <= lngTotalRecords Then
 Get #intCDFile, lngCDID, udtCDToView
 fMainForm.SnippList.ListItems.Add , , udtCDToView.strTitle, , 14
  End If
     End If
 Loop
 Close #intCDFile
 Unload Me
End Sub

Private Sub DirDirectory_Change()
 filFileName.Path = DirDirectory.Path
End Sub

Private Sub drvDrive_Change()
 DirDirectory.Path = drvDrive.Drive
End Sub

Private Sub filFileName_Click()
 Dim intFileNum As Integer
 Dim strTextLine As String, strFilename As String
 If Right(DirDirectory.Path, 1) = "\" Then
  strFilename = filFileName.Path & filFileName.Filename
 Else
  strFilename = filFileName.Path & "\" & filFileName.Filename
 End If
 txtArtist.Text = strFilename
End Sub

Private Sub Form_Load()
 Dim udtCD As AddCD
 Dim intCDFile As Integer, lngRecLength As Long, lngNextCDID As Long
 Dim NumRecords As Integer
 Dim intFileNum As Integer
 intFileNum = FreeFile
 'Open File
 intCDFile = FreeFile
 lngRecLength = LenB(udtCD)
 Open App.Path & "\SnippetIndex.dat" For Random As #intCDFile Len = lngRecLength
 'Next rec (NUMRECORDS)
 If LOF(intFileNum) Mod lngRecLength = 0 Then
  NumRecords = (LOF(intFileNum) \ lngRecLength)
 Else
  NumRecords = (LOF(intFileNum) \ lngRecLength) + 1
 End If
 lngNextCDID = NumRecords + 1
 txtCDID.Text = lngNextCDID
 txtCDID.Enabled = False
 Close #intCDFile
End Sub


Private Sub cmdAdd_Click()
 Dim udtNewCD As AddCD
 Dim intCDFile As Integer, lngRecLength As Long, lngCDID As Long
 'Check if not Title ""
 If txtTitle.Text = "" Then
 MsgBox ("Please Name the Snippet")
 Exit Sub
 End If
 ' check if not path to file ""
 If txtArtist.Text = "" Then
 MsgBox ("Please Select File to Use as a Snippet.")
 Exit Sub
 End If

 'Open File
 intCDFile = FreeFile
 lngRecLength = LenB(udtNewCD)
 Open App.Path & "\SnippetIndex.dat" For Random As #intCDFile Len = lngRecLength

 'Adds New CD
 lngCDID = txtCDID.Text
 udtNewCD.strArtist = txtArtist.Text
 udtNewCD.strTitle = txtTitle.Text
 
 Put #intCDFile, lngCDID, udtNewCD
 'Make txt ""
 txtCDID.Text = lngCDID + 1
 txtTitle.Text = ""
 txtArtist.Text = ""
 Close #intCDFile
End Sub
