VERSION 5.00
Begin VB.Form frmJava 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Java Snippets Library ..."
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   1815
      Left            =   180
      TabIndex        =   4
      Top             =   2940
      Width           =   5415
      Begin VB.CommandButton Command1 
         Caption         =   "OK"
         Height          =   375
         Left            =   3840
         TabIndex        =   10
         Top             =   1110
         Width           =   1455
      End
      Begin VB.TextBox txtCDID 
         Height          =   285
         Left            =   840
         TabIndex        =   9
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtTitle 
         Height          =   285
         Left            =   840
         TabIndex        =   8
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox txtArtist 
         Height          =   285
         Left            =   840
         TabIndex        =   7
         Top             =   1200
         Width           =   2415
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add Script"
         Height          =   375
         Left            =   3840
         TabIndex        =   6
         Top             =   650
         Width           =   1455
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   2160
         Picture         =   "frmJava.frx":0000
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   5
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "Java ID:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Title:"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   12
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "File:"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   11
         Top             =   1200
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   195
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin VB.FileListBox filFileName 
         Height          =   2235
         Left            =   2280
         TabIndex        =   3
         Top             =   240
         Width           =   2895
      End
      Begin VB.DirListBox DirDirectory 
         Height          =   1890
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   2055
      End
      Begin VB.DriveListBox drvDrive 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmJava"
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
For IRef = 1 To fMainForm.JavaList.ListItems.Count
fMainForm.JavaList.ListItems.Remove (1)
Next IRef
' Load it again ! (Our Own)
 fMainForm.JavaList.ListItems.Add , , "Counter Script", , 14
 fMainForm.JavaList.ListItems.Add , , "Gallery Script", , 14
 fMainForm.JavaList.ListItems.Add , , "IP Address Script", , 14
 fMainForm.JavaList.ListItems.Add , , "Text Effect Script", , 14
 fMainForm.JavaList.ListItems.Add , , "Redirection Script", , 14
 fMainForm.JavaList.ListItems.Add , , "Resolution Script", , 14
 fMainForm.JavaList.ListItems.Add , , "Scroller Script", , 14
 fMainForm.JavaList.ListItems.Add , , "Status Bar Script", , 14
 fMainForm.JavaList.ListItems.Add , , "Email Form Script", , 14
 fMainForm.JavaList.ListItems.Add , , "News Ticker Script", , 14

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
Open app_path & "\JavaIndex.dat" For Random As #intJavaFile Len = lngRecLengthJava


If LOF(intFileNumJava) Mod lngRecLengthJava = 0 Then
NumRecordsJava = (LOF(intJavaFile) \ lngRecLengthJava)
Else
NumRecordsJava = (LOF(intJavaFile) \ lngRecLengthJava) + 1
End If
lngTotalRecordsJava = NumRecordsJava

'View Rec if Valid
If lngTotalRecordsJava = 0 Then
Unload Me
Exit Sub
End If
lngJavaID = 0
Do
    If lngJavaID = lngTotalRecordsJava Then
    Close #intJavaFile
    Unload Me
    Exit Sub
    Else
lngJavaID = lngJavaID + 1
 If lngJavaID > 0 And lngJavaID <= lngTotalRecordsJava Then
Get #intJavaFile, lngJavaID, udtJavaToView
fMainForm.JavaList.ListItems.Add , , udtJavaToView.strTitle, , 14
 End If
    End If
Loop
Close #intJavaFile
'FMainForm.























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
Open app_path & "\JavaIndex.dat" For Random As #intCDFile Len = lngRecLength
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
Open app_path & "\JavaIndex.dat" For Random As #intCDFile Len = lngRecLength

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

'#################################
'locates current working directory
'#################################
Function app_path() As String
Dim x As String
    x = App.Path
    If Right$(x, 1) <> "\" Then x = x + "\"
    app_path = UCase$(x)
'app_path is with "\" on the end

End Function


