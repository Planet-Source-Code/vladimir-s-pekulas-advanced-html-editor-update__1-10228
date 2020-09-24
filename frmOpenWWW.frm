VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmOpenWWW 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Open From The Web ...."
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   60
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   195
      Picture         =   "frmOpenWWW.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   45
      Width           =   480
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   75
      TabIndex        =   0
      Top             =   165
      Width           =   7095
      Begin VB.TextBox txtURL 
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   465
         Width           =   5415
      End
      Begin VB.Label Label1 
         Caption         =   "URL:"
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   480
         Width           =   495
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3780
      TabIndex        =   2
      Top             =   1365
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5610
      TabIndex        =   4
      Top             =   1365
      Width           =   1575
   End
End
Attribute VB_Name = "frmOpenWWW"
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

'# Gets a file from entred URL
Private Sub cmdCancel_Click()
 Unload frmOpenWWW
End Sub

Private Sub Command1_Click()
On Error Resume Next
Dim FILE As String, intFileNum As Integer, strTextLine As String, strFileName As String, b() As Byte
    
    FILE = FreeFile
    If txtURL.Text = "" Then
    MsgBox " Please Enter URL ! ", vbInformation
    Exit Sub
    End If

    Inet1.Protocol = icHTTP
    Inet1.URL = txtURL.Text
    b() = Inet1.OpenURL(Inet1.URL, icByteArray)
    Open "c:\Casper~www~open.html" For Binary Access Write As #FILE
    Put #FILE, , b()
    Close #FILE

    Unload frmOpenWWW
        Call fMainForm.LoadNewDocFunction
        fMainForm.ActiveForm.rtfText.Text = ""

 intFileNum = FreeFile
 Open "c:\Casper~www~open.html" For Input As #intFileNum
  Do While Not EOF(intFileNum)
   Line Input #intFileNum, strTextLine
   fMainForm.ActiveForm.rtfText.Text = fMainForm.ActiveForm.rtfText.Text & strTextLine & vbCrLf
  Loop
Close #intFileNum
End Sub

