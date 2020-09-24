VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6480
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4575
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   4575
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer T 
      Interval        =   50
      Left            =   720
      Top             =   7560
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&OK"
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6000
      Width           =   1455
   End
   Begin VB.PictureBox P 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   120
      ScaleHeight     =   1545
      ScaleWidth      =   4305
      TabIndex        =   2
      Top             =   4320
      Width           =   4335
      Begin VB.Label L2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Programmed by: Vladimir S. Pekulas Jr.  July/2000"
         Height          =   255
         Left            =   270
         TabIndex        =   4
         Top             =   2640
         Width           =   3735
      End
      Begin VB.Label L1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   $"frmAbout.frx":0000
         Height          =   735
         Left            =   0
         TabIndex        =   3
         Top             =   1560
         Width           =   4215
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4170
      Left            =   0
      Picture         =   "frmAbout.frx":008C
      ScaleHeight     =   4170
      ScaleWidth      =   4500
      TabIndex        =   0
      Top             =   0
      Width           =   4500
   End
End
Attribute VB_Name = "frmAbout"
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

Private Sub Command1_Click()
Unload Me
End Sub

'*************************************************************
' Project:  IntellIm                                         *
' Filename: ListSearch.ctl                                   *
' Author:   Edward P. Denninger III                          *
' Date:     5/14/2000                                        *
' Copyright © 2000 Edward P. Denninger III                   *
'*************************************************************

'*************************************************************
' Project:  Coloring (original version)                      *
' Filename: n/a                                              *
' Author:   http://www.developersdomain.com                  *
' Date:     n/a                                              *
' Copyright © 2000 http://www.developersdomain.com           *
'*************************************************************


'A little history in bad English:
' Well, I first started with this app about 7 month ago, it
' was pretty much my 1st big app and it looked that way.
' For last couple a weeks I've recoded the whole thing from
' scrach and now it mostly works as I'd like, however it is
' still far far away from being perfect.
' The next version should include "Perfect" error handeling
' and working version of IntellIm (my mistake not author's)
'
'Why Casper HTML Editor:
' Casper Semiramis II was my dog, a boxer, so I named this
' app after him.
'
'Help Wanted:
' If anybody is interested to work on this with me, let me know.



' July 29. 2000
' Vladimir S. Pekulas
Private Sub Form_DblClick()
Unload Me
End Sub

Private Sub Label2_Click()

End Sub

Private Sub Picture1_DblClick()
Unload Me
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call FormDrag(Me)
End Sub

Private Sub T_Timer()
If L1.Top <= -800 Then L1.Top = 1560
If L2.Top <= -240 Then L2.Top = 2640

L1.Top = L1.Top - 15
L2.Top = L2.Top - 15

End Sub
