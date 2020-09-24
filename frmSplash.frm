VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5775
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4515
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   4515
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox picLogo 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4170
      Left            =   0
      Picture         =   "frmSplash.frx":0000
      ScaleHeight     =   4170
      ScaleWidth      =   4500
      TabIndex        =   0
      Top             =   0
      Width           =   4500
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Programmed by: Vladimir S. Pekulas Jr.  July/2000"
      Height          =   255
      Left            =   390
      TabIndex        =   2
      Top             =   5400
      Width           =   3735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmSplash.frx":CBAD
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   4320
      Width           =   4215
   End
End
Attribute VB_Name = "frmSplash"
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

