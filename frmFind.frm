VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Find in Text ..."
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton FindButton 
      Caption         =   "Find"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Left            =   3735
      TabIndex        =   2
      Top             =   105
      Width           =   1230
   End
   Begin VB.CommandButton FindNextButton 
      Caption         =   "Find Again"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Left            =   5070
      TabIndex        =   3
      Top             =   105
      Width           =   1230
   End
   Begin VB.CommandButton ReplaceButton 
      Caption         =   "Replace"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Left            =   3735
      TabIndex        =   4
      Top             =   570
      Width           =   1230
   End
   Begin VB.CommandButton ReplaceAllButton 
      Caption         =   "Replace All"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Left            =   5055
      TabIndex        =   5
      Top             =   570
      Width           =   1230
   End
   Begin VB.CommandButton Command5 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Left            =   5055
      TabIndex        =   8
      Top             =   1095
      Width           =   1230
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Case sensitive"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   150
      TabIndex        =   6
      Top             =   1155
      Value           =   1  'Checked
      Width           =   2040
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Whole word only"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2175
      TabIndex        =   7
      Top             =   1155
      Width           =   2040
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1440
      TabIndex        =   0
      Top             =   105
      Width           =   2115
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1440
      TabIndex        =   1
      Top             =   570
      Width           =   2100
   End
   Begin VB.Label Label1 
      Caption         =   "Find what"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   165
      TabIndex        =   10
      Top             =   120
      Width           =   1410
   End
   Begin VB.Label Label2 
      Caption         =   "Replace with"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   165
      TabIndex        =   9
      Top             =   600
      Width           =   1485
   End
End
Attribute VB_Name = "frmFind"
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
Dim Position As Integer

Private Sub FindButton_Click()
On Error Resume Next
Dim FindFlags As Integer

    Position = 0
    FindFlags = Check1.Value * 4 + Check2.Value * 2
    Position = fMainForm.ActiveForm.rtfText.Find(Text1.Text, Position + 1, , FindFlags)
    'fMainForm.ActiveForm.rtfText.SelLength = 5
    fMainForm.ActiveForm.rtfText.SelLength = Len(Trim(Text1.Text))
    If Position >= 0 Then
        ReplaceButton.Enabled = True
        ReplaceAllButton.Enabled = True
        
    Else
        MsgBox "String not found", vbOKOnly, "Search Help"
        ReplaceButton.Enabled = False
        ReplaceAllButton.Enabled = False
    End If
    
End Sub

Private Sub FindNextButton_Click()
On Error Resume Next
Dim FindFlags

FindFlags = Check1.Value * 4 + Check2.Value * 2
Position = fMainForm.ActiveForm.rtfText.Find(Text1.Text, Position + 1, , FindFlags)
fMainForm.ActiveForm.rtfText.SelLength = Len(Trim(Text1.Text))
If Position > 0 Then
    
Else
    MsgBox "String not found", vbOKOnly, "Search Help"
    ReplaceButton.Enabled = False
    ReplaceAllButton.Enabled = False
End If

End Sub

Private Sub Command5_Click()
 Unload Me
End Sub

Private Sub Form_Activate()
On Error Resume Next
Text1.SetFocus
End Sub

Private Sub Form_GotFocus()
On Error Resume Next
Text1.SetFocus
End Sub
Private Sub ReplaceButton_Click()
On Error Resume Next
Dim FindFlags As Integer

    fMainForm.ActiveForm.rtfText.SelText = Text2.Text
    FindFlags = Check1.Value * 4 + Check2.Value * 2
    Position = fMainForm.ActiveForm.rtfText.Find(Text1.Text, Position + 1, , FindFlags)
    If Position > 0 Then
        
    Else
        'MsgBox "String not found", vbOKOnly, "Search Help"
        ReplaceButton.Enabled = False
        ReplaceAllButton.Enabled = False
    End If
    
End Sub

Private Sub ReplaceAllButton_Click()
On Error Resume Next
Dim FindFlags As Integer, I As Integer
I = 0
    FindFlags = Check1.Value * 4 + Check2.Value * 2
    fMainForm.ActiveForm.rtfText.SelText = Text2.Text
    Position = fMainForm.ActiveForm.rtfText.Find(Text1.Text, Position + 1, , FindFlags)
    While Position > 0
        I = I + 1
        fMainForm.ActiveForm.rtfText.SelText = Text2.Text
        Position = fMainForm.ActiveForm.rtfText.Find(Text1.Text, Position + 1, , FindFlags)
    Wend
    If I = 0 Then I = 1
        ReplaceButton.Enabled = False
        ReplaceAllButton.Enabled = False
        MsgBox I & " item(s) replaced", vbOKOnly, "Search Help"
End Sub


