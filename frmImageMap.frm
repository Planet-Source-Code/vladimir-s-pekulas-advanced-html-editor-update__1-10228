VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6D940288-9F11-11CE-83FD-02608C3EC08A}#2.1#0"; "IMGEDIT.OCX"
Begin VB.Form frmImageMap 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Image Map Tool ..."
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   8445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command8 
      BackColor       =   &H000000FF&
      Height          =   375
      Left            =   1185
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Change Color of Lines"
      Top             =   5055
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   1785
      TabIndex        =   12
      Top             =   4950
      Width           =   6495
      Begin VB.CommandButton Command3 
         Caption         =   "&Browse"
         Height          =   375
         Left            =   4920
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtHREF 
         Height          =   285
         Left            =   960
         TabIndex        =   15
         Text            =   "http://www.example.com"
         Top             =   240
         Width           =   3855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Add"
         Height          =   375
         Left            =   7080
         TabIndex        =   14
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Add"
         Height          =   375
         Left            =   5880
         TabIndex        =   13
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Link To:"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   825
      TabIndex        =   10
      ToolTipText     =   "Delete Selected shape"
      Top             =   5055
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7890
      TabIndex        =   8
      Top             =   765
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7530
      TabIndex        =   7
      Top             =   765
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Locate Image"
      Height          =   375
      Left            =   75
      TabIndex        =   6
      Top             =   135
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   315
      Left            =   1380
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "a"
      Top             =   165
      Width           =   3375
   End
   Begin VB.TextBox Text5 
      Height          =   315
      Left            =   6225
      TabIndex        =   4
      Text            =   "ImgMap"
      Top             =   165
      Width           =   2055
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   105
      TabIndex        =   3
      Top             =   5910
      Width           =   6855
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   4860
      TabIndex        =   2
      Top             =   7860
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&OK"
      Height          =   375
      Left            =   7065
      TabIndex        =   1
      Top             =   5910
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   7065
      TabIndex        =   0
      Top             =   6390
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   195
      Top             =   7065
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ImgeditLibCtl.ImgAnnTool ImgAnnTool2 
      Height          =   375
      Left            =   465
      TabIndex        =   9
      ToolTipText     =   "Draw a polygonal shape"
      Top             =   5055
      Width           =   375
      _Version        =   131073
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   64
      BeginProperty AnnotationFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DestImageControl=   "@ö"
      BeginProperty AnnotationFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ImgeditLibCtl.ImgAnnTool ImgAnnTool1 
      Height          =   375
      Left            =   105
      TabIndex        =   11
      ToolTipText     =   "Draw a rectangular shape"
      Top             =   5055
      Width           =   375
      _Version        =   131073
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   64
      BeginProperty AnnotationFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AnnotationType  =   3
      DestImageControl=   "ImgEdit1"
      BeginProperty AnnotationFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ImgeditLibCtl.ImgEdit ImgEdit1 
      Height          =   4095
      Left            =   105
      TabIndex        =   18
      Top             =   750
      Width           =   8175
      _Version        =   131073
      _ExtentX        =   14420
      _ExtentY        =   7223
      _StockProps     =   96
      BorderStyle     =   1
      ImageControl    =   "ImgEdit1"
      SelectionRectangleEnabled=   0   'False
      BeginProperty AnnotationFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ImagePalette    =   3
      UndoBufferSize  =   134405120
      OcrZoneVisibility=   -4044
      AnnotationOcrType=   127
   End
   Begin VB.Label Label2 
      Caption         =   "Map Name:"
      Height          =   225
      Left            =   5025
      TabIndex        =   20
      Top             =   210
      Width           =   1095
   End
End
Attribute VB_Name = "frmImageMap"
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

'# Image Map tool originaly created by Jamiie someone. (sorry I don't remember full name)
Dim sQoute As String
Dim MapX As Integer
Dim MapY As Integer
Dim MapXStart As Integer
Dim MapYStart As Integer
Dim strPoly As String
Dim I As Integer

Private Sub Combo1_Click()
 ImgEdit1.Zoom = Combo1.Text
 ImgEdit1.Display
End Sub

Private Sub Command1_Click()
 ImgEdit1.DeleteSelectedAnnotations
 Text3.Text = ""
End Sub

Private Sub Command2_Click()
On Error Resume Next
If Text3.Text = "" Then
 MsgBox "No new shapes to map"
Else
 ImgEdit1.DeleteSelectedAnnotations
 List1.AddItem Text3.Text & " HREF =" & sQoute & txtHREF.Text & sQoute & " >"
 Text3.Text = ""
 strPoly = ""
End If
End Sub

Private Sub Command3_Click()
On Error Resume Next
 CommonDialog1.ShowOpen
 txtHREF = CommonDialog1.Filename
End Sub

Private Sub Command4_Click()
On Error GoTo Err
 CommonDialog1.Filter = "JPEG images (*.jpg)|*.jpg;*.jpe;*.jpeg|GIF Images (*.gif)|*.gif|Bitmaps (*.bmp)|*.bmp|"
 CommonDialog1.ShowOpen
 Text4.Text = CommonDialog1.Filename
 ImgEdit1.Image = (CommonDialog1.Filename)
 ImgEdit1.ClearDisplay
 ImgEdit1.Display
 ImgEdit1.Enabled = True
 Exit Sub
Err:
 If Err.Number = 1002 Then
  Exit Sub
 End If
End Sub

Private Sub Command6_Click()
On Error Resume Next
Dim strDone As String
For I = 0 To List1.ListCount - 1
 strDone = strDone & vbCrLf & List1.List(I)
Next I
 fMainForm.ActiveForm.rtfText.SelText = "<map name =" & sQoute & Text5.Text & sQoute & ">" & vbCrLf _
 & strDone & vbCrLf & "</map>" & vbCrLf & "<img src =" & sQoute _
 & ImgEdit1.Image & sQoute & " border = 0 usemap =#" & Text5.Text _
 & ">"
 Unload Me
End Sub

Private Sub Command7_Click()
 Unload Me
End Sub

Private Sub Command8_Click()
On Error Resume Next
 CommonDialog1.ShowColor
 ImgEdit1.AnnotationLineColor = CommonDialog1.Color
 ImgEdit1.AnnotationFillColor = CommonDialog1.Color
 Command8.BackColor = CommonDialog1.Color
End Sub

Private Sub Form_Load()
 I = 0
 sQoute = """"
End Sub

Private Sub ImgAnnTool1_Click()
 ImgEdit1.AnnotationType = wiHollowRect
End Sub

Private Sub ImgAnnTool2_Click()
 ImgEdit1.AnnotationType = wiStraightLine
End Sub

Private Sub ImgEdit1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 45 Then
  ImgEdit1.Zoom = ImgEdit1.Zoom - 20
  ImgEdit1.Display
ElseIf KeyCode = 61 Then
  ImgEdit1.Zoom = ImgEdit1.Zoom + 20
  ImgEdit1.Display
End If
End Sub

Private Sub ImgEdit1_MarkSelect(ByVal Button As Integer, ByVal Shift As Integer, ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, ByVal MarkType As Integer, ByVal GroupName As String)
On Error Resume Next
If ImgEdit1.AnnotationType = wiHollowRect Then
 Text3.Text = "<area shape = " & sQoute & "rect" & sQoute & " COORDS =" & sQoute & Left + ImgEdit1.ScrollPositionX & "," & Top + ImgEdit1.ScrollPositionY & "," & Width + Left + ImgEdit1.ScrollPositionX & "," & Height + Top + ImgEdit1.ScrollPositionY & sQoute
Else
 strPoly = strPoly & MapX & "," & MapY & ","
 Text3.Text = "<area shape = " & sQoute & "poly" & sQoute & " COORDS =" & sQoute & strPoly & sQoute
End If
End Sub

Private Sub ImgEdit1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
 MapX = ScaleX(x, vbTwips, vbPixels) + ImgEdit1.ScrollPositionX
 MapY = ScaleY(y, vbTwips, vbPixels) + ImgEdit1.ScrollPositionY
 Text1.Text = MapX
 Text2.Text = MapY
End Sub

Private Sub ImgEdit1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 On Error Resume Next
 If ImgEdit1.ImageDisplayed = True Then
  MapX = ScaleX(x, vbTwips, vbPixels) + ImgEdit1.ScrollPositionX
  MapY = ScaleY(y, vbTwips, vbPixels) + ImgEdit1.ScrollPositionY
  Text1.Text = MapX
  Text2.Text = MapY
 End If
End Sub
