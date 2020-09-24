VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTags 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Insert Tag ..."
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   8280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   465
      Top             =   2895
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
            Picture         =   "frmTags.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   405
      Left            =   6675
      TabIndex        =   2
      Top             =   4080
      Width           =   1485
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Select Tag"
      Height          =   405
      Left            =   5070
      TabIndex        =   1
      Top             =   4080
      Width           =   1485
   End
   Begin MSComctlLib.ListView TagList 
      Height          =   3840
      Left            =   120
      TabIndex        =   0
      Top             =   105
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   6773
      View            =   2
      Arrange         =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Available HTML 4.0 Tags"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmTags"
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
On Error Resume Next
  If TagList.SelectedItem.Text = "<a>" Then
   Unload Me
   TabNumber = 0
   frmTagEdit.Show 1, fMainForm
   Exit Sub
  End If
 If TagList.SelectedItem.Text = "<br>" Then fMainForm.ActiveForm.rtfText.SelText = "<br>"
 If TagList.SelectedItem.Text = "<blink>" Then fMainForm.ActiveForm.rtfText.SelText = "<blink> </blink>"
 If TagList.SelectedItem.Text = "<center>" Then fMainForm.ActiveForm.rtfText.SelText = "<center> </center>"
 If TagList.SelectedItem.Text = "<dir>" Then fMainForm.ActiveForm.rtfText.SelText = "<dir> </dir>"
  If TagList.SelectedItem.Text = "<div>" Then
   Unload Me
   TabNumber = 3
   frmTagEdit.Show 1, fMainForm
   Exit Sub
  End If
 If TagList.SelectedItem.Text = "<head>" Then fMainForm.ActiveForm.rtfText.SelText = "<head> </head>"
 If TagList.SelectedItem.Text = "<html>" Then fMainForm.ActiveForm.rtfText.SelText = "<html> </html>"
 If TagList.SelectedItem.Text = "<menu>" Then fMainForm.ActiveForm.rtfText.SelText = "<menu> </menu>"
 If TagList.SelectedItem.Text = "<nobr>" Then fMainForm.ActiveForm.rtfText.SelText = "<nobr> </nobr>"
 If TagList.SelectedItem.Text = "<noembad>" Then fMainForm.ActiveForm.rtfText.SelText = "<noembad> </noembad>"
 If TagList.SelectedItem.Text = "<nolayer>" Then fMainForm.ActiveForm.rtfText.SelText = "<nolayer> </nolayer>"
 If TagList.SelectedItem.Text = "<plain text>" Then fMainForm.ActiveForm.rtfText.SelText = "<plain text> </plain text>"
 If TagList.SelectedItem.Text = "<strong>" Then fMainForm.ActiveForm.rtfText.SelText = "<strong> </strong>"
 If TagList.SelectedItem.Text = "<strike>" Then fMainForm.ActiveForm.rtfText.SelText = "<strike> </strike>"
 If TagList.SelectedItem.Text = "<title>" Then fMainForm.ActiveForm.rtfText.SelText = "<title> </title>"
 If TagList.SelectedItem.Text = "<u>" Then fMainForm.ActiveForm.rtfText.SelText = "<u> </u>"
 If TagList.SelectedItem.Text = "<ul>" Then fMainForm.ActiveForm.rtfText.SelText = "<ul> </ul>"
  If TagList.SelectedItem.Text = "<body>" Then
   Unload Me
   TabNumber = 1
   frmTagEdit.Show 1, fMainForm
   Exit Sub
  End If
  If TagList.SelectedItem.Text = "<font>" Then
   Unload Me
   TabNumber = 4
   frmTagEdit.Show 1, fMainForm
   Exit Sub
  End If
  If TagList.SelectedItem.Text = "<img>" Then
   Unload Me
   TabNumber = 6
   frmTagEdit.Show 1, fMainForm
   Exit Sub
  End If
 If TagList.SelectedItem.Text = "<frame>" Then fMainForm.ActiveForm.rtfText.SelText = "<frame> </frame>"
 If TagList.SelectedItem.Text = "<H1>" Then fMainForm.ActiveForm.rtfText.SelText = "<H1> </H1>"
 If TagList.SelectedItem.Text = "<H2>" Then fMainForm.ActiveForm.rtfText.SelText = "<H2> </H2>"
 If TagList.SelectedItem.Text = "<H3>" Then fMainForm.ActiveForm.rtfText.SelText = "<H3> </H3>"
 If TagList.SelectedItem.Text = "<H4>" Then fMainForm.ActiveForm.rtfText.SelText = "<H4> </H4>"
 If TagList.SelectedItem.Text = "<H5>" Then fMainForm.ActiveForm.rtfText.SelText = "<H5> </H5>"
 If TagList.SelectedItem.Text = "<H6>" Then fMainForm.ActiveForm.rtfText.SelText = "<H6> </H6>"
  If TagList.SelectedItem.Text = "<hr>" Then
   Unload Me
   TabNumber = 5
   frmTagEdit.Show 1, fMainForm
   Exit Sub
  End If
   If TagList.SelectedItem.Text = "<table>" Then
    Unload Me
    TabNumber = 13
    frmTagEdit.Show 1, fMainForm
    Exit Sub
   End If
   If TagList.SelectedItem.Text = "<tr>" Then
    Unload Me
    TabNumber = 15
    frmTagEdit.Show 1, fMainForm
    Exit Sub
   End If
   If TagList.SelectedItem.Text = "<td>" Then
    Unload Me
    TabNumber = 14
    frmTagEdit.Show 1, fMainForm
    Exit Sub
   End If
   
 If TagList.SelectedItem.Text = "<th>" Then fMainForm.ActiveForm.rtfText.SelText = "<th> </th>"
 If TagList.SelectedItem.Text = "<Input type=" & Chr(34) & "button" & Chr(34) & ">" Then fMainForm.ActiveForm.rtfText.SelText = "<Input type=" & Chr(34) & "button" & Chr(34) & ">"
  If TagList.SelectedItem.Text = "<Input type=" & Chr(34) & "checkbox" & Chr(34) & ">" Then
   Unload Me
   TabNumber = 2
   frmTagEdit.Show 1, fMainForm
   Exit Sub
  End If
 If TagList.SelectedItem.Text = "<Input type=" & Chr(34) & "File" & Chr(34) & ">" Then fMainForm.ActiveForm.rtfText.SelText = "<Input type=" & Chr(34) & "File" & Chr(34) & ">"
  If TagList.SelectedItem.Text = "<Input type=" & Chr(34) & "hidden" & Chr(34) & ">" Then
   Unload Me
   TabNumber = 11
   frmTagEdit.Show 1, fMainForm
   Exit Sub
  End If
 If TagList.SelectedItem.Text = "<Input type=" & Chr(34) & "image" & Chr(34) & ">" Then fMainForm.ActiveForm.rtfText.SelText = "<Input type=" & Chr(34) & "image" & Chr(34) & ">"
 If TagList.SelectedItem.Text = "<Input type=" & Chr(34) & "password" & Chr(34) & ">" Then fMainForm.ActiveForm.rtfText.SelText = "<Input type=" & Chr(34) & "password" & Chr(34) & ">"
  If TagList.SelectedItem.Text = "<Input type=" & Chr(34) & "radio" & Chr(34) & ">" Then
   Unload Me
   TabNumber = 7
   frmTagEdit.Show 1, fMainForm
   Exit Sub
  End If
 If TagList.SelectedItem.Text = "<Input type=" & Chr(34) & "reset" & Chr(34) & ">" Then fMainForm.ActiveForm.rtfText.SelText = "<Input type=" & Chr(34) & "reset" & Chr(34) & ">"
  If TagList.SelectedItem.Text = "<Input type=" & Chr(34) & "submit" & Chr(34) & ">" Then
   Unload Me
   TabNumber = 9
   frmTagEdit.Show 1, fMainForm
   Exit Sub
  End If
  If TagList.SelectedItem.Text = "<Input type=" & Chr(34) & "text" & Chr(34) & ">" Then
   Unload Me
   TabNumber = 12
   frmTagEdit.Show 1, fMainForm
   Exit Sub
  End If
  If TagList.SelectedItem.Text = "<textarea>" Then
   Unload Me
   TabNumber = 10
   frmTagEdit.Show 1, fMainForm
   Exit Sub
  End If
 Unload frmTags
End Sub
Private Sub Form_Load()
On Error Resume Next
 TagList.ListItems.Add , , "<textarea>", , 1
 TagList.ListItems.Add , , "<img>", , 1
 TagList.ListItems.Add , , "<a>", , 1
 TagList.ListItems.Add , , "<br>", , 1
 TagList.ListItems.Add , , "<blink>", , 1
 TagList.ListItems.Add , , "<center>", , 1
 TagList.ListItems.Add , , "<dir>", , 1
 TagList.ListItems.Add , , "<div>", , 1
 TagList.ListItems.Add , , "<head>", , 1
 TagList.ListItems.Add , , "<html>", , 1
 TagList.ListItems.Add , , "<menu>", , 1
 TagList.ListItems.Add , , "<nobr>", , 1
 TagList.ListItems.Add , , "<noembad>", , 1
 TagList.ListItems.Add , , "<nolayer>", , 1
 TagList.ListItems.Add , , "<plain text>", , 1
 TagList.ListItems.Add , , "<strong>", , 1
 TagList.ListItems.Add , , "<strike>", , 1
 TagList.ListItems.Add , , "<title>", , 1
 TagList.ListItems.Add , , "<u>", , 1
 TagList.ListItems.Add , , "<ul>", , 1
 TagList.ListItems.Add , , "<body>", , 1
 TagList.ListItems.Add , , "<font>", , 1
 TagList.ListItems.Add , , "<frame>", , 1
 TagList.ListItems.Add , , "<H1>", , 1
 TagList.ListItems.Add , , "<H2>", , 1
 TagList.ListItems.Add , , "<H3>", , 1
 TagList.ListItems.Add , , "<H4>", , 1
 TagList.ListItems.Add , , "<H5>", , 1
 TagList.ListItems.Add , , "<H6>", , 1
 TagList.ListItems.Add , , "<hr>", , 1
 TagList.ListItems.Add , , "<table>", , 1
 TagList.ListItems.Add , , "<tr>", , 1
 TagList.ListItems.Add , , "<td>", , 1
 TagList.ListItems.Add , , "<th>", , 1
 TagList.ListItems.Add , , "<Input type=" & Chr(34) & "button" & Chr(34) & ">", , 1
 TagList.ListItems.Add , , "<Input type=" & Chr(34) & "checkbox" & Chr(34) & ">", , 1
 TagList.ListItems.Add , , "<Input type=" & Chr(34) & "File" & Chr(34) & ">", , 1
 TagList.ListItems.Add , , "<Input type=" & Chr(34) & "hidden" & Chr(34) & ">", , 1
 TagList.ListItems.Add , , "<Input type=" & Chr(34) & "image" & Chr(34) & ">", , 1
 TagList.ListItems.Add , , "<Input type=" & Chr(34) & "password" & Chr(34) & ">", , 1
 TagList.ListItems.Add , , "<Input type=" & Chr(34) & "radio" & Chr(34) & ">", , 1
 TagList.ListItems.Add , , "<Input type=" & Chr(34) & "reset" & Chr(34) & ">", , 1
 TagList.ListItems.Add , , "<Input type=" & Chr(34) & "submit" & Chr(34) & ">", , 1
 TagList.ListItems.Add , , "<Input type=" & Chr(34) & "text" & Chr(34) & ">", , 1
End Sub

Private Sub Command2_Click()
 Unload Me
End Sub

'Call Select Button
Private Sub TagList_DblClick()
On Error Resume Next
 Command1_Click
End Sub

'Call Select Button
Private Sub TagList_KeyPress(KeyAscii As Integer)
On Error Resume Next
 If KeyAscii = 13 Then Command1_Click
End Sub
