VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDocuments 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Available Doc..."
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   1665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   1665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   195
      Top             =   1335
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
            Picture         =   "frmDocuments.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView AvDoc 
      Height          =   2475
      Left            =   0
      TabIndex        =   0
      Top             =   -15
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   4366
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Opened Documents"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmDocuments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AvDoc_DblClick()
'# If anybody knows how to open active document just like you'd do
'# from menu, plz let me know at vpekulas@home.com
 MsgBox "This function is not yet available.", vbInformation
 Unload Me
'fMainForm.ActiveForm Trim(Mid(AvDoc.SelectedItem.Text, 5, 4))
End Sub

Private Sub Form_Activate()
 Dim All As Integer, I As Integer
 All = fMainForm.tmpLST.ListCount - 1
 For I = 0 To All
  AvDoc.ListItems.Add , , fMainForm.tmpLST.List(I), , 1
 Next I
End Sub
