VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ListSearch 
   ClientHeight    =   1485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1965
   ScaleHeight     =   99
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   131
   ToolboxBitmap   =   "ListSearch.ctx":0000
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1035
      Top             =   1050
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
            Picture         =   "ListSearch.ctx":0312
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstMain 
      Height          =   1065
      Left            =   0
      TabIndex        =   1
      Top             =   405
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1879
      View            =   2
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "ListSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*************************************************************
' Project:  IntellIm                                         *
' Filename: ListSearch.ctl                                   *
' Author:   Edward P. Denninger III                          *
' Date:     5/14/2000                                        *
' Copyright © 2000 Edward P. Denninger III                   *
'*************************************************************
'*                         NOTICE                            *
'*************************************************************
' You may use and freely distribute this porject and source  *
' at your own leisure as long as I am given credit for my    *
' work.  If you have any comments or ideas for improvement,  *
' you can reach me at: edward3@optonline.net                 *
'*************************************************************
'
' All right this thing doesn't yet work with midi windows.
' All it needs is to change the scalemode to 3 in formload
' but I find out about this too late, so I'll try work it out
' otherwuays.
'                             Comment by: Vladimir S. Pekulas

Option Explicit

' API Calls
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

' API Constants
Private Const GWL_STYLE = (-16)
Private Const LB_ADDSTRING = &H180
Private Const LB_FINDSTRING = &H18F
Private Const LB_RESETCONTENT = &H184
Private Const WS_DLGFRAME = &H400000

' Events
Public Event Done(ByVal Text As String)
Public Event Escape()
'

Public Function FillWithTags(ByVal Filename As String, Optional Uppercase As Boolean = True) As Integer
    On Error GoTo hErr
    
    Dim strTag As String, fFile As Integer
    fFile = FreeFile
    ' Clear the listbox
    SendMessage lstMain.hwnd, LB_RESETCONTENT, 0&, ByVal 0&
    
    ' Fill the listbox with the tags from the file
    Open Filename For Input As fFile
        Do Until EOF(fFile)
            Line Input #fFile, strTag
            
            If Uppercase Then
                ' If there is a tag then add it to the listbox
                If Len(strTag) > 0 Then lstMain.ListItems.Add , , UCase$(strTag), , 1
            Else
                ' If there is a tag then add it to the listbox
                If Len(strTag) > 0 Then lstMain.ListItems.Add , , UCase$(strTag), , 1
            End If
        Loop
    Close fFile
    
    ' Return a value because we completed successfully
    FillWithTags = 1
    
hErr:
    Select Case Err.Number
    Case 0:
        Exit Function
    Case 53:
        MsgBox "Couldn't find the specified tags file!", vbExclamation, "Error"
        FillWithTags = 0
        Exit Function
    Case Else:
        MsgBox Err.Description, vbExclamation, "Error #" & Err.Number
        FillWithTags = 0
        Exit Function
    End Select
End Function

Public Sub FillWithFonts()
    Dim FontCounter As Long
    
    ' Clear the listbox
    SendMessage lstMain.hwnd, LB_RESETCONTENT, 0&, ByVal 0&
    
    ' Add the fonts
    For FontCounter = 0 To Screen.FontCount - 1
        
        SendMessage lstMain.hwnd, LB_ADDSTRING, 0&, ByVal Screen.Fonts(FontCounter) 'LCase$(strTag)
    Next FontCounter
End Sub

Public Sub FillWithDrives(Optional Uppercase As Boolean = True)
    Dim strSave As String
    Dim ret     As Long
    Dim keer    As Integer
    
    ' Clear the listbox
    SendMessage lstMain.hwnd, LB_RESETCONTENT, 0&, ByVal 0&
    
    '--------------------------------------------------------------
    'KPD-Team 1998
    'URL: http://www.allapi.net/
    'E-Mail: KPDTeam@Allapi.net
    
    'Create a buffer to store all the drives
    strSave = String$(255, Chr$(0))
    
    'Get all the drives
    ret& = GetLogicalDriveStrings(255, strSave)
    
    'Extract the drives from the buffer and print them on the form
    For keer = 1 To 100
        If Left$(strSave, InStr(1, strSave, Chr$(0))) = Chr$(0) Then Exit For
        
        If Uppercase Then
            
            SendMessage lstMain.hwnd, LB_ADDSTRING, 0&, ByVal UCase$(Left$(strSave, InStr(1, strSave, Chr$(0)) - 1))
        Else
            
            SendMessage lstMain.hwnd, LB_ADDSTRING, 0&, ByVal LCase$(Left$(strSave, InStr(1, strSave, Chr$(0)) - 1))
        End If
        
        strSave = Right$(strSave, Len(strSave) - InStr(1, strSave, Chr$(0)))
    Next keer
    '--------------------------------------------------------------
End Sub

'---------------------------------------------------------
'- Control stuff -----------------------------------------
'---------------------------------------------------------
Private Sub lstMain_DblClick()
    
    lstMain_KeyPress (vbKeyReturn)
End Sub

Private Sub lstMain_KeyPress(KeyAscii As Integer)
    
    Select Case KeyAscii
    Case vbKeyReturn:
        
        RaiseEvent Done(lstMain.SelectedItem.Text)
        txtSearch.Text = vbNullString
    Case vbKeyEscape:
        
        RaiseEvent Escape
        txtSearch.Text = vbNullString
    Case vbKeyDown:
        
        txtSearch.Text = lstMain.SelectedItem.Text
    End Select
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyDown Then SetFocus lstMain.hwnd
End Sub

Private Sub txtSearch_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    Select Case KeyCode
    Case vbKeyReturn:
        
        RaiseEvent Done(txtSearch.Text)
        txtSearch.Text = vbNullString
    Case vbKeyEscape:
        
        RaiseEvent Escape
        txtSearch.Text = vbNullString
    Case Else:
        Dim lngListIndex As Long
        
        ' Get the list index of the item
        lngListIndex = SendMessage(lstMain.hwnd, LB_FINDSTRING, -1, ByVal txtSearch.Text)
        
        ' If the search string could not be found then...
        If lngListIndex = -1 Then
            
            Exit Sub
        Else    ' If the search string was found...
            
            lstMain.ListItems.Count = lngListIndex
        End If
    End Select
End Sub

Private Sub UserControl_Initialize() ' Put a raised border around the control
    SetWindowLong hwnd, GWL_STYLE, GetWindowLong(hwnd, GWL_STYLE) Or WS_DLGFRAME
End Sub

Private Sub UserControl_Paint()
    On Error Resume Next
    txtSearch.Move 0, 0, ScaleWidth
    lstMain.Move 0, txtSearch.Height + 3, ScaleWidth, ScaleHeight - lstMain.Top + 3
End Sub
