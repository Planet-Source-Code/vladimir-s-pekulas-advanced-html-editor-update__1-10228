Attribute VB_Name = "Wraping"
Private Type TEXTMETRIC
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Byte
    tmLastChar As Byte
    tmDefaultChar As Byte
    tmBreakChar As Byte
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
    End Type
Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long
Private Declare Function SetMapMode Lib "gdi32" (ByVal hdc As Long, ByVal nMapMode As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const MM_TWIPS = 6
Public Const EM_LINEINDEX = &HBB
Public Const EM_LINELENGTH = &HC1
Public Const EM_GETLINECOUNT = &HBA
Public Const EM_LINEFROMCHAR = &HC9
Public mlngWidth As Long ' RightMargin in Richtext control



Public Function FindLongestLine() As Long
    Dim lngfound&, lngBeg&, lnglength&
    Dim lngLargeNum&
    Dim currLine&, lineCount&, lineLength&, lineIndex&
    Dim i%
    
    ' Get Line Count
    lineCount& = SendMessageLong(frmMain.ActiveForm.rtfText.hwnd, EM_GETLINECOUNT, 0&, 0&)
    


    For i = 0 To lineCount - 1
        ' Get character index on line number(wPa
        '     ram)
        lineIndex = SendMessageLong(frmMain.ActiveForm.rtfText.hwnd, EM_LINEINDEX, i, 0&)
        ' Get the lineLength at the lineIndex
        lineLength = SendMessageLong(frmMain.ActiveForm.rtfText.hwnd, EM_LINELENGTH, lineIndex, 0&)


        If lineLength > lngLargeNum Then
            lngLargeNum = lineLength
        End If
    Next
    
    FindLongestLine = lngLargeNum
End Function


Public Function GetFontWidth() As Long
    Dim hdc As Long
    Dim hwnd As Long
    Dim PrevMapMode As Long
    Dim tm As TEXTMETRIC
    
    ' Handle to the text area
    hwnd = frmMain.ActiveForm.rtfText.hwnd
    
    ' Get the device context for the desktop
    '
    hdc = GetWindowDC(hwnd)
    


    If hdc Then
        ' Set the mapping mode to twips
        PrevMapMode = SetMapMode(hdc, MM_TWIPS)
        
        ' Get the size of the system font
        GetTextMetrics hdc, tm
        ' Set the mapping mode back to what it w
        '     as
        PrevMapMode = SetMapMode(hdc, PrevMapMode)
        ' Release the device context
        ReleaseDC hwnd, hdc
    End If
    
    GetFontWidth = tm.tmMaxCharWidth
End Function



