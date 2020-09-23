Attribute VB_Name = "mMsgBox"
Option Explicit

Public goDefaultGUI As New cRichDialogGUI
Public Const CaptionHeight = 19

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type NMHDR
   hwndFrom As Long
   idfrom As Long
   Code As Long
End Type

Private Type REQSIZE
   NMHDR As NMHDR
   RECT As RECT
End Type

Const WM_USER = &H400
Const EM_REQUESTRESIZE = (WM_USER + 65)

Private Declare Sub SetWindowPos Lib "user32" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
    Private Const HWND_TOPMOST = -1
    'Private Const HWND_NOTOPMOST = -2
    Private Const SWP_NOSIZE = &H1
    Private Const SWP_SHOWWINDOW = &H40
    Private Const SWP_NOACTIVATE = &H10
    Private Const SWP_NOMOVE = &H2
    Private Const SWP_NOZORDER = &H4
    Private Const SWP_NOOWNERZORDER = &H200
    Private Const SWP_NOREDRAW = &H8

Private Declare Function LoadStandardIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconNum As StandardIconEnum) As Long
    Private Enum StandardIconEnum
        IDI_ASTERISK = 32516&
        IDI_EXCLAMATION = 32515&
        IDI_HAND = 32513&
        IDI_QUESTION = 32514&
    End Enum

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
    Const SPI_GETWORKAREA& = 48

Private Declare Function DrawCaption Lib "user32" (ByVal Hwnd As Long, ByVal hDC As Long, pcRect As RECT, ByVal un As Long) As Long
    Const DC_ACTIVE = &H1
    Const DC_NOTACTIVE = &H2
    Const DC_ICON = &H4
    Const DC_TEXT = &H8
    Const DC_GRADIENT = &H20

Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal Hwnd As Long, ByVal lpString As String) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal Hwnd As Long, lpRect As RECT) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Sub ReleaseCapture Lib "user32" ()

Private moFindNumber As cFindNumber

Private miTop As Long
Private miLeft As Long

Public Function MsgBoxEx(ByVal Info As cRichDialogInfo, Optional ByVal SubstituteDialog As iRichDialog) As eRichDialogReturn
    Dim lbCreated As Boolean
    If SubstituteDialog Is Nothing Then
        lbCreated = True
        Set SubstituteDialog = DefaultDialog
    End If
    
    With SubstituteDialog
        Set .Info = Info
        .Show
        MsgBoxEx = .Info.Answer
    End With
    If lbCreated And Not Info.IsModeless Then
        Unload SubstituteDialog
        Set SubstituteDialog = Nothing
    End If
End Function

Public Function DefaultDialog() As iRichDialog
    Dim loForm As frmMsgBox
    Set loForm = New frmMsgBox
    Set DefaultDialog = loForm
    Set loForm = Nothing
End Function

Public Function InputBoxEx(ByVal Info As cRichDialogInfo, Optional ByVal SubstituteDialog As iRichDialog) As String
    Dim lbCreated As Boolean
    
    If SubstituteDialog Is Nothing Then
        lbCreated = True
        Set SubstituteDialog = DefaultDialog
    End If
    With SubstituteDialog
        Info.InputBox = True
        Set .Info = Info
        .Show
        InputBoxEx = .Info.ReturnValue & "" 'returnvalue is null if canceled and not (info.attributes and rdcancelreturnerror)
    End With
    If lbCreated And Not Info.IsModeless Then
        Unload SubstituteDialog
        Set SubstituteDialog = Nothing
    End If
End Function

Public Sub SetTopMost(piHwnd As Long)
    SetWindowPos piHwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE
End Sub

Public Sub Main()
    With goDefaultGUI
        .ButtonDrawStyle = rdw95
        .GradientType = rdNoGradient
        .ShowFocusRect = True
        .BackColor = vbButtonFace
        .ShowDivider = True
        Set .ButtonFont = New StdFont
        Set .InputFont = New StdFont
        .ButtonFont.Name = "Ariel"
        .InputFont.Name = "Ariel"
    End With
End Sub

Public Function GetCaptions(piStyle As eRichDialogAttributes, pvButtons)
    On Error Resume Next
    
    Dim lsArray() As String
    
    If Not IsArray(pvButtons) Then
        Select Case piStyle Mod 16
            Case rdOKCancel
                ReDim lsArray(0 To 1)
                lsArray(0) = "&OK"
                lsArray(1) = "&Cancel"
            Case rdYesNo
                ReDim lsArray(0 To 1)
                lsArray(0) = "&Yes"
                lsArray(1) = "&No"
            Case rdYesNoCancel
                ReDim lsArray(0 To 2)
                lsArray(0) = "&Yes"
                lsArray(1) = "&No"
                lsArray(2) = "&Cancel"
            Case rdAbortRetryIgnore
                ReDim lsArray(0 To 2)
                lsArray(0) = "&Abort"
                lsArray(1) = "&Retry"
                lsArray(2) = "&Ignore"
            Case rdRetryCancel
                ReDim lsArray(0 To 1)
                lsArray(0) = "&Retry"
                lsArray(1) = "&Cancel"
            Case Else
                ReDim lsArray(0 To 0)
                lsArray(0) = "&OK"
        End Select
    Else
        Dim i As Long
        i = LBound(pvButtons)
        If Err.Number <> 0 Then Exit Function
        
        For i = LBound(pvButtons) To UBound(pvButtons)
            If i = LBound(pvButtons) Then ReDim lsArray(0 To 0) Else ReDim Preserve lsArray(0 To UBound(lsArray) + 1)
            lsArray(UBound(lsArray)) = pvButtons(i)
            'If UBound(lsArray) = 3 Then Exit For
        Next
        Err.Clear
        i = UBound(lsArray)
        If Err.Number <> 0 Then ReDim lsArray(0 To 0)
    End If
    GetCaptions = lsArray
End Function

Public Function GetAnswerFromCaption(psCaption As String, ByVal piIndex As Long) As Variant
    Select Case psCaption
        Case "&OK"
            GetAnswerFromCaption = rdOK
        Case "&Cancel"
            GetAnswerFromCaption = rdCancel
        Case "&Yes"
            GetAnswerFromCaption = rdYes
        Case "&No"
            GetAnswerFromCaption = rdNo
        Case "&Retry"
            GetAnswerFromCaption = rdRetry
        Case "&Ignore"
            GetAnswerFromCaption = rdIgnore
        Case "&Abort"
            GetAnswerFromCaption = rdAbort
        Case Else
            GetAnswerFromCaption = rdButton1 + piIndex
    End Select
End Function

Public Function GethIcon(piStyle As Long) As Long
    Dim liIconType As Long
    
    If BitIsSet(piStyle, rdCritical) Then
        liIconType = IDI_HAND
    ElseIf BitIsSet(piStyle, rdExclamation) Then
        liIconType = IDI_EXCLAMATION
    ElseIf BitIsSet(piStyle, rdQuestion) Then
        liIconType = IDI_QUESTION
    ElseIf BitIsSet(piStyle, rdInformation) Then
        liIconType = IDI_ASTERISK
    End If
    
    If liIconType <> 0 Then GethIcon = LoadStandardIcon(0&, liIconType)
End Function

Public Sub ForceAutoSize(piRTBhwnd As Long)
    SendMessage piRTBhwnd, EM_REQUESTRESIZE, 0, 0&
End Sub

Public Sub GetRTBDimensions(ByVal Dialog As iRichDialog, Width As Long, Height As Long)
    Dim liRTBhwnd         As Long
    liRTBhwnd = Dialog.RichWnd
            
    If Width > 0 Then
        With Dialog
            SetWindowPos liRTBhwnd, 0, 0, 0, Screen.Height \ Screen.TwipsPerPixelY, Width, SWP_NOZORDER Or SWP_NOMOVE Or SWP_NOACTIVATE
            ForceAutoSize liRTBhwnd
            Height = .RichRequestHeight
        End With
    Else
        Dim liTempWidth       As Long
        Dim liRequestedHeight As Long
        Dim liMinRTFWidth     As Long
        Dim liMaxRTFWidth     As Long

        liMaxRTFWidth = Screen.Width / Screen.TwipsPerPixelX
        liMinRTFWidth = 50


        SetWindowPos liRTBhwnd, 0&, 0, 0, liMaxRTFWidth, liMinRTFWidth, SWP_NOMOVE
        SendMessage liRTBhwnd, EM_REQUESTRESIZE, 0&, 0&
        liRequestedHeight = Dialog.RichRequestHeight
        If liRequestedHeight > 0 Then
            
            If moFindNumber Is Nothing Then Set moFindNumber = New cFindNumber
            With moFindNumber
                .Init liMinRTFWidth, liMaxRTFWidth, 2
                Do While .Diff > 5
                    liTempWidth = .GuessNum
                    SetWindowPos liRTBhwnd, 0&, 0, 0, liTempWidth, liMinRTFWidth, SWP_NOMOVE  'Or SWP_SHOWWINDOW
                    SendMessage liRTBhwnd, EM_REQUESTRESIZE, 0, 0&
                    If liRequestedHeight < Dialog.RichRequestHeight Then .TooSmall Else .TooLarge
                Loop
                Width = .Max
                'liTempWidth = .Max
                'SetWindowPos liRTBhwnd, 0&, 0, 0, liTempWidth, liMinRTFWidth, SWP_SHOWWINDOW Or SWP_NOMOVE
            End With
        
        End If
    End If
    Height = liRequestedHeight
End Sub

'Public Sub zzzzSetMsgBoxPos(poMsgForm As iRichDialog, poUtils As cRTBUtilities, Width As Long, CenterOverHwnd As Long)
'    Const Inc = 25
'
'    Dim liWidth As Long, liHeight As Long
'    Dim liTop As Long, liLeft As Long
'
'    Dim liRTBRequestWidth As Long
'    Dim liRTBRequestHeight As Long
'
'
'    'Set poMsgForm.Utilities = poUtils
'    'poUtils.Utilize(mbRTBAutoSize) = True
'    'GetRTBDimensions poUtils, liRTBRequestWidth, liRTBRequestHeight, Width
'    'poMsgForm.SizeForm liRTBRequestWidth, liRTBRequestHeight
'
'    Dim ltRect As RECT
'    '
'    'GetWindowRect poUtils.FormHwnd, ltRect
'
'    With ltRect
'        liWidth = .Right - .Left
'        liHeight = .Bottom - .Top
'    End With
'
'
'    Dim WorkArea As RECT
'    SystemParametersInfo SPI_GETWORKAREA, 0&, WorkArea, 0&
'
'    If False And CenterOverHwnd = 0 Then
'
''        'If goRecursion Is Nothing Then Set goRecursion = New cRecursionMgr
''
''        'If goRecursion.StackCount = 1 Or (miTop = 0 And miLeft = 0) Then
''        If True Then
''            With ltRect
''                liTop = WorkArea.Top + (WorkArea.Bottom - WorkArea.Top) \ 2 - liHeight \ 2
''                liLeft = WorkArea.Left + (WorkArea.Right - WorkArea.Left) \ 2 - liWidth \ 2
''            End With
''
''        Else
''
''            liLeft = miLeft
''            liTop = miTop
''
''            With WorkArea
''                If liTop + liHeight > WorkArea.Bottom Then
''                    liLeft = liLeft * 0.3
''                    liTop = WorkArea.Top
''                ElseIf liLeft + liWidth > .Right Then
''                    liLeft = .Right - liWidth
''                End If
''            End With
''
''        End If
''
''        miLeft = liLeft + Inc
''        miTop = liTop + Inc
'    Else
'
'        Dim ltCenterRect As RECT
'        GetWindowRect CenterOverHwnd, ltCenterRect
'
'        With ltCenterRect
'            liLeft = .Left + (.Right - .Left) \ 2 - liWidth \ 2
'            liTop = .Top + (.Bottom - .Top) \ 2 - liHeight \ 2
'        End With
'
'        With ltRect
'            .Top = liTop
'            .Left = liLeft
'            .Bottom = liTop + liHeight
'            .Right = liLeft + liWidth
'            If .Right > WorkArea.Right Then
'                liLeft = liLeft - (.Right - WorkArea.Right)
'            ElseIf .Left < WorkArea.Left Then
'                liLeft = WorkArea.Left
'            End If
'
'            If .Bottom > WorkArea.Bottom Then
'                liTop = liTop - (.Bottom - WorkArea.Bottom)
'            ElseIf .Top < WorkArea.Top Then
'                liTop = WorkArea.Top
'            End If
'        End With
'
'    End If
'
'    'SetWindowPos poUtils.FormHwnd, 0, liLeft, liTop, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
'End Sub

Public Sub GetDefaultCancelIndex(Attributes As Long, piDefault As Long, piCancel As Long)
    If BitIsSet(Attributes, rdDefaultButton1) Then
        piDefault = 1
    ElseIf BitIsSet(Attributes, rdDefaultButton2) Then
        piDefault = 2
    ElseIf BitIsSet(Attributes, rdDefaultButton3) Then
        piDefault = 3
    ElseIf BitIsSet(Attributes, rdDefaultButton4) Then
        piDefault = 4
    ElseIf BitIsSet(Attributes, rdDefaultButton5) Then
        piDefault = 5
    ElseIf BitIsSet(Attributes, rdDefaultButton6) Then
        piDefault = 6
    End If
    
    If BitIsSet(Attributes, rdCancelButton1) Then
        piCancel = 1
    ElseIf BitIsSet(Attributes, rdCancelButton2) Then
        piCancel = 2
    ElseIf BitIsSet(Attributes, rdCancelButton3) Then
        piCancel = 3
    ElseIf BitIsSet(Attributes, rdCancelButton4) Then
        piCancel = 4
    ElseIf BitIsSet(Attributes, rdCancelButton5) Then
        piCancel = 5
    ElseIf BitIsSet(Attributes, rdCancelButton6) Then
        piCancel = 6
    End If
End Sub

Public Sub DrawTitleBar(pbActive As Boolean, poForm As Form, psCaption As String) '  piPixelWidth As Long, piPixelHeight As Long, piHwnd As Long, pihDC As Long, psCaption As String)
    On Error Resume Next
    Dim R As RECT
    Dim liVal As Long
    If pbActive Then liVal = DC_ACTIVE
    With poForm
        SetRect R, 0, 0, .ScaleX(.Width, vbTwips, vbPixels), CaptionHeight
        SetWindowText .Hwnd, psCaption & vbNullString
        .Cls
        DrawCaption .Hwnd, .hDC, R, liVal Or DC_TEXT Or DC_GRADIENT
    End With
End Sub

Public Sub MoveForm(piHwnd As Long)
    Const WM_NCLBUTTONDOWN = &HA1
    Const HTCAPTION = 2
    ReleaseCapture
    SendMessage piHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Public Function GetRequestHeightFromNM(ByVal wParam As Long, ByVal lParam As Long) As Long
    Const EN_REQUESTRESIZE = &H701
    Dim rResize As REQSIZE
    Dim MaskHdr As NMHDR
    
    Call CopyMemory(MaskHdr, ByVal lParam, Len(MaskHdr))
    
    If MaskHdr.Code = EN_REQUESTRESIZE Then
        Call CopyMemory(rResize, ByVal lParam, Len(rResize))
        GetRequestHeightFromNM = rResize.RECT.Bottom
    Else
        Err.Raise 5
    End If

End Function

Public Function CenterInWorkspace(piHwnd As Long) As Boolean
    Dim ltRect As RECT
    Dim liHeight As Long, liWidth As Long, liTop As Long, liLeft As Long
    
    If GetWindowRect(piHwnd, ltRect) = 0 Then Exit Function
    With ltRect
        liWidth = .Right - .Left
        liHeight = .Bottom - .Top
    End With
    
    If SystemParametersInfo(SPI_GETWORKAREA, 0&, ltRect, 0&) = 0 Then Exit Function

    With ltRect
        liLeft = .Left + (.Right - .Left) \ 2 - liWidth \ 2
        liTop = .Top + (.Bottom - .Top) \ 2 - liHeight \ 2
    End With
    
    SetWindowPos piHwnd, 0, liLeft, liTop, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
    CenterInWorkspace = True
End Function

Public Function SetInPosition(ByVal piHwndSetPos As Long, ByVal piHwndParent As Long, ByVal piPos As eRichDialogPosition) As Boolean
    If piHwndSetPos = 0 Or piHwndParent = 0 Then Exit Function
    
    Dim ltRectSetPos As RECT
    Dim ltRectParent As RECT
    Dim liWidth As Long, liHeight As Long, liLeft As Long, liTop As Long
    
    If GetWindowRect(piHwndSetPos, ltRectSetPos) = 0 Then Exit Function
    If GetWindowRect(piHwndParent, ltRectParent) = 0 Then Exit Function
    
    SetInPosition = True
    
    With ltRectSetPos
        liHeight = .Bottom - liTop
        liWidth = .Right - liLeft
    End With
    
    Select Case piPos
        Case rdCenterCenter
            With ltRectParent
                liLeft = .Left + (.Right - .Left) \ 2 - liWidth \ 2
                liTop = .Top + (.Bottom - .Top) \ 2 - liHeight \ 2
            End With
        Case rdCenterAbove
            With ltRectParent
                liLeft = .Left + (.Right - .Left) \ 2 - liWidth \ 2
                liTop = .Top - liHeight \ 2
            End With
        Case rdCenterBelow
            With ltRectParent
                liLeft = .Left + (.Right - .Left) \ 2 - liWidth \ 2
                liTop = .Bottom - liHeight \ 2
            End With
        Case rdLeftBelow
            With ltRectParent
                liLeft = .Left - liWidth \ 2
                liTop = .Bottom - liHeight \ 2
            End With
        Case rdLeftCenter
            With ltRectParent
                liLeft = .Left - liWidth \ 2
                liTop = .Top + (.Bottom - .Top) \ 2 - liHeight \ 2
            End With
        Case rdLeftAbove
            With ltRectParent
                liLeft = .Left - liWidth \ 2
                liTop = .Top - liHeight \ 2
            End With
        Case rdRightBelow
            With ltRectParent
                liLeft = .Right - liWidth \ 2
                liTop = .Bottom - liHeight \ 2
            End With
        Case rdRightCenter
            With ltRectParent
                liLeft = .Right - liWidth \ 2
                liTop = .Top + (.Bottom - .Top) \ 2 - liHeight \ 2
            End With
        Case rdRightAbove
            With ltRectParent
                liLeft = .Right - liWidth \ 2
                liTop = .Top - liHeight \ 2
            End With
    End Select
    
    EnsureInWorkArea liLeft, liTop, liWidth, liHeight
    
    
    SetWindowPos piHwndSetPos, 0, liLeft, liTop, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
    
End Function



Public Function EnsureInWorkArea(piLeft As Long, piTop As Long, ByVal piWidth As Long, ByVal piHeight As Long) As Boolean
    Dim WorkArea As RECT, piBottom As Long, piRight As Long
    Dim liRight As Long, liBottom As Long
    liRight = piLeft + piWidth
    liBottom = piTop + piHeight
    
    If SystemParametersInfo(SPI_GETWORKAREA, 0&, WorkArea, 0&) = 0 Then Exit Function
      
    If liRight > WorkArea.Right Then
        piLeft = piLeft - (liRight - WorkArea.Right)
    ElseIf piLeft < WorkArea.Left Then
        piLeft = WorkArea.Left
    End If
    
    If liBottom > WorkArea.Bottom Then
        piTop = piTop - (liBottom - WorkArea.Bottom)
    ElseIf piTop < WorkArea.Top Then
        piTop = WorkArea.Top
    End If
    EnsureInWorkArea = True
End Function

Public Sub GetWorkspaceSize(Width As Long, Height As Long)
    
    Dim WorkArea As RECT
    
    If SystemParametersInfo(SPI_GETWORKAREA, 0&, WorkArea, 0&) = 0 Then Exit Sub
    
    With WorkArea
        Width = .Right - .Left
        Height = .Bottom - .Top
    End With

End Sub


Public Sub Test()
    Main
    Dim loTemp As New gRichDialog
    goDefaultGUI.ButtonDrawStyle = rdhover
    loTemp.MsgBoxEx String(1500, "S"), rdAbortRetryIgnore + rdCancelButton3 + rdDefaultButton2, "Title"
    Exit Sub
        
    Dim loDialog As iRichDialog
    Dim loForm As frmMsgBox
    
    Do
        
        Set loForm = New frmMsgBox
        Set loDialog = loForm


        With loDialog.Info
            .InputBox = True
            .Message = String(1000, "M")
            .Title = "Title"
            .Timeout = 10
            .Attributes = rdAbortRetryIgnore + rdCancelButton3 + rdDefaultButton2 + rdBeep
            goDefaultGUI.ShowDivider = True
            .ReturnValue = "Defaultval"
        End With
        loDialog.Show
        Set loDialog = Nothing
        Unload loForm
        Set loForm = Nothing
    Loop
End Sub


Public Function BitIsSet(ByVal piVal As Long, ByVal piBit As Long) As Boolean
    BitIsSet = CBool(piVal And piBit)
End Function

Public Sub SetBit(piVal As Long, ByVal piBit As Long, ByVal pbState As Boolean)
    Dim lbVal As Boolean
    lbVal = BitIsSet(piVal, piBit)
    If Not lbVal = pbState Then
        If pbState Then
            piVal = piVal + piBit
        Else
            piVal = piVal - piBit
        End If
    End If
End Sub

Public Function EnsureBetween(piVal As Long, ByVal Min As Long, ByVal Max As Long) As Boolean
    If piVal < Min Then
        piVal = Min
    ElseIf piVal > Max Then
        piVal = Max
    Else
        Exit Function
    End If
    
    EnsureBetween = True
End Function

Public Sub LockWindowRedraw(piHwnd As Long, pbState As Boolean)
    Const WM_SETREDRAW = &HB
    SendMessage piHwnd, WM_SETREDRAW, CInt(pbState), 0&
End Sub
