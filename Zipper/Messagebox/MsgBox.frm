VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMsgBox 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2175
   ClientLeft      =   1395
   ClientTop       =   75
   ClientWidth     =   3885
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MsgBox.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   145
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   259
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Index           =   0
      Left            =   1620
      TabIndex        =   0
      Top             =   1200
      Width           =   1695
      Begin VB.CommandButton cmdButton 
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.CheckBox chk 
      Caption         =   "Do not show this dialog again"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   1200
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   420
      Width           =   495
   End
   Begin VB.TextBox txtInput 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1440
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin RichTextLib.RichTextBox txtMessage 
      Height          =   360
      Left            =   1440
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   240
      Width           =   2820
      _ExtentX        =   4974
      _ExtentY        =   635
      _Version        =   393217
      BackColor       =   -2147483633
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      MousePointer    =   1
      Appearance      =   0
      TextRTF         =   $"MsgBox.frx":000C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fra 
      Height          =   1095
      Index           =   1
      Left            =   720
      TabIndex        =   5
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label lblAutoAnswer 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   1500
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   1
      ToolTipText     =   "Double-click to halt the countdown."
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Menu mnu 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuMessage 
         Caption         =   "&Copy Message"
         Index           =   0
      End
      Begin VB.Menu mnuMessage 
         Caption         =   "&Print Message"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements iRichDialog
Implements iSubclass
Implements iTimer

Private moInfo             As cRichDialogInfo
Private moSubClass         As cSubclass
Private moTimer            As cTimer
Private moUtils            As cRTBUtilities
Private moGUI              As cRichDialogGUI

Private Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

Private moLastFocus As Control
Private miRichRequestHeight As Long
Private miMinButtonsWidth As Long
Private miTimeout As Long

Private Sub chk_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then MouseDown Button
End Sub

Private Sub cmdButton_Click(Index As Integer)
    On Error Resume Next
    With moInfo
        .Answer = GetAnswerFromCaption(cmdButton(Index).Caption, Index)
        If moInfo.InputBox Then
            Select Case .Answer
                Case rdCancel
                    .ReturnValue = Null
                Case Else
                    .ReturnValue = txtInput.Text
            End Select
        End If
    End With
    GoAway
End Sub

Private Sub GoAway()
    On Error Resume Next
    Me.Hide
    'DoEvents
    moTimer.TmrStop
    If LenB(chk.Caption) <> 0 Then
        moInfo.CheckBoxValue = CBool(chk.Value)
    End If
    moInfo.Notify.HasReturned Me
    txtMessage.TextRTF = ""
End Sub

Private Sub cmdButton_GotFocus(Index As Integer)
    Set moLastFocus = cmdButton(Index)
End Sub

Private Sub cmdButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnu, vbPopupMenuRightButton
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If moInfo.InputBox Then
        With txtInput
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
    End If
End Sub
'Private Sub Form_Activate()
'    On Error Resume Next
'    If ActiveControl Is Nothing Then
'        Dim Cham As Control
'        For Each Cham In cmdButton
'            If Cham.TabIndex = 0 Then
'                Cham.SetFocus
'                Exit For
'            End If
'        Next
'    End If
'End Sub

Private Sub Form_Initialize()
    On Error Resume Next
    Set moInfo = New cRichDialogInfo
    Set moUtils = New cRTBUtilities
    moUtils.Attach txtMessage.hWnd, Me.hWnd
    Set moSubClass = New cSubclass
    Set moTimer = New cTimer
    With moSubClass
        .Subclass Me.hWnd, Me
        .AddMsg WM_ACTIVATE, MSG_AFTER
        .AddMsg WM_MOUSEWHEEL, MSG_BEFORE
        .AddMsg WM_NOTIFY, MSG_BEFORE
    End With
    Set moGUI = goDefaultGUI.Clone
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim liCode As Long
    If KeyCode = vbKeyTab Then
        If BitIsSet(Shift, vbShiftMask) Then liCode = vbKeyUp Else liCode = vbKeyDown
    Else
        liCode = KeyCode
    End If
    If Not moUtils.ForwardScrollKey(liCode) Then
        If moInfo.InputBox Then Exit Sub
        Dim lsStr As String
        lsStr = "&" & Chr$(KeyCode)
        Dim loB As Control
        Dim liFound As Integer
        liFound = -1
        For Each loB In cmdButton
            If InStr(1, loB.Caption, lsStr, vbTextCompare) Then
                If liFound > -1 Then Exit Sub Else liFound = loB.Index
            End If
        Next
        If liFound > -1 Then cmdButton_Click liFound
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseDown Button
End Sub

Private Sub MouseDown(ByVal piButton As Long)
    If piButton = 2 Then
        PopupMenu mnu, vbPopupMenuRightButton
    ElseIf piButton = 1 Then
        If Not moInfo.Attributes And rdDisallowMove Then MoveForm hWnd
    End If
End Sub

'Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'    On Error Resume Next
'    If Button = 1 Then MoveForm hWnd
'End Sub

Private Sub Form_Terminate()
    On Error Resume Next
    Set moInfo = Nothing
    moUtils.Detach
    Set moUtils = Nothing
    moSubClass.UnSubclass
    Set moSubClass = Nothing
    Set moTimer = Nothing
    Set moGUI = Nothing
    Set moLastFocus = Nothing
End Sub

Private Sub fra_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseDown Button
End Sub

Private Property Get iRichDialog_GUI() As cRichDialogGUI
    Set iRichDialog_GUI = moGUI
End Property

Private Property Get iRichDialog_hWnd() As Long
    iRichDialog_hWnd = Me.hWnd
End Property

Private Property Set iRichDialog_Info(ByVal RHS As cRichDialogInfo)
    If RHS Is Nothing Then Set moInfo = New cRichDialogInfo Else Set moInfo = RHS
End Property

Private Property Get iRichDialog_Info() As cRichDialogInfo
    Set iRichDialog_Info = moInfo
End Property

Private Function iRichDialog_ReActivate() As Boolean
    On Error Resume Next
    If Me.Visible Then
        Show
        iRichDialog_ReActivate = Err.Number = 0
    End If
End Function

Private Property Get iRichDialog_RichRequestHeight() As Long
    iRichDialog_RichRequestHeight = miRichRequestHeight
End Property

Private Property Get iRichDialog_RichWnd() As Long
    iRichDialog_RichWnd = txtMessage.hWnd
End Property

Private Function iRichDialog_Show() As eRichDialogReturn
    On Error Resume Next
    Dim loC As Control
    Dim lbVal As Boolean
    
    With moInfo
        
        If Not .Notify Is Nothing Then
            .Notify.QueryInfo Me, lbVal
            If lbVal Then Exit Function
        End If
        
        txtMessage.TextRTF = .Message
        With moGUI
            If Not .InputFont Is Nothing Then Set txtInput.Font = .InputFont
            For Each loC In Controls
                If TypeOf loC Is CommandButton Then
                    If Not .ButtonFont Is Nothing Then Set loC.Font = .ButtonFont
                    loC.CaptionStyle = .CaptionEffect
                    loC.BackColor = .BackColor
                    loC.ButtonStyle = .ButtonDrawStyle
                    loC.GradientColor = .GradientColor
                    loC.GradientMode = .GradientType
                    loC.ShowFocusRect = .ShowFocusRect
                    loC.FontStyle = .FontStyle
                Else
                    If Not loC Is txtInput Then loC.BackColor = .BackColor
                End If
            Next
            Me.BackColor = .BackColor
            With fra(1)
                If moGUI.ShowDivider Then fra(1).BorderStyle = 1 Else fra(1).BorderStyle = 0
                .Visible = .BorderStyle = 1
            End With
        End With
        
        SizeArrange
        

        If Not SetInPosition(Me.hWnd, .hWndParent, .Position) Then CenterInWorkspace Me.hWnd
        
        If Not .Notify Is Nothing Then
            .Notify.WillShow Me, lbVal
            If lbVal Then Exit Function
        End If
        
        If miTimeout > 0 Then moTimer.TmrStart Me, 1000
        'If cmdButton.UBound = 0 Then cmdButton(0).ShowFocusRect = False
        If BitIsSet(.Attributes, rdTopMost) Then SetTopMost hWnd
        If BitIsSet(.Attributes, rdBeep) Then Beep
        If Not .IsModeless Then
            Show vbModal
        Else
            .Answer = rdAnswerPending
            Show vbModeless
        End If
        On Error GoTo 0
        If moInfo.Answer = rdCancel Then
            If moInfo.Attributes And rdCancelRaiseError Then Err.Raise 20001, , "User pressed cancel."
        End If
        iRichDialog_Show = moInfo.Answer
    End With

End Function

Private Sub iSubclass_Proc(ByVal bBefore As Boolean, bHandled As Boolean, lReturn As Long, hWnd As Long, uMsg As WinSubHook2.eMsg, wParam As Long, lParam As Long)
    On Error Resume Next
    Select Case uMsg
        Case WM_ACTIVATE
            DrawTitleBar CBool(wParam), Me, moInfo.Title
        Case WM_MOUSEWHEEL
            SendMessage txtMessage.hWnd, uMsg, wParam, lParam
        Case WM_NOTIFY
            On Error Resume Next
            miRichRequestHeight = GetRequestHeightFromNM(wParam, lParam)
    End Select
End Sub


Private Sub iTimer_Proc(ByVal lElapsedMS As Long, ByVal lTimerID As Long)
    miTimeout = miTimeout - 1
    If Len(lblAutoAnswer.Caption) > 0 Then
        lblAutoAnswer.Caption = "This message will close in " & miTimeout & IIf(miTimeout = 1, " second.", " seconds.")
        lblAutoAnswer.Refresh
    End If
    If miTimeout = 0 Then
        moInfo.Answer = rdAutoAnswer
        If moInfo.InputBox Then moInfo.ReturnValue = Null
        GoAway
    End If
End Sub

Private Sub lblAutoAnswer_DblClick()
    lblAutoAnswer.Visible = False
    moTimer.TmrStop
End Sub

Private Sub lblAutoAnswer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseDown Button
End Sub

Private Sub mnuMessage_Click(Index As Integer)
    On Error Resume Next
    With txtMessage
        LockWindowRedraw .hWnd, False
        Select Case Index
            Case 0
                .SelStart = 0
                .SelLength = Len(txtMessage.Text)
                Const WM_COPY = &H301
                SendMessage .hWnd, WM_COPY, 0&, 0&
                .SelStart = 0
            Case 1
                .SelStart = 0
                .SelAlignment = rtfLeft
                .SelBold = True
                .SelBullet = False
                .SelCharOffset = 0
                .SelColor = vbBlack
                .SelFontName = "Tahoma"
                .SelFontSize = 10
                .SelHangingIndent = 0
                .SelIndent = 0
                .SelItalic = False
                .SelRightIndent = 0
                .SelStrikeThru = False
                .SelUnderline = False
                
                .SelText = "Title:    " & Caption & vbCrLf & "Printed: " & Format(Now, "m/d/yy h:mm AMPM") & vbCrLf & vbCrLf & vbCrLf
                
                Printer.Print "";
                .SelPrint Printer.hDC
                If Err.Number > 0 Then Printer.KillDoc Else Printer.EndDoc
                .TextRTF = moInfo.Message
                
                    'MsgBoxEx "An error occurred while trying to print the message.  Be sure that a printer is installed and connected." & vbCrLf & "If this does not work, copy the message and paste it into another application." & vbCrLf & vbCrLf & "Error: " & Err.Number & vbCrLf & Err.Description, rdOKOnly + rdCritical + rdBeep, "Printer Error"
        End Select
        LockWindowRedraw .hWnd, True
        .Refresh
    End With
End Sub

Private Sub Picture1_GotFocus()
    On Error Resume Next
    If moLastFocus Is Nothing Then
        For Each moLastFocus In cmdButton
            If moLastFocus.Cancel Then
                moLastFocus.SetFocus
                Exit For
            End If
        Next
    Else
        moLastFocus.SetFocus
    End If
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseDown Button
End Sub

Private Sub ArrangeButtons()
    Const Spacing = 200
    On Error Resume Next
    
    Dim liNumButtons As Long
    Dim DefaultIndex As Long
    Dim CancelIndex As Long
    Dim i As Long
    Dim j As Long
        
    With moInfo
        .Data = GetCaptions(.Attributes, .Data)
        If Not IsArray(moInfo.Data) Then moInfo.Data = Array("&OK")
        liNumButtons = UBound(.Data) + 1
        miMinButtonsWidth = liNumButtons * ScaleX(cmdButton(0).Width + Spacing, vbTwips, vbPixels) - ScaleX(Spacing, vbTwips, vbPixels)
        GetDefaultCancelIndex .Attributes, DefaultIndex, CancelIndex
    End With

    If cmdButton.UBound + 1 < liNumButtons Then
        For i = cmdButton.Count To liNumButtons - 1
            Load cmdButton(i)
        Next
    Else
        For i = cmdButton.UBound To liNumButtons Step -1
            Unload cmdButton(i)
        Next
    End If
    
    fra(0).Width = miMinButtonsWidth
    miMinButtonsWidth = miMinButtonsWidth + 10
    With cmdButton
        For i = .LBound To .UBound
            With .Item(i)
                If i > liNumButtons - 1 Then
                    .Visible = False
                    .Caption = ""
                Else
                    .Visible = True
                    .Caption = moInfo.Data(i)
                    .Default = i = DefaultIndex - 1
                    .Cancel = i = CancelIndex - 1
                End If
            End With
        Next
    
        i = DefaultIndex - 1
        If i < 0 Then i = 0
        For j = 0 To liNumButtons - 1
            .Item(i).TabIndex = j
            i = i + 1
            If i >= liNumButtons Then i = 0
        Next
    
        
        
        Dim liTemp As Long
        Dim liButtonWidth As Long
        
        liButtonWidth = .Item(0).Width

        .Item(0).Left = 0
    
        For i = 1 To liNumButtons - 1
            With .Item(i)
                .Left = cmdButton(i - 1).Left + liButtonWidth + Spacing
                .Visible = True
            End With
        Next
    
        For i = liNumButtons To cmdButton.UBound
            With .Item(i)
                .Visible = False
                .Left = -liButtonWidth * 2
            End With
        Next
        
        If liNumButtons = 1 Then
            .Item(0).Default = True
            .Item(0).Cancel = True
        End If
    
    End With
End Sub

Private Sub iRichDialog_Activate()
    On Error Resume Next
    Show
End Sub


Private Sub DrawMyIcon()
    Dim hIcon As Long
    Dim lbCreated As Boolean
    hIcon = moInfo.hIcon
    lbCreated = hIcon = 0
    If lbCreated Then hIcon = GethIcon(moInfo.Attributes)
    With Picture1
        If hIcon <> 0 Then
            .Visible = True
            .Left = 8
            DrawIcon .hDC, 0, 0, hIcon
        Else
            .Visible = False
            .Left = 0 - .Width
        End If
    End With
    If lbCreated Then DestroyIcon hIcon
End Sub

Private Sub SizeArrange()
    On Error Resume Next
    ArrangeButtons
    DrawMyIcon
    
    Const FrameYOffset = 58
    Const RTBXPadding = 15
    Const RTBIconXPadding = 50
    Const RTBYPadding = 5
    Const MinWidth = 220
    Const MinHeight = 100
    
    Dim liHeight As Long
    Dim liWidth As Long
    Dim liRTBLeft As Long
    Dim liRTBHeight As Long
    Dim liRTBWidth As Long
    Dim liMaxWorkspaceWidth As Long
    Dim liMaxWorkspaceHeight As Long
    Dim liMiddleControlSpace As Long
    Dim liTemp As Long
    
    
    Dim lbShowCountdown As Boolean
    
    miTimeout = moInfo.Timeout
    If miTimeout < 0 Then miTimeout = 0
    If miTimeout > 0 And Not BitIsSet(moInfo.Attributes, rdHideTimeOutCountdown) Then
        lblAutoAnswer.Caption = "This message will close in " & miTimeout & " seconds."
        lblAutoAnswer.Visible = True
        lbShowCountdown = True
        liMiddleControlSpace = liMiddleControlSpace + lblAutoAnswer.Height + 4
    Else
        lblAutoAnswer.Caption = ""
        lblAutoAnswer.Visible = False
    End If
        
    If Not moInfo.InputBox Then
        txtInput.Visible = False
    Else
        With txtInput
            .Tag = " "
            .Text = moInfo.ReturnValue
            .SelStart = 0
            .SelLength = Len(.Text)
            .Tag = ""
            .TabIndex = 0
            .Visible = True
        End With
        liMiddleControlSpace = liMiddleControlSpace + txtInput.Height + 4
    End If
    
    If LenB(moInfo.CheckBoxStatement) <> 0 Then
        With chk
            .Value = Abs(moInfo.CheckBoxValue)
            .Visible = True
            .Caption = moInfo.CheckBoxStatement
            liMiddleControlSpace = liMiddleControlSpace + .Height + 4
        End With
    Else
        chk.Visible = False
    End If
    GetRTBDimensions Me, liRTBWidth, liRTBHeight
    If Picture1.Left > 0 Then _
        liRTBLeft = RTBIconXPadding _
    Else _
        liRTBLeft = RTBXPadding
    
    liWidth = liRTBWidth + liRTBLeft + RTBXPadding
    GetWorkspaceSize liMaxWorkspaceWidth, liMaxWorkspaceHeight

    If liWidth < miMinButtonsWidth Then liWidth = miMinButtonsWidth
    If EnsureBetween(liWidth, MinWidth, liMaxWorkspaceWidth) Then
        liRTBWidth = liWidth - liRTBLeft - RTBXPadding
        txtMessage.Width = liRTBWidth
        ForceAutoSize txtMessage.hWnd
        liRTBHeight = miRichRequestHeight
    End If
    chk.Width = ScaleWidth - chk.Left
    liHeight = liRTBHeight + FrameYOffset + RTBYPadding + CaptionHeight + liMiddleControlSpace
    If EnsureBetween(liHeight, MinHeight, liMaxWorkspaceHeight) Then _
        liRTBHeight = liHeight - CaptionHeight - RTBYPadding - FrameYOffset - liMiddleControlSpace
    
    moUtils.SetAutoSizeEventMask False
    
    liTemp = liHeight * 0.2
    EnsureBetween liTemp, CaptionHeight, liHeight - FrameYOffset
    Picture1.Top = liTemp
    txtMessage.Move liRTBLeft, RTBYPadding + CaptionHeight, liRTBWidth, liRTBHeight
    fra(1).Move -4, liHeight - FrameYOffset, liWidth + 20, FrameYOffset + 26
    fra(0).Top = liHeight - FrameYOffset + RTBXPadding
    Me.Move 0, 0, ScaleX(liWidth, vbPixels, vbTwips), ScaleY(liHeight, vbPixels, vbTwips)
    Dim liRTBBottom As Long
    liRTBBottom = RTBYPadding * 2 + liRTBHeight + CaptionHeight
    
    If moInfo.InputBox Then
        txtInput.Move liRTBLeft, liRTBBottom + 2, liWidth - liRTBLeft - RTBXPadding, txtInput.Height
        liRTBBottom = liRTBBottom + txtInput.Height + 4
    End If
    
    If lbShowCountdown Then
        lblAutoAnswer.Move liRTBLeft, liRTBBottom + 2, liWidth - liRTBLeft * 2, lblAutoAnswer.Height
        liRTBBottom = liRTBBottom + lblAutoAnswer.Height + 4
    End If
    
    chk.Move liRTBLeft, liRTBBottom + 2
    
    With fra(0)
        .Left = (liWidth \ 2 - .Width \ 2) - 2
    End With
    If BitIsSet(moInfo.Attributes, rdDisallowBlankInput) Then cmdButton(0).Enabled = txtInput.Text <> ""
End Sub

Private Sub txtInput_Change()
    If Len(txtInput.Tag) > 0 Then Exit Sub
    If BitIsSet(moInfo.Attributes, rdDisallowBlankInput) Then cmdButton(0).Enabled = txtInput.Text <> ""
    If miTimeout > 0 Then lblAutoAnswer_DblClick
End Sub

'Private Sub txtMessage_GotFocus()
'    Exit Sub
'    If moLastFocus Is Nothing Then
'        For Each moLastFocus In cmdButton
'            If moLastFocus.Default Then
'                moLastFocus.SetFocus
'                Exit For
'            End If
'        Next
'    Else
'        moLastFocus.SetFocus
'    End If
'End Sub

Private Sub txtMessage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseDown Button
End Sub
