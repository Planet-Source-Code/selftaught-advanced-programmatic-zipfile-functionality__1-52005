VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRTBUtilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements iSubclass

Private moSubClassRTB As cSubclass

'Constants used for Auto Resize
Const WM_USER = &H400
Const EM_GETEVENTMASK = (WM_USER + 59)
Const EM_SETEVENTMASK = (WM_USER + 69)
Const ENM_REQUESTRESIZE As Long = &H40000
Const EM_REQUESTRESIZE = (WM_USER + 65)

Private miRTBhwnd As Long
Private miFormhwnd As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Sub Attach(piRTBhwnd As Long, piFormHwnd As Long)
    On Error Resume Next
    Detach
    With moSubClassRTB
        .Subclass piRTBhwnd, Me
        .AddMsg WM_LBUTTONDOWN, MSG_BEFORE
        .AddMsg WM_LBUTTONDBLCLK, MSG_BEFORE
        .AddMsg WM_RBUTTONDOWN, MSG_BEFORE
        .AddMsg WM_RBUTTONDBLCLK, MSG_BEFORE
        .AddMsg WM_RBUTTONUP, MSG_BEFORE
        .AddMsg WM_LBUTTONUP, MSG_BEFORE
        .AddMsg WM_SETFOCUS, MSG_BEFORE
        .AddMsg WM_NCRBUTTONDOWN, MSG_BEFORE
    End With
    miRTBhwnd = piRTBhwnd
    miFormhwnd = piFormHwnd
     
    SetAutoSizeEventMask True
End Sub

Public Sub Detach()
    On Error Resume Next
    SetAutoSizeEventMask False
    moSubClassRTB.UnSubclass
    miRTBhwnd = 0
End Sub

Public Function ForwardScrollKey(ByVal piKey As KeyCodeConstants) As Boolean
    On Error Resume Next
    Const WM_VSCROLL = &H115
    Const SB_ENDSCROLL = 8
    Const SB_LINEDOWN = 1
    Const SB_PAGEDOWN = 3
    Const SB_PAGEUP = 2
    Const SB_LINEUP = 0
    Const SB_BOTTOM = 7
    Const SB_TOP = 6
    If miRTBhwnd = 0 Then Exit Function
    Select Case piKey
        Case vbKeyUp
            SendMessage miRTBhwnd, WM_VSCROLL, SB_LINEUP, 0&
        Case vbKeyDown
            SendMessage miRTBhwnd, WM_VSCROLL, SB_LINEDOWN, 0&
        Case vbKeyPageDown
            SendMessage miRTBhwnd, WM_VSCROLL, SB_PAGEDOWN, 0&
        Case vbKeyPageUp
            SendMessage miRTBhwnd, WM_VSCROLL, SB_PAGEUP, 0&
        Case vbKeyHome
            SendMessage miRTBhwnd, WM_VSCROLL, SB_TOP, 0&
        Case vbKeyEnd
            SendMessage miRTBhwnd, WM_VSCROLL, SB_BOTTOM, 0&
        Case Else
            Exit Function
    End Select
    ForwardScrollKey = True
    SendMessage miRTBhwnd, WM_VSCROLL, SB_ENDSCROLL, 0&
End Function

Public Sub ForceCalcAutoSize()
    ForceAutoSize miRTBhwnd
End Sub

Private Sub Class_Initialize()
    Set moSubClassRTB = New cSubclass
End Sub

Private Sub Class_Terminate()
    Detach
    Set moSubClassRTB = Nothing
End Sub


Public Property Get RTBhwnd() As Long
    RTBhwnd = miRTBhwnd
End Property

Public Sub SetAutoSizeEventMask(pbVal As Long)
    If miRTBhwnd = 0 Then Exit Sub
    If pbVal Then
        SendMessage miRTBhwnd, EM_SETEVENTMASK, 0, ByVal ENM_REQUESTRESIZE
    Else
        SendMessage miRTBhwnd, EM_SETEVENTMASK, 0, 0&
    End If
End Sub

Private Sub iSubclass_Proc(ByVal bBefore As Boolean, bHandled As Boolean, lReturn As Long, hwnd As Long, uMsg As WinSubHook2.eMsg, wParam As Long, lParam As Long)
    Select Case uMsg
        Case WM_NCRBUTTONDOWN, WM_RBUTTONDOWN, WM_LBUTTONDOWN
            SendMessage miFormhwnd, uMsg, wParam, lParam
    End Select
    bHandled = True
End Sub
