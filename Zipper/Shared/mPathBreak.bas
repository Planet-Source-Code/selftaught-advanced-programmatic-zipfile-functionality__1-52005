Attribute VB_Name = "mPathBreak"
Option Explicit

Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const EM_GETWORDBREAKPROC = &HD1
Public Const EM_SETWORDBREAKPROC = &HD0

Private Const WB_ISDELIMITER As Long = 2
Private Const WB_LEFT As Long = 0
Private Const WB_RIGHT As Long = 1

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private miEventsOnKey As Long

Public Function PathBreakProc(ByVal piTextPtr As Long, _
                                      ByVal piCurrentPos As Long, _
                                      ByVal piTextLen As Long, _
                                      ByVal piCode As Long) As Long
    On Error Resume Next
    Dim lsText As String
    lsText = Space$(piTextLen)
    CopyMemory ByVal StrPtr(lsText), ByVal piTextPtr, piTextLen
    lsText = Left$(StrConv(lsText, vbUnicode), piTextLen)
    Select Case piCode
        Case WB_ISDELIMITER
            PathBreakProc = Abs(IsPathBreak(Asc(Mid$(lsText, piCurrentPos + 1, 1))))
        Case WB_LEFT
            If piCurrentPos = 0 Then miEventsOnKey = 0
            If piCurrentPos < 1 Then piCurrentPos = 1
            For piCurrentPos = piCurrentPos - 1 To 1 Step -1
                If IsPathBreak(Asc(Mid$(lsText, piCurrentPos, 1))) Then Exit For
            Next
            PathBreakProc = piCurrentPos
            
        Case WB_RIGHT
again:
            For piCurrentPos = piCurrentPos + 1 To piTextLen
                If IsPathBreak(Asc(Mid$(lsText, piCurrentPos, 1))) Then Exit For
            Next
            PathBreakProc = piCurrentPos - 1
            If miEventsOnKey > 1 Then
                miEventsOnKey = 0
                GoTo again
            Else
                miEventsOnKey = 0
            End If
    End Select
    miEventsOnKey = miEventsOnKey + 1
End Function

Public Sub PathKeyDown()
    miEventsOnKey = 0
End Sub

Private Function IsPathBreak(ByVal piVal As Long) As Boolean
    Select Case piVal
        Case 32, 92, 59, 46 ' \;.
            IsPathBreak = True
    End Select
End Function

Public Sub SetPathbreakProc(ByVal piHwnd As Long)
    SendMessageLong piHwnd, EM_SETWORDBREAKPROC, 0, AddressOf PathBreakProc
End Sub
