Attribute VB_Name = "mShared"
Option Explicit
'Public Const gDataKey = "PrimaryKey"
Public Const fldPrimaryKey = "PKey"
Public Const gsKeySuffix = " ID"

Private Const MAX_DWORD = &HFFFF

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetQueueStatus Lib "User32" (ByVal qsFlags As Long) As Long
Private Declare Sub ReleaseCapture Lib "User32" ()

Public Sub LockWindowRedraw(piHwnd As Long, pbState As Boolean)
    Const WM_SETREDRAW = &HB
    SendMessage piHwnd, WM_SETREDRAW, CInt(pbState), 0&
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

Public Function WindowsPath() As String
    WindowsPath = String(MAX_PATH, vbNullChar)
    WindowsPath = Left$(WindowsPath, GetWindowsDirectory(WindowsPath, Len(WindowsPath)))
End Function

Public Sub FilterNumericKeyAscii(KeyAscii As Integer)
    Select Case KeyAscii
        Case 3, 22, 24
        Case Else
            If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) And KeyAscii <> 8 Then KeyAscii = 0
    End Select
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

Public Function Inc(piVal As Long, Optional piInc As Long = 1)
    If piVal > 2147483647 - piInc Then
        piVal = piInc - (2147483647 - piVal)
    Else
        piVal = piVal + piInc
    End If
End Function























'Public Function HiByte(ByVal Word As Integer) As Byte
  'HiByte = (Word And &HFF00&) \ &H100
'End Function

'Function LoByte(byval w As Integer) As Byte
  'LoByte = w And &HFF
'End Function

Public Function HiWord(ByVal lDWord As Long) As Integer
    HiWord = (lDWord And &HFFFF0000) \ &H10000
End Function

Public Function LoWord(ByVal lDWord As Long) As Integer
    If lDWord And &H8000& Then
        LoWord = lDWord Or &HFFFF0000
    Else
        LoWord = lDWord And &HFFFF&
    End If
End Function

Public Function MakeLong(ByVal LoWord As Integer, ByVal HiWord As Integer) As Long
    MakeLong = (HiWord * &H10000) Or (LoWord And &HFFFF&)
End Function

Public Function MakeQWord(ByVal piLow As Long, ByVal piHigh As Long) As Double
    MakeQWord = (piHigh * MAX_DWORD + piLow)
End Function

Public Function CollClone(ByVal Coll As Collection) As Collection
    Set CollClone = New Collection
    Dim lvTemp
    With CollClone
        For Each lvTemp In Coll
            .Add lvTemp
        Next
    End With
End Function

Public Function OptIndex(ByVal Opts As Object) As Long
Dim lvTemp As OptionButton
    For Each lvTemp In Opts
        If lvTemp.Value Then
            OptIndex = lvTemp.Index
            Exit Function
        End If
    Next
End Function

Public Function OnlyNums(ByVal psVal As String) As String
OnlyNums = String(Len(psVal), " ")
Dim i As Integer, ReturnValPlace As Integer
For i = 1 To Len(psVal)
    If IsNumeric(Mid$(psVal, i, 1)) Then
        ReturnValPlace = ReturnValPlace + 1
        Mid$(OnlyNums, ReturnValPlace, 1) = Mid$(psVal, i, 1)
    End If
Next
OnlyNums = Trim$(OnlyNums)
End Function

Public Sub PasteOnlyNums(ByVal poTextBox As TextBox)
    Dim lsTemp As String
    lsTemp = OnlyNums(Clipboard.GetText)
    'Dim lsText As String
    'Dim liTemp As Long
    On Error Resume Next
    If Len(lsTemp) > 0 Then poTextBox.SelText = lsTemp
            'lsText = .Text
            'liTemp = .SelStart
            'lsText = Mid$(lsText, 1, liTemp) & lsTemp & Mid$(lsText, liTemp + .SelLength + 1)
            'liTemp = liTemp + Len(lsTemp)
            '.Text = lsText
            '.SelStart = liTemp
        'End With
    'End If
End Sub

Public Function IfUserInputOrPaintThenDoEvents()
    Const QS_HOTKEY = &H80
    Const QS_KEY = &H1
    Const QS_MOUSEBUTTON = &H4
    Const QS_PAINT = &H20
    
    Const QS = QS_HOTKEY Or QS_KEY Or QS_MOUSEBUTTON Or QS_PAINT
    IfUserInputOrPaintThenDoEvents = GetQueueStatus(QS) <> 0
    If IfUserInputOrPaintThenDoEvents Then DoEvents
End Function

Public Sub MoveForm(piHwnd As Long)
    Const WM_NCLBUTTONDOWN = &HA1
    Const HTCAPTION = 2
    ReleaseCapture
    SendMessage piHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub




Public Function GetInt(pyData() As Byte, ByVal piPlace As Long) As Integer
    If piPlace + 2 <= UBound(pyData) Then CopyMemory GetInt, pyData(piPlace), 2
End Function

Public Function GetLong(pyData() As Byte, ByVal piPlace As Long) As Long
    If piPlace + 4 <= UBound(pyData) Then CopyMemory GetLong, pyData(piPlace), 4
End Function

Public Function GetString(pyData() As Byte, ByVal piPlace As Long) As String
    Dim liInt As Integer
    If piPlace + 2 < UBound(pyData) Then
        CopyMemory liInt, pyData(piPlace), 2
        If liInt + piPlace <= UBound(pyData) Then
            GetString = Space$(liInt)
            CopyMemory ByVal StrPtr(GetString), pyData(piPlace + 2), liInt
            GetString = Left$(StrConv(GetString, vbUnicode), liInt)
        End If
    End If
End Function
