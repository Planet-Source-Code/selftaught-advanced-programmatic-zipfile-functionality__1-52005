Attribute VB_Name = "mFormGeneral"
Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function GetClientRect Lib "user32" (ByVal Hwnd As Long, lpRect As RECT) As Long

Public Sub GetClientDimensions(ByVal Hwnd As Long, Height As Long, Width As Long)
    Dim lR As RECT
    GetClientRect Hwnd, lR
    With lR
        Height = .Bottom - .Top
        Width = .Right - .Left
    End With
End Sub
