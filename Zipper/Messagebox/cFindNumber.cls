VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFindNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private miMaxNum As Long
Private miMinNum As Long
Private miLastTry As Long
Private miDivisor As Long

Public Sub Init(ByVal Min As Long, ByVal Max As Long, Divisor As Long)
    miMaxNum = Max
    miMinNum = Min
    miDivisor = Divisor
    miLastTry = 0
End Sub

Public Function GuessNum() As Long
    If miLastTry = 0 Or miLastTry = miMaxNum Then
        GuessNum = miMinNum + Diff \ miDivisor
    Else
        GuessNum = miMaxNum - Diff \ miDivisor
    End If
    miLastTry = GuessNum
End Function

Public Sub TooLarge()
    miMaxNum = miLastTry
End Sub

Public Sub TooSmall()
    miMinNum = miLastTry
End Sub

Public Function Diff() As Long
    Diff = miMaxNum - miMinNum
End Function

Public Function Max() As Long
    Max = miMaxNum
End Function

Public Function Min() As Long
    Min = miMinNum
End Function
