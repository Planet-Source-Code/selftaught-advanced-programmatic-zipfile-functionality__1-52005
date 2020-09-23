VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRecursionMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private miStackCount As Long
Private miPendingProcs As Long

Private mbCallReturned() As Boolean

Public Function CallReturned(piProcID As Long) As Boolean
    If IsValid(piProcID) Then
        CallReturned = mbCallReturned(piProcID)
    Else
        Err.Raise 5
    End If
End Function

Public Function EndProc(piProcID As Long)
    If miPendingProcs > 1 Then
        If IsValid(piProcID) Then
            If mbCallReturned(piProcID) Then Err.Raise 5
            miPendingProcs = miPendingProcs - 1
            mbCallReturned(piProcID) = True
'            For miStackCount = UBound(mbCallReturned) To 1 Step -1
'                If mbCallReturned(miStackCount) = True Then Exit For
'            Next
'            If miStackCount < UBound(mbCallReturned) And miStackCount > 0 Then ReDim Preserve mbCallReturned(LBound(mbCallReturned) To miStackCount)
        End If
        If miStackCount > 0 Then Exit Function
    End If
    ReDim mbCallReturned(-1 To -1)
    miPendingProcs = 0
    miStackCount = 0
End Function

'Public Function FirstFreeStackPlace() As Long
'    If miPendingProcs > 0 Then
'        For FirstFreeStackPlace = LBound(mbCallReturned) To UBound(mbCallReturned)
'            If mbCallReturned(FirstFreeStackPlace) Then Exit For
'        Next
'        FirstFreeStackPlace = FirstFreeStackPlace - 1
'    End If
'End Function

Public Function StackCount() As Long
    StackCount = miStackCount
End Function

Public Function StartProc() As Long
    'If miStackCount > 2140000000 Then miStackCount = 1 Else miStackCount = miStackCount + 1
    miStackCount = miStackCount + 1
    StartProc = miStackCount
    miPendingProcs = miPendingProcs + 1
    
    If miStackCount = 1 Then
        ReDim mbCallReturned(1 To 1)
    Else
        ReDim Preserve mbCallReturned(1 To miStackCount)
    End If
End Function

Public Function PendingProcs() As Long
    PendingProcs = miPendingProcs
End Function

Private Sub Class_Initialize()
    ReDim mbCallReturned(-1 To -1)
End Sub

Private Function IsValid(ByVal piNum As Long) As Boolean
    On Error Resume Next
    IsValid = piNum >= LBound(mbCallReturned) And piNum <= UBound(mbCallReturned) And piNum > 0
End Function
