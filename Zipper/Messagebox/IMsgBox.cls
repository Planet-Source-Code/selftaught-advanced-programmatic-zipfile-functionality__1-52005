VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iRichDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Show() As eRichDialogReturn
'
End Function

Public Property Get GUI() As cRichDialogGUI
'
End Property

Public Property Get Info() As cRichDialogInfo
'
End Property

Public Property Set Info(ByVal RDInfo As cRichDialogInfo)
'
End Property

Public Property Get hwnd() As Long
'
End Property

Public Property Get RichWnd() As Long
'
End Property

Public Property Get RichRequestHeight() As Long
'
End Property

Public Function ReActivate() As Boolean
'
End Function
