VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRichDialogInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Message As String
Public Attributes As eRichDialogAttributes
Public Title As String
Public hIcon As Long
Public hWndParent As Long
Public Position As eRichDialogPosition
Public Timeout As Long
Public ReturnValue As Variant
Public Tag As Long
Public Data As Variant
Public Notify As iRichDialogParent
Public Answer As eRichDialogReturn
Public IsModeless As Boolean
Public InputBox As Boolean
Public CheckBoxStatement As String
Public CheckBoxValue As Boolean

'Private miID As Long
'
'Friend Property Let ID(piVal As Long)
'    miID = piVal
'End Property
'Friend Property Get ID() As Long
'    ID = miID
'End Property

'Friend Function Clone() As cRichDialogInfo
'    Set Clone = New cRichDialogInfo
'    With Clone
'        .Attributes = Me.Attributes
'        .Data = Me.Data
'        .hIcon = Me.hIcon
'        .hWndParent = Me.hWndParent
'        .Timeout = Me.Timeout
'        .Text = Me.Text
'        .Title = Me.Title
'    End With
'End Function

