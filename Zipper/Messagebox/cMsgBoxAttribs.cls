VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRichDialogGUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public GradientType    As eGradientTypes
Public FontStyle       As eFontStyles
Public CaptionEffect   As eCaptionEffects
Public ButtonDrawStyle As eButtonDrawStyles

Public GradientColor   As OLE_COLOR
Public BackColor       As OLE_COLOR

Public ShowFocusRect   As Boolean
Public ShowDivider     As Boolean

Public ButtonFont      As StdFont
Public InputFont       As StdFont

Public Function Clone() As cRichDialogGUI
    Set Clone = New cRichDialogGUI
    With Clone
        .BackColor = BackColor
        .ButtonDrawStyle = ButtonDrawStyle
        Set .ButtonFont = ButtonFont
        .CaptionEffect = CaptionEffect
        .FontStyle = FontStyle
        .GradientColor = GradientColor
        .GradientType = GradientType
        Set .InputFont = InputFont
        .ShowDivider = ShowDivider
        .ShowFocusRect = ShowFocusRect
    End With
End Function
