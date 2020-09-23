VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "gRichDialog"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum eRichDialogPosition
    rdCenterCenter
    rdCenterAbove
    rdCenterBelow
    rdLeftBelow
    rdLeftCenter
    rdLeftAbove
    rdRightBelow
    rdRightCenter
    rdRightAbove
    rdCustom
End Enum

Public Enum eRichDialogAttributes
    rdOKOnly = vbOKOnly
    rdOKCancel = vbOKCancel
    rdAbortRetryIgnore = vbAbortRetryIgnore
    rdYesNoCancel = vbYesNoCancel
    rdYesNo = vbYesNo
    rdRetryCancel = vbRetryCancel
    rdCritical = 16
    rdQuestion = 32
    rdExclamation = 64
    rdInformation = 128
    rdBeep = 256
    rdTopMost = 512
    rdHideTimeOutCountdown = 1024
    rdCancelRaiseError = 2048
    rdDisallowMove = 4096
    rdDefaultButton1 = 8192
    rdDefaultButton2 = 16384
    rdDefaultButton3 = 32768
    rdDefaultButton4 = 65536
    rdDefaultButton5 = 131072
    rdDefaultButton6 = 262144
    rdCancelButton1 = 524288
    rdCancelButton2 = 1048576
    rdCancelButton3 = 2097152
    rdCancelButton4 = 4194304
    rdCancelButton5 = 8388608
    rdCancelButton6 = 16777216
    rdDisallowBlankInput = 33554432
'67108864
'134217728
'268435456
'536870912
'1073741824
End Enum

Public Enum eRichDialogReturn
    rdOK = vbOK
    rdCancel = vbCancel
    rdAbort = vbAbort
    rdRetry = vbRetry
    rdIgnore = vbIgnore
    rdYes = vbYes
    rdNo = vbNo
    rdButton1
    rdButton2
    rdButton3
    rdButton4
    rdButton5
    rdButton6
    rdAutoAnswer
    rdAnswerPending
End Enum

Public Enum eGradientTypes       ' gradient directions
    rdNoGradient = 0
    rdLeft2Right = 1
    rdRight2Left = 2
    rdTop2Bottom = 3
    rdBottom2Top = 4
End Enum
Public Enum eCaptionEffects  ' caption styles
    rdDefault = 0
    rdSunken = 1
    rdRaised = 2
End Enum
Public Enum eFontStyles
    rdPlainStyle = 0
    rdBold = 2
    rdItalic = 4
    rdUnderline = 8
    rdBoldItalic = 2 Or 4
    rdBoldUnderline = 2 Or 8
    rdItalicUnderline = 4 Or 8
    rdBoldItalicUnderline = 2 Or 4 Or 8
End Enum

Public Enum eButtonDrawStyles      ' button styles
    rdw95 = 0
    rdw31 = 1
    rdXP = 2
    rdJava = 3
    rdFlat = 4
    rdhover = 5
    rdNetscape = 6
    rdMacintosh = 7
End Enum

Public Function MsgBoxEx(Message As String, _
          Optional ByVal Attributes As eRichDialogAttributes = rdOKOnly + rdInformation, _
                Optional Title As String, _
          Optional ByVal hIcon As Long, _
          Optional ByVal hWndParent As Long, _
          Optional ByVal Position As eRichDialogPosition, _
                Optional Buttons, _
          Optional ByVal Timeout As Long, _
          Optional ByVal Tag As Long, _
          Optional ByVal Notify As iRichDialogParent, _
          Optional ByVal IsModeless As Boolean, _
          Optional ByVal SubstituteDialog As iRichDialog _
                ) As eRichDialogReturn
    
    Dim loInfo As cRichDialogInfo
    Set loInfo = New cRichDialogInfo
    With loInfo
        .Message = Message
        .Attributes = Attributes
        .Title = Title
        .hIcon = hIcon
        .hWndParent = hWndParent
        .Position = Position
        .Data = Buttons
        .Timeout = Timeout
        .Tag = Tag
        .IsModeless = IsModeless
        Set .Notify = Notify
    End With
    
    MsgBoxEx = mMsgBox.MsgBoxEx(loInfo, SubstituteDialog)
    Set loInfo = Nothing
End Function

Public Function InputBoxEx(Message As String, _
            Optional ByVal Attributes As eRichDialogAttributes = rdQuestion, _
                  Optional Title As String, _
                  Optional Default As String, _
            Optional ByVal hIcon As Long, _
            Optional ByVal hWndParent As Long, _
            Optional ByVal Position As eRichDialogPosition, _
                  Optional Buttons, _
            Optional ByVal Timeout As Long, _
            Optional ByVal Tag As Long, _
            Optional ByVal Notify As iRichDialogParent, _
            Optional ByVal IsModeless As Boolean, _
            Optional ByVal SubstituteDialog As iRichDialog _
                  ) As String
    Dim loInfo As cRichDialogInfo
    Set loInfo = New cRichDialogInfo
    With loInfo
        .Message = Message
        .Attributes = Attributes
        .Title = Title
        .hIcon = hIcon
        .hWndParent = hWndParent
        .Position = Position
        .Data = Buttons
        .Timeout = Timeout
        .Tag = Tag
        .ReturnValue = Default
        .IsModeless = IsModeless
        Set .Notify = Notify
        .InputBox = True
    End With
    InputBoxEx = mMsgBox.InputBoxEx(loInfo, SubstituteDialog)
    Set loInfo = Nothing
End Function

Public Property Get RichDialogGUIDefaults() As cRichDialogGUI
    Set RichDialogGUIDefaults = goDefaultGUI
End Property
