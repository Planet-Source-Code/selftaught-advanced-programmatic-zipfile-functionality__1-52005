VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmZipper 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   5250
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ListView lv 
      Height          =   2655
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   4683
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   0
   End
   Begin ComctlLib.Toolbar tb 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5250
      _ExtentX        =   9260
      _ExtentY        =   1270
      ButtonWidth     =   1852
      ButtonHeight    =   1164
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      ImageList       =   "il"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   4
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "New"
            Key             =   "new"
            Object.ToolTipText     =   "Create a New Archive"
            Object.Tag             =   ""
            ImageKey        =   "new"
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Open"
            Key             =   "open"
            Object.ToolTipText     =   "Open an Archive"
            Object.Tag             =   ""
            ImageKey        =   "open"
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Caption         =   "Add File(s)"
            Key             =   "add"
            Object.ToolTipText     =   "Add a File or Files to the Archive"
            Object.Tag             =   ""
            ImageKey        =   "add"
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Caption         =   "     Extract     "
            Key             =   "extract"
            Object.ToolTipText     =   "Extract File(s)"
            Object.Tag             =   ""
            ImageKey        =   "extract"
         EndProperty
      EndProperty
   End
   Begin ComctlLib.StatusBar sb 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   2940
      Width           =   5250
      _ExtentX        =   9260
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   "Ready"
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Working..."
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin ComctlLib.ImageList il 
      Left            =   0
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   5
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmZipper.frx":0000
            Key             =   "open"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmZipper.frx":0712
            Key             =   "addfolder"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmZipper.frx":0E24
            Key             =   "extract"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmZipper.frx":1536
            Key             =   "new"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmZipper.frx":1C48
            Key             =   "add"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmZipper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum eTB
    tbNew = 1
    tbOpen
    tbAdd
    tbExtract
End Enum

Private moFileLV As cFileListView
Private moZipFile As cZipFile

Private msFileName As String
Private mbLoaded As Boolean
Private mbWorking As Boolean
Private mbEncrypted As Boolean

Private moLog As cStringBuilder

Implements iZipCallBack

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long

Private Sub Form_Initialize()
    Set moZipFile = New cZipFile
    Set moFileLV = New cFileListView
    Set moLog = New cStringBuilder
    Me.ScaleMode = vbTwips
    With moFileLV
        .Attach lv
        .Columns = flvName Or flvType Or flvSize Or flvFolder
    End With
    Me.ScaleMode = vbPixels
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Dim liHeight As Long
    Dim liWidth As Long
    Dim liTBHeight As Long
    liTBHeight = tb.Height
    GetClientDimensions hwnd, liHeight, liWidth
    liHeight = liHeight - liTBHeight - sb.Height
    lv.Move 0, liTBHeight, liWidth, liHeight
    lbl.Move 0, liTBHeight, liWidth, liHeight
End Sub

Private Sub Form_Terminate()
    Set moZipFile = Nothing
    Set moLog = Nothing
End Sub

Private Sub LoadFile(ByVal psFile As String)
    lv.ListItems.Clear
    mbLoaded = False
    msFileName = vbNullString
    mbEncrypted = False
    If FileExists(psFile) Then
        mbWorking = True
        ShowControls
        msFileName = psFile
        moZipFile.ReadZipFile psFile, Me
    ElseIf FileIsValidToCreate(psFile, True) Then
        mbWorking = False
        mbLoaded = True
        ShowControls
        msFileName = psFile
    End If
End Sub

Private Sub ShowControls()
    sb.SimpleText = "Ready"
    tb.Enabled = Not mbWorking
    lv.Visible = Not mbWorking
    With tb.Buttons
        .Item(tbAdd).Enabled = mbLoaded
        .Item(tbExtract).Enabled = mbLoaded
    End With
End Sub

Private Sub iZipCallBack_OverwriteRequest(ByVal FileName As String, Answer As AsyncZipper.eUnzipOverwrite)
    Dim liAnswer As eRichDialogReturn
    liAnswer = MsgBoxEx("Do you want to overwrite this file?" & vbNewLine & vbNewLine & FileName, rdQuestion + rdDefaultButton3, "Overwrite File?", , hwnd, rdCenterCenter, Array("Yes", "Yes to All", "No", "No to All"))
    Select Case liAnswer
        Case rdButton1
            Answer = zipOverwriteThisFile
        Case rdButton2
            Answer = zipOverwriteAllFiles
        Case rdButton3
            Answer = zipDoNotOverwrite
        Case rdButton4
            Answer = zipOverwriteNone
    End Select
End Sub

Private Sub iZipCallBack_PasswordRequest(ByVal ForFile As String, Password As String, ByVal WasInvalid As Boolean, ApplyToAll As Boolean, Cancel As Boolean)
    Dim lsPassword As String
    Dim lsMsg As String
    On Error Resume Next
    If WasInvalid Then lsMsg = "Password was rejected." & vbNewLine & vbNewLine
    lsMsg = lsMsg & "Enter a password for this file." & vbNewLine & vbNewLine & ForFile
    lsPassword = InputBoxEx(lsMsg, rdCancelRaiseError, "Enter a password", Password, , hwnd, rdCenterCenter)
    If Err.Number = 0 Then
        Password = lsPassword
    Else
        Cancel = True
    End If
End Sub

Private Sub iZipCallBack_ReadComplete(ByVal ErrorCode As AsyncZipper.eUnzipErrorCodes, ByVal HasComment As Boolean, ByVal CompressedSize As Long, ByVal TotalSize As Long, ByVal NumMembers As Long, ByVal CompressionFactor As Long)
    mbWorking = False
    mbLoaded = ErrorCode = zipUErrNone And NumMembers > 0
    Debug.Assert mbLoaded
    ShowControls
End Sub

Private Sub iZipCallBack_ReadFile(ByVal FileName As String, ByVal Size As Long, ByVal CompressedSize As Long, ByVal CompressionFactor As Long, ByVal FileDate As Date, ByVal CRC As Long, ByVal Encrypted As Boolean)
    On Error Resume Next
    sb.SimpleText = "Reading File: " & FileName
    moFileLV.ShowFile FileName, 0, Size, FileDate
    'Debug.Assert InStr(1, FileName, "Create") = 0
    If Encrypted Then
        lv.ListItems(FileName).Text = lv.ListItems(FileName).Text & "+"
        mbEncrypted = True
    End If
End Sub

Private Sub iZipCallBack_UnzipComplete(ByVal ErrorCode As AsyncZipper.eUnzipErrorCodes)
    If ErrorCode > zipUErrNone Then
        If MsgBox("Unzip operation completed with errors.  Do you want to see the log?", vbYesNo + vbQuestion, "Unzip Error") = vbYes Then
            MsgBoxEx moLog.ToString, , "Unzip Log"
        End If
    End If
    mbWorking = False
    Debug.Assert ErrorCode = zipUErrNone
    ShowControls
End Sub

Private Sub iZipCallBack_UnZipMessage(ByVal Msg As String)
    moLog.Append Msg
    moLog.Append vbNewLine
End Sub

Private Sub iZipCallBack_UnzippedFile(ByVal FileName As String, Cancel As Boolean)
    sb.SimpleText = "Unzipped File: " & FileName
End Sub

Private Sub iZipCallBack_ZipComplete(ByVal ErrorCode As AsyncZipper.eZipErrorCodes)
    Debug.Assert ErrorCode = zipErrNone
    If ErrorCode > zipErrNone Then
        If MsgBox("Zip operation completed with errors.  Do you want to see the log?", vbYesNo + vbQuestion, "Zip Error") = vbYes Then
            MsgBoxEx moLog.ToString, , "Zip Log"
        End If
    End If
    LoadFile msFileName
End Sub

Private Sub iZipCallBack_ZipMessage(ByVal Msg As String)
    moLog.Append Msg
    moLog.Append vbNewLine
End Sub

Private Sub iZipCallBack_ZippedFile(ByVal FileName As String, Cancel As Boolean)
    sb.SimpleText = "Zipped File: " & FileName
End Sub

Private Function Filter(pbZipFiles As Boolean) As String
    If pbZipFiles Then Filter = CommonDialogFilter("Zip Files", "*.zip")
    Filter = Filter & CommonDialogFilter("All Files", "*.*")
End Function

Private Sub lv_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
    SortListView lv, ColumnHeader
End Sub

Private Sub lv_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If MsgBox("Delete the selected files from the zip?", vbYesNo + vbQuestion, "Delete Files?") = vbYes Then
            Dim ltZip As tZipInfo
            With ltZip
                .FileName = msFileName
                If GetSelFiles(.FileSpecs) Then
                    .Attributes = zipDeleteFileSpecs
                    mbWorking = True
                    ShowControls
                    moZipFile.ZipFiles ltZip, Me
                End If
            End With
        End If
    End If
End Sub

Private Sub tb_ButtonClick(ByVal Button As ComctlLib.Button)
    Dim lsName As String
    Dim loColl As Collection
    Dim lvTemp
    Dim liTemp As Long
    Dim ltZip As tZipInfo
    Dim ltUnzip As tUnzipInfo
    Select Case Button.Index
        Case tbNew
            lsName = GetSaveFileName(hwnd, "Enter a new zip filename", "", Filter(True), "zip", OFN_PATHMUSTEXIST + OFN_HIDEREADONLY + OFN_EXPLORER)
            If Len(lsName) > 0 Then
                If FileExists(lsName) Then
                    If MsgBox("This file already exists.  Do you want to delete it?", vbYesNo + vbQuestion, "Delete the file?") = vbYes Then Kill lsName Else Exit Sub
                End If
                LoadFile lsName
            End If
        Case tbOpen
            Set loColl = New Collection
            If GetOpenFileNames(loColl, hwnd, "Choose Files to Add", "", Filter(False), "", OFN_EXPLORER + OFN_FILEMUSTEXIST + OFN_HIDEREADONLY) Then LoadFile (loColl(1))
        Case tbAdd
            frmAdd.Show vbModal, Me
        Case tbExtract
            frmExtract.Show vbModal, Me
    End Select
End Sub

Public Sub UnzipFiles(ptInfo As tUnzipInfo)
    ptInfo.FileName = msFileName
    moLog.TheString = ""
    If moZipFile.UnzipFiles(ptInfo, Me) Then
        mbWorking = True
        ShowControls
    Else
        Debug.Assert False
        MsgBox "Could not start Unzipping the files.", vbInformation, "Zip Error"
    End If
End Sub

Public Sub ZipFiles(ptInfo As tZipInfo)
    ptInfo.FileName = msFileName
    moLog.TheString = ""
    If moZipFile.ZipFiles(ptInfo, Me) Then
        mbWorking = True
        ShowControls
    Else
        Debug.Assert False
        MsgBox "Could not start zipping the files.", vbInformation, "Zip Error"
    End If
End Sub

Public Function GetSelFiles(psString() As String) As Boolean
    On Error Resume Next
    Erase psString
    
    Dim liCount As Long
    Dim loLI As ListItem
    
    liCount = ListCount
    If liCount = 0 Then Exit Function
    GetSelFiles = True
    ReDim psString(0 To liCount - 1)
    
    liCount = 0
    For Each loLI In lv.ListItems
        If loLI.Selected Then
            psString(liCount) = loLI.Key
            liCount = liCount + 1
        End If
    Next
End Function

Public Property Get ListCount(Optional pbSelOnly As Boolean = True) As Long
Const LVM_FIRST = &H1000
Const LVM_GETITEMCOUNT = (LVM_FIRST + 4)
Const LVM_GETSELECTEDCOUNT = (LVM_FIRST + 50)

    If pbSelOnly Then
        ListCount = SendMessage(lv.hwnd, LVM_GETSELECTEDCOUNT, 0&, 0&)
    Else
        ListCount = SendMessage(lv.hwnd, LVM_GETITEMCOUNT, 0&, 0&)
    End If
End Property

Public Property Get Encrypted() As Boolean
    Encrypted = mbEncrypted
End Property
