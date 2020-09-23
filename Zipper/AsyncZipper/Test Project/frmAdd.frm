VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAdd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add File(s)"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd 
      Caption         =   "Browse for Folder..."
      Height          =   375
      Index           =   3
      Left            =   3900
      TabIndex        =   3
      Top             =   120
      Width           =   1515
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Browse for File(s)..."
      Height          =   375
      Index           =   2
      Left            =   2280
      TabIndex        =   2
      Top             =   120
      Width           =   1515
   End
   Begin VB.TextBox txt 
      Enabled         =   0   'False
      Height          =   1125
      Index           =   3
      Left            =   5520
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   21
      Top             =   2640
      Width           =   2415
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Add"
      Enabled         =   0   'False
      Height          =   615
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton cmd 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CheckBox chk 
      Caption         =   "Set File Comment:"
      Height          =   255
      Index           =   8
      Left            =   5520
      TabIndex        =   20
      Top             =   2280
      Width           =   2295
   End
   Begin VB.CheckBox chk 
      Caption         =   "Include Directory entries"
      Height          =   255
      Index           =   7
      Left            =   5520
      TabIndex        =   19
      Top             =   2040
      Width           =   2295
   End
   Begin VB.CheckBox chk 
      Caption         =   "Set Zip file Time to latest file"
      Height          =   255
      Index           =   6
      Left            =   5520
      TabIndex        =   18
      Top             =   1800
      Width           =   2295
   End
   Begin VB.CheckBox chk 
      Caption         =   "Freshen: O/W only if newer"
      Height          =   255
      Index           =   5
      Left            =   5520
      TabIndex        =   17
      Top             =   1560
      Width           =   2295
   End
   Begin VB.CheckBox chk 
      Caption         =   "Ignore Hidden/System Files"
      Height          =   255
      Index           =   4
      Left            =   5520
      TabIndex        =   16
      Top             =   1320
      Width           =   2295
   End
   Begin VB.CheckBox chk 
      Caption         =   "Force 8.3 Filenames"
      Height          =   255
      Index           =   3
      Left            =   5520
      TabIndex        =   15
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CheckBox chk 
      Caption         =   "Convert LF to CRLF"
      Height          =   255
      Index           =   2
      Left            =   5520
      TabIndex        =   14
      Top             =   840
      Width           =   2295
   End
   Begin MSComCtl2.UpDown ud 
      Height          =   225
      Left            =   7560
      TabIndex        =   23
      Top             =   3930
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   397
      _Version        =   393216
      AutoBuddy       =   -1  'True
      BuddyControl    =   "txt(4)"
      BuddyDispid     =   196610
      BuddyIndex      =   4
      OrigLeft        =   3840
      OrigTop         =   3600
      OrigRight       =   4080
      OrigBottom      =   3855
      Max             =   9
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   4
      Left            =   7320
      Locked          =   -1  'True
      TabIndex        =   22
      Text            =   "0"
      Top             =   3900
      Width           =   495
   End
   Begin VB.CheckBox chk 
      Caption         =   "Convert CRLF to LF"
      Height          =   255
      Index           =   1
      Left            =   5520
      TabIndex        =   13
      Top             =   600
      Width           =   2295
   End
   Begin MSComCtl2.DTPicker dtp 
      Height          =   285
      Left            =   4320
      TabIndex        =   11
      Top             =   4080
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   19464193
      CurrentDate     =   38038
      MaxDate         =   73050
      MinDate         =   29221
   End
   Begin VB.OptionButton opt 
      Caption         =   "Include files earlier than:"
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   10
      Top             =   4200
      Width           =   2775
   End
   Begin VB.OptionButton opt 
      Caption         =   "Exclude files earlier than:"
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   9
      Top             =   3960
      Width           =   2775
   End
   Begin VB.OptionButton opt 
      Caption         =   "Include all files regardless of date"
      Height          =   255
      Index           =   0
      Left            =   2280
      TabIndex        =   8
      Top             =   3720
      Value           =   -1  'True
      Width           =   2775
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   2
      Left            =   2280
      TabIndex        =   7
      Top             =   3360
      Width           =   3135
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   1
      Left            =   2280
      TabIndex        =   6
      Top             =   2760
      Width           =   3135
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   0
      Left            =   2280
      TabIndex        =   4
      Top             =   840
      Width           =   3135
   End
   Begin VB.ListBox lst 
      Height          =   1035
      Left            =   2280
      MultiSelect     =   1  'Simple
      TabIndex        =   5
      Top             =   1440
      Width           =   3135
   End
   Begin VB.CheckBox chk 
      Caption         =   "Recurse to Sub-Directories"
      Height          =   255
      Index           =   0
      Left            =   5520
      TabIndex        =   12
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label lbl 
      Caption         =   "Current File Specs:  (wildcards allowed)"
      Height          =   255
      Index           =   5
      Left            =   2280
      TabIndex        =   29
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Label lbl 
      Caption         =   "Compression: (0=default)"
      Height          =   255
      Index           =   4
      Left            =   5520
      TabIndex        =   28
      Top             =   3900
      Width           =   1935
   End
   Begin VB.Label lbl 
      Caption         =   "Additional Options:"
      Height          =   375
      Index           =   3
      Left            =   5520
      TabIndex        =   27
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lbl 
      Caption         =   "Password:  (Blank for no encryption)"
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   26
      Top             =   3120
      Width           =   3135
   End
   Begin VB.Label lbl 
      Caption         =   "Base Path:  (Blank to discard folder names)"
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   25
      Top             =   2520
      Width           =   3135
   End
   Begin VB.Label lbl 
      Caption         =   "Type a path specification and press ENTER"
      Height          =   255
      Index           =   0
      Left            =   2280
      TabIndex        =   24
      Top             =   600
      Width           =   3135
   End
End
Attribute VB_Name = "frmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum eOpt
    optAll
    optExclude
    optInclude
End Enum

Private Enum eTxt
    txtFileSpec
    txtPath
    txtPass
    txtComment
    txtCompression
End Enum

Private Enum eChk
    chkRecurse
    chkCRLFtoLF
    chkLFtoCRLF
    chkDOSFilenames
    chkIgnoreHidden
    chkFreshen
    chkLatestTime
    chkIncludeDirectories
    chkComment
End Enum

Private Enum eCmd
    cmdCancel
    cmdAdd
    cmdAddFile
    cmdAddFolder
End Enum

Private moLB As cListBoxAPI

Private Sub chk_Click(Index As Integer)
    Dim lbVal
    lbVal = chk(Index).Value = vbChecked
    Select Case Index
        Case chkCRLFtoLF
            If lbVal Then chk(chkLFtoCRLF).Value = vbUnchecked
        Case chkLFtoCRLF
            If lbVal Then chk(chkCRLFtoLF).Value = vbUnchecked
        Case chkComment
            txt(txtComment).Enabled = lbVal
    End Select
End Sub

Private Sub DoAddFiles()
    Dim ltInfo As tZipInfo
    Dim I As Long
    With ltInfo
        I = lst.ListCount
        ReDim .FileSpecs(0 To I - 1)
        For I = 0 To I - 1
            .FileSpecs(I) = lst.List(I)
        Next
        .BasePath = txt(txtPath).Text
        .Password = txt(txtPass).Text
        If chk(chkComment).Value = vbChecked Then .Comment = txt(txtComment).Text
        .Compression = ud.Value
        If opt(optInclude).Value Then
            .Attributes = zipIncludeEarlierDates
            .DateMark = dtp.Value
        ElseIf opt(optExclude).Value Then
            .Attributes = zipExcludeEarlierDates
            .DateMark = dtp.Value
        End If
        If chk(chkRecurse).Value = vbChecked Then .Attributes = .Attributes Or zipRecurse
        If chk(chkCRLFtoLF).Value = vbChecked Then .Attributes = .Attributes Or zipCRLFtoLF
        If chk(chkLFtoCRLF).Value = vbChecked Then .Attributes = .Attributes Or zipLFtoCRLF
        If chk(chkDOSFilenames).Value = vbChecked Then .Attributes = .Attributes Or zipForceDOSFileNames
        If chk(chkIgnoreHidden).Value = vbChecked Then .Attributes = .Attributes Or zipIgnoreSystemAndHidden
        If chk(chkFreshen).Value = vbChecked Then .Attributes = .Attributes Or zipOnlyIfNewer
        If chk(chkLatestTime).Value = vbChecked Then .Attributes = .Attributes Or zipLatestTime
        If chk(chkIncludeDirectories).Value = vbChecked Then .Attributes = .Attributes Or zipIncludeDirectoryEntries
        frmZipper.ZipFiles ltInfo
    End With
End Sub

Private Sub cmd_Click(Index As Integer)
    Select Case Index
        Case cmdAdd
            Hide
            DoAddFiles
            Unload Me
        Case cmdCancel
            Unload Me
        Case cmdAddFolder
            Dim lsFile As String
            lsFile = BrowseForFolder(Hwnd, "Choose a folder to add.", vbNullString)
            If Len(lsFile) > 0 Then AddFileSpec lsFile
        Case cmdAddFile
            Dim loColl As Collection
            Set loColl = New Collection
            Dim lvTemp
            If GetOpenFileNames(loColl, Hwnd, "Add File(s)", vbNullString, "All Files (*.*)" & vbNullChar & "*.*" & vbNullChar, vbNullString, OFN_FILEMUSTEXIST + OFN_HIDEREADONLY + OFN_PATHMUSTEXIST + OFN_ALLOWMULTISELECT + OFN_EXPLORER) Then
                For Each lvTemp In loColl
                    moLB.AddItem CStr(lvTemp), , True
                Next
                cmd(cmdAdd).Enabled = True
            End If
    End Select
End Sub

Private Sub Form_Initialize()
    Set moLB = New cListBoxAPI
    moLB.Init lst
    dtp.Value = Date
    SetPathbreakProc txt(txtFileSpec).Hwnd
End Sub

Private Sub Form_Terminate()
    Set moLB = Nothing
End Sub

Private Sub lst_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        moLB.RemoveSelection
        cmd(cmdAdd).Enabled = lst.ListCount > 0
    End If
End Sub

Private Sub opt_Click(Index As Integer)
    dtp.Enabled = Index > optAll
End Sub

Private Function AddFileSpec(psSpec As String) As Boolean
    Dim lsFolder As String
    Dim lsFile As String
    Dim lsMsg As String
    If FolderExists(psSpec) Then psSpec = PathBuild(psSpec, "*")
    lsFile = Replace(Replace(PathGetFileName(psSpec), "*", ""), "?", "")
    lsFolder = PathGetParentFolder(psSpec)
    If FileNameIsLegal(lsFile) Then
        If FolderExists(lsFolder) Then
            If Not moLB.AddItem(psSpec, , True) Then lsMsg = "File spec already exists."
        Else
            lsMsg = "The folder is invalid."
        End If
    Else
        lsMsg = "The filename is invalid."
    End If
    If LenB(lsMsg) = 0 Then
        cmd(cmdAdd).Enabled = True
    Else
        MsgBox lsMsg, vbInformation, "Invalid File Spec."
    End If
End Function

Private Sub txt_Change(Index As Integer)
    If Index = txtFileSpec Then mAutoComplete.ac_Change txt(Index), txt(Index).Tag, acbFile
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = txtFileSpec Then
        txt(Index).Tag = KeyCode
        PathKeyDown
    End If
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = txtFileSpec Then
        If KeyAscii = vbKeyReturn Then
            If AddFileSpec(txt(txtFileSpec).Text) Then txt(txtFileSpec).Text = vbNullString
            KeyAscii = 0
        Else
            mAutoComplete.ac_KeyPress txt(Index), KeyAscii, acbFile
        End If
    End If
End Sub
