VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFileListView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Requires ListView Common Control _5.0_

Private Const MAX_PATH = 260

Private Type tFileTypeInfo
    Ext As String
    TypeName As String
    Icon As Long
End Type

Public Enum eFileLVColumns
    flvName = 1
    flvFolder = 2
    flvSize = 4
    flvType = 8
    flvModified = 16
    flvAccessed = 32
    flvCreated = 64
    flvAttributes = 128
End Enum

Private Enum eShellGetFileInfoFlags
    SHGFI_ATTRIBUTES = &H800                   '  get attributes
    SHGFI_DISPLAYNAME = &H200                  '  get display name
    SHGFI_EXETYPE = &H2000                     '  return exe type
    SHGFI_ICON = &H100                         '  get icon
    SHGFI_ICONLOCATION = &H1000                '  get icon location
    SHGFI_LARGEICON = &H0                      '  get large icon
    SHGFI_LINKOVERLAY = &H8000                 '  put a link overlay on icon
    SHGFI_OPENICON = &H2                       '  get open icon
    SHGFI_PIDL = &H8                           '  pszPath is a pidl
    SHGFI_SELECTED = &H10000                   '  show icon in selected state
    SHGFI_SHELLICONSIZE = &H4                  '  get shell size icon
    SHGFI_SMALLICON = &H1                      '  get small icon
    SHGFI_SYSICONINDEX = &H4000                '  get system icon index
    SHGFI_TYPENAME = &H400                     '  get type name
    SHGFI_USEFILEATTRIBUTES = &H10             '  use passed dwFileAttribute
End Enum

Private Type LV_ITEM
    mask As Long
    iItem As Long
    iSubItem As Long
    State As Long
    stateMask As Long
    pszText As String
    cchTextMax As Long
    iImage As Long
    lParam As Long
    iIndent As Long
End Type

Private Enum eRandomStuff
    LVS_SHAREIMAGELISTS = &H40&
    GWL_STYLE = (-16)

    LVM_FIRST = &H1000&
    LVM_SETIMAGELIST = (LVM_FIRST + 3)

    LVSIL_NORMAL = 0
    LVSIL_SMALL = 1

    LVIF_IMAGE = &H2

    LVM_SETITEM = (LVM_FIRST + 6)
End Enum

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As eShellGetFileInfoFlags) As Long


Private Type SHFILEINFO
    hIcon As Long                      '  out: icon
    iIcon As Long          '  out: icon index
    dwAttributes As Long               '  out: SFGAO_ flags
    szDisplayName As String * MAX_PATH '  out: display name (or path)
    szTypeName As String * 80          '  out: type name
End Type

Private Const ColName = "Name"
Private Const ColModified = "Modified"
Private Const ColAccessed = "Accessed"
Private Const ColCreated = "Created"
Private Const ColAttributes = "Attributes"
Private Const ColType = "Type"
Private Const ColSize = "Size"
Private Const ColFolder = "In Folder"

Private Const DateFormat = "M/D/YY H:MM AMPM"

Private moLV As ListView
Private miHwnd As Long
Private mtTypeInfo() As tFileTypeInfo
Private miCols As eFileLVColumns
Private msTempDir As String

Public Sub Attach(poLV As ListView)
    On Error Resume Next
    miHwnd = 0
    Set moLV = poLV
    miHwnd = poLV.hwnd
    SetWindowLong miHwnd, GWL_STYLE, GetWindowLong(miHwnd, GWL_STYLE) Or LVS_SHAREIMAGELISTS
    AssignSysIL
End Sub

Public Sub AssignSysIL()
    Dim hSysIL(0 To 1) As Long 'sys IL handles
    Dim L As Long
    
    For L = 0 To 1
        FileGetIcon App.Path, hSysIL(L), 0, L
        SendMessage miHwnd, LVM_SETIMAGELIST, L, ByVal hSysIL(L)
    Next
End Sub

Private Function GetTypeInfo(psExt As String, psType As String, piIcon As Long) As Boolean
    Dim i As Long
    On Error GoTo ending
    For i = 0 To UBound(mtTypeInfo)
        If StrComp(psExt, mtTypeInfo(i).Ext, vbTextCompare) = 0 Then
            If Len(psType) = 0 Then psType = mtTypeInfo(i).TypeName
            piIcon = mtTypeInfo(i).Icon
            GetTypeInfo = True
            Exit Function
        End If
    Next
ending:
End Function

Private Function AddTypeInfo(psExt As String, psType As String, ByVal piIcon As Long)
    Dim liVal As Long
    On Error Resume Next
    liVal = UBound(mtTypeInfo)
    If Err.Number = 0 Then
        liVal = liVal + 1
        ReDim Preserve mtTypeInfo(0 To liVal)
    Else
        ReDim mtTypeInfo(0 To 0)
    End If
    With mtTypeInfo(liVal)
        .Icon = piIcon
        .Ext = psExt
        .TypeName = psType
    End With
End Function

Public Sub ShowFile(psPath As String, Optional piAttributes As eFileAttributes = -1, Optional ByVal piSize As Long = -1, Optional pdModified As Date, Optional pdAccessed As Date, Optional pdCreated As Date, Optional psType As String)
    If LenB(psPath) > 0 Then
        On Error Resume Next
        
        Dim lbKill As Boolean
        Dim liTemp As Long
        Dim liIcon As Long
        Dim loItem As ListItem
        Dim lvi As LV_ITEM
        
        Dim ltGetInfo As SHFILEINFO
        Dim lsName As String
        
        Dim liFlags As eShellGetFileInfoFlags
        Dim lbGotTypeInfo As Boolean
        Dim lsExt As String
        Dim lsTempPath As String
        
        
        If ((miCols And flvAccessed) And pdAccessed = #12:00:00 AM#) Or _
           ((miCols And flvModified) And pdModified = #12:00:00 AM#) Or _
           ((miCols And flvCreated) And pdCreated = #12:00:00 AM#) Then
            FileGetTime psPath, pdModified, pdCreated, pdAccessed
        End If
        
        If miCols And flvSize And piSize = -1 Then piSize = FileGetLen(psPath)
    
        lsExt = PathGetExtension(psPath)
        lbGotTypeInfo = GetTypeInfo(lsExt, psType, liIcon)
        If Not lbGotTypeInfo Then
            liFlags = liFlags Or SHGFI_SYSICONINDEX 'SHGFI_ICON
            If miCols And flvType Then
                If LenB(psType) = 0 Then liFlags = liFlags Or SHGFI_TYPENAME
            End If
        End If
        'If moLV.View = lvwIcon Then liFlags = liFlags Or SHGFI_LARGEICON Else liFlags = liFlags Or SHGFI_SMALLICON
        'If PathGetExtension(psPath) = "lnk" Then liFlags = liFlags Or SHGFI_LINKOVERLAY
    
        If miCols And flvAttributes Then
            If piAttributes = -1 Then liFlags = liFlags Or SHGFI_ATTRIBUTES
        End If
        'If miCols And flvName Then liFlags = liFlags Or SHGFI_DISPLAYNAME
        

        If FileExists(psPath) Then
            lsTempPath = psPath
        Else
            If Len(msTempDir) = 0 Then
                msTempDir = PathGetTempFolderName
                PathCreate msTempDir
            End If
            lsTempPath = FileGetTempName(msTempDir, , lsExt)
            liTemp = FreeFile
            Open lsTempPath For Binary As #liTemp
            Close #liTemp
            lbKill = True
        End If
        If liFlags <> 0 Then SHGetFileInfo lsTempPath, 0, ltGetInfo, Len(ltGetInfo), liFlags
        If lbKill Then FileDelete lsTempPath, False
        With ltGetInfo
            If liFlags And SHGFI_SYSICONINDEX Then liIcon = .iIcon
            If liFlags And SHGFI_DISPLAYNAME Then
                lsName = .szDisplayName
                StripNulls lsName
            Else
                lsName = PathGetFileName(psPath)
            End If
            If liFlags And SHGFI_TYPENAME Then
                psType = .szTypeName
                StripNulls psType
            End If
            If liFlags And SHGFI_ATTRIBUTES Then piAttributes = .dwAttributes
        End With
        If (miCols And flvType > 0) And Len(psType) = 0 Then psType = StrConv(lsExt, vbUpperCase) & " File"
        If Not lbGotTypeInfo Then AddTypeInfo lsExt, psType, liIcon
        
        Set loItem = moLV.ListItems(psPath)
        If loItem Is Nothing Then Set loItem = moLV.ListItems.Add(, psPath, lsName)
        
        With loItem
            Dim i As Long
            Dim lsText As String
            For i = 2 To moLV.ColumnHeaders.Count
                Select Case moLV.ColumnHeaders(i).Text
                    Case ColName
                        lsText = lsName
                    Case ColModified
                        lsText = Format(pdModified, DateFormat)
                    Case ColAccessed
                        lsText = Format(pdAccessed, DateFormat)
                    Case ColCreated
                        lsText = Format(pdCreated, DateFormat)
                    Case ColAttributes
                        If piAttributes = -1 Then piAttributes = FileGetAttributes(psPath)
                        lsText = ""
                        If piAttributes And FILE_ATTRIBUTE_READONLY Then lsText = lsText & "R"
                        If piAttributes And FILE_ATTRIBUTE_HIDDEN Then lsText = lsText & "H"
                        If piAttributes And FILE_ATTRIBUTE_SYSTEM Then lsText = lsText & "S"
                        If piAttributes And FILE_ATTRIBUTE_ARCHIVE Then lsText = lsText & "A"
                    Case ColType
                        If LenB(psType) = 0 Then FileGetTypeName psPath, psType
                        lsText = psType
                    Case ColSize
                        If piSize = -1 Then piSize = FileGetLen(psPath)
                        lsText = "#,###,###,##0.00 KB"
                        'If Not piSize Mod KB = 0 Then piSize = piSize + 1
                        lsText = Format(piSize / KB, lsText)
                    Case ColFolder
                        lsText = PathGetParentFolder(psPath)
                    Case Else
                        lsText = ""
                End Select
                .SubItems(i - 1) = lsText
            Next
           
            With lvi
                .iItem = loItem.Index - 1
                .iImage = liIcon
                .mask = LVIF_IMAGE
            End With
            SendMessage miHwnd, LVM_SETITEM, 0&, lvi
           
        End With
    End If
End Sub

Public Property Let Columns(piVal As eFileLVColumns)
    miCols = piVal
    With moLV.ColumnHeaders
        .Clear
        If miCols And flvName Then .Add , ColName, ColName, 1800
        If miCols And flvFolder Then .Add , ColFolder, ColFolder, 2000
        If miCols And flvSize Then .Add(, ColSize, ColSize, 800).Alignment = lvwColumnRight
        If miCols And flvType Then .Add , ColType, ColType, 1800
        If miCols And flvModified Then .Add , ColModified, ColModified, 1400
        If miCols And flvAccessed Then .Add , ColAccessed, ColAccessed, 1400
        If miCols And flvCreated Then .Add , ColCreated, ColCreated, 1400
        If miCols And flvAttributes Then .Add(, ColAttributes, ColAttributes, 600).Alignment = lvwColumnRight
    End With
End Property

Public Sub SyncColumns()
    miCols = 0
    On Error Resume Next
    Dim loCol As ColumnHeader
    With moLV.ColumnHeaders
        If Not Column(ColName) Is Nothing Then miCols = miCols Or flvName
        If Not Column(ColFolder) Is Nothing Then miCols = miCols Or flvFolder
        If Not Column(ColSize) Is Nothing Then miCols = miCols Or flvSize
        If Not Column(ColType) Is Nothing Then miCols = miCols Or flvType
        If Not Column(ColModified) Is Nothing Then miCols = miCols Or flvModified
        If Not Column(ColAccessed) Is Nothing Then miCols = miCols Or flvAccessed
        If Not Column(ColCreated) Is Nothing Then miCols = miCols Or flvCreated
        If Not Column(ColAttributes) Is Nothing Then miCols = miCols Or flvAttributes
    End With
End Sub

Private Property Get Column(psKey As String) As ColumnHeader
    On Error Resume Next
    Set Column = moLV.ColumnHeaders.Item(psKey)
End Property

Private Sub Class_Terminate()
    On Error Resume Next
    Set moLV = Nothing
    If Len(msTempDir) > 0 Then FolderDelete msTempDir
End Sub
