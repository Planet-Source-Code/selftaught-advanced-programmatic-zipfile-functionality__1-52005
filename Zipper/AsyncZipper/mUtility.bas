Attribute VB_Name = "mUtility"
Option Explicit

'Public Const errEntryPointNotFound = 453
'Public Const errFileNotFound = 53

'Shared structures between zip and Unzip
Public Type ZipNames
    s(0 To 1023) As String
End Type
Public Type CBCh
    ch(0 To 255) As Byte
End Type

'These are the reasons that we need to callback due to
'VB's lack of direct support for async method calls
Public Enum eCallBackReason
    zipZip
    zipUnzip
    zipRead
End Enum


'Stuff for filesystem API calls
Private Const MAX_PATH = 260&
Private Enum ePathCharTypes
    PCT_INVALID = 0
    PCT_LFNCHAR = 1
    PCT_SHORTCHAR = 2
    PCT_WILD = 4
    PCT_SEPARATOR = 8
End Enum
Private Enum eDriveType
    DRIVE_UNKNOWN
    DRIVE_ABSENT
    DRIVE_REMOVABLE
    DRIVE_FIXED
    DRIVE_REMOTE
    DRIVE_CDROM
    DRIVE_RAMDISK
End Enum
Private Declare Function PathIsRootAPI Lib "shlwapi.dll" Alias "PathIsRootA" (ByVal pszPath As String) As Long
Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long
Private Declare Function PathGetCharTypeAPI Lib "shlwapi.dll" Alias "PathGetCharTypeA" (ByVal ch As Byte) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As eDriveType
Private Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Declare Sub PathStripPath Lib "shlwapi.dll" Alias "PathStripPathA" (ByVal pszPath As String)

'General-purpose
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

'For async method calls
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long

Private moColl       As Collection
Private moCollReason As Collection
Private moCollInfo   As Collection
Private miTimer      As Long

'workaround for AddressOf operator
Public Function AddrFunc(ByVal piPtr As Long) As Long
    AddrFunc = piPtr
End Function

Public Sub Callback( _
               ByVal ForMe As iZipCallBack, _
               ByVal Reason As eCallBackReason, _
               ByRef Info As Variant _
           )
    
    If moColl Is Nothing Then
        'Initialize collections
        Set moColl = New Collection
        Set moCollReason = New Collection
        Set moCollInfo = New Collection
    End If
    
    'store callback info
    moColl.Add ForMe
    moCollReason.Add Reason
    moCollInfo.Add Info
    
    'It is possible to have calls so close together to get here before the timer
    'returns, which is why we're using collections instead of a modular variable

    If miTimer = 0 Then miTimer = SetTimer(0, 0, 1, AddressOf TimerProc)
End Sub

Private Sub TimerProc( _
                ByVal hWnd As Long, _
                ByVal nIDEvent As Long, _
                ByVal uElapse As Long, _
                ByVal lpTimerFunc As Long _
            )
    
    On Error Resume Next
    
    'Don't want the timer calling us again for the same callback!
    KillTimer 0, miTimer
    miTimer = 0
    
    Dim iCB As iZipCallBack
    Dim liReason As eCallBackReason
    Dim Info As Variant
    
    Dim i As Long
    For i = 1 To moColl.Count
        
        'Extract the specs for this callback
        Set iCB = moColl(1)
        liReason = moCollReason(1)
        Info = moCollInfo(1)
        
        'remove the specs for this callback
        moColl.Remove 1
        moCollReason.Remove 1
        moCollInfo.Remove 1
        
        'decide what exactly to do
        Select Case liReason
            Case zipZip
                Dim ltZipInfo As tZipInfo
                ltZipInfo = Info
                mZipper.ZipFiles ltZipInfo, iCB
            Case zipUnzip
                Dim ltUnzipInfo As tUnzipInfo
                ltUnzipInfo = Info
                mUnzipper.UnzipFiles ltUnzipInfo, iCB
            Case zipRead
                mUnzipper.ReadZipFile CStr(Info), iCB
        End Select
        Set iCB = Nothing
    Next
    
End Sub

'Gets the string out of a byte array, when the byte array may be larger than
'necessary.  Assumes that either a length is passed or the string is terminated by vbNullChar
Public Function GetString( _
                    ByRef pyBytes() As Byte, _
           Optional ByVal piPlace As Long = -1 _
                ) As String
    
    Dim lyTemp() As Byte
    
    If piPlace < 0 Then
        'Find the first null char
        For piPlace = 0 To UBound(pyBytes)
            If pyBytes(piPlace) = 0 Then Exit For
        Next
    Else
        'make sure that we are in bounds
        If piPlace > UBound(pyBytes) + 1 Then piPlace = UBound(pyBytes) + 1
    End If
    
    'extract the string from the first piPlace chars of the byte array
    ReDim lyTemp(0 To piPlace - 1)
    CopyMemory lyTemp(0), pyBytes(0), piPlace
    GetString = StrConv(lyTemp, vbUnicode)
End Function

'Puts a regular string array into the zip structure format
Public Function TranslateStringArray( _
                    ByRef psStrings() As String, _
                    ByRef ZipNames() As String, _
                    ByVal pbFlopBackslash As Boolean _
                ) As Long
    
    Dim i As Long
    Dim liUbound As Long
    
    On Error GoTo NoStrings
    'If the array is not dimensioned, then there's nothing to do!
    liUbound = UBound(psStrings)
    On Error Resume Next
    
    i = UBound(ZipNames) - 1
    'make sure that we're in bounds
    If liUbound > i Then liUbound = i
    
    'instead of testing the flag during every iteration of the loop
    If Not pbFlopBackslash Then
        For i = 0 To liUbound
            ZipNames(i) = psStrings(i)
        Next
    Else
        For i = 0 To liUbound
            ZipNames(i) = Replace$(psStrings(i), "\", "/")
        Next
    End If
    ZipNames(i) = vbNullChar
    'return the number of strings
    TranslateStringArray = i
    Exit Function

NoStrings:
    On Error Resume Next
    ZipNames(0) = vbNullChar
End Function

'Puts a string into a byte array in the zip format
Public Sub TranslateString( _
               ByRef psString As String, _
               ByRef pyBytes() As Byte, _
               ByVal piMax As Long)
    
    On Error Resume Next
    
    Dim i As Long
    Dim lyTemp() As Byte
    
    lyTemp = StrConv(psString, vbFromUnicode)
    i = UBound(lyTemp) + 1
    'make sure we're in bounds
    If i > piMax Then i = piMax
    CopyMemory pyBytes(0), lyTemp(0), i
    'add a trailing null char
    pyBytes(i) = 0
End Sub

Public Sub TrimMsg(psMsg As String)
    psMsg = Replace(RTrim$(Replace$(psMsg, vbLf, " ")), "/", "\")
End Sub


'General purpose filesystem functions
Public Function FileExists(psPath As String) As Boolean
    FileExists = PathFileExists(psPath) <> 0
    If FileExists Then FileExists = Not FolderExists(psPath)
    If FileExists Then FileExists = Not GetDriveType(psPath) > DRIVE_ABSENT
End Function

Public Function FolderExists(psPath As String) As Boolean
    FolderExists = PathIsDirectory(psPath) <> 0
End Function

Public Function PathCreate(ByVal psBottomFolder As String) As Boolean
    PathAddBackslash psBottomFolder
    If PathIsRoot(psBottomFolder) Then
        If GetDriveType(psBottomFolder) = DRIVE_ABSENT Then Exit Function
    End If
    PathCreate = MakeSureDirectoryPathExists(psBottomFolder) <> 0
End Function

Public Function FileIsValidToCreate(psPath As String, Optional pbWithoutCreatingPath As Boolean) As Boolean
    If Not pbWithoutCreatingPath Then
        If Not PathCreate(PathGetParentFolder(psPath)) Then Exit Function
    Else
        If Not FolderExists(PathGetParentFolder(psPath)) Then Exit Function
    End If
    If FolderExists(psPath) Then Exit Function
    FileIsValidToCreate = FileNameIsLegal(PathGetFileName(psPath))
End Function

Private Function FileNameIsLegal(psName As String) As Boolean
    On Error GoTo ending
    Dim lyBytes() As Byte
    Dim i As Long
    lyBytes = StrConv(PathGetFileName(psName), vbFromUnicode)
    For i = LBound(lyBytes) To UBound(lyBytes)
        If Not PathGetCharType(lyBytes(i)) And PCT_LFNCHAR Then Exit Function
    Next
    FileNameIsLegal = True
ending:
End Function

Private Function PathGetCharType(ByVal pyChar As Byte) As ePathCharTypes
    PathGetCharType = PathGetCharTypeAPI(pyChar)
End Function

Private Function PathGetParentFolder(psPath As String) As String
    PathGetParentFolder = PathGetFileName(psPath)
    PathGetParentFolder = Left$(psPath, Len(psPath) - Len(PathGetParentFolder))
End Function

Private Function PathGetFileName(psPath As String) As String
    PrepareString PathGetFileName, psPath
    PathStripPath PathGetFileName
    StripNulls PathGetFileName
End Function

Private Sub StripNulls(psString As String)
    Dim liPos As Long
    liPos = InStr(1, psString, Chr$(0))
    If liPos > 0 Then psString = Left$(psString, liPos - 1)
End Sub

Private Sub PrepareString(psString As String, Optional psValue As String)
    psString = String$(MAX_PATH, 0)
    Mid$(psString, 1, Len(psValue)) = psValue
End Sub

Private Function PathIsRoot(psPath As String) As Boolean
    PathIsRoot = PathIsRootAPI(psPath) <> 0
End Function

Private Sub PathAddBackslash(psPath As String)
    If StrComp(Right$(psPath, 1), "\") <> 0 And LenB(psPath) > 0 Then psPath = psPath & "\"
End Sub

