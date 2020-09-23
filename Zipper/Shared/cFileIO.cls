VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFileIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private miChunksize As Long
Private miSize As Long
Private miPointer As Long
Private miAccess As eFileAccess
Private miShare As eFileShare
Private miCreation As eFileCreation
Private miAttributes As eFileAttributes
Private miFlags As eFileFlags

Private hFile As Long

Public Property Let FileAccess(piVal As eFileAccess)
    If hFile = INVALID_HANDLE_VALUE Then miAccess = piVal
End Property
Public Property Get FileAccess() As eFileAccess
    FileAccess = miAccess
End Property
Public Property Let FileShare(piVal As eFileShare)
    If hFile = INVALID_HANDLE_VALUE Then miShare = piVal
End Property
Public Property Get FileShare() As eFileShare
    FileShare = miShare
End Property
Public Property Let FileCreation(piVal As eFileCreation)
    If hFile = INVALID_HANDLE_VALUE Then miCreation = piVal
End Property
Public Property Get FileCreation() As eFileCreation
    FileCreation = miCreation
End Property
Public Property Let FileAttributes(piVal As eFileAttributes)
    If hFile = INVALID_HANDLE_VALUE Then miAttributes = piVal
End Property
Public Property Get FileAttributes() As eFileAttributes
    FileAttributes = miAttributes
End Property
Public Property Let FileFlags(piVal As eFileFlags)
    If hFile = INVALID_HANDLE_VALUE Then miFlags = piVal
End Property
Public Property Get FileFlags() As eFileFlags
    FileFlags = miFlags
End Property

Public Property Get EOF() As Boolean
    EOF = miPointer >= miSize Or miPointer < 0
End Property

Public Property Get ChunkSize() As Long
    ChunkSize = miChunksize
End Property
Public Property Let ChunkSize(ByVal piVal As Long)
    If piVal < 32768 Then piVal = 32768
    miChunksize = piVal
End Property

Public Property Get FilePos() As Long
    FilePos = miPointer
End Property
Public Property Let FilePos(ByVal NewPosition As Long)
    If hFile <> INVALID_HANDLE_VALUE Then FileSetPointer hFile, NewPosition, FILE_BEGIN, miPointer
End Property

Public Property Get FileLength() As Long
    FileLength = miSize
End Property

Public Function OpenFile(FileName As String) As Boolean
    On Error Resume Next
    If hFile <> INVALID_HANDLE_VALUE Then CloseFile
    If Not PathCreate(PathGetParentFolder(FileName)) Then Exit Function
    OpenFile = FileCreate(FileName, miAccess, miShare, miCreation, miFlags, miAttributes, hFile)
    miSize = FileGetSize(hFile)
    miPointer = 0
End Function

Public Function CloseFile() As Boolean
    If hFile > INVALID_HANDLE_VALUE Then CloseFile = FileClose(hFile)
End Function

Public Function GetLong(piLong As Long) As Boolean
    If hFile = INVALID_HANDLE_VALUE Then Exit Function
    Dim liBytes As Long
    GetLong = FileRead_l(hFile, piLong, liBytes)
    miPointer = miPointer + liBytes
End Function
Public Function AppendLong(ByVal piLong As Long) As Boolean
    If hFile = INVALID_HANDLE_VALUE Then Exit Function
    Dim liBytes As Long
    AppendLong = FileWrite_l(hFile, piLong, liBytes)
    miPointer = miPointer + liBytes
End Function

Public Function GetInteger(piInt As Long) As Boolean
    If hFile = INVALID_HANDLE_VALUE Then Exit Function
    Dim liInt As Integer
    Dim liBytes As Long
    GetInteger = FileRead_i(hFile, liInt, liBytes)
    miPointer = miPointer + liBytes
    piInt = liInt
End Function
Public Function AppendInteger(ByVal piInt As Integer) As Boolean
    If hFile = INVALID_HANDLE_VALUE Then Exit Function
    Dim liBytes As Long
    AppendInteger = FileWrite_i(hFile, piInt, liBytes)
    miPointer = miPointer + liBytes
End Function

Public Function GetString(psString As String) As Boolean
    If hFile = INVALID_HANDLE_VALUE Then Exit Function
    Dim liBytes As Long
    Dim liLen As Integer
    FileRead_i hFile, liLen, liBytes
    miPointer = miPointer + liBytes
    liBytes = liLen
    GetString = FileRead_s(hFile, psString, liBytes)
    miPointer = miPointer + liLen
End Function
Public Function AppendString(psString As String) As Boolean
    If hFile = INVALID_HANDLE_VALUE Then Exit Function
    Dim liBytes As Long
    FileWrite_i hFile, Len(psString), liBytes
    miPointer = miPointer + liBytes
    liBytes = Len(psString)
    AppendString = FileWrite_s(hFile, psString, liBytes)
    miPointer = miPointer + liBytes
End Function

Public Function GetBytes(pyBytes() As Byte, Optional ByVal piChunksize As Long) As Boolean
    If hFile = INVALID_HANDLE_VALUE Then Exit Function
    Select Case piChunksize
        Case Is < 0
            piChunksize = 0
            Debug.Assert False
        Case 0
            piChunksize = miChunksize
    End Select
    GetBytes = FileRead_y(hFile, pyBytes, piChunksize)
    miPointer = miPointer + piChunksize
End Function

Public Function AppendBytes(pyBytes() As Byte, Optional ByVal NumBytes As Long) As Boolean
    If hFile = INVALID_HANDLE_VALUE Then Exit Function
    AppendBytes = FileWrite_y(hFile, pyBytes, NumBytes)
    miPointer = miPointer + NumBytes
End Function

Public Function IsFileOpen() As Boolean
    IsFileOpen = hFile <> INVALID_HANDLE_VALUE
End Function

Private Sub Class_Initialize()
    miAccess = GENERIC_READ
    miShare = FILE_SHARE_READ
    miCreation = OPEN_EXISTING
    'miAttributes=
    miFlags = FILE_FLAG_SEQUENTIAL_SCAN
    hFile = INVALID_HANDLE_VALUE
    'miChunksize = 32768
    'miChunksize = 65536
    'miChunksize = 524288
    miChunksize = 1048576
    'miChunksize = 4194304
    'miChunksize = 16777216
End Sub


Private Sub Class_Terminate()
    CloseFile
End Sub

Public Function GetDouble(pDbl As Double) As Boolean
    If hFile = INVALID_HANDLE_VALUE Then Exit Function
    Dim liBytes As Long
    GetDouble = FileRead_d(hFile, pDbl, liBytes)
    miPointer = miPointer + liBytes
End Function
Public Function AppendDouble(ByVal pDbl As Double) As Boolean
    If hFile = INVALID_HANDLE_VALUE Then Exit Function
    Dim liBytes As Long
    AppendDouble = FileWrite_d(hFile, pDbl, liBytes)
    miPointer = miPointer + liBytes
End Function

