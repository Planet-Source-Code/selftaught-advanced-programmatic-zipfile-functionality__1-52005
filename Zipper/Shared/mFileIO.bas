Attribute VB_Name = "mFileIO"
Option Explicit

Public Enum eFileAccess
    GENERIC_READ = &H80000000
    GENERIC_WRITE = &H40000000
End Enum

Public Enum eFileShare
    FILE_SHARE_READ = &H1
    FILE_SHARE_WRITE = &H2
End Enum

Public Enum eFileCreation
    CREATE_ALWAYS = 2
    CREATE_NEW = 1
    OPEN_EXISTING = 3
    OPEN_ALWAYS = 4
    TRUNCATE_EXISTING = 5
End Enum
   
Public Enum eFileFlags
    FILE_FLAG_DELETE_ON_CLOSE = &H4000000
    FILE_FLAG_SEQUENTIAL_SCAN = &H8000000
    FILE_FLAG_RANDOM_ACCESS = &H10000000
End Enum

Public Enum eFilePos
    FILE_BEGIN = 0
    FILE_CURRENT = 1
    FILE_END = 2
End Enum

Public Const INVALID_HANDLE_VALUE = -1

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Function FileCreate(FilePath As String, ByVal AccessType As eFileAccess, ByVal ShareMode As eFileShare, ByVal CreateMode As eFileCreation, ByVal FileFlags As eFileFlags, ByVal FileAttributes As eFileAttributes, OutHandle As Long) As Boolean
    OutHandle = CreateFile(FilePath, AccessType, ShareMode, ByVal 0&, CreateMode, FileFlags Or FileAttributes, 0)
    FileCreate = OutHandle <> INVALID_HANDLE_VALUE
End Function

Public Function FileGetSize(ByVal hFile As Long) As Double
    Dim liVal As Long
    Dim liHigh As Long
    liVal = GetFileSize(hFile, liHigh)
    FileGetSize = MakeQWord(liVal, liHigh)
End Function

Public Function FileSetPointer(ByVal hFile As Long, ByVal Increment As Long, Optional ByVal IncFrom As eFilePos = FILE_CURRENT, Optional OutPosition As Long) As Boolean
    OutPosition = SetFilePointer(hFile, Increment, 0, IncFrom)
    FileSetPointer = OutPosition <> INVALID_HANDLE_VALUE
End Function

Public Function FileGetPointer(ByVal hFile As Long, OutPointer As Long) As Boolean
    FileGetPointer = FileSetPointer(hFile, 0, FILE_CURRENT, OutPointer)
End Function

Public Function FileRead_y(ByVal hFile As Long, pyBytes() As Byte, ByteCount As Long) As Boolean
    Dim lyBytes() As Byte
    If ByteCount < 1 Then Exit Function
    ReDim lyBytes(0 To ByteCount - 1)
    FileRead_y = ReadFile(hFile, lyBytes(0), ByteCount, ByteCount, ByVal 0&) <> 0
    If FileRead_y Then FileRead_y = ByteCount > 0
    If FileRead_y Then
        If ByteCount - 1 < UBound(lyBytes) Then
            ReDim Preserve lyBytes(0 To ByteCount - 1)
        End If
    End If
    pyBytes = lyBytes
End Function

Public Function FileRead_i(ByVal hFile As Long, IntValue As Integer, ByteCount As Long) As Boolean
    FileRead_i = ReadFile(hFile, IntValue, 2, ByteCount, ByVal 0&) <> 0
    If FileRead_i Then FileRead_i = ByteCount = 2
End Function

Public Function FileRead_l(ByVal hFile As Long, LongValue As Long, ByteCount As Long) As Boolean
    FileRead_l = ReadFile(hFile, LongValue, 4, ByteCount, ByVal 0&) <> 0
    If FileRead_l Then FileRead_l = ByteCount = 4
End Function

Public Function FileRead_d(ByVal hFile As Long, DoubleValue As Double, ByteCount As Long) As Boolean
    FileRead_d = ReadFile(hFile, DoubleValue, 8, ByteCount, ByVal 0&) <> 0
    If FileRead_d Then FileRead_d = ByteCount = 8
End Function

Public Function FileRead_s(ByVal hFile As Long, StringValue As String, ByteCount As Long) As Boolean
    If ByteCount < 1 Then Exit Function
    StringValue = Space$(ByteCount + 1 \ 2)
    FileRead_s = ReadFile(hFile, ByVal StrPtr(StringValue), ByteCount, ByteCount, ByVal 0&) <> 0
    StringValue = StrConv(StringValue, vbUnicode)
    If FileRead_s Then FileRead_s = ByteCount > 0
    If FileRead_s Then
        If ByteCount < Len(StringValue) Then StringValue = Left$(StringValue, ByteCount)
    End If
End Function

Public Function FileWrite_y(ByVal hFile As Long, pyBytes() As Byte, ByteCount As Long) As Boolean
    Dim liL As Long, liBytes As Long
    On Error Resume Next
    liL = LBound(pyBytes)
    liBytes = UBound(pyBytes) - liL + 1
    If ByteCount = 0 Or ByteCount > liBytes Then ByteCount = liBytes
    If Err.Number <> 0 Or ByteCount < 1 Then Exit Function
    
    FileWrite_y = WriteFile(hFile, pyBytes(liL), ByteCount, ByteCount, ByVal 0&) <> 0
    If FileWrite_y Then FileWrite_y = ByteCount > 0
End Function

Public Function FileWrite_i(ByVal hFile As Long, ByVal IntValue As Integer, ByteCount As Long) As Boolean
    FileWrite_i = WriteFile(hFile, IntValue, 2, ByteCount, ByVal 0&) <> 0
    If FileWrite_i Then FileWrite_i = ByteCount > 0
End Function

Public Function FileWrite_l(ByVal hFile As Long, ByVal LongValue As Long, ByteCount As Long) As Boolean
    FileWrite_l = WriteFile(hFile, LongValue, 4, ByteCount, ByVal 0&) <> 0
    If FileWrite_l Then FileWrite_l = ByteCount > 0
End Function

Public Function FileWrite_d(ByVal hFile As Long, ByVal DoubleValue As Double, ByteCount As Long) As Boolean
    FileWrite_d = WriteFile(hFile, DoubleValue, 8, ByteCount, ByVal 0&) <> 0
    If FileWrite_d Then FileWrite_d = ByteCount > 0
End Function

Public Function FileWrite_s(ByVal hFile As Long, StringValue As String, ByteCount As Long) As Boolean
    Dim liLen As Long
    If ByteCount = 0 Then ByteCount = Len(StringValue)
    If ByteCount < 1 Then Exit Function
    liLen = Len(StringValue)
    If ByteCount > liLen Then ByteCount = liLen
    FileWrite_s = WriteFile(hFile, ByVal StrPtr(StrConv(StringValue, vbFromUnicode)), liLen, ByteCount, ByVal 0&) <> 0
    If FileWrite_s Then FileWrite_s = ByteCount > 0
End Function

Public Function FileClose(hFile As Long) As Boolean
    FileClose = CloseHandle(hFile) <> 0
    hFile = INVALID_HANDLE_VALUE
End Function

'Public Function InputFile(psFile As String, _
'                          pyBytes() As Byte) _
'                As Boolean
'    Dim liNum As Long
'    Dim lyBytes() As Byte
'    On Error GoTo finish
'    liNum = FreeFile
'    Open psFile For Binary Access Read As #liNum
'    ReDim lyBytes(0 To loF(liNum) - 1)
'    Get #liNum, , lyBytes
'finish:
'    InputFile = Err.Number = 0
'    On Error Resume Next
'    Close #liNum
'    pyBytes = lyBytes
'End Function
'
'Public Function OutputFile(psFile As String, _
'                           pyBytes() As Byte) _
'                As Boolean
'    Dim liNum As Long
'    liNum = FreeFile
'    On Error Resume Next
'    FileDelete psFile, True
'    Err.Clear
'    If Not PathCreate(PathGetParentFolder(psFile)) Then Exit Function
'    Open psFile For Binary Access Write As #liNum
'    Put #liNum, , pyBytes
'    Close #liNum
'    OutputFile = Err.Number = 0
'End Function
