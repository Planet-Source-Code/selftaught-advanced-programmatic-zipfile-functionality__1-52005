Attribute VB_Name = "mDecode"
Option Explicit
Private moBase64 As cBase64
Private moInFile As cFileIO
Private moOutFile As cFileIO
Private myUnencoded() As Byte
Private myEncoded() As Byte

Public Sub Main()
    Set moInFile = New cFileIO
    With moInFile
        .FileAccess = GENERIC_READ
        .FileCreation = OPEN_EXISTING
        .FileFlags = FILE_FLAG_SEQUENTIAL_SCAN
        .FileShare = FILE_SHARE_READ
    End With
    Set moOutFile = New cFileIO
    With moOutFile
        .FileAccess = GENERIC_WRITE
        .FileCreation = CREATE_ALWAYS
        .FileFlags = FILE_FLAG_SEQUENTIAL_SCAN
        .FileShare = FILE_SHARE_READ
    End With
    Set moBase64 = New cBase64
    
    
    Dim lsEncodedPath As String
    Dim lsSystemPath As String
    lsEncodedPath = PathBuild(App.Path, "..\AsyncZipper\")
    lsSystemPath = PathGetSpecial(sfSystem)
    PathAddBackslash lsSystemPath
    If MsgBox("This program will place Unzip32.dll v.5.5 and Zip32.dll v.2.3 in your system folder.  These are open-source dlls (same code used by WinZip) that are available at www.Info-Zip.org." & vbNewLine & vbNewLine & "Are you sure that you want to place these two files into your system folder?  (This will overwrite the files if they already exist)", vbYesNo + vbQuestion, "Confirm File Placement") = vbYes Then
        DecodeFile lsEncodedPath & "ZipDll.txt", lsSystemPath & "Zip32.dll"
        DecodeFile lsEncodedPath & "UnZipDll.txt", lsSystemPath & "Unzip32.dll"
        MsgBox "Yep, done already!"
    End If
End Sub

Private Sub EncodeFile(Source As String, Dest As String)
    On Error Resume Next
    With moInFile
        If .OpenFile(Source) Then
            If moOutFile.OpenFile(Dest) Then
                Do
                    .GetBytes myUnencoded
                    moBase64.EncodeB64 myUnencoded, myEncoded
                    moOutFile.AppendBytes myEncoded
                Loop While Not .EOF
            End If
        End If
    End With
    moInFile.CloseFile
    moOutFile.CloseFile
    Erase myUnencoded
    Erase myEncoded
End Sub

Private Sub DecodeFile(Source As String, Dest As String)
    On Error Resume Next
    With moInFile
        If .OpenFile(Source) Then
            If moOutFile.OpenFile(Dest) Then
                Do
                    .GetBytes myEncoded
                    moBase64.DecodeB64 myEncoded, myUnencoded
                    moOutFile.AppendBytes myUnencoded
                Loop While Not .EOF
            End If
        End If
    End With
    moInFile.CloseFile
    moOutFile.CloseFile
    Erase myUnencoded
    Erase myEncoded
End Sub
