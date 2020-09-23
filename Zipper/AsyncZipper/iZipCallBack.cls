VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iZipCallBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub OverwriteRequest( _
               ByVal FileName As String, _
               ByRef Answer As eUnzipOverwrite _
           )
'Called to request overwriting of a file
End Sub

Public Sub PasswordRequest( _
               ByVal ForFile As String, _
               ByRef Password As String, _
               ByVal WasInvalid As Boolean, _
               ByRef ApplyToAll As Boolean, _
               ByRef Cancel As Boolean _
           )
'Called to request a password if the attribute was specified or if no password
'was specified and the file requires one or if the password specified is invalid
End Sub

Public Sub ReadComplete( _
               ByVal ErrorCode As eUnzipErrorCodes, _
               ByVal HasComment As Boolean, _
               ByVal CompressedSize As Long, _
               ByVal TotalSize As Long, _
               ByVal NumMembers As Long, _
               ByVal CompressionFactor As Long _
           )
'Called when the reading is done
End Sub

Public Sub ReadFile( _
               ByVal FileName As String, _
               ByVal Size As Long, _
               ByVal CompressedSize As Long, _
               ByVal CompressionFactor As Long, _
               ByVal FileDate As Date, _
               ByVal CRC As Long, _
               ByVal Encrypted As Boolean _
           )
'called for each file in the zip after calling ReadZipFile
End Sub

Public Sub UnzipComplete( _
               ByVal ErrorCode As eUnzipErrorCodes _
           )
'called when extracting is done
End Sub

Public Sub UnzippedFile( _
               ByVal FileName As String, _
               ByRef Cancel As Boolean _
           )
'called when a file is extracted, gives an option to cancel
End Sub

Public Sub UnzipMessage( _
               ByVal Msg As String _
           )
'other messages regarding Unzip
End Sub

Public Sub ZipComplete( _
               ByVal ErrorCode As eZipErrorCodes _
           )
'called when zip is done
End Sub

Public Sub ZippedFile( _
               ByVal FileName As String, _
               ByRef Cancel As Boolean _
           )
'called after compressing a file into a zip, gives option to cancel.
End Sub

Public Sub ZipMessage( _
               ByVal Msg As String _
           )
'other messages regarding zip
End Sub



