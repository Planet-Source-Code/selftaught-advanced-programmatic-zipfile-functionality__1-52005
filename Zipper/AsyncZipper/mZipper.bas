Attribute VB_Name = "mZipper"
Option Explicit

' Callback large "string"
Private Type CBChar
    ch(0 To 4096) As Byte
End Type

' Store the callback functions
Private Type ZIPUSERFUNCTIONS
    lptrPrint As Long          ' Pointer to application's print routine
    lptrComment As Long        ' Pointer to application's comment routine
    lptrPassword As Long       ' Pointer to application's password routine.
    lptrService As Long        ' callback function designed to be used for allowing the
End Type                       ' app to process Windows messages, or cancelling the operation
                               ' as well as giving option of progress.  If this function returns
                               ' non-zero, it will terminate what it is doing.  It provides the app
                               ' with the name of the archive member it has just processed, as well
                               ' as the original size.

Private Type ZPOPT
  date           As String ' US Date (8 Bytes Long) "12/31/98"?
  szRootDir      As String ' Root Directory Pathname (Up To 256 Bytes Long)
  szTempDir      As String ' Temp Directory Pathname (Up To 256 Bytes Long)
  fTemp          As Long   ' 1 If Temp dir Wanted, Else 0
  fSuffix        As Long   ' Include Suffixes (Not Yet Implemented!)
  fEncrypt       As Long   ' 1 If Encryption Wanted, Else 0
  fSystem        As Long   ' 1 To Include System/Hidden Files, Else 0
  fVolume        As Long   ' 1 If Storing Volume Label, Else 0
  fExtra         As Long   ' 1 If Excluding Extra Attributes, Else 0
  fNoDirEntries  As Long   ' 1 If Ignoring Directory Entries, Else 0
  fExcludeDate   As Long   ' 1 If Excluding Files Earlier Than Specified Date, Else 0
  fIncludeDate   As Long   ' 1 If Including Files Earlier Than Specified Date, Else 0
  fVerbose       As Long   ' 1 If Full Messages Wanted, Else 0
  fQuiet         As Long   ' 1 If Minimum Messages Wanted, Else 0
  fCRLF_LF       As Long   ' 1 If Translate CR/LF To LF, Else 0
  fLF_CRLF       As Long   ' 1 If Translate LF To CR/LF, Else 0
  fJunkDir       As Long   ' 1 If Junking Directory Names, Else 0
  fGrow          As Long   ' 1 If Allow Appending To Zip File, Else 0
  fForce         As Long   ' 1 If Making Entries Using DOS File Names, Else 0
  fMove          As Long   ' 1 If Deleting Files Added Or Updated, Else 0
  fDeleteEntries As Long   ' 1 If Files Passed Have To Be Deleted, Else 0
  fUpdate        As Long   ' 1 If Updating Zip File-Overwrite Only If Newer, Else 0
  fFreshen       As Long   ' 1 If Freshing Zip File-Overwrite Only, Else 0
  fJunkSFX       As Long   ' 1 If Junking SFX Prefix, Else 0
  fLatestTime    As Long   ' 1 If Setting Zip File Time To Time Of Latest File In Archive, Else 0
  fComment       As Long   ' 1 If Putting Comment In Zip File, Else 0
  fOffsets       As Long   ' 1 If Updating Archive Offsets For SFX Files, Else 0
  fPrivilege     As Long   ' 1 If Not Saving Privileges, Else 0
  fEncryption    As Long   ' Read Only Property!!!
  fRecurse       As Long   ' 1 (-r), 2 (-R) If Recursing Into Sub-Directories, Else 0
  fRepair        As Long   ' 1 = Fix Archive, 2 = Try Harder To Fix, Else 0
  flevel         As Byte   ' Compression Level - 0 = Stored; Asc("6") = Default; Asc("9") = Max
End Type

'Version checking if you want to implement it.  Code in this component was written for Infozip's WinDll Zip32 Version 2.3
'Private Type ZpVerType
'    major        As Byte   'e.g., integer 5
'    minor        As Byte   'e.g., 2
'    patchlevel   As Byte   'e.g., 0
'    not_used     As Byte
'End Type
'Private Type ZpVer
'    structlen    As Long   'length of the struct being passed
'    flag         As Long   'bit 0: is_beta   bit 1: uses_zlib
'    betalevel    As String * 10 'e.g., "g BETA" or ""
'    date         As String * 20 'e.g., "4 Sep 95" (beta) or "4 September 1995"
'    zlib_version As String * 10 'e.g., "0.95" or NULL
'    zip          As ZpVerType
'    os2dll       As ZpVerType
'    windll       As ZpVerType
'End Type
'Private Declare Sub ZpVersion Lib "zip32.dll" ( _
                         ZpVersion As ZpVer _
                     )

' Set Zip Callbacks
'MUST call before EVERY ZpArchive call with a PROCEDURE level ZIPUSERFUNCTIONS var or else it can GPF!!!
Private Declare Function ZpInit Lib "zip32.dll" ( _
                             ByRef tUserFn As ZIPUSERFUNCTIONS _
                         ) As Long
                         
'Set Zip Flags
Private Declare Function ZpSetOptions Lib "zip32.dll" ( _
                             ByRef tOpts As ZPOPT _
                         ) As Long
                         
'Get Zip Flags (for checking encryption flag) but I don't use it here
'Private Declare Function ZpGetOptions Lib "zip32.dll" () As ZPOPT

'Perform action specified by the flags that were set
Private Declare Function ZpArchive Lib "zip32.dll" ( _
                             ByVal argc As Long, _
                             ByVal funame As String, _
                             ByRef argv As ZipNames _
                         ) As eZipErrorCodes

Private mbCancel    As Boolean 'flag that the user has canceled
Private mbZipping   As Boolean 'flag that we are currently busy

Private msComment   As String ' for storing the comment from the message callback
Private msPassword  As String ' for storing the password to pass to the message callback

Private moClient    As iZipCallBack 'Provide notifications to this object

'Tells the outside world whether we are busy or not
Public Property Get Zipping() As Boolean
    Zipping = mbZipping
End Property

'Find the real zipping action here
Public Sub ZipFiles( _
                ByRef ptInfo As tZipInfo, _
                ByVal poClient As iZipCallBack _
            )
    On Error Resume Next
    If mbZipping Then Exit Sub
    Dim ltZipNames As ZipNames 'Tells the DLL which files we are interested in
    Dim ltUser     As ZIPUSERFUNCTIONS 'Tells the DLL where to find us
    Dim ltOpt      As ZPOPT    'Tells the DLL exactly what we want to do
    
    Dim ldDate     As Date     ' a mark for including earlier or later dates
    Dim lsBasePath As String   ' root path for relative folder storage in zip
    
    Dim liAttributes  As eZipAttributes 'Temp variable instead of accessing the structure many times
    Dim liCompression As eZipCompression 'another temp variable
    Dim liMsgLevel    As eZipMessageLevel 'another temp variable
    Dim liReturn      As eZipErrorCodes 'store the error code of the zip function
    
    mbZipping = True
    mbCancel = False
    Set moClient = poClient
    
    With ptInfo
        msComment = Trim$(.Comment) 'Store comment so that it can be passed to the callback
        msPassword = Trim$(.Password) 'Store password so that it can be passed to the callback
        lsBasePath = .BasePath
        liAttributes = .Attributes
        liCompression = .Compression
        ldDate = .DateMark
        liMsgLevel = .MessageLevel
    End With
        
    With ltOpt
        If liMsgLevel = zipMsgMaximum Then
            .fVerbose = 1
        ElseIf liMsgLevel = zipMsgMinimum Or liMsgLevel < zipMsgMaximum Or liMsgLevel > zipMsgMinimum Then
            'If client specifies an invalid message level, he gets the minimum!
            .fQuiet = 1
        End If
        .fComment = Abs(Len(msComment) > 0) 'If we have a comment, indicate it
        .fEncrypt = Abs(Len(msPassword) > 0) 'If we have a password, indicate it
        .szRootDir = lsBasePath 'For storing relative folders in the zip

        'Set the other attributes that were specified
        If liAttributes And zipCRLFtoLF Then .fCRLF_LF = 1
        If liAttributes And zipLFtoCRLF Then .fLF_CRLF = 1
        If liAttributes And zipForceDOSFileNames Then .fForce = 1
        If liAttributes And zipDeleteFileSpecs Then .fDeleteEntries = 1
        If Not (liAttributes And zipIgnoreSystemAndHidden) Then .fSystem = 1
        If liAttributes And zipOnlyIfNewer Then .fUpdate = 1
        If liAttributes And zipLatestTime Then .fLatestTime = 1
        If liAttributes And zipRecurse Then .fRecurse = 1
        If liAttributes And zipForceRepair Then
            .fRepair = 2
        ElseIf liAttributes And zipRepair Then
            .fRepair = 1
        End If
        If Not (liAttributes And zipIncludeDirectoryEntries) Then .fNoDirEntries = 1

        If ldDate <> #12:00:00 AM# Then 'If date has been set
            .date = Format$(ldDate, "MM/DD/YY")
            'zipExcludeEarlierDates = 0 and zipIncludeEarlierDates = 1, so mod by 2 to get just that attribute
            If liAttributes Mod 2 = 1 Then .fIncludeDate = 1 Else .fExcludeDate = 1
        End If

        .fJunkDir = Abs(Len(lsBasePath) = 0 And .fDeleteEntries = 0)  'if no base path, then no need to store folder names.
        
        'DLL wants asc of 1-9 OR 0& for default, so we add vbKey0 if specified compression it is greater than 0
        If liCompression > zipCompressionStored And liCompression <= zipCompression9Maximum Then .flevel = liCompression + vbKey0
    End With
    
    'Yes, it does appear to be necessary
    If LenB(Trim$(lsBasePath)) > 0 Then ChDir lsBasePath

    With ltUser 'Fill the structure that tells the DLL how to find us
        .lptrPrint = AddrFunc(AddressOf ZipPrintCallback)
        .lptrPassword = AddrFunc(AddressOf ZipPasswordCallback)
        .lptrComment = AddrFunc(AddressOf ZipCommentCallback)
        .lptrService = AddrFunc(AddressOf ZipServiceCallback)
    End With
    ZpInit ltUser 'Initialize the DLL
    
    'Set the current options
    ZpSetOptions ltOpt

    With ptInfo
        'Change the string array of filenames to the zipnames format
        liReturn = TranslateStringArray(.FileSpecs, ltZipNames.s, ltOpt.fDeleteEntries = 1)
        'Do it!
        liReturn = ZpArchive(liReturn, .FileName, ltZipNames)
    End With
    
    'Set client to nothing BEFORE calling zipcomplete, so that
    'the client can start another action during this notification
    Set moClient = Nothing
    msComment = vbNullString
    msPassword = vbNullString
    mbZipping = False
    poClient.ZipComplete liReturn
End Sub

Private Function ZipServiceCallback( _
                     ByRef mname As CBChar, _
                     ByVal x As Long _
                 ) As Long
    On Error Resume Next
    If mbCancel Then ZipServiceCallback = 1: Exit Function
    'Nice and simple, just get the filename from the structure and tell the client about it
    'If the client wants to cancel, then tell InfoZip about it by setting return value to 1
    Dim lbCancel As Boolean
    moClient.ZippedFile Replace$(GetString(mname.ch, x), "/", "\"), lbCancel
    If lbCancel Then: mbCancel = True: ZipServiceCallback = 1
End Function

Private Function ZipPrintCallback( _
                     ByRef fname As CBChar, _
                     ByVal x As Long _
                 ) As Long
    
    On Error Resume Next
    'Get the message from the structure, then tell the client about it.
    Dim lsTemp As String
    lsTemp = GetString(fname.ch, x)
    TrimMsg lsTemp
    If LenB(lsTemp) > 0 And StrComp(lsTemp, ".", vbBinaryCompare) <> 0 Then moClient.ZipMessage lsTemp
End Function

Private Function ZipCommentCallback( _
                     ByRef comm As CBChar _
                 ) As Long
    On Error Resume Next
    Dim i As Long

    If mbCancel Or LenB(msComment) = 0 Then
        'Shouldn't happen, but just in case
        comm.ch(0) = 0
        mbCancel = True
        ZipCommentCallback = 1
    Else
        'put the string into the structure
        TranslateString msComment, comm.ch, 254
    End If
    
End Function

Private Function ZipPasswordCallback( _
                     ByRef pwd As CBCh, _
                     ByVal maxPasswordLength As Long, _
                     ByRef s2 As CBCh, _
                     ByRef Name As CBCh _
                 ) As Long
    
    On Error Resume Next
    
    If mbCancel Or LenB(msPassword) = 0 Then
        'shouldn't happen, but just in case
        pwd.ch(0) = 0
        mbCancel = True
        ZipPasswordCallback = 1
    Else
        'put the string into the structure
        TranslateString msPassword, pwd.ch, maxPasswordLength
    End If
    
End Function
