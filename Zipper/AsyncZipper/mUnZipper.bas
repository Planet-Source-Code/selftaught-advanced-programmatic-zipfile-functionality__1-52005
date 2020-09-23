Attribute VB_Name = "mUnzipper"
Option Explicit

' Callback large "string"
Private Type CBChar
    ch(0 To 32800) As Byte
End Type

' DCL structure
Private Type DCLIST
   ExtractOnlyNewer  As Long    ' 1 = Extract Only Newer/New, Else 0
   SpaceToUnderscore As Long    ' 1 = Convert Space To Underscore, Else 0
   PromptToOverwrite As Long    ' 1 = Prompt To Overwrite Required, Else 0
   fQuiet            As Long    ' 2 = No Messages, 1 = Less, 0 = All
   ncflag            As Long    ' 1 = Write To Stdout, Else 0
   ntflag            As Long    ' 1 = Test Zip File, Else 0
   nvflag            As Long    ' 0 = Extract, 1 = List Zip Contents
   nuflag            As Long    ' 1 = Extract Only Newer Over Existing, Else 0
   nzflag            As Long    ' 1 = Display Zip File Comment, Else 0
   ndflag            As Long    ' 1 = Honor Directories, Else 0
   noflag            As Long    ' 1 = Overwrite Files, Else 0
   naflag            As Long    ' 1 = Convert CR To CRLF, Else 0
   nZIflag           As Long    ' 1 = Zip Info Verbose, Else 0
   C_flag            As Long    ' 1 = Case Insensitivity, 0 = Case Sensitivity
   fPrivilege        As Long    ' 1 = ACL, 2 = Privileges
   lpszZipFN         As String  ' The Zip Filename To Extract Files
   lpszExtractDir    As String  ' The Extraction Directory, NULL If Extracting To Current Dir
End Type

Private Type USERFUNCTION
   ' Callbacks:
   lptrPrnt As Long           ' Pointer to application's print routine
   lptrSound As Long          ' Pointer to application's sound routine.  NULL if app doesn't use sound
   lptrReplace As Long        ' Pointer to application's replace routine.
   lptrPassword As Long       ' Pointer to application's password routine.
   lptrMessage As Long        ' Pointer to application's routine for
                              ' displaying information about specific files in the archive
                              ' used for listing the contents of the archive.
   lptrService As Long        ' callback function designed to be used for allowing the
                              ' app to process Windows messages, or cancelling the operation
                              ' as well as giving option of progress.  If this function returns
                              ' non-zero, it will terminate what it is doing.  It provides the app
                              ' with the name of the archive member it has just processed, as well
                              ' as the original size.
                              
   ' Values filled in after processing:
   lTotalSizeComp As Long     ' Value to be filled in for the compressed total size, excluding
                              ' the archive header and central directory list.
   lTotalSize As Long         ' Total size of all files in the archive
   lCompFactor As Long        ' Overall archive compression factor
   lNumMembers As Long        ' Total number of files in the archive
   cchComment As Integer      ' Flag indicating whether comment in archive.
End Type

'Version checking if you want to implement it.  Code in this component was written for Infozip's WinDll Unzip32 Version 5.5
'Note: Unzip32 5.5 seems to identify itself as version 5.4 using the UzpVersion2 method
'Private Type ZIPVERSIONTYPE
'   major As Byte
'   minor As Byte
'   patchlevel As Byte
'   not_used As Byte
'End Type
'Private Type UZPVER
'   structlen       As Long           ' Length Of The Structure Being Passed
'   flag            As Long           ' Bit 0: is_beta  bit 1: uses_zlib
'   beta            As String * 10    ' e.g., "g BETA" or ""
'   date            As String * 20    ' e.g., "4 Sep 95" (beta) or "4 September 1995"
'   zlib            As String * 10    ' e.g., "1.0.5" or NULL
'   Unzip(1 To 4)   As ZIPVERSIONTYPE ' Version Type Unzip
'   ZipInfo(1 To 4) As ZIPVERSIONTYPE ' Version Type Zip Info
'   os2dll          As Long           ' Version Type OS2 DLL
'   windll(1 To 4)  As ZIPVERSIONTYPE ' Version Type Windows DLL
'End Type
'Private Declare Sub UzpVersion2 Lib "vbuzip10.dll" ( _
                         uzpv As UZPVER _
                     )

Private Declare Function Wiz_SingleEntryUnzip Lib "Unzip32.dll" ( _
                             ByVal ifnc As Long, _
                             ByRef ifnv As ZipNames, _
                             ByVal xfnc As Long, _
                             ByRef xfnv As ZipNames, _
                             ByRef dcll As DCLIST, _
                             ByRef Userf As USERFUNCTION _
                         ) As eUnzipErrorCodes
   
Private mbCancel         As Boolean 'If the user has canceled
Private mbUnzipping      As Boolean 'If we are busy working
Private mbGettingComment As Boolean 'for the callback to decide if we are trying to extract the comment
Private mbGotComment     As Boolean 'for the callback to avoid setting the comment more than once
Private mbPromptForPass  As Boolean 'if the client is to be queried for the password

Private msLastFile       As String 'used to only ask for passwords once
Private msLastInvalid    As String 'used to provide only one notification of an invalid password
Private msComment        As String 'for storing the comment between different methods
Private msPassword       As String 'for storing the password between different methods
Private msTempPassword   As String 'for storing the client's response to PasswordRequest
Private moClient         As iZipCallBack 'To provide notifications

'Tells the outside world if we are busy
Public Property Get Unzipping() As Boolean
    Unzipping = mbUnzipping
End Property

'Read a file from the zip
Private Sub UnzipMessageCallBack( _
                ByVal ucsize As Long, _
                ByVal csiz As Long, _
                ByVal cfactor As Integer, _
                ByVal mo As Integer, _
                ByVal dy As Integer, _
                ByVal yr As Integer, _
                ByVal hh As Integer, _
                ByVal mm As Integer, _
                ByVal c As Byte, _
                ByRef fname As CBCh, _
                ByRef meth As CBCh, _
                ByVal CRC As Long, _
                ByVal fCrypt As Byte _
            )
    
    On Error Resume Next
    Dim lsName As String
    lsName = Replace$(GetString(fname.ch), "/", "\")
    
    'only want to tell the client if it is a file and not a directory
    If Not StrComp(Right$(lsName, 1), "\") = 0 Then _
        moClient.ReadFile lsName, ucsize, csiz, cfactor, DateSerial(yr, mo, dy) + TimeSerial(hh, mm, 0), CRC, fCrypt And 64
    'End If
End Sub

Private Function UnzipPrintCallback( _
                 ByRef fname As CBChar, _
                 ByVal x As Long _
                 ) As Long
    On Error Resume Next
    If mbGettingComment Then
        'If we are trying to get the comment then store it
        msComment = GetString(fname.ch, x)
        mbGettingComment = False
        mbGotComment = True
    Else
        'If we were not trying to get the comment, then this is a regular
        'message during an Unzip opertion
        If Not mbGotComment Then
            Dim lsTemp As String
            lsTemp = GetString(fname.ch, x)
            TrimMsg lsTemp
            If LenB(lsTemp) > 0 And StrComp(lsTemp, ".", vbBinaryCompare) <> 0 Then moClient.UnzipMessage lsTemp
        End If
    End If
End Function

Private Function UnzipPasswordCallBack( _
                     ByRef pwd As CBCh, _
                     ByVal x As Long, _
                     ByRef s2 As CBCh, _
                     ByRef Name As CBCh _
                 ) As Long
    On Error Resume Next
    Dim lsTemp As String
    Dim lbWasInvalid As Boolean
    Dim lbAll As Boolean
    Dim lbCancel As Boolean
    
    If mbCancel Then: UnzipPasswordCallBack = 1: pwd.ch(0) = 0: Exit Function
    
    lbWasInvalid = InStr(1, GetString(s2.ch), "incorrect") > 0
    'If we have no password or were asked to prompt for it
    If mbPromptForPass Or Len(msPassword) = 0 Or lbWasInvalid Then
        'Get the name of the file we're asking about
        lsTemp = Replace(GetString(Name.ch), "/", "\")
        'If we didn't just ask or (we just asked once, and now the password is invalid) then
        If StrComp(lsTemp, msLastFile) <> 0 _
                    Or _
           (StrComp(lsTemp, msLastInvalid) <> 0 And lbWasInvalid) Then
            msLastFile = lsTemp
            If lbWasInvalid Then msLastInvalid = lsTemp
            'query the client for the password
            moClient.PasswordRequest msLastFile, msTempPassword, lbWasInvalid, lbAll, lbCancel
            If lbCancel Then
                mbCancel = True
                UnzipPasswordCallBack = 1
                Exit Function
            End If
            If lbAll Then msPassword = msTempPassword
        End If
        'Give the password to the DLL
        TranslateString msTempPassword, pwd.ch, 254
    Else
        'Give the password to the DLL
        TranslateString msPassword, pwd.ch, 254
    End If

End Function

Private Function UnzipReplaceCallback( _
                     ByRef fname As CBChar _
                 ) As eUnzipOverwrite
   On Error Resume Next
   'This is the default
   UnzipReplaceCallback = zipDoNotOverwrite
   'Ask the client for permission to overwrite
   moClient.OverwriteRequest Replace$(GetString(fname.ch), "/", "\"), UnzipReplaceCallback
   'If the client gives an invalid value, then do not overwrite
   If UnzipReplaceCallback < zipDoNotOverwrite Or UnzipReplaceCallback > zipOverwriteNone Then UnzipReplaceCallback = zipDoNotOverwrite
End Function

Private Function UnzipServiceCallback( _
                     ByRef mname As CBChar, _
                     ByVal x As Long _
                 ) As Long
    On Error Resume Next
    If mbCancel Then: UnzipServiceCallback = 1: Exit Function
    'Let the client know about it
    Dim lbCancel As Boolean
    moClient.UnzippedFile Replace$(GetString(mname.ch), "/", "\"), lbCancel
    'if the client wants to cancel, then tell Infozip about it
    If lbCancel Then: mbCancel = True: UnzipServiceCallback = 1
End Function

'Look Here for the real action
Public Sub UnzipFiles( _
               ByRef ptInfo As tUnzipInfo, _
               ByVal poClient As iZipCallBack _
           )
    Dim ltInc        As ZipNames 'Filenames to include
    Dim ltExc        As ZipNames 'Filenames to exclude
    Dim ltUser       As USERFUNCTION
    Dim liIncCount   As Long
    Dim liExcCount   As Long
    Dim lsExtract    As String 'Folder to extract to
    Dim lsFileName   As String 'Filename to extract from
    Dim ltDCL        As DCLIST 'flags to tell infozip exactly what we want to do
    Dim liReturn     As eUnzipErrorCodes 'return value/error code
    Dim liAttributes As eUnzipAttributes 'temporary var to avoid accessing structure repeatedly
    
    If mbUnzipping Then Exit Sub
    On Error Resume Next
    
    mbUnzipping = True
    mbCancel = False
    Set moClient = poClient 'Store the client so we can make callbacks to them
    
    InitUser ltUser 'Fill the structure with our callback addresses

    With ptInfo
        'put the user-specified filenames into the correct format
        liIncCount = TranslateStringArray(.Include, ltInc.s, True)
        liExcCount = TranslateStringArray(.Exclude, ltExc.s, True)
        liAttributes = .Attributes 'store other members of this structure
        lsExtract = .ExtractToPath
        lsFileName = .FileName
        msPassword = Trim$(Left$(.Password, 254)) 'Don't be passing short stories as passwords now!
        msTempPassword = msPassword
    End With
    
    With ltDCL
        If ptInfo.MessageLevel >= zipMsgMaximum And _
           ptInfo.MessageLevel <= zipMsgMinimum Then
            .fQuiet = ptInfo.MessageLevel
        Else
            'If the client specifies an incorrect message level, he gets the minimum!
            .fQuiet = zipMsgMinimum
        End If
        
        .lpszExtractDir = lsExtract 'Directory to extract to
        .lpszZipFN = lsFileName     'Filename to extract from
        'set all other attributes specified by the user
        If Not (liAttributes And zipCaseSensitive) Then .C_flag = 1
        If liAttributes And zipCRtoCRLF Then .naflag = 1
        If Not (liAttributes And zipDisregardFolderNames) Then .ndflag = 1
        If liAttributes And zipExtractOnlyNewer Then .ExtractOnlyNewer = 1
        If liAttributes And zipJustTesting Then .ntflag = 1
        If liAttributes And zipSpaceToUnderscore Then .SpaceToUnderscore = 1
        If ptInfo.OverwriteAll Then .noflag = 1 Else .PromptToOverwrite = 1
        mbPromptForPass = liAttributes And zipCallbackForPassword
        .nvflag = 0
    End With
   
    'Do it!
    liReturn = Wiz_SingleEntryUnzip(liIncCount, ltInc, liExcCount, ltExc, ltDCL, ltUser)
    
    'Set the modular client to nothing BEFORE calling UnzipComplete, so that the
    'client can make another request during this callback.
    
    Set moClient = Nothing
    mbUnzipping = False
    
    poClient.UnzipComplete liReturn

End Sub

'Some more real action, not as interesting though
Public Sub ReadZipFile( _
               ByRef psFile As String, _
               ByVal poClient As iZipCallBack _
           )
    
    Dim ltDCL As DCLIST 'Store the options for reading the zip
    Dim ltUser As USERFUNCTION
    Dim lt As ZipNames  'a blank zipnames to pass to the Unzip function
    Dim liReturn As eUnzipErrorCodes 'return value/error code
    
    If mbUnzipping Then Exit Sub
    
    On Error Resume Next
    mbCancel = False
    lt.s(0) = vbNullChar
    mbUnzipping = True
    Set moClient = poClient
    
    InitUser ltUser 'Fill the structure with the callback addresses
    
    With ltDCL
        'not many options need to be set for just reading the file
        .lpszExtractDir = vbNullChar 'Of course, we're not extracting
        .lpszZipFN = psFile 'This is the file we want to read
        .fQuiet = 1
        .nvflag = 1
    End With
   
    'Do it!
    liReturn = Wiz_SingleEntryUnzip(0, lt, 0, lt, ltDCL, ltUser)
    
    'Set the client to nothing BEFORE making the ReadComplete Notification, so
    'that the client can make another request during this method.
    Set moClient = Nothing
    mbUnzipping = False
    mbCancel = False
    With ltUser
        poClient.ReadComplete liReturn, (.cchComment > 0), .lTotalSizeComp, .lTotalSize, .lNumMembers, .lCompFactor
    End With
    
End Sub

'This is the only public function in mUnzipper and mZipper that is called synchronously.
Public Function GetComment( _
                    ByRef psFile As String, _
                    ByRef Comment As String _
                ) As Boolean
    
    Dim ltDCL As DCLIST 'Options for getting the password
    Dim ltUser As USERFUNCTION
    Dim lt As ZipNames 'blank structure to pass to the DLL
    
    If mbUnzipping Then Exit Function
    
    On Error Resume Next
    lt.s(0) = vbNullChar
    
    'Prepare modular state for callback
    mbUnzipping = True
    mbGettingComment = True
    mbGotComment = False
    msComment = vbNullString
    mbCancel = False
    
    InitUser ltUser 'Fill structure with callback addresses
    
    With ltDCL
        .lpszExtractDir = vbNullChar 'We aren't extracting anything
        .lpszZipFN = psFile 'Get the comment for this file
        .fQuiet = 1
        .nzflag = 1 'Tell Infozip that we would please like to have the comment
    End With
    
    'Do It!
    GetComment = Wiz_SingleEntryUnzip(0, lt, 0, lt, ltDCL, ltUser) = 0
    
    'Reset Modular state
    mbGettingComment = False
    mbGotComment = False
    mbUnzipping = False
    Comment = msComment
    msComment = vbNullString
    
End Function

Private Sub InitUser(ptUser As USERFUNCTION)
    With ptUser
        .lptrPrnt = AddrFunc(AddressOf UnzipPrintCallback)
        .lptrSound = 0& ' not supported
        .lptrReplace = AddrFunc(AddressOf UnzipReplaceCallback)
        .lptrPassword = AddrFunc(AddressOf UnzipPasswordCallBack)
        .lptrMessage = AddrFunc(AddressOf UnzipMessageCallBack)
        .lptrService = AddrFunc(AddressOf UnzipServiceCallback)
    End With
End Sub
