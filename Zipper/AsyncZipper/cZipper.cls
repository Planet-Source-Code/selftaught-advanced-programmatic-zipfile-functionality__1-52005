VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cZipFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Type tZipInfo
    FileSpecs()  As String 'Load this string array with the files to be zipped.  Wildcards are OK
    BasePath     As String 'Blank if discarding folder names, otherwise the relative parent folder to be used
    Password     As String 'Blank if not encrypting, otherwise the password
    Comment      As String 'Blank if no comment, otherwise the comment
    FileName     As String 'Name of the file to create/modify
    DateMark     As Date   'Used in conjunction w/ zipIncludeEarlierDates and zipExcludeEarlierDates
    Attributes   As eZipAttributes 'Other misc. Attributes
    Compression  As eZipCompression 'Compression level 0-9
    MessageLevel As eZipMessageLevel 'Which messages are requested
End Type

Public Type tUnzipInfo
    FileName      As String 'Name of file to extract from
    ExtractToPath As String 'Extract to folder
    Include()     As String 'Files to include
    Exclude()     As String 'folders to include
    Password      As String 'Blank if no password
    OverwriteAll  As Boolean 'If you don't want any confirmation before overwrites
    Attributes    As eUnzipAttributes 'Other attributes
    MessageLevel  As eZipMessageLevel 'Messages that are requested
End Type

Public Enum eUnzipOverwrite
    zipDoNotOverwrite = 100
    zipOverwriteThisFile = 102
    zipOverwriteAllFiles = 103
    zipOverwriteNone = 104
End Enum

Public Enum eUnzipAttributes
    zipExtractOnlyNewer = 1
    zipSpaceToUnderscore = 2
    zipJustTesting = 4
    zipDisregardFolderNames = 8
    zipCRtoCRLF = 16
    zipCaseSensitive = 32
    zipCallbackForPassword = 64
End Enum

Public Enum eZipAttributes
    zipExcludeEarlierDates
    zipIncludeEarlierDates
    zipCRLFtoLF
    zipLFtoCRLF = 4
    zipForceDOSFileNames = 8
    zipDeleteFileSpecs = 16
    zipIgnoreSystemAndHidden = 32
    zipOnlyIfNewer = 64
    zipLatestTime = 128
    zipRecurse = 256
    zipRepair = 512
    zipForceRepair = 1024
    zipIncludeDirectoryEntries = 2048
End Enum

Public Enum eZipErrorCodes
    zipErrNone                  'Normal; no errors or warnings detected.
    zipErrFileDamaged = 2       'The zipfile is either truncated or damaged in some way (e.g., bogus internal offsets) that makes it appear to be truncated.
    zipErrInvalidStructure      'The structure of the zipfile is invalid; for example, it may have been corrupted by a text-mode ("ASCII") transfer.
    zipErrMemFail               'Zip was unable to allocate sufficient memory to complete the command.
    zipErrInternal              'Internal logic error. (This should never happen; it indicates a programming error of some sort.)
    zipErrZipSplitFail          'ZipSplit was unable to create an archive of the specified size because the compressed size of a single included file is larger than the requested size. (Note that Zip and ZipSplit still do not support the creation of PKWARE-style multi-part archives.)
    zipErrInvalidComment        'The format of a zipfile comment was invalid.
    zipErrTestFail              'Testing (-T option) failed due to errors in the archive, insufficient memory to spawn Unzip, or inability to find Unzip.
    zipErrUserCancel            'Zip was interrupted by user (or superuser) action.
    zipErrTempFileFail          'Zip encountered an error creating or using a temporary file.
    zipErrFileJumpFail          'Reading or seeking (jumping) within an input file failed.
    zipErrNothingToDo           'There was nothing for Zip to do (e.g., "zip foo.zip").
    zipErrFileMissing           'The zipfile was missing or empty (typically when updating or freshening).
    zipErrFileWriteFail         'Zip encountered an error writing to an output file (typically the archive); for example, the disk may be full.
    zipErrFileOpenFail          'Zip could not open an output file (typically the archive) for writing.
    zipErrInvalidCommandLine    'The command-line parameters were specified incorrectly.
    zipErrFileReadFail = 18     'Zip could not open a specified file for reading; either it doesn't exist or the user running Zip doesn't have permission to read it.
End Enum

Public Enum eUnzipErrorCodes
    zipUErrNone                 'Normal; no errors or warnings detected. (There may still be errors in the archive, but if so, they weren't particularly relevant to Unzip's processing and are presumably quite minor.)
    zipUErrWarning              'One or more warning errors were encountered, but processing completed successfully anyway. This includes zipfiles where one or more files was skipped due to unsupported compression method or encryption with an unknown password.
    zipUErrMinor                'A generic error in the zipfile format was detected. Processing may have completed successfully anyway; some broken zipfiles created by other archivers have simple work-arounds.
    zipUErrMajor                'A severe error in the zipfile format was detected. Processing probably failed immediately.
    zipUErrMemFailInit          'Unzip was unable to allocate memory for one or more buffers during program initialization.
    zipUErrMemFailPass          'Unzip was unable to allocate memory or unable to obtain a tty (terminal) to read the decryption password(s).
    zipUErrMemFailUnzipDisk     'Unzip was unable to allocate memory during decompression to disk.
    zipUErrMemFailUnzipMem      'Unzip was unable to allocate memory during in-memory decompression.
    zipUErrFileNotFound         'The specified zipfile(s) was not found.
    zipUErrInvalidCommandLine   'Invalid options were specified on the command line.
    zipUErrNoMatchingFiles      'No matching files were found.
    zipUErrDiskIsFull = 50      'The disk is (or was) full during extraction.
    zipUErrUnexpectedEOF        'The end of the ZIP archive was encountered prematurely.
    zipUErrUserCancel = 80      'The user aborted Unzip prematurely with control-C (or similar)
    zipUErrUnsupportedFormat    'Testing or extraction of one or more files failed due to unsupported compression methods or unsupported decryption.
    zipUErrBadPassword          'No files were found due to bad decryption password(s). (If even one file is successfully processed, however, the exit status is 1.)
End Enum

Public Enum eZipMessageLevel
    zipMsgMaximum
    zipMsgMiddle
    zipMsgMinimum
End Enum

Public Enum eZipCompression
    zipCompressionStored
    zipCompression1
    zipCompression2
    zipCompression3
    zipCompression4
    zipCompression5
    zipCompression6Default = 6
    zipCompression7
    zipCompression8
    zipCompression9Maximum = 9
End Enum

'Why can't I pass stupid structures or arrays byval!!!!
'Must copy it again back into the client's address space, but oh well.
Public Function ZipFiles( _
                    ByRef ZipInfo As tZipInfo, _
                    ByVal ForMe As iZipCallBack _
                ) As Boolean
    
    If Zipping Then Exit Function
    
    Dim ltInfo As tZipInfo
    LSet ltInfo = ZipInfo
    
    With ltInfo
        'Some quick validations
        If Len(.BasePath) > 0 Then
            If Not FolderExists(.BasePath) Then Exit Function
        End If
        If Not FileIsValidToCreate(.FileName) Then Exit Function
    End With
    ZipFiles = True
    'Make the request
    mUtility.Callback ForMe, zipZip, ltInfo
End Function

Public Function UnzipFiles( _
                    ByRef UnzipInfo As tUnzipInfo, _
                    ByVal ForMe As iZipCallBack _
                ) As Boolean
    
    If Unzipping Then Exit Function
    'some quick validations
    If Not FileExists(UnzipInfo.FileName) Then Exit Function
    If Not PathCreate(UnzipInfo.ExtractToPath) Then Exit Function
    
    Dim ltInfo As tUnzipInfo
    LSet ltInfo = UnzipInfo
    
    UnzipFiles = True
    'make the request
    mUtility.Callback ForMe, zipUnzip, ltInfo
End Function

Public Function ReadZipFile( _
                    ByVal FileName As String, _
                    ByVal ForMe As iZipCallBack _
                ) As Boolean
    
    If Unzipping Then Exit Function
    'a quick validation
    If Not FileExists(FileName) Then Exit Function
    
    ReadZipFile = True
    'make the request
    mUtility.Callback ForMe, zipRead, FileName
    
End Function

Public Function GetComment( _
                    ByVal FileName As String, _
                    ByRef Comment As String _
                ) As Boolean
    
    If Unzipping Then Exit Function
    'a quick validation
    If Not FileExists(FileName) Then Exit Function
    GetComment = mUnzipper.GetComment(FileName, Comment)
    
End Function
