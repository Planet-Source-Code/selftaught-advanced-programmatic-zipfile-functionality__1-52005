Attribute VB_Name = "mFileSystem"
Option Explicit
Private mCollTempNames As Collection
Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
'Parts of this module contain code that was formed substantially from code seen at http://www.mentalis.org or all-api.net or www.pscode.com

'#######PUBLIC CONSTS################

'Maximum length for some API functions
    Public Const MAX_PATH = 260&
'
    Public Const KB As Long = 1024&

'#######PUBLIC ENUMS#################
'Flags for open and save common dialogs
    Public Enum eOFNFlags
        OFN_ALLOWMULTISELECT = &H200
        OFN_CREATEPROMPT = &H2000
        OFN_ENABLEHOOK = &H20
        OFN_ENABLETEMPLATE = &H40
        OFN_ENABLETEMPLATEHANDLE = &H80
        OFN_EXPLORER = &H80000                         '  new look commdlg
        OFN_EXTENSIONDIFFERENT = &H400
        OFN_FILEMUSTEXIST = &H1000
        OFN_HIDEREADONLY = &H4
        OFN_LONGNAMES = &H200000                       '  force long names for 3.x modules
        OFN_NOCHANGEDIR = &H8
        OFN_NODEREFERENCELINKS = &H100000
        OFN_NOLONGNAMES = &H40000                      '  force no long names for 4.x modules
        OFN_NONETWORKBUTTON = &H20000
        OFN_NOREADONLYRETURN = &H8000
        OFN_NOTESTFILECREATE = &H10000
        OFN_NOVALIDATE = &H100
        OFN_OVERWRITEPROMPT = &H2
        OFN_PATHMUSTEXIST = &H800
        OFN_READONLY = &H1
        OFN_SHAREAWARE = &H4000
        OFN_SHAREFALLTHROUGH = 2
        OFN_SHARENOWARN = 1
        OFN_SHAREWARN = 0
        OFN_SHOWHELP = &H10
        OFS_MAXPATHNAME = 128
    End Enum
    

'Return values from DriveGetType
    Public Enum eDriveType
        DRIVE_UNKNOWN
        DRIVE_ABSENT
        DRIVE_REMOVABLE
        DRIVE_FIXED
        DRIVE_REMOTE
        DRIVE_CDROM
        DRIVE_RAMDISK
    End Enum

'A return value from DriveGetSpecs
    Public Enum eDriveFlags
        FS_CASE_SENSITIVE = 1
        FS_CASE_IS_PRESERVED = 2
        FS_UNICODE_STORED_ON_DISK = 4
        FS_PERSISTENT_ACLS = 8
    End Enum

'Folders accessible through PathGetSpecial
    Public Enum eSpecialFolders
        sfADMINTOOLS = &H30
        sfALTSTARTUP = &H1D  'The file system directory that corresponds to the user's nonlocalized Startup program group.
        sfAPPDATA = &H1A
        sfBITBUCKET = &HA   'The virtual folder containing the objects in the user's Recycle Bin.
        sfCDBURN_AREA = &H3B  ' Version 6.0. The file system directory acting as a staging area for files waiting to be written to CD. A typical path is C:\Documents and Settings\username\Local Settings\Application Data\Microsoft\CD Burning.
        sfCOMMON_ADMINTOOLS = &H2F   'Version 5.0. The file system directory containing administrative tools for all users of the computer.
        sfCOMMON_ALTSTARTUP = &H1E  'Valid only for Microsoft Windows NT® systems.
        sfCOMMON_APPDATA = &H23
        sfCOMMON_DESKTOPDIRECTORY = &H19  'Valid only for Windows NT systems.
        sfCOMMON_DOCUMENTS = &H2E   'Valid for Windows NT systems and Microsoft Windows® 95 and Windows 98 systems with Shfolder.dll installed.
        sfCOMMON_FAVORITES = &H1F
        sfCOMMON_MUSIC = &H35
        sfCOMMON_PICTURES = &H36
        sfCOMMON_PROGRAMS = &H17
        sfCOMMON_STARTMENU = &H16
        sfCOMMON_STARTUP = &H18
        sfCOMMON_TEMPLATES = &H2D
        sfCOMMON_VIDEO = &H37
        sfCONTROLS = &H3   'The virtual folder containing icons for the Control Panel applications.
        sfCOOKIES = &H21
        sfDESKTOP = &H0
        sfDESKTOPDIRECTORY = &H10  'The file system directory used to physically store file objects on the desktop (not to be confused with the desktop folder itself). A typical path is C:\Documents and Settings\username\Desktop.
        sfDRIVES = &H11  'The virtual folder representing My Computer, containing everything on the local computer: storage devices, printers, and Control Panel. The folder may also contain mapped network drives.
        sfFAVORITES = &H6
        sfFONTS = &H14
        sfHISTORY = &H22
        sfINTERNET = &H1
        sfINTERNET_CACHE = &H20
        sfLOCAL_APPDATA = &H1C
        sfMYDOCUMENTS = &HC
        sfMYMUSIC = &HD
        sfMYPICTURES = &H27
        sfMYVIDEO = &HE
        sfNETHOOD = &H13
        sfNETWORK = &H12
        sfPRINTERS = &H4
        sfPRINTHOOD = &H1B
        sfPROFILE = &H28
        sfPROFILES = &H3E
        sfPROGRAM_FILES = &H26
        sfPROGRAM_FILES_COMMON = &H2B
        sfPROGRAMS = &H2
        sfRECENT = &H8
        sfSENDTO = &H9
        sfSTARTMENU = &HB
        sfSTARTUP = &H7
        sfSystem = &H25
        sfTEMPLATES = &H15
        sfTemporary = -1
        sfWindows = &H24
    End Enum
    
'Enum used by FileGetIcon
    Public Enum eShellIconSizes
        siLarge
        siSmall
        siAuto
    End Enum

'Return types from PathGetCharType
    Public Enum ePathCharTypes
        PCT_INVALID = 0
        PCT_LFNCHAR = 1
        PCT_SHORTCHAR = 2
        PCT_WILD = 4
        PCT_SEPARATOR = 8
    End Enum

'Flags passed the the FileMove Function
    Public Enum eMoveFileFlags
        MOVEFILE_COPY_ALLOWED = 2
        MOVEFILE_DELAY_UNTIL_REBOOT = 4
        MOVEFILE_REPLACE_EXISTING = 1
    End Enum

'Flags for file attributes
    Public Enum eFileAttributes
        FILE_ATTRIBUTE_ARCHIVE = &H20
        FILE_ATTRIBUTE_DIRECTORY = &H10
        FILE_ATTRIBUTE_HIDDEN = &H2
        FILE_ATTRIBUTE_NORMAL = &H80
        FILE_ATTRIBUTE_READONLY = &H1
        FILE_ATTRIBUTE_SYSTEM = &H4
        FILE_ATTRIBUTE_TEMPORARY = &H100
        FILE_ATTRIBUTE_COMPRESSED = &H800
    End Enum


'#######PUBLIC TYPES################


'Types used by FindFiles
    Public Type t64BitBetween
        Low As Double
        High As Double
    End Type
    
    Public Type tFindFiles
        Path As String
        Filter As String
        Recurse As Boolean
        FileCount As Long
        DirCount As Long
        TotalFileSize As Variant
        IgnoreReadOnly As Boolean
        IgnoreSystem As Boolean
        IgnoreHidden As Boolean
        IgnoreTemp As Boolean
        Accessed As t64BitBetween
        Modified As t64BitBetween
        Created As t64BitBetween
        Size As t64BitBetween
    End Type
    
    Public Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
    End Type
    
    Public Type Win32FoundData
        Name As String
        Path As String
        Attributes As eFileAttributes
        Created As Date
        Accessed As Date
        Modified As Date
        Size As Double
    End Type
    
    Public Type WIN32_FIND_DATA
        dwFileAttributes As eFileAttributes
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
    End Type
    
'Type used by FileGetVersion
    Public Type tVersionInfo
        FileName As String
        Directory As String
        FileVer As String
        ProdVer As String
        FileFlags As String
        FileOS As String
        FileType As String
        FileSubType As String
    End Type

'########PRIVATE ENUMS###############

'Private enums used in API calls

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
    
    Private Enum eVersionFileFlags
        VS_FFI_SIGNATURE = &HFEEF04BD
        VS_FFI_STRUCVERSION = &H10000
        VS_FFI_FILEFLAGSMASK = &H3F&
        VS_FF_DEBUG = &H1
        VS_FF_PRERELEASE = &H2
        VS_FF_PATCHED = &H4
        VS_FF_PRIVATEBUILD = &H8
        VS_FF_INFOINFERRED = &H10
        VS_FF_SPECIALBUILD = &H20
    End Enum
    
    Private Enum eVersionFileOS
        VOS_UNKNOWN = &H0
        VOS_DOS = &H10000
        VOS_OS216 = &H20000
        VOS_OS232 = &H30000
        VOS_NT = &H40000
        VOS__BASE = &H0
        VOS__WINDOWS16 = &H1
        VOS__PM16 = &H2
        VOS__PM32 = &H3
        VOS__WINDOWS32 = &H4
        VOS_DOS_WINDOWS16 = &H10001
        VOS_DOS_WINDOWS32 = &H10004
        VOS_OS216_PM16 = &H20002
        VOS_OS232_PM32 = &H30003
        VOS_NT_WINDOWS32 = &H40004
    End Enum
    
    Private Enum eVersionFileTypes
        VFT_UNKNOWN = &H0
        VFT_APP = &H1
        VFT_DLL = &H2
        VFT_DRV = &H3
        VFT_FONT = &H4
        VFT_VXD = &H5
        VFT_STATIC_LIB = &H7
    End Enum
    
    Private Enum eVersionFileSubTypes
        VFT2_UNKNOWN = &H0
        VFT2_DRV_PRINTER = &H1
        VFT2_DRV_KEYBOARD = &H2
        VFT2_DRV_LANGUAGE = &H3
        VFT2_DRV_DISPLAY = &H4
        VFT2_DRV_MOUSE = &H5
        VFT2_DRV_NETWORK = &H6
        VFT2_DRV_SYSTEM = &H7
        VFT2_DRV_INSTALLABLE = &H8
        VFT2_DRV_SOUND = &H9
        VFT2_DRV_COMM = &HA
    End Enum

'#######PRIVATE TYPES###############
    Private Type OPENFILENAME
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        flags As eOFNFlags
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
    End Type

    Private Type SHITEMID
        cb As Long
        abID As Byte
    End Type
    Private Type ITEMIDLIST
        mkid As SHITEMID
    End Type

    Private Type BrowseInfo
        hwndOwner As Long
        pIDLRoot As Long
        pszDisplayName As String
        lpszTitle As String
        ulFlags As Long
        lpfnCallback As Long
        lParam As Long
        iImage As Long
    End Type

    Private Type VS_FIXEDFILEINFO
       dwSignature As Long
       dwStrucVersionl As Integer     '  e.g. = &h0000 = 0
       dwStrucVersionh As Integer     '  e.g. = &h0042 = .42
       dwFileVersionMSl As Integer    '  e.g. = &h0003 = 3
       dwFileVersionMSh As Integer    '  e.g. = &h0075 = .75
       dwFileVersionLSl As Integer    '  e.g. = &h0000 = 0
       dwFileVersionLSh As Integer    '  e.g. = &h0031 = .31
       dwProductVersionMSl As Integer '  e.g. = &h0003 = 3
       dwProductVersionMSh As Integer '  e.g. = &h0010 = .1
       dwProductVersionLSl As Integer '  e.g. = &h0000 = 0
       dwProductVersionLSh As Integer '  e.g. = &h0031 = .31
       dwFileFlagsMask As Long        '  = &h3F for version "0.42"
       dwFileFlags As eVersionFileFlags 'e.g. VFF_DEBUG Or VFF_PRERELEASE
       dwFileOS As eVersionFileOS     '  e.g. VOS_DOS_WINDOWS16
       dwFileType As eVersionFileTypes ' e.g. VFT_DRIVER
       dwFileSubtype As eVersionFileSubTypes 'e.g. VFT2_DRV_KEYBOARD
       dwFileDateMS As Long           '  e.g. 0
       dwFileDateLS As Long           '  e.g. 0
    End Type

    Private Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
    End Type
    
    Private Type SHFILEINFO
        hIcon As Long                      '  out: icon
        iIcon As Long          '  out: icon index
        dwAttributes As Long               '  out: SFGAO_ flags
        szDisplayName As String * MAX_PATH '  out: display name (or path)
        szTypeName As String * 80          '  out: type name
    End Type


'Alphabetical API Declares (All are private)
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal Length As Long)
Private Declare Sub PathCreateFromUrl Lib "shlwapi.dll" Alias "PathCreateFromUrlA" (ByVal pszUrl As String, ByVal pszPath As String, ByRef pcchPath As Long, ByVal dwFlags As Long)
Private Declare Sub PathQuoteSpacesAPI Lib "shlwapi.dll" Alias "PathQuoteSpacesA" (ByVal lpsz As String)
Private Declare Sub PathStripPath Lib "shlwapi.dll" Alias "PathStripPathA" (ByVal pszPath As String)
Private Declare Sub PathUnquoteSpacesAPI Lib "shlwapi.dll" Alias "PathUnquoteSpacesA" (ByVal lpsz As String)


Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function FindCloseAPI Lib "kernel32" Alias "FindClose" (ByVal hFindFile As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lplsFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As Currency, lpTotalNumberOfBytes As Currency, lpTotalNumberOfFreeBytes As Currency) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetOpenFileNameAPI Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileNameAPI Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long
Private Declare Function MoveFileEx Lib "kernel32" Alias "MoveFileExA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal dwFlags As Long) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function PathAddExtensionAPI Lib "shlwapi.dll" Alias "PathAddExtensionA" (ByVal pszPath As String, ByVal pszExt As String) As Long
Private Declare Function PathAppend Lib "shlwapi.dll" Alias "PathAppendA" (ByVal pszPath As String, ByVal pMore As String) As Long
Private Declare Function PathCanonicalize Lib "shlwapi.dll" Alias "PathCanonicalizeA" (ByVal pszBuf As String, ByVal pszPath As String) As Long
Private Declare Function PathCommonPrefix Lib "shlwapi.dll" Alias "PathCommonPrefixA" (ByVal pszFile1 As String, ByVal pszFile2 As String, ByVal achPath As String) As Long
Private Declare Function PathCompactPath Lib "shlwapi.dll" Alias "PathCompactPathA" (ByVal hDC As Long, ByVal pszPath As String, ByVal dx As Long) As Long
Private Declare Function PathCompactPathEx Lib "shlwapi.dll" Alias "PathCompactPathExA" (ByVal pszOut As String, ByVal pszSrc As String, ByVal cchMax As Long, ByVal dwFlags As Long) As Long
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Declare Function PathGetCharTypeAPI Lib "shlwapi.dll" Alias "PathGetCharTypeA" (ByVal ch As Byte) As Long
Private Declare Function PathGetDriveNumber Lib "shlwapi.dll" Alias "PathGetDriveNumberA" (ByVal pszPath As String) As Long
Private Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
Private Declare Function PathIsDirectoryEmpty Lib "shlwapi.dll" Alias "PathIsDirectoryEmptyA" (ByVal pszPath As String) As Long
Private Declare Function PathIsNetworkPathAPI Lib "shlwapi.dll" Alias "PathIsNetworkPathA" (ByVal pszPath As String) As Long
Private Declare Function PathIsPrefix Lib "shlwapi.dll" Alias "PathIsPrefixA" (ByVal pszPrefix As String, ByVal pszPath As String) As Long
Private Declare Function PathIsRootAPI Lib "shlwapi.dll" Alias "PathIsRootA" (ByVal pszPath As String) As Long
Private Declare Function PathIsURLAPI Lib "shlwapi.dll" Alias "PathIsURLA" (ByVal pszPath As String) As Long
Private Declare Function PathMatchSpec Lib "shlwapi.dll" Alias "PathMatchSpecA" (ByVal pszFile As String, ByVal pszSpec As String) As Long
Private Declare Function PathRelativePathTo Lib "shlwapi.dll" Alias "PathRelativePathToA" (ByVal pszPath As String, ByVal pszFrom As String, ByVal dwAttrFrom As Long, ByVal pszTo As String, ByVal dwAttrTo As Long) As Long
Private Declare Function PathStripToRoot Lib "shlwapi.dll" Alias "PathStripToRootA" (ByVal pszPath As String) As Long
Private Declare Function RemoveDirectory Lib "kernel32" Alias "RemoveDirectoryA" (ByVal lpPathName As String) As Long
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As eShellGetFileInfoFlags) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long

'Used for Shell Browse for Folder
Private Declare Function CoInitialize Lib "ole32" Alias "CoInitializeEx" (ByVal pvReserved As Long, ByVal dwCoInit As Long) As Long

Private Const COINIT_MULTITHREADED = &H0

Private Const BIF_RETURNONLYFSDIRS = 1&
Private Const BIF_USENEWUI = &H40
'Private Const BIF_DONTGOBELOWDOMAIN = &H2&     ' For starting the Find Computer
'Private Const BIF_STATUSTEXT = &H4&
'Private Const BIF_RETURNFSANCESTORS = &H8&
'Private Const BIF_EDITBOX = &H10&
'Private Const BIF_VALIDATE = &H20& ' insist on valid result (or CANCEL)
'Private Const BIF_BROWSEFORCOMPUTER = &H1000&  ' Browsing for Computers.
'Private Const BIF_BROWSEFORPRINTER = &H2000&   ' Browsing for Printers
'Private Const BIF_BROWSEINCLUDEFILES = &H4000& ' Browsing for Everything

Private Const WM_USER = &H400&
'// message from browser
Private Const BFFM_INITIALIZED = 1&
'Private Const BFFM_SELCHANGED = 2&
'Private Const BFFM_VALIDATEFAILEDA = 3&     '// lParam:szPath ret:1(cont),0(EndDialog)
'Private Const BFFM_VALIDATEFAILEDW = 4&    ' // lParam:wzPath ret:1(cont),0(EndDialog)

'// messages to browser
'Private Const BFFM_SETSTATUSTEXTA = (WM_USER + 100&)
'Private Const BFFM_ENABLEOK = (WM_USER + 101&)
Private Const BFFM_SETSELECTION = (WM_USER + 102&)
'Private Const BFFM_SETSTATUSTEXT = (WM_USER + 100&)
'Private Const BFFM_SETSELECTIONW = (WM_USER + 103&)
'Private Const BFFM_SETSTATUSTEXTW = (WM_USER + 104&)
Dim msInitFolder As String



'Incremented and used for unique temp filenames.
'Private miUnique As Long



'Additional declares that might become useful, but haven't yet.
'Private Declare Function PathBuildRoot Lib "shlwapi.dll" Alias "PathBuildRootA" (ByVal szRoot As String, ByVal iDrive As Long) As Long
'Private Declare Function PathIsUNCServerShare Lib "shlwapi.dll" Alias "PathIsUNCServerShareA" (ByVal pszPath As String) As Long
'Private Declare Function PathIsUNCServer Lib "shlwapi.dll" Alias "PathIsUNCServerA" (ByVal pszPath As String) As Long
'Private Declare Function PathIsUNC Lib "shlwapi.dll" Alias "PathIsUNCA" (ByVal pszPath As String) As Long
'Private Declare Function PathIsSystemFolder Lib "shlwapi.dll" Alias "PathIsSystemFolderA" (ByVal pszPath As String, ByVal dwAttrb As Long) As Long
'Private Declare Function PathIsSameRoot Lib "shlwapi.dll" Alias "PathIsSameRootA" (ByVal pszPath1 As String, ByVal pszPath2 As String) As Long
'Private Declare Function PathIsRelativeAPI Lib "shlwapi.dll" Alias "PathIsRelativeA" (ByVal pszPath As String) As Long Doesn't seem to work at all....
'Private Declare Function PathIsLFNFileSpec Lib "shlwapi.dll" Alias "PathIsLFNFileSpecA" (ByVal lpName As String) As Long ' Only works in a root directory...?????
'Private Declare Function PathAddBackslash Lib "shlwapi.dll" Alias "PathAddBackslashA" (ByVal pszPath As String) As Long
'Private Declare Function PathCombine Lib "shlwapi.dll" Alias "PathCombineA" (ByVal szDest As String, ByVal lpszDir As String, ByVal lpszFile As String) As Long
'Private Declare Function PathFindOnPath Lib "shlwapi.dll" Alias "PathFindOnPathA" (ByVal pszPath As String, ByVal ppszOtherDirs As String) As Boolean

Public Sub StripNulls(psString As String)
    Dim liPos As Long
    liPos = InStr(1, psString, vbNullChar)
    If liPos > 0 Then psString = Left$(psString, liPos - 1)
End Sub

Private Sub PrepareString(psString As String, Optional psValue As String)
    psString = String$(MAX_PATH, 0)
    Mid$(psString, 1, Len(psValue)) = psValue
End Sub


'#####PATH FUNCTIONS
Public Sub PathAddBackslash(psPath As String)
    If StrComp(Right$(psPath, 1), "\") <> 0 And LenB(psPath) > 0 Then psPath = psPath & "\"
End Sub

Public Function PathAddExtension(psPath As String, ByVal psExtension As String) As String
    PrepareString PathAddExtension, psPath
    Select Case Len(psExtension)
        Case 3
            If StrComp(Left$(psExtension, 1), ".") <> 0 Then
                psExtension = "." & psExtension
            Else
                psExtension = psExtension & " "
            End If
        Case Is > 4
            psExtension = Left$(psExtension, 4)
        Case Is < 3
            psExtension = psExtension & Space$(4 - Len(psExtension))
    End Select
    PathAddExtensionAPI PathAddExtension, psExtension
    StripNulls PathAddExtension
End Function

Public Function PathBuild(psPath As String, _
                          psMore As String) _
                As String
    PrepareString PathBuild, psPath
    PathAppend PathBuild, psMore
    StripNulls PathBuild
End Function

Public Function PathCompact(psPath As String, ByVal piChars As Long) As String
    PathCompact = String(Len(psPath), 0)
    PathCompactPathEx PathCompact, psPath, piChars, 0
    StripNulls PathCompact
End Function

Public Function PathCompactPixels(psPath As String, ByVal piHDC As Long, ByVal piPixels As Long) As String
    PrepareString PathCompactPixels, psPath
    PathCompactPath piHDC, PathCompactPixels, piPixels
    StripNulls PathCompactPixels
End Function

Public Function PathCreate(ByVal psBottomFolder As String) As Boolean
    PathAddBackslash psBottomFolder
    If PathIsRoot(psBottomFolder) Then
        If GetDriveType(psBottomFolder) = DRIVE_ABSENT Then Exit Function
    End If
    PathCreate = MakeSureDirectoryPathExists(psBottomFolder) <> 0
End Function

Public Function PathFromURL(psURL As String) As String
    PrepareString PathFromURL
    PathCreateFromUrl psURL, PathFromURL, MAX_PATH, 0
    StripNulls PathFromURL
End Function

Public Function PathGetAbsolute(psPath As String) As String
    PathGetAbsolute = String(MAX_PATH, vbNullChar)
    PathCanonicalize PathGetAbsolute, psPath & vbNullChar
    StripNulls PathGetAbsolute
End Function

Public Function PathGetBaseName(psPath As String) As String
    On Error Resume Next
    PathGetBaseName = PathGetFileName(psPath)
    PathGetBaseName = Left$(PathGetBaseName, InStrRev(PathGetBaseName, ".") - 1)
End Function

Public Function PathGetCharType(ByVal pyChar As Byte) As ePathCharTypes
    PathGetCharType = PathGetCharTypeAPI(pyChar)
End Function

Public Function PathGetCommonPrefix(psPath1 As String, psPath2 As String) As String
    PathGetCommonPrefix = String(MAX_PATH, 0)
    PathCommonPrefix psPath1, psPath2, PathGetCommonPrefix
    StripNulls PathGetCommonPrefix
End Function

Public Function PathGetExtension(psPath As String) As String
    Dim liVal As Long
    liVal = InStrRev(psPath, ".")
    If liVal <> 0 Then liVal = Len(psPath) - liVal
    PathGetExtension = LCase$(Right$(psPath, liVal))
End Function

Public Function PathGetFileName(psPath As String) As String
    PrepareString PathGetFileName, psPath
    PathStripPath PathGetFileName
    StripNulls PathGetFileName
End Function

Public Function PathGetParentFolder(psPath As String) As String
    PathGetParentFolder = PathGetFileName(psPath)
    PathGetParentFolder = Left$(psPath, Len(psPath) - Len(PathGetParentFolder))
End Function

Public Function PathGetFolder(psPath As String) As String
    On Error Resume Next
    Dim liTemp As Long
    liTemp = InStrRev(psPath, "\")
    PathGetFolder = Left$(psPath, liTemp - 1)
End Function

Public Function PathGetRelative(ByVal psDirFrom As String, _
                                ByVal psFileTo As String) _
                As String
    PathGetRelative = String(MAX_PATH, 0)
    psDirFrom = psDirFrom & String(100, 0)
    psFileTo = psFileTo & String(100, 0)
    
    PathRelativePathTo PathGetRelative, psDirFrom, FILE_ATTRIBUTE_DIRECTORY, psFileTo, FILE_ATTRIBUTE_NORMAL
    
    Dim liPos As Long
    liPos = InStr(1, PathGetRelative, vbNullChar)
    If liPos > 0 Then PathGetRelative = Left$(PathGetRelative, liPos - 1)
End Function

Public Function PathRemove(psPath As String) As Boolean
    PathRemove = RemoveDirectory(psPath) <> 0
End Function

Public Function PathGetRoot(psPath As String) As String
    PrepareString PathGetRoot, psPath
    PathStripToRoot PathGetRoot
    StripNulls PathGetRoot
End Function

Public Function PathGetSpecial(WhichOne As eSpecialFolders) As String
    PrepareString PathGetSpecial
    Dim liVal As Long
    Select Case WhichOne
        Case sfWindows
            liVal = GetWindowsDirectory(PathGetSpecial, MAX_PATH)
        Case sfTemporary
            liVal = GetTempPath(MAX_PATH, PathGetSpecial)
        Case sfSystem
            liVal = GetSystemDirectory(PathGetSpecial, MAX_PATH)
        Case Else
            Dim r As Long
            Dim IDL As ITEMIDLIST
            If SHGetSpecialFolderLocation(0, WhichOne, IDL) = 0 Then
                SHGetPathFromIDList ByVal IDL.mkid.cb, ByVal PathGetSpecial
                StripNulls PathGetSpecial
            Else
                PathGetSpecial = vbNullString
            End If
            Exit Function
    End Select
    PathGetSpecial = Left$(PathGetSpecial, liVal)
End Function

Public Function PathIsNetworkPath(psPath As String) As Boolean
    PathIsNetworkPath = PathIsNetworkPathAPI(psPath) <> 0
End Function

Public Function PathIsURL(psPath As String) As Boolean
    PathIsURL = PathIsURLAPI(psPath) <> 0
End Function

Public Function PathMatchRoot(psPath1 As String, psPath2 As String) As Boolean
    Dim liTemp As Long
    liTemp = PathGetDriveNumber(psPath1)
    If liTemp >= -1 Then PathMatchRoot = liTemp = PathGetDriveNumber(psPath2)
End Function

Public Function PathMatchPattern(psPath As String, psPattern As String) As Boolean
    PathMatchPattern = PathMatchSpec(psPath, psPattern) <> 0
End Function

Public Function PathMatchPrefix(psPath As String, psPrefix As String) As Boolean
    If Right$(psPrefix, 1) <> "\" Then
        PathMatchPrefix = PathIsPrefix(psPrefix, psPath) <> 0
    Else
        PathMatchPrefix = PathIsPrefix(Left$(psPrefix, Len(psPrefix) - 1), psPath) <> 0
        If Not PathMatchPrefix Then
            If Len(psPrefix) <= 3 Then
                PathMatchPrefix = PathMatchRoot(psPrefix, psPath)
            End If
        End If
    End If
End Function

Public Function PathQuoteSpaces(psPath As String) As String
    PrepareString PathQuoteSpaces, psPath
    PathQuoteSpacesAPI PathQuoteSpaces
    StripNulls PathQuoteSpaces
End Function

Public Function PathUnquoteSpaces(psPath As String) As String
    PrepareString PathUnquoteSpaces, psPath
    PathUnquoteSpacesAPI PathUnquoteSpaces
    StripNulls PathUnquoteSpaces
End Function

'Public Sub test()
'    Dim ltFind As tFindFiles
'    Dim loColl As Collection
'    With ltFind
'        .Path = "C:\"
'        .Size.High = 98000
'        .Size.Low = 95000
'        .Filter = "*"
'        .Recurse = True
'    End With
'    Set loColl = FindFiles(ltFind)
'    Stop
'End Sub

Public Function FileDelete(psPath As String, _
            Optional ByVal pbForceIfReadOnly As Boolean = False) _
                As Boolean
    If pbForceIfReadOnly Then
        Dim liVal As eFileAttributes
        liVal = FileGetAttributes(psPath)
        If liVal And FILE_ATTRIBUTE_READONLY Then
            liVal = liVal - FILE_ATTRIBUTE_READONLY
            FileSetAttributes psPath, liVal
        End If
    End If
    FileDelete = DeleteFile(psPath) <> 0
End Function

Public Function FileExists(psPath As String) As Boolean
    FileExists = PathFileExists(psPath) <> 0
    If FileExists Then FileExists = Not FolderExists(psPath)
    If FileExists Then FileExists = Not DriveExists(psPath)
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

Public Function FileNameIsLegal(psName As String) As Boolean
    On Error GoTo ending
    Dim lyBytes() As Byte
    Dim I As Long
    lyBytes = StrConv(PathGetFileName(psName), vbFromUnicode)
    For I = LBound(lyBytes) To UBound(lyBytes)
        If Not PathGetCharType(lyBytes(I)) And PCT_LFNCHAR Then Exit Function
    Next
    FileNameIsLegal = True
ending:
End Function

Public Function FileGetAttributes(psFile As String) As eFileAttributes
    FileGetAttributes = GetFileAttributes(psFile)
    If FileGetAttributes < 0 Then FileGetAttributes = 0
End Function

Public Function FileGetTempName(Optional ByVal psPath As String, Optional psPrefix As String = "tmp", Optional ByVal psExt As String = "tmp") As String
    Dim ltTime As SYSTEMTIME
    Dim liLong As Long
    On Error Resume Next
    GetSystemTime ltTime
    Randomize Timer
    If Len(psPath) = 0 Then psPath = PathGetSpecial(sfTemporary)
    PathAddBackslash psPath
    psPath = psPath & psPrefix
    If Asc(Left$(psExt, 1)) <> vbKeyDelete Then psExt = "." & psExt
    Do
        liLong = TotalTime(ltTime)
        FileGetTempName = psPath & Hex$(liLong) & psExt
        ltTime.wDay = ltTime.wDay + Rnd * 100 - 50
    Loop While FileExists(FileGetTempName) Or TempNameGiven(FileGetTempName)
    If mCollTempNames Is Nothing Then Set mCollTempNames = New Collection
    mCollTempNames.Add "", FileGetTempName
    If mCollTempNames.Count > 32767 Then mCollTempNames.Remove 1
End Function

Public Function PathGetTempFolderName(Optional ByVal psInFolder As String, Optional psPrefix As String = "tmp") As String
    Dim ltTime As SYSTEMTIME
    Dim liLong As Long
    On Error Resume Next
    If Len(psInFolder) = 0 Then psInFolder = PathGetSpecial(sfTemporary)
    GetSystemTime ltTime
    Randomize Timer
    PathAddBackslash psInFolder
    psInFolder = psInFolder & psPrefix
    Do
        liLong = TotalTime(ltTime)
        PathGetTempFolderName = psInFolder & Hex$(liLong) & "\"
        ltTime.wDay = ltTime.wDay + Rnd * 100 - 50
    Loop While FolderExists(PathGetTempFolderName) Or TempNameGiven(PathGetTempFolderName)
    If mCollTempNames Is Nothing Then Set mCollTempNames = New Collection
    mCollTempNames.Add "", PathGetTempFolderName
    If mCollTempNames.Count > 32767 Then mCollTempNames.Remove 1
End Function

Private Function TempNameGiven(psName As String) As Boolean
    On Error Resume Next
    IsObject mCollTempNames(psName)
    TempNameGiven = Err.Number = 0
End Function

Private Function TotalTime(ptTime As SYSTEMTIME) As Long
    With ptTime
        TotalTime = .wDay + .wDayOfWeek + .wHour + .wMilliseconds + .wMinute + .wMonth + .wSecond + .wYear
    End With
End Function
Public Function FileGetTypeName(psFile As String, psTypeName As String) As Boolean
    Dim ltInfo As SHFILEINFO
    FileGetTypeName = SHGetFileInfo(psFile, 0, ltInfo, Len(ltInfo), SHGFI_TYPENAME) <> 0
    psTypeName = ltInfo.szTypeName
    StripNulls psTypeName
End Function

Public Function FileGetIcon(psFile As String, hImageList As Long, IconIndex As Long, ByVal piSize As eShellIconSizes) As Boolean
    Dim ltInfo As SHFILEINFO
    Dim liFlags As eShellGetFileInfoFlags
    Select Case piSize
        Case siSmall
            liFlags = SHGFI_SMALLICON
        Case siLarge
            liFlags = SHGFI_LARGEICON
        Case siAuto
            liFlags = SHGFI_SHELLICONSIZE
    End Select
    liFlags = liFlags 'Or SHGFI_SYSICONINDEX
    hImageList = SHGetFileInfo(psFile, 0, ltInfo, Len(ltInfo), SHGFI_SYSICONINDEX Or liFlags)
    FileGetIcon = hImageList <> 0
    IconIndex = ltInfo.iIcon
End Function

Public Function FileGetLen(psFile As String) As Double
    Dim hFile As Long
    On Error Resume Next
    Dim liLong As Long
    If FileCreate(psFile, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, OPEN_EXISTING, 0, 0, hFile) Then
        FileGetLen = FileGetSize(hFile)
        FileClose hFile
    End If
End Function

Public Function FileGetTime(psFile As String, Optional pdModified As Date, Optional pdCreated As Date, Optional pdAccessed As Date) As Boolean
    Dim hFile As Long
    Dim ltAccessed As FILETIME, ltModified As FILETIME, ltCreated As FILETIME
    Dim ltSystem As SYSTEMTIME

    If Not FileCreate(psFile, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, 0, hFile) Then Exit Function
    FileGetTime = GetFileTime(hFile, ltCreated, ltAccessed, ltModified) <> 0
    FileClose hFile
    FileTimetoDate ltAccessed, pdAccessed
    FileTimetoDate ltModified, pdModified
    FileTimetoDate ltCreated, pdCreated
End Function

Public Function FileGetVersion(ptVI As tVersionInfo) As Boolean
    Dim lDummy As Long, sBuffer() As Byte
    Dim lBufferLen As Long, lVerPointer As Long, udtVerBuffer As VS_FIXEDFILEINFO
    Dim lVerbufferLen As Long
    Dim lsFullName As String
   '*** Get size ****
    With ptVI
        lsFullName = .Directory & .FileName
    
        lBufferLen = GetFileVersionInfoSize(lsFullName, lDummy)
        If lBufferLen < 1 Then Exit Function Else FileGetVersion = True
        
        '**** Store info to udtVerBuffer struct ****
        ReDim sBuffer(lBufferLen)
        GetFileVersionInfo lsFullName, 0&, lBufferLen, sBuffer(0)
        VerQueryValue sBuffer(0), "\", lVerPointer, lVerbufferLen
        MoveMemory udtVerBuffer, lVerPointer, Len(udtVerBuffer)
        
        '**** Determine Structure Version number - NOT USED ****
        'StrucVer = Format$(udtVerBuffer.dwStrucVersionh) & "." & Format$(udtVerBuffer.dwStrucVersionl)
        
        '**** Determine File Version number ****
        .FileVer = Format$(udtVerBuffer.dwFileVersionMSh) & "." & Format$(udtVerBuffer.dwFileVersionMSl) & "." & Format$(udtVerBuffer.dwFileVersionLSh) & "." & Format$(udtVerBuffer.dwFileVersionLSl)
        
        '**** Determine Product Version number ****
        .ProdVer = Format$(udtVerBuffer.dwProductVersionMSh) & "." & Format$(udtVerBuffer.dwProductVersionMSl) & "." & Format$(udtVerBuffer.dwProductVersionLSh) & "." & Format$(udtVerBuffer.dwProductVersionLSl)
        
        '**** Determine Boolean attributes of File ****
        .FileFlags = ""
        If udtVerBuffer.dwFileFlags And VS_FF_DEBUG Then .FileFlags = "Debug "
        If udtVerBuffer.dwFileFlags And VS_FF_PRERELEASE Then .FileFlags = .FileFlags & "PreRel "
        If udtVerBuffer.dwFileFlags And VS_FF_PATCHED Then .FileFlags = .FileFlags & "Patched "
        If udtVerBuffer.dwFileFlags And VS_FF_PRIVATEBUILD Then .FileFlags = .FileFlags & "Private "
        If udtVerBuffer.dwFileFlags And VS_FF_INFOINFERRED Then .FileFlags = .FileFlags & "Info Inferred "
        If udtVerBuffer.dwFileFlags And VS_FF_SPECIALBUILD Then .FileFlags = .FileFlags & "Special Build "
        If udtVerBuffer.dwFileFlags And VFT2_UNKNOWN Then .FileFlags = .FileFlags & "Unknown "
        
        '**** Determine OS for which file was designed ****
        Select Case udtVerBuffer.dwFileOS
            Case VOS_DOS_WINDOWS16
                .FileOS = "DOS-Win16"
            Case VOS_DOS_WINDOWS32
                .FileOS = "DOS-Win32"
            Case VOS_OS216_PM16
                .FileOS = "OS/2-16 PM-16"
            Case VOS_OS232_PM32
                .FileOS = "OS/2-16 PM-32"
            Case VOS_NT_WINDOWS32
                .FileOS = "NT-Win32"
            Case Else
                .FileOS = "Unknown"
        End Select
        Select Case udtVerBuffer.dwFileType
            Case VFT_APP
                .FileType = "App"
            Case VFT_DLL
                .FileType = "DLL"
            Case VFT_DRV
                .FileType = "Driver"
                Select Case udtVerBuffer.dwFileSubtype
                    Case VFT2_DRV_PRINTER
                        .FileSubType = "Printer drv"
                    Case VFT2_DRV_KEYBOARD
                        .FileSubType = "Keyboard drv"
                    Case VFT2_DRV_LANGUAGE
                        .FileSubType = "Language drv"
                    Case VFT2_DRV_DISPLAY
                        .FileSubType = "Display drv"
                    Case VFT2_DRV_MOUSE
                        .FileSubType = "Mouse drv"
                    Case VFT2_DRV_NETWORK
                        .FileSubType = "Network drv"
                    Case VFT2_DRV_SYSTEM
                        .FileSubType = "System drv"
                    Case VFT2_DRV_INSTALLABLE
                        .FileSubType = "Installable"
                    Case VFT2_DRV_SOUND
                        .FileSubType = "Sound drv"
                    Case VFT2_DRV_COMM
                        .FileSubType = "Comm drv"
                    Case VFT2_UNKNOWN
                        .FileSubType = "Unknown"
                End Select
           Case VFT_FONT
                .FileType = "Font"
                Select Case udtVerBuffer.dwFileSubtype
                    'Case VFT_FONT_RASTER
                        '.FileSubType = "Raster Font"
                    'Case VFT_FONT_VECTOR
                        '.FileSubType = "Vector Font"
                    'Case VFT_FONT_TRUETYPE
                        '.FileSubType = "TrueType Font"
                    Case Else
                        Debug.Print "FONT!!!"
                        Debug.Print .FileName & udtVerBuffer.dwFileSubtype
                End Select
           Case VFT_VXD
                .FileType = "VxD"
           Case VFT_STATIC_LIB
                .FileType = "Lib"
           Case Else
                .FileType = "Unknown"
        End Select
    End With
End Function

Public Function FileMove(psFrom As String, _
                         psTo As String, _
          Optional ByVal piFlags As eMoveFileFlags = MOVEFILE_REPLACE_EXISTING) _
                As Boolean
    If Not PathCreate(PathGetParentFolder(psTo)) Then Exit Function
    FileMove = MoveFileEx(psFrom, psTo, piFlags) <> 0
End Function

Public Function FileSetAttributes(psFile As String, ByVal piVal As eFileAttributes) As Boolean
    FileSetAttributes = SetFileAttributes(psFile, piVal) <> 0
End Function

Public Sub FileTimetoDate(ptFileTime As FILETIME, pdDate As Date)
    Dim ltTime As FILETIME
    Dim ltSystem As SYSTEMTIME
    FileTimeToLocalFileTime ptFileTime, ltTime
    FileTimeToSystemTime ltTime, ltSystem
    With ltSystem
        pdDate = DateSerial(.wYear, .wMonth, .wDay) + TimeSerial(.wHour, .wMinute, .wSecond)
    End With
End Sub


'#######FOLDER FUNCTIONS#################
Public Function FolderDelete(psPath As String, _
             Optional ByVal pbForceIfReadOnly As Boolean = False) _
                As Boolean
    Dim loColl As Collection
    Dim lvTemp
    Dim I As Long
    If Not FolderExists(psPath) Then Exit Function
    Dim ltFind As tFindFiles
    With ltFind
        .Path = psPath
        .Filter = "*"
        .Recurse = True
    End With
    Set loColl = FindFiles(ltFind)
    For Each lvTemp In loColl
        If Not FileDelete(CStr(lvTemp), pbForceIfReadOnly) Then Exit Function
    Next
    Set loColl = FindFolders(psPath, False)
    For I = 1 To loColl.Count
        If Not FolderDelete(CStr(loColl.Item(I))) Then Exit Function
    Next
    FolderDelete = RemoveDirectory(psPath) <> 0
End Function

Public Function FolderExists(psPath As String) As Boolean
    FolderExists = PathIsDirectory(psPath) <> 0
End Function

Public Function FolderIsEmpty(psPath As String) As Boolean
    FolderIsEmpty = PathIsDirectoryEmpty(psPath)
End Function

Public Function FolderMove(psFrom As String, _
                           psTo As String, _
            Optional ByVal pbOverwrite As Boolean = True) _
                As Boolean
    Dim liFlags As eMoveFileFlags
    If Not FolderExists(psFrom) Then Exit Function
    If pbOverwrite Then liFlags = MOVEFILE_REPLACE_EXISTING
    FolderMove = MoveFileEx(psFrom, psTo, liFlags) <> 0
End Function


'#######DRIVE FUNCTIONS#################
Public Function DriveExists(psPath As String) As Boolean
    DriveExists = DriveGetType(psPath) > DRIVE_ABSENT
End Function

Public Function DriveGetStrings(psStrings() As String) As Boolean
    Dim lsTemp As String
    On Error Resume Next
    lsTemp = String(MAX_PATH, 0)
    GetLogicalDriveStrings MAX_PATH, lsTemp
    lsTemp = Left$(lsTemp, InStr(1, lsTemp, vbNullChar & vbNullChar) - 1)
    psStrings = Split(lsTemp, vbNullChar)
    
    Dim I As Long
    Err.Clear
    I = UBound(psStrings)
    DriveGetStrings = Err.Number = 0
    
End Function

Public Function DriveGetSpace(psPath As String, Optional pdblFreeBytes As Double, Optional pdblTotalBytes As Double) As Boolean
    Dim lcFreeToCaller As Currency, lcTotal As Currency, lcFree As Currency
    DriveGetSpace = GetDiskFreeSpaceEx(psPath, lcFreeToCaller, lcTotal, lcFree) <> 0
    If lcFreeToCaller < lcFree Then lcFree = lcFreeToCaller
    pdblFreeBytes = lcFree * 10000
    pdblTotalBytes = lcTotal * 1000
End Function

Public Function DriveGetSpecs(psPath As String, Optional psVolumeName As String, Optional piSerialNumber As Long, Optional piMaxFileNameLength As Long, Optional piFlags As eDriveFlags, Optional psFileSystemName As String) As Boolean
    Dim Serial As Long, VName As String, FSName As String
    PrepareString psVolumeName
    psFileSystemName = psVolumeName
    DriveGetSpecs = GetVolumeInformation(psPath, psVolumeName, MAX_PATH, piSerialNumber, piMaxFileNameLength, piFlags, psFileSystemName, MAX_PATH) <> 0
    StripNulls psVolumeName
    StripNulls psFileSystemName
End Function

Public Function DriveGetType(psPath As String) As eDriveType
    DriveGetType = GetDriveType(psPath)
End Function


'###Find Files/Folders###############
Public Function FindFiles(ptFind As tFindFiles, Optional ByVal poColl As Collection) As Collection
    If poColl Is Nothing Then Set poColl = New Collection
    Set FindFiles = poColl
    On Error Resume Next
    Dim ltWin32 As WIN32_FIND_DATA
    Dim liDirs As Long, I As Long, hSearch As Long, liStayInLoop As Long
    Dim lsFileName As String, lsDirNames() As String
    Dim lbCountedDir As Boolean
    Dim liSize As Long

    With ptFind
        PathAddBackslash .Path
        liDirs = 0
        ReDim lsDirNames(0 To liDirs)
        liStayInLoop = True
        hSearch = FindFirstFile(.Path & "*", ltWin32)
        If hSearch <> INVALID_HANDLE_VALUE Then
            Do While liStayInLoop
                lsFileName = ltWin32.cFileName
                StripNulls lsFileName
                If StrComp(lsFileName, ".") <> 0 And StrComp(lsFileName, "..") <> 0 Then
                    If ltWin32.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then
                        If .Recurse Then
                            lsDirNames(liDirs) = lsFileName
                            liDirs = liDirs + 1
                            ReDim Preserve lsDirNames(0 To liDirs)
                        End If
                    Else
                        If ValidateFind(ltWin32, ptFind) Then
                            liSize = MakeQWord(ltWin32.nFileSizeLow, ltWin32.nFileSizeHigh)
                            .TotalFileSize = .TotalFileSize + liSize
                            .FileCount = .FileCount + 1
                            If Not lbCountedDir Then
                                lbCountedDir = True
                                .DirCount = .DirCount + 1
                            End If
                            lsFileName = .Path & lsFileName
                            poColl.Add lsFileName, lsFileName
                        End If
                    End If
                End If
                liStayInLoop = FindNextFile(hSearch, ltWin32)
            Loop
            liStayInLoop = FindCloseAPI(hSearch)
        End If

        If liDirs > 0 Then
            Dim lsPath As String
            lsPath = .Path
            For I = 0 To liDirs - 1
                .Path = lsPath & lsDirNames(I) & "\"
                FindFiles ptFind, poColl
            Next I
            .Path = lsPath
        End If
    End With
End Function

Public Function ValidateFind(ptWin32 As WIN32_FIND_DATA, ptFind As tFindFiles) As Boolean
    With ptFind
        Dim ldDate As Date

        ValidateFind = PathMatchPattern(ptWin32.cFileName, ptFind.Filter)
        If Not ValidateFind Then Exit Function
        
        FileTimetoDate ptWin32.ftLastAccessTime, ldDate
        ValidateFind = BetweenVals(ldDate, .Accessed)
        If Not ValidateFind Then Exit Function
        
        FileTimetoDate ptWin32.ftCreationTime, ldDate
        ValidateFind = BetweenVals(ldDate, .Created)
        If Not ValidateFind Then Exit Function
        
        FileTimetoDate ptWin32.ftLastWriteTime, ldDate
        ValidateFind = BetweenVals(ldDate, .Modified)
        If Not ValidateFind Then Exit Function
        
        ValidateFind = BetweenVals(MakeQWord(ptWin32.nFileSizeLow, ptWin32.nFileSizeHigh), .Size)
        If Not ValidateFind Then Exit Function
        ValidateFind = False
        If .IgnoreHidden Then
            If ptWin32.dwFileAttributes And FILE_ATTRIBUTE_HIDDEN Then Exit Function
        End If
        If .IgnoreReadOnly Then
            If ptWin32.dwFileAttributes And FILE_ATTRIBUTE_READONLY Then Exit Function
        End If
        If .IgnoreSystem Then
            If ptWin32.dwFileAttributes And FILE_ATTRIBUTE_SYSTEM Then Exit Function
        End If
        If .IgnoreTemp Then
            If ptWin32.dwFileAttributes And FILE_ATTRIBUTE_TEMPORARY Then Exit Function
        End If
        ValidateFind = True
    End With
End Function

Private Function BetweenVals(ByVal pdblVal As Double, ptBetween As t64BitBetween) As Boolean
    BetweenVals = True
    With ptBetween
        If .High = 0 Then
            If .Low <> 0 Then BetweenVals = pdblVal > .Low
        Else
            If .Low = 0 Then
                BetweenVals = pdblVal < .High
            Else
                BetweenVals = pdblVal < .High And pdblVal > .Low
            End If
        End If
    End With
End Function

Public Function FindFolders(Path As String, Optional pbRecurse As Boolean = True, Optional poColl As Collection) As Collection
    If poColl Is Nothing Then Set poColl = New Collection
    Dim loColl As Collection
    Set loColl = New Collection
    
    Set FindFolders = poColl
    
    Dim ltFind       As WIN32_FIND_DATA

    Dim hSearch      As Long
    Dim liStayInLoop As Long
    
    Dim lsDirName    As String
    Dim lsFileName   As String
    Dim lvTemp
    
    PathAddBackslash Path

    liStayInLoop = True
    hSearch = FindFirstFile(Path & "*", ltFind)
    If hSearch <> INVALID_HANDLE_VALUE Then
        Do While liStayInLoop
            lsDirName = ltFind.cFileName
            StripNulls lsDirName
            If StrComp(lsDirName, ".") <> 0 And StrComp(lsDirName, "..") <> 0 Then
                lsFileName = Path & lsDirName
                If FileGetAttributes(lsFileName) And FILE_ATTRIBUTE_DIRECTORY Then
                    poColl.Add lsFileName, lsFileName
                    loColl.Add lsFileName
                End If
            End If
            liStayInLoop = FindNextFile(hSearch, ltFind)
        Loop
        liStayInLoop = FindCloseAPI(hSearch)
    End If

    If poColl.Count > 0 And pbRecurse Then
        For Each lvTemp In loColl
            FindFolders PathBuild(Path, CStr(lvTemp)) & "\", pbRecurse, poColl
        Next
    End If
End Function

Public Function FindFirst(ByVal psFile As String, ByVal pbFindFolders As Boolean) As String
    Dim ltFind       As WIN32_FIND_DATA

    Dim hSearch      As Long
    Dim lsPath As String
    Dim lsOrigPattern As String
    
    If Len(psFile) = 0 Then Exit Function
    If Not FolderExists(PathGetParentFolder(psFile)) Then Exit Function
    
    If Right$(psFile, 1) = "\" Then
        lsPath = psFile
    Else
        lsPath = PathGetParentFolder(psFile)
        If Len(lsPath) = 0 Then lsPath = PathGetRoot(psFile)
        PathAddBackslash lsPath
    End If
    lsOrigPattern = psFile & "*"
    hSearch = FindFirstFile(lsOrigPattern, ltFind)
    lsOrigPattern = PathGetFileName(lsOrigPattern)
    If hSearch <> INVALID_HANDLE_VALUE Then
        Do
            psFile = ltFind.cFileName
            StripNulls psFile
            If StrComp(psFile, ".") <> 0 And StrComp(psFile, "..") <> 0 And PathMatchPattern(psFile, lsOrigPattern) Then
                psFile = lsPath & psFile
                If ltFind.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then
                    'If pbFindFolders Then
                    FindFirst = psFile
                    PathAddBackslash FindFirst
                    Exit Do
                    'End If
                Else
                    If Not pbFindFolders Then
                        FindFirst = psFile
                        Exit Do
                    End If
                End If
            End If
        Loop While FindNextFile(hSearch, ltFind)
        FindCloseAPI hSearch
    End If
End Function

Public Function FindSpecific(psPath As String, ptData As WIN32_FIND_DATA, Optional phSearch As Long) As Boolean
    Dim hSearch      As Long
    hSearch = FindFirstFile(psPath, ptData)
    If hSearch <> INVALID_HANDLE_VALUE Then
        If hSearch = 0 Then FindCloseAPI hSearch Else phSearch = hSearch
        FindSpecific = True
    End If
End Function

Public Function FindNext(phSearch As Long, ptData As WIN32_FIND_DATA) As Boolean
    FindNext = FindNextFile(phSearch, ptData) <> 0
End Function

Public Sub FindClose(ByVal hSearch As Long)
    FindCloseAPI hSearch
End Sub

Public Sub FindToFriendlyType(psPath As String, ptWin32 As WIN32_FIND_DATA, ptFriendly As Win32FoundData)
    With ptFriendly
        .Path = psPath
        .Name = ptWin32.cFileName
        StripNulls .Name
        FileTimetoDate ptWin32.ftLastAccessTime, .Accessed
        FileTimetoDate ptWin32.ftCreationTime, .Created
        FileTimetoDate ptWin32.ftLastWriteTime, .Modified
        .Attributes = ptWin32.dwFileAttributes
        .Size = MakeQWord(ptWin32.nFileSizeLow, ptWin32.nFileSizeHigh)
    End With
End Sub



'#######Browse for Folder from all-api.net (Modified)##########
Private Function GetCBaddr(ByVal lAddr As Long) As Long
    GetCBaddr = lAddr
End Function

Private Function BFCallBack(ByVal Hwnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
    On Error Resume Next
    If uMsg = BFFM_INITIALIZED Then
        SendMessage Hwnd, BFFM_SETSELECTION, 1&, ByVal msInitFolder
    End If
End Function

Public Function BrowseForFolder(piHwnd As Long, Optional psPrompt As String, Optional psPath As String = vbNullString, Optional psCaption As String = vbNullString) As String
    Dim BINF As BrowseInfo
    Dim lpItem As Long
    Dim sDir As String
    Dim iLen As Integer
    On Error Resume Next
    CoInitialize vbNull, COINIT_MULTITHREADED
    msInitFolder = psPath
    If msInitFolder <> "" Then
        iLen = Len(msInitFolder)
        If iLen >= 2 Then
            If Right(msInitFolder, 1) = "\" Then
                If iLen > 3 Then msInitFolder = Left$(msInitFolder, iLen - 1) 'not root - remove "\"
            Else
                If iLen < 3 Then msInitFolder = msInitFolder & "\" ' root
            End If
            Err.Clear
            If FolderExists(msInitFolder) Then
                If Err.Number = 0 Then BINF.lpfnCallback = GetCBaddr(AddressOf BFCallBack)
            End If
        End If
    End If
    
    With BINF
        .hwndOwner = piHwnd
        '.pidlRoot = 0& 'open at desktop
        '.pidlRoot = CSIDL_DRIVES 'skips desktop
        .pszDisplayName = Space$(MAX_PATH + 1)
        .lpszTitle = psPrompt
        .ulFlags = BIF_RETURNONLYFSDIRS Or BIF_USENEWUI
    End With
    lpItem = SHBrowseForFolder(BINF)
    If lpItem Then
        sDir = Space$(MAX_PATH)
        If SHGetPathFromIDList(ByVal lpItem, sDir) Then
            BrowseForFolder = Left(sDir, InStr(1, sDir, vbNullChar) - 1)
        End If
    End If
    GlobalFree lpItem
    msInitFolder = ""
End Function

Private Sub GetOFN(ByRef OFN As OPENFILENAME, ByVal piHwnd As Long, psTitle As String, psInitDir As String, psFilter As String, ByVal piFlags As Long)
    With OFN
        .lStructSize = Len(OFN)
        .hwndOwner = piHwnd
        .hInstance = App.hInstance
        .lpstrFilter = psFilter
        .lpstrFile = Space$(254)
        .nMaxFile = 255
        .lpstrFileTitle = Space$(254)
        .nMaxFileTitle = 255
        .lpstrInitialDir = psInitDir
        .lpstrTitle = psTitle
        .flags = piFlags
    End With
End Sub

Public Function CommonDialogFilter(ParamArray pvArgs()) As String
    On Error Resume Next
    Dim I As Long
    For I = LBound(pvArgs) To UBound(pvArgs)
        CommonDialogFilter = CommonDialogFilter & pvArgs(I) & vbNullChar
    Next
End Function

Public Function GetSaveFileName(ByVal Hwnd As Long, psTitle As String, psInitDir As String, psFilter As String, psDefExtension As String, piFlags As eOFNFlags) As String
    On Error Resume Next
    Dim ltOFN As OPENFILENAME
    GetOFN ltOFN, Hwnd, psTitle, psInitDir, psFilter, piFlags
    If GetSaveFileNameAPI(ltOFN) <> 0 Then
        GetSaveFileName = ltOFN.lpstrFile
        StripNulls GetSaveFileName
    End If
    If Len(GetSaveFileName) > 0 Then GetSaveFileName = PathAddExtension(GetSaveFileName, psDefExtension)
End Function

Public Function GetOpenFileNames(ByVal poColl As Collection, ByVal Hwnd As Long, psTitle As String, psInitDir As String, psFilter As String, psDefExtension As String, piFlags As eOFNFlags) As Boolean
    On Error Resume Next
    Dim ltOFN As OPENFILENAME
    GetOFN ltOFN, Hwnd, psTitle, psInitDir, psFilter, piFlags
    GetOpenFileNames = GetOpenFileNameAPI(ltOFN) <> 0
    If GetOpenFileNames Then
        Dim I As Long
        Dim lsNames() As String
        Dim lsPath As String
        lsNames = Split(ltOFN.lpstrFile, vbNullChar)
        lsPath = lsNames(0)
        If Len(Trim$(lsNames(1))) = 0 Then
            poColl.Add lsPath
        Else
            Dim lsName As String
            For I = 1 To UBound(lsNames)
                If Len(lsNames(I)) = 0 Then Exit For
                lsName = PathAddExtension(lsPath & lsNames(I), psDefExtension)
                poColl.Add lsName, lsName
            Next
        End If
    End If
End Function

Public Function PathIsRoot(psPath As String) As Boolean
    PathIsRoot = PathIsRootAPI(psPath) <> 0
End Function
