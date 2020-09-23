Attribute VB_Name = "StreamModule"
Option Explicit
'* This module contains the functions needed to enumerate, read and write to NTFS Alternate Data Streams
'* NTFS is only supported natively by NT systems (including XP/.Net) and since build 3.51
'* There are other ways to enumerate and manipulate NTFS ADS but this undocumented way is the most compatible
'* Ok now you must be wondering why I use both methods, and don't just dump the Kernel way...the truth is,
'* if you want to do a system-wide security scan for streams, the kernel APIs are much much faster then using the
'* backup APIs. I've included a scan function which skips past the normal DATA stream, and just displays any files that contain streams.
'* If you see anyhting suspicious, you can use the backup APIs to investigate more in detail.

' APIs we use for the Stream Manipulation
' API Declares for Kernel Seeking of the Streams
Public Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Public Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function NtQueryInformationFile Lib "NTDLL.DLL" (ByVal FileHandle As Long, IoStatusBlock_Out As IO_STATUS_BLOCK, lpFileInformation_Out As Long, ByVal Length As Long, ByVal FileInformationClass As FILE_INFORMATION_CLASS) As Long
Public Declare Function GetVolumeInformation Lib "kernel32.dll" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
' API to read/write and use Backup API to get Streams
Public Declare Function BackupRead Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByVal bAbort As Long, ByVal bProcessSecurity As Long, ByRef lpContext As Long) As Long
Public Declare Function BackupSeek Lib "kernel32" (ByVal hFile As Long, ByVal dwLowBytesToSeek As Long, ByVal dwHighBytesToSeek As Long, ByRef lpdwLowByteSeeked As Long, ByRef lpdwHighByteSeeked As Long, ByRef lpContext As Long) As Long
Public Declare Function DeleteFile Lib "kernel32.dll" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Public Declare Function GetFileSize Lib "kernel32.dll" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Public Declare Function ReadFile Lib "kernel32.dll" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Public Declare Function SetFilePointer Lib "kernel32.dll" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
' API to Find files
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
' API to read registry
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
' End API Block

' Constants Needed for NTQueryInformationFile, organized into an Enumeration
' Credits to http://undocumented.ntinternals.net/ and the people who submitted this there
' Also thanks to the www.mvps.org Win32 Site where some C++ type/enums where available...
Public Enum FILE_INFORMATION_CLASS
    FileDirectoryInformation = 1
    FileFullDirectoryInformation   ' // 2
    FileBothDirectoryInformation   ' // 3
    FileBasicInformation           ' // 4  wdm
    FileStandardInformation        ' // 5  wdm
    FileInternalInformation        ' // 6
    FileEaInformation              ' // 74
    FileAccessInformation          ' // 8
    FileNameInformation            ' // 9
    FileRenameInformation          ' // 10
    FileLinkInformation            ' // 11
    FileNamesInformation           ' // 12
    FileDispositionInformation     ' // 13
    FilePositionInformation        ' // 14 wdm
    FileFullEaInformation          ' // 15
    FileModeInformation            ' // 16
    FileAlignmentInformation       ' // 17
    FileAllInformation             ' // 18
    FileAllocationInformation      ' // 19
    FileEndOfFileInformation       ' // 20 wdm
    FileAlternateNameInformation   ' // 21
    FileStreamInformation          ' // 22
    FilePipeInformation            ' // 23
    FilePipeLocalInformation       ' // 24
    FilePipeRemoteInformation      ' // 25
    FileMailslotQueryInformation   ' // 26
    FileMailslotSetInformation     ' // 27
    FileCompressionInformation     ' // 28
    FileObjectIdInformation        ' // 29
    FileCompletionInformation      ' // 30
    FileMoveClusterInformation     ' // 31
    FileQuotaInformation           ' // 32
    FileReparsePointInformation    ' // 33
    FileNetworkOpenInformation     ' // 34
    FileAttributeTagInformation    ' // 35
    FileTrackingInformation        ' // 36
    FileMaximumInformation
End Enum
' End Constants needed for NTQueryFileInformation

' Structures needed for the NT API
Public Type IO_STATUS_BLOCK
    IoStatus                As Long
    Information             As FILE_INFORMATION_CLASS
End Type
Public Type FILE_STREAM_INFORMATION
    NextEntryOffset         As Long
    StreamNameLength        As Long
    StreamSizeLow           As Long
    StreamSizeHi            As Long
    StreamAllocationSizeLow As Long
    StreamAllocationSizeHi  As Long
    StreamName(259)         As Byte
End Type
' End NTQueryInformationFile Structures

' Constants for Create and Open File APIs, neatly organized into Enumerations
Public Enum OpenFileFlags
  NoFileFlags = 0
  PosixSemantics = &H1000000
  BackupSemantics = &H2000000
  DeleteOnClose = &H4000000
  SequentialScan = &H8000000
  RandomAccess = &H10000000
  NoBuffering = &H20000000
  OverlappedIO = &H40000000
  WriteThrough = &H80000000
End Enum
Public Enum FileMode
  CreateNew = 1
  CreateAlways = 2
  OpenExisting = 3
  OpenOrCreate = 4
  Truncate = 5
  Append = 6
End Enum
Public Enum FileAccess
  AccessRead = &H80000000
  AccessWrite = &H40000000
  AccessReadWrite = &H80000000 Or &H40000000
  AccessDelete = &H10000
  AccessReadControl = &H20000
  AccessWriteDac = &H40000
  AccessWriteOwner = &H80000
  AccessSynchronize = &H100000
  AccessStandardRightsRequired = &HF0000
  AccessStandardRightsAll = &H1F0000
  AccessSystemSecurity = &H1000000
End Enum
Public Enum FileShare
  ShareNone = 0
  ShareRead = 1
  ShareWrite = 2
  ShareReadWrite = 3
  ShareDelete = 4
End Enum
' End Create/Open File Constants

' Stream Structure returned by the Backup APIs
Public Type WIN32_STREAM_ID
  dwStreamId              As Long
  dwStreamAttributes      As Long
  SizeLow                 As Long
  SizeHigh                As Long
  dwStreamNameSize        As Long
End Type

' And the Constants that go along with it...(in enumerations again)
Public Enum FileStreamTypes
  BACKUP_INVALID = 0
  BACKUP_DATA = 1                     ' Standard data stream (NTFS names "::DATA$")
  BACKUP_EA_DATA = 2                  ' Extended attribute data
  BACKUP_SECURITY_DATA = 3            ' Contains ACL's, etc.
  BACKUP_ALTERNATE_DATA = 4           ' Alternative data stream
  BACKUP_LINK = 5                     ' Posix style hard link
  BACKUP_PROPRETY_DATA = 6            ' Property data
  BACKUP_OBJECT_ID = 7                ' Uniquely identifies a file in the file system
  BACKUP_REPARSE_DATA = 8             ' Stream uses reparse points
  BACKUP_SPARSE_BLOCK = 9             ' Stream is a sparse file.
End Enum
Public Enum FileStreamAttributes
  BACKUP_NORMAL_ATTRIBUTE = &H0
  BACKUP_MODIFIED_WHEN_READ = &H1
  BACKUP_CONTAINS_SECURITY = &H2
  BACKUP_CONTAINS_PROPRETIES = &H4
  BACKUP_SPARSE_ATTRIBUTE = &H8
End Enum
Public Enum FileAttributes
    FILE_ATTRIBUTE_DIRECTORY = &H10
End Enum
' End Stream Type/Attributes Constants

' Constants for the SetFilePointer API
Public Enum FilePointerOptions                  ' Options for the SetFilePointer API
    BeginOfFile = 0
    FileCurrentPosition = 1
    EndOfFile = 2
End Enum
' End SetFilePointer API Constants

' Default Windows Buffer Size for Backup APIs
Public Const DefBufferSize     As Long = 128& * 1024&

' Flag to check if the File SYstem supports ADS
Public Const FILE_NAMED_STREAMS As Long = &H40000

' Structures needed for Find File APIs
Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Public Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * 261
    cAlternate As String * 14
End Type
' End Find File API structures

' --------------------------------------------------------------------------------------------------
'                       ENUMERATION FUNCTIONS TO FIND STREAMS IN A FILE OR DISK
' NTQIF = USES KERNEL NATIVE API TO FIND STREAMS IN A FILE, LIMITED BUT ULTRA FAST
' ENUMERATE_STREAMS = USES BACKUP API TO ENUMERATE STREAMS IN A FILE, RETURNS LOTS OF INFORMATION
' SEEKER = GENERAL FUNCTION THAT FINDS ALL THE FILES ON A DISK USING API
' --------------------------------------------------------------------------------------------------
Public Sub NTQIF(File As String)
'Ultra-fast Kernel Native Seeker
Dim IOStatusBlock As IO_STATUS_BLOCK
Dim Buffer() As Byte
Dim StreamInfo As FILE_STREAM_INFORMATION
Dim cbStreamInfo As Long, lpStreamInfo As Long, Handle As Long
Dim ErrorCode As Integer
Dim StreamName As String, Streams As String, StreamItem As Long, StreamSizes As Long

Handle = CreateFileW(StrPtr(File), 0&, 0&, 0&, OpenExisting, 0&, 0&)    ' We are just seeking, no additional options needed
If Handle = -1 Then Exit Sub                                            ' Cancel if we can't open the file
SendMessage StatusBar, SB_SETTEXTA, ByVal 0, ByVal File                 ' You can remove this for even faster speed, but it might look like it's frozen
cbStreamInfo = 4096                                                     ' 4K Buffer
ErrorCode = 234                                                         ' Will change once our buffer is OK
ReDim Buffer(1 To cbStreamInfo)                                         ' Start off with the 4K buffer
Do While ErrorCode = 234                                                ' (STACK BUFFER OVERFLOW)
    ErrorCode = NtQueryInformationFile(Handle, IOStatusBlock, ByVal VarPtr(Buffer(1)), cbStreamInfo, ByVal FileStreamInformation)
    If ErrorCode = 234 Then                                             ' Make buffer if we got that error
        cbStreamInfo = cbStreamInfo + 4096                              ' Add 4 more K
        ReDim Buffer(1 To cbStreamInfo)                                 ' Redimension the buffer
    End If
Loop
lpStreamInfo = VarPtr(Buffer(1))                                        ' Copy all the data to our buffer
Do
    RtlMoveMemory ByVal VarPtr(StreamInfo.NextEntryOffset), ByVal lpStreamInfo, 24                              ' Copy into our structure
    RtlMoveMemory ByVal VarPtr(StreamInfo.StreamName(0)), ByVal lpStreamInfo + 24, StreamInfo.StreamNameLength  ' Tell our struct how long the unicode string is
    StreamName = Left$(StreamInfo.StreamName, StreamInfo.StreamNameLength / 2)                                  ' Turn into a string, remove nulls...
    If StreamName <> "::$DATA" And StreamName <> ":encryptable:$DATA" Then
        Streams = Streams & Mid$(StreamName, 2, Len(StreamName) - 7) & ","                                      ' Add the stream to our stream list, except if it's the default one or encryption, also take off : and :$DATA
        StreamSizes = StreamInfo.StreamSizeLow + StreamSizes                                                     ' Add the cummulative size of all the streams
    End If
    If StreamInfo.NextEntryOffset Then lpStreamInfo = lpStreamInfo + StreamInfo.NextEntryOffset Else Exit Do    ' Jump to next stream
Loop
CloseHandle Handle                                                                                              ' We are done with the file, close it
If LenB(Streams) <> 0 Then                                              ' Check if streams were found
    StreamItem = CreateItem(ListView, File)                             ' Add the file to the listview
    ChangeItemText ListView, StreamItem, 1, Streams                     ' Add the name of the streams
    ChangeItemText ListView, StreamItem, 2, CStr(StreamSizes)           ' If there were any streams, add the sizes to listview
End If
Exit Sub
End Sub
Public Sub Seeker(Folder As String)
' This function uses the Find File APIs to enumerate every file on the disk
Dim FindHandle As Long
Dim FileName As String
Dim FileExists As Boolean
Dim W32_FIND_DATA As WIN32_FIND_DATA

FindHandle = FindFirstFile(Folder & "*.*", W32_FIND_DATA)                                       ' Start the search with the first file
FileExists = True                                                                               ' Make the loop work with the 1st file
While FileExists                                                                                ' As long as there are files...
    FileName = Left$(W32_FIND_DATA.cFileName, InStr(1, W32_FIND_DATA.cFileName, vbNullChar) - 1) ' Get the name, remove the nulls
    If LenB(FileName) Then                                                                      ' Make sure we read the filename, sometimes it doessnt!?
        If AscW(FileName) <> 46 Then                                                            ' Remove the "." and ".." root entries
            If (W32_FIND_DATA.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then               ' Is this a directory?
                Seeker Folder & FileName & "\"                                                  ' It is, start a new search inside it
            Else
                NTQIF Folder & FileName                                                         ' It isn't, it's a file, look for streams
            End If
        End If
    End If
    FileExists = FindNextFile(FindHandle, W32_FIND_DATA)                                        ' Move to the next file/folder
    DoEvents                                                                                    ' Don't freeze the UI
Wend
FindClose FindHandle                                                                            ' This search is finished
End Sub
Public Sub Enumerate_Streams(File As String)
Dim Handle As Long
Dim StreamName() As Byte, Buffer() As Byte
Dim W32_STREAM_ID As WIN32_STREAM_ID
Dim cbRead As Long, lpContext As Long, LoBytes As Long, HiBytes As Long, cbToSeek As Long, BufferLength As Long, StreamItem As Long

Handle = CreateFileW(StrPtr(File), AccessStandardRightsAll, ShareReadWrite, 0&, OpenExisting, BackupSemantics Or SequentialScan, 0&) ' Open the file
ReDim Buffer(DefBufferSize - 1) ' Redimension the buffer
cbRead = 1 ' To start off the loop
Do While cbRead
    LoBytes = SetFilePointer(Handle, 0&, 0&, BeginOfFile)                                                                           ' start at the beginning
    BackupRead Handle, VarPtr(Buffer(0)), LenB(W32_STREAM_ID), cbRead, 0&, 0&, lpContext                                           ' read the header
    If cbRead = 0 Then Exit Do                                                                                                      ' looks like we are done
    BufferLength = cbRead                                                                                                           ' remember where we are
    RtlMoveMemory ByVal VarPtr(W32_STREAM_ID), ByVal VarPtr(Buffer(0)), LenB(W32_STREAM_ID)                                         ' copy the header
    With W32_STREAM_ID
        If .dwStreamNameSize Then                                                                                                  ' this is a named stream
            ReDim StreamName(.dwStreamNameSize - 1)                                                                                 ' redimension the array
            cbRead = 0
            BackupRead Handle, VarPtr(Buffer(0)) + BufferLength, .dwStreamNameSize, cbRead, 0&, 0&, lpContext                       ' read from where we left off
            RtlMoveMemory ByVal VarPtr(StreamName(0)), ByVal VarPtr(Buffer(0)) + BufferLength, .dwStreamNameSize                    ' copy the stream name
            StreamItem = CreateItem(ListView, Left$(Left$(StreamName, .dwStreamNameSize / 2), .dwStreamNameSize / 2 - 6))           ' format it
            ChangeItemText ListView, StreamItem, 1, .SizeLow & " bytes"                                                                ' add information to listview (size/name)
            ChangeItemText ListView, StreamItem, 2, StreamIDToString(.dwStreamId)                                                   ' add information to listview (id type)
            ChangeItemText ListView, StreamItem, 3, StreamAttributeToString(.dwStreamAttributes)                                    ' add information to listview (attribs)
        End If
        cbToSeek = W32_STREAM_ID.SizeLow                                                                                               ' where to seek next
        BackupSeek Handle, cbToSeek, 0&, LoBytes, HiBytes, lpContext                                                                ' Unless you're dealing >4GB files, HiWord will also be 0...if not, you'd need to implement some dirty 64-bit integer hack
        If LoBytes = 0 Then Exit Do
    End With
Loop
' Flush out everything
BackupRead Handle, 0&, 0&, 0&, 1&, 0&, lpContext
CloseHandle Handle
End Sub
' --------------------------------------------------------------------------------------------------
'                       FILE FUNCTIONS TO CREATE, OPEN, VIEW OR DELETE STREAMS
' VIEWSTREAM = DUMPS THE CONTENTS OF A STREAM INTO A BUFFER
' DELETESTREAM = DELETES A STREAM, WITHOUT DELETING THE FILE
' OPENSTREAM = OPENS A STREAM IN WORDPAD FOR EASY EDITING
' CREATESTREAM = CREATES A NEW STREAM AND THEN OPENS IT
' --------------------------------------------------------------------------------------------------
Public Function ViewStream(StreamName As String) As String
Dim hFile As Long, Size As Long, Buffer As String, BytesRead As Long
'Reads a stream into a buffer
hFile = CreateFileW(StrPtr(StreamName), AccessRead, ShareRead, 0&, OpenExisting, 0&, 0&)
Size = GetFileSize(hFile, 0&)
Buffer = String$(Size, 0)
ReadFile hFile, ByVal Buffer, Size, BytesRead, 0
CloseHandle hFile
ViewStream = Buffer
End Function
Public Sub DeleteStream(StreamName As String)
'Deletes a stream
DeleteFile StreamName
End Sub
Public Function OpenStream(StreamName As String) As Long
'Opens a stream with Wordpad (finds where it's installed) and adds quotes around the name so Wordpad can open a stream that has spaces
'Returns the handle of the process so we can refresh the list after the user closed wordpad
Dim hShell As Long
hShell = Shell(Replace$(GetWordPadPath, "%ProgramFiles%", Environ$("PROGRAMFILES")) & " " & """" & StreamName & """", vbNormalFocus)
OpenStream = OpenProcess(&H100000, True, hShell)
End Function
Public Function CreateStream(StreamName As String) As Long
'Creates a New Stream and then Opens it
'Returns the handle of the process so we can refresh the list after the user clsoed wordpad
CloseHandle CreateFileW(StrPtr(StreamName), 0&, 0&, 0&, CreateNew, 0&, 0&)
CreateStream = OpenStream(StreamName)
End Function
' --------------------------------------------------------------------------------------------------
'                       VALUE TO STRING CONVERTERS FOR STREAM PROPRETIES
' STREAMIDTOSTRING = RETURNS THE STRING (MEANING) OF A GIVEN STREAMID
' STREAMATTRIBUTETOSTRING = RETURNS THE STRING (MEANING) OF A GIVEN STREAMATTRIBUTE
' --------------------------------------------------------------------------------------------------
Public Function StreamIDToString(ByVal StreamId As FileStreamTypes) As String
'Returns the ID type as a string
Select Case StreamId
    Case BACKUP_EA_DATA
        StreamIDToString = "Extended Data"
    Case BACKUP_ALTERNATE_DATA
        StreamIDToString = "Alternate Data"
    Case BACKUP_ALTERNATE_DATA
        StreamIDToString = "Hard Link"
    Case BACKUP_SECURITY_DATA
        StreamIDToString = "Security Data"
    Case BACKUP_PROPRETY_DATA
        StreamIDToString = "Proprety Data"
    Case BACKUP_OBJECT_ID
        StreamIDToString = "Object ID"
    Case BACKUP_REPARSE_DATA
        StreamIDToString = "Reparse Data"
    Case BACKUP_SPARSE_BLOCK
        StreamIDToString = "Sparse Block"
End Select
End Function
Public Function StreamAttributeToString(ByVal StreamAttribute As FileStreamAttributes) As String
'Returns the Attribute type as a string
Select Case StreamAttribute
    Case BACKUP_NORMAL_ATTRIBUTE
        StreamAttributeToString = "Normal Attribute"
    Case BACKUP_MODIFIED_WHEN_READ
        StreamAttributeToString = "Modified when Read"
    Case BACKUP_CONTAINS_SECURITY
        StreamAttributeToString = "Contains Security"
    Case BACKUP_CONTAINS_PROPRETIES
        StreamAttributeToString = "Contains Propreties"
    Case BACKUP_SPARSE_ATTRIBUTE
        StreamAttributeToString = "Sparse Attribute"
End Select
End Function
' --------------------------------------------------------------------------------------------------
'                       HELPER FUNCTIONS FOR STREAM MANIPULATION
' GETWORDPADPATH = RETURNS THE PATH OF WORDPAD ON ANY LOCALIZED VERSION OF WINDOWS NT
' CHECKSTREAMCAPABILITY = RETURNS WHETHER THE FILE SYSTEM SUPPORTS ADS
' --------------------------------------------------------------------------------------------------
Public Function GetWordPadPath() As String
' Check the Registry to find the path of wordpad
Dim strBuf As String, keyhand As Long

RegOpenKey &H80000002, "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WORDPAD.EXE", keyhand
strBuf = String$(255, " ")
RegQueryValueEx keyhand, vbNullString, 0&, 0&, ByVal strBuf, 255
GetWordPadPath = Left$(strBuf, InStr(strBuf, vbNullChar) - 1)
End Function
Public Function CheckStreamCapability() As Boolean
'Checks if the current FileSystem supports NTFS
Dim VolName As String * 256, VolSN As Long, MaxCompLen As Long, VolFlags As Long, VolFileSys As String * 256

GetVolumeInformation "C:\", VolName, Len(VolName), VolSN, MaxCompLen, VolFlags, VolFileSys, Len(VolFileSys)
If VolFlags And FILE_NAMED_STREAMS Then CheckStreamCapability = True
End Function
