Attribute VB_Name = "modInstaller"
Option Explicit
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, _
            ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, _
            ByVal dwMilliseconds As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long

Const SYNCHRONIZE = &H100000
Const INFINITE = &HFFFF 'Wait forever
Const WAIT_OBJECT_0 = 0 'The state of the specified object is signaled
Const WAIT_TIMEOUT = &H102 'The time-out interval elapsed & the object’s state is nonsignaled.

'declares for ini controlling
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long

Private Const MAX_PATH = 260
Private Const INVALID_HANDLE_VALUE = -1


Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
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

Public strWorkDir As String
Public setupFile As String
Public SetupFilePath As String
Public setupName As String

'reads ini string
Public Function Read(ByVal Section As String, ByVal Key As String, Optional ByVal strDefault As String) As String
Dim RetVal As String * 255, v As Long
v = GetPrivateProfileString(Section, Key, strDefault, RetVal, 255, SetupFilePath)
If v > 0 Then Read = Left(RetVal, v)
End Function

'reads ini section
Public Function ReadSection(ByVal Section As String) As String
Dim RetVal As String * 4096, v As Long
v = GetPrivateProfileSection(Section, RetVal, 4096, SetupFilePath)
If v > 0 Then ReadSection = Left(RetVal, v - 1)
End Function

'writes ini
Public Sub Save(ByVal Section As String, ByVal Key As String, ByVal Value As String)
WritePrivateProfileString Section, Key, Value, SetupFilePath
End Sub

'writes ini section
Public Sub SaveSection(Section As String, Value As String)
WritePrivateProfileSection Section, Value, SetupFilePath
End Sub

'removes ini section
Public Sub RemoveSection(Section As String)
WritePrivateProfileString Section, vbNullString, "", SetupFilePath
End Sub

'Create the cabinet file
Public Sub CreateCabinet(arrFiles() As String)
Dim fFile As Integer, strHead As String, iCount As Integer, i As Integer
Dim strCMD As String, blnRet As Boolean
Dim lPid As Long
Dim lHnd As Long
Dim lRet As Long

If arrFiles(0) <> "" Then
    iCount = UBound(arrFiles)
    fFile = FreeFile
    
    strHead = ".Option Explicit" & vbCrLf & _
                ".Set Cabinet=on" & vbCrLf & _
                ".Set Compress=on" & vbCrLf & _
                ".Set MaxDiskSize=CDRom" & vbCrLf & _
                ".Set ReservePerCabinetSize=6144" & vbCrLf & _
                ".Set DiskDirectoryTemplate=" & vbCrLf & _
                ".Set CompressionType=MSZip" & vbCrLf & _
                ".Set CompressionLevel=7" & vbCrLf & _
                ".Set CompressionMemory=21" & vbCrLf & _
                ".Set CabinetNameTemplate=" & setupName & ".cab"
                
    Open strWorkDir & "\Setup.ddf" For Output As #fFile
    Print #fFile, strHead
    
    For i = 0 To iCount
         Write #fFile, arrFiles(i)
    Next i
        
    Close #fFile
    
    SetCurrentDirectory strWorkDir
    
    strCMD = Environ$("COMSPEC") & " /c " & "Makecab /F Setup.ddf"
    lPid = Shell(strCMD, vbHide)
    If lPid <> 0 Then
            lHnd = OpenProcess(SYNCHRONIZE, 0, lPid)
            If lHnd <> 0 Then
                lRet = WaitForSingleObject(lHnd, INFINITE)
                CloseHandle (lHnd)
            End If
            Kill strWorkDir & "\setup.rpt"
            Kill strWorkDir & "\setup.inf"
            Kill strWorkDir & "\setup.ddf"
            FileCopy App.Path & "\Setup.exe", strWorkDir & "\" & setupName & ".exe"
    End If
End If
End Sub

'Check if file exist
Public Function FileExist(ByVal FullFileName As String) As Boolean
Dim ret As Long
Dim WFD As WIN32_FIND_DATA
ret = FindFirstFile(FullFileName, WFD)
If ret <> INVALID_HANDLE_VALUE Then FileExist = True
End Function

'Copy file
Public Function FileCopy(ByVal Source As String, ByVal Destination As String) As Boolean
Dim ret As Long
ret = CopyFile(Source, Destination, 0&)
If ret <> 0 Then FileCopy = True
End Function

