Attribute VB_Name = "modInstaller"
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function IsNTAdmin Lib "advpack.dll" (ByVal dwReserved As Long, ByRef lpdwReserved As Long) As Long

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Private Declare Function RemoveDirectory Lib "kernel32" Alias "RemoveDirectoryA" (ByVal lpPathName As String) As Long
Private Declare Function MoveFileEx Lib "kernel32" Alias "MoveFileExA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal dwFlags As Long) As Long

Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal Length As Long)
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Any, ByVal wParam As Any, ByVal lParam As Any) As Long
Private Const ERROR_SUCCESS = &H0
Private Const MOVEFILE_DELAY_UNTIL_REBOOT = &H4

Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long


'declares for ini controlling
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, _
            ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, _
            ByVal dwMilliseconds As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Const SYNCHRONIZE = &H100000
Const INFINITE = &HFFFF 'Wait forever
Const WAIT_OBJECT_0 = 0 'The state of the specified object is signaled
Const WAIT_TIMEOUT = &H102 'The time-out interval elapsed & the object’s state is nonsignaled.

Private Type OSVERSIONINFO 'All Windows Version
        dwOSVersionInfoSize As Long 'Structure size = 148
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128
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
   dwFileFlags As Long            '  e.g. VFF_DEBUG Or VFF_PRERELEASE
   dwFileOS As Long               '  e.g. VOS_DOS_WINDOWS16
   dwFileType As Long             '  e.g. VFT_DRIVER
   dwFileSubtype As Long          '  e.g. VFT2_DRV_KEYBOARD
   dwFileDateMS As Long           '  e.g. 0
   dwFileDateLS As Long           '  e.g. 0
End Type

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

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

Public Type EnvVariables
    System32            As String
    InstallationPath    As String
    CommonProgramFiles  As String
    ProgramFiles        As String
    WinDir              As String
End Type

Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const REG_SZ = 1

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long


Public strTempPath As String

Public IniFile  As String
Public EnvironVariables As EnvVariables
Public arrFiles As Variant
Public blnReboot As Boolean

'Check if the installer is allowed to choose installpath, if so let him do so.
Public Sub ChooseInstallPath()
Dim cd As New CommonDialog, strTemp As String
Dim var As Variant, i As Integer, tmp As Variant

If CBool(Read("APP", "InstallPath", "0")) Then
    strTemp = cd.GetFolderName("What folder do you want to install " & App.EXEName & " in?")
    If strTemp <> "" Then
        EnvironVariables.InstallationPath = strTemp
        var = ReadSection("DESTINATION")
        var = Split(var, Chr(0))
        For i = 0 To UBound(var)
            tmp = Split(var(i), "=")
            If tmp(1) = "%InstallationPath%" Then
                Save IniFile, "DESTINATION", tmp(0), strTemp
            End If
        Next
    End If
End If
End Sub

'Check if the installer must be admin
Public Function CheckAdmin() As Boolean
CheckAdmin = CBool(Read("APP", "Admin", "0"))
End Function

'Check if the installer is an admin, only works on WinNT, on Win9x is everybody admin
Public Function IsAdmin() As Boolean
    IsAdmin = CBool(IsNTAdmin(ByVal 0&, ByVal 0&))
End Function

'Create the key in the registry that adds the program into add or remove programs applet
Public Sub CreateKey()
Dim ret As Long, strUninstallString As String
strUninstallString = EnvironVariables.System32 & "\KUninstall.exe " & EnvironVariables.System32 & "\" & App.EXEName & ".ksf"

RegOpenKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & App.EXEName, ret

If ret = 0 Then RegCreateKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & App.EXEName, ret
If ret Then RegSetValueEx ret, "UninstallString", 0, REG_SZ, ByVal strUninstallString, Len(strUninstallString)
If ret Then RegSetValueEx ret, "DisplayName", 0, REG_SZ, ByVal App.EXEName, Len(App.EXEName)

RegCloseKey ret
End Sub

'Make links (shortcuts)
Public Sub MakeLinks()
Dim var As Variant, tmp As Variant, i As Integer, strLink As String
var = ReadSection("LINKS")
If var <> "" Then
    var = Split(var, Chr(0))
    For i = 0 To UBound(var)
        tmp = Split(var(i), "=")
        CreateLink tmp(1), Read("DESTINATION", tmp(0)) & "\" & tmp(0)
    Next
End If
End Sub

'Extract the cabinet file
Public Function Extract() As Boolean
Dim strCMD As String, blnRet As Boolean
Dim lPid As Long
Dim lHnd As Long
Dim lRet As Long

strCMD = Environ$("COMSPEC") & " /c " & "Expand " & App.Path & "\" & App.EXEName & ".cab -f:*.* " & strTempPath
CreateFolders (strTempPath)
lPid = Shell(strCMD, vbHide)
If lPid <> 0 Then
        'Get a handle to the shelled process.
        lHnd = OpenProcess(SYNCHRONIZE, 0, lPid)
        'If successful, wait for the application to end and close the handle.
        If lHnd <> 0 Then
            lRet = WaitForSingleObject(lHnd, INFINITE)
            CloseHandle (lHnd)
        End If
End If
Extract = CBool(lPid)
End Function

'Add the files being installed to an array
Public Sub SetFiles()
Dim strTemp As String, arrVar As Variant, i As Integer
strTemp = ReadSection("FILES")
If strTemp <> "" Then
    arrFiles = Split(strTemp, Chr(0))

    For i = 0 To UBound(arrFiles)
        arrVar = Split(arrFiles(i), "=")
        arrFiles(i) = arrVar(0)
    Next
End If
End Sub

'Delete the files in the temppath, removedirectory works bad
Public Sub DeleteTempPath()
Kill strTempPath & "\*.*"
RemoveDirectory strTempPath
RemoveDirectory "c:\KSetup"
End Sub


'Set environment paths and initial variables
Public Sub SetEnviron()
Dim strEnv As String
With EnvironVariables
    .CommonProgramFiles = Environ("CommonProgramFiles")
    .ProgramFiles = Environ("ProgramFiles")
    .InstallationPath = .ProgramFiles & "\" & App.EXEName
    .WinDir = Environ("WinDir")
    .System32 = .WinDir & "\System32"
End With
strTempPath = "c:\KSetup\tmp"
IniFile = strTempPath & "\" & App.EXEName & ".ksf"
End Sub

'Translate the environmentpath into a legal path on the installers computer
Public Function GetEnvPath(ByVal strPath As String) As String
Dim strTemp As String, i As Integer


With EnvironVariables
    strPath = Replace(strPath, "%System32%", .System32)
    strPath = Replace(strPath, "%InstallationPath%", .InstallationPath)
    strPath = Replace(strPath, "%CommonProgramFiles%", .CommonProgramFiles)
    strPath = Replace(strPath, "%ProgramFiles%", .ProgramFiles)
    strPath = Replace(strPath, "%WinDir%", .WinDir)
End With

GetEnvPath = strPath
End Function

'Copy all files, check for existing, register and unregister, if file in use add them to be replaced during next boot
'replace files during next boot is not tested for win9x, let me know if it works ok.
Public Sub CopyFiles()
Dim i As Integer, strDest As String

If IsArray(arrFiles) Then
    For i = 0 To UBound(arrFiles)
        strDest = GetEnvPath(Read("DESTINATION", arrFiles(i)))
        If strDest <> "" Then
            CreateFolders strDest
            If FileExist(strDest & "\" & arrFiles(i)) Then
                If GetFileVersion(strTempPath & "\" & arrFiles(i)) >= GetFileVersion(strDest & "\" & arrFiles(i)) Then
                    If FileExist(strTempPath & "\" & arrFiles(i)) Then
                        If Not FileInUse(strTempPath & "\" & arrFiles(i)) Then
                            RegisterServer strDest & "\" & arrFiles(i), False
                            FileCopy strTempPath & "\" & arrFiles(i), strDest & "\" & arrFiles(i)
                            RegisterServer strDest & "\" & arrFiles(i), True
                        Else
                            If IsWin9x Then
                                Save EnvironVariables.WinDir & "\Wininet.ini", "rename", strDest & "\" & arrFiles(i), strTempPath & "\" & arrFiles(i)
                                blnReboot = True
                            Else
                                MoveFileEx strTempPath & "\" & arrFiles(i), strDest & "\" & arrFiles(i), MOVEFILE_DELAY_UNTIL_REBOOT
                                blnReboot = True
                            End If
                        End If
                    End If
                End If
            Else
                If FileExist(strTempPath & "\" & arrFiles(i)) Then
                    FileCopy strTempPath & "\" & arrFiles(i), strDest & "\" & arrFiles(i)
                    RegisterServer strDest & "\" & arrFiles(i), True
                End If
            End If
        End If
    Next
End If
FileCopy IniFile, EnvironVariables.System32 & "\" & App.EXEName & ".ksf"
FileCopy strTempPath & "\KUninstall.exe", EnvironVariables.System32 & "\KUninstall.exe"

End Sub

'Create folders
Public Sub CreateFolders(ByVal strFullPath As String)
Dim ret As Long
Dim Security As SECURITY_ATTRIBUTES
Dim arrFolders As Variant
Dim tmpFolder As String
Dim i As Integer
If Right(strFullPath, 1) = "\" Then strFullPath = Mid(strFullPath, 1, Len(strFullPath) - 1)

arrFolders = Split(strFullPath, "\")
tmpFolder = arrFolders(0)
For i = 1 To UBound(arrFolders)
    tmpFolder = tmpFolder & "\" & arrFolders(i)
    ret = CreateDirectory(tmpFolder, Security)
Next
End Sub

'Check fileversion
Public Function GetFileVersion(ByVal FullFileName As String) As String
Dim rc As Long, lDummy As Long, sBuffer() As Byte
Dim lBufferLen As Long, lVerPointer As Long, udtVerBuffer As VS_FIXEDFILEINFO
Dim lVerbufferLen As Long

'*** Get size ****
lBufferLen = GetFileVersionInfoSize(FullFileName, lDummy)
If lBufferLen < 1 Then
    '**** Store info to udtVerBuffer struct ****
    ReDim sBuffer(lBufferLen)
    rc = GetFileVersionInfo(FullFileName, 0&, lBufferLen, sBuffer(0))
    rc = VerQueryValue(sBuffer(0), "\", lVerPointer, lVerbufferLen)
    MoveMemory udtVerBuffer, lVerPointer, Len(udtVerBuffer)
    
    '**** Determine File Version number ****
    GetFileVersion = Format$(udtVerBuffer.dwFileVersionMSh) & "." & Format$(udtVerBuffer.dwFileVersionMSl) & "." & Format$(udtVerBuffer.dwFileVersionLSh) & "." & Format$(udtVerBuffer.dwFileVersionLSl)
End If
End Function

'Copy files
Public Function FileCopy(ByVal Source As String, ByVal Destination As String) As Boolean
Dim ret As Long
ret = CopyFile(Source, Destination, 0&)
If ret <> 0 Then FileCopy = True
End Function

'Check if file exist
Public Function FileExist(ByVal FullFileName As String) As Boolean
Dim ret As Long
Dim WFD As WIN32_FIND_DATA
ret = FindFirstFile(FullFileName, WFD)
If ret <> INVALID_HANDLE_VALUE Then FileExist = True
End Function

'Check if file being replaced is in use
Public Function FileInUse(ByVal Filename As String) As Boolean
Dim iFile As Long

If Not FileExist(Filename) Then
    FileInUse = False
Else
    On Error Resume Next
    iFile = FreeFile
    Open Filename For Binary Access Read Lock Read Write As #iFile
    FileInUse = (Err.Number <> 0)
    Close iFile
End If
End Function

'Register and unregister
Public Function RegisterServer(DllServerPath As String, bRegister As Boolean)
Dim lb As Long, pa As Long
On Error GoTo errHandler
lb = LoadLibrary(DllServerPath)

If lb <> 0 Then
    If bRegister Then
        pa = GetProcAddress(lb, "DllRegisterServer")
    Else
        pa = GetProcAddress(lb, "DllUnregisterServer")
    End If
    
    If CallWindowProc(pa, 0&, ByVal 0&, ByVal 0&, ByVal 0&) = ERROR_SUCCESS Then
       RegisterServer = True
    Else
        RegisterServer = False
    End If
    FreeLibrary lb
End If

errHandler:
End Function

'add path to system variables path, not used (yet)
Public Function SetEnv(VarName As String, VarVal As String) As String
Dim SysEnvObj
On Error Resume Next
Set SysEnvObj = CreateObject("WScript.Shell")
    SysEnvObj.Environment.Item(VarName) = SysEnvObj.Environment.Item(VarName) & ";" & VarVal
    SetEnv = SysEnvObj.Environment.Item(VarName)
On Error GoTo 0
End Function

'reads ini string
Public Function Read(ByVal Section As String, ByVal Key As String, Optional ByVal strDefault As String) As String
Dim RetVal As String * 255, v As Long
v = GetPrivateProfileString(Section, Key, strDefault, RetVal, 255, IniFile)
If v > 0 Then Read = Left(RetVal, v)
End Function

'reads ini section
Public Function ReadSection(ByVal Section As String) As String
Dim RetVal As String * 4096, v As Long
v = GetPrivateProfileSection(Section, RetVal, 4096, IniFile)
If v > 0 Then ReadSection = Left(RetVal, v - 1)
End Function

'writes ini
Public Sub Save(ByVal objIni As String, ByVal Section As String, ByVal Key As String, ByVal Value As String)
WritePrivateProfileString Section, Key, Value, objIni
End Sub

'writes ini section
Public Sub SaveSection(Section As String, Value As String)
WritePrivateProfileSection Section, Value, IniFile
End Sub

'removes ini section
Public Sub RemoveSection(Section As String)
WritePrivateProfileString Section, vbNullString, "", IniFile
End Sub
