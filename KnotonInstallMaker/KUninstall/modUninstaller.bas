Attribute VB_Name = "modUninstaller"
Option Explicit
'declares for ini controlling
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long


Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const REG_SZ = 1

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long

Private Declare Function GetVersion Lib "kernel32" () As Long

Public Type EnvVariables
    System32            As String
    InstallationPath    As String
    CommonProgramFiles  As String
    ProgramFiles        As String
    WinDir              As String
End Type

Public EnvironVariables As EnvVariables


Public iniFile As String
Public setupName As String

'initiate
Public Sub Main()
iniFile = Command
If iniFile <> "" Then
    SetEnviron
    DeleteLinks
    DeleteFiles
    DeleteKey
    MsgBox setupName & " is uninstalled.", vbInformation, "KUninstaller"
End If
End
End Sub

'Delete the application being uninstalled from the add or remove programs applet
Public Sub DeleteKey()
Dim ret As Long
RegOpenKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & setupName, ret
If ret Then RegDeleteKey ret, ""
End Sub

'Set environment paths and initate variables
Public Sub SetEnviron()
Dim strEnv As String, tmp As String
tmp = Mid(iniFile, InStrRev(iniFile, "\") + 1)
setupName = Mid(tmp, 1, Len(tmp) - 4)

With EnvironVariables
    .CommonProgramFiles = Environ("CommonProgramFiles")
    .ProgramFiles = Environ("ProgramFiles")
    .InstallationPath = .ProgramFiles & "\" & setupName
    .WinDir = Environ("WinDir")
    .System32 = .WinDir & "\System32"
End With
End Sub

'Translate into a legal path on the installers machine
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

'Translate into legal specialfolder paths on the installers machine
Private Function GetPath(ByVal SpecialFolder As String) As String
Dim oShell As Object
Set oShell = CreateObject("WScript.Shell")

If IsWin9x Then
    If InStr(1, SpecialFolder, "%AllUsersDesktop%") Then
        GetPath = oShell.specialfolders.Item("Desktop")
    ElseIf InStr(1, SpecialFolder, "%AllUsersStartUp%") Then
        GetPath = oShell.specialfolders.Item("Startup")
    ElseIf InStr(1, SpecialFolder, "%AllUsersPrograms%") Then
        GetPath = Replace(SpecialFolder, "%AllUsersPrograms%", oShell.specialfolders.Item("Programs"))
    ElseIf InStr(1, SpecialFolder, "%Desktop%") Then
        GetPath = oShell.specialfolders.Item("Desktop")
    ElseIf InStr(1, SpecialFolder, "%StartUp%") Then
        GetPath = oShell.specialfolders.Item("Startup")
    ElseIf InStr(1, SpecialFolder, "%Programs%") Then
        GetPath = Replace(SpecialFolder, "%Programs%", oShell.specialfolders.Item("Programs"))
    End If
Else
    If InStr(1, SpecialFolder, "%AllUsersDesktop%") Then
        GetPath = oShell.specialfolders.Item("AllUsersDesktop")
    ElseIf InStr(1, SpecialFolder, "%AllUsersStartUp%") Then
        GetPath = oShell.specialfolders.Item("AllUsersStartup")
    ElseIf InStr(1, SpecialFolder, "%Desktop%") Then
        GetPath = oShell.specialfolders.Item("Desktop")
    ElseIf InStr(1, SpecialFolder, "%StartUp%") Then
        GetPath = oShell.specialfolders.Item("Startup")
    ElseIf InStr(1, SpecialFolder, "%QuickLaunch%") Then
        GetPath = oShell.specialfolders.Item("AppData") & "\Microsoft\Internet Explorer\Quick Launch"
    ElseIf InStr(1, SpecialFolder, "%AllUsersPrograms%") Then
        GetPath = Replace(SpecialFolder, "%AllUsersPrograms%", oShell.specialfolders.Item("AllUsersPrograms"))
    ElseIf InStr(1, SpecialFolder, "%Programs%") Then
        GetPath = Replace(SpecialFolder, "%Programs%", oShell.specialfolders.Item("Programs"))
    End If
End If

End Function


Private Function IsWin9x() As Boolean
  IsWin9x = CBool(GetVersion() And &H80000000)
End Function

'Delete links that belongs to the app being uninstalled
Public Sub DeleteLinks()
Dim var As Variant, tmp As Variant, i As Integer
On Error Resume Next
var = ReadSection("LINKS")

If var <> "" Then
    var = Split(var, Chr(0))
    For i = 0 To UBound(var)
        tmp = Split(var(i), "=")
        Kill GetPath(tmp(1)) & "\" & Mid$(tmp(0), 1, InStrRev(tmp(0), ".") - 1) & ".lnk"
    Next
End If

On Error GoTo 0
End Sub

'Delete files that was installed with the app being uninstalled, leave files marked shared
Public Sub DeleteFiles()
Dim var As Variant, tmp As Variant, i As Integer, strTemp As String
On Error Resume Next
var = ReadSection("DESTINATION")

If var <> "" Then
    var = Split(var, Chr(0))
    For i = 0 To UBound(var)
        tmp = Split(var(i), "=")
        If Read("SHARED", tmp(0)) = "" Then
            Kill GetEnvPath(Read("DESTINATION", tmp(0))) & "\" & tmp(0)
        End If
    Next
End If
Kill iniFile
On Error GoTo 0
End Sub

'reads ini string
Public Function Read(ByVal Section As String, ByVal Key As String, Optional ByVal strDefault As String) As String
Dim RetVal As String * 255, v As Long
v = GetPrivateProfileString(Section, Key, strDefault, RetVal, 255, iniFile)
If v > 0 Then Read = Left(RetVal, v)
End Function

'reads ini section
Public Function ReadSection(ByVal Section As String) As String
Dim RetVal As String * 4096, v As Long
v = GetPrivateProfileSection(Section, RetVal, 4096, iniFile)
If v > 0 Then ReadSection = Left(RetVal, v - 1)
End Function

'writes ini
Public Sub Save(ByVal Section As String, ByVal Key As String, ByVal Value As String)
WritePrivateProfileString Section, Key, Value, iniFile
End Sub

'writes ini section
Public Sub SaveSection(Section As String, Value As String)
WritePrivateProfileSection Section, Value, iniFile
End Sub

'removes ini section
Public Sub RemoveSection(Section As String)
WritePrivateProfileString Section, vbNullString, "", iniFile
End Sub

