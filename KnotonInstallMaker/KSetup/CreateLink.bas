Attribute VB_Name = "modCreateLink"
Option Explicit
Private Declare Function GetVersion Lib "kernel32" () As Long

Private oShell    As Object
Private oShortCut As Object

'Translate to a legal path on the installers computer
Private Function GetPath(ByVal SpecialFolder As String) As String
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

'creates a link, shortcut
Public Sub CreateLink(ByVal SpecialFolder As String, ByVal strTarget As String)
Dim strTemp As String
Dim tmp As String
On Error GoTo errHandler
tmp = strTarget
tmp = Mid$(tmp, InStrRev(tmp, "\") + 1)
tmp = Mid$(tmp, 1, InStrRev(tmp, ".") - 1)

Set oShell = CreateObject("WScript.Shell")
strTemp = GetPath(SpecialFolder)
CreateFolders strTemp
strTarget = GetEnvPath(strTarget)
If strTemp <> "" Then
    CreateFolders strTemp
    Set oShortCut = oShell.CreateShortcut(strTemp & "\" & Trim$(tmp) & ".lnk")
    With oShortCut
        .TargetPath = Trim$(strTarget)
        .Description = ""
        .Arguments = ""
        .WorkingDirectory = Trim$(Mid$(strTarget, 1, InStrRev(strTarget, "\") - 1))
        .WindowStyle = 4
        .Save
    End With
End If

errHandler:
Set oShell = Nothing
Set oShortCut = Nothing
End Sub

'Check if the installers machine is win9x or not
Public Function IsWin9x() As Boolean
  IsWin9x = CBool(GetVersion() And &H80000000)
End Function

