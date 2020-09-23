VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmInstaller 
   Caption         =   "KInstaller"
   ClientHeight    =   6870
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   11850
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkInstallPath 
      Caption         =   "Allow User to choose %InstallationPath%"
      Height          =   555
      Left            =   9720
      TabIndex        =   8
      Top             =   6240
      Width           =   1935
   End
   Begin VB.CheckBox chkAdmin 
      Caption         =   "User must be Admin"
      Height          =   315
      Left            =   9720
      TabIndex        =   7
      Top             =   5880
      Width           =   1935
   End
   Begin VB.TextBox txtLinks 
      Height          =   285
      Left            =   3120
      TabIndex        =   5
      Top             =   6480
      Width           =   6135
   End
   Begin VB.ComboBox cboLinks 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   6480
      Width           =   2895
   End
   Begin VB.ComboBox cboEnvironPath 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   5880
      Width           =   2895
   End
   Begin VB.TextBox txtEnvironPath 
      Height          =   285
      Left            =   3120
      TabIndex        =   1
      Top             =   5880
      Width           =   6135
   End
   Begin MSComctlLib.ListView lvFiles 
      Height          =   5535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   9763
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Links"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   6240
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Environ Path variables"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   5640
      Width           =   2655
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAS 
         Caption         =   "Save As"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuAction 
      Caption         =   "Action"
      Enabled         =   0   'False
      Begin VB.Menu mnuAddFiles 
         Caption         =   "Add Files"
      End
      Begin VB.Menu mnuChangeDestination 
         Caption         =   "Change Destination"
      End
      Begin VB.Menu mnuDelFiles 
         Caption         =   "Delete checked Files"
      End
      Begin VB.Menu mnuFilesShared 
         Caption         =   "Mark Checked Files Shared"
      End
      Begin VB.Menu mnuFilesUnshare 
         Caption         =   "Mark Checked Files Unshared"
      End
      Begin VB.Menu mnuSetLink 
         Caption         =   "Set Link"
      End
      Begin VB.Menu mnuRemoveLink 
         Caption         =   "Remove Link"
      End
      Begin VB.Menu mnuCreateCab 
         Caption         =   "Create cabinet"
      End
   End
End
Attribute VB_Name = "frmInstaller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const LVM_FIRST = &H1000
Private Const LVM_DELETEALLITEMS = (LVM_FIRST + 9)

'Setup the listviews
Private Sub FixaListViews()

lvFiles.ColumnHeaders.Add 1, , "Filename"
lvFiles.ColumnHeaders.Add 2, , "Source"
lvFiles.ColumnHeaders.Add 3, , "Destination"
lvFiles.ColumnHeaders.Add 4, , "Shared"
lvFiles.ColumnHeaders.Add 5, , "Links"

lvFiles.View = lvwReport
lvFiles.FullRowSelect = True
lvFiles.Sorted = True
lvFiles.SortKey = 1
End Sub

'Set up the combo with environpaths
Private Sub FixaEnvironPaths()
cboEnvironPath.AddItem "%System32%"
cboEnvironPath.AddItem "%InstallationPath%"
cboEnvironPath.AddItem "%CommonProgramFiles%"
cboEnvironPath.AddItem "%ProgramFiles%"
cboEnvironPath.AddItem "%WinDir%"

End Sub

'Set up the combo with links
Private Sub FixaLinks()
cboLinks.AddItem "%Programs%"
cboLinks.AddItem "%AllUsersPrograms%"
cboLinks.AddItem "%AllUsersStartUp%"
cboLinks.AddItem "%AllUsersDesktop%"
cboLinks.AddItem "%Desktop%"
cboLinks.AddItem "%SendTo%"
cboLinks.AddItem "%StartUp%"
cboLinks.AddItem "%QuickLaunch%"
End Sub

'add to the textbox that holds the environpath to set
Private Sub cboEnvironPath_Click()
If txtEnvironPath = "" Then
    txtEnvironPath = cboEnvironPath.Text
Else
    txtEnvironPath = txtEnvironPath & "\" & cboEnvironPath.Text
End If
End Sub

'Add to the textbox that holds the linkpath to set
Private Sub cboLinks_Click()
txtLinks = cboLinks.Text
If cboLinks.ListIndex >= 2 Then
    txtLinks.Locked = True
Else
    txtLinks.Locked = False
End If
End Sub

'Set if the setup being made only can be used with an admin or not
Private Sub chkAdmin_Click()
If mnuAction.Enabled Then
    Save "APP", "Admin", CStr(chkAdmin.Value)
End If
End Sub

'Set if the user are allowed to choose installationpath
Private Sub chkInstallPath_Click()
If mnuAction.Enabled Then
    Save "APP", "InstallPath", CStr(chkInstallPath.Value)
End If
End Sub

Private Sub Form_Load()
FixaListViews
FixaEnvironPaths
FixaLinks
End Sub

'Resize the listview
Private Sub Form_Resize()
Dim x As Double

If Me.WindowState <> vbMinimized Then
    lvFiles.Width = Me.Width - 150
    x = (lvFiles.Width - 5200) / 2
    lvFiles.ColumnHeaders(1).Width = 2150
    lvFiles.ColumnHeaders(2).Width = x + 500
    lvFiles.ColumnHeaders(3).Width = x - 600
    lvFiles.ColumnHeaders(4).Width = 900
    lvFiles.ColumnHeaders(5).Width = 2150
End If
End Sub

'Make sure nothing is selected in the listview
Private Sub AvmarkeraListView(Lv As ListView)
Dim i As Integer
For i = 1 To Lv.ListItems.Count
    Lv.ListItems(i).Selected = False
Next
End Sub

'Refresh the gui with saved data
Private Sub RefreshFiles()
Dim var As Variant, i As Integer, itmx As ListItem, tmp As Variant
Call SendMessage(lvFiles.hwnd, LVM_DELETEALLITEMS, 0, ByVal 0&)

var = ReadSection("FILES")
var = Split(var, Chr(0))
For i = 0 To UBound(var)
    tmp = Split(var(i), "=")
    Set itmx = lvFiles.ListItems.Add(, , tmp(0))
    itmx.SubItems(1) = tmp(1)
    itmx.SubItems(2) = Read("DESTINATION", tmp(0), "")
    itmx.SubItems(3) = Read("SHARED", tmp(0), "")
    itmx.SubItems(4) = Read("LINKS", tmp(0), "")
Next

chkAdmin.Value = Read("APP", "Admin", "0")
chkInstallPath.Value = Read("APP", "InstallPath", "0")

AvmarkeraListView lvFiles
End Sub

'Save data to the .ksf file (inifile)
Private Sub SaveFiles()
Dim itmx As ListItem, i As Integer
RemoveSection "FILES"
RemoveSection "DESTINATION"
RemoveSection "SHARED"
RemoveSection "LINKS"
RemoveSection "APP"
Save "APP", "Admin", CStr(chkAdmin.Value)
Save "APP", "InstallPath", CStr(chkInstallPath.Value)

For i = 1 To lvFiles.ListItems.Count
    Set itmx = lvFiles.ListItems.Item(i)
    If FileExist(itmx.SubItems(1) & "\" & itmx) Then
        Save "FILES", itmx, itmx.SubItems(1)
        Save "DESTINATION", itmx, itmx.SubItems(2)
        If itmx.SubItems(3) <> "" Then Save "SHARED", itmx, itmx.SubItems(3)
        If itmx.SubItems(4) <> "" Then Save "LINKS", itmx, itmx.SubItems(4)
    End If
Next
End Sub

'Delete all the checked files from the .ksf file
Private Sub DeleteCheckedFiles()
Dim itmx As ListItem, i As Integer
RemoveSection "FILES"
RemoveSection "DESTINATION"
RemoveSection "SHARED"
RemoveSection "LINKS"

For i = 1 To lvFiles.ListItems.Count
    Set itmx = lvFiles.ListItems.Item(i)
    If Not itmx.Checked Then
        Set itmx = lvFiles.ListItems.Item(i)
        Save "FILES", itmx, itmx.SubItems(1)
        Save "DESTINATION", itmx, itmx.SubItems(2)
        
        If itmx.SubItems(3) <> "" Then
            Save "SHARED", itmx, itmx.SubItems(3)
        End If
        
        If itmx.SubItems(4) <> "" Then
            Save "LINKS", itmx, itmx.SubItems(4)
        End If
    End If
Next

RefreshFiles
End Sub

'Mark all checked files as shared, will not be deleted during uninstall
Private Sub MarkCheckedFilesShared()
Dim itmx As ListItem, i As Integer
SaveFiles
RemoveSection "SHARED"

For i = 1 To lvFiles.ListItems.Count
    Set itmx = lvFiles.ListItems.Item(i)
    If itmx.Checked Or itmx.SubItems(3) <> "" Then
        Set itmx = lvFiles.ListItems.Item(i)
        Save "SHARED", itmx, "SHARED"
    End If
Next
RefreshFiles
End Sub

'Mark checked files as not shared
Private Sub MarkCheckedFilesUnShared()
Dim itmx As ListItem, i As Integer
SaveFiles
RemoveSection "SHARED"

For i = 1 To lvFiles.ListItems.Count
    Set itmx = lvFiles.ListItems.Item(i)
    If Not itmx.Checked And itmx.SubItems(3) <> "" Then
        Set itmx = lvFiles.ListItems.Item(i)
        Save "SHARED", itmx, "SHARED"
    End If
Next
RefreshFiles
End Sub

'Add files to the setup being made
Private Sub SelectFiles()
Dim tmp As String, dirPath As String
Dim mFiles As Variant
Dim i As Integer, x As Integer, itmx As ListItem
Dim CD As New CommonDialog

CD.Filter = "All Files (*.*)|*.*"
CD.DialogTitle = "Select Files"
CD.AllowMultiSelect = True
CD.ShowOpen
mFiles = CD.Filename
If IsArray(mFiles) Then
    x = UBound(mFiles)
    tmp = mFiles(1)
    For i = 1 To x - 1
        mFiles(i) = mFiles(i + 1)
    Next
    mFiles(x) = tmp
    dirPath = mFiles(0)
    
    For i = 1 To x
        Set itmx = lvFiles.ListItems.Add(, , mFiles(i))
        itmx.SubItems(1) = dirPath
        itmx.SubItems(2) = "%InstallationPath%"
    Next
    SaveFiles
    RefreshFiles
Else
    If mFiles <> "" Then
        dirPath = Mid(mFiles, 1, InStrRev(mFiles, "\") - 1)
        Set itmx = lvFiles.ListItems.Add(, , Replace(mFiles, dirPath & "\", ""))
        itmx.SubItems(1) = dirPath
        itmx.SubItems(2) = "%InstallationPath%"
        SaveFiles
        RefreshFiles
    End If
End If


End Sub

Private Sub mnuAddFiles_Click()
SelectFiles
End Sub

'Alter the destination path (computer being installed to) for checked files
Private Sub mnuChangeDestination_Click()
Dim i As Integer, itmx As ListItem
If Len(txtEnvironPath.Text) > 0 Then
    If Right(txtEnvironPath, 1) = "\" Then txtEnvironPath = Mid(txtEnvironPath, 1, Len(txtEnvironPath) - 1)
    For i = 1 To lvFiles.ListItems.Count
        Set itmx = lvFiles.ListItems.Item(i)
        If itmx.Checked Then
            Save "DESTINATION", itmx, txtEnvironPath.Text
        End If
    Next
    RefreshFiles
End If
End Sub

'Create the cab file, add .ksf file and uninstaller
Private Sub mnuCreateCab_Click()
Dim i As Integer, itmx As ListItem, strFiles() As String
ReDim strFiles(lvFiles.ListItems.Count + 1)
For i = 1 To lvFiles.ListItems.Count
    Set itmx = lvFiles.ListItems.Item(i)
    strFiles(i - 1) = itmx.SubItems(1) & "\" & itmx
Next
FileCopy App.Path & "\Uninstall.exe", strWorkDir & "\KUninstall.exe"
strFiles(lvFiles.ListItems.Count) = SetupFilePath
strFiles(UBound(strFiles)) = strWorkDir & "\KUninstall.exe"
Me.MousePointer = vbHourglass
CreateCabinet strFiles
Kill strWorkDir & "\KUninstall.exe"
Me.MousePointer = vbDefault
End Sub

Private Sub mnuDelFiles_Click()
DeleteCheckedFiles
End Sub

Private Sub mnuFilesShared_Click()
MarkCheckedFilesShared
End Sub

Private Sub mnuFilesUnshare_Click()
MarkCheckedFilesUnShared
End Sub

'Make a new .ksf file, remember to give it the name of your program since it is being used in other places
Private Sub mnuNew_Click()
Dim CD As New CommonDialog, tmp As String
CD.InitDir = App.Path
CD.Filter = "KInstall Files (*.ksf)|*.ksf"
CD.DialogTitle = "Create KInstaller Setup File"
CD.ShowSave
tmp = CD.Filename
If tmp <> "" Then
    If LCase(Right(tmp, 4)) <> ".ksf" Then tmp = tmp & ".ksf"
    Call SendMessage(lvFiles.hwnd, LVM_DELETEALLITEMS, 0, ByVal 0&)
    SetupFilePath = tmp
    Me.Caption = SetupFilePath
    strWorkDir = Mid(tmp, 1, InStrRev(tmp, "\") - 1)
    setupFile = Mid(tmp, InStrRev(tmp, "\") + 1)
    setupName = Mid(setupFile, 1, Len(setupFile) - 4)
    RefreshFiles
    mnuSave.Enabled = True
    mnuSaveAS.Enabled = True
    mnuAction.Enabled = True
End If

End Sub

'Open a .ksf file to do further work with
Private Sub mnuOpen_Click()
Dim CD As New CommonDialog, tmp As String
CD.InitDir = App.Path
CD.Filter = "KInstall Files (*.ksf)|*.ksf"
CD.DialogTitle = "Select KInstaller Setup File to work with"
CD.ShowOpen
tmp = CD.Filename
If tmp <> "" Then
    If LCase(Right(tmp, 4)) <> ".ksf" Then tmp = tmp & ".ksf"
    Call SendMessage(lvFiles.hwnd, LVM_DELETEALLITEMS, 0, ByVal 0&)
    SetupFilePath = tmp
    Me.Caption = SetupFilePath
    strWorkDir = Mid(tmp, 1, InStrRev(tmp, "\") - 1)
    setupFile = Mid(tmp, InStrRev(tmp, "\") + 1)
    setupName = Mid(setupFile, 1, Len(setupFile) - 4)
    RefreshFiles
    mnuSave.Enabled = True
    mnuSaveAS.Enabled = True
    mnuAction.Enabled = True
End If

End Sub

'Remove links from the .ksf file
Private Sub mnuRemoveLink_Click()
Dim itmx As ListItem, i As Integer
SaveFiles
RemoveSection "LINKS"

For i = 1 To lvFiles.ListItems.Count
    Set itmx = lvFiles.ListItems.Item(i)
    If Not itmx.Checked And itmx.SubItems(4) <> "" Then
        Set itmx = lvFiles.ListItems.Item(i)
        Save "LINKS", itmx, itmx.SubItems(4)
    End If
Next

RefreshFiles
End Sub

'Save the data
Private Sub mnuSave_Click()
SaveFiles
RefreshFiles
End Sub

'Save as (use current projekt to make a new projekt)
Private Sub mnuSaveAS_Click()
Dim CD As New CommonDialog, tmp As String
CD.InitDir = App.Path
CD.Filter = "KInstall Files (*.ksf)|*.ksf"
CD.DialogTitle = "Save KInstaller Setup File as"
CD.ShowOpen
tmp = CD.Filename
If tmp <> "" Then
    If LCase(Right(tmp, 4)) <> ".ksf" Then tmp = tmp & ".ksf"
    SetupFilePath = tmp
    Me.Caption = SetupFilePath
    strWorkDir = Mid(tmp, 1, InStrRev(tmp, "\") - 1)
    SaveFiles
    RefreshFiles
    mnuSave.Enabled = True
    mnuSaveAS.Enabled = True
    mnuAction.Enabled = True
End If

End Sub

'Set links for checked files
Private Sub mnuSetLink_Click()
Dim itmx As ListItem, i As Integer
If Len(txtLinks) > 0 Then
    If Right(txtLinks, 1) = "\" Then txtLinks = Mid(txtLinks, 1, Len(txtLinks) - 1)
    For i = 1 To lvFiles.ListItems.Count
        Set itmx = lvFiles.ListItems.Item(i)
        If itmx.Checked Then
            Set itmx = lvFiles.ListItems.Item(i)
            Save "LINKS", itmx, txtLinks
        End If
    Next
    RefreshFiles
End If
End Sub

'Weak protection to not make bad links
Private Sub txtLinks_Change()
If txtLinks = "" Then
    txtLinks.Locked = True
Else
    txtLinks.Locked = False
End If
End Sub
