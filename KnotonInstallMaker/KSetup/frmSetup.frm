VERSION 5.00
Begin VB.Form frmSetup 
   Caption         =   "KSetup"
   ClientHeight    =   1425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   3765
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSetup 
      Caption         =   "Setup"
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   420
      Width           =   1215
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetupPromptReboot Lib "setupapi.dll" (ByRef FileQueue As Long, ByVal Owner As Long, ByVal ScanOnly As Long) As Long

Private Sub cmdSetup_Click()
On Error GoTo errHandler
SetEnviron
If Extract Then
    If CheckAdmin Then
        If Not IsWin9x Then
            If Not IsAdmin Then
                MsgBox "You must be local admin to install " & App.EXEName
                DeleteTempPath
                End
            End If
        End If
    End If
    ChooseInstallPath
    SetFiles
    CopyFiles
    MakeLinks
    CreateKey
    If Not blnReboot Then
        DeleteTempPath
        MsgBox "The installation is done!", vbInformation
    Else
        MsgBox "The installation is done!" & vbCrLf & _
                "But some files where in use and need to be replaced" & vbCrLf & _
                "during next reboot." & vbCrLf & _
                "You must reboot before using " & App.EXEName, vbInformation
                
        SetupPromptReboot ByVal 0&, Me.hwnd, 0

    End If
End If

Exit Sub
errHandler:
MsgBox Err.Description
End Sub

Private Sub Form_Load()
SetEnviron
Me.Caption = "Setup " & App.EXEName
End Sub
