VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Test Exception Handler"
   ClientHeight    =   2310
   ClientLeft      =   4575
   ClientTop       =   2070
   ClientWidth     =   5595
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   5595
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "C&lose"
      Height          =   330
      Left            =   3120
      TabIndex        =   7
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdUnInstall 
      Caption         =   "&UnInstall"
      Height          =   330
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdInstall 
      Caption         =   "&Install"
      Height          =   330
      Left            =   1320
      TabIndex        =   3
      Top             =   1920
      Width           =   1095
   End
   Begin VB.ComboBox cmbException 
      Height          =   315
      ItemData        =   "frmMain.frx":0000
      Left            =   1650
      List            =   "frmMain.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1320
      Width           =   3840
   End
   Begin VB.CommandButton cmdCrash 
      Caption         =   "&Crash"
      Default         =   -1  'True
      Height          =   330
      Left            =   4320
      TabIndex        =   0
      Top             =   1920
      Width           =   1140
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMain.frx":0004
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   1200
      TabIndex        =   6
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "WARNING:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Exception type:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   1365
      Width           =   1305
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdInstall_Click()
Call InstallExceptionHandler
UpdateInstallState
End Sub

Private Sub cmdUnInstall_Click()
Call UninstallExceptionHandler
UpdateInstallState
End Sub

Private Sub Form_Load()
    cmbException.AddItem GetExceptionName(enumExceptionType.enumExceptionType_AccessViolation)
    cmbException.ItemData(0) = enumExceptionType_AccessViolation
    cmbException.AddItem GetExceptionName(enumExceptionType.enumExceptionType_IllegalInstruction)
    cmbException.ItemData(1) = enumExceptionType_IllegalInstruction
    cmbException.AddItem GetExceptionName(enumExceptionType.enumExceptionType_PriviledgedInstruction)
    cmbException.ItemData(2) = enumExceptionType_PriviledgedInstruction
    cmbException.AddItem GetExceptionName(enumExceptionType.enumExceptionType_ArrayBoundsExceeded)
    cmbException.ItemData(3) = enumExceptionType_ArrayBoundsExceeded
    cmbException.AddItem GetExceptionName(enumExceptionType.enumExceptionType_Breakpoint)
    cmbException.ItemData(4) = enumExceptionType_Breakpoint
    cmbException.AddItem GetExceptionName(enumExceptionType.enumExceptionType_ControlCExit)
    cmbException.ItemData(5) = enumExceptionType_ControlCExit
    cmbException.AddItem GetExceptionName(enumExceptionType.enumExceptionType_DataTypeMisalignment)
    cmbException.ItemData(6) = enumExceptionType_DataTypeMisalignment
    cmbException.AddItem GetExceptionName(enumExceptionType.enumExceptionType_NoncontinuableException)
    cmbException.ItemData(7) = enumExceptionType_NoncontinuableException
    cmbException.AddItem GetExceptionName(enumExceptionType.enumExceptionType_SingleStep)
    cmbException.ItemData(8) = enumExceptionType_SingleStep
    cmbException.ListIndex = 0
    cmdInstall_Click
End Sub
Private Sub UpdateInstallState()
cmdInstall.Enabled = Not blnIsHandlerInstalled
cmdUnInstall.Enabled = blnIsHandlerInstalled
End Sub
Private Sub cmdCrash_Click()
On Error GoTo hErr
    Call RaiseAnException(cmbException.ItemData(cmbException.ListIndex))
Exit Sub
hErr:
    HandleTheException Err.Description, "frmMain.cmdCrash_Click()"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If Not blnIsHandlerInstalled Then UninstallExceptionHandler
End Sub

