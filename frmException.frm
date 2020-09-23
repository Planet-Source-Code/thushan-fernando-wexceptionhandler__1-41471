VERSION 5.00
Begin VB.Form frmException 
   Caption         =   "Exception Error"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7320
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmException.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4800
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtException 
      Appearance      =   0  'Flat
      Height          =   1215
      Left            =   1320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   2280
      Width           =   5775
   End
   Begin VB.CheckBox chkAutoStart 
      Caption         =   "&Auto Start Application to previous state"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   4560
      Value           =   1  'Checked
      Width           =   3495
   End
   Begin VB.CommandButton cmdContinue 
      Cancel          =   -1  'True
      Caption         =   "&Continue.."
      Height          =   330
      Left            =   4920
      TabIndex        =   5
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Default         =   -1  'True
      Height          =   330
      Left            =   6120
      TabIndex        =   4
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   35
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   7335
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Exception:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00981F0A&
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   2280
      Width           =   870
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmException.frx":0CCA
      Height          =   615
      Left            =   240
      TabIndex        =   7
      Top             =   3720
      Width           =   6975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmException.frx":0D78
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   7095
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   6720
      Picture         =   "frmException.frx":0F35
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "An Exception Error Occured in &wsAppName;"
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   3150
   End
   Begin VB.Label lblException 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exception ErrorHandler"
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
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1950
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "frmException"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private blnAutoStart As Boolean
Private blnContinue As Boolean

Public Property Get bContinue() As Boolean
bContinue = blnContinue
End Property
Public Property Get bAutoStart() As Boolean
bAutoStart = blnAutoStart
End Property

Private Sub chkAutoStart_Click()
blnAutoStart = CBool(chkAutoStart.Value)
End Sub

Private Sub cmdExit_Click()
blnContinue = False
Hide
End Sub

Private Sub cmdContinue_Click()
blnContinue = True
Hide
End Sub
