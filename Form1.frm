VERSION 5.00
Object = "*\AProject1.vbp"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Audio LEDs & How to use them."
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   3480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Play"
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   600
      Width           =   2655
   End
   Begin LED.hLED hLED1 
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   1200
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   450
   End
   Begin MSComDlg.CommonDialog cdb 
      Left            =   1920
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load"
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin LED.vLED vLED1 
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   2355
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(): On Error GoTo Canceled
cdb.Filter = "Windows Waveform Files (*.wav)|*.wav"
cdb.ShowOpen
OpenFile cdb.FileName
Canceled:
End Sub

Private Sub Command2_Click()
Play
End Sub

Private Sub Form_Load()
Initialize_Engine
End Sub

Private Sub Form_Unload(Cancel As Integer)
Terminate_Engine
End
End Sub
