VERSION 5.00
Begin VB.UserControl hLED 
   ClientHeight    =   270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2685
   ScaleHeight     =   270
   ScaleWidth      =   2685
   Begin VB.PictureBox LEDPB 
      BackColor       =   &H0000C000&
      Height          =   255
      Index           =   0
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   10
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox LEDPB 
      BackColor       =   &H0000C000&
      Height          =   255
      Index           =   1
      Left            =   240
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   9
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox LEDPB 
      BackColor       =   &H0000C000&
      Height          =   255
      Index           =   2
      Left            =   480
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   8
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox LEDPB 
      BackColor       =   &H0000C000&
      Height          =   255
      Index           =   3
      Left            =   720
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   7
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox LEDPB 
      BackColor       =   &H0000C000&
      Height          =   255
      Index           =   4
      Left            =   960
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   6
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox LEDPB 
      BackColor       =   &H0000C000&
      Height          =   255
      Index           =   5
      Left            =   1200
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   5
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox LEDPB 
      BackColor       =   &H0000C0C0&
      Height          =   255
      Index           =   6
      Left            =   1440
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   4
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox LEDPB 
      BackColor       =   &H0000C0C0&
      Height          =   255
      Index           =   7
      Left            =   1680
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   3
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox LEDPB 
      BackColor       =   &H0000C0C0&
      Height          =   255
      Index           =   8
      Left            =   1920
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox LEDPB 
      BackColor       =   &H000000C0&
      Height          =   255
      Index           =   9
      Left            =   2160
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox LEDPB 
      BackColor       =   &H000000C0&
      Height          =   255
      Index           =   10
      Left            =   2400
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "hLED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim P_Val As Integer
Property Get Value() As Integer
Value = P_Val
End Property
Property Let Value(NewVal As Integer)
P_Val = NewVal
For i = 0 To 10
If i <= P_Val Then
If i < 6 Then LEDPB(i).BackColor = &HFF00&
If i > 5 And i < 9 Then LEDPB(i).BackColor = &HFFFF&
If i > 8 Then LEDPB(i).BackColor = &HFF
Else
If i < 6 Then LEDPB(i).BackColor = &HC000&
If i > 5 And i < 9 Then LEDPB(i).BackColor = &HC0C0&
If i > 8 Then LEDPB(i).BackColor = &HC0
End If
Next
PropertyChanged "Value"
End Property
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
P_Val = PropBag.ReadProperty("Value", 0)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Call PropBag.WriteProperty("Value", P_Val, 0)
End Sub

