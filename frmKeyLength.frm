VERSION 5.00
Begin VB.Form frmKeyLength 
   Caption         =   "Select Key Size"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4305
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4305
   Begin VB.OptionButton opt16384 
      Caption         =   "16384"
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
   End
   Begin VB.OptionButton opt8192 
      Caption         =   "8192"
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
   End
   Begin VB.OptionButton opt4096 
      Caption         =   "4096"
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.OptionButton opt2048 
      Caption         =   "2048"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.OptionButton opt1024 
      Caption         =   "1024"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   1200
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton opt512 
      Caption         =   "512"
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   $"frmKeyLength.frx":0000
      Height          =   855
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmKeyLength"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngInternalKey As Long

Public Property Get Key() As Long

Key = lngInternalKey

End Property

Private Sub Form_Load()

lngInternalKey = 1024 'default value

End Sub

Private Sub opt512_Click()
lngInternalKey = 512
End Sub

Private Sub opt1024_Click()
lngInternalKey = 1024
End Sub

Private Sub opt2048_Click()
lngInternalKey = 2048
End Sub

Private Sub opt4096_Click()
lngInternalKey = 4096
End Sub

Private Sub opt8192_Click()
lngInternalKey = 8192
End Sub

Private Sub opt16384_Click()
lngInternalKey = 16384
End Sub
