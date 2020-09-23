VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEncryptSingleKey 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CryptAPI Single-Key Encryption Demo"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   6450
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5040
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      InitDir         =   "c:"
   End
   Begin VB.CommandButton cmdSwitch 
      Caption         =   "Switch to Key-Pair Example"
      Height          =   495
      Left            =   2040
      TabIndex        =   12
      Top             =   6600
      Width           =   2415
   End
   Begin VB.TextBox txtStringEncrypt 
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   4440
      Width           =   6135
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "Example password"
      Top             =   1080
      Width           =   6135
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "Encrypt && Decrypt File Using Single Key"
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      Top             =   5880
      Width           =   2415
   End
   Begin VB.CommandButton cmdDecrypt 
      Caption         =   "Decrypt Text Using Single Key"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   3800
      Width           =   2535
   End
   Begin VB.CommandButton cmdEncrypt 
      Caption         =   "Encrypt Text Using Single Key"
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   1720
      Width           =   2535
   End
   Begin VB.TextBox txtStringPlain 
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmEncryptSingleKey.frx":0000
      Top             =   2400
      Width           =   6135
   End
   Begin VB.Line Line3 
      Index           =   2
      X1              =   120
      X2              =   6360
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line Line3 
      Index           =   1
      X1              =   120
      X2              =   6360
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6360
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label1 
      Caption         =   "File Example"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   11
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "String Example"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   $"frmEncryptSingleKey.frx":0010
      Height          =   615
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   6135
   End
   Begin VB.Label Label1 
      Caption         =   "Password (This is the ""Key"")- This can be any length"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Encrypted Text - Note that VB eliminates invisible characters so this isn't all of the data"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   3960
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Plain Text - This can be any length"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   2535
   End
End
Attribute VB_Name = "frmEncryptSingleKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objEncryption As clsCryptoAPIandCompression
Dim strIn As String
Dim strOut As String
'you must use variables to store the encrypted data because VB
'eliminates invisible characters if you use the "text" property of a text box

Private Sub cmdDecrypt_Click()

cmdDecrypt.Enabled = False
cmdEncrypt.Enabled = True

strIn = strOut 'strOut already set in cmdEncrypt_Click
strOut = objEncryption.DecryptString(strIn, txtPassword) 'decrypt it

With txtStringPlain
    .Text = strOut
    .SelStart = 0
    .SelLength = 65535
    .SetFocus
End With

End Sub

Private Sub cmdEncrypt_Click()

cmdDecrypt.Enabled = True
cmdEncrypt.Enabled = False

strIn = txtStringPlain.Text
strOut = objEncryption.EncryptString(strIn, txtPassword)

With txtStringEncrypt
    .Text = strOut
    .SelStart = 0
    .SelLength = 65535
    .SetFocus
End With

End Sub

Private Sub cmdFile_Click()
Dim sngTim As Single
Dim strFileAndPathName As String

cmdFile.Enabled = False

CommonDialog1.FileName = ""
CommonDialog1.ShowOpen
strFileAndPathName = CommonDialog1.FileName

sngTim = Timer

objEncryption.EncryptFile strFileAndPathName, strFileAndPathName & "ENCRYPTED", txtPassword.Text

objEncryption.DecryptFile strFileAndPathName & "ENCRYPTED", strFileAndPathName & "DECRYPTED", txtPassword.Text

MsgBox "Finished encrypting and decrypting file.  Remove the 'DECRYPTED' file extension to test the decrypted file. " & vbNewLine & "Time elapsed: " & Timer - sngTim

cmdFile.Enabled = True

End Sub

Private Sub cmdSwitch_Click()

Unload frmEncryptSingleKey
frmEncryptKeyPair.Show

End Sub

Private Sub Form_Load()

Set objEncryption = New clsCryptoAPIandCompression
objEncryption.SessionStart

End Sub

Private Sub Form_Unload(Cancel As Integer)

objEncryption.SessionEnd
Set objEncryption = Nothing

End Sub
