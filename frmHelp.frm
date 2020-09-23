VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Help with Key Pair Encryption"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   7215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Label1 = "Key pair encryption is different than the more-intuitive session key encryption.  Session key encryption uses the same key (password) to encrypt and decrypt plaintext.  With key pair encryption, two keys are generated which are mathematically related, yet it is impossible to figure out one key just because you have its pair.  To exchange data, you use an 'exchange key pair'.  One is a public key that you can give out to anyone.  The other is a private key which is password-protected and should be stored where others can not get it.  Typically you use someone else's public key to encrypt plaintext which only they can then decrypt using their private key.  If you want to encrypt something for your own use, use your own public key to encrypt it and then only you can decrypt it using your private key."
Label2 = "The CryptoAPI also generates another key pair for signatures, but you use the keys in the opposite order.  To ensure that I am the one that sent you something, I will sign the data using my private signature key, and you will check the signature using my public signature key."

End Sub
