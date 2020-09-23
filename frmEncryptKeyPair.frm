VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEncryptKeyPair 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CryptAPI Key-Pair Encryption Demo"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   8295
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdListProviders 
      Caption         =   "List Available Providers on This Machine"
      Height          =   195
      Left            =   2160
      TabIndex        =   21
      Top             =   2160
      Width           =   3015
   End
   Begin VB.CommandButton cmdKeyPairValidate 
      Caption         =   "Validate Text (Signature Key)"
      Enabled         =   0   'False
      Height          =   615
      Left            =   7080
      TabIndex        =   20
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdKeyPairSign 
      Caption         =   "Sign Text (Signature Key)"
      Enabled         =   0   'False
      Height          =   615
      Left            =   7080
      TabIndex        =   19
      Top             =   120
      Width           =   1095
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Generate New Signature Key Pair and Export It for Future Use"
      Height          =   735
      Left            =   120
      TabIndex        =   18
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton cmdKeyPairFileValidate 
      Caption         =   "Validate File"
      Enabled         =   0   'False
      Height          =   615
      Left            =   6840
      TabIndex        =   17
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdKeyPairFileSign 
      Caption         =   "Sign File"
      Enabled         =   0   'False
      Height          =   615
      Left            =   5640
      TabIndex        =   16
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdHelp 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Help with Key Pair Encryption"
      Height          =   495
      Left            =   240
      TabIndex        =   15
      Top             =   5280
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2160
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      InitDir         =   "c:"
   End
   Begin VB.CommandButton cmdKeyPairFileDecrypt 
      Caption         =   "Decrypt File"
      Enabled         =   0   'False
      Height          =   615
      Left            =   4320
      TabIndex        =   14
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdSwitch 
      Caption         =   "Switch to Single-Key Example"
      Height          =   495
      Left            =   4080
      TabIndex        =   13
      Top             =   5280
      Width           =   2295
   End
   Begin VB.CommandButton cmdEnableKeyPair 
      Caption         =   "Generate or Import Key"
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   4680
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Generate New Exchange Key Pair and Export It for Future Use"
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Import Key"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdKeyPairFileEncrypt 
      Caption         =   "Encrypt File"
      Enabled         =   0   'False
      Height          =   615
      Left            =   3120
      TabIndex        =   6
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdKeyPairDecrypt 
      Caption         =   "Decrypt Text Using Exchange Key Pair (your private key)"
      Enabled         =   0   'False
      Height          =   615
      Left            =   5400
      TabIndex        =   5
      Top             =   2155
      Width           =   1575
   End
   Begin VB.CommandButton cmdKeyPairEncrypt 
      Caption         =   "Encrypt Text Using Exchange Key Pair (your public key)"
      Enabled         =   0   'False
      Height          =   615
      Left            =   5400
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txtStringEncrypt 
      Height          =   1215
      Left            =   2040
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2880
      Width           =   6135
   End
   Begin VB.TextBox txtStringPlain 
      Height          =   1215
      Left            =   2040
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmEncryptKeyPair.frx":0000
      Top             =   840
      Width           =   6135
   End
   Begin VB.Line Line1 
      X1              =   5520
      X2              =   5520
      Y1              =   4440
      Y2              =   5040
   End
   Begin VB.Line Line3 
      Index           =   2
      X1              =   1920
      X2              =   8280
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line3 
      Index           =   1
      X1              =   1920
      X2              =   8280
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Label Label1 
      Caption         =   "File Example"
      Height          =   255
      Index           =   6
      Left            =   2040
      TabIndex        =   12
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "String Example"
      Height          =   255
      Index           =   5
      Left            =   2040
      TabIndex        =   11
      Top             =   240
      Width           =   1215
   End
   Begin VB.Line Line2 
      Index           =   1
      X1              =   1920
      X2              =   1920
      Y1              =   120
      Y2              =   5760
   End
   Begin VB.Label Label1 
      Caption         =   $"frmEncryptKeyPair.frx":0010
      Height          =   2535
      Index           =   4
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Encrypted Text - Note that VB eliminates invisible characters so this isn't all of the data"
      Height          =   375
      Index           =   1
      Left            =   2040
      TabIndex        =   3
      Top             =   2400
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Plain Text - This can be any length"
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   2
      Top             =   600
      Width           =   2535
   End
End
Attribute VB_Name = "frmEncryptKeyPair"
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

Private Sub cmdEnableKeyPair_Click()
Dim intNextFreeFile As Integer
Dim strFileAndPathName As String
Dim strKey As String

cmdEnableKeyPair.Enabled = False
frmEncryptKeyPair.MousePointer = vbHourglass

objEncryption.KeyLength = frmKeyLength.Key

If Option1.Value = True Then 'generate and export Exchange keys
    objEncryption.Generate_KeyPair True 'generate Exchange key pair
    objEncryption.Export_KeyPair InputBox("Enter a password to encrypt the private key.") 'export the key pair (makes it available in the ValuePublicKey and ValuePublicPrivateKey properties)
    
    'save the key pair to files; add Timer value to easily get a somewhat unique file name
    intNextFreeFile = FreeFile
    Open "C:\" & Timer & " PublicPrivateKey.prk" For Binary Access Write As #intNextFreeFile
    Put #intNextFreeFile, , objEncryption.ValuePublicPrivateKey
    Close #intNextFreeFile
    
    intNextFreeFile = FreeFile
    Open "C:\" & Timer & " PublicKey.pbk" For Binary Access Write As #intNextFreeFile
    Put #intNextFreeFile, , objEncryption.ValuePublicKey
    Close #intNextFreeFile
    
    MsgBox "Generated new key pair and exported the keys to C:\?????.?? PublicPrivateKey.prk and C:\?????.?? PublicKey.pbk"
    'Note that the PublicPrivate Key contains your private key and
    'should be stored securely...the class encrypts it for you like
    'PGP does.  You must know the Password to decrypt it and thus
    'decrypt files with it.
    'See http://www.pgp.com/products/freeware/default.asp
    'for the best (and free) commercial-grade file and email
    'encryption program around.
    
    'enable/disable the buttons
    cmdKeyPairSign.Enabled = False
    cmdKeyPairValidate.Enabled = False
    cmdKeyPairFileSign.Enabled = False
    cmdKeyPairFileValidate.Enabled = False
    cmdKeyPairEncrypt.Enabled = True
    cmdKeyPairDecrypt.Enabled = False
    cmdKeyPairFileEncrypt.Enabled = True
    cmdKeyPairFileDecrypt.Enabled = True
    
ElseIf Option3.Value = True Then 'generate and export Signature keys
    objEncryption.Generate_KeyPair False 'generate Signature key pair
    objEncryption.Export_KeyPair InputBox("Enter a password to encrypt the private key.") 'export the key pair (makes it available in the ValuePublicKey and ValuePublicPrivateKey properties)
    
    'save the key pair to files; add Timer value to easily get a somewhat unique file name
    intNextFreeFile = FreeFile
    Open "C:\" & Timer & " PublicPrivateKey.srk" For Binary Access Write As #intNextFreeFile
    Put #intNextFreeFile, , objEncryption.ValuePublicPrivateKey
    Close #intNextFreeFile
    
    intNextFreeFile = FreeFile
    Open "C:\" & Timer & " PublicKey.spk" For Binary Access Write As #intNextFreeFile
    Put #intNextFreeFile, , objEncryption.ValuePublicKey
    Close #intNextFreeFile
    
    MsgBox "Generated new key pair and exported the keys to C:\?????.?? PublicPrivateKey.srk and C:\?????.?? PublicKey.spk"
    'Note that the PublicPrivate Key contains your private key and
    'should be stored securely.
    
    'enable/disable the buttons
    cmdKeyPairSign.Enabled = True
    cmdKeyPairValidate.Enabled = False
    cmdKeyPairFileSign.Enabled = True
    cmdKeyPairFileValidate.Enabled = True
    cmdKeyPairEncrypt.Enabled = False
    cmdKeyPairDecrypt.Enabled = False
    cmdKeyPairFileEncrypt.Enabled = False
    cmdKeyPairFileDecrypt.Enabled = False
    
Else 'option2 is selected so import a key
    CommonDialog1.Filter = "Keys (.prk; .pbk; .srk; .spk)|*.prk;*.pbk;*.srk;*.spk" 'only show keys
    CommonDialog1.FileName = ""
    CommonDialog1.ShowOpen
    strFileAndPathName = CommonDialog1.FileName

    'read key from file
    intNextFreeFile = FreeFile
    Open strFileAndPathName For Binary As #intNextFreeFile
    strKey = String(LOF(intNextFreeFile), vbNullChar) 'must initialize the variable
    Get #intNextFreeFile, , strKey
    Close #intNextFreeFile
    
    'pass the data to the correct property
    Select Case Right(strFileAndPathName, 4) 'look at the file extension
        Case ".prk" 'Exchange publicprivate key
            objEncryption.ValuePublicPrivateKey = String(Len(strKey), vbNullChar) 'initialize the variable
            objEncryption.ValuePublicPrivateKey = strKey
            objEncryption.Import_KeyPair InputBox("Enter the password to decrypt the private key."), True
            MsgBox "Imported the PublicPrivate key pair."
            'enable/disable the buttons
            cmdKeyPairSign.Enabled = False
            cmdKeyPairValidate.Enabled = False
            cmdKeyPairFileSign.Enabled = False
            cmdKeyPairFileValidate.Enabled = False
            cmdKeyPairEncrypt.Enabled = True
            cmdKeyPairDecrypt.Enabled = False
            cmdKeyPairFileEncrypt.Enabled = True
            cmdKeyPairFileDecrypt.Enabled = True
            
        Case ".pbk" 'Exchange public key
            objEncryption.ValuePublicKey = String(Len(strKey), vbNullChar) 'initialize the variable
            objEncryption.ValuePublicKey = strKey
            objEncryption.Import_KeyPair , True
            MsgBox "Imported the Public key."
            'enable/disable the buttons
            cmdKeyPairSign.Enabled = False
            cmdKeyPairValidate.Enabled = False
            cmdKeyPairFileSign.Enabled = False
            cmdKeyPairFileValidate.Enabled = False
            cmdKeyPairEncrypt.Enabled = True
            cmdKeyPairDecrypt.Enabled = False
            cmdKeyPairFileEncrypt.Enabled = True
            cmdKeyPairFileDecrypt.Enabled = False
            
        Case ".srk" 'Signature publicprivate key
            objEncryption.ValuePublicPrivateKey = String(Len(strKey), vbNullChar) 'initialize the variable
            objEncryption.ValuePublicPrivateKey = strKey
            objEncryption.Import_KeyPair InputBox("Enter the password to decrypt the private key."), False
            MsgBox "Imported the PublicPrivate key pair."
            'enable/disable the buttons
            cmdKeyPairSign.Enabled = True
            cmdKeyPairValidate.Enabled = True
            cmdKeyPairFileSign.Enabled = True
            cmdKeyPairFileValidate.Enabled = True
            cmdKeyPairEncrypt.Enabled = False
            cmdKeyPairDecrypt.Enabled = False
            cmdKeyPairFileEncrypt.Enabled = False
            cmdKeyPairFileDecrypt.Enabled = False
            
        Case ".spk" 'Signature public key
            objEncryption.ValuePublicKey = String(Len(strKey), vbNullChar) 'initialize the variable
            objEncryption.ValuePublicKey = strKey
            objEncryption.Import_KeyPair , False
            MsgBox "Imported the Public key."
            'enable/disable the buttons
            cmdKeyPairSign.Enabled = False
            cmdKeyPairValidate.Enabled = True
            cmdKeyPairFileSign.Enabled = False
            cmdKeyPairFileValidate.Enabled = True
            cmdKeyPairEncrypt.Enabled = False
            cmdKeyPairDecrypt.Enabled = False
            cmdKeyPairFileEncrypt.Enabled = False
            cmdKeyPairFileDecrypt.Enabled = False
            
        Case Else 'not a key file
            MsgBox "Not a key file.  Did not import a key."

    End Select 'file extension

End If 'generate new or import

cmdEnableKeyPair.Enabled = True
frmEncryptKeyPair.MousePointer = vbArrow

End Sub

Private Sub cmdHelp_Click()
frmHelp.Show

End Sub

Private Sub cmdKeyPairDecrypt_Click()

cmdKeyPairDecrypt.Enabled = False
cmdKeyPairEncrypt.Enabled = True

strIn = strOut 'strOut already set in cmdKeyPairEncrypt_Click
strOut = objEncryption.DecryptString_KeyPair(strIn) 'decrypt it using the Exchange key pair

With txtStringPlain
    .Text = strOut
    .SelStart = 0
    .SelLength = 65535
    .SetFocus
End With

End Sub

Private Sub cmdKeyPairEncrypt_Click()

cmdKeyPairDecrypt.Enabled = True
cmdKeyPairEncrypt.Enabled = False

strIn = txtStringPlain.Text
strOut = objEncryption.EncryptString_KeyPair(strIn) 'encrypt it using the Exchange key pair

With txtStringEncrypt
    .Text = strOut
    .SelStart = 0
    .SelLength = 65535
    .SetFocus
End With

End Sub

Private Sub cmdKeyPairFileDecrypt_Click()
Dim sngTim As Single
Dim strFileAndPathName As String

cmdKeyPairFileDecrypt.Enabled = False

CommonDialog1.Filter = "" 'no filter
CommonDialog1.FileName = ""
CommonDialog1.ShowOpen
strFileAndPathName = CommonDialog1.FileName

sngTim = Timer

objEncryption.DecryptFile_KeyPair strFileAndPathName, strFileAndPathName & "DECRYPTED"  'decrypt it using the Exchange key pair

MsgBox "Finished decrypting the file.  Remove the 'DECRYPTED' file extension to test the decrypted file. " & vbNewLine & "Time elapsed: " & Timer - sngTim

cmdKeyPairFileDecrypt.Enabled = True

End Sub

Private Sub cmdKeyPairFileEncrypt_Click()
Dim sngTim As Single
Dim strFileAndPathName As String

cmdKeyPairFileEncrypt.Enabled = False

CommonDialog1.Filter = "" 'no filter
CommonDialog1.FileName = ""
CommonDialog1.ShowOpen
strFileAndPathName = CommonDialog1.FileName

sngTim = Timer

objEncryption.EncryptFile_KeyPair strFileAndPathName, strFileAndPathName & "ENCRYPTED" 'encrypt it using the Exchange key pair

MsgBox "Finished encrypting the file.  Remove the 'ENCRYPTED' file extension to test the encrypted file. " & vbNewLine & "Time elapsed: " & Timer - sngTim

cmdKeyPairFileEncrypt.Enabled = True

End Sub

Private Sub cmdKeyPairFileSign_Click()
Dim strFileAndPathName As String

cmdKeyPairFileSign.Enabled = False

CommonDialog1.Filter = "" 'no filter
CommonDialog1.FileName = ""
CommonDialog1.ShowOpen
strFileAndPathName = CommonDialog1.FileName

objEncryption.SignFile_KeyPair strFileAndPathName, strFileAndPathName & "SIGNED" 'sign it using the Signature key pair

MsgBox "Finished signing the file."

cmdKeyPairFileSign.Enabled = True

End Sub

Private Sub cmdKeyPairFileValidate_Click()
Dim strFileAndPathName As String

cmdKeyPairFileValidate.Enabled = False

CommonDialog1.Filter = "" 'no filter
CommonDialog1.FileName = ""
CommonDialog1.ShowOpen
strFileAndPathName = CommonDialog1.FileName

objEncryption.ValidateFile_KeyPair strFileAndPathName, strFileAndPathName & "VALIDATED"  'validate it using the Signature key pair

MsgBox "Finished validating the file.  The signature has been stripped from the file with the 'VALIDATED' file extension."

cmdKeyPairFileValidate.Enabled = True

End Sub

Private Sub cmdKeyPairSign_Click()

cmdKeyPairValidate.Enabled = True
cmdKeyPairSign.Enabled = False

objEncryption.SignString_KeyPair txtStringPlain
MsgBox "The text in the upper text box has been signed.  The signature has been stored internally."

End Sub

Private Sub cmdKeyPairValidate_Click()

cmdKeyPairSign.Enabled = True
cmdKeyPairValidate.Enabled = False

objEncryption.ValidateString_KeyPair txtStringPlain
MsgBox "The text in the upper text box has been validated."

End Sub

Private Sub cmdListProviders_Click()
txtStringPlain.Text = objEncryption.ListAvailableProviders

End Sub

Private Sub cmdSwitch_Click()

Unload frmEncryptKeyPair
frmEncryptSingleKey.Show

End Sub

Private Sub Form_Load()

Set objEncryption = New clsCryptoAPIandCompression
objEncryption.SessionStart

frmKeyLength.Visible = True

End Sub

Private Sub Form_Unload(Cancel As Integer)

objEncryption.SessionEnd
Set objEncryption = Nothing

Unload frmKeyLength
Set frmKeyLength = Nothing

End Sub

