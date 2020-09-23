VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   2295
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   720
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get Hash"
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This code shows how to generate a hash of a
'string of data. Hashes are extremely usefull for determining
'whether a transmission or file has been altered.' There are 2
'algorithms available in this example - MD5 and SHA

'The MD5 returns a 16 character (128 bit) hash and the
'SHA returns a 20 character (160 bit) hash.
'You can use hashes to create crypto keys and
'to verify integrity of packets when using
'winsock (UDP especially).

'Inputs: The function takes two parameters: (1) data - string
'of data to get hash of & (2) hashType (0 or 1). 0 is MD5
'algorithm and 1 is SHA.

'Returns: returns ASCII string which is the hash.  Some characters
'may not display properly.  Convert to hex if you want to display
'a hash.

Private Const ALG_CLASS_HASH = 32768
Private Const ALG_TYPE_ANY = 0
Private Const ALG_SID_MD5 = 3
Private Const ALG_SID_SHA = 4
Private Const MS_DEF_PROV = "Microsoft Base Cryptographic Provider v1.0"
Private Const PROV_RSA_FULL = 1
Private Const HP_HASHVAL = &H2
Private Const CRYPT_NEWKEYSET = &H8
Private Const CALG_MD5 = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_MD5)
Private Const CALG_SHA = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_SHA)

Private Declare Function CryptCreateHash Lib "advapi32.dll" ( _
    ByVal hProv As Long, _
    ByVal Algid As Long, _
    ByVal hKey As Long, _
    ByVal dwFlags As Long, _
    phHash As Long) As Long

Private Declare Function CryptHashData Lib "advapi32.dll" ( _
    ByVal hHash As Long, _
    ByVal pbData As String, _
    ByVal dwDataLen As Long, _
    ByVal dwFlags As Long) As Long
 
Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" ( _
    phProv As Long, _
    ByVal pszContainer As String, _
    ByVal pszProvider As String, _
    ByVal dwProvType As Long, _
    ByVal dwFlags As Long) As Long
    
Private Declare Function CryptReleaseContext Lib "advapi32.dll" ( _
    ByVal hProv As Long, _
    ByVal dwFlags As Long) As Long

Private Declare Function CryptGetHashParam Lib "advapi32.dll" ( _
    ByVal hHash As Long, _
    ByVal dwParam As Long, _
    ByVal pbData As String, _
    pdwDataLen As Long, _
    ByVal dwFlags As Long) As Long
    
    Private cryptContext As Long
    
                            
Public Function getHash(data As String, hashType As Integer) As String
Dim ht As Long
Dim sTemp As String
Dim sProv As String
Dim hLen As Long
Dim h As String
Dim hl As Long

'get hash type
If hashType = 0 Then
    'MD5
    ht = CALG_MD5
    hLen = 16
ElseIf hashType = 1 Then
    'SHA
    hLen = 20
    ht = CALG_SHA
Else
    getHash = ""
    Exit Function
End If
'--- Prepare string buffers
sTemp = vbNullChar
sProv = MS_DEF_PROV & vbNullChar
'---Gain Access To CryptoAPI


If Not CBool(CryptAcquireContext(cryptContext, sTemp, sProv, PROV_RSA_FULL, 0)) Then
    If Not CBool(CryptAcquireContext(cryptContext, sTemp, sProv, PROV_RSA_FULL, CRYPT_NEWKEYSET)) Then
        getHash = ""
        Exit Function
    End If
End If
'Create Empty hash object

If Not CBool(CryptCreateHash(cryptContext, ht, 0, 0, hl)) Then
    getHash = ""
    Exit Function
End If
'Hash the input string.

If Not CBool(CryptHashData(hl, data, Len(data), 0)) Then
    getHash = ""
    Exit Function
End If
h = String(20, vbNull)
'Get hash val

If Not CBool(CryptGetHashParam(hl, HP_HASHVAL, h, hLen, 0)) Then
    getHash = ""
    Exit Function
End If
getHash = h

'Release provider handle
If (cryptContext <> 0) Then Call CryptReleaseContext(cryptContext, 0)

End Function

Private Sub Command1_Click()
Text1 = getHash(Text1.Text, 1)

End Sub
