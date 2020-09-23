VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   2145
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Form1.frx":0000
      Top             =   720
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
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

Private cryptContext As Long
Private Const MS_DEF_PROV = "Microsoft Base Cryptographic Provider v1.0"
Private Const PROV_RSA_FULL = 1
Private Const CRYPT_NEWKEYSET = &H8

Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" ( _
    phProv As Long, _
    ByVal pszContainer As String, _
    ByVal pszProvider As String, _
    ByVal dwProvType As Long, _
    ByVal dwFlags As Long) As Long
    
Private Declare Function CryptReleaseContext Lib "advapi32.dll" ( _
    ByVal hProv As Long, _
    ByVal dwFlags As Long) As Long

Private Declare Function CryptGenRandom Lib "advapi32.dll" _
    (ByVal hProv As Long, _
    ByVal dwLen As Long, _
    ByVal pbBuffer As String) As Long

' Inputs: saltLen - Length of data to be returned
'
' Returns: returns a string of random data.  Some data may not
'          be displayable.  Convert to hex first if you must
'          display the data accurately.

Public Function getSalt(saltLen As Long) As String
    Dim sv As String
    Dim sTemp As String
    Dim sProv As String
    
    'make sure valid positive long is sent
    If saltLen <= 0 Then
        getSalt = ""
        Exit Function
    End If
    '--- Prepare string buffers
    sTemp = vbNullChar
    sProv = MS_DEF_PROV & vbNullChar
    '---Gain Access To CryptoAPI
    If Not CBool(CryptAcquireContext(cryptContext, sTemp, sProv, PROV_RSA_FULL, 0)) Then
        If Not CBool(CryptAcquireContext(cryptContext, sTemp, sProv, PROV_RSA_FULL, CRYPT_NEWKEYSET)) Then
            getSalt = ""
            Exit Function
        End If
    End If
    
    'Make the Salt
    sv = String(saltLen, vbNull)
        
    If Not CBool(CryptGenRandom(cryptContext, saltLen, sv)) Then
        getSalt = ""
    Else
        getSalt = sv
    End If
End Function

Private Sub Command1_Click()

Text1 = getSalt(Val(Text1))

End Sub
