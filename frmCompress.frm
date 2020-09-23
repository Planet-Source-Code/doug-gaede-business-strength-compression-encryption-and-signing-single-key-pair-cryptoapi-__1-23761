VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCompress 
   Caption         =   "Zlib Compression Demo"
   ClientHeight    =   2760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   ScaleHeight     =   2760
   ScaleWidth      =   5445
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4800
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      InitDir         =   "c:"
   End
   Begin VB.TextBox txtLevel 
      Height          =   285
      Left            =   2880
      TabIndex        =   3
      Text            =   "9"
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Compress String (CompressString example)"
      Height          =   435
      Left            =   720
      TabIndex        =   2
      Top             =   2160
      Width           =   3915
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Compress Byte Array (CompressByteArray example)"
      Height          =   435
      Left            =   720
      TabIndex        =   1
      Top             =   1680
      Width           =   3915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Compress && Decompress File"
      Height          =   435
      Left            =   720
      TabIndex        =   0
      Top             =   1080
      Width           =   3915
   End
   Begin VB.Label Label2 
      Caption         =   $"frmCompress.frx":0000
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label Label1 
      Caption         =   "Compression:"
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmCompress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objCompress As clsCryptoAPIandCompression
Dim lnglngResult As Long

Private Sub Command1_Click()
Dim strFileAndPathName As String

CommonDialog1.FileName = ""
CommonDialog1.ShowOpen
strFileAndPathName = CommonDialog1.FileName

lnglngResult = objCompress.CompressFile(strFileAndPathName, strFileAndPathName & "COMPRESSED", Val(txtLevel))
lnglngResult = objCompress.DecompressFile(strFileAndPathName & "COMPRESSED", strFileAndPathName & "DECOMPRESSED")

MsgBox "Done!  Remove the DECOMPRESSED file extension to test the file."

End Sub

Private Sub Command2_Click()

Dim TheBytes() As Byte
Dim lngCnt As Long
Dim intChar As Integer
Dim lngC As Long

ReDim TheBytes(102400 - 1) 'allocate precisely 100K

'Fill our bytes from random junk.
lngCnt = 10
For lngC = 0 To UBound(TheBytes)
    If lngCnt = 10 Then
      intChar = Int((256) * Rnd + 0)
      lngCnt = 0
    End If
    TheBytes(lngC) = intChar
Next lngC

MsgBox "Original size: " & CStr(UBound(TheBytes) + 1) & " bytes"

lnglngResult = objCompress.CompressByteArray(TheBytes(), Val(txtLevel))

MsgBox "Compressed size: " & objCompress.ValueCompressedSize & " bytes"

lnglngResult = objCompress.DecompressByteArray(TheBytes(), objCompress.ValueDecompressedSize)

MsgBox "Decompressed size: " & objCompress.ValueDecompressedSize & " bytes"

'cleanup
Erase TheBytes

End Sub

Private Sub Command3_Click()
Dim strOurString As String

strOurString = "I'm just a bill, just a lonely old bill.  Sittin' up on Capital Hill.  This is a string. It is just a normal ordinary string."

MsgBox strOurString & " [Length: " & Len(strOurString) & "]"

lnglngResult = objCompress.CompressString(strOurString, Val(txtLevel))

MsgBox "Compressed string (may not display all characters due to invisible characters [Length: " & objCompress.ValueCompressedSize & "]: " & strOurString

lnglngResult = objCompress.DecompressString(strOurString, objCompress.ValueDecompressedSize)

MsgBox "The string, decompressed: " & strOurString

End Sub

Private Sub Form_Load()
Set objCompress = New clsCryptoAPIandCompression

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set objCompress = Nothing

End Sub

