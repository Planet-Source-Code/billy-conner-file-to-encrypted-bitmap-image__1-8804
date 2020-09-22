VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Encrypt 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File to Picture Encrypter"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   265
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   395
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   5775
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   240
         TabIndex        =   16
         Top             =   1320
         Width           =   3855
      End
      Begin VB.CommandButton Command4 
         Caption         =   "..."
         Height          =   255
         Left            =   4200
         TabIndex        =   3
         Top             =   1320
         Width           =   255
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Encrypt"
         Height          =   495
         Index           =   0
         Left            =   4560
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Decrypt"
         Height          =   495
         Index           =   1
         Left            =   4560
         TabIndex        =   6
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   255
         Left            =   4200
         TabIndex        =   1
         Top             =   480
         Width           =   255
      End
      Begin VB.CommandButton Command3 
         Caption         =   "..."
         Height          =   255
         Left            =   4200
         TabIndex        =   5
         Top             =   2160
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   3855
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2040
         MaxLength       =   20
         TabIndex        =   7
         Top             =   2880
         Width           =   2055
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   240
         TabIndex        =   4
         Top             =   2160
         Width           =   3855
      End
      Begin VB.Label Label5 
         Caption         =   "Picture To Decrypt:"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "File To Encrypt:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Encryption String:"
         Height          =   255
         Left            =   600
         TabIndex        =   12
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "File To Decrypt Picture To:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1920
         Width           =   1935
      End
   End
   Begin MSComctlLib.ProgressBar PB1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   3720
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   5280
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   0
      ScaleHeight     =   7
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   7
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label4 
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   3360
      Width           =   4455
   End
End
Attribute VB_Name = "Encrypt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This Program does not have error handling.
'some Parts of the Bitmap headers may not be totally correct in all cases.
' i am still learning the header values.  but this works for me.

Option Explicit
Private Type BITMAPFILEHEADER '14 bytes
        bfType As Integer '4d42 backwards is 424d = BM
        bfSize As Long 'File Size
        bfReserved1 As Integer '0
        bfReserved2 As Integer '0
        bfOffBits As Long '54
End Type

Private Type BITMAPINFOHEADER '40 bytes
        biSize As Long '40
        biWidth As Long
        biHeight As Long
        biPlanes As Integer ' 1
        biBitCount As Integer '1 4 8 24
        biCompression As Long '0
        biSizeImage As Long '0 if the bitmap is in the BI_RGB format.
        biXPelsPerMeter As Long '3780
        biYPelsPerMeter As Long '3780
        biClrUsed As Long
        biClrImportant As Long '0
End Type
Dim BmpF As BITMAPFILEHEADER
Dim BmpI As BITMAPINFOHEADER
'FL is file length
'FLD3 is Filelength Divided By 3
'Leftover Is the amount left overafter dividing FL by 3
'SQ is the sq height/Width Of The Picture
Dim dfile$
Private Sub Command1_Click(Index As Integer)
Dim TimerBeginPos As Double
TimerBeginPos = Timer
If Len(Text2) = 0 Then
    Text2.SetFocus: Exit Sub
End If
If Len(Text4) = 0 Then
    Text4.SetFocus: Exit Sub
End If
Frame1.Enabled = False
Select Case Index
    Case 0
        If Len(Text1) = 0 Then
            Text1.SetFocus: Exit Sub
        End If
        Call EncryptIt
        Label4.Caption = "Finished Encrypting... " & Format(Timer - TimerBeginPos, "###.##") & " Seconds."
    Case 1
        If Len(Text3) = 0 Then
            Text3.SetFocus: Exit Sub
        End If
        dfile$ = Text3
        Call Decrypt
        Label4.Caption = "Finished Decrypting... " & Format(Timer - TimerBeginPos, "###.##") & " Seconds."
End Select
Frame1.Enabled = True
End Sub
Function Crypt(KeyCode As String, pos As Long) As Byte
'This is the Function that encrypts the data.
Dim tmp As String
Static Counter As Byte
If pos = 1 Then Counter = 0
Counter = ((Counter + 1) Mod 2) + 1
tmp = Mid$(KeyCode, (pos Mod Len(KeyCode)) + 1, 1)
Crypt = ((Asc(StrReverse(tmp)) ^ Counter)) Mod 256
End Function

Sub EncryptIt()
Dim FF As Byte
Dim FL As Long
Dim FLD3 As Long
Dim header As String
Dim filedata As String
Dim LeftOver As Byte
Dim Sq As Long
Dim BmpW As Long, BmpH As Long
Dim KeyCodeData As String
Dim i As Long
Dim KeyCodeAsc As Byte
Dim red As Byte, green As Byte, blue As Byte
Dim MainCryptingData As String
'On Local Error Resume Next
FF = FreeFile
'---Get the file data---
Open Text1.Text For Binary As FF
    FL = LOF(FF) ' find length of file
    If FL = 0 Then
        Close FF
        Exit Sub
    End If
    filedata = String(FL, Chr(0))
    Get FF, 1, filedata
Close FF
' end of the Getting

FLD3 = Int(FL / 3) ' divide by 3
LeftOver = FL - (FLD3 * 3) 'get remaining bytes
PB1.Value = 0
PB1.Max = FL ' + Leftover
Sq = Int(Sqr(FLD3 + LeftOver))
If Sq < 2 Then Sq = 2
Label4.Caption = "Encrypting File...     " & FL & " Bytes."
BmpW = Sq
BmpH = Sq
With BmpF
    .bfType = &H4D42
    .bfSize = 54 + ((BmpW * BmpH) * 3)
    .bfReserved1 = 0
    .bfReserved2 = 0
    .bfOffBits = 54
End With
With BmpI
    .biSize = 40
    .biWidth = BmpW
    .biHeight = BmpH
    .biPlanes = 1
    .biBitCount = 24
    .biCompression = 0
    .biSizeImage = 0
    .biXPelsPerMeter = 3780
    .biYPelsPerMeter = 3780
    .biClrUsed = 0
    .biClrImportant = 0
End With

header = Convert(BmpF.bfType, 2)
header = header & Convert(BmpF.bfSize, 4)
header = header & Convert(BmpF.bfReserved1, 2)
header = header & Convert(BmpF.bfReserved2, 2)
header = header & Convert(BmpF.bfOffBits, 4)
header = header & Convert(BmpI.biSize, 4)
header = header & Convert(BmpI.biWidth, 4)
header = header & Convert(BmpI.biHeight, 4)
header = header & Convert(BmpI.biPlanes, 2)
header = header & Convert(BmpI.biBitCount, 2)
header = header & Convert(BmpI.biCompression, 4)
header = header & Convert(BmpI.biSizeImage, 4)
header = header & Convert(BmpI.biXPelsPerMeter, 4)
header = header & Convert(BmpI.biYPelsPerMeter, 4)
header = header & Convert(BmpI.biClrUsed, 4)
header = header & Convert(BmpI.biClrImportant, 4)
KeyCodeData = Text2
KeyCodeData = KeyCodeData & Chr(StringSumWithMod(KeyCodeData))
KeyCodeData = Chr(StringSumWithMod(KeyCodeData, 15)) & KeyCodeData
Open Text4 For Output As #1
    Print #1, header;
Dim LeftOverChar As String
i = 1
Do
    KeyCodeAsc = Crypt(KeyCodeData, i)
    If i + 1 < FL Then
    red = Asc(Mid$(filedata, i, 1)) Xor KeyCodeAsc
    green = Asc(Mid$(filedata, i + 1, 1)) Xor KeyCodeAsc + 1
    blue = Asc(Mid$(filedata, i + 2, 1)) Xor KeyCodeAsc + 2
    LeftOverChar = Chr(blue) & Chr(green) & Chr(red)
    Else
        Dim pow As Byte
        pow = IIf(LeftOver = 1, 2, 1)
        red = Asc(Mid$(filedata, i, 1)) Xor KeyCodeAsc + pow
        LeftOverChar = Chr(red)
        If LeftOver = 2 Then
            green = Asc(Mid$(filedata, i + 1, 1)) Xor KeyCodeAsc + 2
            LeftOverChar = Chr(green) & LeftOverChar
        End If
    End If
    MainCryptingData = MainCryptingData & LeftOverChar
    If (i Mod 100) = 0 Then
        DoEvents
        Print #1, MainCryptingData;
        MainCryptingData = ""
    End If
    PB1.Value = i
    i = i + 3
Loop Until i >= FL + 1
If Len(MainCryptingData) Then Print #1, MainCryptingData;
Close #1
Exit Sub
End Sub
Sub Decrypt()
Dim FF As Byte 'file number
Dim header As String * 54
Dim BmpType As String * 2
Dim info As String
Dim KeyCodeData As String 'the password is stored here temporarily
Dim i As Long
Dim KeyCodeAsc As Byte
Dim red As Byte, green As Byte, blue As Byte
Dim MainCryptingData As String
dfile$ = Text3
'Px = 0: Py = 0
On Local Error Resume Next
FF = FreeFile
Open Text4.Text For Binary As FF
If LOF(FF) < 55 Then GoTo InvalidBitmap
Get FF, 1, header
BmpType = Left(header, 2)
If BmpType <> "BM" Then GoTo InvalidBitmap
info = String(LOF(FF) - 54, Chr(0))
Get FF, 55, info
Close FF
Label4.Caption = "Decrypting Picture...     " & Len(info) & " Bytes."
PB1.Value = 0
PB1.Max = Len(info)
KeyCodeData = Text2
KeyCodeData = KeyCodeData & Chr(StringSumWithMod(KeyCodeData))
KeyCodeData = Chr(StringSumWithMod(KeyCodeData, 15)) & KeyCodeData
Open dfile$ For Output As #1
For i = 1 To Len(info) Step 3
    KeyCodeAsc = Crypt(KeyCodeData, i)
    MainCryptingData = MainCryptingData & Chr(Asc(Mid(info, i + 2, 1)) Xor KeyCodeAsc)
    MainCryptingData = MainCryptingData & Chr(Asc(Mid(info, i + 1, 1)) Xor KeyCodeAsc + 1)
    MainCryptingData = MainCryptingData & Chr(Asc(Mid(info, i, 1)) Xor KeyCodeAsc + 2)
    If (i Mod 100) = 0 Then
        DoEvents
        Print #1, MainCryptingData;
        MainCryptingData = ""
    End If
    PB1.Value = i
Next i
PB1.Value = i
If Len(MainCryptingData) Then Print #1, MainCryptingData;
Close #1
Exit Sub
InvalidBitmap:
Close FF
MsgBox "Invalid Bitmap!"
End Sub

Private Sub Command4_Click()
CD1.InitDir = Text4
CD1.ShowSave
If CD1.FileName <> "" Then Text4.Text = CD1.FileName
End Sub

Private Sub Command2_Click()
CD1.FileName = Text1
CD1.ShowOpen
If CD1.FileName <> "" Then Text1.Text = CD1.FileName
End Sub

Private Sub Command3_Click()
CD1.InitDir = Text3
CD1.ShowOpen
If CD1.FileName <> "" Then Text3.Text = CD1.FileName
End Sub

Private Sub Form_Load()
Randomize Timer
Text1 = App.Path & "\File.txt"
Text4 = App.Path & "\Encrypt.bmp"
Text3 = App.Path & "\Back2File.txt"
End Sub

Private Function Convert(Num As Variant, Chars As Byte) As String
Dim hx As String
Dim i As Integer
Dim cdata As String
hx = Hex(Num)
If Len(hx) Mod 2 Then hx = "0" & hx
hx = String$((Chars * 2) - Len(hx), "0") & hx
For i = ((Chars * 2) - 1) To 1 Step -2
    cdata = cdata & Chr(hex2dec(Mid$(hx, i, 2)))
Next i
Convert = cdata
End Function

Private Function hex2dec(hx As String) As Byte
hex2dec = Val("&H" & hx)
End Function

Private Function StringSumWithMod(Str As String, Optional DaMod As Byte = 255) As Byte
Dim Daval As Integer
Dim i As Integer
For i = 1 To Len(Str)
    Daval = Daval + Asc(Mid(Str, i, 1))
Next i
StringSumWithMod = Daval Mod DaMod
End Function
Private Sub Form_Unload(Cancel As Integer)
End
End Sub
