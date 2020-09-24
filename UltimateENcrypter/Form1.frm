VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   7125
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   2415
      Left            =   120
      TabIndex        =   6
      Top             =   3480
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   7440
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   480
      Width           =   2535
   End
   Begin RichTextLib.RichTextBox text2 
      Height          =   2415
      Left            =   120
      TabIndex        =   3
      Top             =   3480
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   4260
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":0000
   End
   Begin RichTextLib.RichTextBox text1 
      Height          =   2535
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   4471
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":00BA
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Decrypt"
      Height          =   495
      Left            =   4920
      TabIndex        =   1
      Top             =   2760
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Encrypt"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label passwordf 
      Caption         =   "Label1"
      Height          =   495
      Left            =   2280
      TabIndex        =   5
      Top             =   2760
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
     Dim Serpent As New clsSerpent
     Dim Blowfish As New clsBlowfish
     Dim cheshire As New clsCheshire
     Dim a As String
Dim filepathina As String
Dim filepathouta As String
Dim le As Integer
Private Sub Command1_Click()
vbf = text1.Text
Form1.passwordf.Caption = InputBox("What do you want the password to be")
DoEvents
DoEvents
    '  encryptit (text1.text)
whereareweat = 1
Randomize
a = text1.Text
'a = UCase(a)
For i = 1 To Len(text1.Text)
    mystr = Left(text1.Text, whereareweat)
    mystr = Right(mystr, 1)
    'MsgBox (mystr)
        newname = ""
For Counter = 1 To 10
    Char = Int(Rnd * 86) + 6
    newname = newname & Chr(Char)
Next Counter
    text2.Text = text2.Text & newname & mystr
        whereareweat = whereareweat + 1
    Next i
       '   BTMEncrypt (text2.text)

    CaesarShiftencode (text2.Text)
text2.Text = Serpent.EncryptString(text2.Text, passwordf.Caption)
    ReverseString (text2.Text)
text2.Text = cheshire.CEncode(text2.Text, passwordf.Caption)
HexEncode (text2.Text)
Eyncryptlevel1 (text2.Text)
'Form1.text2.Text = Blowfish.EncryptString(Form1.text2.Text, Form1.passwordf.Caption, False)

TEncrypt (text2.Text)
Open App.Path & "\" & "temp.txt" For Output As #1
Write #1, text2.Text
Close #1
'Form1.text2.Text = Blowfish.EncryptString(Form1.text2.Text, Form1.passwordf.Caption, False)
filepathina = App.Path & "\" & "temp.txt"
filepathouta = InputBox("file out")
le = Val(9)
text2.Text = CompressFile(filepathina, filepathouta, le)
text1.Text = vbf
'Steganalysis/
End Sub

Private Sub Command2_Click()
filepathina = InputBox("file in")
filepathouta = App.Path & "\" & "yo.txt"
le = Val(9)
text1.Text = DecompressFile(filepathina, filepathouta)
Form1.passwordf.Caption = InputBox("What do you want the password to be")
DoEvents
'Form1.text1.Text = Blowfish.DecryptString(Form1.text1.Text, Form1.passwordf.Caption, False)
Open App.Path & "\" & "yo.txt" For Input As #1
Input #1, a
Close #1
text1.Text = a
TDecrypt (text1.Text)
'Form1.text1.Text = Blowfish.DecryptString(Form1.text1.Text, Form1.passwordf.Caption, False)
UnEyncryptlevel1 (text1.Text)
Form1.text1.Text = HexDecode(text1.Text)
text1.Text = cheshire.CDecode(text1.Text, passwordf.Caption)
ReverseString (text1.Text)
text1.Text = Serpent.DecryptString(text1.Text, passwordf.Caption)
CaesarShiftdecode (text1.Text)

'decryptit (text1.text)
'BTMdEcrypt (text1.text)
    lenghtofstring = Len(Form1.text1.Text)
    anystring = Form1.text1.Text
        lenghtofstring = lenghtofstring - 10
        anystring = Right(anystring, lenghtofstring)
        Form1.Text3.Text = ""

For i = 1 To 1000
letter = Left(anystring, 1)
Form1.Text3.Text = Form1.Text3.Text & letter
If lenghtofstring - 11 < 1 Then
i = 1000
Else
        lenghtofstring = lenghtofstring - 11
        anystring = Right(anystring, lenghtofstring)
End If
Next i
Form1.text2.Text = Form1.Text3.Text

End Sub

