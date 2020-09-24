VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCheshire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Encoding/Decoding Written by Säure (Davy Cook) 2001
'http://ghostlyembers.8m.com/
'Utilizes Fredrik Qvarfort's Huffman compression and CRC routines (C) 2000

Public Enum CRCAlgorithms
  CRC32
End Enum
Private Type HUFFMANTREE
  ParentNode As Integer
  RightNode As Integer
  LeftNode As Integer
  Value As Integer
  Weight As Long
End Type
Private Type ByteArray
  Count As Byte
  Data() As Byte
End Type

Private Const PROGRESS_CALCFREQUENCY = 7
Private Const PROGRESS_CALCCRC = 5
Private Const PROGRESS_ENCODING = 88
Private Const PROGRESS_DECODING = 89
Private Const PROGRESS_CHECKCRC = 11
Private m_CRC32 As Long
Private m_CRC32Asm() As Byte
Private m_CRC32Init As Boolean
Private m_CRC32Table(0 To 255) As Long

Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Event Prog(Prec As Integer)
Event Progress(Procent As Integer)

Function CEncode(istring As String, pw As String) As String
On Error GoTo Err_CEncode
Dim temp As String, x As Long, outp As String
Dim crcc As String
crcc = StrCRC(istring)
temp = Dis(istring, True)
Dim z As Long
For x = 1 To Len(temp)
z = z + 1
If z = (Len(pw) + 1) Then z = 1
outp = outp + Chr$(Asc(Mid$(temp, x, 1)) Xor Asc(Mid$(pw, z, 1)))
Next
outp = outp + crcc
CEncode = Dis(outp)
Exit Function
Err_CEncode:
CEncode = "Program Error!"
End Function

Function CDecode(istring As String, pw As String) As String
On Error GoTo Err_CDecode
Dim temp As String, x As Long, outp As String
Dim z As Long, crcc As String
temp = Asm(istring)
crcc = Right$(temp, 4)
temp = Left$(temp, Len(temp) - 4)
For x = 1 To Len(temp)
z = z + 1
If z = (Len(pw) + 1) Then z = 1
outp = outp + Chr$(Asc(Mid$(temp, x, 1)) Xor Asc(Mid$(pw, z, 1)))
Next
CDecode = Asm(outp, True)
If StrCRC(CDecode) <> crcc Then CDecode = "String CRC Error!"
Exit Function
Err_CDecode:
CDecode = "Program Error!"
End Function

'4 bytes
Private Function StrCRC(istring As String) As String
Dim m As String
m = Hex2Chr(Hex$(CRC_CalculateString(istring)))
If Len(m) = 3 Then m = Chr$(0) + m
If Len(m) = 2 Then m = Chr$(0) + Chr$(0) + m
If Len(m) = 1 Then m = Chr$(0) + Chr$(0) + Chr$(0) + m
If Len(m) = 0 Then m = Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0)
StrCRC = m
End Function

'8 bytes
Private Function FileCRC(ifile As String) As String
Dim m As String
m = Hex2Chr(Hex$(CRC_CalculateFile(ifile)))
If Len(m) = 3 Then m = Chr$(0) + m
If Len(m) = 2 Then m = Chr$(0) + Chr$(0) + m
If Len(m) = 1 Then m = Chr$(0) + Chr$(0) + Chr$(0) + m
If Len(m) = 0 Then m = Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0)
FileCRC = m
m = Hex2Chr(Hex$((FileLen(ifile))))
If Len(m) = 3 Then m = Chr$(0) + m
If Len(m) = 2 Then m = Chr$(0) + Chr$(0) + m
If Len(m) = 1 Then m = Chr$(0) + Chr$(0) + Chr$(0) + m
If Len(m) = 0 Then m = Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0)
FileCRC = FileCRC + m
End Function

Private Function Hex2Chr(istring As String) As String
Dim m As String, x As Byte, v As String * 2, r As Byte, f As Byte, y As Byte
Dim a As String * 1
For x = 1 To Len(istring) / 2
v = Mid$(istring, x * 2 - 1, 2)
f = 0
For y = 1 To 2
a = Mid$(v, y, 1)
Select Case a
 Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0"
  r = Val(a)
 Case "A"
  r = 10
 Case "B"
  r = 11
 Case "C"
  r = 12
 Case "D"
  r = 13
 Case "E"
  r = 14
 Case "F"
  r = 15
End Select
f = f + (r * (16 ^ (2 - y)))
Next
m = m + Chr$(f)
Next
Hex2Chr = m
End Function

Private Function Dis(istring As String, Optional padd As Boolean = False) As String
Dim strlen As Long, len16 As Long, lenext As Long, x As Byte
strlen = Len(istring)
If strlen < 16 And padd Then
 GoTo Short_Padd
ElseIf strlen < 16 And Not padd Then
 Dis = istring
 Exit Function
End If
Dim parts(1 To 21) As String
len16 = strlen \ 16
lenext = strlen - (len16 * 16)
For x = 1 To 16
parts(x) = Mid$(istring, (x - 1) * len16 + 1, len16)
Next
parts(17) = Right$(istring, lenext)
If padd Then
Dim ai As Integer
ai = len16
parts(18) = RC(ai)
parts(19) = RC(ai)
End If
If Not padd Then
Dis = parts(17) + parts(14) + parts(10) + parts(3) + parts(7) + parts(16) + parts(11) + parts(2) + parts(6) + parts(1) + parts(4) + parts(5) + parts(15) + parts(9) + parts(12) + parts(8) + parts(13)
Else
Dis = parts(17) + parts(14) + parts(10) + parts(3) + parts(7) + parts(18) + parts(16) + parts(11) + parts(2) + parts(6) + parts(1) + parts(4) + parts(5) + parts(15) + parts(9) + parts(19) + parts(12) + parts(8) + parts(13)
End If
Erase parts
Exit Function
Short_Padd:
Dis = RC(3) + istring + RC(3)
End Function

Private Function Asm(istring As String, Optional padded As Boolean = False) As String
Dim strlen As Double, len16 As Double, lenext As Double, x As Double
strlen = Len(istring)
If strlen < 16 And Not padded Then
 Asm = istring
 Exit Function
ElseIf strlen < 16 And padded Then
 GoTo paddremove
End If
Dim parts(1 To 19) As String
If Not padded Then
 len16 = strlen \ 16
 lenext = strlen - (len16 * 16)
 For x = 1 To 16
 parts(x) = Mid$(istring, lenext + 1 + len16 * (x - 1), len16)
 Next
 parts(17) = Left$(istring, lenext)
 Asm = parts(9) + parts(7) + parts(3) + parts(10) + parts(11) + parts(8) + parts(4) + parts(15) + parts(13) + parts(2) + parts(6) + parts(14) + parts(16) + parts(1) + parts(12) + parts(5) + parts(17)
Else
 len16 = strlen \ 18
 lenext = strlen - (len16 * 18)
 For x = 1 To 18
 parts(x) = Mid$(istring, lenext + 1 + len16 * (x - 1), len16)
 Next
 parts(19) = Left$(istring, lenext)
 Asm = parts(10) + parts(8) + parts(3) + parts(11) + parts(12) + parts(9) + parts(4) + parts(17) + parts(14) + parts(2) + parts(7) + parts(16) + parts(18) + parts(1) + parts(13) + parts(6) + parts(19)
End If
Erase parts
Exit Function
paddremove:
Asm = Mid$(istring, 4, strlen - 6)
End Function
Private Function RC(num As Integer) As String
Randomize Timer
Dim x As Integer
For x = 1 To num
RC = RC + Chr$(Int(Rnd * 256))
Next
End Function
Private Function PWM(istring As String) As String
Dim x As Integer, y As Integer
x = Len(istring)
If x > 39 Then istring = "mErcUry" + istring + "dOObEr"
If x < 40 And x > 31 Then istring = "~[BlAB" + istring + "wLR`{"
If x < 32 And x > 23 Then istring = "{dDk|e" + istring + "{}_]cK"
If x < 24 And x > 15 Then istring = "Ll33UF" + istring + "^wrjOW"
If x < 16 And x > 7 Then istring = "MosF`S" + istring + "`sa|ev"
If x < 8 And x > 0 Then istring = "ABm~@V" + istring + "Nj`Zfq"
x = Len(istring)
PWM = ""
For y = 1 To x
PWM = PWM + Chr$(Asc(Mid$(istring, y, 1)) Xor 42)
Next
End Function

Public Function Encode(ifile As String, ofile As String, pw As String, Optional comp As Boolean = True, Optional overwrite As Boolean = False) As String
On Error GoTo Encode_Err
If Not FileExist(ifile) Then
 Encode = "Failed: Inputted File does not exist."
 Exit Function
End If
If FileExist(ofile) And Not overwrite Then
 Encode = "Failed: File to Output to exists."
 Exit Function
ElseIf FileExist(ofile) And overwrite Then
 If FileExist(ofile) Then Kill ofile
End If
pw = PWM(pw)
Dim tfile As String
tfile = "$$temp$$"
If FileExist(tfile) Then Kill tfile
If comp Then
Huffman_EncodeFile ifile, tfile
Else
FileCopy ifile, tfile
End If
Dim chunk8k As Long, chunkex As Double, chunk As String * 8192
chunk8k = FileLen(tfile) \ 8192
chunkex = FileLen(tfile) - (chunk8k * 8192)
If chunkex = 0 Then GoTo GoMunk
ReDim chunkext(1 To chunkex) As String * 1
Dim chunkextall As String
GoMunk:
Dim ff As Integer, outa As String, fg As Integer, pos As Double, pos2 As Double, curr As Integer
ff = FreeFile()
Open tfile For Binary As #ff
fg = FreeFile()
Open ofile For Binary As #fg
pos = 1
pos2 = 1
If chunk8k <> 0 Then
For x = 1 To chunk8k
Get #ff, pos2, chunk
pos2 = pos2 + 8192
outa = CEncode(chunk, pw)
If outa = "Program Error!" Then GoTo Encode_Err
Put #fg, pos, outa
pos = pos + Len(outa)
curr = x / chunk8k * 100
RaiseEvent Prog(curr)
Next
End If
If chunkex <> 0 Then
chunkextall = ""
Get #ff, pos2, chunkext()
For x = 1 To UBound(chunkext())
chunkextall = chunkextall + chunkext(x)
Next
outa = CEncode(chunkextall, pw)
If outa = "Program Error!" Then GoTo Encode_Err
Put #fg, pos, outa
End If
Close #ff
Close #fg
If FileExist(tfile) Then Kill tfile
Encode = "Success"
Exit Function
Encode_Err:
Close #fg
Close #ff
If FileExist(ofile) Then Kill ofile
If FileExist(ofile) Then Kill tfile
Encode = "Program Error in Encoding!"
End Function

Public Function Decode(ifile As String, ofile As String, pw As String, Optional comp As Boolean = True, Optional overwrite As Boolean = False) As String
On Error GoTo Err_Decode
If Not FileExist(ifile) Then
 Decode = "Failed: Inputted File does not exist."
 Exit Function
End If
If FileExist(ofile) And Not overwrite Then
 Decode = "Failed: File to Output to exists."
 Exit Function
ElseIf FileExist(ofile) And overwrite Then
 If FileExist(ofile) Then Kill ofile
End If
Dim tfile As String
pw = PWM(pw)
tfile = "$$temp$$"
If FileExist(tfile) Then Kill tfile
Dim chunknum As Double, chunkext As Double
Dim ff As Integer, outa As String, fg As Integer, pos As Double, pos2 As Double
Dim chunk As String * 9220, curr As Integer
chunknum = FileLen(ifile) \ 9220
chunkext = FileLen(ifile) - (chunknum * 9220)
If chunkext = 0 Then GoTo Munk
ReDim chunke(1 To chunkext) As String * 1
Dim chunkex As String
Munk:
ff = FreeFile()
Open ifile For Binary As #ff
fg = FreeFile()
Open tfile For Binary As #fg
pos = 1
pos2 = 1
If chunknum <> 0 Then
For x = 1 To chunknum
Get #ff, pos2, chunk
pos2 = pos2 + 9220
outa = CDecode(chunk, pw)
If outa = "String CRC Error!" Then GoTo Err_Dec
If outa = "Program Error!" Then GoTo Funky_Err
Put #fg, pos, outa
pos = pos + Len(outa)
curr = x / chunknum * 100
RaiseEvent Prog(curr)
Next
End If
If chunkext <> 0 Then
chunkex = ""
Get #ff, pos2, chunke()
Get #ff, pos2, chunke()
For x = 1 To UBound(chunke())
chunkex = chunkex + chunke(x)
Next
outa = CDecode(chunkex, pw)
If outa = "String CRC Error!" Then GoTo Err_Dec1
If outa = "Program Error!" Then GoTo Funky_Err
Put #fg, pos, outa
End If
Close #ff
Close #fg
If comp Then
Huffman_Decodefile tfile, ofile
Else
FileCopy tfile, ofile
End If
If FileExist(tfile) Then Kill tfile
Decode = "Success"
Exit Function
Err_Dec:
Close #ff
Close #fg
If FileExist(tfile) Then Kill tfile
If FileExist(ofile) Then Kill ofile
Decode = "String CRC Error at the number" + Str$(x) + " chunk!"
Exit Function
Err_Dec1:
Close #ff
Close #fg
If FileExist(tfile) Then Kill tfile
If FileExist(ofile) Then Kill ofile
Decode = "String CRC Error at last chunk!"
Exit Function
Funky_Err:
Decode = "Program Error in Decoding!"
Close #ff
Close #fg
If FileExist(tfile) Then Kill tfile
If FileExist(ofile) Then Kill ofile
Exit Function
Err_Decode:
Decode = "Program Error in Decoding!"
Close #ff
Close #fg
If FileExist(tfile) Then Kill tfile
If FileExist(ofile) Then Kill ofile
End Function

Private Function AddBytes(ByteArrayA() As Byte) As Variant
  Dim ByteSize As Long
  On Local Error GoTo NoData
  ByteSize = UBound(ByteArrayA) - LBound(ByteArrayA) + 1
  On Local Error GoTo 0
  Call CallWindowProc(VarPtr(m_CRC32Asm(0)), VarPtr(m_CRC32), VarPtr(ByteArrayA(LBound(ByteArrayA))), VarPtr(m_CRC32Table(0)), ByteSize)
NoData:
  AddBytes = (Not m_CRC32)
End Function

Public Function CRC_CalculateFile(Filename As String) As Variant
  Dim Filenr As Integer
  Dim ByteArrayA() As Byte
  If (FileLen(Filename) = 0) Then Exit Function
  Filenr = FreeFile
  Open Filename For Binary As #Filenr
  ReDim ByteArrayA(0 To LOF(Filenr) - 1)
  Get #Filenr, , ByteArrayA()
  Close #Filenr
  m_CRC32 = &HFFFFFFFF
  CRC_CalculateFile = AddBytes(ByteArrayA)
End Function

Public Function CRC_CalculateString(Text As String)
  m_CRC32 = &HFFFFFFFF
  CRC_CalculateString = AddBytes(StrConv(Text, vbFromUnicode))
End Function

Private Sub InitializeCRC32()
  Dim i As Long
  Dim sASM As String
  m_CRC32Table(0) = &H0
  m_CRC32Table(1) = &H77073096
  m_CRC32Table(2) = &HEE0E612C
  m_CRC32Table(3) = &H990951BA
  m_CRC32Table(4) = &H76DC419
  m_CRC32Table(5) = &H706AF48F
  m_CRC32Table(6) = &HE963A535
  m_CRC32Table(7) = &H9E6495A3
  m_CRC32Table(8) = &HEDB8832
  m_CRC32Table(9) = &H79DCB8A4
  m_CRC32Table(10) = &HE0D5E91E
  m_CRC32Table(11) = &H97D2D988
  m_CRC32Table(12) = &H9B64C2B
  m_CRC32Table(13) = &H7EB17CBD
  m_CRC32Table(14) = &HE7B82D07
  m_CRC32Table(15) = &H90BF1D91
  m_CRC32Table(16) = &H1DB71064
  m_CRC32Table(17) = &H6AB020F2
  m_CRC32Table(18) = &HF3B97148
  m_CRC32Table(19) = &H84BE41DE
  m_CRC32Table(20) = &H1ADAD47D
  m_CRC32Table(21) = &H6DDDE4EB
  m_CRC32Table(22) = &HF4D4B551
  m_CRC32Table(23) = &H83D385C7
  m_CRC32Table(24) = &H136C9856
  m_CRC32Table(25) = &H646BA8C0
  m_CRC32Table(26) = &HFD62F97A
  m_CRC32Table(27) = &H8A65C9EC
  m_CRC32Table(28) = &H14015C4F
  m_CRC32Table(29) = &H63066CD9
  m_CRC32Table(30) = &HFA0F3D63
  m_CRC32Table(31) = &H8D080DF5
  m_CRC32Table(32) = &H3B6E20C8
  m_CRC32Table(33) = &H4C69105E
  m_CRC32Table(34) = &HD56041E4
  m_CRC32Table(35) = &HA2677172
  m_CRC32Table(36) = &H3C03E4D1
  m_CRC32Table(37) = &H4B04D447
  m_CRC32Table(38) = &HD20D85FD
  m_CRC32Table(39) = &HA50AB56B
  m_CRC32Table(40) = &H35B5A8FA
  m_CRC32Table(41) = &H42B2986C
  m_CRC32Table(42) = &HDBBBC9D6
  m_CRC32Table(43) = &HACBCF940
  m_CRC32Table(44) = &H32D86CE3
  m_CRC32Table(45) = &H45DF5C75
  m_CRC32Table(46) = &HDCD60DCF
  m_CRC32Table(47) = &HABD13D59
  m_CRC32Table(48) = &H26D930AC
  m_CRC32Table(49) = &H51DE003A
  m_CRC32Table(50) = &HC8D75180
  m_CRC32Table(51) = &HBFD06116
  m_CRC32Table(52) = &H21B4F4B5
  m_CRC32Table(53) = &H56B3C423
  m_CRC32Table(54) = &HCFBA9599
  m_CRC32Table(55) = &HB8BDA50F
  m_CRC32Table(56) = &H2802B89E
  m_CRC32Table(57) = &H5F058808
  m_CRC32Table(58) = &HC60CD9B2
  m_CRC32Table(59) = &HB10BE924
  m_CRC32Table(60) = &H2F6F7C87
  m_CRC32Table(61) = &H58684C11
  m_CRC32Table(62) = &HC1611DAB
  m_CRC32Table(63) = &HB6662D3D
  m_CRC32Table(64) = &H76DC4190
  m_CRC32Table(65) = &H1DB7106
  m_CRC32Table(66) = &H98D220BC
  m_CRC32Table(67) = &HEFD5102A
  m_CRC32Table(68) = &H71B18589
  m_CRC32Table(69) = &H6B6B51F
  m_CRC32Table(70) = &H9FBFE4A5
  m_CRC32Table(71) = &HE8B8D433
  m_CRC32Table(72) = &H7807C9A2
  m_CRC32Table(73) = &HF00F934
  m_CRC32Table(74) = &H9609A88E
  m_CRC32Table(75) = &HE10E9818
  m_CRC32Table(76) = &H7F6A0DBB
  m_CRC32Table(77) = &H86D3D2D
  m_CRC32Table(78) = &H91646C97
  m_CRC32Table(79) = &HE6635C01
  m_CRC32Table(80) = &H6B6B51F4
  m_CRC32Table(81) = &H1C6C6162
  m_CRC32Table(82) = &H856530D8
  m_CRC32Table(83) = &HF262004E
  m_CRC32Table(84) = &H6C0695ED
  m_CRC32Table(85) = &H1B01A57B
  m_CRC32Table(86) = &H8208F4C1
  m_CRC32Table(87) = &HF50FC457
  m_CRC32Table(88) = &H65B0D9C6
  m_CRC32Table(89) = &H12B7E950
  m_CRC32Table(90) = &H8BBEB8EA
  m_CRC32Table(91) = &HFCB9887C
  m_CRC32Table(92) = &H62DD1DDF
  m_CRC32Table(93) = &H15DA2D49
  m_CRC32Table(94) = &H8CD37CF3
  m_CRC32Table(95) = &HFBD44C65
  m_CRC32Table(96) = &H4DB26158
  m_CRC32Table(97) = &H3AB551CE
  m_CRC32Table(98) = &HA3BC0074
  m_CRC32Table(99) = &HD4BB30E2
  m_CRC32Table(100) = &H4ADFA541
  m_CRC32Table(101) = &H3DD895D7
  m_CRC32Table(102) = &HA4D1C46D
  m_CRC32Table(103) = &HD3D6F4FB
  m_CRC32Table(104) = &H4369E96A
  m_CRC32Table(105) = &H346ED9FC
  m_CRC32Table(106) = &HAD678846
  m_CRC32Table(107) = &HDA60B8D0
  m_CRC32Table(108) = &H44042D73
  m_CRC32Table(109) = &H33031DE5
  m_CRC32Table(110) = &HAA0A4C5F
  m_CRC32Table(111) = &HDD0D7CC9
  m_CRC32Table(112) = &H5005713C
  m_CRC32Table(113) = &H270241AA
  m_CRC32Table(114) = &HBE0B1010
  m_CRC32Table(115) = &HC90C2086
  m_CRC32Table(116) = &H5768B525
  m_CRC32Table(117) = &H206F85B3
  m_CRC32Table(118) = &HB966D409
  m_CRC32Table(119) = &HCE61E49F
  m_CRC32Table(120) = &H5EDEF90E
  m_CRC32Table(121) = &H29D9C998
  m_CRC32Table(122) = &HB0D09822
  m_CRC32Table(123) = &HC7D7A8B4
  m_CRC32Table(124) = &H59B33D17
  m_CRC32Table(125) = &H2EB40D81
  m_CRC32Table(126) = &HB7BD5C3B
  m_CRC32Table(127) = &HC0BA6CAD
  m_CRC32Table(128) = &HEDB88320
  m_CRC32Table(129) = &H9ABFB3B6
  m_CRC32Table(130) = &H3B6E20C
  m_CRC32Table(131) = &H74B1D29A
  m_CRC32Table(132) = &HEAD54739
  m_CRC32Table(133) = &H9DD277AF
  m_CRC32Table(134) = &H4DB2615
  m_CRC32Table(135) = &H73DC1683
  m_CRC32Table(136) = &HE3630B12
  m_CRC32Table(137) = &H94643B84
  m_CRC32Table(138) = &HD6D6A3E
  m_CRC32Table(139) = &H7A6A5AA8
  m_CRC32Table(140) = &HE40ECF0B
  m_CRC32Table(141) = &H9309FF9D
  m_CRC32Table(142) = &HA00AE27
  m_CRC32Table(143) = &H7D079EB1
  m_CRC32Table(144) = &HF00F9344
  m_CRC32Table(145) = &H8708A3D2
  m_CRC32Table(146) = &H1E01F268
  m_CRC32Table(147) = &H6906C2FE
  m_CRC32Table(148) = &HF762575D
  m_CRC32Table(149) = &H806567CB
  m_CRC32Table(150) = &H196C3671
  m_CRC32Table(151) = &H6E6B06E7
  m_CRC32Table(152) = &HFED41B76
  m_CRC32Table(153) = &H89D32BE0
  m_CRC32Table(154) = &H10DA7A5A
  m_CRC32Table(155) = &H67DD4ACC
  m_CRC32Table(156) = &HF9B9DF6F
  m_CRC32Table(157) = &H8EBEEFF9
  m_CRC32Table(158) = &H17B7BE43
  m_CRC32Table(159) = &H60B08ED5
  m_CRC32Table(160) = &HD6D6A3E8
  m_CRC32Table(161) = &HA1D1937E
  m_CRC32Table(162) = &H38D8C2C4
  m_CRC32Table(163) = &H4FDFF252
  m_CRC32Table(164) = &HD1BB67F1
  m_CRC32Table(165) = &HA6BC5767
  m_CRC32Table(166) = &H3FB506DD
  m_CRC32Table(167) = &H48B2364B
  m_CRC32Table(168) = &HD80D2BDA
  m_CRC32Table(169) = &HAF0A1B4C
  m_CRC32Table(170) = &H36034AF6
  m_CRC32Table(171) = &H41047A60
  m_CRC32Table(172) = &HDF60EFC3
  m_CRC32Table(173) = &HA867DF55
  m_CRC32Table(174) = &H316E8EEF
  m_CRC32Table(175) = &H4669BE79
  m_CRC32Table(176) = &HCB61B38C
  m_CRC32Table(177) = &HBC66831A
  m_CRC32Table(178) = &H256FD2A0
  m_CRC32Table(179) = &H5268E236
  m_CRC32Table(180) = &HCC0C7795
  m_CRC32Table(181) = &HBB0B4703
  m_CRC32Table(182) = &H220216B9
  m_CRC32Table(183) = &H5505262F
  m_CRC32Table(184) = &HC5BA3BBE
  m_CRC32Table(185) = &HB2BD0B28
  m_CRC32Table(186) = &H2BB45A92
  m_CRC32Table(187) = &H5CB36A04
  m_CRC32Table(188) = &HC2D7FFA7
  m_CRC32Table(189) = &HB5D0CF31
  m_CRC32Table(190) = &H2CD99E8B
  m_CRC32Table(191) = &H5BDEAE1D
  m_CRC32Table(192) = &H9B64C2B0
  m_CRC32Table(193) = &HEC63F226
  m_CRC32Table(194) = &H756AA39C
  m_CRC32Table(195) = &H26D930A
  m_CRC32Table(196) = &H9C0906A9
  m_CRC32Table(197) = &HEB0E363F
  m_CRC32Table(198) = &H72076785
  m_CRC32Table(199) = &H5005713
  m_CRC32Table(200) = &H95BF4A82
  m_CRC32Table(201) = &HE2B87A14
  m_CRC32Table(202) = &H7BB12BAE
  m_CRC32Table(203) = &HCB61B38
  m_CRC32Table(204) = &H92D28E9B
  m_CRC32Table(205) = &HE5D5BE0D
  m_CRC32Table(206) = &H7CDCEFB7
  m_CRC32Table(207) = &HBDBDF21
  m_CRC32Table(208) = &H86D3D2D4
  m_CRC32Table(209) = &HF1D4E242
  m_CRC32Table(210) = &H68DDB3F8
  m_CRC32Table(211) = &H1FDA836E
  m_CRC32Table(212) = &H81BE16CD
  m_CRC32Table(213) = &HF6B9265B
  m_CRC32Table(214) = &H6FB077E1
  m_CRC32Table(215) = &H18B74777
  m_CRC32Table(216) = &H88085AE6
  m_CRC32Table(217) = &HFF0F6A70
  m_CRC32Table(218) = &H66063BCA
  m_CRC32Table(219) = &H11010B5C
  m_CRC32Table(220) = &H8F659EFF
  m_CRC32Table(221) = &HF862AE69
  m_CRC32Table(222) = &H616BFFD3
  m_CRC32Table(223) = &H166CCF45
  m_CRC32Table(224) = &HA00AE278
  m_CRC32Table(225) = &HD70DD2EE
  m_CRC32Table(226) = &H4E048354
  m_CRC32Table(227) = &H3903B3C2
  m_CRC32Table(228) = &HA7672661
  m_CRC32Table(229) = &HD06016F7
  m_CRC32Table(230) = &H4969474D
  m_CRC32Table(231) = &H3E6E77DB
  m_CRC32Table(232) = &HAED16A4A
  m_CRC32Table(233) = &HD9D65ADC
  m_CRC32Table(234) = &H40DF0B66
  m_CRC32Table(235) = &H37D83BF0
  m_CRC32Table(236) = &HA9BCAE53
  m_CRC32Table(237) = &HDEBB9EC5
  m_CRC32Table(238) = &H47B2CF7F
  m_CRC32Table(239) = &H30B5FFE9
  m_CRC32Table(240) = &HBDBDF21C
  m_CRC32Table(241) = &HCABAC28A
  m_CRC32Table(242) = &H53B39330
  m_CRC32Table(243) = &H24B4A3A6
  m_CRC32Table(244) = &HBAD03605
  m_CRC32Table(245) = &HCDD70693
  m_CRC32Table(246) = &H54DE5729
  m_CRC32Table(247) = &H23D967BF
  m_CRC32Table(248) = &HB3667A2E
  m_CRC32Table(249) = &HC4614AB8
  m_CRC32Table(250) = &H5D681B02
  m_CRC32Table(251) = &H2A6F2B94
  m_CRC32Table(252) = &HB40BBE37
  m_CRC32Table(253) = &HC30C8EA1
  m_CRC32Table(254) = &H5A05DF1B
  m_CRC32Table(255) = &H2D02EF8D
  sASM = "5589E557565053518B45088B008B750C8B7D108B4D1431DB8A1E30C3C1E80833049F464975F28B4D088901595B585E5F89EC5DC21000"
  ReDim m_CRC32Asm(0 To Len(sASM) \ 2 - 1)
  For i = 1 To Len(sASM) Step 2
    m_CRC32Asm(i \ 2) = Val("&H" & Mid$(sASM, i, 2))
  Next
  m_CRC32Init = True
End Sub
Private Sub Class_Initialize()
If (Not m_CRC32Init) Then Call InitializeCRC32
m_CRC32 = &HFFFFFFFF
End Sub
Public Sub Huffman_EncodeFile(SourceFile As String, DestFile As String)
  Dim ByteArray() As Byte
  Dim Filenr As Integer
  If (Not FileExist(SourceFile)) Then
    Err.Raise vbObjectError, "clsHuffman.Huffman_EncodeFile()", "Source file does not exist"
  End If
  Filenr = FreeFile
  Open SourceFile For Binary As #Filenr
  ReDim ByteArray(0 To LOF(Filenr) - 1)
  Get #Filenr, , ByteArray()
  Close #Filenr
  Call EncodeByte(ByteArray(), UBound(ByteArray) + 1)
  If (FileExist(DestFile)) Then Kill DestFile
  Open DestFile For Binary As #Filenr
  Put #Filenr, , ByteArray()
  Close #Filenr
End Sub
Public Sub Huffman_Decodefile(SourceFile As String, DestFile As String)
  Dim ByteArray() As Byte
  Dim Filenr As Integer
  If (Not FileExist(SourceFile)) Then
    Err.Raise vbObjectError, "clsHuffman.Huffman_Decodefile()", "Source file does not exist"
  End If
  Filenr = FreeFile
  Open SourceFile For Binary As #Filenr
  ReDim ByteArray(0 To LOF(Filenr) - 1)
  Get #Filenr, , ByteArray()
  Close #Filenr
  Call DecodeByte(ByteArray(), UBound(ByteArray) + 1)
  If (FileExist(DestFile)) Then Kill DestFile
  Open DestFile For Binary As #Filenr
  Put #Filenr, , ByteArray()
  Close #Filenr
End Sub
Private Sub CreateTree(Nodes() As HUFFMANTREE, NodesCount As Long, Char As Long, Bytes As ByteArray)
  Dim a As Integer
  Dim NodeIndex As Long
  NodeIndex = 0
  For a = 0 To (Bytes.Count - 1)
    If (Bytes.Data(a) = 0) Then
      If (Nodes(NodeIndex).LeftNode = -1) Then
        Nodes(NodeIndex).LeftNode = NodesCount
        Nodes(NodesCount).ParentNode = NodeIndex
        Nodes(NodesCount).LeftNode = -1
        Nodes(NodesCount).RightNode = -1
        Nodes(NodesCount).Value = -1
        NodesCount = NodesCount + 1
      End If
      NodeIndex = Nodes(NodeIndex).LeftNode
    ElseIf (Bytes.Data(a) = 1) Then
      If (Nodes(NodeIndex).RightNode = -1) Then
        Nodes(NodeIndex).RightNode = NodesCount
        Nodes(NodesCount).ParentNode = NodeIndex
        Nodes(NodesCount).LeftNode = -1
        Nodes(NodesCount).RightNode = -1
        Nodes(NodesCount).Value = -1
        NodesCount = NodesCount + 1
      End If
      NodeIndex = Nodes(NodeIndex).RightNode
    Else
      Stop
    End If
  Next
Nodes(NodeIndex).Value = Char
End Sub
Private Sub EncodeByte(ByteArray() As Byte, ByteLen As Long)
  Dim i As Long
  Dim j As Long
  Dim Char As Byte
  Dim BitPos As Byte
  Dim lNode1 As Long
  Dim lNode2 As Long
  Dim lNodes As Long
  Dim lLength As Long
  Dim Count As Integer
  Dim lWeight1 As Long
  Dim lWeight2 As Long
  Dim Result() As Byte
  Dim ByteValue As Byte
  Dim ResultLen As Long
  Dim Bytes As ByteArray
  Dim NodesCount As Integer
  Dim NewProgress As Integer
  Dim CurrProgress As Integer
  Dim BitValue(0 To 7) As Byte
  Dim CharCount(0 To 255) As Long
  Dim Nodes(0 To 511) As HUFFMANTREE
  Dim CharValue(0 To 255) As ByteArray
  If (ByteLen = 0) Then
    ReDim Preserve ByteArray(0 To ByteLen + 3)
    If (ByteLen > 0) Then
      Call CopyMem(ByteArray(4), ByteArray(0), ByteLen)
    End If
    ByteArray(0) = 72 '"H"
    ByteArray(1) = 69 '"E"
    ByteArray(2) = 48 '"0"
    ByteArray(3) = 13 'vbCr
    Exit Sub
  End If
  ReDim Result(0 To 522)
  Result(0) = 72 'H
  Result(1) = 69 'E
  Result(2) = 51 '3
  Result(3) = 13 'vbCr
  ResultLen = 4
  For i = 0 To (ByteLen - 1)
    CharCount(ByteArray(i)) = CharCount(ByteArray(i)) + 1
    If (i Mod 1000 = 0) Then
      NewProgress = i / ByteLen * PROGRESS_CALCFREQUENCY
      If (NewProgress <> CurrProgress) Then
        CurrProgress = NewProgress
        RaiseEvent Progress(CurrProgress)
      End If
    End If
  Next
  For i = 0 To 255
    If (CharCount(i) > 0) Then
      With Nodes(NodesCount)
        .Weight = CharCount(i)
        .Value = i
        .LeftNode = -1
        .RightNode = -1
        .ParentNode = -1
      End With
      NodesCount = NodesCount + 1
    End If
  Next
  For lNodes = NodesCount To 2 Step -1
    lNode1 = -1: lNode2 = -1
    For i = 0 To (NodesCount - 1)
      If (Nodes(i).ParentNode = -1) Then
        If (lNode1 = -1) Then
          lWeight1 = Nodes(i).Weight
          lNode1 = i
        ElseIf (lNode2 = -1) Then
          lWeight2 = Nodes(i).Weight
          lNode2 = i
        ElseIf (Nodes(i).Weight < lWeight1) Then
          If (Nodes(i).Weight < lWeight2) Then
            If (lWeight1 < lWeight2) Then
              lWeight2 = Nodes(i).Weight
              lNode2 = i
            Else
              lWeight1 = Nodes(i).Weight
              lNode1 = i
            End If
          Else
            lWeight1 = Nodes(i).Weight
            lNode1 = i
          End If
        ElseIf (Nodes(i).Weight < lWeight2) Then
          lWeight2 = Nodes(i).Weight
          lNode2 = i
        End If
      End If
    Next
    With Nodes(NodesCount)
      .Weight = lWeight1 + lWeight2
      .LeftNode = lNode1
      .RightNode = lNode2
      .ParentNode = -1
      .Value = -1
    End With
    Nodes(lNode1).ParentNode = NodesCount
    Nodes(lNode2).ParentNode = NodesCount
    NodesCount = NodesCount + 1
  Next
ReDim Bytes.Data(0 To 255)
  Call CreateBitSequences(Nodes(), NodesCount - 1, Bytes, CharValue)
  For i = 0 To 255
    If (CharCount(i) > 0) Then
      lLength = lLength + CharValue(i).Count * CharCount(i)
    End If
  Next
  lLength = IIf(lLength Mod 8 = 0, lLength \ 8, lLength \ 8 + 1)
  If ((lLength = 0) Or (lLength > ByteLen)) Then
    Call CopyMem(ByteArray(4), ByteArray(0), ByteLen)
    ByteArray(0) = 72
    ByteArray(1) = 69
    ByteArray(2) = 48
    ByteArray(3) = 13
    Exit Sub
  End If
  Char = 0
  For i = 0 To (ByteLen - 1)
    Char = Char Xor ByteArray(i)
    If (i Mod 10000 = 0) Then
      NewProgress = i / ByteLen * PROGRESS_CALCCRC + PROGRESS_CALCFREQUENCY
      If (NewProgress <> CurrProgress) Then
        CurrProgress = NewProgress
        RaiseEvent Progress(CurrProgress)
      End If
    End If
  Next
  Result(ResultLen) = Char
  ResultLen = ResultLen + 1
  Call CopyMem(Result(ResultLen), ByteLen, 4)
  ResultLen = ResultLen + 4
  For i = 0 To 7
    BitValue(i) = 2 ^ i
  Next
  Count = 0
  For i = 0 To 255
    If (CharValue(i).Count > 0) Then
      Count = Count + 1
    End If
  Next
  Call CopyMem(Result(ResultLen), Count, 2)
  ResultLen = ResultLen + 2
  Count = 0
  For i = 0 To 255
    If (CharValue(i).Count > 0) Then
      Result(ResultLen) = i
      ResultLen = ResultLen + 1
      Result(ResultLen) = CharValue(i).Count
      ResultLen = ResultLen + 1
      Count = Count + 16 + CharValue(i).Count
    End If
  Next
  ReDim Preserve Result(0 To ResultLen + Count \ 8)
  BitPos = 0
  ByteValue = 0
  For i = 0 To 255
    With CharValue(i)
      If (.Count > 0) Then
        For j = 0 To (.Count - 1)
          If (.Data(j)) Then ByteValue = ByteValue + BitValue(BitPos)
          BitPos = BitPos + 1
          If (BitPos = 8) Then
            Result(ResultLen) = ByteValue
            ResultLen = ResultLen + 1
            ByteValue = 0
            BitPos = 0
          End If
        Next
      End If
    End With
  Next
  If (BitPos > 0) Then
    Result(ResultLen) = ByteValue
    ResultLen = ResultLen + 1
  End If
  ReDim Preserve Result(0 To ResultLen - 1 + lLength)
  Char = 0
  BitPos = 0
  For i = 0 To (ByteLen - 1)
    With CharValue(ByteArray(i))
      For j = 0 To (.Count - 1)
        If (.Data(j) = 1) Then Char = Char + BitValue(BitPos)
        BitPos = BitPos + 1
        If (BitPos = 8) Then
          Result(ResultLen) = Char
          ResultLen = ResultLen + 1
          BitPos = 0
          Char = 0
        End If
      Next
    End With
    If (i Mod 10000 = 0) Then
      NewProgress = i / ByteLen * PROGRESS_ENCODING + PROGRESS_CALCCRC + PROGRESS_CALCFREQUENCY
      If (NewProgress <> CurrProgress) Then
        CurrProgress = NewProgress
        RaiseEvent Progress(CurrProgress)
      End If
    End If
  Next
  If (BitPos > 0) Then
    Result(ResultLen) = Char
    ResultLen = ResultLen + 1
  End If
  ReDim ByteArray(0 To ResultLen - 1)
  Call CopyMem(ByteArray(0), Result(0), ResultLen)
  If (CurrProgress <> 100) Then
    RaiseEvent Progress(100)
  End If
End Sub

Private Function DecodeByte(ByteArray() As Byte, ByteLen As Long)
  Dim i As Long
  Dim j As Long
  Dim pos As Long
  Dim Char As Byte
  Dim CurrPos As Long
  Dim Count As Integer
  Dim CheckSum As Byte
  Dim Result() As Byte
  Dim BitPos As Integer
  Dim NodeIndex As Long
  Dim ByteValue As Byte
  Dim ResultLen As Long
  Dim NodesCount As Long
  Dim lResultLen As Long
  Dim NewProgress As Integer
  Dim CurrProgress As Integer
  Dim BitValue(0 To 7) As Byte
  Dim Nodes(0 To 511) As HUFFMANTREE
  Dim CharValue(0 To 255) As ByteArray
  If (ByteArray(0) <> 72) Or (ByteArray(1) <> 69) Or (ByteArray(3) <> 13) Then
  ElseIf (ByteArray(2) = 48) Then
    Call CopyMem(ByteArray(0), ByteArray(4), ByteLen - 4)
    ReDim Preserve ByteArray(0 To ByteLen - 5)
    Exit Function
  ElseIf (ByteArray(2) <> 51) Then
    Err.Raise vbObjectError, "HuffmanDecode()", "The data either was not compressed with HE3 or is corrupt (identification string not found)"
    Exit Function
  End If
  CurrPos = 5
  CheckSum = ByteArray(CurrPos - 1)
  CurrPos = CurrPos + 1
  Call CopyMem(ResultLen, ByteArray(CurrPos - 1), 4)
  CurrPos = CurrPos + 4
  lResultLen = ResultLen
  If (ResultLen = 0) Then Exit Function
  ReDim Result(0 To ResultLen - 1)
  Call CopyMem(Count, ByteArray(CurrPos - 1), 2)
  CurrPos = CurrPos + 2
  For i = 1 To Count
    With CharValue(ByteArray(CurrPos - 1))
      CurrPos = CurrPos + 1
      .Count = ByteArray(CurrPos - 1)
      CurrPos = CurrPos + 1
      ReDim .Data(0 To .Count - 1)
    End With
  Next
  For i = 0 To 7
    BitValue(i) = 2 ^ i
  Next
  ByteValue = ByteArray(CurrPos - 1)
  CurrPos = CurrPos + 1
  BitPos = 0
  For i = 0 To 255
    With CharValue(i)
      If (.Count > 0) Then
        For j = 0 To (.Count - 1)
          If (ByteValue And BitValue(BitPos)) Then .Data(j) = 1
          BitPos = BitPos + 1
          If (BitPos = 8) Then
            ByteValue = ByteArray(CurrPos - 1)
            CurrPos = CurrPos + 1
            BitPos = 0
          End If
        Next
      End If
    End With
  Next
  If (BitPos = 0) Then CurrPos = CurrPos - 1
  NodesCount = 1
  Nodes(0).LeftNode = -1
  Nodes(0).RightNode = -1
  Nodes(0).ParentNode = -1
  Nodes(0).Value = -1
  For i = 0 To 255
    Call CreateTree(Nodes(), NodesCount, i, CharValue(i))
  Next
  ResultLen = 0
  For CurrPos = CurrPos To ByteLen
    ByteValue = ByteArray(CurrPos - 1)
    For BitPos = 0 To 7
      If (ByteValue And BitValue(BitPos)) Then
        NodeIndex = Nodes(NodeIndex).RightNode
      Else
        NodeIndex = Nodes(NodeIndex).LeftNode
      End If
      If (Nodes(NodeIndex).Value > -1) Then
        Result(ResultLen) = Nodes(NodeIndex).Value
        ResultLen = ResultLen + 1
        If (ResultLen = lResultLen) Then GoTo DecodeFinished
        NodeIndex = 0
      End If
    Next
    If (CurrPos Mod 10000 = 0) Then
      NewProgress = CurrPos / ByteLen * PROGRESS_DECODING
      If (NewProgress <> CurrProgress) Then
        CurrProgress = NewProgress
        RaiseEvent Progress(CurrProgress)
      End If
    End If
  Next
DecodeFinished:
Char = 0
  For i = 0 To (ResultLen - 1)
    Char = Char Xor Result(i)
    If (i Mod 10000 = 0) Then
      NewProgress = i / ResultLen * PROGRESS_CHECKCRC + PROGRESS_DECODING
      If (NewProgress <> CurrProgress) Then
        CurrProgress = NewProgress
        RaiseEvent Progress(CurrProgress)
      End If
    End If
  Next
  If (Char <> CheckSum) Then
    Err.Raise vbObjectError, "clsHuffman.Decode()", "The data might be corrupted (checksum did not match expected value)"
  End If
  ReDim ByteArray(0 To ResultLen - 1)
  Call CopyMem(ByteArray(0), Result(0), ResultLen)
  If (CurrProgress <> 100) Then
    RaiseEvent Progress(100)
  End If
End Function
Private Sub CreateBitSequences(Nodes() As HUFFMANTREE, ByVal NodeIndex As Integer, Bytes As ByteArray, CharValue() As ByteArray)
  Dim NewBytes As ByteArray
  If (Nodes(NodeIndex).Value > -1) Then
    CharValue(Nodes(NodeIndex).Value) = Bytes
    Exit Sub
  End If
  If (Nodes(NodeIndex).LeftNode > -1) Then
    NewBytes = Bytes
    NewBytes.Data(NewBytes.Count) = 0
    NewBytes.Count = NewBytes.Count + 1
    Call CreateBitSequences(Nodes(), Nodes(NodeIndex).LeftNode, NewBytes, CharValue)
  End If
  If (Nodes(NodeIndex).RightNode > -1) Then
    NewBytes = Bytes
    NewBytes.Data(NewBytes.Count) = 1
    NewBytes.Count = NewBytes.Count + 1
    Call CreateBitSequences(Nodes(), Nodes(NodeIndex).RightNode, NewBytes, CharValue)
  End If
End Sub
Public Function FileExist(Filename As String) As Boolean
  On Error GoTo FileDoesNotExist
  Call FileLen(Filename)
  FileExist = True
  Exit Function
FileDoesNotExist:
  FileExist = False
End Function


