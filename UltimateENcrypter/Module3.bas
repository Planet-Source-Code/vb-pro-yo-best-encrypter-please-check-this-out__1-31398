Attribute VB_Name = "Module3"
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private ByteArray() As Byte
Private hiByte As Long
Private hiBound As Long
Dim intBin(7) As Integer

 Sub Class_Initialize()
intBin(0) = 1
intBin(1) = 2
intBin(2) = 4
intBin(3) = 8
intBin(4) = 16
intBin(5) = 32
intBin(6) = 64
intBin(7) = 128
End Sub
 Function Ascii2Binary(Text As String) As String
On Error GoTo errorhandler
GoSub begin

errorhandler:
Reset
Exit Function

begin:
    Dim iAsc As Integer, hLoop As Integer, tInt As Integer, mLoop As Integer
    Reset
    For mLoop = 1 To Len(Text)
        iAsc = Asc(Mid$(Text, mLoop, 1))
        For hLoop = 7 To 0 Step -1
            If intBin(hLoop) <= iAsc - tInt Then
                Append "1"
                tInt = tInt + intBin(hLoop)
            Else
                Append "0"
            End If
        Next
        tInt = 0
    Next
    Ascii2Binary = Trim(GData)
    Reset
End Function
Private Sub Reset()
    hiByte = 0
    hiBound = 1024
    ReDim ByteArray(hiBound)
End Sub
 Sub Append(ByRef StringData As String, Optional Length As Long)
    Dim DataLength As Long
    If Length > 0 Then DataLength = Length Else DataLength = Len(StringData)
    If DataLength + hiByte > hiBound Then
        hiBound = hiBound + 1024
        ReDim Preserve ByteArray(hiBound)
    End If
    CopyMem ByVal VarPtr(ByteArray(hiByte)), ByVal StringData, DataLength
    hiByte = hiByte + DataLength
End Sub
Sub translatexor(datastring)
    Dim temp As String
    Dim i As Integer
    Dim location As Integer
    'DataString$ is the string you want to e
    '     ncode/decode.
    temp$ = ""
    'this temporarily holds the encrypted st
    '     ring


    For i% = 1 To Len(datastring$)
        location% = (i% Mod Len(Code$)) + 1
        'this little operation gives you the nex
        '     t byte location in the
        'CODE string, looping back to the beginn
        '     ing when it reaches the
        'end
        temp$ = temp$ + Chr$(Asc(Mid$(datastring$, i%, 1)) Xor Asc(Mid$(Code$, location%, 1)))
        'perform the XOR operation
    Next i%


    Form1.textbox1.Text = temp$
    End Sub
 Function HexDecode(Text As String) As String
    Dim iCount As Double
    Reset
    For iCount = 1 To Len(Text) Step 2
        Append Chr$(Val("&H" & Mid$(Text, iCount, 2)))
    Next
    HexDecode = GData
    Reset
    DoEvents
'    Form1.text1.Text = HexDecode
End Function
 Function HexEncode(Text As String) As String
 
    Dim iCount As Double, sTemp As String
    Reset
    For iCount = 1 To Len(Text)
        sTemp = Hex$(Asc(Mid$(Text, iCount, 1)))
        If Len(sTemp) < 2 Then sTemp = "0" & sTemp
        Append sTemp
    Next
    HexEncode = GData
    Reset
    Form1.text2.Text = HexEncode
End Function
 Property Get GData() As String
    Dim StringData As String
    StringData = Space(hiByte)
    CopyMem ByVal StringData, ByVal VarPtr(ByteArray(0)), hiByte
    GData = StringData
End Property
 Function Binary2Ascii(BinaryText As String) As String
On Error GoTo errorhandler
GoSub begin

errorhandler:
Exit Function

begin:
    Dim iLoop As Integer, t As Integer, mLoop As Integer, s As String, n As Integer
    For mLoop = 1 To Len(BinaryText) / 8
        For iLoop = ((mLoop - 1) * 8) + 1 To ((mLoop - 1) * 8) + 8
            n = n + 1
            t = t + (Val(Mid(BinaryText, iLoop, 1)) * intBin(8 - n))
        Next
        s = s & Chr(t)
        t = 0: n = 0
    Next
    Binary2Ascii = s
End Function
 Function ReverseString(Text As String) As String
On Error GoTo errorhandler
GoSub begin

errorhandler:
Reset
Exit Function

begin:
    Dim iLoop As Integer
    Reset
    For iLoop = Len(Text) To 1 Step -1
        Append Mid$(Text, iLoop, 1)
    Next
    ReverseString = GData
    Reset
End Function
Function BTMEncrypt(Text)
    For god = 1 To Len(Text)
            Current$ = Asc(Mid(Text, god, 1)) - god
        Process$ = Process$ & Chr(Current$)
    Next god
    BTMEncrypt = Process$
    Form1.text2.Text = BTMEncrypt
End Function
Function BTMdEcrypt(Text)
    For god = 1 To Len(Text)
            Current$ = Asc(Mid(Text, god, 1)) + god
        Process$ = Process$ & Chr(Current$)
    Next god
    BTMEncrypt = Process$
    Form1.text1.Text = BTMEncrypt
End Function
Function encryptit(ToEncode As String) As String
    Dim a As String, b As String, x As Long, s() As Byte
    Dim c As String, d As String
    'If The Value Is Uneven Simply Add A Space
    If Not Len(ToEncode) Mod 2 = 0 Then ToEncode = ToEncode & " "
    'Create An Array To Hold The Encoded String
    ReDim s(Len(ToEncode) - 1)
    For x = 1 To Len(ToEncode)
        'a Holds The Left Hex Value Of The Current Char
        a = Left(Right("0" & Hex(Asc(Mid(ToEncode, x, 1))), 2), 1)
        'b Holds The Right Hex Value Of The Current Char
        b = Right(Right("0" & Hex(Asc(Mid(ToEncode, x, 1))), 2), 1)
        'c Holds The Left Hex Value Of The Last Char - Current Char
        c = Left(Right("0" & Hex(Asc(Mid(ToEncode, Len(ToEncode) - x + 1, 1))), 2), 1)
        'd Holds The Right Hex Value Of The Last Char - Current Char
        d = Right(Right("0" & Hex(Asc(Mid(ToEncode, Len(ToEncode) - x + 1, 1))), 2), 1)
        'We Combine Half Of The First Char And Half Of The Last Char Into A New Char
        s(x - 1) = Val("&H" & a & c)
        'We Do It Again For The Last Char
        s(Len(ToEncode) - x) = Val("&H" & b & d)
    Next
    'Make The Byte Array Into A String
    Code = StrConv(s, vbUnicode)
    Form1.Text.Text = Code
End Function
Function decryptit(ToEncode As String) As String
    Dim a As String, b As String, x As Long, s() As Byte
    Dim c As String, d As String
    'If The Value Is Uneven Simply Add A Space
    If Not Len(ToEncode) Mod 2 = 0 Then ToEncode = ToEncode & " "
    'Create An Array To Hold The Encoded String
    ReDim s(Len(ToEncode) - 1)
    For x = 1 To Len(ToEncode)
        'a Holds The Left Hex Value Of The Current Char
        a = Left(Right("0" & Hex(Asc(Mid(ToEncode, x, 1))), 2), 1)
        'b Holds The Right Hex Value Of The Current Char
        b = Right(Right("0" & Hex(Asc(Mid(ToEncode, x, 1))), 2), 1)
        'c Holds The Left Hex Value Of The Last Char - Current Char
        c = Left(Right("0" & Hex(Asc(Mid(ToEncode, Len(ToEncode) - x + 1, 1))), 2), 1)
        'd Holds The Right Hex Value Of The Last Char - Current Char
        d = Right(Right("0" & Hex(Asc(Mid(ToEncode, Len(ToEncode) - x + 1, 1))), 2), 1)
        'We Combine Half Of The First Char And Half Of The Last Char Into A New Char
        s(x - 1) = Val("&H" & a & c)
        'We Do It Again For The Last Char
        s(Len(ToEncode) - x) = Val("&H" & b & d)
    Next
    'Make The Byte Array Into A String
    Code = StrConv(s, vbUnicode)
    Form1.text1.Text = Code
End Function
 Function CaesarShiftencode(Text As String) As String
'On Error GoTo errorhandler

GoSub begin


'errorhandler:
'Reset
'Exit Function
Difference = 4
begin:
    While Difference > 26
        Difference = Difference - 26
    Wend
    Reset
    Dim iLoop As Integer, tAsc As String
    For iLoop = 1 To Len(Text)
        tAsc = Mid$(Text, iLoop, 1)
        If tAsc = " " Then Append " "
        If tAsc = LCase$(tAsc) And tAsc <> " " Then
            If Asc(tAsc) + Difference > 122 Then Append Chr$(97 + ((Asc(tAsc) + Difference) - 122)) Else Append Chr$(Asc(tAsc) + Difference)
        ElseIf tAsc = UCase$(tAsc) And tAsc <> " " Then
            If Asc(tAsc) + Difference > 90 Then Append Chr$(65 + ((Asc(tAsc) + Difference) - 90)) Else Append Chr$(Asc(tAsc) + Difference)
        End If
    Next
    CaesarShift = GData
    Form1.text2.Text = CaesarShift
    Reset
End Function
 Function CaesarShiftdecode(Text As String) As String
'On Error GoTo errorhandler
GoSub begin

'errorhandler:
'Reset
'Exit Function
Difference = 22
begin:
    While Difference > 26
        Difference = Difference - 26
    Wend
    Reset
    Dim iLoop As Integer, tAsc As String
    For iLoop = 1 To Len(Text)
        tAsc = Mid$(Text, iLoop, 1)
        If tAsc = " " Then Append " "
        If tAsc = LCase$(tAsc) And tAsc <> " " Then
            If Asc(tAsc) + Difference > 122 Then Append Chr$(97 + ((Asc(tAsc) + Difference) - 122)) Else Append Chr$(Asc(tAsc) + Difference)
        ElseIf tAsc = UCase$(tAsc) And tAsc <> " " Then
            If Asc(tAsc) + Difference > 90 Then Append Chr$(65 + ((Asc(tAsc) + Difference) - 90)) Else Append Chr$(Asc(tAsc) + Difference)
        End If
    Next
    CaesarShift = GData
    Form1.text1.Text = CaesarShift
    Reset
End Function
 Function Ascii2Transposition(Text As String) As String
On Error GoTo errorhandler
GoSub begin

errorhandler:
Reset
Exit Function

begin:
    Dim iLoop As Integer, dInt As Integer
    Reset
    dInt = Format$(Len(Text) / 2, 0)
    For iLoop = 1 To Len(Text) / 2
        Append Mid$(Text, iLoop, 1)
        Append Mid$(Text, iLoop + dInt, 1)
    Next
    If Len(Text) / 2 < dInt Then Append Mid$(Text, dInt, 1)
    Ascii2Transposition = GData
    Reset
End Function
Function Transposition2Ascii(Text As String) As String
On Error GoTo errorhandler
GoSub begin

errorhandler:
Reset
Exit Function

begin:
    Dim iLoop As Integer, dInt As Integer
    dInt = Len(Text) / 2
    Reset
    For iLoop = 1 To Len(Text) Step 2
        Append Mid$(Text, iLoop, 1)
    Next
    For iLoop = 2 To Len(Text) Step 2
        Append Mid$(Text, iLoop, 1)
    Next iLoop
    Transposition2Ascii = GData
    Reset
End Function
Function Ascii2PigLatin(Text As String) As String
On Error GoTo errorhandler
GoSub begin

errorhandler:
Exit Function

begin:
    Dim iInt As Integer, sString As String, n() As String, zInt As Integer
    sString = Text
    Do
        DoEvents
        iInt = InStr(sString, " ")
        If iInt = 0 Then Exit Do
        sString = Mid$(sString, iInt + 1)
        zInt = zInt + 1
    Loop
    n = Split(Text, " ")
    sString = ""
    For iInt = 0 To zInt
        sString = sString & Right$(n(iInt), Len(n(iInt)) - 1) & Left$(n(iInt), 1) & "ay "
    Next
    sString = Trim(sString)
    Ascii2PigLatin = sString
End Function

 Function PigLatin2Ascii(Text As String) As String
On Error GoTo errorhandler
GoSub begin

errorhandler:
Reset
Exit Function

begin:
    Dim iInt As Integer, sString As String, n() As String, zInt As Integer
    sString = Text
    Do
        DoEvents
        iInt = InStr(sString, " ")
        If iInt = 0 Then Exit Do
        sString = Mid$(sString, iInt + 1)
        zInt = zInt + 1
    Loop
    n = Split(Text, " ")
    sString = ""
    For iInt = 0 To zInt
        n(iInt) = Left$(n(iInt), Len(n(iInt)) - 2)
        sString = sString & Right$(n(iInt), 1) & Left$(n(iInt), Len(n(iInt)) - 1) & " "
    Next
    sString = Trim(sString)
    PigLatin2Ascii = sString
End Function

Function Ascii2Random(CustomValue As String, EncryptCode As String) As String
On Error GoTo errorhandler
GoSub begin

errorhandler:
Reset
Exit Function

begin:
    Dim a As String, i As Long, b As String, d As String, Q%, c As String
    Q = Val(Trigger(EncryptCode))
    a = CustomValue
    Reset
    For i = 1 To Len(a)
        DoEvents
        b = Mid$(a, i, 1)
        c = Asc(b) + Q
        If c < 0 Then c = c - c - c
        If Len(c) = 1 Then c = "00" & c
        If Len(c) = 2 Then c = "0" & c
        Append c
    Next i
    d = GData
    Reset
    Dim f As Integer, e As String, g As Integer
    For i = 1 To Len(d)
        DoEvents
        f = Mid$(d, i, 1)
        Randomize
        g = Rnd * 19
        Select Case f
            Case 0: Append Chr(g + 55)
            Case 1: Append Chr(g + 75)
            Case 2: Append Chr(g + 95)
            Case 3: Append Chr(g + 115)
            Case 4: Append Chr(g + 135)
            Case 5: Append Chr(g + 155)
            Case 6: Append Chr(g + 175)
            Case 7: Append Chr(g + 195)
            Case 8: Append Chr(g + 215)
            Case 9: Append Chr(g + 235)
        End Select
    Next i
    Ascii2Random = GData
    Reset
End Function

Function Random2Ascii(ByVal CustomValue As String, ByVal EncryptCode As String) As String
On Error GoTo errorhandler
GoSub begin

errorhandler:
Reset
Exit Function

begin:
    Dim x As String, y As String, z As String, Q As Integer, a As String, i As Long, b As String, c As String
    x = Left$(EncryptCode, 1)
    y = Right$(EncryptCode, 1)
    z = Mid$(EncryptCode, Format$(Len(EncryptCode) / 2, "#"), 1)
    Q = Format$((Asc(x) + Asc(y) + Asc(z)) / 6, "##")
    a = CustomValue
    For i = 1 To Len(a)
        DoEvents
        b = Val(Asc(Mid$(a, i, 1)))
        If b >= 55 And b < 75 Then c = c & 0
        If b >= 75 And b < 95 Then c = c & 1
        If b >= 95 And b < 115 Then c = c & 2
        If b >= 115 And b < 135 Then c = c & 3
        If b >= 135 And b < 155 Then c = c & 4
        If b >= 155 And b < 175 Then c = c & 5
        If b >= 175 And b < 195 Then c = c & 6
        If b >= 195 And b < 215 Then c = c & 7
        If b >= 215 And b < 235 Then c = c & 8
        If b >= 235 And b < 255 Then c = c & 9
        If b < 55 Or b >= 255 Then
            Random2Ascii = ""
            Exit Function
        End If
    Next i
    Reset
    Dim d As Integer, f As Integer
    For i = 1 To Len(c) Step 3
        DoEvents
        d = Mid$(c, i, 3)
        f = d - Q
        If f < 0 Then f = f - f - f
        If f > 255 Or f < 0 Then
            Random2Ascii = ""
            Exit Function
        End If
        Append Chr$(f)
    Next i
    Random2Ascii = GData
    Reset
End Function

 Function Trigger(CodeWord As String) As String
On Error GoTo errorhandler
GoSub begin

errorhandler:
Exit Function

begin:
    If CodeWord = "" Then Trigger = "": Exit Function
    Dim x$, y$, z$, Q%
    x = Left$(CodeWord, 1)
    y = Right$(CodeWord, 1)
    z = Mid$(CodeWord, Format$(Len(CodeWord) / 2, "#"), 1)
    Q = Format$((Asc(x) + Asc(y) + Asc(z)) / 6, "##")
    Trigger = Q
End Function
