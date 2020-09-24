Attribute VB_Name = "Compress"
'* LZSS Visual Basic port from ANSI C
'* Translation Copyright 1999 Patterson Programming

Option Explicit
DefInt A-Z

Const N = 4096
Const F = 18
Const NIL = N
Const THRESHOLD = 2
Dim Infinity&

Dim textsize As Long
Dim codesize As Long
Dim Allocated As Long
Dim textbuf(N + F - 1) As Byte
Dim matchposition As Integer
Dim matchlength As Byte
Dim lson(N + 1) As Integer
Dim rson(N + 257) As Integer
Dim dad(N + 1) As Integer
Dim InPointer&, OutPointer&

Private Sub InitTree()
    Dim i As Integer
    For i = N + 1 To (N + 256): rson(i) = NIL: Next
    For i = 0 To N - 1: dad(i) = NIL: Next
End Sub

Private Sub InsertNode(ByVal r As Integer)

    Dim i As Integer, p As Integer, cmp As Integer
    Dim key As Integer
    Dim x&, x1 As Integer, x2 As Integer

    cmp = 1
    key = r
    p = N + 1 + textbuf(r)
    rson(r) = NIL: lson(r) = NIL
    matchlength = 0

    Infinity& = 2147483647
    For x& = 0 To Infinity&
        If cmp >= 0 Then
            If rson(p) <> NIL Then
                p = rson(p)
            Else
                rson(p) = r
                dad(r) = p
                Exit Sub
            End If
        Else
            If lson(p) <> NIL Then
                p = lson(p)
            Else
                lson(p) = r
                dad(r) = p
                Exit Sub
            End If
        End If
        For i = 1 To F - 1
        x1 = textbuf(r + i): x2 = textbuf(p + i)
            cmp = x1 - x2
            If cmp <> 0 Then Exit For
        Next
        If i > matchlength Then
            matchposition = p
            matchlength = i
            If matchlength >= F Then
                Exit For
            End If
        End If
    Next
    dad(r) = dad(p): lson(r) = lson(p): rson(r) = rson(p)
    dad(lson(p)) = r: dad(rson(p)) = r
    If rson(dad(p)) = p Then
        rson(dad(p)) = r
    Else
        lson(dad(p)) = r
    End If
    dad(p) = NIL

End Sub

Private Sub DeleteNode(ByVal p As Integer)

    Dim q As Integer

    If dad(p) = NIL Then Exit Sub
    If rson(p) = NIL Then
        q = lson(p)
    ElseIf lson(p) = NIL Then
        q = rson(p)
    Else
        q = lson(p)
        If rson(q) <> NIL Then
            Do
                q = rson(q)
            Loop While rson(q) <> NIL
            rson(dad(q)) = lson(q)
            dad(lson(q)) = dad(q)
            lson(q) = lson(p)
            dad(lson(p)) = q
        End If
        rson(q) = rson(p)
        dad(rson(p)) = q
    End If
    dad(q) = dad(p)
    If rson(dad(p)) = p Then
        rson(dad(p)) = q
    Else
        lson(dad(p)) = q
    End If
    dad(p) = NIL

End Sub

Sub LZCompress(InLen&, OutLen&)

    Dim codebuf(17) As Byte, codebufptr As Integer
    Dim i As Integer, cc As Byte, r As Integer, s As Integer
    Dim lenn As Integer, lastmatchlength As Integer, mask As Byte

    InitTree
    codebuf(0) = 0
    InPointer& = 0: OutPointer& = 0
    textsize = 0: codesize = 0

    codebufptr = 1: mask = 1
    s = 0: r = N - F
    For i = s To r - 1: textbuf(i) = 32: Next
    For lenn = 0 To F - 1
        If InPointer& >= InLen& Then Exit For
        cc = InBuffer(InPointer&)
        InPointer& = InPointer& + 1
        textbuf(r + lenn) = cc
    Next
    textsize = lenn
    If textsize = 0 Then Exit Sub
    For i = 1 To F
        InsertNode (r - i)
    Next
    InsertNode (r)
    Do
        If matchlength > lenn Then matchlength = lenn
        If matchlength <= THRESHOLD Then
            matchlength = 1
            codebuf(0) = (codebuf(0) Or mask)
            codebuf(codebufptr) = textbuf(r)
            codebufptr = codebufptr + 1
        Else
            codebuf(codebufptr) = (matchposition And 255)
            codebufptr = codebufptr + 1
            codebuf(codebufptr) = (((matchposition \ 16) And 240) Or _
                               (matchlength - (THRESHOLD + 1))) And 255
            codebufptr = codebufptr + 1
        End If
        mask = (mask * 2) And 255
        If mask = 0 Then
            For i = 0 To codebufptr - 1
                OutBuffer(OutPointer&) = codebuf(i)
                OutPointer& = OutPointer& + 1
            Next
            codesize = codesize + codebufptr
            codebuf(0) = 0
            mask = 1
            codebufptr = mask
        End If
        lastmatchlength = matchlength
        For i = 0 To lastmatchlength - 1
            If InPointer& >= InLen& Then Exit For
            cc = InBuffer(InPointer&)
            InPointer& = InPointer& + 1

            DeleteNode (s)

            textbuf(s) = cc
            If s < (F - 1) Then textbuf(s + N) = cc
            s = (s + 1) And (N - 1): r = (r + 1) And (N - 1)
            InsertNode (r)
        Next
        textsize = textsize + i
        While i < lastmatchlength
            i = i + 1

            DeleteNode (s)

            s = (s + 1) And (N - 1): r = (r + 1) And (N - 1)
            lenn = lenn - 1
            If lenn <> 0 Then InsertNode (r)
        Wend
    Loop While lenn > 0

    If codebufptr > 1 Then
        For i = 0 To codebufptr - 1
            OutBuffer(OutPointer&) = codebuf(i)
            OutPointer& = OutPointer& + 1
        Next
        codesize = codesize + codebufptr
    End If

    OutLen& = codesize

End Sub

Sub LZDecompress(InLen&, OutLen&)

    Dim flags As Byte, flagcount As Byte
    Dim x&, bb As Byte, b1 As Byte, b2 As Byte
    Dim i As Integer, j As Integer, k As Integer, r As Integer

    InPointer& = 0: OutPointer& = 0: Allocated = 0
    For i = 0 To (N - F) - 1: textbuf(i) = 32: Next
    r = N - F: flags = 0
    flagcount = 8

    Infinity& = 2147483647
    For x& = 0 To Infinity&

        If flagcount = 8 Then
            If InPointer& >= InLen& Then Exit For
            bb = InBuffer(InPointer&)
            InPointer& = InPointer& + 1
            flagcount = 0: flags = bb
        End If
        If (flags And 1) = 1 Then
            If InPointer& >= InLen& Then Exit For
            bb = InBuffer(InPointer&)
            InPointer& = InPointer& + 1
            If OutPointer& = Allocated Then
                Allocated = Allocated + 4096
                ReDim Preserve OutBuffer(Allocated - 1)
            End If
            OutBuffer(OutPointer&) = bb
            OutPointer& = OutPointer& + 1
            textbuf(r) = bb: r = r + 1
            r = (r And (N - 1))
        Else
            If InPointer& >= InLen& Then Exit For
            b1 = InBuffer(InPointer&)
            InPointer& = InPointer& + 1
            If InPointer& >= InLen& Then Exit For
            b2 = InBuffer(InPointer&)
            InPointer& = InPointer& + 1
            i = b1: j = b2
            i = i Or ((j And 240) * 16)
            j = (j And 15) + THRESHOLD
            For k = 0 To j
                bb = textbuf((i + k) And (N - 1))
                If OutPointer& = Allocated Then
                    Allocated = Allocated + 4096
                    ReDim Preserve OutBuffer(Allocated - 1)
                End If
                OutBuffer(OutPointer&) = bb
                OutPointer& = OutPointer& + 1
                textbuf(r) = bb
                r = r + 1
                r = (r And (N - 1))
            Next
        End If
        flags = (flags \ 2)
        flagcount = flagcount + 1
    Next

    OutLen& = OutPointer&

End Sub
