Attribute VB_Name = "Module1"
Function Eyncryptlevel1(sData As String) As String
    Dim sTemp As String, sTemp1 As String
    For i = 1 To Len(sData)
        sTemp = Mid(sData, i, 1)
        lT = Asc(sTemp) * 2
        sTemp1 = sTemp1 & Chr(lT)
    Next i
    Eyncrypt = sTemp1
    Form1.text2.Text = Eyncrypt

End Function


Function UnEyncryptlevel1(sData As String) As String

    Dim sTemp As String, sTemp1 As String
    For iI% = 1 To Len(sData$)
        sTemp$ = Mid$(sData$, iI%, 1)
        lT = Asc(sTemp$) \ 2
        sTemp1$ = sTemp1$ & Chr(lT)
    Next iI%
    UnEyncrypt$ = sTemp1$
    Form1.text1.Text = UnEyncrypt$
    lenghtofstring = Len(Form1.text1.Text)
    anystring = Form1.text1.Text
        lenghtofstring = lenghtofstring - 10
        anystring = Right(anystring, lenghtofstring)
        Form1.Text3.Text = ""


End Function
