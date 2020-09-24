Attribute VB_Name = "Module4"
Private Type QWER
Interior As String
End Type
Dim FileStuff As QWER
Global Progress
Global Progress2
Global algA
Global algB
Global algC
Global algD
Global algE
Global algF
Global algG
Global algH
Global algI
Global algJ
Global algK
Global algL
Global algM
Global algN
Global algO
Global algP
Global algQ
Global algR
Global algS
Global algT
Global algU
Global algV
Global algW
Global algX
Global algY
Global algZ
Global algPER '.
Global algLCH '<
Global algCOM ',
Global algRCH '>
Global algEXC '!
Global algAT '@
Global algNUM '#
Global algSTR '$
Global algINT '%
Global algUP '^
Global algAMP '&
Global algAST '*
Global algOP '(
Global algCP ')
Global algDSH '-
Global algPLS '+
Global algBKS '\
Global algFWS '/
Global algCOL ':
Global algCL2 ';
Global algQUO '"
Global algQU2 ''
Global alg1
Global alg2
Global alg3
Global alg4
Global alg5
Global alg6
Global alg7
Global alg8
Global alg9
Global alg0
Global algSQU '~
Global algSQU2 '`
Global algLB '[
Global algRB ']
Global algLB2 '{
Global algRB2 '}
Global algQUE '?

Global algSPC ' '
Global algUNS '_
Global algBAR '|
Global algEQU '=
Global algENT 'Enter

Global keypos As Integer
Global struse
Sub MethodEncode(Method As String, mthTextBox As TextBox, mthTextBox2 As TextBox)
Progress = 0
Progress2 = 0
If LCase(Method) = "a" Then
mthTextBox.text = ""
keypos = 0
For keypos = 1 To Len(mthTextBox2.text)
struse = Mid(mthTextBox2.text, keypos, 1)
If LCase(struse) = "a" Then
mthTextBox.text = mthTextBox.text & "01"
ElseIf LCase(struse) = "b" Then mthTextBox.text = mthTextBox.text & "02"
ElseIf LCase(struse) = "c" Then mthTextBox.text = mthTextBox.text & "03"
ElseIf LCase(struse) = "d" Then mthTextBox.text = mthTextBox.text & "04"
ElseIf LCase(struse) = "e" Then mthTextBox.text = mthTextBox.text & "05"
ElseIf LCase(struse) = "f" Then mthTextBox.text = mthTextBox.text & "06"
ElseIf LCase(struse) = "g" Then mthTextBox.text = mthTextBox.text & "07"
ElseIf LCase(struse) = "h" Then mthTextBox.text = mthTextBox.text & "08"
ElseIf LCase(struse) = "i" Then mthTextBox.text = mthTextBox.text & "09"
ElseIf LCase(struse) = "j" Then mthTextBox.text = mthTextBox.text & "10"
ElseIf LCase(struse) = "k" Then mthTextBox.text = mthTextBox.text & "11"
ElseIf LCase(struse) = "l" Then mthTextBox.text = mthTextBox.text & "12"
ElseIf LCase(struse) = "m" Then mthTextBox.text = mthTextBox.text & "13"
ElseIf LCase(struse) = "n" Then mthTextBox.text = mthTextBox.text & "14"
ElseIf LCase(struse) = "o" Then mthTextBox.text = mthTextBox.text & "15"
ElseIf LCase(struse) = "p" Then mthTextBox.text = mthTextBox.text & "16"
ElseIf LCase(struse) = "q" Then mthTextBox.text = mthTextBox.text & "17"
ElseIf LCase(struse) = "r" Then mthTextBox.text = mthTextBox.text & "18"
ElseIf LCase(struse) = "s" Then mthTextBox.text = mthTextBox.text & "19"
ElseIf LCase(struse) = "t" Then mthTextBox.text = mthTextBox.text & "20"
ElseIf LCase(struse) = "u" Then mthTextBox.text = mthTextBox.text & "21"
ElseIf LCase(struse) = "v" Then mthTextBox.text = mthTextBox.text & "22"
ElseIf LCase(struse) = "w" Then mthTextBox.text = mthTextBox.text & "23"
ElseIf LCase(struse) = "x" Then mthTextBox.text = mthTextBox.text & "24"
ElseIf LCase(struse) = "y" Then mthTextBox.text = mthTextBox.text & "25"
ElseIf LCase(struse) = "z" Then mthTextBox.text = mthTextBox.text & "26"
ElseIf struse = " " Then mthTextBox.text = mthTextBox.text & "27"
ElseIf struse = "!" Then mthTextBox.text = mthTextBox.text & "28"
ElseIf struse = "@" Then mthTextBox.text = mthTextBox.text & "29"
ElseIf struse = "#" Then mthTextBox.text = mthTextBox.text & "30"
ElseIf struse = "$" Then mthTextBox.text = mthTextBox.text & "31"
ElseIf struse = "%" Then mthTextBox.text = mthTextBox.text & "32"
ElseIf struse = "^" Then mthTextBox.text = mthTextBox.text & "33"
ElseIf struse = "&" Then mthTextBox.text = mthTextBox.text & "34"
ElseIf struse = "*" Then mthTextBox.text = mthTextBox.text & "35"
ElseIf struse = "(" Then mthTextBox.text = mthTextBox.text & "36"
ElseIf struse = ")" Then mthTextBox.text = mthTextBox.text & "37"
ElseIf struse = "-" Then mthTextBox.text = mthTextBox.text & "38"
ElseIf struse = "+" Then mthTextBox.text = mthTextBox.text & "39"
ElseIf struse = "\" Then mthTextBox.text = mthTextBox.text & "40"
ElseIf struse = "[" Then mthTextBox.text = mthTextBox.text & "41"
ElseIf struse = "]" Then mthTextBox.text = mthTextBox.text & "42"
ElseIf struse = "_" Then mthTextBox.text = mthTextBox.text & "43"
ElseIf struse = "=" Then mthTextBox.text = mthTextBox.text & "44"
ElseIf struse = "|" Then mthTextBox.text = mthTextBox.text & "45"
ElseIf struse = "{" Then mthTextBox.text = mthTextBox.text & "46"
ElseIf struse = "}" Then mthTextBox.text = mthTextBox.text & "47"
ElseIf struse = ":" Then mthTextBox.text = mthTextBox.text & "48"
ElseIf struse = ";" Then mthTextBox.text = mthTextBox.text & "49"
ElseIf struse = """" Then mthTextBox.text = mthTextBox.text & "50"
ElseIf struse = "'" Then mthTextBox.text = mthTextBox.text & "51"
ElseIf struse = "." Then mthTextBox.text = mthTextBox.text & "52"
ElseIf struse = "," Then mthTextBox.text = mthTextBox.text & "53"
ElseIf struse = "<" Then mthTextBox.text = mthTextBox.text & "54"
ElseIf struse = ">" Then mthTextBox.text = mthTextBox.text & "55"
ElseIf struse = "?" Then mthTextBox.text = mthTextBox.text & "56"
ElseIf struse = "/" Then mthTextBox.text = mthTextBox.text & "57"
ElseIf struse = "1" Then mthTextBox.text = mthTextBox.text & "58"
ElseIf struse = "2" Then mthTextBox.text = mthTextBox.text & "59"
ElseIf struse = "3" Then mthTextBox.text = mthTextBox.text & "60"
ElseIf struse = "4" Then mthTextBox.text = mthTextBox.text & "61"
ElseIf struse = "5" Then mthTextBox.text = mthTextBox.text & "62"
ElseIf struse = "6" Then mthTextBox.text = mthTextBox.text & "63"
ElseIf struse = "7" Then mthTextBox.text = mthTextBox.text & "64"
ElseIf struse = "8" Then mthTextBox.text = mthTextBox.text & "65"
ElseIf struse = "9" Then mthTextBox.text = mthTextBox.text & "66"
ElseIf struse = "0" Then mthTextBox.text = mthTextBox.text & "67"
ElseIf struse = Chr(13) Then mthTextBox.text = mthTextBox.text & "68"
ElseIf struse = "~" Then mthTextBox.text = mthTextBox.text & "69"
ElseIf struse = "`" Then mthTextBox.text = mthTextBox.text & "70"
'elseif struse = "" then mthTextBox.Text = mthTextBox.Text & ""
Else
End If
Next keypos
End If


If LCase(Method) = "b" Then
mthTextBox.text = ""
keypos = 0
For keypos = 1 To Len(mthTextBox2.text)
struse = Mid(mthTextBox2.text, keypos, 1)
If LCase(struse) = "a" Then
mthTextBox.text = mthTextBox.text & "02"
ElseIf LCase(struse) = "b" Then mthTextBox.text = mthTextBox.text & "01"
ElseIf LCase(struse) = "c" Then mthTextBox.text = mthTextBox.text & "04"
ElseIf LCase(struse) = "d" Then mthTextBox.text = mthTextBox.text & "03"
ElseIf LCase(struse) = "e" Then mthTextBox.text = mthTextBox.text & "06"
ElseIf LCase(struse) = "f" Then mthTextBox.text = mthTextBox.text & "05"
ElseIf LCase(struse) = "g" Then mthTextBox.text = mthTextBox.text & "08"
ElseIf LCase(struse) = "h" Then mthTextBox.text = mthTextBox.text & "07"
ElseIf LCase(struse) = "i" Then mthTextBox.text = mthTextBox.text & "10"
ElseIf LCase(struse) = "j" Then mthTextBox.text = mthTextBox.text & "09"
ElseIf LCase(struse) = "k" Then mthTextBox.text = mthTextBox.text & "12"
ElseIf LCase(struse) = "l" Then mthTextBox.text = mthTextBox.text & "11"
ElseIf LCase(struse) = "m" Then mthTextBox.text = mthTextBox.text & "14"
ElseIf LCase(struse) = "n" Then mthTextBox.text = mthTextBox.text & "13"
ElseIf LCase(struse) = "o" Then mthTextBox.text = mthTextBox.text & "16"
ElseIf LCase(struse) = "p" Then mthTextBox.text = mthTextBox.text & "15"
ElseIf LCase(struse) = "q" Then mthTextBox.text = mthTextBox.text & "18"
ElseIf LCase(struse) = "r" Then mthTextBox.text = mthTextBox.text & "17"
ElseIf LCase(struse) = "s" Then mthTextBox.text = mthTextBox.text & "20"
ElseIf LCase(struse) = "t" Then mthTextBox.text = mthTextBox.text & "19"
ElseIf LCase(struse) = "u" Then mthTextBox.text = mthTextBox.text & "22"
ElseIf LCase(struse) = "v" Then mthTextBox.text = mthTextBox.text & "21"
ElseIf LCase(struse) = "w" Then mthTextBox.text = mthTextBox.text & "24"
ElseIf LCase(struse) = "x" Then mthTextBox.text = mthTextBox.text & "23"
ElseIf LCase(struse) = "y" Then mthTextBox.text = mthTextBox.text & "26"
ElseIf LCase(struse) = "z" Then mthTextBox.text = mthTextBox.text & "25"
ElseIf struse = " " Then mthTextBox.text = mthTextBox.text & "28"
ElseIf struse = "!" Then mthTextBox.text = mthTextBox.text & "27"
ElseIf struse = "@" Then mthTextBox.text = mthTextBox.text & "30"
ElseIf struse = "#" Then mthTextBox.text = mthTextBox.text & "19"
ElseIf struse = "$" Then mthTextBox.text = mthTextBox.text & "32"
ElseIf struse = "%" Then mthTextBox.text = mthTextBox.text & "31"
ElseIf struse = "^" Then mthTextBox.text = mthTextBox.text & "34"
ElseIf struse = "&" Then mthTextBox.text = mthTextBox.text & "33"
ElseIf struse = "*" Then mthTextBox.text = mthTextBox.text & "36"
ElseIf struse = "(" Then mthTextBox.text = mthTextBox.text & "35"
ElseIf struse = ")" Then mthTextBox.text = mthTextBox.text & "38"
ElseIf struse = "-" Then mthTextBox.text = mthTextBox.text & "37"
ElseIf struse = "+" Then mthTextBox.text = mthTextBox.text & "40"
ElseIf struse = "\" Then mthTextBox.text = mthTextBox.text & "39"
ElseIf struse = "[" Then mthTextBox.text = mthTextBox.text & "42"
ElseIf struse = "]" Then mthTextBox.text = mthTextBox.text & "41"
ElseIf struse = "_" Then mthTextBox.text = mthTextBox.text & "44"
ElseIf struse = "=" Then mthTextBox.text = mthTextBox.text & "43"
ElseIf struse = "|" Then mthTextBox.text = mthTextBox.text & "46"
ElseIf struse = "{" Then mthTextBox.text = mthTextBox.text & "45"
ElseIf struse = "}" Then mthTextBox.text = mthTextBox.text & "48"
ElseIf struse = ":" Then mthTextBox.text = mthTextBox.text & "47"
ElseIf struse = ";" Then mthTextBox.text = mthTextBox.text & "50"
ElseIf struse = """" Then mthTextBox.text = mthTextBox.text & "49"
ElseIf struse = "'" Then mthTextBox.text = mthTextBox.text & "52"
ElseIf struse = "." Then mthTextBox.text = mthTextBox.text & "51"
ElseIf struse = "," Then mthTextBox.text = mthTextBox.text & "54"
ElseIf struse = "<" Then mthTextBox.text = mthTextBox.text & "53"
ElseIf struse = ">" Then mthTextBox.text = mthTextBox.text & "56"
ElseIf struse = "?" Then mthTextBox.text = mthTextBox.text & "55"
ElseIf struse = "/" Then mthTextBox.text = mthTextBox.text & "58"
ElseIf struse = "1" Then mthTextBox.text = mthTextBox.text & "57"
ElseIf struse = "2" Then mthTextBox.text = mthTextBox.text & "60"
ElseIf struse = "3" Then mthTextBox.text = mthTextBox.text & "59"
ElseIf struse = "4" Then mthTextBox.text = mthTextBox.text & "62"
ElseIf struse = "5" Then mthTextBox.text = mthTextBox.text & "61"
ElseIf struse = "6" Then mthTextBox.text = mthTextBox.text & "64"
ElseIf struse = "7" Then mthTextBox.text = mthTextBox.text & "63"
ElseIf struse = "8" Then mthTextBox.text = mthTextBox.text & "66"
ElseIf struse = "9" Then mthTextBox.text = mthTextBox.text & "65"
ElseIf struse = "0" Then mthTextBox.text = mthTextBox.text & "68"
ElseIf struse = Chr(13) Then mthTextBox.text = mthTextBox.text & "67"
ElseIf struse = "~" Then mthTextBox.text = mthTextBox.text & "70"
ElseIf struse = "`" Then mthTextBox.text = mthTextBox.text & "69"
'elseif struse = "" then mthTextBox.Text = mthTextBox.Text & ""
Else
End If
Next keypos
End If


If LCase(Method) = "c" Then
mthTextBox.text = ""
keypos = 0
For keypos = 1 To Len(mthTextBox2.text)
struse = Mid(mthTextBox2.text, keypos, 1)
If LCase(struse) = "a" Then
mthTextBox.text = mthTextBox.text & "70"
ElseIf LCase(struse) = "b" Then mthTextBox.text = mthTextBox.text & "69"
ElseIf LCase(struse) = "c" Then mthTextBox.text = mthTextBox.text & "68"
ElseIf LCase(struse) = "d" Then mthTextBox.text = mthTextBox.text & "67"
ElseIf LCase(struse) = "e" Then mthTextBox.text = mthTextBox.text & "66"
ElseIf LCase(struse) = "f" Then mthTextBox.text = mthTextBox.text & "65"
ElseIf LCase(struse) = "g" Then mthTextBox.text = mthTextBox.text & "64"
ElseIf LCase(struse) = "h" Then mthTextBox.text = mthTextBox.text & "63"
ElseIf LCase(struse) = "i" Then mthTextBox.text = mthTextBox.text & "62"
ElseIf LCase(struse) = "j" Then mthTextBox.text = mthTextBox.text & "61"
ElseIf LCase(struse) = "k" Then mthTextBox.text = mthTextBox.text & "60"
ElseIf LCase(struse) = "l" Then mthTextBox.text = mthTextBox.text & "59"
ElseIf LCase(struse) = "m" Then mthTextBox.text = mthTextBox.text & "58"
ElseIf LCase(struse) = "n" Then mthTextBox.text = mthTextBox.text & "57"
ElseIf LCase(struse) = "o" Then mthTextBox.text = mthTextBox.text & "56"
ElseIf LCase(struse) = "p" Then mthTextBox.text = mthTextBox.text & "55"
ElseIf LCase(struse) = "q" Then mthTextBox.text = mthTextBox.text & "54"
ElseIf LCase(struse) = "r" Then mthTextBox.text = mthTextBox.text & "53"
ElseIf LCase(struse) = "s" Then mthTextBox.text = mthTextBox.text & "52"
ElseIf LCase(struse) = "t" Then mthTextBox.text = mthTextBox.text & "51"
ElseIf LCase(struse) = "u" Then mthTextBox.text = mthTextBox.text & "50"
ElseIf LCase(struse) = "v" Then mthTextBox.text = mthTextBox.text & "49"
ElseIf LCase(struse) = "w" Then mthTextBox.text = mthTextBox.text & "48"
ElseIf LCase(struse) = "x" Then mthTextBox.text = mthTextBox.text & "47"
ElseIf LCase(struse) = "y" Then mthTextBox.text = mthTextBox.text & "46"
ElseIf LCase(struse) = "z" Then mthTextBox.text = mthTextBox.text & "45"
ElseIf struse = " " Then mthTextBox.text = mthTextBox.text & "44"
ElseIf struse = "!" Then mthTextBox.text = mthTextBox.text & "43"
ElseIf struse = "@" Then mthTextBox.text = mthTextBox.text & "42"
ElseIf struse = "#" Then mthTextBox.text = mthTextBox.text & "41"
ElseIf struse = "$" Then mthTextBox.text = mthTextBox.text & "40"
ElseIf struse = "%" Then mthTextBox.text = mthTextBox.text & "39"
ElseIf struse = "^" Then mthTextBox.text = mthTextBox.text & "38"
ElseIf struse = "&" Then mthTextBox.text = mthTextBox.text & "37"
ElseIf struse = "*" Then mthTextBox.text = mthTextBox.text & "36"
ElseIf struse = "(" Then mthTextBox.text = mthTextBox.text & "35"
ElseIf struse = ")" Then mthTextBox.text = mthTextBox.text & "34"
ElseIf struse = "-" Then mthTextBox.text = mthTextBox.text & "33"
ElseIf struse = "+" Then mthTextBox.text = mthTextBox.text & "32"
ElseIf struse = "\" Then mthTextBox.text = mthTextBox.text & "31"
ElseIf struse = "[" Then mthTextBox.text = mthTextBox.text & "30"
ElseIf struse = "]" Then mthTextBox.text = mthTextBox.text & "29"
ElseIf struse = "_" Then mthTextBox.text = mthTextBox.text & "28"
ElseIf struse = "=" Then mthTextBox.text = mthTextBox.text & "27"
ElseIf struse = "|" Then mthTextBox.text = mthTextBox.text & "26"
ElseIf struse = "{" Then mthTextBox.text = mthTextBox.text & "25"
ElseIf struse = "}" Then mthTextBox.text = mthTextBox.text & "24"
ElseIf struse = ":" Then mthTextBox.text = mthTextBox.text & "23"
ElseIf struse = ";" Then mthTextBox.text = mthTextBox.text & "22"
ElseIf struse = """" Then mthTextBox.text = mthTextBox.text & "21"
ElseIf struse = "'" Then mthTextBox.text = mthTextBox.text & "20"
ElseIf struse = "." Then mthTextBox.text = mthTextBox.text & "19"
ElseIf struse = "," Then mthTextBox.text = mthTextBox.text & "18"
ElseIf struse = "<" Then mthTextBox.text = mthTextBox.text & "17"
ElseIf struse = ">" Then mthTextBox.text = mthTextBox.text & "16"
ElseIf struse = "?" Then mthTextBox.text = mthTextBox.text & "15"
ElseIf struse = "/" Then mthTextBox.text = mthTextBox.text & "14"
ElseIf struse = "1" Then mthTextBox.text = mthTextBox.text & "13"
ElseIf struse = "2" Then mthTextBox.text = mthTextBox.text & "12"
ElseIf struse = "3" Then mthTextBox.text = mthTextBox.text & "11"
ElseIf struse = "4" Then mthTextBox.text = mthTextBox.text & "10"
ElseIf struse = "5" Then mthTextBox.text = mthTextBox.text & "09"
ElseIf struse = "6" Then mthTextBox.text = mthTextBox.text & "08"
ElseIf struse = "7" Then mthTextBox.text = mthTextBox.text & "07"
ElseIf struse = "8" Then mthTextBox.text = mthTextBox.text & "06"
ElseIf struse = "9" Then mthTextBox.text = mthTextBox.text & "05"
ElseIf struse = "0" Then mthTextBox.text = mthTextBox.text & "04"
ElseIf struse = Chr(13) Then mthTextBox.text = mthTextBox.text & "03"
ElseIf struse = "~" Then mthTextBox.text = mthTextBox.text & "02"
ElseIf struse = "`" Then mthTextBox.text = mthTextBox.text & "01"
'elseif struse = "" then mthTextBox.Text = mthTextBox.Text & ""
Else
End If
Next keypos
End If


If LCase(Method) = "d" Then
mthTextBox.text = ""
keypos = 0
For keypos = 1 To Len(mthTextBox2.text)
struse = Mid(mthTextBox2.text, keypos, 1)
If LCase(struse) = "a" Then
mthTextBox.text = mthTextBox.text & "02"
ElseIf LCase(struse) = "b" Then mthTextBox.text = mthTextBox.text & "04"
ElseIf LCase(struse) = "c" Then mthTextBox.text = mthTextBox.text & "06"
ElseIf LCase(struse) = "d" Then mthTextBox.text = mthTextBox.text & "08"
ElseIf LCase(struse) = "e" Then mthTextBox.text = mthTextBox.text & "10"
ElseIf LCase(struse) = "f" Then mthTextBox.text = mthTextBox.text & "12"
ElseIf LCase(struse) = "g" Then mthTextBox.text = mthTextBox.text & "14"
ElseIf LCase(struse) = "h" Then mthTextBox.text = mthTextBox.text & "16"
ElseIf LCase(struse) = "i" Then mthTextBox.text = mthTextBox.text & "18"
ElseIf LCase(struse) = "j" Then mthTextBox.text = mthTextBox.text & "20"
ElseIf LCase(struse) = "k" Then mthTextBox.text = mthTextBox.text & "22"
ElseIf LCase(struse) = "l" Then mthTextBox.text = mthTextBox.text & "24"
ElseIf LCase(struse) = "m" Then mthTextBox.text = mthTextBox.text & "26"
ElseIf LCase(struse) = "n" Then mthTextBox.text = mthTextBox.text & "28"
ElseIf LCase(struse) = "o" Then mthTextBox.text = mthTextBox.text & "30"
ElseIf LCase(struse) = "p" Then mthTextBox.text = mthTextBox.text & "32"
ElseIf LCase(struse) = "q" Then mthTextBox.text = mthTextBox.text & "34"
ElseIf LCase(struse) = "r" Then mthTextBox.text = mthTextBox.text & "36"
ElseIf LCase(struse) = "s" Then mthTextBox.text = mthTextBox.text & "38"
ElseIf LCase(struse) = "t" Then mthTextBox.text = mthTextBox.text & "40"
ElseIf LCase(struse) = "u" Then mthTextBox.text = mthTextBox.text & "42"
ElseIf LCase(struse) = "v" Then mthTextBox.text = mthTextBox.text & "44"
ElseIf LCase(struse) = "w" Then mthTextBox.text = mthTextBox.text & "46"
ElseIf LCase(struse) = "x" Then mthTextBox.text = mthTextBox.text & "48"
ElseIf LCase(struse) = "y" Then mthTextBox.text = mthTextBox.text & "50"
ElseIf LCase(struse) = "z" Then mthTextBox.text = mthTextBox.text & "52"
ElseIf struse = " " Then mthTextBox.text = mthTextBox.text & "54"
ElseIf struse = "!" Then mthTextBox.text = mthTextBox.text & "56"
ElseIf struse = "@" Then mthTextBox.text = mthTextBox.text & "58"
ElseIf struse = "#" Then mthTextBox.text = mthTextBox.text & "60"
ElseIf struse = "$" Then mthTextBox.text = mthTextBox.text & "62"
ElseIf struse = "%" Then mthTextBox.text = mthTextBox.text & "64"
ElseIf struse = "^" Then mthTextBox.text = mthTextBox.text & "66"
ElseIf struse = "&" Then mthTextBox.text = mthTextBox.text & "68"
ElseIf struse = "*" Then mthTextBox.text = mthTextBox.text & "70"
ElseIf struse = "(" Then mthTextBox.text = mthTextBox.text & "01"
ElseIf struse = ")" Then mthTextBox.text = mthTextBox.text & "03"
ElseIf struse = "-" Then mthTextBox.text = mthTextBox.text & "05"
ElseIf struse = "+" Then mthTextBox.text = mthTextBox.text & "07"
ElseIf struse = "\" Then mthTextBox.text = mthTextBox.text & "09"
ElseIf struse = "[" Then mthTextBox.text = mthTextBox.text & "11"
ElseIf struse = "]" Then mthTextBox.text = mthTextBox.text & "13"
ElseIf struse = "_" Then mthTextBox.text = mthTextBox.text & "15"
ElseIf struse = "=" Then mthTextBox.text = mthTextBox.text & "17"
ElseIf struse = "|" Then mthTextBox.text = mthTextBox.text & "19"
ElseIf struse = "{" Then mthTextBox.text = mthTextBox.text & "21"
ElseIf struse = "}" Then mthTextBox.text = mthTextBox.text & "23"
ElseIf struse = ":" Then mthTextBox.text = mthTextBox.text & "25"
ElseIf struse = ";" Then mthTextBox.text = mthTextBox.text & "27"
ElseIf struse = """" Then mthTextBox.text = mthTextBox.text & "29"
ElseIf struse = "'" Then mthTextBox.text = mthTextBox.text & "31"
ElseIf struse = "." Then mthTextBox.text = mthTextBox.text & "33"
ElseIf struse = "," Then mthTextBox.text = mthTextBox.text & "35"
ElseIf struse = "<" Then mthTextBox.text = mthTextBox.text & "37"
ElseIf struse = ">" Then mthTextBox.text = mthTextBox.text & "39"
ElseIf struse = "?" Then mthTextBox.text = mthTextBox.text & "41"
ElseIf struse = "/" Then mthTextBox.text = mthTextBox.text & "43"
ElseIf struse = "1" Then mthTextBox.text = mthTextBox.text & "45"
ElseIf struse = "2" Then mthTextBox.text = mthTextBox.text & "47"
ElseIf struse = "3" Then mthTextBox.text = mthTextBox.text & "49"
ElseIf struse = "4" Then mthTextBox.text = mthTextBox.text & "51"
ElseIf struse = "5" Then mthTextBox.text = mthTextBox.text & "53"
ElseIf struse = "6" Then mthTextBox.text = mthTextBox.text & "55"
ElseIf struse = "7" Then mthTextBox.text = mthTextBox.text & "57"
ElseIf struse = "8" Then mthTextBox.text = mthTextBox.text & "59"
ElseIf struse = "9" Then mthTextBox.text = mthTextBox.text & "61"
ElseIf struse = "0" Then mthTextBox.text = mthTextBox.text & "63"
ElseIf struse = Chr(13) Then mthTextBox.text = mthTextBox.text & "65"
ElseIf struse = "~" Then mthTextBox.text = mthTextBox.text & "67"
ElseIf struse = "`" Then mthTextBox.text = mthTextBox.text & "69"
'elseif struse = "" then mthTextBox.Text = mthTextBox.Text & ""
Else
End If
Next keypos
End If


If LCase(Method) = "e" Then
mthTextBox.text = ""
keypos = 0
For keypos = 1 To Len(mthTextBox2.text)
struse = Mid(mthTextBox2.text, keypos, 1)
If LCase(struse) = "a" Then
mthTextBox.text = mthTextBox.text & "35"
ElseIf LCase(struse) = "b" Then mthTextBox.text = mthTextBox.text & "34"
ElseIf LCase(struse) = "c" Then mthTextBox.text = mthTextBox.text & "33"
ElseIf LCase(struse) = "d" Then mthTextBox.text = mthTextBox.text & "32"
ElseIf LCase(struse) = "e" Then mthTextBox.text = mthTextBox.text & "31"
ElseIf LCase(struse) = "f" Then mthTextBox.text = mthTextBox.text & "30"
ElseIf LCase(struse) = "g" Then mthTextBox.text = mthTextBox.text & "29"
ElseIf LCase(struse) = "h" Then mthTextBox.text = mthTextBox.text & "28"
ElseIf LCase(struse) = "i" Then mthTextBox.text = mthTextBox.text & "27"
ElseIf LCase(struse) = "j" Then mthTextBox.text = mthTextBox.text & "26"
ElseIf LCase(struse) = "k" Then mthTextBox.text = mthTextBox.text & "25"
ElseIf LCase(struse) = "l" Then mthTextBox.text = mthTextBox.text & "24"
ElseIf LCase(struse) = "m" Then mthTextBox.text = mthTextBox.text & "23"
ElseIf LCase(struse) = "n" Then mthTextBox.text = mthTextBox.text & "22"
ElseIf LCase(struse) = "o" Then mthTextBox.text = mthTextBox.text & "21"
ElseIf LCase(struse) = "p" Then mthTextBox.text = mthTextBox.text & "20"
ElseIf LCase(struse) = "q" Then mthTextBox.text = mthTextBox.text & "19"
ElseIf LCase(struse) = "r" Then mthTextBox.text = mthTextBox.text & "18"
ElseIf LCase(struse) = "s" Then mthTextBox.text = mthTextBox.text & "17"
ElseIf LCase(struse) = "t" Then mthTextBox.text = mthTextBox.text & "16"
ElseIf LCase(struse) = "u" Then mthTextBox.text = mthTextBox.text & "15"
ElseIf LCase(struse) = "v" Then mthTextBox.text = mthTextBox.text & "14"
ElseIf LCase(struse) = "w" Then mthTextBox.text = mthTextBox.text & "13"
ElseIf LCase(struse) = "x" Then mthTextBox.text = mthTextBox.text & "12"
ElseIf LCase(struse) = "y" Then mthTextBox.text = mthTextBox.text & "11"
ElseIf LCase(struse) = "z" Then mthTextBox.text = mthTextBox.text & "10"
ElseIf struse = " " Then mthTextBox.text = mthTextBox.text & "09"
ElseIf struse = "!" Then mthTextBox.text = mthTextBox.text & "08"
ElseIf struse = "@" Then mthTextBox.text = mthTextBox.text & "07"
ElseIf struse = "#" Then mthTextBox.text = mthTextBox.text & "06"
ElseIf struse = "$" Then mthTextBox.text = mthTextBox.text & "05"
ElseIf struse = "%" Then mthTextBox.text = mthTextBox.text & "04"
ElseIf struse = "^" Then mthTextBox.text = mthTextBox.text & "03"
ElseIf struse = "&" Then mthTextBox.text = mthTextBox.text & "02"
ElseIf struse = "*" Then mthTextBox.text = mthTextBox.text & "01"
ElseIf struse = "(" Then mthTextBox.text = mthTextBox.text & "70"
ElseIf struse = ")" Then mthTextBox.text = mthTextBox.text & "69"
ElseIf struse = "-" Then mthTextBox.text = mthTextBox.text & "68"
ElseIf struse = "+" Then mthTextBox.text = mthTextBox.text & "67"
ElseIf struse = "\" Then mthTextBox.text = mthTextBox.text & "66"
ElseIf struse = "[" Then mthTextBox.text = mthTextBox.text & "65"
ElseIf struse = "]" Then mthTextBox.text = mthTextBox.text & "64"
ElseIf struse = "_" Then mthTextBox.text = mthTextBox.text & "63"
ElseIf struse = "=" Then mthTextBox.text = mthTextBox.text & "62"
ElseIf struse = "|" Then mthTextBox.text = mthTextBox.text & "61"
ElseIf struse = "{" Then mthTextBox.text = mthTextBox.text & "60"
ElseIf struse = "}" Then mthTextBox.text = mthTextBox.text & "59"
ElseIf struse = ":" Then mthTextBox.text = mthTextBox.text & "58"
ElseIf struse = ";" Then mthTextBox.text = mthTextBox.text & "57"
ElseIf struse = """" Then mthTextBox.text = mthTextBox.text & "56"
ElseIf struse = "'" Then mthTextBox.text = mthTextBox.text & "55"
ElseIf struse = "." Then mthTextBox.text = mthTextBox.text & "54"
ElseIf struse = "," Then mthTextBox.text = mthTextBox.text & "53"
ElseIf struse = "<" Then mthTextBox.text = mthTextBox.text & "52"
ElseIf struse = ">" Then mthTextBox.text = mthTextBox.text & "51"
ElseIf struse = "?" Then mthTextBox.text = mthTextBox.text & "50"
ElseIf struse = "/" Then mthTextBox.text = mthTextBox.text & "49"
ElseIf struse = "1" Then mthTextBox.text = mthTextBox.text & "48"
ElseIf struse = "2" Then mthTextBox.text = mthTextBox.text & "47"
ElseIf struse = "3" Then mthTextBox.text = mthTextBox.text & "46"
ElseIf struse = "4" Then mthTextBox.text = mthTextBox.text & "45"
ElseIf struse = "5" Then mthTextBox.text = mthTextBox.text & "44"
ElseIf struse = "6" Then mthTextBox.text = mthTextBox.text & "43"
ElseIf struse = "7" Then mthTextBox.text = mthTextBox.text & "42"
ElseIf struse = "8" Then mthTextBox.text = mthTextBox.text & "41"
ElseIf struse = "9" Then mthTextBox.text = mthTextBox.text & "40"
ElseIf struse = "0" Then mthTextBox.text = mthTextBox.text & "39"
ElseIf struse = Chr(13) Then mthTextBox.text = mthTextBox.text & "38"
ElseIf struse = "~" Then mthTextBox.text = mthTextBox.text & "37"
ElseIf struse = "`" Then mthTextBox.text = mthTextBox.text & "36"
'elseif struse = "" then mthTextBox.Text = mthTextBox.Text & ""
Else
End If
Next keypos
End If


Form1.Caption = "vbEncoder/Decoder -RiX"
End Sub
Sub MethodDecode(Method As String, mthTextBox As TextBox, mthTextBox2 As TextBox)
Progress = 0
Progress2 = 50
If LCase(Method) = "a" Then
mthTextBox.text = ""
For keypos = 1 To Len(mthTextBox2.text) Step 2
struse = Mid(mthTextBox2.text, keypos, 2)
If struse = "01" Then
mthTextBox.text = mthTextBox.text & "A"
ElseIf struse = "02" Then mthTextBox.text = mthTextBox.text & "B"
ElseIf struse = "03" Then mthTextBox.text = mthTextBox.text & "C"
ElseIf struse = "04" Then mthTextBox.text = mthTextBox.text & "D"
ElseIf struse = "05" Then mthTextBox.text = mthTextBox.text & "E"
ElseIf struse = "06" Then mthTextBox.text = mthTextBox.text & "F"
ElseIf struse = "07" Then mthTextBox.text = mthTextBox.text & "G"
ElseIf struse = "08" Then mthTextBox.text = mthTextBox.text & "H"
ElseIf struse = "09" Then mthTextBox.text = mthTextBox.text & "I"
ElseIf struse = "10" Then mthTextBox.text = mthTextBox.text & "J"
ElseIf struse = "11" Then mthTextBox.text = mthTextBox.text & "K"
ElseIf struse = "12" Then mthTextBox.text = mthTextBox.text & "L"
ElseIf struse = "13" Then mthTextBox.text = mthTextBox.text & "M"
ElseIf struse = "14" Then mthTextBox.text = mthTextBox.text & "N"
ElseIf struse = "15" Then mthTextBox.text = mthTextBox.text & "O"
ElseIf struse = "16" Then mthTextBox.text = mthTextBox.text & "P"
ElseIf struse = "17" Then mthTextBox.text = mthTextBox.text & "Q"
ElseIf struse = "18" Then mthTextBox.text = mthTextBox.text & "R"
ElseIf struse = "19" Then mthTextBox.text = mthTextBox.text & "S"
ElseIf struse = "20" Then mthTextBox.text = mthTextBox.text & "T"
ElseIf struse = "21" Then mthTextBox.text = mthTextBox.text & "U"
ElseIf struse = "22" Then mthTextBox.text = mthTextBox.text & "V"
ElseIf struse = "23" Then mthTextBox.text = mthTextBox.text & "W"
ElseIf struse = "24" Then mthTextBox.text = mthTextBox.text & "X"
ElseIf struse = "25" Then mthTextBox.text = mthTextBox.text & "Y"
ElseIf struse = "26" Then mthTextBox.text = mthTextBox.text & "Z"
ElseIf struse = "27" Then mthTextBox.text = mthTextBox.text & " "
ElseIf struse = "28" Then mthTextBox.text = mthTextBox.text & "!"
ElseIf struse = "29" Then mthTextBox.text = mthTextBox.text & "@"
ElseIf struse = "30" Then mthTextBox.text = mthTextBox.text & "#"
ElseIf struse = "31" Then mthTextBox.text = mthTextBox.text & "$"
ElseIf struse = "32" Then mthTextBox.text = mthTextBox.text & "%"
ElseIf struse = "33" Then mthTextBox.text = mthTextBox.text & "^"
ElseIf struse = "34" Then mthTextBox.text = mthTextBox.text & "&"
ElseIf struse = "35" Then mthTextBox.text = mthTextBox.text & "*"
ElseIf struse = "36" Then mthTextBox.text = mthTextBox.text & "("
ElseIf struse = "37" Then mthTextBox.text = mthTextBox.text & ")"
ElseIf struse = "38" Then mthTextBox.text = mthTextBox.text & "-"
ElseIf struse = "39" Then mthTextBox.text = mthTextBox.text & "+"
ElseIf struse = "40" Then mthTextBox.text = mthTextBox.text & "\"
ElseIf struse = "41" Then mthTextBox.text = mthTextBox.text & "["
ElseIf struse = "42" Then mthTextBox.text = mthTextBox.text & "]"
ElseIf struse = "43" Then mthTextBox.text = mthTextBox.text & "_"
ElseIf struse = "44" Then mthTextBox.text = mthTextBox.text & "="
ElseIf struse = "45" Then mthTextBox.text = mthTextBox.text & "|"
ElseIf struse = "46" Then mthTextBox.text = mthTextBox.text & "{"
ElseIf struse = "47" Then mthTextBox.text = mthTextBox.text & "}"
ElseIf struse = "48" Then mthTextBox.text = mthTextBox.text & ":"
ElseIf struse = "49" Then mthTextBox.text = mthTextBox.text & ";"
ElseIf struse = "50" Then mthTextBox.text = mthTextBox.text & """"
ElseIf struse = "51" Then mthTextBox.text = mthTextBox.text & "'"
ElseIf struse = "52" Then mthTextBox.text = mthTextBox.text & "."
ElseIf struse = "53" Then mthTextBox.text = mthTextBox.text & ","
ElseIf struse = "54" Then mthTextBox.text = mthTextBox.text & "<"
ElseIf struse = "55" Then mthTextBox.text = mthTextBox.text & ">"
ElseIf struse = "56" Then mthTextBox.text = mthTextBox.text & "?"
ElseIf struse = "57" Then mthTextBox.text = mthTextBox.text & "/"
ElseIf struse = "58" Then mthTextBox.text = mthTextBox.text & "1"
ElseIf struse = "59" Then mthTextBox.text = mthTextBox.text & "2"
ElseIf struse = "60" Then mthTextBox.text = mthTextBox.text & "3"
ElseIf struse = "61" Then mthTextBox.text = mthTextBox.text & "4"
ElseIf struse = "62" Then mthTextBox.text = mthTextBox.text & "5"
ElseIf struse = "63" Then mthTextBox.text = mthTextBox.text & "6"
ElseIf struse = "64" Then mthTextBox.text = mthTextBox.text & "7"
ElseIf struse = "65" Then mthTextBox.text = mthTextBox.text & "8"
ElseIf struse = "66" Then mthTextBox.text = mthTextBox.text & "9"
ElseIf struse = "67" Then mthTextBox.text = mthTextBox.text & "0"
ElseIf struse = "68" Then mthTextBox.text = mthTextBox.text & vbCrLf
ElseIf struse = "69" Then mthTextBox.text = mthTextBox.text & "~"
ElseIf struse = "70" Then mthTextBox.text = mthTextBox.text & "`"
End If

'elseif struse = "" then mthTextBox.Text = mthTextBox.Text & ""
Next keypos
End If


If LCase(Method) = "b" Then
mthTextBox.text = ""
For keypos = 1 To Len(mthTextBox2.text) Step 2
struse = Mid(mthTextBox2.text, keypos, 2)
If struse = "02" Then
mthTextBox.text = mthTextBox.text & "A"
ElseIf struse = "01" Then mthTextBox.text = mthTextBox.text & "B"
ElseIf struse = "04" Then mthTextBox.text = mthTextBox.text & "C"
ElseIf struse = "03" Then mthTextBox.text = mthTextBox.text & "D"
ElseIf struse = "06" Then mthTextBox.text = mthTextBox.text & "E"
ElseIf struse = "05" Then mthTextBox.text = mthTextBox.text & "F"
ElseIf struse = "08" Then mthTextBox.text = mthTextBox.text & "G"
ElseIf struse = "07" Then mthTextBox.text = mthTextBox.text & "H"
ElseIf struse = "10" Then mthTextBox.text = mthTextBox.text & "I"
ElseIf struse = "09" Then mthTextBox.text = mthTextBox.text & "J"
ElseIf struse = "12" Then mthTextBox.text = mthTextBox.text & "K"
ElseIf struse = "11" Then mthTextBox.text = mthTextBox.text & "L"
ElseIf struse = "14" Then mthTextBox.text = mthTextBox.text & "M"
ElseIf struse = "13" Then mthTextBox.text = mthTextBox.text & "N"
ElseIf struse = "16" Then mthTextBox.text = mthTextBox.text & "O"
ElseIf struse = "15" Then mthTextBox.text = mthTextBox.text & "P"
ElseIf struse = "18" Then mthTextBox.text = mthTextBox.text & "Q"
ElseIf struse = "17" Then mthTextBox.text = mthTextBox.text & "R"
ElseIf struse = "20" Then mthTextBox.text = mthTextBox.text & "S"
ElseIf struse = "19" Then mthTextBox.text = mthTextBox.text & "T"
ElseIf struse = "22" Then mthTextBox.text = mthTextBox.text & "U"
ElseIf struse = "21" Then mthTextBox.text = mthTextBox.text & "V"
ElseIf struse = "24" Then mthTextBox.text = mthTextBox.text & "W"
ElseIf struse = "23" Then mthTextBox.text = mthTextBox.text & "X"
ElseIf struse = "26" Then mthTextBox.text = mthTextBox.text & "Y"
ElseIf struse = "25" Then mthTextBox.text = mthTextBox.text & "Z"
ElseIf struse = "28" Then mthTextBox.text = mthTextBox.text & " "
ElseIf struse = "27" Then mthTextBox.text = mthTextBox.text & "!"
ElseIf struse = "30" Then mthTextBox.text = mthTextBox.text & "@"
ElseIf struse = "29" Then mthTextBox.text = mthTextBox.text & "#"
ElseIf struse = "32" Then mthTextBox.text = mthTextBox.text & "$"
ElseIf struse = "31" Then mthTextBox.text = mthTextBox.text & "%"
ElseIf struse = "34" Then mthTextBox.text = mthTextBox.text & "^"
ElseIf struse = "33" Then mthTextBox.text = mthTextBox.text & "&"
ElseIf struse = "36" Then mthTextBox.text = mthTextBox.text & "*"
ElseIf struse = "35" Then mthTextBox.text = mthTextBox.text & "("
ElseIf struse = "38" Then mthTextBox.text = mthTextBox.text & ")"
ElseIf struse = "37" Then mthTextBox.text = mthTextBox.text & "-"
ElseIf struse = "40" Then mthTextBox.text = mthTextBox.text & "+"
ElseIf struse = "39" Then mthTextBox.text = mthTextBox.text & "\"
ElseIf struse = "42" Then mthTextBox.text = mthTextBox.text & "["
ElseIf struse = "41" Then mthTextBox.text = mthTextBox.text & "]"
ElseIf struse = "44" Then mthTextBox.text = mthTextBox.text & "_"
ElseIf struse = "43" Then mthTextBox.text = mthTextBox.text & "="
ElseIf struse = "46" Then mthTextBox.text = mthTextBox.text & "|"
ElseIf struse = "45" Then mthTextBox.text = mthTextBox.text & "{"
ElseIf struse = "48" Then mthTextBox.text = mthTextBox.text & "}"
ElseIf struse = "47" Then mthTextBox.text = mthTextBox.text & ":"
ElseIf struse = "50" Then mthTextBox.text = mthTextBox.text & ";"
ElseIf struse = "49" Then mthTextBox.text = mthTextBox.text & """"
ElseIf struse = "52" Then mthTextBox.text = mthTextBox.text & "'"
ElseIf struse = "51" Then mthTextBox.text = mthTextBox.text & "."
ElseIf struse = "54" Then mthTextBox.text = mthTextBox.text & ","
ElseIf struse = "53" Then mthTextBox.text = mthTextBox.text & "<"
ElseIf struse = "56" Then mthTextBox.text = mthTextBox.text & ">"
ElseIf struse = "55" Then mthTextBox.text = mthTextBox.text & "?"
ElseIf struse = "58" Then mthTextBox.text = mthTextBox.text & "/"
ElseIf struse = "57" Then mthTextBox.text = mthTextBox.text & "1"
ElseIf struse = "60" Then mthTextBox.text = mthTextBox.text & "2"
ElseIf struse = "59" Then mthTextBox.text = mthTextBox.text & "3"
ElseIf struse = "62" Then mthTextBox.text = mthTextBox.text & "4"
ElseIf struse = "61" Then mthTextBox.text = mthTextBox.text & "5"
ElseIf struse = "64" Then mthTextBox.text = mthTextBox.text & "6"
ElseIf struse = "63" Then mthTextBox.text = mthTextBox.text & "7"
ElseIf struse = "66" Then mthTextBox.text = mthTextBox.text & "8"
ElseIf struse = "65" Then mthTextBox.text = mthTextBox.text & "9"
ElseIf struse = "68" Then mthTextBox.text = mthTextBox.text & "0"
ElseIf struse = "67" Then mthTextBox.text = mthTextBox.text & vbCrLf
ElseIf struse = "70" Then mthTextBox.text = mthTextBox.text & "~"
ElseIf struse = "69" Then mthTextBox.text = mthTextBox.text & "`"
End If

'elseif struse = "" then mthTextBox.Text = mthTextBox.Text & ""
Next keypos
End If


If LCase(Method) = "c" Then
mthTextBox.text = ""
For keypos = 1 To Len(mthTextBox2.text) Step 2
struse = Mid(mthTextBox2.text, keypos, 2)
If struse = "70" Then
mthTextBox.text = mthTextBox.text & "A"
ElseIf struse = "69" Then mthTextBox.text = mthTextBox.text & "B"
ElseIf struse = "68" Then mthTextBox.text = mthTextBox.text & "C"
ElseIf struse = "67" Then mthTextBox.text = mthTextBox.text & "D"
ElseIf struse = "66" Then mthTextBox.text = mthTextBox.text & "E"
ElseIf struse = "65" Then mthTextBox.text = mthTextBox.text & "F"
ElseIf struse = "64" Then mthTextBox.text = mthTextBox.text & "G"
ElseIf struse = "63" Then mthTextBox.text = mthTextBox.text & "H"
ElseIf struse = "62" Then mthTextBox.text = mthTextBox.text & "I"
ElseIf struse = "61" Then mthTextBox.text = mthTextBox.text & "J"
ElseIf struse = "60" Then mthTextBox.text = mthTextBox.text & "K"
ElseIf struse = "59" Then mthTextBox.text = mthTextBox.text & "L"
ElseIf struse = "58" Then mthTextBox.text = mthTextBox.text & "M"
ElseIf struse = "57" Then mthTextBox.text = mthTextBox.text & "N"
ElseIf struse = "56" Then mthTextBox.text = mthTextBox.text & "O"
ElseIf struse = "55" Then mthTextBox.text = mthTextBox.text & "P"
ElseIf struse = "54" Then mthTextBox.text = mthTextBox.text & "Q"
ElseIf struse = "53" Then mthTextBox.text = mthTextBox.text & "R"
ElseIf struse = "52" Then mthTextBox.text = mthTextBox.text & "S"
ElseIf struse = "51" Then mthTextBox.text = mthTextBox.text & "T"
ElseIf struse = "50" Then mthTextBox.text = mthTextBox.text & "U"
ElseIf struse = "49" Then mthTextBox.text = mthTextBox.text & "V"
ElseIf struse = "48" Then mthTextBox.text = mthTextBox.text & "W"
ElseIf struse = "47" Then mthTextBox.text = mthTextBox.text & "X"
ElseIf struse = "46" Then mthTextBox.text = mthTextBox.text & "Y"
ElseIf struse = "45" Then mthTextBox.text = mthTextBox.text & "Z"
ElseIf struse = "44" Then mthTextBox.text = mthTextBox.text & " "
ElseIf struse = "43" Then mthTextBox.text = mthTextBox.text & "!"
ElseIf struse = "42" Then mthTextBox.text = mthTextBox.text & "@"
ElseIf struse = "41" Then mthTextBox.text = mthTextBox.text & "#"
ElseIf struse = "40" Then mthTextBox.text = mthTextBox.text & "$"
ElseIf struse = "39" Then mthTextBox.text = mthTextBox.text & "%"
ElseIf struse = "38" Then mthTextBox.text = mthTextBox.text & "^"
ElseIf struse = "37" Then mthTextBox.text = mthTextBox.text & "&"
ElseIf struse = "36" Then mthTextBox.text = mthTextBox.text & "*"
ElseIf struse = "35" Then mthTextBox.text = mthTextBox.text & "("
ElseIf struse = "34" Then mthTextBox.text = mthTextBox.text & ")"
ElseIf struse = "33" Then mthTextBox.text = mthTextBox.text & "-"
ElseIf struse = "32" Then mthTextBox.text = mthTextBox.text & "+"
ElseIf struse = "31" Then mthTextBox.text = mthTextBox.text & "\"
ElseIf struse = "30" Then mthTextBox.text = mthTextBox.text & "["
ElseIf struse = "29" Then mthTextBox.text = mthTextBox.text & "]"
ElseIf struse = "28" Then mthTextBox.text = mthTextBox.text & "_"
ElseIf struse = "27" Then mthTextBox.text = mthTextBox.text & "="
ElseIf struse = "26" Then mthTextBox.text = mthTextBox.text & "|"
ElseIf struse = "25" Then mthTextBox.text = mthTextBox.text & "{"
ElseIf struse = "24" Then mthTextBox.text = mthTextBox.text & "}"
ElseIf struse = "23" Then mthTextBox.text = mthTextBox.text & ":"
ElseIf struse = "22" Then mthTextBox.text = mthTextBox.text & ";"
ElseIf struse = "21" Then mthTextBox.text = mthTextBox.text & """"
ElseIf struse = "20" Then mthTextBox.text = mthTextBox.text & "'"
ElseIf struse = "19" Then mthTextBox.text = mthTextBox.text & "."
ElseIf struse = "18" Then mthTextBox.text = mthTextBox.text & ","
ElseIf struse = "17" Then mthTextBox.text = mthTextBox.text & "<"
ElseIf struse = "16" Then mthTextBox.text = mthTextBox.text & ">"
ElseIf struse = "15" Then mthTextBox.text = mthTextBox.text & "?"
ElseIf struse = "14" Then mthTextBox.text = mthTextBox.text & "/"
ElseIf struse = "13" Then mthTextBox.text = mthTextBox.text & "1"
ElseIf struse = "12" Then mthTextBox.text = mthTextBox.text & "2"
ElseIf struse = "11" Then mthTextBox.text = mthTextBox.text & "3"
ElseIf struse = "10" Then mthTextBox.text = mthTextBox.text & "4"
ElseIf struse = "09" Then mthTextBox.text = mthTextBox.text & "5"
ElseIf struse = "08" Then mthTextBox.text = mthTextBox.text & "6"
ElseIf struse = "07" Then mthTextBox.text = mthTextBox.text & "7"
ElseIf struse = "06" Then mthTextBox.text = mthTextBox.text & "8"
ElseIf struse = "05" Then mthTextBox.text = mthTextBox.text & "9"
ElseIf struse = "04" Then mthTextBox.text = mthTextBox.text & "0"
ElseIf struse = "03" Then mthTextBox.text = mthTextBox.text & vbCrLf
ElseIf struse = "02" Then mthTextBox.text = mthTextBox.text & "~"
ElseIf struse = "01" Then mthTextBox.text = mthTextBox.text & "`"
End If

'elseif struse = "" then mthTextBox.Text = mthTextBox.Text & ""
Next keypos
End If


If LCase(Method) = "d" Then
mthTextBox.text = ""
For keypos = 1 To Len(mthTextBox2.text) Step 2
struse = Mid(mthTextBox2.text, keypos, 2)
If struse = "02" Then
mthTextBox.text = mthTextBox.text & "A"
ElseIf struse = "04" Then mthTextBox.text = mthTextBox.text & "B"
ElseIf struse = "06" Then mthTextBox.text = mthTextBox.text & "C"
ElseIf struse = "08" Then mthTextBox.text = mthTextBox.text & "D"
ElseIf struse = "10" Then mthTextBox.text = mthTextBox.text & "E"
ElseIf struse = "12" Then mthTextBox.text = mthTextBox.text & "F"
ElseIf struse = "14" Then mthTextBox.text = mthTextBox.text & "G"
ElseIf struse = "16" Then mthTextBox.text = mthTextBox.text & "H"
ElseIf struse = "18" Then mthTextBox.text = mthTextBox.text & "I"
ElseIf struse = "20" Then mthTextBox.text = mthTextBox.text & "J"
ElseIf struse = "22" Then mthTextBox.text = mthTextBox.text & "K"
ElseIf struse = "24" Then mthTextBox.text = mthTextBox.text & "L"
ElseIf struse = "26" Then mthTextBox.text = mthTextBox.text & "M"
ElseIf struse = "28" Then mthTextBox.text = mthTextBox.text & "N"
ElseIf struse = "30" Then mthTextBox.text = mthTextBox.text & "O"
ElseIf struse = "32" Then mthTextBox.text = mthTextBox.text & "P"
ElseIf struse = "34" Then mthTextBox.text = mthTextBox.text & "Q"
ElseIf struse = "36" Then mthTextBox.text = mthTextBox.text & "R"
ElseIf struse = "38" Then mthTextBox.text = mthTextBox.text & "S"
ElseIf struse = "40" Then mthTextBox.text = mthTextBox.text & "T"
ElseIf struse = "42" Then mthTextBox.text = mthTextBox.text & "U"
ElseIf struse = "44" Then mthTextBox.text = mthTextBox.text & "V"
ElseIf struse = "46" Then mthTextBox.text = mthTextBox.text & "W"
ElseIf struse = "48" Then mthTextBox.text = mthTextBox.text & "X"
ElseIf struse = "50" Then mthTextBox.text = mthTextBox.text & "Y"
ElseIf struse = "52" Then mthTextBox.text = mthTextBox.text & "Z"
ElseIf struse = "54" Then mthTextBox.text = mthTextBox.text & " "
ElseIf struse = "56" Then mthTextBox.text = mthTextBox.text & "!"
ElseIf struse = "58" Then mthTextBox.text = mthTextBox.text & "@"
ElseIf struse = "60" Then mthTextBox.text = mthTextBox.text & "#"
ElseIf struse = "62" Then mthTextBox.text = mthTextBox.text & "$"
ElseIf struse = "64" Then mthTextBox.text = mthTextBox.text & "%"
ElseIf struse = "66" Then mthTextBox.text = mthTextBox.text & "^"
ElseIf struse = "68" Then mthTextBox.text = mthTextBox.text & "&"
ElseIf struse = "70" Then mthTextBox.text = mthTextBox.text & "*"
ElseIf struse = "01" Then mthTextBox.text = mthTextBox.text & "("
ElseIf struse = "03" Then mthTextBox.text = mthTextBox.text & ")"
ElseIf struse = "05" Then mthTextBox.text = mthTextBox.text & "-"
ElseIf struse = "07" Then mthTextBox.text = mthTextBox.text & "+"
ElseIf struse = "09" Then mthTextBox.text = mthTextBox.text & "\"
ElseIf struse = "11" Then mthTextBox.text = mthTextBox.text & "["
ElseIf struse = "13" Then mthTextBox.text = mthTextBox.text & "]"
ElseIf struse = "15" Then mthTextBox.text = mthTextBox.text & "_"
ElseIf struse = "17" Then mthTextBox.text = mthTextBox.text & "="
ElseIf struse = "19" Then mthTextBox.text = mthTextBox.text & "|"
ElseIf struse = "21" Then mthTextBox.text = mthTextBox.text & "{"
ElseIf struse = "23" Then mthTextBox.text = mthTextBox.text & "}"
ElseIf struse = "25" Then mthTextBox.text = mthTextBox.text & ":"
ElseIf struse = "27" Then mthTextBox.text = mthTextBox.text & ";"
ElseIf struse = "29" Then mthTextBox.text = mthTextBox.text & """"
ElseIf struse = "31" Then mthTextBox.text = mthTextBox.text & "'"
ElseIf struse = "33" Then mthTextBox.text = mthTextBox.text & "."
ElseIf struse = "35" Then mthTextBox.text = mthTextBox.text & ","
ElseIf struse = "37" Then mthTextBox.text = mthTextBox.text & "<"
ElseIf struse = "39" Then mthTextBox.text = mthTextBox.text & ">"
ElseIf struse = "41" Then mthTextBox.text = mthTextBox.text & "?"
ElseIf struse = "43" Then mthTextBox.text = mthTextBox.text & "/"
ElseIf struse = "45" Then mthTextBox.text = mthTextBox.text & "1"
ElseIf struse = "47" Then mthTextBox.text = mthTextBox.text & "2"
ElseIf struse = "49" Then mthTextBox.text = mthTextBox.text & "3"
ElseIf struse = "51" Then mthTextBox.text = mthTextBox.text & "4"
ElseIf struse = "53" Then mthTextBox.text = mthTextBox.text & "5"
ElseIf struse = "55" Then mthTextBox.text = mthTextBox.text & "6"
ElseIf struse = "57" Then mthTextBox.text = mthTextBox.text & "7"
ElseIf struse = "59" Then mthTextBox.text = mthTextBox.text & "8"
ElseIf struse = "61" Then mthTextBox.text = mthTextBox.text & "9"
ElseIf struse = "63" Then mthTextBox.text = mthTextBox.text & "0"
ElseIf struse = "65" Then mthTextBox.text = mthTextBox.text & vbCrLf
ElseIf struse = "67" Then mthTextBox.text = mthTextBox.text & "~"
ElseIf struse = "69" Then mthTextBox.text = mthTextBox.text & "`"
End If
Next keypos
End If


If LCase(Method) = "e" Then
mthTextBox.text = ""
For keypos = 1 To Len(mthTextBox2.text) Step 2
struse = Mid(mthTextBox2.text, keypos, 2)
If struse = "35" Then
mthTextBox.text = mthTextBox.text & "A"
ElseIf struse = "34" Then mthTextBox.text = mthTextBox.text & "B"
ElseIf struse = "33" Then mthTextBox.text = mthTextBox.text & "C"
ElseIf struse = "32" Then mthTextBox.text = mthTextBox.text & "D"
ElseIf struse = "31" Then mthTextBox.text = mthTextBox.text & "E"
ElseIf struse = "30" Then mthTextBox.text = mthTextBox.text & "F"
ElseIf struse = "29" Then mthTextBox.text = mthTextBox.text & "G"
ElseIf struse = "28" Then mthTextBox.text = mthTextBox.text & "H"
ElseIf struse = "27" Then mthTextBox.text = mthTextBox.text & "I"
ElseIf struse = "26" Then mthTextBox.text = mthTextBox.text & "J"
ElseIf struse = "25" Then mthTextBox.text = mthTextBox.text & "K"
ElseIf struse = "24" Then mthTextBox.text = mthTextBox.text & "L"
ElseIf struse = "23" Then mthTextBox.text = mthTextBox.text & "M"
ElseIf struse = "22" Then mthTextBox.text = mthTextBox.text & "N"
ElseIf struse = "21" Then mthTextBox.text = mthTextBox.text & "O"
ElseIf struse = "20" Then mthTextBox.text = mthTextBox.text & "P"
ElseIf struse = "19" Then mthTextBox.text = mthTextBox.text & "Q"
ElseIf struse = "18" Then mthTextBox.text = mthTextBox.text & "R"
ElseIf struse = "17" Then mthTextBox.text = mthTextBox.text & "S"
ElseIf struse = "16" Then mthTextBox.text = mthTextBox.text & "T"
ElseIf struse = "15" Then mthTextBox.text = mthTextBox.text & "U"
ElseIf struse = "14" Then mthTextBox.text = mthTextBox.text & "V"
ElseIf struse = "13" Then mthTextBox.text = mthTextBox.text & "W"
ElseIf struse = "12" Then mthTextBox.text = mthTextBox.text & "X"
ElseIf struse = "11" Then mthTextBox.text = mthTextBox.text & "Y"
ElseIf struse = "10" Then mthTextBox.text = mthTextBox.text & "Z"
ElseIf struse = "09" Then mthTextBox.text = mthTextBox.text & " "
ElseIf struse = "08" Then mthTextBox.text = mthTextBox.text & "!"
ElseIf struse = "07" Then mthTextBox.text = mthTextBox.text & "@"
ElseIf struse = "06" Then mthTextBox.text = mthTextBox.text & "#"
ElseIf struse = "05" Then mthTextBox.text = mthTextBox.text & "$"
ElseIf struse = "04" Then mthTextBox.text = mthTextBox.text & "%"
ElseIf struse = "03" Then mthTextBox.text = mthTextBox.text & "^"
ElseIf struse = "02" Then mthTextBox.text = mthTextBox.text & "&"
ElseIf struse = "01" Then mthTextBox.text = mthTextBox.text & "*"
ElseIf struse = "70" Then mthTextBox.text = mthTextBox.text & "("
ElseIf struse = "69" Then mthTextBox.text = mthTextBox.text & ")"
ElseIf struse = "68" Then mthTextBox.text = mthTextBox.text & "-"
ElseIf struse = "67" Then mthTextBox.text = mthTextBox.text & "+"
ElseIf struse = "66" Then mthTextBox.text = mthTextBox.text & "\"
ElseIf struse = "65" Then mthTextBox.text = mthTextBox.text & "["
ElseIf struse = "64" Then mthTextBox.text = mthTextBox.text & "]"
ElseIf struse = "63" Then mthTextBox.text = mthTextBox.text & "_"
ElseIf struse = "62" Then mthTextBox.text = mthTextBox.text & "="
ElseIf struse = "61" Then mthTextBox.text = mthTextBox.text & "|"
ElseIf struse = "60" Then mthTextBox.text = mthTextBox.text & "{"
ElseIf struse = "59" Then mthTextBox.text = mthTextBox.text & "}"
ElseIf struse = "58" Then mthTextBox.text = mthTextBox.text & ":"
ElseIf struse = "57" Then mthTextBox.text = mthTextBox.text & ";"
ElseIf struse = "56" Then mthTextBox.text = mthTextBox.text & """"
ElseIf struse = "55" Then mthTextBox.text = mthTextBox.text & "'"
ElseIf struse = "54" Then mthTextBox.text = mthTextBox.text & "."
ElseIf struse = "53" Then mthTextBox.text = mthTextBox.text & ","
ElseIf struse = "52" Then mthTextBox.text = mthTextBox.text & "<"
ElseIf struse = "51" Then mthTextBox.text = mthTextBox.text & ">"
ElseIf struse = "50" Then mthTextBox.text = mthTextBox.text & "?"
ElseIf struse = "49" Then mthTextBox.text = mthTextBox.text & "/"
ElseIf struse = "48" Then mthTextBox.text = mthTextBox.text & "1"
ElseIf struse = "47" Then mthTextBox.text = mthTextBox.text & "2"
ElseIf struse = "46" Then mthTextBox.text = mthTextBox.text & "3"
ElseIf struse = "45" Then mthTextBox.text = mthTextBox.text & "4"
ElseIf struse = "44" Then mthTextBox.text = mthTextBox.text & "5"
ElseIf struse = "43" Then mthTextBox.text = mthTextBox.text & "6"
ElseIf struse = "42" Then mthTextBox.text = mthTextBox.text & "7"
ElseIf struse = "41" Then mthTextBox.text = mthTextBox.text & "8"
ElseIf struse = "40" Then mthTextBox.text = mthTextBox.text & "9"
ElseIf struse = "39" Then mthTextBox.text = mthTextBox.text & "0"
ElseIf struse = "38" Then mthTextBox.text = mthTextBox.text & vbCrLf
ElseIf struse = "37" Then mthTextBox.text = mthTextBox.text & "~"
ElseIf struse = "36" Then mthTextBox.text = mthTextBox.text & "`"
End If
Next keypos
End If
End Sub
Sub ExportMethod(MethodName As String, mthComboBox As ComboBox)
If mthComboBox.text = MethodName Then
FileName$ = InputBox("Enter filename", "Export:[" & mthComboBox.text & "]")
If FileName$ = "" Then Exit Sub
If Right(App.Path, 1) = "\" Then Open App.Path & FileName$ & ".mth" For Binary As #1: GoTo nextstep1
Open App.Path & "\" & FileName$ & ".mth" For Binary As #1
nextstep1:
FileStuff.Interior = "[FN:" & MethodName & "/FN]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALG0[" & alg0 & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALG1[" & alg1 & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALG2[" & alg2 & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALG3[" & alg3 & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALG4[" & alg4 & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALG5[" & alg5 & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALG6[" & alg6 & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALG7[" & alg7 & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALG8[" & alg8 & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALG9[" & alg9 & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGA[" & algA & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGAMP[" & algAMP & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGAST[" & algAST & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGAT[" & algAT & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGB[" & algB & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGBKS[" & algBKS & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGC[" & algC & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGCL2[" & algCL2 & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGCOL[" & algCOL & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGCOM[" & algCOM & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGCP[" & algCP & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGD[" & algD & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGDSH[" & algDSH & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGE[" & algE & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGEXC[" & algEXC & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGF[" & algF & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGFWS[" & algFWS & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGG[" & algG & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGH[" & algH & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGI[" & algI & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGINT[" & algINT & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGJ[" & algJ & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGK[" & algK & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGL[" & algL & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGLB[" & algLB & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGLB2[" & algLB2 & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGLCH[" & algLCH & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGM[" & algM & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGN[" & algN & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGNUM[" & algNUM & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGO[" & algO & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGOP[" & algOP & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGP[" & algP & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGPER[" & algPER & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGPLS[" & algPLS & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGQ[" & algQ & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGQU2[" & algQU2 & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGQUE[" & algQUE & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGQUO[" & algQUO & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGR[" & algR & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGRB[" & algRB & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGRB2[" & algRB2 & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGRCH[" & algRCH & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGS[" & algS & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGSQU[" & algSQU & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGSQU2[" & algSQU2 & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGSTR[" & algSTR & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGT[" & algT & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGU[" & algU & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGUP[" & algUP & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGV[" & algV & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGW[" & algW & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGX[" & algX & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGY[" & algY & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGZ[" & algZ & "]"
FileStuff.Interior = FileStuff.Interior & "ALGSPC[" & algSPC & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGUNS[" & algUNS & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGBAR[" & algBAR & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGEQU[" & algEQU & "]" & vbCrLf
FileStuff.Interior = FileStuff.Interior & "ALGENT[" & algENT & "]" & vbCrLf
'filestuff.Interior = filestuff.Interior & "ALG[" & alg & "]" & vbcrlf
Put #1, 1, FileStuff.Interior
DoEvents
Close #1
End If
End Sub
Sub RefreshMethodChrValues(Method As String)

End Sub
Sub LoadMethodChrValues(mthComboBox As ComboBox, mthFileName As String)
If Right(App.Path, 1) = "\" Then Open App.Path & mthFileName & ".mth" For Binary As #1: GoTo nextme2
Open App.Path & "\" & mthFileName & ".mth" For Binary As #1
nextme2:

Get #1, 1, FileStuff.Interior

tmpfilename$ = Mid(FileStuff.Interior, 5, (InStr(FileStuff.Interior, "/FN:") + 1) - 8)
algA = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGA["), 2)
algB = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGB["), 2)
algC = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGC["), 2)
algD = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGD["), 2)
algE = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGE["), 2)
algF = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGF["), 2)
algG = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGG["), 2)
algH = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGH["), 2)
algI = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGI["), 2)
algJ = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGJ["), 2)
algK = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGK["), 2)
algL = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGL["), 2)
algM = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGM["), 2)
algN = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGN["), 2)
algO = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGO["), 2)
algP = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGP["), 2)
algQ = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGQ["), 2)
algR = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGR["), 2)
algS = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGS["), 2)
algT = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGT["), 2)
algU = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGU["), 2)
algV = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGV["), 2)
algW = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGW["), 2)
algX = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGX["), 2)
algY = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGY["), 2)
algZ = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGZ["), 2)
alg1 = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALG1["), 2)
alg2 = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALG2["), 2)
alg3 = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALG3["), 2)
alg4 = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALG4["), 2)
alg5 = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALG5["), 2)
alg6 = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALG6["), 2)
alg7 = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALG7["), 2)
alg8 = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALG8["), 2)
alg9 = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALG9["), 2)
alg0 = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALG0["), 2)
algPER = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGPER["), 2)
algLCH = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGLCH["), 2)
algCOM = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGCOM["), 2)
algRCH = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGRCH["), 2)
algEXC = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGEXC["), 2)
algAT = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGAT["), 2)
algNUM = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGNUM["), 2)
algSTR = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGSTR["), 2)
algINT = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGINT["), 2)
algUP = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGUP["), 2)
algAMP = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGAMP["), 2)
algAST = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGAST["), 2)
algOP = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGOP["), 2)
algCP = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGCP["), 2)
algDSH = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGDSH["), 2)
algPLS = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGPLS["), 2)
algBKS = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGBKS["), 2)
algFWS = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGFWS["), 2)
algCOL = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGCOL["), 2)
algCL2 = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGCL2["), 2)
algQUO = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGQUO["), 2)
algQU2 = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGQU2["), 2)
algSQU = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGSQU["), 2)
algSQU2 = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGSQU2["), 2)
algLB = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGLB["), 2)
algRB = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGRB["), 2)
algLB2 = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGLB2["), 2)
algRB2 = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGRB2["), 2)
algQUE = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGQUE["), 2)
algSPC = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGSPC["), 2)
algUNS = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGUNS["), 2)
algBAR = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGBAR["), 2)
algEQU = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGEQU["), 2)
algENT = Mid(FileStuff.Interior, InStr(FileStuff.Interior, "ALGENT["), 2)
Close #1
DoEvents
'mthComboBox.AddItem tmpfilename$
End Sub
