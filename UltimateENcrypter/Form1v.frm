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
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   360
      TabIndex        =   3
      Top             =   1800
      Width           =   3855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   0
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   1320
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim blowfish As New clsBlowfish
Private Sub Command1_Click()
Form1.Label1.Caption = InputBox("ff")
Text2.Text = blowfish.EncryptString(Form1.Text1.Text, Form1.Label1.Caption, False)
End Sub

Private Sub Command2_Click()
Form1.Label1.Caption = InputBox("fidkgljlm")
Text2.Text = blowfish.DecryptString(Form1.Text1.Text, Form1.Label1.Caption, False)
End Sub
