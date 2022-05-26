VERSION 5.00
Begin VB.Form Tela_TesteS 
   Caption         =   "Form1"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   5205
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Muda Seletor"
      Height          =   495
      Left            =   2760
      TabIndex        =   6
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Text            =   "0"
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4200
      Top             =   240
   End
   Begin VB.CommandButton BT 
      Caption         =   "LIGA"
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Seletor"
      Height          =   195
      Index           =   2
      Left            =   960
      TabIndex        =   7
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Resultado"
      Height          =   195
      Index           =   1
      Left            =   1080
      TabIndex        =   4
      Top             =   1920
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Porta"
      Height          =   195
      Left            =   960
      TabIndex        =   3
      Top             =   240
      Width           =   375
   End
   Begin VB.Label LB 
      AutoSize        =   -1  'True
      Caption         =   "LB"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1080
      TabIndex        =   2
      Top             =   2280
      Width           =   435
   End
End
Attribute VB_Name = "Tela_TesteS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim P As Long
Private Sub BT_Click()
    If BT.Caption = "LIGA" Then
        BT.Caption = "DESLIGA"
        P = Val("&H" & Text1.Text)
        Timer1.Enabled = True
    Else
        BT.Caption = "LIGA"
        Timer1.Enabled = False
   End If
End Sub

Private Sub Command1_Click()
    MsgBox LePorta(&H37A)
End Sub

Private Sub Command2_Click()
    EscrevePorta &H37A, Text3.Text
End Sub

Private Sub Form_Load()
    EscrevePorta &H37A, 32
End Sub

Private Sub Form_Unload(Cancel As Integer)
    EscrevePorta &H37A, 0
    EscrevePorta P, 0
End Sub

Private Sub Timer1_Timer()
    'LB.Caption = LePorta(P)
        EscrevePorta &H37A, Text3.Text
        'Text3.Text = Text3.Text + 1
       ' MsgBox LePorta(&H37A)
        'If Int(Text3.Text) = 255 Then Timer1.Enabled = False
End Sub
