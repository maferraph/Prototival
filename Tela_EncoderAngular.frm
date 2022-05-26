VERSION 5.00
Begin VB.Form Tela_EncoderAngular 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Encoder Angular"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   3510
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame9 
      Caption         =   "Tempo"
      Height          =   615
      Left            =   960
      TabIndex        =   17
      Top             =   1560
      Width           =   735
      Begin VB.Label LBTEMPO 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   585
      End
   End
   Begin VB.CommandButton BT 
      Caption         =   "ZERAR"
      Height          =   375
      Left            =   1800
      TabIndex        =   16
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Frame Frame5 
      Caption         =   "RPM"
      Height          =   615
      Left            =   120
      TabIndex        =   14
      Top             =   1560
      Width           =   735
      Begin VB.Label LBVEL 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   585
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Ângulo"
      Height          =   615
      Left            =   1800
      TabIndex        =   12
      Top             =   840
      Width           =   1575
      Begin VB.Label LBANGULO 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   585
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Pulso"
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   735
      Begin VB.Label LBPULSO 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   585
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Sentido"
      Height          =   615
      Left            =   960
      TabIndex        =   8
      Top             =   840
      Width           =   735
      Begin VB.Label LBSENTIDO 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   585
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Binv"
      Height          =   615
      Left            =   2640
      TabIndex        =   6
      Top             =   120
      Width           =   735
      Begin VB.Label LBBI 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   585
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Ainv"
      Height          =   615
      Left            =   1800
      TabIndex        =   4
      Top             =   120
      Width           =   735
      Begin VB.Label LBAI 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   585
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "B"
      Height          =   615
      Left            =   960
      TabIndex        =   2
      Top             =   120
      Width           =   735
      Begin VB.Label LBB 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   585
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "A"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
      Begin VB.Label LBA 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   585
      End
   End
   Begin VB.Timer TEMPO 
      Interval        =   1
      Left            =   3120
      Top             =   1440
   End
End
Attribute VB_Name = "Tela_EncoderAngular"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IVAR_TEMPO, IVAR_PORTA, IVAR_PULSO, IVAR_A, IVAR_B, IVAR_DA, IVAR_DB As Integer
Dim SVAR_SENTIDO, SVAR_SH, SVAR_SAH As String
Private Sub BT_Click()
    IVAR_A = 0
    IVAR_B = 0
    IVAR_AI = 0
    IVAR_BI = 0
    IVAR_DA = 0
    IVAR_DB = 0
    IVAR_DAI = 0
    IVAR_DBI = 0
    IVAR_PULSO = 0
    LBA.Caption = "0"
    LBB.Caption = "0"
    LBAI.Caption = "0"
    LBBI.Caption = "0"
    LBPULSO.Caption = "0"
    LBSENTIDO.Caption = "0"
    LBANGULO.Caption = "0"
    LBVEL.Caption = "0"
End Sub
Private Sub Form_Load()
    TEMPO.Interval = 1
    
    IVAR_TEMPO = 0
    SVAR_SH = "H"
    SVAR_SAH = "AH"
End Sub
Private Sub TEMPO_Timer()
    'A = LARANJA = VERMELHO = 1 bit
    'B = AMARELO = LARANJA = 2 bit
    IVAR_PORTA = LePorta(&H37C)
    If IVAR_PORTA = 1 Then
        IVAR_A = 1
        IVAR_B = 0
    ElseIf IVAR_PORTA = 2 Then
        IVAR_A = 0
        IVAR_B = 1
    ElseIf IVAR_PORTA = 3 Then
        IVAR_A = 1
        IVAR_B = 1
    ElseIf IVAR_PORTA = 4 Then
        IVAR_A = 0
        IVAR_B = 0
    End If
    'verifica se a condicao dos pinos mudou, portanto, teve pulso
    'condicoes abaixo na sequencia A-B-AI-BI -> Para direita = antihorário, para esquerda = horário
    '1100
    If IVAR_A = 0 And IVAR_B = 0 Then
        If IVAR_DA = 1 And IVAR_DB = 0 Then
            SVAR_SENTIDO = SVAR_SH
            IVAR_PULSO = IVAR_PULSO + 1
        ElseIf IVAR_DA = 0 And IVAR_DB = 1 Then
            SVAR_SENTIDO = SVAR_SAH
            IVAR_PULSO = IVAR_PULSO - 1
        End If
    ElseIf IVAR_A = 1 And IVAR_B = 0 Then
        If IVAR_DA = 1 And IVAR_DB = 1 Then
            SVAR_SENTIDO = SVAR_SH
            IVAR_PULSO = IVAR_PULSO + 1
        ElseIf IVAR_DA = 0 And IVAR_DB = 1 Then
            SVAR_SENTIDO = SVAR_SAH
            IVAR_PULSO = IVAR_PULSO - 1
        End If
    ElseIf IVAR_A = 0 And IVAR_B = 1 Then
        If IVAR_DA = 1 And IVAR_DB = 0 Then
            SVAR_SENTIDO = SVAR_SH
            IVAR_PULSO = IVAR_PULSO + 1
        ElseIf IVAR_DA = 1 And IVAR_DB = 1 Then
            SVAR_SENTIDO = SVAR_SAH
            IVAR_PULSO = IVAR_PULSO - 1
        End If
    ElseIf IVAR_A = 1 And IVAR_B = 1 Then
        If IVAR_DA = 0 And IVAR_DB = 1 Then
            SVAR_SENTIDO = SVAR_SH
            IVAR_PULSO = IVAR_PULSO + 1
        ElseIf IVAR_DA = 1 And IVAR_DB = 0 Then
            SVAR_SENTIDO = SVAR_SAH
            IVAR_PULSO = IVAR_PULSO - 1
        End If
    End If
    'confirma estado anterior
    IVAR_DA = IVAR_A
    IVAR_DB = IVAR_B
   
   
    If IVAR_PULSO = 0 Then
   
    ElseIf IVAR_PULSO = 1024 Then
   
    End If

    'verifica tempo para ver velocidade
    IVAR_TEMPO = IVAR_TEMPO + 1
    If IVAR_TEMPO = 1000 Then
        IVAR_TEMPO = 0
    End If
    'altera captions
    LBA.Caption = IVAR_A
    LBB.Caption = IVAR_B
    LBPULSO.Caption = IVAR_PULSO
    LBSENTIDO.Caption = SVAR_SENTIDO
    LBTEMPO.Caption = IVAR_TEMPO
End Sub
