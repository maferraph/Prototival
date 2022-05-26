VERSION 5.00
Begin VB.Form Tela_Ciclos 
   AutoRedraw      =   -1  'True
   Caption         =   "Configuração dos ciclos"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9735
   ControlBox      =   0   'False
   Icon            =   "Tela_Ciclos.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7200
   ScaleWidth      =   9735
   WindowState     =   2  'Maximized
   Begin VB.Frame FR 
      BorderStyle     =   0  'None
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9375
      Begin VB.Frame Frame1 
         Caption         =   "Paradas programadas para plotar a assinatura da válvula:"
         Height          =   1935
         Left            =   0
         TabIndex        =   5
         Top             =   600
         Width           =   9255
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   6960
            TabIndex        =   13
            Text            =   "Text1"
            Top             =   1440
            Width           =   2175
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   4440
            TabIndex        =   11
            Text            =   "Text1"
            Top             =   1440
            Width           =   2175
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   1800
            TabIndex        =   9
            Text            =   "Text1"
            Top             =   1440
            Width           =   2175
         End
         Begin VB.Frame Frame2 
            Caption         =   "Ciclos escolhidos para assinatura:"
            Height          =   855
            Left            =   1800
            TabIndex        =   7
            Top             =   240
            Width           =   7335
            Begin VB.Label Label1 
               Caption         =   "Label1"
               Height          =   495
               Left            =   120
               TabIndex        =   8
               Top             =   240
               Width           =   7095
            End
         End
         Begin VB.ListBox LT_Parada 
            Height          =   1425
            ItemData        =   "Tela_Ciclos.frx":030A
            Left            =   120
            List            =   "Tela_Ciclos.frx":030C
            TabIndex        =   6
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Nº Assinaturas Alta Pressão:"
            Height          =   195
            Index           =   4
            Left            =   6960
            TabIndex        =   14
            Top             =   1200
            Width           =   2010
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Nº Assinaturas Média Pressão:"
            Height          =   195
            Index           =   3
            Left            =   4440
            TabIndex        =   12
            Top             =   1200
            Width           =   2175
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Nº Assinaturas Baixa Pressão:"
            Height          =   195
            Index           =   2
            Left            =   1800
            TabIndex        =   10
            Top             =   1200
            Width           =   2130
         End
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   4800
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   240
         Width           =   4455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   0
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   240
         Width           =   4455
      End
      Begin VB.Frame Frame3 
         Caption         =   "Ciclos programados para executar o teste de sobrecarga de torque:"
         Height          =   1935
         Left            =   0
         TabIndex        =   15
         Top             =   2520
         Width           =   9255
         Begin VB.ListBox List1 
            Height          =   1425
            ItemData        =   "Tela_Ciclos.frx":030E
            Left            =   120
            List            =   "Tela_Ciclos.frx":0310
            TabIndex        =   18
            Top             =   360
            Width           =   1575
         End
         Begin VB.Frame Frame4 
            Caption         =   "Ciclos escolhidos para sobrecarga:"
            Height          =   1575
            Left            =   1800
            TabIndex        =   16
            Top             =   240
            Width           =   7335
            Begin VB.Label Label2 
               Caption         =   "Label1"
               Height          =   1215
               Left            =   120
               TabIndex        =   17
               Top             =   240
               Width           =   7095
            End
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Tempos durante a ciclagem:"
         Height          =   2175
         Left            =   0
         TabIndex        =   19
         Top             =   4440
         Width           =   4695
         Begin VB.TextBox Text11 
            Height          =   285
            Left            =   2400
            TabIndex        =   30
            Text            =   "Text1"
            Top             =   1680
            Width           =   2175
         End
         Begin VB.TextBox Text10 
            Height          =   285
            Left            =   120
            TabIndex        =   29
            Text            =   "Text1"
            Top             =   1680
            Width           =   2175
         End
         Begin VB.TextBox Text9 
            Height          =   285
            Left            =   2400
            TabIndex        =   26
            Text            =   "Text1"
            Top             =   480
            Width           =   2175
         End
         Begin VB.TextBox Text8 
            Height          =   285
            Left            =   2400
            TabIndex        =   24
            Text            =   "Text1"
            Top             =   1080
            Width           =   2175
         End
         Begin VB.TextBox Text7 
            Height          =   285
            Left            =   120
            TabIndex        =   22
            Text            =   "Text1"
            Top             =   1080
            Width           =   2175
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   120
            TabIndex        =   20
            Text            =   "Text1"
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Tempo Alivio Carga (s):"
            Height          =   195
            Index           =   10
            Left            =   2400
            TabIndex        =   31
            Top             =   1440
            Width           =   1635
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Tempo Estanqueidade (s):"
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   28
            Top             =   1440
            Width           =   1860
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Tempo Torque V.Aberta (s):"
            Height          =   195
            Index           =   8
            Left            =   2400
            TabIndex        =   27
            Top             =   840
            Width           =   1965
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Tempo Torque V.Fechada (s):"
            Height          =   195
            Index           =   7
            Left            =   2400
            TabIndex        =   25
            Top             =   240
            Width           =   2130
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Dados durante assinatura (Hz):"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   23
            Top             =   840
            Width           =   2190
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Dados durante ciclagem (Hz):"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   2100
         End
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Número total de ciclos executados à temperatura extrema:"
         Height          =   195
         Index           =   1
         Left            =   4800
         TabIndex        =   4
         Top             =   0
         Width           =   4110
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Número total de ciclos executados à temperatura ambiente:"
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   4200
      End
   End
End
Attribute VB_Name = "Tela_Ciclos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BT_Fechar_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    FR.Top = 100
    FR.Left = (Screen.Width / 2) - (FR.Width / 2)
    VAR_Tela = "Tela_Ciclos"
End Sub


'*************************
' FUNCOES DESTA TELA
'*************************
Public Sub Fechar()
    Unload Tela_Ciclos
End Sub
Public Sub Salvar()
    VAR_PodeSalvar = True

End Sub
