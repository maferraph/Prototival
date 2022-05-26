VERSION 5.00
Begin VB.Form Tela_Sensores 
   AutoRedraw      =   -1  'True
   Caption         =   "Configurações dos Sensores"
   ClientHeight    =   8850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9675
   ControlBox      =   0   'False
   Icon            =   "Tela_Sensores.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8850
   ScaleWidth      =   9675
   WindowState     =   2  'Maximized
   Begin VB.Frame FR 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7815
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   8775
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Left            =   3600
         TabIndex        =   112
         Top             =   6960
         Width           =   1095
      End
      Begin VB.PictureBox PB 
         AutoRedraw      =   -1  'True
         CausesValidation=   0   'False
         ClipControls    =   0   'False
         Height          =   1020
         Left            =   1800
         ScaleHeight     =   960
         ScaleWidth      =   1680
         TabIndex        =   111
         Top             =   6720
         Width           =   1740
      End
      Begin VB.Frame Frame1 
         Caption         =   "Configuração dos sensores da máquina de validação do protótipo:"
         Height          =   6135
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   8655
         Begin VB.Frame Frame2 
            Height          =   375
            Index           =   16
            Left            =   4080
            TabIndex        =   80
            Top             =   480
            Width           =   1335
            Begin VB.OptionButton RB_01 
               Caption         =   "1ª"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   82
               Top             =   120
               Width           =   495
            End
            Begin VB.OptionButton RB_01 
               Caption         =   "2ª"
               Height          =   195
               Index           =   1
               Left            =   720
               TabIndex        =   81
               Top             =   120
               Width           =   495
            End
         End
         Begin VB.Frame Frame2 
            Height          =   375
            Index           =   15
            Left            =   4080
            TabIndex        =   79
            Top             =   840
            Width           =   1335
            Begin VB.OptionButton RB_02 
               Caption         =   "1ª"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   84
               Top             =   120
               Width           =   495
            End
            Begin VB.OptionButton RB_02 
               Caption         =   "2ª"
               Height          =   195
               Index           =   1
               Left            =   720
               TabIndex        =   83
               Top             =   120
               Width           =   495
            End
         End
         Begin VB.Frame Frame2 
            Height          =   375
            Index           =   14
            Left            =   4080
            TabIndex        =   78
            Top             =   1200
            Width           =   1335
            Begin VB.OptionButton RB_03 
               Caption         =   "1ª"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   86
               Top             =   120
               Width           =   495
            End
            Begin VB.OptionButton RB_03 
               Caption         =   "2ª"
               Height          =   195
               Index           =   1
               Left            =   720
               TabIndex        =   85
               Top             =   120
               Width           =   495
            End
         End
         Begin VB.Frame Frame2 
            Height          =   375
            Index           =   13
            Left            =   4080
            TabIndex        =   77
            Top             =   1560
            Width           =   1335
            Begin VB.OptionButton RB_04 
               Caption         =   "1ª"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   88
               Top             =   120
               Width           =   495
            End
            Begin VB.OptionButton RB_04 
               Caption         =   "2ª"
               Height          =   195
               Index           =   1
               Left            =   720
               TabIndex        =   87
               Top             =   120
               Width           =   495
            End
         End
         Begin VB.Frame Frame2 
            Height          =   375
            Index           =   12
            Left            =   4080
            TabIndex        =   76
            Top             =   1920
            Width           =   1335
            Begin VB.OptionButton RB_05 
               Caption         =   "1ª"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   90
               Top             =   120
               Width           =   495
            End
            Begin VB.OptionButton RB_05 
               Caption         =   "2ª"
               Height          =   195
               Index           =   1
               Left            =   720
               TabIndex        =   89
               Top             =   120
               Width           =   495
            End
         End
         Begin VB.Frame Frame2 
            Height          =   375
            Index           =   11
            Left            =   4080
            TabIndex        =   75
            Top             =   2280
            Width           =   1335
            Begin VB.OptionButton RB_06 
               Caption         =   "1ª"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   92
               Top             =   120
               Width           =   495
            End
            Begin VB.OptionButton RB_06 
               Caption         =   "2ª"
               Height          =   195
               Index           =   1
               Left            =   720
               TabIndex        =   91
               Top             =   120
               Width           =   495
            End
         End
         Begin VB.Frame Frame2 
            Height          =   375
            Index           =   10
            Left            =   4080
            TabIndex        =   74
            Top             =   2640
            Width           =   1335
            Begin VB.OptionButton RB_07 
               Caption         =   "1ª"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   94
               Top             =   120
               Width           =   495
            End
            Begin VB.OptionButton RB_07 
               Caption         =   "2ª"
               Height          =   195
               Index           =   1
               Left            =   720
               TabIndex        =   93
               Top             =   120
               Width           =   495
            End
         End
         Begin VB.Frame Frame2 
            Height          =   375
            Index           =   9
            Left            =   4080
            TabIndex        =   73
            Top             =   3000
            Width           =   1335
            Begin VB.OptionButton RB_08 
               Caption         =   "1ª"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   96
               Top             =   120
               Width           =   495
            End
            Begin VB.OptionButton RB_08 
               Caption         =   "2ª"
               Height          =   195
               Index           =   1
               Left            =   720
               TabIndex        =   95
               Top             =   120
               Width           =   495
            End
         End
         Begin VB.Frame Frame2 
            Height          =   375
            Index           =   8
            Left            =   4080
            TabIndex        =   72
            Top             =   3360
            Width           =   1335
            Begin VB.OptionButton RB_09 
               Caption         =   "1ª"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   98
               Top             =   120
               Width           =   495
            End
            Begin VB.OptionButton RB_09 
               Caption         =   "2ª"
               Height          =   195
               Index           =   1
               Left            =   720
               TabIndex        =   97
               Top             =   120
               Width           =   495
            End
         End
         Begin VB.Frame Frame2 
            Height          =   375
            Index           =   7
            Left            =   4080
            TabIndex        =   71
            Top             =   3720
            Width           =   1335
            Begin VB.OptionButton RB_10 
               Caption         =   "1ª"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   100
               Top             =   120
               Width           =   495
            End
            Begin VB.OptionButton RB_10 
               Caption         =   "2ª"
               Height          =   195
               Index           =   1
               Left            =   720
               TabIndex        =   99
               Top             =   120
               Width           =   495
            End
         End
         Begin VB.Frame Frame2 
            Height          =   375
            Index           =   6
            Left            =   4080
            TabIndex        =   70
            Top             =   4080
            Width           =   1335
            Begin VB.OptionButton RB_11 
               Caption         =   "1ª"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   102
               Top             =   120
               Width           =   495
            End
            Begin VB.OptionButton RB_11 
               Caption         =   "2ª"
               Height          =   195
               Index           =   1
               Left            =   720
               TabIndex        =   101
               Top             =   120
               Width           =   495
            End
         End
         Begin VB.Frame Frame2 
            Height          =   375
            Index           =   5
            Left            =   4080
            TabIndex        =   69
            Top             =   4440
            Width           =   1335
            Begin VB.OptionButton RB_12 
               Caption         =   "1ª"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   104
               Top             =   120
               Width           =   495
            End
            Begin VB.OptionButton RB_12 
               Caption         =   "2ª"
               Height          =   195
               Index           =   1
               Left            =   720
               TabIndex        =   103
               Top             =   120
               Width           =   495
            End
         End
         Begin VB.Frame Frame2 
            Height          =   375
            Index           =   4
            Left            =   4080
            TabIndex        =   68
            Top             =   4800
            Width           =   1335
            Begin VB.OptionButton RB_13 
               Caption         =   "1ª"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   106
               Top             =   120
               Width           =   495
            End
            Begin VB.OptionButton RB_13 
               Caption         =   "2ª"
               Height          =   195
               Index           =   1
               Left            =   720
               TabIndex        =   105
               Top             =   120
               Width           =   495
            End
         End
         Begin VB.Frame Frame2 
            Height          =   375
            Index           =   3
            Left            =   4080
            TabIndex        =   67
            Top             =   5160
            Width           =   1335
            Begin VB.OptionButton RB_14 
               Caption         =   "1ª"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   108
               Top             =   120
               Width           =   495
            End
            Begin VB.OptionButton RB_14 
               Caption         =   "2ª"
               Height          =   195
               Index           =   1
               Left            =   720
               TabIndex        =   107
               Top             =   120
               Width           =   495
            End
         End
         Begin VB.Frame Frame2 
            Height          =   375
            Index           =   1
            Left            =   4080
            TabIndex        =   66
            Top             =   5520
            Width           =   1335
            Begin VB.OptionButton RB_15 
               Caption         =   "1ª"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   110
               Top             =   120
               Width           =   495
            End
            Begin VB.OptionButton RB_15 
               Caption         =   "2ª"
               Height          =   195
               Index           =   1
               Left            =   720
               TabIndex        =   109
               Top             =   120
               Width           =   495
            End
         End
         Begin VB.TextBox TXT_MAX15 
            Height          =   285
            Left            =   7080
            TabIndex        =   64
            Top             =   5640
            Width           =   1455
         End
         Begin VB.TextBox TXT_MIN15 
            Height          =   285
            Left            =   5520
            TabIndex        =   63
            Top             =   5640
            Width           =   1455
         End
         Begin VB.TextBox TXT_LPT15 
            Height          =   285
            Left            =   2520
            TabIndex        =   62
            Top             =   5640
            Width           =   1455
         End
         Begin VB.TextBox TXT_MAX14 
            Height          =   285
            Left            =   7080
            TabIndex        =   60
            Top             =   5280
            Width           =   1455
         End
         Begin VB.TextBox TXT_MIN14 
            Height          =   285
            Left            =   5520
            TabIndex        =   59
            Top             =   5280
            Width           =   1455
         End
         Begin VB.TextBox TXT_LPT14 
            Height          =   285
            Left            =   2520
            TabIndex        =   58
            Top             =   5280
            Width           =   1455
         End
         Begin VB.TextBox TXT_MAX13 
            Height          =   285
            Left            =   7080
            TabIndex        =   56
            Top             =   4920
            Width           =   1455
         End
         Begin VB.TextBox TXT_MIN13 
            Height          =   285
            Left            =   5520
            TabIndex        =   55
            Top             =   4920
            Width           =   1455
         End
         Begin VB.TextBox TXT_LPT13 
            Height          =   285
            Left            =   2520
            TabIndex        =   54
            Top             =   4920
            Width           =   1455
         End
         Begin VB.TextBox TXT_MAX12 
            Height          =   285
            Left            =   7080
            TabIndex        =   52
            Top             =   4560
            Width           =   1455
         End
         Begin VB.TextBox TXT_MIN12 
            Height          =   285
            Left            =   5520
            TabIndex        =   51
            Top             =   4560
            Width           =   1455
         End
         Begin VB.TextBox TXT_LPT12 
            Height          =   285
            Left            =   2520
            TabIndex        =   50
            Top             =   4560
            Width           =   1455
         End
         Begin VB.TextBox TXT_MAX11 
            Height          =   285
            Left            =   7080
            TabIndex        =   48
            Top             =   4200
            Width           =   1455
         End
         Begin VB.TextBox TXT_MIN11 
            Height          =   285
            Left            =   5520
            TabIndex        =   47
            Top             =   4200
            Width           =   1455
         End
         Begin VB.TextBox TXT_LPT11 
            Height          =   285
            Left            =   2520
            TabIndex        =   46
            Top             =   4200
            Width           =   1455
         End
         Begin VB.TextBox TXT_MAX10 
            Height          =   285
            Left            =   7080
            TabIndex        =   44
            Top             =   3840
            Width           =   1455
         End
         Begin VB.TextBox TXT_MIN10 
            Height          =   285
            Left            =   5520
            TabIndex        =   43
            Top             =   3840
            Width           =   1455
         End
         Begin VB.TextBox TXT_LPT10 
            Height          =   285
            Left            =   2520
            TabIndex        =   42
            Top             =   3840
            Width           =   1455
         End
         Begin VB.TextBox TXT_MAX09 
            Height          =   285
            Left            =   7080
            TabIndex        =   40
            Top             =   3480
            Width           =   1455
         End
         Begin VB.TextBox TXT_MIN09 
            Height          =   285
            Left            =   5520
            TabIndex        =   39
            Top             =   3480
            Width           =   1455
         End
         Begin VB.TextBox TXT_LPT09 
            Height          =   285
            Left            =   2520
            TabIndex        =   38
            Top             =   3480
            Width           =   1455
         End
         Begin VB.TextBox TXT_MAX08 
            Height          =   285
            Left            =   7080
            TabIndex        =   36
            Top             =   3120
            Width           =   1455
         End
         Begin VB.TextBox TXT_MIN08 
            Height          =   285
            Left            =   5520
            TabIndex        =   35
            Top             =   3120
            Width           =   1455
         End
         Begin VB.TextBox TXT_LPT08 
            Height          =   285
            Left            =   2520
            TabIndex        =   34
            Top             =   3120
            Width           =   1455
         End
         Begin VB.TextBox TXT_MAX07 
            Height          =   285
            Left            =   7080
            TabIndex        =   32
            Top             =   2760
            Width           =   1455
         End
         Begin VB.TextBox TXT_MIN07 
            Height          =   285
            Left            =   5520
            TabIndex        =   31
            Top             =   2760
            Width           =   1455
         End
         Begin VB.TextBox TXT_LPT07 
            Height          =   285
            Left            =   2520
            TabIndex        =   30
            Top             =   2760
            Width           =   1455
         End
         Begin VB.TextBox TXT_MAX06 
            Height          =   285
            Left            =   7080
            TabIndex        =   28
            Top             =   2400
            Width           =   1455
         End
         Begin VB.TextBox TXT_MIN06 
            Height          =   285
            Left            =   5520
            TabIndex        =   27
            Top             =   2400
            Width           =   1455
         End
         Begin VB.TextBox TXT_LPT06 
            Height          =   285
            Left            =   2520
            TabIndex        =   26
            Top             =   2400
            Width           =   1455
         End
         Begin VB.TextBox TXT_MAX05 
            Height          =   285
            Left            =   7080
            TabIndex        =   24
            Top             =   2040
            Width           =   1455
         End
         Begin VB.TextBox TXT_MIN05 
            Height          =   285
            Left            =   5520
            TabIndex        =   23
            Top             =   2040
            Width           =   1455
         End
         Begin VB.TextBox TXT_LPT05 
            Height          =   285
            Left            =   2520
            TabIndex        =   22
            Top             =   2040
            Width           =   1455
         End
         Begin VB.TextBox TXT_MAX04 
            Height          =   285
            Left            =   7080
            TabIndex        =   20
            Top             =   1680
            Width           =   1455
         End
         Begin VB.TextBox TXT_MIN04 
            Height          =   285
            Left            =   5520
            TabIndex        =   19
            Top             =   1680
            Width           =   1455
         End
         Begin VB.TextBox TXT_LPT04 
            Height          =   285
            Left            =   2520
            TabIndex        =   18
            Top             =   1680
            Width           =   1455
         End
         Begin VB.TextBox TXT_MAX03 
            Height          =   285
            Left            =   7080
            TabIndex        =   16
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox TXT_MIN03 
            Height          =   285
            Left            =   5520
            TabIndex        =   15
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox TXT_LPT03 
            Height          =   285
            Left            =   2520
            TabIndex        =   14
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox TXT_MAX02 
            Height          =   285
            Left            =   7080
            TabIndex        =   12
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox TXT_MIN02 
            Height          =   285
            Left            =   5520
            TabIndex        =   11
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox TXT_LPT02 
            Height          =   285
            Left            =   2520
            TabIndex        =   10
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox TXT_MAX01 
            Height          =   285
            Left            =   7080
            TabIndex        =   7
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox TXT_MIN01 
            Height          =   285
            Left            =   5520
            TabIndex        =   5
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox TXT_LPT01 
            Height          =   285
            Left            =   2520
            TabIndex        =   2
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "15) Vazamento na Gaxeta"
            Height          =   195
            Index           =   18
            Left            =   120
            TabIndex        =   65
            Top             =   5640
            Width           =   1845
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "14) Vazamento na Junta"
            Height          =   195
            Index           =   17
            Left            =   120
            TabIndex        =   61
            Top             =   5280
            Width           =   1725
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "13) Vazamento na Passagem"
            Height          =   195
            Index           =   16
            Left            =   120
            TabIndex        =   57
            Top             =   4920
            Width           =   2070
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "12) Vazão"
            Height          =   195
            Index           =   15
            Left            =   120
            TabIndex        =   53
            Top             =   4560
            Width           =   720
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "11) Velocidade de Acionamento"
            Height          =   195
            Index           =   14
            Left            =   120
            TabIndex        =   49
            Top             =   4200
            Width           =   2265
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "10) Vibração"
            Height          =   195
            Index           =   13
            Left            =   120
            TabIndex        =   45
            Top             =   3840
            Width           =   900
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "9) Ruído"
            Height          =   195
            Index           =   12
            Left            =   120
            TabIndex        =   41
            Top             =   3480
            Width           =   630
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "8) Deformação (Straingage)"
            Height          =   195
            Index           =   11
            Left            =   120
            TabIndex        =   37
            Top             =   3120
            Width           =   1950
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "7) Deslocamento Angular"
            Height          =   195
            Index           =   10
            Left            =   120
            TabIndex        =   33
            Top             =   2760
            Width           =   1785
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "6) Deslocamento Linea"
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   29
            Top             =   2400
            Width           =   1635
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "5) Temperatura do Fluído"
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   25
            Top             =   2040
            Width           =   1800
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "4) Torque de Acionamento"
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   21
            Top             =   1680
            Width           =   1890
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "3) Pressão na Jusante"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   17
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "2) Pressão no Corpo"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   13
            Top             =   960
            Width           =   2295
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "1) Pressão na Montante"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   9
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Valor Máximo:"
            Height          =   195
            Index           =   3
            Left            =   7080
            TabIndex        =   8
            Top             =   360
            Width           =   990
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Valor Mínimo:"
            Height          =   195
            Index           =   2
            Left            =   5520
            TabIndex        =   6
            Top             =   360
            Width           =   975
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Número porta LPT:"
            Height          =   195
            Index           =   0
            Left            =   2520
            TabIndex        =   3
            Top             =   360
            Width           =   1350
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Sequência de bits:"
            Height          =   195
            Index           =   1
            Left            =   4080
            TabIndex        =   4
            Top             =   240
            Width           =   1320
         End
      End
   End
End
Attribute VB_Name = "Tela_Sensores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'    PB.Line (500, 500)-(2000, 2000), RGB(0, 0, 255)
'    VB.SavePicture PB.Image, App.Path & "\teste.jpg"
 Tela_Simulacao.Show
End Sub

Private Sub Form_Load()
    FR.Top = 100
    FR.Left = (Screen.Width / 2) - (FR.Width / 2)
    VAR_Salvar = False
    VAR_Tela = "Tela_Sensores"
    'Pega Valores de Configuracao do Arquivo
    'Sensor 1
    TXT_LPT01.Text = LeINI(VAR_ArquivoINI, "Sensor1", "LPT")
    If LeINI(VAR_ArquivoINI, "Sensor1", "Bits") = "1" Then
        RB_01(0).Value = True
    ElseIf LeINI(VAR_ArquivoINI, "Sensor1", "Bits") = "2" Then
        RB_01(1).Value = True
    End If
    TXT_MIN01.Text = LeINI(VAR_ArquivoINI, "Sensor1", "ValMin")
    TXT_MAX01.Text = LeINI(VAR_ArquivoINI, "Sensor1", "ValMax")
    'Sensor 2
    TXT_LPT02.Text = LeINI(VAR_ArquivoINI, "Sensor2", "LPT")
    If LeINI(VAR_ArquivoINI, "Sensor2", "Bits") = "1" Then
        RB_02(0).Value = True
    ElseIf LeINI(VAR_ArquivoINI, "Sensor2", "Bits") = "2" Then
        RB_02(1).Value = True
    End If
    TXT_MIN02.Text = LeINI(VAR_ArquivoINI, "Sensor2", "ValMin")
    TXT_MAX02.Text = LeINI(VAR_ArquivoINI, "Sensor2", "ValMax")
    'Sensor 3
    TXT_LPT03.Text = LeINI(VAR_ArquivoINI, "Sensor3", "LPT")
    If LeINI(VAR_ArquivoINI, "Sensor3", "Bits") = "1" Then
        RB_03(0).Value = True
    ElseIf LeINI(VAR_ArquivoINI, "Sensor3", "Bits") = "2" Then
        RB_03(1).Value = True
    End If
    TXT_MIN03.Text = LeINI(VAR_ArquivoINI, "Sensor3", "ValMin")
    TXT_MAX03.Text = LeINI(VAR_ArquivoINI, "Sensor3", "ValMax")
    'Sensor 4
    TXT_LPT04.Text = LeINI(VAR_ArquivoINI, "Sensor4", "LPT")
    If LeINI(VAR_ArquivoINI, "Sensor4", "Bits") = "1" Then
        RB_04(0).Value = True
    ElseIf LeINI(VAR_ArquivoINI, "Sensor4", "Bits") = "2" Then
        RB_04(1).Value = True
    End If
    TXT_MIN04.Text = LeINI(VAR_ArquivoINI, "Sensor4", "ValMin")
    TXT_MAX04.Text = LeINI(VAR_ArquivoINI, "Sensor4", "ValMax")
    'Sensor 5
    TXT_LPT05.Text = LeINI(VAR_ArquivoINI, "Sensor5", "LPT")
    If LeINI(VAR_ArquivoINI, "Sensor4", "Bits") = "1" Then
        RB_05(0).Value = True
    ElseIf LeINI(VAR_ArquivoINI, "Sensor4", "Bits") = "2" Then
        RB_05(1).Value = True
    End If
    TXT_MIN05.Text = LeINI(VAR_ArquivoINI, "Sensor5", "ValMin")
    TXT_MAX05.Text = LeINI(VAR_ArquivoINI, "Sensor5", "ValMax")
     'Sensor 6
    TXT_LPT06.Text = LeINI(VAR_ArquivoINI, "Sensor6", "LPT")
    If LeINI(VAR_ArquivoINI, "Sensor6", "Bits") = "1" Then
        RB_06(0).Value = True
    ElseIf LeINI(VAR_ArquivoINI, "Sensor6", "Bits") = "2" Then
        RB_06(1).Value = True
    End If
    TXT_MIN06.Text = LeINI(VAR_ArquivoINI, "Sensor6", "ValMin")
    TXT_MAX06.Text = LeINI(VAR_ArquivoINI, "Sensor6", "ValMax")
    'Sensor 7
    TXT_LPT07.Text = LeINI(VAR_ArquivoINI, "Sensor7", "LPT")
    If LeINI(VAR_ArquivoINI, "Sensor7", "Bits") = "1" Then
        RB_07(0).Value = True
    ElseIf LeINI(VAR_ArquivoINI, "Sensor7", "Bits") = "2" Then
        RB_07(1).Value = True
    End If
    TXT_MIN07.Text = LeINI(VAR_ArquivoINI, "Sensor7", "ValMin")
    TXT_MAX07.Text = LeINI(VAR_ArquivoINI, "Sensor7", "ValMax")
    'Sensor 8
    TXT_LPT08.Text = LeINI(VAR_ArquivoINI, "Sensor8", "LPT")
    If LeINI(VAR_ArquivoINI, "Sensor8", "Bits") = "1" Then
        RB_08(0).Value = True
    ElseIf LeINI(VAR_ArquivoINI, "Sensor8", "Bits") = "2" Then
        RB_08(1).Value = True
    End If
    TXT_MIN08.Text = LeINI(VAR_ArquivoINI, "Sensor8", "ValMin")
    TXT_MAX08.Text = LeINI(VAR_ArquivoINI, "Sensor8", "ValMax")
     'Sensor 9
    TXT_LPT09.Text = LeINI(VAR_ArquivoINI, "Sensor9", "LPT")
    If LeINI(VAR_ArquivoINI, "Sensor9", "Bits") = "1" Then
        RB_09(0).Value = True
    ElseIf LeINI(VAR_ArquivoINI, "Sensor9", "Bits") = "2" Then
        RB_09(1).Value = True
    End If
    TXT_MIN09.Text = LeINI(VAR_ArquivoINI, "Sensor9", "ValMin")
    TXT_MAX09.Text = LeINI(VAR_ArquivoINI, "Sensor9", "ValMax")
    'Sensor 10
    TXT_LPT10.Text = LeINI(VAR_ArquivoINI, "Sensor10", "LPT")
    If LeINI(VAR_ArquivoINI, "Sensor10", "Bits") = "1" Then
        RB_10(0).Value = True
    ElseIf LeINI(VAR_ArquivoINI, "Sensor10", "Bits") = "2" Then
        RB_10(1).Value = True
    End If
    TXT_MIN10.Text = LeINI(VAR_ArquivoINI, "Sensor10", "ValMin")
    TXT_MAX10.Text = LeINI(VAR_ArquivoINI, "Sensor10", "ValMax")
    'Sensor 11
    TXT_LPT11.Text = LeINI(VAR_ArquivoINI, "Sensor11", "LPT")
    If LeINI(VAR_ArquivoINI, "Sensor11", "Bits") = "1" Then
        RB_11(0).Value = True
    ElseIf LeINI(VAR_ArquivoINI, "Sensor11", "Bits") = "2" Then
        RB_11(1).Value = True
    End If
    TXT_MIN11.Text = LeINI(VAR_ArquivoINI, "Sensor11", "ValMin")
    TXT_MAX11.Text = LeINI(VAR_ArquivoINI, "Sensor11", "ValMax")
    'Sensor 12
    TXT_LPT12.Text = LeINI(VAR_ArquivoINI, "Sensor12", "LPT")
    If LeINI(VAR_ArquivoINI, "Sensor12", "Bits") = "1" Then
        RB_12(0).Value = True
    ElseIf LeINI(VAR_ArquivoINI, "Sensor12", "Bits") = "2" Then
        RB_12(1).Value = True
    End If
    TXT_MIN12.Text = LeINI(VAR_ArquivoINI, "Sensor12", "ValMin")
    TXT_MAX12.Text = LeINI(VAR_ArquivoINI, "Sensor12", "ValMax")
    'Sensor 13
    TXT_LPT13.Text = LeINI(VAR_ArquivoINI, "Sensor13", "LPT")
    If LeINI(VAR_ArquivoINI, "Sensor13", "Bits") = "1" Then
        RB_13(0).Value = True
    ElseIf LeINI(VAR_ArquivoINI, "Sensor13", "Bits") = "2" Then
        RB_13(1).Value = True
    End If
    TXT_MIN13.Text = LeINI(VAR_ArquivoINI, "Sensor13", "ValMin")
    TXT_MAX13.Text = LeINI(VAR_ArquivoINI, "Sensor13", "ValMax")
    'Sensor 10
    TXT_LPT14.Text = LeINI(VAR_ArquivoINI, "Sensor14", "LPT")
    If LeINI(VAR_ArquivoINI, "Sensor14", "Bits") = "1" Then
        RB_14(0).Value = True
    ElseIf LeINI(VAR_ArquivoINI, "Sensor14", "Bits") = "2" Then
        RB_14(1).Value = True
    End If
    TXT_MIN14.Text = LeINI(VAR_ArquivoINI, "Sensor14", "ValMin")
    TXT_MAX14.Text = LeINI(VAR_ArquivoINI, "Sensor14", "ValMax")
    'Sensor 150
    TXT_LPT15.Text = LeINI(VAR_ArquivoINI, "Sensor15", "LPT")
    If LeINI(VAR_ArquivoINI, "Sensor15", "Bits") = "1" Then
        RB_15(0).Value = True
    ElseIf LeINI(VAR_ArquivoINI, "Sensor15", "Bits") = "2" Then
        RB_15(1).Value = True
    End If
    TXT_MIN15.Text = LeINI(VAR_ArquivoINI, "Sensor15", "ValMin")
    TXT_MAX15.Text = LeINI(VAR_ArquivoINI, "Sensor15", "ValMax")
    'desabilita botão salvar
    HabilitaSalvarDados False
End Sub
Private Sub RB_01_Click(Index As Integer)
    HabilitaSalvarDados True
End Sub

Private Sub RB_02_Click(Index As Integer)
    HabilitaSalvarDados True
End Sub

Private Sub RB_03_Click(Index As Integer)
    HabilitaSalvarDados True
End Sub

Private Sub RB_04_Click(Index As Integer)
    HabilitaSalvarDados True
End Sub

Private Sub RB_05_Click(Index As Integer)
    HabilitaSalvarDados True
End Sub

Private Sub RB_06_Click(Index As Integer)
    HabilitaSalvarDados True
End Sub

Private Sub RB_07_Click(Index As Integer)
    HabilitaSalvarDados True
End Sub

Private Sub RB_08_Click(Index As Integer)
    HabilitaSalvarDados True
End Sub

Private Sub RB_09_Click(Index As Integer)
    HabilitaSalvarDados True
End Sub

Private Sub RB_10_Click(Index As Integer)
    HabilitaSalvarDados True
End Sub

Private Sub RB_11_Click(Index As Integer)
    HabilitaSalvarDados True
End Sub

Private Sub RB_12_Click(Index As Integer)
    HabilitaSalvarDados True
End Sub

Private Sub RB_13_Click(Index As Integer)
    HabilitaSalvarDados True
End Sub

Private Sub RB_14_Click(Index As Integer)
    HabilitaSalvarDados True
End Sub

Private Sub RB_15_Click(Index As Integer)
    HabilitaSalvarDados True
End Sub

Private Sub TXT_LPT01_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_LPT01_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_LPT02_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_LPT02_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_LPT03_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_LPT03_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_LPT04_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_LPT04_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_LPT05_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_LPT05_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_LPT06_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_LPT06_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_LPT07_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_LPT07_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_LPT08_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_LPT08_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_LPT09_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_LPT09_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_LPT10_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_LPT10_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_LPT11_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_LPT11_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_LPT12_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_LPT12_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_LPT13_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_LPT13_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_LPT14_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_LPT14_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_LPT15_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_LPT15_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_MAX01_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_MAX01_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_MAX02_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_MAX02_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_MAX03_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_MAX03_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_MAX04_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_MAX04_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_MAX05_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_MAX05_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_MAX06_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_MAX06_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_MAX07_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_MAX07_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_MAX08_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_MAX08_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_MAX09_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_MAX09_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_MAX10_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_MAX10_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_MAX11_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_MAX11_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_MAX12_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_MAX12_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_MAX13_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_MAX13_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_MAX14_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_MAX14_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_MAX15_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_MAX15_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_MIN01_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_MIN01_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_MIN02_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_MIN02_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_MIN03_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_MIN03_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_MIN04_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_MIN04_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_MIN05_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_MIN05_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_MIN06_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_MIN06_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_MIN07_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_MIN07_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_MIN08_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_MIN08_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_MIN09_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_MIN09_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_MIN10_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_MIN10_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_MIN11_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_MIN11_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_MIN12_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_MIN12_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_MIN13_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_MIN13_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_MIN14_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_MIN14_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_MIN15_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_MIN15_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub


'*************************
' FUNCOES DESTA TELA
'*************************
Public Sub Fechar()
    If VAR_Salvar = True Then
        VAR_RespMsg = MsgBox("Você alterou os dados de configuração dos sensores e não salvou as informações. Tem certeza que deseja sair sem gravar as modificações?", vbInformation + vbYesNo, "Sair sem salvar")
        If VAR_RespMsg = vbNo Then
            Salvar
            Exit Sub
        End If
    End If
    Unload Tela_Sensores
End Sub
Public Sub Salvar()
    Dim VAR_MSGBOX_LPT, VAR_MSGBOX_RB, VAR_MSGBOX_VALMIN, VAR_MSGBOX_VALMAX As String
    VAR_MSGBOX_LPT = "É necessário digitar somente o número da porta no sensor "
    VAR_MSGBOX_RB = "É necessário selecionar a sequência de bits no sensor "
    VAR_MSGBOX_VALMIN = "É necessário digitar o menor valor medido pelo sensor "
    VAR_MSGBOX_VALMAX = "É necessário digitar o maior valor medido pelo sensor "
    VAR_PodeSalvar = True
    'Sensor 1
    If IsNumeric(TXT_LPT01.Text) = False Then
        MsgBox VAR_MSGBOX_LPT & "1.", vbOKOnly + vbCritical, "Erro"
        TXT_LPT01.SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        EscreveINI VAR_ArquivoINI, "Sensor1", "LPT", TXT_LPT01.Text
    End If
    If RB_01(0).Value = False And RB_01(1).Value = False Then
        MsgBox VAR_MSGBOX_RB & "1.", vbOKOnly + vbCritical, "Erro"
        RB_01(0).SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        If RB_01(0).Value = True Then
            EscreveINI VAR_ArquivoINI, "Sensor1", "Bits", "1"
        Else
            EscreveINI VAR_ArquivoINI, "Sensor1", "Bits", "2"
        End If
    End If
    If IsNumeric(TXT_MIN01.Text) = False Then
        MsgBox VAR_MSGBOX_VALMIN & "1.", vbOKOnly + vbCritical, "Erro"
        TXT_MIN01.SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        EscreveINI VAR_ArquivoINI, "Sensor1", "ValMin", TXT_MIN01.Text
    End If
    If IsNumeric(TXT_MAX01.Text) = False Then
        MsgBox VAR_MSGBOX_VALMAX & "1.", vbOKOnly + vbCritical, "Erro"
        TXT_MAX01.SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        EscreveINI VAR_ArquivoINI, "Sensor1", "ValMax", TXT_MAX01.Text
    End If
    'Sensor 2
    If IsNumeric(TXT_LPT02.Text) = False Then
        MsgBox VAR_MSGBOX_LPT & "2.", vbOKOnly + vbCritical, "Erro"
        TXT_LPT02.SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        EscreveINI VAR_ArquivoINI, "Sensor2", "LPT", TXT_LPT02.Text
    End If
    If RB_02(0).Value = False And RB_02(1).Value = False Then
        MsgBox VAR_MSGBOX_RB & "2.", vbOKOnly + vbCritical, "Erro"
        RB_02(0).SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        If RB_02(0).Value = True Then
            EscreveINI VAR_ArquivoINI, "Sensor2", "Bits", "1"
        Else
            EscreveINI VAR_ArquivoINI, "Sensor2", "Bits", "2"
        End If
    End If
    If IsNumeric(TXT_MIN02.Text) = False Then
        MsgBox VAR_MSGBOX_VALMIN & "2.", vbOKOnly + vbCritical, "Erro"
        TXT_MIN02.SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        EscreveINI VAR_ArquivoINI, "Sensor2", "ValMin", TXT_MIN02.Text
    End If
    If IsNumeric(TXT_MAX02.Text) = False Then
        MsgBox VAR_MSGBOX_VALMAX & "2.", vbOKOnly + vbCritical, "Erro"
        TXT_MAX02.SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        EscreveINI VAR_ArquivoINI, "Sensor2", "ValMax", TXT_MAX02.Text
    End If
    'Sensor 3
    If IsNumeric(TXT_LPT03.Text) = False Then
        MsgBox VAR_MSGBOX_LPT & "3.", vbOKOnly + vbCritical, "Erro"
        TXT_LPT03.SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        EscreveINI VAR_ArquivoINI, "Sensor3", "LPT", TXT_LPT03.Text
    End If
    If RB_03(0).Value = False And RB_03(1).Value = False Then
        MsgBox VAR_MSGBOX_RB & "3.", vbOKOnly + vbCritical, "Erro"
        RB_03(0).SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        If RB_03(0).Value = True Then
            EscreveINI VAR_ArquivoINI, "Sensor3", "Bits", "1"
        Else
            EscreveINI VAR_ArquivoINI, "Sensor3", "Bits", "2"
        End If
    End If
    If IsNumeric(TXT_MIN03.Text) = False Then
        MsgBox VAR_MSGBOX_VALMIN & "3.", vbOKOnly + vbCritical, "Erro"
        TXT_MIN03.SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        EscreveINI VAR_ArquivoINI, "Sensor3", "ValMin", TXT_MIN03.Text
    End If
    If IsNumeric(TXT_MAX03.Text) = False Then
        MsgBox VAR_MSGBOX_VALMAX & "3.", vbOKOnly + vbCritical, "Erro"
        TXT_MAX03.SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        EscreveINI VAR_ArquivoINI, "Sensor3", "ValMax", TXT_MAX03.Text
    End If
    'Sensor 4
    If IsNumeric(TXT_LPT04.Text) = False Then
        MsgBox VAR_MSGBOX_LPT & "4.", vbOKOnly + vbCritical, "Erro"
        TXT_LPT04.SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        EscreveINI VAR_ArquivoINI, "Sensor4", "LPT", TXT_LPT04.Text
    End If
    If RB_04(0).Value = False And RB_04(1).Value = False Then
        MsgBox VAR_MSGBOX_RB & "4.", vbOKOnly + vbCritical, "Erro"
        RB_04(0).SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        If RB_04(0).Value = True Then
            EscreveINI VAR_ArquivoINI, "Sensor4", "Bits", "1"
        Else
            EscreveINI VAR_ArquivoINI, "Sensor4", "Bits", "2"
        End If
    End If
    If IsNumeric(TXT_MIN04.Text) = False Then
        MsgBox VAR_MSGBOX_VALMIN & "4.", vbOKOnly + vbCritical, "Erro"
        TXT_MIN04.SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        EscreveINI VAR_ArquivoINI, "Sensor4", "ValMin", TXT_MIN04.Text
    End If
    If IsNumeric(TXT_MAX04.Text) = False Then
        MsgBox VAR_MSGBOX_VALMAX & "4.", vbOKOnly + vbCritical, "Erro"
        TXT_MAX04.SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        EscreveINI VAR_ArquivoINI, "Sensor4", "ValMax", TXT_MAX04.Text
    End If
    'Sensor 5
    If IsNumeric(TXT_LPT05.Text) = False Then
        MsgBox VAR_MSGBOX_LPT & "5.", vbOKOnly + vbCritical, "Erro"
        TXT_LPT05.SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        EscreveINI VAR_ArquivoINI, "Sensor5", "LPT", TXT_LPT05.Text
    End If
    If RB_05(0).Value = False And RB_05(1).Value = False Then
        MsgBox VAR_MSGBOX_RB & "5.", vbOKOnly + vbCritical, "Erro"
        RB_05(0).SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        If RB_05(0).Value = True Then
            EscreveINI VAR_ArquivoINI, "Sensor5", "Bits", "1"
        Else
            EscreveINI VAR_ArquivoINI, "Sensor5", "Bits", "2"
        End If
    End If
    If IsNumeric(TXT_MIN05.Text) = False Then
        MsgBox VAR_MSGBOX_VALMIN & "5.", vbOKOnly + vbCritical, "Erro"
        TXT_MIN05.SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        EscreveINI VAR_ArquivoINI, "Sensor5", "ValMin", TXT_MIN05.Text
    End If
    If IsNumeric(TXT_MAX05.Text) = False Then
        MsgBox VAR_MSGBOX_VALMAX & "5.", vbOKOnly + vbCritical, "Erro"
        TXT_MAX05.SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        EscreveINI VAR_ArquivoINI, "Sensor5", "ValMax", TXT_MAX05.Text
    End If
    'Sensor 6
    If IsNumeric(TXT_LPT06.Text) = False Then
        MsgBox VAR_MSGBOX_LPT & "6.", vbOKOnly + vbCritical, "Erro"
        TXT_LPT06.SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        EscreveINI VAR_ArquivoINI, "Sensor6", "LPT", TXT_LPT06.Text
    End If
    If RB_06(0).Value = False And RB_06(1).Value = False Then
        MsgBox VAR_MSGBOX_RB & "6.", vbOKOnly + vbCritical, "Erro"
        RB_06(0).SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        If RB_06(0).Value = True Then
            EscreveINI VAR_ArquivoINI, "Sensor6", "Bits", "1"
        Else
            EscreveINI VAR_ArquivoINI, "Sensor6", "Bits", "2"
        End If
    End If
    If IsNumeric(TXT_MIN06.Text) = False Then
        MsgBox VAR_MSGBOX_VALMIN & "6.", vbOKOnly + vbCritical, "Erro"
        TXT_MIN06.SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        EscreveINI VAR_ArquivoINI, "Sensor6", "ValMin", TXT_MIN06.Text
    End If
    If IsNumeric(TXT_MAX06.Text) = False Then
        MsgBox VAR_MSGBOX_VALMAX & "6.", vbOKOnly + vbCritical, "Erro"
        TXT_MAX06.SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        EscreveINI VAR_ArquivoINI, "Sensor6", "ValMax", TXT_MAX06.Text
    End If
    'Sensor 7
    If IsNumeric(TXT_LPT07.Text) = False Then
        MsgBox VAR_MSGBOX_LPT & "7.", vbOKOnly + vbCritical, "Erro"
        TXT_LPT07.SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        EscreveINI VAR_ArquivoINI, "Sensor7", "LPT", TXT_LPT07.Text
    End If
    If RB_07(0).Value = False And RB_07(1).Value = False Then
        MsgBox VAR_MSGBOX_RB & "7.", vbOKOnly + vbCritical, "Erro"
        RB_07(0).SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        If RB_07(0).Value = True Then
            EscreveINI VAR_ArquivoINI, "Sensor7", "Bits", "1"
        Else
            EscreveINI VAR_ArquivoINI, "Sensor7", "Bits", "2"
        End If
    End If
    If IsNumeric(TXT_MIN07.Text) = False Then
        MsgBox VAR_MSGBOX_VALMIN & "7.", vbOKOnly + vbCritical, "Erro"
        TXT_MIN07.SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        EscreveINI VAR_ArquivoINI, "Sensor7", "ValMin", TXT_MIN07.Text
    End If
    If IsNumeric(TXT_MAX07.Text) = False Then
        MsgBox VAR_MSGBOX_VALMAX & "7.", vbOKOnly + vbCritical, "Erro"
        TXT_MAX07.SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        EscreveINI VAR_ArquivoINI, "Sensor7", "ValMax", TXT_MAX07.Text
    End If
    'Sensor 8
    If IsNumeric(TXT_LPT08.Text) = False Then
        MsgBox VAR_MSGBOX_LPT & "8.", vbOKOnly + vbCritical, "Erro"
        TXT_LPT08.SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        EscreveINI VAR_ArquivoINI, "Sensor8", "LPT", TXT_LPT08.Text
    End If
    If RB_08(0).Value = False And RB_08(1).Value = False Then
        MsgBox VAR_MSGBOX_RB & "8.", vbOKOnly + vbCritical, "Erro"
        RB_08(0).SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        If RB_08(0).Value = True Then
            EscreveINI VAR_ArquivoINI, "Sensor8", "Bits", "1"
        Else
            EscreveINI VAR_ArquivoINI, "Sensor8", "Bits", "2"
        End If
    End If
    If IsNumeric(TXT_MIN08.Text) = False Then
        MsgBox VAR_MSGBOX_VALMIN & "8.", vbOKOnly + vbCritical, "Erro"
        TXT_MIN08.SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        EscreveINI VAR_ArquivoINI, "Sensor8", "ValMin", TXT_MIN08.Text
    End If
    If IsNumeric(TXT_MAX08.Text) = False Then
        MsgBox VAR_MSGBOX_VALMAX & "8.", vbOKOnly + vbCritical, "Erro"
        TXT_MAX08.SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        EscreveINI VAR_ArquivoINI, "Sensor8", "ValMax", TXT_MAX08.Text
    End If
    'Sensor 9
    If IsNumeric(TXT_LPT09.Text) = False Then
        MsgBox VAR_MSGBOX_LPT & "9.", vbOKOnly + vbCritical, "Erro"
        TXT_LPT09.SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        EscreveINI VAR_ArquivoINI, "Sensor9", "LPT", TXT_LPT09.Text
    End If
    If RB_09(0).Value = False And RB_09(1).Value = False Then
        MsgBox VAR_MSGBOX_RB & "9.", vbOKOnly + vbCritical, "Erro"
        RB_09(0).SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        If RB_09(0).Value = True Then
            EscreveINI VAR_ArquivoINI, "Sensor9", "Bits", "1"
        Else
            EscreveINI VAR_ArquivoINI, "Sensor9", "Bits", "2"
        End If
    End If
    If IsNumeric(TXT_MIN09.Text) = False Then
        MsgBox VAR_MSGBOX_VALMIN & "9.", vbOKOnly + vbCritical, "Erro"
        TXT_MIN09.SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        EscreveINI VAR_ArquivoINI, "Sensor9", "ValMin", TXT_MIN09.Text
    End If
    If IsNumeric(TXT_MAX09.Text) = False Then
        MsgBox VAR_MSGBOX_VALMAX & "9.", vbOKOnly + vbCritical, "Erro"
        TXT_MAX09.SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        EscreveINI VAR_ArquivoINI, "Sensor9", "ValMax", TXT_MAX09.Text
    End If
    'Sensor 10
    If IsNumeric(TXT_LPT10.Text) = False Then
        MsgBox VAR_MSGBOX_LPT & "10.", vbOKOnly + vbCritical, "Erro"
        TXT_LPT10.SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        EscreveINI VAR_ArquivoINI, "Sensor10", "LPT", TXT_LPT10.Text
    End If
    If RB_10(0).Value = False And RB_10(1).Value = False Then
        MsgBox VAR_MSGBOX_RB & "10.", vbOKOnly + vbCritical, "Erro"
        RB_10(0).SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        If RB_10(0).Value = True Then
            EscreveINI VAR_ArquivoINI, "Sensor10", "Bits", "1"
        Else
            EscreveINI VAR_ArquivoINI, "Sensor10", "Bits", "2"
        End If
    End If
    If IsNumeric(TXT_MIN10.Text) = False Then
        MsgBox VAR_MSGBOX_VALMIN & "10.", vbOKOnly + vbCritical, "Erro"
        TXT_MIN10.SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        EscreveINI VAR_ArquivoINI, "Sensor10", "ValMin", TXT_MIN10.Text
    End If
    If IsNumeric(TXT_MAX10.Text) = False Then
        MsgBox VAR_MSGBOX_VALMAX & "10.", vbOKOnly + vbCritical, "Erro"
        TXT_MAX10.SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        EscreveINI VAR_ArquivoINI, "Sensor10", "ValMax", TXT_MAX10.Text
    End If
    'Sensor 11
    If IsNumeric(TXT_LPT11.Text) = False Then
        MsgBox VAR_MSGBOX_LPT & "11.", vbOKOnly + vbCritical, "Erro"
        TXT_LPT11.SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        EscreveINI VAR_ArquivoINI, "Sensor11", "LPT", TXT_LPT11.Text
    End If
    If RB_11(0).Value = False And RB_11(1).Value = False Then
        MsgBox VAR_MSGBOX_RB & "11.", vbOKOnly + vbCritical, "Erro"
        RB_11(0).SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        If RB_11(0).Value = True Then
            EscreveINI VAR_ArquivoINI, "Sensor11", "Bits", "1"
        Else
            EscreveINI VAR_ArquivoINI, "Sensor11", "Bits", "2"
        End If
    End If
    If IsNumeric(TXT_MIN11.Text) = False Then
        MsgBox VAR_MSGBOX_VALMIN & "11.", vbOKOnly + vbCritical, "Erro"
        TXT_MIN11.SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        EscreveINI VAR_ArquivoINI, "Sensor11", "ValMin", TXT_MIN11.Text
    End If
    If IsNumeric(TXT_MAX11.Text) = False Then
        MsgBox VAR_MSGBOX_VALMAX & "11.", vbOKOnly + vbCritical, "Erro"
        TXT_MAX11.SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        EscreveINI VAR_ArquivoINI, "Sensor11", "ValMax", TXT_MAX11.Text
    End If
    'Sensor 12
    If IsNumeric(TXT_LPT12.Text) = False Then
        MsgBox VAR_MSGBOX_LPT & "12.", vbOKOnly + vbCritical, "Erro"
        TXT_LPT12.SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        EscreveINI VAR_ArquivoINI, "Sensor12", "LPT", TXT_LPT12.Text
    End If
    If RB_12(0).Value = False And RB_12(1).Value = False Then
        MsgBox VAR_MSGBOX_RB & "12.", vbOKOnly + vbCritical, "Erro"
        RB_12(0).SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        If RB_12(0).Value = True Then
            EscreveINI VAR_ArquivoINI, "Sensor12", "Bits", "1"
        Else
            EscreveINI VAR_ArquivoINI, "Sensor12", "Bits", "2"
        End If
    End If
    If IsNumeric(TXT_MIN12.Text) = False Then
        MsgBox VAR_MSGBOX_VALMIN & "12.", vbOKOnly + vbCritical, "Erro"
        TXT_MIN12.SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        EscreveINI VAR_ArquivoINI, "Sensor12", "ValMin", TXT_MIN12.Text
    End If
    If IsNumeric(TXT_MAX12.Text) = False Then
        MsgBox VAR_MSGBOX_VALMAX & "12.", vbOKOnly + vbCritical, "Erro"
        TXT_MAX12.SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        EscreveINI VAR_ArquivoINI, "Sensor12", "ValMax", TXT_MAX12.Text
    End If
    'Sensor 13
    If IsNumeric(TXT_LPT13.Text) = False Then
        MsgBox VAR_MSGBOX_LPT & "13.", vbOKOnly + vbCritical, "Erro"
        TXT_LPT13.SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        EscreveINI VAR_ArquivoINI, "Sensor13", "LPT", TXT_LPT13.Text
    End If
    If RB_13(0).Value = False And RB_13(1).Value = False Then
        MsgBox VAR_MSGBOX_RB & "13.", vbOKOnly + vbCritical, "Erro"
        RB_13(0).SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        If RB_13(0).Value = True Then
            EscreveINI VAR_ArquivoINI, "Sensor13", "Bits", "1"
        Else
            EscreveINI VAR_ArquivoINI, "Sensor13", "Bits", "2"
        End If
    End If
    If IsNumeric(TXT_MIN13.Text) = False Then
        MsgBox VAR_MSGBOX_VALMIN & "13.", vbOKOnly + vbCritical, "Erro"
        TXT_MIN13.SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        EscreveINI VAR_ArquivoINI, "Sensor13", "ValMin", TXT_MIN13.Text
    End If
    If IsNumeric(TXT_MAX13.Text) = False Then
        MsgBox VAR_MSGBOX_VALMAX & "13.", vbOKOnly + vbCritical, "Erro"
        TXT_MAX13.SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        EscreveINI VAR_ArquivoINI, "Sensor13", "ValMax", TXT_MAX13.Text
    End If
    'Sensor 14
    If IsNumeric(TXT_LPT14.Text) = False Then
        MsgBox VAR_MSGBOX_LPT & "14.", vbOKOnly + vbCritical, "Erro"
        TXT_LPT14.SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        EscreveINI VAR_ArquivoINI, "Sensor14", "LPT", TXT_LPT14.Text
    End If
    If RB_14(0).Value = False And RB_14(1).Value = False Then
        MsgBox VAR_MSGBOX_RB & "14.", vbOKOnly + vbCritical, "Erro"
        RB_14(0).SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        If RB_14(0).Value = True Then
            EscreveINI VAR_ArquivoINI, "Sensor14", "Bits", "1"
        Else
            EscreveINI VAR_ArquivoINI, "Sensor14", "Bits", "2"
        End If
    End If
    If IsNumeric(TXT_MIN14.Text) = False Then
        MsgBox VAR_MSGBOX_VALMIN & "14.", vbOKOnly + vbCritical, "Erro"
        TXT_MIN14.SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        EscreveINI VAR_ArquivoINI, "Sensor14", "ValMin", TXT_MIN14.Text
    End If
    If IsNumeric(TXT_MAX14.Text) = False Then
        MsgBox VAR_MSGBOX_VALMAX & "14.", vbOKOnly + vbCritical, "Erro"
        TXT_MAX14.SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        EscreveINI VAR_ArquivoINI, "Sensor14", "ValMax", TXT_MAX14.Text
    End If
    'Sensor 15
    If IsNumeric(TXT_LPT15.Text) = False Then
        MsgBox VAR_MSGBOX_LPT & "15.", vbOKOnly + vbCritical, "Erro"
        TXT_LPT15.SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        EscreveINI VAR_ArquivoINI, "Sensor15", "LPT", TXT_LPT15.Text
    End If
    If RB_15(0).Value = False And RB_15(1).Value = False Then
        MsgBox VAR_MSGBOX_RB & "15.", vbOKOnly + vbCritical, "Erro"
        RB_15(0).SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        If RB_15(0).Value = True Then
            EscreveINI VAR_ArquivoINI, "Sensor15", "Bits", "1"
        Else
            EscreveINI VAR_ArquivoINI, "Sensor15", "Bits", "2"
        End If
    End If
    If IsNumeric(TXT_MIN15.Text) = False Then
        MsgBox VAR_MSGBOX_VALMIN & "15.", vbOKOnly + vbCritical, "Erro"
        TXT_MIN15.SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        EscreveINI VAR_ArquivoINI, "Sensor15", "ValMin", TXT_MIN15.Text
    End If
    If IsNumeric(TXT_MAX15.Text) = False Then
        MsgBox VAR_MSGBOX_VALMAX & "15.", vbOKOnly + vbCritical, "Erro"
        TXT_MAX15.SetFocus
        VAR_PodeSalvar = False
        Exit Sub
    Else
        EscreveINI VAR_ArquivoINI, "Sensor15", "ValMax", TXT_MAX15.Text
    End If
    'Desabilita botão salvar
    HabilitaSalvarDados False
End Sub
