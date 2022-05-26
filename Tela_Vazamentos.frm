VERSION 5.00
Begin VB.Form Tela_Vazamentos 
   Caption         =   "Tabela de Vazamentos permitidos durante teste de estanqueidade"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11760
   ControlBox      =   0   'False
   Icon            =   "Tela_Vazamentos.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8355
   ScaleWidth      =   11760
   WindowState     =   2  'Maximized
   Begin VB.Frame FR 
      BorderStyle     =   0  'None
      Height          =   5535
      Left            =   0
      TabIndex        =   80
      Top             =   120
      Width           =   11295
      Begin VB.Frame Frame4 
         Caption         =   "Teste Pneumático:"
         Height          =   4095
         Left            =   7200
         TabIndex        =   88
         Top             =   0
         Width           =   4095
         Begin VB.TextBox TXT_PS10 
            Height          =   285
            Left            =   2760
            TabIndex        =   79
            Top             =   3720
            Width           =   1215
         End
         Begin VB.TextBox TXT_PR10 
            Height          =   285
            Left            =   1440
            TabIndex        =   78
            Top             =   3720
            Width           =   1215
         End
         Begin VB.TextBox TXT_PO10 
            Height          =   285
            Left            =   120
            TabIndex        =   77
            Top             =   3720
            Width           =   1215
         End
         Begin VB.TextBox TXT_PS9 
            Height          =   285
            Left            =   2760
            TabIndex        =   71
            Top             =   3360
            Width           =   1215
         End
         Begin VB.TextBox TXT_PR9 
            Height          =   285
            Left            =   1440
            TabIndex        =   70
            Top             =   3360
            Width           =   1215
         End
         Begin VB.TextBox TXT_PO9 
            Height          =   285
            Left            =   120
            TabIndex        =   69
            Top             =   3360
            Width           =   1215
         End
         Begin VB.TextBox TXT_PS8 
            Height          =   285
            Left            =   2760
            TabIndex        =   63
            Top             =   3000
            Width           =   1215
         End
         Begin VB.TextBox TXT_PR8 
            Height          =   285
            Left            =   1440
            TabIndex        =   62
            Top             =   3000
            Width           =   1215
         End
         Begin VB.TextBox TXT_PO8 
            Height          =   285
            Left            =   120
            TabIndex        =   61
            Top             =   3000
            Width           =   1215
         End
         Begin VB.TextBox TXT_PS7 
            Height          =   285
            Left            =   2760
            TabIndex        =   55
            Top             =   2640
            Width           =   1215
         End
         Begin VB.TextBox TXT_PR7 
            Height          =   285
            Left            =   1440
            TabIndex        =   54
            Top             =   2640
            Width           =   1215
         End
         Begin VB.TextBox TXT_PO7 
            Height          =   285
            Left            =   120
            TabIndex        =   53
            Top             =   2640
            Width           =   1215
         End
         Begin VB.TextBox TXT_PS6 
            Height          =   285
            Left            =   2760
            TabIndex        =   47
            Top             =   2280
            Width           =   1215
         End
         Begin VB.TextBox TXT_PR6 
            Height          =   285
            Left            =   1440
            TabIndex        =   46
            Top             =   2280
            Width           =   1215
         End
         Begin VB.TextBox TXT_PO6 
            Height          =   285
            Left            =   120
            TabIndex        =   45
            Top             =   2280
            Width           =   1215
         End
         Begin VB.TextBox TXT_PS5 
            Height          =   285
            Left            =   2760
            TabIndex        =   39
            Top             =   1920
            Width           =   1215
         End
         Begin VB.TextBox TXT_PR5 
            Height          =   285
            Left            =   1440
            TabIndex        =   38
            Top             =   1920
            Width           =   1215
         End
         Begin VB.TextBox TXT_PO5 
            Height          =   285
            Left            =   120
            TabIndex        =   37
            Top             =   1920
            Width           =   1215
         End
         Begin VB.TextBox TXT_PS4 
            Height          =   285
            Left            =   2760
            TabIndex        =   31
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox TXT_PR4 
            Height          =   285
            Left            =   1440
            TabIndex        =   30
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox TXT_PO4 
            Height          =   285
            Left            =   120
            TabIndex        =   29
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox TXT_PS3 
            Height          =   285
            Left            =   2760
            TabIndex        =   23
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox TXT_PR3 
            Height          =   285
            Left            =   1440
            TabIndex        =   22
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox TXT_PO3 
            Height          =   285
            Left            =   120
            TabIndex        =   21
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox TXT_PS2 
            Height          =   285
            Left            =   2760
            TabIndex        =   15
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox TXT_PR2 
            Height          =   285
            Left            =   1440
            TabIndex        =   14
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox TXT_PO2 
            Height          =   285
            Left            =   120
            TabIndex        =   13
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox TXT_PO1 
            Height          =   285
            Left            =   120
            TabIndex        =   5
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox TXT_PR1 
            Height          =   285
            Left            =   1440
            TabIndex        =   6
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox TXT_PS1 
            Height          =   285
            Left            =   2760
            TabIndex        =   7
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Metal - Outras:"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   91
            ToolTipText     =   "Pressão Máxima de Trabalho"
            Top             =   240
            Width           =   1035
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Metal - Retenção"
            Height          =   195
            Index           =   3
            Left            =   1440
            TabIndex        =   90
            ToolTipText     =   "Temperatura Máxima de Trabalho"
            Top             =   240
            Width           =   1230
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Sede Resiliente:"
            Height          =   195
            Index           =   2
            Left            =   2760
            TabIndex        =   89
            ToolTipText     =   "Temperatura Máxima de Trabalho"
            Top             =   240
            Width           =   1155
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Teste Hidrostático:"
         Height          =   4095
         Left            =   3000
         TabIndex        =   84
         Top             =   0
         Width           =   4095
         Begin VB.TextBox TXT_HS10 
            Height          =   285
            Left            =   2760
            TabIndex        =   76
            Top             =   3720
            Width           =   1215
         End
         Begin VB.TextBox TXT_HR10 
            Height          =   285
            Left            =   1440
            TabIndex        =   75
            Top             =   3720
            Width           =   1215
         End
         Begin VB.TextBox TXT_HO10 
            Height          =   285
            Left            =   120
            TabIndex        =   74
            Top             =   3720
            Width           =   1215
         End
         Begin VB.TextBox TXT_HS9 
            Height          =   285
            Left            =   2760
            TabIndex        =   68
            Top             =   3360
            Width           =   1215
         End
         Begin VB.TextBox TXT_HR9 
            Height          =   285
            Left            =   1440
            TabIndex        =   67
            Top             =   3360
            Width           =   1215
         End
         Begin VB.TextBox TXT_HO9 
            Height          =   285
            Left            =   120
            TabIndex        =   66
            Top             =   3360
            Width           =   1215
         End
         Begin VB.TextBox TXT_HS8 
            Height          =   285
            Left            =   2760
            TabIndex        =   60
            Top             =   3000
            Width           =   1215
         End
         Begin VB.TextBox TXT_HR8 
            Height          =   285
            Left            =   1440
            TabIndex        =   59
            Top             =   3000
            Width           =   1215
         End
         Begin VB.TextBox TXT_HO8 
            Height          =   285
            Left            =   120
            TabIndex        =   58
            Top             =   3000
            Width           =   1215
         End
         Begin VB.TextBox TXT_HS7 
            Height          =   285
            Left            =   2760
            TabIndex        =   52
            Top             =   2640
            Width           =   1215
         End
         Begin VB.TextBox TXT_HR7 
            Height          =   285
            Left            =   1440
            TabIndex        =   51
            Top             =   2640
            Width           =   1215
         End
         Begin VB.TextBox TXT_HO7 
            Height          =   285
            Left            =   120
            TabIndex        =   50
            Top             =   2640
            Width           =   1215
         End
         Begin VB.TextBox TXT_HS6 
            Height          =   285
            Left            =   2760
            TabIndex        =   44
            Top             =   2280
            Width           =   1215
         End
         Begin VB.TextBox TXT_HR6 
            Height          =   285
            Left            =   1440
            TabIndex        =   43
            Top             =   2280
            Width           =   1215
         End
         Begin VB.TextBox TXT_HO6 
            Height          =   285
            Left            =   120
            TabIndex        =   42
            Top             =   2280
            Width           =   1215
         End
         Begin VB.TextBox TXT_HS5 
            Height          =   285
            Left            =   2760
            TabIndex        =   36
            Top             =   1920
            Width           =   1215
         End
         Begin VB.TextBox TXT_HR5 
            Height          =   285
            Left            =   1440
            TabIndex        =   35
            Top             =   1920
            Width           =   1215
         End
         Begin VB.TextBox TXT_HO5 
            Height          =   285
            Left            =   120
            TabIndex        =   34
            Top             =   1920
            Width           =   1215
         End
         Begin VB.TextBox TXT_HS4 
            Height          =   285
            Left            =   2760
            TabIndex        =   28
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox TXT_HR4 
            Height          =   285
            Left            =   1440
            TabIndex        =   27
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox TXT_HO4 
            Height          =   285
            Left            =   120
            TabIndex        =   26
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox TXT_HS3 
            Height          =   285
            Left            =   2760
            TabIndex        =   20
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox TXT_HR3 
            Height          =   285
            Left            =   1440
            TabIndex        =   19
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox TXT_HO3 
            Height          =   285
            Left            =   120
            TabIndex        =   18
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox TXT_HS2 
            Height          =   285
            Left            =   2760
            TabIndex        =   12
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox TXT_HR2 
            Height          =   285
            Left            =   1440
            TabIndex        =   11
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox TXT_HO2 
            Height          =   285
            Left            =   120
            TabIndex        =   10
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox TXT_HS1 
            Height          =   285
            Left            =   2760
            TabIndex        =   4
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox TXT_HR1 
            Height          =   285
            Left            =   1440
            TabIndex        =   3
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox TXT_HO1 
            Height          =   285
            Left            =   120
            TabIndex        =   2
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Sede Resiliente:"
            Height          =   195
            Index           =   4
            Left            =   2760
            TabIndex        =   87
            ToolTipText     =   "Temperatura Máxima de Trabalho"
            Top             =   240
            Width           =   1155
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Metal - Retenção"
            Height          =   195
            Index           =   1
            Left            =   1440
            TabIndex        =   86
            ToolTipText     =   "Temperatura Máxima de Trabalho"
            Top             =   240
            Width           =   1230
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Metal - Outras:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   85
            ToolTipText     =   "Pressão Máxima de Trabalho"
            Top             =   240
            Width           =   1035
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Intervalos dos Ciclos:"
         Height          =   4095
         Left            =   120
         TabIndex        =   81
         Top             =   0
         Width           =   2775
         Begin VB.TextBox TXT_C10E 
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   72
            Top             =   3720
            Width           =   1215
         End
         Begin VB.TextBox TXT_C10A 
            Height          =   285
            Left            =   1440
            TabIndex        =   73
            Top             =   3720
            Width           =   1215
         End
         Begin VB.TextBox TXT_C9E 
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   64
            Top             =   3360
            Width           =   1215
         End
         Begin VB.TextBox TXT_C9A 
            Height          =   285
            Left            =   1440
            TabIndex        =   65
            Top             =   3360
            Width           =   1215
         End
         Begin VB.TextBox TXT_C8E 
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   56
            Top             =   3000
            Width           =   1215
         End
         Begin VB.TextBox TXT_C8A 
            Height          =   285
            Left            =   1440
            TabIndex        =   57
            Top             =   3000
            Width           =   1215
         End
         Begin VB.TextBox TXT_C7E 
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   48
            Top             =   2640
            Width           =   1215
         End
         Begin VB.TextBox TXT_C7A 
            Height          =   285
            Left            =   1440
            TabIndex        =   49
            Top             =   2640
            Width           =   1215
         End
         Begin VB.TextBox TXT_C6E 
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   40
            Top             =   2280
            Width           =   1215
         End
         Begin VB.TextBox TXT_C6A 
            Height          =   285
            Left            =   1440
            TabIndex        =   41
            Top             =   2280
            Width           =   1215
         End
         Begin VB.TextBox TXT_C5E 
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   32
            Top             =   1920
            Width           =   1215
         End
         Begin VB.TextBox TXT_C5A 
            Height          =   285
            Left            =   1440
            TabIndex        =   33
            Top             =   1920
            Width           =   1215
         End
         Begin VB.TextBox TXT_C4E 
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   24
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox TXT_C4A 
            Height          =   285
            Left            =   1440
            TabIndex        =   25
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox TXT_C3E 
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   16
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox TXT_C3A 
            Height          =   285
            Left            =   1440
            TabIndex        =   17
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox TXT_C2E 
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   8
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox TXT_C2A 
            Height          =   285
            Left            =   1440
            TabIndex        =   9
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox TXT_C1E 
            Height          =   285
            Left            =   120
            TabIndex        =   0
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox TXT_C1A 
            Height          =   285
            Left            =   1440
            TabIndex        =   1
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Maior que:"
            Height          =   195
            Index           =   17
            Left            =   120
            TabIndex        =   83
            ToolTipText     =   "Pressão Máxima de Trabalho"
            Top             =   240
            Width           =   750
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Menor Igual que:"
            Height          =   195
            Index           =   16
            Left            =   1440
            TabIndex        =   82
            ToolTipText     =   "Temperatura Máxima de Trabalho"
            Top             =   240
            Width           =   1200
         End
      End
   End
End
Attribute VB_Name = "Tela_Vazamentos"
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
    VAR_Tela = "Tela_Vazamentos"
    'Le todos valores dos campos
    'Ciclo 1
    TXT_C1E.Text = LeINI(VAR_ArquivoINI, "Vazamento1", "CicloMenor")
    TXT_C1A.Text = LeINI(VAR_ArquivoINI, "Vazamento1", "CicloMaior")
    TXT_HO1.Text = LeINI(VAR_ArquivoINI, "Vazamento1", "HidrostaticoOutras")
    TXT_HR1.Text = LeINI(VAR_ArquivoINI, "Vazamento1", "HidrostaticoRetencao")
    TXT_HS1.Text = LeINI(VAR_ArquivoINI, "Vazamento1", "HidrostaticoSedes")
    TXT_PO1.Text = LeINI(VAR_ArquivoINI, "Vazamento1", "PneumaticoOutras")
    TXT_PR1.Text = LeINI(VAR_ArquivoINI, "Vazamento1", "PneumaticoRetencao")
    TXT_PS1.Text = LeINI(VAR_ArquivoINI, "Vazamento1", "PneumaticoSedes")
    'Ciclo 2
    TXT_C2E.Text = LeINI(VAR_ArquivoINI, "Vazamento2", "CicloMenor")
    TXT_C2A.Text = LeINI(VAR_ArquivoINI, "Vazamento2", "CicloMaior")
    TXT_HO2.Text = LeINI(VAR_ArquivoINI, "Vazamento2", "HidrostaticoOutras")
    TXT_HR2.Text = LeINI(VAR_ArquivoINI, "Vazamento2", "HidrostaticoRetencao")
    TXT_HS2.Text = LeINI(VAR_ArquivoINI, "Vazamento2", "HidrostaticoSedes")
    TXT_PO2.Text = LeINI(VAR_ArquivoINI, "Vazamento2", "PneumaticoOutras")
    TXT_PR2.Text = LeINI(VAR_ArquivoINI, "Vazamento2", "PneumaticoRetencao")
    TXT_PS2.Text = LeINI(VAR_ArquivoINI, "Vazamento2", "PneumaticoSedes")
    'Ciclo 3
    TXT_C3E.Text = LeINI(VAR_ArquivoINI, "Vazamento3", "CicloMenor")
    TXT_C3A.Text = LeINI(VAR_ArquivoINI, "Vazamento3", "CicloMaior")
    TXT_HO3.Text = LeINI(VAR_ArquivoINI, "Vazamento3", "HidrostaticoOutras")
    TXT_HR3.Text = LeINI(VAR_ArquivoINI, "Vazamento3", "HidrostaticoRetencao")
    TXT_HS3.Text = LeINI(VAR_ArquivoINI, "Vazamento3", "HidrostaticoSedes")
    TXT_PO3.Text = LeINI(VAR_ArquivoINI, "Vazamento3", "PneumaticoOutras")
    TXT_PR3.Text = LeINI(VAR_ArquivoINI, "Vazamento3", "PneumaticoRetencao")
    TXT_PS3.Text = LeINI(VAR_ArquivoINI, "Vazamento3", "PneumaticoSedes")
    'Ciclo 4
    TXT_C4E.Text = LeINI(VAR_ArquivoINI, "Vazamento4", "CicloMenor")
    TXT_C4A.Text = LeINI(VAR_ArquivoINI, "Vazamento4", "CicloMaior")
    TXT_HO4.Text = LeINI(VAR_ArquivoINI, "Vazamento4", "HidrostaticoOutras")
    TXT_HR4.Text = LeINI(VAR_ArquivoINI, "Vazamento4", "HidrostaticoRetencao")
    TXT_HS4.Text = LeINI(VAR_ArquivoINI, "Vazamento4", "HidrostaticoSedes")
    TXT_PO4.Text = LeINI(VAR_ArquivoINI, "Vazamento4", "PneumaticoOutras")
    TXT_PR4.Text = LeINI(VAR_ArquivoINI, "Vazamento4", "PneumaticoRetencao")
    TXT_PS4.Text = LeINI(VAR_ArquivoINI, "Vazamento4", "PneumaticoSedes")
    'Ciclo 5
    TXT_C5E.Text = LeINI(VAR_ArquivoINI, "Vazamento5", "CicloMenor")
    TXT_C5A.Text = LeINI(VAR_ArquivoINI, "Vazamento5", "CicloMaior")
    TXT_HO5.Text = LeINI(VAR_ArquivoINI, "Vazamento5", "HidrostaticoOutras")
    TXT_HR5.Text = LeINI(VAR_ArquivoINI, "Vazamento5", "HidrostaticoRetencao")
    TXT_HS5.Text = LeINI(VAR_ArquivoINI, "Vazamento5", "HidrostaticoSedes")
    TXT_PO5.Text = LeINI(VAR_ArquivoINI, "Vazamento5", "PneumaticoOutras")
    TXT_PR5.Text = LeINI(VAR_ArquivoINI, "Vazamento5", "PneumaticoRetencao")
    TXT_PS5.Text = LeINI(VAR_ArquivoINI, "Vazamento5", "PneumaticoSedes")
    'Ciclo 6
    TXT_C6E.Text = LeINI(VAR_ArquivoINI, "Vazamento6", "CicloMenor")
    TXT_C6A.Text = LeINI(VAR_ArquivoINI, "Vazamento6", "CicloMaior")
    TXT_HO6.Text = LeINI(VAR_ArquivoINI, "Vazamento6", "HidrostaticoOutras")
    TXT_HR6.Text = LeINI(VAR_ArquivoINI, "Vazamento6", "HidrostaticoRetencao")
    TXT_HS6.Text = LeINI(VAR_ArquivoINI, "Vazamento6", "HidrostaticoSedes")
    TXT_PO6.Text = LeINI(VAR_ArquivoINI, "Vazamento6", "PneumaticoOutras")
    TXT_PR6.Text = LeINI(VAR_ArquivoINI, "Vazamento6", "PneumaticoRetencao")
    TXT_PS6.Text = LeINI(VAR_ArquivoINI, "Vazamento6", "PneumaticoSedes")
    'Ciclo 7
    TXT_C7E.Text = LeINI(VAR_ArquivoINI, "Vazamento7", "CicloMenor")
    TXT_C7A.Text = LeINI(VAR_ArquivoINI, "Vazamento7", "CicloMaior")
    TXT_HO7.Text = LeINI(VAR_ArquivoINI, "Vazamento7", "HidrostaticoOutras")
    TXT_HR7.Text = LeINI(VAR_ArquivoINI, "Vazamento7", "HidrostaticoRetencao")
    TXT_HS7.Text = LeINI(VAR_ArquivoINI, "Vazamento7", "HidrostaticoSedes")
    TXT_PO7.Text = LeINI(VAR_ArquivoINI, "Vazamento7", "PneumaticoOutras")
    TXT_PR7.Text = LeINI(VAR_ArquivoINI, "Vazamento7", "PneumaticoRetencao")
    TXT_PS7.Text = LeINI(VAR_ArquivoINI, "Vazamento7", "PneumaticoSedes")
    'Ciclo 8
    TXT_C8E.Text = LeINI(VAR_ArquivoINI, "Vazamento8", "CicloMenor")
    TXT_C8A.Text = LeINI(VAR_ArquivoINI, "Vazamento8", "CicloMaior")
    TXT_HO8.Text = LeINI(VAR_ArquivoINI, "Vazamento8", "HidrostaticoOutras")
    TXT_HR8.Text = LeINI(VAR_ArquivoINI, "Vazamento8", "HidrostaticoRetencao")
    TXT_HS8.Text = LeINI(VAR_ArquivoINI, "Vazamento8", "HidrostaticoSedes")
    TXT_PO8.Text = LeINI(VAR_ArquivoINI, "Vazamento8", "PneumaticoOutras")
    TXT_PR8.Text = LeINI(VAR_ArquivoINI, "Vazamento8", "PneumaticoRetencao")
    TXT_PS8.Text = LeINI(VAR_ArquivoINI, "Vazamento8", "PneumaticoSedes")
    'Ciclo 9
    TXT_C9E.Text = LeINI(VAR_ArquivoINI, "Vazamento9", "CicloMenor")
    TXT_C9A.Text = LeINI(VAR_ArquivoINI, "Vazamento9", "CicloMaior")
    TXT_HO9.Text = LeINI(VAR_ArquivoINI, "Vazamento9", "HidrostaticoOutras")
    TXT_HR9.Text = LeINI(VAR_ArquivoINI, "Vazamento9", "HidrostaticoRetencao")
    TXT_HS9.Text = LeINI(VAR_ArquivoINI, "Vazamento9", "HidrostaticoSedes")
    TXT_PO9.Text = LeINI(VAR_ArquivoINI, "Vazamento9", "PneumaticoOutras")
    TXT_PR9.Text = LeINI(VAR_ArquivoINI, "Vazamento9", "PneumaticoRetencao")
    TXT_PS9.Text = LeINI(VAR_ArquivoINI, "Vazamento9", "PneumaticoSedes")
    'Ciclo 10
    TXT_C10E.Text = LeINI(VAR_ArquivoINI, "Vazamento10", "CicloMenor")
    TXT_C10A.Text = LeINI(VAR_ArquivoINI, "Vazamento10", "CicloMaior")
    TXT_HO10.Text = LeINI(VAR_ArquivoINI, "Vazamento10", "HidrostaticoOutras")
    TXT_HR10.Text = LeINI(VAR_ArquivoINI, "Vazamento10", "HidrostaticoRetencao")
    TXT_HS10.Text = LeINI(VAR_ArquivoINI, "Vazamento10", "HidrostaticoSedes")
    TXT_PO10.Text = LeINI(VAR_ArquivoINI, "Vazamento10", "PneumaticoOutras")
    TXT_PR10.Text = LeINI(VAR_ArquivoINI, "Vazamento10", "PneumaticoRetencao")
    TXT_PS10.Text = LeINI(VAR_ArquivoINI, "Vazamento10", "PneumaticoSedes")
    'Desabilita Botao Salvar Dados
    HabilitaSalvarDados False
End Sub

Private Sub TXT_C10A_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_C10A_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_C10E_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_C10E_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_C1A_Change()
    HabilitaSalvarDados True
    TXT_C2E.Text = TXT_C1A.Text
End Sub

Private Sub TXT_C1A_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_C1E_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_C1E_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
    'If KeyAscii = 13 Then TXT_PorcIPI.SetFocus
End Sub

Private Sub TXT_C2A_Change()
    HabilitaSalvarDados True
    TXT_C3E.Text = TXT_C2A.Text
End Sub

Private Sub TXT_C2A_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_C2E_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_C2E_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_C3A_Change()
    HabilitaSalvarDados True
    TXT_C4E.Text = TXT_C3A.Text
End Sub

Private Sub TXT_C3A_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_C3E_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_C3E_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_C4A_Change()
    HabilitaSalvarDados True
    TXT_C5E.Text = TXT_C4A.Text
End Sub

Private Sub TXT_C4A_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_C4E_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_C4E_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_C5A_Change()
    HabilitaSalvarDados True
    TXT_C6E.Text = TXT_C5A.Text
End Sub

Private Sub TXT_C5A_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_C5E_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_C5E_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_C6A_Change()
    HabilitaSalvarDados True
    TXT_C7E.Text = TXT_C6A.Text
End Sub

Private Sub TXT_C6A_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_C6E_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_C6E_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_C7A_Change()
    HabilitaSalvarDados True
    TXT_C8E.Text = TXT_C7A.Text
End Sub

Private Sub TXT_C7A_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_C7E_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_C7E_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_C8A_Change()
    HabilitaSalvarDados True
    TXT_C9E.Text = TXT_C8A.Text
End Sub

Private Sub TXT_C8A_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_C8E_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_C8E_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_C9A_Change()
    HabilitaSalvarDados True
    TXT_C10E.Text = TXT_C9A.Text
End Sub

Private Sub TXT_C9A_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_C9E_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_C9E_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_HO1_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_HO1_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_HO10_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_HO10_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_HO2_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_HO2_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_HO3_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_HO3_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_HO4_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_HO4_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_HO5_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_HO5_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_HO6_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_HO6_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_HO7_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_HO7_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_HO8_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_HO8_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_HO9_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_HO9_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_HR1_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_HR1_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_HR10_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_HR10_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_HR2_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_HR2_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_HR3_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_HR3_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_HR4_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_HR4_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_HR5_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_HR5_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_HR6_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_HR6_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_HR7_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_HR7_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_HR8_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_HR8_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_HR9_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_HR9_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_HS1_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_HS1_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_HS10_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_HS10_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_HS2_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_HS2_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_HS3_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_HS3_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_HS4_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_HS4_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_HS5_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_HS5_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_HS6_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_HS6_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_HS7_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_HS7_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_HS8_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_HS8_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_HS9_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_HS9_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_PO1_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_PO1_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_PO10_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_PO10_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_PO2_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_PO2_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_PO3_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_PO3_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_PO4_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_PO4_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_PO5_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_PO5_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_PO6_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_PO6_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_PO7_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_PO7_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_PO8_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_PO8_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_PO9_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_PO9_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_PR1_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_PR1_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_PR10_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_PR10_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_PR2_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_PR2_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_PR3_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_PR3_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_PR4_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_PR4_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_PR5_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_PR5_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_PR6_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_PR6_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_PR7_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_PR7_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_PR8_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_PR8_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_PR9_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_PR9_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_PS1_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_PS1_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_PS10_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_PS10_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_PS2_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_PS2_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_PS3_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_PS3_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_PS4_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_PS4_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_PS5_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_PS5_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_PS6_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_PS6_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_PS7_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_PS7_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_PS8_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_PS8_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub

Private Sub TXT_PS9_Change()
    HabilitaSalvarDados True
End Sub

Private Sub TXT_PS9_KeyPress(KeyAscii As Integer)
    KeyAscii = ValidaTexto(KeyAscii)
End Sub



'*************************
' FUNCOES DESTA TELA
'*************************
Public Sub Fechar()
    If VAR_Salvar = True Then
        VAR_RespMsg = MsgBox("Você editou os dados sobre vazamento permitido durante os testes. Tem certeza que desejar sair sem gravar as modificações?", vbQuestion + vbYesNo, "Sair sem salvar")
        If VAR_RespMsg = vbNo Then
            Salvar
            Exit Sub
        End If
   End If
   Unload Tela_Vazamentos
End Sub
Public Sub Salvar()
    'Valida Textos Digitados
    'Ciclo 1
    If IsNumeric(TXT_C1E.Text) = False Then
        MsgBox "Você não digitou o valor do início do 1º ciclo.", vbCritical + vbOKOnly, "Erro"
        TXT_C1E.SetFocus
        Exit Sub
    End If
    'Salva os dados de vazamento
    'Ciclo 1
    EscreveINI VAR_ArquivoINI, "Vazamento1", "CicloMenor", TXT_C1E.Text
    EscreveINI VAR_ArquivoINI, "Vazamento1", "CicloMaior", TXT_C1A.Text
    EscreveINI VAR_ArquivoINI, "Vazamento1", "HidrostaticoOutras", TXT_HO1.Text
    EscreveINI VAR_ArquivoINI, "Vazamento1", "HidrostaticoRetencao", TXT_HR1.Text
    EscreveINI VAR_ArquivoINI, "Vazamento1", "HidrostaticoSedes", TXT_HS1.Text
    EscreveINI VAR_ArquivoINI, "Vazamento1", "PneumaticoOutras", TXT_PO1.Text
    EscreveINI VAR_ArquivoINI, "Vazamento1", "PneumaticoRetencao", TXT_PR1.Text
    EscreveINI VAR_ArquivoINI, "Vazamento1", "PneumaticoSedes", TXT_PS1.Text
    'Ciclo 2
    EscreveINI VAR_ArquivoINI, "Vazamento2", "CicloMenor", TXT_C2E.Text
    EscreveINI VAR_ArquivoINI, "Vazamento2", "CicloMaior", TXT_C2A.Text
    EscreveINI VAR_ArquivoINI, "Vazamento2", "HidrostaticoOutras", TXT_HO2.Text
    EscreveINI VAR_ArquivoINI, "Vazamento2", "HidrostaticoRetencao", TXT_HR2.Text
    EscreveINI VAR_ArquivoINI, "Vazamento2", "HidrostaticoSedes", TXT_HS2.Text
    EscreveINI VAR_ArquivoINI, "Vazamento2", "PneumaticoOutras", TXT_PO2.Text
    EscreveINI VAR_ArquivoINI, "Vazamento2", "PneumaticoRetencao", TXT_PR2.Text
    EscreveINI VAR_ArquivoINI, "Vazamento2", "PneumaticoSedes", TXT_PS2.Text
    'Ciclo 3
    EscreveINI VAR_ArquivoINI, "Vazamento3", "CicloMenor", TXT_C3E.Text
    EscreveINI VAR_ArquivoINI, "Vazamento3", "CicloMaior", TXT_C3A.Text
    EscreveINI VAR_ArquivoINI, "Vazamento3", "HidrostaticoOutras", TXT_HO3.Text
    EscreveINI VAR_ArquivoINI, "Vazamento3", "HidrostaticoRetencao", TXT_HR3.Text
    EscreveINI VAR_ArquivoINI, "Vazamento3", "HidrostaticoSedes", TXT_HS3.Text
    EscreveINI VAR_ArquivoINI, "Vazamento3", "PneumaticoOutras", TXT_PO3.Text
    EscreveINI VAR_ArquivoINI, "Vazamento3", "PneumaticoRetencao", TXT_PR3.Text
    EscreveINI VAR_ArquivoINI, "Vazamento3", "PneumaticoSedes", TXT_PS3.Text
    'Ciclo 4
    EscreveINI VAR_ArquivoINI, "Vazamento4", "CicloMenor", TXT_C4E.Text
    EscreveINI VAR_ArquivoINI, "Vazamento4", "CicloMaior", TXT_C4A.Text
    EscreveINI VAR_ArquivoINI, "Vazamento4", "HidrostaticoOutras", TXT_HO4.Text
    EscreveINI VAR_ArquivoINI, "Vazamento4", "HidrostaticoRetencao", TXT_HR4.Text
    EscreveINI VAR_ArquivoINI, "Vazamento4", "HidrostaticoSedes", TXT_HS4.Text
    EscreveINI VAR_ArquivoINI, "Vazamento4", "PneumaticoOutras", TXT_PO4.Text
    EscreveINI VAR_ArquivoINI, "Vazamento4", "PneumaticoRetencao", TXT_PR4.Text
    EscreveINI VAR_ArquivoINI, "Vazamento4", "PneumaticoSedes", TXT_PS4.Text
    'Ciclo 5
    EscreveINI VAR_ArquivoINI, "Vazamento5", "CicloMenor", TXT_C5E.Text
    EscreveINI VAR_ArquivoINI, "Vazamento5", "CicloMaior", TXT_C5A.Text
    EscreveINI VAR_ArquivoINI, "Vazamento5", "HidrostaticoOutras", TXT_HO5.Text
    EscreveINI VAR_ArquivoINI, "Vazamento5", "HidrostaticoRetencao", TXT_HR5.Text
    EscreveINI VAR_ArquivoINI, "Vazamento5", "HidrostaticoSedes", TXT_HS5.Text
    EscreveINI VAR_ArquivoINI, "Vazamento5", "PneumaticoOutras", TXT_PO5.Text
    EscreveINI VAR_ArquivoINI, "Vazamento5", "PneumaticoRetencao", TXT_PR5.Text
    EscreveINI VAR_ArquivoINI, "Vazamento5", "PneumaticoSedes", TXT_PS5.Text
    'Ciclo 6
    EscreveINI VAR_ArquivoINI, "Vazamento6", "CicloMenor", TXT_C6E.Text
    EscreveINI VAR_ArquivoINI, "Vazamento6", "CicloMaior", TXT_C6A.Text
    EscreveINI VAR_ArquivoINI, "Vazamento6", "HidrostaticoOutras", TXT_HO6.Text
    EscreveINI VAR_ArquivoINI, "Vazamento6", "HidrostaticoRetencao", TXT_HR6.Text
    EscreveINI VAR_ArquivoINI, "Vazamento6", "HidrostaticoSedes", TXT_HS6.Text
    EscreveINI VAR_ArquivoINI, "Vazamento6", "PneumaticoOutras", TXT_PO6.Text
    EscreveINI VAR_ArquivoINI, "Vazamento6", "PneumaticoRetencao", TXT_PR6.Text
    EscreveINI VAR_ArquivoINI, "Vazamento6", "PneumaticoSedes", TXT_PS6.Text
    'Ciclo 7
    EscreveINI VAR_ArquivoINI, "Vazamento7", "CicloMenor", TXT_C7E.Text
    EscreveINI VAR_ArquivoINI, "Vazamento7", "CicloMaior", TXT_C7A.Text
    EscreveINI VAR_ArquivoINI, "Vazamento7", "HidrostaticoOutras", TXT_HO7.Text
    EscreveINI VAR_ArquivoINI, "Vazamento7", "HidrostaticoRetencao", TXT_HR7.Text
    EscreveINI VAR_ArquivoINI, "Vazamento7", "HidrostaticoSedes", TXT_HS7.Text
    EscreveINI VAR_ArquivoINI, "Vazamento7", "PneumaticoOutras", TXT_PO7.Text
    EscreveINI VAR_ArquivoINI, "Vazamento7", "PneumaticoRetencao", TXT_PR7.Text
    EscreveINI VAR_ArquivoINI, "Vazamento7", "PneumaticoSedes", TXT_PS7.Text
    'Ciclo 8
    EscreveINI VAR_ArquivoINI, "Vazamento8", "CicloMenor", TXT_C8E.Text
    EscreveINI VAR_ArquivoINI, "Vazamento8", "CicloMaior", TXT_C8A.Text
    EscreveINI VAR_ArquivoINI, "Vazamento8", "HidrostaticoOutras", TXT_HO8.Text
    EscreveINI VAR_ArquivoINI, "Vazamento8", "HidrostaticoRetencao", TXT_HR8.Text
    EscreveINI VAR_ArquivoINI, "Vazamento8", "HidrostaticoSedes", TXT_HS8.Text
    EscreveINI VAR_ArquivoINI, "Vazamento8", "PneumaticoOutras", TXT_PO8.Text
    EscreveINI VAR_ArquivoINI, "Vazamento8", "PneumaticoRetencao", TXT_PR8.Text
    EscreveINI VAR_ArquivoINI, "Vazamento8", "PneumaticoSedes", TXT_PS8.Text
    'Ciclo 9
    EscreveINI VAR_ArquivoINI, "Vazamento9", "CicloMenor", TXT_C9E.Text
    EscreveINI VAR_ArquivoINI, "Vazamento9", "CicloMaior", TXT_C9A.Text
    EscreveINI VAR_ArquivoINI, "Vazamento9", "HidrostaticoOutras", TXT_HO9.Text
    EscreveINI VAR_ArquivoINI, "Vazamento9", "HidrostaticoRetencao", TXT_HR9.Text
    EscreveINI VAR_ArquivoINI, "Vazamento9", "HidrostaticoSedes", TXT_HS9.Text
    EscreveINI VAR_ArquivoINI, "Vazamento9", "PneumaticoOutras", TXT_PO9.Text
    EscreveINI VAR_ArquivoINI, "Vazamento9", "PneumaticoRetencao", TXT_PR9.Text
    EscreveINI VAR_ArquivoINI, "Vazamento9", "PneumaticoSedes", TXT_PS9.Text
    'Ciclo 10
    EscreveINI VAR_ArquivoINI, "Vazamento10", "CicloMenor", TXT_C10E.Text
    EscreveINI VAR_ArquivoINI, "Vazamento10", "CicloMaior", TXT_C10A.Text
    EscreveINI VAR_ArquivoINI, "Vazamento10", "HidrostaticoOutras", TXT_HO10.Text
    EscreveINI VAR_ArquivoINI, "Vazamento10", "HidrostaticoRetencao", TXT_HR10.Text
    EscreveINI VAR_ArquivoINI, "Vazamento10", "HidrostaticoSedes", TXT_HS10.Text
    EscreveINI VAR_ArquivoINI, "Vazamento10", "PneumaticoOutras", TXT_PO10.Text
    EscreveINI VAR_ArquivoINI, "Vazamento10", "PneumaticoRetencao", TXT_PR10.Text
    EscreveINI VAR_ArquivoINI, "Vazamento10", "PneumaticoSedes", TXT_PS10.Text
    'desabilita botão salvar
    HabilitaSalvarDados False
End Sub

