VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Tela_Simulacao 
   AutoRedraw      =   -1  'True
   Caption         =   "c"
   ClientHeight    =   9690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14760
   ControlBox      =   0   'False
   Icon            =   "Tela_Simulacao.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9690
   ScaleWidth      =   14760
   WindowState     =   2  'Maximized
   Begin VB.Timer TIMER_TEMPOS 
      Interval        =   1000
      Left            =   4440
      Top             =   120
   End
   Begin TabDlg.SSTab ST 
      Height          =   9255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   16325
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Monitoramento da Máquina"
      TabPicture(0)   =   "Tela_Simulacao.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame8"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "TIMER_SIMULACAO"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "TIMER_SMD"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "TIMER_SMD_AUX"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame4"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Frame9"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Frame7"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Plotagem do Gráfico"
      TabPicture(1)   =   "Tela_Simulacao.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "PICB"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Controles da Máquina"
      TabPicture(2)   =   "Tela_Simulacao.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "BT_TesteSimulacao"
      Tab(2).Control(1)=   "BT_LIGA"
      Tab(2).ControlCount=   2
      Begin VB.Frame Frame7 
         Caption         =   "Bomba de Vácuo:"
         Height          =   615
         Left            =   9720
         TabIndex        =   148
         Top             =   1080
         Width           =   4575
         Begin VB.Label LB_BombaVacuo 
            AutoSize        =   -1  'True
            Caption         =   "DESLIGADO"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   270
            Left            =   120
            TabIndex        =   151
            Top             =   240
            Width           =   1260
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Velocidade Motor:"
            Height          =   195
            Index           =   41
            Left            =   2400
            TabIndex        =   150
            Top             =   0
            Width           =   1290
         End
         Begin VB.Label LB_VelocidadeBombaVacuo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "3600 RPM"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   2400
            TabIndex        =   149
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.CommandButton BT_TesteSimulacao 
         Caption         =   "SIMULACAO"
         Height          =   495
         Left            =   -71400
         TabIndex        =   118
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton BT_LIGA 
         Caption         =   "LIGA1"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -73320
         TabIndex        =   117
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Frame Frame9 
         Caption         =   "Número de ensaios neste ciclo:"
         Height          =   3375
         Left            =   9720
         TabIndex        =   115
         Top             =   5640
         Width           =   4575
         Begin VB.Frame Frame10 
            Caption         =   "Ensaio de Estanqueidade:"
            Height          =   1695
            Left            =   120
            TabIndex        =   119
            Top             =   480
            Width           =   4335
            Begin VB.Label LB_PAP_CV 
               AutoSize        =   -1  'True
               Caption         =   "NA"
               BeginProperty Font 
                  Name            =   "Arial Black"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   225
               Left            =   3960
               TabIndex        =   147
               Top             =   1440
               Width           =   270
            End
            Begin VB.Label LB_PMP_CV 
               AutoSize        =   -1  'True
               Caption         =   "NA"
               BeginProperty Font 
                  Name            =   "Arial Black"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   225
               Left            =   3960
               TabIndex        =   146
               Top             =   1200
               Width           =   270
            End
            Begin VB.Label LB_PAP_2S 
               AutoSize        =   -1  'True
               Caption         =   "NA"
               BeginProperty Font 
                  Name            =   "Arial Black"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   225
               Left            =   3240
               TabIndex        =   145
               Top             =   1440
               Width           =   270
            End
            Begin VB.Label LB_PMP_2S 
               AutoSize        =   -1  'True
               Caption         =   "NA"
               BeginProperty Font 
                  Name            =   "Arial Black"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   225
               Left            =   3240
               TabIndex        =   144
               Top             =   1200
               Width           =   270
            End
            Begin VB.Label LB_PAP_1S 
               AutoSize        =   -1  'True
               Caption         =   "NA"
               BeginProperty Font 
                  Name            =   "Arial Black"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   225
               Left            =   2520
               TabIndex        =   143
               Top             =   1440
               Width           =   270
            End
            Begin VB.Label LB_PMP_1S 
               AutoSize        =   -1  'True
               Caption         =   "NA"
               BeginProperty Font 
                  Name            =   "Arial Black"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   225
               Left            =   2520
               TabIndex        =   142
               Top             =   1200
               Width           =   270
            End
            Begin VB.Label LB 
               AutoSize        =   -1  'True
               Caption         =   "Pneumático em Alta Pressão:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   210
               Index           =   74
               Left            =   120
               TabIndex        =   141
               Top             =   1440
               Width           =   2175
            End
            Begin VB.Label LB 
               AutoSize        =   -1  'True
               Caption         =   "Pneumático em Média Pressão:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   210
               Index           =   54
               Left            =   120
               TabIndex        =   140
               Top             =   1200
               Width           =   2340
            End
            Begin VB.Label LB_PBP_CV 
               AutoSize        =   -1  'True
               Caption         =   "NA"
               BeginProperty Font 
                  Name            =   "Arial Black"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   225
               Left            =   3960
               TabIndex        =   138
               Top             =   960
               Width           =   270
            End
            Begin VB.Label LB_HAP_CV 
               AutoSize        =   -1  'True
               Caption         =   "NA"
               BeginProperty Font 
                  Name            =   "Arial Black"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   225
               Left            =   3960
               TabIndex        =   137
               Top             =   720
               Width           =   270
            End
            Begin VB.Label LB_HMP_CV 
               AutoSize        =   -1  'True
               Caption         =   "NA"
               BeginProperty Font 
                  Name            =   "Arial Black"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   225
               Left            =   3960
               TabIndex        =   136
               Top             =   480
               Width           =   270
            End
            Begin VB.Label LB_PBP_2S 
               AutoSize        =   -1  'True
               Caption         =   "NA"
               BeginProperty Font 
                  Name            =   "Arial Black"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   225
               Left            =   3240
               TabIndex        =   135
               Top             =   960
               Width           =   270
            End
            Begin VB.Label LB_HAP_2S 
               AutoSize        =   -1  'True
               Caption         =   "NA"
               BeginProperty Font 
                  Name            =   "Arial Black"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   225
               Left            =   3240
               TabIndex        =   134
               Top             =   720
               Width           =   270
            End
            Begin VB.Label LB_HMP_2S 
               AutoSize        =   -1  'True
               Caption         =   "NA"
               BeginProperty Font 
                  Name            =   "Arial Black"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   225
               Left            =   3240
               TabIndex        =   133
               Top             =   480
               Width           =   270
            End
            Begin VB.Label LB_PBP_1S 
               AutoSize        =   -1  'True
               Caption         =   "NA"
               BeginProperty Font 
                  Name            =   "Arial Black"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   225
               Left            =   2520
               TabIndex        =   132
               Top             =   960
               Width           =   270
            End
            Begin VB.Label LB_HAP_1S 
               AutoSize        =   -1  'True
               Caption         =   "NA"
               BeginProperty Font 
                  Name            =   "Arial Black"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   225
               Left            =   2520
               TabIndex        =   131
               Top             =   720
               Width           =   270
            End
            Begin VB.Label LB_HMP_1S 
               AutoSize        =   -1  'True
               Caption         =   "NA"
               BeginProperty Font 
                  Name            =   "Arial Black"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   225
               Left            =   2520
               TabIndex        =   130
               Top             =   480
               Width           =   270
            End
            Begin VB.Label LB_HBP_CV 
               AutoSize        =   -1  'True
               Caption         =   "NA"
               BeginProperty Font 
                  Name            =   "Arial Black"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   225
               Left            =   3960
               TabIndex        =   129
               Top             =   240
               Width           =   270
            End
            Begin VB.Label LB_HBP_2S 
               AutoSize        =   -1  'True
               Caption         =   "NA"
               BeginProperty Font 
                  Name            =   "Arial Black"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   225
               Left            =   3240
               TabIndex        =   128
               Top             =   240
               Width           =   270
            End
            Begin VB.Label LB 
               AutoSize        =   -1  'True
               Caption         =   "C&&CV:"
               Height          =   195
               Index           =   62
               Left            =   3840
               TabIndex        =   127
               Top             =   0
               Width           =   450
            End
            Begin VB.Label LB 
               AutoSize        =   -1  'True
               Caption         =   "P - 2ª S:"
               Height          =   195
               Index           =   61
               Left            =   3120
               TabIndex        =   126
               Top             =   0
               Width           =   585
            End
            Begin VB.Label LB 
               AutoSize        =   -1  'True
               Caption         =   "P - 1ª S:"
               Height          =   195
               Index           =   60
               Left            =   2400
               TabIndex        =   125
               Top             =   0
               Width           =   585
            End
            Begin VB.Label LB_HBP_1S 
               AutoSize        =   -1  'True
               Caption         =   "NA"
               BeginProperty Font 
                  Name            =   "Arial Black"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   225
               Left            =   2520
               TabIndex        =   124
               Top             =   240
               Width           =   270
            End
            Begin VB.Label LB 
               AutoSize        =   -1  'True
               Caption         =   "Pneumático em Baixa Pressão:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   210
               Index           =   59
               Left            =   120
               TabIndex        =   123
               Top             =   960
               Width           =   2310
            End
            Begin VB.Label LB 
               AutoSize        =   -1  'True
               Caption         =   "Hidrostático em Alta Pressão:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   210
               Index           =   58
               Left            =   120
               TabIndex        =   122
               Top             =   720
               Width           =   2205
            End
            Begin VB.Label LB 
               AutoSize        =   -1  'True
               Caption         =   "Hidrostático em Média Pressão:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   210
               Index           =   57
               Left            =   120
               TabIndex        =   121
               Top             =   480
               Width           =   2370
            End
            Begin VB.Label LB 
               AutoSize        =   -1  'True
               Caption         =   "Hidrostático em Baixa Pressão:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   210
               Index           =   56
               Left            =   120
               TabIndex        =   120
               Top             =   240
               Width           =   2340
            End
         End
         Begin VB.Label LB_AAP 
            AutoSize        =   -1  'True
            Caption         =   "SIM"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   3360
            TabIndex        =   159
            Top             =   3000
            Width           =   330
         End
         Begin VB.Label LB_AMP 
            AutoSize        =   -1  'True
            Caption         =   "SIM"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   3360
            TabIndex        =   158
            Top             =   2760
            Width           =   330
         End
         Begin VB.Label LB_ABP 
            AutoSize        =   -1  'True
            Caption         =   "SIM"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   3360
            TabIndex        =   157
            Top             =   2520
            Width           =   330
         End
         Begin VB.Label LB_AC 
            AutoSize        =   -1  'True
            Caption         =   "SIM"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   3360
            TabIndex        =   156
            Top             =   2280
            Width           =   330
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Assinatura em Alta Pressão:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   210
            Index           =   84
            Left            =   120
            TabIndex        =   155
            Top             =   3000
            Width           =   2100
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Assinatura em Média Pressão:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   210
            Index           =   83
            Left            =   120
            TabIndex        =   154
            Top             =   2760
            Width           =   2265
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Assinatura em Baixa Pressão:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   210
            Index           =   82
            Left            =   120
            TabIndex        =   153
            Top             =   2520
            Width           =   2235
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Alívio de Cavidade do Corpo:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   210
            Index           =   81
            Left            =   120
            TabIndex        =   152
            Top             =   2280
            Width           =   2130
         End
         Begin VB.Label LB_CicloParada 
            AutoSize        =   -1  'True
            Caption         =   "SIM"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   3480
            TabIndex        =   139
            Top             =   240
            Width           =   330
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Este ciclo tem parada para ensaios:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   210
            Index           =   55
            Left            =   120
            TabIndex        =   116
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Variável Monitorada:"
         Height          =   5775
         Left            =   240
         TabIndex        =   63
         Top             =   360
         Width           =   9255
         Begin MSComctlLib.ProgressBar PB_S1 
            Height          =   375
            Left            =   6360
            TabIndex        =   64
            Top             =   240
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
         End
         Begin MSComctlLib.ProgressBar PB_S3 
            Height          =   375
            Left            =   6360
            TabIndex        =   65
            Top             =   960
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
         End
         Begin MSComctlLib.ProgressBar PB_S5 
            Height          =   375
            Left            =   6360
            TabIndex        =   66
            Top             =   1680
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
         End
         Begin MSComctlLib.ProgressBar PB_S4 
            Height          =   375
            Left            =   6360
            TabIndex        =   67
            Top             =   1320
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
         End
         Begin MSComctlLib.ProgressBar PB_S2 
            Height          =   375
            Left            =   6360
            TabIndex        =   68
            Top             =   600
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
         End
         Begin MSComctlLib.ProgressBar PB_S6 
            Height          =   375
            Left            =   6360
            TabIndex        =   69
            Top             =   2040
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
         End
         Begin MSComctlLib.ProgressBar PB_S7 
            Height          =   375
            Left            =   6360
            TabIndex        =   70
            Top             =   2400
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
         End
         Begin MSComctlLib.ProgressBar PB_S8 
            Height          =   375
            Left            =   6360
            TabIndex        =   71
            Top             =   2760
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
         End
         Begin MSComctlLib.ProgressBar PB_S9 
            Height          =   375
            Left            =   6360
            TabIndex        =   72
            Top             =   3120
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
         End
         Begin MSComctlLib.ProgressBar PB_S10 
            Height          =   375
            Left            =   6360
            TabIndex        =   73
            Top             =   3480
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
         End
         Begin MSComctlLib.ProgressBar PB_S11 
            Height          =   375
            Left            =   6360
            TabIndex        =   74
            Top             =   3840
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
         End
         Begin MSComctlLib.ProgressBar PB_S12 
            Height          =   375
            Left            =   6360
            TabIndex        =   75
            Top             =   4200
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
         End
         Begin MSComctlLib.ProgressBar PB_S13 
            Height          =   375
            Left            =   6360
            TabIndex        =   76
            Top             =   4560
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
         End
         Begin MSComctlLib.ProgressBar PB_S14 
            Height          =   375
            Left            =   6360
            TabIndex        =   77
            Top             =   4920
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
         End
         Begin MSComctlLib.ProgressBar PB_S15 
            Height          =   375
            Left            =   6360
            TabIndex        =   78
            Top             =   5280
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
         End
         Begin VB.Label LB_S15 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "1023109213,12"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   3240
            TabIndex        =   110
            Top             =   5280
            Width           =   3015
         End
         Begin VB.Label LB_S14 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "1023109213,12"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   3240
            TabIndex        =   109
            Top             =   4920
            Width           =   3015
         End
         Begin VB.Label LB_S13 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "1023109213,12"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   3240
            TabIndex        =   108
            Top             =   4560
            Width           =   3015
         End
         Begin VB.Label LB_S12 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "1023109213,12"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   3240
            TabIndex        =   107
            Top             =   4200
            Width           =   3015
         End
         Begin VB.Label LB_S11 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "1023109213,12"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   3240
            TabIndex        =   106
            Top             =   3840
            Width           =   3015
         End
         Begin VB.Label LB_S10 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "1023109213,12"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   3240
            TabIndex        =   105
            Top             =   3480
            Width           =   3015
         End
         Begin VB.Label LB_S9 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "1023109213,12"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   3240
            TabIndex        =   104
            Top             =   3120
            Width           =   3015
         End
         Begin VB.Label LB_S8 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "1023109213,12"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   3240
            TabIndex        =   103
            Top             =   2760
            Width           =   3015
         End
         Begin VB.Label LB_S7 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "1023109213,12"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   3240
            TabIndex        =   102
            Top             =   2400
            Width           =   3015
         End
         Begin VB.Label LB_S6 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "1023109213,12"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   3240
            TabIndex        =   101
            Top             =   2040
            Width           =   3015
         End
         Begin VB.Label LB_S5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "1023109213,12"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   3240
            TabIndex        =   100
            Top             =   1680
            Width           =   3015
         End
         Begin VB.Label LB_S4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "1023109213,12"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   3240
            TabIndex        =   99
            Top             =   1320
            Width           =   3015
         End
         Begin VB.Label LB_S3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "1023109213,12"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   3240
            TabIndex        =   98
            Top             =   960
            Width           =   3015
         End
         Begin VB.Label LB_S2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "1023109213,12"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   3240
            TabIndex        =   97
            Top             =   600
            Width           =   3015
         End
         Begin VB.Label LB_S1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "1023109213,12"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   3240
            TabIndex        =   96
            Top             =   240
            Width           =   3015
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "15) Vazamento na Gaxeta"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Index           =   18
            Left            =   120
            TabIndex        =   95
            Top             =   5400
            Width           =   2340
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "14) Vazamento na Junta"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Index           =   17
            Left            =   120
            TabIndex        =   94
            Top             =   5040
            Width           =   2220
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "13) Vazamento na Passagem"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Index           =   16
            Left            =   120
            TabIndex        =   93
            Top             =   4680
            Width           =   2625
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "12) Vazão"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Index           =   15
            Left            =   120
            TabIndex        =   92
            Top             =   4320
            Width           =   900
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "11) Velocidade de Acionamento"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Index           =   14
            Left            =   120
            TabIndex        =   91
            Top             =   3960
            Width           =   2955
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "10) Vibração"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Index           =   13
            Left            =   120
            TabIndex        =   90
            Top             =   3600
            Width           =   1155
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "9) Ruído"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Index           =   12
            Left            =   120
            TabIndex        =   89
            Top             =   3240
            Width           =   780
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "8) Deformação (Straingage)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Index           =   11
            Left            =   120
            TabIndex        =   88
            Top             =   2880
            Width           =   2520
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "7) Deslocamento Angular"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Index           =   10
            Left            =   120
            TabIndex        =   87
            Top             =   2520
            Width           =   2340
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "6) Deslocamento Linear"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Index           =   9
            Left            =   120
            TabIndex        =   86
            Top             =   2160
            Width           =   2175
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "5) Temperatura do Fluído"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Index           =   8
            Left            =   120
            TabIndex        =   85
            Top             =   1800
            Width           =   2355
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "4) Torque de Acionamento"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Index           =   7
            Left            =   120
            TabIndex        =   84
            Top             =   1440
            Width           =   2475
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "3) Pressão na Jusante"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Index           =   6
            Left            =   120
            TabIndex        =   83
            Top             =   1080
            Width           =   1980
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "2) Pressão no Corpo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Index           =   5
            Left            =   120
            TabIndex        =   82
            Top             =   720
            Width           =   1860
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "1) Pressão na Montante"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Index           =   4
            Left            =   120
            TabIndex        =   81
            Top             =   360
            Width           =   2145
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Resposta do Sensor (Valor):"
            Height          =   195
            Index           =   0
            Left            =   3240
            TabIndex        =   80
            Top             =   0
            Width           =   1980
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Percentual sobre o Valor:"
            Height          =   195
            Index           =   1
            Left            =   6360
            TabIndex        =   79
            Top             =   0
            Width           =   1785
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Tempos de Ciclagem:"
         Height          =   855
         Left            =   240
         TabIndex        =   55
         Top             =   6240
         Width           =   9255
         Begin VB.Label LB_TempoEnsaio 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "20h:10m:01s"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   255
            Left            =   7200
            TabIndex        =   160
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label LB_TempoEspera 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "20h:10m:01s"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   255
            Left            =   4800
            TabIndex        =   62
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label LB_TempoCiclo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "20h:10m:01s"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   255
            Left            =   2400
            TabIndex        =   61
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label LB_TempoTotalTeste 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "20h:10m:01s"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Tempo de Espera:"
            Height          =   195
            Index           =   20
            Left            =   4800
            TabIndex        =   59
            Top             =   240
            Width           =   1305
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Tempo do Ciclo:"
            Height          =   195
            Index           =   19
            Left            =   2400
            TabIndex        =   58
            Top             =   240
            Width           =   1155
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Tempo de Ensaio:"
            Height          =   195
            Index           =   3
            Left            =   7200
            TabIndex        =   57
            Top             =   240
            Width           =   1290
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Tempo Total de Teste:"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   56
            Top             =   240
            Width           =   1620
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Torques especiais monitorados neste ciclo (N.m):"
         Height          =   855
         Left            =   240
         TabIndex        =   42
         Top             =   8160
         Width           =   9255
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "TRFQ:"
            Height          =   195
            Index           =   30
            Left            =   7800
            TabIndex        =   54
            ToolTipText     =   "Torque real de fechamento na quebra de movimento"
            Top             =   240
            Width           =   480
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "TRFS:"
            Height          =   195
            Index           =   29
            Left            =   4800
            TabIndex        =   53
            ToolTipText     =   "Torque real de fechamento sem diferencial de pressão"
            Top             =   240
            Width           =   465
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "TRAS:"
            Height          =   195
            Index           =   27
            Left            =   3240
            TabIndex        =   52
            ToolTipText     =   "Torque real da abertura sem diferencial de pressão"
            Top             =   240
            Width           =   480
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "TRAC:"
            Height          =   195
            Index           =   26
            Left            =   1680
            TabIndex        =   51
            ToolTipText     =   "Torque real da abertura com diferencial de pressão"
            Top             =   240
            Width           =   480
         End
         Begin VB.Label LB_TRFQ 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "250,00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   255
            Left            =   7800
            TabIndex        =   50
            ToolTipText     =   "Torque real de fechamento na quebra de movimento"
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label LB_TRFC 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "250,00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   255
            Left            =   6360
            TabIndex        =   49
            ToolTipText     =   "Torque real do fechamento com diferencial de pressão"
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label LB_TRFS 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "250,00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   255
            Left            =   4800
            TabIndex        =   48
            ToolTipText     =   "Torque real de fechamento sem diferencial de pressão"
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label LB_TRAS 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "250,00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   255
            Left            =   3240
            TabIndex        =   47
            ToolTipText     =   "Torque real da abertura sem diferencial de pressão"
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label LB_TRAC 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "250,00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   255
            Left            =   1680
            TabIndex        =   46
            ToolTipText     =   "Torque real da abertura com diferencial de pressão"
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label LB_TRAQ 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "250,00"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   255
            Left            =   120
            TabIndex        =   45
            ToolTipText     =   "Torque real de abertura na quebra de movimento"
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "TRFC:"
            Height          =   195
            Index           =   28
            Left            =   6360
            TabIndex        =   44
            ToolTipText     =   "Torque real do fechamento com diferencial de pressão"
            Top             =   240
            Width           =   465
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "TRAQ:"
            Height          =   195
            Index           =   25
            Left            =   120
            TabIndex        =   43
            ToolTipText     =   "Torque real de abertura na quebra de movimento"
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.PictureBox PICB 
         Height          =   8535
         Left            =   -74880
         ScaleHeight     =   8475
         ScaleWidth      =   14235
         TabIndex        =   41
         Top             =   480
         Width           =   14295
      End
      Begin VB.Timer TIMER_SMD_AUX 
         Interval        =   1
         Left            =   5880
         Top             =   0
      End
      Begin VB.Timer TIMER_SMD 
         Interval        =   1
         Left            =   5400
         Top             =   0
      End
      Begin VB.Timer TIMER_SIMULACAO 
         Left            =   4920
         Top             =   0
      End
      Begin VB.Frame Frame5 
         Caption         =   "Bomba Fluído:"
         Height          =   615
         Left            =   9720
         TabIndex        =   37
         Top             =   360
         Width           =   4575
         Begin VB.Label LB_VelocidadeBombaFluido 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "3600 RPM"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   2400
            TabIndex        =   40
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Velocidade Motor:"
            Height          =   195
            Index           =   34
            Left            =   2400
            TabIndex        =   39
            Top             =   0
            Width           =   1290
         End
         Begin VB.Label LB_BombaFluido 
            AutoSize        =   -1  'True
            Caption         =   "DESLIGADO"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   270
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   1260
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "  ENTRADA:       VÁLVULA DA MÁQUINA:          SAÍDA:"
         Height          =   2775
         Left            =   9720
         TabIndex        =   15
         Top             =   1800
         Width           =   4575
         Begin VB.Label LB 
            Alignment       =   2  'Center
            Caption         =   "Montante do Protótipo"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   31
            Left            =   1200
            TabIndex        =   36
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label LB 
            Alignment       =   2  'Center
            Caption         =   "Jusante do Protótipo"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   33
            Left            =   1200
            TabIndex        =   35
            Top             =   600
            Width           =   2175
         End
         Begin VB.Label LB 
            Alignment       =   2  'Center
            Caption         =   "Tanque de Água Fria"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   35
            Left            =   1200
            TabIndex        =   34
            Top             =   960
            Width           =   2115
         End
         Begin VB.Label LB 
            Alignment       =   2  'Center
            Caption         =   "Tanque de Água Quente"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   36
            Left            =   1200
            TabIndex        =   33
            Top             =   1320
            Width           =   2175
         End
         Begin VB.Label LB 
            Alignment       =   2  'Center
            Caption         =   "Tanque de Nitrogênio"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   37
            Left            =   1200
            TabIndex        =   32
            Top             =   1680
            Width           =   2115
         End
         Begin VB.Label LB 
            Alignment       =   2  'Center
            Caption         =   "Tanque de Retorno"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   38
            Left            =   1200
            TabIndex        =   31
            Top             =   2040
            Width           =   2115
         End
         Begin VB.Label LB 
            Alignment       =   2  'Center
            Caption         =   "Bomba de Vácuo"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   39
            Left            =   1200
            TabIndex        =   30
            Top             =   2400
            Width           =   2115
         End
         Begin VB.Label LB_VM_Montante_Entrada 
            Alignment       =   2  'Center
            Caption         =   "ABERTA"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   270
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   1065
         End
         Begin VB.Label LB_VM_Montante_Saida 
            Alignment       =   2  'Center
            Caption         =   "FECHADA"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   270
            Left            =   3360
            TabIndex        =   28
            Top             =   240
            Width           =   1065
         End
         Begin VB.Label LB_VM_Jusante_Entrada 
            Alignment       =   2  'Center
            Caption         =   "ABERTA"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   270
            Left            =   120
            TabIndex        =   27
            Top             =   600
            Width           =   1065
         End
         Begin VB.Label LB_VM_AguaFria_Entrada 
            Alignment       =   2  'Center
            Caption         =   "ABERTA"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   270
            Left            =   120
            TabIndex        =   26
            Top             =   960
            Width           =   1065
         End
         Begin VB.Label LB_VM_AguaQuente_Entrada 
            Alignment       =   2  'Center
            Caption         =   "ABERTA"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   270
            Left            =   120
            TabIndex        =   25
            Top             =   1320
            Width           =   1065
         End
         Begin VB.Label LB_VM_Nitrogenio_Entrada 
            Alignment       =   2  'Center
            Caption         =   "ABERTA"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   270
            Left            =   120
            TabIndex        =   24
            Top             =   1680
            Width           =   1065
         End
         Begin VB.Label LB_VM_TanqueRetorno_Entrada 
            Alignment       =   2  'Center
            Caption         =   "ABERTA"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   270
            Left            =   120
            TabIndex        =   23
            Top             =   2040
            Width           =   1065
         End
         Begin VB.Label LB_VM_BombaVacuo_Entrada 
            Alignment       =   2  'Center
            Caption         =   "ABERTA"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   270
            Left            =   120
            TabIndex        =   22
            Top             =   2400
            Width           =   1065
         End
         Begin VB.Label LB_VM_Jusante_Saida 
            Alignment       =   2  'Center
            Caption         =   "FECHADA"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   270
            Left            =   3360
            TabIndex        =   21
            Top             =   600
            Width           =   1065
         End
         Begin VB.Label LB_VM_AguaFria_Saida 
            Alignment       =   2  'Center
            Caption         =   "FECHADA"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   270
            Left            =   3360
            TabIndex        =   20
            Top             =   960
            Width           =   1065
         End
         Begin VB.Label LB_VM_AguaQuente_Saida 
            Alignment       =   2  'Center
            Caption         =   "FECHADA"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   270
            Left            =   3360
            TabIndex        =   19
            Top             =   1320
            Width           =   1065
         End
         Begin VB.Label LB_VM_Nitrogenio_Saida 
            Alignment       =   2  'Center
            Caption         =   "FECHADA"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   270
            Left            =   3360
            TabIndex        =   18
            Top             =   1680
            Width           =   1065
         End
         Begin VB.Label LB_VM_TanqueRetorno_Saida 
            Alignment       =   2  'Center
            Caption         =   "FECHADA"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   270
            Left            =   3360
            TabIndex        =   17
            Top             =   2040
            Width           =   1065
         End
         Begin VB.Label LB_VM_BombaVacuo_Saida 
            Alignment       =   2  'Center
            Caption         =   "FECHADA"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   270
            Left            =   3360
            TabIndex        =   16
            Top             =   2400
            Width           =   1065
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Ensaio que está sendo executado na máquina:"
         Height          =   855
         Left            =   9720
         TabIndex        =   10
         Top             =   4680
         Width           =   4575
         Begin VB.Label LB_EnsaioBP 
            Alignment       =   2  'Center
            Caption         =   "Baixa Pressão"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   600
            Width           =   1395
         End
         Begin VB.Label LB_Ensaio 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Ensaio Hidrostático"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   4335
         End
         Begin VB.Label LB_EnsaioMP 
            Alignment       =   2  'Center
            Caption         =   "Média Pressão"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1560
            TabIndex        =   12
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label LB_EnsaioAP 
            Alignment       =   1  'Right Justify
            Caption         =   "Alta Pressão"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   3120
            TabIndex        =   11
            Top             =   600
            Width           =   1335
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Informações sobre o ciclo:"
         Height          =   855
         Left            =   240
         TabIndex        =   1
         Top             =   7200
         Width           =   9255
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "T.O. neste Ciclo:"
            Height          =   195
            Index           =   50
            Left            =   7560
            TabIndex        =   114
            Top             =   240
            Width           =   1185
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Situação da Válvula:"
            Height          =   195
            Index           =   48
            Left            =   5400
            TabIndex        =   113
            Top             =   240
            Width           =   1470
         End
         Begin VB.Label LB_TorqueValvula 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Sobrecarga"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   7560
            TabIndex        =   112
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label LB_SituacaoValvula 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Parada Aberta"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   5400
            TabIndex        =   111
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label LB_Ciclos 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "5000"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   480
            Width           =   735
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Ciclo nº:"
            Height          =   195
            Index           =   23
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   585
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "T.Assin.:"
            Height          =   195
            Index           =   24
            Left            =   960
            TabIndex        =   7
            Top             =   240
            Width           =   615
         End
         Begin VB.Label LB_TotalAssinaturas 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "5000"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   960
            TabIndex        =   6
            Top             =   480
            Width           =   735
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "T.Ensaios:"
            Height          =   195
            Index           =   42
            Left            =   1800
            TabIndex        =   5
            Top             =   240
            Width           =   750
         End
         Begin VB.Label LB_TotalEnsaios 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "5000"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   1800
            TabIndex        =   4
            Top             =   480
            Width           =   735
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Situação deste Ciclo:"
            Height          =   195
            Index           =   43
            Left            =   2640
            TabIndex        =   3
            Top             =   240
            Width           =   1500
         End
         Begin VB.Label LB_SituacaoCiclo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "Aguardando Ensaio"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   2640
            TabIndex        =   2
            Top             =   480
            Width           =   2655
         End
      End
   End
End
Attribute VB_Name = "Tela_Simulacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VALOR_CONTAGEM, VAR_SENSOR As Integer
Dim VAR_INTERVALO_S1, VAR_INTERVALO_S2, VAR_INTERVALO_S3, VAR_INTERVALO_S4, _
    VAR_INTERVALO_S5, VAR_INTERVALO_S6, VAR_INTERVALO_S7, VAR_INTERVALO_S8, _
    VAR_INTERVALO_S9, VAR_INTERVALO_S10, VAR_INTERVALO_S11, VAR_INTERVALO_S12, _
    VAR_INTERVALO_S13, VAR_INTERVALO_S14, VAR_INTERVALO_S15 As Integer
Dim VAR_LEITURASENSOR, VAR_LEITURASENSOR_AUXILIAR As Integer
Dim VAR_TEMPOTOTAL, VAR_TEMPOCICLO, VAR_TEMPOESPERA, VAR_TEMPOENSAIO As Long
Dim VAR_MEDE_TEMPOTOTAL, VAR_MEDE_TEMPOCICLO, VAR_MEDE_TEMPOESPERA, VAR_MEDE_TEMPOENSAIO As Boolean

Private Sub BT_TesteSimulacao_Click()
    TIMER_SIMULACAO.Enabled = True
    BT_TesteSimulacao.Enabled = False
End Sub
Private Sub Form_Load()
    ST.Top = 100
    ST.Left = (Screen.Width / 2) - (ST.Width / 2)
    VAR_Tela = "Tela_Simulacao"
    LimpaGrafico
    VALOR_CONTAGEM = 0
    'Carrega valores dos TIMER
    VAR_MEDE_TEMPOTOTAL = False
    VAR_MEDE_TEMPOCICLO = False
    VAR_MEDE_TEMPOESPERA = False
    VAR_MEDE_TEMPOENSAIO = False
    VAR_TEMPOTOTAL = 0
    VAR_TEMPOCICLO = 0
    VAR_TEMPOESPERA = 0
    VAR_TEMPOENSAIO = 0
    TIMER_TEMPOS.Enabled = True
    TIMER_TEMPOS.Interval = 1000 'muda labels de tempo a cada segundo
    TIMER_SIMULACAO.Interval = 200
    TIMER_SIMULACAO.Enabled = False
    VAR_LEITURASENSOR = 1
    TIMER_SMD_AUX.Enabled = False
    'Limpa campos dos sensores
    LimpaDadosSensores
End Sub
Private Sub TIMER_SIMULACAO_Timer()
    'Carrega modo EPP
    'EscrevePorta &H37A, 32
    'Escreve informações na barra de status
    Tela_Principal.ST.SimpleText = "Passo número: " & VALOR_CONTAGEM
    VALOR_CONTAGEM = VALOR_CONTAGEM + 1
    'Le valor do sensor
    LeSensor (VAR_LEITURASENSOR)
    'verifica número máximo de sensores na máquina
    If VAR_LEITURASENSOR = 9 Then
        VAR_LEITURASENSOR = 1
    End If
    'verifica final da simulacao
    If VALOR_CONTAGEM = 100 Then
        TIMER_SIMULACAO.Enabled = False
        VALOR_CONTAGEM = 0
        BT_TesteSimulacao.Enabled = True
    End If
End Sub
Private Sub TIMER_SMD_AUX_Timer()
    VAR_LEITURASENSOR = VAR_LEITURASENSOR + 1
    TIMER_SMD_AUX.Enabled = False
End Sub
Private Sub TIMER_SMD_Timer()
    If TIMER_SMD_AUX.Enabled = False Then
        LeSensor (VAR_LEITURASENSOR)
    End If
    'verifica número máximo de sensores na máquina
    If VAR_LEITURASENSOR = 9 Then
        VAR_LEITURASENSOR = 1
    Else
       'TIMER AUXILIAR PARA RETARDAR TEMPO DE LEITURA DA PORTA DO PC (PLACA É MAIS DEVAGAR)
       TIMER_SMD_AUX.Enabled = True
    End If
End Sub
Private Sub TIMER_TEMPOS_Timer()
    'verifica tempo total de teste
    If VAR_MEDE_TEMPOTOTAL = True Then
        VAR_TEMPOTOTAL = VAR_TEMPOTOTAL + 1
        LB_TempoTotalTeste.Caption = TempoString(VAR_TEMPOTOTAL)
    End If
    'verifica tempo de ciclo
    If VAR_MEDE_TEMPOCICLO = True Then
        VAR_TEMPOCICLO = VAR_TEMPOCICLO + 1
        LB_TempoCiclo.Caption = TempoString(VAR_TEMPOCICLO)
    End If
    'verifica tempo de espera
    If VAR_MEDE_TEMPOESPERA = True Then
        VAR_TEMPOESPERA = VAR_TEMPOESPERA + 1
        LB_TempoEspera.Caption = TempoString(VAR_TEMPOESPERA)
    End If
    'verifica tempo de ensaio
    If VAR_MEDE_TEMPOENSAIO = True Then
        VAR_TEMPOENSAIO = VAR_TEMPOENSAIO + 1
        LB_TempoEnsaio.Caption = TempoString(VAR_TEMPOENSAIO)
    End If
End Sub


'*************************
' FUNCOES DESTA TELA
'*************************
Public Sub Fechar()
    Unload Tela_Simulacao
End Sub
Public Sub Salvar()
    VAR_PodeSalvar = True
End Sub
Private Sub LeSensor(VARSUB_Sensor As Integer)
    If VARSUB_Sensor = 1 Then
        'Verifica sensor 1
        EscrevePorta &H37A, 2
        VAR_SENSOR = LePorta(&H37C) * VAR_INTERVALO_S1
        LB_S1.Caption = VAR_SENSOR & " PSI"
        PB_S1.Value = VAR_SENSOR
    ElseIf VARSUB_Sensor = 2 Then
        'Verifica sensor 2
        EscrevePorta &H37A, 10
        VAR_SENSOR = LePorta(&H37C) * VAR_INTERVALO_S2
        LB_S2.Caption = VAR_SENSOR & " PSI"
        PB_S2.Value = VAR_SENSOR
    ElseIf VARSUB_Sensor = 3 Then
        'Verifica sensor 3
        EscrevePorta &H37A, 8
        VAR_SENSOR = LePorta(&H37C) * VAR_INTERVALO_S3
        LB_S3.Caption = VAR_SENSOR & " PSI"
        PB_S3.Value = VAR_SENSOR
    ElseIf VARSUB_Sensor = 4 Then
        'Verifica sensor 4
        EscrevePorta &H37A, 1
        VAR_SENSOR = LePorta(&H37C) * VAR_INTERVALO_S4
        LB_S4.Caption = VAR_SENSOR & " N.m"
        PB_S4.Value = VAR_SENSOR
    ElseIf VARSUB_Sensor = 5 Then
        'Verifica sensor 5
        EscrevePorta &H37A, 14
        VAR_SENSOR = LePorta(&H37C) * VAR_INTERVALO_S5
        LB_S5.Caption = VAR_SENSOR & " ºC"
        PB_S5.Value = VAR_SENSOR
    ElseIf VARSUB_Sensor = 6 Then
        'Verifica sensor 6
        EscrevePorta &H37A, 6
        VAR_SENSOR = LePorta(&H37C) * VAR_INTERVALO_S6
        LB_S6.Caption = VAR_SENSOR & " mm"
        PB_S6.Value = VAR_SENSOR
    ElseIf VARSUB_Sensor = 7 Then
        'Verifica sensor 7
        EscrevePorta &H37A, 12
        VAR_SENSOR = LePorta(&H37C) * VAR_INTERVALO_S7
        LB_S7.Caption = VAR_SENSOR & " Graus"
        PB_S7.Value = VAR_SENSOR
    ElseIf VARSUB_Sensor = 8 Then
        'Verifica sensor 8
        EscrevePorta &H37A, 4
        VAR_SENSOR = LePorta(&H37C) * VAR_INTERVALO_S8
        LB_S8.Caption = VAR_SENSOR & " KKK"
        PB_S8.Value = VAR_SENSOR
    End If
End Sub
Private Sub LimpaDadosSensores()
    'Limpa valores dos sensores
    LB_S1.Caption = ""
    LB_S2.Caption = ""
    LB_S3.Caption = ""
    LB_S4.Caption = ""
    LB_S5.Caption = ""
    LB_S6.Caption = ""
    LB_S7.Caption = ""
    LB_S8.Caption = ""
    LB_S9.Caption = ""
    LB_S10.Caption = ""
    LB_S11.Caption = ""
    LB_S12.Caption = ""
    LB_S13.Caption = ""
    LB_S14.Caption = ""
    LB_S15.Caption = ""
    'Carrega valores minimos das barras de progresso dos sensores
    PB_S1.Min = LeINI(VAR_ArquivoINI, "Sensor1", "ValMin")
    PB_S2.Min = LeINI(VAR_ArquivoINI, "Sensor2", "ValMin")
    PB_S3.Min = LeINI(VAR_ArquivoINI, "Sensor3", "ValMin")
    PB_S4.Min = LeINI(VAR_ArquivoINI, "Sensor4", "ValMin")
    PB_S5.Min = LeINI(VAR_ArquivoINI, "Sensor5", "ValMin")
    PB_S6.Min = LeINI(VAR_ArquivoINI, "Sensor6", "ValMin")
    PB_S7.Min = LeINI(VAR_ArquivoINI, "Sensor7", "ValMin")
    PB_S8.Min = LeINI(VAR_ArquivoINI, "Sensor8", "ValMin")
    PB_S9.Min = LeINI(VAR_ArquivoINI, "Sensor9", "ValMin")
    PB_S10.Min = LeINI(VAR_ArquivoINI, "Sensor10", "ValMin")
    PB_S11.Min = LeINI(VAR_ArquivoINI, "Sensor11", "ValMin")
    PB_S12.Min = LeINI(VAR_ArquivoINI, "Sensor12", "ValMin")
    PB_S13.Min = LeINI(VAR_ArquivoINI, "Sensor13", "ValMin")
    PB_S14.Min = LeINI(VAR_ArquivoINI, "Sensor14", "ValMin")
    PB_S15.Min = LeINI(VAR_ArquivoINI, "Sensor15", "ValMin")
    'Carrega valores maximos das barras de progresso dos sensores
    PB_S1.Max = LeINI(VAR_ArquivoINI, "Sensor1", "ValMax")
    PB_S2.Max = LeINI(VAR_ArquivoINI, "Sensor2", "ValMax")
    PB_S3.Max = LeINI(VAR_ArquivoINI, "Sensor3", "ValMax")
    PB_S4.Max = LeINI(VAR_ArquivoINI, "Sensor4", "ValMax")
    PB_S5.Max = LeINI(VAR_ArquivoINI, "Sensor5", "ValMax")
    PB_S6.Max = LeINI(VAR_ArquivoINI, "Sensor6", "ValMax")
    PB_S7.Max = LeINI(VAR_ArquivoINI, "Sensor7", "ValMax")
    PB_S8.Max = LeINI(VAR_ArquivoINI, "Sensor8", "ValMax")
    PB_S9.Max = LeINI(VAR_ArquivoINI, "Sensor9", "ValMax")
    PB_S10.Max = LeINI(VAR_ArquivoINI, "Sensor10", "ValMax")
    PB_S11.Max = LeINI(VAR_ArquivoINI, "Sensor11", "ValMax")
    PB_S12.Max = LeINI(VAR_ArquivoINI, "Sensor12", "ValMax")
    PB_S13.Max = LeINI(VAR_ArquivoINI, "Sensor13", "ValMax")
    PB_S14.Max = LeINI(VAR_ArquivoINI, "Sensor14", "ValMax")
    PB_S15.Max = LeINI(VAR_ArquivoINI, "Sensor15", "ValMax")
    'Zera barras de progresso dos sensores
    PB_S1.Value = 0
    PB_S2.Value = 0
    PB_S3.Value = 0
    PB_S4.Value = 0
    PB_S5.Value = 0
    PB_S6.Value = 0
    PB_S7.Value = 0
    PB_S8.Value = 0
    PB_S9.Value = 0
    PB_S10.Value = 0
    PB_S11.Value = 0
    PB_S12.Value = 0
    PB_S13.Value = 0
    PB_S14.Value = 0
    PB_S15.Value = 0
    'Carrega Intervalos dos Sensores
    VAR_INTERVALO_S1 = (PB_S1.Max - PB_S1.Min) / VAR_RESOLUCAO
    VAR_INTERVALO_S2 = (PB_S2.Max - PB_S2.Min) / VAR_RESOLUCAO
    VAR_INTERVALO_S3 = (PB_S3.Max - PB_S3.Min) / VAR_RESOLUCAO
    VAR_INTERVALO_S4 = (PB_S4.Max - PB_S4.Min) / VAR_RESOLUCAO
    VAR_INTERVALO_S5 = (PB_S5.Max - PB_S5.Min) / VAR_RESOLUCAO
    VAR_INTERVALO_S6 = (PB_S6.Max - PB_S6.Min) / VAR_RESOLUCAO
    VAR_INTERVALO_S7 = (PB_S7.Max - PB_S7.Min) / VAR_RESOLUCAO
    VAR_INTERVALO_S8 = (PB_S8.Max - PB_S8.Min) / VAR_RESOLUCAO
    VAR_INTERVALO_S9 = (PB_S9.Max - PB_S9.Min) / VAR_RESOLUCAO
    VAR_INTERVALO_S10 = (PB_S10.Max - PB_S10.Min) / VAR_RESOLUCAO
    VAR_INTERVALO_S11 = (PB_S11.Max - PB_S11.Min) / VAR_RESOLUCAO
    VAR_INTERVALO_S12 = (PB_S12.Max - PB_S12.Min) / VAR_RESOLUCAO
    VAR_INTERVALO_S13 = (PB_S13.Max - PB_S13.Min) / VAR_RESOLUCAO
    VAR_INTERVALO_S14 = (PB_S14.Max - PB_S14.Min) / VAR_RESOLUCAO
    VAR_INTERVALO_S15 = (PB_S15.Max - PB_S15.Min) / VAR_RESOLUCAO
    'Contadores de Tempo
    LB_TempoTotalTeste.Caption = TempoString(0)
    LB_TempoCiclo.Caption = TempoString(0)
    LB_TempoEspera.Caption = TempoString(0)
    LB_TempoEnsaio.Caption = TempoString(0)
    'informacoes sobre o ciclo
    LB_Ciclos.Caption = "0"
    LB_TotalAssinaturas.Caption = "0"
    LB_TotalEnsaios.Caption = "0"
    LB_SituacaoCiclo.Caption = "Parado"
    LB_SituacaoValvula.Caption = "-"
    LB_TorqueValvula.Caption = "-"
    'Torques
    LB_TRAQ.Caption = "-"
    LB_TRAC.Caption = "-"
    LB_TRAS.Caption = "-"
    LB_TRFS.Caption = "-"
    LB_TRFC.Caption = "-"
    LB_TRFQ.Caption = "-"
    'desliga bombas
    BombaFluido False
    BombaVacuo False
    'fecha valvulas maquina
    VM_Montante "E", False
    VM_Montante "S", False
    VM_Jusante "E", False
    VM_Jusante "S", False
    VM_AguaFria "E", False
    VM_AguaFria "S", False
    VM_AguaQuente "E", False
    VM_AguaQuente "S", False
    VM_Nitrogenio "E", False
    VM_Nitrogenio "S", False
    VM_TanqueRetorno "E", False
    VM_TanqueRetorno "S", False
    VM_BombaVacuo "E", False
    VM_BombaVacuo "S", False
End Sub
Private Sub LimpaGrafico()
    With PICB
        .Cls
        .CurrentX = 100
        .CurrentY = 100
        ' .Line -100, 2000
    End With
End Sub
Private Sub BombaFluido(Ligado As Boolean, Optional Velocidade As Integer)
    If Ligado = False Then
        LB_BombaFluido.Caption = "DESLIGADO"
        LB_BombaFluido.ForeColor = &HFF&
        LB_VelocidadeBombaFluido.Caption = "-"
    Else
        LB_BombaFluido.Caption = "LIGADO"
        LB_BombaFluido.ForeColor = &HC000&
        LB_VelocidadeBombaFluido.Caption = Trim(Str(Velocidade)) & " RPM"
    End If
End Sub
Private Sub BombaVacuo(Ligado As Boolean, Optional Velocidade As Integer)
    If Ligado = False Then
        LB_BombaVacuo.Caption = "DESLIGADO"
        LB_BombaVacuo.ForeColor = &HFF&
        LB_VelocidadeBombaVacuo.Caption = "-"
    Else
        LB_BombaVacuo.Caption = "LIGADO"
        LB_BombaVacuo.ForeColor = &HC000&
        LB_VelocidadeBombaVacuo.Caption = Trim(Str(Velocidade)) & " RPM"
    End If
End Sub
Private Sub VM_Montante(EntradaSaida As String, Ligado As Boolean)
    If EntradaSaida = "E" Then 'ENTRADA
        If Ligado = True Then
            LB_VM_Montante_Entrada.Caption = "ABERTA"
            LB_VM_Montante_Entrada.ForeColor = &HC000&
        Else
            LB_VM_Montante_Entrada.Caption = "FECHADA"
            LB_VM_Montante_Entrada.ForeColor = &HFF&
        End If
    Else 'SAIDA
        If Ligado = True Then
            LB_VM_Montante_Saida.Caption = "ABERTA"
            LB_VM_Montante_Saida.ForeColor = &HC000&
        Else
            LB_VM_Montante_Saida.Caption = "FECHADA"
            LB_VM_Montante_Saida.ForeColor = &HFF&
        End If
    End If
End Sub
Private Sub VM_Jusante(EntradaSaida As String, Ligado As Boolean)
    If EntradaSaida = "E" Then 'ENTRADA
        If Ligado = True Then
            LB_VM_Jusante_Entrada.Caption = "ABERTA"
            LB_VM_Jusante_Entrada.ForeColor = &HC000&
        Else
            LB_VM_Jusante_Entrada.Caption = "FECHADA"
            LB_VM_Jusante_Entrada.ForeColor = &HFF&
        End If
    Else 'SAIDA
        If Ligado = True Then
            LB_VM_Jusante_Saida.Caption = "ABERTA"
            LB_VM_Jusante_Saida.ForeColor = &HC000&
        Else
            LB_VM_Jusante_Saida.Caption = "FECHADA"
            LB_VM_Jusante_Saida.ForeColor = &HFF&
        End If
    End If
End Sub
Private Sub VM_AguaFria(EntradaSaida As String, Ligado As Boolean)
    If EntradaSaida = "E" Then 'ENTRADA
        If Ligado = True Then
            LB_VM_AguaFria_Entrada.Caption = "ABERTA"
            LB_VM_AguaFria_Entrada.ForeColor = &HC000&
        Else
            LB_VM_AguaFria_Entrada.Caption = "FECHADA"
            LB_VM_AguaFria_Entrada.ForeColor = &HFF&
        End If
    Else 'SAIDA
        If Ligado = True Then
            LB_VM_AguaFria_Saida.Caption = "ABERTA"
            LB_VM_AguaFria_Saida.ForeColor = &HC000&
        Else
            LB_VM_AguaFria_Saida.Caption = "FECHADA"
            LB_VM_AguaFria_Saida.ForeColor = &HFF&
        End If
    End If
End Sub
Private Sub VM_AguaQuente(EntradaSaida As String, Ligado As Boolean)
    If EntradaSaida = "E" Then 'ENTRADA
        If Ligado = True Then
            LB_VM_AguaQuente_Entrada.Caption = "ABERTA"
            LB_VM_AguaQuente_Entrada.ForeColor = &HC000&
        Else
            LB_VM_AguaQuente_Entrada.Caption = "FECHADA"
            LB_VM_AguaQuente_Entrada.ForeColor = &HFF&
        End If
    Else 'SAIDA
        If Ligado = True Then
            LB_VM_AguaQuente_Saida.Caption = "ABERTA"
            LB_VM_AguaQuente_Saida.ForeColor = &HC000&
        Else
            LB_VM_AguaQuente_Saida.Caption = "FECHADA"
            LB_VM_AguaQuente_Saida.ForeColor = &HFF&
        End If
    End If
End Sub
Private Sub VM_Nitrogenio(EntradaSaida As String, Ligado As Boolean)
    If EntradaSaida = "E" Then 'ENTRADA
        If Ligado = True Then
            LB_VM_Nitrogenio_Entrada.Caption = "ABERTA"
            LB_VM_Nitrogenio_Entrada.ForeColor = &HC000&
        Else
            LB_VM_Nitrogenio_Entrada.Caption = "FECHADA"
            LB_VM_Nitrogenio_Entrada.ForeColor = &HFF&
        End If
    Else 'SAIDA
        If Ligado = True Then
            LB_VM_Nitrogenio_Saida.Caption = "ABERTA"
            LB_VM_Nitrogenio_Saida.ForeColor = &HC000&
        Else
            LB_VM_Nitrogenio_Saida.Caption = "FECHADA"
            LB_VM_Nitrogenio_Saida.ForeColor = &HFF&
        End If
    End If
End Sub
Private Sub VM_TanqueRetorno(EntradaSaida As String, Ligado As Boolean)
    If EntradaSaida = "E" Then 'ENTRADA
        If Ligado = True Then
            LB_VM_TanqueRetorno_Entrada.Caption = "ABERTA"
            LB_VM_TanqueRetorno_Entrada.ForeColor = &HC000&
        Else
            LB_VM_TanqueRetorno_Entrada.Caption = "FECHADA"
            LB_VM_TanqueRetorno_Entrada.ForeColor = &HFF&
        End If
    Else 'SAIDA
        If Ligado = True Then
            LB_VM_TanqueRetorno_Saida.Caption = "ABERTA"
            LB_VM_TanqueRetorno_Saida.ForeColor = &HC000&
        Else
            LB_VM_TanqueRetorno_Saida.Caption = "FECHADA"
            LB_VM_TanqueRetorno_Saida.ForeColor = &HFF&
        End If
    End If
End Sub
Private Sub VM_BombaVacuo(EntradaSaida As String, Ligado As Boolean)
    If EntradaSaida = "E" Then 'ENTRADA
        If Ligado = True Then
            LB_VM_BombaVacuo_Entrada.Caption = "ABERTA"
            LB_VM_BombaVacuo_Entrada.ForeColor = &HC000&
        Else
            LB_VM_BombaVacuo_Entrada.Caption = "FECHADA"
            LB_VM_BombaVacuo_Entrada.ForeColor = &HFF&
        End If
    Else 'SAIDA
        If Ligado = True Then
            LB_VM_BombaVacuo_Saida.Caption = "ABERTA"
            LB_VM_BombaVacuo_Saida.ForeColor = &HC000&
        Else
            LB_VM_BombaVacuo_Saida.Caption = "FECHADA"
            LB_VM_BombaVacuo_Saida.ForeColor = &HFF&
        End If
    End If
End Sub

