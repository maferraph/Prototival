VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Tela_Valvula 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   7020
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7020
   ScaleWidth      =   8520
   WindowState     =   2  'Maximized
   Begin VB.Frame FR 
      BorderStyle     =   0  'None
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      Begin VB.Frame Frame3 
         Caption         =   "Ensaios que serão executados nesta válvula (na sequência abaixo):"
         Height          =   2775
         Left            =   0
         TabIndex        =   31
         Top             =   2760
         Width           =   7935
         Begin VB.Frame Frame4 
            Caption         =   "Ciclagem"
            Height          =   495
            Left            =   360
            TabIndex        =   44
            Top             =   2040
            Width           =   7455
            Begin VB.OptionButton Option3 
               Caption         =   "Slam-test (válvulas auto-operadas)"
               Height          =   195
               Left            =   4560
               TabIndex        =   47
               Top             =   240
               Width           =   2775
            End
            Begin VB.OptionButton Option2 
               Caption         =   "Ciclagem de torque (válvulas com volante)"
               Height          =   195
               Left            =   1200
               TabIndex        =   46
               Top             =   240
               Width           =   3375
            End
            Begin VB.OptionButton Option1 
               Caption         =   "Nenhuma"
               Height          =   195
               Left            =   120
               TabIndex        =   45
               Top             =   240
               Width           =   1215
            End
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Assinatura de torque da válvula"
            Height          =   255
            Left            =   360
            TabIndex        =   41
            Top             =   1800
            Width           =   7455
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Alívio de cavidade do corpo (válvulas com alívio como p.ex. esferas)"
            Height          =   255
            Left            =   360
            TabIndex        =   39
            Top             =   1440
            Width           =   7455
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Estanqueidade na passagem (segunda sede) (válvulas bidirecionais)"
            Height          =   255
            Left            =   360
            TabIndex        =   37
            Top             =   1080
            Width           =   7455
         End
         Begin VB.CheckBox Check10 
            Caption         =   "Estanqueidade na passagem (primeira sede) (válvulas unidirecionais e bidirecionais)"
            Height          =   255
            Left            =   360
            TabIndex        =   33
            Top             =   720
            Width           =   7455
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Coeficiente de Vazão"
            Height          =   255
            Left            =   360
            TabIndex        =   32
            Top             =   360
            Width           =   2415
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "6)"
            Height          =   195
            Index           =   19
            Left            =   120
            TabIndex        =   43
            ToolTipText     =   "Pressão Máxima de Trabalho"
            Top             =   2160
            Width           =   135
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "5)"
            Height          =   195
            Index           =   18
            Left            =   120
            TabIndex        =   42
            ToolTipText     =   "Pressão Máxima de Trabalho"
            Top             =   1800
            Width           =   135
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "4)"
            Height          =   195
            Index           =   12
            Left            =   120
            TabIndex        =   40
            ToolTipText     =   "Pressão Máxima de Trabalho"
            Top             =   1440
            Width           =   135
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "3)"
            Height          =   195
            Index           =   11
            Left            =   120
            TabIndex        =   38
            ToolTipText     =   "Pressão Máxima de Trabalho"
            Top             =   1080
            Width           =   135
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "2)"
            Height          =   195
            Index           =   10
            Left            =   120
            TabIndex        =   36
            ToolTipText     =   "Pressão Máxima de Trabalho"
            Top             =   720
            Width           =   135
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "1)"
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   35
            ToolTipText     =   "Pressão Máxima de Trabalho"
            Top             =   360
            Width           =   135
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Dimensionamento da válvula para simulação:"
         Height          =   975
         Left            =   0
         TabIndex        =   20
         Top             =   1680
         Width           =   7935
         Begin VB.TextBox Text18 
            Height          =   285
            Left            =   120
            TabIndex        =   25
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox Text17 
            Height          =   285
            Left            =   1680
            TabIndex        =   24
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox Text16 
            Height          =   285
            Left            =   3240
            TabIndex        =   23
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox Text15 
            Height          =   285
            Left            =   4800
            TabIndex        =   22
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox Text14 
            Height          =   285
            Left            =   6360
            TabIndex        =   21
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "PMT (PSI):"
            Height          =   195
            Index           =   17
            Left            =   120
            TabIndex        =   30
            ToolTipText     =   "Pressão Máxima de Trabalho"
            Top             =   240
            Width           =   780
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "TMT (ºC):"
            Height          =   195
            Index           =   16
            Left            =   1680
            TabIndex        =   29
            ToolTipText     =   "Temperatura Máxima de Trabalho"
            Top             =   240
            Width           =   690
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "TNO (N.m):"
            Height          =   195
            Index           =   15
            Left            =   3240
            TabIndex        =   28
            ToolTipText     =   "Torque Normal de Operação"
            Top             =   240
            Width           =   810
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "-"
            Height          =   195
            Index           =   14
            Left            =   4800
            TabIndex        =   27
            Top             =   240
            Width           =   45
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Voltas do Volante:"
            Height          =   195
            Index           =   13
            Left            =   6360
            TabIndex        =   26
            Top             =   240
            Width           =   1290
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Dados gerais sobre a válvula:"
         Height          =   1575
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   7935
         Begin VB.TextBox Text9 
            Height          =   285
            Left            =   4800
            TabIndex        =   19
            Top             =   1080
            Width           =   3015
         End
         Begin VB.TextBox Text8 
            Height          =   285
            Left            =   3240
            TabIndex        =   18
            Top             =   1080
            Width           =   1455
         End
         Begin VB.TextBox Text7 
            Height          =   285
            Left            =   1680
            TabIndex        =   17
            Top             =   1080
            Width           =   1455
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   120
            TabIndex        =   16
            Top             =   1080
            Width           =   1455
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   6360
            TabIndex        =   11
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   4800
            TabIndex        =   8
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   3240
            TabIndex        =   7
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   1680
            TabIndex        =   5
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   120
            TabIndex        =   3
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Dados adicionais da válvula:"
            Height          =   195
            Index           =   8
            Left            =   4800
            TabIndex        =   15
            Top             =   840
            Width           =   2040
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Número da Peça:"
            Height          =   195
            Index           =   7
            Left            =   3240
            TabIndex        =   14
            Top             =   840
            Width           =   1245
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Ordem de Montagem:"
            Height          =   195
            Index           =   6
            Left            =   1680
            TabIndex        =   13
            Top             =   840
            Width           =   1530
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Extremidade:"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   12
            Top             =   840
            Width           =   915
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Bitola:"
            Height          =   195
            Index           =   4
            Left            =   6360
            TabIndex        =   10
            Top             =   240
            Width           =   435
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Material da Válvula:"
            Height          =   195
            Index           =   3
            Left            =   4800
            TabIndex        =   9
            Top             =   240
            Width           =   1395
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Material dos Internos:"
            Height          =   195
            Index           =   2
            Left            =   3240
            TabIndex        =   6
            Top             =   240
            Width           =   1515
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Classe de Pressão:"
            Height          =   195
            Index           =   1
            Left            =   1680
            TabIndex        =   4
            Top             =   240
            Width           =   1350
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Válvula:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   570
         End
      End
      Begin MSForms.CommandButton CommandButton1 
         Height          =   615
         Left            =   5880
         TabIndex        =   34
         Top             =   5640
         Width           =   2175
         Size            =   "3836;1085"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
End
Attribute VB_Name = "Tela_Valvula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

 Dim xl As New Excel.Application
    Dim xlw As Excel.Workbook
    Dim xlp As Excel.Worksheets
    
    Set xlw = xl.Workbooks.Add
    xlw.Sheets(3).Delete
    xlw.Sheets(2).Delete
    xlw.Worksheets(1).Name = "xuxu"
    
    xlw.Worksheets(1).Select
    xlw.Application.Cells(1, 1) = "TORQUE"
    xlw.Application.Cells(1, 2) = "PRESSAO"
    xlw.Application.Cells(1, 3) = "TEMPERATURA"
    xlw.Application.Cells(1, 4) = "POSICAO"
    
    xlw.Application.Cells(2, 1) = "1"
    xlw.Application.Cells(3, 1) = "2"
    xlw.Application.Cells(4, 1) = "3"
    
    xlw.Application.Range("A1:A4").Select
    xlw.Application.Charts.Add
    xlw.Application.ActiveChart.ApplyCustomType xlLine
    xlw.Application.Charts(1).Name = "É Assim..."
    
    xlw.SaveAs App.Path & "\teste.xls"
    xlw.Close
    
    xl.Quit
End Sub

Private Sub Form_Load()
    FR.Top = 100
    FR.Left = (Screen.Width / 2) - (FR.Width / 2)
    VAR_Tela = "Tela_Valvula"
End Sub


'*************************
' FUNCOES DESTA TELA
'*************************
Public Sub Fechar()
    Unload Tela_Valvula
End Sub
Public Sub Salvar()

End Sub

