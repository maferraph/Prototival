VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm Tela_Principal 
   BackColor       =   &H8000000C&
   Caption         =   "Prototival"
   ClientHeight    =   5250
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9105
   Icon            =   "Tela_Principal.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList IL 
      Left            =   0
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Principal.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Principal.frx":0BE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Principal.frx":1836
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Principal.frx":1C88
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Principal.frx":1FA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Principal.frx":22BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Principal.frx":25D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Principal.frx":28F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Principal.frx":2C0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Principal.frx":305C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Principal.frx":3376
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Principal.frx":37C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Principal.frx":3C1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Principal.frx":406C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Principal.frx":44BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Principal.frx":47D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Principal.frx":4C2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Principal.frx":5504
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Principal.frx":581E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TB 
      Align           =   1  'Align Top
      Height          =   1800
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9105
      _ExtentX        =   16060
      _ExtentY        =   3175
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "IL"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Novo"
            Description     =   "Nova ficha de válvula para validação"
            Object.ToolTipText     =   "Nova ficha de válvula para validação"
            Object.Tag             =   "Nova ficha de válvula para validação"
            ImageIndex      =   7
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Abrir"
            Description     =   "Abrir uma ficha de válvula para validação"
            Object.ToolTipText     =   "Abrir uma ficha de válvula para validação"
            Object.Tag             =   "Abrir uma ficha de válvula para validação"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Salvar"
            Description     =   "Salvar uma ficha de válvula"
            Object.ToolTipText     =   "Salvar uma ficha de válvula"
            Object.Tag             =   "Salvar uma ficha de válvula"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SalvarComo"
            Description     =   "Salvar como..."
            Object.ToolTipText     =   "Salvar como..."
            Object.Tag             =   "Salvar como..."
            ImageIndex      =   11
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Imprimir"
            Description     =   "Imprimir"
            Object.ToolTipText     =   "Imprimir"
            Object.Tag             =   "Imprimir"
            ImageIndex      =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PDF"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Sair"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Ciclo"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Vazamento"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Sensor"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Validacao"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Resultados"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SalvarDados"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SairTela"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Demais"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar ST 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   0
      Top             =   5010
      Width           =   9105
      _ExtentX        =   16060
      _ExtentY        =   423
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Enabled         =   0   'False
            Object.Width           =   15531
         EndProperty
      EndProperty
   End
   Begin VB.Menu Menu_Arquivo 
      Caption         =   "&Arquivo"
      Begin VB.Menu Menu_Arquivo_Novo 
         Caption         =   "&Novo"
         Shortcut        =   ^N
      End
      Begin VB.Menu Menu_Arquivo_Abrir 
         Caption         =   "&Abrir"
         Shortcut        =   ^A
      End
      Begin VB.Menu B2 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_Arquivo_Salvar 
         Caption         =   "Salvar"
         Shortcut        =   ^S
      End
      Begin VB.Menu Menu_Arquivo_SalvarComo 
         Caption         =   "Salvar como..."
      End
      Begin VB.Menu B1 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_Arquivo_Sair 
         Caption         =   "Sair"
      End
   End
   Begin VB.Menu Menu_Configuracoes 
      Caption         =   "&Configurações"
      Begin VB.Menu Menu_Configuracoes_Ciclos 
         Caption         =   "Ciclos"
      End
      Begin VB.Menu Menu_Configuracoes_Vazamanto 
         Caption         =   "Vazamentos"
      End
      Begin VB.Menu Menu_Configuracoes_Sensores 
         Caption         =   "Sensores"
      End
   End
   Begin VB.Menu Menu_Validacao 
      Caption         =   "&Validação"
      Begin VB.Menu Menu_Validacao_Iniciar 
         Caption         =   "Iniciar"
      End
   End
   Begin VB.Menu Menu_Ajuda 
      Caption         =   "&Ajuda"
   End
End
Attribute VB_Name = "Tela_Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VALOR As Integer
Private Sub BT_Sair_Click()
    Menu_Arquivo_Sair_Click
End Sub
Private Sub BT_Simulacao_Click()
    Menu_Validacao_Iniciar_Click
End Sub
Private Sub MDIForm_Load()
    BotoesMenus "FecharSalvarDados"
VALOR = 0

    Me.Menu_Arquivo_Salvar.Enabled = True
End Sub

Private Sub Menu_Arquivo_Abrir_Click()
    Tela_Simulacao.Show
End Sub

Private Sub Menu_Arquivo_Novo_Click()
    Tela_Valvula.Show
'Tela_TesteS.Show
End Sub
Private Sub Menu_Arquivo_Sair_Click()
    End
End Sub

Private Sub Menu_Arquivo_Salvar_Click()
    Tela_EncoderAngular.Show
End Sub

Private Sub Menu_Configuracoes_Ciclos_Click()
    Tela_Ciclos.Show
End Sub
Private Sub Menu_Configuracoes_Sensores_Click()
    Tela_Sensores.Show
End Sub
Private Sub Menu_Configuracoes_Vazamanto_Click()
    Tela_Vazamentos.Show
End Sub
Private Sub Menu_Validacao_Iniciar_Click()
    Tela_Simulacao.Show
End Sub
Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then 'Novo
        BotoesMenus "AbrirSalvarDados"
        Menu_Arquivo_Novo_Click
    ElseIf Button.Index = 2 Then 'Abrir
        BotoesMenus "AbrirSalvarDados"
    ElseIf Button.Index = 3 Then 'Salvar
    ElseIf Button.Index = 4 Then 'Salvar como
    ElseIf Button.Index = 5 Then 'Imprimir
    ElseIf Button.Index = 6 Then 'PDF
    ElseIf Button.Index = 7 Then 'Sair
        Menu_Arquivo_Sair_Click
    ElseIf Button.Index = 9 Then 'Ciclo
        BotoesMenus "AbrirSalvarDados"
        Menu_Configuracoes_Ciclos_Click
    ElseIf Button.Index = 10 Then 'Vazamento
        BotoesMenus "AbrirSalvarDados"
        Menu_Configuracoes_Vazamanto_Click
    ElseIf Button.Index = 11 Then 'Sensor
        BotoesMenus "AbrirSalvarDados"
        Menu_Configuracoes_Sensores_Click
    ElseIf Button.Index = 13 Then 'Validacao
        Menu_Validacao_Iniciar_Click
    ElseIf Button.Index = 14 Then 'Resultado
    ElseIf Button.Index = 16 Then 'SalvarDados
        If VAR_Tela = "Tela_Ciclos" Then
            Tela_Ciclos.Salvar
        ElseIf VAR_Tela = "Tela_Sensores" Then
            Tela_Sensores.Salvar
        ElseIf VAR_Tela = "Tela_Simulacao" Then
            Tela_Simulacao.Salvar
        ElseIf VAR_Tela = "Tela_Simulacao" Then
            Tela_Simulacao.Salvar
        ElseIf VAR_Tela = "Tela_Valvula" Then
            Tela_Valvula.Salvar
        ElseIf VAR_Tela = "Tela_Vazamentos" Then
            Tela_Vazamentos.Salvar
        End If
    ElseIf Button.Index = 17 Then 'SairTela
        FechaTela
        BotoesMenus "FecharSalvarDados"
    ElseIf Button.Index = 18 Then 'DemaisConfig
    
    End If
End Sub

'*************************
' FUNCOES DESTA TELA
'*************************

Private Sub FechaTela()
    If VAR_Tela = "Tela_Ciclos" Then
        Tela_Ciclos.Fechar
    ElseIf VAR_Tela = "Tela_Sensores" Then
        Tela_Sensores.Fechar
    ElseIf VAR_Tela = "Tela_Simulacao" Then
        Tela_Simulacao.Fechar
    ElseIf VAR_Tela = "Tela_Simulacao" Then
        Tela_Simulacao.Fechar
    ElseIf VAR_Tela = "Tela_Valvula" Then
        Tela_Valvula.Fechar
    ElseIf VAR_Tela = "Tela_Vazamentos" Then
        Tela_Vazamentos.Fechar
    End If
    If VAR_Salvar = False Then
        VAR_Tela = ""
    End If
End Sub
Private Sub BotoesMenus(VALOR As String)
    If VALOR = "FecharSalvarDados" Then
        TB.Buttons(1).Enabled = True
        TB.Buttons(2).Enabled = True
        TB.Buttons(3).Enabled = False
        TB.Buttons(4).Enabled = False
        TB.Buttons(5).Enabled = False
        TB.Buttons(6).Enabled = False
        TB.Buttons(7).Enabled = True
        TB.Buttons(9).Enabled = True
        TB.Buttons(10).Enabled = True
        TB.Buttons(11).Enabled = True
        TB.Buttons(13).Enabled = False
        TB.Buttons(14).Enabled = False
        TB.Buttons(16).Enabled = False
        TB.Buttons(17).Enabled = False
        TB.Buttons(18).Enabled = False
        With Me
            .Menu_Arquivo_Novo.Enabled = True
            .Menu_Arquivo_Abrir.Enabled = True
            .Menu_Arquivo_Salvar.Enabled = False
            .Menu_Arquivo_SalvarComo.Enabled = False
            .Menu_Arquivo_Sair.Enabled = True
            .Menu_Configuracoes_Ciclos.Enabled = True
            .Menu_Configuracoes_Vazamanto.Enabled = True
            .Menu_Configuracoes_Sensores.Enabled = True
            .Menu_Validacao_Iniciar.Enabled = False
        End With
    ElseIf VALOR = "AbrirSalvarDados" Then
        TB.Buttons(1).Enabled = False
        TB.Buttons(2).Enabled = False
        TB.Buttons(3).Enabled = False
        TB.Buttons(4).Enabled = False
        TB.Buttons(5).Enabled = False
        TB.Buttons(6).Enabled = False
        TB.Buttons(7).Enabled = False
        TB.Buttons(9).Enabled = False
        TB.Buttons(10).Enabled = False
        TB.Buttons(11).Enabled = False
        TB.Buttons(13).Enabled = False
        TB.Buttons(14).Enabled = False
        TB.Buttons(16).Enabled = False
        TB.Buttons(17).Enabled = True
        TB.Buttons(18).Enabled = False
        With Me
            .Menu_Arquivo_Novo.Enabled = False
            .Menu_Arquivo_Abrir.Enabled = False
            .Menu_Arquivo_Salvar.Enabled = True
            .Menu_Arquivo_SalvarComo.Enabled = True
            .Menu_Arquivo_Sair.Enabled = False
            .Menu_Configuracoes_Ciclos.Enabled = False
            .Menu_Configuracoes_Vazamanto.Enabled = False
            .Menu_Configuracoes_Sensores.Enabled = False
            .Menu_Validacao_Iniciar.Enabled = False
        End With
    ElseIf VALOR = "SalvarArquivo" Then
        TB.Buttons(1).Enabled = False
        TB.Buttons(2).Enabled = False
        TB.Buttons(3).Enabled = True
        TB.Buttons(4).Enabled = True
        TB.Buttons(5).Enabled = True
        TB.Buttons(6).Enabled = True
        TB.Buttons(7).Enabled = False
        TB.Buttons(9).Enabled = False
        TB.Buttons(10).Enabled = False
        TB.Buttons(11).Enabled = False
        TB.Buttons(13).Enabled = False
        TB.Buttons(14).Enabled = False
        TB.Buttons(16).Enabled = True
        TB.Buttons(17).Enabled = True
        TB.Buttons(18).Enabled = False
        With Me
            .Menu_Arquivo_Novo.Enabled = True
            .Menu_Arquivo_Abrir.Enabled = True
            .Menu_Arquivo_Salvar.Enabled = False
            .Menu_Arquivo_SalvarComo.Enabled = False
            .Menu_Arquivo_Sair.Enabled = True
            .Menu_Configuracoes_Ciclos.Enabled = True
            .Menu_Configuracoes_Vazamanto.Enabled = True
            .Menu_Configuracoes_Sensores.Enabled = True
            .Menu_Validacao_Iniciar.Enabled = False
        End With
    End If
End Sub
