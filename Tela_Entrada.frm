VERSION 5.00
Begin VB.Form Tela_Entrada 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   4245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   0
      Top             =   0
   End
   Begin VB.Image Image1 
      Height          =   4215
      Left            =   0
      Picture         =   "Tela_Entrada.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7095
   End
End
Attribute VB_Name = "Tela_Entrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Texto As String

Private Sub Form_Load()
    ' Texto = LeINI(Text1.Text, Text2.Text, Text7.Text)
    VAR_ArquivoINI = App.Path & "\Prototival.con"
End Sub
Private Sub Timer1_Timer()
    Tela_Principal.Show
    Unload Tela_Entrada
End Sub
