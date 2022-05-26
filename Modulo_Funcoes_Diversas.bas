Attribute VB_Name = "Modulo_Funcoes_Diversas"
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Ret As String
Public VAR_Salvar As Boolean
Public VAR_RespMsg As String
Public VAR_ArquivoINI As String
Public VAR_Tela As String
Public VAR_PodeSalvar As Boolean

Public Sub EscreveINI(Arquivo As String, Secao As String, Chave As String, VALOR As String)
    WritePrivateProfileString Secao, Chave, VALOR, Arquivo
End Sub

Public Function LeINI(Arquivo As String, Secao As String, Chave As String)
    Ret = Space$(255)
    RetLen = GetPrivateProfileString(Secao, Chave, "", Ret, Len(Ret), Arquivo)
    Ret = Left$(Ret, RetLen)
    LeINI = Ret
End Function

Public Sub HabilitaSalvarDados(Estado As Boolean)
    VAR_Salvar = Estado
    Tela_Principal.TB.Buttons(16).Enabled = Estado
End Sub
Public Function ValidaTexto(KeyAscii As Integer) As Integer
    If Chr(KeyAscii) <> "0" And _
       Chr(KeyAscii) <> "1" And _
       Chr(KeyAscii) <> "2" And _
       Chr(KeyAscii) <> "3" And _
       Chr(KeyAscii) <> "4" And _
       Chr(KeyAscii) <> "5" And _
       Chr(KeyAscii) <> "6" And _
       Chr(KeyAscii) <> "7" And _
       Chr(KeyAscii) <> "8" And _
       Chr(KeyAscii) <> "9" And _
       Chr(KeyAscii) <> "," And _
       Chr(KeyAscii) <> "/" And _
       KeyAscii > 30 Then
        KeyAscii = 27
    End If
    ValidaTexto = KeyAscii
End Function
Public Function TempoString(ByVal TEMPO As Long) As String
    Dim VAR_H, VAR_M, VAR_S, VAR_T As Long, VAR_TEMP As String
    If TEMPO = 0 Then
        TempoString = "-"
    ElseIf TEMPO < 60 Then
        If Len(Trim(Str(TEMPO))) = 1 Then
            TempoString = "0h:0m:0" & Trim(Str(TEMPO)) & "s"
        Else
            TempoString = "0h:0m:" & Trim(Str(TEMPO)) & "s"
        End If
    ElseIf TEMPO = 60 Then
        TempoString = "0h:01m:00s"
    ElseIf TEMPO > 60 And TEMPO < 3600 Then
        VAR_M = Int(TEMPO / 60)
        VAR_S = TEMPO - (VAR_M * 60)
        If Len(Trim(Str(VAR_M))) = 1 Then
            VAR_TEMP = "0h:0" & Trim(Str(VAR_M))
        Else
            VAR_TEMP = "0h:" & Trim(Str(VAR_M))
        End If
        If Len(Trim(Str(VAR_S))) = 1 Then
            TempoString = VAR_TEMP & "m:0" & Trim(Str(VAR_S)) & "s"
        Else
            TempoString = VAR_TEMP & "m:" & Trim(Str(VAR_S)) & "s"
        End If
    ElseIf TEMPO = 3600 Then
        TempoString = "1h:00m:00s"
    ElseIf TEMPO > 3600 Then
        VAR_H = Int(TEMPO / 3600)
        VAR_M = TEMPO - (VAR_H * 3600)
        If VAR_M < 60 Then
            If Len(Trim(Str(VAR_M))) = 1 Then
                TempoString = Trim(Str(VAR_H)) & "h:00m" & ":0" & Trim(Str(VAR_M)) & "s"
            Else
                TempoString = Trim(Str(VAR_H)) & "h:00m" & ":" & Trim(Str(VAR_M)) & "s"
            End If
        ElseIf VAR_M = 60 Then
            TempoString = Trim(Str(VAR_H)) & "h:01m:00s"
        ElseIf VAR_M > 60 Then
            VAR_T = Int(VAR_M / 60)
            VAR_S = VAR_M - (VAR_T * 60)
            VAR_M = VAR_T
            If Len(Trim(Str(VAR_M))) = 1 Then
                VAR_TEMP = Trim(Str(VAR_H)) & "h:0" & Trim(Str(VAR_M))
            Else
                VAR_TEMP = Trim(Str(VAR_H)) & "h:" & Trim(Str(VAR_M))
            End If
            If Len(Trim(Str(VAR_S))) = 1 Then
                TempoString = VAR_TEMP & "m:0" & Trim(Str(VAR_S)) & "s"
            Else
                TempoString = VAR_TEMP & "m:" & Trim(Str(VAR_S)) & "s"
            End If
        End If
    End If
End Function
