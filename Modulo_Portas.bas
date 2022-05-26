Attribute VB_Name = "Modulo_Portas"
Public Declare Function LePorta Lib "inpout32.dll" _
Alias "Inp32" (ByVal EnderecoPortaH As Integer) As Integer

Public Declare Sub EscrevePorta Lib "inpout32.dll" _
Alias "Out32" (ByVal EnderecoPortaH As Integer, ByVal VALOR As Integer)

'Resolucao do conversor ADC
Public Const VAR_RESOLUCAO As Integer = 256

