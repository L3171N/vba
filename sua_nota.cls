VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Public Sub sua_nota()

    Dim N1 As Integer
    Dim N2 As Integer
    Dim N3 As Integer
    Dim resultado
    
    N1 = InputBox("Digite Sua nota do primeiro bimestre: ")
    N2 = InputBox("Digite Sau nota do segundo bimestre: ")
    N3 = InputBox("Digite Sua nota do terceiro bimestre: ")
     
     resultado = N1 * 0.4 + N2 * 0.3 + N3 * 0.3
     
     If resultado > 6 Then
      MsgBox ("Aprovado ")
     ElseIf resultado <= 6 Then
      MsgBox ("Reprovado")
    End If

End Sub

