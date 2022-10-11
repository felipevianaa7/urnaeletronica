VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cxFim 
   Caption         =   "UserForm2"
   ClientHeight    =   7095
   ClientLeft      =   12135
   ClientTop       =   465
   ClientWidth     =   10440
   OleObjectBlob   =   "cxFim.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cxFim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btLiberar_Click()
    
    Dim iLiberar As String
    
    iLiberar = InputBox("Digite a senha para liberar.", "Tela de liberação")
    If iLiberar = "123" Then
        Unload Me
    Else
        MsgBox "Senha inválida. Verifique!", vbExclamation, "Informação"
    End If

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode <> 1 Then Cancel = 1
End Sub
