VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   10590
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16965
   OleObjectBlob   =   "Arquivo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub bt0_Click()
    Dim iLinha As Integer
    Dim i As Integer
    
    cdNumero = cdNumero & 0
    
    iLinha = Planilha1.Cells(Planilha1.Cells.Rows.Count, "a").End(xlUp).Row
    For i = 2 To iLinha
        cdNome = ""
        Image1.Picture = LoadPicture("G:\Projeto Urna Eletônica\Imagens\oculto.bmp")
        If cdNumero = Planilha1.Cells(i, 1) Then
            cdNome = Planilha1.Cells(i, 2)
            Image1.Picture = LoadPicture("G:\Projeto Urna Eletônica\Imagens\" & cdNome & ".jpg")
            Exit Sub
        End If
    Next i
End Sub

Private Sub bt1_Click()
    Dim iLinha As Integer
    Dim i As Integer
    cdNumero = cdNumero & 1
    iLinha = Planilha1.Cells(Planilha1.Cells.Rows.Count, "a").End(xlUp).Row
    For i = 2 To iLinha
    cdNome = ""
    Image1.Picture = LoadPicture("G:\Projeto Urna Eletônica\Imagens\oculto.bmp")
        If cdNumero = Planilha1.Cells(i, 1) Then
            cdNome = Planilha1.Cells(i, 2)
            Image1.Picture = LoadPicture("G:\Projeto Urna Eletônica\Imagens\" & cdNome & ".jpg")
            Exit Sub
        End If
    Next i
End Sub

Private Sub bt2_Click()
    Dim iLinha As Integer
    Dim i As Integer
    cdNumero = cdNumero & 2
    iLinha = Planilha1.Cells(Planilha1.Cells.Rows.Count, "a").End(xlUp).Row
    For i = 2 To iLinha
    cdNome = ""
    Image1.Picture = LoadPicture("G:\Projeto Urna Eletônica\Imagens\oculto.bmp")
        If cdNumero = Planilha1.Cells(i, 1) Then
            cdNome = Planilha1.Cells(i, 2)
            Image1.Picture = LoadPicture("G:\Projeto Urna Eletônica\Imagens\" & cdNome & ".jpg")
            Exit Sub
        End If
    Next i
End Sub

Private Sub bt3_Click()
    Dim iLinha As Integer
    Dim i As Integer
    cdNumero = cdNumero & 3
    iLinha = Planilha1.Cells(Planilha1.Cells.Rows.Count, "a").End(xlUp).Row
    For i = 2 To iLinha
    cdNome = ""
    Image1.Picture = LoadPicture("G:\Projeto Urna Eletônica\Imagens\oculto.bmp")
        If cdNumero = Planilha1.Cells(i, 1) Then
            cdNome = Planilha1.Cells(i, 2)
            Image1.Picture = LoadPicture("G:\Projeto Urna Eletônica\Imagens\" & cdNome & ".jpg")
            Exit Sub
        End If
    Next i
End Sub

Private Sub bt4_Click()
    Dim iLinha As Integer
    Dim i As Integer
    cdNumero = cdNumero & 4
    iLinha = Planilha1.Cells(Planilha1.Cells.Rows.Count, "a").End(xlUp).Row
    For i = 2 To iLinha
    cdNome = ""
    Image1.Picture = LoadPicture("G:\Projeto Urna Eletônica\Imagens\oculto.bmp")
        If cdNumero = Planilha1.Cells(i, 1) Then
            cdNome = Planilha1.Cells(i, 2)
            Image1.Picture = LoadPicture("G:\Projeto Urna Eletônica\Imagens\" & cdNome & ".jpg")
            Exit Sub
        End If
    Next i
End Sub

Private Sub bt5_Click()
    Dim iLinha As Integer
    Dim i As Integer
    cdNumero = cdNumero & 5
    iLinha = Planilha1.Cells(Planilha1.Cells.Rows.Count, "a").End(xlUp).Row
    For i = 2 To iLinha
    cdNome = ""
    Image1.Picture = LoadPicture("G:\Projeto Urna Eletônica\Imagens\oculto.bmp")
        If cdNumero = Planilha1.Cells(i, 1) Then
            cdNome = Planilha1.Cells(i, 2)
            Image1.Picture = LoadPicture("G:\Projeto Urna Eletônica\Imagens\" & cdNome & ".jpg")
            Exit Sub
        End If
    Next i
End Sub

Private Sub bt6_Click()
    Dim iLinha As Integer
    Dim i As Integer
    cdNumero = cdNumero & 6
    iLinha = Planilha1.Cells(Planilha1.Cells.Rows.Count, "a").End(xlUp).Row
    For i = 2 To iLinha
    cdNome = ""
    Image1.Picture = LoadPicture("G:\Projeto Urna Eletônica\Imagens\oculto.bmp")
        If cdNumero = Planilha1.Cells(i, 1) Then
            cdNome = Planilha1.Cells(i, 2)
            Image1.Picture = LoadPicture("G:\Projeto Urna Eletônica\Imagens\" & cdNome & ".jpg")
            Exit Sub
        End If
    Next i
End Sub

Private Sub bt7_Click()
    Dim iLinha As Integer
    Dim i As Integer
    cdNumero = cdNumero & 7
    iLinha = Planilha1.Cells(Planilha1.Cells.Rows.Count, "a").End(xlUp).Row
    For i = 2 To iLinha
    cdNome = ""
    Image1.Picture = LoadPicture("G:\Projeto Urna Eletônica\Imagens\oculto.bmp")
        If cdNumero = Planilha1.Cells(i, 1) Then
            cdNome = Planilha1.Cells(i, 2)
            Image1.Picture = LoadPicture("G:\Projeto Urna Eletônica\Imagens\" & cdNome & ".jpg")
            Exit Sub
        End If
    Next i
End Sub

Private Sub bt8_Click()
    Dim iLinha As Integer
    Dim i As Integer
    cdNumero = cdNumero & 8
    iLinha = Planilha1.Cells(Planilha1.Cells.Rows.Count, "a").End(xlUp).Row
    For i = 2 To iLinha
    cdNome = ""
    Image1.Picture = LoadPicture("G:\Projeto Urna Eletônica\Imagens\oculto.bmp")
        If cdNumero = Planilha1.Cells(i, 1) Then
            cdNome = Planilha1.Cells(i, 2)
            Image1.Picture = LoadPicture("G:\Projeto Urna Eletônica\Imagens\" & cdNome & ".jpg")
            Exit Sub
        End If
    Next i
End Sub

Private Sub bt9_Click()
    Dim iLinha As Integer
    Dim i As Integer
    cdNumero = cdNumero & 9
    iLinha = Planilha1.Cells(Planilha1.Cells.Rows.Count, "a").End(xlUp).Row
    For i = 2 To iLinha
    cdNome = ""
    Image1.Picture = LoadPicture("G:\Projeto Urna Eletônica\Imagens\oculto.bmp")
        If cdNumero = Planilha1.Cells(i, 1) Then
            cdNome = Planilha1.Cells(i, 2)
            Image1.Picture = LoadPicture("G:\Projeto Urna Eletônica\Imagens\" & cdNome & ".jpg")
            Exit Sub
        End If
    Next i
End Sub

Private Sub btApagar_Click()
    Dim iLinha As Integer
    Dim i As Integer
    
    cdNumero = Mid(cdNumero, 1, Len(cdNumero) - 1)
    cdNome = ""
    
    iLinha = Planilha1.Cells(Planilha1.Cells.Rows.Count, "a").End(xlUp).Row
    For i = 2 To iLinha
        cdNome = ""
        Image1.Picture = LoadPicture("G:\Projeto Urna Eletônica\Imagens\oculto.bmp")
        If cdNumero = Planilha1.Cells(i, 1) Then
            cdNome = Planilha1.Cells(i, 2)
            Image1.Picture = LoadPicture("G:\Projeto Urna Eletônica\Imagens\" & cdNome & ".jpg")
            Exit Sub
        End If
    Next i
End Sub

Private Sub btBranco_Click()
 Dim Pergunta As String
    Dim iLinha As Integer
    
    cdNumero = ""
    cdNome = ""
    Image1.Picture = LoadPicture("G:\Projeto Urna Eletônica\Imagens\oculto.bmp")
    
    Pergunta = MsgBox("Você confirma seu voto em BRANCO ?", vbYesNo + vbInformation, "Informação")
    
    If Pergunta = vbNo Then
        Exit Sub
    End If
    
    iLinha = Planilha2.Cells(Planilha1.Cells.Rows.Count, "a").End(xlUp).Row + 1
    Planilha2.Cells(iLinha, 1) = "Nulo"
    
    
    btLimpar_Click
    
    Call Som
    cxFim.Show
    
End Sub
 
 Private Sub btConfirma_Click()
    Dim Pergunta As String
    Dim iLinha As Integer
    
    
    Pergunta = MsgBox("Você confirma seu voto ?", vbYesNo + vbInformation, "Informação")
    
    If Pergunta = vbNo Then
        Exit Sub
    End If
    
    iLinha = Planilha2.Cells(Planilha1.Cells.Rows.Count, "a").End(xlUp).Row + 1
    If cdNome = "" Then
    Planilha2.Cells(iLinha, 1) = "Nulo"
    Else
    Planilha2.Cells(iLinha, 1) = cdNome
    End If
    
    btLimpar_Click
    
    Call Som
    cxFim.Show
    
End Sub
Private Sub btLimpar_Click()
    cdNumero = ""
    cdNome = ""
    Image1.Picture = LoadPicture("G:\Projeto Urna Eletônica\Imagens\oculto.bmp")
End Sub

Private Sub CommandButton5_Click()

End Sub

