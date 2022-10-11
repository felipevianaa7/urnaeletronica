VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   10590
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16200
   OleObjectBlob   =   "Form1.frx":0000
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
    cdNumero = Mid(cdNumero, 1, Len(cdNumero) - 1)
    cdNome = ""
    Image1.Picture = LoadPicture("G:\Projeto Urna Eletônica\Imagens\oculto.bmp")
End Sub

Private Sub btLimpar_Click()
    cdNumero = ""
    cdNome = ""
    Image1.Picture = LoadPicture("G:\Projeto Urna Eletônica\Imagens\oculto.bmp")
End Sub

Private Sub CommandButton2_Click()

End Sub
