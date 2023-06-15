Attribute VB_Name = "Módulo1"
Option Explicit

Private Declare PtrSafe Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
 
Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Const SND_FILENAME = &H20000

'Defina os nomes dos arquivos de som'
Const NoPadraoWav = "som.wav"

Dim SomWave
Sub Som()

    SomWave = "C:\Users\Usuário\Documents\Projeto Urna Eletônica\" & NoPadraoWav
    AtivarAlarme
    
End Sub
Sub AtivarAlarme()
    
    Call PlaySound(SomWave, 0&, SND_ASYNC Or SND_FILENAME)
    
End Sub
