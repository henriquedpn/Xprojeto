Attribute VB_Name = "modLock"
'==============================================================================
' Módulo: modLock - bloqueio exclusivo para gravações concorrentes
'
' Como funciona:
'   Antes de gravar no banco, o usuário abre o arquivo "<banco>.lock" com
'   acesso EXCLUSIVO (Lock Read Write). Enquanto o arquivo estiver aberto,
'   nenhum outro usuário consegue abri-lo, então as gravações ficam em fila.
'   Se o Excel de quem segura o bloqueio travar ou a rede cair, o Windows
'   fecha o arquivo automaticamente: não existe "bloqueio órfão".
'==============================================================================
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private mHandle As Integer
Private mArquivoLock As String

Public Function AdquirirLock(ByVal caminhoBanco As String, ByVal timeoutSeg As Long) As Boolean
    Dim f As Integer, t0 As Double
    If mHandle <> 0 Then AdquirirLock = True: Exit Function
    mArquivoLock = caminhoBanco & ".lock"
    t0 = Timer
    Do
        f = FreeFile
        If TentarAbrir(f) Then
            mHandle = f
            AdquirirLock = True
            Application.StatusBar = False
            Exit Function
        End If
        Application.StatusBar = APP_NOME & ": aguardando outro usuário concluir a gravação..."
        Esperar 400
    Loop While Decorrido(t0) < timeoutSeg
    Application.StatusBar = False
End Function

Public Sub LiberarLock()
    On Error Resume Next
    If mHandle <> 0 Then
        Close #mHandle
        mHandle = 0
        Kill mArquivoLock   ' falha silenciosamente se outro usuário já o abriu
    End If
End Sub

Private Function TentarAbrir(ByVal f As Integer) As Boolean
    On Error GoTo Falha
    Open mArquivoLock For Output Lock Read Write As #f
    Print #f, UsuarioAtual() & " | " & Environ$("COMPUTERNAME") & " | " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    TentarAbrir = True
    Exit Function
Falha:
    TentarAbrir = False
End Function

Private Function Decorrido(ByVal t0 As Double) As Double
    Decorrido = Timer - t0
    If Decorrido < 0 Then Decorrido = Decorrido + 86400#   ' virada da meia-noite
End Function

Public Sub Esperar(ByVal ms As Long)
    Dim i As Long
    For i = 1 To ms \ 50
        Sleep 50
        DoEvents
    Next i
End Sub
