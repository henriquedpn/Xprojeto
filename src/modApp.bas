Attribute VB_Name = "modApp"
'==============================================================================
' Módulo: modApp - inicialização e atualização automática
'==============================================================================
Option Explicit

Private mProxima As Date
Private mAtivo As Boolean

' Executado automaticamente ao abrir a pasta de trabalho
Public Sub Auto_Open()
    Iniciar
End Sub

' Executado automaticamente ao fechar a pasta de trabalho
Public Sub Auto_Close()
    PararAutoAtualizacao
    LiberarLock
End Sub

' Também é o botão "Conectar / Sincronizar" da aba Config
Public Sub Iniciar()
    On Error GoTo Falha
    If Not EstruturaOK() Then ConfigurarPastaDeTrabalho
    ProtegerTudo
    LimparFormulario

    If Len(CaminhoBD()) = 0 Then
        ThisWorkbook.Worksheets(SH_CFG).Activate
        NR("cfgCaminho").Select
        MsgBox "Bem-vindo ao " & APP_NOME & "!" & vbLf & vbLf & _
               "1) Informe na célula destacada o caminho do banco compartilhado na rede" & vbLf & _
               "    (ex.: \\servidor\projetos\BD_Projetos.xlsx);" & vbLf & _
               "2) Clique em 'Conectar'." & vbLf & vbLf & _
               "Primeira instalação? O administrador deve clicar em 'Criar banco na rede'.", _
               vbInformation, APP_NOME
        Exit Sub
    End If

    If Not ArquivoExiste(CaminhoBD()) Then
        ThisWorkbook.Worksheets(SH_CFG).Activate
        MsgBox "Não foi possível acessar o banco de projetos:" & vbLf & CaminhoBD() & vbLf & vbLf & _
               "Verifique a conexão com a rede e o caminho na aba Config.", vbExclamation, APP_NOME
        Exit Sub
    End If

    If CarregarDoServidor(False) Then
        VerificarCadastroUsuario
        IniciarAutoAtualizacao
        ThisWorkbook.Worksheets(SH_PROJ).Activate
    End If
    Exit Sub
Falha:
    ModoRapidoReset
    MsgBox "Erro ao iniciar: " & Err.Description, vbCritical, APP_NOME
End Sub

'------------------------------------------------------------------------------
' Atualização automática: a cada N segundos verifica se o arquivo do banco
' mudou (data de modificação). Se mudou, recarrega a lista e o Dashboard.
' A verificação é leve: só abre o banco quando há alteração.
'------------------------------------------------------------------------------
Public Sub IniciarAutoAtualizacao()
    PararAutoAtualizacao
    mAtivo = True
    AgendarProxima
End Sub

Public Sub PararAutoAtualizacao()
    On Error Resume Next
    If mProxima <> 0 Then
        Application.OnTime EarliestTime:=mProxima, Procedure:=NomeRotina(), Schedule:=False
    End If
    mProxima = 0
    mAtivo = False
End Sub

Private Sub AgendarProxima()
    mProxima = Now + TimeSerial(0, 0, IntervaloSeg())
    Application.OnTime EarliestTime:=mProxima, Procedure:=NomeRotina()
End Sub

Private Function NomeRotina() As String
    NomeRotina = "'" & ThisWorkbook.Name & "'!VerificarAtualizacoes"
End Function

Public Sub VerificarAtualizacoes()
    mProxima = 0
    On Error GoTo Fim
    If Not gOcupado Then
        If ServidorMudou() Then CarregarDoServidor True
    End If
Fim:
    On Error Resume Next
    If mAtivo Then AgendarProxima
End Sub
