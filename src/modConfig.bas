Attribute VB_Name = "modConfig"
'==============================================================================
' GERENCIADOR DE PROJETOS MULTIDISCIPLINAR
' Módulo: modConfig - constantes, parâmetros e utilitários gerais
'==============================================================================
Option Explicit

Public Const APP_NOME As String = "Gerenciador de Projetos"
Public Const SENHA_PROTECAO As String = "gp2026"

' ---- Abas da pasta LOCAL (front-end, uma cópia por usuário) ----
Public Const SH_PROJ As String = "Projetos"
Public Const SH_DASH As String = "Dashboard"
Public Const SH_USR As String = "Usuarios"
Public Const SH_CFG As String = "Config"
Public Const TBL_PROJ As String = "tblProjetos"
Public Const TBL_USR As String = "tblUsuarios"

' ---- Abas do BANCO compartilhado na rede (back-end .xlsx) ----
Public Const BD_PROJ As String = "Projetos"
Public Const BD_USR As String = "Usuarios"
Public Const BD_CTRL As String = "Controle"
Public Const BD_HIST As String = "Historico"

' ---- Colunas do banco: aba Projetos ----
Public Const B_ID As Long = 1
Public Const B_COD As Long = 2
Public Const B_DESC As Long = 3
Public Const B_COLAB As Long = 4
Public Const B_DISC As Long = 5
Public Const B_INI As Long = 6
Public Const B_DUR As Long = 7
Public Const B_FIM As Long = 8
Public Const B_STATUS As Long = 9
Public Const B_PERC As Long = 10
Public Const B_CRIADO_POR As Long = 11
Public Const B_CRIADO_EM As Long = 12
Public Const B_ATU_POR As Long = 13
Public Const B_ATU_EM As Long = 14
Public Const B_NCOLS As Long = 14

' ---- Colunas de usuários (banco e tabela local) ----
Public Const U_LOGIN As Long = 1
Public Const U_NOME As Long = 2
Public Const U_DISC As Long = 3
Public Const U_GESTOR As Long = 4
Public Const U_PERFIL As Long = 5
Public Const U_NCOLS As Long = 5

' ---- Colunas da tabela local tblProjetos (aba Projetos) ----
Public Const L_COD As Long = 1
Public Const L_DESC As Long = 2
Public Const L_COLAB As Long = 3
Public Const L_NOME As Long = 4
Public Const L_DISC As Long = 5
Public Const L_INI As Long = 6
Public Const L_DUR As Long = 7
Public Const L_FIM As Long = 8
Public Const L_STATUS As Long = 9
Public Const L_PERC As Long = 10
Public Const L_SIT As Long = 11
Public Const L_ATU_POR As Long = 12
Public Const L_ATU_EM As Long = 13
Public Const L_ID As Long = 14
Public Const L_NCOLS As Long = 14

' ---- Status do projeto ----
Public Const ST_NAO_INICIADO As String = "Não iniciado"
Public Const ST_ANDAMENTO As String = "Em andamento"
Public Const ST_PAUSADO As String = "Pausado"
Public Const ST_CONCLUIDO As String = "Concluído"
Public Const ST_CANCELADO As String = "Cancelado"

' ---- Situação de prazo (calculada) ----
Public Const SIT_ATRASADO As String = "Atrasado"
Public Const SIT_VENCE As String = "Vence em até 7 dias"
Public Const SIT_NO_PRAZO As String = "No prazo"
Public Const SIT_CONCLUIDO As String = "Concluído"
Public Const SIT_CANCELADO As String = "Cancelado"

' ---- Perfis ----
Public Const PERFIL_COLAB As String = "Colaborador"
Public Const PERFIL_GESTOR As String = "Gestor"
Public Const PERFIL_ADMIN As String = "Admin"

' Indica que uma gravação está em andamento (evita a auto-atualização concorrente)
Public gOcupado As Boolean

Private mProfundidade As Long

'------------------------------------------------------------------------------
' Acesso aos intervalos nomeados
'------------------------------------------------------------------------------
Public Function NR(ByVal nome As String) As Range
    Set NR = ThisWorkbook.Names(nome).RefersToRange
End Function

Public Function CaminhoBD() As String
    CaminhoBD = Texto(NR("cfgCaminho").Value)
End Function

Public Function IntervaloSeg() As Long
    IntervaloSeg = NumLng(NR("cfgIntervalo").Value, 30)
    If IntervaloSeg < 10 Then IntervaloSeg = 10
End Function

Public Function TimeoutLockSeg() As Long
    TimeoutLockSeg = NumLng(NR("cfgTimeout").Value, 20)
    If TimeoutLockSeg < 3 Then TimeoutLockSeg = 3
End Function

Public Function UsaDiasUteis() As Boolean
    UsaDiasUteis = (UCase$(Texto(NR("cfgDiasUteis").Value)) <> "NAO")
End Function

Public Function VerTodos() As Boolean
    VerTodos = (UCase$(Texto(NR("cfgVerTodos").Value)) = "SIM")
End Function

Public Function EditandoUsuarios() As Boolean
    EditandoUsuarios = (UCase$(Texto(NR("cfgEditandoUsuarios").Value)) = "SIM")
End Function

'------------------------------------------------------------------------------
' Usuário atual = login do Windows (em minúsculas)
'------------------------------------------------------------------------------
Public Function UsuarioAtual() As String
    UsuarioAtual = LCase$(Trim$(Environ$("USERNAME")))
    If Len(UsuarioAtual) = 0 Then UsuarioAtual = LCase$(Trim$(Application.UserName))
End Function

'------------------------------------------------------------------------------
' Regras de negócio
'------------------------------------------------------------------------------
' Data fim a partir da data de início e da duração estimada.
' Em dias úteis: considera fins de semana e a lista de Feriados (aba Config).
' Duração 1 = termina no próprio dia de início (ou no próximo dia útil).
Public Function CalcularDataFim(ByVal inicio As Date, ByVal duracao As Long) As Date
    If duracao < 1 Then duracao = 1
    If UsaDiasUteis() Then
        CalcularDataFim = Application.WorksheetFunction.WorkDay(inicio - 1, duracao, NR("Feriados"))
    Else
        CalcularDataFim = inicio + duracao - 1
    End If
End Function

Public Function SituacaoPrazo(ByVal status As Variant, ByVal dataFim As Variant) As String
    Dim st As String
    st = Texto(status)
    If st = ST_CONCLUIDO Then SituacaoPrazo = SIT_CONCLUIDO: Exit Function
    If st = ST_CANCELADO Then SituacaoPrazo = SIT_CANCELADO: Exit Function
    If Not IsDate(dataFim) Then Exit Function
    If CDate(dataFim) < Date Then
        SituacaoPrazo = SIT_ATRASADO
    ElseIf CDate(dataFim) - Date <= 7 Then
        SituacaoPrazo = SIT_VENCE
    Else
        SituacaoPrazo = SIT_NO_PRAZO
    End If
End Function

'------------------------------------------------------------------------------
' Conversões seguras
'------------------------------------------------------------------------------
Public Function Texto(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function
    Texto = Trim$(CStr(v))
End Function

Public Function NumLng(ByVal v As Variant, Optional ByVal padrao As Long = 0) As Long
    NumLng = padrao
    If Len(Texto(v)) = 0 Then Exit Function
    If IsNumeric(v) Then NumLng = CLng(v)
End Function

Public Function NumDbl(ByVal v As Variant, Optional ByVal padrao As Double = 0) As Double
    NumDbl = padrao
    If Len(Texto(v)) = 0 Then Exit Function
    If IsNumeric(v) Then NumDbl = CDbl(v)
End Function

' Carimbo de versão (usado no controle de concorrência otimista)
Public Function DataTexto(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    If IsDate(v) Then DataTexto = Format$(CDate(v), "yyyy-mm-dd hh:nn:ss")
End Function

'------------------------------------------------------------------------------
' Proteção das abas
'------------------------------------------------------------------------------
Public Sub Desproteger(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Unprotect Password:=SENHA_PROTECAO
End Sub

Public Sub Proteger(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Protect Password:=SENHA_PROTECAO, DrawingObjects:=True, Contents:=True, _
        Scenarios:=True, UserInterfaceOnly:=True, AllowFiltering:=True, _
        AllowFormattingColumns:=True
End Sub

Public Sub ProtegerUsuarios()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_USR)
    If EditandoUsuarios() Then Desproteger ws Else Proteger ws
End Sub

Public Sub ProtegerTudo()
    Proteger ThisWorkbook.Worksheets(SH_PROJ)
    Proteger ThisWorkbook.Worksheets(SH_DASH)
    ProtegerUsuarios
End Sub

'------------------------------------------------------------------------------
' Desempenho (aninhável)
'------------------------------------------------------------------------------
Public Sub ModoRapido(ByVal ligar As Boolean)
    If ligar Then
        If mProfundidade = 0 Then
            Application.ScreenUpdating = False
            Application.EnableEvents = False
            Application.DisplayAlerts = False
        End If
        mProfundidade = mProfundidade + 1
    Else
        If mProfundidade > 0 Then mProfundidade = mProfundidade - 1
        If mProfundidade = 0 Then
            Application.ScreenUpdating = True
            Application.EnableEvents = True
            Application.DisplayAlerts = True
        End If
    End If
End Sub

Public Sub ModoRapidoReset()
    mProfundidade = 1
    ModoRapido False
End Sub

'------------------------------------------------------------------------------
' Formatação condicional da coluna "Situação"
'------------------------------------------------------------------------------
Public Sub AplicarCoresSituacao(ByVal rg As Range)
    If rg Is Nothing Then Exit Sub
    rg.FormatConditions.Delete
    AddCor rg, SIT_ATRASADO, RGB(255, 199, 206), RGB(156, 0, 6)
    AddCor rg, SIT_VENCE, RGB(255, 235, 156), RGB(156, 87, 0)
    AddCor rg, SIT_NO_PRAZO, RGB(221, 235, 247), RGB(31, 78, 121)
    AddCor rg, SIT_CONCLUIDO, RGB(198, 239, 206), RGB(0, 97, 0)
    AddCor rg, SIT_CANCELADO, RGB(217, 217, 217), RGB(89, 89, 89)
End Sub

Private Sub AddCor(ByVal rg As Range, ByVal valor As String, ByVal fundo As Long, ByVal fonte As Long)
    With rg.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""" & valor & """")
        .Interior.Color = fundo
        .Font.Color = fonte
        .Font.Bold = True
    End With
End Sub
