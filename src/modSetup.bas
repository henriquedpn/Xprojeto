Attribute VB_Name = "modSetup"
'==============================================================================
' Módulo: modSetup - cria as abas, tabelas, nomes, validações e botões
' Executado automaticamente na primeira abertura (Iniciar).
'==============================================================================
Option Explicit

Private Const COR_BOTAO_AZUL As Long = 7949855     ' RGB(31, 78, 121)
Private Const COR_BOTAO_VERDE As Long = 3506772    ' RGB(84, 130, 53)
Private Const COR_BOTAO_VERMELHO As Long = 2763442 ' RGB(178, 42, 42)
Private Const COR_BOTAO_CINZA As Long = 8421504    ' RGB(128, 128, 128)
Private Const COR_ENTRADA As Long = 13434879       ' RGB(255, 255, 204)

Public Function EstruturaOK() As Boolean
    Dim n As Variant, ws As Worksheet, nm As Name, lo As ListObject
    On Error GoTo Nao
    For Each n In Array(SH_PROJ, SH_DASH, SH_USR, SH_CFG)
        Set ws = ThisWorkbook.Worksheets(CStr(n))
    Next n
    For Each n In Array("cfgCaminho", "cfgIntervalo", "cfgDiasUteis", "cfgTimeout", "cfgVerTodos", _
                        "cfgUltimoCarimbo", "cfgUltimaSync", "cfgEditandoUsuarios", "inCodigo", "inModo", _
                        "inID", "inVersao", "dashFiltro", "lstColaboradores", "lstStatus", "Feriados")
        Set nm = ThisWorkbook.Names(CStr(n))
    Next n
    Set lo = ThisWorkbook.Worksheets(SH_PROJ).ListObjects(TBL_PROJ)
    Set lo = ThisWorkbook.Worksheets(SH_USR).ListObjects(TBL_USR)
    EstruturaOK = True
    Exit Function
Nao:
    EstruturaOK = False
End Function

' Pode ser executada manualmente (Alt+F8) para recriar o layout.
' Os dados de projetos/usuários ficam no servidor e são recarregados.
Public Sub ConfigurarPastaDeTrabalho()
    Dim wsP As Worksheet, wsD As Worksheet, wsU As Worksheet, wsC As Worksheet
    On Error GoTo Falha
    ModoRapido True
    Set wsP = ObterAba(SH_PROJ, Nothing)
    Set wsD = ObterAba(SH_DASH, wsP)
    Set wsU = ObterAba(SH_USR, wsD)
    Set wsC = ObterAba(SH_CFG, wsU)
    MontarConfig wsC
    MontarUsuarios wsU
    MontarProjetos wsP
    MontarDashboard wsD
    RemoverAbasVazias
    ModoRapido False
    wsP.Activate
    Exit Sub
Falha:
    ModoRapidoReset
    MsgBox "Erro ao configurar a pasta de trabalho: " & Err.Description, vbCritical, APP_NOME
End Sub

'------------------------------------------------------------------------------
Private Function ObterAba(ByVal nome As String, ByVal depoisDe As Worksheet) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nome)
    On Error GoTo 0
    If ws Is Nothing Then
        If depoisDe Is Nothing Then
            Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Sheets(1))
        Else
            Set ws = ThisWorkbook.Worksheets.Add(After:=depoisDe)
        End If
        ws.Name = nome
    End If
    Set ObterAba = ws
End Function

Private Sub LimparAba(ByVal ws As Worksheet)
    Dim i As Long
    Desproteger ws
    For i = ws.ListObjects.Count To 1 Step -1
        ws.ListObjects(i).Delete
    Next i
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
    ws.Cells.Clear
    ws.Cells.Validation.Delete
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 10
End Sub

Private Sub RemoverAbasVazias()
    Dim ws As Worksheet, i As Long
    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        Set ws = ThisWorkbook.Worksheets(i)
        Select Case ws.Name
            Case SH_PROJ, SH_DASH, SH_USR, SH_CFG
            Case Else
                If Application.WorksheetFunction.CountA(ws.Cells) = 0 And ws.Shapes.Count = 0 Then ws.Delete
        End Select
    Next i
End Sub

Private Sub DefinirNome(ByVal nome As String, ByVal rg As Range)
    ThisWorkbook.Names.Add Name:=nome, RefersTo:="='" & rg.Parent.Name & "'!" & rg.Address
End Sub

Private Sub CriarBotao(ByVal ws As Worksheet, ByVal esquerda As Double, ByVal topo As Double, _
                       ByVal largura As Double, ByVal legenda As String, ByVal macro As String, ByVal cor As Long)
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, esquerda, topo, largura, 24)
    shp.Name = "btn" & macro
    shp.Fill.ForeColor.RGB = cor
    shp.Line.Visible = msoFalse
    shp.Placement = xlFreeFloating
    With shp.TextFrame2
        .VerticalAnchor = msoAnchorMiddle
        .MarginLeft = 2
        .MarginRight = 2
        With .TextRange
            .Text = legenda
            .Font.Size = 10
            .Font.Bold = msoTrue
            .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With
    shp.OnAction = macro
End Sub

Private Sub TituloAba(ByVal ws As Worksheet, ByVal legenda As String)
    With ws.Range("A1")
        .Value = legenda
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(31, 78, 121)
    End With
    ws.Rows(1).RowHeight = 26
End Sub

Private Sub SemGrade(ByVal ws As Worksheet, Optional ByVal linhaCongelar As Long = 0)
    ws.Activate
    With ActiveWindow
        .DisplayGridlines = False
        .FreezePanes = False
        If linhaCongelar > 0 Then
            .ScrollRow = 1
            .SplitColumn = 0
            .SplitRow = linhaCongelar
            .FreezePanes = True
        End If
    End With
End Sub

'------------------------------------------------------------------------------
' CONFIG
'------------------------------------------------------------------------------
Private Sub MontarConfig(ByVal ws As Worksheet)
    Dim caminhoAnt As String, rg As Range, x As Double

    On Error Resume Next
    caminhoAnt = CStr(ThisWorkbook.Names("cfgCaminho").RefersToRange.Value)
    On Error GoTo 0

    LimparAba ws
    TituloAba ws, "Configurações"
    ws.Columns("A").ColumnWidth = 52
    ws.Columns("B").ColumnWidth = 55
    ws.Columns("C").ColumnWidth = 3
    ws.Columns("D").ColumnWidth = 3
    ws.Columns("E").ColumnWidth = 14
    ws.Columns("F").ColumnWidth = 3
    ws.Columns("G").ColumnWidth = 16
    ws.Columns("H").ColumnWidth = 14
    ws.Columns("I").ColumnWidth = 8
    ws.Columns("J").ColumnWidth = 3
    ws.Columns("K").ColumnWidth = 20

    ws.Range("A3").Value = "Caminho do banco compartilhado na rede (.xlsx)"
    ws.Range("A4").Value = "Intervalo da atualização automática (segundos)"
    ws.Range("A5").Value = "Data fim em dias úteis? (SIM/NAO)"
    ws.Range("A6").Value = "Tempo máximo de espera pelo bloqueio (segundos)"
    ws.Range("A7").Value = "Colaboradores veem projetos de todos? (SIM/NAO)"
    ws.Range("B3").Value = caminhoAnt
    ws.Range("B4").Value = 30
    ws.Range("B5").Value = "SIM"
    ws.Range("B6").Value = 20
    ws.Range("B7").Value = "NAO"
    ws.Range("B3:B7").Interior.Color = COR_ENTRADA
    ws.Range("B3:B7").Borders.Color = RGB(191, 191, 191)
    ws.Range("A8").Value = "Use sempre o caminho UNC (\\servidor\pasta\arquivo.xlsx), igual em todas as máquinas."
    ws.Range("A8").Font.Italic = True
    ws.Range("A8").Font.Color = RGB(128, 128, 128)

    ws.Range("A13").Value = "Controle interno (não editar)"
    ws.Range("A13").Font.Bold = True
    ws.Range("A14").Value = "Versão do servidor sincronizada"
    ws.Range("A15").Value = "Última sincronização"
    ws.Range("A16").Value = "ID do projeto em edição"
    ws.Range("A17").Value = "Versão do projeto em edição"
    ws.Range("A18").Value = "Admin editando usuários (SIM/NAO)"
    ws.Range("B18").Value = "NAO"
    ws.Range("B15").NumberFormat = "dd/mm/yyyy hh:mm:ss"
    ws.Range("A14:B18").Font.Color = RGB(128, 128, 128)

    ws.Range("E2").Value = "Feriados"
    ws.Range("G2").Value = "Status"
    ws.Range("H2").Value = "Perfis"
    ws.Range("I2").Value = "Sim/Não"
    ws.Range("K2").Value = "Disciplinas"
    ws.Range("E2:K2").Font.Bold = True
    ws.Range("E3:E200").NumberFormat = "dd/mm/yyyy"
    ws.Range("G3:G7").Value = Application.WorksheetFunction.Transpose( _
        Array(ST_NAO_INICIADO, ST_ANDAMENTO, ST_PAUSADO, ST_CONCLUIDO, ST_CANCELADO))
    ws.Range("H3:H5").Value = Application.WorksheetFunction.Transpose(Array(PERFIL_COLAB, PERFIL_GESTOR, PERFIL_ADMIN))
    ws.Range("I3:I4").Value = Application.WorksheetFunction.Transpose(Array("SIM", "NAO"))
    ws.Range("K3:K11").Value = Application.WorksheetFunction.Transpose(Array("Arquitetura", "Civil", _
        "Estrutural", "Elétrica", "Hidráulica", "Mecânica", "Automação", "Planejamento", "Coordenação"))

    DefinirNome "cfgCaminho", ws.Range("B3")
    DefinirNome "cfgIntervalo", ws.Range("B4")
    DefinirNome "cfgDiasUteis", ws.Range("B5")
    DefinirNome "cfgTimeout", ws.Range("B6")
    DefinirNome "cfgVerTodos", ws.Range("B7")
    DefinirNome "cfgUltimoCarimbo", ws.Range("B14")
    DefinirNome "cfgUltimaSync", ws.Range("B15")
    DefinirNome "inID", ws.Range("B16")
    DefinirNome "inVersao", ws.Range("B17")
    DefinirNome "cfgEditandoUsuarios", ws.Range("B18")
    DefinirNome "Feriados", ws.Range("E3:E200")
    DefinirNome "lstStatus", ws.Range("G3:G7")
    DefinirNome "lstPerfil", ws.Range("H3:H5")
    DefinirNome "lstSimNao", ws.Range("I3:I4")
    ThisWorkbook.Names.Add Name:="lstDisciplinas", _
        RefersTo:="=OFFSET('" & SH_CFG & "'!$K$3,0,0,MAX(1,COUNTA('" & SH_CFG & "'!$K$3:$K$60)),1)"

    ValidarLista ws.Range("B5"), "lstSimNao", True
    ValidarLista ws.Range("B7"), "lstSimNao", True

    x = ws.Range("A10").Left
    CriarBotao ws, x, ws.Range("A10").Top, 150, "Selecionar banco...", "SelecionarArquivoBD", COR_BOTAO_AZUL
    CriarBotao ws, x + 158, ws.Range("A10").Top, 150, "Conectar / Sincronizar", "Iniciar", COR_BOTAO_VERDE
    CriarBotao ws, x + 316, ws.Range("A10").Top, 150, "Criar banco na rede", "CriarBancoDeDados", COR_BOTAO_CINZA
    ws.Rows(10).RowHeight = 30
    SemGrade ws
End Sub

'------------------------------------------------------------------------------
' USUARIOS
'------------------------------------------------------------------------------
Private Sub MontarUsuarios(ByVal ws As Worksheet)
    Dim lo As ListObject, x As Double

    LimparAba ws
    TituloAba ws, "Usuários e hierarquia"
    ws.Range("A2").Value = "Login = usuário de rede do Windows. Gestor = login do chefe imediato " & _
        "(o chefe pode cadastrar projetos para toda a sua equipe). Alterações somente por Admin."
    ws.Range("A2").Font.Color = RGB(89, 89, 89)
    ws.Columns("A").ColumnWidth = 20
    ws.Columns("B").ColumnWidth = 32
    ws.Columns("C").ColumnWidth = 20
    ws.Columns("D").ColumnWidth = 20
    ws.Columns("E").ColumnWidth = 14

    ws.Range("A6").Resize(1, U_NCOLS).Value = Array("Login", "Nome", "Disciplina", "Gestor (login)", "Perfil")
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A6").Resize(2, U_NCOLS), , xlYes)
    lo.Name = TBL_USR
    lo.TableStyle = "TableStyleMedium2"
    ThisWorkbook.Names.Add Name:="lstColaboradores", RefersTo:="=" & TBL_USR & "[Login]"

    x = ws.Range("A4").Left
    CriarBotao ws, x, ws.Range("A4").Top, 140, "Editar usuários", "EditarUsuarios", COR_BOTAO_AZUL
    CriarBotao ws, x + 148, ws.Range("A4").Top, 140, "Publicar usuários", "PublicarUsuarios", COR_BOTAO_VERDE
    CriarBotao ws, x + 296, ws.Range("A4").Top, 140, "Cancelar edição", "CancelarEdicaoUsuarios", COR_BOTAO_CINZA
    ws.Rows(4).RowHeight = 30
    SemGrade ws, 6
End Sub

'------------------------------------------------------------------------------
' PROJETOS
'------------------------------------------------------------------------------
Private Sub MontarProjetos(ByVal ws As Worksheet)
    Dim lo As ListObject, x As Double, t As Double, larg As Variant, i As Long
    Dim cab As Variant

    LimparAba ws
    TituloAba ws, "Projetos"
    ws.Range("A2").Font.Color = RGB(89, 89, 89)

    larg = Array(14, 42, 16, 26, 16, 13, 13, 13, 15, 12, 20, 15, 17, 7)
    For i = 0 To UBound(larg)
        ws.Columns(i + 1).ColumnWidth = larg(i)
    Next i

    ' ---- formulário de cadastro ----
    ws.Range("A3").Font.Bold = True
    DefinirNome "inModo", ws.Range("A3")
    cab = Array("Código do projeto", "Descrição do projeto", "Colaborador (login)", "Nome", "Disciplina", _
                "Data início", "Duração estimada (dias)", "Data fim (calculada)", "Status", "% concluído")
    ws.Range("A4").Resize(1, 10).Value = cab
    With ws.Range("A4:J4")
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(84, 130, 53)
        .WrapText = True
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
    End With
    ws.Rows(4).RowHeight = 30
    ws.Rows(5).RowHeight = 22
    With ws.Range("A5:J5")
        .Interior.Color = COR_ENTRADA
        .Borders.Color = RGB(191, 191, 191)
        .VerticalAlignment = xlCenter
        .Locked = False
    End With
    ws.Range("D5,H5").Locked = True
    ws.Range("D5,H5").Interior.Color = RGB(242, 242, 242)

    DefinirNome "inCodigo", ws.Range("A5")
    DefinirNome "inDescricao", ws.Range("B5")
    DefinirNome "inColaborador", ws.Range("C5")
    DefinirNome "inNome", ws.Range("D5")
    DefinirNome "inDisciplina", ws.Range("E5")
    DefinirNome "inInicio", ws.Range("F5")
    DefinirNome "inDuracao", ws.Range("G5")
    DefinirNome "inFim", ws.Range("H5")
    DefinirNome "inStatus", ws.Range("I5")
    DefinirNome "inPerc", ws.Range("J5")

    ws.Range("A5").NumberFormat = "@"
    ws.Range("F5,H5").NumberFormat = "dd/mm/yyyy"
    ws.Range("G5").NumberFormat = "0"
    ws.Range("J5").NumberFormat = "0%"
    ws.Range("D5").Formula = "=IFERROR(INDEX(" & TBL_USR & "[Nome],MATCH(C5," & TBL_USR & "[Login],0)),"""")"
    ws.Range("H5").Formula = "=IF(OR(F5="""",G5=""""),"""",IF(cfgDiasUteis=""NAO"",F5+G5-1,WORKDAY(F5-1,G5,Feriados)))"

    ValidarLista ws.Range("C5"), "lstColaboradores", True
    ValidarLista ws.Range("E5"), "lstDisciplinas", False
    ValidarLista ws.Range("I5"), "lstStatus", True
    With ws.Range("F5").Validation
        .Delete
        .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, Operator:=xlGreater, _
             Formula1:=CStr(CLng(DateSerial(2000, 1, 1)))
        .ErrorMessage = "Informe uma data válida."
    End With
    With ws.Range("G5").Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
             Formula1:="1", Formula2:="3650"
        .ErrorMessage = "Informe a duração em dias (número inteiro de 1 a 3650)."
    End With
    With ws.Range("J5").Validation
        .Delete
        .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
             Formula1:="0", Formula2:="1"
        .ErrorMessage = "Informe um percentual entre 0% e 100%."
    End With

    ' ---- botões ----
    ws.Rows(6).RowHeight = 32
    x = ws.Range("A6").Left + 2
    t = ws.Range("A6").Top + 4
    CriarBotao ws, x, t, 125, "Salvar projeto", "SalvarProjeto", COR_BOTAO_VERDE
    CriarBotao ws, x + 133, t, 125, "Novo / Limpar", "NovoProjeto", COR_BOTAO_CINZA
    CriarBotao ws, x + 266, t, 125, "Editar selecionado", "EditarSelecionado", COR_BOTAO_AZUL
    CriarBotao ws, x + 399, t, 125, "Excluir selecionado", "ExcluirSelecionado", COR_BOTAO_VERMELHO
    CriarBotao ws, x + 532, t, 125, "Atualizar agora", "AtualizarAgora", COR_BOTAO_AZUL
    CriarBotao ws, x + 665, t, 125, "Ver Dashboard", "IrParaDashboard", COR_BOTAO_AZUL

    ws.Range("A7").Value = "Para editar ou excluir: clique em uma linha da lista e use os botões. " & _
        "A lista e o Dashboard são atualizados automaticamente quando qualquer usuário grava."
    ws.Range("A7").Font.Italic = True
    ws.Range("A7").Font.Color = RGB(128, 128, 128)

    ' ---- lista (tabela) ----
    ws.Range("A8").Value = "Lista de projetos"
    ws.Range("A8").Font.Bold = True
    ws.Range("A8").Font.Size = 12
    ws.Range("A9").Resize(1, L_NCOLS).Value = Array("Código", "Descrição", "Colaborador", "Nome", "Disciplina", _
        "Data início", "Duração (dias)", "Data fim", "Status", "% concluído", "Situação", "Atualizado por", _
        "Atualizado em", "ID")
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A9").Resize(2, L_NCOLS), , xlYes)
    lo.Name = TBL_PROJ
    lo.TableStyle = "TableStyleMedium2"
    lo.HeaderRowRange.WrapText = True
    ws.Rows(9).RowHeight = 30

    SemGrade ws, 9
End Sub

'------------------------------------------------------------------------------
' DASHBOARD
'------------------------------------------------------------------------------
Private Sub MontarDashboard(ByVal ws As Worksheet)
    Dim larg As Variant, i As Long

    LimparAba ws
    TituloAba ws, "Dashboard de projetos por colaborador"
    larg = Array(2, 16, 34, 18, 12, 12, 12, 12, 12, 12, 12, 14, 12)
    For i = 0 To UBound(larg)
        ws.Columns(i + 1).ColumnWidth = larg(i)
    Next i
    ws.Range("B2").Font.Color = RGB(89, 89, 89)

    ws.Range("B3").Value = "Filtrar colaborador (vazio = todos):"
    ws.Range("B3").Font.Bold = True
    With ws.Range("D3")
        .Interior.Color = COR_ENTRADA
        .Borders.Color = RGB(191, 191, 191)
        .Locked = False
    End With
    DefinirNome "dashFiltro", ws.Range("D3")
    ValidarLista ws.Range("D3"), "lstColaboradores", True
    ws.Rows(3).RowHeight = 28
    CriarBotao ws, ws.Range("E3").Left + 6, ws.Range("E3").Top + 2, 130, "Atualizar dashboard", "AtualizarDashboardBotao", COR_BOTAO_AZUL
    CriarBotao ws, ws.Range("E3").Left + 144, ws.Range("E3").Top + 2, 130, "Ir para Projetos", "IrParaProjetos", COR_BOTAO_VERDE

    SemGrade ws
End Sub
