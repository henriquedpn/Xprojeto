Attribute VB_Name = "modBanco"
'==============================================================================
' Módulo: modBanco - leitura e gravação no banco compartilhado (.xlsx na rede)
'
' Regras de concorrência:
'   - LEITURA: o banco é aberto SOMENTE LEITURA por alguns instantes e fechado.
'     Vários usuários podem ler ao mesmo tempo.
'   - GRAVAÇÃO: o usuário obtém o bloqueio exclusivo (modLock), abre o banco,
'     grava, salva, fecha e libera o bloqueio. As gravações entram em fila.
'   - EDIÇÃO: controle otimista. Ao salvar uma edição, compara-se a data/hora
'     da última alteração do registro com a que o usuário carregou. Se outra
'     pessoa alterou nesse meio tempo, a gravação é recusada.
'==============================================================================
Option Explicit

Private mAnterior As Workbook   ' pasta que estava ativa antes de abrir o banco

Public Function ArquivoExiste(ByVal caminho As String) As Boolean
    On Error Resume Next
    If Len(caminho) = 0 Then Exit Function
    ArquivoExiste = (Len(Dir$(caminho, vbNormal)) > 0)
    If Err.Number <> 0 Then ArquivoExiste = False
End Function

Public Function AbrirBD(ByVal somenteLeitura As Boolean) As Workbook
    Dim caminho As String, nomeArq As String, wb As Workbook
    caminho = CaminhoBD()
    If Not ArquivoExiste(caminho) Then
        Err.Raise vbObjectError + 1001, APP_NOME, "Banco de dados não encontrado:" & vbLf & caminho
    End If
    nomeArq = Mid$(caminho, InStrRev(caminho, "\") + 1)
    For Each wb In Application.Workbooks
        If StrComp(wb.Name, nomeArq, vbTextCompare) = 0 Then
            Err.Raise vbObjectError + 1002, APP_NOME, "O arquivo do banco (" & nomeArq & _
                ") está aberto neste Excel. Feche-o e tente novamente."
        End If
    Next wb
    Set mAnterior = ActiveWorkbook
    Set AbrirBD = Workbooks.Open(Filename:=caminho, UpdateLinks:=0, ReadOnly:=somenteLeitura, _
                                 Notify:=False, AddToMru:=False)
End Function

Public Sub FecharBD(ByVal wb As Workbook, ByVal salvar As Boolean)
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=salvar
    If Not mAnterior Is Nothing Then mAnterior.Activate   ' devolve o foco ao usuário
    Set mAnterior = Nothing
End Sub

' Lê a aba a partir da linha 2 e devolve uma matriz 2D (ou Empty se vazia)
Public Function LerTabela(ByVal ws As Worksheet, ByVal nCols As Long) As Variant
    Dim ultima As Long
    ultima = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ultima < 2 Then
        LerTabela = Empty
    Else
        LerTabela = ws.Range(ws.Cells(2, 1), ws.Cells(ultima, nCols)).Value
    End If
End Function

'------------------------------------------------------------------------------
' Verifica se o banco foi alterado desde a última sincronização
'------------------------------------------------------------------------------
Public Function ServidorMudou() As Boolean
    Dim c As String
    On Error GoTo Erro
    c = Format$(FileDateTime(CaminhoBD()), "yyyy-mm-dd hh:nn:ss")
    ServidorMudou = (c <> Texto(NR("cfgUltimoCarimbo").Value))
    Exit Function
Erro:
    ServidorMudou = False
End Function

'------------------------------------------------------------------------------
' Baixa projetos e usuários do servidor e atualiza lista + Dashboard
'------------------------------------------------------------------------------
Public Function CarregarDoServidor(ByVal silencioso As Boolean) As Boolean
    Dim wb As Workbook, vProj As Variant, vUsr As Variant, carimbo As Date, msg As String
    On Error GoTo Falha
    ModoRapido True
    carimbo = FileDateTime(CaminhoBD())     ' capturado ANTES de ler (seguro p/ concorrência)
    Set wb = AbrirBD(True)
    vProj = LerTabela(wb.Worksheets(BD_PROJ), B_NCOLS)
    vUsr = LerTabela(wb.Worksheets(BD_USR), U_NCOLS)
    FecharBD wb, False
    Set wb = Nothing

    GravarUsuariosLocal vUsr
    PreencherProjetosLocal vProj
    AtualizarDashboard

    NR("cfgUltimoCarimbo").Value = "'" & Format$(carimbo, "yyyy-mm-dd hh:nn:ss")
    NR("cfgUltimaSync").Value = Now
    AtualizarStatus
    Application.StatusBar = False
    CarregarDoServidor = True
Sair:
    ModoRapido False
    Exit Function
Falha:
    msg = Err.Description
    FecharBD wb, False
    If silencioso Then
        Application.StatusBar = APP_NOME & " - falha na sincronização (" & Format$(Now, "hh:nn") & "): " & msg
    Else
        MsgBox "Não foi possível carregar os dados do servidor:" & vbLf & msg, vbExclamation, APP_NOME
    End If
    Resume Sair
End Function

'------------------------------------------------------------------------------
' Abertura/fechamento para gravação (com bloqueio)
'------------------------------------------------------------------------------
Private Function AbrirParaEscrita(ByRef msgErro As String) As Workbook
    Dim wb As Workbook
    If Not AdquirirLock(CaminhoBD(), TimeoutLockSeg()) Then
        msgErro = "O banco está sendo atualizado por outro usuário. Tente novamente em alguns segundos."
        Exit Function
    End If
    On Error GoTo Falha
    Set wb = AbrirBD(False)
    If wb.ReadOnly Then
        FecharBD wb, False
        LiberarLock
        msgErro = "O arquivo do banco foi aberto diretamente no Excel por outra pessoa." & vbLf & _
                  "Peça para fechá-lo e tente novamente."
        Exit Function
    End If
    Set AbrirParaEscrita = wb
    Exit Function
Falha:
    msgErro = Err.Description
    FecharBD wb, False
    LiberarLock
End Function

Private Sub ConcluirEscrita(ByVal wb As Workbook)
    wb.Save
    FecharBD wb, False
    LiberarLock
End Sub

Private Sub AbortarEscrita(ByVal wb As Workbook)
    FecharBD wb, False
    LiberarLock
End Sub

Private Sub RegistrarAlteracao(ByVal wb As Workbook, ByVal acao As String, ByVal id As Long, _
                               ByVal codigo As String, ByVal colab As String)
    Dim c As Worksheet, h As Worksheet, r As Long
    Set c = wb.Worksheets(BD_CTRL)
    c.Range("B1").Value = NumLng(c.Range("B1").Value) + 1
    c.Range("B3").Value = Now
    c.Range("B4").Value = UsuarioAtual()
    On Error Resume Next
    Set h = wb.Worksheets(BD_HIST)
    On Error GoTo 0
    If Not h Is Nothing Then
        r = h.Cells(h.Rows.Count, 1).End(xlUp).Row + 1
        h.Cells(r, 1).Resize(1, 6).Value = Array(Now, UsuarioAtual(), acao, id, codigo, colab)
    End If
End Sub

'------------------------------------------------------------------------------
' Inclui (idEdicao = 0) ou altera um projeto. Retorna o ID gravado (0 = erro).
'------------------------------------------------------------------------------
Public Function GravarProjetoServidor(ByVal idEdicao As Long, ByVal versaoOriginal As String, _
        ByVal codigo As String, ByVal descricao As String, ByVal colab As String, _
        ByVal disciplina As String, ByVal inicio As Date, ByVal duracao As Long, _
        ByVal fim As Date, ByVal status As String, ByVal perc As Double, _
        ByRef msgErro As String) As Long

    Dim wb As Workbook, sh As Worksheet, shC As Worksheet, v As Variant
    Dim ultima As Long, r As Long, linha As Long, novoId As Long, maxId As Long, eu As String

    gOcupado = True
    ModoRapido True
    Set wb = AbrirParaEscrita(msgErro)
    If wb Is Nothing Then GoTo Sair

    On Error GoTo Falha
    eu = UsuarioAtual()
    Set sh = wb.Worksheets(BD_PROJ)
    Set shC = wb.Worksheets(BD_CTRL)
    ultima = sh.Cells(sh.Rows.Count, B_ID).End(xlUp).Row
    If ultima < 1 Then ultima = 1

    If ultima >= 2 Then
        v = sh.Range(sh.Cells(2, 1), sh.Cells(ultima, B_NCOLS)).Value
        For r = 1 To UBound(v, 1)
            If NumLng(v(r, B_ID)) > maxId Then maxId = NumLng(v(r, B_ID))
            If idEdicao > 0 And NumLng(v(r, B_ID)) = idEdicao Then linha = r + 1
            ' Um mesmo código pode ter vários colaboradores (multidisciplinar),
            ' mas não pode repetir para o MESMO colaborador.
            If StrComp(Texto(v(r, B_COD)), codigo, vbTextCompare) = 0 _
               And StrComp(Texto(v(r, B_COLAB)), colab, vbTextCompare) = 0 _
               And NumLng(v(r, B_ID)) <> idEdicao Then
                Err.Raise vbObjectError + 1011, APP_NOME, "O colaborador '" & colab & _
                    "' já possui o projeto '" & codigo & "'."
            End If
        Next r
    End If

    If idEdicao > 0 Then
        If linha = 0 Then Err.Raise vbObjectError + 1012, APP_NOME, _
            "Este projeto foi excluído por outro usuário."
        If DataTexto(sh.Cells(linha, B_ATU_EM).Value) <> versaoOriginal Then
            Err.Raise vbObjectError + 1013, APP_NOME, "Este projeto foi alterado por '" & _
                Texto(sh.Cells(linha, B_ATU_POR).Value) & "' em " & _
                Format$(sh.Cells(linha, B_ATU_EM).Value, "dd/mm/yyyy hh:nn") & _
                " depois que você o abriu para edição." & vbLf & _
                "A lista foi recarregada: clique em 'Editar selecionado' novamente."
        End If
        novoId = idEdicao
    Else
        novoId = NumLng(shC.Range("B2").Value)
        If maxId > novoId Then novoId = maxId
        novoId = novoId + 1
        shC.Range("B2").Value = novoId
        linha = ultima + 1
        sh.Cells(linha, B_CRIADO_POR).Value = eu
        sh.Cells(linha, B_CRIADO_EM).Value = Now
    End If

    sh.Cells(linha, B_ID).Value = novoId
    sh.Cells(linha, B_COD).Value = codigo
    sh.Cells(linha, B_DESC).Value = descricao
    sh.Cells(linha, B_COLAB).Value = colab
    sh.Cells(linha, B_DISC).Value = disciplina
    sh.Cells(linha, B_INI).Value = inicio
    sh.Cells(linha, B_DUR).Value = duracao
    sh.Cells(linha, B_FIM).Value = fim
    sh.Cells(linha, B_STATUS).Value = status
    sh.Cells(linha, B_PERC).Value = perc
    sh.Cells(linha, B_ATU_POR).Value = eu
    sh.Cells(linha, B_ATU_EM).Value = Now
    sh.Cells(linha, B_INI).NumberFormat = "dd/mm/yyyy"
    sh.Cells(linha, B_FIM).NumberFormat = "dd/mm/yyyy"

    RegistrarAlteracao wb, IIf(idEdicao > 0, "ALTERAR", "INCLUIR"), novoId, codigo, colab
    ConcluirEscrita wb
    GravarProjetoServidor = novoId
Sair:
    gOcupado = False
    ModoRapido False
    Exit Function
Falha:
    msgErro = Err.Description
    AbortarEscrita wb
    Resume Sair
End Function

'------------------------------------------------------------------------------
' Exclui um projeto
'------------------------------------------------------------------------------
Public Function ExcluirProjetoServidor(ByVal id As Long, ByVal versaoOriginal As String, _
                                       ByRef msgErro As String) As Boolean
    Dim wb As Workbook, sh As Worksheet, ultima As Long, r As Long, linha As Long

    gOcupado = True
    ModoRapido True
    Set wb = AbrirParaEscrita(msgErro)
    If wb Is Nothing Then GoTo Sair

    On Error GoTo Falha
    Set sh = wb.Worksheets(BD_PROJ)
    ultima = sh.Cells(sh.Rows.Count, B_ID).End(xlUp).Row
    For r = 2 To ultima
        If NumLng(sh.Cells(r, B_ID).Value) = id Then linha = r: Exit For
    Next r
    If linha = 0 Then Err.Raise vbObjectError + 1020, APP_NOME, "O projeto já havia sido excluído."
    If DataTexto(sh.Cells(linha, B_ATU_EM).Value) <> versaoOriginal Then
        Err.Raise vbObjectError + 1021, APP_NOME, "O projeto foi alterado por '" & _
            Texto(sh.Cells(linha, B_ATU_POR).Value) & "' após a última sincronização." & vbLf & _
            "A lista foi recarregada. Confira e tente novamente."
    End If

    RegistrarAlteracao wb, "EXCLUIR", id, Texto(sh.Cells(linha, B_COD).Value), Texto(sh.Cells(linha, B_COLAB).Value)
    sh.Rows(linha).Delete
    ConcluirEscrita wb
    ExcluirProjetoServidor = True
Sair:
    gOcupado = False
    ModoRapido False
    Exit Function
Falha:
    msgErro = Err.Description
    AbortarEscrita wb
    Resume Sair
End Function

'------------------------------------------------------------------------------
' Substitui o cadastro de usuários do servidor (somente Admin)
'------------------------------------------------------------------------------
Public Function PublicarUsuariosServidor(ByVal v As Variant, ByVal n As Long, _
                                         ByRef msgErro As String) As Boolean
    Dim wb As Workbook, sh As Worksheet, ultima As Long

    gOcupado = True
    ModoRapido True
    Set wb = AbrirParaEscrita(msgErro)
    If wb Is Nothing Then GoTo Sair

    On Error GoTo Falha
    Set sh = wb.Worksheets(BD_USR)
    ultima = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
    If ultima >= 2 Then sh.Range(sh.Cells(2, 1), sh.Cells(ultima, U_NCOLS)).ClearContents
    If n > 0 Then sh.Cells(2, 1).Resize(n, U_NCOLS).Value = v
    RegistrarAlteracao wb, "USUARIOS", 0, "", CStr(n) & " usuário(s)"
    ConcluirEscrita wb
    PublicarUsuariosServidor = True
Sair:
    gOcupado = False
    ModoRapido False
    Exit Function
Falha:
    msgErro = Err.Description
    AbortarEscrita wb
    Resume Sair
End Function

'------------------------------------------------------------------------------
' Cria o banco compartilhado (primeira instalação - administrador)
'------------------------------------------------------------------------------
Public Sub CriarBancoDeDados()
    Dim caminho As Variant, wb As Workbook, ws As Worksheet, eu As String

    caminho = CaminhoBD()
    If Len(caminho) = 0 Then
        caminho = Application.GetSaveAsFilename(InitialFileName:="BD_Projetos.xlsx", _
                  FileFilter:="Pasta de trabalho do Excel (*.xlsx), *.xlsx", _
                  Title:="Onde salvar o banco compartilhado? (use uma pasta da rede)")
        If VarType(caminho) = vbBoolean Then Exit Sub
        NR("cfgCaminho").Value = CStr(caminho)
    End If
    caminho = CStr(caminho)

    If ArquivoExiste(CStr(caminho)) Then
        MsgBox "Já existe um banco em:" & vbLf & caminho & vbLf & vbLf & _
               "Clique em 'Conectar' para usá-lo.", vbInformation, APP_NOME
        Exit Sub
    End If
    If MsgBox("Criar o banco de projetos em:" & vbLf & caminho & " ?", vbYesNo + vbQuestion, APP_NOME) = vbNo Then Exit Sub

    On Error GoTo Falha
    eu = UsuarioAtual()
    ModoRapido True
    Set wb = Workbooks.Add(xlWBATWorksheet)

    Set ws = wb.Worksheets(1)
    ws.Name = BD_PROJ
    ws.Range("A1").Resize(1, B_NCOLS).Value = Array("ID", "Codigo", "Descricao", "Colaborador", _
        "Disciplina", "DataInicio", "DuracaoDias", "DataFim", "Status", "PercConcluido", _
        "CriadoPor", "CriadoEm", "AtualizadoPor", "AtualizadoEm")
    ws.Rows(1).Font.Bold = True

    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = BD_USR
    ws.Range("A1").Resize(1, U_NCOLS).Value = Array("Login", "Nome", "Disciplina", "Gestor", "Perfil")
    ws.Range("A2").Resize(1, U_NCOLS).Value = Array(eu, Application.UserName, "Coordenação", "", PERFIL_ADMIN)
    ws.Rows(1).Font.Bold = True

    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = BD_CTRL
    ws.Range("A1:A4").Value = Application.WorksheetFunction.Transpose( _
        Array("Versao", "UltimoID", "UltimaAlteracao", "AlteradoPor"))
    ws.Range("B1").Value = 0
    ws.Range("B2").Value = 0
    ws.Range("B3").Value = Now
    ws.Range("B4").Value = eu

    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = BD_HIST
    ws.Range("A1").Resize(1, 6).Value = Array("DataHora", "Usuario", "Acao", "ID", "Codigo", "Colaborador")
    ws.Rows(1).Font.Bold = True

    wb.Worksheets(1).Activate
    wb.SaveAs Filename:=CStr(caminho), FileFormat:=xlOpenXMLWorkbook
    wb.Close SaveChanges:=False
    ModoRapido False

    MsgBox "Banco criado com sucesso!" & vbLf & vbLf & _
           "Você (" & eu & ") foi cadastrado como Admin." & vbLf & _
           "Cadastre os demais colaboradores na aba Usuarios.", vbInformation, APP_NOME
    Iniciar
    Exit Sub
Falha:
    MsgBox "Erro ao criar o banco: " & Err.Description, vbCritical, APP_NOME
    FecharBD wb, False
    ModoRapidoReset
End Sub

Public Sub SelecionarArquivoBD()
    Dim f As Variant
    f = Application.GetOpenFilename("Pasta de trabalho do Excel (*.xlsx), *.xlsx", , _
                                    "Selecione o banco de projetos compartilhado")
    If VarType(f) = vbBoolean Then Exit Sub
    NR("cfgCaminho").Value = CStr(f)
    Iniciar
End Sub
