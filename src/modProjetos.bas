Attribute VB_Name = "modProjetos"
'==============================================================================
' Módulo: modProjetos - aba "Projetos" (formulário de cadastro + lista)
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Monta a lista local a partir dos dados do servidor (respeitando permissões)
'------------------------------------------------------------------------------
Public Sub PreencherProjetosLocal(ByVal vProj As Variant)
    Dim ws As Worksheet, lo As ListObject, out() As Variant, i As Long, n As Long, colab As String

    Set ws = ThisWorkbook.Worksheets(SH_PROJ)
    Set lo = ws.ListObjects(TBL_PROJ)

    If IsArray(vProj) Then
        ReDim out(1 To UBound(vProj, 1), 1 To L_NCOLS)
        For i = 1 To UBound(vProj, 1)
            If NumLng(vProj(i, B_ID)) > 0 Then
                colab = LCase$(Texto(vProj(i, B_COLAB)))
                If PodeVer(colab) Then
                    n = n + 1
                    out(n, L_COD) = vProj(i, B_COD)
                    out(n, L_DESC) = vProj(i, B_DESC)
                    out(n, L_COLAB) = colab
                    out(n, L_NOME) = NomeUsuario(colab)
                    out(n, L_DISC) = vProj(i, B_DISC)
                    out(n, L_INI) = vProj(i, B_INI)
                    out(n, L_DUR) = vProj(i, B_DUR)
                    out(n, L_FIM) = vProj(i, B_FIM)
                    out(n, L_STATUS) = vProj(i, B_STATUS)
                    out(n, L_PERC) = NumDbl(vProj(i, B_PERC))
                    out(n, L_SIT) = SituacaoPrazo(vProj(i, B_STATUS), vProj(i, B_FIM))
                    out(n, L_ATU_POR) = vProj(i, B_ATU_POR)
                    out(n, L_ATU_EM) = vProj(i, B_ATU_EM)
                    out(n, L_ID) = vProj(i, B_ID)
                End If
            End If
        Next i
    End If

    ModoRapido True
    Desproteger ws
    EscreverTabela lo, out, n, L_NCOLS

    If n > 1 Then
        With lo.Sort
            .SortFields.Clear
            .SortFields.Add Key:=lo.ListColumns(L_NOME).DataBodyRange, Order:=xlAscending
            .SortFields.Add Key:=lo.ListColumns(L_INI).DataBodyRange, Order:=xlAscending
            .Header = xlYes
            .Apply
        End With
    End If

    lo.ListColumns(L_INI).DataBodyRange.NumberFormat = "dd/mm/yyyy"
    lo.ListColumns(L_FIM).DataBodyRange.NumberFormat = "dd/mm/yyyy"
    lo.ListColumns(L_ATU_EM).DataBodyRange.NumberFormat = "dd/mm/yyyy hh:mm"
    lo.ListColumns(L_PERC).DataBodyRange.NumberFormat = "0%"
    lo.ListColumns(L_DUR).DataBodyRange.NumberFormat = "0"

    AplicarCoresSituacao lo.ListColumns(L_SIT).DataBodyRange
    With lo.ListColumns(L_PERC).DataBodyRange
        .FormatConditions.Delete
        With .FormatConditions.AddDatabar
            .MinPoint.Modify newtype:=xlConditionValueNumber, newvalue:=0
            .MaxPoint.Modify newtype:=xlConditionValueNumber, newvalue:=1
            .BarColor.Color = RGB(91, 155, 213)
        End With
    End With

    Proteger ws
    ModoRapido False
End Sub

Public Sub AtualizarStatus()
    Dim ws As Worksheet, eu As String
    Set ws = ThisWorkbook.Worksheets(SH_PROJ)
    eu = UsuarioAtual()
    Desproteger ws
    ws.Range("A2").Value = "Usuário: " & NomeUsuario(eu) & " (" & eu & ")  |  Perfil: " & PerfilUsuario(eu) & _
        "  |  Última sincronização: " & Format$(Now, "dd/mm/yyyy hh:nn:ss") & _
        "  |  Atualização automática a cada " & IntervaloSeg() & " s"
    Proteger ws
End Sub

'------------------------------------------------------------------------------
' Formulário (linha 5 da aba Projetos)
'------------------------------------------------------------------------------
Public Sub NovoProjeto()
    LimparFormulario
    ThisWorkbook.Worksheets(SH_PROJ).Activate
    NR("inCodigo").Select
End Sub

Public Sub LimparFormulario()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_PROJ)
    Desproteger ws
    NR("inCodigo").Value = ""
    NR("inDescricao").Value = ""
    NR("inColaborador").Value = UsuarioAtual()
    NR("inDisciplina").Value = ""
    NR("inInicio").Value = Date
    NR("inDuracao").Value = ""
    NR("inStatus").Value = ST_NAO_INICIADO
    NR("inPerc").Value = 0
    NR("inID").Value = ""
    NR("inVersao").Value = ""
    NR("inModo").Value = "NOVO PROJETO - preencha a linha amarela e clique em 'Salvar projeto'"
    NR("inModo").Font.Color = RGB(0, 97, 0)
    Proteger ws
End Sub

Public Sub SalvarProjeto()
    Dim cod As String, desc As String, colab As String, disc As String, status As String
    Dim ini As Date, dur As Long, fim As Date, perc As Double, msg As String
    Dim idEd As Long, versao As String, novoId As Long, v As Variant

    If Len(CaminhoBD()) = 0 Then MsgBox "Configure o banco na aba Config.", vbExclamation, APP_NOME: Exit Sub

    cod = Texto(NR("inCodigo").Value)
    desc = Texto(NR("inDescricao").Value)
    colab = LCase$(Texto(NR("inColaborador").Value))
    disc = Texto(NR("inDisciplina").Value)
    status = Texto(NR("inStatus").Value)

    If Len(cod) = 0 Then msg = msg & "- Informe o código do projeto." & vbLf
    If Len(desc) = 0 Then msg = msg & "- Informe a descrição do projeto." & vbLf
    If Len(colab) = 0 Then
        msg = msg & "- Informe o colaborador." & vbLf
    ElseIf Not UsuarioExiste(colab) Then
        msg = msg & "- Colaborador '" & colab & "' não cadastrado." & vbLf
    End If
    v = NR("inInicio").Value
    If IsDate(v) Then ini = CDate(v) Else msg = msg & "- Data de início inválida." & vbLf
    v = NR("inDuracao").Value
    If NumDbl(v) >= 1 And NumDbl(v) = Int(NumDbl(v)) Then
        dur = CLng(v)
    Else
        msg = msg & "- Duração estimada deve ser um número inteiro de dias (>= 1)." & vbLf
    End If
    If Len(msg) > 0 Then
        MsgBox "Verifique os campos:" & vbLf & vbLf & msg, vbExclamation, APP_NOME
        Exit Sub
    End If

    If Not PodeGerenciar(colab) Then
        MsgBox "Você não tem permissão para cadastrar/alterar projetos de " & NomeUsuario(colab) & "." & vbLf & _
               "Somente o próprio colaborador, seus gestores ou um Admin podem fazê-lo.", vbExclamation, APP_NOME
        Exit Sub
    End If

    If Len(disc) = 0 Then disc = DisciplinaUsuario(colab)
    If Len(status) = 0 Then status = ST_NAO_INICIADO
    perc = NumDbl(NR("inPerc").Value)
    If perc > 1 Then perc = perc / 100
    If perc < 0 Then perc = 0
    If perc > 1 Then perc = 1
    If status = ST_CONCLUIDO Then perc = 1
    fim = CalcularDataFim(ini, dur)
    idEd = NumLng(NR("inID").Value)
    versao = Texto(NR("inVersao").Value)

    Application.StatusBar = APP_NOME & ": gravando no servidor..."
    novoId = GravarProjetoServidor(idEd, versao, cod, desc, colab, disc, ini, dur, fim, status, perc, msg)
    Application.StatusBar = False

    If novoId = 0 Then
        CarregarDoServidor True
        MsgBox "Não foi possível salvar:" & vbLf & vbLf & msg, vbExclamation, APP_NOME
        Exit Sub
    End If

    LimparFormulario
    CarregarDoServidor True
    MsgBox "Projeto '" & cod & "' salvo para " & NomeUsuario(colab) & "." & vbLf & _
           "Término previsto: " & Format$(fim, "dd/mm/yyyy"), vbInformation, APP_NOME
End Sub

'------------------------------------------------------------------------------
' Linha selecionada na lista (0 = nenhuma)
'------------------------------------------------------------------------------
Private Function LinhaSelecionada(ByVal lo As ListObject) As Long
    If ActiveSheet Is Nothing Then Exit Function
    If ActiveSheet.Name <> SH_PROJ Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    If Intersect(ActiveCell, lo.DataBodyRange) Is Nothing Then Exit Function
    LinhaSelecionada = ActiveCell.Row - lo.DataBodyRange.Row + 1
    If NumLng(lo.DataBodyRange.Cells(LinhaSelecionada, L_ID).Value) = 0 Then LinhaSelecionada = 0
End Function

Public Sub EditarSelecionado()
    Dim ws As Worksheet, lo As ListObject, i As Long, colab As String, rg As Range

    Set ws = ThisWorkbook.Worksheets(SH_PROJ)
    Set lo = ws.ListObjects(TBL_PROJ)
    i = LinhaSelecionada(lo)
    If i = 0 Then
        MsgBox "Clique em uma linha da lista de projetos e tente novamente.", vbInformation, APP_NOME
        Exit Sub
    End If
    Set rg = lo.DataBodyRange.Rows(i)
    colab = LCase$(Texto(rg.Cells(1, L_COLAB).Value))
    If Not PodeGerenciar(colab) Then
        MsgBox "Você pode visualizar, mas não alterar, os projetos de " & NomeUsuario(colab) & ".", vbInformation, APP_NOME
        Exit Sub
    End If

    Desproteger ws
    NR("inCodigo").Value = rg.Cells(1, L_COD).Value
    NR("inDescricao").Value = rg.Cells(1, L_DESC).Value
    NR("inColaborador").Value = colab
    NR("inDisciplina").Value = rg.Cells(1, L_DISC).Value
    NR("inInicio").Value = rg.Cells(1, L_INI).Value
    NR("inDuracao").Value = rg.Cells(1, L_DUR).Value
    NR("inStatus").Value = rg.Cells(1, L_STATUS).Value
    NR("inPerc").Value = rg.Cells(1, L_PERC).Value
    NR("inID").Value = rg.Cells(1, L_ID).Value
    NR("inVersao").Value = "'" & DataTexto(rg.Cells(1, L_ATU_EM).Value)
    NR("inModo").Value = "EDITANDO o projeto '" & rg.Cells(1, L_COD).Value & "' de " & NomeUsuario(colab) & _
                         " - altere a linha amarela e clique em 'Salvar projeto' (ou 'Novo / Limpar' para cancelar)"
    NR("inModo").Font.Color = RGB(192, 80, 0)
    Proteger ws
    NR("inDescricao").Select
End Sub

Public Sub ExcluirSelecionado()
    Dim ws As Worksheet, lo As ListObject, i As Long, colab As String, cod As String
    Dim id As Long, versao As String, msg As String

    Set ws = ThisWorkbook.Worksheets(SH_PROJ)
    Set lo = ws.ListObjects(TBL_PROJ)
    i = LinhaSelecionada(lo)
    If i = 0 Then
        MsgBox "Clique em uma linha da lista de projetos e tente novamente.", vbInformation, APP_NOME
        Exit Sub
    End If
    colab = LCase$(Texto(lo.DataBodyRange.Cells(i, L_COLAB).Value))
    cod = Texto(lo.DataBodyRange.Cells(i, L_COD).Value)
    id = NumLng(lo.DataBodyRange.Cells(i, L_ID).Value)
    versao = DataTexto(lo.DataBodyRange.Cells(i, L_ATU_EM).Value)

    If Not PodeGerenciar(colab) Then
        MsgBox "Você não tem permissão para excluir projetos de " & NomeUsuario(colab) & ".", vbExclamation, APP_NOME
        Exit Sub
    End If
    If MsgBox("Excluir o projeto '" & cod & "' de " & NomeUsuario(colab) & "?" & vbLf & _
              "Esta ação não pode ser desfeita.", vbYesNo + vbExclamation + vbDefaultButton2, APP_NOME) = vbNo Then Exit Sub

    If ExcluirProjetoServidor(id, versao, msg) Then
        If NumLng(NR("inID").Value) = id Then LimparFormulario
        CarregarDoServidor True
        MsgBox "Projeto excluído.", vbInformation, APP_NOME
    Else
        CarregarDoServidor True
        MsgBox "Não foi possível excluir:" & vbLf & vbLf & msg, vbExclamation, APP_NOME
    End If
End Sub

Public Sub AtualizarAgora()
    If CarregarDoServidor(False) Then
        Application.StatusBar = APP_NOME & ": dados atualizados às " & Format$(Now, "hh:nn:ss")
    End If
End Sub

Public Sub IrParaDashboard()
    ThisWorkbook.Worksheets(SH_DASH).Activate
End Sub

Public Sub IrParaProjetos()
    ThisWorkbook.Worksheets(SH_PROJ).Activate
End Sub
