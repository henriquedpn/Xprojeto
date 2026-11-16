Attribute VB_Name = "modDashboard"
'==============================================================================
' Módulo: modDashboard - indicadores por colaborador
' É reconstruído automaticamente sempre que a lista de projetos é atualizada.
'==============================================================================
Option Explicit

Private Const LIN_KPI As Long = 5
Private Const LIN_RESUMO As Long = 9

' Índices das estatísticas por colaborador
Private Const E_TOTAL As Long = 1
Private Const E_NAOINI As Long = 2
Private Const E_ANDAM As Long = 3
Private Const E_PAUSA As Long = 4
Private Const E_ATRAS As Long = 5
Private Const E_CONCL As Long = 6
Private Const E_CANC As Long = 7
Private Const E_SOMAPERC As Long = 8
Private Const E_QTDPERC As Long = 9
Private Const E_PROXFIM As Long = 10
Private Const E_VENCE As Long = 11
Private Const E_N As Long = 11

Private Const COR_AZUL As Long = 7949855      ' RGB(31, 78, 121)
Private Const COR_VERDE As Long = 5287936     ' RGB(0, 176, 80)
Private Const COR_VERMELHO As Long = 3937500  ' RGB(220, 20, 60)
Private Const COR_AMBAR As Long = 49407       ' RGB(255, 192, 0)
Private Const COR_CINZA As Long = 10921638    ' RGB(166, 166, 166)

Public Sub AtualizarDashboardBotao()
    AtualizarDashboard
    ThisWorkbook.Worksheets(SH_DASH).Activate
End Sub

Public Sub AtualizarDashboard()
    Dim ws As Worksheet, lo As ListObject, v As Variant
    Dim filtro As String, colab As String, st As String, sit As String
    Dim idx As Object, logins() As String, est() As Double, tot(1 To E_N) As Double
    Dim i As Long, j As Long, k As Long, nCol As Long, nLin As Long, nProj As Long
    Dim res() As Variant, det() As Variant, grupos() As Long, nDet As Long
    Dim lin1 As Long, linFim As Long, linGraf As Long, linDet As Long, altura As Double
    Dim co As ChartObject, pm As Double

    On Error GoTo Falha
    Set ws = ThisWorkbook.Worksheets(SH_DASH)
    Set lo = ThisWorkbook.Worksheets(SH_PROJ).ListObjects(TBL_PROJ)
    filtro = LCase$(Texto(NR("dashFiltro").Value))

    ModoRapido True
    Desproteger ws
    If ws.ChartObjects.Count > 0 Then ws.ChartObjects.Delete
    With ws.Range(ws.Cells(4, 1), ws.Cells(ws.UsedRange.Row + ws.UsedRange.Rows.Count + 5, 26))
        .UnMerge
        .Clear
        .EntireRow.RowHeight = ws.StandardHeight
    End With
    ws.Range("B2").Value = "Atualizado em " & Format$(Now, "dd/mm/yyyy hh:nn:ss") & _
                           "  |  Usuário: " & NomeUsuario(UsuarioAtual()) & _
                           IIf(Len(filtro) > 0, "  |  Filtro: " & NomeUsuario(filtro), "")

    '--------------------------------------------------------------- coleta
    Set idx = CreateObject("Scripting.Dictionary")
    idx.CompareMode = vbTextCompare
    If Not lo.DataBodyRange Is Nothing Then
        v = lo.DataBodyRange.Value
        nLin = UBound(v, 1)
        ReDim logins(1 To nLin)
        ReDim est(1 To nLin, 1 To E_N)
    End If

    For i = 1 To nLin
        If NumLng(v(i, L_ID)) > 0 Then
            colab = LCase$(Texto(v(i, L_COLAB)))
            If Len(filtro) = 0 Or colab = filtro Then
                If Not idx.Exists(colab) Then
                    nCol = nCol + 1
                    idx.Add colab, nCol
                    logins(nCol) = colab
                End If
                k = idx(colab)
                nProj = nProj + 1
                st = Texto(v(i, L_STATUS))
                sit = Texto(v(i, L_SIT))

                est(k, E_TOTAL) = est(k, E_TOTAL) + 1
                If sit = SIT_ATRASADO Then
                    j = E_ATRAS
                ElseIf st = ST_CONCLUIDO Then
                    j = E_CONCL
                ElseIf st = ST_CANCELADO Then
                    j = E_CANC
                ElseIf st = ST_PAUSADO Then
                    j = E_PAUSA
                ElseIf st = ST_ANDAMENTO Then
                    j = E_ANDAM
                Else
                    j = E_NAOINI
                End If
                est(k, j) = est(k, j) + 1
                If sit = SIT_VENCE Then est(k, E_VENCE) = est(k, E_VENCE) + 1
                If st <> ST_CANCELADO Then
                    est(k, E_SOMAPERC) = est(k, E_SOMAPERC) + NumDbl(v(i, L_PERC))
                    est(k, E_QTDPERC) = est(k, E_QTDPERC) + 1
                End If
                If st <> ST_CONCLUIDO And st <> ST_CANCELADO And IsDate(v(i, L_FIM)) Then
                    If CDate(v(i, L_FIM)) >= Date Then
                        If est(k, E_PROXFIM) = 0 Or CDbl(CDate(v(i, L_FIM))) < est(k, E_PROXFIM) Then
                            est(k, E_PROXFIM) = CDbl(CDate(v(i, L_FIM)))
                        End If
                    End If
                End If
            End If
        End If
    Next i

    For k = 1 To nCol
        For j = 1 To E_N
            If j <> E_PROXFIM Then tot(j) = tot(j) + est(k, j)
        Next j
    Next k

    '--------------------------------------------------------------- KPIs
    Kpi ws, 2, "Projetos", tot(E_TOTAL), "0", COR_AZUL
    Kpi ws, 4, "Em andamento", tot(E_ANDAM), "0", RGB(91, 155, 213)
    Kpi ws, 6, "Atrasados", tot(E_ATRAS), "0", COR_VERMELHO
    Kpi ws, 8, "Vencem em até 7 dias", tot(E_VENCE), "0", COR_AMBAR
    Kpi ws, 10, "Concluídos", tot(E_CONCL), "0", COR_VERDE
    If tot(E_QTDPERC) > 0 Then pm = tot(E_SOMAPERC) / tot(E_QTDPERC)
    Kpi ws, 12, "% médio concluído", pm, "0%", COR_CINZA

    '--------------------------------------------------------------- resumo
    ws.Cells(LIN_RESUMO - 1, 2).Value = "Resumo por colaborador"
    Titulo ws.Cells(LIN_RESUMO - 1, 2)
    ws.Cells(LIN_RESUMO, 2).Resize(1, 11).Value = Array("Login", "Nome", "Disciplina", "Total", _
        "Não iniciados", "Em andamento", "Pausados", "Atrasados", "Concluídos", "% médio", "Próximo término")
    Cabecalho ws.Cells(LIN_RESUMO, 2).Resize(1, 11)

    If nCol = 0 Then
        ws.Cells(LIN_RESUMO + 1, 2).Value = "Nenhum projeto para exibir."
        ws.Cells(LIN_RESUMO + 1, 2).Font.Italic = True
        GoTo Finalizar
    End If

    ReDim res(1 To nCol + 1, 1 To 11)
    For k = 1 To nCol
        res(k, 1) = logins(k)
        res(k, 2) = NomeUsuario(logins(k))
        res(k, 3) = DisciplinaUsuario(logins(k))
        res(k, 4) = est(k, E_TOTAL)
        res(k, 5) = est(k, E_NAOINI)
        res(k, 6) = est(k, E_ANDAM)
        res(k, 7) = est(k, E_PAUSA)
        res(k, 8) = est(k, E_ATRAS)
        res(k, 9) = est(k, E_CONCL)
        If est(k, E_QTDPERC) > 0 Then res(k, 10) = est(k, E_SOMAPERC) / est(k, E_QTDPERC) Else res(k, 10) = 0
        If est(k, E_PROXFIM) > 0 Then res(k, 11) = CDate(est(k, E_PROXFIM)) Else res(k, 11) = "-"
    Next k
    res(nCol + 1, 1) = "TOTAL"
    For j = 4 To 9
        res(nCol + 1, j) = 0
    Next j
    res(nCol + 1, 4) = tot(E_TOTAL): res(nCol + 1, 5) = tot(E_NAOINI): res(nCol + 1, 6) = tot(E_ANDAM)
    res(nCol + 1, 7) = tot(E_PAUSA): res(nCol + 1, 8) = tot(E_ATRAS): res(nCol + 1, 9) = tot(E_CONCL)
    If tot(E_QTDPERC) > 0 Then res(nCol + 1, 10) = tot(E_SOMAPERC) / tot(E_QTDPERC) Else res(nCol + 1, 10) = 0
    res(nCol + 1, 11) = ""

    lin1 = LIN_RESUMO + 1
    linFim = lin1 + nCol            ' linha do TOTAL
    With ws.Cells(lin1, 2).Resize(nCol + 1, 11)
        .Value = res
        .Borders.Color = RGB(217, 217, 217)
        .VerticalAlignment = xlCenter
    End With
    ws.Range(ws.Cells(lin1, 5), ws.Cells(linFim, 10)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(lin1, 11), ws.Cells(linFim, 11)).NumberFormat = "0%"
    ws.Range(ws.Cells(lin1, 12), ws.Cells(linFim, 12)).NumberFormat = "dd/mm/yyyy"
    ws.Range(ws.Cells(lin1, 12), ws.Cells(linFim, 12)).HorizontalAlignment = xlCenter
    With ws.Cells(linFim, 2).Resize(1, 11)
        .Font.Bold = True
        .Interior.Color = RGB(242, 242, 242)
        .Borders(xlEdgeTop).Weight = xlMedium
    End With
    ' destaques
    For k = 1 To nCol
        If est(k, E_ATRAS) > 0 Then
            ws.Cells(lin1 + k - 1, 9).Font.Color = RGB(192, 0, 0)
            ws.Cells(lin1 + k - 1, 9).Font.Bold = True
        End If
    Next k
    With ws.Range(ws.Cells(lin1, 11), ws.Cells(linFim - 1, 11)).FormatConditions.AddDatabar
        .MinPoint.Modify newtype:=xlConditionValueNumber, newvalue:=0
        .MaxPoint.Modify newtype:=xlConditionValueNumber, newvalue:=1
        .BarColor.Color = RGB(91, 155, 213)
    End With

    '--------------------------------------------------------------- gráfico
    linGraf = linFim + 2
    altura = 70 + nCol * 26
    If altura < 220 Then altura = 220
    If altura > 600 Then altura = 600
    Set co = ws.ChartObjects.Add(Left:=ws.Cells(linGraf, 2).Left, Top:=ws.Cells(linGraf, 2).Top, _
                                 Width:=ws.Range(ws.Cells(1, 2), ws.Cells(1, 12)).Width, Height:=altura)
    co.Name = "chtColaboradores"
    With co.Chart
        NovaSerie co.Chart, "Não iniciados", ws.Range(ws.Cells(lin1, 6), ws.Cells(linFim - 1, 6)), ws.Range(ws.Cells(lin1, 3), ws.Cells(linFim - 1, 3)), COR_CINZA
        NovaSerie co.Chart, "Em andamento", ws.Range(ws.Cells(lin1, 7), ws.Cells(linFim - 1, 7)), ws.Range(ws.Cells(lin1, 3), ws.Cells(linFim - 1, 3)), RGB(91, 155, 213)
        NovaSerie co.Chart, "Pausados", ws.Range(ws.Cells(lin1, 8), ws.Cells(linFim - 1, 8)), ws.Range(ws.Cells(lin1, 3), ws.Cells(linFim - 1, 3)), COR_AMBAR
        NovaSerie co.Chart, "Atrasados", ws.Range(ws.Cells(lin1, 9), ws.Cells(linFim - 1, 9)), ws.Range(ws.Cells(lin1, 3), ws.Cells(linFim - 1, 3)), COR_VERMELHO
        NovaSerie co.Chart, "Concluídos", ws.Range(ws.Cells(lin1, 10), ws.Cells(linFim - 1, 10)), ws.Range(ws.Cells(lin1, 3), ws.Cells(linFim - 1, 3)), COR_VERDE
        .ChartType = xlBarStacked
        .HasTitle = True
        .ChartTitle.Text = "Projetos por colaborador e situação"
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        .Axes(xlCategory).ReversePlotOrder = True
        .ChartGroups(1).GapWidth = 60
    End With

    ' primeira linha livre abaixo do gráfico
    linDet = linGraf
    Do While ws.Cells(linDet, 2).Top < co.Top + co.Height
        linDet = linDet + 1
    Loop
    linDet = linDet + 1

    '--------------------------------------------------------------- detalhe
    ws.Cells(linDet, 2).Value = "Projetos por colaborador"
    Titulo ws.Cells(linDet, 2)
    ws.Cells(linDet + 1, 2).Resize(1, 9).Value = Array("Código", "Descrição", "Disciplina", "Início", _
        "Fim previsto", "Dias restantes", "Status", "% concluído", "Situação")
    Cabecalho ws.Cells(linDet + 1, 2).Resize(1, 9)

    ReDim det(1 To nProj + nCol, 1 To 9)
    ReDim grupos(1 To nCol)
    For k = 1 To nCol
        nDet = nDet + 1
        grupos(k) = nDet
        det(nDet, 1) = NomeUsuario(logins(k)) & "  (" & logins(k) & ")  -  " & est(k, E_TOTAL) & _
                       " projeto(s)" & IIf(est(k, E_ATRAS) > 0, "  |  " & est(k, E_ATRAS) & " atrasado(s)", "")
        For i = 1 To nLin
            If NumLng(v(i, L_ID)) > 0 Then
                If StrComp(Texto(v(i, L_COLAB)), logins(k), vbTextCompare) = 0 Then
                    nDet = nDet + 1
                    st = Texto(v(i, L_STATUS))
                    det(nDet, 1) = v(i, L_COD)
                    det(nDet, 2) = v(i, L_DESC)
                    det(nDet, 3) = v(i, L_DISC)
                    det(nDet, 4) = v(i, L_INI)
                    det(nDet, 5) = v(i, L_FIM)
                    If st <> ST_CONCLUIDO And st <> ST_CANCELADO And IsDate(v(i, L_FIM)) Then
                        det(nDet, 6) = CLng(CDate(v(i, L_FIM)) - Date)
                    Else
                        det(nDet, 6) = ""
                    End If
                    det(nDet, 7) = st
                    det(nDet, 8) = NumDbl(v(i, L_PERC))
                    det(nDet, 9) = v(i, L_SIT)
                End If
            End If
        Next i
    Next k

    lin1 = linDet + 2
    With ws.Cells(lin1, 2).Resize(nDet, 9)
        .Value = det
        .Borders.Color = RGB(217, 217, 217)
    End With
    ws.Range(ws.Cells(lin1, 5), ws.Cells(lin1 + nDet - 1, 6)).NumberFormat = "dd/mm/yyyy"
    ws.Range(ws.Cells(lin1, 5), ws.Cells(lin1 + nDet - 1, 7)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(lin1, 9), ws.Cells(lin1 + nDet - 1, 9)).NumberFormat = "0%"
    ws.Range(ws.Cells(lin1, 9), ws.Cells(lin1 + nDet - 1, 10)).HorizontalAlignment = xlCenter
    AplicarCoresSituacao ws.Range(ws.Cells(lin1, 10), ws.Cells(lin1 + nDet - 1, 10))
    For k = 1 To nCol
        With ws.Cells(lin1 + grupos(k) - 1, 2).Resize(1, 9)
            .Interior.Color = RGB(221, 235, 247)
            .Font.Bold = True
            .Font.Color = COR_AZUL
            .NumberFormat = "General"
        End With
    Next k

Finalizar:
    Proteger ws
    ModoRapido False
    Exit Sub
Falha:
    Application.StatusBar = APP_NOME & " - erro ao montar o Dashboard: " & Err.Description
    Proteger ws
    ModoRapido False
End Sub

Private Sub Kpi(ByVal ws As Worksheet, ByVal col As Long, ByVal rotulo As String, ByVal valor As Double, _
                ByVal formato As String, ByVal cor As Long)
    With ws.Cells(LIN_KPI, col).Resize(1, 2)
        .Merge
        .Value = rotulo
        .Font.Size = 9
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = cor
        .HorizontalAlignment = xlCenter
    End With
    With ws.Cells(LIN_KPI + 1, col).Resize(1, 2)
        .Merge
        .Value = valor
        .NumberFormat = formato
        .Font.Size = 22
        .Font.Bold = True
        .Font.Color = cor
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders(xlEdgeBottom).Color = cor
        .Borders(xlEdgeBottom).Weight = xlThick
    End With
    ws.Rows(LIN_KPI + 1).RowHeight = 36
End Sub

Private Sub Titulo(ByVal rg As Range)
    rg.Font.Size = 13
    rg.Font.Bold = True
    rg.Font.Color = COR_AZUL
End Sub

Private Sub Cabecalho(ByVal rg As Range)
    rg.Interior.Color = COR_AZUL
    rg.Font.Color = RGB(255, 255, 255)
    rg.Font.Bold = True
    rg.HorizontalAlignment = xlCenter
    rg.VerticalAlignment = xlCenter
    rg.WrapText = True
    rg.RowHeight = 30
End Sub

Private Sub NovaSerie(ByVal ch As Chart, ByVal nome As String, ByVal valores As Range, _
                      ByVal categorias As Range, ByVal cor As Long)
    Dim s As Series
    Set s = ch.SeriesCollection.NewSeries
    s.Name = nome
    s.Values = valores
    s.XValues = categorias
    s.Format.Fill.ForeColor.RGB = cor
End Sub
