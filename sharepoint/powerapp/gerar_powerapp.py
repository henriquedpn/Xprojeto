#!/usr/bin/env python3
"""
Gera o Power App (canvas) "Gerenciador de Projetos" a partir da definição da
lista do SharePoint (../definicao-lista.json) e empacota em um arquivo .msapp
com a Power Platform CLI (pac).

Telas geradas
-------------
  scrProjetos   lista de projetos (Meus / Minha equipe / Todos), busca, status,
                barra de progresso e situação do prazo
  scrFormulario inclusão/edição/exclusão com cálculo ao vivo da data fim
  scrDashboard  indicadores, resumo por colaborador com barra empilhada e lista
                de atrasados / projetos do colaborador selecionado

Uso
---
    python gerar_powerapp.py                     # gera fontes + GerenciadorProjetos.msapp
    python gerar_powerapp.py --somente-fontes    # só gera a pasta de fontes (sem pac)
    python gerar_powerapp.py --saida meuapp.msapp --nome "Projetos Engenharia"

Requisito para empacotar: Power Platform CLI ("pac"), versão 2.15.x
    dotnet tool install --global Microsoft.PowerApps.CLI.Tool --version 2.15.1

Depois de gerar: Power Apps Studio > Abrir > Procurar arquivos > selecione o
.msapp > Dados > Adicionar dados > SharePoint > site > lista "Projetos".
"""
from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import sys
import uuid
from dataclasses import dataclass, field
from pathlib import Path

AQUI = Path(__file__).resolve().parent
BASE = AQUI / "base"
DEFINICAO = AQUI.parent / "definicao-lista.json"

LARGURA, ALTURA = 1366, 768

# --------------------------------------------------------------------------- cores
AZUL = "RGBA(0, 120, 212, 1)"
AZUL_ESCURO = "RGBA(31, 78, 121, 1)"
VERDE = "RGBA(16, 124, 16, 1)"
VERMELHO = "RGBA(164, 38, 44, 1)"
AMBAR = "RGBA(193, 156, 0, 1)"
AMBAR_TEXTO = "RGBA(138, 83, 0, 1)"
CINZA = "RGBA(161, 159, 157, 1)"
CINZA_ESCURO = "RGBA(96, 94, 92, 1)"
CINZA_CLARO = "RGBA(243, 242, 241, 1)"
LINHA = "RGBA(225, 223, 221, 1)"
BRANCO = "RGBA(255, 255, 255, 1)"
TEXTO = "RGBA(50, 49, 48, 1)"


# --------------------------------------------------------------------------- modelo
@dataclass
class Controle:
    nome: str
    modelo: str                      # ex.: "label", "gallery.galleryVertical", "icon.Add"
    props: dict = field(default_factory=dict)
    filhos: list = field(default_factory=list)

    def todos(self):
        yield self
        for f in self.filhos:
            yield from f.todos()


def _formula_yaml(prop: str, formula: str, recuo: str) -> list[str]:
    """Serializa uma propriedade no formato .fx.yaml (layout 0.30)."""
    formula = formula.strip()
    if "\n" in formula or ":" in formula or "#" in formula:
        linhas = formula.split("\n")
        saida = [f"{recuo}{prop}: |", f"{recuo}    ={linhas[0]}"]
        saida += [f"{recuo}    {ln}" if ln.strip() else "" for ln in linhas[1:]]
        return saida
    return [f"{recuo}{prop}: ={formula}"]


def controle_yaml(c: Controle, nivel: int = 0) -> list[str]:
    recuo = "    " * nivel
    linhas = [f"{recuo}{c.nome} As {c.modelo}:"]
    for prop, formula in c.props.items():
        linhas += _formula_yaml(prop, formula, recuo + "    ")
    for z, filho in enumerate(c.filhos, start=1):
        filho.props.setdefault("ZIndex", str(z))
        linhas.append("")
        linhas += controle_yaml(filho, nivel + 1)
    return linhas


# --------------------------------------------------------------------------- atalhos
def txt(s: str) -> str:
    """Literal de texto Power Fx."""
    return '"' + s.replace('"', '""') + '"'


def pos(x, y, w, h) -> dict:
    return {"X": str(x), "Y": str(y), "Width": str(w), "Height": str(h)}


def rotulo(nome, texto, x, y, w, h, tamanho=12, cor=TEXTO, negrito=False, **extra) -> Controle:
    p = {"Text": texto, **pos(x, y, w, h), "Size": str(tamanho), "Color": cor}
    if negrito:
        p["FontWeight"] = "FontWeight.Semibold"
    p.update(extra)
    return Controle(nome, "label", p)


def botao(nome, texto, x, y, w, ao_clicar, cor=AZUL, h=40, **extra) -> Controle:
    p = {"Text": txt(texto), **pos(x, y, w, h), "Fill": cor, "Color": BRANCO,
         "HoverFill": f"ColorFade({cor}, -15%)", "PressedFill": f"ColorFade({cor}, -30%)",
         "BorderThickness": "0", "Size": "12",
         "RadiusTopLeft": "4", "RadiusTopRight": "4", "RadiusBottomLeft": "4", "RadiusBottomRight": "4",
         "OnSelect": ao_clicar}
    p.update(extra)
    return Controle(nome, "button", p)


def retangulo(nome, x, y, w, h, cor, **extra) -> Controle:
    return Controle(nome, "rectangle", {**pos(x, y, w, h), "Fill": cor, **extra})


def cabecalho(prefixo: str, titulo: str, botoes: list[Controle]) -> list[Controle]:
    return [retangulo(f"rec{prefixo}Topo", 0, 0, "Parent.Width", 64, AZUL_ESCURO),
            rotulo(f"lbl{prefixo}Titulo", txt(titulo), 24, 0, 600, 64, tamanho=20, cor=BRANCO, negrito=True),
            *botoes]


# --------------------------------------------------------------------------- regras (Power Fx)
def aberto(r: str) -> str:
    return f'!({r}StatusProjeto.Value in ["Concluído", "Cancelado"])'


def atrasado(r: str) -> str:
    return f"({aberto(r)} && {r}DataFim < Today())"


def situacao(r: str) -> str:
    return (f'If({r}StatusProjeto.Value = "Concluído", "Concluído", '
            f'{r}StatusProjeto.Value = "Cancelado", "Cancelado", '
            f'{r}DataFim < Today(), "Atrasado", '
            f'DateDiff(Today(), {r}DataFim, Days) <= 7, "Vence em até 7 dias", "No prazo")')


def cor_situacao(r: str) -> str:
    return (f'Switch({situacao(r)}, "Atrasado", {VERMELHO}, "Vence em até 7 dias", {AMBAR_TEXTO}, '
            f'"Concluído", {VERDE}, {CINZA_ESCURO})')


def cor_status(expr: str) -> str:
    return (f'Switch({expr}, "Concluído", {VERDE}, "Em andamento", {AZUL}, "Pausado", {AMBAR}, '
            f'"Cancelado", {CINZA_ESCURO}, {CINZA})')


def data_fim(dias_uteis: bool) -> str:
    """Mesma regra da coluna calculada DataFim (seg-sex; início no fim de semana vai para segunda)."""
    if not dias_uteis:
        return "DateAdd(dteInicio.SelectedDate, Value(txtDuracao.Text) - 1, Days)"
    return ("With({ini: dteInicio.SelectedDate, n: Value(txtDuracao.Text) - 1, "
            "w: Mod(Weekday(dteInicio.SelectedDate) + 5, 7)},\n"
            "    DateAdd(ini, If(w > 4, 7 - w, 0) + n + RoundDown((If(w > 4, 0, w) + n) / 5, 0) * 2, Days))")


# --------------------------------------------------------------------------- telas
def app_onstart() -> Controle:
    return Controle("App", "appinfo", {
        "BackEnabled": "false",
        "OnStart": "Set(gItem, Blank());\nSet(gEmail, User().Email)",
    })


def tela_projetos(lista: str) -> Controle:
    filtro = (f"Sort(\n"
              f"    Filter({lista},\n"
              f'        (drpVisao.Selected.Value = "Todos"\n'
              f'            || (drpVisao.Selected.Value = "Meus projetos" && Colaborador.Email = gEmail)\n'
              f'            || (drpVisao.Selected.Value = "Minha equipe" && Gestor.Email = gEmail))\n'
              f"        && (IsBlank(txtBusca.Text) || txtBusca.Text in Title || txtBusca.Text in Descricao)\n"
              f"    ),\n"
              f"    DataInicio, SortOrder.Ascending)")

    sel = "Select(Parent)"
    modelo = [
        retangulo("recPrjStatus", 0, 8, 6, "Parent.TemplateHeight - 16",
                  cor_status("ThisItem.StatusProjeto.Value"), OnSelect=sel),
        rotulo("lblPrjCodigo", "ThisItem.Title", 20, 6, 200, 26, tamanho=13, negrito=True, OnSelect=sel),
        rotulo("lblPrjDescricao", "ThisItem.Descricao", 220, 6, 560, 26, tamanho=12, OnSelect=sel),
        rotulo("lblPrjPessoa",
               'ThisItem.Colaborador.DisplayName & If(IsBlank(ThisItem.Disciplina.Value), "", "  •  " & ThisItem.Disciplina.Value)',
               20, 34, 420, 22, tamanho=11, cor=CINZA_ESCURO, OnSelect=sel),
        rotulo("lblPrjDatas",
               '"Início " & Text(ThisItem.DataInicio, "dd/mm/yyyy") & "   Fim " & Text(ThisItem.DataFim, "dd/mm/yyyy") '
               '& "   (" & ThisItem.DuracaoDias & " dias)"',
               20, 58, 520, 22, tamanho=11, cor=CINZA_ESCURO, OnSelect=sel),
        rotulo("lblPrjStatus", "ThisItem.StatusProjeto.Value", 800, 10, 130, 26, tamanho=11, cor=BRANCO,
               negrito=True, Fill=cor_status("ThisItem.StatusProjeto.Value"), Align="Align.Center", OnSelect=sel),
        retangulo("recPrjProgFundo", 800, 50, 130, 10, LINHA, OnSelect=sel),
        retangulo("recPrjProg", 800, 50, "130 * Min(100, Coalesce(ThisItem.PercConcluido, 0)) / 100", 10,
                  f"If(Coalesce(ThisItem.PercConcluido, 0) >= 100, {VERDE}, {AZUL})", OnSelect=sel),
        rotulo("lblPrjPerc", 'Coalesce(ThisItem.PercConcluido, 0) & "%"', 936, 42, 60, 26, tamanho=11, OnSelect=sel),
        rotulo("lblPrjSituacao", situacao("ThisItem."), 1010, 26, 210, 30, tamanho=12,
               cor=cor_situacao("ThisItem."), negrito=True, OnSelect=sel),
        Controle("icoPrjAbrir", "icon.ChevronRight", {**pos(1240, 30, 28, 28), "Color": CINZA_ESCURO,
                                                       "OnSelect": sel}),
        retangulo("recPrjLinha", 0, "Parent.TemplateHeight - 1", "Parent.TemplateWidth", 1, LINHA),
    ]
    galeria = Controle("galProjetos", "gallery.galleryVertical", {
        "Items": filtro, **pos(24, 168, 1318, 580), "TemplateSize": "88", "TemplatePadding": "0",
        "ShowScrollbar": "true", "Fill": BRANCO,
        "OnSelect": "Set(gItem, ThisItem);\nNavigate(scrFormulario, ScreenTransition.None)",
    }, modelo)

    filhos = cabecalho("Prj", "Projetos", [
        botao("btnPrjNovo", "+ Novo projeto", 930, 12, 180,
              "Set(gItem, Blank());\nNavigate(scrFormulario, ScreenTransition.None)", cor=VERDE),
        botao("btnPrjDashboard", "Dashboard", 1124, 12, 180, "Navigate(scrDashboard, ScreenTransition.None)"),
    ]) + [
        rotulo("lblPrjVisao", txt("Exibir"), 24, 80, 200, 24, tamanho=11, cor=CINZA_ESCURO),
        Controle("drpVisao", "dropdown", {"Items": '["Meus projetos", "Minha equipe", "Todos"]',
                                          "Default": txt("Meus projetos"), **pos(24, 106, 260, 40)}),
        rotulo("lblPrjBusca", txt("Buscar código ou descrição"), 304, 80, 300, 24, tamanho=11, cor=CINZA_ESCURO),
        Controle("txtBusca", "text", {"Default": '""', "HintText": txt("Digite para filtrar..."),
                                      **pos(304, 106, 420, 40)}),
        rotulo("lblPrjContagem", 'CountRows(galProjetos.AllItems) & " projeto(s)"', 1100, 106, 240, 40,
               tamanho=12, cor=CINZA_ESCURO, Align="Align.Right"),
        galeria,
        rotulo("lblPrjVazio", txt("Nenhum projeto encontrado."), 24, 200, 600, 40, tamanho=13,
               cor=CINZA_ESCURO, Visible="CountRows(galProjetos.AllItems) = 0"),
    ]
    return Controle("scrProjetos", "screen", {"Fill": CINZA_CLARO}, filhos)


def tela_formulario(lista: str, dias_uteis: bool) -> Controle:
    eu = ('{Claims: "i:0#.f|membership|" & Lower(User().Email), DisplayName: User().FullName, '
          'Email: User().Email, Department: "", JobTitle: "", Picture: ""}')

    def campo(nome, rotulo_txt, x, y, w=560):
        return rotulo(f"lbl{nome}", txt(rotulo_txt), x, y, w, 24, tamanho=11, cor=CINZA_ESCURO, negrito=True)

    ex, dx = 40, 700          # colunas esquerda / direita
    novo = "IsBlank(gItem)"
    entradas = [
        campo("FrmCodigo", "Código do projeto *", ex, 84),
        Controle("txtCodigo", "text", {"Default": f'If({novo}, "", gItem.Title)', **pos(ex, 110, 600, 40)}),
        campo("FrmDescricao", "Descrição", ex, 160),
        Controle("txtDescricao", "text", {"Default": f'If({novo}, "", gItem.Descricao)', "Mode": "TextMode.MultiLine",
                                          **pos(ex, 186, 600, 90)}),
        campo("FrmColaborador", "Colaborador *", ex, 286),
        Controle("cmbColaborador", "combobox", {
            "Items": f"Choices({lista}.Colaborador)", "SelectMultiple": "false", "IsSearchable": "true",
            "DisplayFields": '["DisplayName"]', "SearchFields": '["DisplayName", "Email"]',
            "DefaultSelectedItems": f"If({novo}, Table({eu}), Table(gItem.Colaborador))",
            **pos(ex, 312, 600, 40)}),
        campo("FrmGestor", "Gestor (chefe imediato)", ex, 362),
        Controle("cmbGestor", "combobox", {
            "Items": f"Choices({lista}.Gestor)", "SelectMultiple": "false", "IsSearchable": "true",
            "DisplayFields": '["DisplayName"]', "SearchFields": '["DisplayName", "Email"]',
            "DefaultSelectedItems": f"If({novo} || IsBlank(gItem.Gestor), Blank(), Table(gItem.Gestor))",
            **pos(ex, 388, 600, 40)}),
        campo("FrmDisciplina", "Disciplina", ex, 438),
        Controle("drpDisciplina", "dropdown", {
            "Items": f"Choices({lista}.Disciplina)", "AllowEmptySelection": "true",
            "Default": f'If({novo}, "", gItem.Disciplina.Value)', **pos(ex, 464, 600, 40)}),

        campo("FrmInicio", "Data início *", dx, 84, 280),
        Controle("dteInicio", "datepicker", {"DefaultDate": f"If({novo}, Today(), gItem.DataInicio)",
                                             **pos(dx, 110, 280, 40)}),
        campo("FrmDuracao", "Duração estimada (dias) *", dx + 300, 84, 326),
        Controle("txtDuracao", "text", {"Default": f'If({novo}, "", Text(gItem.DuracaoDias))',
                                        "Format": "TextFormat.Number", **pos(dx + 300, 110, 326, 40)}),
        rotulo("lblFrmFim",
               '"Data fim prevista: " & If(IsBlank(dteInicio.SelectedDate) || Value(txtDuracao.Text) < 1, "-",\n'
               f'    Text({data_fim(dias_uteis)}, "dd/mm/yyyy"))',
               dx, 160, 626, 36, tamanho=14, cor=AZUL_ESCURO, negrito=True),
        campo("FrmStatus", "Status", dx, 210),
        Controle("drpStatus", "dropdown", {"Items": f"Choices({lista}.StatusProjeto)",
                                           "Default": f'If({novo}, "Não iniciado", gItem.StatusProjeto.Value)',
                                           **pos(dx, 236, 626, 40)}),
        campo("FrmPerc", "% concluído", dx, 286),
        Controle("sldPerc", "slider", {"Min": "0", "Max": "100", "ShowValue": "true",
                                       "Default": f"If({novo}, 0, Coalesce(gItem.PercConcluido, 0))",
                                       **pos(dx, 318, 626, 40)}),
        campo("FrmObs", "Observações", dx, 362),
        Controle("txtObs", "text", {"Default": f'If({novo}, "", gItem.Observacoes)', "Mode": "TextMode.MultiLine",
                                    **pos(dx, 388, 626, 116)}),
    ]

    salvar = (
        "If(\n"
        "    IsBlank(Trim(txtCodigo.Text)) || IsBlank(cmbColaborador.Selected) || IsBlank(dteInicio.SelectedDate) || Value(txtDuracao.Text) < 1,\n"
        '    Notify("Preencha código, colaborador, data de início e duração (mínimo 1 dia).", NotificationType.Error),\n'
        f"    Set(gSalvo, Patch({lista}, If({novo}, Defaults({lista}), gItem), {{\n"
        "        Title: Trim(txtCodigo.Text),\n"
        "        Descricao: txtDescricao.Text,\n"
        "        Colaborador: cmbColaborador.Selected,\n"
        "        Gestor: cmbGestor.Selected,\n"
        "        Disciplina: drpDisciplina.Selected,\n"
        "        DataInicio: dteInicio.SelectedDate,\n"
        "        DuracaoDias: Value(txtDuracao.Text),\n"
        "        StatusProjeto: drpStatus.Selected,\n"
        '        PercConcluido: If(drpStatus.Selected.Value = "Concluído", 100, sldPerc.Value),\n'
        "        Observacoes: txtObs.Text\n"
        "    }));\n"
        f"    If(IsEmpty(Errors({lista})),\n"
        '        Notify("Projeto salvo. Término previsto: " & Text(gSalvo.DataFim, "dd/mm/yyyy"), NotificationType.Success);\n'
        "        Set(gItem, Blank());\n"
        "        Navigate(scrProjetos, ScreenTransition.None),\n"
        f'        Notify("Não foi possível salvar: " & First(Errors({lista})).Message, NotificationType.Error)\n'
        "    )\n"
        ")"
    )
    excluir = (
        f"Remove({lista}, gItem);\n"
        f"If(IsEmpty(Errors({lista})),\n"
        '    Notify("Projeto excluído.", NotificationType.Success);\n'
        "    Set(gItem, Blank());\n"
        "    Navigate(scrProjetos, ScreenTransition.None),\n"
        f'    Notify("Não foi possível excluir: " & First(Errors({lista})).Message, NotificationType.Error)\n'
        ")"
    )
    acoes = [
        retangulo("recFrmRodape", 0, 680, "Parent.Width", 88, BRANCO),
        botao("btnFrmSalvar", "Salvar", ex, 704, 160, salvar, cor=VERDE),
        botao("btnFrmCancelar", "Cancelar", ex + 176, 704, 160, "Set(gItem, Blank());\nNavigate(scrProjetos, ScreenTransition.None)",
              cor=CINZA_ESCURO),
        botao("btnFrmExcluir", "Excluir", 1000, 704, 140, "UpdateContext({locConfirmar: true})", cor=VERMELHO,
              Visible=f"!{novo} && !locConfirmar"),
        botao("btnFrmConfirmar", "Confirmar exclusão", 1000, 704, 200, excluir, cor=VERMELHO,
              Visible=f"!{novo} && locConfirmar"),
        botao("btnFrmDesistir", "Não excluir", 1210, 704, 120, "UpdateContext({locConfirmar: false})",
              cor=CINZA_ESCURO, Visible=f"!{novo} && locConfirmar"),
    ]
    reset = ";\n".join(f"Reset({n})" for n in ("txtCodigo", "txtDescricao", "cmbColaborador", "cmbGestor",
                                                "drpDisciplina", "dteInicio", "txtDuracao", "drpStatus",
                                                "sldPerc", "txtObs"))
    filhos = cabecalho("Frm", "Projeto", []) + [
        rotulo("lblFrmModo", f'If({novo}, "Novo projeto", "Editando " & gItem.Title)', 700, 0, 640, 64,
               tamanho=14, cor=BRANCO, Align="Align.Right"),
        *entradas, *acoes,
    ]
    return Controle("scrFormulario", "screen", {
        "Fill": BRANCO, "OnVisible": reset + ";\nUpdateContext({locConfirmar: false})"}, filhos)


def tela_dashboard(lista: str) -> Controle:
    atualizar = (
        f"Refresh({lista});\n"
        f"ClearCollect(colProj, {lista});\n"
        "ClearCollect(colPessoas,\n"
        "    Sort(\n"
        "        Filter(\n"
        "            ForAll(Sequence(CountRows(colProj)) As i,\n"
        "                With({p: Last(FirstN(colProj, i.Value))},\n"
        "                    {Email: p.Colaborador.Email, Nome: p.Colaborador.DisplayName,\n"
        "                     Primeiro: IsBlank(LookUp(FirstN(colProj, i.Value - 1), Colaborador.Email = p.Colaborador.Email))})),\n"
        "            Primeiro),\n"
        "        Nome));\n"
        "UpdateContext({locPessoa: Blank()})"
    )

    kpis = [
        ("Total", "CountRows(colProj)", AZUL_ESCURO),
        ("Em andamento", f'CountIf(colProj, StatusProjeto.Value = "Em andamento" && !{atrasado("")})', AZUL),
        ("Atrasados", f"CountIf(colProj, {atrasado('')})", VERMELHO),
        ("Vencem em 7 dias", f"CountIf(colProj, {aberto('')} && DataFim >= Today() && DateDiff(Today(), DataFim, Days) <= 7)",
         AMBAR_TEXTO),
        ("Concluídos", 'CountIf(colProj, StatusProjeto.Value = "Concluído")', VERDE),
        ("% médio concluído",
         'Round(Coalesce(Average(Filter(colProj, StatusProjeto.Value <> "Cancelado"), PercConcluido), 0), 0) & "%"',
         CINZA_ESCURO),
    ]
    cards = []
    for i, (titulo, valor, cor) in enumerate(kpis):
        x = 24 + i * 222
        cards += [
            retangulo(f"recKpi{i}", x, 80, 210, 84, BRANCO),
            retangulo(f"recKpiCor{i}", x, 80, 6, 84, cor),
            rotulo(f"lblKpiTit{i}", txt(titulo), x + 16, 86, 190, 24, tamanho=11, cor=CINZA_ESCURO),
            rotulo(f"lblKpiVal{i}", valor, x + 16, 110, 190, 48, tamanho=26, cor=cor, negrito=True),
        ]

    p = "ThisItem.Email"
    def conta(cond):
        return f"CountIf(colProj, Colaborador.Email = {p} && {cond})"
    n_atr = conta(atrasado(""))
    n_ini = conta(f'StatusProjeto.Value = "Não iniciado" && !{atrasado("")}')
    n_and = conta(f'StatusProjeto.Value = "Em andamento" && !{atrasado("")}')
    n_pau = conta(f'StatusProjeto.Value = "Pausado" && !{atrasado("")}')
    n_con = conta('StatusProjeto.Value = "Concluído"')
    total = f"CountIf(colProj, Colaborador.Email = {p})"

    colunas = [("Colaborador", 16, 220), ("Total", 240, 60), ("Não inic.", 300, 70), ("Andamento", 370, 80),
               ("Pausados", 450, 70), ("Atrasados", 520, 70), ("Concluídos", 590, 80)]
    cab_resumo = [retangulo("recResCab", 24, 180, 840, 36, AZUL_ESCURO)] + [
        rotulo(f"lblResCab{i}", txt(t), 24 + x, 180, w, 36, tamanho=10, cor=BRANCO, negrito=True,
               Align="Align.Left" if i == 0 else "Align.Center")
        for i, (t, x, w) in enumerate(colunas)
    ] + [rotulo("lblResCabBarra", txt("Distribuição"), 24 + 680, 180, 150, 36, tamanho=10, cor=BRANCO, negrito=True)]

    valores = [f"{p}", total, n_ini, n_and, n_pau, n_atr, n_con]
    sel = "Select(Parent)"
    modelo = [
        retangulo("recResSel", 0, 0, "Parent.TemplateWidth", "Parent.TemplateHeight",
                  f"If(locPessoa = {p}, RGBA(222, 236, 249, 1), RGBA(0, 0, 0, 0))", OnSelect=sel),
        rotulo("lblResNome", "ThisItem.Nome", 16, 0, 220, 40, tamanho=11, negrito=True, OnSelect=sel),
    ]
    for i, (expr, (_, x, w)) in enumerate(zip(valores[1:], colunas[1:]), start=1):
        cor = VERMELHO if i == 5 else TEXTO
        modelo.append(rotulo(f"lblResVal{i}", expr, x, 0, w, 40, tamanho=11, cor=cor, Align="Align.Center",
                             OnSelect=sel))
    # barra empilhada 100% (não iniciados | andamento | pausados | atrasados | concluídos)
    larg = 140
    seg = [("Ini", n_ini, CINZA), ("And", n_and, AZUL), ("Pau", n_pau, AMBAR), ("Atr", n_atr, VERMELHO),
           ("Con", n_con, VERDE)]
    x_atual = "680"
    for nome, expr, cor in seg:
        w = f"{larg} * {expr} / Max(1, {total})"
        modelo.append(retangulo(f"recBar{nome}", x_atual, 12, w, 16, cor, OnSelect=sel))
        x_atual = f"recBar{nome}.X + recBar{nome}.Width"
    modelo.append(retangulo("recResLinha", 0, "Parent.TemplateHeight - 1", "Parent.TemplateWidth", 1, LINHA))

    galeria_resumo = Controle("galResumo", "gallery.galleryVertical", {
        "Items": "colPessoas", **pos(24, 216, 840, 532), "TemplateSize": "40", "TemplatePadding": "0",
        "Fill": BRANCO, "ShowScrollbar": "true",
        "OnSelect": "UpdateContext({locPessoa: If(locPessoa = ThisItem.Email, Blank(), ThisItem.Email)})",
    }, modelo)

    itens_dir = (
        "Sort(\n"
        "    Filter(colProj,\n"
        f"        If(IsBlank(locPessoa), {atrasado('')}, Colaborador.Email = locPessoa)),\n"
        "    DataFim, SortOrder.Ascending)"
    )
    modelo_dir = [
        retangulo("recDetCor", 0, 6, 5, "Parent.TemplateHeight - 12", cor_status("ThisItem.StatusProjeto.Value")),
        rotulo("lblDetCodigo", 'ThisItem.Title & "  -  " & ThisItem.Descricao', 14, 2, 410, 26, tamanho=11,
               negrito=True),
        rotulo("lblDetInfo",
               'If(IsBlank(locPessoa), ThisItem.Colaborador.DisplayName & "  •  ", "") & "Fim " & '
               'Text(ThisItem.DataFim, "dd/mm/yyyy") & "  •  " & ThisItem.StatusProjeto.Value & "  •  " & '
               'Coalesce(ThisItem.PercConcluido, 0) & "%"',
               14, 28, 300, 24, tamanho=10, cor=CINZA_ESCURO),
        rotulo("lblDetSit", situacao("ThisItem."), 314, 28, 110, 24, tamanho=10, cor=cor_situacao("ThisItem."),
               negrito=True, Align="Align.Right"),
        retangulo("recDetLinha", 0, "Parent.TemplateHeight - 1", "Parent.TemplateWidth", 1, LINHA),
    ]
    painel = [
        retangulo("recDetFundo", 880, 180, 462, 568, BRANCO),
        rotulo("lblDetTitulo",
               'If(IsBlank(locPessoa), "Projetos atrasados (todos)", "Projetos de " & '
               'LookUp(colPessoas, Email = locPessoa).Nome)',
               896, 184, 430, 32, tamanho=13, cor=AZUL_ESCURO, negrito=True),
        rotulo("lblDetDica", txt("Clique em um colaborador ao lado para ver os projetos dele."), 896, 214, 430, 22,
               tamanho=9, cor=CINZA_ESCURO),
        Controle("galDetalhe", "gallery.galleryVertical", {
            "Items": itens_dir, **pos(888, 240, 446, 500), "TemplateSize": "56", "TemplatePadding": "0",
            "ShowScrollbar": "true",
            "OnSelect": f"Set(gItem, LookUp({lista}, ID = ThisItem.ID));\nNavigate(scrFormulario, ScreenTransition.None)",
        }, modelo_dir),
        rotulo("lblDetVazio", txt("Nenhum projeto."), 896, 250, 430, 30, tamanho=11, cor=CINZA_ESCURO,
               Visible="CountRows(galDetalhe.AllItems) = 0"),
    ]
    filhos = cabecalho("Dash", "Dashboard por colaborador", [
        botao("btnDashAtualizar", "Atualizar", 930, 12, 180, atualizar),
        botao("btnDashProjetos", "Projetos", 1124, 12, 180, "Navigate(scrProjetos, ScreenTransition.None)", cor=VERDE),
    ]) + cards + cab_resumo + [galeria_resumo] + painel
    return Controle("scrDashboard", "screen", {"Fill": CINZA_CLARO, "OnVisible": "Select(btnDashAtualizar)"},
                    filhos)


# --------------------------------------------------------------------------- montagem
def ler_definicao(caminho: Path) -> tuple[str, bool]:
    d = json.loads(caminho.read_text(encoding="utf-8"))
    xml_fim = next(c["xml"] for c in d["campos"] if c["nome"] == "DataFim")
    dias_uteis = "WEEKDAY" in xml_fim.upper()
    nomes = {c["nome"] for c in d["campos"]}
    faltando = {"Descricao", "Colaborador", "Gestor", "Disciplina", "DataInicio", "DuracaoDias", "DataFim",
                "StatusProjeto", "PercConcluido", "Observacoes"} - nomes
    if faltando:
        raise SystemExit(f"definicao-lista.json não tem as colunas esperadas: {sorted(faltando)}")
    return d["titulo"], dias_uteis


def nome_fonte(lista: str) -> str:
    """Nome da fonte de dados como o Power Apps a referencia em fórmulas."""
    return lista if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", lista) else "'" + lista.replace("'", "''") + "'"


def validar_propriedades(telas: list[Controle], pasta_pkgs: Path) -> list[str]:
    """Confere se toda propriedade usada existe na definição do controle (pkgs/*.xml)."""
    conhecidas: dict[str, set[str]] = {}
    for xml in pasta_pkgs.glob("*.xml"):
        modelo = xml.stem.rsplit("_", 1)[0]
        conhecidas[modelo] = set(re.findall(r'name="([A-Za-z]+)"', xml.read_text(encoding="utf-8")))
    comuns = {"ZIndex"}
    tela = {"Fill", "OnVisible", "OnHidden", "BackgroundImage", "LoadingSpinner", "LoadingSpinnerColor"}
    app = {"OnStart", "BackEnabled", "ConfirmExit", "ConfirmExitMessage", "Theme"}
    erros = []
    for t in telas:
        for c in t.todos():
            base = c.modelo.split(".")[0]
            validas = {"screen": tela, "appinfo": app}.get(base) or conhecidas.get(base)
            if validas is None:
                erros.append(f"{c.nome}: modelo '{base}' sem definição em pkgs/")
                continue
            for prop in c.props:
                if prop not in validas and prop not in comuns:
                    erros.append(f"{c.nome} ({base}): propriedade desconhecida '{prop}'")
    nomes = [c.nome for t in telas for c in t.todos()]
    erros += [f"nome de controle repetido: {n}" for n in sorted({n for n in nomes if nomes.count(n) > 1})]
    return erros


def gerar_fontes(destino: Path, nome_app: str, lista: str, dias_uteis: bool) -> list[Controle]:
    if destino.exists():
        shutil.rmtree(destino)
    (destino / "Src").mkdir(parents=True)
    shutil.copytree(BASE / "pkgs", destino / "pkgs")
    shutil.copy(BASE / "ControlTemplates.json", destino)
    shutil.copy(BASE / "Themes.json", destino / "Src")
    (destino / "ComponentReferences.json").write_text("[]", encoding="utf-8")

    fonte = nome_fonte(lista)
    app = app_onstart()
    telas = [tela_projetos(fonte), tela_formulario(fonte, dias_uteis), tela_dashboard(fonte)]

    erros = validar_propriedades([app, *telas], BASE / "pkgs")
    if erros:
        raise SystemExit("Erros na definição do app:\n  " + "\n  ".join(erros))

    for c in [app, *telas]:
        conteudo = "\n".join(controle_yaml(c)) + "\n"
        (destino / "Src" / f"{c.nome}.fx.yaml").write_text(conteudo, encoding="utf-8")

    m = json.loads((BASE / "CanvasManifest.json").read_text(encoding="utf-8"))
    m["Properties"].update({
        "AppCreationSource": "AppFromScratch", "AppDescription": "Gerenciador de projetos multidisciplinar",
        "Author": "", "DefaultConnectedDataSourceMaxGetRowsCount": 2000,
        "DocumentAppType": "DesktopOrTablet", "DocumentLayoutWidth": LARGURA, "DocumentLayoutHeight": ALTURA,
        "DocumentLayoutOrientation": "landscape", "FileID": str(uuid.uuid4()), "Id": str(uuid.uuid4()),
        "Name": f"{nome_app}.msapp",
    })
    m["PublishInfo"].update({"AppName": nome_app, "BackgroundColor": "RGBA(31, 78, 121, 1)",
                             "IconColor": "RGBA(255, 255, 255, 1)", "IconName": "Edit", "LogoFileName": ""})
    m["ScreenOrder"] = [t.nome for t in telas]
    (destino / "CanvasManifest.json").write_text(json.dumps(m, indent=2, ensure_ascii=False), encoding="utf-8")
    return [app, *telas]


def empacotar(fontes: Path, saida: Path) -> None:
    pac = shutil.which("pac") or shutil.which("pac.exe") or shutil.which("pac.cmd")
    if not pac:
        raise SystemExit("A Power Platform CLI (pac) não foi encontrada. Instale com:\n"
                         "  dotnet tool install --global Microsoft.PowerApps.CLI.Tool --version 2.15.1\n"
                         f"ou empacote depois:  pac canvas pack --sources \"{fontes}\" --msapp \"{saida}\" "
                         "--layout Experimental")
    cmd = [pac, "canvas", "pack", "--sources", str(fontes), "--msapp", str(saida), "--layout", "Experimental"]
    r = subprocess.run(cmd, capture_output=True, text=True)
    if r.returncode != 0 or not saida.exists():
        raise SystemExit(f"Falha ao empacotar:\n{r.stdout}\n{r.stderr}")


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Gera o Power App do gerenciador de projetos.")
    ap.add_argument("--definicao", type=Path, default=DEFINICAO, help="definicao-lista.json")
    ap.add_argument("--nome", default="Gerenciador de Projetos", help="nome do app")
    ap.add_argument("--saida", type=Path, default=AQUI / "GerenciadorProjetos.msapp", help="arquivo .msapp")
    ap.add_argument("--fontes", type=Path, default=AQUI / "fontes", help="pasta onde as fontes são geradas")
    ap.add_argument("--somente-fontes", action="store_true", help="não chama o pac")
    a = ap.parse_args(argv)

    lista, dias_uteis = ler_definicao(a.definicao)
    telas = gerar_fontes(a.fontes, a.nome, lista, dias_uteis)
    n = sum(1 for t in telas for _ in t.todos())
    print(f"Fontes geradas em {a.fontes} ({len(telas) - 1} telas, {n} controles; lista '{lista}', "
          f"data fim em {'dias úteis' if dias_uteis else 'dias corridos'}).")
    if a.somente_fontes:
        return 0
    if a.saida.exists():
        a.saida.unlink()
    empacotar(a.fontes, a.saida)
    print(f"App gerado: {a.saida}")
    print("Próximo passo: Power Apps Studio > Abrir > Procurar arquivos > selecione o .msapp; "
          f"depois Dados > Adicionar dados > SharePoint > seu site > lista '{lista}'.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
