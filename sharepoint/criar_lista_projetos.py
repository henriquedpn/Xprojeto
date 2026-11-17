#!/usr/bin/env python3
"""
Cria a lista "Projetos" no SharePoint Online (versão Python do Criar-ListaProjetos.ps1).

Lê as definições de ``definicao-lista.json`` (as mesmas usadas pelo script do
navegador e pelo PowerShell) e cria, via API REST do SharePoint:
colunas, data fim calculada, formatação JSON, modos de exibição e o layout
do formulário.

Requisitos
----------
    pip install -r requirements.txt        # msal + requests

Um App Registration no Entra ID (o mesmo usado pelo PnP serve):
  * Plataforma "Aplicativos móveis e da área de trabalho" com o redirect
    URI ``http://localhost``
  * "Permitir fluxos de clientes públicos" = Sim
  * Permissão delegada SharePoint > ``AllSites.Manage`` (com consentimento)
Quem executa precisa ser Proprietário ou Membro do site.

Exemplos
--------
    python criar_lista_projetos.py --site https://empresa.sharepoint.com/sites/Projetos --client-id <GUID>
    python criar_lista_projetos.py --site ... --client-id <GUID> --device-code   # sem navegador local
    python criar_lista_projetos.py --site ... --token <access_token>             # token já obtido
"""
from __future__ import annotations

import argparse
import json
import sys
import time
from pathlib import Path
from urllib.parse import quote, urlparse

import requests

ARQUIVO_DEFINICAO = Path(__file__).resolve().parent / "definicao-lista.json"


# --------------------------------------------------------------------------- autenticação
def obter_token(site: str, client_id: str, tenant: str, device_code: bool) -> str:
    import msal  # importado aqui para permitir --token sem ter o msal instalado

    host = urlparse(site).netloc
    scopes = [f"https://{host}/.default"]
    app = msal.PublicClientApplication(client_id, authority=f"https://login.microsoftonline.com/{tenant}")

    if device_code:
        fluxo = app.initiate_device_flow(scopes=scopes)
        if "user_code" not in fluxo:
            raise RuntimeError(f"Falha ao iniciar o login: {fluxo.get('error_description', fluxo)}")
        print(fluxo["message"], flush=True)
        resultado = app.acquire_token_by_device_flow(fluxo)
    else:
        print("Abrindo o navegador para login no Microsoft 365...", flush=True)
        resultado = app.acquire_token_interactive(scopes=scopes, prompt="select_account")

    if "access_token" not in resultado:
        raise RuntimeError(f"Login não concluído: {resultado.get('error_description', resultado)}")
    return resultado["access_token"]


# --------------------------------------------------------------------------- cliente REST
class ErroSharePoint(Exception):
    def __init__(self, status: int, mensagem: str):
        super().__init__(mensagem)
        self.status = status


class ClienteSharePoint:
    def __init__(self, site: str, token: str, sessao: requests.Session | None = None):
        self.site = site.rstrip("/")
        self.sessao = sessao or requests.Session()
        self.sessao.headers.update({
            "Authorization": f"Bearer {token}",
            "Accept": "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
        })

    def _chamar(self, metodo: str, caminho: str, corpo=None, extras: dict | None = None):
        url = self.site + caminho
        dados = None if corpo is None else json.dumps(corpo, ensure_ascii=False).encode("utf-8")
        for tentativa in range(5):
            r = self.sessao.request(metodo, url, data=dados, headers=extras or {}, timeout=60)
            if r.status_code in (429, 503) and tentativa < 4:   # limitação do SharePoint
                espera = int(r.headers.get("Retry-After", 5))
                print(f"   SharePoint pediu para aguardar {espera}s...", flush=True)
                time.sleep(espera)
                continue
            break
        if not r.ok:
            try:
                msg = r.json()["error"]["message"]["value"]
            except (ValueError, KeyError, TypeError):
                msg = r.text
            raise ErroSharePoint(r.status_code, f"{metodo} {caminho} -> {r.status_code}: {msg}")
        return r.json() if r.content else None

    def get(self, caminho):
        return self._chamar("GET", caminho)

    def post(self, caminho, corpo=None):
        return self._chamar("POST", caminho, corpo)

    def merge(self, caminho, corpo):
        return self._chamar("POST", caminho, corpo, {"X-HTTP-Method": "MERGE", "IF-MATCH": "*"})


def q(texto: str) -> str:
    """Escapa um valor para uso dentro de getbytitle('...')."""
    return quote(texto.replace("'", "''"), safe="")


# --------------------------------------------------------------------------- criação
def passo(msg: str) -> None:
    print(f"► {msg}", flush=True)


def aviso(msg: str) -> None:
    print(f"   AVISO: {msg}", file=sys.stderr, flush=True)


def configurar_view(sp: ClienteSharePoint, base: str, campos: list[str], extras: dict) -> None:
    sp.post(base + "/viewfields/removeallviewfields")
    for f in campos:
        sp.post(base + f"/viewfields/addviewfield('{f}')")
    sp.merge(base, {"__metadata": {"type": "SP.View"}, **extras})


def criar_lista(sp: ClienteSharePoint, d: dict) -> str:
    lista = f"/_api/web/lists/getbytitle('{q(d['titulo'])}')"

    try:
        sp.get(lista + "?$select=Id")
        raise SystemExit(f"Já existe uma lista '{d['titulo']}' neste site. "
                         "Exclua-a ou altere 'titulo' em definicao-lista.json.")
    except ErroSharePoint as e:
        if e.status != 404:
            raise

    passo(f"Criando a lista {d['titulo']}")
    sp.post("/_api/web/lists", {"__metadata": {"type": "SP.List"}, "BaseTemplate": 100,
                                "Title": d["titulo"], "Description": d["descricao"]})
    sp.merge(lista, {"__metadata": {"type": "SP.List"}, "EnableVersioning": True})

    passo("Criando as colunas")
    for c in d["campos"]:
        # Options 12 = AddToAllContentTypes + AddFieldInternalNameHint (nome interno sem acento)
        sp.post(lista + "/fields/createfieldasxml", {"parameters": {
            "__metadata": {"type": "SP.XmlSchemaFieldCreationInformation"},
            "SchemaXml": c["xml"], "Options": 12}})
        print(f"   coluna {c['titulo']}", flush=True)

    passo("Renomeando colunas e aplicando formatação")
    sp.merge(lista + "/fields/getbyinternalnameortitle('Title')",
             {"__metadata": {"type": "SP.Field"}, "Title": d["tituloCodigo"]})
    for c in d["campos"]:
        corpo = {"__metadata": {"type": "SP.Field"}, "Title": c["titulo"]}
        if c.get("formatador"):
            corpo["CustomFormatter"] = c["formatador"]
        try:
            sp.merge(lista + f"/fields/getbyinternalnameortitle('{c['nome']}')", corpo)
        except ErroSharePoint as e:
            aviso(f"coluna {c['titulo']}: {e}")

    passo("Configurando o modo de exibição padrão")
    configurar_view(sp, lista + "/defaultview", d["camposView"], {
        "ViewQuery": d["consultaPadrao"], "CustomFormatter": d["formatadorLinha"], "RowLimit": 100})

    for v in d["views"]:
        passo(f"Criando o modo de exibição \"{v['titulo']}\"")
        try:
            sp.post(lista + "/views", {"__metadata": {"type": "SP.View"}, "Title": v["titulo"],
                                       "PersonalView": False, "ViewQuery": v["query"],
                                       "RowLimit": 100, "Paged": True})
            extras = {"CustomFormatter": d["formatadorLinha"]}
            if v.get("totais"):
                extras.update(Aggregations=d["totais"], AggregationsStatus="On")
            configurar_view(sp, lista + f"/views/getbytitle('{q(v['titulo'])}')", d["camposView"], extras)
        except ErroSharePoint as e:
            aviso(f"modo de exibição \"{v['titulo']}\": {e}")

    passo("Organizando o formulário em seções")
    try:
        cts = sp.get(lista + "/contenttypes?$select=StringId,Name")["d"]["results"]
        ct = next((c for c in cts if c["StringId"].upper().startswith("0X0100")
                   or c["StringId"].upper() == "0X01"), None)
        if ct:
            sp.merge(lista + f"/contenttypes('{ct['StringId']}')",
                     {"__metadata": {"type": "SP.ContentType"}, "ClientFormCustomFormatter": d["formulario"]})
    except ErroSharePoint as e:
        aviso(f"layout do formulário não aplicado (opcional): {e}")

    return f"{sp.site}/Lists/{quote(d['titulo'])}"


# --------------------------------------------------------------------------- main
def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description="Cria a lista Projetos no SharePoint Online.")
    p.add_argument("--site", required=True, help="URL do site, ex.: https://empresa.sharepoint.com/sites/Projetos")
    p.add_argument("--client-id", help="Application (client) ID do App Registration no Entra ID")
    p.add_argument("--tenant", default="organizations", help="ID ou domínio do tenant (padrão: organizations)")
    p.add_argument("--device-code", action="store_true", help="login por código (sem abrir navegador local)")
    p.add_argument("--token", help="usar um access token já obtido (dispensa --client-id)")
    p.add_argument("--definicao", type=Path, default=ARQUIVO_DEFINICAO, help="arquivo de definição da lista")
    a = p.parse_args(argv)

    if not a.token and not a.client_id:
        p.error("informe --client-id (ou --token)")
    if not a.site.lower().startswith("https://"):
        p.error("--site deve começar com https://")

    definicao = json.loads(a.definicao.read_text(encoding="utf-8"))
    token = a.token or obter_token(a.site, a.client_id, a.tenant, a.device_code)
    sp = ClienteSharePoint(a.site, token)

    try:
        url = criar_lista(sp, definicao)
    except ErroSharePoint as e:
        print(f"ERRO: {e}\nSe a lista chegou a ser criada, exclua-a antes de rodar novamente.", file=sys.stderr)
        return 1
    print(f"CONCLUÍDO! {url}", flush=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
