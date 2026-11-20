# JSON Compartilhado na Rede

App HTML + JS + CSS (sem servidor) que lê e grava um JSON numa pasta da rede,
com controle de concorrência feito **inteiramente pelos clientes**.

## Como usar

1. Abra `index.html` no **Chrome** ou **Edge** (desktop). Firefox/Safari não têm a File System Access API.
2. Informe seu nome.
3. Clique em **Escolher pasta da rede** e selecione a pasta compartilhada
   (unidade mapeada, ex.: `Z:\Projetos\dados`). O navegador lembra a pasta; nas próximas
   vezes basta clicar em **Reconectar**.
4. **Editar** → altera → **Salvar** (ou **Cancelar**).

> Dica: se o navegador bloquear em `file://`, sirva a pasta localmente
> (`python -m http.server`) e abra `http://localhost:8000`.

## Arquivos criados na pasta da rede

| Arquivo | Função |
|---|---|
| `dados.json` | Os dados + `meta.versao`, `atualizadoPor`, `atualizadoEm` |
| `dados.lock.json` | Quem está usando o arquivo, fazendo o quê (`editando`, `gravando`, `criando`, `restaurando`) e desde quando |
| `dados.backup.json` | Versão anterior, gravada antes de cada salvamento |

## Protocolo de bloqueio (todos os clientes seguem)

1. **Ler o lock.** Se há lock vivo de outra pessoa → não entra, mostra quem está usando.
2. **Gravar o próprio lock** com token aleatório, **esperar um tempo aleatório e reler** (duas vezes).
   Se duas pessoas tentaram ao mesmo tempo, só uma vê o próprio token; a outra desiste.
3. **Heartbeat** a cada 5 s enquanto edita.
4. **Lock abandonado:** se o conteúdo do lock não muda por 30 s (medido pelo relógio de quem observa,
   então não depende de relógios sincronizados), qualquer cliente pode assumir. Fechar o navegador
   ou travar a máquina libera o arquivo sozinho.
5. **Ao salvar:** confirma que o lock é seu (ou recupera, se estiver livre), confere que `meta.versao`
   não mudou, grava backup, grava `versao + 1`, relê para validar e apaga o lock.
6. **JSON ilegível:** tenta ler de novo algumas vezes; se persistir, oferece **Restaurar backup**.
7. **Forçar liberação:** botão manual para casos extremos.

A gravação usa `createWritable()`, que escreve num arquivo temporário e só substitui o original
ao fechar — quem lê nunca vê o arquivo pela metade.

## Limitações

Sem servidor não há lock 100% atômico: o protocolo reduz a chance de colisão a quase zero e o
controle de versão impede sobrescrever o trabalho de outra pessoa, mas todos precisam usar este app
(alguém editando o JSON direto no Bloco de Notas ignora o lock). Ajuste tempos em `CFG` no `app.js`.
