# Lista "Projetos" para SharePoint Online

Versão SharePoint do gerenciador de projetos multidisciplinar. Um script cria a lista pronta no site da
sua empresa: colunas, data fim calculada, cores, modos de exibição e formulário em seções.

## O que é criado

**Colunas**

| Coluna | Tipo | Observação |
|---|---|---|
| Código do projeto | Texto (coluna Título) | Obrigatória |
| Descrição | Várias linhas | |
| Disciplina | Escolha | Aceita valores digitados além da lista |
| Colaborador | Pessoa | Obrigatória |
| Gestor | Pessoa | Chefe imediato, usado em "Minha equipe" |
| Data início | Data | Padrão = hoje |
| Duração (dias) | Número inteiro ≥ 1 | Duração estimada |
| **Data fim** | **Calculada** | Dias úteis (seg-sex); se começar no fim de semana, conta a partir de segunda |
| Status | Escolha | Etiqueta colorida |
| % concluído | Número 0-100 | Barra de progresso |
| Observações | Várias linhas | |

**Modos de exibição** (todos com linhas **vermelhas para atrasados** e **amarelas para os que vencem em até 7 dias**)

- **Todos os itens**: ordenado pela data fim.
- **Meus projetos**: só os projetos em que você é o Colaborador.
- **Minha equipe**: projetos em que você é o Gestor, agrupados por colaborador.
- **Por colaborador**: todos, agrupados por pessoa, com contagem.
- **Atrasados**: data fim já passou e o projeto não está concluído nem cancelado.

**Formulário**: cabeçalho "Projeto X" e seções Identificação / Responsáveis / Prazo / Andamento.

**Histórico de versões** ativado: cada alteração fica registrada com autor e data.

## Instalação (recomendado: pelo navegador, sem instalar nada)

Requisito: ser **Proprietário ou Membro** do site do SharePoint.

1. Abra no **Edge ou Chrome** o site onde a lista vai ficar, ex.: `https://suaempresa.sharepoint.com/sites/Projetos`.
2. Pressione **F12** e clique na guia **Console**.
3. Abra o arquivo [`criar-lista-projetos.js`](criar-lista-projetos.js), copie **todo** o conteúdo, cole no Console e
   pressione **Enter**. Se o navegador bloquear a colagem, digite `allow pasting` (ou `permitir colar`),
   pressione Enter e cole de novo.
4. Aguarde cerca de 30 segundos até aparecer **CONCLUÍDO!** com o link da lista.

> Quer **dias corridos**, outro nome de lista ou outras disciplinas? Altere o bloco `CONFIG` no início
> do script antes de colar.

## Alternativa para a TI (PnP PowerShell)

Se a empresa bloqueia o Console do navegador:

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser
.\Criar-ListaProjetos.ps1 -SiteUrl https://suaempresa.sharepoint.com/sites/Projetos -ClientId <id-do-app-PnP>
```

O script lê [`definicao-lista.json`](definicao-lista.json), que tem as mesmas definições do script do navegador.

## Alternativa em Python

Faz o mesmo que o PowerShell, em Windows, Linux ou Mac:

```bash
pip install -r requirements.txt
python criar_lista_projetos.py --site https://suaempresa.sharepoint.com/sites/Projetos --client-id <id-do-app>
```

- O login abre o navegador (conta Microsoft 365 da empresa). Em servidor sem navegador, use `--device-code`.
- `--tenant suaempresa.onmicrosoft.com`: necessário se o app for de um único tenant.
- `--token <token>`: usa um token de acesso já obtido por outro meio.

O App Registration no Entra ID precisa de:

1. plataforma **Aplicativos móveis e da área de trabalho** com o redirect URI `http://localhost`;
2. **Permitir fluxos de clientes públicos = Sim**;
3. permissão delegada **SharePoint › AllSites.Manage**, com consentimento do administrador.

## Alternativa manual

Os arquivos em [`formatacao/`](formatacao) podem ser colados direto na interface:

| Arquivo | Onde colar |
|---|---|
| `coluna-status.json` | Cabeçalho da coluna Status › Configurações da coluna › Formatar esta coluna › Modo avançado |
| `coluna-percentual.json` | Idem, na coluna % concluído |
| `coluna-data-fim.json` | Idem, na coluna Data fim |
| `modo-exibicao-linhas.json` | Exibir opções › Formatar modo de exibição atual › Modo avançado |
| `formulario-cabecalho.json` / `formulario-corpo.json` | Abrir um item › Editar formulário › Configurar layout |
| `formula-data-fim.txt` | Fórmula da coluna calculada |

> Os JSONs de formatação usam os **nomes internos** `DataInicio`, `DuracaoDias` e `StatusProjeto`. Na montagem
> manual, crie as colunas primeiro com esses nomes (sem espaço nem acento) e só depois renomeie para o nome de exibição.

## Depois de criar (2 minutos, opcional)

1. **Quadro (Kanban)**: Exibir opções › Criar novo modo de exibição › **Quadro**, organizado por **Status**.
2. **Calendário**: Criar novo modo de exibição › **Calendário**, início = Data início, fim = Data fim.
3. **Dashboard**: crie uma página no site e adicione a web part **Lista** (modo "Por colaborador") e a
   web part **Gráfico Rápido** (contagem por Status ou por Colaborador). Para indicadores avançados, conecte o
   **Power BI** à lista.
4. **Tornar "Meus projetos" o padrão**: abra o modo de exibição › Exibir opções › Definir como padrão.

## Limitações e próximos passos

- **Feriados** não entram no cálculo (colunas calculadas não consultam outras listas). Para considerá-los, use
  um fluxo do Power Automate que calcula a data fim e grava em uma coluna de data comum.
- **Permissões por hierarquia** (chefe edita só os projetos da equipe): por padrão, quem tem acesso ao site
  edita a lista. Para restringir por item, use um fluxo do Power Automate que, ao criar o item, dá acesso
  de edição apenas ao Colaborador e ao Gestor.
- O cálculo de "atrasado" nas cores é feito a partir de Data início + Duração (mesma regra da coluna Data fim),
  porque a formatação JSON recebe colunas calculadas como texto.
