# Gerenciador de Projetos Multidisciplinar (Excel + VBA)

Gerenciador de projetos em Excel/VBA para **vários usuários ao mesmo tempo**. Cada colaborador
usa uma **cópia local** do arquivo `.xlsm` (o front-end). Os dados ficam em **um único arquivo
`.xlsx` compartilhado na rede** (o banco).

```
 PC do João  ─┐                                   ┌─ aba Projetos (cadastro + lista)
 PC da Maria ─┼──►  \\servidor\projetos\BD_Projetos.xlsx  ──►  ├─ aba Dashboard (por colaborador)
 PC do chefe ─┘     (+ BD_Projetos.xlsx.lock)       └─ atualização automática
```

## Funcionalidades

| Item | Como funciona |
|---|---|
| **Aba Projetos** | Formulário (linha amarela) com **Código do projeto, Descrição, Colaborador, Disciplina, Data início, Duração estimada (dias), Data fim (calculada)**, Status e % concluído. Abaixo, a tabela `tblProjetos` com todos os projetos que o usuário pode ver. |
| **Data fim** | Calculada pela duração estimada: em **dias úteis** (descontando fins de semana e a lista de **Feriados** da aba Config) ou em dias corridos (`Config > Data fim em dias úteis = NAO`). |
| **Chefe cadastra para a equipe** | A aba Usuarios tem a coluna **Gestor**. O colaborador cadastra os próprios projetos; o chefe (direto ou indireto) cadastra, edita e exclui projetos de qualquer pessoa da equipe; o Admin gerencia todos. |
| **Atualização automática** | A cada N segundos (padrão 30) o Excel de cada usuário verifica se o banco mudou. Se mudou, recarrega a lista e reconstrói o **Dashboard** sem nenhuma ação do usuário. Assim, quando o chefe inclui um projeto para o colaborador, ele aparece no Excel do colaborador em segundos. |
| **Dashboard** | Indicadores gerais (total, em andamento, atrasados, vencendo em 7 dias, concluídos, % médio), **resumo por colaborador**, gráfico de barras empilhadas e a **lista de projetos agrupada por colaborador** com dias restantes e situação (No prazo / Vence em até 7 dias / Atrasado / Concluído). Possui filtro por colaborador. |
| **Multidisciplinar** | O mesmo código de projeto pode ser atribuído a vários colaboradores de disciplinas diferentes (só não pode repetir para o mesmo colaborador). Lista de disciplinas editável na aba Config. |
| **Histórico** | Toda inclusão, alteração e exclusão é registrada na aba `Historico` do banco (data/hora, usuário, ação). |

## Como o acesso simultâneo é tratado

1. **Leitura** – o banco é aberto **somente leitura** por uma fração de segundo e fechado. Várias pessoas
   podem ler ao mesmo tempo.
2. **Gravação** – antes de gravar, o Excel abre o arquivo `BD_Projetos.xlsx.lock` com **acesso exclusivo**
   do Windows. Quem chegar depois espera na fila (até 20 s, configurável). Se o Excel de quem está gravando
   travar ou a rede cair, o Windows libera o arquivo sozinho, então **não existe bloqueio preso**.
3. **Edição concorrente** – controle otimista: ao salvar uma edição, o sistema confere se o registro foi
   alterado por outra pessoa desde que você o carregou. Se foi, a gravação é recusada e a lista recarregada,
   e nada é sobrescrito sem aviso.
4. **Detecção de mudanças** – só a data de modificação do arquivo é consultada a cada ciclo, o que é
   muito leve. O banco só é aberto quando algo realmente mudou.

## Estrutura do repositório

```
src/
  modConfig.bas     constantes, parâmetros, regras (data fim, situação), proteção
  modLock.bas       bloqueio exclusivo para gravação (fila entre usuários)
  modBanco.bas      leitura/gravação no banco da rede, criação do banco
  modUsuarios.bas   cadastro, hierarquia chefe/colaborador e permissões
  modProjetos.bas   aba Projetos: salvar, editar, excluir, lista
  modDashboard.bas  aba Dashboard: KPIs, resumo, gráfico e detalhe por colaborador
  modSetup.bas      cria abas, tabelas, nomes, validações e botões
  modApp.bas        Auto_Open / Auto_Close e atualização automática (OnTime)
```

Os arquivos `.bas` estão em **Windows-1252 com quebras de linha CRLF**, que é o formato que o editor
do VBA importa sem estragar os acentos.

## Instalação

### 1. Montar o arquivo `.xlsm` (uma única vez, pelo administrador)

1. Abra um Excel em branco e salve como **`GerenciadorProjetos.xlsm`** (Pasta de Trabalho Habilitada para Macro).
2. Pressione **Alt + F11** para abrir o editor do VBA.
3. Menu **Arquivo > Importar arquivo...** e importe os 8 arquivos da pasta `src/`.
4. Feche o editor, pressione **Alt + F8**, execute a macro **`Iniciar`**. As abas *Projetos, Dashboard,
   Usuarios e Config* são criadas automaticamente.
5. Salve o arquivo.

> Nenhuma referência extra é necessária (`Scripting.Dictionary` é usado por late binding).

### 2. Criar o banco na rede (administrador)

1. Na aba **Config**, informe em *Caminho do banco* o caminho **UNC** da pasta compartilhada, por exemplo
   `\\servidor\projetos\BD_Projetos.xlsx`. Evite letras de unidade (`Z:\`), pois podem variar entre máquinas.
2. Clique em **Criar banco na rede**. Você é cadastrado como **Admin**.
3. Na aba **Usuarios**, clique em **Editar usuários**, preencha uma linha por colaborador e clique em
   **Publicar usuários**:

| Login | Nome | Disciplina | Gestor (login) | Perfil |
|---|---|---|---|---|
| ana.costa | Ana Costa | Coordenação | | Admin |
| carlos.lima | Carlos Lima | Elétrica | ana.costa | Gestor |
| joao.silva | João Silva | Elétrica | carlos.lima | Colaborador |
| maria.souza | Maria Souza | Civil | ana.costa | Colaborador |

   *Login* = nome de usuário do Windows (o que aparece em `%USERNAME%`).

### 3. Distribuir aos colaboradores

Copie o `GerenciadorProjetos.xlsm` (já com o caminho configurado) para cada usuário. Cada um abre a
**própria cópia local**. Ao abrir, o arquivo conecta ao banco, carrega os projetos e liga a atualização
automática.

> **Importante:** cada usuário precisa ter permissão de leitura e gravação na pasta da rede, e as macros
> precisam estar habilitadas. O mais simples é adicionar a pasta onde fica o `.xlsm` como **Local Confiável**
> (*Arquivo > Opções > Central de Confiabilidade > Locais Confiáveis*). Arquivos baixados da internet ou de
> e-mail podem vir bloqueados: *botão direito no arquivo > Propriedades > Desbloquear*.

## Uso diário

**Cadastrar um projeto** – na aba *Projetos*, preencha a linha amarela (o colaborador já vem preenchido
com o seu login; o chefe escolhe qualquer pessoa da equipe na lista) e clique em **Salvar projeto**. A data
fim é calculada na hora.

**Editar / excluir** – clique em uma linha da lista e use **Editar selecionado** (os dados sobem para a
linha amarela; altere e salve) ou **Excluir selecionado**.

**Dashboard** – é reconstruído automaticamente a cada atualização. Use o filtro (célula amarela) para ver
um único colaborador.

## Configurações (aba Config)

| Parâmetro | Padrão | Descrição |
|---|---|---|
| Caminho do banco | – | Arquivo `.xlsx` compartilhado na rede |
| Intervalo da atualização automática | 30 s | Mínimo de 10 s |
| Data fim em dias úteis | SIM | `NAO` = dias corridos |
| Tempo máximo de espera pelo bloqueio | 20 s | Fila de gravação |
| Colaboradores veem projetos de todos | NAO | `SIM` = todos veem tudo (só editam o que é permitido) |
| Feriados (coluna E) | – | Datas desconsideradas no cálculo em dias úteis |
| Disciplinas (coluna K) | – | Lista sugerida no cadastro |

Senha de proteção das abas: `gp2026` (constante `SENHA_PROTECAO` em `modConfig`).

## Limitações e observações

- Projetado para **Excel para Windows** (usa a API `Sleep` do kernel32 e `%USERNAME%`).
- Ninguém deve abrir o `BD_Projetos.xlsx` diretamente no Excel em modo de edição. Se alguém o fizer, as
  gravações são recusadas com uma mensagem clara até que o arquivo seja fechado. Para consulta, abra
  como somente leitura.
- Adequado para equipes de até algumas dezenas de usuários e milhares de projetos. Para volumes maiores,
  o mesmo front-end pode ser adaptado para gravar em Access/SQL Server via ADO.
- Para recriar o layout das abas (por exemplo, após apagar uma aba sem querer), execute a macro
  `ConfigurarPastaDeTrabalho` (Alt + F8) e depois `Iniciar`. Os dados ficam no servidor e são recarregados.
