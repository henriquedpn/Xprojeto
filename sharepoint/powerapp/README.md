# Power App "Gerenciador de Projetos" gerado por Python

`gerar_powerapp.py` gera um app de tela (canvas) para a lista **Projetos** do SharePoint e o empacota
em **`GerenciadorProjetos.msapp`**, que você abre no Power Apps Studio.

> **Atalho:** o arquivo `GerenciadorProjetos.msapp` já está gerado nesta pasta. Se a sua lista foi criada
> com os scripts da pasta `sharepoint/` sem alterações, pule direto para **"Importar no Power Apps"**.

## Telas

| Tela | O que faz |
|---|---|
| **Projetos** | Lista com filtro *Meus projetos / Minha equipe / Todos*, busca por código ou descrição, faixa colorida por status, barra de % concluído e situação do prazo (No prazo / Vence em até 7 dias / Atrasado). Clique em um projeto para editar. |
| **Formulário** | Inclui, edita e exclui (com confirmação). Mostra a **data fim prevista ao vivo**, com a mesma regra de dias úteis da coluna calculada. Novo projeto já vem com você como colaborador. |
| **Dashboard** | 6 indicadores (total, em andamento, atrasados, vencem em 7 dias, concluídos, % médio), **resumo por colaborador** com barra empilhada por situação e painel lateral com os atrasados. Clique em um colaborador para ver os projetos dele. |

## Importar no Power Apps (2 minutos)

1. Acesse **make.powerapps.com** › **Criar** › **Aplicativo de tela em branco** para entrar no Studio.
   No Studio, use **Arquivo** › **Abrir** › **Procurar arquivos** e selecione `GerenciadorProjetos.msapp`.
2. O Studio vai mostrar erros nas fórmulas, porque o app ainda não conhece a lista. Isso é esperado.
3. No painel **Dados** › **Adicionar dados** › **SharePoint** › escolha o site › marque a lista **Projetos**.
   Os erros desaparecem.
4. **Salvar** e **Publicar**. Compartilhe o app com a equipe; cada pessoa precisa ter acesso ao site do
   SharePoint.

Avisos de **delegação** (triângulo amarelo) são esperados na busca por texto e no Dashboard. O app lê até
2.000 projetos, o que é suficiente para o uso típico.

## Gerar de novo (após alterar a lista ou o app)

```bash
dotnet tool install --global Microsoft.PowerApps.CLI.Tool --version 2.15.1   # uma vez (.NET 8+)
python gerar_powerapp.py                                                      # gera fontes/ e o .msapp
python gerar_powerapp.py --nome "Projetos Engenharia" --saida eng.msapp       # outro nome
python gerar_powerapp.py --somente-fontes                                     # só as fontes, sem pac
```

O gerador lê `../definicao-lista.json` (nome da lista e regra de dias úteis ou corridos). Para mudar telas,
cores ou fórmulas, edite as funções `tela_projetos`, `tela_formulario` e `tela_dashboard`. Cada controle é
um objeto `Controle(nome, modelo, propriedades, filhos)` com fórmulas Power Fx em texto.

Antes de empacotar, o gerador confere se toda propriedade usada existe na definição do controle e se não há
nomes repetidos.

## Verificação das fórmulas

`verificacao/verificar.sh` (requer .NET 8+) checa **tipos e nomes** de todas as fórmulas com o motor oficial
do Power Fx (`Microsoft.PowerFx.Interpreter`). A verificação simula a lista Projetos com as colunas reais
(pessoa, escolha, data, número) e os controles do app. Também valida o tipo esperado de cada propriedade:
número em `X`/`Width`, cor em `Fill`, booleano em `Visible`, tabela em `Items` etc.

Situação na versão atual:

- 861 fórmulas verificadas, 0 erros;
- 5.065 regras do `.msapp` analisadas pelo parser do Power Fx, 0 erros de sintaxe;
- a fórmula da data fim, executada pelo Power Fx, bateu com a contagem dia a dia em 1.475 casos;
- nenhum controle fora dos limites da tela ou da galeria.

## Como funciona e limitações

- O app usa o formato de fontes "0.30" do `pac canvas pack`, com as definições de controles da pasta `base/`
  (MIT, Microsoft, veja `base/NOTICE.md`). O `pac` gera um `.msapp` completo, preenchendo os valores padrão
  de cada controle. O Studio atualiza o documento para a versão atual ao abrir.
- A Microsoft marcou esse formato de fontes como **obsoleto** no `pac`, a favor do novo formato `.pa.yaml`.
  Por isso a versão **2.15.1** está fixada acima. O `.msapp` gerado não depende do `pac` depois de pronto.
- A conexão com o SharePoint não vai embutida, porque exige o ID da lista e uma autorização sua. Por isso
  ela é adicionada no passo 3.
- O app usa os **nomes internos** das colunas (`DataInicio`, `StatusProjeto`...). O Studio os exibe com os
  nomes de exibição após conectar a lista.
- **Não foi testado num Power Apps real**: as verificações acima rodam fora do Studio. Faça a primeira
  importação num ambiente de teste. Se o Studio apontar algum erro, me envie a mensagem.
