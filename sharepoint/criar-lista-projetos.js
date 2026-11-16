/*
 * ============================================================================
 *  CRIA A LISTA "Projetos" NO SHAREPOINT ONLINE
 *  Gerenciador de projetos multidisciplinar
 * ============================================================================
 *  COMO USAR
 *   1. Abra no navegador (Edge ou Chrome) o site do SharePoint onde a lista
 *      deve ficar, ex.: https://suaempresa.sharepoint.com/sites/Projetos
 *      (você precisa ser Proprietário ou Membro do site).
 *   2. Pressione F12 e abra a guia "Console".
 *   3. Se o navegador pedir, digite:  allow pasting  (ou "permitir colar").
 *   4. Cole TODO o conteúdo deste arquivo e pressione Enter.
 *   5. Aguarde a mensagem "CONCLUÍDO" (cerca de 30 segundos).
 *
 *  Ajuste o bloco CONFIG abaixo se quiser outro nome de lista, outras
 *  disciplinas ou cálculo em dias corridos.
 * ============================================================================
 */
(function (root) {
  'use strict';

  // ------------------------------------------------------------------ CONFIG
  const CONFIG = {
    titulo: 'Projetos',
    descricao: 'Gerenciador de projetos multidisciplinar',
    diasUteis: true, // true = data fim em dias úteis (seg-sex); false = dias corridos
    disciplinas: ['Arquitetura', 'Civil', 'Estrutural', 'Elétrica', 'Hidráulica',
                  'Mecânica', 'Automação', 'Planejamento', 'Coordenação'],
    status: ['Não iniciado', 'Em andamento', 'Pausado', 'Concluído', 'Cancelado'],
  };

  // ------------------------------------------------------------ DEFINIÇÕES
  const MS_DIA = 86400000;

  function xml(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  // Fórmula da coluna calculada (sintaxe invariável: funções em inglês e vírgula)
  function formulaDataFim(diasUteis) {
    if (!diasUteis) return '=[DataInicio]+[DuracaoDias]-1';
    const wd = 'WEEKDAY([DataInicio],2)'; // 1 = segunda ... 7 = domingo
    // Se começar no fim de semana, empurra para a segunda-feira seguinte
    return '=[DataInicio]+IF(' + wd + '>5,8-' + wd + ',0)+([DuracaoDias]-1)' +
           '+INT((IF(' + wd + '>5,0,' + wd + '-1)+[DuracaoDias]-1)/5)*2';
  }

  // Mesma conta em "expressão de formatação JSON", feita a partir de colunas
  // tipadas (o valor de colunas calculadas chega como texto na formatação).
  function exprFimMs(diasUteis) {
    const ini = 'Number([$DataInicio])';
    const n = '([$DuracaoDias]-1)';
    if (!diasUteis) return '(' + ini + '+' + n + '*' + MS_DIA + ')';
    const dias = 'floor((' + ini + '+43200000)/' + MS_DIA + ')';      // dias desde 01/01/1970
    const wd = '((' + dias + '+3)-floor((' + dias + '+3)/7)*7)';       // 0 = segunda ... 6 = domingo
    const desloc = '(if(' + wd + '>4,7-' + wd + ',0))';
    const wd2 = '(if(' + wd + '>4,0,' + wd + '))';
    return '(' + ini + '+(' + desloc + '+' + n + '+floor((' + wd2 + '+' + n + ')/5)*2)*' + MS_DIA + ')';
  }

  function definicao(cfg) {
    const fim = exprFimMs(cfg.diasUteis);
    const aberto = "([$StatusProjeto]!='Concluído'&&[$StatusProjeto]!='Cancelado'&&[$DataInicio]!='')";
    const atrasado = '(' + aberto + '&&' + fim + '+' + MS_DIA + '<=Number(@now))';
    const vence = '(' + aberto + '&&' + fim + '<=Number(@now)+' + 7 * MS_DIA + ')';

    const fmtStatus = {
      $schema: 'https://developer.microsoft.com/json-schemas/sp/v2/column-formatting.schema.json',
      elmType: 'div',
      txtContent: '@currentField',
      style: {
        padding: '2px 10px', 'border-radius': '12px', color: 'white', 'font-weight': '600',
        'text-align': 'center', 'white-space': 'nowrap',
        'background-color': "=if(@currentField=='Concluído','#107c10',if(@currentField=='Em andamento','#0078d4'," +
          "if(@currentField=='Pausado','#c19c00',if(@currentField=='Cancelado','#7a7574','#a19f9d'))))",
      },
    };

    const fmtPerc = {
      $schema: 'https://developer.microsoft.com/json-schemas/sp/v2/column-formatting.schema.json',
      elmType: 'div',
      style: { width: '100%', 'background-color': '#edebe9', 'border-radius': '4px', height: '18px' },
      children: [{
        elmType: 'div',
        txtContent: "=if(@currentField=='','0%',@currentField+'%')",
        style: {
          width: "=if(@currentField=='','0%',@currentField+'%')",
          'min-width': '28px', height: '18px', 'line-height': '18px', 'border-radius': '4px',
          'padding-left': '4px', color: 'white', 'font-size': '11px',
          'background-color': "=if(@currentField>=100,'#107c10','#0078d4')",
        },
      }],
    };

    const fmtDataFim = {
      $schema: 'https://developer.microsoft.com/json-schemas/sp/v2/column-formatting.schema.json',
      elmType: 'div',
      style: {
        display: 'flex', 'align-items': 'center',
        color: "=if(" + atrasado + ",'#a4262c',if(" + vence + ",'#8a5300',''))",
        'font-weight': "=if(" + atrasado + "||" + vence + ",'600','')",
      },
      children: [
        { elmType: 'span', style: { 'padding-right': '4px' },
          attributes: { iconName: "=if(" + atrasado + ",'Warning',if(" + vence + ",'Clock',''))" } },
        { elmType: 'span', txtContent: '@currentField' },
      ],
    };

    const fmtLinha = {
      $schema: 'https://developer.microsoft.com/json-schemas/sp/v2/row-formatting.schema.json',
      additionalRowClass: "=if(" + atrasado + ",'sp-field-severity--severeWarning',if(" + vence +
        ",'sp-field-severity--warning',''))",
    };

    const formulario = JSON.stringify({
      headerJSONFormatter: JSON.stringify({
        elmType: 'div',
        style: { display: 'flex', 'align-items': 'center', 'margin-bottom': '12px' },
        children: [
          { elmType: 'span', attributes: { iconName: 'ProjectCollection' },
            style: { 'font-size': '28px', 'padding-right': '10px', color: '#0078d4' } },
          { elmType: 'div', style: { 'font-size': '22px', 'font-weight': '600' },
            txtContent: "=if([$Title]=='','Novo projeto','Projeto ' + [$Title])" },
        ],
      }),
      bodyJSONFormatter: JSON.stringify({
        sections: [
          { displayname: 'Identificação', fields: ['Código do projeto', 'Descrição', 'Disciplina'] },
          { displayname: 'Responsáveis', fields: ['Colaborador', 'Gestor'] },
          { displayname: 'Prazo', fields: ['Data início', 'Duração (dias)'] },
          { displayname: 'Andamento', fields: ['Status', '% concluído', 'Observações'] },
        ],
      }),
      footerJSONFormatter: '',
    });

    const choices = (arr) => arr.map((c) => '<CHOICE>' + xml(c) + '</CHOICE>').join('');

    // Os campos são criados com DisplayName = nome interno e renomeados no fim
    // (assim os nomes internos ficam sem acento e a fórmula sempre resolve).
    const campos = [
      { nome: 'Descricao', titulo: 'Descrição',
        xml: '<Field Type="Note" Name="Descricao" StaticName="Descricao" DisplayName="Descricao" NumLines="4" RichText="FALSE" />' },
      { nome: 'Disciplina', titulo: 'Disciplina',
        xml: '<Field Type="Choice" Name="Disciplina" StaticName="Disciplina" DisplayName="Disciplina" Format="Dropdown" FillInChoice="TRUE"><CHOICES>' +
             choices(cfg.disciplinas) + '</CHOICES></Field>' },
      { nome: 'Colaborador', titulo: 'Colaborador',
        xml: '<Field Type="User" Name="Colaborador" StaticName="Colaborador" DisplayName="Colaborador" Required="TRUE" UserSelectionMode="PeopleOnly" UserSelectionScope="0" List="UserInfo" ShowField="ImnName" />' },
      { nome: 'Gestor', titulo: 'Gestor',
        xml: '<Field Type="User" Name="Gestor" StaticName="Gestor" DisplayName="Gestor" Required="FALSE" UserSelectionMode="PeopleOnly" UserSelectionScope="0" List="UserInfo" ShowField="ImnName" Description="Chefe imediato do colaborador (usado no modo de exibição Minha equipe)" />' },
      { nome: 'DataInicio', titulo: 'Data início',
        xml: '<Field Type="DateTime" Name="DataInicio" StaticName="DataInicio" DisplayName="DataInicio" Format="DateOnly" Required="TRUE"><Default>[today]</Default></Field>' },
      { nome: 'DuracaoDias', titulo: 'Duração (dias)',
        xml: '<Field Type="Number" Name="DuracaoDias" StaticName="DuracaoDias" DisplayName="DuracaoDias" Required="TRUE" Min="1" Max="3650" Decimals="0" Description="Duração estimada em dias ' +
             (cfg.diasUteis ? 'úteis' : 'corridos') + '" />' },
      { nome: 'DataFim', titulo: 'Data fim', formatador: fmtDataFim,
        xml: '<Field Type="Calculated" Name="DataFim" StaticName="DataFim" DisplayName="DataFim" ResultType="DateTime" Format="DateOnly" ReadOnly="TRUE"><Formula>' +
             xml(formulaDataFim(cfg.diasUteis)) +
             '</Formula><FieldRefs><FieldRef Name="DataInicio" /><FieldRef Name="DuracaoDias" /></FieldRefs></Field>' },
      { nome: 'StatusProjeto', titulo: 'Status', formatador: fmtStatus,
        xml: '<Field Type="Choice" Name="StatusProjeto" StaticName="StatusProjeto" DisplayName="StatusProjeto" Required="TRUE" Format="Dropdown" FillInChoice="FALSE"><Default>' +
             xml(cfg.status[0]) + '</Default><CHOICES>' + choices(cfg.status) + '</CHOICES></Field>' },
      { nome: 'PercConcluido', titulo: '% concluído', formatador: fmtPerc,
        xml: '<Field Type="Number" Name="PercConcluido" StaticName="PercConcluido" DisplayName="PercConcluido" Min="0" Max="100" Decimals="0"><Default>0</Default></Field>' },
      { nome: 'Observacoes', titulo: 'Observações',
        xml: '<Field Type="Note" Name="Observacoes" StaticName="Observacoes" DisplayName="Observacoes" NumLines="4" RichText="FALSE" />' },
    ];

    const camposView = ['LinkTitle', 'Descricao', 'Colaborador', 'Disciplina', 'DataInicio',
                        'DuracaoDias', 'DataFim', 'StatusProjeto', 'PercConcluido'];
    const eu = '<Value Type="Integer"><UserID Type="Integer" /></Value>';
    const naoFechado = '<And><Neq><FieldRef Name="StatusProjeto" /><Value Type="Choice">Concluído</Value></Neq>' +
                       '<Neq><FieldRef Name="StatusProjeto" /><Value Type="Choice">Cancelado</Value></Neq></And>';

    const views = [
      { titulo: 'Meus projetos',
        query: '<Where><Eq><FieldRef Name="Colaborador" LookupId="TRUE" />' + eu + '</Eq></Where>' +
               '<OrderBy><FieldRef Name="DataInicio" /></OrderBy>' },
      { titulo: 'Minha equipe',
        query: '<Where><Eq><FieldRef Name="Gestor" LookupId="TRUE" />' + eu + '</Eq></Where>' +
               '<GroupBy Collapse="FALSE" GroupLimit="100"><FieldRef Name="Colaborador" /></GroupBy>' +
               '<OrderBy><FieldRef Name="DataInicio" /></OrderBy>',
        totais: true },
      { titulo: 'Por colaborador',
        query: '<GroupBy Collapse="FALSE" GroupLimit="100"><FieldRef Name="Colaborador" /></GroupBy>' +
               '<OrderBy><FieldRef Name="DataFim" /></OrderBy>',
        totais: true },
      { titulo: 'Atrasados',
        query: '<Where><And><Lt><FieldRef Name="DataFim" /><Value Type="DateTime"><Today /></Value></Lt>' +
               naoFechado + '</And></Where><OrderBy><FieldRef Name="DataFim" /></OrderBy>' },
    ];

    return {
      titulo: cfg.titulo,
      descricao: cfg.descricao,
      tituloCodigo: 'Código do projeto',
      campos: campos.map((c) => ({
        nome: c.nome, titulo: c.titulo, xml: c.xml,
        formatador: c.formatador ? JSON.stringify(c.formatador) : null,
      })),
      camposView: camposView,
      views: views,
      consultaPadrao: '<OrderBy><FieldRef Name="DataFim" /></OrderBy>',
      formatadorLinha: JSON.stringify(fmtLinha),
      totais: '<FieldRef Name="LinkTitle" Type="COUNT" />',
      formulario: formulario,
      formatadoresLegiveis: { status: fmtStatus, percentual: fmtPerc, dataFim: fmtDataFim, linha: fmtLinha },
    };
  }

  // ------------------------------------------------------------- REST API
  function urlDoSite() {
    if (root._spPageContextInfo && root._spPageContextInfo.webAbsoluteUrl) {
      return root._spPageContextInfo.webAbsoluteUrl;
    }
    const m = root.location.pathname.match(/^\/(sites|teams)\/[^/]+/i);
    return root.location.origin + (m ? m[0] : '');
  }

  function criarCliente(site) {
    let digest = null;
    async function chamar(metodo, caminho, corpo, extras) {
      const headers = Object.assign({
        Accept: 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
      }, digest ? { 'X-RequestDigest': digest } : {}, extras || {});
      const r = await root.fetch(site + caminho, {
        method: metodo, headers: headers, credentials: 'same-origin',
        body: corpo === undefined ? undefined : JSON.stringify(corpo),
      });
      const txt = await r.text();
      if (!r.ok) {
        let msg = txt;
        try { msg = JSON.parse(txt).error.message.value; } catch (e) { /* mantém texto */ }
        const err = new Error(metodo + ' ' + caminho + ' -> ' + r.status + ': ' + msg);
        err.status = r.status;
        throw err;
      }
      return txt ? JSON.parse(txt) : null;
    }
    return {
      async iniciar() {
        const r = await chamar('POST', '/_api/contextinfo');
        digest = r.d.GetContextWebInformation.FormDigestValue;
      },
      get: (c) => chamar('GET', c),
      post: (c, corpo) => chamar('POST', c, corpo),
      merge: (c, corpo) => chamar('POST', c, corpo, { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' }),
    };
  }

  const q = (s) => encodeURIComponent(String(s).replace(/'/g, "''"));

  async function main() {
    const def = definicao(CONFIG);
    const site = urlDoSite();
    const L = "/_api/web/lists/getbytitle('" + q(def.titulo) + "')";
    const sp = criarCliente(site);
    const passo = (m) => console.log('%c► ' + m, 'color:#0078d4;font-weight:bold');
    const aviso = (m) => console.warn('⚠ ' + m);

    console.log('%cCriando a lista "' + def.titulo + '" em ' + site, 'font-size:14px;font-weight:bold');
    await sp.iniciar();

    try {
      await sp.get(L + '?$select=Id');
      console.error('Já existe uma lista chamada "' + def.titulo + '" neste site. ' +
                    'Exclua-a ou altere CONFIG.titulo no início do script e rode de novo.');
      return;
    } catch (e) {
      if (e.status !== 404) throw e;
    }

    passo('Criando a lista');
    await sp.post('/_api/web/lists', {
      __metadata: { type: 'SP.List' }, BaseTemplate: 100, Title: def.titulo, Description: def.descricao,
    });
    await sp.merge(L, { __metadata: { type: 'SP.List' }, EnableVersioning: true });

    passo('Criando as colunas');
    for (const c of def.campos) {
      await sp.post(L + '/fields/createfieldasxml', {
        parameters: { __metadata: { type: 'SP.XmlSchemaFieldCreationInformation' }, SchemaXml: c.xml, Options: 12 },
      });
      console.log('   coluna ' + c.titulo);
    }

    passo('Renomeando colunas e aplicando formatação');
    await sp.merge(L + "/fields/getbyinternalnameortitle('Title')",
      { __metadata: { type: 'SP.Field' }, Title: def.tituloCodigo });
    for (const c of def.campos) {
      const corpo = { __metadata: { type: 'SP.Field' }, Title: c.titulo };
      if (c.formatador) corpo.CustomFormatter = c.formatador;
      try {
        await sp.merge(L + "/fields/getbyinternalnameortitle('" + c.nome + "')", corpo);
      } catch (e) {
        aviso('Não foi possível ajustar a coluna ' + c.titulo + ': ' + e.message);
      }
    }

    async function configurarView(base, campos, extras) {
      await sp.post(base + '/viewfields/removeallviewfields');
      for (const f of campos) {
        await sp.post(base + "/viewfields/addviewfield('" + f + "')");
      }
      await sp.merge(base, Object.assign({ __metadata: { type: 'SP.View' } }, extras));
    }

    passo('Configurando o modo de exibição padrão');
    await configurarView(L + '/defaultview', def.camposView, {
      ViewQuery: def.consultaPadrao, CustomFormatter: def.formatadorLinha, RowLimit: 100,
    });

    for (const v of def.views) {
      passo('Criando o modo de exibição "' + v.titulo + '"');
      try {
        await sp.post(L + '/views', {
          __metadata: { type: 'SP.View' }, Title: v.titulo, PersonalView: false,
          ViewQuery: v.query, RowLimit: 100, Paged: true,
        });
        const extras = { CustomFormatter: def.formatadorLinha };
        if (v.totais) { extras.Aggregations = def.totais; extras.AggregationsStatus = 'On'; }
        await configurarView(L + "/views/getbytitle('" + q(v.titulo) + "')", def.camposView, extras);
      } catch (e) {
        aviso('Modo de exibição "' + v.titulo + '": ' + e.message);
      }
    }

    passo('Organizando o formulário em seções');
    try {
      const cts = await sp.get(L + '/contenttypes?$select=StringId,Name');
      const ct = cts.d.results.find((x) => /^0x01(00|$)/i.test(x.StringId) && !/^0x0120/i.test(x.StringId));
      if (ct) {
        await sp.merge(L + "/contenttypes('" + ct.StringId + "')",
          { __metadata: { type: 'SP.ContentType' }, ClientFormCustomFormatter: def.formulario });
      }
    } catch (e) {
      aviso('Layout do formulário não aplicado (opcional): ' + e.message);
    }

    const url = site + '/Lists/' + encodeURIComponent(def.titulo);
    console.log('%cCONCLUÍDO! Abra a lista: ' + url, 'color:#107c10;font-size:14px;font-weight:bold');
    return url;
  }

  // -------------------------------------------------------------- execução
  if (typeof module !== 'undefined' && module.exports) {
    module.exports = { CONFIG: CONFIG, definicao: definicao, formulaDataFim: formulaDataFim,
                       exprFimMs: exprFimMs, main: main };
  } else {
    main().catch((e) => console.error('ERRO: ' + e.message + '\nNada foi apagado. Corrija e rode novamente ' +
      '(se a lista chegou a ser criada, exclua-a antes).'));
  }
})(typeof window !== 'undefined' ? window : globalThis);
