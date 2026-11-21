/*
 * JSON Compartilhado na Rede — controle de concorrência 100% no cliente.
 *
 * Arquivos na pasta da rede:
 *   dados.json         -> os dados (com meta.versao incrementada a cada gravação)
 *   dados.lock.json    -> registro de quem está usando o arquivo e fazendo o quê
 *   dados.backup.json  -> cópia da versão anterior (gravada antes de cada salvamento)
 *
 * Protocolo (todos os clientes seguem as mesmas regras):
 *   1. Para editar, o cliente lê o lock. Se existir um lock "vivo" de outra pessoa, não entra.
 *   2. Se estiver livre, grava o próprio lock com um token aleatório, espera um tempo
 *      aleatório e relê. Só considera o lock seu se o token lido for o seu (desempate
 *      quando duas pessoas tentam ao mesmo tempo: quem gravou por último ganha, o outro desiste).
 *   3. Enquanto edita, renova o lock (heartbeat) a cada poucos segundos.
 *   4. Um lock é considerado abandonado se o conteúdo dele não muda por LOCK_TTL_MS,
 *      medido pelo relógio de QUEM OBSERVA (assim não depende dos relógios das máquinas
 *      estarem sincronizados). Navegador fechado / máquina travada => lock expira sozinho.
 *   5. Ao salvar: confere que o lock ainda é seu, confere que meta.versao não mudou,
 *      grava backup, grava dados com versao+1, relê para validar e remove o lock.
 *   6. Leitura de JSON inválido (arquivo pela metade) => tenta de novo; se persistir,
 *      oferece restaurar o backup.
 */
'use strict';

const CFG = {
  ARQ_DADOS: 'dados.json',
  ARQ_LOCK: 'dados.lock.json',
  ARQ_BACKUP: 'dados.backup.json',
  LOCK_TTL_MS: 30000,          // lock sem alteração por esse tempo = abandonado
  HEARTBEAT_MS: 5000,          // renovação do lock enquanto edita
  POLL_MS: 3000,               // atualização da tela / observação do lock
  ESPERA_CONFIRMACAO_MS: [800, 1600], // espera aleatória antes de confirmar o lock
  TENTATIVAS_LEITURA: 4,
};

// ---------------------------------------------------------------------------
// Estado
// ---------------------------------------------------------------------------
const estado = {
  dir: null,                 // FileSystemDirectoryHandle da pasta da rede
  sessao: novoId(),          // identifica esta aba
  meuLock: null,             // objeto do lock que eu gravei (quando tenho o lock)
  dados: null,               // última versão lida do dados.json
  editando: false,
  versaoBase: null,          // versão do arquivo quando comecei a editar
  rascunho: null,            // cópia dos registros sendo editada
  observacaoLock: null,      // { raw, vistoEm } para detectar lock abandonado
  timerPoll: null,
  timerHeartbeat: null,
  dadosCorrompidos: false,
};

// Fila simples para que operações de arquivo desta aba não se atropelem.
let fila = Promise.resolve();
function serializar(fn) {
  const p = fila.then(fn, fn);
  fila = p.catch(() => {});
  return p;
}

// ---------------------------------------------------------------------------
// Utilidades
// ---------------------------------------------------------------------------
function novoId() {
  if (crypto.randomUUID) return crypto.randomUUID();
  return Array.from(crypto.getRandomValues(new Uint8Array(16)), b => b.toString(16).padStart(2, '0')).join('');
}
const dormir = ms => new Promise(r => setTimeout(r, ms));
const aleatorio = (min, max) => Math.floor(min + Math.random() * (max - min));
const agoraISO = () => new Date().toISOString();
const $ = id => document.getElementById(id);

function usuario() {
  return ($('nomeUsuario').value || '').trim() || 'Anônimo';
}

function log(msg, tipo = '') {
  const li = document.createElement('li');
  li.className = tipo;
  li.textContent = `[${new Date().toLocaleTimeString()}] ${msg}`;
  $('log').prepend(li);
  while ($('log').children.length > 200) $('log').lastChild.remove();
}

function formatarData(iso) {
  if (!iso) return '—';
  const d = new Date(iso);
  return isNaN(d) ? iso : d.toLocaleString();
}

// ---------------------------------------------------------------------------
// Acesso a arquivos (File System Access API)
// ---------------------------------------------------------------------------
async function lerTexto(nome) {
  try {
    const fh = await estado.dir.getFileHandle(nome);
    const arquivo = await fh.getFile();      // sempre um "retrato" novo do disco
    return await arquivo.text();
  } catch (e) {
    if (e.name === 'NotFoundError') return null;
    throw e;
  }
}

/** Lê e faz parse; repete se o JSON vier quebrado (alguém gravando no mesmo instante). */
async function lerJSON(nome, tentativas = CFG.TENTATIVAS_LEITURA) {
  let ultimoErro;
  for (let i = 0; i < tentativas; i++) {
    const texto = await lerTexto(nome);
    if (texto === null) return { existe: false, dados: null, raw: null };
    try {
      return { existe: true, dados: JSON.parse(texto), raw: texto };
    } catch (e) {
      ultimoErro = e;
      await dormir(aleatorio(300, 700));
    }
  }
  return { existe: true, dados: null, raw: null, erro: ultimoErro };
}

/**
 * Grava o arquivo inteiro. O createWritable() do Chrome/Edge grava num arquivo
 * temporário (.crswap) e só substitui o original no close() — então quem lê
 * nunca vê o arquivo pela metade.
 */
async function gravarJSON(nome, obj) {
  const fh = await estado.dir.getFileHandle(nome, { create: true });
  const w = await fh.createWritable({ keepExistingData: false });
  try {
    await w.write(JSON.stringify(obj, null, 2));
    await w.close();
  } catch (e) {
    try { await w.abort(); } catch (_) {}
    throw e;
  }
}

async function removerArquivo(nome) {
  try {
    await estado.dir.removeEntry(nome);
  } catch (e) {
    if (e.name !== 'NotFoundError') throw e;
  }
}

// ---------------------------------------------------------------------------
// Persistência do handle da pasta (IndexedDB) para não escolher toda vez
// ---------------------------------------------------------------------------
function abrirIDB() {
  return new Promise((ok, falha) => {
    const req = indexedDB.open('json-rede', 1);
    req.onupgradeneeded = () => req.result.createObjectStore('handles');
    req.onsuccess = () => ok(req.result);
    req.onerror = () => falha(req.error);
  });
}
async function idbSet(chave, valor) {
  const db = await abrirIDB();
  return new Promise((ok, falha) => {
    const tx = db.transaction('handles', 'readwrite');
    tx.objectStore('handles').put(valor, chave);
    tx.oncomplete = () => ok();
    tx.onerror = () => falha(tx.error);
  });
}
async function idbGet(chave) {
  const db = await abrirIDB();
  return new Promise((ok, falha) => {
    const req = db.transaction('handles').objectStore('handles').get(chave);
    req.onsuccess = () => ok(req.result);
    req.onerror = () => falha(req.error);
  });
}

// ---------------------------------------------------------------------------
// LOCK
// ---------------------------------------------------------------------------

/** Lê o lock e atualiza a observação usada para detectar lock abandonado. */
async function lerLock() {
  const r = await lerJSON(CFG.ARQ_LOCK, 2);
  if (!r.existe) {
    estado.observacaoLock = null;
    return null;
  }
  // Lock ilegível: trata como ocupado (por segurança) até expirar pela observação.
  const raw = r.raw ?? '__lock_ilegivel__';
  const obs = estado.observacaoLock;
  if (!obs || obs.raw !== raw) {
    estado.observacaoLock = { raw, vistoEm: Date.now() };
  }
  return r.dados ?? { ilegivel: true };
}

/** Lock está vivo se o conteúdo mudou (heartbeat) há menos de TTL, pelo MEU relógio. */
function lockEstaVivo() {
  const obs = estado.observacaoLock;
  if (!obs) return false;
  return Date.now() - obs.vistoEm < CFG.LOCK_TTL_MS;
}

function lockEhMeu(lock) {
  return !!(lock && estado.meuLock && lock.token === estado.meuLock.token);
}

function montarLock(operacao) {
  return {
    token: novoId(),
    sessao: estado.sessao,
    usuario: usuario(),
    operacao,                       // "editando", "gravando", "criando", "restaurando"...
    desde: agoraISO(),
    heartbeat: 0,
    atualizadoEm: agoraISO(),
    navegador: navigator.userAgent.slice(0, 120),
  };
}

/**
 * Tenta adquirir o lock.
 * Retorna { ok: true } ou { ok: false, dono } com quem está segurando.
 */
function adquirirLock(operacao) {
  return serializar(async () => {
    const atual = await lerLock();
    if (atual && lockEhMeu(atual)) return { ok: true };
    if (atual && lockEstaVivo()) return { ok: false, dono: atual };
    if (atual) log(`Lock abandonado de "${atual.usuario ?? '?'}" detectado — assumindo.`, 'alerta');

    const meu = montarLock(operacao);
    await gravarJSON(CFG.ARQ_LOCK, meu);

    // Espera aleatória: se outra pessoa gravou ao mesmo tempo, o último a gravar vence.
    await dormir(aleatorio(...CFG.ESPERA_CONFIRMACAO_MS));
    const conf1 = await lerLock();
    if (!conf1 || conf1.token !== meu.token) {
      return { ok: false, dono: conf1 };
    }
    // Segunda confirmação curta (pega gravações que chegaram atrasadas pela rede).
    await dormir(aleatorio(200, 500));
    const conf2 = await lerLock();
    if (!conf2 || conf2.token !== meu.token) {
      return { ok: false, dono: conf2 };
    }

    estado.meuLock = meu;
    iniciarHeartbeat();
    log(`Lock adquirido (${operacao}).`, 'ok');
    return { ok: true };
  });
}

/** Atualiza o lock (heartbeat e/ou operação). Retorna false se perdi o lock. */
function renovarLock(novaOperacao) {
  return serializar(async () => {
    if (!estado.meuLock) return false;
    const atual = await lerLock();
    if (!lockEhMeu(atual)) {
      perdiOLock(atual);
      return false;
    }
    estado.meuLock = {
      ...estado.meuLock,
      operacao: novaOperacao ?? estado.meuLock.operacao,
      heartbeat: estado.meuLock.heartbeat + 1,
      atualizadoEm: agoraISO(),
    };
    await gravarJSON(CFG.ARQ_LOCK, estado.meuLock);
    return true;
  });
}

function liberarLock() {
  return serializar(async () => {
    pararHeartbeat();
    if (!estado.meuLock) return;
    const atual = await lerLock();
    if (lockEhMeu(atual)) {
      await removerArquivo(CFG.ARQ_LOCK);
      estado.observacaoLock = null;
      log('Lock liberado.', 'ok');
    }
    estado.meuLock = null;
  });
}

function perdiOLock(dono) {
  pararHeartbeat();
  estado.meuLock = null;
  log(`ATENÇÃO: o lock foi perdido${dono ? ` para "${dono.usuario}"` : ''}. ` +
      'Ao salvar, o sistema tentará recuperar o lock se ninguém tiver alterado o arquivo.', 'erro');
}

function iniciarHeartbeat() {
  pararHeartbeat();
  estado.timerHeartbeat = setInterval(() => {
    renovarLock().catch(e => log('Falha no heartbeat: ' + e.message, 'erro'));
  }, CFG.HEARTBEAT_MS);
}
function pararHeartbeat() {
  clearInterval(estado.timerHeartbeat);
  estado.timerHeartbeat = null;
}

// ---------------------------------------------------------------------------
// DADOS
// ---------------------------------------------------------------------------
function estruturaVazia() {
  return {
    meta: { versao: 0, criadoEm: agoraISO(), atualizadoEm: agoraISO(), atualizadoPor: usuario() },
    registros: [],
  };
}

function validarEstrutura(d) {
  return d && typeof d === 'object' && d.meta && typeof d.meta.versao === 'number' && Array.isArray(d.registros);
}

/** Lê dados.json. Retorna o objeto, null se não existe, ou lança erro se corrompido. */
async function lerDados() {
  const r = await lerJSON(CFG.ARQ_DADOS);
  if (!r.existe) return null;
  if (!validarEstrutura(r.dados)) {
    const e = new Error('dados.json está corrompido ou em formato inesperado.');
    e.corrompido = true;
    throw e;
  }
  return r.dados;
}

async function garantirArquivoDados() {
  if (await lerDados() !== null) return;
  log('dados.json não existe — criando arquivo novo.', 'alerta');
  const r = await adquirirLock('criando');
  if (!r.ok) return; // outra pessoa deve estar criando
  try {
    if (await lerDados() === null) await gravarJSON(CFG.ARQ_DADOS, estruturaVazia());
  } finally {
    await liberarLock();
  }
}

// ---------------------------------------------------------------------------
// Fluxo de edição
// ---------------------------------------------------------------------------
async function iniciarEdicao() {
  if (!estado.dir) return;
  if (!($('nomeUsuario').value || '').trim()) {
    alert('Informe seu nome antes de editar (ele aparece para os outros usuários).');
    $('nomeUsuario').focus();
    return;
  }
  travarBotoes(true);
  try {
    const r = await adquirirLock('editando');
    if (!r.ok) {
      const d = r.dono;
      alert(d && !d.ilegivel
        ? `O arquivo está em uso por "${d.usuario}" (${d.operacao}) desde ${formatarData(d.desde)}.\nTente novamente mais tarde.`
        : 'Outra pessoa acabou de pegar o arquivo. Tente novamente em instantes.');
      return;
    }
    // Com o lock em mãos, relê a versão mais recente antes de editar.
    const dados = await lerDados() ?? estruturaVazia();
    estado.dados = dados;
    estado.versaoBase = dados.meta.versao;
    estado.rascunho = structuredClone(dados.registros);
    estado.editando = true;
    log(`Edição iniciada sobre a versão ${estado.versaoBase}.`);
  } catch (e) {
    log('Erro ao iniciar edição: ' + e.message, 'erro');
    await liberarLock().catch(() => {});
  } finally {
    travarBotoes(false);
    renderizar();
  }
}

async function salvar() {
  if (!estado.editando) return;
  travarBotoes(true);
  try {
    // 1. Garante que o lock ainda é meu (ou recupera, se estiver livre).
    let tenho = await renovarLock('gravando');
    if (!tenho) {
      const r = await adquirirLock('gravando');
      if (!r.ok) {
        alert(`Não foi possível salvar: o arquivo está bloqueado por "${r.dono?.usuario ?? '?'}".\n` +
              'Suas alterações continuam na tela; tente salvar de novo depois.');
        return;
      }
      tenho = true;
    }

    // 2. Confere se ninguém alterou o arquivo desde que comecei (controle otimista).
    const atual = await lerDados();
    const versaoAtual = atual ? atual.meta.versao : 0;
    if (versaoAtual !== estado.versaoBase) {
      alert(`Conflito: o arquivo foi alterado por "${atual.meta.atualizadoPor}" ` +
            `(versão ${versaoAtual}, você editou a ${estado.versaoBase}).\n` +
            'Suas alterações não foram gravadas. Copie o que precisar e cancele a edição.');
      log('Conflito de versão — gravação abortada.', 'erro');
      await renovarLock('editando');
      return;
    }

    // 3. Backup da versão atual.
    if (atual) await gravarJSON(CFG.ARQ_BACKUP, atual);

    // 4. Grava a nova versão.
    const novo = {
      meta: {
        ...(atual?.meta ?? estruturaVazia().meta),
        versao: versaoAtual + 1,
        atualizadoEm: agoraISO(),
        atualizadoPor: usuario(),
      },
      registros: estado.rascunho,
    };
    await gravarJSON(CFG.ARQ_DADOS, novo);

    // 5. Valida relendo.
    const conferido = await lerDados();
    if (!conferido || conferido.meta.versao !== novo.meta.versao) {
      throw new Error('A verificação após gravar falhou (versão lida diferente da gravada).');
    }

    estado.dados = conferido;
    encerrarEdicaoLocal();
    await liberarLock();
    log(`Salvo com sucesso — versão ${novo.meta.versao}.`, 'ok');
  } catch (e) {
    log('Erro ao salvar: ' + e.message, 'erro');
    alert('Erro ao salvar: ' + e.message + '\nO lock continua seu; tente salvar novamente.');
    await renovarLock('editando').catch(() => {});
  } finally {
    travarBotoes(false);
    renderizar();
  }
}

async function cancelar() {
  if (estado.editando && !confirm('Descartar as alterações não salvas?')) return;
  encerrarEdicaoLocal();
  await liberarLock().catch(e => log('Erro ao liberar: ' + e.message, 'erro'));
  await atualizar();
}

function encerrarEdicaoLocal() {
  estado.editando = false;
  estado.rascunho = null;
  estado.versaoBase = null;
}

async function forcarLiberacao() {
  const lock = await serializar(lerLock);
  if (!lock) return atualizar();
  const quem = lock.ilegivel ? '(lock ilegível)' : `"${lock.usuario}" (${lock.operacao}) desde ${formatarData(lock.desde)}`;
  if (!confirm(`Remover à força o bloqueio de ${quem}?\n\nSó faça isso se tiver certeza de que essa pessoa não está mais editando.`)) return;
  await serializar(() => removerArquivo(CFG.ARQ_LOCK));
  estado.observacaoLock = null;
  log(`Lock de ${quem} removido à força.`, 'alerta');
  await atualizar();
}

async function restaurarBackup() {
  const bkp = await lerJSON(CFG.ARQ_BACKUP);
  if (!bkp.existe || !validarEstrutura(bkp.dados)) {
    alert('Não há backup válido disponível.');
    return;
  }
  if (!confirm(`Restaurar o backup (versão ${bkp.dados.meta.versao}, de ${formatarData(bkp.dados.meta.atualizadoEm)})?`)) return;
  const r = await adquirirLock('restaurando');
  if (!r.ok) { alert(`Arquivo bloqueado por "${r.dono?.usuario ?? '?'}".`); return; }
  try {
    const restaurado = structuredClone(bkp.dados);
    restaurado.meta.versao = bkp.dados.meta.versao + 1000; // garante versão nova e maior
    restaurado.meta.atualizadoEm = agoraISO();
    restaurado.meta.atualizadoPor = usuario() + ' (restauração de backup)';
    await gravarJSON(CFG.ARQ_DADOS, restaurado);
    estado.dadosCorrompidos = false;
    log('Backup restaurado.', 'ok');
  } finally {
    await liberarLock();
    await atualizar();
  }
}

// ---------------------------------------------------------------------------
// Atualização periódica
// ---------------------------------------------------------------------------
let atualizando = false;
async function atualizar() {
  if (!estado.dir || atualizando) return;
  atualizando = true;
  try {
    await serializar(lerLock); // mantém a observação do lock em dia
    if (!estado.editando) {
      try {
        estado.dados = await lerDados();
        estado.dadosCorrompidos = false;
      } catch (e) {
        if (e.corrompido) {
          if (!estado.dadosCorrompidos) log(e.message + ' Use "Restaurar backup" se não se resolver sozinho.', 'erro');
          estado.dadosCorrompidos = true;
        } else throw e;
      }
    }
  } catch (e) {
    log('Erro ao atualizar: ' + e.message, 'erro');
  } finally {
    atualizando = false;
    renderizar();
  }
}

function iniciarPolling() {
  clearInterval(estado.timerPoll);
  estado.timerPoll = setInterval(atualizar, CFG.POLL_MS);
}

// ---------------------------------------------------------------------------
// Conexão com a pasta
// ---------------------------------------------------------------------------
async function conectar(dir) {
  estado.dir = dir;
  $('nomePasta').textContent = 'Pasta: ' + dir.name;
  $('btnReconectar').classList.add('oculto');
  try { await idbSet('pasta', dir); } catch (_) {}
  log(`Conectado à pasta "${dir.name}".`, 'ok');
  try {
    await garantirArquivoDados();
  } catch (e) {
    log('Erro ao preparar dados.json: ' + e.message, 'erro');
  }
  await atualizar();
  iniciarPolling();
}

async function escolherPasta() {
  try {
    const dir = await window.showDirectoryPicker({ id: 'json-rede', mode: 'readwrite' });
    await conectar(dir);
  } catch (e) {
    if (e.name === 'AbortError') return;
    if (window.self !== window.top) {
      $('semSuporte').innerHTML = 'Esta página está aberta dentro de outra (ex.: visualização do claude.ai) e o navegador ' +
        'não deixa escolher pastas assim. Baixe o <b>index.html</b>, <b>app.js</b> e <b>style.css</b> para a mesma pasta ' +
        'e abra o <b>index.html</b> direto no Chrome ou Edge.';
      $('semSuporte').classList.remove('oculto');
    }
    log('Erro ao abrir pasta: ' + e.message, 'erro');
  }
}

async function reconectar() {
  const dir = await idbGet('pasta').catch(() => null);
  if (!dir) return escolherPasta();
  const perm = await dir.requestPermission({ mode: 'readwrite' });
  if (perm === 'granted') await conectar(dir);
  else log('Permissão negada para a pasta.', 'erro');
}

// ---------------------------------------------------------------------------
// Renderização
// ---------------------------------------------------------------------------
let botoesTravados = false;
function travarBotoes(v) { botoesTravados = v; renderizar(); }

function renderizar() {
  renderizarStatus();
  renderizarBotoes();
  if (!estado.editando || !$('corpoTabela').dataset.modoEdicao) renderizarTabela();
}

function renderizarStatus() {
  const d = estado.dados;
  $('infoVersao').textContent = estado.dadosCorrompidos
    ? 'CORROMPIDO'
    : d ? `versão ${d.meta.versao} — por ${d.meta.atualizadoPor} em ${formatarData(d.meta.atualizadoEm)}` : '—';

  const pill = $('statusLock');
  const obs = estado.observacaoLock;
  if (!estado.dir) { pill.className = 'pill neutro'; pill.textContent = '—'; return; }
  if (estado.meuLock) {
    pill.className = 'pill meu';
    pill.textContent = `Você está ${estado.meuLock.operacao}`;
  } else if (estado.editando) {
    pill.className = 'pill erro';
    pill.textContent = 'Lock perdido — tente salvar';
  } else if (obs) {
    let lock = null;
    try { lock = JSON.parse(obs.raw); } catch (_) {}
    const vivo = lockEstaVivo();
    const restante = Math.max(0, Math.ceil((CFG.LOCK_TTL_MS - (Date.now() - obs.vistoEm)) / 1000));
    pill.className = vivo ? 'pill outro' : 'pill erro';
    pill.textContent = lock
      ? `${lock.usuario} está ${lock.operacao} desde ${formatarData(lock.desde)}` +
        (vivo ? '' : ' (sem sinal — lock abandonado)')
      : 'Lock ilegível';
    pill.title = vivo ? `Expira em ${restante}s se não houver sinal` : '';
  } else {
    pill.className = 'pill livre';
    pill.textContent = 'Livre';
    pill.title = '';
  }
}

function renderizarBotoes() {
  const conectado = !!estado.dir;
  const ocupadoPorOutro = !estado.meuLock && !!estado.observacaoLock && lockEstaVivo();
  $('btnEditar').disabled = botoesTravados || !conectado || estado.editando || ocupadoPorOutro || estado.dadosCorrompidos;
  $('btnAdicionar').disabled = botoesTravados || !estado.editando;
  $('btnSalvar').disabled = botoesTravados || !estado.editando;
  $('btnCancelar').disabled = botoesTravados || !estado.editando;
  $('btnForcar').disabled = botoesTravados || !conectado || !!estado.meuLock || !estado.observacaoLock;
  $('btnRestaurar').classList.toggle('oculto', !estado.dadosCorrompidos);
}

const STATUS = ['Aberto', 'Em andamento', 'Concluído', 'Cancelado'];

function renderizarTabela() {
  const tbody = $('corpoTabela');
  tbody.innerHTML = '';
  tbody.dataset.modoEdicao = estado.editando ? '1' : '';

  const registros = estado.editando ? estado.rascunho : (estado.dados?.registros ?? []);
  if (!estado.dir) {
    tbody.innerHTML = '<tr><td colspan="5" class="muted centro">Selecione a pasta da rede para começar.</td></tr>';
    return;
  }
  if (!registros.length) {
    tbody.innerHTML = `<tr><td colspan="5" class="muted centro">Nenhum registro.${estado.editando ? ' Clique em "Novo registro".' : ''}</td></tr>`;
    return;
  }

  for (const reg of registros) {
    const tr = document.createElement('tr');
    if (estado.editando) {
      tr.className = 'editando';
      tr.append(
        celulaTexto(reg.id),
        celulaInput(reg, 'titulo', 'input'),
        celulaInput(reg, 'descricao', 'textarea'),
        celulaSelect(reg, 'status', STATUS),
        celulaExcluir(reg),
      );
    } else {
      tr.append(
        celulaTexto(reg.id),
        celulaTexto(reg.titulo),
        celulaTexto(reg.descricao),
        celulaTexto(reg.status),
        celulaTexto(''),
      );
    }
    tbody.append(tr);
  }
}

function celulaTexto(txt) {
  const td = document.createElement('td');
  td.textContent = txt ?? '';
  return td;
}
function celulaInput(reg, campo, tag) {
  const td = document.createElement('td');
  const el = document.createElement(tag);
  el.value = reg[campo] ?? '';
  el.addEventListener('input', () => { reg[campo] = el.value; });
  td.append(el);
  return td;
}
function celulaSelect(reg, campo, opcoes) {
  const td = document.createElement('td');
  const sel = document.createElement('select');
  for (const o of opcoes) sel.add(new Option(o, o, false, reg[campo] === o));
  sel.addEventListener('change', () => { reg[campo] = sel.value; });
  td.append(sel);
  return td;
}
function celulaExcluir(reg) {
  const td = document.createElement('td');
  const b = document.createElement('button');
  b.textContent = '🗑';
  b.title = 'Excluir';
  b.addEventListener('click', () => {
    estado.rascunho = estado.rascunho.filter(r => r !== reg);
    renderizarTabela();
  });
  td.append(b);
  return td;
}

function adicionarRegistro() {
  const proximoId = estado.rascunho.reduce((m, r) => Math.max(m, Number(r.id) || 0), 0) + 1;
  estado.rascunho.push({ id: proximoId, titulo: '', descricao: '', status: STATUS[0] });
  renderizarTabela();
  const inputs = $('corpoTabela').querySelectorAll('tr:last-child input');
  if (inputs[0]) inputs[0].focus();
}

// ---------------------------------------------------------------------------
// Inicialização
// ---------------------------------------------------------------------------
async function iniciar() {
  if (!('showDirectoryPicker' in window)) {
    $('semSuporte').classList.remove('oculto');
    $('btnEscolherPasta').disabled = true;
    return;
  }

  $('nomeUsuario').value = localStorage.getItem('json-rede-usuario') || '';
  $('nomeUsuario').addEventListener('change', () => {
    try { localStorage.setItem('json-rede-usuario', $('nomeUsuario').value.trim()); } catch (_) {}
  });

  $('btnEscolherPasta').addEventListener('click', escolherPasta);
  $('btnReconectar').addEventListener('click', reconectar);
  $('btnEditar').addEventListener('click', iniciarEdicao);
  $('btnAdicionar').addEventListener('click', adicionarRegistro);
  $('btnSalvar').addEventListener('click', salvar);
  $('btnCancelar').addEventListener('click', cancelar);
  $('btnForcar').addEventListener('click', forcarLiberacao);
  $('btnRestaurar').addEventListener('click', restaurarBackup);

  // Avisa antes de fechar com edição aberta. Se fechar mesmo assim, o lock
  // expira sozinho depois de LOCK_TTL_MS (o heartbeat para).
  window.addEventListener('beforeunload', ev => {
    if (estado.editando) { ev.preventDefault(); ev.returnValue = ''; }
  });
  window.addEventListener('pagehide', () => {
    if (estado.meuLock) liberarLock(); // melhor esforço
  });

  // Já houve uma pasta antes? Oferece reconectar (a permissão exige um clique).
  try {
    const anterior = await idbGet('pasta');
    if (anterior) {
      const perm = await anterior.queryPermission({ mode: 'readwrite' });
      if (perm === 'granted') await conectar(anterior);
      else {
        $('btnReconectar').classList.remove('oculto');
        $('btnReconectar').textContent = `🔄 Reconectar a "${anterior.name}"`;
      }
    }
  } catch (_) {}

  renderizar();
}

iniciar();
