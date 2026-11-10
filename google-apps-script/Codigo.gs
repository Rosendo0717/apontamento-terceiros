/**
 * Apontamento de Terceiros — servidor (Google Apps Script)
 *
 * O que faz:
 *  - Recebe o apontamento do formulário, envia por e-mail com a planilha anexa,
 *    guarda o .xlsx no Google Drive e grava cada linha na planilha BASE (histórico).
 *  - Guarda o cadastro central de empresas e funcionários (só o administrador altera, com a senha mestre).
 *  - Não envia duas vezes o mesmo envio (o formulário pode tentar de novo quando volta a internet).
 *
 * IMPORTANTE: troque a SENHA_MESTRE abaixo SOMENTE aqui no Apps Script.
 * Não coloque a senha no GitHub (lá o código é público).
 */

const DESTINO = 'pcm@linsagro.com.br';       // para onde vão os e-mails (vários: separe por vírgula)
const CHAVE = 'apont-terceiros-2026';           // igual à do formulário
const SENHA_MESTRE = 'TROQUE_AQUI';             // senha da área do administrador (troque antes de implantar)
const PASTA = 'Apontamentos Terceiros';

// E-mails a mais SÓ para uma empresa (vão em cópia, além do DESTINO).
// Use o nome da empresa exatamente como está no cadastro. Vários e-mails: separe por vírgula.
const EMAILS_EXTRAS = {
  // 'Montec': 'fiscal@exemplo.com.br',
  // 'Jato Dois Irmãos': 'pessoa1@exemplo.com.br, pessoa2@exemplo.com.br',
};         // pasta no seu Google Drive

const ABAS = {
  REGISTRO: ['RECEBIDO EM', 'ID ENVIO', 'EMPRESA', 'DATA', 'ENVIO Nº', 'SITUAÇÃO', 'MOTIVO DO REENVIO', 'ENCARREGADO', 'PESSOAS', 'ARQUIVO', 'ANEXOS'],
  BASE: ['ID ENVIO', 'RECEBIDO EM', 'ENVIO Nº', 'SITUAÇÃO', 'EMPRESA', 'DATA', 'MATRÍCULA', 'COLABORADOR', 'ENTRADA', 'SAÍDA ALMOÇO',
         'VOLTA ALMOÇO', 'SAÍDA', 'TOTAL', 'ORDEM', 'SETOR', 'RESPONSÁVEL DO SETOR', 'DESCRIÇÃO SERVIÇOS', 'STATUS DA ORDEM', 'OBSERVAÇÃO',
         'ENCARREGADO', 'MOTIVO DO REENVIO', 'EQUIPE', 'INÍCIO ESPERA LT', 'FIM ESPERA LT', 'TEMPO ESPERA LT'],
  SETORES: ['SETOR', 'E-MAILS (CÓPIA)'],
  RASCUNHOS: ['EMPRESA', 'DATA', 'APARELHO', 'ID DO APARELHO', 'ATUALIZADO', 'TELA', 'ARQUIVO'],
  EMPRESAS: ['EMPRESA', 'LINK (#)'],
  FUNCIONARIOS: ['EMPRESA', 'COLABORADOR', 'MATRÍCULA']
};

/* ---------------- entradas ---------------- */

function doGet(e) {
  const acao = e && e.parameter && e.parameter.acao;
  if (acao === 'cadastro') {
    if (e.parameter.chave !== CHAVE) return resp({ ok: false, erro: 'chave inválida' });
    return resp(Object.assign({ ok: true }, lerCadastro_()));
  }
  return resp({ ok: true, status: 'Apontamento de Terceiros no ar' });
}

function doPost(e) {
  let p;
  try { p = JSON.parse(e.postData.contents); } catch (err) { return resp({ ok: false, erro: 'pedido inválido' }); }
  if (p.chave !== CHAVE) return resp({ ok: false, erro: 'chave inválida' });

  if (p.acao === 'login') return resp(senhaOk_(p.senha) ? { ok: true } : { ok: false, erro: 'senha incorreta' });

  // rascunho do dia: para continuar o preenchimento em outro aparelho (sem as fotos)
  if (p.acao === 'rascunho') return comTrava_(() => salvarRascunho_(p));
  if (p.acao === 'buscarRascunho') return resp({ ok: true, rascunho: buscarRascunho_(p.empresa, p.data) });
  if (p.acao === 'apagarRascunho') return comTrava_(() => { apagarRascunho_(p.empresa, p.data); return { ok: true }; });

  if (p.acao === 'setores') {
    if (!senhaOk_(p.senha)) return resp({ ok: false, erro: 'senha incorreta' });
    return resp({ ok: true, setores: lerSetores_() });
  }
  if (p.acao === 'salvarSetores') {
    if (!senhaOk_(p.senha)) return resp({ ok: false, erro: 'senha incorreta' });
    return comTrava_(() => { salvarSetores_(p.setores || {}); return { ok: true }; });
  }
  if (p.acao === 'relatorio') {
    if (!senhaOk_(p.senha)) return resp({ ok: false, erro: 'senha incorreta' });
    return resp({ ok: true, linhas: relatorio_(p.empresa, p.de, p.ate) });
  }

  if (p.acao === 'salvarCadastro') {
    if (!senhaOk_(p.senha)) return resp({ ok: false, erro: 'senha incorreta' });
    return comTrava_(() => { salvarCadastro_(p.cadastro); return { ok: true }; });
  }

  return comTrava_(() => receberEnvio_(p));
}

/* ---------------- envio do apontamento ---------------- */

function receberEnvio_(p) {
  if (!p.id || !p.arquivo || !p.nomeArquivo) return { ok: false, erro: 'envio incompleto' };
  const ss = planilha_();
  const reg = ss.getSheetByName('REGISTRO');

  // já recebido? (o formulário reenviou depois de cair a internet) → não manda de novo
  const ids = reg.getLastRow() > 1 ? reg.getRange(2, 2, reg.getLastRow() - 1, 1).getValues().map(r => r[0]) : [];
  if (ids.indexOf(p.id) >= 0) return { ok: true, repetido: true };

  const xlsx = Utilities.newBlob(Utilities.base64Decode(p.arquivo),
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', p.nomeArquivo);

  const extras = (p.anexos || []).map(a => Utilities.newBlob(Utilities.base64Decode(a.b64), a.tipo || 'application/octet-stream', a.nome || 'anexo'));

  MailApp.sendEmail({
    to: DESTINO,
    cc: copias_(p),
    subject: p.assunto || ('APONTAMENTO | ' + p.empresa + ' | ' + p.data),
    body: p.corpo || '',
    attachments: [xlsx].concat(extras),
    name: 'Apontamento de Terceiros'
  });

  const arquivo = pasta_().createFile(xlsx);
  let linksAnexos = '';
  if (extras.length) {
    const sub = subpasta_(pasta_(), 'Anexos');
    linksAnexos = extras.map(b => { b.setName(p.empresa + ' ' + String(p.data).replace(/\//g, '-') + ' - ' + b.getName()); return sub.createFile(b).getUrl(); }).join('\n');
  }
  const agora = new Date();

  // envios anteriores da mesma empresa + data passam a "SUBSTITUÍDO"
  marcarSubstituidos_(reg, 3, 4, 6, p.empresa, p.data);
  const base = ss.getSheetByName('BASE');
  marcarSubstituidos_(base, 5, 6, 4, p.empresa, p.data);

  reg.appendRow([agora, p.id, p.empresa, "'" + p.data, p.envio || 1, 'VÁLIDO', p.motivo || '', p.encarregado || '', p.pessoas || '', arquivo.getUrl(), linksAnexos]);

  const linhas = (p.linhas || []).map(l => [p.id, agora, p.envio || 1, 'VÁLIDO', p.empresa, "'" + p.data, l.mat ? "'" + l.mat : '', l.nome,
    l.entrada, l.saidaAlm || '', l.voltaAlm || '', l.saida, l.total, l.ordem, l.setor || '', l.solicitante || '', l.desc || '',
    l.status || '', l.obs || '', p.encarregado || '', p.motivo || '', l.equipe || '', l.ltIni || '', l.ltFim || '', l.ltTempo || '']);
  if (linhas.length) base.getRange(base.getLastRow() + 1, 1, linhas.length, linhas[0].length).setValues(linhas);

  // o dia foi enviado: o rascunho não serve mais
  const m = String(p.data || '').match(/(\d{2})\/(\d{2})\/(\d{4})/);
  if (m) apagarRascunho_(p.empresa, m[3] + '-' + m[2] + '-' + m[1]);

  return { ok: true };
}

function marcarSubstituidos_(aba, colEmp, colData, colSit, empresa, data) {
  const n = aba.getLastRow() - 1;
  if (n < 1) return;
  const vals = aba.getRange(2, 1, n, aba.getLastColumn()).getDisplayValues();
  vals.forEach((r, i) => {
    if (r[colEmp - 1] === empresa && r[colData - 1] === data && r[colSit - 1] === 'VÁLIDO') {
      aba.getRange(i + 2, colSit).setValue('SUBSTITUÍDO');
    }
  });
}

/* ---------------- cadastro central ---------------- */

function lerCadastro_() {
  const ss = planilha_();
  const emp = ss.getSheetByName('EMPRESAS'), fun = ss.getSheetByName('FUNCIONARIOS');
  const empresas = emp.getLastRow() > 1
    ? emp.getRange(2, 1, emp.getLastRow() - 1, 2).getDisplayValues().filter(r => r[0]).map(r => ({ nome: r[0], slug: r[1] }))
    : [];
  const funcionarios = {};
  empresas.forEach(e => funcionarios[e.nome] = []);
  if (fun.getLastRow() > 1) {
    fun.getRange(2, 1, fun.getLastRow() - 1, 3).getDisplayValues().forEach(r => {
      if (r[0] && r[1]) (funcionarios[r[0]] = funcionarios[r[0]] || []).push({ nome: r[1], mat: r[2] });
    });
  }
  return { empresas: empresas, funcionarios: funcionarios };
}

function salvarCadastro_(c) {
  const ss = planilha_();
  const emp = ss.getSheetByName('EMPRESAS'), fun = ss.getSheetByName('FUNCIONARIOS');
  [emp, fun].forEach(a => { if (a.getLastRow() > 1) a.getRange(2, 1, a.getLastRow() - 1, a.getLastColumn()).clearContent(); });
  const e = (c.empresas || []).map(x => [x.nome, x.slug || '']);
  if (e.length) emp.getRange(2, 1, e.length, 2).setValues(e);
  const f = [];
  Object.keys(c.funcionarios || {}).forEach(n => (c.funcionarios[n] || []).forEach(x => f.push([n, x.nome, x.mat || ''])));
  if (f.length) {
    fun.getRange(2, 3, f.length, 1).setNumberFormat('@'); // matrícula como texto
    fun.getRange(2, 1, f.length, 3).setValues(f);
  }
}

/* ---------------- apoio ---------------- */

/* ---------------- e-mails em cópia: por empresa (código) + por setor (área do administrador) ---------------- */

function copias_(p) {
  const lista = extrasDe_(p.empresa).split(',');
  const porSetor = lerSetores_();
  const setores = {};
  (p.linhas || []).forEach(l => { if (l.setor) setores[String(l.setor).trim().toLowerCase()] = true; });
  Object.keys(porSetor).forEach(s => { if (setores[s.trim().toLowerCase()]) lista.push.apply(lista, String(porSetor[s]).split(/[,;\s]+/)); });
  const vistos = {};
  return lista.map(e => e.trim()).filter(e => /@/.test(e) && !vistos[e.toLowerCase()] && (vistos[e.toLowerCase()] = true)).join(',');
}

function lerSetores_() {
  const aba = planilha_().getSheetByName('SETORES'), out = {};
  if (aba.getLastRow() > 1) aba.getRange(2, 1, aba.getLastRow() - 1, 2).getDisplayValues().forEach(r => { if (r[0]) out[r[0]] = r[1]; });
  return out;
}

function salvarSetores_(m) {
  const aba = planilha_().getSheetByName('SETORES');
  if (aba.getLastRow() > 1) aba.getRange(2, 1, aba.getLastRow() - 1, 2).clearContent();
  const linhas = Object.keys(m).filter(s => String(m[s] || '').trim()).map(s => [s, String(m[s]).trim()]);
  if (linhas.length) aba.getRange(2, 1, linhas.length, 2).setValues(linhas);
}

/* ---------------- rascunhos (continuar em outro aparelho) ---------------- */

function linhaRascunho_(aba, empresa, data) {
  if (aba.getLastRow() < 2) return 0;
  const v = aba.getRange(2, 1, aba.getLastRow() - 1, 2).getDisplayValues();
  for (let i = 0; i < v.length; i++) if (v[i][0] === empresa && v[i][1] === data) return i + 2;
  return 0;
}

function salvarRascunho_(p) {
  if (!p.empresa || !/^\d{4}-\d{2}-\d{2}$/.test(p.data || '') || !p.conteudo) return { ok: false, erro: 'rascunho incompleto' };
  const aba = planilha_().getSheetByName('RASCUNHOS'), n = linhaRascunho_(aba, p.empresa, p.data);
  if (n) {
    const r = aba.getRange(n, 1, 1, 7).getDisplayValues()[0];
    // outro aparelho salvou depois da última versão que este aparelho conhece: avisa em vez de sobrescrever
    if (!p.forcar && r[3] !== p.aparelho && r[4] > String(p.base || '')) {
      return { ok: false, conflito: true, rascunho: { aparelhoNome: r[2], aparelho: r[3], atualizado: r[4], passo: r[5], conteudo: lerArquivo_(r[6]) } };
    }
  }
  const agora = new Date().toISOString();
  const pasta = subpasta_(pasta_(), 'Rascunhos'), nome = p.empresa + ' ' + p.data + '.json';
  let id = n ? aba.getRange(n, 7).getDisplayValue() : '';
  let arq = null;
  if (id) { try { arq = DriveApp.getFileById(id); arq.setContent(p.conteudo); } catch (e) { arq = null; } }
  if (!arq) arq = pasta.createFile(nome, p.conteudo, 'application/json');
  const linha = [p.empresa, "'" + p.data, p.aparelhoNome || '', p.aparelho || '', "'" + agora, p.passo || '', arq.getId()];
  if (n) aba.getRange(n, 1, 1, 7).setValues([linha]); else aba.appendRow(linha);
  return { ok: true, atualizado: agora };
}

function buscarRascunho_(empresa, data) {
  const aba = planilha_().getSheetByName('RASCUNHOS'), n = linhaRascunho_(aba, empresa, data);
  if (!n) return null;
  const r = aba.getRange(n, 1, 1, 7).getDisplayValues()[0];
  const conteudo = lerArquivo_(r[6]);
  return conteudo ? { aparelhoNome: r[2], aparelho: r[3], atualizado: r[4], passo: r[5], conteudo: conteudo } : null;
}

function apagarRascunho_(empresa, data) {
  const aba = planilha_().getSheetByName('RASCUNHOS'), n = linhaRascunho_(aba, empresa, data);
  if (!n) return;
  const id = aba.getRange(n, 7).getDisplayValue();
  try { if (id) DriveApp.getFileById(id).setTrashed(true); } catch (e) {}
  aba.deleteRow(n);
}

function lerArquivo_(id) { try { return id ? DriveApp.getFileById(id).getBlob().getDataAsString('UTF-8') : ''; } catch (e) { return ''; } }

/* ---------------- planilha do período (área do administrador) ---------------- */

function relatorio_(empresa, de, ate) {
  const iso = s => { const m = String(s || '').match(/(\d{2})\/(\d{2})\/(\d{4})/); return m ? m[3] + '-' + m[2] + '-' + m[1] : ''; };
  const a = iso(de), b = iso(ate), base = planilha_().getSheetByName('BASE');
  if (base.getLastRow() < 2) return [];
  const n = ABAS.BASE.length;
  return base.getRange(2, 1, base.getLastRow() - 1, n).getDisplayValues()
    .filter(r => r[3] === 'VÁLIDO' && r[4] === empresa && iso(r[5]) >= a && iso(r[5]) <= b)
    .map(r => ({ envio: r[2], data: r[5], mat: r[6], nome: r[7], entrada: r[8], saidaAlm: r[9], voltaAlm: r[10], saida: r[11], total: r[12],
      ordem: r[13], setor: r[14], solicitante: r[15], desc: r[16], status: r[17], obs: r[18], encarregado: r[19],
      equipe: r[21], ltIni: r[22], ltFim: r[23], ltTempo: r[24] }));
}

function extrasDe_(empresa) {
  const k = Object.keys(EMAILS_EXTRAS).find(n => n.trim().toLowerCase() === String(empresa || '').trim().toLowerCase());
  return k ? String(EMAILS_EXTRAS[k]).split(/[,;\s]+/).filter(e => /@/.test(e)).join(',') : '';
}

function senhaOk_(s) { return SENHA_MESTRE !== 'TROQUE_AQUI' && s === SENHA_MESTRE; }

function comTrava_(fn) {
  const trava = LockService.getScriptLock();
  try { trava.waitLock(20000); return resp(fn()); }
  catch (err) { return resp({ ok: false, erro: String(err) }); }
  finally { try { trava.releaseLock(); } catch (e) {} }
}

function pasta_() {
  const it = DriveApp.getFoldersByName(PASTA);
  return it.hasNext() ? it.next() : DriveApp.createFolder(PASTA);
}

function subpasta_(pai, nome) {
  const it = pai.getFoldersByName(nome);
  return it.hasNext() ? it.next() : pai.createFolder(nome);
}

function planilha_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('BASE_ID');
  let ss = null;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) {
    ss = SpreadsheetApp.create('Base - Apontamento Terceiros');
    DriveApp.getFileById(ss.getId()).moveTo(pasta_());
    props.setProperty('BASE_ID', ss.getId());
  }
  Object.keys(ABAS).forEach(nome => {
    let aba = ss.getSheetByName(nome);
    if (!aba) {
      aba = ss.insertSheet(nome);
      aba.appendRow(ABAS[nome]);
      aba.setFrozenRows(1);
      aba.getRange(1, 1, 1, ABAS[nome].length).setFontWeight('bold').setBackground('#1d6a48').setFontColor('#ffffff');
    } else if (aba.getLastColumn() < ABAS[nome].length) { // coluna nova numa versão mais recente
      aba.getRange(1, 1, 1, ABAS[nome].length).setValues([ABAS[nome]]).setFontWeight('bold').setBackground('#1d6a48').setFontColor('#ffffff');
    }
  });
  const sobra = ss.getSheetByName('Página1') || ss.getSheetByName('Sheet1');
  if (sobra && ss.getSheets().length > 1) ss.deleteSheet(sobra);
  return ss;
}

function resp(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Rode uma vez pelo editor: autoriza Gmail/Drive/Planilhas e cria a planilha "Base - Apontamento Terceiros"
function testarAutorizacao() {
  const ss = planilha_();
  Logger.log('Planilha: ' + ss.getUrl());
  Logger.log('E-mails restantes hoje: ' + MailApp.getRemainingDailyQuota());
  if (SENHA_MESTRE === 'TROQUE_AQUI') Logger.log('ATENÇÃO: troque a SENHA_MESTRE no topo do código.');
}
