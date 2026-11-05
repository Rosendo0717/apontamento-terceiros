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
         'VOLTA ALMOÇO', 'SAÍDA', 'TOTAL', 'ORDEM', 'SETOR', 'SOLICITANTE', 'DESCRIÇÃO SERVIÇOS', 'STATUS DA ORDEM', 'OBSERVAÇÃO',
         'ENCARREGADO', 'MOTIVO DO REENVIO'],
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
    cc: extrasDe_(p.empresa),
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
    l.status || '', l.obs || '', p.encarregado || '', p.motivo || '']);
  if (linhas.length) base.getRange(base.getLastRow() + 1, 1, linhas.length, linhas[0].length).setValues(linhas);

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
