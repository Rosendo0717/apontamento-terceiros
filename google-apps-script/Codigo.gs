/**
 * Apontamento de Terceiros — recebe o apontamento do formulário e envia por e-mail com a planilha anexa.
 * Também guarda o .xlsx numa pasta do Google Drive e registra cada envio numa planilha.
 *
 * Como usar: cole este código em script.google.com, ajuste o DESTINO abaixo e implante como
 * "App da Web" (Executar como: Eu · Quem pode acessar: Qualquer pessoa).
 */

const DESTINO = 'gabriel07rosendo@gmail.com';      // para onde vão os e-mails (pode pôr vários separados por vírgula)
const CHAVE = 'apont-terceiros-2026';              // precisa ser igual à do formulário
const PASTA = 'Apontamentos Terceiros';            // pasta criada no seu Google Drive

function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents);
    if (p.chave !== CHAVE) return resp({ ok: false, erro: 'chave inválida' });
    if (!p.arquivo || !p.nomeArquivo) return resp({ ok: false, erro: 'sem anexo' });

    const xlsx = Utilities.newBlob(
      Utilities.base64Decode(p.arquivo),
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      p.nomeArquivo
    );

    MailApp.sendEmail({
      to: DESTINO,
      subject: p.assunto || ('APONTAMENTO | ' + p.empresa + ' | ' + p.data),
      body: p.corpo || '',
      attachments: [xlsx],
      name: 'Apontamento de Terceiros'
    });

    const pasta = pegarPasta_();
    const salvo = pasta.createFile(xlsx);
    registrar_([new Date(), p.empresa, p.data, p.envio || 1, p.motivo || '', p.encarregado || '', p.pessoas || '', salvo.getUrl()]);

    return resp({ ok: true });
  } catch (err) {
    return resp({ ok: false, erro: String(err) });
  }
}

// Abrir o endereço no navegador mostra se o script está no ar
function doGet() {
  return resp({ ok: true, status: 'Apontamento de Terceiros no ar' });
}

function pegarPasta_() {
  const it = DriveApp.getFoldersByName(PASTA);
  return it.hasNext() ? it.next() : DriveApp.createFolder(PASTA);
}

function registrar_(linha) {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('PLANILHA_ID');
  let ss;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) {
    ss = SpreadsheetApp.create('Registro de envios - Apontamento Terceiros');
    ss.getSheets()[0].appendRow(['RECEBIDO EM', 'EMPRESA', 'DATA', 'ENVIO Nº', 'MOTIVO DO REENVIO', 'ENCARREGADO', 'PESSOAS', 'ARQUIVO']);
    DriveApp.getFileById(ss.getId()).moveTo(pegarPasta_());
    props.setProperty('PLANILHA_ID', ss.getId());
  }
  ss.getSheets()[0].appendRow(linha);
}

function resp(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Rode esta função uma vez pelo editor para autorizar Gmail, Drive e Planilhas
function testarAutorizacao() {
  pegarPasta_();
  Logger.log('Pasta OK. Cota de e-mails restante hoje: ' + MailApp.getRemainingDailyQuota());
}
