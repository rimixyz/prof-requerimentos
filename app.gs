/**
 * Central de dados dos Requerimentos dos Professores
 * --------------------------------------------------
 * Este Google Apps Script recebe, de maneira independente, uma cópia completa
 * de cada formulário enviado pelo index.html.
 *
 * Estrutura criada automaticamente na planilha:
 *
 * 1) REQUERIMENTOS
 *    Uma linha por envio, com os principais campos já prontos para dashboard.
 *
 * 2) CAMPOS
 *    Formato normalizado: uma linha para cada campo de cada requerimento.
 *    Essa aba é ótima para consultas, tabelas dinâmicas e Looker Studio.
 *
 * 3) EVENTOS_RAW
 *    Guarda o JSON integral recebido e o status do processamento. Serve como
 *    auditoria e permite recuperar informações mesmo que o dashboard mude.
 *
 * 4) ERROS
 *    Registra falhas internas do Apps Script sem interferir no site/fórum.
 *
 * IMPORTANTE PARA A INSTALAÇÃO
 * ---------------------------
 * - A forma mais simples é criar este script DENTRO da planilha desejada em:
 *   Extensões > Apps Script.
 * - Se preferir um projeto separado, preencha SPREADSHEET_ID abaixo.
 * - Depois, implante como "Aplicativo da Web", executando como proprietário e
 *   permitindo acesso a qualquer pessoa que tenha o link.
 * - Copie a URL terminada em /exec para CONFIG.appsScriptUrl no index.html.
 */

const APP_CONFIG = Object.freeze({
  // Deixe vazio quando o script estiver vinculado diretamente à planilha.
  SPREADSHEET_ID: '',

  SCHEMA_VERSION: 1,
  LOCK_TIMEOUT_MS: 30000,
  MAX_CELL_LENGTH: 45000,

  SHEETS: Object.freeze({
    REQUESTS: 'REQUERIMENTOS',
    FIELDS: 'CAMPOS',
    RAW_EVENTS: 'EVENTOS_RAW',
    ERRORS: 'ERROS'
  })
});

const REQUEST_HEADERS = Object.freeze([
  'event_id',
  'schema_version',
  'recebido_em',
  'enviado_em',
  'form_id',
  'tipo_requerimento',
  'autor_forum',
  'nicknames',
  'cargo_atual',
  'novo_cargo',
  'motivo',
  'permissao',
  'dias',
  'data_requerimento',
  'enviar_mp',
  'grupos_internos',
  'pagina_origem',
  'forum_origem',
  'total_campos',
  'bloco_rascunho'
]);

const FIELD_HEADERS = Object.freeze([
  'event_id',
  'recebido_em',
  'form_id',
  'tipo_requerimento',
  'autor_forum',
  'ordem',
  'chave',
  'name',
  'id',
  'rotulo',
  'tipo_campo',
  'valor_exibido',
  'valor_bruto',
  'marcado',
  'oculto',
  'somente_mp',
  'bloco_rascunho'
]);

const RAW_HEADERS = Object.freeze([
  'event_id',
  'status',
  'recebido_em',
  'processado_em',
  'form_id',
  'tipo_requerimento',
  'autor_forum',
  'payload_json',
  'erro'
]);

const ERROR_HEADERS = Object.freeze([
  'data',
  'event_id',
  'form_id',
  'mensagem',
  'stack',
  'payload_parcial'
]);

/**
 * Rota de diagnóstico. Abrir a URL /exec no navegador deve retornar este JSON.
 */
function doGet() {
  return jsonResponse_({
    ok: true,
    service: 'RCC Professores - Requerimentos',
    schemaVersion: APP_CONFIG.SCHEMA_VERSION,
    timestamp: new Date().toISOString()
  });
}

/**
 * Ponto de entrada dos eventos enviados pelo index.html.
 *
 * A trava impede que dois envios simultâneos misturem linhas. O event_id torna
 * o processamento idempotente: o mesmo evento concluído não é gravado de novo.
 */
function doPost(e) {
  const receivedAt = new Date();
  const lock = LockService.getScriptLock();
  let payload = null;
  let eventId = '';
  let rawSheet = null;
  let rawRow = 0;

  try {
    payload = parsePayload_(e);
    validatePayload_(payload);
    eventId = String(payload.eventId);

    lock.waitLock(APP_CONFIG.LOCK_TIMEOUT_MS);

    const spreadsheet = getSpreadsheet_();
    const requestSheet = ensureSheet_(spreadsheet, APP_CONFIG.SHEETS.REQUESTS, REQUEST_HEADERS);
    const fieldSheet = ensureSheet_(spreadsheet, APP_CONFIG.SHEETS.FIELDS, FIELD_HEADERS);
    rawSheet = ensureSheet_(spreadsheet, APP_CONFIG.SHEETS.RAW_EVENTS, RAW_HEADERS);
    ensureSheet_(spreadsheet, APP_CONFIG.SHEETS.ERRORS, ERROR_HEADERS);

    rawRow = findEventRow_(rawSheet, eventId);
    if (rawRow && String(rawSheet.getRange(rawRow, 2).getValue()) === 'SUCESSO') {
      return jsonResponse_({ ok: true, duplicate: true, eventId: eventId });
    }

    // Em uma repetição após erro, limpamos qualquer gravação parcial anterior.
    removeRowsByEventId_(requestSheet, eventId);
    removeRowsByEventId_(fieldSheet, eventId);

    rawRow = upsertRawEvent_(rawSheet, rawRow, payload, receivedAt, 'PROCESSANDO', '');

    // Um lote do modo rascunho é uma única postagem no fórum, porém cada
    // bloco continua sendo uma linha independente para facilitar dashboards.
    const normalizedPayloads = expandDraftPayload_(payload);
    normalizedPayloads.forEach(function(normalizedPayload) {
      appendRequest_(requestSheet, normalizedPayload, receivedAt);
      appendFields_(fieldSheet, normalizedPayload, receivedAt);
    });

    rawSheet.getRange(rawRow, 1, 1, RAW_HEADERS.length).setValues([rawEventRow_(
      payload,
      receivedAt,
      'SUCESSO',
      ''
    )]);

    SpreadsheetApp.flush();

    return jsonResponse_({
      ok: true,
      duplicate: false,
      eventId: eventId,
      fieldsSaved: normalizedPayloads.reduce(function(total, item) {
        return total + (Array.isArray(item.fields) ? item.fields.length : 0);
      }, 0),
      draftBlocksSaved: payload.draftMode === true ? normalizedPayloads.length : 0
    });
  } catch (error) {
    const message = error && error.message ? error.message : String(error);

    try {
      const spreadsheet = getSpreadsheet_();
      const errorSheet = ensureSheet_(spreadsheet, APP_CONFIG.SHEETS.ERRORS, ERROR_HEADERS);
      appendError_(errorSheet, payload, eventId, receivedAt, error);

      if (rawSheet && rawRow && payload) {
        rawSheet.getRange(rawRow, 1, 1, RAW_HEADERS.length).setValues([rawEventRow_(
          payload,
          receivedAt,
          'ERRO',
          message
        )]);
      }
    } catch (loggingError) {
      console.error('Não foi possível registrar o erro:', loggingError);
    }

    // A resposta de erro é útil para testes manuais. O index.html não depende
    // dela e continuará o fluxo normal mesmo quando o Apps Script falhar.
    return jsonResponse_({ ok: false, eventId: eventId, error: message });
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

/**
 * Converte o lote temporário em registros normalizados sem perder o evento
 * bruto original. Todos os blocos compartilham o event_id da postagem e são
 * diferenciados pela coluna bloco_rascunho.
 */
function expandDraftPayload_(payload) {
  if (payload.draftMode !== true || !Array.isArray(payload.draftItems) || !payload.draftItems.length) {
    return [Object.assign({}, payload, { draftIndex: '' })];
  }

  return payload.draftItems.map(function(item, index) {
    const details = item.details || {};
    return Object.assign({}, payload, {
      formId: item.formId || payload.formId,
      formTitle: item.formTitle || payload.formTitle,
      sentAt: item.capturedAt || payload.sentAt,
      fields: Array.isArray(item.fields) ? item.fields : [],
      summary: {
        nicknames: details.nicknames || '',
        currentRole: details.currentRole || '',
        newRole: details.newRole || '',
        reason: details.reason || '',
        permission: details.permission || '',
        days: details.days || '',
        requestDate: details.date || '',
        sendPrivateMessage: Boolean(item.privateMessage),
        internalGroups: []
      },
      draftIndex: index + 1
    });
  });
}

/** Lê JSON enviado como corpo text/plain ou application/json. */
function parsePayload_(e) {
  if (!e || !e.postData || !e.postData.contents) {
    throw new Error('Corpo da requisição ausente.');
  }

  try {
    return JSON.parse(e.postData.contents);
  } catch (error) {
    throw new Error('JSON inválido: ' + error.message);
  }
}

/** Validação mínima para evitar linhas impossíveis de relacionar depois. */
function validatePayload_(payload) {
  if (!payload || typeof payload !== 'object') throw new Error('Payload inválido.');
  if (!payload.eventId) throw new Error('eventId ausente.');
  if (!payload.formId) throw new Error('formId ausente.');
  if (!Array.isArray(payload.fields)) throw new Error('Lista fields ausente ou inválida.');
}

/** Retorna a planilha vinculada ou abre a configurada por ID. */
function getSpreadsheet_() {
  if (APP_CONFIG.SPREADSHEET_ID) {
    return SpreadsheetApp.openById(APP_CONFIG.SPREADSHEET_ID);
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error('Nenhuma planilha vinculada. Informe APP_CONFIG.SPREADSHEET_ID.');
  }
  return spreadsheet;
}

/** Cria uma aba quando necessário e garante seu cabeçalho. */
function ensureSheet_(spreadsheet, sheetName, headers) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) sheet = spreadsheet.insertSheet(sheetName);

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  const currentHeaders = headerRange.getValues()[0];
  const needsHeader = headers.some(function(header, index) {
    return currentHeaders[index] !== header;
  });

  if (needsHeader) {
    headerRange.setValues([headers]);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#821F88');
    headerRange.setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
  }

  return sheet;
}

/** Grava a linha resumida usada pelo dashboard. */
function appendRequest_(sheet, payload, receivedAt) {
  const summary = payload.summary || {};
  const row = [
    payload.eventId,
    payload.schemaVersion || APP_CONFIG.SCHEMA_VERSION,
    receivedAt,
    parseDateOrText_(payload.sentAt),
    payload.formId,
    payload.formTitle || '',
    payload.author || '',
    summary.nicknames || '',
    summary.currentRole || '',
    summary.newRole || '',
    summary.reason || '',
    summary.permission || '',
    summary.days || '',
    summary.requestDate || '',
    summary.sendPrivateMessage === true,
    arrayToText_(summary.internalGroups),
    payload.pageUrl || '',
    payload.forumOrigin || '',
    payload.fields.length,
    payload.draftIndex || ''
  ].map(safeCell_);

  sheet.appendRow(row);
}

/**
 * Grava todos os campos em formato longo/normalizado.
 * Mesmo campos ocultos e exclusivos de MP são mantidos e identificados.
 */
function appendFields_(sheet, payload, receivedAt) {
  if (!payload.fields.length) return;

  const rows = payload.fields.map(function(field, index) {
    return [
      payload.eventId,
      receivedAt,
      payload.formId,
      payload.formTitle || '',
      payload.author || '',
      index + 1,
      field.key || '',
      field.name || '',
      field.id || '',
      field.label || '',
      field.type || '',
      field.displayValue !== undefined ? field.displayValue : '',
      field.rawValue !== undefined ? field.rawValue : '',
      field.checked === true,
      field.hidden === true,
      field.privateMessageOnly === true,
      payload.draftIndex || ''
    ].map(safeCell_);
  });

  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, FIELD_HEADERS.length).setValues(rows);
}

/** Monta a linha do log bruto. */
function rawEventRow_(payload, receivedAt, status, errorMessage) {
  return [
    payload.eventId,
    status,
    receivedAt,
    new Date(),
    payload.formId,
    payload.formTitle || '',
    payload.author || '',
    truncate_(JSON.stringify(payload)),
    errorMessage || ''
  ].map(safeCell_);
}

/** Cria ou atualiza o evento bruto. */
function upsertRawEvent_(sheet, existingRow, payload, receivedAt, status, errorMessage) {
  const row = existingRow || sheet.getLastRow() + 1;
  sheet.getRange(row, 1, 1, RAW_HEADERS.length).setValues([rawEventRow_(
    payload,
    receivedAt,
    status,
    errorMessage
  )]);
  return row;
}

/** Procura event_id na primeira coluna. */
function findEventRow_(sheet, eventId) {
  if (sheet.getLastRow() < 2) return 0;
  const match = sheet
    .getRange(2, 1, sheet.getLastRow() - 1, 1)
    .createTextFinder(String(eventId))
    .matchEntireCell(true)
    .findNext();
  return match ? match.getRow() : 0;
}

/** Remove gravações parciais do mesmo evento antes de uma nova tentativa. */
function removeRowsByEventId_(sheet, eventId) {
  if (sheet.getLastRow() < 2) return;
  const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues();
  for (let index = ids.length - 1; index >= 0; index--) {
    if (ids[index][0] === String(eventId)) sheet.deleteRow(index + 2);
  }
}

/** Registra uma falha sem perder o payload que a originou. */
function appendError_(sheet, payload, eventId, receivedAt, error) {
  sheet.appendRow([
    receivedAt,
    eventId || '',
    payload && payload.formId ? payload.formId : '',
    error && error.message ? error.message : String(error),
    error && error.stack ? truncate_(error.stack) : '',
    payload ? truncate_(JSON.stringify(payload)) : ''
  ].map(safeCell_));
}

/** Evita fórmulas acidentais/injetadas em células vindas de texto livre. */
function safeCell_(value) {
  if (value instanceof Date || typeof value === 'boolean' || typeof value === 'number') {
    return value;
  }

  const text = truncate_(value === null || value === undefined ? '' : String(value));
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function truncate_(value) {
  const text = String(value || '');
  return text.length > APP_CONFIG.MAX_CELL_LENGTH
    ? text.substring(0, APP_CONFIG.MAX_CELL_LENGTH) + '…'
    : text;
}

function arrayToText_(value) {
  if (Array.isArray(value)) return value.join(' / ');
  return value || '';
}

function parseDateOrText_(value) {
  if (!value) return '';
  const date = new Date(value);
  return isNaN(date.getTime()) ? String(value) : date;
}

function jsonResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
