/*
 * Kunduz Diamond CRM — Google Apps Script API, version 4.
 *
 * Installation:
 * 1. Open Extensions > Apps Script from the Google spreadsheet used by the CRM.
 * 2. Replace the old server code with this file's contents and save.
 * 3. Deploy > Manage deployments > Edit the web app deployment > select New version > Deploy.
 * 4. Keep the existing /exec URL in the CRM's Google Sheets settings.
 *
 * The KD_ sheet names deliberately avoid modifying any existing sheets. Each
 * row has readable columns plus a JSON column used to round-trip every field.
 */
var API_VERSION = 4;
var SHEETS = {
  clients: {
    name: 'KD_Клиенты',
    headers: ['ID', 'Добавлена', 'Обновлена', 'Имя', 'Контакт', 'Ниша', 'Стаж', 'О компании',
      'Расходы на рекламу', 'Средний чек', 'Ответственный', 'Напомнить', 'Текст напоминания',
      'Реклама запускалась', 'Этап', 'Ярлыки', 'Данные JSON']
  },
  tasks: {
    name: 'KD_Дела',
    headers: ['ID', 'Создана', 'Обновлена', 'Задача', 'Описание', 'Ответственный', 'Срок',
      'Статус', 'ID клиента', 'Клиент', 'Данные JSON']
  },
  accounting: {
    name: 'KD_Бухгалтерия',
    headers: ['ID', 'Создана', 'Обновлена', 'Тип', 'Категория', 'Дата', 'Сумма',
      'Валюта', 'Описание', 'Данные JSON']
  }
};

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('Пустой запрос.');
    }
    var request = JSON.parse(e.postData.contents);
    var result;
    if (request.action === 'list') {
      result = listAll_();
    } else if (request.action === 'save') {
      result = saveAll_(request);
    } else {
      throw new Error('Неизвестное действие API.');
    }
    result.result = 'success';
    result.version = API_VERSION;
    return json_(result);
  } catch (error) {
    return json_({ result: 'error', version: API_VERSION, message: error.message || String(error) });
  }
}

function doGet() {
  return json_({ result: 'success', version: API_VERSION, name: 'Kunduz Diamond CRM API' });
}

function listAll_() {
  var spreadsheet = spreadsheet_();
  var response = { clients: [], tasks: [], accounting: [] };
  Object.keys(SHEETS).forEach(function(key) {
    response[key] = readRecords_(getSheet_(spreadsheet, SHEETS[key]));
  });
  return response;
}

function saveAll_(request) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var spreadsheet = spreadsheet_();
    saveRecords_(getSheet_(spreadsheet, SHEETS.clients), request.clients, request.deletes);
    saveRecords_(getSheet_(spreadsheet, SHEETS.tasks), request.tasks, request.taskDeletes);
    saveRecords_(getSheet_(spreadsheet, SHEETS.accounting), request.accounting, request.accountingDeletes);
    return {};
  } finally {
    lock.releaseLock();
  }
}

function spreadsheet_() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error('Скрипт должен быть создан из Google Таблицы через Расширения > Apps Script.');
  }
  return spreadsheet;
}

function getSheet_(spreadsheet, config) {
  var sheet = spreadsheet.getSheetByName(config.name);
  if (!sheet) sheet = spreadsheet.insertSheet(config.name);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, config.headers.length).setValues([config.headers]);
    sheet.setFrozenRows(1);
  } else {
    var actual = sheet.getRange(1, 1, 1, config.headers.length).getValues()[0];
    if (config.headers.some(function(header, index) { return actual[index] !== header; })) {
      throw new Error('Структура листа «' + config.name + '» не совпадает с версией 4.');
    }
  }
  return sheet;
}

function readRecords_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  var jsonColumn = sheet.getLastColumn() - 1;
  return rows.filter(function(row) { return row[0] && row[jsonColumn]; }).map(function(row, index) {
    try {
      var record = JSON.parse(row[jsonColumn]);
      if (!record.id) record.id = String(row[0]);
      return record;
    } catch (error) {
      throw new Error('Не удалось прочитать JSON в листе «' + sheet.getName() + '», строка ' + (index + 2) + '.');
    }
  });
}

function saveRecords_(sheet, incoming, deleteIds) {
  incoming = Array.isArray(incoming) ? incoming : [];
  deleteIds = Array.isArray(deleteIds) ? deleteIds.map(String) : [];
  var records = readRecords_(sheet);
  var byId = {};
  records.forEach(function(record) { byId[String(record.id)] = record; });
  deleteIds.forEach(function(id) { delete byId[id]; });

  incoming.forEach(function(record) {
    if (!record || !record.id) throw new Error('Операция содержит запись без ID.');
    if (sheet.getName() === SHEETS.accounting.name) validateAccounting_(record);
    byId[String(record.id)] = record;
  });

  var nextRecords = Object.keys(byId).map(function(id) { return byId[id]; });
  var rows = nextRecords.map(function(record) { return rowFor_(sheet.getName(), record); });
  var oldRows = Math.max(0, sheet.getLastRow() - 1);
  if (oldRows) sheet.getRange(2, 1, oldRows, sheet.getLastColumn()).clearContent();
  if (rows.length) sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function rowFor_(sheetName, record) {
  var value = function(key) { return record[key] == null ? '' : record[key]; };
  if (sheetName === SHEETS.clients.name) {
    return [value('id'), value('createdAt'), value('updatedAt'), value('name'), value('contact'),
      value('niche'), value('years'), value('about'), value('adSpend'), value('avgCheck'),
      value('owner'), value('remindAt'), value('remindText'), value('adsRun'), value('stage'),
      JSON.stringify(record.labels || []), JSON.stringify(record)];
  }
  if (sheetName === SHEETS.tasks.name) {
    return [value('id'), value('createdAt'), value('updatedAt'), value('title'), value('description'),
      value('assignee'), value('due'), value('status'), value('clientId'), value('clientName'),
      JSON.stringify(record)];
  }
  return [value('id'), value('createdAt'), value('updatedAt'),
    value('type') === 'income' ? 'Доход' : 'Расход', value('category'), value('date'),
    Number(value('amount')), value('currency') === 'USD' ? 'USD ($)' : 'KGS (сом)',
    value('note'), JSON.stringify(record)];
}

function validateAccounting_(record) {
  var amount = Number(record.amount);
  if (record.type !== 'income' && record.type !== 'expense') {
    throw new Error('У бухгалтерской операции должен быть тип «Доход» или «Расход».');
  }
  if (!isFinite(amount) || amount <= 0) {
    throw new Error('Сумма бухгалтерской операции должна быть больше нуля.');
  }
  if (record.currency !== 'KGS' && record.currency !== 'USD') {
    throw new Error('Поддерживаются только валюты KGS и USD.');
  }
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(record.date || ''))) {
    throw new Error('Укажите дату операции в формате ГГГГ-ММ-ДД.');
  }
}

function json_(value) {
  return ContentService.createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}
