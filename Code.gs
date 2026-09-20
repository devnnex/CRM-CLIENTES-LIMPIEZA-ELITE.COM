const SHEET = "clientes";
const COLUMNS = [
  "id",
  "nombre",
  "telefono",
  "ciudad",
  "servicio",
  "frecuencia",
  "valor",
  "fecha",
  "contactoMes",
  "expiracion",
  "direccion",
  "mensajeEnviado",
  "mensajeEnviadoFecha",
  "mensajeCiclo",
  "alertaOculta",
  "alertaOcultaFecha",
  "alertaCiclo",
  "actualizadoEn"
];

const LIST_CACHE_KEY = "limpiezaEliteClientesV3";
const LIST_CACHE_SECONDS = 60;
const SCHEMA_VERSION = "3";

/* =========================
   RESPUESTAS
========================= */
function json(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function errorJson(error) {
  return json({
    status: "error",
    message: error && error.message ? error.message : String(error)
  });
}

/* =========================
   ESQUEMA AUTOMATICO
   Se ejecuta solo en cada entrada y nunca borra ni reordena columnas.
========================= */
function ensureSchema() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) throw new Error("El script debe estar vinculado a una hoja de calculo.");

  let sheet = spreadsheet.getSheetByName(SHEET);
  let schemaChanged = false;
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET);
    schemaChanged = true;
  }

  let headers = [];
  if (sheet.getLastColumn() > 0 && sheet.getLastRow() > 0) {
    headers = sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getDisplayValues()[0]
      .map(value => String(value || "").trim());
  }

  if (!headers.some(Boolean)) {
    headers = COLUMNS.slice();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    schemaChanged = true;
  } else {
    const missing = COLUMNS.filter(column => !headers.includes(column));
    if (missing.length) {
      sheet.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
      headers = headers.concat(missing);
      schemaChanged = true;
    }
  }

  const index = getHeaderIndex(headers);
  const properties = PropertiesService.getScriptProperties();
  if (schemaChanged || properties.getProperty("schemaVersion") !== SCHEMA_VERSION) {
    if (index.telefono !== undefined) {
      sheet.getRange(1, index.telefono + 1, sheet.getMaxRows(), 1).setNumberFormat("@");
    }
    if (index.contactoMes !== undefined) {
      sheet.getRange(1, index.contactoMes + 1, sheet.getMaxRows(), 1).setNumberFormat("@");
    }
    sheet.setFrozenRows(1);
    properties.setProperty("schemaVersion", SCHEMA_VERSION);
  }
  return { sheet, headers, index };
}

function getHeaderIndex(headers) {
  return headers.reduce((map, header, index) => {
    if (header) map[header] = index;
    return map;
  }, {});
}

/* =========================
   NORMALIZACION
========================= */
function normalizeContactoMes(value, fechaFallback) {
  const raw = String(value || "").trim();

  if (/^\d{4}-(0[1-9]|1[0-2])$/.test(raw)) return raw;

  if (/^(0[1-9]|1[0-2])$/.test(raw)) {
    return `${new Date().getFullYear()}-${raw}`;
  }

  const date = fechaFallback ? new Date(fechaFallback) : new Date();
  if (!isNaN(date.getTime())) {
    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
  }

  const now = new Date();
  return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`;
}

function normalizeCycleValue(value) {
  const raw = String(value || "").trim();
  if (!raw) return "sin-fecha";

  const isoDate = raw.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (isoDate) return `${isoDate[1]}-${isoDate[2]}-${isoDate[3]}`;

  const parsed = new Date(raw);
  if (!isNaN(parsed.getTime())) {
    return `${parsed.getFullYear()}-${String(parsed.getMonth() + 1).padStart(2, "0")}-${String(parsed.getDate()).padStart(2, "0")}`;
  }

  return raw.toLowerCase();
}

function getMessageCycle(record) {
  return `${String(record.id || "")}::${normalizeCycleValue(record.expiracion || record.contactoMes)}`;
}

function normalizeBoolean(value) {
  return value === true || value === 1 || ["true", "1", "si", "sí", "yes"].includes(String(value || "").toLowerCase());
}

function valueOr(source, key, fallback) {
  return source[key] !== undefined && source[key] !== null ? source[key] : fallback;
}

function normalizeRecord(data, existing) {
  const current = existing || {};
  const now = new Date().toISOString();
  const fecha = valueOr(data, "fecha", current.fecha || now) || now;
  const contactoMes = normalizeContactoMes(
    valueOr(data, "contactoMes", current.contactoMes || ""),
    fecha
  );

  return {
    id: valueOr(data, "id", current.id || Utilities.getUuid()),
    nombre: valueOr(data, "nombre", current.nombre || ""),
    telefono: String(valueOr(data, "telefono", current.telefono || "")),
    ciudad: valueOr(data, "ciudad", current.ciudad || ""),
    servicio: valueOr(data, "servicio", current.servicio || ""),
    frecuencia: valueOr(data, "frecuencia", current.frecuencia || ""),
    valor: valueOr(data, "valor", current.valor || ""),
    fecha,
    contactoMes,
    expiracion: valueOr(data, "expiracion", current.expiracion || ""),
    direccion: valueOr(data, "direccion", current.direccion || ""),
    mensajeEnviado: normalizeBoolean(valueOr(data, "mensajeEnviado", current.mensajeEnviado || false)),
    mensajeEnviadoFecha: valueOr(data, "mensajeEnviadoFecha", current.mensajeEnviadoFecha || ""),
    mensajeCiclo: valueOr(data, "mensajeCiclo", current.mensajeCiclo || ""),
    alertaOculta: normalizeBoolean(valueOr(data, "alertaOculta", current.alertaOculta || false)),
    alertaOcultaFecha: valueOr(data, "alertaOcultaFecha", current.alertaOcultaFecha || ""),
    alertaCiclo: valueOr(data, "alertaCiclo", current.alertaCiclo || ""),
    actualizadoEn: now
  };
}

/* =========================
   FILAS Y CACHE
========================= */
function serializeCell(value) {
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
    return value.toISOString();
  }
  return value;
}

function rowToObject(headers, row) {
  return headers.reduce((record, header, index) => {
    if (header) record[header] = serializeCell(row[index]);
    return record;
  }, {});
}

function objectToRow(headers, record, currentRow) {
  return headers.map((header, index) => {
    if (Object.prototype.hasOwnProperty.call(record, header)) return record[header];
    return currentRow ? currentRow[index] : "";
  });
}

function getAllClients() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(LIST_CACHE_KEY);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (error) {}
  }

  const context = ensureSchema();
  const lastRow = context.sheet.getLastRow();
  if (lastRow < 2) return [];

  const rows = context.sheet
    .getRange(2, 1, lastRow - 1, context.headers.length)
    .getValues()
    .map(row => rowToObject(context.headers, row))
    .filter(record => String(record.id || "").trim());

  try {
    cache.put(LIST_CACHE_KEY, JSON.stringify(rows), LIST_CACHE_SECONDS);
  } catch (error) {}

  return rows;
}

function invalidateClientsCache() {
  CacheService.getScriptCache().remove(LIST_CACHE_KEY);
}

function findClientRow(context, id) {
  if (!id || context.sheet.getLastRow() < 2 || context.index.id === undefined) return null;

  const range = context.sheet.getRange(
    2,
    context.index.id + 1,
    context.sheet.getLastRow() - 1,
    1
  );
  const cell = range
    .createTextFinder(String(id))
    .matchEntireCell(true)
    .matchCase(true)
    .findNext();

  return cell ? cell.getRow() : null;
}

function readClientAtRow(context, rowNumber) {
  const row = context.sheet
    .getRange(rowNumber, 1, 1, context.headers.length)
    .getValues()[0];
  return { row, record: rowToObject(context.headers, row) };
}

function writeClientAtRow(context, rowNumber, record, currentRow) {
  context.sheet
    .getRange(rowNumber, 1, 1, context.headers.length)
    .setValues([objectToRow(context.headers, record, currentRow)]);
}

/* =========================
   FILTROS OPCIONALES DEL BACKEND
========================= */
function normalizeSearch(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .trim();
}

function getStatus(expiration) {
  const date = new Date(expiration);
  if (isNaN(date.getTime())) return "active";
  const days = Math.ceil((date.getTime() - Date.now()) / 86400000);
  if (days <= 0) return "expired";
  if (days <= 7) return "soon";
  return "active";
}

function filterClients(rows, parameters) {
  const queryTerms = normalizeSearch(parameters.q).split(/\s+/).filter(Boolean);
  const city = normalizeSearch(parameters.city);
  const month = String(parameters.contactoMes || "").trim();
  const status = String(parameters.state || "").trim();
  const contact = String(parameters.contact || "").trim();

  const filtered = rows.filter(record => {
    if (queryTerms.length) {
      const searchable = normalizeSearch([
        record.nombre,
        record.telefono,
        record.ciudad,
        record.direccion,
        record.servicio,
        record.contactoMes
      ].join(" "));
      if (!queryTerms.every(term => searchable.includes(term))) return false;
    }

    if (city && normalizeSearch(record.ciudad) !== city) return false;
    if (month && String(record.contactoMes || "") !== month) return false;
    if (status && getStatus(record.expiracion) !== status) return false;

    const contacted = normalizeBoolean(record.mensajeEnviado) &&
      (!record.mensajeCiclo || record.mensajeCiclo === getMessageCycle(record));
    if (contact === "sent" && !contacted) return false;
    if (contact === "pending" && contacted) return false;

    return true;
  });

  const limit = Math.max(0, Math.min(Number(parameters.limit) || 0, 1000));
  return limit ? filtered.slice(0, limit) : filtered;
}

/* =========================
   GET
========================= */
function doGet(e) {
  try {
    const parameters = e && e.parameter ? e.parameter : {};
    const action = parameters.action || "health";

    if (action === "list") {
      return json(filterClients(getAllClients(), parameters));
    }

    if (action === "health") {
      const context = ensureSchema();
      return json({
        status: "ok",
        sheet: SHEET,
        columns: context.headers,
        version: 3
      });
    }

    return json({ status: "ok", version: 3 });
  } catch (error) {
    return errorJson(error);
  }
}

/* =========================
   POST
========================= */
function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);
    const context = ensureSchema();
    const data = JSON.parse(e.postData.contents || "{}");
    const action = data.action || "create";

    if (action === "create") {
      data.id = data.id || Utilities.getUuid();
      const record = normalizeRecord(data, null);
      const newRow = Math.max(context.sheet.getLastRow() + 1, 2);
      writeClientAtRow(context, newRow, record, null);
      invalidateClientsCache();
      return json({ status: "created", id: record.id, contactoMes: record.contactoMes });
    }

    const rowNumber = findClientRow(context, data.id);
    if (!rowNumber) return json({ status: "not_found" });

    if (action === "delete") {
      context.sheet.deleteRow(rowNumber);
      invalidateClientsCache();
      return json({ status: "deleted", id: data.id });
    }

    const current = readClientAtRow(context, rowNumber);

    if (action === "update") {
      const previousCycle = getMessageCycle(current.record);
      const record = normalizeRecord(data, current.record);
      const nextCycle = getMessageCycle(record);

      if (previousCycle !== nextCycle) {
        record.mensajeEnviado = false;
        record.mensajeEnviadoFecha = "";
        record.mensajeCiclo = "";
        record.alertaOculta = false;
        record.alertaOcultaFecha = "";
        record.alertaCiclo = "";
      }

      writeClientAtRow(context, rowNumber, record, current.row);
      invalidateClientsCache();
      return json({ status: "updated", id: record.id, contactoMes: record.contactoMes });
    }

    if (action === "markContacted") {
      const record = normalizeRecord({}, current.record);
      record.mensajeEnviado = true;
      record.mensajeEnviadoFecha = new Date().toISOString();
      record.mensajeCiclo = data.mensajeCiclo || getMessageCycle(record);
      record.alertaOculta = false;
      record.alertaOcultaFecha = "";
      record.alertaCiclo = "";
      writeClientAtRow(context, rowNumber, record, current.row);
      invalidateClientsCache();
      return json({
        status: "contacted",
        id: record.id,
        mensajeEnviadoFecha: record.mensajeEnviadoFecha,
        mensajeCiclo: record.mensajeCiclo
      });
    }

    if (action === "dismissAlert") {
      const record = normalizeRecord({}, current.record);
      record.alertaOculta = true;
      record.alertaOcultaFecha = new Date().toISOString();
      record.alertaCiclo = data.alertaCiclo || getMessageCycle(record);
      writeClientAtRow(context, rowNumber, record, current.row);
      invalidateClientsCache();
      return json({ status: "alert_dismissed", id: record.id, alertaCiclo: record.alertaCiclo });
    }

    return json({ status: "invalid_action" });
  } catch (error) {
    return errorJson(error);
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}
