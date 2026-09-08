/**
 * SimpleQuote (Phase 1 MVP) — Google Apps Script Web App
 *
 * Quotes + Convert-to-Invoice + a customizable Quote Request intake
 * form, backed entirely by a Google Sheet. No Customers/Services
 * library yet — that's Phase 2. Every quote/invoice keeps its own
 * customer info directly on the row.
 *
 * PRINTING/PDF: quotes and invoices are printed/saved as PDF using
 * the browser's own Print dialog (via website/print-view.html), not
 * a Google Docs template — no external template setup required.
 * Branding (logo, color, terms language) is configured through
 * website/settings-wizard.html and stored in the Settings tab.
 *
 * NETWORKING: everything — reads AND writes — goes through GET
 * requests using JSONP (a <script> tag, not fetch()). This sidesteps
 * browser CORS restrictions entirely (which normally block reading
 * Apps Script responses) and, unlike a fire-and-forget POST, lets the
 * app show a real confirmation (the assigned Quote/Invoice Number)
 * the moment something saves.
 *
 * SETUP:
 * 1. Create a new Google Sheet with tabs: Quotes, Quote_Items,
 *    Invoices, Invoice_Items, Settings, Quote_Requests (see
 *    sheet-templates/ for ready-made CSVs to import for each tab).
 * 2. Extensions > Apps Script, paste this file in.
 * 3. Fill in the Settings tab (one row of values — see
 *    sheet-templates/settings-template.csv for the columns), or use
 *    website/settings-wizard.html once deployed to do this visually.
 * 4. Deploy > New deployment > Web app. Execute as: Me. Who has
 *    access: Anyone. Copy the /exec URL into website/app.html's
 *    CONFIG.apiUrl (and website/settings-wizard.html,
 *    website/print-view.html, website/request-builder.html's
 *    generated form).
 */

var SHEET_NAMES = {
  CUSTOMERS: 'Customers',
  LEADS: 'Leads',
  BOOKINGS: 'Bookings',
  CATALOG: 'Products_Services',
  QUOTES: 'Quotes',
  QUOTE_ITEMS: 'Quote_Items',
  INVOICES: 'Invoices',
  INVOICE_ITEMS: 'Invoice_Items',
  SETTINGS: 'Settings',
  QUOTE_REQUESTS: 'Quote_Requests'
};

var CALENDAR_ID = 'default';
var BOOKING_WINDOW_DAYS = 14;
var SERVICES = [{
  name: 'Technology Help Visit', durationMinutes: 60, requiresAddress: true,
  availability: [
    { day: 1, start: '09:00', end: '17:00' }, { day: 2, start: '09:00', end: '17:00' },
    { day: 3, start: '09:00', end: '17:00' }, { day: 4, start: '09:00', end: '17:00' },
    { day: 5, start: '09:00', end: '17:00' }
  ]
}];

// ============================================================
// Generic sheet helpers — read/write rows as plain objects using
// the header row as keys. Every tab in this app follows this shape.
// ============================================================

function getSheet_(name) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(name);
  // Be forgiving when Google imports a CSV using hyphens, spaces, or
  // different capitalization. Quote-Items, quote_items, and
  // "Quote Items" all resolve to the canonical Quote_Items tab.
  if (!sheet) {
    var normalizedExpected = normalizeSheetName_(name);
    sheet = spreadsheet.getSheets().filter(function (candidate) {
      return normalizeSheetName_(candidate.getName()) === normalizedExpected;
    })[0] || null;
  }
  if (!sheet) throw new Error('Sheet tab not found: ' + name);
  return sheet;
}

function normalizeSheetName_(name) {
  return String(name || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

function readAllRows_(sheetName) {
  var sheet = getSheet_(sheetName);
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0];
  return values.slice(1).map(function (row) {
    var obj = {};
    headers.forEach(function (h, i) { obj[h] = row[i]; });
    return obj;
  }).filter(function (obj) {
    // Skip fully blank rows
    return Object.keys(obj).some(function (k) { return obj[k] !== '' && obj[k] !== null; });
  });
}

function appendRow_(sheetName, rowObj) {
  var sheet = getSheet_(sheetName);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var row = headers.map(function (h) { return rowObj.hasOwnProperty(h) ? rowObj[h] : ''; });
  sheet.appendRow(row);
}

function findRowIndexById_(sheetName, idColumn, idValue) {
  var sheet = getSheet_(sheetName);
  var values = sheet.getDataRange().getValues();
  var headers = values[0];
  var idColIndex = headers.indexOf(idColumn);
  if (idColIndex === -1) throw new Error('Column not found: ' + idColumn);
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][idColIndex]) === String(idValue)) return i + 1; // 1-based sheet row
  }
  return -1;
}

function updateRowById_(sheetName, idColumn, idValue, updates) {
  var sheet = getSheet_(sheetName);
  var rowIndex = findRowIndexById_(sheetName, idColumn, idValue);
  if (rowIndex === -1) throw new Error('Record not found: ' + idValue);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var currentRow = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
  var newRow = headers.map(function (h, i) {
    return updates.hasOwnProperty(h) ? updates[h] : currentRow[i];
  });
  sheet.getRange(rowIndex, 1, 1, headers.length).setValues([newRow]);
}

function deleteRowById_(sheetName, idColumn, idValue) {
  var sheet = getSheet_(sheetName);
  var rowIndex = findRowIndexById_(sheetName, idColumn, idValue);
  if (rowIndex === -1) return false;
  sheet.deleteRow(rowIndex);
  return true;
}

function deleteRowsByForeignKey_(sheetName, fkColumn, fkValue) {
  var sheet = getSheet_(sheetName);
  var values = sheet.getDataRange().getValues();
  var headers = values[0];
  var fkColIndex = headers.indexOf(fkColumn);
  if (fkColIndex === -1) return;
  // Delete from the bottom up so row indices don't shift under us
  for (var i = values.length - 1; i >= 1; i--) {
    if (String(values[i][fkColIndex]) === String(fkValue)) {
      sheet.deleteRow(i + 1);
    }
  }
}

// ============================================================
// Settings
// ============================================================

function getSettings_() {
  var rows = readAllRows_(SHEET_NAMES.SETTINGS);
  return rows[0] || {};
}

// ============================================================
// ID / Number generation
// Quote_ID / Invoice_ID: internal, unique, never shown to customers.
// Quote_Number / Invoice_Number: human-facing, sequential per year,
// using the configurable prefix from Settings.
// ============================================================

function generateInternalId_(prefix) {
  return prefix + '_' + new Date().getTime() + '_' + Math.floor(Math.random() * 10000);
}

// Uses LockService + PropertiesService so two people saving at the
// exact same moment can never be assigned the same number.
function getNextSequentialNumber_(counterKey) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var props = PropertiesService.getScriptProperties();
    var current = parseInt(props.getProperty(counterKey) || '0', 10);
    var next = current + 1;
    props.setProperty(counterKey, String(next));
    return next;
  } finally {
    lock.releaseLock();
  }
}

function formatDocumentNumber_(prefix, year, sequence) {
  var padded = ('0000' + sequence).slice(-4);
  return prefix + '-' + year + '-' + padded;
}

function generateQuoteNumber_(settings) {
  var year = new Date().getFullYear();
  var prefix = settings.Quote_Prefix || 'Q';
  var seq = getNextSequentialNumber_('quoteCounter_' + year);
  return formatDocumentNumber_(prefix, year, seq);
}

function generateInvoiceNumber_(settings) {
  var year = new Date().getFullYear();
  var prefix = settings.Invoice_Prefix || 'INV';
  var seq = getNextSequentialNumber_('invoiceCounter_' + year);
  return formatDocumentNumber_(prefix, year, seq);
}

// ============================================================
// Calculations — pulled into their own functions so they're easy
// to test and impossible to get subtly wrong in two different places.
// ============================================================

function calculateLineItemAmount_(quantity, rate) {
  return round2_(parseFloat(quantity || 0) * parseFloat(rate || 0));
}

function calculateTotals_(items, discount, taxRate) {
  var subtotal = 0;
  var taxableSubtotal = 0;

  items.forEach(function (item) {
    var amount = calculateLineItemAmount_(item.quantity, item.rate);
    subtotal += amount;
    if (item.taxable) taxableSubtotal += amount;
  });

  discount = parseFloat(discount || 0);
  taxRate = parseFloat(taxRate || 0);

  var tax = round2_(taxableSubtotal * (taxRate / 100));
  var total = round2_(subtotal - discount + tax);

  return {
    subtotal: round2_(subtotal),
    discount: round2_(discount),
    tax: tax,
    total: total
  };
}

function calculateBalanceDue_(total, amountPaid) {
  return round2_(parseFloat(total || 0) - parseFloat(amountPaid || 0));
}

function computeInvoiceStatus_(total, amountPaid, dueDate, currentStatus) {
  if (currentStatus === 'Void') return 'Void';
  var balance = calculateBalanceDue_(total, amountPaid);
  if (total > 0 && balance <= 0) return 'Paid';
  if (parseFloat(amountPaid || 0) > 0) return 'Partially Paid';
  if (dueDate && new Date(dueDate) < new Date(new Date().toDateString())) return 'Overdue';
  return currentStatus || 'Draft';
}

function round2_(n) {
  return Math.round((n + Number.EPSILON) * 100) / 100;
}

function addDays_(date, days) {
  var d = new Date(date);
  d.setDate(d.getDate() + days);
  return d;
}

// Leads and bookings imported from the existing Neighbor Tech tools.
function normalizedFields_(fields) {
  var normalized = {};
  Object.keys(fields || {}).forEach(function (key) {
    normalized[String(key).trim().toLowerCase()] = fields[key];
  });
  return normalized;
}

function appendFormRow_(sheetName, fields, systemValues) {
  var sheet = getSheet_(sheetName);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var normalized = normalizedFields_(fields);
  var row = headers.map(function (header) {
    var key = String(header).trim().toLowerCase();
    if (systemValues.hasOwnProperty(key)) return systemValues[key];
    return normalized.hasOwnProperty(key) ? normalized[key] : '';
  });
  sheet.appendRow(row);
}

function listLeads_() {
  return readAllRows_(SHEET_NAMES.LEADS).sort(function (a, b) { return new Date(b.Timestamp) - new Date(a.Timestamp); });
}

function listBookings_() {
  return readAllRows_(SHEET_NAMES.BOOKINGS).sort(function (a, b) { return new Date(b.Timestamp) - new Date(a.Timestamp); });
}

function normalizeEmail_(value) { return String(value || '').trim().toLowerCase(); }
function normalizePhone_(value) { return String(value || '').replace(/\D/g, '').slice(-10); }

function listCustomers_() {
  return readAllRows_(SHEET_NAMES.CUSTOMERS).filter(function (customer) {
    return String(customer.Active).toLowerCase() !== 'false';
  }).sort(function (a, b) { return String(a.Name || '').localeCompare(String(b.Name || '')); });
}

function findCustomer_(email, phone) {
  var normalizedEmail = normalizeEmail_(email);
  var normalizedPhone = normalizePhone_(phone);
  return listCustomers_().filter(function (customer) {
    return (normalizedEmail && normalizeEmail_(customer.Email) === normalizedEmail) ||
      (normalizedPhone && normalizePhone_(customer.Phone) === normalizedPhone);
  })[0] || null;
}

function saveCustomer_(payload) {
  var now = new Date();
  var isNew = !payload.id;
  var id = isNew ? generateInternalId_('CUST') : payload.id;
  var existing = isNew ? null : listCustomers_().filter(function (customer) { return customer.Customer_ID === id; })[0];
  var row = {
    Customer_ID: id, Name: String(payload.name || '').trim(), Company: String(payload.company || '').trim(),
    Email: String(payload.email || '').trim(), Phone: String(payload.phone || '').trim(), Address: String(payload.address || '').trim(),
    Source: String(payload.source || (existing ? existing.Source : 'Manual')), Notes: String(payload.notes || ''),
    Created_At: existing ? existing.Created_At : now, Updated_At: now, Active: true
  };
  if (!row.Name) throw new Error('Customer name is required.');
  if (isNew) {
    var match = findCustomer_(row.Email, row.Phone);
    if (match) return match;
    appendRow_(SHEET_NAMES.CUSTOMERS, row);
  } else updateRowById_(SHEET_NAMES.CUSTOMERS, 'Customer_ID', id, row);
  return row;
}

function findOrCreateCustomer_(fields, source) {
  var match = findCustomer_(fields.Email || fields.email, fields.Phone || fields.phone);
  if (match) return match;
  return saveCustomer_({
    name: fields.Name || fields.name || 'Unknown customer', company: fields['Business Name'] || '',
    email: fields.Email || fields.email || '', phone: fields.Phone || fields.phone || '',
    address: fields.Address || fields.address || '', source: source, notes: ''
  });
}

function getCustomerHistory_(customerId) {
  var customer = listCustomers_().filter(function (item) { return item.Customer_ID === customerId; })[0];
  if (!customer) throw new Error('Customer not found.');
  function related(sheetName) { return readAllRows_(sheetName).filter(function (row) { return row.Customer_ID === customerId; }); }
  return { customer: customer, leads: related(SHEET_NAMES.LEADS), bookings: related(SHEET_NAMES.BOOKINGS), quotes: related(SHEET_NAMES.QUOTES), invoices: related(SHEET_NAMES.INVOICES) };
}

function savePhoneInquiry_(payload) {
  var customer = saveCustomer_(payload.customer || {});
  var leadId = generateInternalId_('LEAD');
  appendRow_(SHEET_NAMES.LEADS, {
    Lead_ID: leadId, Customer_ID: customer.Customer_ID, Timestamp: new Date(), Name: customer.Name,
    'Business Name': customer.Company, Email: customer.Email, Phone: customer.Phone,
    'Business Type': '', 'Service Interested': payload.service || '', 'Preferred Contact': payload.preferredContact || 'Phone',
    Message: payload.notes || '', 'Source Page': 'Phone call', Status: 'New', 'Follow-up Date': payload.followUpDate || ''
  });
  return { customer: customer, leadId: leadId };
}

function listCatalog_() {
  return readAllRows_(SHEET_NAMES.CATALOG).filter(function (item) { return String(item.Active).toLowerCase() !== 'false'; });
}

function saveCatalogItem_(payload) {
  var isNew = !payload.id;
  var id = isNew ? generateInternalId_('ITEM') : payload.id;
  var row = {
    Item_ID: id, Name: String(payload.name || '').trim(), Description: String(payload.description || '').trim(),
    Type: payload.type === 'Product' ? 'Product' : 'Service', Unit: String(payload.unit || 'each'),
    Duration_Minutes: Math.max(0, parseInt(payload.durationMinutes || 0, 10)),
    Rate: Math.max(0, parseFloat(payload.rate || 0)), Taxable: !!payload.taxable, Active: true
  };
  if (!row.Name) throw new Error('Item name is required.');
  if (isNew) appendRow_(SHEET_NAMES.CATALOG, row);
  else updateRowById_(SHEET_NAMES.CATALOG, 'Item_ID', id, row);
  return row;
}

function archiveCatalogItem_(id) {
  updateRowById_(SHEET_NAMES.CATALOG, 'Item_ID', id, { Active: false });
  return { archived: true };
}

function createQuoteFromSource_(sourceType, sourceId) {
  var sheetName = sourceType === 'Booking' ? SHEET_NAMES.BOOKINGS : SHEET_NAMES.LEADS;
  var idKey = sourceType === 'Booking' ? 'Booking_ID' : 'Lead_ID';
  var source = readAllRows_(sheetName).filter(function (row) { return String(row[idKey]) === String(sourceId); })[0];
  if (!source) throw new Error(sourceType + ' not found');
  var notes = Object.keys(source).filter(function (key) {
    return ['Timestamp', 'Name', 'Email', 'Phone', 'Status'].indexOf(key) === -1 && source[key] !== '';
  }).map(function (key) { return key + ': ' + source[key]; }).join('\n');
  return saveQuote_({
    customerId: source.Customer_ID || '',
    customerName: source.Name || '', customerEmail: source.Email || '', customerPhone: source.Phone || '',
    sourceType: sourceType, sourceReference: sourceId, status: 'Draft',
    notes: notes, discount: 0, items: []
  });
}

function findService_(name) {
  for (var i = 0; i < SERVICES.length; i++) if (SERVICES[i].name === name) return SERVICES[i];
  return null;
}

function getCalendar_() {
  return CALENDAR_ID === 'default' ? CalendarApp.getDefaultCalendar() : CalendarApp.getCalendarById(CALENDAR_ID);
}

function computeAvailableSlots_(serviceName) {
  var service = findService_(serviceName);
  if (!service) return [];
  var calendar = getCalendar_();
  var slots = [];
  var now = new Date();
  for (var d = 0; d < BOOKING_WINDOW_DAYS; d++) {
    var day = new Date(now.getFullYear(), now.getMonth(), now.getDate() + d);
    var blocks = service.availability.filter(function (b) { return b.day === day.getDay(); });
    var events = calendar.getEventsForDay(day);
    blocks.forEach(function (block) {
      var startParts = block.start.split(':');
      var endParts = block.end.split(':');
      var startMinutes = parseInt(startParts[0], 10) * 60 + parseInt(startParts[1], 10);
      var endMinutes = parseInt(endParts[0], 10) * 60 + parseInt(endParts[1], 10);
      for (var t = startMinutes; t + service.durationMinutes <= endMinutes; t += service.durationMinutes) {
        var start = new Date(day.getFullYear(), day.getMonth(), day.getDate(), 0, t, 0);
        var end = new Date(start.getTime() + service.durationMinutes * 60000);
        if (start < now) continue;
        var conflict = events.some(function (event) { return start < event.getEndTime() && end > event.getStartTime(); });
        if (!conflict) slots.push({ startIso: start.toISOString(), endIso: end.toISOString(), label: Utilities.formatDate(start, Session.getScriptTimeZone(), 'EEE, MMM d - h:mm a') });
      }
    });
  }
  return slots;
}

function createOwnerBooking_(payload) {
  var customer = listCustomers_().filter(function (item) { return item.Customer_ID === payload.customerId; })[0];
  if (!customer) throw new Error('Choose a valid customer.');
  var service = findService_(payload.service);
  if (!service) throw new Error('Choose a valid service.');
  var start = new Date(payload.slotStart);
  var end = new Date(payload.slotEnd);
  if (isNaN(start.getTime()) || isNaN(end.getTime()) || end <= start) throw new Error('Choose a valid appointment time.');
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (getCalendar_().getEvents(start, end).length) throw new Error('That time is no longer available. Refresh the available times.');
    var event = getCalendar_().createEvent(service.name + ' - ' + customer.Name, start, end, {
      description: ['Customer: ' + customer.Name, 'Phone: ' + customer.Phone, 'Email: ' + customer.Email, 'Notes: ' + (payload.notes || '')].join('\n')
    });
    var bookingId = generateInternalId_('BKG');
    appendRow_(SHEET_NAMES.BOOKINGS, {
      Booking_ID: bookingId, Customer_ID: customer.Customer_ID, Timestamp: new Date(), Name: customer.Name,
      Phone: customer.Phone, Email: customer.Email, Service: service.name, 'Appointment Time': start,
      Address: payload.address || customer.Address || '', Notes: payload.notes || '', 'Source Page': 'CRM dashboard',
      Status: 'Confirmed', 'Calendar Event ID': event.getId()
    });
    return { Booking_ID: bookingId, Customer_ID: customer.Customer_ID, Appointment_Time: start };
  } finally { lock.releaseLock(); }
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents || '{}');
    var fields = data.fields || {};
    if (data.service && data.slotStart && data.slotEnd) {
      var service = findService_(data.service);
      if (!service) throw new Error('Unknown service');
      var lock = LockService.getScriptLock();
      lock.waitLock(10000);
      try {
        var start = new Date(data.slotStart);
        var end = new Date(data.slotEnd);
        if (getCalendar_().getEvents(start, end).length) throw new Error('That time was just booked. Please choose another.');
        var title = service.name + ' - ' + (fields.Name || 'Customer');
        var event = getCalendar_().createEvent(title, start, end, { description: Object.keys(fields).map(function (k) { return k + ': ' + fields[k]; }).join('\n') });
        var bookingCustomer = findOrCreateCustomer_(fields, 'Booking form');
        appendFormRow_(SHEET_NAMES.BOOKINGS, fields, {
          booking_id: generateInternalId_('BKG'), customer_id: bookingCustomer.Customer_ID, timestamp: new Date(), status: 'Confirmed', 'source page': data.sourcePage || '',
          service: data.service, 'appointment time': start, 'calendar event id': event.getId()
        });
      } finally { lock.releaseLock(); }
      return jsonOutput_({ result: 'success' });
    }
    var leadCustomer = findOrCreateCustomer_(fields, 'Lead form');
    appendFormRow_(SHEET_NAMES.LEADS, fields, {
      lead_id: generateInternalId_('LEAD'), customer_id: leadCustomer.Customer_ID, timestamp: new Date(), status: 'New', 'follow-up date': '', 'source page': data.sourcePage || ''
    });
    return jsonOutput_({ result: 'success' });
  } catch (error) {
    return jsonOutput_({ result: 'error', message: error.message || error.toString() });
  }
}

function jsonOutput_(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON);
}

function ownerAuthorized_(params) {
  var expected = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
  return !!expected && String(params.adminKey || '') === expected;
}

// ============================================================
// Quote Requests — a customizable public intake form. Fields are
// matched by column header name (same pattern as the lead-capture
// and booking-request tools), so the business owner can add/remove
// fields freely without touching this file.
// ============================================================

function submitQuoteRequest_(formFields, sourcePage) {
  var sheet = getSheet_(SHEET_NAMES.QUOTE_REQUESTS);
  var lastColumn = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];

  var normalizedFields = {};
  for (var key in formFields) {
    normalizedFields[key.toString().trim().toLowerCase()] = formFields[key];
  }

  var row = headers.map(function (header) {
    var headerKey = String(header).trim().toLowerCase();
    if (headerKey === 'request_id') return generateInternalId_('REQ');
    if (headerKey === 'timestamp') return new Date();
    if (headerKey === 'status') return 'New';
    if (headerKey === 'source page') return sourcePage || '';
    return normalizedFields.hasOwnProperty(headerKey) ? normalizedFields[headerKey] : '';
  });

  sheet.appendRow(row);
  return { submitted: true };
}

function listQuoteRequests_() {
  return readAllRows_(SHEET_NAMES.QUOTE_REQUESTS).sort(function (a, b) {
    return new Date(b.Timestamp) - new Date(a.Timestamp);
  });
}

// Best-effort pulls out Name/Email/Phone if the business included
// columns with those names; everything else submitted gets folded
// into the new quote's Notes so nothing is ever lost, regardless of
// what custom fields were configured on the intake form.
function createQuoteFromRequest_(requestId) {
  var requests = readAllRows_(SHEET_NAMES.QUOTE_REQUESTS);
  var request = requests.filter(function (r) { return r.Request_ID === requestId; })[0];
  if (!request) throw new Error('Request not found: ' + requestId);

  var knownKeys = { name: 'customerName', email: 'customerEmail', phone: 'customerPhone' };
  var customerName = '', customerEmail = '', customerPhone = '';
  var noteLines = [];

  Object.keys(request).forEach(function (key) {
    var lower = key.trim().toLowerCase();
    var value = request[key];
    if (lower === 'name') { customerName = value; return; }
    if (lower === 'email') { customerEmail = value; return; }
    if (lower === 'phone') { customerPhone = value; return; }
    if (lower === 'request_id' || lower === 'timestamp' || lower === 'status' || lower === 'source page') return;
    if (value !== '' && value !== null && value !== undefined) {
      noteLines.push(key + ': ' + value);
    }
  });

  var requestCustomer = findOrCreateCustomer_({ Name: customerName, Email: customerEmail, Phone: customerPhone }, 'Quote request');
  var quote = saveQuote_({
    customerId: requestCustomer.Customer_ID,
    customerName: customerName,
    customerEmail: customerEmail,
    customerPhone: customerPhone,
    sourceType: 'Request',
    sourceReference: requestId,
    status: 'Draft',
    discount: 0,
    notes: noteLines.join('\n'),
    items: []
  });

  updateRowById_(SHEET_NAMES.QUOTE_REQUESTS, 'Request_ID', requestId, { Status: 'Converted' });

  return quote;
}

// ============================================================
// Quotes
// ============================================================

function listQuotes_() {
  return readAllRows_(SHEET_NAMES.QUOTES).sort(function (a, b) {
    return new Date(b.Created_At) - new Date(a.Created_At);
  });
}

function getQuoteWithItems_(quoteId) {
  var quotes = readAllRows_(SHEET_NAMES.QUOTES);
  var quote = quotes.filter(function (q) { return q.Quote_ID === quoteId; })[0];
  if (!quote) throw new Error('Quote not found: ' + quoteId);
  var items = readAllRows_(SHEET_NAMES.QUOTE_ITEMS).filter(function (i) { return i.Quote_ID === quoteId; });
  quote.items = items;
  return quote;
}

function saveQuote_(payload) {
  var settings = getSettings_();
  var now = new Date();
  var totals = calculateTotals_(payload.items || [], payload.discount, payload.taxRate != null ? payload.taxRate : settings.Default_Tax_Rate);

  var isNew = !payload.id;
  var quoteId = isNew ? generateInternalId_('Q') : payload.id;
  var quoteNumber = isNew ? generateQuoteNumber_(settings) : payload.quoteNumber;
  var createdDate = isNew ? now : new Date(payload.createdDate || now);
  var expirationDate = payload.expirationDate
    ? new Date(payload.expirationDate)
    : addDays_(now, parseInt(settings.Quote_Expiration_Days || 30, 10));

  var row = {
    Quote_ID: quoteId,
    Customer_ID: payload.customerId || '',
    Quote_Number: quoteNumber,
    Customer_Name: payload.customerName || '',
    Customer_Email: payload.customerEmail || '',
    Customer_Phone: payload.customerPhone || '',
    Source_Type: payload.sourceType || '',
    Source_Reference: payload.sourceReference || '',
    Created_Date: createdDate,
    Expiration_Date: expirationDate,
    Status: payload.status || 'Draft',
    Subtotal: totals.subtotal,
    Discount: totals.discount,
    Tax: totals.tax,
    Total: totals.total,
    Notes: payload.notes || '',
    Terms: payload.terms || '',
    Created_At: isNew ? now : payload.createdAt,
    Updated_At: now
  };

  if (isNew) {
    appendRow_(SHEET_NAMES.QUOTES, row);
  } else {
    updateRowById_(SHEET_NAMES.QUOTES, 'Quote_ID', quoteId, row);
  }

  // Simplest correct way to handle add/edit/reorder/remove of line
  // items: replace the whole set every save.
  deleteRowsByForeignKey_(SHEET_NAMES.QUOTE_ITEMS, 'Quote_ID', quoteId);
  (payload.items || []).forEach(function (item, index) {
    appendRow_(SHEET_NAMES.QUOTE_ITEMS, {
      Item_ID: generateInternalId_('QI'),
      Quote_ID: quoteId,
      Description: item.description || '',
      Quantity: item.quantity || 0,
      Rate: item.rate || 0,
      Amount: calculateLineItemAmount_(item.quantity, item.rate),
      Taxable: !!item.taxable
    });
  });

  return getQuoteWithItems_(quoteId);
}

function deleteQuote_(quoteId) {
  var quotes = readAllRows_(SHEET_NAMES.QUOTES);
  var quote = quotes.filter(function (q) { return q.Quote_ID === quoteId; })[0];
  if (!quote) throw new Error('Quote not found: ' + quoteId);
  if (quote.Status !== 'Draft') {
    throw new Error('Only draft quotes can be deleted. This quote is: ' + quote.Status);
  }
  deleteRowsByForeignKey_(SHEET_NAMES.QUOTE_ITEMS, 'Quote_ID', quoteId);
  deleteRowById_(SHEET_NAMES.QUOTES, 'Quote_ID', quoteId);
  return { deleted: true };
}

// ============================================================
// Convert Quote -> Invoice
// The quote is left completely unchanged. A new invoice is created
// with its own new ID/Number, remaining independently editable
// afterward (actual work performed may differ from the quote).
// ============================================================

function convertQuoteToInvoice_(quoteId) {
  var settings = getSettings_();
  var quote = getQuoteWithItems_(quoteId);
  var now = new Date();

  var invoiceId = generateInternalId_('INV');
  var invoiceNumber = generateInvoiceNumber_(settings);
  var dueDate = addDays_(now, parseInt(settings.Invoice_Due_Days || 14, 10));

  var invoiceRow = {
    Invoice_ID: invoiceId,
    Invoice_Number: invoiceNumber,
    Customer_ID: quote.Customer_ID || '',
    Customer_Name: quote.Customer_Name,
    Customer_Email: quote.Customer_Email,
    Customer_Phone: quote.Customer_Phone,
    Source_Quote_ID: quote.Quote_ID,
    Created_Date: now,
    Due_Date: dueDate,
    Status: 'Draft',
    Subtotal: quote.Subtotal,
    Discount: quote.Discount,
    Tax: quote.Tax,
    Total: quote.Total,
    Amount_Paid: 0,
    Balance_Due: quote.Total,
    Payment_Date: '',
    Payment_Method: '',
    Notes: quote.Notes,
    Payment_Instructions: settings.Payment_Instructions || '',
    Created_At: now,
    Updated_At: now
  };
  appendRow_(SHEET_NAMES.INVOICES, invoiceRow);

  quote.items.forEach(function (item) {
    appendRow_(SHEET_NAMES.INVOICE_ITEMS, {
      Item_ID: generateInternalId_('II'),
      Invoice_ID: invoiceId,
      Description: item.Description,
      Quantity: item.Quantity,
      Rate: item.Rate,
      Amount: item.Amount,
      Taxable: item.Taxable
    });
  });

  return getInvoiceWithItems_(invoiceId);
}

// ============================================================
// Invoices
// ============================================================

function listInvoices_() {
  return readAllRows_(SHEET_NAMES.INVOICES).sort(function (a, b) {
    return new Date(b.Created_At) - new Date(a.Created_At);
  });
}

function getInvoiceWithItems_(invoiceId) {
  var invoices = readAllRows_(SHEET_NAMES.INVOICES);
  var invoice = invoices.filter(function (inv) { return inv.Invoice_ID === invoiceId; })[0];
  if (!invoice) throw new Error('Invoice not found: ' + invoiceId);
  var items = readAllRows_(SHEET_NAMES.INVOICE_ITEMS).filter(function (i) { return i.Invoice_ID === invoiceId; });
  invoice.items = items;
  return invoice;
}

function saveInvoice_(payload) {
  var settings = getSettings_();
  var now = new Date();
  var totals = calculateTotals_(payload.items || [], payload.discount, payload.taxRate != null ? payload.taxRate : settings.Default_Tax_Rate);
  var balanceDue = calculateBalanceDue_(totals.total, payload.amountPaid);

  var isNew = !payload.id;
  var invoiceId = isNew ? generateInternalId_('INV') : payload.id;
  var invoiceNumber = isNew ? generateInvoiceNumber_(settings) : payload.invoiceNumber;
  var createdDate = isNew ? now : new Date(payload.createdDate || now);
  var dueDate = payload.dueDate
    ? new Date(payload.dueDate)
    : addDays_(now, parseInt(settings.Invoice_Due_Days || 14, 10));

  var status = computeInvoiceStatus_(totals.total, payload.amountPaid, dueDate, payload.status);

  var row = {
    Invoice_ID: invoiceId,
    Customer_ID: payload.customerId || '',
    Invoice_Number: invoiceNumber,
    Customer_Name: payload.customerName || '',
    Customer_Email: payload.customerEmail || '',
    Customer_Phone: payload.customerPhone || '',
    Source_Quote_ID: payload.sourceQuoteId || '',
    Created_Date: createdDate,
    Due_Date: dueDate,
    Status: status,
    Subtotal: totals.subtotal,
    Discount: totals.discount,
    Tax: totals.tax,
    Total: totals.total,
    Amount_Paid: parseFloat(payload.amountPaid || 0),
    Balance_Due: balanceDue,
    Payment_Date: payload.paymentDate || '',
    Payment_Method: payload.paymentMethod || '',
    Notes: payload.notes || '',
    Payment_Instructions: payload.paymentInstructions || settings.Payment_Instructions || '',
    Created_At: isNew ? now : payload.createdAt,
    Updated_At: now
  };

  if (isNew) {
    appendRow_(SHEET_NAMES.INVOICES, row);
  } else {
    updateRowById_(SHEET_NAMES.INVOICES, 'Invoice_ID', invoiceId, row);
  }

  deleteRowsByForeignKey_(SHEET_NAMES.INVOICE_ITEMS, 'Invoice_ID', invoiceId);
  (payload.items || []).forEach(function (item) {
    appendRow_(SHEET_NAMES.INVOICE_ITEMS, {
      Item_ID: generateInternalId_('II'),
      Invoice_ID: invoiceId,
      Description: item.description || '',
      Quantity: item.quantity || 0,
      Rate: item.rate || 0,
      Amount: calculateLineItemAmount_(item.quantity, item.rate),
      Taxable: !!item.taxable
    });
  });

  return getInvoiceWithItems_(invoiceId);
}

// ============================================================
// Settings (branding, terms, business info) — read/write.
// Used by the branding wizard and the print view.
// ============================================================

function saveSettings_(payload) {
  var sheet = getSheet_(SHEET_NAMES.SETTINGS);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var existingRows = readAllRows_(SHEET_NAMES.SETTINGS);
  var row = headers.map(function (h) {
    return payload.hasOwnProperty(h) ? payload[h] : (existingRows[0] ? existingRows[0][h] : '');
  });
  if (existingRows.length === 0) {
    sheet.appendRow(row);
  } else {
    sheet.getRange(2, 1, 1, headers.length).setValues([row]);
  }
  return getSettings_();
}

function formatDate_(d) {
  if (!d) return '';
  return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), 'MMM d, yyyy');
}

function formatCurrency_(amount, currency) {
  var symbol = (currency === 'EUR') ? '\u20ac' : (currency === 'GBP') ? '\u00a3' : '$';
  return symbol + Number(amount || 0).toFixed(2);
}

// ============================================================
// Web app entry point — everything (reads AND writes) is a GET
// request answered as JSONP.
// ============================================================

function doGet(e) {
  var params = (e && e.parameter) || {};

  if (!params.callback) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'SimpleQuote endpoint is live' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var callback = params.callback;
  if (!/^[A-Za-z_$][0-9A-Za-z_$\.]*$/.test(callback)) {
    return jsonOutput_({ success: false, error: 'Invalid callback.' });
  }
  var result;

  try {
    var publicAction = params.action === 'services' || params.action === 'availability' || params.action === 'submit_request';
    if (!publicAction && !ownerAuthorized_(params)) throw new Error('Owner authorization required.');
    switch (params.action) {
      case 'list_customers':
        result = { success: true, data: listCustomers_() };
        break;
      case 'get_customer_history':
        result = { success: true, data: getCustomerHistory_(params.id) };
        break;
      case 'save_customer':
        result = { success: true, data: saveCustomer_(JSON.parse(decodeURIComponent(params.data))) };
        break;
      case 'save_phone_inquiry':
        result = { success: true, data: savePhoneInquiry_(JSON.parse(decodeURIComponent(params.data))) };
        break;
      case 'create_owner_booking':
        result = { success: true, data: createOwnerBooking_(JSON.parse(decodeURIComponent(params.data))) };
        break;
      case 'services':
        result = { services: SERVICES.map(function (s) { return { name: s.name, durationMinutes: s.durationMinutes, requiresAddress: !!s.requiresAddress }; }) };
        break;
      case 'availability':
        result = { slots: computeAvailableSlots_(params.service) };
        break;
      case 'list_leads':
        result = { success: true, data: listLeads_() };
        break;
      case 'list_bookings':
        result = { success: true, data: listBookings_() };
        break;
      case 'list_catalog':
        result = { success: true, data: listCatalog_() };
        break;
      case 'save_catalog_item':
        result = { success: true, data: saveCatalogItem_(JSON.parse(decodeURIComponent(params.data))) };
        break;
      case 'archive_catalog_item':
        result = { success: true, data: archiveCatalogItem_(params.id) };
        break;
      case 'create_quote_from_source':
        result = { success: true, data: createQuoteFromSource_(params.sourceType, params.sourceId) };
        break;
      case 'list_quotes':
        result = { success: true, data: listQuotes_() };
        break;
      case 'get_quote':
        result = { success: true, data: getQuoteWithItems_(params.id) };
        break;
      case 'save_quote':
        result = { success: true, data: saveQuote_(JSON.parse(decodeURIComponent(params.data))) };
        break;
      case 'delete_quote':
        result = { success: true, data: deleteQuote_(params.id) };
        break;
      case 'convert_to_invoice':
        result = { success: true, data: convertQuoteToInvoice_(params.quoteId) };
        break;
      case 'list_invoices':
        result = { success: true, data: listInvoices_() };
        break;
      case 'get_invoice':
        result = { success: true, data: getInvoiceWithItems_(params.id) };
        break;
      case 'save_invoice':
        result = { success: true, data: saveInvoice_(JSON.parse(decodeURIComponent(params.data))) };
        break;
      case 'get_settings':
        result = { success: true, data: getSettings_() };
        break;
      case 'save_settings':
        result = { success: true, data: saveSettings_(JSON.parse(decodeURIComponent(params.data))) };
        break;
      case 'submit_request':
        result = { success: true, data: submitQuoteRequest_(JSON.parse(decodeURIComponent(params.data)).fields, params.sourcePage) };
        break;
      case 'list_requests':
        result = { success: true, data: listQuoteRequests_() };
        break;
      case 'create_quote_from_request':
        result = { success: true, data: createQuoteFromRequest_(params.requestId) };
        break;
      default:
        result = { success: false, error: 'Unknown action: ' + params.action };
    }
  } catch (error) {
    result = { success: false, error: error.message || error.toString() };
  }

  return ContentService
    .createTextOutput(callback + '(' + JSON.stringify(result) + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}
