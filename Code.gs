const DASHBOARD_SPREADSHEET_ID = '1uGXAkSiv8DOYVwvk9naa42_JAbdimVqvptZQybuOAQE';
const DASHBOARD_TZ = 'Asia/Bangkok';
const DASHBOARD_CACHE_SECONDS = 600;
const DASHBOARD_CACHE_KEY = 'origin-dashboard-v8';
const DASHBOARD_SHEETS = [
  { name: 'AISPM MA Z2', type: 'MA' },
  { name: 'AISPM MA Z17', type: 'MA' },
  { name: 'AISPL MA Z2Z17', type: 'MA' },
  { name: 'AISPM Install Z2', type: 'INSTALL' },
  { name: 'AISPM Install Z17', type: 'INSTALL' },
  { name: 'AISPL FTTB FTTR', type: 'MA' }
];

/**
 * LINE Daily Report
 * ส่งเฉพาะ 5 โซนด้านล่างไปยังกลุ่มชื่อ ORF Report เท่านั้น
 * คีย์ลับทั้งหมดเก็บใน Script Properties ไม่เก็บไว้ใน source code
 */
const LINE_REPORT_GROUP_NAME = 'ORF Report';
const LINE_REPORT_DASHBOARD_URL = 'https://script.google.com/macros/s/AKfycbyzSkhJir63IR7EuXSSspRAsjEtKRU6PESK3qOf7l15WnVMzI6R-wAnFCLeWeEeV9yE/exec';
const LINE_REPORT_ZONES = [
  'AISPM MA Z2',
  'AISPM MA Z17',
  'AISPL MA Z2Z17',
  'AISPM Install Z2',
  'AISPM Install Z17'
];
const LINE_PROP_TOKEN = 'LINE_CHANNEL_ACCESS_TOKEN';
const LINE_PROP_GROUP_ID = 'LINE_REPORT_GROUP_ID';
const LINE_PROP_WEBHOOK_KEY = 'LINE_WEBHOOK_KEY';
const LINE_DAILY_FUNCTION = 'sendDailyLineReport';
const LINE_DAILY_HOUR = 20;

function doGet(e) {
  if (e && e.parameter && e.parameter.api === 'dashboard') {
    return serveDashboardApi_(e);
  }
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Executive Dashboard | Origin Fiber 2026')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/**
 * Public read-only endpoint for the separately hosted dashboard.
 * JSONP is used because Apps Script ContentService cannot attach custom CORS headers.
 */
function serveDashboardApi_(e) {
  const callback = clean_(e.parameter.callback);
  if (!/^[A-Za-z_$][0-9A-Za-z_$\.]{0,127}$/.test(callback)) {
    return ContentService.createTextOutput(JSON.stringify({ error: 'Invalid callback' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  const payload = getDashboardData(false);
  return ContentService.createTextOutput(callback + '(' + JSON.stringify(payload) + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

/**
 * LINE webhook receiver.
 * Apps Script ไม่เปิด HTTP headers ให้ doPost จึงใช้ secret query parameter เพิ่มอีกชั้น
 * และยืนยันชื่อกลุ่มกับ LINE API ก่อนบันทึก groupId เสมอ
 */
function doPost(e) {
  const properties = PropertiesService.getScriptProperties();
  const expectedKey = properties.getProperty(LINE_PROP_WEBHOOK_KEY);
  const receivedKey = e && e.parameter ? clean_(e.parameter.key) : '';
  if (!expectedKey || !receivedKey || !safeEquals_(expectedKey, receivedKey)) {
    return ContentService.createTextOutput('ignored');
  }

  let payload;
  try {
    payload = JSON.parse(e.postData && e.postData.contents ? e.postData.contents : '{}');
  } catch (error) {
    console.error('Invalid LINE webhook JSON: ' + error.message);
    return ContentService.createTextOutput('ok');
  }

  const token = properties.getProperty(LINE_PROP_TOKEN);
  if (!token) return ContentService.createTextOutput('ok');

  (payload.events || []).forEach(function (event) {
    const source = event.source || {};
    if (source.type !== 'group' || !source.groupId) return;

    const text = event.type === 'message' && event.message && event.message.type === 'text'
      ? clean_(event.message.text)
      : '';
    const isSetupMessage = /^#?(setup|ตั้งค่า|เชื่อมต่อ)(\s+orf)?$/i.test(text);
    properties.setProperty('LINE_LAST_GROUP_CANDIDATE_ID', source.groupId);
    properties.setProperty('LINE_LAST_GROUP_CANDIDATE_AT', new Date().toISOString());

    const summary = getLineGroupSummary_(source.groupId, token);
    if (summary && clean_(summary.groupName) !== LINE_REPORT_GROUP_NAME) return;
    if (!summary && !isSetupMessage) return;

    properties.setProperty(LINE_PROP_GROUP_ID, source.groupId);
    properties.setProperty('LINE_REPORT_GROUP_VERIFIED_NAME', LINE_REPORT_GROUP_NAME);
    properties.setProperty('LINE_REPORT_GROUP_VERIFIED_AT', new Date().toISOString());

    if (event.replyToken && isSetupMessage) {
      replyLine_(event.replyToken, '✅ เชื่อมต่อ Daily Report กับกลุ่ม ORF Report สำเร็จ\nเวลาส่งอัตโนมัติ: ทุกวันประมาณ 20:00 น.');
    }
  });

  return ContentService.createTextOutput('ok');
}

/** สร้าง secret สำหรับ webhook และคืน URL ที่ต้องนำไปตั้งใน LINE Developers Console */
function prepareLineWebhook() {
  const properties = PropertiesService.getScriptProperties();
  let key = properties.getProperty(LINE_PROP_WEBHOOK_KEY);
  if (!key) {
    key = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
    properties.setProperty(LINE_PROP_WEBHOOK_KEY, key);
  }
  const webAppUrl = ScriptApp.getService().getUrl() || LINE_REPORT_DASHBOARD_URL;
  return webAppUrl + '?key=' + encodeURIComponent(key);
}

/** สร้าง Trigger ส่งรายงานทุกวันประมาณ 20:00 น. เวลาไทย โดยไม่สร้างซ้ำ */
function setupDailyLineTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (trigger.getHandlerFunction() === LINE_DAILY_FUNCTION) ScriptApp.deleteTrigger(trigger);
  });
  const trigger = ScriptApp.newTrigger(LINE_DAILY_FUNCTION)
    .timeBased()
    .atHour(LINE_DAILY_HOUR)
    .nearMinute(0)
    .everyDays(1)
    .inTimezone(DASHBOARD_TZ)
    .create();
  PropertiesService.getScriptProperties().setProperty('LINE_DAILY_TRIGGER_ID', trigger.getUniqueId());
  return 'ตั้งเวลาส่งรายวันประมาณ 20:00 น. สำเร็จ';
}

/** ฟังก์ชันที่ Time-driven trigger เรียกทุกวัน โดยส่งข้อมูลย้อนหลัง 1 วัน */
function sendDailyLineReport() {
  const reportDate = new Date();
  reportDate.setDate(reportDate.getDate() - 1);
  const dateStr = Utilities.formatDate(reportDate, DASHBOARD_TZ, 'yyyy-MM-dd');
  return sendLineReportForDate(dateStr);
}

/** ใช้ทดสอบข้อมูลตัวอย่างที่ผู้ใช้ตรวจสอบแล้ว */
function sendTestLineReport_17092026() {
  return sendLineReportForDate('2026-09-17');
}

/** ส่งภาพรวม 5 โซนและรายละเอียดทุกโซน */
function sendLineReportForDate(dateInput) {
  const dateStr = normalizeReportDate_(dateInput);
  const properties = PropertiesService.getScriptProperties();
  const token = properties.getProperty(LINE_PROP_TOKEN);
  const groupId = properties.getProperty(LINE_PROP_GROUP_ID);
  if (!token) throw new Error('ยังไม่ได้ตั้ง Script Property: ' + LINE_PROP_TOKEN);
  if (!groupId) throw new Error('ยังไม่พบ groupId ของ ORF Report กรุณาส่งคำว่า #setup ในกลุ่มหลังเปิด webhook');

  // ป้องกันการส่งผิดกลุ่ม แม้มีคนแก้ค่า groupId ใน Script Properties
  const summary = getLineGroupSummary_(groupId, token);
  const verifiedName = properties.getProperty('LINE_REPORT_GROUP_VERIFIED_NAME');
  if ((summary && clean_(summary.groupName) !== LINE_REPORT_GROUP_NAME) ||
      (!summary && verifiedName !== LINE_REPORT_GROUP_NAME)) {
    throw new Error('ยกเลิกการส่ง: groupId ไม่ใช่กลุ่ม ' + LINE_REPORT_GROUP_NAME);
  }

  const rows = readLineReportRows_(dateStr);
  const messages = composeLineReportMessages_(dateStr, rows);
  pushLineMessages_(groupId, messages, token);

  properties.setProperty('LINE_LAST_SENT_DATE', dateStr);
  properties.setProperty('LINE_LAST_SENT_AT', new Date().toISOString());
  properties.setProperty('LINE_LAST_SENT_MESSAGE_COUNT', String(messages.length));
  return { date: dateStr, groupName: summary ? summary.groupName : verifiedName, rows: rows.length, messages: messages.length };
}

/** ตรวจสถานะได้โดยไม่คืนค่า token หรือ groupId เต็ม */
function getLineBotStatus() {
  const p = PropertiesService.getScriptProperties();
  const token = p.getProperty(LINE_PROP_TOKEN);
  const groupId = p.getProperty(LINE_PROP_GROUP_ID);
  let groupName = '';
  if (token && groupId) {
    const summary = getLineGroupSummary_(groupId, token);
    groupName = summary ? summary.groupName : '';
  }
  return {
    tokenConfigured: Boolean(token),
    groupConfigured: Boolean(groupId),
    groupName: groupName,
    groupIdMasked: groupId ? groupId.slice(0, 5) + '…' + groupId.slice(-4) : '',
    schedule: 'ทุกวันประมาณ 20:00 น. (' + DASHBOARD_TZ + ')',
    zones: LINE_REPORT_ZONES.slice(),
    lastSentDate: p.getProperty('LINE_LAST_SENT_DATE') || '',
    lastSentAt: p.getProperty('LINE_LAST_SENT_AT') || ''
  };
}

function readLineReportRows_(dateStr) {
  const ss = SpreadsheetApp.openById(DASHBOARD_SPREADSHEET_ID);
  const configByName = {};
  DASHBOARD_SHEETS.forEach(function (item) { configByName[item.name] = item; });
  const rows = [];
  LINE_REPORT_ZONES.forEach(function (name) {
    const config = configByName[name];
    if (!config) throw new Error('ไม่พบการตั้งค่าโซน ' + name);
    readSheet_(ss.getSheetByName(name), config).forEach(function (row) {
      if (row.DateStr === dateStr) rows.push(row);
    });
  });
  return rows;
}

function composeLineReportMessages_(dateStr, rows) {
  const byZone = {};
  LINE_REPORT_ZONES.forEach(function (zone) { byZone[zone] = []; });
  rows.forEach(function (row) {
    if (byZone[row.Sheet]) byZone[row.Sheet].push(row);
  });

  const overview = buildLineOverview_(dateStr, rows, byZone);
  const zoneMessages = LINE_REPORT_ZONES.map(function (zone) {
    return buildLineZoneDetail_(dateStr, zone, byZone[zone]);
  });

  // LINE รับได้สูงสุด 5 message objects ต่อหนึ่ง push request
  // รวมภาพรวมกับโซนแรกเพื่อให้รายงานปกติส่งได้ครบใน request เดียว
  const first = overview + '\n\n━━━━━━━━━━━━━━━━━━━━\n\n' + zoneMessages.shift();
  const output = splitLineText_(first, 4900);
  zoneMessages.forEach(function (message) {
    Array.prototype.push.apply(output, splitLineText_(message, 4900));
  });
  return output;
}

function buildLineOverview_(dateStr, rows, byZone) {
  const stats = reportStats_(rows);
  const lines = [
    '🟠 ORIGIN FIBER · 5 ZONES',
    'DAILY PERFORMANCE — ' + formatLineDate_(dateStr),
    '',
    'งานทั้งหมด  ' + stats.total + ' งาน',
    '✅ Completed  ' + stats.completed + ' งาน',
    'Success Rate  ' + stats.successRate + '%',
    '🔴 Cancel  ' + stats.cancel + ' งาน',
    '🟡 Postpone  ' + stats.postponed + ' งาน'
  ];
  if (stats.other) lines.push('⚪ สถานะอื่น  ' + stats.other + ' งาน');
  lines.push('', '⚠️ งานที่ต้องติดตาม  ' + stats.followUp + ' งาน');
  lines.push('คิดเป็น ' + stats.followUpRate + '% ของงานทั้งหมด', '', 'ZONE OVERVIEW');
  LINE_REPORT_ZONES.forEach(function (zone) {
    const s = reportStats_(byZone[zone] || []);
    lines.push('', zone, '✅ ' + s.completed + '  🔴 ' + s.cancel + '  🟡 ' + s.postponed + '  | รวม ' + s.total);
  });
  return lines.join('\n');
}

function buildLineZoneDetail_(dateStr, zone, rows) {
  const stats = reportStats_(rows);
  const lines = [
    '🟠 ' + zone,
    'DAILY PERFORMANCE — ' + formatLineDate_(dateStr),
    '',
    'งานทั้งหมด  ' + stats.total + ' งาน',
    '✅ Completed  ' + stats.completed + ' งาน',
    'Success Rate  ' + stats.successRate + '%',
    '🔴 Cancel  ' + stats.cancel + ' งาน',
    '🟡 Postpone  ' + stats.postponed + ' งาน'
  ];
  if (stats.other) lines.push('⚪ สถานะอื่น  ' + stats.other + ' งาน');
  lines.push('', '⚠️ งานที่ต้องติดตาม  ' + stats.followUp + ' งาน');
  lines.push('คิดเป็น ' + stats.followUpRate + '% ของงานทั้งหมด');

  const reasons = countReasons_(rows);
  lines.push('', 'สาเหตุงานไม่สำเร็จ');
  if (!reasons.length) {
    lines.push('– ไม่มีงาน Cancel / Postpone');
  } else {
    reasons.slice(0, 5).forEach(function (item, index) {
      lines.push((index + 1) + '. ' + item.name + '  ' + item.count + ' งาน');
    });
  }

  lines.push('', '────────────────────', 'TEAM PERFORMANCE', '────────────────────');
  const teams = aggregateLineTeams_(rows);
  if (!teams.length) {
    lines.push('', 'ไม่มีข้อมูลทีมช่างในวันนี้');
  } else {
    teams.forEach(function (team) {
      lines.push('', team.name, '✅ ' + team.completed + '  🔴 ' + team.cancel + '  🟡 ' + team.postponed);
      if (team.other) lines.push('⚪ สถานะอื่น ' + team.other);
    });
  }
  lines.push('', '────────────────────', '✅ Completed  🔴 Cancel  🟡 Postpone');
  lines.push('', 'เปิด Dashboard', LINE_REPORT_DASHBOARD_URL);
  return lines.join('\n');
}

function reportStats_(rows) {
  const counts = rows.reduce(function (memo, row) {
    memo[row.Status] = (memo[row.Status] || 0) + 1;
    return memo;
  }, {});
  const total = rows.length;
  const completed = counts.Completed || 0;
  const cancel = counts.Cancel || 0;
  const postponed = counts.Postponed || 0;
  const other = Math.max(0, total - completed - cancel - postponed);
  const followUp = total - completed;
  return {
    total: total,
    completed: completed,
    cancel: cancel,
    postponed: postponed,
    other: other,
    followUp: followUp,
    successRate: total ? (completed * 100 / total).toFixed(1) : '0.0',
    followUpRate: total ? (followUp * 100 / total).toFixed(1) : '0.0'
  };
}

function countReasons_(rows) {
  const counts = {};
  rows.forEach(function (row) {
    if (row.Status === 'Completed' || !row.Reason) return;
    const reason = clean_(row.Reason);
    counts[reason] = (counts[reason] || 0) + 1;
  });
  return Object.keys(counts).map(function (name) {
    return { name: name, count: counts[name] };
  }).sort(function (a, b) {
    return b.count - a.count || a.name.localeCompare(b.name, 'th');
  });
}

function aggregateLineTeams_(rows) {
  const teams = {};
  rows.forEach(function (row) {
    const name = displayEngineerName_(row.Engineer);
    if (!teams[name]) teams[name] = { name: name, total: 0, completed: 0, cancel: 0, postponed: 0, other: 0 };
    const team = teams[name];
    team.total++;
    if (row.Status === 'Completed') team.completed++;
    else if (row.Status === 'Cancel') team.cancel++;
    else if (row.Status === 'Postponed') team.postponed++;
    else team.other++;
  });
  return Object.keys(teams).map(function (name) { return teams[name]; }).sort(function (a, b) {
    return b.total - a.total || b.completed - a.completed || a.name.localeCompare(b.name, 'th');
  });
}

function displayEngineerName_(name) {
  return clean_(name).replace(/\s*\([^)]*[A-Za-z][^)]*\)\s*$/g, '').trim();
}

function splitLineText_(text, maxLength) {
  if (text.length <= maxLength) return [text];
  const lines = text.split('\n');
  const parts = [];
  let current = '';
  lines.forEach(function (line) {
    const candidate = current ? current + '\n' + line : line;
    if (candidate.length > maxLength && current) {
      parts.push(current);
      current = line;
    } else {
      current = candidate;
    }
  });
  if (current) parts.push(current);
  return parts;
}

function pushLineMessages_(to, texts, token) {
  for (let index = 0; index < texts.length; index += 5) {
    const messages = texts.slice(index, index + 5).map(function (text) {
      return { type: 'text', text: text };
    });
    callLineApi_('https://api.line.me/v2/bot/message/push', 'post', token, {
      to: to,
      messages: messages,
      notificationDisabled: false
    });
  }
}

function replyLine_(replyToken, text) {
  const token = PropertiesService.getScriptProperties().getProperty(LINE_PROP_TOKEN);
  if (!token) return;
  callLineApi_('https://api.line.me/v2/bot/message/reply', 'post', token, {
    replyToken: replyToken,
    messages: [{ type: 'text', text: text }]
  });
}

function getLineGroupSummary_(groupId, token) {
  try {
    return callLineApi_('https://api.line.me/v2/bot/group/' + encodeURIComponent(groupId) + '/summary', 'get', token);
  } catch (error) {
    console.error('อ่านชื่อกลุ่ม LINE ไม่สำเร็จ: ' + error.message);
    return null;
  }
}

function callLineApi_(url, method, token, payload) {
  const options = {
    method: method,
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  };
  if (payload) {
    options.contentType = 'application/json';
    options.payload = JSON.stringify(payload);
  }
  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();
  const body = response.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error('LINE API ' + code + ': ' + body.slice(0, 500));
  }
  return body ? JSON.parse(body) : {};
}

function normalizeReportDate_(value) {
  if (value instanceof Date && !isNaN(value)) return Utilities.formatDate(value, DASHBOARD_TZ, 'yyyy-MM-dd');
  const parsed = parseDate_(value);
  if (!parsed) throw new Error('รูปแบบวันที่ไม่ถูกต้อง: ' + value);
  return Utilities.formatDate(parsed, DASHBOARD_TZ, 'yyyy-MM-dd');
}

function formatLineDate_(dateStr) {
  const parts = dateStr.split('-');
  return Number(parts[2]) + '/' + Number(parts[1]) + '/' + parts[0];
}

function safeEquals_(left, right) {
  left = String(left || '');
  right = String(right || '');
  if (left.length !== right.length) return false;
  let result = 0;
  for (let index = 0; index < left.length; index++) result |= left.charCodeAt(index) ^ right.charCodeAt(index);
  return result === 0;
}

/** จุดเรียกหลักของหน้าเว็บ: อ่านข้อมูลเพียงรอบเดียวและใช้ cache 10 นาที */
function getDashboardData(forceRefresh) {
  if (!forceRefresh) {
    const cached = readChunkedCache_(DASHBOARD_CACHE_KEY);
    if (cached) return packPayload_(cached);
  }

  const lock = LockService.getScriptLock();
  lock.tryLock(8000);
  try {
    if (!forceRefresh) {
      const cached = readChunkedCache_(DASHBOARD_CACHE_KEY);
      if (cached) return packPayload_(cached);
    }
    const ss = SpreadsheetApp.openById(DASHBOARD_SPREADSHEET_ID);
    const rows = [];
    const errors = [];
    DASHBOARD_SHEETS.forEach(function (config) {
      try {
        Array.prototype.push.apply(rows, readSheet_(ss.getSheetByName(config.name), config));
      } catch (error) {
        errors.push(config.name + ': ' + error.message);
      }
    });
    const payload = {
      rows: rows,
      meta: {
        updatedAt: Utilities.formatDate(new Date(), DASHBOARD_TZ, "yyyy-MM-dd'T'HH:mm:ssXXX"),
        sheets: DASHBOARD_SHEETS.map(function (item) { return item.name; }),
        rowCount: rows.length,
        warnings: errors
      }
    };
    writeChunkedCache_(DASHBOARD_CACHE_KEY, payload, DASHBOARD_CACHE_SECONDS);
    return packPayload_(payload);
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

// คงชื่อฟังก์ชันเดิมไว้ เผื่อมีหน้า/สคริปต์อื่นเรียกใช้อยู่
function getAllData(forceRefresh) {
  return getDashboardData(forceRefresh).rows.map(toLegacyOverview_);
}

function getInstallData(forceRefresh) {
  return getDashboardData(forceRefresh).rows
    .filter(function (row) { return row.Type === 'INSTALL'; })
    .map(toLegacyInstall_);
}

function readSheet_(sheet, config) {
  if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) return [];
  const values = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  const headers = values.shift().map(normalizeHeader_);
  const indexes = resolveIndexes_(headers, config.type);
  const output = [];
  values.forEach(function (row) {
    const date = parseDate_(valueAt_(row, indexes.date));
    const engineer = clean_(valueAt_(row, indexes.engineer));
    if (!date || !engineer) return;
    const rawStatus = clean_(valueAt_(row, indexes.status));
    const status = normalizeStatus_(rawStatus);
    const reason = clean_(valueAt_(row, indexes.reason)) || (status === 'Completed' ? '' : rawStatus);
    const cableLen = number_(valueAt_(row, indexes.cable));
    const basePay = config.type === 'INSTALL' && status === 'Completed' ? 1250 : 0;
    const cableSurcharge = basePay && cableLen > 325 ? Math.ceil(cableLen - 325) * 6 : 0;
    const dateStr = Utilities.formatDate(date, DASHBOARD_TZ, 'yyyy-MM-dd');
    output.push({
      Sheet: config.name, Type: config.type, DateStr: dateStr, Timestamp: date.getTime(),
      Day: Number(Utilities.formatDate(date, DASHBOARD_TZ, 'd')), Engineer: engineer,
      Status: status, Reason: reason, Product: clean_(valueAt_(row, indexes.product)),
      FibreID: clean_(valueAt_(row, indexes.fibreId)), PlanDate: dateStr, CableLen: cableLen,
      BasePay: basePay, CableSurcharge: cableSurcharge, TotalPay: basePay + cableSurcharge
    });
  });
  return output;
}

function resolveIndexes_(headers, type) {
  if (type === 'INSTALL') {
    return { date: 8, engineer: 9, status: 10, reason: findHeader_(headers, [/^foa reason$/, /^lmr reason$/, /^ถูกเลื่อน ถูกยกเลิกเนื่องจาก$/, /^ถูกยกเลิกเนื่องจากอะไร$/, /^reject reason$/], -1), product: 0, fibreId: 1, cable: 15 };
  }
  return {
    date: findHeader_(headers, [/time/, /วันที่/], -1),
    engineer: findHeader_(headers, [/field engineer/, /ช่าง/], -1),
    status: findHeader_(headers, [/^status$/], -1),
    reason: findHeader_(headers, [/^foa reason$/, /^lmr reason$/, /^ถูกเลื่อน ถูกยกเลิกเนื่องจาก$/, /^ถูกยกเลิกเนื่องจากอะไร$/, /^reject reason$/], -1),
    product: findHeader_(headers, [/product/, /service/, /ผลิตภัณฑ์/, /ประเภทงาน/], -1),
    fibreId: findHeader_(headers, [/fibre.*id/, /fiber.*id/, /access.*number/, /access.*no/, /หมายเลขวงจร/], type === 'INSTALL' ? 1 : 2),
    cable: findHeader_(headers, [/cable.*(length|meter)/, /(length|ระยะ).*cable/, /ระยะสาย/, /ความยาวสาย/], -1)
  };
}

function findHeader_(headers, patterns, fallback) {
  for (let i = 0; i < patterns.length; i++) {
    const index = headers.findIndex(function (header) { return patterns[i].test(header); });
    if (index >= 0) return index;
  }
  return fallback >= 0 && fallback < headers.length ? fallback : -1;
}

function normalizeStatus_(value) {
  const text = clean_(value).toLowerCase();
  if (/complete|completed|สำเร็จ|ปิดงาน|done/.test(text)) return 'Completed';
  if (/cancel|ยกเลิก/.test(text)) return 'Cancel';
  if (/postpone|เลื่อน|นัดใหม่/.test(text)) return 'Postponed';
  if (/onsite|หน้างาน|เข้าพื้นที่/.test(text)) return 'Onsite';
  if (/assign|รับงาน|มอบหมาย/.test(text)) return 'Assign';
  return text ? 'Other' : 'Assign';
}

function parseDate_(value) {
  if (value instanceof Date && !isNaN(value)) return value;
  const text = clean_(value);
  if (!text) return null;
  let match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (match) return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
  match = text.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/);
  if (match) {
    let year = Number(match[3]);
    if (year > 2400) year -= 543;
    if (year < 100) year += 2000;
    return new Date(year, Number(match[2]) - 1, Number(match[1]));
  }
  const parsed = new Date(text);
  return isNaN(parsed) ? null : parsed;
}

function toLegacyOverview_(row) {
  return { Sheet: row.Sheet, DateStr: row.DateStr, Timestamp: row.Timestamp, Day: row.Day, Engineer: row.Engineer, Status: row.Status, Reason: row.Reason };
}

function toLegacyInstall_(row) {
  return { Sheet: row.Sheet, Product: row.Product, FibreID: row.FibreID, PlanDate: row.PlanDate, Timestamp: row.Timestamp, Engineer: row.Engineer, Status: row.Status, CableLen: row.CableLen, BasePay: row.BasePay, CableSurcharge: row.CableSurcharge, TotalPay: row.TotalPay };
}

function writeChunkedCache_(key, value, seconds) {
  const cache = CacheService.getScriptCache();
  // ใช้ 28,000 ตัวอักษรต่อก้อนเพื่อไม่ให้ข้อความภาษาไทยเกินเพดาน 100 KB/key
  const chunks = JSON.stringify(value).match(/[\s\S]{1,28000}/g) || [''];
  const entries = {};
  chunks.forEach(function (chunk, index) { entries[key + ':' + index] = chunk; });
  cache.putAll(entries, seconds);
  cache.put(key + ':meta', String(chunks.length), seconds);
}

function readChunkedCache_(key) {
  const cache = CacheService.getScriptCache();
  const count = Number(cache.get(key + ':meta'));
  if (!count) return null;
  const keys = Array.from({ length: count }, function (_, index) { return key + ':' + index; });
  const chunks = cache.getAll(keys);
  if (keys.some(function (item) { return !chunks[item]; })) return null;
  try { return JSON.parse(keys.map(function (item) { return chunks[item]; }).join('')); }
  catch (error) { return null; }
}

/** ลดขนาดข้อมูลก่อนส่งผ่าน google.script.run เพื่อไม่ให้ payload ขนาดใหญ่หลุดระหว่างทาง */
function packPayload_(value) {
  const json = JSON.stringify(value);
  const bytes = Utilities.gzip(Utilities.newBlob(json, 'application/json')).getBytes();
  return { encoded: true, payload: Utilities.base64Encode(bytes) };
}

function valueAt_(row, index) { return index >= 0 ? row[index] : ''; }
function clean_(value) { return String(value == null ? '' : value).replace(/\s+/g, ' ').trim(); }
function number_(value) { const result = Number(String(value == null ? '' : value).replace(/[^0-9.\-]/g, '')); return isFinite(result) ? result : 0; }
function normalizeHeader_(value) { return clean_(value).toLowerCase().replace(/[\s_\-]+/g, ' '); }
