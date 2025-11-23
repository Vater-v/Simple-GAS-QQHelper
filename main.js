// =========================================================================================
// I. КОНФИГУРАЦИЯ ПРОЕКТА QQ
// =========================================================================================
const CONFIG = {
  sheet_log: 'QQpoker', 
  sheet_accs: 'Accs',
  sheet_core: 'QQ_CORE',
  sheet_helper: 'QQ_Helper',
  
  accs_row_limit: 75,
  balance_warning_threshold: 45,
  hands_warning_threshold: 150,
  rest_threshold_hours: 7,
  
  HEADER_SYNONYMS: {
    accountStatus: ['статус', 'status'],
    accountSet: ['комплект', 'pc'],
    accountNickname: ['никнейм', 'ник', 'account'],
    brainId: ['brain id', 'bid'],
    date: ['date', 'дата'],
    limits: ['limit', 'лимит'],
    startTime: ['start time', 'время начала сессии'],
    endTime: ['end time', 'время окончания'],
    balanceStart: ['start br', 'начальный баланс'],
    balanceEnd: ['end br', 'конечный'],
    balanceTotal: ['total br', 'total'],
    hands: ['hands', 'руки'],
  }
};

// =========================================================================================
// II. ГЛОБАЛЬНЫЕ ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
// =========================================================================================

function getHeaderIndices_(headerRow) {
  const indices = {};
  const headerMap = new Map(headerRow.map((h, i) => [String(h).trim().toLowerCase(), i]));
  for (const key in CONFIG.HEADER_SYNONYMS) {
      const synonyms = CONFIG.HEADER_SYNONYMS[key];
      for (const synonym of synonyms) {
          if (headerMap.has(synonym.toLowerCase())) {
              indices[key] = headerMap.get(synonym.toLowerCase());
              break;
          }
      }
  }
  return indices;
}

function parseDateTime_(dateInput, timeStr) {
  if (!dateInput || !timeStr || (typeof timeStr !== 'string' && !(timeStr instanceof Date))) return null;
  if (timeStr instanceof Date) return timeStr;
  if (String(timeStr).trim() === '') return null;
  if (!(dateInput instanceof Date)) {
    const dateParts = String(dateInput).split('.');
    if (dateParts.length !== 3) {
      const isoAttempt = new Date(dateInput);
      if(isNaN(isoAttempt.getTime())) return null;
      dateInput = isoAttempt;
    } else {
       const formattedStr = `${dateParts[1]}.${dateParts[0]}.${dateParts[2]} ${timeStr}`;
       const d = new Date(formattedStr);
       return isNaN(d.getTime()) ? null : d;
    }
  }
  const timeParts = timeStr.split(':');
  if (timeParts.length < 2) return null;
  const newDate = new Date(dateInput.getTime());
  newDate.setHours(parseInt(timeParts[0], 10));
  newDate.setMinutes(parseInt(timeParts[1], 10));
  newDate.setSeconds(0);
  newDate.setMilliseconds(0);
  return newDate;
}

function formatDuration_(minutes) {
  if (minutes < 0 || isNaN(minutes)) return "0 м";
  const hours = Math.floor(minutes / 60);
  const mins = Math.round(minutes % 60);
  let result = [];
  if (hours > 0) result.push(`${hours} ч`);
  result.push(`${mins} м`);
  return result.join(' ');
}

// =========================================================================================
// III. ТОЧКА ВХОДА И UI
// =========================================================================================
function onOpen(e) {
  SpreadsheetApp.getUi().createMenu(`QQ Helper`)
    .addItem("Обновить отчёты", "mainUpdate")
    .addToUi();
}

function mainUpdate() {
  const now = new Date();
  console.log(`--- ЗАПУСК ОБНОВЛЕНИЯ (${now.toLocaleString()}) ---`);
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    let coreSheet = ss.getSheetByName(CONFIG.sheet_core);
    if (!coreSheet) coreSheet = ss.insertSheet(CONFIG.sheet_core);
    
    let helperSheet = ss.getSheetByName(CONFIG.sheet_helper);
    if (!helperSheet) helperSheet = ss.insertSheet(CONFIG.sheet_helper);
    
    const accsSheet = ss.getSheetByName(CONFIG.sheet_accs);
    const logSheet = ss.getSheetByName(CONFIG.sheet_log);
    if (!accsSheet || !logSheet) throw new Error("Не найдены листы Accs или QQpoker!");

    const accsHeaders = getHeaderIndices_(accsSheet.getRange(1, 1, 1, accsSheet.getLastColumn()).getValues()[0]);
    const logHeaders = getHeaderIndices_(logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues()[0]);
    
    const accsData = accsSheet.getRange(2, 1, CONFIG.accs_row_limit, accsSheet.getLastColumn()).getValues();
    const logData = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, logSheet.getLastColumn()).getValues();
    
    const logMap = new Map();
    logData.forEach(row => {
      const nick = String(row[logHeaders.accountNickname] || '').trim();
      if (nick) {
        if (!logMap.has(nick)) logMap.set(nick, []);
        logMap.get(nick).push(row);
      }
    });

    const allAccountsData = accsData.map(row => {
      const nick = String(row[accsHeaders.accountNickname] || '').trim();
      const komplekt = String(row[accsHeaders.accountSet] || '').trim();
      const brainId = String(row[accsHeaders.brainId] || '').trim();
      
      const isS59 = (komplekt.toLowerCase() === 's59');
      if (!nick || (!brainId && !isS59)) return null;

      const sessions = logMap.get(nick) || [];
      let latestSession = null;

      sessions.forEach(sessionRow => {
        const sessionDate = parseDateTime_(sessionRow[logHeaders.date], sessionRow[logHeaders.startTime]);
        if (sessionDate && (!latestSession || sessionDate.getTime() > latestSession.startTime.getTime())) {
          let endTime = parseDateTime_(sessionRow[logHeaders.date], sessionRow[logHeaders.endTime]);
          if (endTime && endTime < sessionDate) endTime.setDate(endTime.getDate() + 1);
          latestSession = { row: sessionRow, startTime: sessionDate, endTime: endTime, limit: sessionRow[logHeaders.limits] };
        }
      });
      
      let computedStatus = 'unknown', pauseTime = '', inPlayTime = '', isInGame = false;
      let balanceInfo = { value: null };

      if (latestSession) {
        isInGame = !latestSession.endTime;
        if (isInGame) {
          computedStatus = 'in-play';
          inPlayTime = formatDuration_((now - latestSession.startTime) / 6e4);
        } else {
          const diffHours = (now - latestSession.endTime) / 36e5;
          computedStatus = (diffHours >= CONFIG.rest_threshold_hours) ? 'rested' : 'paused';
          pauseTime = formatDuration_((now - latestSession.endTime) / 6e4);
        }
      }
      
      return {
        nick: nick, komplekt: komplekt, 
        status: String(row[accsHeaders.accountStatus] || '').trim(),
        pauseTime: pauseTime, inPlayTime: inPlayTime, computedStatus: computedStatus, 
        isInGame: isInGame, latestSession: latestSession
      };
    }).filter(Boolean);

    generateDashboard_(helperSheet, allAccountsData, now);
    SpreadsheetApp.getActive().toast('Отчёты QQ успешно обновлены!');

  } catch (e) {
    console.error(`КРИТИЧЕСКАЯ ОШИБКА: ${e.message} \nСтек: ${e.stack}`);
    SpreadsheetApp.getActive().toast(`Произошла ошибка: ${e.message}`);
  }
}

// =========================================================================================
// IV. ГЕНЕРАТОР ДАШБОРДА
// =========================================================================================
function generateDashboard_(helperSheet, allAccountsData, now) {
  helperSheet.clear();
  allAccountsData.sort((a, b) => String(a.komplekt).localeCompare(String(b.komplekt)));
  
  let currentRow = 1;
  
  // --- Таблица 1: Основная ---
  const t1_headers = ["Аккаунт (комплект)", "Статус", "Время паузы", "Статус состояния"];
  const t1_data = allAccountsData.filter(acc => acc.computedStatus !== 'unknown').map(acc => [
    `${acc.nick} (${acc.komplekt})`, acc.status, acc.pauseTime, acc.computedStatus
  ]);
  currentRow = drawTable_(helperSheet, currentRow, 1, "1. Основная таблица", t1_headers, t1_data);

  // --- Таблица 2: Готовы к запуску ---
  const t2_headers = ["Аккаунт (комплект)", "Время паузы"];
  const t2_data = allAccountsData.filter(acc => String(acc.status).toLowerCase() === 'ready' && acc.computedStatus === 'rested').map(acc => [`${acc.nick} (${acc.komplekt})`, acc.pauseTime]);
  currentRow = drawTable_(helperSheet, currentRow, 1, "2. Готовы к запуску", t2_headers, t2_data);

  // --- Таблица 3: В игре ---
  const t3_headers = ["Аккаунт (комплект)", "Время в игре"];
  const t3_data = allAccountsData.filter(acc => acc.isInGame).map(acc => [`${acc.nick} (${acc.komplekt})`, acc.inPlayTime]);
  currentRow = drawTable_(helperSheet, currentRow, 1, `3. В игре`, t3_headers, t3_data);

  // --- Таблица 4: Обратить внимание ---
  const t4_headers = ["Аккаунт (комплект)", "Статус"];
  const t4_data = allAccountsData.filter(acc => String(acc.status).toLowerCase() === 'sla').map(acc => [`${acc.nick} (${acc.komplekt})`, acc.status]);
  drawTable_(helperSheet, currentRow, 1, "4. Обратить внимание", t4_headers, t4_data);
}

function drawTable_(sheet, startRow, startCol, title, headers, data) {
  sheet.getRange(startRow, startCol).setValue(title).setFontWeight('bold');
  startRow++;
  const headerRange = sheet.getRange(startRow, startCol, 1, headers.length);
  headerRange.setValues([headers]).setFontWeight('bold').setBackground('#efefef').setHorizontalAlignment('center');
  startRow++;
  if (data.length > 0) {
    const dataRange = sheet.getRange(startRow, startCol, data.length, headers.length);
    dataRange.setValues(data).setVerticalAlignment('middle').setHorizontalAlignment('center');
    dataRange.offset(0,0, data.length, 1).setHorizontalAlignment('left');
    sheet.getRange(startRow - 1, startCol, data.length + 1, headers.length).setBorder(true, true, true, true, true, true, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);
    return startRow + data.length + 2;
  } else {
    sheet.getRange(startRow, startCol, 1, headers.length).merge().setValue("Нет данных").setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.getRange(startRow - 1, startCol, 2, headers.length).setBorder(true, true, true, true, true, true, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);
    return startRow + 3;
  }
}
