var CONFIG = {
  voteSheetName: '角色投票',               // 留空 = 用第一個分頁
  resultSheetName: '抽獎紀錄',
  defaultWinnerCount: 5,                  // 輸入框留空時的預設人數
  seatHeaders: ['樓層/區域', '排', '號'],  // 必須與投票分頁的標題列文字完全一致
  dedupeBySeat: true,                     // 同座位多筆投票只算一張券
  excludePreviousWinners: true            // 已中獎的座位不再參加後續梯次
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🎁 抽獎')
    .addItem('開始抽獎', 'runLottery')
    .addItem('查看候選名單', 'showCandidates')
    .addSeparator()
    .addItem('檢查資料分頁', 'showSheetInfo')
    .addToUi();
}

function runLottery() {
  var ui = SpreadsheetApp.getUi();
  var candidates;
  try {
    candidates = getCandidates_();
  } catch (e) {
    ui.alert('無法抽獎', e.message, ui.ButtonSet.OK);
    return;
  }

  if (!candidates.length) {
    ui.alert('無法抽獎', '目前沒有符合條件的候選者。', ui.ButtonSet.OK);
    return;
  }

  var resp = ui.prompt(
    '抽獎',
    '候選人數：' + candidates.length + ' 位\n要抽幾位？（直接按確定 = ' + CONFIG.defaultWinnerCount + ' 位）',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  var text = resp.getResponseText().trim();
  var count = text === '' ? CONFIG.defaultWinnerCount : parseInt(text, 10);
  if (!(count > 0)) {
    ui.alert('請輸入大於 0 的數字。', ui.ButtonSet.OK);
    return;
  }
  if (count > candidates.length) {
    ui.alert('人數不足', '候選人只有 ' + candidates.length + ' 位。', ui.ButtonSet.OK);
    return;
  }

  var winners = pickRandom_(candidates, count);
  var round = recordWinners_(winners);

  ui.alert(
    '第 ' + round + ' 梯次中獎名單',
    winners.map(function(w, i) { return (i + 1) + '. ' + w.seat; }).join('\n'),
    ui.ButtonSet.OK
  );
}

function showCandidates() {
  var ui = SpreadsheetApp.getUi();
  try {
    var candidates = getCandidates_();
    ui.alert(
      '候選名單',
      '符合資格：' + candidates.length + ' 位\n\n' +
      (CONFIG.dedupeBySeat ? '・同座位只算一張券' : '・每筆投票各算一張券') + '\n' +
      (CONFIG.excludePreviousWinners ? '・已排除先前中獎的座位' : '・先前中獎的座位仍可再中'),
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('無法讀取', e.message, ui.ButtonSet.OK);
  }
}

function showSheetInfo() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var lines = sheets.map(function(sheet, i) {
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    var headers = '（空白）';
    if (lastRow >= 1 && lastCol >= 1) {
      var found = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
        .map(function(h) { return String(h).trim(); })
        .filter(function(h) { return h !== ''; });
      if (found.length) headers = found.join(' / ');
    }
    return (i + 1) + '.「' + sheet.getName() + '」' + (i === 0 ? '（目前讀這個）' : '') +
      '\n　資料列數：' + Math.max(lastRow - 1, 0) +
      '\n　標題列：' + headers;
  });

  SpreadsheetApp.getUi().alert('資料分頁檢查', lines.join('\n\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}

function getVoteSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!CONFIG.voteSheetName) return ss.getSheets()[0];
  var sheet = ss.getSheetByName(CONFIG.voteSheetName);
  if (!sheet) throw new Error('找不到分頁「' + CONFIG.voteSheetName + '」。');
  return sheet;
}

function getCandidates_() {
  var sheet = getVoteSheet_();
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    throw new Error(
      '分頁「' + sheet.getName() + '」沒有資料。\n\n' +
      '如果投票資料在別的分頁，請把最上面的 CONFIG.voteSheetName 改成那個分頁的名稱\n' +
      '（可先用選單的「檢查資料分頁」看有哪些分頁）。'
    );
  }

  var headers = values[0].map(function(h) { return String(h).trim(); });
  var seatCols = CONFIG.seatHeaders.map(function(name) {
    var index = headers.indexOf(name);
    if (index === -1) {
      throw new Error(
        '找不到欄位「' + name + '」。\n\n目前的標題列：\n' + headers.join(' / ') +
        '\n\n請把 CONFIG.seatHeaders 改成與實際標題一致。'
      );
    }
    return index;
  });

  var excluded = CONFIG.excludePreviousWinners ? getPreviousWinners_() : {};
  var seen = {};
  var candidates = [];

  for (var r = 1; r < values.length; r++) {
    var parts = seatCols.map(function(c) { return String(values[r][c]).trim(); });
    var incomplete = parts.some(function(p) { return p === ''; });
    if (incomplete) continue;

    var seat = parts.join(' ');
    if (excluded[seat]) continue;
    if (CONFIG.dedupeBySeat && seen[seat]) continue;

    seen[seat] = true;
    candidates.push({ seat: seat, parts: parts, row: r + 1 });
  }
  return candidates;
}

function getPreviousWinners_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.resultSheetName);
  var winners = {};
  if (!sheet || sheet.getLastRow() < 2) return winners;

  sheet.getRange(2, 3, sheet.getLastRow() - 1, CONFIG.seatHeaders.length).getValues().forEach(function(row) {
    var seat = row.map(function(v) { return String(v).trim(); }).join(' ');
    if (seat.trim()) winners[seat] = true;
  });
  return winners;
}

function pickRandom_(list, count) {
  var pool = list.slice();
  var picked = [];
  for (var i = 0; i < count; i++) {
    var j = Math.floor(Math.random() * pool.length);
    picked.push(pool.splice(j, 1)[0]);
  }
  return picked;
}

function recordWinners_(winners) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.resultSheetName);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.resultSheetName);
    sheet.appendRow(['抽獎時間', '梯次'].concat(CONFIG.seatHeaders).concat(['投票列號']));
    sheet.setFrozenRows(1);
  }

  var round = 1;
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues().forEach(function(row) {
      var value = parseInt(row[0], 10);
      if (value >= round) round = value + 1;
    });
  }

  var now = new Date();
  var rows = winners.map(function(w) {
    return [now, round].concat(w.parts).concat([w.row]);
  });
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  return round;
}
