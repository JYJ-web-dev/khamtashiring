const TARGET_GOAL = 110000;

// [외부 fetch API 지원] GET 및 POST 요청을 수신하여 JSON으로 응답
function doGet(e) {
  if (e && e.parameter && e.parameter.action) {
    return handleApiRequest_(e.parameter);
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('캄따시링 수행 일지')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

function doPost(e) {
  let params = {};
  if (e && e.postData && e.postData.contents) {
    try {
      params = JSON.parse(e.postData.contents);
    } catch (err) {
      params = e.parameter || {};
    }
  } else if (e && e.parameter) {
    params = e.parameter;
  }
  return handleApiRequest_(params);
}

function handleApiRequest_(params) {
  const action = params.action;
  let result = { success: false, message: '유효하지 않은 요청입니다.' };

  try {
    if (action === 'accessDashboard') {
      result = accessDashboard(params.name, params.phoneLast4);
    } else if (action === 'saveNgondroLog') {
      result = saveNgondroLog(params.payload || params);
    } else if (action === 'saveOtherLog') {
      result = saveOtherLog(params.payload || params);
    } else if (action === 'deleteOtherLog') {
      result = deleteOtherLog(params.memberId, params.logId, params.phoneLast4);
    } else if (action === 'registerNewMemberAndNgondro') {
      result = registerNewMemberAndNgondro(params.payload || params);
    } else if (action === 'updateNgondroAspiration') {
      result = updateNgondroAspiration(params.memberId, params.aspirationText);
    } else if (action === 'updateOtherAspiration') {
      result = updateOtherAspiration(params.practiceId, params.aspirationText);
    } else if (action === 'registerOtherPractice') {
      result = registerOtherPractice(params.payload || params);
    } else if (action === 'completeOtherPractice') {
      result = completeOtherPractice(params.practiceId, params.finalTotal);
    } else if (action === 'proceedNgondroNextRound') {
      result = proceedNgondroNextRound(params.memberId, params.memberName, params.currentRound, params.phoneLast4);
    }
  } catch (err) {
    result = { success: false, message: err.toString() };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// 6개 시트 1회 적재 맵핑
function getSheetsMap_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const map = {};
  for (let i = 0; i < sheets.length; i++) {
    map[sheets[i].getName()] = sheets[i];
  }
  return {
    ss: ss,
    memSheet: map['회원_명부'] || ss.insertSheet('회원_명부'),
    ngondroConfig: map['사가행_설정'] || ss.insertSheet('사가행_설정'),
    ngondroLog: map['사가행_기록'] || ss.insertSheet('사가행_기록'),
    ngondroSummary: map['사가행_누적집계'] || ss.insertSheet('사가행_누적집계'),
    otherConfig: map['기타수행_설정'] || ss.insertSheet('기타수행_설정'),
    otherLog: map['기타수행_기록'] || ss.insertSheet('기타수행_기록')
  };
}

function formatDateFast_(val) {
  if (!val) return '';
  if (typeof val === 'string') return val.substring(0, 10);
  if (val instanceof Date) {
    const y = val.getFullYear();
    const m = ('0' + (val.getMonth() + 1)).slice(-2);
    const d = ('0' + val.getDate()).slice(-2);
    return y + '-' + m + '-' + d;
  }
  return String(val).substring(0, 10);
}

function formatDateTimeFast_(val) {
  if (!val) return '';
  if (typeof val === 'string') return val;
  if (val instanceof Date) {
    const y = val.getFullYear();
    const m = ('0' + (val.getMonth() + 1)).slice(-2);
    const d = ('0' + val.getDate()).slice(-2);
    const hh = ('0' + val.getHours()).slice(-2);
    const mm = ('0' + val.getMinutes()).slice(-2);
    const ss = ('0' + val.getSeconds()).slice(-2);
    return `${y}-${m}-${d} ${hh}:${mm}:${ss}`;
  }
  return String(val);
}

function accessDashboard(name, phoneLast4) {
  try {
    const sheets = getSheetsMap_();
    const cleanName = (name || '').trim();
    const cleanPhone = ('' + phoneLast4).replace(/\D/g, '').slice(-4);
    
    if (cleanPhone.length !== 4) {
      return { success: false, message: '전화번호 뒷자리 4자리를 정확히 입력해 주세요.' };
    }

    const memRows = sheets.memSheet.getDataRange().getValues();
    let member = null;
    for (let i = 1; i < memRows.length; i++) {
      if (String(memRows[i][1]).trim() === cleanName && String(memRows[i][2]).trim() === cleanPhone) {
        member = { id: memRows[i][0], name: cleanName, phoneLast4: cleanPhone };
        break;
      }
    }

    const todayStr = formatDateFast_(new Date());
    if (!member) {
      return { success: true, isNewMember: true, name: cleanName, phoneLast4: cleanPhone, todayDate: todayStr };
    }

    return fetchDashboardDataWithSheets_(sheets, member, todayStr);
  } catch (err) {
    return { success: false, message: '조회 실패: ' + err.toString() };
  }
}

function fetchDashboardDataWithSheets_(sheets, member, todayStr) {
  const currentMonth = todayStr.substring(0, 7);

  if (!member.name) {
    const mRows = sheets.memSheet.getDataRange().getValues();
    for (let m = 1; m < mRows.length; m++) {
      if (mRows[m][0] === member.id) {
        member.name = String(mRows[m][1]).trim();
        member.phoneLast4 = String(mRows[m][2]).trim();
        break;
      }
    }
  }

  const nConfigRows = sheets.ngondroConfig.getDataRange().getValues();
  let round = 1;
  let startDate = todayStr;
  let ngondroAspiration = '';
  let totalCounts = { prostrations: 0, vajrasattva: 0, mandala: 0, guruYoga: 0 };
  let monthCounts = { prostrations: 0, vajrasattva: 0, mandala: 0, guruYoga: 0 };

  for (let j = nConfigRows.length - 1; j >= 1; j--) {
    if (nConfigRows[j][0] === member.id && nConfigRows[j][8] === '진행중') {
      round = Number(nConfigRows[j][2]) || 1;
      startDate = formatDateFast_(nConfigRows[j][3]);
      totalCounts.prostrations = Number(nConfigRows[j][4]) || 0;
      totalCounts.vajrasattva = Number(nConfigRows[j][5]) || 0;
      totalCounts.mandala = Number(nConfigRows[j][6]) || 0;
      totalCounts.guruYoga = Number(nConfigRows[j][7]) || 0;
      ngondroAspiration = nConfigRows[j][9] || '';
      break;
    }
  }

  const nLogRows = sheets.ngondroLog.getDataRange().getValues();
  const allNgondroLogs = [];
  for (let k = 1; k < nLogRows.length; k++) {
    if (nLogRows[k][2] === member.id && Number(nLogRows[k][4]) === round) {
      const p = Number(nLogRows[k][5]) || 0;
      const v = Number(nLogRows[k][6]) || 0;
      const m = Number(nLogRows[k][7]) || 0;
      const g = Number(nLogRows[k][8]) || 0;
      const memo = nLogRows[k][9] || '';
      const logDate = formatDateFast_(nLogRows[k][1]);

      totalCounts.prostrations += p;
      totalCounts.vajrasattva += v;
      totalCounts.mandala += m;
      totalCounts.guruYoga += g;

      if (logDate.startsWith(currentMonth)) {
        monthCounts.prostrations += p;
        monthCounts.vajrasattva += v;
        monthCounts.mandala += m;
        monthCounts.guruYoga += g;
      }
      allNgondroLogs.push({ date: logDate, p, v, m, g, memo });
    }
  }

  const oConfigRows = sheets.otherConfig.getDataRange().getValues();
  const activePractices = [];
  const completedPractices = [];

  for (let x = 1; x < oConfigRows.length; x++) {
    if (oConfigRows[x][1] === member.id) {
      const pId = oConfigRows[x][0];
      const pName = oConfigRows[x][3];
      const pPeriod = oConfigRows[x][4];
      const pStart = formatDateFast_(oConfigRows[x][5]);
      const pTarget = Number(oConfigRows[x][6]) || 0;
      const pUnit = oConfigRows[x][7];
      const pStatus = oConfigRows[x][8];
      const pAspiration = oConfigRows[x][9] || '';
      const pEndDate = oConfigRows[x][10] ? formatDateFast_(oConfigRows[x][10]) : '';
      const pFinalTotal = Number(oConfigRows[x][11]) || 0;

      if (pStatus === '진행중') {
        activePractices.push({ id: pId, name: pName, period: pPeriod, startDate: pStart, target: pTarget, unit: pUnit, aspiration: pAspiration, total: 0 });
      } else if (pStatus === '완료') {
        completedPractices.push({ id: pId, name: pName, period: pPeriod, startDate: pStart, endDate: pEndDate, total: pFinalTotal, unit: pUnit, aspiration: pAspiration });
      }
    }
  }

  const oLogRows = sheets.otherLog.getDataRange().getValues();
  const allOtherLogs = [];
  for (let y = 1; y < oLogRows.length; y++) {
    if (oLogRows[y][2] === member.id) {
      const rowTs = formatDateTimeFast_(oLogRows[y][0]);
      const logDate = formatDateFast_(oLogRows[y][1]);
      const loggedName = oLogRows[y][4];
      const amt = Number(oLogRows[y][5]) || 0;
      const memo = oLogRows[y][6] || '';

      const match = activePractices.find(ap => ap.name === loggedName);
      if (match) match.total += amt;

      allOtherLogs.push({ id: rowTs, date: logDate, name: loggedName, amount: amt, memo: memo });
    }
  }

  const allCompleted = 
    totalCounts.prostrations >= TARGET_GOAL &&
    totalCounts.vajrasattva >= TARGET_GOAL &&
    totalCounts.mandala >= TARGET_GOAL &&
    totalCounts.guruYoga >= TARGET_GOAL;

  return {
    success: true,
    isNewMember: false,
    member: member,
    todayDate: todayStr,
    round: round,
    startDate: startDate,
    targetGoal: TARGET_GOAL,
    totalCounts: totalCounts,
    monthCounts: monthCounts,
    allCompleted: allCompleted,
    ngondroAspiration: ngondroAspiration,
    allNgondroLogs: allNgondroLogs.reverse(),
    activePractices: activePractices,
    completedPractices: completedPractices.reverse(),
    allOtherLogs: allOtherLogs.reverse()
  };
}

function saveNgondroLog(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheets = getSheetsMap_();
    const nowStr = formatDateTimeFast_(new Date());
    const memo = (payload.memo || '').trim().substring(0, 200);
    const targetRound = Number(payload.round) || 1;

    sheets.ngondroLog.appendRow([
      nowStr,
      payload.date,
      payload.memberId,
      payload.memberName,
      targetRound,
      Number(payload.pros) || 0,
      Number(payload.vajra) || 0,
      Number(payload.mandala) || 0,
      Number(payload.guru) || 0,
      memo
    ]);

    const updatedDashboard = fetchDashboardDataWithSheets_(
      sheets,
      { id: payload.memberId, name: payload.memberName, phoneLast4: payload.phoneLast4 },
      payload.date
    );

    const tot = updatedDashboard.totalCounts;
    const avgPct = (((tot.prostrations + tot.vajrasattva + tot.mandala + tot.guruYoga) / (TARGET_GOAL * 4)) * 100).toFixed(1) + '%';
    
    const sumData = sheets.ngondroSummary.getDataRange().getValues();
    let targetRow = -1;
    for (let s = 1; s < sumData.length; s++) {
      if (sumData[s][0] === payload.memberId && Number(sumData[s][3]) === targetRound) {
        targetRow = s + 1;
        break;
      }
    }

    const rowVal = [payload.memberId, payload.memberName, payload.phoneLast4 || '', targetRound, tot.prostrations, tot.vajrasattva, tot.mandala, tot.guruYoga, avgPct, payload.date];
    if (targetRow > 0) {
      sheets.ngondroSummary.getRange(targetRow, 1, 1, 10).setValues([rowVal]);
    } else {
      sheets.ngondroSummary.appendRow(rowVal);
    }

    return { success: true, updatedDashboard: updatedDashboard };
  } catch (err) {
    return { success: false, message: '저장 처리 지연: ' + err.toString() };
  } finally {
    lock.releaseLock();
  }
}

function saveOtherLog(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheets = getSheetsMap_();
    const nowStr = formatDateTimeFast_(new Date());
    const memo = (payload.memo || '').trim().substring(0, 200);

    sheets.otherLog.appendRow([
      nowStr,
      payload.date,
      payload.memberId,
      payload.memberName,
      payload.practiceName,
      Number(payload.amount) || 0,
      memo
    ]);

    const updatedDashboard = fetchDashboardDataWithSheets_(
      sheets,
      { id: payload.memberId, name: payload.memberName, phoneLast4: payload.phoneLast4 },
      payload.date
    );

    return { success: true, updatedDashboard: updatedDashboard };
  } catch (err) {
    return { success: false, message: '저장 처리 지연: ' + err.toString() };
  } finally {
    lock.releaseLock();
  }
}

function deleteOtherLog(memberId, logId, phoneLast4) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheets = getSheetsMap_();
    const rows = sheets.otherLog.getDataRange().getValues();
    let targetRow = -1;

    for (let i = 1; i < rows.length; i++) {
      const rowTs = formatDateTimeFast_(rows[i][0]);
      if (rows[i][2] === memberId && rowTs === logId) {
        targetRow = i + 1;
        break;
      }
    }

    if (targetRow > 0) {
      sheets.otherLog.deleteRow(targetRow);
      const todayStr = formatDateFast_(new Date());
      const updatedDashboard = fetchDashboardDataWithSheets_(sheets, { id: memberId, name: '', phoneLast4: phoneLast4 }, todayStr);
      return { success: true, updatedDashboard: updatedDashboard };
    }
    return { success: false, message: '삭제 대상을 찾을 수 없습니다.' };
  } catch (err) {
    return { success: false, message: '삭제 실패: ' + err.toString() };
  } finally {
    lock.releaseLock();
  }
}

function updateOtherAspiration(practiceId, aspirationText) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheets = getSheetsMap_();
    const rows = sheets.otherConfig.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === practiceId) {
        sheets.otherConfig.getRange(i + 1, 10).setValue(aspirationText.trim());
        return { success: true };
      }
    }
    return { success: false, message: '기도 설정을 찾을 수 없습니다.' };
  } catch (err) {
    return { success: false, message: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

function registerNewMemberAndNgondro(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheets = getSheetsMap_();
    const cleanPhone = ('' + payload.phoneLast4).replace(/\D/g, '').slice(-4);
    const nowStr = formatDateTimeFast_(new Date());
    const memberId = 'M-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMddHHmmss');

    sheets.memSheet.appendRow([memberId, payload.name.trim(), cleanPhone, nowStr]);

    const p = Number(payload.initPros) || 0;
    const v = Number(payload.initVajra) || 0;
    const m = Number(payload.initMandala) || 0;
    const g = Number(payload.initGuru) || 0;

    sheets.ngondroConfig.appendRow([memberId, payload.name.trim(), 1, payload.startDate, p, v, m, g, '진행중', (payload.aspiration || '').trim()]);

    const avgPct = (((p + v + m + g) / (TARGET_GOAL * 4)) * 100).toFixed(1) + '%';
    sheets.ngondroSummary.appendRow([memberId, payload.name.trim(), cleanPhone, 1, p, v, m, g, avgPct, payload.startDate]);

    return { success: true, member: { id: memberId, name: payload.name.trim(), phoneLast4: cleanPhone } };
  } catch (err) {
    return { success: false, message: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

function updateNgondroAspiration(memberId, aspirationText) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheets = getSheetsMap_();
    const rows = sheets.ngondroConfig.getDataRange().getValues();
    for (let i = rows.length - 1; i >= 1; i--) {
      if (rows[i][0] === memberId && rows[i][8] === '진행중') {
        sheets.ngondroConfig.getRange(i + 1, 10).setValue(aspirationText.trim());
        return { success: true };
      }
    }
    return { success: false, message: '설정을 찾을 수 없습니다.' };
  } catch (err) {
    return { success: false, message: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

function registerOtherPractice(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheets = getSheetsMap_();
    const newId = 'OTH-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMddHHmmss');
    sheets.otherConfig.appendRow([newId, payload.memberId, payload.memberName, payload.practiceName.trim(), payload.periodType, payload.startDate, 0, payload.unit.trim() || '독', '진행중', (payload.aspiration || '').trim(), '', 0]);
    return { success: true };
  } catch (err) {
    return { success: false, message: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

function completeOtherPractice(practiceId, finalTotal) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheets = getSheetsMap_();
    const rows = sheets.otherConfig.getDataRange().getValues();
    const nowStr = formatDateFast_(new Date());
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === practiceId) {
        sheets.otherConfig.getRange(i + 1, 9).setValue('완료');
        sheets.otherConfig.getRange(i + 1, 11).setValue(nowStr);
        sheets.otherConfig.getRange(i + 1, 12).setValue(finalTotal);
        return { success: true };
      }
    }
    return { success: false, message: '기록을 찾을 수 없습니다.' };
  } catch (err) {
    return { success: false, message: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

function proceedNgondroNextRound(memberId, memberName, currentRound, phoneLast4) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheets = getSheetsMap_();
    const rows = sheets.ngondroConfig.getDataRange().getValues();
    const todayStr = formatDateFast_(new Date());
    const currR = Number(currentRound) || 1;
    let oldAsp = '';

    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === memberId && Number(rows[i][2]) === currR) {
        sheets.ngondroConfig.getRange(i + 1, 9).setValue('완료');
        oldAsp = rows[i][9] || '';
        break;
      }
    }

    const nextR = currR + 1;
    sheets.ngondroConfig.appendRow([memberId, memberName, nextR, todayStr, 0, 0, 0, 0, '진행중', oldAsp]);

    let pLast4 = phoneLast4;
    if (!pLast4) {
      const mRows = sheets.memSheet.getDataRange().getValues();
      for (let m = 1; m < mRows.length; m++) {
        if (mRows[m][0] === memberId) {
          pLast4 = mRows[m][2];
          break;
        }
      }
    }

    sheets.ngondroSummary.appendRow([memberId, memberName, pLast4 || '', nextR, 0, 0, 0, 0, '0.0%', todayStr]);
    return { success: true };
  } catch (err) {
    return { success: false, message: err.toString() };
  } finally {
    lock.releaseLock();
  }
}
