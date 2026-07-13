/**
 * ============================================================
 * 근무표 QR 공유 웹앱 - Code.gs
 * ============================================================
 * Google Sheets + Apps Script 기반
 * v0.4.5 - 계급 표시, 근무 통계 추가, 입력 제한 적용, 실시간 조회 생성
 * ============================================================
 */

// ── 상수: 스프레드시트 ID ────────────────────────────────────
var SPREADSHEET_ID = '1kNhqHfrLQxwyONuahpdaz934LLjWxMIixbc2ws-2A0E';

// ── 상수: 시트 이름 ──────────────────────────────────────────
var SHEET_NAMES = {
  SETTINGS:  '설정',
  WORKERS:   '근무자',
  VACATION:  '휴가',
  PINNED:    '말뚝근무',
  RESULTS:   '생성결과'
};

// ── 상수: 각 시트의 헤더 행 ──────────────────────────────────
var HEADERS = {};
HEADERS[SHEET_NAMES.SETTINGS]  = ['항목', '값'];
HEADERS[SHEET_NAMES.WORKERS]   = ['순번', '계급', '이름', '소속', '사용여부'];
HEADERS[SHEET_NAMES.VACATION]  = ['날짜', '이름', '유형', '기간', '비고'];
HEADERS[SHEET_NAMES.PINNED]    = ['날짜', '시간', '이름', '비고'];
HEADERS[SHEET_NAMES.RESULTS]   = ['sid', '생성일시', '날짜', '제목', 'html', 'json'];


// ============================================================
// setupSheets()
// 필요한 시트 5개를 생성하고, 각 시트에 헤더 행을 세팅한다.
// ============================================================
function setupSheets() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var created = [];   // 새로 만든 시트 목록
  var skipped = [];   // 이미 존재해서 건너뛴 시트 목록

  var names = [
    SHEET_NAMES.SETTINGS,
    SHEET_NAMES.WORKERS,
    SHEET_NAMES.VACATION,
    SHEET_NAMES.PINNED,
    SHEET_NAMES.RESULTS
  ];

  for (var i = 0; i < names.length; i++) {
    var name = names[i];
    var sheet = ss.getSheetByName(name);

    if (sheet) {
      skipped.push(name);
    } else {
      sheet = ss.insertSheet(name);
      created.push(name);
    }

    var header = HEADERS[name];
    if (header && header.length > 0) {
      var lastCol = sheet.getLastColumn();
      var firstCell = sheet.getRange(1, 1).getValue();

      // 시트가 아예 비어있는 경우 헤더 새로 생성
      if (!firstCell || firstCell === '') {
        sheet.getRange(1, 1, 1, header.length).setValues([header]);
      } else {
        // 기존 시트 업그레이드 로직 (열 삽입)
        if (name === SHEET_NAMES.WORKERS && lastCol === 4) {
          sheet.insertColumnBefore(2);
          sheet.getRange(1, 2).setValue('계급');
          Logger.log('Upgraded WORKERS sheet: inserted 계급 column at B');
        } else if (name === SHEET_NAMES.VACATION && lastCol === 4) {
          sheet.insertColumnBefore(4);
          sheet.getRange(1, 4).setValue('기간');
          Logger.log('Upgraded VACATION sheet: inserted 기간 column at D');
        }
        // 헤더 덮어쓰기 (정확한 텍스트로 보장)
        sheet.getRange(1, 1, 1, header.length).setValues([header]);
      }

      sheet.getRange(1, 1, 1, header.length)
        .setFontWeight('bold')
        .setBackground('#4a86c8')
        .setFontColor('#ffffff');

      for (var c = 1; c <= header.length; c++) {
        sheet.autoResizeColumn(c);
      }
    }
  }

  // 데이터 유효성 검사 (드롭다운) 설정
  try {
    var workersSheet = ss.getSheetByName(SHEET_NAMES.WORKERS);
    if (workersSheet) {
      var rankRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['소방령', '소방경', '소방위', '소방장', '소방교', '소방사'])
        .setAllowInvalid(false)
        .setHelpText('소방령, 소방경, 소방위, 소방장, 소방교, 소방사 중 하나를 선택해 주세요.')
        .build();
      workersSheet.getRange(2, 2, 999, 1).setDataValidation(rankRule);
    }
    
    var vacationSheet = ss.getSheetByName(SHEET_NAMES.VACATION);
    if (vacationSheet) {
      var vacRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['주간연가', '야간연가', '전일연가', '장기재직휴가', '교육', '출장', '지각', '육아시간', '휴직'])
        .setAllowInvalid(false)
        .setHelpText('주간연가, 야간연가, 전일연가, 장기재직휴가, 교육, 출장, 지각, 육아시간, 휴직 중 하나를 선택해 주세요.')
        .build();
      vacationSheet.getRange(2, 3, 999, 1).setDataValidation(vacRule);
    }
  } catch (e) {
    Logger.log('⚠️ 유효성 검사 규칙 설정 중 오류: ' + e.message);
  }

  var msg = '✅ setupSheets 완료\n';
  msg += '  생성됨: ' + (created.length > 0 ? created.join(', ') : '없음') + '\n';
  msg += '  업그레이드/유지됨: ' + (skipped.length > 0 ? skipped.join(', ') : '없음');
  Logger.log(msg);

  return { created: created, skipped: skipped };
}


// ============================================================
// seedSampleData()
// 요구사항에 맞춰 시트를 비우고 62명(사고자 2명 포함) 및 설정을 입력한다.
// ============================================================
function seedSampleData() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var seeded = [];
    var skipped = [];

    // ── 1) 설정 시트 ──────────────────────────────────────────
    var settingsSheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
    if (settingsSheet) {
      if (settingsSheet.getLastRow() >= 2) {
        settingsSheet.getRange(2, 1, settingsSheet.getLastRow() - 1, Math.max(1, settingsSheet.getLastColumn())).clearContent();
      }
      var settingsData = [
        ['근무일자',         '2026-06-24'],
        ['근무시작',         '09:00'],
        ['근무종료',         '09:00'],
        ['시간단위(분)',     '60'],
        ['시간대별필요인원', '1']
      ];
      settingsSheet.getRange(2, 1, settingsData.length, 2).setValues(settingsData);
      seeded.push(SHEET_NAMES.SETTINGS);
    }

    // ── 2) 근무자 시트 (총 62명: 1팀 20명, 2팀 20명, 3팀 20명, 사고자 2명) ──
    var workersSheet = ss.getSheetByName(SHEET_NAMES.WORKERS);
    if (workersSheet) {
      if (workersSheet.getLastRow() >= 2) {
        workersSheet.getRange(2, 1, workersSheet.getLastRow() - 1, Math.max(1, workersSheet.getLastColumn())).clearContent();
      }
      // 기존 유효성 검사 규칙 모두 제거 (프로그램식 입력을 방해하지 않도록)
      workersSheet.getRange(2, 1, 999, 5).clearDataValidations();
      
      var workersData = [];
      var seq = 1;
      var ranks = ['소방령', '소방경', '소방위', '소방장', '소방교', '소방사'];

      // 1팀 20명
      for (var i = 1; i <= 20; i++) {
        var rank = ranks[Math.floor(Math.random() * ranks.length)];
        workersData.push([seq, rank, '1팀' + (i < 10 ? '0' + i : i), '1팀', 'Y']);
        seq++;
      }
      // 2팀 20명
      for (var i = 1; i <= 20; i++) {
        var rank = ranks[Math.floor(Math.random() * ranks.length)];
        workersData.push([seq, rank, '2팀' + (i < 10 ? '0' + i : i), '2팀', 'Y']);
        seq++;
      }
      // 3팀 20명
      for (var i = 1; i <= 20; i++) {
        var rank = ranks[Math.floor(Math.random() * ranks.length)];
        workersData.push([seq, rank, '3팀' + (i < 10 ? '0' + i : i), '3팀', 'Y']);
        seq++;
      }
      // 사고자 2명 (휴직 1명, 출장 1명)
      workersData.push([seq, '소방교', '사고자1', '1팀', 'Y']);
      seq++;
      workersData.push([seq, '소방사', '사고자2', '2팀', 'Y']);
      seq++;

      // 일괄 데이터 쓰기
      workersSheet.getRange(2, 1, workersData.length, 5).setValues(workersData);
      
      // 유효성 검사 재설정
      var rankRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['소방령', '소방경', '소방위', '소방장', '소방교', '소방사'])
        .setAllowInvalid(false)
        .setHelpText('소방령, 소방경, 소방위, 소방장, 소방교, 소방사 중 하나를 선택해 주세요.')
        .build();
      workersSheet.getRange(2, 2, 999, 1).setDataValidation(rankRule);
      
      seeded.push(SHEET_NAMES.WORKERS);
    }

    // ── 3) 휴가 시트 (사고자 2명 정보 등록) ──────────────────────────────────
    var vacationSheet = ss.getSheetByName(SHEET_NAMES.VACATION);
    if (vacationSheet) {
      if (vacationSheet.getLastRow() >= 2) {
        vacationSheet.getRange(2, 1, vacationSheet.getLastRow() - 1, Math.max(1, vacationSheet.getLastColumn())).clearContent();
      }
      // 기존 유효성 검사 규칙 모두 제거
      vacationSheet.getRange(2, 1, 999, 5).clearDataValidations();
      
      var vacationData = [
        ['2026-06-24', '사고자1', '휴직', '2026-06-24 ~ 2026-07-24', '장기휴직'],
        ['2026-06-24', '사고자2', '출장', '2026-06-24', '서울 출장']
      ];
      vacationSheet.getRange(2, 1, vacationData.length, 5).setValues(vacationData);
      
      // 유효성 검사 재설정
      var vacRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['주간연가', '야간연가', '전일연가', '장기재직휴가', '교육', '출장', '지각', '육아시간', '휴직'])
        .setAllowInvalid(false)
        .setHelpText('주간연가, 야간연가, 전일연가, 장기재직휴가, 교육, 출장, 지각, 육아시간, 휴직 중 하나를 선택해 주세요.')
        .build();
      vacationSheet.getRange(2, 3, 999, 1).setDataValidation(vacRule);
      
      seeded.push(SHEET_NAMES.VACATION);
    }

    // ── 4) 말뚝근무 시트 (말뚝근무자 제거) ──────────────────────────────────
    var pinnedSheet = ss.getSheetByName(SHEET_NAMES.PINNED);
    if (pinnedSheet) {
      if (pinnedSheet.getLastRow() >= 2) {
        pinnedSheet.getRange(2, 1, pinnedSheet.getLastRow() - 1, Math.max(1, pinnedSheet.getLastColumn())).clearContent();
      }
      seeded.push(SHEET_NAMES.PINNED);
    }

    skipped.push(SHEET_NAMES.RESULTS + ' (자동 생성 대상)');

    var msg = '✅ seedSampleData 완료 (62명 근무자, 사고자 2명 설정, 말뚝근무자 제거 완료)\n';
    msg += '  입력됨: ' + seeded.join(', ');
    Logger.log(msg);

    return { success: true, seeded: seeded, skipped: skipped };
  } catch (e) {
    Logger.log('❌ seedSampleData 에러: ' + e.message);
    return { success: false, error: e.message };
  }
}

function _hasData(sheet) {
  return sheet.getLastRow() > 1;
}


// ============================================================
//  2단계: 근무표 생성
// ============================================================

function generateSchedule(optDate) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  var config = _readSettings(ss);
  if (!config) return null;

  // 특정 날짜 조회 요청 시 해당 날짜로 덮어씌움
  if (optDate) {
    config.date = _normalizeDate(optDate);
  }

  var allWorkers = _readWorkers(ss);
  var vacationList = _readVacationList(ss, config.date); 
  
  var candidates = [];
  var vacMap = {};
  for (var i = 0; i < vacationList.length; i++) {
    vacMap[vacationList[i].name] = vacationList[i];
    vacationList[i].rank = _findWorkerRank(allWorkers, vacationList[i].name);
  }

  var stats = _getWorkerStats(ss);
  var activeTeam = _getTeamOnDuty(config.date);
  Logger.log('Active team on duty for ' + config.date + ': ' + activeTeam);

  for (var i = 0; i < allWorkers.length; i++) {
    var worker = allWorkers[i];
    
    // 해당 날짜의 당번 팀만 근무 후보(candidates)에 등록
    if (worker.team !== activeTeam) {
      continue;
    }
    
    var vac = vacMap[worker.name];
    var cCount = stats.counts[worker.name] || { count1214: 0, count0003: 0 };
    candidates.push({
      seq: worker.seq,
      rank: worker.rank,
      name: worker.name,
      team: worker.team,
      isFull: vac ? vac.isFull : false,
      isDayExempt: vac ? vac.isDay : false,
      isNightExempt: vac ? vac.isNight : false,
      count1214: cCount.count1214,
      count0003: cCount.count0003,
      isAssignedToday: false
    });
  }

  if (candidates.length === 0) {
    Logger.log('⚠️ 근무 가능한 인원이 없습니다.');
    return null;
  }

  var timeSlots = _buildTimeSlots();
  config.requiredPerSlot = 1; 

  var pinnedList = _readPinnedList(ss, config.date); 
  
  var pinnedMap = {};
  for (var i = 0; i < pinnedList.length; i++) {
    var p = pinnedList[i];
    if (!pinnedMap[p.time]) pinnedMap[p.time] = [];
    pinnedMap[p.time].push({ name: p.name, comment: p.comment });
  }

  var schedule = [];
  for (var t = 0; t < timeSlots.length; t++) {
    schedule.push({
      time: timeSlots[t].time,
      label: timeSlots[t].label,
      isDay: timeSlots[t].isDay,
      isSpecial: timeSlots[t].isSpecial,
      specialId: timeSlots[t].specialId,
      workers: []
    });
  }

  // 5-a) 말뚝근무 먼저 배치
  for (var s = 0; s < schedule.length; s++) {
    var slot = schedule[s];
    if (pinnedMap[slot.time]) {
      for (var p = 0; p < pinnedMap[slot.time].length; p++) {
        var pName = pinnedMap[slot.time][p].name;
        var pRank = _findWorkerRank(allWorkers, pName);
        for (var c = 0; c < candidates.length; c++) {
          if (candidates[c].name === pName) {
            candidates[c].isAssignedToday = true;
          }
        }
        slot.workers.push({
          name: pName,
          rank: pRank,
          team: _findWorkerTeam(allWorkers, pName),
          isPinned: true,
          comment: pinnedMap[slot.time][p].comment
        });
      }
    }
  }

  // 5-b) 특별 시간대 배정 (12~14, 00~03)
  for (var s = 0; s < schedule.length; s++) {
    var slot = schedule[s];
    if (!slot.isSpecial) continue;
    
    var needed = config.requiredPerSlot - slot.workers.length;
    if (needed <= 0) continue;

    var available = [];
    for (var c = 0; c < candidates.length; c++) {
      var w = candidates[c];
      if (w.isAssignedToday || w.isFull) continue;
      if (slot.isDay && w.isDayExempt) continue;
      if (!slot.isDay && w.isNightExempt) continue;
      available.push(w);
    }

    available.sort(function(a, b) {
      var aCount = a[slot.specialId];
      var bCount = b[slot.specialId];
      if (aCount !== bCount) return aCount - bCount;
      return a.seq - b.seq;
    });

    for (var i = 0; i < needed && i < available.length; i++) {
      var picked = available[i];
      slot.workers.push({
        name: picked.name,
        rank: picked.rank || '',
        team: picked.team,
        isPinned: false,
        comment: ''
      });
      picked.isAssignedToday = true;
    }
  }

  // 5-c) 일반 시간대 배정 (Sliding Window)
  var vBaseIdx = 0;
  var lastBaseName = _getLastVirtualBaseNameForTeam(ss, activeTeam);
  if (lastBaseName) {
    for (var c = 0; c < candidates.length; c++) {
      if (candidates[c].name === lastBaseName) {
        vBaseIdx = c; break;
      }
    }
  }
  
  var candidateIdx = vBaseIdx;
  
  for (var s = 0; s < schedule.length; s++) {
    var slot = schedule[s];
    if (slot.isSpecial) continue; 

    var needed = config.requiredPerSlot - slot.workers.length;
    var tried = 0;
    while (needed > 0 && tried < candidates.length) {
      var w = candidates[candidateIdx % candidates.length];
      candidateIdx++;
      tried++;

      if (w.isAssignedToday || w.isFull) continue;
      if (slot.isDay && w.isDayExempt) continue;
      if (!slot.isDay && w.isNightExempt) continue;

      slot.workers.push({
        name: w.name,
        rank: w.rank || '',
        team: w.team,
        isPinned: false,
        comment: ''
      });
      w.isAssignedToday = true;
      needed--;
      tried = 0; 
    }
  }

  var nextBaseIdx = (vBaseIdx + 1) % candidates.length;
  var nextVirtualBaseName = candidates[nextBaseIdx].name;

  var sid = _generateSid();

  // ── 통계 요약 계산 ─────────────────────────────────────
  var totalCount = allWorkers.length;
  var vacationCount = vacationList.length;

  // 당번팀 소속 근무자 중 오늘 휴가(사고)가 아닌 전체 인원을 당번 수에 집계
  var onDutyCount = 0;
  for (var w = 0; w < allWorkers.length; w++) {
    var worker = allWorkers[w];
    if (worker.team === activeTeam) {
      var isOnVacation = false;
      for (var v = 0; v < vacationList.length; v++) {
        if (vacationList[v].name === worker.name) {
          isOnVacation = true; break;
        }
      }
      if (!isOnVacation) {
        onDutyCount++;
      }
    }
  }
  var offDutyCount = totalCount - vacationCount - onDutyCount;

  var resultJson = {
    sid: sid,
    date: config.date,
    title: '신평119안전센터 소내근무표',
    startTime: config.startTime,
    endTime: config.endTime,
    intervalMin: config.intervalMin,
    requiredPerSlot: config.requiredPerSlot,
    vacationList: vacationList, 
    schedule: schedule,
    nextVirtualBaseName: nextVirtualBaseName,
    statsSummary: {
      totalCount: totalCount,
      vacationCount: vacationCount,
      onDutyCount: onDutyCount,
      offDutyCount: offDutyCount
    }
  };
  
  var jsonStr = JSON.stringify(resultJson);
  var htmlStr = _buildScheduleHtml(resultJson);

  var resultsSheet = ss.getSheetByName(SHEET_NAMES.RESULTS);
  if (!resultsSheet) {
    Logger.log('⚠️ "생성결과" 시트가 없습니다.');
    return null;
  }

  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  var newRow = [sid, now, config.date, resultJson.title, htmlStr, jsonStr];
  resultsSheet.appendRow(newRow);

  return resultJson;
}

function _readSettings(ss) {
  var sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (!sheet || sheet.getLastRow() < 2) return null;

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  var map = {};
  for (var i = 0; i < data.length; i++) {
    var key = String(data[i][0]).trim();
    var val = data[i][1];
    if (key === '근무일자') {
      map[key] = _normalizeDate(val);
    } else if (key === '근무시작' || key === '근무종료') {
      map[key] = _normalizeTime(val);
    } else {
      map[key] = String(val).trim();
    }
  }

  return {
    date:            map['근무일자']         || '',
    startTime:       map['근무시작']         || '09:00',
    endTime:         map['근무종료']         || '09:00',
    intervalMin:     parseInt(map['시간단위(분)'])     || 60,
    requiredPerSlot: parseInt(map['시간대별필요인원']) || 1
  };
}

function _readWorkers(ss) {
  var sheet = ss.getSheetByName(SHEET_NAMES.WORKERS);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  var workers = [];
  for (var i = 0; i < data.length; i++) {
    var useYn = String(data[i][4]).trim().toUpperCase();
    if (useYn === 'Y') {
      workers.push({
        seq:  data[i][0],
        rank: String(data[i][1]).trim(),
        name: String(data[i][2]).trim(),
        team: String(data[i][3]).trim()
      });
    }
  }
  return workers;
}

function _readVacationList(ss, dateStr) {
  var sheet = ss.getSheetByName(SHEET_NAMES.VACATION);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  var list = [];
  var seen = {};

  for (var i = 0; i < data.length; i++) {
    // 해당 날짜가 휴가기간 내에 포함되는지 검사
    if (_isDateInPeriod(dateStr, data[i][0], data[i][3])) {
      var name = String(data[i][1]).trim();
      if (!seen[name]) {
        seen[name] = true;
        var type = String(data[i][2]).trim();
        var period = String(data[i][3]).trim();
        var comment = String(data[i][4]).trim();
        var fullText = type + ' ' + comment;
        
        var isDay = fullText.indexOf('주간') !== -1;
        var isNight = fullText.indexOf('야간') !== -1;
        var isFull = fullText.indexOf('전일') !== -1 || (!isDay && !isNight);
        
        list.push({
          name: name,
          type: type,
          period: period,
          comment: comment,
          isDay: isDay,
          isNight: isNight,
          isFull: isFull
        });
      }
    }
  }
  return list;
}

function _getWorkerStats(ss) {
  var sheet = ss.getSheetByName(SHEET_NAMES.RESULTS);
  var stats = { counts: {}, nextVirtualBaseName: null };
  if (!sheet || sheet.getLastRow() < 2) return stats;

  var data = sheet.getRange(2, 6, sheet.getLastRow() - 1, 1).getValues(); 
  var lastValidBaseName = null;

  for (var i = 0; i < data.length; i++) {
    try {
      var jsonStr = data[i][0];
      if (!jsonStr) continue;
      var obj = JSON.parse(jsonStr);
      
      if (obj.schedule) {
        for (var s = 0; s < obj.schedule.length; s++) {
          var slot = obj.schedule[s];
          var specialId = slot.specialId;
          
          if (specialId) {
            for (var w = 0; w < slot.workers.length; w++) {
              var wName = slot.workers[w].name;
              if (!stats.counts[wName]) stats.counts[wName] = { count1214: 0, count0003: 0 };
              stats.counts[wName][specialId] = (stats.counts[wName][specialId] || 0) + 1;
            }
          }
        }
      }
      if (obj.nextVirtualBaseName) {
        lastValidBaseName = obj.nextVirtualBaseName;
      }
    } catch (e) {}
  }
  
  stats.nextVirtualBaseName = lastValidBaseName;
  return stats;
}

function _readPinnedList(ss, dateStr) {
  var sheet = ss.getSheetByName(SHEET_NAMES.PINNED);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  var list = [];
  for (var i = 0; i < data.length; i++) {
    var rowDate = _normalizeDate(data[i][0]);
    if (rowDate === dateStr) {
      list.push({
        time: _normalizeTime(data[i][1]),
        name: String(data[i][2]).trim(),
        comment: String(data[i][3]).trim()
      });
    }
  }
  return list;
}

function _findWorkerTeam(allWorkers, name) {
  for (var i = 0; i < allWorkers.length; i++) {
    if (allWorkers[i].name === name) {
      return allWorkers[i].team;
    }
  }
  return '기타';
}

function _findWorkerRank(allWorkers, name) {
  for (var i = 0; i < allWorkers.length; i++) {
    if (allWorkers[i].name === name) {
      return allWorkers[i].rank || '';
    }
  }
  return '';
}

function _buildTimeSlots() {
  return [
    { time: '09:00', label: '09~12', isDay: true,  isSpecial: false, specialId: null },
    { time: '12:00', label: '12~14', isDay: true,  isSpecial: true,  specialId: 'count1214' },
    { time: '14:00', label: '14~16', isDay: true,  isSpecial: false, specialId: null },
    { time: '16:00', label: '16~18', isDay: true,  isSpecial: false, specialId: null },
    { time: '18:00', label: '18~20', isDay: false, isSpecial: false, specialId: null },
    { time: '20:00', label: '20~22', isDay: false, isSpecial: false, specialId: null },
    { time: '22:00', label: '22~00', isDay: false, isSpecial: false, specialId: null },
    { time: '00:00', label: '00~03', isDay: false, isSpecial: true,  specialId: 'count0003' },
    { time: '03:00', label: '03~06', isDay: false, isSpecial: false, specialId: null },
    { time: '06:00', label: '06~09', isDay: false, isSpecial: false, specialId: null }
  ];
}

function _generateSid() {
  var chars = 'abcdefghijklmnopqrstuvwxyz0123456789';
  var sid = '';
  for (var i = 0; i < 8; i++) {
    sid += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return sid;
}

function _buildScheduleHtml(result) {
  var html = '';
  html += '<div style="max-width:480px;margin:0 auto;font-family:\'Apple SD Gothic Neo\',sans-serif;">';
  html += '<h2 style="text-align:center;color:#1a73e8;">📋 ' + result.title + '</h2>';

  if (result.statsSummary) {
    html += '<div style="display:flex;justify-content:space-around;background:#f0f4f9;padding:10px;border-radius:8px;margin-bottom:12px;font-size:13px;color:#333;text-align:center;">';
    html += '<div><strong>총원</strong><br>' + result.statsSummary.totalCount + '명</div>';
    html += '<div style="border-left:1px solid #ccc;padding-left:10px;"><strong>사고자</strong><br>' + result.statsSummary.vacationCount + '명</div>';
    html += '<div style="border-left:1px solid #ccc;padding-left:10px;"><strong>당번</strong><br>' + result.statsSummary.onDutyCount + '명</div>';
    html += '<div style="border-left:1px solid #ccc;padding-left:10px;"><strong>비번</strong><br>' + result.statsSummary.offDutyCount + '명</div>';
    html += '</div>';
  }

  if (result.vacationList && result.vacationList.length > 0) {
    html += '<div style="margin:8px 0;padding:8px 12px;background:#fff3f3;border-radius:8px;border-left:4px solid #e53935;">';
    html += '<strong>🚫 사고자 현황:</strong> ';
    for (var v = 0; v < result.vacationList.length; v++) {
      var vac = result.vacationList[v];
      var nameWithRank = (vac.rank ? vac.rank + ' ' : '') + vac.name;
      var detail = nameWithRank;
      if (vac.type) {
        detail += ' (' + vac.type;
        if (vac.period && vac.period !== vac.date) {
          detail += ' | ' + vac.period;
        }
        if (vac.comment) {
          detail += ' - ' + vac.comment;
        }
        detail += ')';
      }
      html += '<span style="display:inline-block;padding:2px 8px;margin:2px;background:#ffcdd2;border-radius:12px;font-size:13px;color:#c62828;">';
      html += detail + '</span>';
    }
    html += '</div>';
  }

  html += '<table style="width:100%;border-collapse:collapse;margin-top:12px;">';
  html += '<thead><tr style="background:#1a73e8;color:#fff;">';
  html += '<th style="padding:10px 8px;text-align:center;border:1px solid #ddd;">시간</th>';
  html += '<th style="padding:10px 8px;text-align:center;border:1px solid #ddd;">근무자</th>';
  html += '</tr></thead><tbody>';

  for (var s = 0; s < result.schedule.length; s++) {
    var slot = result.schedule[s];
    var bgColor = (s % 2 === 0) ? '#ffffff' : '#f5f5f5';
    html += '<tr style="background:' + bgColor + ';">';
    html += '<td style="padding:8px;text-align:center;border:1px solid #e0e0e0;font-weight:bold;white-space:nowrap;">' + (slot.label || slot.time) + '</td>';
    html += '<td style="padding:8px;border:1px solid #e0e0e0;">';

    for (var w = 0; w < slot.workers.length; w++) {
      var worker = slot.workers[w];
      if (!worker) continue;

      var name = worker.name || '';
      var rank = worker.rank || '';
      var team = worker.team || '';
      var isPinned = !!worker.isPinned;
      var comment = worker.comment || '';

      var dispName = (rank ? rank + ' ' : '') + name;
      if (team) {
        dispName += ' (' + team + ')';
      }
      
      if (isPinned) {
        var pinLabel = '📌 ' + dispName;
        if (comment) {
          pinLabel += ' [' + comment + ']';
        }
        html += '<span style="display:inline-block;padding:3px 10px;margin:2px;background:#bbdefb;color:#1565c0;border-radius:12px;font-size:13px;">' + pinLabel + '</span>';
      } else {
        html += '<span style="display:inline-block;padding:3px 10px;margin:2px;background:#e8f5e9;color:#2e7d32;border-radius:12px;font-size:13px;">' + dispName + '</span>';
      }
    }

    html += '</td></tr>';
  }

  html += '</tbody></table>';
  html += '<p style="text-align:center;color:#999;font-size:11px;margin-top:16px;">생성일시: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') + '</p>';
  html += '</div>';

  return html;
}

// ── 신규 API 함수: 계급이 누락된 사용중인 근무자 조회 ──────────────────────
function getMissingRankWorkers() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAMES.WORKERS);
    if (!sheet || sheet.getLastRow() < 2) return [];
    
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
    var missing = [];
    for (var i = 0; i < data.length; i++) {
      var useYn = String(data[i][4]).trim().toUpperCase();
      var rank = String(data[i][1]).trim();
      var name = String(data[i][2]).trim();
      if (useYn === 'Y' && (!rank || rank === '')) {
        missing.push({
          seq: data[i][0],
          name: name,
          team: String(data[i][3]).trim()
        });
      }
    }
    return missing;
  } catch (e) {
    Logger.log('Error in getMissingRankWorkers: ' + e.message);
    return [];
  }
}

// ── 신규 API 함수: 특정 근무자의 계급 업데이트 ──────────────────────
function updateWorkerRank(seq, rank) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAMES.WORKERS);
    if (!sheet || sheet.getLastRow() < 2) return false;
    
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
    for (var i = 0; i < data.length; i++) {
      if (parseInt(data[i][0]) === parseInt(seq)) {
        sheet.getRange(i + 2, 2).setValue(rank);
        return true;
      }
    }
    return false;
  } catch (e) {
    Logger.log('Error in updateWorkerRank: ' + e.message);
    return false;
  }
}


// ============================================================
//  3단계: 웹앱 라우팅 & 조회
// ============================================================

function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};
  var queryString = (e && e.queryString) ? String(e.queryString).trim() : '';

  // ── 외부 API 요청 처리 (CORS 대응 JSON 반환) ─────────────────
  if (params.action) {
    var responseOutput = '';
    try {
      var responseData = null;
      if (params.action === 'getLatestSidByDate') {
        responseData = getLatestSidByDate(params.date);
      } else if (params.action === 'getScheduleBySid') {
        responseData = getScheduleBySid(params.sid);
      } else if (params.action === 'getScheduleByD') {
        responseData = decodeQrData(params.d);
      } else if (params.action === 'getMissingRankWorkers') {
        responseData = getMissingRankWorkers();
      } else if (params.action === 'getWebAppUrl') {
        responseData = getWebAppUrl();
      } else {
        responseData = { error: 'Unknown action: ' + params.action };
      }
      responseOutput = JSON.stringify(responseData);
    } catch (err) {
      responseOutput = JSON.stringify({ error: err.message });
    }
    return ContentService.createTextOutput(responseOutput)
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  if (params.mode === 'debug') {
    try {
      var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      var logText = '🔍 Spreadsheet Diagnostic Log\n============================\n\n';
      
      var sheets = ss.getSheets();
      for (var i = 0; i < sheets.length; i++) {
        var sh = sheets[i];
        var name = sh.getName();
        var lastRow = sh.getLastRow();
        var lastCol = sh.getLastColumn();
        logText += 'Sheet: ' + name + ' (Rows: ' + lastRow + ', Cols: ' + lastCol + ')\n';
        
        if (lastRow > 0 && lastCol > 0) {
          var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
          logText += 'Headers: ' + JSON.stringify(headers) + '\n';
          
          if (lastRow >= 2) {
            var sample = sh.getRange(2, 1, Math.min(5, lastRow - 1), lastCol).getValues();
            logText += 'Sample Data (up to 5 rows):\n' + JSON.stringify(sample, null, 2) + '\n';
          }
        } else {
          logText += 'Empty sheet\n';
        }
        
        var rules = (lastRow > 0 && lastCol > 0) ? sh.getRange(1, 1, lastRow, lastCol).getDataValidations() : [];
        var ruleList = [];
        for (var r = 0; r < rules.length; r++) {
          for (var c = 0; c < rules[r].length; c++) {
            var rule = rules[r][c];
            if (rule) {
              ruleList.push({
                cell: String.fromCharCode(65 + c) + (r + 1),
                criteria: rule.getCriteriaType().toString()
              });
            }
          }
        }
        if (ruleList.length > 0) {
          logText += 'Data Validation Rules Count: ' + ruleList.length + '\n';
        }
        logText += '--------------------------------------------------\n\n';
      }
      return ContentService.createTextOutput(logText)
        .setMimeType(ContentService.MimeType.TEXT);
    } catch (err) {
      return ContentService.createTextOutput('Error: ' + err.message)
        .setMimeType(ContentService.MimeType.TEXT);
    }
  }

  if (params.mode === 'testSeed') {
    try {
      var log = [];
      log.push('Starting testSeed...');
      
      var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      log.push('Opened Spreadsheet.');
      
      // 1) Settings
      var settingsSheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
      if (settingsSheet) {
        log.push('Clearing settingsSheet...');
        if (settingsSheet.getLastRow() >= 2) {
          settingsSheet.getRange(2, 1, settingsSheet.getLastRow() - 1, Math.max(1, settingsSheet.getLastColumn())).clearContent();
        }
        var settingsData = [
          ['근무일자',         '2026-06-24'],
          ['근무시작',         '09:00'],
          ['근무종료',         '09:00'],
          ['시간단위(분)',     '60'],
          ['시간대별필요인원', '1']
        ];
        log.push('Writing settingsData...');
        settingsSheet.getRange(2, 1, settingsData.length, 2).setValues(settingsData);
        log.push('Settings seeded successfully.');
      } else {
        log.push('Settings sheet not found!');
      }

      // 2) Workers
      var workersSheet = ss.getSheetByName(SHEET_NAMES.WORKERS);
      if (workersSheet) {
        log.push('Clearing workersSheet...');
        if (workersSheet.getLastRow() >= 2) {
          workersSheet.getRange(2, 1, workersSheet.getLastRow() - 1, Math.max(1, workersSheet.getLastColumn())).clearContent();
        }
        
        var workersData = [];
        var seq = 1;
        var ranks = ['소방령', '소방경', '소방위', '소방장', '소방교', '소방사'];

        for (var i = 1; i <= 20; i++) {
          var rank = ranks[Math.floor(Math.random() * ranks.length)];
          workersData.push([seq, rank, '1팀' + (i < 10 ? '0' + i : i), '1팀', 'Y']);
          seq++;
        }
        for (var i = 1; i <= 20; i++) {
          var rank = ranks[Math.floor(Math.random() * ranks.length)];
          workersData.push([seq, rank, '2팀' + (i < 10 ? '0' + i : i), '2팀', 'Y']);
          seq++;
        }
        for (var i = 1; i <= 20; i++) {
          var rank = ranks[Math.floor(Math.random() * ranks.length)];
          workersData.push([seq, rank, '3팀' + (i < 10 ? '0' + i : i), '3팀', 'Y']);
          seq++;
        }
        workersData.push([seq, '소방교', '사고자1', '1팀', 'Y']);
        seq++;
        workersData.push([seq, '소방사', '사고자2', '2팀', 'Y']);
        seq++;

        log.push('Writing workersData (count: ' + workersData.length + ')...');
        log.push('Sample row 0: ' + JSON.stringify(workersData[0]));
        log.push('Sample row 60: ' + JSON.stringify(workersData[60]));
        
        workersSheet.getRange(2, 1, workersData.length, 5).setValues(workersData);
        log.push('Workers seeded successfully.');
      } else {
        log.push('Workers sheet not found!');
      }

      // 3) Vacation
      var vacationSheet = ss.getSheetByName(SHEET_NAMES.VACATION);
      if (vacationSheet) {
        log.push('Clearing vacationSheet...');
        if (vacationSheet.getLastRow() >= 2) {
          vacationSheet.getRange(2, 1, vacationSheet.getLastRow() - 1, Math.max(1, vacationSheet.getLastColumn())).clearContent();
        }
        var vacationData = [
          ['2026-06-24', '사고자1', '휴직', '2026-06-24 ~ 2026-07-24', '장기휴직'],
          ['2026-06-24', '사고자2', '출장', '2026-06-24', '서울 출장']
        ];
        log.push('Writing vacationData...');
        vacationSheet.getRange(2, 1, vacationData.length, 5).setValues(vacationData);
        log.push('Vacation seeded successfully.');
      } else {
        log.push('Vacation sheet not found!');
      }

      // 4) Pinned
      var pinnedSheet = ss.getSheetByName(SHEET_NAMES.PINNED);
      if (pinnedSheet) {
        log.push('Clearing pinnedSheet...');
        if (pinnedSheet.getLastRow() >= 2) {
          pinnedSheet.getRange(2, 1, pinnedSheet.getLastRow() - 1, Math.max(1, pinnedSheet.getLastColumn())).clearContent();
        }
        log.push('Pinned cleared successfully.');
      } else {
        log.push('Pinned sheet not found!');
      }
      
      log.push('All steps completed successfully!');
      return HtmlService.createHtmlOutput('<h3>Success</h3><pre>' + log.join('\n') + '</pre>');
    } catch (e) {
      log.push('ERROR: ' + e.message);
      return HtmlService.createHtmlOutput('<h3>Failure</h3><pre>' + log.join('\n') + '</pre>');
    }
  }

  if (params.mode === 'admin') {
    var template = HtmlService.createTemplateFromFile('admin');
    return template.evaluate()
      .setTitle('근무표 관리자')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  var sid = '';
  if (params.sid) {
    sid = String(params.sid).trim();
  } 
  else if (params[''] && String(params['']).trim().length === 8) {
    sid = String(params['']).trim();
  }
  else if (queryString.length === 8) {
    sid = queryString;
  } 
  else {
    for (var key in params) {
      var trimmedKey = String(key).trim();
      if (trimmedKey.length === 8 && (params[key] === '' || params[key] === undefined)) {
        sid = trimmedKey;
        break;
      }
    }
  }

  if (params.d) {
    var template = HtmlService.createTemplateFromFile('view');
    template.sid = '';
    template.scheduleData = JSON.stringify(decodeQrData(params.d));
    return template.evaluate()
      .setTitle('근무표 조회')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (sid) {
    var template = HtmlService.createTemplateFromFile('view');
    template.sid = sid;
    template.scheduleData = JSON.stringify(getScheduleBySid(sid));
    return template.evaluate()
      .setTitle('근무표 조회')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  var template = HtmlService.createTemplateFromFile('index');
  return template.evaluate()
    .setTitle('근무표 QR 공유')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function getScheduleBySid(sid) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAMES.RESULTS);
    if (!sheet || sheet.getLastRow() < 2) return null;

    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === sid) {
        try {
          var parsed = JSON.parse(data[i][5]);
          
          // 구글 스프레드시트의 최신 정보(총원, 계급, 소속팀, 휴가 현황)를 실시간 반영
          var allWorkers = _readWorkers(ss);
          var vacationList = _readVacationList(ss, parsed.date);
          var activeTeam = _getTeamOnDuty(parsed.date);
          
          var totalCount = allWorkers.length;
          var vacationCount = vacationList.length;
          
          // 당번팀 소속 근무자 중 오늘 휴가(사고)가 아닌 전체 인원을 당번 수에 집계
          var onDutyCount = 0;
          for (var w = 0; w < allWorkers.length; w++) {
            var worker = allWorkers[w];
            if (worker.team === activeTeam) {
              var isOnVacation = false;
              for (var v = 0; v < vacationList.length; v++) {
                if (vacationList[v].name === worker.name) {
                  isOnVacation = true; break;
                }
              }
              if (!isOnVacation) {
                onDutyCount++;
              }
            }
          }
          var offDutyCount = totalCount - vacationCount - onDutyCount;

          // 순수 안전 타입으로만 데이터 재가공 (직렬화 크래시 원천 방지)
          var cleanedSchedule = [];
          if (parsed.schedule) {
            for (var s = 0; s < parsed.schedule.length; s++) {
              var slot = parsed.schedule[s];
              var cleanedWorkers = [];
              if (slot.workers) {
                for (var w = 0; w < slot.workers.length; w++) {
                  if (slot.workers[w]) {
                    cleanedWorkers.push({
                      name: String(slot.workers[w].name || ''),
                      rank: String(_findWorkerRank(allWorkers, slot.workers[w].name)),
                      team: String(_findWorkerTeam(allWorkers, slot.workers[w].name)),
                      isPinned: !!slot.workers[w].isPinned,
                      comment: String(slot.workers[w].comment || '')
                    });
                  }
                }
              }
              cleanedSchedule.push({
                time: String(slot.time || slot.label || ''),
                label: String(slot.label || slot.time || ''),
                isDay: !!slot.isDay,
                isSpecial: !!slot.isSpecial,
                workers: cleanedWorkers
              });
            }
          }

          var cleanedVacations = [];
          if (vacationList) {
            for (var v = 0; v < vacationList.length; v++) {
              var vac = vacationList[v];
              if (vac) {
                cleanedVacations.push({
                  name: String(vac.name || ''),
                  rank: String(vac.rank || _findWorkerRank(allWorkers, vac.name)),
                  type: String(vac.type || ''),
                  period: String(vac.period || ''),
                  comment: String(vac.comment || '')
                });
              }
            }
          }

          var cleanResult = {
            sid: String(sid),
            date: String(parsed.date || ''),
            title: String(parsed.title || ''),
            requiredPerSlot: parseInt(parsed.requiredPerSlot) || 1,
            intervalMin: parseInt(parsed.intervalMin) || 60,
            vacationList: cleanedVacations,
            schedule: cleanedSchedule,
            statsSummary: {
              totalCount: parseInt(totalCount) || 0,
              vacationCount: parseInt(vacationCount) || 0,
              onDutyCount: parseInt(onDutyCount) || 0,
              offDutyCount: parseInt(offDutyCount) || 0
            },
            isQrMode: !!parsed.isQrMode
          };
          
          return cleanResult;
        } catch (e) {
          Logger.log('⚠️ JSON 파싱 오류 (sid: ' + sid + '): ' + e.message);
          return null;
        }
      }
    }
  } catch (globalErr) {
    Logger.log('⚠️ getScheduleBySid 치명적 오류: ' + globalErr.message);
    return { error: globalErr.message };
  }
  return null;
}

// ── 조회 및 실시간 자동 생성 처리 기능 ──────────────────────
function getLatestSidByDate(dateStr) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAMES.RESULTS);
  if (!sheet) return null;

  var normDate = _normalizeDate(dateStr);
  
  if (sheet.getLastRow() >= 2) {
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
    for (var i = data.length - 1; i >= 0; i--) {
      var rowDate = _normalizeDate(data[i][2]);
      if (rowDate === normDate) {
        return String(data[i][0]).trim(); // 기존 생성된 sid 반환
      }
    }
  }

  // 기존에 생성된 근무표가 없다면 실시간 자동 생성 진행!
  try {
    var result = generateSchedule(normDate);
    if (result && result.sid) {
      return result.sid;
    }
  } catch (e) {
    Logger.log('실시간 근무표 자동 생성 오류 (날짜: ' + normDate + '): ' + e.message);
  }

  return null;
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}

function getPinnedMapForView(dateStr) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return _readPinnedMap(ss, dateStr);
}

function _normalizeDate(value) {
  if (value instanceof Date || (value && typeof value.getMonth === 'function')) {
    try {
      return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } catch(err) {
      return Utilities.formatDate(value, "GMT+09:00", 'yyyy-MM-dd');
    }
  }
  var str = String(value).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(str)) {
    return str;
  }
  
  // 구글 표준 긴 날짜 형식 문자열 파싱 지원 (예: Wed Jun 24 2026 00:00:00 GMT+0900)
  var parts = str.split(' ');
  if (parts.length >= 4) {
    var months = { Jan:'01', Feb:'02', Mar:'03', Apr:'04', May:'05', Jun:'06', Jul:'07', Aug:'08', Sep:'09', Oct:'10', Nov:'11', Dec:'12' };
    var yyyy = '', mm = '', dd = '';
    
    // 포맷 1: Wed Jun 24 2026
    if (months[parts[1]] && !isNaN(parts[2]) && !isNaN(parts[3])) {
      yyyy = parts[3];
      mm = months[parts[1]];
      dd = String(parts[2]).padStart(2, '0');
    }
    // 포맷 2: 24 Jun 2026
    else if (months[parts[2]] && !isNaN(parts[1]) && !isNaN(parts[3])) {
      yyyy = parts[3];
      mm = months[parts[2]];
      dd = String(parts[1]).padStart(2, '0');
    }
    
    if (yyyy && mm && dd) {
      return yyyy + '-' + mm + '-' + dd;
    }
  }

  try {
    var d = new Date(str);
    if (!isNaN(d.getTime())) {
      try {
        return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } catch(err) {
        return Utilities.formatDate(d, "GMT+09:00", 'yyyy-MM-dd');
      }
    }
  } catch (e) {}
  return str;
}

function _normalizeTime(value) {
  if (value instanceof Date || (value && typeof value.getHours === 'function')) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'HH:mm');
  }
  var str = String(value).trim();
  var match = str.match(/(\d{1,2}):(\d{2})/);
  if (match) {
    var h = parseInt(match[1], 10);
    var m = parseInt(match[2], 10);
    return (h < 10 ? '0' + h : '' + h) + ':' + (m < 10 ? '0' + m : '' + m);
  }
  try {
    var d = new Date(str);
    if (!isNaN(d.getTime())) {
      return Utilities.formatDate(d, Session.getScriptTimeZone(), 'HH:mm');
    }
  } catch (e) {}
  if (str.length === 4 && str.indexOf(':') === 1) {
    str = '0' + str;
  }
  return str;
}

// ── 3일 주기 당번 팀 계산 ──────────────────────────────────
function _getTeamOnDuty(dateStr) {
  // dateStr is 'YYYY-MM-DD'
  var normDate = _normalizeDate(dateStr);
  var parts = normDate.split('-');
  var targetDate = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
  
  // 기준일: 2026-06-25일은 1팀 근무일
  var baseDate = new Date(2026, 5, 25);
  
  targetDate.setHours(12, 0, 0, 0);
  baseDate.setHours(12, 0, 0, 0);
  
  var diffTime = targetDate.getTime() - baseDate.getTime();
  var diffDays = Math.round(diffTime / (1000 * 60 * 60 * 24));
  
  var mod = diffDays % 3;
  if (mod < 0) {
    mod += 3;
  }
  
  if (mod === 0) return '1팀';
  if (mod === 1) return '2팀';
  return '3팀';
}

// ── 특정 팀의 직전 VirtualBaseName 조회 ────────────────────────
function _getLastVirtualBaseNameForTeam(ss, teamName) {
  var sheet = ss.getSheetByName(SHEET_NAMES.RESULTS);
  if (!sheet || sheet.getLastRow() < 2) return null;

  var data = sheet.getRange(2, 6, sheet.getLastRow() - 1, 1).getValues();
  for (var i = data.length - 1; i >= 0; i--) {
    try {
      var jsonStr = data[i][0];
      if (!jsonStr) continue;
      var obj = JSON.parse(jsonStr);
      
      var prevTeam = null;
      if (obj.schedule) {
        for (var s = 0; s < obj.schedule.length; s++) {
          var slot = obj.schedule[s];
          if (slot.workers && slot.workers.length > 0) {
            for (var w = 0; w < slot.workers.length; w++) {
              if (!slot.workers[w].isPinned) {
                prevTeam = slot.workers[w].team;
                break;
              }
            }
          }
          if (prevTeam) break;
        }
      }
      
      if (prevTeam === teamName && obj.nextVirtualBaseName) {
        return obj.nextVirtualBaseName;
      }
    } catch (e) {}
  }
  return null;
}

// ── 특정 날짜가 휴가 기간 범위 내에 속하는지 판별 ─────────────────
function _isDateInPeriod(targetDateStr, dateColVal, periodColVal) {
  var targetStr = _normalizeDate(targetDateStr);
  var target = new Date(targetStr);
  target.setHours(12, 0, 0, 0);
  
  var periodStr = String(periodColVal).trim();
  if (!periodStr) {
    return _normalizeDate(dateColVal) === targetStr;
  }
  
  // 1) 'YYYY-MM-DD ~ YYYY-MM-DD' 또는 'YYYY-MM-DD~YYYY-MM-DD' 범위 형태 검사
  var parts = periodStr.split('~');
  if (parts.length === 2) {
    var startStr = _normalizeDate(parts[0].trim());
    var endStr = _normalizeDate(parts[1].trim());
    
    var start = new Date(startStr);
    var end = new Date(endStr);
    
    start.setHours(12, 0, 0, 0);
    end.setHours(12, 0, 0, 0);
    
    return target.getTime() >= start.getTime() && target.getTime() <= end.getTime();
  }
  
  // 2) 단일 날짜 형태 검사
  if (/^\d{4}-\d{2}-\d{2}$/.test(periodStr)) {
    return _normalizeDate(periodStr) === targetStr;
  }
  
  // 3) 그 외 텍스트 등은 기본 등록 날짜와 일치하는지만 검사
  return _normalizeDate(dateColVal) === targetStr;
}

/**
 * QR코드에 포함된 압축 데이터를 파싱하여 원래의 JSON 스케줄 데이터 객체 구조로 리빌딩한다.
 * @param {string} d 압축 데이터 문자열
 * @return {Object|null}
 */
function decodeQrData(d) {
  try {
    if (!d) return null;
    
    // 포맷: date|team|total,vac,onduty,offduty|vacName,vacRank,vacType,vacPeriod,vacComment;...|slotLabel:wName.wRank.wTeam.isPinned(1/0).wComment,wName...;...
    var parts = d.split('|');
    if (parts.length < 5) return null;
    
    var date = parts[0];
    var team = parts[1];
    
    var statsParts = parts[2].split(',');
    var statsSummary = null;
    if (statsParts.length >= 4) {
      statsSummary = {
        totalCount: parseInt(statsParts[0], 10) || 0,
        vacationCount: parseInt(statsParts[1], 10) || 0,
        onDutyCount: parseInt(statsParts[2], 10) || 0,
        offDutyCount: parseInt(statsParts[3], 10) || 0
      };
    }
    
    var vacationList = [];
    if (parts[3]) {
      var vacEntries = parts[3].split(';');
      for (var i = 0; i < vacEntries.length; i++) {
        var vParts = vacEntries[i].split(',');
        if (vParts[0]) {
          vacationList.push({
            name: vParts[0] || '',
            rank: vParts[1] || '',
            type: vParts[2] || '',
            period: vParts[3] || '',
            comment: vParts[4] || ''
          });
        }
      }
    }
    
    var schedule = [];
    if (parts[4]) {
      var slotEntries = parts[4].split(';');
      for (var i = 0; i < slotEntries.length; i++) {
        var sParts = slotEntries[i].split(':');
        if (sParts.length >= 2) {
          var label = sParts[0];
          var workersStr = sParts[1];
          var workers = [];
          if (workersStr) {
            var wEntries = workersStr.split(',');
            for (var j = 0; j < wEntries.length; j++) {
              var wParts = wEntries[j].split('.');
              if (wParts[0]) {
                workers.push({
                  name: wParts[0] || '',
                  rank: wParts[1] || '',
                  team: wParts[2] || '',
                  isPinned: wParts[3] === '1',
                  comment: wParts[4] || ''
                });
              }
            }
          }
          schedule.push({
            label: label,
            isSpecial: (label === '12~14' || label === '00~03'),
            workers: workers
          });
        }
      }
    }
    
    return {
      sid: '',
      date: date,
      title: '신평119안전센터 소내근무표 (' + team + ' 당번)',
      statsSummary: statsSummary,
      vacationList: vacationList,
      schedule: schedule,
      isQrMode: true
    };
  } catch (err) {
    Logger.log('Error decoding QR data: ' + err.message);
    return null;
  }
}

// ── 외부 API 요청 처리 (POST) ──────────────────────────────
function doPost(e) {
  var params = {};
  var postDataContent = '';
  
  if (e && e.postData && e.postData.contents) {
    postDataContent = e.postData.contents;
    try {
      params = JSON.parse(postDataContent);
    } catch (err) {
      params = e.parameter || {};
    }
  } else if (e && e.parameter) {
    params = e.parameter;
  }

  var action = params.action || (e && e.parameter ? e.parameter.action : '');
  var responseData = null;
  
  try {
    if (action === 'setupSheets') {
      responseData = setupSheets();
    } else if (action === 'seedSampleData') {
      responseData = seedSampleData();
    } else if (action === 'generateSchedule') {
      var dateVal = params.date || (e && e.parameter ? e.parameter.date : '');
      responseData = generateSchedule(dateVal);
    } else if (action === 'updateWorkerRank') {
      var seqVal = params.seq || (e && e.parameter ? e.parameter.seq : '');
      var rankVal = params.rank || (e && e.parameter ? e.parameter.rank : '');
      var success = updateWorkerRank(seqVal, rankVal);
      responseData = { success: success };
    } else {
      responseData = { error: 'Unknown POST action: ' + action };
    }
  } catch (err) {
    responseData = { error: err.message };
  }
  
  return ContentService.createTextOutput(JSON.stringify(responseData))
    .setMimeType(ContentService.MimeType.JSON);
}
