// ============================================================
// [공통 설정]
// ============================================================
var SHEET_NAME = 'mails';
var HEADERS = ["Date", "Sender", "Recipient", "Labels", "Subject", "Urgency",
               "Summary", "Snippet", "Note", "Thread ID", "Full Content"];
var NUM_COLS = HEADERS.length;
var COL = {
  DATE: 1, SENDER: 2, RECIPIENT: 3, LABELS: 4, SUBJECT: 5, URGENCY: 6,
  SUMMARY: 7, SNIPPET: 8, NOTE: 9, THREAD_ID: 10, FULL: 11
};

// 💡 F열(Urgency) 분류 규칙: 제목·본문·라벨(D열)을 합친 글자에서 검사 (위에서부터 먼저 매칭되는 것 적용)
var URGENCY_RULES = [
  { pattern: /URGENT|긴급|AOG|결함|CANCELLED|중요/i, label: "🔴 긴급" },
  { pattern: /요청|의뢰|문의|지연|Recovery/i,        label: "🟡 요청/확인" },
  { pattern: /공지|부고|결혼|축하|보고/i,             label: "⚪ 단순공지" }
];
var URGENCY_DEFAULT = "일반";

// 💡 G열(Summary) =AI() 프롬프트 (수식에서 K열 셀 뒤에 & 로 이어 붙임)
var SUMMARY_PROMPT = "는 메일의 제목과 본문 일부를 추출한 텍스트인데 '음슴체' 2문장으로 요약해 줘";

/**
 * [메뉴 생성 함수]
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Email Tools')
      .addItem('Get Emails (Popup)', 'fetchUnreadEmails')
      .addItem('Freeze AI() Summaries to Values', 'freezeAiSummaries')
      .addSeparator()
      .addItem('Clear Emails Contents', 'clearSpecificFormats')
      .addToUi();
}

/**
 * [이메일 수집 메인 함수] - ⚡초고속 50개씩 끊어치기 튜닝 버전 + 긴급도 분류(값) + AI() 요약⚡
 */
function fetchUnreadEmails() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

  // --- 사용자 설정 변수 ---
  var num_of_mails_in_thread = 1;
  var num_of_maxMails_collected = 50;
  var len_of_mailContents = 400;
  var recipient_to_be_Bold = 'thlim@koreanair.com';

  // --- STEP 1: 날짜 입력 팝업 ---
  var inputDates = ui.prompt(
    '조회 기간 입력',
    '시작일과 종료일을 물결(~) 또는 쉼표(,)로 구분하여 입력하세요.\n(예: 2024/01/01 ~ 2024/01/02)',
    ui.ButtonSet.OK_CANCEL
  );
  if (inputDates.getSelectedButton() != ui.Button.OK) return;

  var rawInput = inputDates.getResponseText().trim();
  var dateParts = rawInput.split(/[\~,\s]+/).map(function(s) { return s.trim(); }).filter(Boolean);

  if (dateParts.length < 2) {
    ui.alert("시작일과 종료일을 모두 입력해주세요.\n(예: 2024/01/01 ~ 2024/01/02)");
    return;
  }

  var startDate = new Date(dateParts[0]);
  var endDate = new Date(dateParts[1]);

  if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
    ui.alert("잘못된 날짜 형식입니다. 'YYYY/MM/DD' 형식으로 다시 실행해주세요.");
    return;
  }

  var startTimestamp = Math.floor(startDate.getTime() / 1000);
  var endTimestamp = Math.floor(endDate.getTime() / 1000);

  // --- STEP 2: 시트 헤더 설정 + 기존 데이터 정리 ---
  var isNewSheet = (sheet.getLastRow() === 0 || sheet.getRange("A1").getValue() === "");
  sheet.getRange(1, 1, 1, NUM_COLS).setValues([HEADERS]).setFontWeight("bold");
  if (isNewSheet) sheet.setFrozenRows(1);

  // 💡 이전 실행에서 남은 =AI() 수식 중 결과가 나온 것은 값으로 고정 (정렬로 섞이기 전에 처리)
  var freezeResult = freezeAiSummaries_(sheet);
  // 💡 F열에 남아 있는 MAP 배열 수식 제거 + 비어 있는 긴급도 채우기 (기존 데이터 마이그레이션)
  fillMissingUrgency_(sheet);

  // --- STEP 3: Gmail 검색 실행 ---
  var query = 'in:inbox -is:starred after:' + startTimestamp + ' before:' + endTimestamp;
  var threads = GmailApp.search(query, 0, num_of_maxMails_collected);

  if (threads.length === 0) {
    ui.alert("해당 기간의 받은편지함에 처리할 메일이 없습니다. (모두 처리되었거나 조건에 맞는 메일이 없음)");
    return;
  }

  var threadsToArchive = [];
  var threadBundles = [];

  // --- STEP 4 & 5 통합: 데이터 추출 및 즉시 그룹화 ---
  var chunkSize = 50;

  for (var c = 0; c < threads.length; c += chunkSize) {
    var threadChunk = threads.slice(c, c + chunkSize);
    var messages2D = GmailApp.getMessagesForThreads(threadChunk);

    for (var i = 0; i < threadChunk.length; i++) {
      var thread = threadChunk[i];
      var messages = messages2D[i];
      threadsToArchive.push(thread);

      var labels = thread.getLabels();
      var labelString = labels.map(function(l) { return l.getName(); }).join(", ");
      var bundle = [];

      var startIndex = Math.max(0, messages.length - num_of_mails_in_thread);
      for (var j = startIndex; j < messages.length; j++) {
        var msg = messages[j];
        var cleanSnippet = msg.getPlainBody().replace(/[\r\n]+/g, ' ').trim().substring(0, len_of_mailContents);

        bundle.push({
          date: msg.getDate(),
          sender: msg.getFrom(),
          recipient: msg.getTo(),
          labels: labelString,
          subject: msg.getSubject(),
          snippet: cleanSnippet,
          threadId: thread.getId() // 💡 메일 그룹핑을 위한 Gmail 고유 Thread ID
        });
      }
      bundle.reverse();
      if (bundle.length > 0) {
        threadBundles.push(bundle);
      }
    }
  }

  threadBundles.sort(function(bundleA, bundleB) {
    var subjectA = bundleA[0].subject.toUpperCase();
    var subjectB = bundleB[0].subject.toUpperCase();
    return subjectA < subjectB ? -1 : (subjectA > subjectB ? 1 : 0);
  });

  // --- STEP 6: A~K열 전체 행 데이터 + 서식 준비 (수식 없이 값으로) ---
  var rows = [];          // A~K열 값
  var weights = [];
  var styles = [];
  var colors = [];

  threadBundles.forEach(function(bundle) {
    bundle.forEach(function(msg) {
      // 💡 K열: 기존 CONCATENATE(E$1,": ",E,char(10),H$1,": ",H) 수식의 결과를 값으로 직접 생성
      var fullContent = HEADERS[COL.SUBJECT - 1] + ": " + msg.subject + "\n" +
                        HEADERS[COL.SNIPPET - 1] + ": " + msg.snippet;

      rows.push([
        msg.date,
        safeText_(msg.sender),
        safeText_(msg.recipient),
        safeText_(msg.labels),
        safeText_(msg.subject),
        classifyUrgency(msg.subject, msg.snippet, msg.labels), // 💡 F열: 제목·본문·라벨 정규식 분류 결과를 값으로
        "",                                        // G열: 정렬 후 STEP 10에서 =AI() 수식 입력
        safeText_(msg.snippet),
        "",
        msg.threadId,
        safeText_(fullContent)
      ]);

      var isTarget = msg.recipient.indexOf(recipient_to_be_Bold) !== -1;
      weights.push(fillRow_(isTarget ? "bold" : "normal"));
      styles.push(fillRow_(isTarget ? "italic" : "normal"));
      colors.push(fillRow_(isTarget ? "blue" : "black"));
    });
  });

  // --- STEP 7: 스프레드시트에 한 번에 쓰기 ---
  var startRow = getLastDataRow(sheet, 1) + 1;
  var numRows = rows.length;
  var dataRange = sheet.getRange(startRow, 1, numRows, NUM_COLS);
  dataRange.setValues(rows);
  dataRange.setBackground(null);
  dataRange.setFontWeights(weights);
  dataRange.setFontStyles(styles);
  dataRange.setFontColors(colors);
  sheet.getRange(startRow, COL.DATE, numRows, 1).setNumberFormat("yyyy-mm-dd hh:mm:ss");

  // --- STEP 8: Gmail 후처리 (보관) ---
  var batchSize = 100;
  for (var k = 0; k < threadsToArchive.length; k += batchSize) {
    GmailApp.moveThreadsToArchive(threadsToArchive.slice(k, k + batchSize));
  }

  // --- STEP 9: 데이터 자동 정렬 (Thread ID 오름차순 -> Date 오름차순) ---
  var finalDataRow = getLastDataRow(sheet, 1);
  if (finalDataRow > 1) {
    sheet.getRange(2, 1, finalDataRow - 1, NUM_COLS).sort([
      {column: COL.THREAD_ID, ascending: true}, // 1순위 정렬: J열 (Thread ID)
      {column: COL.DATE, ascending: true}       // 2순위 정렬: A열 (Date)
    ]);
  }

  // --- STEP 10: G열(Summary)이 빈 행에 =AI() 수식 입력 ---
  // 💡 정렬이 끝난 뒤에 넣어야 수식의 행 참조(K열)가 정렬로 꼬이지 않음
  //    수식 입력만 하고 바로 종료 → AI 생성은 시트에서 비동기로 진행되므로 스크립트 실행 시간에 포함되지 않음
  var aiFormulaCount = applyAiFormulas_(sheet);

  var message = "작업 완료! " + numRows + "개의 메일을 처리했습니다.\n" +
                "같은 메일 그룹끼리 일시에 따라 자동으로 정렬되었습니다.\n\n" +
                "• =AI() 요약 수식 입력: " + aiFormulaCount + "건\n";
  if (freezeResult.frozen > 0 || freezeResult.pending > 0) {
    message += "• 이전 실행분 값 고정: " + freezeResult.frozen + "건 (아직 생성 중/오류: " + freezeResult.pending + "건)\n";
  }
  message += "\n요약 생성이 끝나면 다음 실행 때 자동으로, 또는 'Freeze AI() Summaries to Values' 메뉴로 값으로 고정됩니다.";
  ui.alert(message);
}

/**
 * [긴급도 분류]
 * 제목이 비어 있으면 "", 아니면 "제목 + 본문 + 라벨"을 이어 붙인 글자에 대해 규칙을 순서대로 검사
 * @param {string} subject  E열 Subject
 * @param {string} snippet  H열 Snippet
 * @param {string} labels   D열 Labels (예: "업무, 정비문서") — 생략 가능
 */
function classifyUrgency(subject, snippet, labels) {
  subject = (subject == null) ? "" : String(subject);
  if (subject === "") return "";
  var text = [subject, snippet, labels]
      .map(function(v) { return v == null ? "" : String(v); })
      .join(" ");
  for (var i = 0; i < URGENCY_RULES.length; i++) {
    if (URGENCY_RULES[i].pattern.test(text)) return URGENCY_RULES[i].label;
  }
  return URGENCY_DEFAULT;
}

/**
 * [F열 마이그레이션]
 * F열에 수식(예: 기존 MAP 배열 수식)이 남아 있으면 F열 내용을 지우고,
 * 비어 있는 행의 긴급도를 E열(Subject)/H열(Snippet)/D열(Labels) 기준으로 값으로 채움
 */
function fillMissingUrgency_(sheet) {
  var lastRow = getLastDataRow(sheet, 1);
  if (lastRow <= 1) return;
  var n = lastRow - 1;

  var urgencyRange = sheet.getRange(2, COL.URGENCY, n, 1);
  var hasFormula = urgencyRange.getFormulas().some(function(r) { return r[0] !== ""; });
  if (hasFormula) {
    sheet.getRange(2, COL.URGENCY, sheet.getMaxRows() - 1, 1).clearContent();
  }

  var current = urgencyRange.getValues();
  var subjects = sheet.getRange(2, COL.SUBJECT, n, 1).getValues();
  var snippets = sheet.getRange(2, COL.SNIPPET, n, 1).getValues();
  var labels = sheet.getRange(2, COL.LABELS, n, 1).getValues();
  var changed = false;

  var output = current.map(function(r, i) {
    if (r[0] !== "") return [r[0]];
    var urgency = classifyUrgency(subjects[i][0], snippets[i][0], labels[i][0]);
    if (urgency !== "") changed = true;
    return [urgency];
  });

  if (changed) urgencyRange.setValues(output);
}

/**
 * [=AI() 수식 입력]
 * G열(Summary)이 비어 있고 K열(Full Content)이 있는 행에만 =AI() 수식을 넣음
 * @return {number} 수식을 넣은 행 수
 */
function applyAiFormulas_(sheet) {
  var lastRow = getLastDataRow(sheet, 1);
  if (lastRow <= 1) return 0;
  var n = lastRow - 1;

  var summaryRange = sheet.getRange(2, COL.SUMMARY, n, 1);
  var formulas = summaryRange.getFormulas();
  var values = summaryRange.getValues();
  var fullContents = sheet.getRange(2, COL.FULL, n, 1).getValues();
  var fullColLetter = columnToLetter_(COL.FULL);
  var count = 0;

  for (var i = 0; i < n; i++) {
    if (formulas[i][0] === "" && values[i][0] === "" && fullContents[i][0] !== "") {
      var row = i + 2;
      sheet.getRange(row, COL.SUMMARY).setFormula('=AI(' + fullColLetter + row + ' & "' + SUMMARY_PROMPT + '")');
      count++;
    }
  }
  return count;
}

/**
 * [메뉴] G열의 =AI() 수식 중 결과가 나온 셀을 값으로 고정
 */
function freezeAiSummaries() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var result = freezeAiSummaries_(sheet);
  SpreadsheetApp.getUi().alert("값 고정: " + result.frozen + "건\n아직 생성 중이거나 오류: " + result.pending + "건");
}

function freezeAiSummaries_(sheet) {
  var result = { frozen: 0, pending: 0 };
  var lastRow = getLastDataRow(sheet, 1);
  if (lastRow <= 1) return result;

  var summaryRange = sheet.getRange(2, COL.SUMMARY, lastRow - 1, 1);
  var formulas = summaryRange.getFormulas();
  var displays = summaryRange.getDisplayValues();

  for (var i = 0; i < formulas.length; i++) {
    if (!/^=\s*AI\s*\(/i.test(formulas[i][0])) continue;
    var value = displays[i][0];
    if (isAiResultReady_(value)) {
      sheet.getRange(i + 2, COL.SUMMARY).setValue(safeText_(value));
      result.frozen++;
    } else {
      result.pending++;
    }
  }
  return result;
}

function isAiResultReady_(value) {
  if (value === "") return false;
  if (value.charAt(0) === "#") return false;               // #ERROR!, #N/A, #NAME? 등
  if (/^(loading|generating)/i.test(value)) return false;  // 생성 중 표시
  return true;
}

/**
 * [실제 데이터의 마지막 행 반환 헬퍼 함수]
 * 시트 하단의 빈 서식 행 때문에 getLastRow()가 잘못된 값을 뱉는 것을 방지
 */
function getLastDataRow(sheet, colIndex) {
  var values = sheet.getRange(1, colIndex, sheet.getMaxRows(), 1).getValues();
  for (var i = values.length - 1; i >= 0; i--) {
    if (values[i][0] !== "") {
      return i + 1;
    }
  }
  return 1;
}

/**
 * [텍스트 안전 처리] "=" 또는 "+"로 시작하는 문자열이 수식/숫자로 해석되지 않도록 ' 접두
 */
function safeText_(value) {
  var text = (value == null) ? "" : String(value);
  return /^[=+]/.test(text) ? "'" + text : text;
}

function fillRow_(value) {
  var row = [];
  for (var i = 0; i < NUM_COLS; i++) row.push(value);
  return row;
}

function columnToLetter_(column) {
  var letter = "";
  while (column > 0) {
    var mod = (column - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    column = Math.floor((column - 1) / 26);
  }
  return letter;
}

/**
 * [데이터 초기화 함수]
 * 헤더를 제외한 A~K열 내용 삭제 (F열도 이제 수식이 아닌 값이므로 함께 삭제)
 */
function clearSpecificFormats() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var lastDataRow = getLastDataRow(sheet, 1); // A열 기준 마지막 데이터 탐색

  if (lastDataRow <= 1) return;

  var range = sheet.getRange(2, 1, lastDataRow - 1, Math.min(NUM_COLS, sheet.getMaxColumns()));
  range.clearContent();
  range.setBackground(null);
  range.setFontWeight('normal');
  range.setFontStyle('normal');
  range.setVerticalAlignment('top');
  range.setHorizontalAlignment('left');
  range.setWrap(true);
}
