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

// 💡 F열(Urgency) 분류 규칙: 기존 MAP/REGEXMATCH 수식과 동일한 순서·패턴 (위에서부터 먼저 매칭되는 것 적용)
var URGENCY_RULES = [
  { pattern: /URGENT|긴급|AOG|결함|CANCELLED|중요/i, label: "🔴 긴급" },
  { pattern: /요청|의뢰|문의|지연|Recovery/i,        label: "🟡 요청/확인" },
  { pattern: /공지|부고|결혼|축하|보고/i,             label: "⚪ 단순공지" }
];
var URGENCY_DEFAULT = "일반";

// 💡 G열(Summary) 요약 설정
var SUMMARY_PROMPT = "메일의 제목과 내용을 보고 2문장으로 요약해 줘";
var GEMINI_MODEL_DEFAULT = 'gemini-2.5-flash-lite'; // 스크립트 속성 GEMINI_MODEL 로 변경 가능
var GEMINI_BATCH_SIZE = 25;                          // 요청 1건에 담을 메일 수 (요청들은 병렬 전송)

/**
 * [메뉴 생성 함수]
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Email Tools')
      .addItem('Get Emails (Popup)', 'fetchUnreadEmails')
      .addItem('Freeze AI() Summaries to Values', 'freezeAiSummaries')
      .addItem('Set Gemini API Key', 'setGeminiApiKey')
      .addSeparator()
      .addItem('Clear Emails Contents', 'clearSpecificFormats')
      .addToUi();
}

/**
 * [이메일 수집 메인 함수] - ⚡초고속 50개씩 끊어치기 튜닝 버전 + 긴급도 분류(값) + Gemini 일괄 요약⚡
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
  var fullContents = [];  // Gemini 요약 입력 (K열과 동일)
  var weights = [];
  var styles = [];
  var colors = [];

  threadBundles.forEach(function(bundle) {
    bundle.forEach(function(msg) {
      // 💡 K열: 기존 CONCATENATE(E$1,": ",E,char(10),H$1,": ",H) 수식의 결과를 값으로 직접 생성
      var fullContent = HEADERS[COL.SUBJECT - 1] + ": " + msg.subject + "\n" +
                        HEADERS[COL.SNIPPET - 1] + ": " + msg.snippet;
      fullContents.push(fullContent);

      rows.push([
        msg.date,
        safeText_(msg.sender),
        safeText_(msg.recipient),
        safeText_(msg.labels),
        safeText_(msg.subject),
        classifyUrgency(msg.subject, msg.snippet), // 💡 F열: 정규식 분류 결과를 값으로
        "",                                        // G열: 아래 STEP 7에서 채움
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

  // --- STEP 7: G열(Summary) Gemini API 일괄 요약 (메일 25개당 요청 1건, 병렬 전송) ---
  var summaryResult = summarizeWithGemini_(fullContents);
  var apiSummarized = 0;
  summaryResult.summaries.forEach(function(summary, idx) {
    if (summary) {
      rows[idx][COL.SUMMARY - 1] = safeText_(summary);
      apiSummarized++;
    }
  });

  // --- STEP 8: 스프레드시트에 한 번에 쓰기 ---
  var startRow = getLastDataRow(sheet, 1) + 1;
  var numRows = rows.length;
  var dataRange = sheet.getRange(startRow, 1, numRows, NUM_COLS);
  dataRange.setValues(rows);
  dataRange.setBackground(null);
  dataRange.setFontWeights(weights);
  dataRange.setFontStyles(styles);
  dataRange.setFontColors(colors);
  sheet.getRange(startRow, COL.DATE, numRows, 1).setNumberFormat("yyyy-mm-dd hh:mm:ss");

  // --- STEP 9: Gmail 후처리 (보관) ---
  var batchSize = 100;
  for (var k = 0; k < threadsToArchive.length; k += batchSize) {
    GmailApp.moveThreadsToArchive(threadsToArchive.slice(k, k + batchSize));
  }

  // --- STEP 10: 데이터 자동 정렬 (Thread ID 오름차순 -> Date 오름차순) ---
  var finalDataRow = getLastDataRow(sheet, 1);
  if (finalDataRow > 1) {
    sheet.getRange(2, 1, finalDataRow - 1, NUM_COLS).sort([
      {column: COL.THREAD_ID, ascending: true}, // 1순위 정렬: J열 (Thread ID)
      {column: COL.DATE, ascending: true}       // 2순위 정렬: A열 (Date)
    ]);
  }

  // --- STEP 11: API 요약이 비어 있는 행만 =AI() 수식으로 대체 (정렬 후에 넣어 행 참조가 꼬이지 않게) ---
  var fallbackCount = applyAiFormulaFallback_(sheet);

  var message = "작업 완료! " + numRows + "개의 메일을 처리했습니다.\n" +
                "같은 메일 그룹끼리 일시에 따라 자동으로 정렬되었습니다.\n\n" +
                "• Gemini API 요약: " + apiSummarized + "건\n";
  if (fallbackCount > 0) {
    message += "• =AI() 수식으로 대체: " + fallbackCount + "건 (결과가 나오면 다음 실행 때 또는 " +
               "'Freeze AI() Summaries to Values' 메뉴로 값 고정)\n";
  }
  if (freezeResult.frozen > 0) {
    message += "• 이전 =AI() 결과 값 고정: " + freezeResult.frozen + "건\n";
  }
  if (summaryResult.error) {
    message += "\n⚠️ Gemini API: " + summaryResult.error;
  }
  ui.alert(message);
}

/**
 * [긴급도 분류]
 * 기존 F열 수식과 동일: 제목이 비어 있으면 "", 아니면 "제목 + 공백 + 본문"에 대해 규칙을 순서대로 검사
 */
function classifyUrgency(subject, snippet) {
  subject = (subject == null) ? "" : String(subject);
  if (subject === "") return "";
  var text = subject + " " + (snippet == null ? "" : String(snippet));
  for (var i = 0; i < URGENCY_RULES.length; i++) {
    if (URGENCY_RULES[i].pattern.test(text)) return URGENCY_RULES[i].label;
  }
  return URGENCY_DEFAULT;
}

/**
 * [F열 마이그레이션]
 * F열에 수식(예: 기존 MAP 배열 수식)이 남아 있으면 F열 내용을 지우고,
 * 비어 있는 행의 긴급도를 E열(Subject)/H열(Snippet) 기준으로 값으로 채움
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
  var changed = false;

  var output = current.map(function(r, i) {
    if (r[0] !== "") return [r[0]];
    var urgency = classifyUrgency(subjects[i][0], snippets[i][0]);
    if (urgency !== "") changed = true;
    return [urgency];
  });

  if (changed) urgencyRange.setValues(output);
}

/**
 * [Gemini API 일괄 요약]
 * 메일 여러 개를 요청 1건에 묶고(JSON 응답), 요청들은 UrlFetchApp.fetchAll 로 병렬 전송
 * @return {{summaries: string[], error: string}} summaries[i] 는 실패 시 ""
 */
function summarizeWithGemini_(contents) {
  var summaries = contents.map(function() { return ""; });
  if (contents.length === 0) return { summaries: summaries, error: "" };

  var props = PropertiesService.getScriptProperties();
  var apiKey = props.getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    return { summaries: summaries, error: "API 키가 없어 =AI() 수식으로 대체했습니다. ('Set Gemini API Key' 메뉴에서 설정)" };
  }
  var model = props.getProperty('GEMINI_MODEL') || GEMINI_MODEL_DEFAULT;
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent';

  var systemText =
    "너는 업무 메일 요약 도우미야. 입력으로 [번호]가 붙은 여러 통의 메일(Subject, Snippet)이 주어진다.\n" +
    "각 메일을 서로 독립적으로, 다음 지시에 따라 한국어로 요약해: \"" + SUMMARY_PROMPT + "\"\n" +
    "입력의 번호를 id 로 그대로 사용하고, 모든 메일에 대해 빠짐없이 결과를 반환해.";

  var requests = [];
  for (var s = 0; s < contents.length; s += GEMINI_BATCH_SIZE) {
    var userText = contents.slice(s, s + GEMINI_BATCH_SIZE).map(function(content, k) {
      return "[" + (s + k) + "]\n" + content;
    }).join("\n\n");

    var payload = {
      systemInstruction: { parts: [{ text: systemText }] },
      contents: [{ role: "user", parts: [{ text: userText }] }],
      generationConfig: {
        temperature: 0.2,
        responseMimeType: "application/json",
        responseSchema: {
          type: "ARRAY",
          items: {
            type: "OBJECT",
            properties: {
              id: { type: "INTEGER" },
              summary: { type: "STRING" }
            },
            required: ["id", "summary"]
          }
        }
      }
    };

    requests.push({
      url: url,
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-goog-api-key': apiKey },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
  }

  var errors = [];
  var responses;
  try {
    responses = UrlFetchApp.fetchAll(requests);
  } catch (e) {
    return { summaries: summaries, error: "요청 실패 (" + e.message + ") → =AI() 수식으로 대체했습니다." };
  }

  responses.forEach(function(res) {
    var code = res.getResponseCode();
    if (code !== 200) {
      errors.push("HTTP " + code + " " + res.getContentText().substring(0, 200));
      return;
    }
    try {
      var body = JSON.parse(res.getContentText());
      var text = body.candidates[0].content.parts.map(function(p) { return p.text || ""; }).join("");
      JSON.parse(text).forEach(function(item) {
        if (item && typeof item.id === "number" && item.id >= 0 && item.id < summaries.length && item.summary) {
          summaries[item.id] = String(item.summary).trim();
        }
      });
    } catch (e) {
      errors.push("응답 해석 실패 (" + e.message + ")");
    }
  });

  if (errors.length > 0) console.warn("Gemini API errors: " + errors.join(" | "));
  var error = errors.length > 0 ? errors[0] + " → 실패한 행은 =AI() 수식으로 대체했습니다." : "";
  return { summaries: summaries, error: error };
}

/**
 * [=AI() 수식 대체]
 * G열(Summary)이 비어 있고 K열(Full Content)이 있는 행에만 =AI() 수식을 넣음
 * @return {number} 수식을 넣은 행 수
 */
function applyAiFormulaFallback_(sheet) {
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
      sheet.getRange(row, COL.SUMMARY).setFormula('=AI("' + SUMMARY_PROMPT + '", ' + fullColLetter + row + ')');
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
 * [Gemini API 키 설정] 스크립트 속성 GEMINI_API_KEY 에 저장
 */
function setGeminiApiKey() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    'Gemini API 키 설정',
    'Google AI Studio(https://aistudio.google.com/apikey)에서 발급한 API 키를 입력하세요.\n(비워 두고 OK를 누르면 키가 삭제됩니다)',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() != ui.Button.OK) return;

  var key = response.getResponseText().trim();
  var props = PropertiesService.getScriptProperties();
  if (key) {
    props.setProperty('GEMINI_API_KEY', key);
    ui.alert("API 키가 저장되었습니다.");
  } else {
    props.deleteProperty('GEMINI_API_KEY');
    ui.alert("API 키가 삭제되었습니다. 요약은 =AI() 수식으로 대체됩니다.");
  }
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
