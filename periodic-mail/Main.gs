/**
 * ============================================================
 * [정기 메일 발송기]
 *  - 트리거 1개(1시간마다)가 runScheduledMails 를 실행 → MAIL_JOBS 중 오늘 발송 대상만 발송
 *  - 같은 메일은 하루 1번만 발송 (스크립트 속성에 마지막 발송일 기록)
 *  - 본문은 HTML 템플릿 파일 + 공통 레이아웃(layout.html) + 공통 서명(signature.html) 으로 조립
 *
 * [편집기에서 직접 실행하는 함수]
 *   setupTrigger       : 1시간 주기 트리거 설치 (최초 1회)
 *   removeTrigger      : 트리거 삭제
 *   testSendToMe       : TEST_JOB_ID 메일을 나에게만 [TEST] 로 발송 (일정·링크 검사 무시)
 *   testSendAllToMe    : 모든 활성 메일을 나에게만 [TEST] 로 발송
 *   checkSchedule      : 앞으로 30일간 발송 예정일을 로그로 확인
 *   resetSentHistory   : 발송 이력 초기화 (오늘 다시 보내야 할 때)
 * ============================================================
 */
var TRIGGER_FUNCTION = 'runScheduledMails';
var TEST_JOB_ID = 'svc_tag_hardcopy';
var DAY_CODES = ['MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT', 'SUN'];
var LINK_STYLE = 'color: #1155cc;';

/**
 * [트리거 진입점] 1시간마다 실행
 */
function runScheduledMails() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) return;

  try {
    var now = new Date();
    var today = formatDate_(now, 'yyyy-MM-dd');
    var hour = Number(formatDate_(now, 'H'));
    var props = PropertiesService.getScriptProperties();
    var errors = [];

    MAIL_JOBS.forEach(function(job) {
      if (!isScheduledOn_(job, today) || hour < (job.schedule.hour || 0)) return;

      var sentKey = 'lastSent_' + job.id;
      if (props.getProperty(sentKey) === today) return; // 오늘 이미 발송

      try {
        sendJob_(job, now, false);
        props.setProperty(sentKey, today);
        console.log('발송 완료: ' + job.id + ' (' + today + ')');
      } catch (e) {
        errors.push(job.id + ': ' + e.message);
      }
    });

    // 💡 오류를 던져야 Apps Script 의 "트리거 실패 알림" 메일로 통지됨 (다른 메일 발송은 계속 진행된 뒤)
    if (errors.length > 0) {
      throw new Error('정기 메일 발송 실패\n' + errors.join('\n'));
    }
  } finally {
    lock.releaseLock();
  }
}

/**
 * [발송 일정 판단] 해당 날짜(yyyy-MM-dd)가 발송일인지 (시각은 따로 확인)
 */
function isScheduledOn_(job, ymd) {
  if (job.enabled === false) return false;
  if (job.startDate && ymd < job.startDate) return false;
  if (job.endDate && ymd > job.endDate) return false;

  var s = job.schedule || {};
  var dayNum = dayNumber_(ymd);
  var dayCode = DAY_CODES[(dayNum + 3) % 7]; // 1970-01-01 = 목요일

  switch (s.type) {
    case 'daily':
      return !(s.weekdaysOnly && (dayCode === 'SAT' || dayCode === 'SUN'));

    case 'weekly':
      if ((s.days || []).indexOf(dayCode) === -1) return false;
      if (s.everyNWeeks > 1) {
        if (!job.startDate) throw new Error(job.id + ': everyNWeeks 를 쓰려면 startDate 가 필요합니다.');
        var weeks = (weekStart_(dayNum) - weekStart_(dayNumber_(job.startDate))) / 7;
        return weeks % s.everyNWeeks === 0;
      }
      return true;

    case 'monthly':
      var parts = ymd.split('-').map(Number);
      var lastDay = new Date(Date.UTC(parts[0], parts[1], 0)).getUTCDate();
      return (s.dates || []).some(function(d) {
        return d === 'LAST' ? parts[2] === lastDay : Number(d) === parts[2];
      });

    default:
      throw new Error(job.id + ': 알 수 없는 schedule.type "' + s.type + '"');
  }
}

/**
 * [메일 조립 + 발송]
 * @param {boolean} isTest true 면 나에게만 [TEST] 로 발송 (참조 없음, 빈 링크 허용)
 */
function sendJob_(job, now, isTest) {
  if (!isTest) {
    var missing = Object.keys(job.links || {}).filter(function(k) { return !job.links[k]; });
    if (missing.length > 0) throw new Error('링크가 비어 있습니다: ' + missing.join(', '));
  }

  var mail = buildMail_(job, now);
  var options = { htmlBody: mail.html };
  var to = job.to;
  var subject = mail.subject;

  if (isTest) {
    to = Session.getEffectiveUser().getEmail();
    subject = '[TEST] ' + subject;
  } else {
    if (job.cc) options.cc = job.cc;
    if (job.bcc) options.bcc = job.bcc;
    if (job.replyTo) options.replyTo = job.replyTo;
  }

  GmailApp.sendEmail(to, subject, mail.text, options);
}

function buildMail_(job, now) {
  var vars = {
    date: formatDate_(now, 'yyMMdd'),      // 예: 260915
    today: formatDate_(now, 'yyyy-MM-dd')  // 예: 2026-09-15
  };
  var data = { job: job, links: job.links || {}, vars: vars, sender: SENDER };

  var content = include(job.template, data);
  var html = include('layout', { job: job, sender: SENDER, vars: vars, content: content });

  return {
    subject: fillText_(job.subject, vars),
    html: html,
    text: htmlToText_(html)
  };
}

/**
 * [HTML 템플릿 렌더링] 템플릿 안에서도 <?!= include('파일명', {...}) ?> 로 부분 템플릿 삽입 가능
 */
function include(name, data) {
  var template = HtmlService.createTemplateFromFile(name);
  Object.keys(data || {}).forEach(function(key) { template[key] = data[key]; });
  return template.evaluate().getContent();
}

/**
 * [링크 HTML 생성] 템플릿에서 <?!= linkHtml(links.attach1, '링크') ?> 로 사용 → 모든 메일의 링크 스타일 통일
 */
function linkHtml(url, label) {
  if (!url) return '<span style="color: #d93025;">(링크 미입력)</span>';
  return '<a href="' + escapeHtml_(url) + '" style="' + LINK_STYLE + '">' + escapeHtml_(label || url) + '</a>';
}

// ------------------------------------------------------------
// 편집기에서 직접 실행하는 관리 함수
// ------------------------------------------------------------
function setupTrigger() {
  removeTrigger();
  ScriptApp.newTrigger(TRIGGER_FUNCTION).timeBased().everyHours(1).create();
  console.log('트리거 설치 완료: ' + TRIGGER_FUNCTION + ' (1시간마다)');
}

function removeTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === TRIGGER_FUNCTION) ScriptApp.deleteTrigger(trigger);
  });
}

function testSendToMe() {
  var job = MAIL_JOBS.filter(function(j) { return j.id === TEST_JOB_ID; })[0];
  if (!job) throw new Error('TEST_JOB_ID 에 해당하는 메일이 없습니다: ' + TEST_JOB_ID);
  sendJob_(job, new Date(), true);
  console.log('테스트 발송 완료: ' + job.id);
}

function testSendAllToMe() {
  MAIL_JOBS.forEach(function(job) {
    if (job.enabled === false) return;
    sendJob_(job, new Date(), true);
    console.log('테스트 발송 완료: ' + job.id);
  });
}

function checkSchedule() {
  var start = dayNumber_(formatDate_(new Date(), 'yyyy-MM-dd'));
  MAIL_JOBS.forEach(function(job) {
    var dates = [];
    for (var d = 0; d < 30; d++) {
      var ymd = new Date((start + d) * 86400000).toISOString().slice(0, 10);
      if (isScheduledOn_(job, ymd)) dates.push(ymd);
    }
    console.log(job.id + ' (' + (job.schedule.hour || 0) + '시 이후): ' + (dates.join(', ') || '30일 내 발송 없음'));
  });
}

function resetSentHistory() {
  var props = PropertiesService.getScriptProperties();
  MAIL_JOBS.forEach(function(job) { props.deleteProperty('lastSent_' + job.id); });
  console.log('발송 이력 초기화 완료');
}

// ------------------------------------------------------------
// 내부 헬퍼
// ------------------------------------------------------------
function formatDate_(date, pattern) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), pattern);
}

/** 'yyyy-MM-dd' → 1970-01-01 기준 일수 (시간대 영향 없음) */
function dayNumber_(ymd) {
  var p = ymd.split('-').map(Number);
  return Math.floor(Date.UTC(p[0], p[1] - 1, p[2]) / 86400000);
}

/** 해당 일이 속한 주의 월요일 일수 */
function weekStart_(dayNum) {
  return dayNum - (dayNum + 3) % 7;
}

/** 제목 등 일반 텍스트의 {{key}} 치환 */
function fillText_(text, vars) {
  return String(text).replace(/\{\{(\w+)\}\}/g, function(match, key) {
    return vars.hasOwnProperty(key) ? vars[key] : match;
  });
}

function escapeHtml_(text) {
  return String(text)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

/** HTML 미지원 메일 환경용 텍스트 본문 (링크는 "텍스트 (URL)" 로 남김) */
function htmlToText_(html) {
  return html
      .replace(/\s+/g, ' ')                                   // HTML 처럼 소스 줄바꿈·공백은 공백 1칸으로
      .replace(/<a\s[^>]*href="([^"]*)"[^>]*>([\s\S]*?)<\/a>/gi, '$2 ($1)')
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/<\/(div|p|li|tr|h\d)>/gi, '\n')
      .replace(/<[^>]+>/g, '')
      .replace(/&nbsp;/g, ' ').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&amp;/g, '&')
      .replace(/[ \t]+/g, ' ')
      .replace(/ *\n */g, '\n')
      .replace(/\n{3,}/g, '\n\n')
      .trim();
}
