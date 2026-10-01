/**
 * ============================================================
 * [보내는 사람 정보] - 모든 메일의 서명(signature.html)에 공통 사용
 * ============================================================
 */
var SENDER = {
  name: '임태현',
  dept: '김해중정비공장 중정비지원팀',
  email: 'thlim@koreanair.com',
  phone: '+82 51 970-6094'
};

/**
 * ============================================================
 * [메일 작업(Job) 목록]
 *  - 메일 1종 = 객체 1개.
 *  - 새 메일 추가: ① 아래 배열에 객체 추가  ② 본문 HTML 파일(template 이름) 1개 생성
 *  - 링크·수신자·일정 변경은 이 파일만 수정 (HTML 파일은 건드리지 않음)
 *
 * [schedule 작성법]  hour: 이 시각(0~23시) 이후 첫 트리거 실행 때 발송 (트리거는 1시간마다 실행)
 *   { type: 'daily',   hour: 8 }                                매일
 *   { type: 'daily',   hour: 8, weekdaysOnly: true }            평일(월~금)만
 *   { type: 'weekly',  days: ['MON', 'THU'], hour: 8 }          매주 지정 요일 (MON TUE WED THU FRI SAT SUN)
 *   { type: 'weekly',  days: ['MON'], hour: 8, everyNWeeks: 2 } 격주 (startDate 가 속한 주부터 2주마다)
 *   { type: 'monthly', dates: [1, 15, 'LAST'], hour: 8 }        매월 지정 일 ('LAST' = 말일)
 *
 * [제목에서 쓸 수 있는 치환값]  {{date}} → 260915 (YYMMDD),  {{today}} → 2026-09-15
 * [본문 HTML 에서 쓸 수 있는 값] vars.date, vars.today, links.xxx, job.xxx, sender.xxx
 * ============================================================
 */
var MAIL_JOBS = [
  {
    id: 'svc_tag_hardcopy',            // 고유 ID (발송 이력 저장 키로 사용, 변경 시 이력 초기화됨)
    enabled: true,
    template: 'mail_svc_tag',          // 본문 HTML 파일 이름 (mail_svc_tag.html)
    subject: '[요청] 정비문서 관리 - 부품수리 SVC Tag 사본, 하드카피 수행본',
    to: 'thlim@koreanair.com',         // 여러 명이면 쉼표로 구분
    cc: 'thlim@koreanair.com',
    // 본문 상단에 "수신/참조"로 표시만 되는 목록 (실제 발송 대상 아님, 필요 없으면 삭제)
    displayTo: ['iljkim@koreanair.com', 'kimthyun@koreanair.com', 'milkim@koreanair.com',
                'syunkang@koreanair.com', 'seongkpark@koreanair.com'],
    displayCc: ['juclee@koreanair.com', 'heebongjung@koreanair.com'],
    // ★ 실제 연결할 파일 링크 (비어 있으면 실제 발송이 차단되고 오류 알림)
    links: {
      attach1: '',   // 첨부1. 부품수리 SVC Tag 사본 보관_PUSMH
      attach2: ''    // 첨부2. 하드카피 수행본 보관_PUSMH
    },
    schedule: { type: 'weekly', days: ['MON'], hour: 8 },
    startDate: '',                     // 'YYYY-MM-DD' (비우면 제한 없음)
    endDate: '2026-10-26'              // 이 날짜까지 발송 (당일 포함)
  }

  // , { id: '...', template: '...', ... }   ← 새 메일은 여기에 추가
];
