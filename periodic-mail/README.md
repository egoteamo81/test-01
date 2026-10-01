# 정기 메일 발송기 (Google Apps Script)

## 파일 구성 — "코드 / 설정 / HTML" 분리
| 파일 | 역할 | 수정 빈도 |
|---|---|---|
| `Config.gs` | 메일 목록(`MAIL_JOBS`): 제목·수신자·**링크**·발송 주기·종료일, 보내는 사람(`SENDER`) | 자주 |
| `mail_*.html` | 메일별 **본문**만 (메일 1종 = 파일 1개) | 내용 바뀔 때 |
| `layout.html` | 공통 틀: 글꼴, 본문 상단 수신/참조 표시, 서명 위치 | 거의 없음 |
| `signature.html` | 공통 서명 | 거의 없음 |
| `Main.gs` | 일정 판단, 템플릿 조립, 발송, 트리거 관리 | 없음 |
| `appsscript.json` | 시간대 `Asia/Seoul` | 없음 |

Apps Script 편집기에서 **파일 추가(+) → HTML** 로 `layout`, `signature`, `mail_svc_tag` 를 만들고 내용을 붙여 넣습니다. (파일명에 `.html` 은 자동으로 붙음)

## HTML 관리 원칙
1. **본문 HTML 에는 바뀌는 값을 직접 쓰지 않는다.** 링크·날짜는 `Config.gs` 에 두고 템플릿에서 꺼내 씀
   - 날짜: `<?= vars.date ?>` (YYMMDD), `<?= vars.today ?>` (YYYY-MM-DD)
   - 링크: `<?!= linkHtml(links.attach1, '링크') ?>` → 모든 메일의 링크 색/스타일이 한 곳(`LINK_STYLE`)에서 통일
2. **공통 부분은 한 번만 작성.** 글꼴·서명·수신/참조 표시는 `layout.html`/`signature.html` 이 모든 메일에 자동 적용
3. **스타일은 인라인(style="…")으로.** Outlook 등 사내 메일 클라이언트는 `<style>` 블록을 무시하는 경우가 많음
4. **주석은 `<? /* … */ ?>` 로.** `<!-- -->` 는 메일 원본에 그대로 실려 나감
5. `<?= ?>` 는 자동 이스케이프(일반 값), `<?!= ?>` 는 HTML 그대로 출력(`linkHtml`, `include` 결과에만 사용)

## 새 메일 추가 방법
1. `mail_xxx.html` 파일을 만들고 본문만 작성 (기존 `mail_svc_tag.html` 복사해서 수정)
2. `Config.gs` 의 `MAIL_JOBS` 에 객체 1개 추가 (`template: 'mail_xxx'`, `links`, `schedule` 등)
3. `TEST_JOB_ID` 를 새 id 로 바꾸고 `testSendToMe` 실행 → 내 메일함에서 확인

## 최초 설정
1. `setupTrigger` 1회 실행 (권한 승인) → 1시간마다 `runScheduledMails` 실행
2. `checkSchedule` 실행 → 앞으로 30일 발송 예정일을 로그로 확인

## 동작 방식
- 트리거는 **1개**만 사용 (메일 종류가 늘어도 트리거 추가 불필요, 트리거 개수 제한 회피)
- 발송일에 `schedule.hour` 시 이후 첫 실행 때 발송, 같은 메일은 **하루 1번만** 발송
- 실제 발송 시 `links` 중 빈 값이 있으면 **발송하지 않고 오류** → Apps Script 트리거 실패 알림 메일로 통지
  (테스트 발송은 빈 링크를 빨간색 "(링크 미입력)" 으로 표시)
- `endDate` 가 지나면 자동으로 발송 중단
