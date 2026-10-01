# Gmail → Google Sheets 수집기 (Google Apps Script)

`gmail_to_sheets.gs` 를 스프레드시트의 Apps Script 편집기에 붙여 넣어 사용합니다. (시트 이름: `mails`)

## 열 구성
| 열 | 내용 | 생성 방식 |
|---|---|---|
| A~E | Date, Sender, Recipient, Labels, Subject | 값 |
| F | Urgency | 스크립트에서 정규식으로 분류한 **값** (수식 없음) |
| G | Summary | 정렬 후 `=AI(K행 & "는 메일의 제목과 본문 일부를 추출한 텍스트인데 '음슴체' 2문장으로 요약해 줘")` 입력 → 결과가 나오면 **값으로 고정** |
| H | Snippet | 값 |
| I | Note | (직접 입력) |
| J | Thread ID | 값 |
| K | Full Content | `Subject: …` + 줄바꿈 + `Snippet: …` **값** |

## 동작 방식
- 스크립트는 `=AI()` 수식을 **입력만 하고 바로 종료**합니다. 요약 생성은 시트에서 비동기로 진행되므로 스크립트 실행 시간에 포함되지 않습니다.
- 다음 `Get Emails` 실행 시작 시(정렬 전에) 결과가 나온 `=AI()` 셀을 값으로 고정합니다.
- 마지막 실행 후에는 **Freeze AI() Summaries to Values** 메뉴로 남은 수식을 값으로 고정하세요.
- 기존 F2의 `=MAP(...)` 수식은 첫 실행 때 자동으로 제거되고 기존 행의 긴급도는 값으로 채워집니다.

## 메뉴
- **Get Emails (Popup)**: 기간 입력 → 최대 50개 수집·분류·요약 → 보관(Archive) → 정렬
- **Freeze AI() Summaries to Values**: G열의 `=AI()` 수식 중 결과가 나온 셀을 값으로 고정 (Get Emails 실행 시에도 자동 수행)
- **Clear Emails Contents**: 헤더 제외 A~K열 내용 삭제
