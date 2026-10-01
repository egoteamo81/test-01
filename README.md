# Gmail → Google Sheets 수집기 (Google Apps Script)

`gmail_to_sheets.gs` 를 스프레드시트의 Apps Script 편집기에 붙여 넣어 사용합니다. (시트 이름: `mails`)

## 열 구성
| 열 | 내용 | 생성 방식 |
|---|---|---|
| A~E | Date, Sender, Recipient, Labels, Subject | 값 |
| F | Urgency | 스크립트에서 정규식으로 분류한 **값** (수식 없음) |
| G | Summary | Gemini API 일괄 요약 **값** / 실패 시 `=AI()` 수식 → 나중에 값으로 고정 |
| H | Snippet | 값 |
| I | Note | (직접 입력) |
| J | Thread ID | 값 |
| K | Full Content | `Subject: …` + 줄바꿈 + `Snippet: …` **값** |

## 최초 설정
1. Google AI Studio(https://aistudio.google.com/apikey)에서 API 키 발급
2. 시트 메뉴 **Email Tools → Set Gemini API Key** 에서 키 입력 (스크립트 속성 `GEMINI_API_KEY` 에 저장)
3. 첫 실행 시 외부 요청(UrlFetchApp) 권한 재승인
4. 기존 F2의 `=MAP(...)` 수식은 첫 실행 때 자동으로 제거되고 기존 행의 긴급도는 값으로 채워집니다.

모델을 바꾸려면 스크립트 속성 `GEMINI_MODEL` 을 추가하세요. (기본값 `gemini-2.5-flash-lite`)

## 메뉴
- **Get Emails (Popup)**: 기간 입력 → 최대 50개 수집·분류·요약 → 보관(Archive) → 정렬
- **Freeze AI() Summaries to Values**: G열의 `=AI()` 수식 중 결과가 나온 셀을 값으로 고정 (Get Emails 실행 시에도 자동 수행)
- **Set Gemini API Key**: API 키 저장/삭제
- **Clear Emails Contents**: 헤더 제외 A~K열 내용 삭제
