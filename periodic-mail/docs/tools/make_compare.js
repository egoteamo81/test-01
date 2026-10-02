// "내가 쓴 HTML(코드)" 와 "실제 메일에서 보이는 모습" 을 나란히 보여 주는 이미지 생성기.
//  - 오른쪽 그림은 손으로 그린 것이 아니라, 진짜 Main.gs(linkHtml 등)와 진짜 템플릿 규칙으로 계산한 결과다.
//  - 사용:  node tools/make_compare.js   → img/cmp_*.png 생성 (크롬 필요)
const fs = require('fs'), path = require('path'), cp = require('child_process');
const ROOT = path.resolve(__dirname, '..', '..');          // periodic-mail/
const OUT = path.resolve(__dirname, '..', 'img');
const CHROME = '/opt/pw-browsers/chromium-1194/chrome-linux/chrome';

// ── Apps Script 환경 흉내 (Main.gs 를 그대로 읽어서 쓴다) ──
function compile(src) {
  let code = "var __o='';\n", re = /<\?(!=|=)?([\s\S]*?)\?>/g, last = 0, m;
  while ((m = re.exec(src))) {
    code += "__o+=" + JSON.stringify(src.slice(last, m.index)) + ";\n";
    if (m[1] === '!=') code += "__o+=String(" + m[2] + ");\n";
    else if (m[1] === '=') code += "__o+=escapeHtml_(String(" + m[2] + "));\n";
    else code += m[2] + "\n";
    last = re.lastIndex;
  }
  return code + "__o+=" + JSON.stringify(src.slice(last)) + ";\nreturn __o;";
}
function render(src, data) { return new Function('__d', 'with(__d){' + compile(src) + '}')(data); }
global.HtmlService = { createTemplateFromFile(n) {
  const src = fs.readFileSync(path.join(ROOT, n + '.html'), 'utf8'); const t = {};
  t.evaluate = () => { const d = {}; for (const k in t) if (k !== 'evaluate') d[k] = t[k]; const out = render(src, d); return { getContent: () => out }; };
  return t; } };
global.Session = { getScriptTimeZone: () => 'Asia/Seoul', getEffectiveUser: () => ({ getEmail: () => 'me@example.com' }) };
global.Utilities = { formatDate(d, tz, p) { const s = new Date(d.getTime() + 9 * 3600e3), pad = n => String(n).padStart(2, '0');
  return p.replace('yyyy', s.getUTCFullYear()).replace('yy', String(s.getUTCFullYear()).slice(-2)).replace('MM', pad(s.getUTCMonth() + 1)).replace('dd', pad(s.getUTCDate())); } };
for (const f of ['Config.gs', 'Main.gs']) (0, eval)(fs.readFileSync(path.join(ROOT, f), 'utf8'));

const NOW = new Date('2026-09-14T23:30:00Z');                // 한국 시간 2026-09-15
const VARS = { date: '260915', today: '2026-09-15' };
const baseData = (extra) => Object.assign({ job: { deadline: '10월 31일' }, links: { sheet: 'https://docs.google.com/spreadsheets/d/abc123' }, vars: VARS, sender: SENDER }, extra || {});

// ── 코드 색칠 ──
const esc = s => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
function colorize(src) {
  return src.split(/(<\?[\s\S]*?\?>)/).map(part => {
    if (/^<\?/.test(part)) return '<span class="sc">' + esc(part) + '</span>';
    return esc(part).replace(/(&lt;\/?[a-zA-Z][^]*?&gt;)/g, m => '<span class="tg">' + m + '</span>');
  }).join('');
}
const CSS = `
*{box-sizing:border-box}body{margin:0;background:#fff;font-family:'Noto Sans KR','Malgun Gothic',sans-serif;color:#1A1A1A}
.wrap{width:var(--w);padding:20px 24px 24px}
.hd{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:8px}
.hd div{font-weight:700;font-size:15px;padding:7px 12px;border-radius:8px 8px 0 0}
.h1{background:#1E3A5F;color:#fff}.h2{background:#15803D;color:#fff}
.row{margin-bottom:14px}.lab{font-size:13.5px;font-weight:700;color:#0F5C8C;margin:0 0 5px 2px}
.grid{display:grid;grid-template-columns:1fr 1fr;gap:16px;align-items:stretch}
pre{margin:0;background:#F4F4F5;border:1px solid #C9D2DC;border-left:5px solid #1E3A5F;border-radius:6px;padding:10px 12px;font-family:'DejaVu Sans Mono',Consolas,monospace;font-size:12.5px;line-height:1.55;white-space:pre-wrap;word-break:break-all;overflow:hidden}
.tg{color:#0B5CAD}.sc{color:#B32020;background:#FDEAEA;border-radius:3px}
.mail{border:1px solid #C9D2DC;border-left:5px solid #15803D;border-radius:6px;overflow:hidden;background:#fff}
.mh{background:#F2F5F8;border-bottom:1px solid #C9D2DC;padding:5px 12px;font-size:12px;color:#6B7280}
.mb{padding:12px 14px;font-family:'Malgun Gothic','맑은 고딕','Noto Sans KR',sans-serif;font-size:14px;color:#000;line-height:1.6;overflow-wrap:anywhere}
.note{font-size:12.5px;color:#6B7280;margin:5px 2px 0}
`;
function page(rows, head = ['내가 쓴 HTML (코드)', '받는 사람이 보는 메일'], width = 780) {
  const body = rows.map(r => {
    const out = r.result != null ? r.result : render(r.src, r.data || baseData());
    return `<div class="row">${r.label ? `<div class="lab">${esc(r.label)}</div>` : ''}<div class="grid">
      <pre>${colorize(r.src.replace(/\n$/, ''))}</pre>
      <div class="mail"><div class="mh">${esc(r.subject || '제목: [요청] 예시 메일')}</div><div class="mb">${out}</div></div></div>${r.note ? `<div class="note">${esc(r.note)}</div>` : ''}</div>`;
  }).join('');
  return `<!doctype html><meta charset="utf-8"><style>${CSS}:root{--w:${width}px}</style><div class="wrap"><div class="hd"><div class="h1">${head[0]}</div><div class="h2">${head[1]}</div></div>${body}</div>`;
}
function shoot(name, html, width = 780) {
  const f = path.join(OUT, name + '.html'); fs.writeFileSync(f, html);
  cp.execFileSync(CHROME, ['--headless=new', '--no-sandbox', '--disable-gpu', '--hide-scrollbars', '--force-device-scale-factor=2',
    '--window-size=' + width + ',4200', '--screenshot=' + path.join(OUT, name + '.png'), 'file://' + f], { stdio: 'ignore' });
  cp.execFileSync('python3', ['-c', `
from PIL import Image, ImageChops
p='${path.join(OUT, name + '.png')}'
im=Image.open(p).convert('RGB'); b=ImageChops.difference(im,Image.new('RGB',im.size,(255,255,255))).getbbox()
im.crop((0,0,im.width,min(im.height,b[3]+30))).save(p)`]);
  fs.unlinkSync(f); console.log('● img/' + name + '.png');
}

// ══════════ 예제들 ══════════
shoot('cmp_1_linebreak', page([
  { label: '① 엔터만 친 경우 — 메일에서는 한 줄로 이어 붙는다', src: '안녕하십니까.\n아래와 같이 요청 드립니다.\n제출 기한은 10월 31일 입니다.',
    note: '코드에서 엔터(줄 바꿈)를 쳐도 메일은 무시합니다. 공백 한 칸으로 취급해요.' },
  { label: '② <br> 을 쓴 경우 — 줄이 바뀐다', src: '안녕하십니까.<br>\n아래와 같이 요청 드립니다.<br>\n제출 기한은 10월 31일 입니다.',
    note: '<br> 이 "줄 바꿈" 입니다. 코드에서 엔터는 보기 좋게 정리하려고 치는 것일 뿐입니다.' },
  { label: '③ <br> 을 두 번 — 한 줄을 비운다 (문단 나누기)', src: '안녕하십니까.<br><br>\n아래와 같이 요청 드립니다.<br>' },
]));

shoot('cmp_2_text', page([
  { label: '굵게 · 색 · 배경 강조', src: '<strong>1. 제출 기한</strong><br>\n<span style="color: #c00;">10월 31일</span> 까지 제출<br>\n<span style="background: #ffeb3b;">반드시</span> 회신 바랍니다.<br>\n<u>밑줄</u> 과 <i>기울임</i> 도 가능합니다.' },
  { label: '들여쓰기 — 한 칸씩 안으로', src: '<strong>1. 항목</strong><br>\n<div style="margin-left: 20px;">\n  가. 첫째 설명<br>\n  나. 둘째 설명<br>\n  <div style="margin-left: 20px;">\n    - 더 자세한 설명<br>\n  </div>\n</div>',
    note: 'margin-left: 20px 가 "한 단계", 40px 가 "두 단계" 입니다. <div> 를 겹치면 단계가 쌓입니다.' },
]));

shoot('cmp_3_table', page([
  { label: '표 — 줄 · 칸 · 테두리', src: '<table border="1" cellpadding="6" cellspacing="0"\n       style="border-collapse: collapse; border-color: #999;">\n  <tr style="background: #eee;">\n    <th>구분</th><th>기한</th>\n  </tr>\n  <tr><td>SVC Tag</td><td>10/31</td></tr>\n  <tr><td>하드카피</td><td>11/15</td></tr>\n</table>',
    note: '<tr> 한 줄, <th> 머리 칸, <td> 일반 칸. 메일에서는 border 속성을 직접 써 주는 것이 가장 안전합니다.' },
]));

shoot('cmp_4_values', page([
  { label: '값 끼워 넣기 — 코드의 <?= … ?> 가 실제 값으로 바뀐다',
    src: '시트 이름: <?= vars.date ?><br>\n발송일: <?= vars.today ?><br>\n제출 기한: <?= job.deadline ?><br>\n보내는 사람: <?= sender.name ?>',
    note: '이 예에서 vars.date = 260915, vars.today = 2026-09-15, job.deadline = "10월 31일", sender.name = "임태현" 이었습니다. 발송하는 날마다 날짜가 자동으로 바뀝니다.' },
]));

shoot('cmp_5_link', page([
  { label: '① 링크 — linkHtml 로 넣기 (권장)', src: '첨부. 점검표 (<?!= linkHtml(links.sheet, \'링크\') ?>)',
    note: '파란 글씨 "링크"를 누르면 links.sheet 에 넣어 둔 주소로 이동합니다.' },
  { label: '② 링크 칸을 비워 둔 경우 — 실제 발송은 막히고, 테스트에서는 빨간 글씨로 알려 준다', src: '첨부. 점검표 (<?!= linkHtml(links.sheet, \'링크\') ?>)',
    data: baseData({ links: { sheet: '' } }) },
  { label: '③ 링크를 직접 쓴 경우 — 되지만, 색·빈 칸 검사를 못 받는다 (비권장)', src: '첨부. 점검표 (<a href="https://example.com">링크</a>)',
    note: '주소가 바뀔 때마다 HTML 을 고쳐야 하고, 빈 주소 검사도 받지 못합니다.' },
]));

shoot('cmp_6_logic', page([
  { label: '조건 — 값이 있을 때만 보이기 (job.deadline 이 있는 경우)',
    src: '<? if (job.deadline) { ?>\n제출 기한: <?= job.deadline ?><br>\n<? } ?>\n문의: 회신 바랍니다.', data: baseData() },
  { label: '조건 — 같은 코드, 값이 없는 경우 (그 줄이 통째로 사라진다)',
    src: '<? if (job.deadline) { ?>\n제출 기한: <?= job.deadline ?><br>\n<? } ?>\n문의: 회신 바랍니다.', data: baseData({ job: {} }) },
  { label: '반복 — 목록을 한 줄씩 찍기 (job.items 가 3개인 경우)',
    src: '<? for (var i = 0; i < job.items.length; i++) { ?>\n<?= (i + 1) ?>. <?= job.items[i] ?><br>\n<? } ?>',
    data: baseData({ job: { items: ['SVC Tag 사본', '하드카피 수행본', '검사 성적서'] } }),
    note: 'Config.gs 의 메일 항목에 items: [\'SVC Tag 사본\', \'하드카피 수행본\', \'검사 성적서\'] 를 추가해 두면, 개수가 달라져도 HTML 은 그대로입니다.' },
]));

shoot('cmp_7_mistakes', page([
  { label: '실수 ① <?= ?> 로 링크를 출력 — 태그가 글자 그대로 보인다', src: '첨부 (<?= linkHtml(links.sheet, \'링크\') ?>)',
    note: '<?= ?> 는 안전하게 "글자로" 출력합니다. 링크처럼 HTML 을 그대로 쓸 값은 <?!= ?> 를 써야 합니다.' },
  { label: '정답 — <?!= ?>', src: '첨부 (<?!= linkHtml(links.sheet, \'링크\') ?>)' },
  { label: '실수 ② 꺾쇠(<)를 그대로 쓴 경우 — <A열> 이 태그로 오인되어 글자가 사라진다', src: '1 < 2 이면 보고<br>\n<A열> 을 갱신<br>\n첨부1 > 확인',
    note: '"<" 바로 뒤에 글자가 오면 메일 프로그램이 태그로 읽어서 그 글자가 화면에서 사라집니다. (A열 이 안 보임)' },
  { label: '정답 — &lt; 와 &gt; 로 쓴다', src: '1 &lt; 2 이면 보고<br>\n&lt;A열&gt; 을 갱신<br>\n첨부1 &gt; 확인' },
]));

// ── 실제 파일 ↔ 실제 메일 (진짜 buildMail_ 결과) ──
const job = JSON.parse(JSON.stringify(MAIL_JOBS[0]));
job.links = { attach1: 'https://docs.google.com/spreadsheets/d/aaa', attach2: 'https://docs.google.com/spreadsheets/d/bbb' };
const mail = buildMail_(job, NOW);
const realSrc = fs.readFileSync(path.join(ROOT, 'mail_svc_tag.html'), 'utf8').replace(/^<\?[\s\S]*?\?>\n/, '<? /* (주석은 메일에 안 나옴) */ ?>\n');
shoot('cmp_8_real', page([{ label: '', src: realSrc, result: mail.html, subject: '제목: ' + mail.subject,
  note: '왼쪽은 mail_svc_tag.html(본문만) 입니다. 오른쪽 메일의 맨 위 "수신/참조" 와 맨 아래 "임태현 드림" 서명은 layout.html·signature.html 이 자동으로 붙인 것입니다.' }],
  ['mail_svc_tag.html (내가 쓰는 본문)', '실제 발송되는 메일 (본문 + 공통 틀 + 서명)'], 900), 900);
