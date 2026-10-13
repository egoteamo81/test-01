#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""drive_publish.py — 문서를 구글 드라이브 'Projects_Claude Code' 폴더의 최신본으로 갈아 끼운다.

왜 있나 (2026-10-01 사용자 요청)
  가이드북이 자주 개정되는데, 그때마다 사용자가 드라이브에 손으로 올리고 있었다.
  드라이브에는 버전 관리가 필요 없고 최신본 한 개만 있으면 된다.

언제 도나
  docx_kit 의 Doc.save() 가 저장 직후 maybe_publish() 를 부른다. 규칙은 하나다.
  **/root/my_project/scripts/ 아래 어떤 `docs/` 폴더에 저장된 .docx 는 올린다.**
  모든 프로젝트의 문서가 그 프로젝트의 docs/ 에 저장되므로(scripts/CLAUDE.md 규칙),
  build 스크립트를 따로 고치지 않아도 전 프로젝트에 걸린다.
  docs/ 밖(시험 문서, light_screen 의 종목별 문서 등)은 올리지 않는다.
  올리는 곳은 'Projects_Claude Code' 아래 하위 폴더다 (2026-10-07): 할 일 문서는 'followup', 나머지는 '가이드북'. folder_for() 참고.
  끄려면 환경변수 DRIVE_PUBLISH=0.

어떻게 갈아 끼우나
  같은 이름의 파일이 폴더에 있으면 **그 파일의 내용을 덮어쓴다.** 없으면 새로 만든다.
  사용자는 '지우고 새로 올려 달라'고 했는데 덮어쓰기를 택한 이유:
    - 결과는 같다. 폴더에는 최신본 한 개만 있다.
    - 링크(파일 ID)가 그대로라 북마크·공유 링크가 안 깨진다.
    - 사용자가 손으로 올린 기존 파일은 사용자 소유라 서비스 계정이 지울 수 없다
      (내 드라이브에서는 소유자만 휴지통에 넣을 수 있다. 2026-10-01 확인: canTrash=False,
      canEdit=True). 덮어쓰기는 편집 권한만으로 된다.
  구글 문서로 변환해서 올린다. 사용자가 손으로 올릴 때도 변환해 왔다(기존 9개 전부
  google-apps.document). 구글 문서는 드라이브 용량을 쓰지 않아 서비스 계정 용량 문제도 없다.
  같은 이름이 여러 개면 가장 최근 것을 덮어쓰고 나머지는 휴지통에 넣는다(권한이 되면).

실패해도 빌드는 멈추지 않는다. 워드 파일은 이미 디스크에 저장된 뒤다. 대신 크게 알린다.

손으로 올리기
  /root/myenv/bin/python3 drive_publish.py <파일.docx> [...]
"""
import os
import sys
import time

FOLDER_ID = '17CZdHwjVrB2RDBzDEeugT-wvuAT0PEt7'      # 구글 드라이브 'Projects_Claude Code'
# 개인정보가 든 자료는 여기 두지 않는다. private 프로젝트는 'Private_Claude Code' 아래 자기 폴더를 folder_id 로 넘긴다
# (2026-10-04. 규칙은 scripts/private/CLAUDE.md). publish() 의 이름 검색은 지정한 폴더 '바로 아래' 만 본다.
# 2026-10-07 사용자 지시: 'Projects_Claude Code' 바로 아래에 두지 않고 하위 폴더 둘로 나눈다.
# 폴더는 사용자 계정(드라이브 커넥터)으로 만들었다. 서비스 계정이 만들면 소유자가 서비스 계정이 되어 사용자가 정리하기 어렵다.
GUIDE_FOLDER_ID = '1kQdMvkSbwlL6Av49GOt6m-DdJdJXYA05'      # 'Projects_Claude Code/가이드북' 가이드북·보고서·학습노트 등 일반 문서
FOLLOWUP_FOLDER_ID = '16kjMdyGOxNUkQUZMfSoNSPY4VyAAT5L3'   # 'Projects_Claude Code/followup' 할 일 문서 (shared/followup_doc.py)


def folder_for(name):
    """파일 이름으로 올릴 폴더를 고른다. 할 일 문서는 이름에 '_followup' 이 들어간다 (followup_doc.py 가 그렇게 짓는다)."""
    return FOLLOWUP_FOLDER_ID if '_followup' in name else GUIDE_FOLDER_ID


SA_FILE = '/root/my_project/advance-rush-301620-c0f342bab7b4.json'   # Stocks 동기화와 같은 서비스 계정
SCRIPTS_ROOT = '/root/my_project/scripts/'
DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
GDOC_MIME = 'application/vnd.google-apps.document'

_drive = None


def _svc():
    global _drive
    if _drive is None:
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
        creds = service_account.Credentials.from_service_account_file(
            SA_FILE, scopes=['https://www.googleapis.com/auth/drive'])
        _drive = build('drive', 'v3', credentials=creds, cache_discovery=False)
    return _drive


def should_publish(path):
    """docx_kit.save() 가 묻는다. 이 파일을 드라이브에 올릴 대상인가."""
    if os.environ.get('DRIVE_PUBLISH', '1') == '0':
        return False
    p = os.path.abspath(path)
    # 개인정보 구역은 절대 올리지 않는다 (2026-10-01). 결과물은 사용자가 말할 때만 개인 폴더로.
    if p.startswith(SCRIPTS_ROOT + 'private/') or p.startswith('/root/my_project/private_data/'):
        return False
    return (p.startswith(SCRIPTS_ROOT) and p.lower().endswith('.docx')
            and f'{os.sep}docs{os.sep}' in p[len(SCRIPTS_ROOT) - 1:])


def publish(path, folder_id=None):
    """폴더의 같은 이름 파일을 이 파일 내용으로 갈아 끼운다. 결과 dict 를 돌려준다.
    folder_id 를 안 주면 folder_for() 가 이름으로 고른다 (가이드북 / followup)."""
    from googleapiclient.http import MediaFileUpload
    d = _svc()
    name = os.path.basename(path)
    folder_id = folder_id or folder_for(name)
    # 구글 문서로 변환해 새로 만들면 드라이브가 이름에서 '.docx' 를 떼어 버린다
    # (2026-10-01 실측: '_ZZ_훅시험_삭제예정.docx' 로 만들었는데 '_ZZ_훅시험_삭제예정' 으로 저장됨).
    # 그래서 두 이름을 다 찾고, 새로 만든 뒤에는 이름을 '.docx' 붙은 원래대로 되돌린다.
    # 사용자가 손으로 올린 기존 파일들이 '.docx' 를 달고 있어 그쪽에 맞춘다.
    # 사용자는 읽은 문서 이름 앞에 '✅' 를 붙인다 (2026-10-12 운영가이드에서 발견: 못 알아보고 새 판을 따로 만들었다).
    # 그 이름도 같은 문서로 본다. 갱신은 내용만 바꾸므로 사용자가 붙인 표시는 그대로 남는다.
    stem = os.path.splitext(name)[0]
    esc = lambda x: x.replace('\\', '\\\\').replace("'", "\\'")
    names = ' or '.join(f"name = '{esc(p + n)}'" for p in ('', '✅') for n in (name, stem))
    found = d.files().list(
        q=(f"'{folder_id}' in parents and trashed = false and ({names})"),
        fields='files(id,mimeType,modifiedTime,capabilities(canEdit,canTrash))',
        pageSize=50).execute().get('files', [])
    if not found and folder_id in (GUIDE_FOLDER_ID, FOLLOWUP_FOLDER_ID):
        # 하위 폴더로 나누기 전(2026-10-07) 루트에 남은 같은 이름이 있으면 새로 만들지 않고 옮겨 와서 갈아 끼운다.
        # 새로 만들면 링크가 바뀌고 루트에 옛 판이 남는다.
        old = d.files().list(
            q=(f"'{FOLDER_ID}' in parents and trashed = false and ({names})"),
            fields='files(id,mimeType,modifiedTime,capabilities(canEdit,canTrash))', pageSize=50).execute().get('files', [])
        for f in old:
            d.files().update(fileId=f['id'], addParents=folder_id, removeParents=FOLDER_ID, fields='id').execute()
        found = old
    found.sort(key=lambda f: f['modifiedTime'], reverse=True)
    media = MediaFileUpload(path, mimetype=DOCX_MIME, resumable=False)
    t = time.time()
    target = next((f for f in found if f['capabilities']['canEdit']), None)
    if target:
        f = d.files().update(fileId=target['id'], media_body=media,
                             fields='id,mimeType').execute()
        action = '갱신'
    else:
        f = d.files().create(body={'name': name, 'parents': [folder_id], 'mimeType': GDOC_MIME},
                             media_body=media, fields='id,mimeType,name').execute()
        if f['name'] != name:
            d.files().update(fileId=f['id'], body={'name': name}).execute()
        action = '새로 올림'
    extra, cannot = 0, []
    for old in found:
        if old['id'] == f['id']:
            continue
        if old['capabilities']['canTrash']:
            d.files().update(fileId=old['id'], body={'trashed': True}).execute()
            extra += 1
        else:
            cannot.append(old['id'])
    kind = 'document' if f['mimeType'] == GDOC_MIME else 'file'
    url = (f"https://docs.google.com/document/d/{f['id']}/edit" if kind == 'document'
           else f"https://drive.google.com/file/d/{f['id']}/view")
    return {'action': action, 'name': name, 'id': f['id'], 'url': url,
            'sec': round(time.time() - t, 1), 'trashed_dups': extra, 'stuck_dups': cannot}


def maybe_publish(path):
    """docx_kit.save() 에서 부른다. 대상이면 올리고 한 줄 알린다. 예외를 밖으로 내지 않는다."""
    if not should_publish(path):
        return None
    try:
        r = publish(path)
    except Exception as e:                       # noqa: BLE001
        print(f'[drive] ※ 업로드 실패. 워드 파일은 저장됨. {os.path.basename(path)}: '
              f'{type(e).__name__}: {str(e)[:200]}', file=sys.stderr)
        return None
    msg = f"[drive] {r['action']}: {r['name']} ({r['sec']}초) {r['url']}"
    if r['trashed_dups']:
        msg += f" / 같은 이름 중복 {r['trashed_dups']}개 휴지통"
    if r['stuck_dups']:
        msg += f" / ※ 지울 권한 없는 중복 {len(r['stuck_dups'])}개 남음. 드라이브에서 직접 지울 것"
    print(msg)
    return r


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    for p in sys.argv[1:]:
        r = publish(p)
        print(f"[drive] {r['action']}: {r['name']} ({r['sec']}초) {r['url']}")
