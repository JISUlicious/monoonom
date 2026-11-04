# web-excel

로컬 Python 서버 + 브라우저 UI로 동작하는 **엑셀 호환 웹 스프레드시트**입니다.
에이전트 SW(Claude Code 데스크톱, VS Code/Cursor Simple Browser 등)의 미리보기 창에서 `xl.py serve`가 알려 주는 주소(`http://localhost:8765/#t=토큰`)로 접속해 xlsx 파일을 읽고 쓰고, xls 파일을 읽을 수 있습니다(xls는 xlsx로 저장).

## 실행

이 프로젝트는 **Agent Skill** 형태로 패키징되어 있습니다: `.agents/skills/web-excel/` (Claude Code용 링크: `.claude/skills/web-excel`).
Python 패키지는 함께 넣지 않으므로 처음 한 번 설치합니다: `python -m pip install -r requirements.txt` (openpyxl 필수, xlrd는 .xls를 열 때만). 설치되어 있지 않으면 `xl.py serve`가 설치 명령을 안내하고 종료 코드 6으로 끝납니다.

```bash
XL=.agents/skills/web-excel/scripts/xl.py
python3 $XL serve                 # 작업 경로(현재 폴더)용 서버를 분리 실행하거나 기존 서버 재사용 → URL 출력
python3 $XL status                # 상태·열린 문서
python3 $XL launch-config         # Claude Code 데스크톱: .claude/launch.json에 미리보기 구성 기록
python3 $XL stop                  # 종료 (저장 안 된 문서가 있으면 거부, --force)
python3 run.py --port 8765 --root .   # 개발용: 포그라운드 실행
python3 tests/make_samples.py     # 예제 xlsx 생성 (samples/sample.xls는 저장소에 있는 고정 파일)
```

### 에이전트가 쓰는 명령 (요약)

작업 폴더(사용자 프로젝트 폴더)를 현재 폴더로 두고, skill의 `xl.py`는 절대 경로로 부릅니다. skill 폴더로
`cd`해서 실행하면 거부됩니다(종료 코드 2). 다른 곳에서 부를 때는 `--root <작업 폴더>`.

```bash
cd ~/my-project                          # 작업 폴더
XL=/path/to/.agents/skills/web-excel/scripts/xl.py
python3 $XL open report.xlsx            # 열기 → 미리보기에 탭으로 표시
python3 $XL describe                    # 시트·사용 범위·머리글·열 유형·차트
python3 $XL read '매출!A1:F20'           # 필요한 범위만 (날짜는 YYYY-MM-DD)
python3 $XL fill 'G2:G7' '=E2*0.3'      # 수식 채우기(상대 참조 이동)
python3 $XL format G1:G7 b=true nf='#,##0' --border outline
python3 $XL chart add A1:D7 --type line --title 추이 --at H20
python3 $XL rows insert 5 --count 2     # 행 삽입 (수식·차트·필터 범위가 따라 이동)
python3 $XL sort A1:F50 --by C:desc     # 정렬
python3 $XL changes --since 0 --actor user   # 사용자가 미리보기에서 바꾼 것
python3 $XL undo BATCH                  # 에이전트 변경 되돌리기
python3 $XL save                        # 저장(이전 파일은 .web-excel/backups/)
```

전체 명령은 `.agents/skills/web-excel/references/cli.md`, 일괄 편집 형식은 `references/edits.md`, UI 기능과의 대응은 `docs/agent-coverage.md`.

| `serve` 옵션 | 설명 |
|---|---|
| `--root DIR` | 작업 경로(기본: `$WEBEXCEL_ROOT`, `$CLAUDE_PROJECT_DIR`, 현재 폴더). 이 밖의 파일은 접근 차단 |
| `--port P` | 선호 포트(기본 8765, 사용 중이면 다음 빈 포트) |
| `--idle-timeout S` | 클라이언트·요청이 없으면 S초 후 종료(기본 3600, 0=안 함) |
| `--foreground` / `--exit-with-parent` | 현재 프로세스에서 실행 / 부모 프로세스가 끝나면 종료(launch.json용) |
| `--allow-host NAME` | 프록시/원격 미리보기에서 다른 Host 헤더 허용 |
| `--no-sandbox` | 작업 경로 밖 파일 허용 |

서버 상태는 `<작업 경로>/.web-excel/`(server.json: pid·port·url, token, server.log)에 있고 이 폴더는 git에서 제외됩니다.
API는 토큰이 필요합니다. CLI와 페이지 모두 `X-WebExcel-Token` 헤더로 보냅니다. 페이지는 주소의 `#t=토큰`에서 토큰을 읽으므로, 페이지를 다른 사이트의 iframe 안에 띄우는 미리보기(VS Code Simple Browser, Tauri 앱 등)에서도 동작합니다. 토큰 없이 연 일반 탭은 포트별 쿠키(`wx_token_<포트>`, `SameSite=Strict`)로 동작합니다. 실시간 이벤트 스트림은 헤더를 보낼 수 없어 1분짜리 일회용 티켓을 씁니다. 토큰은 작업 경로별로 유지되어 서버를 다시 시작해도 열린 미리보기가 계속 동작합니다.

## 요구사항 대응

| # | 요구사항 | 구현 |
|---|---|---|
| 1 | 로컬 서버 + 미리보기 창에서 xlsx 읽기·쓰기, xls 읽기 | `run.py` → `server/app.py` (stdlib `ThreadingHTTPServer`). xlsx/xlsm: openpyxl, xls: xlrd(읽기 전용 — 저장하면 같은 이름의 .xlsx) |
| 2 | **DRM: 파일 읽기는 반드시 Python 함수** | 브라우저는 파일 바이트에 접근하지 않고 **경로만** 전송. 모든 디스크 I/O는 `server/fileio.py`의 `read_bytes()/write_bytes()` 한 곳을 통과. 드래그&드롭 업로드는 의도적으로 차단. DRM SDK 호출이 필요하면 `WEBEXCEL_READ_HOOK=pkg.mod:func`, `WEBEXCEL_WRITE_HOOK=...`로 교체 |
| 3 | 웹에서 조작 가능한 스프레드시트 | Canvas 가상화 그리드(1,048,576×16,384), 틀 고정, 병합, 행/열 크기 조절·숨기기·자동 맞춤, 채우기 핸들(선형 추세/요일·월 목록/숫자 접미사), 복사·붙여넣기(Excel과 TSV/HTML 상호 운용), 실행 취소/다시 실행 |
| 4 | 여러 탭/윈도우 | 상단 통합 문서 탭(다중 파일), 하단 시트 탭(추가·삭제·이름 변경·복사·이동·드래그·탭 색·숨기기). **서버가 문서를 소유**: 브라우저 창·에이전트 CLI가 모두 서버에 변경(op)을 보내고, 서버가 순서를 정해 각 창에 실시간 전달(SSE). 같은 문서를 여러 창·외부 브라우저에서 열어도 동기화되고, 에이전트가 연 파일은 미리보기에 탭으로 자동 표시 |
| 5 | 호환 차트 | 세로/가로/누적 막대, 꺾은선, 영역, 원형, 도넛, 분산형, 방사형, **콤보(막대+꺾은선)**. xlsx의 차트 정의(계열 참조 수식 `'시트'!$B$2:$B$7`)를 그대로 읽어 Chart.js로 렌더링, 저장 시 네이티브 Excel 차트로 기록. 수정하지 않은 차트는 원본 객체를 유지 |
| 6 | 스타일링 | 글꼴/크기/굵게/기울임/밑줄(이중)/취소선/색, 채우기, 테두리 13종, 맞춤·들여쓰기·줄 바꿈, 표시 형식(섹션·조건·색·날짜·시간·`[h]`·분수·지수·회계·₩), 서식 복사, 셀 서식(Ctrl+1). 테마 색+tint 해석(ECMA-376) |
| 7 | 필터 | 자동 필터(값 목록·검색·사용자 지정 조건 AND/OR·와일드카드), 정렬(오름/내림, 다중 기준), 다시 적용, 중복 제거. 파일의 `autoFilter`/`filterColumn`과 숨김 행으로 저장 |
| 8 | 호환 수식 | 자체 엔진(약 200개 함수): SUMIFS·XLOOKUP·INDEX/MATCH·OFFSET·INDIRECT·TEXT·DATEDIF·NETWORKDAYS·PMT·IRR·SUBTOTAL·LET 등. 배열 연산(SUMPRODUCT 조건식), 교차 시트·정의된 이름, 순환 참조 감지, 행/열 삽입·삭제·시트 이름 변경 시 참조 자동 조정, 미지원 함수는 Excel이 저장한 캐시 값으로 대체. `_xlfn.` 접두사 자동 처리 |
| 9 | 호환 단축키 | Excel 단축키 대부분(아래 표 / 앱에서 `⌨` 버튼). 한글 IME 조합 입력 안전 처리 |

### 주요 단축키

`Ctrl+N/O/S`, `F12`, `Ctrl+Z/Y`, `Ctrl+C/X/V`, `Ctrl+Alt+V`, `Ctrl+B/I/U/5`, `Ctrl+1`, `F2`, `F4`(참조 전환/반복), `Alt+Enter`, `Ctrl+Enter`, `Ctrl+D/R`, `Ctrl+;`, `Ctrl+Shift+;`, `Alt+=`, `Ctrl+Shift+L`, `Alt+↓`, `Ctrl+F/H`, `Ctrl+방향키`, `Ctrl+Shift+방향키`, `Ctrl+Home/End`, `Ctrl+A`, `Ctrl+Space`, `Shift+Space`, `Ctrl++/-`, `Ctrl+9/0`, `Ctrl+PgUp/PgDn`, `Shift+F11`, `Alt+F1`, `Shift+F2`, `Ctrl+K`, `` Ctrl+` ``, `Ctrl+Shift+1~7`, `F9`, `Ctrl+F4`, `Ctrl+Tab`, `Ctrl+휠`.
일반 브라우저 탭에서는 `Ctrl+N/W/Tab/PgUp/PgDn` 등을 브라우저가 먼저 가로챌 수 있습니다(대체: `Alt+N`, 탭 클릭). 미리보기 창/앱 모드에서는 대부분 동작합니다.

## 구조

```
.agents/skills/web-excel/          ← Agent Skill 패키지 (원본)
  SKILL.md                         이식 가능판 (명세 필드만)
  SKILL.claude-code.md             Claude Code판 (allowed-tools, 상태 주입, launch.json 안내)
  scripts/
    xl.py                          ★ 단일 진입점 CLI: serve/status/stop/launch-config
    cli_docs.py                    문서 명령: open/describe/read/write/fill/format/chart/sheet/apply/changes/undo/save …
    calc/worker.mjs                브라우저가 없을 때 같은 JS 엔진으로 수식 계산(Node 18+)
    requirements.txt               openpyxl(필수), xlrd(.xls용, 선택)
    server/
      fileio.py                    ★ 파일 바이트 I/O 단일 진입점 (DRM 훅)
      app.py                       HTTP 라우팅·인증·정적 파일
      session.py                   서버 문서 세션: op 적용·버전·로그·충돌 규칙·되돌리기
      docs_api.py                  /api/docs·/api/stream(SSE)·저장(백업·디스크 변경 감지)
      agent_api.py                 에이전트 편집(/edit → op 변환)·읽기(/range, /describe, /changes, /selection)·계산 최신화
      a1.py                        A1 참조 유틸
      agent_more.py                필터 조건·차트 편집·시트 관리·연속 채우기·복사/이동·찾기/바꾸기 등 나머지 편집
      region.py                    표 영역 판정(붙어 있는 데이터, 합계 행 제외) — filter.js와 같은 규칙
      refs.py                      수식 참조 재작성(브라우저 formula.js 이식: 채우기·행/열 이동·시트 이름)
      xlsx_io.py / xls_io.py       openpyxl ↔ JSON 모델, xlrd → JSON 모델(읽기만)
      colors.py, units.py, formula_compat.py
  assets/web/                      브라우저 UI
    index.html, css/app.css, vendor/chart.umd.min.js (Chart.js 4.4.6, MIT)
    js/app.js                      컨트롤러: 문서/창, 명령, 리본, 단축키
    js/grid.js                     Canvas 그리드 렌더링·히트 테스트·드래그
    js/editor.js                   셀 편집기/수식 입력줄
    js/formula.js / engine.js      토크나이저·파서·참조 재작성 / 계산 엔진 + 함수
    js/model.js                    통합 문서 모델, 트랜잭션 실행 취소
    js/docsync.js                  서버 문서 동기화: 낙관적 적용, 대기 중 변경 추적, 창당 이벤트 스트림 1개, 계산값 보고
    js/numfmt.js, charts.js, filter.js, clipboard.js, dialogs.js, api.js
  references/cli.md, edits.md, troubleshooting.md
.claude/skills/web-excel -> ../../.agents/skills/web-excel
run.py                             개발용 포그라운드 실행
tools/build_claude_code_skill.py   SKILL.md에서 SKILL.claude-code.md 생성
tests/                             make_samples.py, test_roundtrip.py, engine.test.mjs, e2e.cjs, sync.e2e.cjs, agent_cli.test.py,
                                   agent_full.test.py, structural.test.py, refs_parity.test.py(+.mjs),
                                   excel_compat.test.py, preserve_formulas.test.py, preserve.e2e.cjs
docs/                              project-setup.md, skill-spec.md, skill-implementation.md
```

### 설계 결정과 근거

1. **경로 기반 I/O (DRM)**: 국내 문서 DRM은 대개 허가된 프로세스에만 투명 복호화를 제공하므로, 브라우저(`<input type=file>`/드롭)가 아닌 Python 프로세스가 직접 파일을 열어야 합니다. 서버는 zip 시그니처(`PK`)가 아니면 "DRM 훅을 지정하라"는 오류를 반환합니다. 쓰기도 임시 파일+rename 없이 같은 경로에 직접 기록합니다(파일 식별자를 추적하는 DRM 대비).
2. **원본 패치 저장**: xlsx 저장 시 원본을 다시 로드해 시트 내용만 교체하므로 UI가 모델링하지 않는 조건부 서식·데이터 유효성·이미지·인쇄 설정·정의된 이름·VBA(xlsm)가 보존됩니다. 수식 결과는 캐시 값으로 함께 기록하고(계산하지 않는 프로그램도 값을 봄), `fullCalcOnLoad`를 설정해 Excel이 열 때 다시 계산합니다.
3. **재계산**: 매 변경 시 전체 재계산(메모이제이션)을 수행하되, 재귀 깊이를 제한하고 가장 깊은 미계산 셀부터 다시 시작하는 반복 심화 방식으로 20,000단계 체인도 스택 오버플로 없이 계산합니다(Sestoft의 Corecalc 보고서가 다루는 스프레드시트 재계산·의존 그래프 문제 참고 [4]).
4. **Excel 호환 세부 규칙**: 1900 윤년 버그를 포함한 날짜 일련번호 [3], 15자리 유효숫자 비교, 단항 `-`가 `^`보다 우선(`=-2^2`=4), 숫자<텍스트<논리값 비교 순서, `_xlfn.` 미래 함수 접두사 [2].
5. **보안**: 기본 127.0.0.1 바인딩, Host 헤더 검사(DNS 리바인딩 방지), 변경 API는 커스텀 헤더 `X-WebExcel` 필수(교차 출처 요청은 CORS 프리플라이트에서 차단됨 [8]), 선택적 `--root` 샌드박스.
6. **동기화**: 서버가 문서를 소유하고 변경(op)에 버전을 붙여 순서를 정합니다. 같은 셀은 서버에 나중에 도착한 쪽이 이기고, 좌표를 바꾸는 변경(행 삽입·정렬 등)은 그 사이 다른 클라이언트가 바꿨으면 거절(409)해 다시 불러오게 합니다. 창마다 이벤트 스트림(SSE)을 하나만 열어 브라우저의 호스트당 연결 수 제한을 피합니다. 자세한 내용은 docs/skill-implementation.md.

## 테스트

```bash
python -m unittest tests.test_roundtrip   # Python 읽기/쓰기 왕복 (xlsx 패치, 차트, xls 읽기 전용, _xlfn)
node tests/engine.test.mjs                # 수식 엔진 70개 검사
python3 run.py --port 8765 --root . & node tests/e2e.cjs   # 브라우저 E2E 37개 (npm i playwright 필요)
# 브라우저 테스트 4개(e2e → sync → preserve → startup)는 새로 띄운 서버에서 한 번씩 실행 (e2e는 열린 문서 수 등 초기 상태를 가정)
node tests/sync.e2e.cjs                   # 서버 동기화 17개 (에이전트↔브라우저, 동시 편집, 다중 창, 되돌리기)
python3 tests/agent_cli.test.py           # 에이전트 CLI 시나리오 39개 (브라우저 없이, 저장 파일을 openpyxl로 확인)
python3 tests/structural.test.py          # 행/열 삽입·삭제·정렬·충돌 좌표 변환 20개
python3 tests/region.test.py              # A1이 아닌 곳에서 시작하는 표(제목·빈 행·합계 행)
python3 tests/agent_full.test.py          # UI 기능 전부를 에이전트 명령으로 (대응표: docs/agent-coverage.md)
python3 tests/refs_parity.test.py         # 수식 참조 재작성 JS↔Python 동등성 (무작위 1만 건)
python3 tests/preserve_formulas.test.py    # 계산할 수 없는 함수: 열 때 경고, 고치지 않은 수식은 원래대로 저장
node tests/preserve.e2e.cjs               # 같은 내용을 미리보기에서 (경고 표시, F2+Enter, 수식 수정 경고)
node tests/startup.e2e.cjs                # 미리보기 시작·재연결 중 에이전트가 열고 닫은 문서 (탭 추가·제거, ?doc 포커스, 빈 통합 문서 교체)
node tests/auth.e2e.cjs                   # 인증: 다른 사이트 iframe + #t 토큰, 토큰 없을 때 안내, 서버 2개, 재시작 (서버를 스스로 띄움)
python3 tests/excel_compat.test.py        # 저장한 수식 229개: _xlfn. 접두사·캐시 값·LibreOffice 재계산 대조 (soffice 필요)
```

## 알려진 제한

- **xls는 읽기 전용**: xlrd는 BIFF 수식 텍스트를 제공하지 않아 계산된 값만 읽습니다. 저장하면 같은 이름의 .xlsx를 새로 만들고 원본 .xls는 그대로 둡니다(xls로 다시 쓰면 원본 수식이 값으로 바뀌므로).
- 조건부 서식·이미지·피벗·스파크라인은 **표시하지 않지만 xlsx 저장 시 보존**됩니다. 행 삽입 시 이들의 범위는 이동하지 않습니다.
- 동적 배열 수식은 결과를 분산(spill)하지 않고 왼쪽 위 값만 표시합니다.
- 차트 세부 서식(개별 색, 데이터 레이블, 보조 축)은 수정 시 기본 스타일로 다시 생성됩니다.
- 셀 단위 "셀 밀기" 삽입/삭제는 다른 수식의 참조를 조정하지 않습니다(행/열 전체 삽입·삭제는 조정).

## 참고 문헌

1. Ecma International, *ECMA-376: Office Open XML File Formats* (SpreadsheetML: 셀·스타일·테마 색 tint·numFmt·autoFilter·DrawingML 차트). https://ecma-international.org/publications-and-standards/standards/ecma-376/
2. Microsoft, *[MS-XLSX] Excel Extensions to the Office Open XML SpreadsheetML File Format*, §2.2.3 Functions(미래 함수 `_xlfn.`/`_xlws.`). https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/5d1b6d44-6fc1-4ecd-8fef-0b27406cc2bf
3. Microsoft Support, *Excel incorrectly assumes that the year 1900 is a leap year* (KB214326). https://support.microsoft.com/kb/214326
4. P. Sestoft, *A Spreadsheet Core Implementation in C#*, IT University of Copenhagen, Technical Report TR-2006-91, 2006. https://www.itu.dk/~sestoft/funcalc/ITU-TR-2006-91.pdf
5. Microsoft, *[MS-XLS] Excel Binary File Format (.xls) Structure* (xls 읽기). https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/cd03cb5f-ca02-4934-a391-bb674cb8aa06
6. Microsoft Support, *Keyboard shortcuts in Excel*. https://support.microsoft.com/en-us/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f
7. WHATWG, *HTML Living Standard — Broadcasting to other browsing contexts (BroadcastChannel)*. https://html.spec.whatwg.org/multipage/web-messaging.html#broadcasting-to-other-browsing-contexts
8. WHATWG, *Fetch Standard — CORS-preflight fetch*. https://fetch.spec.whatwg.org/#cors-preflight-fetch
9. W3C, *UI Events — Composition Events* (IME 입력 처리). https://www.w3.org/TR/uievents/#events-compositionevents
10. Microsoft Support, *Number format codes*. https://support.microsoft.com/en-us/office/number-format-codes-5026bbd6-04bc-48cd-bf33-80f18b4eae68

## 라이선스

MIT (`LICENSE`). 함께 넣은 외부 코드는 Chart.js(MIT, `assets/web/vendor/`) 하나이고 고지는
`.agents/skills/web-excel/THIRD_PARTY_NOTICES.md`에 있다. Python 패키지(openpyxl, xlrd)는 사용자가 PyPI에서 설치한다.
