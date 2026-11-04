# web-excel

로컬 Python 서버 + 브라우저 UI로 동작하는 **엑셀 호환 웹 스프레드시트**입니다.
에이전트 SW(Claude Code 데스크톱, VS Code/Cursor Simple Browser 등)의 미리보기 창에서 `http://127.0.0.1:8765`로 접속해 xlsx/xls 파일을 읽고 쓸 수 있습니다.

## 실행

```bash
pip install -r requirements.txt          # openpyxl, xlrd, xlwt (그 외 의존성 없음: 표준 라이브러리 HTTP 서버)
python run.py --start ./samples          # → http://127.0.0.1:8765
python tests/make_samples.py             # 예제 파일(samples/sample.xlsx, sample.xls) 생성
```

| 옵션 | 설명 |
|---|---|
| `--port 8765` / `--host 127.0.0.1` | 바인딩 주소 (기본: 로컬 전용) |
| `--start DIR` | 파일 열기 대화상자 시작 폴더 |
| `--root DIR` | 이 폴더 밖 파일 접근 차단 (선택) |
| `--allow-host NAME` | 프록시/원격 미리보기에서 다른 Host 헤더로 접속할 때 허용 |
| `?open=<경로>` | URL 파라미터로 시작 시 파일 열기 |

Claude Code 데스크톱 미리보기용 설정은 `.claude/launch.json`에 있습니다.

## 요구사항 대응

| # | 요구사항 | 구현 |
|---|---|---|
| 1 | 로컬 서버 + 미리보기 창에서 xlsx/xls 읽기·쓰기 | `run.py` → `server/app.py` (stdlib `ThreadingHTTPServer`). xlsx/xlsm: openpyxl, xls: xlrd(읽기)/xlwt(쓰기) |
| 2 | **DRM: 파일 읽기는 반드시 Python 함수** | 브라우저는 파일 바이트에 접근하지 않고 **경로만** 전송. 모든 디스크 I/O는 `server/fileio.py`의 `read_bytes()/write_bytes()` 한 곳을 통과. 드래그&드롭 업로드는 의도적으로 차단. DRM SDK 호출이 필요하면 `WEBEXCEL_READ_HOOK=pkg.mod:func`, `WEBEXCEL_WRITE_HOOK=...`로 교체 |
| 3 | 웹에서 조작 가능한 스프레드시트 | Canvas 가상화 그리드(1,048,576×16,384), 틀 고정, 병합, 행/열 크기 조절·숨기기·자동 맞춤, 채우기 핸들(선형 추세/요일·월 목록/숫자 접미사), 복사·붙여넣기(Excel과 TSV/HTML 상호 운용), 실행 취소/다시 실행 |
| 4 | 여러 탭/윈도우 | 상단 통합 문서 탭(다중 파일), 하단 시트 탭(추가·삭제·이름 변경·복사·이동·드래그·탭 색·숨기기). **새 창**: 같은 문서를 다른 창에서 열고 `BroadcastChannel`로 실시간 동기화(셀/서식/시트 구조) |
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
run.py                  서버 실행
server/
  fileio.py             ★ 파일 바이트 I/O 단일 진입점 (DRM 훅)
  app.py                HTTP 라우팅: /api/info, /api/fs, /api/open, /api/save, /api/session
  xlsx_io.py            openpyxl ↔ JSON 모델 (원본 패치 방식 저장)
  xls_io.py             xlrd/xlwt ↔ JSON 모델
  colors.py, units.py, formula_compat.py
static/
  index.html, css/app.css, vendor/chart.umd.min.js (Chart.js 4.4.6, MIT)
  js/app.js             컨트롤러: 문서/창, 명령, 리본, 단축키
  js/grid.js            Canvas 그리드 렌더링·히트 테스트·드래그
  js/editor.js          셀 편집기/수식 입력줄 (입력·편집·참조 지정 모드, 자동 완성)
  js/formula.js         토크나이저/파서/참조 재작성
  js/engine.js          계산 엔진 + 함수 라이브러리
  js/model.js           통합 문서 모델, 트랜잭션 실행 취소, 행/열 이동
  js/numfmt.js          Excel 표시 형식
  js/charts.js, filter.js, clipboard.js, dialogs.js, sync.js, api.js
tests/                  make_samples.py, test_roundtrip.py, engine.test.mjs, e2e.cjs
```

### 설계 결정과 근거

1. **경로 기반 I/O (DRM)**: 국내 문서 DRM은 대개 허가된 프로세스에만 투명 복호화를 제공하므로, 브라우저(`<input type=file>`/드롭)가 아닌 Python 프로세스가 직접 파일을 열어야 합니다. 서버는 zip 시그니처(`PK`)가 아니면 "DRM 훅을 지정하라"는 오류를 반환합니다. 쓰기도 임시 파일+rename 없이 같은 경로에 직접 기록합니다(파일 식별자를 추적하는 DRM 대비).
2. **원본 패치 저장**: xlsx 저장 시 원본을 다시 로드해 시트 내용만 교체하므로 UI가 모델링하지 않는 조건부 서식·데이터 유효성·이미지·인쇄 설정·정의된 이름·VBA(xlsm)가 보존됩니다. 수식의 캐시 값은 기록하지 않고 `fullCalcOnLoad`를 설정해 Excel이 열 때 재계산합니다.
3. **재계산**: 매 변경 시 전체 재계산(메모이제이션)을 수행하되, 재귀 깊이를 제한하고 가장 깊은 미계산 셀부터 다시 시작하는 반복 심화 방식으로 20,000단계 체인도 스택 오버플로 없이 계산합니다(Sestoft의 Corecalc 보고서가 다루는 스프레드시트 재계산·의존 그래프 문제 참고 [4]).
4. **Excel 호환 세부 규칙**: 1900 윤년 버그를 포함한 날짜 일련번호 [3], 15자리 유효숫자 비교, 단항 `-`가 `^`보다 우선(`=-2^2`=4), 숫자<텍스트<논리값 비교 순서, `_xlfn.` 미래 함수 접두사 [2].
5. **보안**: 기본 127.0.0.1 바인딩, Host 헤더 검사(DNS 리바인딩 방지), 변경 API는 커스텀 헤더 `X-WebExcel` 필수(교차 출처 요청은 CORS 프리플라이트에서 차단됨 [8]), 선택적 `--root` 샌드박스.
6. **다중 창**: 같은 출처 창 간 `BroadcastChannel`로 트랜잭션 패치를 전송 [7]. 팝업이 외부 브라우저로 열리는 환경을 위해 서버 세션(`/api/session`)으로 초기 상태를 넘기는 대체 경로가 있습니다.

## 테스트

```bash
python -m unittest tests.test_roundtrip   # Python 읽기/쓰기 왕복 (xlsx 패치, 차트, xls, _xlfn)
node tests/engine.test.mjs                # 수식 엔진 70개 검사
python run.py & node tests/e2e.cjs        # 브라우저 E2E 34개 (npm i playwright 필요)
```

## 알려진 제한

- **xls 읽기**: xlrd는 BIFF 수식 텍스트를 제공하지 않아 계산된 값만 읽습니다(수식을 유지하려면 xlsx로 저장). xls 쓰기는 65,536행×256열, 56색 팔레트, 차트/자동 필터 미저장.
- 조건부 서식·이미지·피벗·스파크라인은 **표시하지 않지만 xlsx 저장 시 보존**됩니다. 행 삽입 시 이들의 범위는 이동하지 않습니다.
- 동적 배열 수식은 결과를 분산(spill)하지 않고 왼쪽 위 값만 표시합니다.
- 차트 세부 서식(개별 색, 데이터 레이블, 보조 축)은 수정 시 기본 스타일로 다시 생성됩니다.
- 셀 단위 "셀 밀기" 삽입/삭제는 다른 수식의 참조를 조정하지 않습니다(행/열 전체 삽입·삭제는 조정).

## 참고 문헌

1. Ecma International, *ECMA-376: Office Open XML File Formats* (SpreadsheetML: 셀·스타일·테마 색 tint·numFmt·autoFilter·DrawingML 차트). https://ecma-international.org/publications-and-standards/standards/ecma-376/
2. Microsoft, *[MS-XLSX] Excel Extensions to the Office Open XML SpreadsheetML File Format*, §2.2.3 Functions(미래 함수 `_xlfn.`/`_xlws.`). https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/5d1b6d44-6fc1-4ecd-8fef-0b27406cc2bf
3. Microsoft Support, *Excel incorrectly assumes that the year 1900 is a leap year* (KB214326). https://support.microsoft.com/kb/214326
4. P. Sestoft, *A Spreadsheet Core Implementation in C#*, IT University of Copenhagen, Technical Report TR-2006-91, 2006. https://www.itu.dk/~sestoft/funcalc/ITU-TR-2006-91.pdf
5. Microsoft, *[MS-XLS] Excel Binary File Format (.xls) Structure* (BIFF8 FORMULA 레코드의 캐시 결과값 형식). https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/cd03cb5f-ca02-4934-a391-bb674cb8aa06
6. Microsoft Support, *Keyboard shortcuts in Excel*. https://support.microsoft.com/en-us/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f
7. WHATWG, *HTML Living Standard — Broadcasting to other browsing contexts (BroadcastChannel)*. https://html.spec.whatwg.org/multipage/web-messaging.html#broadcasting-to-other-browsing-contexts
8. WHATWG, *Fetch Standard — CORS-preflight fetch*. https://fetch.spec.whatwg.org/#cors-preflight-fetch
9. W3C, *UI Events — Composition Events* (IME 입력 처리). https://www.w3.org/TR/uievents/#events-compositionevents
10. Microsoft Support, *Number format codes*. https://support.microsoft.com/en-us/office/number-format-codes-5026bbd6-04bc-48cd-bf33-80f18b4eae68
