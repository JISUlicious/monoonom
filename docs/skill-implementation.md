# Agent Skill 구현 기록

[skill-spec.md](skill-spec.md)의 계획(P0–P2: 최소 기능)을 단계별로 구현한 기록. 단계마다 무엇을 했는지, 결정과 이유, 검증 방법, spec과 달라진 점을 남긴다.

## P0. 패키징·서버 수명주기 (2026-11-04)

### 한 일
- 코드를 skill 폴더로 이동: `server/` → `.agents/skills/web-excel/scripts/server/`, `static/` → `.agents/skills/web-excel/assets/web/`.
- `.claude/skills/web-excel` → `../../.agents/skills/web-excel` 심볼릭 링크(이 저장소를 연 Claude Code가 프로젝트 skill로 인식).
- 의존성 동봉: `scripts/_vendor/`에 openpyxl 3.1.5, et_xmlfile 2.0.0(MIT), xlrd 2.0.2, xlwt 1.3.0(BSD)과 라이선스. 약 2.3MB.
- `scripts/xl.py`: `serve`(분리 실행·재사용), `serve --foreground [--exit-with-parent]`, `status [--brief]`, `stop [--force]`, `launch-config`.
- 서버(`server/app.py`): 라우트 표 방식으로 재작성, 토큰 인증, `/api/health`, `/api/shutdown`, 활동 시각 기록.
- `SKILL.md`(이식판), `SKILL.claude-code.md`(Claude Code판), `references/troubleshooting.md`.
- `run.py`는 개발용 래퍼(`xl.py serve --foreground`)로 변경. 이전 다중 창 대체 경로였던 `/api/session`은 제거(P1에서 서버 경유 동기화로 대체).

### 결정과 이유
- **원본 위치 `.agents/skills/`**: Codex·VS Code·Cursor·Gemini CLI가 공통으로 읽는 경로(spec §2.1.1). Claude Code는 링크로 연결.
- **상태 파일 `<작업 경로>/.web-excel/server.json`**: 작업 경로마다 서버 하나. 임시 파일 + `os.replace`로 원자적 기록, 권한 0600, 폴더에 `.gitignore`(`*`)를 자동 생성해 토큰이 커밋되지 않게 함.
- **포트**: 8765부터 빈 포트를 찾음. 고정 포트만 쓰면 다른 작업 경로의 서버나 잔존 프로세스와 충돌.
- **인증**: CLI는 `X-WebExcel-Token` 헤더. 브라우저는 `/` 로드 시 받는 `HttpOnly; SameSite=Strict` 쿠키(+POST에는 기존 `X-WebExcel` 헤더). `/api/health`는 토큰 없이 `{ok, app}`만, 토큰이 있으면 pid·root까지 반환(`alive()`가 pid로 같은 서버인지 확인).
- **수명주기 자기방어**: 클라이언트·요청 없이 `--idle-timeout`(기본 3600초) 지나면 종료, `--exit-with-parent`면 부모 PID가 바뀔 때 종료(POSIX만; Windows는 미지원으로 둠).
- **분리 실행**: POSIX `start_new_session=True`, Windows `DETACHED_PROCESS | CREATE_NEW_PROCESS_GROUP | CREATE_NO_WINDOW`. 로그는 `.web-excel/server.log`.
- **`SKILL.claude-code.md`를 패키지 안에 둠**: 검증기는 링크되지 않은 .md라고 경고만 한다(오류 아님). 이름이 `SKILL.md`가 아니라 "SKILL.md 2개" 오류에는 걸리지 않는다. 링크로 연결된 Claude Code에는 이식판 `SKILL.md`가 로드되며 그대로 동작한다.

### 검증
- `serve` → 새 서버 시작, 재실행 시 재사용, `status --brief`, 토큰 없는 `/api/info` 401, 쿠키로 200, `launch-config`가 기존 구성을 보존하며 엔트리 추가, `stop`.
- `--idle-timeout 3` → 약 3초 뒤 종료 로그 확인. Python 부모를 강제 종료 → `--exit-with-parent` 서버가 2초 안에 종료.
- 발견·수정한 버그: 없는 POST 경로로 요청이 오면 본문을 읽지 않아 같은 keep-alive 연결의 다음 요청이 깨짐 → 라우팅 전에 항상 본문을 읽도록 수정.
- `validate_skill.py --portable` 통과(오류 0, 경고 1: `SKILL.claude-code.md` 미링크). Python 왕복 5/5, 엔진 70/70, 브라우저 E2E 34/34.

### spec과 달라진 점
- `skills-ref validate`는 이 세션 네트워크에서 설치 불가로 아직 돌리지 못함(P6에서).
- Windows 분리 실행과 `--exit-with-parent`는 Linux에서만 검증.

## P1. 서버가 문서를 소유하는 동기화 (2026-11-04)

### 한 일
- `server/session.py`: `DocSession`(시트·셀·서식을 서버 메모리에 보관, 버전, op 로그 5,000개, 되돌리기용 이전 상태), `Registry`(열린 문서 목록), `Hub`(SSE 구독자 큐), 저장 전 백업.
- `server/docs_api.py`: `GET/POST /api/docs`, `GET /api/docs/{id}`·`/snapshot`, `POST /ops`·`/presence`·`/calc`·`/save`·`/undo`, `DELETE /api/docs/{id}`, 창당 이벤트 스트림 `GET /api/stream` + `POST /api/stream/subscribe|unsubscribe`.
- 브라우저: `js/docsync.js`(문서별 동기화 + 창 하나당 이벤트 스트림), `model.js`에 시트 JSON 변환·원격 op 적용(`applyRemote`)·좌표 변경 표시(`markStructural`), `app.js`의 열기/새로 만들기/저장/닫기/새 창을 서버 문서 기준으로 변경, 다른 클라이언트가 바꾼 셀을 노란색으로 잠깐 강조하고 상태 표시줄에 표시.
- 제거: `js/sync.js`(BroadcastChannel), `/api/open`·`/api/save`·`/api/new`(문서 API로 대체).

### op 형식
`setCells {sheet, cells:[[r,c,cell|null]]}`, `setSheetMeta {sheet, meta}`, `sheets {order, add}`, `setWorkbook {props}`. 시트는 이름이 아닌 id로 지칭. 브라우저는 실행 취소 기록(트랜잭션 항목)을 그대로 이 op들로 바꿔 보낸다(변경 후 상태를 보내므로 서버는 의미를 해석할 필요가 없다).

### 결정과 이유
- **같은 셀 충돌은 나중 도착 우선(LWW)**: 서버가 순서를 정하므로 타임스탬프가 필요 없다(Figma 방식, spec [4]). 브라우저는 자기 변경의 echo가 SSE로 돌아올 때까지 그 셀을 "대기 중"으로 두고, 그 사이 도착한 남의 변경은 그 셀에 적용하지 않는다. 서버는 내 변경을 나중에 적용하므로 결과가 서버와 같아진다.
- **좌표를 바꾸는 변경은 거절 후 재동기화**: 행/열 삽입·삭제, 정렬, 시트 이름 변경, 셀 이동, 중복 제거, 시트 구조 변경이 해당. 기준 버전 이후 다른 클라이언트의 변경이 있으면 409 → 브라우저가 스냅샷을 다시 받고 알림. spec의 MVP 규칙 그대로(P3에서 좌표 변환으로 개선 예정).
- **창당 이벤트 스트림 1개**: 처음엔 문서마다 EventSource를 열었는데, 두 번째 창의 요청이 멈췄다. 브라우저는 HTTP/1.1에서 호스트당 연결을 약 6개로 제한하고 이 한도를 탭 전체가 공유하기 때문. 창마다 `/api/stream` 하나를 열고 문서는 POST로 구독하도록 바꿨다. 구독 시 재전송(since 이후 op)과 구독 등록을 문서 잠금 안에서 함께 해 빠지거나 중복되는 이벤트가 없게 했다.
- **브라우저가 계산값을 보고**: 계산 엔진이 JS라 서버는 수식 결과를 모른다. 열린 창이 재계산 후 바뀐 수식 결과만 `/calc`로 보내고, 서버는 버전과 함께 저장한다(P2의 `read`가 사용, 저장 시 xls 캐시 값으로도 사용).
- **저장 안전장치**: 같은 경로에 덮어쓸 때 열 때의 (mtime, size)와 다르면 409(다른 프로그램이 저장). 덮어쓰기 전 `.web-excel/backups/`에 사본(DRM 정책상 사본 금지면 `backup:false`). 유휴 종료 시 저장 안 된 문서는 `.web-excel/recovery/`에 JSON으로 남김.
- **에이전트 변경 되돌리기**: 각 batch에 이전 상태를 기록해 `POST /undo {batchId}`로 되돌림. 이후 같은 셀이 바뀌었으면 409(force로 강제).
- **창을 닫아도 문서는 서버에 남음**: 그래서 "저장 안 하고 나가기" 경고(beforeunload)를 없앴다. 탭의 × 는 서버 문서를 닫는다(저장 안 됐으면 확인).

### 검증
- `tests/sync.e2e.cjs` 11개 통과: 에이전트 op가 브라우저에 반영·강조, 브라우저 입력이 서버 스냅샷에 반영, 계산값 보고 버전, 같은 셀 동시 쓰기 수렴, 오래된 기준 버전의 좌표 변경 409, 두 번째 창 실시간 반영, 에이전트가 연 파일이 탭으로 표시, batch 되돌리기.
- 기존 `tests/e2e.cjs` 34개(열기·필터·정렬·차트·저장·재열기·xls·다중 창 포함), Python 왕복 5개, 엔진 70개 통과.

### spec과 달라진 점
- 문서별 SSE(`/api/docs/{id}/events`) 대신 창당 스트림을 기본으로 했다(문서별 엔드포인트는 CLI용으로 남김).
- 다른 클라이언트의 선택 영역 표시는 아직 하지 않음(서버에 presence는 저장, CLI `selection`에서 사용 예정).

## P2. 에이전트 CLI (2026-11-04)

### 한 일
- `server/agent_api.py`: `POST /api/docs/{id}/edit`(사람 수준의 편집 → 상태 op 변환, 한 batch), `GET /range`, `/describe`, `/changes`, `/selection`, 수식 결과 최신화(`ensure_calc`).
- `server/a1.py`: A1 참조 파싱, 시트 이름 변경 시 다른 시트 수식·차트 범위·정의된 이름 재작성.
- `scripts/calc/worker.mjs`: 브라우저가 쓰는 엔진(`model.js`)을 Node에서 그대로 돌려 수식 결과 계산.
- `scripts/cli_docs.py`: `ls, open, new, docs, describe, read, write, fill, clear, format, merge, unmerge, width, height, hide, unhide, freeze, filter, sheet, chart, apply, changes, selection, undo, save, close` (+ 기존 서버 명령).
- `SKILL.md` 본문 완성(작업 순서·명령표·값 해석 규칙·계산·종료 코드, 85줄·약 1.7k 토큰), `references/cli.md`·`edits.md`, `tools/build_claude_code_skill.py`로 `SKILL.claude-code.md` 생성.

### 결정과 이유
- **편집은 서버에서 op로 변환**: CLI는 얇게 두고(stdlib만), 범위 해석·값 해석·서식 병합은 서버가 한다. 결과는 P1과 같은 상태 op라서 브라우저에 그대로 실시간 반영되고 batch 단위로 되돌릴 수 있다.
- **값은 Excel 입력처럼 해석**: `=`는 수식, `2026-03-15`는 날짜(일련번호 + yyyy-mm-dd 서식), `15%`, `1,234`, `₩1,000`, `TRUE`. `'`로 시작하거나 `--raw`면 문자열 그대로. 에이전트가 날짜를 문자열로 넣는 실수를 막기 위함.
- **읽기 결과는 사람이 읽는 형태**: 날짜 서식 숫자는 ISO 문자열, 실수는 15자리 유효숫자(`110.39999999999999` → `110.4`), 전체 열 범위는 사용 범위로 자르고 기본 200행에서 끊어 다음 범위를 알려 준다(토큰 절약).
- **채우기(fill)는 상대 참조 이동**: 한 수식을 범위에 채우면 행/열마다 상대 참조를 옮긴다(Python 정규식 구현, 문자열·따옴표 시트 이름은 건드리지 않음). 전체 참조 재작성(행 삽입 등)은 여전히 P3.
- **수식 결과의 출처 3단계**: ① 미리보기 창이 보고한 값(최대 2초 대기) ② 없으면 Node로 같은 엔진 실행 ③ 둘 다 없으면 `stale`로 표시하고 파일의 마지막 계산값 사용. 응답에 `calc.status/source`를 함께 준다.
- **문서 지정**: 문서가 하나면 `--doc` 생략. 여러 개면 id·이름·경로로 지정하고, 모호하면 목록과 함께 exit 2.
- **상대 경로는 작업 경로 기준**: 에이전트의 셸 위치와 서버 root가 다를 수 있어서 `open`, `save --as`, `ls`, `apply`, `--tsv` 모두 root 기준으로 해석.
- **종료 코드**: 0 성공, 2 사용법, 3 서버 없음, 4 충돌(409), 5 파일·DRM·시트 오류 — 에이전트가 다음 행동을 고르기 쉽게.

### 검증
- `tests/agent_cli.test.py` 34개 통과: 기동·재사용, 열기, describe, 읽기(날짜 ISO), fill의 상대 참조 이동, 브라우저 없이 Node 계산, Excel식 값 해석, 서식·되돌리기, 열 너비, 차트, 시트 이름 변경 시 다른 시트 수식 갱신, 교차 시트 수식, `edits.md` 예제 그대로 적용, 변경 이력, 백업 저장 후 openpyxl로 시트 이름·수식·차트·열 너비·날짜 확인, 오류 종료 코드, 저장 안 된 상태의 stop 거부.
- 기존 스위트 모두 통과: Python 왕복 5, 엔진 70, 브라우저 E2E 34, 동기화 11.
- 테스트로 찾아 고친 버그: (1) 상대 경로를 CLI 위치 기준으로 해석, (2) 빈 시트에 한 batch로 값을 쓰고 서식·차트를 만들면 앞서 쓴 셀을 못 보고 범위를 A1로 잘라 버림, (3) 포트 검사에 `SO_REUSEADDR`가 없어 TIME_WAIT 포트를 사용 중으로 판단.
- `validate_skill.py --portable` 통과(경고 1: `SKILL.claude-code.md` 미링크, 의도된 것).

### spec과 달라진 점·남은 것
- 행/열 삽입·삭제와 정렬은 에이전트 CLI에 아직 없음(P3: 참조 재작성 Python 포팅 필요). 미리보기에서 사용자는 할 수 있음.
- 미리보기에 다른 클라이언트의 선택 영역 표시, Windows 실기기 검증, `skills-ref validate`, 트리거 평가(P6)는 남음.

## 이후 수정 (2026-11-04)

- **차트 DPI 버그**: 2배 해상도 화면에서 차트가 상자의 왼쪽 위 1/4에만 그려졌다. 캔버스 픽셀 크기를 직접 `크기×devicePixelRatio`로 잡은 뒤 Chart.js가 다시 배율을 곱했기 때문. 캔버스에는 논리 크기만 주도록 수정(`charts.js`). 이전 E2E는 1배 화면이라 못 잡았고, 아래 캡처(2배)에서 발견.
- 에이전트가 CLI로 G열에 수식·서식을 넣는 순간의 미리보기 화면: [agent-edit-preview.png](agent-edit-preview.png) (바뀐 셀이 노란색으로 잠깐 강조되고 상태 표시줄에 "에이전트: …"가 뜬다).

## P3. 구조 변경(행/열 삽입·삭제, 정렬)과 충돌 좌표 변환 (2026-11-04)

### 한 일
- `server/refs.py`: 브라우저 `formula.js`의 토큰화·참조 재작성(`translate`, `adjust`, 시트 이름 변경)을 Python으로 그대로 이식. `agent_api.py`의 정규식 기반 구현을 이것으로 교체.
- 에이전트 편집 `insertRows/deleteRows/insertCols/deleteCols/sort`: 모든 시트의 수식, 차트 계열·위치, 병합, 필터 범위, 숨긴 행·필터로 숨긴 행, 행 높이·열 너비, 정의된 이름을 함께 이동(브라우저 `Workbook.shift()`와 같은 규칙). 정렬은 필터로 숨긴 행 제외, 빈 셀 맨 뒤, 수식은 행과 함께 이동.
- CLI: `rows insert|delete`, `cols insert|delete`, `sort RANGE --by C:desc …`. `/edit`는 구조 변경을 각각 별도 batch로 나눠 적용(뒤 편집의 좌표 기준이 바뀌므로).
- 충돌 규칙 개선(`session.apply`): 오래된 기준 버전의 **셀 편집**이 다른 클라이언트의 행/열 삽입·삭제 뒤에 도착하면, 거절하지 않고 그 이동만큼 좌표와 수식을 옮겨 적용. 지워진 행/열에 쓴 값은 버림. 정렬·시트 삭제 뒤이거나, 이동된 시트의 속성(열 너비 등)을 통째로 바꾸는 batch는 여전히 409.
- 브라우저: 행/열 삽입·삭제 batch에 이동 정보(`shifts`)를 함께 보냄. 내 변경이 서버에서 옮겨졌거나, 내 변경이 대기 중일 때 원격 구조 변경이 오면 대기분이 다 확인된 뒤 조용히 다시 불러옴(편집 중이면 끝날 때까지 미룸).
- 구조 변경 batch 되돌리기: 반대 방향 이동 정보로 기록해 이후 충돌 변환에도 쓰임.

### 결정과 이유
- **재작성기를 새로 짜지 않고 이식**: 같은 수식이 브라우저와 서버에서 다르게 바뀌면 동기화가 깨진다. JS의 토큰화 순서(문자열 → 오류 → 숫자 → 참조 → 식별자)와 "알 수 없는 문자면 수식을 그대로 둔다"는 동작까지 맞췄다.
- **셀 편집만 변환**: 셀 좌표는 이동 규칙이 명확하지만, 시트 속성 op는 상태 전체를 보내므로 변환하면 이동 결과를 덮어쓴다. 정렬은 행 순서를 섞으므로 "같은 내용의 새 위치"를 정할 수 없다. 이 경우만 거절하고 다시 불러오게 했다.
- **구조 변경마다 별도 batch**: Excel에서 행 삽입 뒤 입력하는 것처럼, 뒤따르는 편집은 삽입 후 좌표로 쓰는 것이 자연스럽다. 되돌리기도 단계별로 가능.

### 검증
- `tests/refs_parity.test.py`: 무작위 수식 10,000개(따옴표·한글·숫자 시트 이름, 열/행 범위, 문자열, 오류, 함수 이름 LOG10 등 포함)를 JS와 Python으로 재작성 → 불일치 0(4,391개가 실제로 바뀜).
- `tests/structural.test.py` 20개: 삽입·삭제 시 데이터·합계 수식·다른 시트 수식·필터 범위·차트 계열·병합 이동, 참조 행 삭제 시 범위 축소, 정렬, 오래된 셀 편집의 좌표·수식 변환, 삭제된 행 편집 버림, 정렬 뒤 409, 이동된 시트 속성 변경 409, 구조 batch 되돌리기.
- `tests/sync.e2e.cjs` 14개(+3): 에이전트 행 삽입이 미리보기에 반영, **사용자 입력 요청을 일부러 지연시킨 사이 에이전트가 위에 행 삽입 → 서버와 브라우저 모두 옮겨진 위치에 같은 값**, 브라우저 행 삽입이 에이전트 읽기에 반영.
- `tests/agent_cli.test.py` 39개(+5), 브라우저 E2E 34, Python 왕복 5, 엔진 70 통과. 검증기 오류 0.

### 남은 것
- 서로 다른 클라이언트가 **동시에 둘 다 구조 변경**을 하면 늦은 쪽은 여전히 409 후 다시 불러오기(드묾).
- 정렬 뒤의 동시 셀 편집은 409(위치를 정할 수 없음).

## P4–P6 마무리 (2026-11-04)

### P4. 수식 계산 출처 — P2에서 구현됨
미리보기 창의 보고값(최대 2초 대기) → Node 워커(같은 JS 엔진) → `stale` 표시. `read`·`describe` 응답에 `calc.status/source`. 추가 작업 없음.

### P5. 안전·보안
- P1–P2에서 구현: 토큰(헤더)·`SameSite=Strict` 쿠키, Host 검사, 작업 경로 샌드박스, 저장 전 백업, 디스크 변경 감지, 유휴 종료 시 복구 스냅샷, batch 되돌리기.
- 이번에 추가: **감사 로그** `<작업 경로>/.web-excel/audit.log` — 열기·편집(버전, 행위자, batch, 라벨, 셀 개수, 구조 변경 여부, 좌표 변환 여부)·저장(경로, 형식, 백업 위치)·닫기를 JSON 줄로 기록. **셀 내용은 기록하지 않음**(DRM 문서 내용이 평문으로 복사되지 않게).

### P6. 문서·평가
- `evals/evals.json`: 시나리오 4개(이익률 열 추가, 사용자 선택 영역 서식, 행 삽입+정렬+저장, 사용자 변경 요약). 스키마는 analysis-docs 작성 표준(§11)의 `skill_name` + `evals[{id, prompt, expected_output, files, expectations}]`. 스킬 루트의 `evals/`는 패키징에서 제외됨.
- `evals/trigger_queries.json`: 써야 할 질문 10개 + 키워드가 겹치는 쓰면 안 되는 질문 10개.
- 검증기: `validate_skill.py --portable` 오류 0. 공식 `skills-ref`(PyPI 0.1.1, agentskills.io / anthropics/agentskills)는 내려받았지만 **외부 코드 실행이 권한 검사에서 막혀 실행하지 않았다.** 대신 소스(`validator.py`)를 읽어 같은 규칙(허용 필드 6개, name 소문자·문자/숫자/하이픈·폴더 이름 일치·64자, description 1,024자, compatibility 500자, metadata 문자열 값)을 직접 확인: 모두 만족(description 549자, compatibility 363자). 차이는 `strictyaml` 파서뿐이라 실제 실행으로 한 번 더 확인 필요.
- **평가는 아직 실행하지 않음**: 표준대로 하려면 새 세션에서 skill 있음/없음 A/B, 트리거 질문 각 3회·60/40 분할, Haiku·Sonnet·Opus 각각 실행이 필요하다. 모델 호출 비용이 들고 이 세션에서 다른 에이전트 세션을 돌리는 것이므로 사용자 승인 후 진행.

## P7. UI 기능 전부를 에이전트 명령으로 (2026-11-04)

### 배경
P2–P3까지는 값·수식·서식·병합·크기·필터 켜기·행/열·정렬·시트 추가/이름/삭제·차트 추가/삭제만 가능했다. 미리보기에서 사용자가 할 수 있는 작업을 에이전트도 모두 할 수 있어야 한다는 요구에 맞춰 빠진 기능을 채웠다. 대응표: [agent-coverage.md](agent-coverage.md).

### 한 일
- `server/agent_more.py`: 필터 조건(값 목록·사용자 지정 조건·지우기·다시 적용, filter.js 이식), 차트 편집, 시트 복사·이동·숨기기·탭 색·눈금선, 연속 데이터 채우기(app.js의 makeSeries 이식: 선형 추세·요일/월/분기 목록·숫자 접미사·날짜 +1), 복사·선택하여 붙여넣기·행열 바꿈·이동, 셀 삽입/삭제(밀기), 중복 제거, 바꾸기, 메모, 하이퍼링크, 자동 맞춤, 서식 복사, 소수 자릿수, 전체 병합, 메모 지우기.
- 읽기·보기: `read --what text`(화면에 보이는 텍스트), `find`, `goto`(사용자 미리보기를 시트/범위로 이동), `view`(확대/축소·수식 표시).
- Node 워커에 표시 텍스트 모드 추가: 필터의 값 목록은 Excel처럼 **화면 표시 텍스트**로 비교하므로 브라우저와 같은 표시 형식 처리기(numfmt.js)를 Node에서 돌린다. Node가 없으면 Python 근사 처리(일반·소수·천 단위·%·통화 기호·날짜).

### 테스트로 찾아 고친 버그
- 시트를 복사·이동하면 서버의 "활성 시트" 번호가 다른 시트를 가리켜, 그 뒤 `--sheet` 없는 편집이 복사본에 들어감 → 활성 시트를 번호가 아닌 시트 id로 유지.
- 아무것도 바뀌지 않은 편집이 `undo None`을 안내 → "변경 없음"으로 표시.

### 검증
- `tests/agent_full.test.py` 34개: 각 기능을 CLI로 실행하고 결과를 읽어 확인, 마지막에 저장한 파일을 openpyxl로 열어 시트 순서·숨긴 시트·탭 색·눈금선·메모·하이퍼링크·자동 맞춤·병합·차트를 확인.
- `tests/sync.e2e.cjs` 17개(+3): 에이전트 `goto`가 미리보기를 해당 시트·셀로 이동, `view` 확대/축소, 에이전트 필터가 미리보기에서 행을 숨김.
- 기존 스위트 모두 통과(왕복 5, 엔진 70, 동등성 1만 건, 구조 20, CLI 39, E2E 34). 검증기 오류 0. SKILL.md 100줄·약 2.2k 토큰.

### 에이전트로 할 수 없는 것
사용자 PC 클립보드(다른 앱과 주고받기), 새 브라우저 창 열기, 각 창의 Ctrl+Z 기록(대신 서버 기록 기준 `undo BATCH`), 키보드 탐색 같은 화면 조작. 셀 밀기·이동은 UI와 마찬가지로 다른 수식의 참조를 고치지 않는다.

## 표가 A1에서 시작하지 않을 때 (2026-11-04)

### 확인 방법
제목(B2, B2:F2 병합), 빈 행, 머리글 C5:F5, 데이터 C6:F12, 바로 아래 합계 행 C13(`=SUM(D6:D12)` 등), 옆에 메모(H5)가 있는 파일(`samples/offset.xlsx`)로 UI와 에이전트 동작을 모두 확인했다.

### 발견한 문제
- **UI**: 셀 하나에서 Ctrl+Shift+L을 누르면 Excel처럼 붙어 있는 영역(C5:F13)을 잡는데, 합계 행까지 필터 범위에 들어갔다. 이어서 정렬하면 합계 행이 데이터 사이로 섞이고 `=SUM(D6:D12)`가 `#REF!`로 깨졌다.
- **UI**: 영역 확장 루프가 최대 200번만 돌아, 200행이 넘는 표는 셀 하나에서 필터/정렬할 때 잘릴 수 있었다.
- **에이전트**: `filter D7`(셀 하나)은 D7:D8, `filter C:F`(열 전체)는 제목·빈 행을 포함한 C1:F13이 됐고, `sort D7`은 아무것도 하지 않았다. `describe`는 사용 범위를 A1부터로, 제목 행을 머리글로 보고했다.

### 고친 것
- 영역 판정을 UI(`filter.js`)와 서버(`server/region.py`)에 같은 규칙으로 구현: 붙어 있는 데이터 영역을 찾고(빈 행·열에서 끊김), 끝의 **합계 행**은 필터·정렬·차트·중복 제거 대상에서 뺀다. 합계 행 판정은 첫 열 텍스트가 합계/총계/소계/평균/Total 등이거나, 위쪽 행들을 SUM·SUBTOTAL·AVERAGE·COUNT·MAX·MIN으로 묶는 수식이 있는 행.
- 에이전트 명령에 **셀 하나나 열/행 전체**를 주면 그 위치의 표로 확장. 여러 셀 범위를 명시하면(예: `filter C5:F13`) 합계 행 포함이라도 그대로 존중.
- 필터가 켜진 범위 안에서 정렬하면 필터 범위의 표(합계 제외, 머리글 유지)를 정렬. 필터 "다시 적용"은 아래에 추가된 행은 포함하되 합계 행은 포함하지 않음.
- `describe`: 사용 범위를 실제 첫 셀부터 보고하고, 시트의 **표 목록**(범위·머리글 행·합계 행)을 따로 보여 준다. 머리글과 열 유형은 가장 큰 표 기준.
- Ctrl+A는 Excel과 같이 합계 행을 포함한 영역 전체를 선택(바꾸지 않음).

### 검증
- `tests/region.test.py` 14개(영역·합계 판정, 표 목록, 셀 하나/열 전체로 필터, 조건이 표의 첫 열 기준인지, 정렬 시 합계 유지, 다시 적용, 중복 제거, 차트, 명시 범위 존중).
- `tests/e2e.cjs` 37개(+3): UI에서 셀 하나로 필터 → C5:F12, 정렬해도 합계·SUM 유지, Ctrl+A는 합계 포함.
- 나머지 스위트 모두 통과.
