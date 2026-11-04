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
