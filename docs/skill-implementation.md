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
