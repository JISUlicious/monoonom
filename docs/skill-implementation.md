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
