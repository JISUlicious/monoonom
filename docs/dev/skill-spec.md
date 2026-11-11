# web-excel Agent Skill: 기술 조사 · 실현 가능성 · 계획 · 명세 (v0.1)

> 대상: 현재 `feature/web-excel` 구현(Python 서버 + 브라우저 UI, 약 8,000줄)을 **Agent Skill**로 포장하여,
> 에이전트가 스크립트로 로컬 서버를 띄우고 API로 작업 경로의 엑셀 파일을 조작하며,
> 사용자가 에이전트 SW의 미리보기 창(`http://localhost:{port}`)에서 하는 조작과 **실시간 동기화**되게 하는 것.

---

## 1. 결론 요약

| 질문 | 답 |
|---|---|
| Skill로 구현 가능한가? | **가능.** Skill은 "지시문(SKILL.md) + 번들 스크립트"이고, 에이전트가 셸 도구로 스크립트를 실행한다. 스크립트가 서버를 **분리(detach) 실행**하면 세션 동안 유지되며, 에이전트는 같은 스크립트의 CLI 하위 명령으로 HTTP API를 호출한다. |
| 미리보기 창 연동은? | 미리보기 창은 단순히 `localhost:{port}` 웹을 보여주므로, 서버가 UI를 같은 포트에서 서빙하면 된다. 추가 통합 불필요. |
| 사용자 조작 ↔ 에이전트 조작 동기화는? | 가능하지만 **현재 구조 변경이 필요**하다. 지금은 브라우저가 문서의 진실 원천(source of truth)이고 서버는 열기/저장만 한다. → **서버가 문서 세션을 소유**하고, 브라우저와 에이전트 CLI가 모두 "연산(op)"을 서버에 보내며, 서버가 순서를 정해 SSE로 모든 클라이언트에 방송하는 구조로 바꿔야 한다. |
| 가장 어려운 부분 | ① 서버 권한(server-authoritative) 동기화와 구조 변경(행 삽입·정렬) 충돌 처리, ② 브라우저가 없을 때의 **수식 계산**(엔진이 JS). Skill 포장 자체는 쉽다. |
| 예상 공수 | 약 **24–36 인일**(1인 기준, 아래 §6). 기존 코드 재사용률은 UI·엔진·파일 I/O 기준 약 70%. |

---

## 2. 전제와 확인된 사실

### 2.1 Agent Skills (확인한 범위)

출처: Claude Code 문서 [1] 원문, 사용자의 `analysis-docs` 저장소 [12], 그리고 **analysis-docs 세션의 회신(2026-11-04)** [15]. 이 세션에서는 agentskills.io가 막혀 있지만, analysis-docs 세션이 명세 원문 [16]을 직접 열어 [12]와 대조해 주었다.

- **구조**: 폴더당 `SKILL.md` 정확히 1개(루트). 선택 디렉터리 `scripts/`(실행), `references/`(필요 시 읽음), `assets/`(산출물용 파일), `evals/`(패키징 제외). 중첩 `SKILL.md`는 거부 → 여러 스킬이면 플러그인 [12].
- **점진적 공개 3단계**: L1 name+description 항상 로드(스킬당 약 100토큰) / L2 본문 트리거 시(**5,000토큰 미만**, 500줄 미만 — 둘 다 명세 권고 [16]) / L3 번들은 읽거나 실행할 때만, 스크립트 코드는 컨텍스트에 들어가지 않음 [1][12].
- **이식 가능(P0) 필드 6개**: `name`, `description`, `license`, `compatibility`, `metadata`, `allowed-tools`. 그 밖의 Claude Code 확장 필드(약 14개)는 claude.ai·Skills API 업로드에서 **하드 에러** [1][12].
  - `allowed-tools`는 명세가 "**Experimental. Support for this field may vary between agent implementations**"로 표시한다 [16]. → P0판에서는 쓰지 않고 Claude Code판에만 둔다.
  - `metadata`는 **문자열 키 → 문자열 값** 맵으로 한정된다 [16]. 숫자·리스트·중첩 객체 금지(`version: "0.1"`처럼 따옴표).
- **하드 제약** [12]
  - `name`: `^[a-z0-9-]+$`, 64자 이하, 앞뒤 하이픈·`--` 금지, `anthropic`/`claude` 포함 금지. **상위 디렉터리 이름과 일치해야 한다**(명세상 규칙 [16]; [12]는 권고로만 적음). VS Code는 불일치 시 로드 실패 [17].
  - `description`: 필수, 1,024자 이하, **`<` `>` 금지**(프롬프트 인젝션 방어). Claude Code 목록에서는 `when_to_use`와 합쳐 1,536자에서 잘림 [1].
  - `compatibility`: 500자 이하. 업로드 전체 30MB 미만(압축 해제 기준).
  - 참조는 SKILL.md에서 1단계까지, 100줄 넘는 참조 파일은 맨 위에 목차.
- **Claude Code 전용(P1)**: `${CLAUDE_SKILL_DIR}`·`$ARGUMENTS` 치환, `` !`cmd` `` 동적 주입(실패 시 호출 전체 중단 → `|| true`, 기본 2분 타임아웃), `allowed-tools`에 스크립트 경로 지정 [1]. claude.ai에서 동기화된 스킬은 `!` 명령이 실행되지 않고 글자 그대로 남아 조용히 깨진다 [12].
- **권한 수명**: `allowed-tools` 승인은 **다음 사용자 메시지에서 해제**된다 [1][12]. → 매 턴 `xl.py` 실행 권한 확인이 뜰 수 있으므로 설치 안내에 권한 규칙(예: `Bash(python */scripts/xl.py *)`) 추가 방법을 넣는다.
- **배포 위치**: 개인 `~/.claude/skills/`, 프로젝트 `.claude/skills/`, 플러그인. **Cowork와 클라우드 세션은 로컬 `~/.claude/skills/`를 읽지 않는다**(저장소 커밋·플러그인·claude.ai 업로드로 배포) [12].
- **실행 환경 차이** [12]: Claude Code는 네트워크·설치 가능(전역 설치는 지양), claude.ai는 설정에 따름, **API는 네트워크 없음·패키지 설치 불가**.
- **검증 도구**: [12]의 `validate_skill.py <dir> [--portable]`(오류 시 exit 1, `--portable`이면 CC 전용 필드도 오류)와 명세의 참조 구현 **`skills-ref validate ./web-excel`** [18]을 CI에서 함께 돌려 교차 검증한다.
- **스크립트 언어**: 지원 언어는 에이전트 구현에 달려 있고(보통 Python·Bash·JavaScript), 스크립트는 자기완결적이거나 의존성을 명확히 문서화해야 한다 [16]. → Python 단일 진입점 + 의존성 동봉 방침과 맞음.

### 2.1.1 에이전트 SW별 차이 [15]

| 제품 | 스킬 위치 | 확인 수준 | 이 skill에 미치는 영향 |
|---|---|---|---|
| Claude Code | `.claude/skills/`, `~/.claude/skills/`, 플러그인 [1] | 원문 | 주 대상. 데스크톱은 `.claude/launch.json`으로 서버 수명주기·미리보기를 앱이 관리 |
| Codex CLI | `$CODEX_HOME/skills/`(기본 `~/.codex/skills`), VS Code 문서 기준 `.agents/skills/`·`~/.agents/skills/` | 실측(codex-cli 0.162.1, macOS) | **`workspace-write` 샌드박스에서 127.0.0.1 바인딩조차 EPERM**, `network_access=true`로도 실패, `danger-full-access`에서만 성공(`codex sandbox` 경로 측정, 에이전트 exec 경로는 미검증). 스킬 메타데이터가 컨텍스트 예산을 넘으면 description을 자름 |
| VS Code / Copilot | 프로젝트 `.github/skills/`·`.claude/skills/`·`.agents/skills/`, 개인 `~/.copilot/skills/`·`~/.claude/skills/`·`~/.agents/skills/` [17] | 원문 | name 불일치·허용 외 문자 시 조용히 로드 실패. 스크립트 실행·승인 흐름은 문서에 없음(모름) |
| Cursor | `.agents/skills/`·`.cursor/skills/`(+홈), 호환으로 `.claude/skills/`·`.codex/skills/` [19] | 원문 요약 | 재귀 탐색 허용 |
| Gemini CLI | `.gemini/skills/`·`.agents/skills/`(+홈), SKILL.md는 루트나 1단계 깊이까지 [20] | 원문 요약 | **스킬 활성화 전에 사용자 확인** |

이식 권고 [15]: 원본을 `.agents/skills/web-excel/`에 두고 `.claude/skills/`·`.github/skills/`는 심볼릭 링크나 CI 복사로 맞춘다. description의 **앞쪽 100–200자에 트리거 문구**를 둔다(Codex 절단 대비).


### 2.2 미리보기 창
- 사용자 제공 전제: `localhost:{port}` 웹을 보여주는 뷰. → 서버가 UI·API·SSE를 **하나의 포트, 같은 출처**로 제공하면 된다.
- 웹뷰 특성상 주의: 팝업(`window.open`)이 외부 브라우저로 열릴 수 있음, 클립보드 권한 제한 가능, 일부 단축키를 호스트 앱이 가로챌 수 있음. → 다중 창 동기화를 `BroadcastChannel`(같은 브라우저 한정) 대신 **서버 경유**(SSE)로 바꾸면 미리보기 창과 외부 브라우저가 함께 동기화된다.
- Claude Code 데스크톱은 `.claude/launch.json`에 선언한 서버를 **앱이 띄우고 정리**하며 미리보기를 `http://localhost:<port>`로 연다 [15](analysis-docs 세션의 preview 도구 사용 경험). 세부:
  - 엔트리 예: `{"version":"0.0.1","configurations":[{"name":"web-excel","runtimeExecutable":"python3","runtimeArgs":["<skill>/scripts/xl.py","serve","--foreground","--port","8765"],"port":8765}]}`
  - 명령 없이 `url`만 있는 엔트리는 이미 떠 있는 서버에 붙는다(attach).
  - localhost `url`은 origin만 허용(경로·쿼리 불가)하고 `port`와 일치해야 한다. 특정 문서는 연 뒤 이동한다.

### 2.3 서버 수명주기: 두 가지 모드

| 모드 | 언제 | 누가 프로세스를 소유 | 방식 |
|---|---|---|---|
| **A. 호스트 관리 (권장)** | Claude Code 데스크톱 | 에이전트 앱 | `xl.py launch-config`가 빈 포트를 골라 `.claude/launch.json` 엔트리를 만들거나 갱신 → 미리보기 시작. 서버는 `serve --foreground`로 돌고 **부모 프로세스가 사라지면 종료** |
| **B. 자체 분리 실행** | CLI·타사 에이전트 | 스크립트 | `serve`가 분리 프로세스로 재기동(POSIX `start_new_session=True`, Windows `DETACHED_PROCESS`·`CREATE_NEW_PROCESS_GROUP`·`CREATE_NO_WINDOW` [9]) → 출력한 URL을 사용자가 미리보기에서 엶 |

공통 자기방어(세션 종료 시 호스트가 프로세스를 정리해 준다는 보장은 확인되지 않았다 [15]):
- **유휴 종료**: 연결된 클라이언트 0개이고 마지막 요청 후 N분(기본 60분)이면 종료. 단 dirty 문서가 있으면 복구 스냅샷을 먼저 기록.
- **단일 인스턴스와 포트 공고**: 작업 경로별 `.web-excel/server.json`에 `{pid, port, token, started_at, root}`를 임시 파일 + 원자적 rename으로 기록. 재진입 시 health 확인 후 재사용, 죽었으면 정리. 고정 포트만 쓰면 동시 실행·잔존 프로세스에서 깨지므로 선호 포트(8765)가 사용 중이면 다음 빈 포트를 쓴다(모드 A는 고른 포트를 launch.json에 기록).
- **바인딩은 127.0.0.1만**(0.0.0.0 금지, LAN 노출 방지). API는 토큰 필수.

검토한 대안: 서버를 **stdio MCP 서버로 감싸면** 호스트가 프로세스를 소유해 세션과 함께 정리되고 Codex 샌드박스 문제도 MCP 기동 정책으로 옮겨진다 [15]. 다만 MCP 클라이언트 설정이 필요하고 미리보기는 어차피 localhost HTTP가 필요하므로, v1은 "skill + CLI"로 하고 MCP 래퍼는 P7(선택)로 둔다.

---

## 3. 필요 기술과 수준

난이도: 상(전문 지식·설계 판단 필요) / 중(경험 있는 개발자) / 하(일반 구현).  상태: 완료 = 현 브랜치에 있음.

| 영역 | 필요 기술 | 난이도 | 상태 | 남은 작업 |
|---|---|---|---|---|
| 파일 형식 | OOXML(SpreadsheetML·DrawingML 차트), BIFF8(읽기), openpyxl/xlrd 내부 구조 [6][7] | 상 | 완료 | 연산을 파일 모델에 반영하는 경로 재사용 |
| DRM 대응 | 경로 기반 I/O, 프로세스 기반 투명 복호화 이해, 훅 설계 | 중 | 완료(`fileio.py`) | 분리 실행된 python 프로세스도 DRM 허용 대상인지 **현장 검증** |
| 수식 엔진 | 파서/평가기, Excel 의미론(형변환·오류 전파·1900 날짜), 의존·재계산 [8] | 상 | 완료(JS, 함수 약 200개) | 서버 측 계산 경로(§5.6) |
| 참조 재작성 | 행/열 삽입·삭제·정렬 시 수식 참조 조정 | 중~상 | 완료(JS) | **Python 포팅 + JS와 동등성 테스트** |
| 그리드 UI | Canvas 가상화, IME, 히트 테스트, 단축키 | 상 | 완료 | 원격 변경 하이라이트, 상대방 선택 표시 |
| 동시 편집 | 서버 순서화 + LWW, 구조 연산 충돌 처리(OT 개념) [3][4][5] | **상** | 부분(창 간 BroadcastChannel) | **서버 권한 세션·op 로그·SSE·재동기화** |
| 실시간 전송 | SSE(서버→클라), HTTP POST(클라→서버) [10] | 중 | 없음 | stdlib 서버에서 SSE 스트리밍 |
| 프로세스 관리 | 분리 실행, PID/포트/락 파일, 헬스 체크, 유휴 종료, Windows 차이 [9] | 중 | 없음 | `xl serve/stop/status` |
| 패키징 | Skill 구조, Python 의존성 설치 안내(`scripts/requirements.txt`: openpyxl 필수, xlrd 선택) | 하~중 | 없음 | 없으면 `xl.py serve`가 설치 명령 안내(종료 코드 6). 처음 계획은 `_vendor` 동봉이었으나 라이선스 고지 부담 때문에 변경 |
| 에이전트 인터페이스 | 토큰 효율적 CLI 출력, 오류 코드, 멱등성, 일괄 처리 | 중 | 없음 | §5.2 CLI |
| 보안 | 로컬 서버 CSRF/DNS 리바인딩, 토큰, 경로 샌드박스 [11] | 중 | 부분(Host 검사·커스텀 헤더) | 세션 토큰·작업 경로 기본 샌드박스 |
| 테스트 | Node/Playwright E2E, 동등성 테스트, Skill 트리거 평가 | 중 | 부분 | 동시 편집 시나리오, 교차 OS |

**필요 인력 수준**: 동시 편집과 파일 형식을 다룰 수 있는 시니어 1명이면 충분. Skill 작성 자체는 문서 작성 수준의 난이도다.

---

## 4. 목표 아키텍처

```
 ┌────────────── 사용자 PC (작업 경로 = 프로젝트 루트) ───────────────┐
 │                                                                     │
 │  에이전트 SW                                                         │
 │   ├─ 에이전트 ──(셸)──> python xl.py <cmd>  ──HTTP + token──┐        │
 │   └─ 미리보기 창 ── http://localhost:{port} ──HTTP/SSE──┐   │        │
 │                                                         ▼   ▼        │
 │                     ┌──── web-excel 서버 (분리 실행, Python) ────┐   │
 │                     │ 정적 UI / REST API / SSE /events           │   │
 │                     │ DocSession: 스냅샷 + version + op 로그     │   │
 │                     │   ├─ apply(op) → 순서 부여 → 방송          │   │
 │                     │   ├─ 계산: Node 워커(engine.js) | 브라우저 │   │
 │                     │   └─ 저장/열기: fileio.read/write_bytes ───┼──> xlsx/xls
 │                     └────────────────────────────────────────────┘   │
 │                     .web-excel/ (state.json, logs, backups, recovery)│
 └─────────────────────────────────────────────────────────────────────┘
```

핵심 원칙
1. **서버가 진실 원천**: 열린 문서는 서버 메모리의 `DocSession`이 소유. 브라우저·CLI는 모두 클라이언트.
2. **모든 변경은 op**: 브라우저 편집도, 에이전트 명령도 같은 op 스키마로 서버에 보낸다. 서버가 `version`을 1씩 증가시키며 순서를 정한다 [4].
3. **파일 접근은 Python만**: 기존 DRM 원칙 유지. 저장은 명시적(`save`)이 기본.
4. **행위자(actor) 기록**: 모든 op에 `actor: user|agent`와 `batchId`를 남겨, 서로의 변경을 보고 되돌릴 수 있게 한다.

---

## 5. 명세 (Spec)

### 5.1 Skill 패키지 구조

```
web-excel/
├─ SKILL.md                     # 500줄·5,000토큰 미만: 언제 쓰는지, 시작/종료, 핵심 명령, 안전 규칙
├─ SKILL.claude-code.md         # Claude Code용 판 (allowed-tools, 상태 주입, launch.json 안내)
├─ scripts/
│  ├─ xl.py                     # 단일 진입점 CLI (stdlib만 사용)
│  ├─ server/                   # 현재 server/ 패키지 + session.py, ops.py, refs.py(참조 재작성 포팅)
│  ├─ calc/worker.mjs           # Node가 있을 때 engine.js로 재계산하는 워커
│  └─ requirements.txt          # openpyxl(필수), xlrd(.xls용, 선택) — 동봉하지 않고 pip으로 설치
├─ assets/web/                  # 현재 static/ (index.html, js/, css/, vendor/chart.umd.min.js)
├─ references/
│  ├─ cli.md                    # 전체 명령·옵션·출력 형식
│  ├─ ops.md                    # op 스키마
│  ├─ formulas.md               # 지원 함수·제한
│  └─ troubleshooting.md        # 포트 충돌, DRM, Windows, 프록시
└─ LICENSE, THIRD_PARTY_NOTICES
```

`SKILL.md`는 [12]의 템플릿처럼 **두 판**으로 낸다.

- `SKILL.md`(P0, 기본): 명세 필드만 사용하되 실험적 필드인 `allowed-tools`는 뺀다. 스크립트 경로는 "이 스킬 폴더의 `scripts/xl.py`"로 서술. 본문 5,000토큰 미만.
- `SKILL.claude-code.md`(P1): `${CLAUDE_SKILL_DIR}`를 쓰는 `allowed-tools`와 `` !`python ${CLAUDE_SKILL_DIR}/scripts/xl.py status --brief || true` `` 상태 주입을 추가. Claude Code 설치 시 이 파일을 `SKILL.md`로 배치.

P0 프론트매터 초안(검증기 `--portable` 통과 기준: `<>` 없음, 1,024자 이하, 3인칭, 배제 조건 포함):

```yaml
---
name: web-excel
description: >
  Opens, inspects, edits, formats, charts, filters and saves Excel workbooks (.xlsx, .xlsm, .xls)
  in the workspace through a local web-excel server, while the user watches and edits the same
  workbook live in the agent's preview pane at a localhost URL. Use when the user asks to read or
  change spreadsheet data, build formulas or charts, or wants to see or edit an Excel file
  together with the agent. Do not use for CSV-only data processing, Google Sheets, or one-off
  conversions where no live view is wanted; use a plain spreadsheet library for those.
license: MIT
compatibility: >
  Requires Python 3.9+ and a local agent session on the user's machine with a shell; binds a
  port on 127.0.0.1 (sandboxes that block loopback binding, e.g. Codex workspace-write, need
  approval or full access). Optional Node.js 18+ for server-side recalculation. Not usable from
  the Skills API or cloud sessions, which cannot serve the user's localhost.
metadata:
  version: "0.1"
---
```

본문 핵심 지시(요지):
1. 시작: Claude Code 데스크톱이면 `xl.py launch-config` 후 미리보기 시작(모드 A). 그 밖에는 `xl.py serve`(모드 B) 후 출력된 URL을 사용자에게 알리고 미리보기 창에서 열도록 안내.
2. 탐색은 `describe` → 필요한 범위만 `read`(대용량 전체 덤프 금지).
3. 변경은 가능한 한 `apply`(일괄 op)로 한 번에. 변경 후 `changes`로 결과 확인.
4. 사용자가 미리보기에서 편집 중일 수 있으므로 쓰기 전 `changes --since`로 최신 상태 확인.
5. **저장은 사용자가 요청했거나 명시적으로 동의한 경우에만** `save`. 덮어쓰기 전 자동 백업.
6. 작업 종료 시 서버는 유지(사용자가 계속 볼 수 있게). 사용자가 원하면 `stop`.

### 5.2 CLI (`xl.py`)

공통: 상태 파일 `.web-excel/state.json`(port, pid, token, root)을 찾아 서버에 연결. 출력 기본은 **간결한 텍스트/TSV**, `--json`으로 기계 판독 형식. 종료 코드: 0 성공, 2 사용법, 3 서버 없음, 4 충돌(409), 5 파일/DRM 오류.

| 명령 | 설명 |
|---|---|
| `serve [--port 8765] [--root DIR] [--no-open-browser] [--foreground]` | 멱등 기동. 이미 있으면 상태만 출력. 포트가 사용 중이면 다음 포트 탐색. `--foreground`는 launch.json용 |
| `launch-config [--port P]` | 빈 포트를 골라 `.claude/launch.json`에 web-excel 엔트리를 추가·갱신(기존 엔트리 보존) |
| `status` / `stop` | 서버 상태(포트, 열린 문서, 연결 클라이언트 수, dirty) / 종료(dirty 문서가 있으면 거부, `--force` 시 복구 파일 기록) |
| `ls [DIR]` | 작업 경로의 엑셀 파일 목록 |
| `open PATH` | 서버에서 문서 세션 열기 → `docId` 출력. 미리보기 창에 탭으로 나타남 |
| `describe DOC` | 시트별 사용 범위, 머리글 추정, 열 유형, 수식/차트/필터/병합 개수 (토큰 절약용) |
| `read DOC RANGE [--what values\|formulas\|styles\|all] [--max-rows 200]` | `시트!A1:D50` 형식. 초과 시 잘림과 다음 페이지 안내 |
| `write DOC RANGE --tsv FILE\|- \| --json VALUES` | 값/수식 입력(`=`로 시작하면 수식) |
| `apply DOC OPS.json` | 일괄 op(§5.4). 원자적 적용, 하나의 `batchId` |
| `format DOC RANGE --set k=v ...` | 예: `--set b=true nf="#,##0" bg=#FFF2CC` |
| `sheet DOC add\|rm\|rename\|move ...` | 시트 관리 |
| `rows/cols DOC insert\|delete\|width\|height\|hide ...` | 구조 변경 |
| `sort DOC RANGE --by C:desc [--header]` / `filter DOC RANGE --col B --gt 150` | 정렬·필터 |
| `chart DOC add --type bar --data 시트!A1:C7 [--title T] [--at H2]` / `chart DOC rm ID` | 차트 |
| `changes DOC [--since VERSION] [--actor user\|agent]` | 변경 요약(누가, 어떤 범위). 사용자 편집 감지에 사용 |
| `selection DOC` | 사용자가 미리보기에서 선택 중인 시트/범위 ("이 부분 고쳐줘" 지원) |
| `undo DOC --batch ID` | 특정 일괄 변경 되돌리기(이후 같은 셀이 바뀌었으면 경고 후 `--force` 필요) |
| `save DOC [--as PATH] [--format xlsx\|xls]` | Python으로 저장. 기본 덮어쓰기 전 `.web-excel/backups/`에 사본 |
| `close DOC` | 세션 닫기(dirty면 거부) |

### 5.3 HTTP API

모든 `/api/*` 요청: `Host` 검사, `X-WebExcel-Token: <token>` 필수(CLI는 state.json, 브라우저는 `/` 로드 시 발급되는 `HttpOnly; SameSite=Strict` 쿠키로 대체) [11].

| 메서드 · 경로 | 요청 | 응답 |
|---|---|---|
| `GET /api/health` | – | `{ok, version, pid, root}` |
| `GET /api/fs?path=` | – | 기존과 동일(작업 경로 샌드박스 기본) |
| `POST /api/docs` | `{path}` | `{docId, version, snapshot}` (이미 열려 있으면 기존 세션) |
| `GET /api/docs` | – | 열린 세션 목록 `{docId, path, version, dirty, clients}` |
| `GET /api/docs/{id}/snapshot` | – | `{version, workbook(JSON 모델), calc: {source, version}}` |
| `GET /api/docs/{id}/describe` | – | §5.2 describe 데이터 |
| `GET /api/docs/{id}/range?ref=&what=` | – | `{ref, values[][], formulas[][]?, styles?, calcVersion}` |
| `POST /api/docs/{id}/ops` | `{baseVersion, actor, clientId, batchId, label, ops[]}` | `200 {version}` / `409 {currentVersion, reason}` |
| `GET /api/docs/{id}/events?since=N` | SSE | `event: ops` `{version, actor, batchId, ops}` / `event: calc` `{version, values patch}` / `event: presence` / `event: file` (외부 변경) / `event: saved` |
| `POST /api/docs/{id}/presence` | `{clientId, sheet, sel}` | – |
| `GET /api/docs/{id}/changes?since=` | – | 범위 단위 변경 요약 |
| `POST /api/docs/{id}/undo` | `{batchId, force}` | `{version}` / `409` |
| `POST /api/docs/{id}/save` | `{path?, format?, overwrite, backup}` | `{path, warnings}` |
| `DELETE /api/docs/{id}` | `{force}` | – |

### 5.4 op 스키마

두 종류로 나눈다.

**상태 op (멱등, 셀/속성 단위 LWW)**: 서버 도착 순서상 마지막 값이 이긴다(Figma 방식, 서버가 순서를 정하므로 타임스탬프 불필요) [4].
```jsonc
{ "op": "setCells", "sheet": "s1", "cells": [[r, c, {"v": 1, "f": "=A1*2", "s": {...}} | null]] }
{ "op": "patchStyle", "sheet": "s1", "range": "B2:D9", "set": {"b": true, "nf": "#,##0"}, "unset": ["bg"] }
{ "op": "setSheetProps", "sheet": "s1", "props": {"freeze": "B2", "showGrid": false, "tabColor": "#FF0000"} }
{ "op": "setDims", "sheet": "s1", "axis": "col", "items": [[3, {"w": 120, "hidden": false}]] }
{ "op": "setMerges", "sheet": "s1", "add": ["A1:C1"], "remove": [] }
{ "op": "setFilter", "sheet": "s1", "filter": {"ref": "A1:F7", "cols": {"1": {"custom": [...]}}} | null }
{ "op": "chartUpsert", "sheet": "s1", "chart": {...} }  { "op": "chartDelete", "sheet": "s1", "id": "ch1" }
```
**구조 op (비가환)**: 좌표계를 바꾼다.
```jsonc
{ "op": "insertRows"|"deleteRows"|"insertCols"|"deleteCols", "sheet": "s1", "at": 4, "count": 2 }
{ "op": "sort", "sheet": "s1", "range": "A1:F7", "keys": [{"c": 2, "asc": false}], "header": true }
{ "op": "addSheet"|"deleteSheet"|"renameSheet"|"moveSheet", ... }
```
- 시트는 이름이 아닌 안정 ID(`s1`)로 지칭(CLI는 이름을 ID로 변환).
- 서버는 구조 op 적용 시 **모든 수식·차트 참조·병합·필터 범위**를 재작성(현재 JS `adjustFormula`의 Python 포팅).
- 브라우저는 기존 트랜잭션 엔트리(`cells/meta/structure`)를 위 op로 직렬화하여 전송하도록 `sync.js`를 교체.

### 5.5 동기화 프로토콜

1. 클라이언트는 `snapshot`(version V)을 받고 `events?since=V`로 SSE 구독.
2. 로컬 편집은 **낙관적 적용** 후 `pending` 큐에 넣고 `POST ops {baseVersion}`.
3. 서버 규칙
   - 상태 op만 포함: `baseVersion`과 무관하게 수락, 순서 부여, 방송(LWW).
   - 구조 op 포함 또는 `baseVersion` 이후 **같은 시트에 구조 op가 있었던** 경우: v1에서는 `409` 반환 → 클라이언트는 스냅샷 재수신 후 pending op를 **좌표 변환(행/열 인덱스 이동)해 재전송**하거나, 변환 불가 시 버리고 사용자에게 알림. 이는 Ellis & Gibbs의 OT에서 1차원 인덱스 변환만 취한 단순화다 [3].
4. 자신이 보낸 op의 방송(echo)은 `clientId`로 식별해 pending에서 제거(재적용하지 않음).
5. 재연결: `since=V`로 누락분 재전송, op 로그 범위를 벗어나면 스냅샷 재수신.
6. 표시: 상대 행위자 변경은 셀을 잠시 강조하고 상태 표시줄에 "에이전트가 B2:D9를 수정함 (되돌리기)" 표시. 상대방 선택 영역은 다른 색 테두리로 표시(presence).
7. 실행 취소: 브라우저 Ctrl+Z는 **자기 행위자의 batch**만 되돌린다(서버 `undo` 사용). 상대가 이후 같은 셀을 바꿨다면 확인 요청.

### 5.6 수식 계산 전략

엔진은 JS(`engine.js`, `model.js`)이며 DOM 의존이 없다(이미 Node로 테스트됨). 서버에서 값을 제공하는 단계:

| 순위 | 계산 원천 | 조건 | 비고 |
|---|---|---|---|
| 1 | **Node 워커**(`calc/worker.mjs`가 동일 엔진 실행) | `node` 18+ 존재 | 결과가 브라우저와 동일. 서버가 op 적용 후 증분 전달, `calc` 이벤트 방송 |
| 2 | **연결된 브라우저** | 미리보기 창이 열려 있음 | 브라우저가 계산 후 `calc` 패치를 서버로 보고(리더 1개 선출) |
| 3 | 파일의 **캐시 값** | 위 둘 없음 | `read` 응답에 `calc: "stale"` 표시, SKILL.md에서 사용자에게 미리보기를 열거나 Node 설치를 권하도록 지시 |

검토 후 제외: Python 수식 라이브러리(`formulas`, `pycel`)는 무거운 의존성·라이선스(EUPL/GPL)·엔진 이중화로 결과 불일치 위험.

### 5.7 파일·DRM·안전

- 모든 읽기/쓰기는 `fileio.read_bytes/write_bytes`(훅 교체 가능) — 분리 실행된 서버 프로세스도 DRM 허용 목록에 포함되는지 **현장 검증 필수**.
- 저장 전 백업 `.web-excel/backups/<name>.<timestamp>.xlsx`(DRM 정책상 사본 생성이 금지된 환경을 위해 `--no-backup` 설정 제공).
- 외부 변경 감지: 열 때 `(mtime, size, sha256)` 기록 → 저장 전/주기적으로 비교, 다르면 `event: file`로 알리고 덮어쓰기 거부(사용자 확인 필요).
- Office 잠금 파일(`~$name.xlsx`) 존재 시 경고(Excel에서 열려 있음).
- 비정상 종료 대비: dirty 세션은 주기적으로 `.web-excel/recovery/`에 JSON 스냅샷.
- 기본 샌드박스: `root` = 작업 경로(`${CLAUDE_PROJECT_DIR}` 또는 cwd). 밖의 경로는 `--allow-path`로 명시.

### 5.8 보안

| 위협 | 대책 |
|---|---|
| 다른 웹사이트의 요청(CSRF) | 토큰 헤더 필수, 쿠키는 `SameSite=Strict` [11], CORS 미허용 |
| DNS 리바인딩 | `Host` 헤더 허용 목록(현재 구현 유지) |
| 같은 PC의 다른 사용자 | 127.0.0.1 바인딩, `state.json` 권한 0600 |
| 경로 탈출 | 정규화 후 root 비교(현재 구현 유지) |
| 에이전트의 의도치 않은 덮어쓰기 | 명시적 save, 자동 백업, batch 단위 undo, 감사 로그 `.web-excel/audit.log` |

### 5.9 실행 표면(surface)별 지원 범위

| 표면 | 스킬 로드 | 서버 실행·미리보기 | 판정 |
|---|---|---|---|
| Claude Code (사용자 PC: CLI/데스크톱) | 개인·프로젝트·플러그인 | 가능 (loopback, 셸) | **주 대상** |
| Cowork / Claude Code 클라우드 세션 | 저장소 `.claude/skills/`·동기화 스킬만 [12] | 서버는 원격 컨테이너에서 돌아 사용자 PC 파일·미리보기와 분리 | 미리보기 요구 충족 불가 |
| claude.ai | 업로드 스킬(P0 필드만) | 장시간 서버·사용자 localhost 불가 | 대상 외 |
| Skills API | 업로드 스킬 | 네트워크 없음 [12] | 대상 외 |
| Codex CLI | `~/.codex/skills`, `.agents/skills/` | `workspace-write`에서 loopback 바인딩 차단(실측 [15]) | **조건부**: 승인 또는 full access 필요 |
| VS Code / Copilot, Cursor | `.github`·`.agents`·`.cursor`·`.claude` 경로 [17][19] | 셸+loopback이면 가능, 실행 승인 흐름 미확인 | **미검증**, P0판으로 시험 |
| Gemini CLI | `.gemini/skills/`·`.agents/skills/` [20] | 활성화 전 사용자 확인 | **미검증** |

### 5.10 비기능 요구사항(목표치)

| 항목 | 목표 |
|---|---|
| 로컬 op 왕복(브라우저→서버→다른 클라이언트) | p95 < 150 ms |
| `serve` 기동 | < 2 s (Node 워커는 지연 기동) |
| 문서 크기 | 셀 100만 개까지 열기·편집(Python 메모리 약 0.5–1 GB 예상 → 초과 시 경고) |
| 지원 OS | Windows 10+, macOS 13+, Linux / Python 3.9+ |
| 외부 의존 | 런타임 `pip` 설치 불필요(동봉), Node 선택 |

---

## 6. 구현 계획

| 단계 | 내용 | 산출물 / 완료 기준 | 공수(인일) |
|---|---|---|---|
| P0 패키징·수명주기 | Skill 폴더 구조(원본 `.agents/skills/web-excel/` + 링크), 의존성 동봉, `xl.py serve/stop/status/launch-config`, 모드 A·B, server.json 원자 기록, 포트 탐색, 유휴 종료·부모 종료 감지 | Python만 있는 새 PC에서 미리보기에 UI 표시(모드 A·B 각각). 재실행 멱등. Windows/macOS 검증. Codex에서 바인딩 가능 조건 재현 | 2–4 |
| P1 서버 권한 세션 | `DocSession`, 스냅샷 API, 상태 op 적용(Python), op 로그, SSE, 브라우저 `sync.js` 교체(낙관적 적용 + echo 처리) | 브라우저 2개 + CLI 동시 셀 편집이 수렴. 재연결 시 누락 없음 | 5–7 |
| P2 에이전트 CLI | `open/describe/read/write/apply/format/sheet/chart/changes/selection/save` + 오류 코드 + TSV/JSON | 스크립트 시나리오(열기→요약→수정→차트→저장) 무인 통과 | 4–6 |
| P3 구조 op | 참조 재작성 Python 포팅, insert/delete/sort/rename, 409 + 재동기화·좌표 변환 | JS와 Python 재작성 결과 **동등성 테스트**(무작위 수식 1만 건) 일치 | 5–8 |
| P4 계산 단계화 | Node 워커, 브라우저 계산 보고(리더 선출), stale 표시 | 세 계산 원천의 결과가 테스트 코퍼스에서 일치 | 3–5 |
| P5 안전·보안 | 토큰/쿠키, 백업, 외부 변경 감지, 복구, 감사 로그, batch undo | 보안 체크리스트 통과, 충돌 시 데이터 손실 0 | 3–4 |
| P6 문서·평가 | SKILL.md(P0/P1 두 판, 본문 5,000토큰 미만), references/(1단계·목차), `validate_skill.py --portable` + `skills-ref validate` CI, `evals/evals.json`(시나리오 ≥3, 새 세션 스킬 유/무 A/B), 트리거 쿼리 20개(해당 8–10·키워드가 겹치는 비해당 8–10, 각 3회, 60/40 분할 후 test 점수로 선택), Haiku·Sonnet·Opus 각각 [12], 동시 편집 E2E | 검증기 오류 0, 트리거 정확도(test) ≥ 90%, E2E 통과 | 2–3 |
| P7 (선택) MCP 래퍼 | stdio MCP 서버가 같은 서버 프로세스를 소유·노출 | MCP 클라이언트에서 같은 시나리오 통과 | 2–3 |
| **합계(P0–P6)** | | | **24–37** |

우선순위: P0 → P1 → P2만으로도 "에이전트가 조작하고 사용자가 실시간으로 보는" 최소 기능(MVP, 약 11–16 인일)이 된다. 구조 op 충돌(P3)은 MVP에서 "구조 op는 서버 단일 순서로만 허용, 충돌 시 재동기화"로 단순 처리 가능.

---

## 7. 위험과 대응

| 위험 | 영향 | 대응 |
|---|---|---|
| DRM 에이전트가 분리 실행 python 또는 백업 사본을 다르게 취급 | 열기/저장 실패, 정책 위반 | P0에서 현장 PC 검증, 훅·`--no-backup` 제공 |
| 에이전트 SW가 세션 종료 시 프로세스 그룹 정리(또는 정리하지 않아 잔존) | 서버 중단 / 좀비 서버 | 모드 A는 앱이 소유, 모드 B는 유휴 종료·server.json 재사용, 복구 스냅샷 |
| 에이전트 샌드박스가 loopback 바인딩 차단(Codex 실측) | 서버 기동 불가 | compatibility에 명시, 기동 실패 시 원인과 승인 방법을 출력 |
| 제품별 스킬 경로·description 절단 차이 | 스킬이 안 보이거나 트리거 실패 | `.agents/skills/` 원본 + 링크, 트리거 문구를 앞 100–200자에 |
| 미리보기 웹뷰 제약(팝업·클립보드·단축키) | 일부 기능 저하 | 서버 경유 다중 창, 클립보드는 Ctrl+C/V 이벤트 사용(이미 그렇게 구현), 단축키 대체 키 |
| Node 부재 | 브라우저 미연결 시 값 계산 불가 | stale 표시 + 안내, 필요 시 P4에서 브라우저 헤드리스 대안 검토 |
| 동시 편집 버그 | 데이터 불일치 | 서버 단일 순서·LWW로 범위 축소, 동등성/퍼즈 테스트 |
| 대용량 파일 | 메모리·지연 | 크기 상한 경고, read 페이지네이션, 증분 계산 |
| 포트 충돌·다중 프로젝트 | 다른 서버에 연결 | state.json을 작업 경로별로, health 응답의 root 확인 |

---

## 8. 참고 자료

1. Anthropic, *Extend Claude with skills* (Claude Code 문서: 위치, 프론트매터, `${CLAUDE_SKILL_DIR}`, 동적 주입, 한도). https://code.claude.com/docs/en/skills
2. B. Zhang, K. Lazuka, M. Murag, *Equipping agents for the real world with Agent Skills*, Anthropic Engineering, 2025. https://www.anthropic.com/engineering/equipping-agents-for-the-real-world-with-agent-skills
3. C. A. Ellis, S. J. Gibbs, "Concurrency Control in Groupware Systems," *Proc. ACM SIGMOD*, pp. 399–407, 1989. https://doi.org/10.1145/67544.66963
4. E. Wallace, *How Figma's multiplayer technology works*, Figma Blog, 2019 (서버 순서화 + 속성 단위 LWW). https://www.figma.com/blog/how-figmas-multiplayer-technology-works/
5. M. Shapiro, N. Preguiça, C. Baquero, M. Zawirski, "Conflict-free Replicated Data Types," *SSS 2011*, LNCS 6976, pp. 386–400. https://doi.org/10.1007/978-3-642-24550-3_29
6. Ecma International, *ECMA-376 Office Open XML File Formats*. https://ecma-international.org/publications-and-standards/standards/ecma-376/
7. Microsoft, *[MS-XLSX] §2.2.3 Functions* (`_xlfn.` 미래 함수). https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/5d1b6d44-6fc1-4ecd-8fef-0b27406cc2bf
8. P. Sestoft, *A Spreadsheet Core Implementation in C#*, ITU Technical Report TR-2006-91, 2006 (재계산·의존 그래프). https://www.itu.dk/~sestoft/funcalc/ITU-TR-2006-91.pdf
9. Python Software Foundation, *subprocess — Subprocess management* (`start_new_session`, `DETACHED_PROCESS`). https://docs.python.org/3/library/subprocess.html
10. WHATWG, *HTML Living Standard — Server-sent events*. https://html.spec.whatwg.org/multipage/server-sent-events.html
11. OWASP, *Cross-Site Request Forgery Prevention Cheat Sheet*. https://cheatsheetseries.owasp.org/cheatsheets/Cross-Site_Request_Forgery_Prevention_Cheat_Sheet.html

12. JISUlicious, *Agent Skills 작성 표준* (analysis-docs, 2026-09-04; 근거: Claude Code skills 문서, platform.claude.com의 best-practices·overview·skills-guide, `anthropics/skills`의 skill-creator, 배포 스킬 40개 실측) 및 `validate_skill.py`, `templates/`. https://github.com/JISUlicious/analysis-docs/tree/main/agent-skills
13. Anthropic, *Skill authoring best practices*. https://platform.claude.com/docs/en/agents-and-tools/agent-skills/best-practices
14. Anthropic, *Agent Skills overview*. https://platform.claude.com/docs/en/agents-and-tools/agent-skills/overview

15. analysis-docs 세션(session_018uoA4GNoSMJodFx4mAzhVm)의 회신, 2026-11-04. 등급 표기: 원문 확인 / 그 PC에서 실측 / 2차 / 모름. Codex 항목은 그 세션의 실측.
16. Agent Skills, *Specification*. https://agentskills.io/specification (analysis-docs 세션이 원문 확인)
17. Microsoft, *Agent skills in VS Code*. https://code.visualstudio.com/docs/agent-customization/agent-skills
18. agentskills, *skills-ref* (참조 검증 도구). https://github.com/agentskills/agentskills/tree/main/skills-ref
19. Cursor, *Skills*. https://cursor.com/docs/context/skills
20. Google, *Gemini CLI: Creating skills*. https://geminicli.com/docs/cli/creating-skills

검증 메모: [1]과 [12]는 이 세션에서 원문을 열람해 확인했고([13][14]는 [12]가 인용한 1차 문서), [16]–[20]은 analysis-docs 세션이 확인한 것을 전달받았다(이 세션에서 직접 열람하지 않음), [2][3][4][8]은 검색으로 서지 정보를 확인했다. agentskills.io 명세는 네트워크 정책으로 열람하지 못했다. [5][6][9][10][11]은 널리 알려진 표준·문서로, 링크는 이 환경에서 직접 열어 보지 못했다.
