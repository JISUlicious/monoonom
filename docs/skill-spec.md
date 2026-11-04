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
- Skill = 폴더 + `SKILL.md`(YAML 프론트매터 + 지시문) + 선택적 지원 파일(scripts 등). 에이전트는 먼저 이름·설명만 보고, 필요할 때 본문을 읽고, 그다음 필요한 파일만 읽거나 실행한다(**점진적 공개**) [1][2].
- Claude Code 기준 위치: 개인 `~/.claude/skills/<name>/SKILL.md`, 프로젝트 `.claude/skills/<name>/SKILL.md`, 플러그인 `<plugin>/skills/<name>/` [1].
- 번들 스크립트는 **컨텍스트에 로드되지 않고 실행**된다. 문서화된 패턴: 본문과 `allowed-tools`에 `${CLAUDE_SKILL_DIR}/scripts/...`를 사용 [1].
- 치환 변수: `${CLAUDE_SKILL_DIR}`, `${CLAUDE_PROJECT_DIR}`, `$ARGUMENTS` 등 [1].
- `` !`command` `` 동적 주입: Skill 호출 전에 명령을 실행해 결과를 본문에 넣는다. 기본 2분 타임아웃, 실패 시 호출 중단 [1]. → **서버 상태 확인(`xl status`)을 주입**하는 데 적합. 서버 기동 자체는 주입보다 명시적 단계가 안전.
- 이식성: Claude Code 밖(claude.ai, Skills API)에서는 `name`, `description`, `license`, `compatibility`, `metadata`, `allowed-tools`만 허용되고 다른 키는 오류 [1]. → **프론트매터는 이 6개만 사용**한다.
- 한도: `SKILL.md` 500줄 이하 권장, description(+when_to_use)은 목록에서 1,536자로 잘림 [1].
- 오픈 표준 명세(agentskills.io)는 이 세션의 네트워크 정책상 열람하지 못했다. 다른 에이전트 SW의 지원 범위는 각 제품 문서로 재확인 필요.

### 2.2 미리보기 창
- 사용자 제공 전제: `localhost:{port}` 웹을 보여주는 뷰. → 서버가 UI·API·SSE를 **하나의 포트, 같은 출처**로 제공하면 된다.
- 웹뷰 특성상 주의: 팝업(`window.open`)이 외부 브라우저로 열릴 수 있음, 클립보드 권한 제한 가능, 일부 단축키를 호스트 앱이 가로챌 수 있음. → 다중 창 동기화를 `BroadcastChannel`(같은 브라우저 한정) 대신 **서버 경유(SSE)**로 바꾸면 미리보기 창과 외부 브라우저가 함께 동기화된다.
- Claude Code 데스크톱은 `.claude/launch.json`으로 개발 서버를 띄우고 미리보기한다(제3자 자료 기준, 공식 문서 재확인 필요). 지원하되 필수로 두지 않는다.

### 2.3 "에이전트가 서버를 실행" 시 실제 동작
- 에이전트 셸 도구로 실행한 프로세스는 셸/세션 종료 시 함께 종료될 수 있다. → 스크립트가 스스로 **분리 프로세스**로 재기동해야 한다:
  POSIX `subprocess.Popen(..., start_new_session=True)`, Windows `creationflags=DETACHED_PROCESS | CREATE_NEW_PROCESS_GROUP | CREATE_NO_WINDOW` [9].
- 같은 작업 경로에 서버가 이미 떠 있으면 재사용(멱등 `serve`).

---

## 3. 필요 기술과 수준

난이도: 상(전문 지식·설계 판단 필요) / 중(경험 있는 개발자) / 하(일반 구현).  상태: 완료 = 현 브랜치에 있음.

| 영역 | 필요 기술 | 난이도 | 상태 | 남은 작업 |
|---|---|---|---|---|
| 파일 형식 | OOXML(SpreadsheetML·DrawingML 차트), BIFF8, openpyxl/xlrd/xlwt 내부 구조 [6][7] | 상 | 완료 | 연산을 파일 모델에 반영하는 경로 재사용 |
| DRM 대응 | 경로 기반 I/O, 프로세스 기반 투명 복호화 이해, 훅 설계 | 중 | 완료(`fileio.py`) | 분리 실행된 python 프로세스도 DRM 허용 대상인지 **현장 검증** |
| 수식 엔진 | 파서/평가기, Excel 의미론(형변환·오류 전파·1900 날짜), 의존·재계산 [8] | 상 | 완료(JS, 함수 약 200개) | 서버 측 계산 경로(§5.6) |
| 참조 재작성 | 행/열 삽입·삭제·정렬 시 수식 참조 조정 | 중~상 | 완료(JS) | **Python 포팅 + JS와 동등성 테스트** |
| 그리드 UI | Canvas 가상화, IME, 히트 테스트, 단축키 | 상 | 완료 | 원격 변경 하이라이트, 상대방 선택 표시 |
| 동시 편집 | 서버 순서화 + LWW, 구조 연산 충돌 처리(OT 개념) [3][4][5] | **상** | 부분(창 간 BroadcastChannel) | **서버 권한 세션·op 로그·SSE·재동기화** |
| 실시간 전송 | SSE(서버→클라), HTTP POST(클라→서버) [10] | 중 | 없음 | stdlib 서버에서 SSE 스트리밍 |
| 프로세스 관리 | 분리 실행, PID/포트/락 파일, 헬스 체크, 유휴 종료, Windows 차이 [9] | 중 | 없음 | `xl serve/stop/status` |
| 패키징 | Skill 구조, 순수 Python 의존성 동봉(openpyxl·et_xmlfile MIT, xlrd·xlwt BSD) | 하~중 | 없음 | `scripts/_vendor` 동봉으로 `pip` 없이 실행 |
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
├─ SKILL.md                     # 500줄 이하: 언제 쓰는지, 시작/종료, 핵심 명령, 안전 규칙
├─ scripts/
│  ├─ xl.py                     # 단일 진입점 CLI (stdlib만 사용)
│  ├─ server/                   # 현재 server/ 패키지 + session.py, ops.py, refs.py(참조 재작성 포팅)
│  ├─ calc/worker.mjs           # Node가 있을 때 engine.js로 재계산하는 워커
│  └─ _vendor/                  # openpyxl, et_xmlfile, xlrd, xlwt (순수 Python, 오프라인 설치 불필요)
├─ assets/web/                  # 현재 static/ (index.html, js/, css/, vendor/chart.umd.min.js)
├─ references/
│  ├─ cli.md                    # 전체 명령·옵션·출력 형식
│  ├─ ops.md                    # op 스키마
│  ├─ formulas.md               # 지원 함수·제한
│  └─ troubleshooting.md        # 포트 충돌, DRM, Windows, 프록시
└─ LICENSE, THIRD_PARTY_NOTICES
```

`SKILL.md` 프론트매터 초안(이식 가능한 키만 사용):

```yaml
---
name: web-excel
description: >
  Open, inspect, edit, format, chart, filter and save Excel workbooks (.xlsx/.xlsm/.xls) in the
  user's workspace through a local web-excel server, while the user watches and edits the same
  workbook live in the agent's preview pane (http://localhost:<port>). Use when the user asks to
  read or change spreadsheet data, build formulas or charts, or wants to see/edit an Excel file.
license: MIT
compatibility: Requires Python 3.9+. Optional Node.js 18+ for server-side formula recalculation.
metadata:
  version: "0.1"
allowed-tools: Bash(python ${CLAUDE_SKILL_DIR}/scripts/xl.py *)
---
```

본문 핵심 지시(요지):
1. 시작: `python ${CLAUDE_SKILL_DIR}/scripts/xl.py serve` → 출력된 URL을 사용자에게 알리고 미리보기 창에서 열도록 안내.
2. 탐색은 `describe` → 필요한 범위만 `read`(대용량 전체 덤프 금지).
3. 변경은 가능한 한 `apply`(일괄 op)로 한 번에. 변경 후 `changes`로 결과 확인.
4. 사용자가 미리보기에서 편집 중일 수 있으므로 쓰기 전 `changes --since`로 최신 상태 확인.
5. **저장은 사용자가 요청했거나 명시적으로 동의한 경우에만** `save`. 덮어쓰기 전 자동 백업.
6. 작업 종료 시 서버는 유지(사용자가 계속 볼 수 있게). 사용자가 원하면 `stop`.

### 5.2 CLI (`xl.py`)

공통: 상태 파일 `.web-excel/state.json`(port, pid, token, root)을 찾아 서버에 연결. 출력 기본은 **간결한 텍스트/TSV**, `--json`으로 기계 판독 형식. 종료 코드: 0 성공, 2 사용법, 3 서버 없음, 4 충돌(409), 5 파일/DRM 오류.

| 명령 | 설명 |
|---|---|
| `serve [--port 8765] [--root DIR] [--no-open-browser] [--foreground]` | 멱등 기동. 이미 있으면 상태만 출력. 포트가 사용 중이면 다음 포트 탐색. `--foreground`는 launch.json 용 |
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

### 5.9 비기능 요구사항(목표치)

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
| P0 패키징·수명주기 | Skill 폴더 구조, 의존성 동봉, `xl.py serve/stop/status`, 분리 실행, state/lock, 포트 탐색, 유휴 종료, launch.json 도우미 | Python만 있는 새 PC에서 `serve` 후 미리보기에 UI 표시. 재실행 멱등. Windows/macOS 검증 | 2–3 |
| P1 서버 권한 세션 | `DocSession`, 스냅샷 API, 상태 op 적용(Python), op 로그, SSE, 브라우저 `sync.js` 교체(낙관적 적용 + echo 처리) | 브라우저 2개 + CLI 동시 셀 편집이 수렴. 재연결 시 누락 없음 | 5–7 |
| P2 에이전트 CLI | `open/describe/read/write/apply/format/sheet/chart/changes/selection/save` + 오류 코드 + TSV/JSON | 스크립트 시나리오(열기→요약→수정→차트→저장) 무인 통과 | 4–6 |
| P3 구조 op | 참조 재작성 Python 포팅, insert/delete/sort/rename, 409 + 재동기화·좌표 변환 | JS와 Python 재작성 결과 **동등성 테스트**(무작위 수식 1만 건) 일치 | 5–8 |
| P4 계산 단계화 | Node 워커, 브라우저 계산 보고(리더 선출), stale 표시 | 세 계산 원천의 결과가 테스트 코퍼스에서 일치 | 3–5 |
| P5 안전·보안 | 토큰/쿠키, 백업, 외부 변경 감지, 복구, 감사 로그, batch undo | 보안 체크리스트 통과, 충돌 시 데이터 손실 0 | 3–4 |
| P6 문서·평가 | SKILL.md, references/, Skill 트리거 평가(사용해야 할/말아야 할 요청 각 20개), 동시 편집 E2E | 트리거 정확도 ≥ 90%, E2E 통과 | 2–3 |
| **합계** | | | **24–36** |

우선순위: P0 → P1 → P2만으로도 "에이전트가 조작하고 사용자가 실시간으로 보는" 최소 기능(MVP, 약 11–16 인일)이 된다. 구조 op 충돌(P3)은 MVP에서 "구조 op는 서버 단일 순서로만 허용, 충돌 시 재동기화"로 단순 처리 가능.

---

## 7. 위험과 대응

| 위험 | 영향 | 대응 |
|---|---|---|
| DRM 에이전트가 분리 실행 python 또는 백업 사본을 다르게 취급 | 열기/저장 실패, 정책 위반 | P0에서 현장 PC 검증, 훅·`--no-backup` 제공 |
| 에이전트 SW가 세션 종료 시 프로세스 그룹 정리 | 서버 중단 | 새 세션/프로세스 그룹으로 분리, 복구 스냅샷 |
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

검증 메모: [1]은 이 세션에서 원문을 열람해 확인했고, [2][3][4][8]은 검색으로 서지 정보를 확인했다. agentskills.io 명세는 네트워크 정책으로 열람하지 못했다. [5][6][9][10][11]은 널리 알려진 표준·문서로, 링크는 이 환경에서 직접 열어 보지 못했다.
