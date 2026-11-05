# monoonom skill 평가

실행기: `tools/run_skill_evals.py`(`trigger` / `scenarios` / `report`). Claude Code CLI(`claude -p`) 2.1.318으로 실행.
실행마다 새 작업 폴더를 만들고, Claude Code가 읽는 형태(`.claude/skills/monoonom`: 생성된 SKILL.md + 패키지 링크)로 skill을 설치한다.
결과 원본: `docs/evals/<날짜>/`(trigger.json, scenarios.json, report.md).

## 2026-11-05 (첫 실행, Haiku·Sonnet·Opus)

### 트리거 (질문 20개: 써야 함 10, 쓰면 안 됨 10; 모델별 1회)
설정: Bash 금지, 3턴 이내에 Skill 도구로 monoonom을 부르면 "호출"로 센다. 같은 환경에 claude.ai의 `anthropic-skills:xlsx` 등 다른 skill도 설치되어 있어 실제 경쟁 조건이다.

| 모델 | 써야 할 때 호출 | 쓰면 안 될 때 안 부름 | 정확도 | 비용 |
|---|---|---|---|---|
| Haiku | 5/10 | 10/10 | 15/20 | $0.06 |
| Sonnet | 8/10 | 10/10 | 18/20 | $1.25 |
| Opus | 8/10 | 10/10 | 18/20 | $2.46 |

- 오발동은 0건이다.
- 세 모델이 모두 놓친 질문은 둘이다. 둘 다 화면이나 미리보기 언급이 없는 영어 단순 편집 요청이다.
  - "insert two rows above row 10 … stock.xlsx"
  - "Rename the '매출' sheet … without breaking formulas"
  - Sonnet·Opus는 이 경우 `anthropic-skills:xlsx`를 골랐다.
- Haiku는 skill 없이 Bash나 openpyxl로 바로 처리하려는 경향이 강하다.

### 시나리오 (6개 × 3개 모델, 채점: opus가 기록과 최종 통합 문서 상태로 기대 항목별 판정)

| 시나리오 | Haiku | Sonnet | Opus |
|---|---|---|---|
| add-margin-column | 4/5 | 4/5 | 4/5 |
| react-to-user-selection | 3/3 | 3/3 (skill 미호출) | 2/3 |
| insert-and-sort-then-save | 0/4 (skill 미호출) | 4/4 | 4/4 |
| user-edits-meanwhile | 0/2 (skill 미호출) | 2/2 (skill 미호출) | 2/2 |
| give-preview-url | 4/4 | 4/4 | 4/4 |
| summary-with-formulas | 1/5 (skill 미호출) | 4/5 | 5/5 |
| 합계 | 12/23 ($0.05) | 21/23 ($0.57) | 21/23 ($1.15) |

해석:
- **평가 기준 문제**: add-margin-column의 "xl.py serve를 실행해 URL 전달" 항목은 세 모델 모두 실패로 채점됐다. 평가 환경이 서버를 미리 띄워 두었으므로, `xl.py url`로 받은 주소를 그대로 준 것은 올바른 동작이다. 기대 항목을 "serve 또는 url"로 고쳐야 한다.
- **skill을 쓰면 좋은 결과**: skill을 쓴 Sonnet·Opus는 행 삽입·정렬·저장(백업 안내 포함), 사용자 선택 범위 처리, 수식 요약을 모두 xl.py로 했다.
- **skill을 쓰지 않으면 위험**: Haiku는 skill을 부르지 않은 3개 시나리오에서 xlsx XML을 직접 고쳤다. 그리고 `cp`로 **원본 파일을 묻지 않고 덮어썼다**. 미리보기에는 아무것도 반영되지 않았다.
- **작은 문제**: Opus는 편집 후 되돌리는 방법을 사용자에게 알리지 않은 적이 있다(react-to-user-selection).

### 첫 시나리오 실행이 무효였던 이유 (같은 날 재실행)
- Claude Code판 SKILL.md에는 `allowed-tools`가 있다. 그래서 **skill을 호출하는 것 자체가 권한 요청이 된다**(stream-json에 `permission_denied … "Execute skill: monoonom"`).
- `-p` 실행에서 `--allowedTools`에 `Skill`이 없으면 이 요청이 거부되고, 에이전트는 지시 없이 작업했다. 그 결과 XML을 직접 고치거나 원본을 덮어썼다.
- 이식판(`allowed-tools` 없음)은 정상 로드되고, `allowed-tools` 줄만 빼도 정상 로드됨을 확인했다. `Skill` 또는 `Skill(monoonom)`을 허용하면 해결된다.
- 대화형 세션에서는 사용자에게 "Execute skill" 허용을 묻는다. 비대화형(`-p`, SDK, CI)에서는 아무 표시 없이 실패하므로 사용자 안내가 필요하다.
- 무효였던 결과는 저장하지 않았다. 트리거 결과는 skill을 고르는 판단이 로드보다 먼저라서 유효하다.

## 후속 검토 항목 (2026-11-05, 보류)
평가와 내용 검토에서 나온 개선 후보입니다. 필수 변경은 아니라고 판단해 기록만 해 둡니다.

| # | 항목 | 근거 | 영향 |
|---|---|---|---|
| A | 비대화형 실행 안내: README에 "`-p`/SDK/CI에서는 `Skill(monoonom)`을 허용"을 적고, 이 저장소 `.claude/settings.json`에 허용 규칙을 넣는다 | Claude Code판의 `allowed-tools` 때문에 skill 호출이 권한 요청이 되고, 비대화형에서는 아무 표시 없이 거부됨(위 "첫 시나리오 실행이 무효였던 이유") | 대화형 세션은 사용자에게 한 번 허용을 묻는 것으로 끝나므로 문제없음. 자동화에서만 해당 |
| B | description 보강: "통합 문서를 openpyxl/XML로 직접 고치는 대신 이 skill을 쓸 것, 차트·서식이 보존되고 사용자가 화면으로 확인 가능" | 트리거 평가에서 화면 언급이 없는 단순 편집 요청을 놓침. Haiku는 skill을 부르지 않고 원본 파일을 덮어씀. 오발동은 0건이라 여유가 있음 | 트리거 동작이 바뀌므로 고치면 트리거 평가를 다시 돌릴 것 |
| C | 이식판 SKILL.md의 `python scripts/xl.py`를 `python3`(Windows는 `python`)로 | macOS·Linux에는 `python` 명령이 없는 경우가 많음(Claude Code판은 이미 `python3`) | 에이전트가 대개 스스로 고쳐 실행하므로 작음 |
| D | `xl.py changes --since VERSION`의 VERSION이 마지막 명령 출력의 `vN`임을 명시 | 본문에 정의가 없음 | 작음 |
| E | 편집 후 되돌리는 명령(`xl.py undo BATCH`)을 사용자에게 알리라는 지시 | Opus가 한 번 빠뜨림(react-to-user-selection) | 작음 |
| F | 평가 기대 항목 "xl.py serve를 실행해 URL 전달" → "serve 또는 url로 받은 URL을 그대로 전달" | 서버가 이미 떠 있으면 `url`이 올바른 동작인데 실패로 채점됨 | 평가 기준만 |

내용 검토에서 확인한 것: SKILL.md의 명령 예시 63개, `references/cli.md`의 옵션, `references/edits.md` 예시, 문서에 적힌 동작(저장 안 된 변경이 있으면 `stop`이 종료 코드 4로 거부, `.xls`는 같은 이름의 `.xlsx`로 저장, 백업, 의존성 누락 시 종료 코드 6)이 모두 실제 CLI와 맞다.
