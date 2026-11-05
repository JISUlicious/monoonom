# web-excel skill 평가

실행기: `tools/run_skill_evals.py`(`trigger` / `scenarios` / `report`). Claude Code CLI(`claude -p`) 2.1.318으로 실행.
실행마다 새 작업 폴더를 만들고, Claude Code가 읽는 형태(`.claude/skills/web-excel`: 생성된 SKILL.md + 패키지 링크)로 skill을 설치한다.
결과 원본: `docs/evals/<날짜>/`(trigger.json, scenarios.json, report.md).

## 2026-11-05 (첫 실행, Haiku·Sonnet·Opus)

### 트리거 (질문 20개: 써야 함 10, 쓰면 안 됨 10; 모델별 1회)
설정: Bash 금지, 3턴 이내에 Skill 도구로 web-excel을 부르면 "호출"로 센다. 같은 환경에 claude.ai의 `anthropic-skills:xlsx` 등 다른 skill도 설치되어 있어 실제 경쟁 조건이다.

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
- Claude Code판 SKILL.md에는 `allowed-tools`가 있다. 그래서 **skill을 호출하는 것 자체가 권한 요청이 된다**(stream-json에 `permission_denied … "Execute skill: web-excel"`).
- `-p` 실행에서 `--allowedTools`에 `Skill`이 없으면 이 요청이 거부되고, 에이전트는 지시 없이 작업했다. 그 결과 XML을 직접 고치거나 원본을 덮어썼다.
- 이식판(`allowed-tools` 없음)은 정상 로드되고, `allowed-tools` 줄만 빼도 정상 로드됨을 확인했다. `Skill` 또는 `Skill(web-excel)`을 허용하면 해결된다.
- 대화형 세션에서는 사용자에게 "Execute skill" 허용을 묻는다. 비대화형(`-p`, SDK, CI)에서는 아무 표시 없이 실패하므로 사용자 안내가 필요하다.
- 무효였던 결과는 저장하지 않았다. 트리거 결과는 skill을 고르는 판단이 로드보다 먼저라서 유효하다.
