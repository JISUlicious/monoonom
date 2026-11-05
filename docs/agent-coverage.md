# UI 기능 ↔ 에이전트 명령 대응표

미리보기(브라우저 UI)에서 할 수 있는 작업을 에이전트가 `xl.py`로 할 수 있는지 정리한 표 (2026-11-04, P7 기준).
명령 형식은 `.agents/skills/web-excel/references/cli.md`, 일괄 편집(JSON)은 `references/edits.md`.

## 파일·문서

| UI 기능 | 에이전트 명령 | 비고 |
|---|---|---|
| 새로 만들기 / 열기 / 닫기 | `new`, `open PATH`, `close` | 에이전트가 연 문서는 미리보기에 탭으로 자동 표시 |
| 저장 / 다른 이름으로 저장 / xls·xlsm | `save`, `save --as P --format xls` | 덮어쓰기 전 백업, 디스크 변경 감지 |
| 문서 탭 전환, 새 창 | `goto`(문서 지정 시 해당 탭 활성화) | 새 브라우저 창 자체는 에이전트가 열 수 없음 |
| 파일 열기 대화상자 | `ls [DIR]` | |

## 홈 탭

| UI 기능 | 에이전트 명령 |
|---|---|
| 값·수식 입력, Ctrl+Enter(선택 영역 모두) | `write`, `fill RANGE VALUE` |
| 잘라내기 / 복사 / 붙여넣기 | `move FROM TO`, `copy FROM TO` |
| 선택하여 붙여넣기(값/서식/수식/행열 바꿈) | `copy … --what values\|formats\|formulas --transpose` |
| 서식 복사 | `format-copy FROM TO` |
| 글꼴·크기·크게/작게·굵게·기울임·밑줄·취소선·색 | `format RANGE ff= fs= b= i= u= st= fc=` |
| 채우기 색 | `format RANGE bg=#RRGGBB` |
| 테두리(13종) | `format RANGE --border all\|outline\|inside\|top\|bottom\|left\|right\|none[:style[:#color]]` |
| 맞춤(가로/세로), 줄 바꿈, 들여쓰기 | `format RANGE ha= va= wrap=true ind=N` |
| 병합하고 가운데 맞춤 / 전체 병합 / 셀 분할 | `merge`, `merge --across`, `unmerge` |
| 표시 형식(일반·숫자·통화·회계·날짜·시간·%·분수·지수·텍스트·사용자 지정) | `format RANGE nf='…'` |
| 쉼표·%·₩ 스타일, 자릿수 늘림/줄임 | `format nf=…`, `decimals RANGE ±N` |
| 셀 삽입/삭제(밀기), 행/열 삽입/삭제 | `cells insert\|delete --shift …`, `rows`, `cols` |
| 행 높이·열 너비·자동 맞춤·숨기기 | `height`, `width`, `autofit`, `hide`, `unhide` |
| 시트 삽입·삭제·이름·복사·이동·탭 색·숨기기 | `sheet add\|rm\|rename\|copy\|move\|color\|hide\|unhide` |
| 자동 합계, 채우기(아래로/오른쪽), 연속 데이터 | `write '=SUM(…)'`, `copy`, `series SOURCE TARGET` |
| 지우기(모두/서식/내용/메모) | `clear --what all\|formats\|contents\|comments` |
| 정렬, 사용자 지정 정렬(3단계) | `sort RANGE --by C:desc --by A …` |
| 찾기 / 바꾸기 | `find`, `replace` |

## 삽입·수식·데이터·보기

| UI 기능 | 에이전트 명령 |
|---|---|
| 차트 삽입(9종) | `chart add DATA --type …` |
| 차트 편집(종류·데이터·제목·범례·축 제목·위치·크기)·삭제 | `chart set ID …`, `chart rm ID` |
| 메모, 하이퍼링크 | `comment`, `link` |
| 함수 삽입 | `write '=FUNC(…)'` (함수 목록은 SKILL 참고 문서) |
| 수식 표시, 확대/축소 | `view --formulas`, `view --zoom 150` |
| 지금 계산(F9) | 자동(읽을 때 계산값 최신화) |
| 필터 켜기/끄기, 값 목록·조건 필터, 지우기, 다시 적용 | `filter RANGE\|off`, `filter col C --values\|--where`, `filter clear`, `filter reapply` |
| 중복된 항목 제거 | `dedupe` |
| 눈금선 | `sheet gridlines NAME off` |
| 틀 고정 / 첫 행 / 첫 열 / 취소 | `freeze B2`, `freeze A2`, `freeze B1`, `freeze none` |
| 이름 상자로 이동 | `goto RANGE` (사용자 화면 이동) |

## 읽기·협업 (UI에는 없고 에이전트에 필요한 것)

| 목적 | 명령 |
|---|---|
| 구조 파악 | `describe` |
| 값/화면 텍스트/수식/서식 읽기 | `read --what values\|text\|raw\|formulas\|all` |
| 사용자가 바꾼 것 | `changes --since V --actor user` |
| 사용자가 선택한 범위 | `selection` |
| 되돌리기 | `undo BATCH` (에이전트·사용자 변경 모두 batch id로) |

## 에이전트로 할 수 없는 것

| 기능 | 이유 / 대안 |
|---|---|
| 사용자 PC 클립보드(다른 앱과 복사·붙여넣기) | 브라우저 클립보드는 사용자 조작에서만 접근 가능. 시트 안 복사는 `copy` |
| 새 브라우저 창 열기 | 브라우저 보안. `goto`로 기존 미리보기를 이동 |
| 사용자 화면의 Ctrl+Z 기록 | 창마다 따로 관리됨. 서버 기록 기준 `undo BATCH`는 가능 |
| 키보드 탐색·편집 모드 같은 화면 조작 | 결과를 직접 쓰는 명령으로 대체 |

## 화면에는 표시되지 않지만 보존되는 것 (UI·에이전트 공통)

조건부 서식, 데이터 유효성, 이미지, VBA(.xlsm): 편집 기능은 없고 xlsx 저장 시 원본에서 보존된다. 피벗 테이블: 보존·`describe` 표시·영역 보호·위치 이동·원본 변경 시 열 때 새로 고침(1단계). 새로 고침·만들기는 없다.
