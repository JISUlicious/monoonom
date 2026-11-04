# 작업 기록

최신 항목이 위에 온다.

- 2026-11-04: 코드 리뷰 8건 반영(xls 저장 대상 이름·형식, 409에 덮어쓸 파일 이름, 값 변환 경고 등), 미리보기 시작 중 에이전트가 연 문서가 탭으로 안 뜨던 버그 수정 (→ docs/skill-implementation.md)
- 2026-11-04: Python 패키지 동봉 중단 — `scripts/_vendor` 삭제, `scripts/requirements.txt`로 설치 안내(없으면 종료 코드 6과 설치 명령), 라이선스 고지는 Chart.js만 남음 (→ docs/skill-implementation.md)
- 2026-11-04: xls 읽기 전용화·xlwt 제거 — xls 저장은 같은 이름의 .xlsx로(원본 유지), LGPL 의존성 제거, xlrd 고지 문구 누락 보완 (→ docs/skill-implementation.md)
- 2026-11-04: 공개 준비 — MIT `LICENSE`, 동봉 라이브러리 고지(`THIRD_PARTY_NOTICES.md`, xlwt 라이선스 원문·LGPL-2.1 원문 추가), 로컬 경로 제거 (→ README.md)
- 2026-11-04: 계산할 수 없는 함수가 있는 파일 — 열 때 경고(CLI·미리보기·describe), 고치지 않은 수식은 원래 텍스트·결과 그대로 저장, 고친 수식도 파일의 `_xlfn.` 접두사 유지, 파서가 못 읽는 수식에서 Enter가 막히던 버그 수정 (→ docs/skill-implementation.md)
- 2026-11-04: Excel 호환 저장 점검·수정 — 수식 결과를 캐시 값으로 저장(저장 전 계산 갱신), LET 인자 `_xlpm.` 접두사, `_xlfn.` 누락 함수 26개 추가, NORM.DIST 정밀도(1e-7→1e-15), YEARFRAC 5개 기준·DAYS360 US/EU 방식, LibreOffice 대조 테스트 229개 (→ docs/skill-implementation.md)
- 2026-11-04: A1이 아닌 곳에서 시작하는 표 처리 — 합계 행을 필터·정렬 대상에서 제외(정렬 시 #REF! 버그 수정), 200행 넘는 표 영역 잘림 수정, 에이전트 명령의 셀 하나/열 전체 → 표 자동 확장, describe에 표 목록 (→ docs/skill-implementation.md)
- 2026-11-04: Skill P7 — UI 기능 전부를 에이전트 명령으로(필터 조건·차트 편집·시트 복사/이동/숨기기·연속 채우기·복사/붙여넣기·셀 밀기·중복 제거·찾기/바꾸기·메모·링크·자동 맞춤·서식 복사·화면 이동 등), 대응표 작성 (→ docs/agent-coverage.md, docs/skill-implementation.md)
- 2026-11-04: Skill P4–P6 마무리 — 감사 로그(셀 내용 제외), 평가 시나리오 4개·트리거 질문 20개, 공식 검증기 규칙 확인(실행은 권한상 보류), 평가 실행은 승인 대기 (→ docs/skill-implementation.md)
- 2026-11-04: Skill P3 구현 — 수식 참조 재작성 Python 이식(JS와 1만 건 동등성 확인), 에이전트 행/열 삽입·삭제·정렬, 동시 편집 시 셀 편집 좌표 변환 (→ docs/skill-implementation.md)
- 2026-11-04: 고해상도 화면에서 차트가 상자의 1/4 크기로 그려지던 버그 수정, 에이전트 편집 중 미리보기 화면 캡처 추가 (→ docs/skill-implementation.md)
- 2026-11-04: Skill P2 구현 — 에이전트 CLI(열기·요약·읽기·쓰기·채우기·서식·차트·시트·일괄 적용·변경 이력·되돌리기·저장), Node 계산, SKILL.md 완성, 참조 문서 (→ docs/skill-implementation.md)
- 2026-11-04: Skill P1 구현 — 서버가 문서를 소유(op·버전·로그), 창당 SSE 스트림, 브라우저 낙관적 적용·대기 추적, 계산값 보고, 저장 백업·디스크 변경 감지, batch 되돌리기 (→ docs/skill-implementation.md)
- 2026-11-04: Skill P0 구현 — skill 폴더 구조(`.agents/skills/web-excel`), 의존성 동봉, `xl.py serve/status/stop/launch-config`, 토큰 인증, 유휴·부모 종료 감지, SKILL.md 2판 (→ docs/skill-implementation.md)
- 2026-11-04: Agent Skill spec 2차 보강 — analysis-docs 세션 회신 반영(name=디렉터리 규칙, allowed-tools 실험적, metadata 문자열 맵, 제품별 경로, Codex loopback 차단 실측, 서버 수명주기 모드 A/B, skills-ref 검증) (→ docs/skill-spec.md)
- 2026-11-04: Agent Skill spec 보강 — analysis-docs의 작성 표준 반영(description `<>` 금지, P0/P1 두 판, 실행 표면별 지원 범위, 평가 방식) (→ docs/skill-spec.md)
- 2026-11-04: Agent Skill 포장 조사·계획·spec 작성 — 결론: 구현 가능, 서버가 문서를 소유하는 동기화 구조로 변경 필요, 24–36인일 (→ docs/skill-spec.md)
- 2026-11-04: 웹 엑셀 1차 구현 — Python 서버(xlsx/xls, DRM 훅), Canvas 그리드, 수식 엔진, 차트, 서식, 필터, 다중 탭·창, 단축키, 테스트 (→ README.md)
- 2026-11-04: 문서화 규칙 적용 — `index.md`, `logs.md`, `docs/` 추가 (→ docs/project-setup.md)
- 2026-11-04: 프로젝트 생성 — GitHub 저장소 `JISUlicious/web-excel` (→ docs/project-setup.md)
