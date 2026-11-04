# 작업 기록

최신 항목이 위에 온다.

- 2026-11-04: 고해상도 화면에서 차트가 상자의 1/4 크기로 그려지던 버그 수정, 에이전트 편집 중 미리보기 화면 캡처 추가 (→ docs/skill-implementation.md)
- 2026-11-04: Skill P2 구현 — 에이전트 CLI(열기·요약·읽기·쓰기·채우기·서식·차트·시트·일괄 적용·변경 이력·되돌리기·저장), Node 계산, SKILL.md 완성, 참조 문서 (→ docs/skill-implementation.md)
- 2026-11-04: Skill P1 구현 — 서버가 문서를 소유(op·버전·로그), 창당 SSE 스트림, 브라우저 낙관적 적용·대기 추적, 계산값 보고, 저장 백업·디스크 변경 감지, batch 되돌리기 (→ docs/skill-implementation.md)
- 2026-11-04: Skill P0 구현 — skill 폴더 구조(`.agents/skills/web-excel`), 의존성 동봉, `xl.py serve/status/stop/launch-config`, 토큰 인증, 유휴·부모 종료 감지, SKILL.md 2판 (→ docs/skill-implementation.md)
- 2026-11-04: Agent Skill spec 2차 보강 — analysis-docs 세션 회신 반영(name=디렉터리 규칙, allowed-tools 실험적, metadata 문자열 맵, 제품별 경로, Codex loopback 차단 실측, 서버 수명주기 모드 A/B, skills-ref 검증) (→ docs/skill-spec.md)
- 2026-11-04: Agent Skill spec 보강 — analysis-docs의 작성 표준 반영(description `<>` 금지, P0/P1 두 판, 실행 표면별 지원 범위, 평가 방식) (→ docs/skill-spec.md)
- 2026-11-04: Agent Skill 포장 조사·계획·spec 작성 — 결론: 구현 가능, 서버가 문서를 소유하는 동기화 구조로 변경 필요, 24–36인일 (→ docs/skill-spec.md)
- 2026-11-04: 웹 엑셀 1차 구현 — Python 서버(xlsx/xls, DRM 훅), Canvas 그리드, 수식 엔진, 차트, 서식, 필터, 다중 탭·창, 단축키, 테스트 (→ README.md)
- 2026-11-04: 문서화 규칙 적용 — `index.md`, `logs.md`, `docs/` 추가 (→ docs/project-setup.md)
- 2026-11-04: 프로젝트 생성 — GitHub 비공개 저장소 `JISUlicious/web-excel`, 로컬 `~/data/workspace/vibe/apps/web-excel` (→ docs/project-setup.md)
