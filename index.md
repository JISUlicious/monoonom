# web-excel 문서 색인

## 기본 문서

- [README.md](README.md) — 프로젝트 소개, 실행 방법, 요구사항 대응표, 구조·설계 결정, 제한, 참고 문헌
- [logs.md](logs.md) — 작업 기록

## docs/

- [project-setup.md](docs/project-setup.md) — 저장소·폴더 구성과 그 이유
- [skill-spec.md](docs/skill-spec.md) — Agent Skill로 포장하기 위한 기술 조사·실현 가능성·구현 계획·명세
- [skill-implementation.md](docs/skill-implementation.md) — 단계별(P0–) 구현 기록: 한 일, 결정과 이유, 검증, spec과 달라진 점

## Agent Skill

- [.agents/skills/web-excel/](.agents/skills/web-excel/) — skill 패키지 원본 (`SKILL.md`, `scripts/xl.py`, `assets/web/`, `references/`)
- [.claude/skills/web-excel](.claude/skills/web-excel) — Claude Code용 심볼릭 링크

## tests/

- [tests/](tests/) — Python 왕복 테스트, 수식 엔진 테스트, 브라우저 E2E (실행 방법은 README의 "테스트")
