# monoonom 문서 색인

## 기본 문서

- [README.md](README.md) — 영문 첫 화면: 소개, 빠른 시작, 문서 목록
- [README.ko.md](README.ko.md) — README 한국어판
- [logs.md](logs.md) — 작업 기록

## docs/guide/ — 사용자 문서 (영문 + `.ko.md` 한국어판)

- [install.md](docs/guide/install.md) · [install.ko.md](docs/guide/install.ko.md) — 설치: 필요한 것, 플러그인, `npx skills`, 직접 복사, 업데이트·제거
- [usage.md](docs/guide/usage.md) · [usage.ko.md](docs/guide/usage.ko.md) — 사용: 미리보기, 에이전트 따라가기, 화면 언어, 명령, 단축키
- [security.md](docs/guide/security.md) · [security.ko.md](docs/guide/security.ko.md) — 보안과 DRM
- [limitations.md](docs/guide/limitations.md) · [limitations.ko.md](docs/guide/limitations.ko.md) — 알려진 제한
- `docs/guide/images/` — README·가이드 화면

## docs/dev/ — 개발 문서 (한국어)

- [development.md](docs/dev/development.md) — 배포 구성, 저장소 구조, 설계 결정과 근거, 테스트 실행, 최초 요구사항 대응표, 참고 문헌
- [project-setup.md](docs/dev/project-setup.md) — 저장소·폴더 구성과 그 이유
- [skill-spec.md](docs/dev/skill-spec.md) — Agent Skill로 포장하기 위한 기술 조사·실현 가능성·구현 계획·명세
- [agent-coverage.md](docs/dev/agent-coverage.md) — UI 기능 ↔ 에이전트 명령 대응표
- [skill-implementation.md](docs/dev/skill-implementation.md) — 단계별(P0–) 구현 기록: 한 일, 결정과 이유, 검증, spec과 달라진 점
- [skill-evals.md](docs/dev/skill-evals.md) — skill 평가 방법·결과, 미뤄 둔 후속 항목 (원본: `docs/dev/evals/`)
- `docs/dev/images/` — 구현 기록용 화면

## Agent Skill

- [.agents/skills/monoonom/](.agents/skills/monoonom/) — skill 패키지 원본 (`SKILL.md`, `scripts/xl.py`, `assets/web/`, `references/`)
- [.agents/skills/monoonom/references/cli.md](.agents/skills/monoonom/references/cli.md) — 에이전트 CLI 전체 명령
- [.agents/skills/monoonom/references/edits.md](.agents/skills/monoonom/references/edits.md) — 일괄 편집(JSON) 형식
- [.claude/skills/monoonom](.claude/skills/monoonom) — 이 저장소의 Claude Code 프로젝트 skill (SKILL.md + 원본 링크)
- [.claude-plugin/](.claude-plugin/) — Claude Code 플러그인·마켓플레이스 (`plugin.json`, `marketplace.json`, 플러그인 skill)
- [evals/monoonom/](evals/monoonom/) — skill 평가 시나리오·트리거 질문 (skill 패키지 밖)
- [tools/build_claude_code_skill.py](tools/build_claude_code_skill.py) — SKILL.md에서 Claude Code판 생성

## tests/

- [tests/](tests/) — Python 왕복 테스트, 수식 엔진 테스트, 브라우저 E2E (실행 방법은 docs/dev/development.md의 "테스트"). 브라우저 테스트는 `tests/korean.cjs`로 한국어 화면에서 실행
