# 프로젝트 구성

## 저장소

- GitHub: `JISUlicious/web-excel-agent-skills` (공개), 기본 브랜치 `main`. 옛 이름 `JISUlicious/web-excel`은 GitHub가 새 주소로 연결한다.
- 로컬: 개발자 작업 폴더에 clone
- Claude GitHub 앱이 "All repositories"로 설치돼 있어 클라우드 세션에서 바로 쓸 수 있다.

## 결정과 이유

- **로컬 경로 `vibe/apps/`**: 공부(`study/`)나 게임(`games/`)에 해당하지 않아 앱 분류 폴더 `apps/`를 새로 만들어 넣었다.
  고려한 대안은 분류 없이 `vibe/web-excel`에 두는 것과 `vibe/web/web-excel`이었다.
- **비공개 → 공개**: 처음에는 공개 여부를 정하지 않아 비공개로 만들었다. 공개 전에 MIT `LICENSE`를 추가하고, Python 패키지 동봉을 없애 고지 대상을 Chart.js 하나로 줄인 뒤 공개했다.
- **이름 `web-excel-agent-skills`**: 공개하면서 Agent Skill이라는 점이 드러나도록 바꿨다.

## 변경 이력

- 2026-11-04: 공개 전환, 저장소 이름 변경(`web-excel` → `web-excel-agent-skills`), PR #1 squash 병합
- 2026-11-04: 문서 작성 (저장소 생성, 문서화 규칙 적용)
