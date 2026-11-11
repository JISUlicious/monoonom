# 설치

[English](install.md) | **한국어** · [README로](../../README.ko.md)

## 필요한 것

- **Python 3.9+**, `openpyxl`과 `Pillow`(필수: openpyxl은 Pillow가 있어야 통합 문서의 그림을 보존합니다), `xlrd`(`.xls`를 열 때만).
- 선택: **Node.js 18 이상**이 있으면 미리보기를 열지 않았을 때도 수식을 계산합니다.
- 에이전트가 **같은 컴퓨터**에서 실행되어야 합니다(미리보기가 `localhost`에 접속). 클라우드에서 도는 에이전트나 Skills API에서는 쓸 수 없습니다.

```bash
python3 -m pip install openpyxl pillow xlrd
```

인터넷이 없는 PC: 연결된 PC에서 `pip download openpyxl pillow xlrd -d wheels`를 실행하고 `wheels/`를 복사한 뒤
`pip install --no-index --find-links wheels openpyxl pillow xlrd`. 패키지가 없으면 `xl.py serve`가 설치 명령을 안내하고 종료 코드 6으로 끝납니다.

## Claude Code: 플러그인 (권장)

```bash
claude plugin marketplace add JISUlicious/monoonom
claude plugin install monoonom@monoonom
```

- 세션 안에서는 `/plugin install monoonom --marketplace JISUlicious/monoonom`(Claude Code 2.1.275 이상)로 한 번에 설치합니다.
- skill 이름은 `/monoonom:monoonom`입니다. 스프레드시트 작업을 요청하면 Claude가 알아서 씁니다.
- 업데이트: `claude plugin update monoonom@monoonom`. 제거: `claude plugin uninstall monoonom@monoonom`.
- 플러그인판은 자기 `xl.py` 명령을 미리 허용하고(`allowed-tools`), skill을 불러올 때 서버 상태를 보여 주며, Claude Code 데스크톱에서 미리보기 창을 열도록 안내합니다.

## 다른 에이전트 (Codex, Cursor, Copilot 등): `npx skills`

```bash
npx skills add JISUlicious/monoonom            # 현재 프로젝트에 (에이전트 선택)
npx skills add JISUlicious/monoonom -g -a codex # 모든 프로젝트, 특정 에이전트
```

- [skills CLI](https://github.com/vercel-labs/skills)가 `.agents/skills/monoonom/`(이식판 `SKILL.md`)를 설치합니다.
- 업데이트: `npx skills update monoonom`. 제거: `npx skills remove monoonom`.
- Claude Code에 이 방법으로 설치해도 동작하지만 위의 플러그인 기능은 빠집니다.
- Codex의 `workspace-write` 모드는 로컬 포트 바인딩을 막으므로 명령을 승인하거나 전체 권한을 주어야 합니다.

## 직접 복사

`.agents/skills/monoonom/` 폴더를 에이전트의 skill 폴더(예: `~/.agents/skills/`, Claude Code는 `~/.claude/skills/`)에
복사합니다. Claude Code에 복사할 때는 `SKILL.claude-code.md`를 `SKILL.md`로 바꿔 넣습니다.
