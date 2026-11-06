# MoNooNoM

[English](README.md) | **한국어**

<img src=".agents/skills/monoonom/assets/web/img/monoonom.svg" alt="" width="72" align="right">

사람과 AI 에이전트가 **같은 스프레드시트를 함께 보면서 작업하는** 로컬 웹 스프레드시트입니다.
에이전트가 Excel 파일(xlsx·xlsm, xls 읽기)을 열고 고치면, 사용자는 에이전트 앱의 미리보기 창(또는 브라우저)에서
그 변화를 실시간으로 보고 직접 편집할 수 있습니다. 이름은 '모눈(종이)'을 영어로 쓰고 어느 쪽에서도 읽히게 만든 것입니다(MoNooN + NooNoM).

![사용자가 보는 동안 에이전트가 비율 열과 차트를 추가하는 화면](docs/guide/images/monoonom-agent.png)
*에이전트가 차트를 추가한 뒤 Share(비율) 열을 채우는 중입니다. 바뀐 셀이 강조되고 상태 표시줄에 에이전트가 한 일이 표시됩니다.*

- **함께 편집**: 에이전트와 사용자의 편집이 순서대로 합쳐져 열린 모든 창에 반영됩니다. 에이전트는 사용자가 바꾼 셀을 읽습니다.
- **Excel 호환**: 수식 약 200개 함수, 서식, 필터·정렬, 차트, 피벗 테이블, Excel 단축키. 저장하면 앱이 다루지 않는 조건부 서식·이미지·VBA 등은 그대로 보존됩니다.
- **로컬 전용**: 서버는 `127.0.0.1`에서만 동작하고, 파일은 작업 폴더 밖으로 나가지 않으며, 파일 읽기·쓰기에 DRM 훅을 붙일 수 있습니다.

## 빠른 시작

에이전트와 같은 컴퓨터에 Python 3.9+가 필요합니다: `python3 -m pip install openpyxl xlrd`.

```bash
# Claude Code
claude plugin marketplace add JISUlicious/monoonom
claude plugin install monoonom@monoonom

# Codex, Cursor, Copilot 등 다른 에이전트
npx skills add JISUlicious/monoonom
```

그다음 에이전트에게 요청하면 됩니다. 예: "report.xlsx 열어서 같이 보자", "월별 합계 열 추가하고 꺾은선 차트 만들어 줘".
에이전트가 열어 볼 미리보기 주소를 알려 줍니다.

## 문서

| | |
|---|---|
| [설치](docs/guide/install.ko.md) | 필요한 것, 플러그인, `npx skills`, 직접 복사, 업데이트·제거 |
| [사용](docs/guide/usage.ko.md) | 미리보기, 에이전트 따라가기, 화면 언어(`?lang=`), 명령, 단축키 |
| [보안과 DRM](docs/guide/security.ko.md) | 토큰, 로컬 전용 서버, 작업 폴더 제한, DRM 훅 |
| [알려진 제한](docs/guide/limitations.ko.md) | 표시하지 않지만 보존하는 것, xls, 피벗, 차트 |
| [명령 참조](.agents/skills/monoonom/references/cli.md) | `xl.py`의 모든 명령 (에이전트가 읽는 문서, 영문) |
| [개발](docs/dev/development.md) | 저장소 구조, 설계 결정, 테스트 |

## 라이선스

MIT ([LICENSE](LICENSE)). 함께 넣은 외부 코드는 Chart.js(MIT) 하나이며 고지는
[`THIRD_PARTY_NOTICES.md`](.agents/skills/monoonom/THIRD_PARTY_NOTICES.md)에 있습니다.

## 상표

Microsoft, Excel은 미국 및 기타 국가에서 Microsoft Corporation의 상표입니다. MoNooNoM은 Microsoft와 관련이 없으며 Microsoft가 후원하거나 보증하지 않습니다. 이 문서에서 'Excel'은 파일 형식과 호환되는 동작을 설명하는 데에만 씁니다.
