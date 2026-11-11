# 보안과 DRM

[English](security.md) | **한국어** · [README로](../../README.ko.md)

## 보안

- 서버는 `127.0.0.1`에만 바인딩하고 `Host` 헤더를 검사합니다(DNS 리바인딩 방지).
- 모든 API는 작업 폴더별 토큰이 필요합니다. 미리보기 주소의 `#t=…`(헤더로 보냄)로, 또는 토큰 없이 연 탭은 포트별 쿠키로 확인합니다. 변경 요청에는 사용자 지정 헤더도 필요해서, 다른 사이트의 요청은 CORS 사전 요청 없이 보낼 수 없습니다.
- 작업 폴더 밖의 파일은 열 수 없습니다(`--no-sandbox`로 해제).
- 서버 상태(`server.json`, 토큰, 로그, 백업)는 `<작업 폴더>/.monoonom/`에 있습니다. 버전 관리에 넣지 마세요.
- 토큰이 붙은 미리보기 주소가 유출되었다면 `xl.py stop` 뒤 `xl.py serve --new-token`을 실행하고 새 주소로 다시 여세요.

## DRM 환경

브라우저는 파일 내용을 읽지 않고 경로만 보냅니다. 모든 읽기·쓰기가 `scripts/server/fileio.py`를 지나므로, 허가된
프로세스(Python)에 투명 복호화를 해 주는 DRM이면 그대로 동작합니다. DRM SDK를 호출해야 하면 함수를 끼웁니다:

```bash
MONOONOM_READ_HOOK=mypkg.drm:read_file     # (path) -> bytes
MONOONOM_WRITE_HOOK=mypkg.drm:write_file   # (path, data: bytes) -> None
```

파일 식별자를 추적하는 DRM을 위해 임시 파일+이름 바꾸기 없이 같은 경로에 직접 씁니다. 읽은 파일이 zip(xlsx)이 아니면
읽기 훅을 지정하라는 안내와 함께 오류를 냅니다. 페이지에 파일을 끌어 놓는 것은 일부러 막았습니다(파일은 서버를 통해 엽니다).
