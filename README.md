# UNI&CORE 재고 대시보드 - GitHub Pages 자동 배포

SharePoint의 재고 데이터를 매일 자동으로 읽어 웹 대시보드를 업데이트합니다.

---

## 세팅 가이드 (처음 한 번만 하면 됩니다)

### 1단계: GitHub 저장소 생성

1. [github.com](https://github.com) 에 로그인합니다
2. 우측 상단 **+** → **New repository** 클릭
3. 설정:
   - **Repository name**: `unicore-inventory` (원하는 이름)
   - **Public** 선택 (GitHub Pages 무료 사용을 위해)
   - 나머지는 기본값
4. **Create repository** 클릭

### 2단계: 파일 업로드

저장소 생성 후 나오는 화면에서 **uploading an existing file** 링크를 클릭하고,
`github-deploy` 폴더 안의 **모든 파일과 폴더**를 드래그하여 업로드합니다:

```
업로드할 파일/폴더 구조:
├── .github/
│   └── workflows/
│       └── sync.yml          ← 자동 실행 스케줄
├── scripts/
│   └── sync_sharepoint.py    ← SharePoint 동기화 스크립트
├── data/                     ← 출고 이력 데이터 (자동 생성)
├── index.html                ← 대시보드 메인 파일
├── requirements.txt          ← Python 패키지 목록
└── README.md                 ← 이 파일
```

> **참고**: `.github` 폴더는 숨김 폴더입니다. 윈도우 탐색기에서 **보기 → 숨긴 항목** 체크를 해야 보입니다.

**Commit changes** 를 클릭하여 업로드를 완료합니다.

### 3단계: GitHub Pages 활성화

1. 저장소 페이지에서 **Settings** (⚙️) 탭 클릭
2. 왼쪽 메뉴에서 **Pages** 클릭
3. **Source** 를 **Deploy from a branch** 로 선택
4. **Branch** 를 `main`, 폴더를 `/ (root)` 로 설정
5. **Save** 클릭

몇 분 후 대시보드가 아래 주소에서 접근 가능합니다:
```
https://[GitHub사용자이름].github.io/unicore-inventory/
```

### 4단계: SharePoint 연동 시크릿 설정 (자동 동기화용)

1. 저장소 **Settings** → 왼쪽 **Secrets and variables** → **Actions** 클릭
2. **New repository secret** 을 눌러 아래 3개를 추가합니다:

| Name | Value | 설명 |
|------|-------|------|
| `SHAREPOINT_USERNAME` | `your-email@uniandcore.com` | SharePoint 로그인 이메일 |
| `SHAREPOINT_PASSWORD` | `your-password` | SharePoint 비밀번호 |
| `SHAREPOINT_FILE_PATH` | `/sites/msteams_b2be8e/Shared Documents/General/12. SC/F'cst/UNI&CORE Inventory Report_(3월).xlsb` | 엑셀 파일 경로 |

> **보안 안내**: Secrets는 암호화되어 저장되며, 로그에도 표시되지 않습니다.

> **참고**: SharePoint에 MFA(다중인증)가 설정되어 있는 경우, 앱 비밀번호 생성이 필요할 수 있습니다. IT 담당자에게 문의하세요.

### 5단계: 동기화 테스트

1. 저장소 **Actions** 탭 클릭
2. 왼쪽에서 **SharePoint 데이터 동기화** 워크플로우 클릭
3. **Run workflow** → **Run workflow** 클릭
4. 실행이 완료되면 대시보드에 최신 데이터가 반영됩니다

---

## 운영 가이드

### 자동 실행 스케줄
- **매일 KST 23:50** 에 자동으로 SharePoint 데이터를 읽어 대시보드를 업데이트합니다
- Actions 탭에서 실행 이력을 확인할 수 있습니다

### 매월 파일 변경 시
SharePoint의 재고 보고서 파일이 매월 바뀝니다 (예: 3월→4월).
파일이 변경되면 **Settings → Secrets** 에서 `SHAREPOINT_FILE_PATH` 값을 새 파일 경로로 업데이트하세요.

예시:
```
/sites/msteams_b2be8e/Shared Documents/General/12. SC/F'cst/UNI&CORE Inventory Report_(4월).xlsb
```

### 수동 실행
긴급하게 데이터를 업데이트해야 할 때는 **Actions** 탭에서 **Run workflow** 를 클릭하면 즉시 실행됩니다.

### 문제 해결
- **Actions 실패 시**: Actions 탭에서 실패한 실행을 클릭하면 로그를 볼 수 있습니다
- **인증 오류**: SHAREPOINT_USERNAME, SHAREPOINT_PASSWORD 시크릿 값을 확인하세요
- **파일 경로 오류**: SHAREPOINT_FILE_PATH가 정확한지 확인하세요
- **대시보드가 안 보일 때**: Settings → Pages에서 배포 상태를 확인하세요
