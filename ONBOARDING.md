# MEDIPEEL 품질경영대시보드(QMS) 인수인계 문서

> 작성일: 2026-07-23 (최종 인수인계 시점 기준)
> 대상: 이 시스템을 새로 담당하게 될 실무자

---

## 0. 가장 먼저 읽어야 할 것 — 긴급 보안 조치 2건

담당을 넘겨받으면 **다른 작업보다 먼저** 아래 두 가지를 처리하세요. 둘 다 이미 발견된 채로 방치되어 있던 문제입니다.

### 0-1. GitHub 저장소가 Public 상태이고, 코드에 평문 비밀번호가 그대로 들어있습니다
- 저장소: `github.com/pyj090909-svg/medipeel-qms` — **Public**, 확인된 크기 122KB
- 루트 `index.html` 안의 `DEFAULT_ACCOUNTS` 배열에 전 직원의 로그인 아이디·비밀번호가 **평문**으로 박혀 있습니다 (예: `yjpark/yjpark`, `jack/jack`, `admin/admin1234` 등 30여 개 계정).
- 즉 인터넷에 있는 누구나 이 저장소를 열어보면 전 직원 로그인 정보를 그대로 볼 수 있는 상태입니다.
- **권장 조치**: ① 저장소를 Private으로 전환, ② 그것만으로는 이미 노출된 비밀번호가 무효화되지 않으므로 전 계정 비밀번호 재발급, ③ 장기적으로는 비밀번호를 클라이언트 코드에서 빼고 서버(Apps Script 등) 쪽에서 검증하는 구조로 개선 검토.

### 0-2. Git 원격 URL에 GitHub 토큰이 평문 저장돼 있습니다
- `.git/config`의 `origin` URL 형식이 `https://<계정>:<토큰>@github.com/...` 로, OAuth 토큰(`gho_...`)이 그대로 박혀 있습니다.
- 이 파일을 열어보는 사람 누구나 이 토큰으로 저장소에 쓰기 권한을 얻습니다.
- **권장 조치**: GitHub 설정에서 해당 토큰 즉시 폐기(revoke) → Git Credential Manager나 SSH 키 방식으로 재설정.

---

## 1. 이 시스템이 무엇인가

**MEDIPEEL 품질경영대시보드(QMS v1.0)** — 클레임 처리, 품질 지표, 인허가(RA) 진행 관리, 환경마크, 임상시험보고서(CSR) 등 여러 업무를 한 화면에서 볼 수 있게 모아둔 사내 대시보드 허브입니다.

- **배포 주소**: https://pyj090909-svg.github.io/medipeel-qms/
- **배포 방식**: GitHub Pages (저장소 `main` 브랜치, 루트 경로에서 정적 파일을 그대로 서빙)
- **로그인 필요**: 사내 계정 (아이디/비밀번호는 루트 `index.html`에 하드코딩)
- **기술 스택**: 순수 HTML/CSS/JS (프레임워크 없음) + 개별 모듈의 데이터 연동은 Google Apps Script(GAS) + Google Sheets
- **주의**: 로그인 로직은 100% 브라우저에서 실행되는 JS이며, 런타임에 GitHub을 호출하지 않습니다. 다만 GitHub Pages가 `Cache-Control: max-age=600`(10분)을 걸어두므로, 코드를 푸시한 뒤 최대 10분간은 이미 캐시된 브라우저에서 이전 버전이 보일 수 있습니다 (강력 새로고침으로 우회 가능).

---

## 2. 전체 구조 — 허브 + 개별 모듈

루트 `index.html`은 로그인 화면과 "메뉴" 역할만 합니다. 메뉴에서 각 카드를 클릭하면:
- 같은 저장소 안의 하위 폴더(`quality/`, `claim/` 등)로 이동하거나,
- **완전히 다른 도메인의 외부 사이트**(Netlify, Deskroom)로 새 탭이 열립니다.

**★ 가장 헷갈리는 부분(반드시 이해해야 함)**: 메뉴에 표시되는 번호(QMS_01, QMS_02 …)와, 로그인 권한을 판별하는 내부 코드(`data-module` 속성값)가 **서로 일치하지 않습니다**. 과거 세션에서 메뉴 순서를 재배치하면서 화면 번호만 바뀌었고, 내부 권한 코드는 원래 값을 그대로 유지했기 때문입니다. 계정에 권한을 부여할 때는 반드시 아래 표의 "내부 코드" 열을 기준으로 해야 합니다. 화면 번호로 착각하고 넣으면 완전히 다른 세션에 권한을 주는 사고가 납니다 (이번 인수인계 직전에 실제로 이 문제로 한 번 혼동이 있었습니다).

| 화면 표시 | 내부 코드(data-module) | 세션명 | 연결 위치 | 종류 |
|---|---|---|---|---|
| QMS_01 | `QMS_01` | RA_PMS 신규 등록 관리 | `https://medipeel-ra-progress.netlify.app/` | **외부 사이트**, 자체 로그인 |
| QMS_02 | `QMS_10` | RA_Ai 에이전트 (Deskroom) | `https://app.deskroom.so/auth/login` | **외부 SaaS** |
| QMS_03 | `QMS_02` | RA_인허가 완료 현황 | `https://medipeel-ra.netlify.app/` | **외부 사이트**, 로그인 없음 |
| QMS_04 | `QMS_03` | RA_비용 관리 | `Cost/index.html` | 저장소 내부 |
| QMS_05 | `QMS_09` | RA_대행사 계약 관리 | `contract/index.html` | 저장소 내부 |
| QMS_06 | `QMS_04` | 클레임 관련 | `claim/index.html` | 저장소 내부 |
| QMS_07 | `QMS_05` | 품질관리 일반 지표 | `quality/index.html` | 저장소 내부 |
| QMS_08 | `QMS_07` | 포장재 상용성 시험 관리 | `CT_Test/index.html` | 저장소 내부 |
| QMS_09 | `QMS_06` | 글로벌 환경마크 | `E_mark/index.html` | 저장소 내부 |
| QMS_10 | `QMS_11` | CSR (임상시험보고서) | `Clinical Test/index.html` | 저장소 내부 |
| QMS_11, QMS_12 | 없음 | 예비 세션 | `href="#"` (비활성) | 미사용 슬롯 |

계정별 접근 권한은 `index.html`의 `DEFAULT_ACCOUNTS` 배열, `allowedModules` 필드에 **내부 코드** 배열로 지정합니다. `allowedModules: null`이면 전체 허용입니다.

---

## 3. 계정 시스템 상세

- 계정 데이터는 `index.html` 안 `DEFAULT_ACCOUNTS` (하드코딩 기본값) + 브라우저 `localStorage` (관리자가 화면에서 추가/수정한 내용, 브라우저별로 따로 저장됨) 두 곳에 있습니다.
- 페이지 로드 시 `loadAccounts()`가 `localStorage`를 `DEFAULT_ACCOUNTS` 기준으로 동기화합니다 — 코드에 없는 계정은 유지, 코드에 있는 계정은 `allowedModules`/`group`/`role`을 코드 값으로 항상 덮어씁니다. **즉 계정 권한을 바꾸려면 코드(`DEFAULT_ACCOUNTS`)를 고쳐서 배포해야 하며, 로그인 화면의 "계정관리" 메뉴(관리자 전용)로 추가한 계정은 새 계정 추가에는 쓰이지만 기존 코드 계정의 권한을 영구히 바꾸지는 못합니다.**
- 그룹(`GROUPS` 배열): 품질구매팀, 경영지원본부, 상품본부, 상품기획팀, 디자인팀, 스킨이데아, 윈게이트코리아, 해외영업팀, 영업지원팀, MSPI
- 최근 추가된 계정: `salessupport` / `salessupport` (영업지원팀 공용계정, 내부 코드 `QMS_01`·`QMS_02`, 즉 화면상 QMS_01·QMS_03 두 세션 접근 가능)

### 미해결 이슈 — RA_PMS(QMS_01) 세션은 별도 로그인이 필요함
`QMS_01`(화면 번호)이 연결되는 `https://medipeel-ra-progress.netlify.app/`는 **QMS 허브와 완전히 독립된 자체 로그인 시스템**을 가지고 있습니다 (아이디/비밀번호 입력창이 별도로 있고, 등록되지 않은 계정은 "등록되지 않은 아이디입니다"라는 문구가 뜹니다). QMS 허브의 `DEFAULT_ACCOUNTS`에 계정을 추가해도 이 사이트에는 반영되지 않습니다.
- 이 사이트의 소스 코드를 로컬 다운로드 폴더 전체에서 찾아봤으나 **발견하지 못했습니다.**
- 유력한 후보: 이 저장소의 `RA Task process/` 폴더 — `index.html` 제목이 "QMS_01 RA 진행 관리"이고 Code.gs가 참조하는 시트 구조(DATA/PROGRESS/SCHEDULE/BLOCKER/설정, 시트ID `1axlGq2PbeumWqunb3EKAMFBWHBhEbXBBrYUq-uNccDI`)가 개념적으로 일치합니다. 다만 이 폴더의 `index.html`에는 로그인 화면 코드가 없어서, 실제 배포된 사이트와 100% 같은 파일인지는 확인하지 못했습니다 (배포 후 별도로 로그인 기능이 추가됐을 가능성).
- **다음 담당자가 할 일**: Netlify 계정(또는 이 사이트를 실제로 배포한 사람)에 접근해서 `medipeel-ra-progress` 사이트의 실제 연결된 저장소/소스를 확인하고, 거기에 필요한 계정을 추가해야 합니다.

---

## 4. 저장소 내부 모듈 상세 (Google Apps Script 연동)

각 폴더는 `index.html`(화면) + `Code.gs`(Google Apps Script 백엔드, 웹앱으로 배포되어 있음)로 구성됩니다. GAS 코드는 이 저장소에 **원본 백업용**으로만 있고, 실제 실행되는 코드는 각 담당 Google 계정의 Apps Script 프로젝트에 배포돼 있습니다 (`Code.gs` 상단 주석에 배포 방법이 적혀 있음).

| 폴더 | 화면 번호 | Google Sheet ID | 비고 |
|---|---|---|---|
| `claim/` | QMS_06 | `1Pv8mcM80TuVeY_mJa1Ivwr_BCpHaC--EedR8auib-aw` (시트1) | 클레임 접수/처리. 최근 로컬 수정사항 미커밋 상태 |
| `quality/` | QMS_07 | (SETUP_GUIDE.gs만 있고 Code.gs는 없음, 시트명 `상품목록`) | 품질 KPI. 최근 로컬 수정사항 미커밋 상태 |
| `Cost/` | QMS_04 | `1nW1VomsorXvXaXenqV51qqQ_XXnalNCKbAfdhve8TA0` (RA_비용관리_DATA/설정) | RA 대행사 비용 관리 |
| `contract/` | QMS_05 | `1_y9epi5qkhYeg5-uPxcYFMSOGjRu1Vl7u9UWg06iFG4` (gid 1492902388) | 대행사 계약 관리 |
| `CT_Test/` | QMS_08 | `1trBtUq6kSTHgeG_7wTSwnklc3br3DCp9HowtfNvOuBY` | 포장재 상용성 시험 |
| `E_mark/` | QMS_09 | (Code.gs에 SHEET_ID 명시 없음 — 컨테이너 바인딩 시트 사용 추정) | 환경마크 서류 링크 관리 |
| `Clinical Test/` | QMS_10 | `1sNDUhQusj0Kmjnw2jBMk2tXlQva4ww95_h319SRk8Is` | GAS 없이 Google Sheets CSV를 직접 fetch (`gviz/tq` 엔드포인트). 기획서 문서(`스킨이데아_보고서_대시보드_기획서_v1.md`) 포함 |
| `RA Task process/` | **메뉴 미연결** | `1axlGq2PbeumWqunb3EKAMFBWHBhEbXBBrYUq-uNccDI` | 위 3절 참고 — QMS_01 관련 추정 프로토타입, 현재 허브 메뉴에서 링크가 빠져 있음 |
| `artwork/` | **메뉴 미연결** | `1YDB57GE3gG-IxhOxxgh4EacQKPU8c82NeKjjS_Ub3Hw` (문안검토로그) | RA-AMS 아트웍 스크리닝. 예전엔 QMS_02였으나 "사용 안 함" 요청으로 예비 세션 처리, 폴더는 남아있음 |

---

## 5. 저장소 밖의 외부 사이트 소스 위치

- `https://medipeel-ra.netlify.app/` (화면 QMS_03) 와 `https://medipeel-ra-stock.netlify.app/` (재고 현황, QMS_03 화면 안의 숨겨진 버튼) 의 소스는 **이 저장소 밖**, `C:\Users\user\Downloads\1_RA.i platform.zip` 안에 있습니다. 압축을 풀면:
  - `1_RA.i platform/플랫폼 포털/index/index.html` → medipeel-ra.netlify.app 추정
  - `1_RA.i platform/재고정보 연동/index/index.html` → medipeel-ra-stock.netlify.app 추정
  - ⚠ 이 zip은 macOS에서 만들어져 한글 폴더명이 NFD 방식으로 인코딩돼 있어 Windows 도구에서 파일을 못 찾는 경우가 있습니다. 폴더명을 한 번 다른 이름으로 rename한 뒤 다시 원래 이름으로 rename하면 정상화됩니다 (Windows가 재생성하면서 NFC로 바뀜).
- `https://medipeel-ra-progress.netlify.app/` (화면 QMS_01) 소스 위치는 **미확인** — 4절 참고.
- `https://app.deskroom.so/` (화면 QMS_02) 는 Deskroom이라는 외부 SaaS 제품으로, 자체 소스가 없습니다 (제3자 서비스).

---

## 6. Git 상태 — 인수인계 시점 기준

```
커밋됨, 라이브 반영 완료: index.html (메뉴 리디자인, 예비 세션 정리, salessupport 계정 추가 — 커밋 fd1691f)

미커밋 (다음 담당자가 검토 필요):
  M claim/Code.gs
  M claim/index.html
  M quality/index.html

한 번도 커밋 안 된 폴더 (untracked):
  CT_Test/, Clinical Test/, Cost/, E_mark/, RA Task process/, artwork/, contract/,
  images/deskroom_rogo.jpg, quality/SETUP_GUIDE.gs
```
이 미추적 폴더들은 실제로는 라이브 사이트에서 정상적으로 쓰이고 있는 파일들입니다 (허브 메뉴가 이 폴더들의 `index.html`을 직접 링크). GitHub Pages는 로컬 파일 시스템을 그대로 배포하는 게 아니라 **Git에 커밋된 내용만** 배포하므로, 만약 이 폴더들이 실제로 한 번도 push된 적이 없다면 배포된 사이트에는 어떻게 존재하는지 확인이 필요합니다 — 가능성은:
1. 과거 이 폴더들을 커밋했다가 이후 `.gitignore`나 다른 이유로 추적 해제됐고, 배포본에는 그 시점 커밋이 남아있음, 또는
2. 실제 배포본과 로컬 폴더 내용이 100% 동일하지 않을 수 있음

**다음 담당자가 반드시 확인할 것**: `git log --all -- "claim/index.html"` 등으로 각 폴더가 실제로 커밋된 적이 있는지, 그리고 로컬 파일과 배포된 라이브 파일이 내용상 일치하는지 diff 확인.

---

## 7. 인수인계 체크리스트

- [ ] **(긴급)** GitHub 저장소 Private 전환 검토
- [ ] **(긴급)** 노출된 Git 토큰(`gho_...`) 폐기 및 재설정
- [ ] **(긴급)** 이미 노출된 전 직원 비밀번호 재발급 검토
- [ ] `medipeel-ra-progress.netlify.app` (QMS_01) 실제 소스/Netlify 계정 위치 확인 → salessupport 등 필요 계정 추가
- [ ] 미커밋 상태인 `claim/Code.gs`, `claim/index.html`, `quality/index.html` 검토 후 커밋 여부 결정
- [ ] untracked 폴더들의 실제 커밋 여부와 배포본 일치 여부 확인 (6절 참고)
- [ ] `RA Task process/`, `artwork/` 폴더를 계속 보관할지, 완전히 정리할지 결정 (현재 허브 메뉴에서 링크 없음)
- [ ] 계정 권한 변경 시 반드시 2절 표의 "내부 코드" 열 기준으로 작업 (화면 번호 기준으로 넣으면 오작동)
