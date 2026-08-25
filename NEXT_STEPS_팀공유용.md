# 팀 동료 공유 전, 지금 당장 해야 할 일

발주처 배포는 나중 얘기고, 우선 팀 동료 한 명한테 "이 웹 배포 버전 계산이 맞게 나오는지" 테스트 받는 게 목표. 그러려면 Add-in이 실제로 어딘가에서 실행되고 있어야 함(로컬 파일만으로는 PUMPCALC 함수가 인식 안 돼서 전부 #NAME? 뜸).

## 1. Add-in을 실행할 곳 정하기

Custom Functions는 `functions.html`/`functions.js`/`functions.json`을 HTTPS로 서빙해야 동작한다. 사내 테스트 단계라 가장 간단한 선택지부터:

- **GitHub Pages / Netlify / Vercel 같은 무료 정적 호스팅** — `pumpcalc-addin` 폴더를 통째로 올리면 끝. 가장 빠름, 팀 테스트 용도로 충분.
- **회사 사내 웹서버(있다면)** — 이미 쓰는 정적 파일 서버가 있으면 거기에 폴더 하나 만들어서 올리기만 하면 됨.
- **로컬 PC + ngrok/localtunnel** — 지금 당장 뭘 설치하기 애매하면, 로컬에서 `python -m http.server` 같은 걸로 띄우고 ngrok으로 임시 HTTPS URL 만드는 것도 됨. 단 정배님 PC가 켜져 있어야만 동료가 파일을 열 수 있음 (진짜 임시용).

→ 오늘 팀 동료한테 공유하는 목적이면 **GitHub Pages나 Netlify**가 제일 무난하다. 계정만 있으면 5분 안에 URL 나옴.

## 2. manifest.xml의 PLACEHOLDER 채우기

`pumpcalc-addin/manifest.xml` 열어서 아래 4곳을 실제 값으로 바꿔야 함:

- `PLACEHOLDER-GENERATE-NEW-GUID` → 새 GUID 하나 생성해서 교체
  (터미널 있으면 `python3 -c "import uuid; print(uuid.uuid4())"`, 없으면 아무 온라인 GUID 생성기 사용)
- `PLACEHOLDER-COMPANY-NAME` → 팀/회사 이름
- `PLACEHOLDER-HOSTING-DOMAIN` → 1번에서 정한 호스팅 URL (예: `내계정.github.io/pumpcalc-addin`)
  - 이 도메인이 파일 4곳에서 등장함: `IconUrl`, `HighResolutionIconUrl`, `SupportUrl`, `AppDomains`, `SourceLocation`, `Functions.Script.Url`, `Functions.Page.Url`, `Functions.Metadata.Url` — 전부 같은 도메인으로 통일

아이콘(`icon-32.png`, `icon-80.png`)은 테스트 단계에선 대충 아무 이미지나 넣어도 무방. 없으면 그 두 줄은 지워도 Add-in 자체 동작에는 지장 없음(사이드로드 시 아이콘만 깨져 보임).

## 3. Add-in 사이드로드해서 동료 PC에서 열기

발주처처럼 M365 관리자 승인 절차 없이, 팀 내부 테스트는 **"내 추가 기능" 수동 업로드**로 바로 확인 가능:

1. 동료가 데스크톱 Excel(또는 Excel Online)에서 `Pump Hydraulics_WebDeploy_v1.xlsx` 파일을 연다
2. 리본 메뉴 **삽입(Insert) → 추가 기능(Add-ins) → 내 추가 기능(My Add-ins) → 업로드한 추가 기능(Upload My Add-in)**
3. 2번에서 완성한 `manifest.xml` 파일을 업로드
4. Add-in이 로드되면 셀의 `=PUMPCALC.CALC_HEAD(...)` 같은 수식이 정상 계산되는지 확인

이 방식은 M365 관리자 승인이 필요 없고, 그 파일을 연 사람 각자가 한 번씩 업로드해줘야 하는 방식이라 팀 내부 테스트에 딱 맞음. (발주처 대상 정식 배포는 이후 단계에서 Integrated Apps로 전환)

## 4. 동료한테 넘길 것

- `Pump Hydraulics_WebDeploy_v1.xlsx` (계산 시트)
- 채운 `manifest.xml` (Add-in 등록용)
- 위 "3. Add-in 사이드로드" 방법 안내 (또는 이 문서 자체를 같이 전달)

## 순서 요약

1. 정적 호스팅 하나 정하기 (GitHub Pages 추천)
2. `pumpcalc-addin` 폴더 통째로 업로드
3. `manifest.xml`의 PLACEHOLDER 4곳을 실제 URL/이름으로 수정
4. xlsx 파일 + manifest.xml을 동료에게 전달, 사이드로드 방법 안내
5. 동료가 열어서 계산값이 정상적으로 나오는지 확인 → 문제 있으면 알려달라고 요청
