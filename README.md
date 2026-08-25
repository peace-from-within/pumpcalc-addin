# PUMPCALC Custom Functions Add-in

Pump Hydraulics 계산서(Module_For_Equation.bas)의 21개 계산 함수를 Excel Custom
Functions Add-in으로 이식한 것. 웹 배포본 워크북에는 계산 로직이 전혀 포함되지
않고, 셀 수식에는 `=PUMPCALC.CV_PRESS_DROP(...)` 같은 함수 호출만 남는다. 실제
계산 코드(functions.js)는 이 폴더를 호스팅하는 서버/스토리지에만 존재하며
발주처에게 전달되는 xlsm/xlsx 파일 안에는 들어가지 않는다.

## 웹 배포용 워크북 (`Pump Hydraulics_WebDeploy_v1.xlsx`)

`Pump Hydraulics_Case 반영_개선.xlsm`을 기준으로 `build_scripts/build_webdeploy3.py`
+ `build_webdeploy4_xlsx.py`를 순서대로 실행해 만든 최종 산출물. 처리 내역:

- Line DP / Pump Calculation 시트의 21개 UDF 호출(총 612건) 전부 `PUMPCALC.*`로 치환
- Fixed DP(Total Loss) 모드로 고정 — AB22 토글 드롭다운·PumpDiff_fixed 매크로 제거,
  기존에 이미 Fixed DP 상태로 저장돼 있던 AL41/AN41/AP41(`=Total Loss - Pump DP`) 수식 그대로 유지
- Path 관리 기능 제거: Sketch 시트의 Path 드롭다운(AK3:AK11) 비움, Path 관련
  버튼/컨트롤 삭제. **Paths 시트 자체는 완전 삭제 대신 `state="veryHidden"`으로
  처리** — OOXML 구조상 시트를 완전히 삭제하면 Excel/LibreOffice가 파일을 손상된
  것으로 판단해 열리지 않는 문제가 실측 확인되어(단일 `<sheet>` 태그 제거만으로도
  재현), 데이터/구조를 그대로 둔 채 사용자에게만 안 보이게 하는 방식으로 우회함
- Multi Case 드롭다운(AB39), Insert/Delete Stream 버튼, Show_EQUIVL 등 나머지
  매크로 전용 Form Control 19개 전부 삭제 (ctrlProp 파일까지 완전 제거)
- vbaProject.bin 삭제, `.xlsm` → `.xlsx`로 전환, `fullCalcOnLoad` 설정

LibreOffice로 열림/재계산 검증 완료(구조 오류 없음). 단, Add-in 없이 계산하면
`PUMPCALC.*` 호출부가 `#NAME?`로 뜨는 게 정상이며 — 이는 Add-in을 실제로 설치한
Excel 환경에서만 정상 계산된다. 21개 함수 자체는 별도로 VBA 원본 대비 496개
테스트 케이스 전수 검증(496/496 통과) 완료된 상태.

## 파일 구성

- `manifest.xml` — Office Add-in 매니페스트. 발주처 M365 관리자가 Integrated
  Apps에 등록할 때 쓰는 파일. **배포 전에 `PLACEHOLDER-*` 값들을 실제 값으로
  교체해야 한다** (아래 "배포 전 체크리스트" 참고).
- `functions.js` — 22개 Custom Function 구현체 (Office.js `CustomFunctions.associate`로 등록).
  Pipe_ID 룩업 테이블(35행×11스케줄)이 인라인 JSON으로 내장되어 있어 별도 파일
  의존성이 없다.
- `functions.json` — Custom Functions 메타데이터(함수명/파라미터/설명). Excel이
  자동완성·툴팁에 사용.
- `functions.html` — Custom Functions 런타임을 로드하는 빈 페이지. 사용자에게
  보이지 않음.
- `src/core.js` — 계산 로직 원본(Node.js 모듈 형태, `require`로 단위테스트 가능).
  **로직을 수정할 때는 이 파일을 고치고, 다시 functions.js에 조립해 넣어야 한다**
  (현재는 수작업 조립 — 아래 "로직 수정 시" 참고).
- `src/pipeIdTable.json` — Pipe_ID 시트에서 추출한 룩업 테이블 원본.
- `src/verify.js`, `src/verify2.js` — VBA 원본 대비 검증 스크립트 (아래 참고).

## 검증 내역

VBA 원본(Module_For_Equation.bas)을 그대로 옮긴 Python ground-truth(`vba_port_faithful.py`,
`extra_vba_port.py`)와 JS 결과를 무작위 입력값으로 대조했다.

| 대상 | 테스트 수 | 결과 |
|---|---|---|
| Calc_RA, VaporVolFrac, AvgDensity, DPHomogeneous, DPDukler, InplaceDensity, TwoPhaseVelocity, BakerXval, BakerYval, Baker, GriffithWallisXval, GriffithWallisYval, Griffith, CalcPres, CalcHead, CalcPumpEff, EstNPSHR, CVPressDrop | 396 | 396 통과 |
| CalcPumpEffAnother, SelMotorPower, EqLenFactor | 100 | 100 통과 |
| Pipe_ID (실제 시트 셀 대조) | 15 (+5 빈 셀) | 15 통과 |

**중요한 발견 - VBA의 `Max(a,b,c)` 버그 재현**: 원본 VBA는
`If b>Max Then Max=b ElseIf c>Max Then Max=c` 구조라서 실제로는 3개 값의
진짜 최댓값이 아니다 (예: a=1,b=2,c=5 → 진짜 최댓값은 5지만 VBA는 2를 반환).
`CVPressDrop`의 Lummus Method 분기가 이 함수를 쓰므로, `core.js`의 `Max_VBA`도
이 버그를 의도적으로 그대로 재현했다. (과거 세션에서 만든 `vba_port.py`는 이
버그를 놓치고 진짜 `max()`를 썼던 것을 이번에 발견해 `vba_port_faithful.py`로
수정 후 재검증함.)

또 하나: DPDukler 최종항이 처음엔 `Wtot/3600`으로 잘못 구현되어 있었는데
(`GT/3600`이어야 함, GT=Wtot/AREA), 검증 과정에서 20개 중 0개 통과로 바로 잡혔다.

## 배포 전 체크리스트

1. `manifest.xml`의 `PLACEHOLDER-GENERATE-NEW-GUID`를 실제 GUID로 교체
   (`python3 -c "import uuid; print(uuid.uuid4())"` 등으로 생성)
2. `PLACEHOLDER-COMPANY-NAME`, `PLACEHOLDER-HOSTING-DOMAIN`을 실제 값으로 교체
   (functions.js/functions.html/functions.json을 HTTPS로 서빙할 도메인 필요 —
   Azure Static Web Apps, GitHub Pages 등 정적 호스팅으로 충분)
3. 아이콘 파일(`assets/icon-32.png`, `icon-80.png`) 준비 후 업로드
4. 호스팅 후 `manifest.xml`을 발주처 M365 관리자에게 전달, Admin center →
   Settings → Integrated Apps → Upload custom apps 에서 매니페스트 URL 등록
   요청 (자세한 배경은 이전 세션 논의 참고)
5. 워크북 셀 수식은 `=PUMPCALC.CV_PRESS_DROP(...)`처럼 네임스페이스를 붙여 작성

## 로직 수정 시

지금은 빌드 파이프라인 없이 `src/core.js`의 본문을 `functions.js` 상단에 수작업
조립해 넣은 상태다(pipeIdTable도 JSON 상수로 인라인됨). 계산 로직을 고칠 일이
생기면:

1. `src/core.js`를 수정
2. `node src/verify.js`, `node src/verify2.js`로 회귀 테스트
3. `functions.js`의 해당 함수 본문(주석 없는 순수 로직 부분)을 동일하게 갱신
   (또는 이번에 썼던 조립 스크립트 방식을 재사용해 자동화 가능)

## 함수 목록 (22개)

CALC_RA, PIPE_ID, VAPOR_VOL_FRAC, AVG_DENSITY, DP_HOMOGENEOUS, DP_DUKLER,
INPLACE_DENSITY, TWO_PHASE_VELOCITY, BAKER_XVAL, BAKER_YVAL, BAKER,
GRIFFITH_WALLIS_XVAL, GRIFFITH_WALLIS_YVAL, GRIFFITH, CALC_PRES, CALC_HEAD,
CALC_PUMP_EFF, CALC_PUMP_EFF_ALT, SEL_MOTOR_POWER, EST_NPSHR, EQ_LEN_FACTOR,
CV_PRESS_DROP
