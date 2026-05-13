# code.gs — Apps Script 관리 문서

> 미래오토메이션(주) 현금흐름 관리 앱 백엔드  
> 스프레드시트: `1VxYrCD3eZr5PpTORFPCEQPfWM5QSr-tNFNnc_W1C5qM`

---

## Apps Script 배포 방법

```
1. script.google.com → 해당 프로젝트 열기
2. 기존 코드 전체 선택(Ctrl+A) → 삭제
3. code.gs 내용 전체 복사 → 붙여넣기
4. 저장(Ctrl+S)
5. 배포 → 배포 관리 → ✏️ 버전 수정 → [새 버전] → 배포
```

> ⚠️ **저장 후 반드시 "새 버전"으로 재배포해야 변경사항이 반영됩니다.**  
> 단순 저장만 하면 기존 배포 URL은 이전 버전을 계속 사용합니다.

---

## 시트명 상수 목록

| 상수명 | 시트명 | 용도 |
|--------|--------|------|
| `PAYABLES_SHEET` | `미지급_raw` | 미지급 원본 데이터 |
| `RECEIVABLES_SHEET` | `raw` | 미수금 원본 데이터 |
| `MANAGER_SHEET` | `담당자` | 담당자 마스터 |
| `PLAN_SHEET` | `결제계획` | 결제계획 저장 |
| `MASTER_SHEET` | `업체마스터` | 업체 마스터 |
| `HISTORY_SHEET` | `결제이력` | 결제이력 저장 |
| `UPDATE_HISTORY_SHEET` | `업데이트이력` | 변경이력 로그 |
| `TAX_INVOICE_SHEET` | `세금계산서_raw` | 세금계산서 (대사용) |
| `LEDGER_SALES_SHEET` | `계정별원장_매출_raw` | 계정원장 매출 (대사용) |
| `LEDGER_BUY_SHEET` | `계정별원장_매입_raw` | 계정원장 매입 (대사용) |
| `LEDGER_PAY_SHEET` | `계정별원장_미지급_raw` | 계정원장 미지급 (대사용) |
| `DAILY_SALES_SHEET` | `영업현황_raw` | 영업현황 일별 (대사용) |
| `BIZ_DIVISION_SHEET` | `사업부문마스터` | 사업부문 마스터 |
| `FIXED_SHEET` | `고정지출` | 고정지출 |
| `PNL_SHEET` | `경영손익_data` | 경영손익 월별 데이터 |

---

## API action 목록

### GET (읽기)

| action | 반환 | 설명 |
|--------|------|------|
| *(없음)* | `{ data: [...] }` | 미지급 raw 데이터 (기본값) |
| `getPaymentPlans` | `{ rows: [...] }` | 결제계획 |
| `getVendorMaster` | `{ rows: [...] }` | 업체마스터 |
| `getPaymentHistory` | `{ rows: [...] }` | 결제이력 |
| `getReceivables` | `{ rows: [...] }` | 미수금 raw |
| `getManagerMaster` | `{ rows: [...] }` | 담당자 마스터 |
| `getTaxInvoices` | `{ rows: [...] }` | 세금계산서 |
| `getLedgerSales` | `{ rows: [...] }` | 계정원장 매출 |
| `getLedgerPurchase` | `{ rows: [...] }` | 계정원장 매입 |
| `getLedgerPayable` | `{ rows: [...] }` | 계정원장 미지급 |
| `getDailySales` | `{ rows: [...] }` | 영업현황 |
| `getBizDivision` | `{ rows: [...] }` | 사업부문 마스터 |
| `getFixed` | `{ rows: [...] }` | 고정지출 |
| `getPnlData` | `{ rows: [...] }` | 경영손익 월별 데이터 |
| `getDaesaAll` | `{ taxInvoices, ledgerSales, ledgerPurchase, ledgerPayable, dailySales, bizDivision }` | 대사 데이터 6종 일괄 (스프레드시트 1회 오픈) |
| `getMautoData` | `{ data: {...} }` | 엠오토 JSON (`엠오토_json` 시트 A1) |
| `getAvailableFundsJson` | `{ updatedAt, data }` | 가용자금 JSON (`가용자금_json` 시트) |

### POST (쓰기)

| action | 필수 필드 | 설명 |
|--------|-----------|------|
| `appendPaymentPlans` | `rows[]` | 결제계획 추가 |
| `appendPaymentHistory` | `rows[]` | 결제이력 추가 |
| `appendUpdateHistory` | `rows[]` | 업데이트이력 추가 |
| `upsertVendorMaster` | `rows[]` | 업체마스터 저장 (복합키 upsert) |
| `upsertManagerMaster` | `rows[]` | 담당자마스터 저장 |
| `upsertTaxInvoices` | `rows[]` | 세금계산서 저장 (`_row_key` 기준) |
| `upsertLedger` | `rows[]`, `ledgerType` | 계정원장 저장 (`ledgerType`: `매출`/`매입`/`미지급`) |
| `upsertDailySales` | `rows[]` | 영업현황 저장 |
| `upsertBizDivision` | `rows[]` | 사업부문마스터 저장 |
| `savePnlData` | `rows[]` 또는 `row{}` | 경영손익 저장 (`_key` 기준 upsert) |
| `saveMautoData` | `data{}` | 엠오토 JSON 저장 |
| `upsertAvailableFunds` | `updatedAt`, `data{}` | 가용자금 JSON 저장 |
| `sendReceivableEmails` | *(params)* | 미수금 이메일 발송 |
| `sendRawDiffEmail` | `diff[]`, `recipients[]` | 미지급 변경 감지 이메일 |
| `sendPaymentWarningEmail` | `warnings[]`, `recipients[]` | 결제 경고 이메일 (은행정보 누락) |

---

## 헬퍼 함수 설명

### `getSheetRows(sheetName, ss?)`
시트 전체를 객체 배열로 반환.  
`ss`를 전달하면 `openById` 재호출 생략 → `getDaesaAll` 같은 일괄 요청에서 속도 향상.  
헤더 행은 최대 10행 내에서 키워드(`거래처코드`, `year` 등)로 자동 감지.

### `upsertRowsByKey(sheetName, keyField, rows)`
`keyField` 기준으로 기존 행 업데이트, 신규 행 추가.  
전체 데이터를 메모리에 올린 뒤 `setValues` 2회(기존/신규)로 처리해 API 호출 최소화.  
키 비교 시 앞자리 0 및 아포스트로피 제거 (`00101` = `101` = `'00101`).

### `upsertVendorMasterRows(sheetName, newRows)`
업체마스터 전용 upsert.  
복합키: `거래처코드_norm` → `사업자번호` → `거래처명` 순서로 fallback.  
`거래처코드_norm`, `거래처코드_raw`, `vendor_id`, `사업자번호`, `계좌번호` 컬럼에  
`setNumberFormat("@")` 적용 → Google Sheets 자동 숫자 변환 방지 (앞자리 0 보존).

### `appendRows(sheetName, rows)`
시트 맨 아래에 행 추가. 중복 체크 없음.

---

## 변경 이력

### 2026-05-13 — code.gs 전체 정리 (현재)

**수정 내용:**
- `doGet`: `GET_SHEET_MAP` 룩업 테이블로 action 라우팅 통합 (`if` 13개 → `map` 1개)
- `doPost`: 상단에서 `rows` 변수 1회 선언 후 전체 재사용 (중복 제거)
- `upsertVendorMasterRows`: `applyTextFormat` 내부 함수로 텍스트 포맷 적용 코드 중복 제거
- `getSheetRows`: `HEADER_KEYWORDS` 배열 + `some()` 으로 가독성 개선
- `wrapEmail` 파라미터명 `body` → `innerHtml` (외부 `body` 변수 shadowing 해소)
- `appendRows`: 중간 변수 제거, 인라인 map 으로 간결화
- 섹션 구분선 및 JSDoc 주석 정비

---

### 2026-05-13 — `getPnlData` / `savePnlData` 추가

**수정 내용:**
- `const PNL_SHEET = "경영손익_data"` 상수 추가
- `doGet`에 `getPnlData` action 추가 → `경영손익_data` 시트 반환
- `doPost`에 `savePnlData` action 추가 → `rows[]` 배열 또는 단건 `row{}` 처리
  - `_key` 컬럼 기준 upsert (`${year}_${month}` 형식)
  - pnl.html 서명 처리 및 app.js `[☁️ 동기화]` 버튼 양쪽에서 호출

**이 버전 전 code.gs에 있던 문제 (주의):**
```
// ❌ 잘못된 예 — 이런 중복이 있으면 SyntaxError 발생
const PNL_SHEET = "월간손익";     // 구버전 (삭제 필요)
const PNL_SHEET = "경영손익_data"; // 신버전
```
Apps Script 편집기에 이 두 줄이 동시에 존재하면 `SyntaxError: Identifier 'PNL_SHEET' has already been declared` 오류로 모든 API 호출이 실패합니다.

---

### 2026-05-12 — `getDaesaAll` 추가 (대사 탭 속도 개선)

**수정 내용:**
- `getSheetRows(sheetName, ss?)` — `ss` 파라미터 추가 (스프레드시트 객체 재사용)
- `getDaesaAll` action 추가 — `SpreadsheetApp.openById` 1회 호출로 6개 시트 일괄 반환
- 효과: HTTP 왕복 6회 → 1회 (체감 2~5배 빠름)

---

### 2026-05-12 — 업체마스터 앞자리 0 소실 버그 수정

**수정 내용 (`upsertVendorMasterRows`):**  
`거래처코드_norm`, `거래처코드_raw`, `vendor_id`, `사업자번호`, `계좌번호` 컬럼에  
`setNumberFormat("@")` 적용 추가.

**원인:** Google Sheets `setValues()`가 `"04159"` 같은 문자열을 숫자 `4159`로 자동 변환.  
`setNumberFormat("@")`로 텍스트 포맷 강제 지정하면 앞자리 0이 유지됨.

---

## 경영손익 데이터 구조 (`경영손익_data` 시트)

| 컬럼 | 타입 | 설명 |
|------|------|------|
| `_key` | string | 기본키 (`2026_03` 형식) |
| `year` | number | 연도 |
| `month` | number | 월 |
| `revenue` | number | 매출액 |
| `targetRevenue` | number | 목표매출액 |
| `cogs` | number | 상품매출원가 |
| `mfg` | number | 당기총제조비용 |
| `sga` | number | 판매관리비 |
| `interest` | number | 영업외비용 |
| `approvalStatus` | string | 결재상태 (`draft` / `기안` / `합의1` / `합의2` / `결재완료`) |
| `draftDate` | string | 기안일 (`yyyy-MM-dd`) |
| `agree1Date` | string | 합의①일 |
| `agree2Date` | string | 합의②일 |
| `ceoDate` | string | 최종결재일 |
| `docNo` | string | 문서번호 |
| `ceoComment` | string | 대표이사 의견 |

---

## 인증 토큰

모든 요청에 `token` 파라미터(GET) 또는 `body.token`(POST) 필요.

```
토큰값: miraeautomation2026
localStorage 키: receivable-payable-webapp.api-token.v1
```
