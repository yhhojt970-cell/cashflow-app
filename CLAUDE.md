# 현금흐름 관리 앱 — CLAUDE.md

## 프로젝트 개요

미래오토메이션(주) 내부 현금흐름 관리 웹앱 (정적 HTML).  
Google Sheets 연동 + 로컬 Excel 붙여넣기 방식.

- **GitHub**: `https://github.com/yhhojt970-cell/cashflow-app`
- **배포**: GitHub Pages → `https://yhhojt970-cell.github.io/cashflow-app/`
- **파일 구조**: `index.html` + `app.js` + `style.css` (빌드 없음, 순수 HTML/JS/CSS)

---

## 탭 구성

| 탭 ID | 탭명 | 설명 |
|-------|------|------|
| `home` | 홈 | 금융 통합 대시보드 |
| `funds` | 가용자금 | **Excel 붙여넣기 입력** (localStorage 저장) |
| `receivables` | 미수금 | Google Sheets `raw` 시트에서 로드 |
| `payables` | 미지급 | Google Sheets `미지급_raw` 시트에서 로드 |
| `fixed` | 고정지출 | Google Sheets `고정지출` 시트에서 로드 |
| `daesa` | 대사 | 입출금 매칭 |
| `mauto` | 엠오토 | 엠오토 전용 현금흐름 (가용자금·미수금·미지급·고정지출 붙여넣기, 구글시트 원격 저장/로드) |

---

## Google Sheets 연동 구조

```js
const SHEET_SPREADSHEET_ID = "1VxYrCD3eZr5PpTORFPCEQPfWM5QSr-tNFNnc_W1C5qM";
const SHEET_APP_SCRIPT_URL = "https://script.google.com/macros/s/..."; // Apps Script WebApp URL
```

### Apps Script action 목록

| action | 방향 | 설명 |
|--------|------|------|
| (없음, GET) | 읽기 | 미지급 raw 데이터 (`미지급_raw` 시트) |
| `getReceivables` | 읽기 | 미수금 raw (`raw` 시트) |
| `getVendorMaster` | 읽기 | 업체마스터 |
| `getManagerMaster` | 읽기 | 담당자 마스터 |
| `getPaymentPlans` | 읽기 | 결제계획 |
| `getPaymentHistory` | 읽기 | 결제이력 |
| `getFixed` | 읽기 | 고정지출 |
| `getAvailableFunds` | 읽기 | 가용자금 (gviz fallback 우선) |
| `getTaxInvoices` | 읽기 | 세금계산서 (대사용) |
| `getLedgerSales` | 읽기 | 계정별원장 — 외상매출금 (대사용) |
| `getLedgerPurchase` | 읽기 | 계정별원장 — 외상매입금 (대사용) |
| `getLedgerPayable` | 읽기 | 계정별원장 — 미지급금 (대사용) |
| `getDailySales` | 읽기 | 영업현황 일별 (대사용) |
| `getBizDivision` | 읽기 | 사업부문 마스터 (대사용) |
| `getDaesaAll` | 읽기 | 대사 데이터 6종 일괄 반환 (스프레드시트 1회 오픈, 속도 최적화) |
| `getMautoData` | 읽기 | 엠오토 데이터 (`엠오토_json` 시트 A1 셀 JSON) |
| `getAvailableFundsJson` | 읽기 | 가용자금 JSON (`가용자금_json` 시트, `{updatedAt, data}`) |
| `upsertVendorMaster` | 쓰기 | 업체마스터 저장 (중복 시 덮어쓰기) |
| `upsertBizDivision` | 쓰기 | 사업부문 마스터 저장 |
| `upsertTaxInvoices` | 쓰기 | 세금계산서 저장 |
| `upsertLedger` | 쓰기 | 계정별원장 저장 (ledgerType: 매출/매입/미지급) |
| `upsertDailySales` | 쓰기 | 영업현황 저장 |
| `appendPaymentPlans` | 쓰기 | 결제계획 저장 |
| `appendPaymentHistory` | 쓰기 | 결제이력 저장 |
| `appendUpdateHistory` | 쓰기 | 변경 이력 기록 |
| `sendReceivableEmails` | 쓰기 | 미수금 이메일 발송 |
| `sendRawDiffEmail` | 쓰기 | 미지급 변경 감지 이메일 |
| `sendPaymentWarningEmail` | 쓰기 | 결제 경고 이메일 |
| `saveMautoData` | 쓰기 | 엠오토 데이터 저장 (`엠오토_json` 시트 A1 셀) |
| `upsertAvailableFunds` | 쓰기 | 가용자금 JSON 저장 (`가용자금_json` 시트, A1=updatedAt, B1=data) |

**읽기**: gviz(공개 URL) 우선, 실패 시 Apps Script fallback  
**쓰기**: 항상 Apps Script (`postSheetWebApp` 함수)

---

## 미수금 / 미지급 데이터 흐름

### 미지급 (`loadSheetPayables`)

**1단계** — raw + 업체마스터 동시 로드 → diff 없으면 즉시 렌더링:
```
fetchVendorMasterRowsFromApi()  → action=getVendorMaster
fetchSheetWebApp()              → Apps Script 기본 URL (미지급_raw 시트)
                                   └→ 인증 실패 시 gviz fallback
```

**2단계** — 백그라운드 로드 후 재적용:
```
fetchSavedPaymentPlansFromApi() → action=getPaymentPlans
fetchPaymentHistoryRowsFromApi()→ action=getPaymentHistory
```

**데이터 처리 파이프라인:**
```
rows
→ parsePayableRow()               각 행 파싱 (거래처코드, 연월, 매입, 지급합 등)
→ detectPayablesRawDiff()         이전 스냅샷과 변경 비교 (금액변경/삭제 감지)
→ applySavedPayablesState()       localStorage 상태 복원 (결제계획, 선택, paidOverride)
→ applySavedPaymentPlansFromApi() 원격 결제계획 적용 (로컬보다 최신인 경우만)
→ applyPaymentHistoryRows()       결제이력 기반 완료/부분결제 처리
→ ensureAutoPaymentPlans()        납기 그룹별 자동 결제일 계산
→ enrichPayablesWithVendorMaster()업체마스터에서 은행/계좌 정보 매칭
→ persistPayablesState()          localStorage 저장 + Apps Script 동기화 (700ms debounce)
```

**로컬 저장 키:** `receivable-payable-webapp.payables-state.v1`

### 미수금 (`loadSheetReceivables`)

**동시 로드:**
```
fetchReceivablesFromApi()    → action=getReceivables → 실패 시 gviz (raw 시트)
fetchManagerMasterFromApi()  → action=getManagerMaster
```

**처리:**
```
rows → parseReceivableRow() → enrichReceivablesWithManager() → renderReceivables()
```

---

## 자료업로드 버튼

헤더의 **"자료업로드"** 버튼 → 패널 토글 → 5종 파일 업로드

| 항목 | 파일 형식 | 파서 | 헤더 위치 | Apps Script action |
|------|-----------|------|-----------|-------------------|
| 세금계산서 (매출/매입 통합) | `.xls/.xlsx` | `parseTaxInvoiceFile` | 7행 (index 6) | `upsertTaxInvoices` |
| 계정별원장 — 외상매출금 | `.xls/.xlsx` | `parseLedgerFile` | 1행 (index 0) | `upsertLedger` (ledgerType=매출) |
| 계정별원장 — 외상매입금 | `.xls/.xlsx` | `parseLedgerFile` | 1행 (index 0) | `upsertLedger` (ledgerType=매입) |
| 계정별원장 — 미지급금 | `.xls/.xlsx` | `parseLedgerFile` | 1행 (index 0) | `upsertLedger` (ledgerType=미지급) |
| 영업현황 (일별) | `.xls/.xlsx` | `parseDailySalesFile` | 8행 (index 7) | `upsertDailySales` |

**동작 순서:**
1. 파일 선택 → 로컬에서 즉시 파싱 (XLSX.js)
2. 업체마스터와 매칭 (`matchVendorEntry` — 사업자번호 or 거래처코드 기준)
3. "구글시트 저장" 클릭 → `postSheetWebApp(action, rows)` → 중복 행은 자동 덮어쓰기

> ⚠️ **자료업로드는 대사 탭 전용 시트만 채웁니다** (세금계산서_raw, 계정별원장_*_raw, 영업현황_raw).  
> 미지급(`미지급_raw`) / 미수금(`raw`) 시트는 ERP에서 별도 내려받아야 합니다.  
> 고정지출은 localStorage(`cashflow-app.fixed-v1`) 우선 → 구글시트 순서이므로,  
> localStorage에 이전 데이터가 남아있으면 구글시트 신규 데이터가 무시됩니다.

**중복 방지 키 (`_row_key`):**
- 세금계산서: 승인번호 (없으면 작성일자+사업자번호+합계)
- 계정별원장: 일자+견표번호+거래처코드
- 영업현황: 전표번호 (없으면 거래일자+거래처코드+판매금액+구매금액)

---

## 대사 탭

**목적:** 세금계산서 ↔ 계정별원장 ↔ 영업현황 수치를 업체별·월별로 교차 비교

**데이터 로드 (`loadDaesaData`):**

탭 진입 시 버튼 클릭으로 수동 로드 (앱 시작 시 자동 로드 안 함)

```
Apps Script 동시 6개 요청:
  getTaxInvoices    → 세금계산서 (매출/매입)
  getLedgerSales    → 계정별원장 외상매출금
  getLedgerPurchase → 계정별원장 외상매입금
  getLedgerPayable  → 계정별원장 미지급금
  getDailySales     → 영업현황 (일별)
  getBizDivision    → 사업부문 마스터 (없으면 빈 배열)
```

**표시 구조 (`buildDaesaMap`):**

업체코드 기준으로 집계:

| 컬럼 | 출처 |
|------|------|
| 세금계산서 매출 | `taxInvoices` 중 매출 행 |
| 계정원장 매출 | `ledgerSales` |
| 영업현황 매출 | `dailySales` |
| 세금계산서 매입 | `taxInvoices` 중 매입 행 |
| 계정원장 매입+미지급 | `ledgerPurchase` + `ledgerPayable` |
| 영업현황 매입 | `dailySales` |

필터: 연도·월 선택, 검색창 연동  
정렬: 컬럼 헤더 클릭 (거래처명/각 금액 컬럼)

---

## 입출금 매칭 버튼

헤더의 **"입출금 매칭"** 버튼 → 은행 입출금 엑셀 파일 업로드

```
파일 선택 → XLSX.js 파싱 → parseBankSheet()
  헤더: 날짜 / 적요 / 출금 / 입금 (인식 필수)
→ openBankImportDialog()  → 대사 탭에 결과 표시
```

---

## 업체마스터

---

## 가용자금 탭 (2026-04-27 구현)

Google Sheets 대신 **Excel 복사+붙여넣기**로 데이터 입력. localStorage에 저장되어 새로고침 후에도 유지.

### 데이터 구조

```js
availableFunds = {
  accounts: [],        // 계좌: [{bank, accountNo, balance}]
  b2bLoans: [],        // B2B 대출: [{latestExpiry, execNo, finalExpiry, used}]
  purchaseVendors: [], // 구매자금 업체: [{date, name, amount}]
  eBonds: [],          // 전자채권: [{expiry, client, receiptDate, amount}]
  eNotes: [],          // 전자어음: [{bank, client, receiptDate, expiry, amount}]
}
```

### 각 섹션 엑셀 헤더 (정확히 일치해야 함)

| 섹션 | 헤더 |
|------|------|
| ① 계좌 | `은행` / `계좌번호` / `가용자금` |
| ② B2B 대출 | `최신만기일` / `실행번호` / `최종만기` / `합계` |
| ③ 구매자금 사용가능 업체 | `작성일자` / `업체명` / `금액` |
| ④ 전자채권 | `만기일` / `거래처명` / `수납일` / `합계` |
| ⑤ 전자어음 | `은행` / `거래처명` / `수납일` / `만기일` / `합계` |

### B2B 총대출액

`B2B_TOTAL_LIMIT = 500000000` (5억 고정)  
사용가능액 = 5억 − 현사용액(합계 합산)

### localStorage 키

`cashflow-app.available-funds-v2`

---

## 헤더 버튼 기능

| 버튼 | 기능 |
|------|------|
| 🔑 토큰 | Apps Script API 토큰 설정 |
| 📗 구글시트 | 연결된 스프레드시트 열기 |
| 🛠 마스터 관리 | 업체마스터 업로드 / 원장→업체마스터 / 사업부문마스터 |
| 자료업로드 | 미지급 등 raw 데이터 업로드 |
| 입출금 매칭 | 대사 탭용 은행 입출금 파일 업로드 |

---

## 버그 수정 이력

---

### 2026-07-28 (9): 엠오토 탭 미수/미지급 화면에도 세금계산서·입출금 상세 드릴다운 추가

**문제:** 바로 아래(8) 드릴다운을 대사 탭에만 추가했더니, 정작 실사용하는 엠오토 탭 자체 미수/미지급 화면(연월별·업체별)에는 안 나타남 — 두 화면이 `buildArRecap()` 계산 로직은 공유하지만 렌더링 코드는 완전히 별개였기 때문.

**수정 (`app.js`):**
- `arRecapDetailColsHtml(발생상세, 충당상세)` 공용 헬퍼로 분리 (대사 탭·엠오토 탭 양쪽 재사용)
- `arRecapToMautoRows()`에 발생상세/충당상세 전달 추가
- `renderMautoAccountingTable()` / `renderMautoPayablesByVendor()`의 리프 행에 상세 행 추가 — 기존 연/월(업체) 펼침 토글에 자연스럽게 편입되어 별도 JS 불필요

**수정 파일:** `app.js` (버전 `?v=20260728l`)  
**커밋:** `1638e2f`

---

### 2026-07-28 (8): 미수/미지급 자동계산 뷰에 거래처×연월 상세 드릴다운 추가

**요청 배경:** 세금계산서 발생액과 입출금 충당액이 일치(잔액 0)해도 "이 거래처 이 달 물품대를 실제로 언제 줬는지" 확인하고 싶다는 요청.

**추가 (`app.js`, `style.css`):**
- `buildArRecap()`이 거래처×연월 집계 시 `발생상세`(세금계산서 개별 라인: 작성일자·승인번호·금액)와 `충당상세`(입출금 개별 라인: 실제 지급일·금액)도 함께 반환하도록 확장
- 대사 탭 "📊 미수/미지급" 표에서 행 클릭 → 세금계산서 상세 / 입출금 상세를 나란히 펼쳐 보여줌 (`daesaState.arRecapExpanded`로 펼침 상태 관리)
- 잔액 0(완납) 건도 행이 그대로 보여서 클릭 시 상세 확인 가능

**수정 파일:** `app.js`, `style.css`, `index.html` (버전 `?v=20260728h`)  
**커밋:** `b8a9b29`

---

### 2026-07-28 (7): 엠오토 입출금 분류에서 학습한 규칙이 분류규칙 관리에 안 보이던 문제 수정

**증상:** 엠오토 입출금 분류 다이얼로그에서 "분류규칙에 추가"로 새 규칙을 저장해도, 분류규칙 관리 패널에 안 보임.

**원인:** 저장 후 `분류규칙 관리` 패널의 **현재 사업체 필터(bizFilter)** 기준으로 규칙 목록을 재조회했는데, 필터가 "미래"로 남아있으면 방금 저장한 엠오토 규칙이 서버엔 저장됐어도 재조회 결과엔 안 잡힘.

**수정 (`app.js`):** 저장된 새 규칙을 필터와 무관하게 `rulesState.rows`에 직접 병합하도록 변경.

**수정 파일:** `app.js` (버전 `?v=20260728j`)  
**커밋:** `ffe2a66`

---

### 2026-07-28 (6): 분류규칙 관리에 거래처명·매칭키·고정분류 검색 기능 추가

사업체 필터만 있던 분류규칙 관리 패널에 실시간 텍스트 검색창 추가 (거래처명/매칭키/고정분류 대상, 사업체 필터와 동시 적용).

**수정 파일:** `app.js` (버전 `?v=20260728i`)  
**커밋:** `47c8991`

---

### 2026-07-28 (5): 분류규칙 저장/삭제 시 고정지출 자동계산(mautoFixedRules) 캐시 미동기화 수정

**증상:** 분류규칙 관리에서 결제예정일·예정금액을 수정해도 고정지출 자동계산 화면에 반영이 안 되고, 입출금 파일을 재업로드해야만 반영되는 것처럼 보임.

**원인:** 고정지출 자동계산 화면이 분류규칙 관리와 별도의 `mautoFixedRules` 캐시 변수를 써서, `rulesState.rows` 갱신과 무관하게 그대로 남아있었음.

**수정 (`app.js`):** `saveRule()`/`deleteRule()`에서 `syncMautoFixedRulesFromRulesState()` 호출 추가 — 저장/삭제 즉시 `mautoFixedRules`도 동기화.

**수정 파일:** `app.js` (버전 `?v=20260728h`)  
**커밋:** `46b2dd3`

---

### 2026-07-28 (4): 분류규칙 결제예정일·예정금액 변경이력 기록 기능 추가

**요청 배경:** 같은 항목이라도 결제일·금액이 시점에 따라 바뀔 수 있는데(예: 국민415310 결제일이 2025-09부터 2일→15일로 변경), 이를 추적할 방법이 없었음. 매달 전체 스냅샷을 쌓는 대신, 값이 실제로 바뀐 시점만 기록하는 변경이력(체인지로그) 방식으로 구현.

**추가 (`code.gs`):**
- `RULE_HISTORY_SHEET = "분류규칙_이력"` 시트 상수 추가
- `getRuleHistory`(GET) / `appendRuleHistory`(POST) 액션 추가

**추가 (`app.js`):**
- `saveRule()`에서 결제예정일·예정금액이 실제로 바뀐 필드만 감지해 `{사업체, 매칭방식, 매칭키, 거래처명, 필드, 이전값, 신규값, 변경일시}` 기록 (`appendRuleHistoryIfChanged`)
- 분류규칙 관리 툴바에 "📜 변경이력" 버튼 → 시계열 조회 패널 (사업체 필터 연동)

**⚠️ Apps Script 재배포 완료됨** (code.gs 변경)

**수정 파일:** `app.js`, `code.gs`, `style.css`, `index.html` (버전 `?v=20260728g`)  
**커밋:** `8ba0e46`

---

### 2026-07-28 (3): 분류규칙 관리에서 매칭키 수정 시 규칙이 중복 생성되던 버그 수정

**증상:** 분류규칙 관리에서 기존 규칙의 매칭키(예: 계좌번호)를 고쳐서 저장하면, 수정이 아니라 새 규칙이 추가되고 예전 규칙은 그대로 남음.

**원인:** 규칙의 고유 식별자(`_rule_key`)가 `사업체+매칭방식+매칭키`로 구성되는데, 매칭키를 바꾸면 식별자 자체가 바뀌어 저장 시 "기존 걸 수정"이 아니라 "새 규칙 추가"로 처리됨.

**수정 (`app.js`):** 저장 버튼 클릭 시 원래 키(`data-edit-key`)를 함께 넘겨, 키가 바뀐 경우 예전 규칙을 자동 삭제하도록 `saveRule()` 수정.

**⚠️ 이 버그로 이전에 중복 생성된 규칙은 자동으로 안 없어짐 — 수동 확인/삭제 필요**

**수정 파일:** `app.js` (버전 `?v=20260728e`)  
**커밋:** `22809a4`

---

### 2026-07-28 (2): 엠오토 고정지출 완료항목 이중집계 수정 + 분류규칙 저장/삭제 시 자동 재분류

**증상 1 — 고정지출 이중집계:** 날짜 그룹 안에 이미 결제완료(✓)된 항목과 미완료 항목이 섞여 있으면, 그룹 전체가 안 끝났다는 이유로 완료된 항목 금액까지 체크 합계에 계속 포함됨.

**수정 (`app.js`):** `renderMautoFixedAutoView`의 `grpExpected`, `calcFixedCheckedTotal`의 `byDate[d].amt` 계산에서 개별 항목의 `status==="완료"`를 그룹 완료 여부와 무관하게 제외하도록 수정.

**증상 2 — 분류규칙 미반영:** 분류규칙 관리에서 규칙을 저장/삭제해도 이미 업로드된 입출금 내역에는 반영이 안 되고, 엑셀 일괄가져오기만 자동 반영됐음.

**수정 (`app.js`):** `saveRule()`/`deleteRule()`에 `rebuildMautoRows()` 호출 추가 — 저장/삭제 즉시 기존 내역 재분류.

**수정 파일:** `app.js` (버전 `?v=20260728e`)  
**커밋:** `3e3c16c`

---

### 2026-07-28 (1): 엠오토 분류 결과 목록 날짜 내림차순 정렬 + 연월 분배 안내 문구 개선

- `openMautoClassifyResultView(rows)`: 날짜(및 시간) 내림차순 정렬 추가 — 최근 거래가 위로
- "연월=금액" 다중 연월 분배 입력 안내 문구를 실제 사용 형식(괄호 설명, 콤마 포맷 금액)에 맞게 수정

**수정 파일:** `app.js`  
**커밋:** `5fb9eca`

---

### 2026-07-08: 다른 컴퓨터에서 엠오토 미수/미지급 안 뜨는 근본 원인 수정 (`fetchSheetWebApp` action 무시 버그)

**증상:** 입력 컴퓨터에선 엠오토 미수금/미지급이 정상 표시되는데, 다른 컴퓨터에선 항상 비어있음. 강력 새로고침(Ctrl+Shift+R) 해도 그대로 → 캐시 문제 아님.

**진단:**
- 엠오토 미수/미지급은 붙여넣기(`mautoData.receivables/payables`)가 아니라 **세금계산서 − 입출금분류로 계산**됨 (`renderMautoTab` 내 `hasTax` 분기 → `buildArRecap`).
- 서버엔 세금계산서(`엠오토_세금계산서` 145건)·분류(`엠오토_분류`) 데이터가 정상 존재.
- **진짜 원인:** `fetchSheetWebApp()`이 파라미터를 선언하지 않아, 호출부의 `fetchSheetWebApp({ action: "getMautoTaxInvoices" })` 인자를 **통째로 무시**. 항상 action 없는 기본 응답(미지급_raw)만 받아옴 → `_row_key` 없는 행 → `loadMautoTaxRemote` 병합 0건 → `mautoTaxInvoices` 계속 빈 배열 → `hasTax=false` → 미수/미지급 빈 화면.
- 부가 버그: 함수가 `{rows:[...]}` 응답을 처리 못 하고 예외 throw (`getMautoTaxInvoices`/`getClassifiedRows`/`getMautoSourceRows` 모두 `{rows}` 형태).
- 입력 컴퓨터는 `loadMautoTaxSource()`로 localStorage에서 로드돼 정상으로 보였음.

**수정 (`app.js`):**
- `fetchSheetWebApp(params = {})`: `params`의 `action` 등을 URL 쿼리에 반영. 반환값을 원본 `body`로 통일(`{rows}`·`{data}`·배열 모두 호출부에서 처리). 인증 실패 시 `fetchPublicSheet()` 폴백 유지.
- `loadSheetPayables`의 no-arg 호출(970행): 배열 언랩 `Array.isArray(b) ? b : (b.data||b.rows||[])`로 기존 미지급 로드 호환 유지.
- 이 수정으로 세금계산서뿐 아니라 **입출금분류·소스 행 원격 로드(컴퓨터 간 공유)도 처음으로 정상 동작**.

**수정 파일:** `app.js`, `index.html` (버전 `?v=20260708a`)

---

### 2026-06-16 (5): 입출금 원본 행 구글시트 공유 연동 (`엠오토_소스` 시트)

**문제:** 입출금 파일 업로드 시 파싱된 원본 거래 행(`mauto-source-files-v1`)이 localStorage에만 저장되어 다른 컴퓨터에서 규칙 변경 후 재분류가 불가능함.

**해법:**
- `code.gs`: `MAUTO_SOURCE_SHEET = "엠오토_소스"` 시트 상수 추가
  - `getMautoSourceRows` GET action 추가 → `엠오토_소스` 시트 전체 반환
  - `upsertMautoSourceRows` POST action 추가 → `_txKey` 기준 upsert
- `app.js`:
  - `saveSourceFiles()`: localStorage 저장 후 3초 debounce로 `엠오토_소스` 시트에 upsert (공유 필드: `_txKey, fileKey, filename, date, time, _memo, _memo2, _bank, _account, credit, debit`)
  - `loadMautoSourceRemote()` 함수 신규 추가: 로컬에 없는 `_txKey` 행만 병합 → `rebuildMautoRows()` 호출 → `renderMautoTab()`
  - `switchTab("mauto")` + `setupTabs` mauto 분기: 탭 진입 시 `loadMautoSourceRemote()` 자동 호출

**⚠️ Apps Script 재배포 필요** (code.gs 변경)

**수정 파일:** `app.js`, `code.gs`, `index.html` (버전 `?v=20260616e`)

---

### 2026-06-16 (4): 고정지출 체크박스·예정금액 수동수정 구글시트 공유 연동

**문제:** 고정지출 체크박스 상태(`mauto-fixed-checked-v1`)와 예정금액 수동 수정(`mauto-fixed-amount-overrides-v1`)이 localStorage에만 저장되어 다른 컴퓨터에서 반영 안 됨.

**해법 (`app.js`):**
- `saveFixedChecked()`: localStorage 저장 후 `_scheduleMautoRemoteSave()` 추가 호출
- `saveFixedAmountOverrides()`: localStorage 저장 후 `_scheduleMautoRemoteSave()` 추가 호출
- `_scheduleMautoRemoteSave()`: payload에 `fixedChecked`, `fixedAmountOverrides` 추가 포함 (`엠오토_json` A1 셀 JSON에 저장)
- `loadMautoDataRemote().then()` 핸들러 (switchTab + setupTabs 2곳): 원격 데이터에서 `fixedChecked`/`fixedAmountOverrides` 추출하여 전역 변수 및 localStorage 갱신

**저장 위치:** `엠오토_json` 시트 A1 셀 JSON 내 `fixedChecked` / `fixedAmountOverrides` 필드 — code.gs 수정 불필요

**수정 파일:** `app.js`, `index.html` (버전 `?v=20260616d`)

---

### 2026-06-16 (3): 엠오토 제외설정 구글시트 공유 연동

**문제:** 미수/미지급 탭의 제외설정(`mautoExcludeVendorsRcv`, `mautoExcludeVendorsPay`)이 localStorage에만 저장되어 다른 컴퓨터/사용자에서 공유 안 됨.

**해법 (`app.js`):**
- `saveMautoExcludeVendors(side)`: localStorage 저장 후 `_scheduleMautoRemoteSave()` 추가 호출
- `_scheduleMautoRemoteSave()`: `{ data: mautoData }` → `{ data: { ...mautoData, excludeRcv, excludePay } }` 로 수정 (제외설정 포함 저장)
- `loadMautoDataRemote().then()` 핸들러 (switchTab + setupTabs 2곳 모두):
  - `normalizeMautoData(remote)` 호출 전 `remote.excludeRcv` / `remote.excludePay` 추출하여 전역 변수 및 localStorage 갱신

**저장 위치:** `엠오토_json` 시트 A1 셀 JSON 내 `excludeRcv` / `excludePay` 필드 — code.gs 수정 불필요

**수정 파일:** `app.js`, `index.html` (버전 `?v=20260616c`)

---

### 2026-06-16 (2): 엠오토 세금계산서 구글시트 공유 + 대사 탭 자동 로드

**문제:** 다른 컴퓨터에서 엠오토 탭 세금계산서가 없어 미수/미지급 잔액이 0으로 표시됨. 대사 탭도 소스 파일 없으면 수동 버튼 클릭 필요.

**수정 (`app.js`):**
- `saveMautoTaxSource()`: localStorage 저장 후 구글시트 debounce 저장(3초) 추가
  - `postSheetWebApp("upsertMautoTaxInvoices", { rows: mautoTaxInvoices })` → `엠오토_세금계산서` 시트 upsert
- `loadMautoTaxRemote()` 함수 신규 추가: `getMautoTaxInvoices` action으로 구글시트에서 세금계산서 로드
  - 로컬에 없는 행만 `mautoTaxInvoices`에 병합 → `renderMautoTab()` 갱신
  - 가상 소스 항목 `__remote__` 생성 (파일 뱃지에 "원격 로드"로 표시)
- `switchTab("mauto")`: 탭 진입 시 `loadMautoTaxRemote()` 자동 호출 추가
- `switchTab("daesa")`: 로컬 소스 없고(`!hasMiraeSources()`) 데이터 미로드 시 `loadDaesaData()` 자동 호출

**수정 (`code.gs`):**
- `MAUTO_TAX_SHEET = "엠오토_세금계산서"` 시트명 상수 추가
- `doGet`: `getMautoTaxInvoices` action 추가 → `엠오토_세금계산서` 시트 전체 반환
- `doPost`: `upsertMautoTaxInvoices` action 추가 → `_row_key` 기준 upsert

**공유 범위 변경:**
- 세금계산서 파일: `mauto-tax-source-v1` ❌ → `엠오토_세금계산서` 시트 ✅
- 대사 탭: 소스 파일 없으면 자동으로 구글시트에서 로드

**⚠️ Apps Script 재배포 필요** (code.gs 변경)

**수정 파일:** `app.js`, `code.gs`, `index.html` (버전 `?v=20260616b`)

---

### 2026-06-16: 엠오토 미지급 업체별 보기 + 고정지출 예정금액 인라인 수정

**기능 추가 (`app.js`):**
- `renderMautoPayablesByVendor(payRows)`: 엠오토 미지급 업체별 집계 뷰
  - 잔액 절댓값 내림차순 정렬 (잔액 큰 업체 위로)
  - 업체 행 클릭 → 연월별 상세 펼침/접힘 (기존 toggle 핸들러 재사용)
  - 총합계 행에 업체 수 표시
  - `mautoPayViewMode = "ym" | "vendor"` 상태 변수 추가
  - 미지급 섹션 헤더 **[연월별] [업체별]** 토글 버튼 — 전환 시 섹션 열린 상태 유지
- 고정지출 예정금액 인라인 직접 수정 (localStorage 영구저장):
  - `MAUTO_FIXED_AMOUNT_KEY = "mauto-fixed-amount-overrides-v1"` — `{ "YYYY-MM||거래처명||예정일": amount }`
  - `applyFixedAmountOverrides(monthData)`: `buildFixedFromRules` 결과에 오버라이드 적용
  - 항목 행 예정금액 셀 → `<input>` 교체 (클릭 즉시 수정)
    - 빈칸 저장 시 자동계산 복원 / 수정된 항목은 주황 하단선 + `수정` 뱃지
  - blur/Enter 시 소계 행·카드 합계·예상 잔액 실시간 갱신
  - **⚠️ localStorage 전용 — 다른 컴퓨터와 공유 안 됨** (구글시트 연동 미구현)

**공유 범위 정리 (2026-06-16 수정 기준):**
| 데이터 | 저장 위치 | 공유 여부 |
|--------|-----------|-----------|
| 입출금 분류 결과 | `엠오토_분류` 시트 | ✅ 공유 |
| 분류규칙 | `분류규칙` 시트 | ✅ 공유 |
| 가용자금·미수금·미지급·고정지출(붙여넣기) | `엠오토_json` 시트 | ✅ 공유 |
| 세금계산서 파일(파싱 행) | `엠오토_세금계산서` 시트 | ✅ 공유 (2026-06-16 수정) |
| 미수/미지급 제외설정 (`excludeRcv`/`excludePay`) | `엠오토_json` 시트 내 JSON | ✅ 공유 (2026-06-16 수정) |
| 체크박스 상태 | `엠오토_json` 시트 내 JSON (`fixedChecked`) | ✅ 공유 (2026-06-16 수정) |
| 예정금액 수동 수정 | `엠오토_json` 시트 내 JSON (`fixedAmountOverrides`) | ✅ 공유 (2026-06-16 수정) |
| 입출금 원본 행 (파싱 데이터) | `엠오토_소스` 시트 | ✅ 공유 (2026-06-16 수정) |

**수정 파일:** `app.js`, `index.html` (버전 `?v=20260615ah`)  
**커밋:** `bfa583e`, `067ec64`

---

### 2026-06-15 (추가): 엠오토 탭 부가세 보고서 + 모바일 반응형 + 버그 수정

**기능 추가 (`app.js`, `style.css`):**
- `renderMautoVatView()`: 엠오토 세금계산서 기반 부가세 납부세액 집계 보고서
  - 기간 모드: **반기(기본, 개인사업자 기준)** / 월간 / 연간
  - 엠오토 탭 상단 **📊 부가세** 버튼으로 토글 (카드 그리드 아래 표시)
  - `mautoVatView / mautoVatMode / mautoVatYear` 상태 변수 추가
  - `buildVatSummary` / `buildVatPeriods` 기존 함수 재사용
  - 세금계산서 없으면 파일 업로드 안내 메시지
- 엠오토 탭 모바일 반응형 (`@media max-width: 768px/480px`):
  - `.app-shell` min-width 1180px → 320px로 해제
  - 탭 바: `overflow-x: auto` (가로 스크롤, 한 줄 유지)
  - 엠오토 상단 버튼 4개 → 2×2 그리드
  - 카드 그리드: 768px=3열, 480px=2열
  - `.mauto-top-actions` CSS 클래스 (inline style 제거)
  - 고정지출 자동계산 테이블: `overflow-x: auto` 래퍼 + `min-width: 480px`

**버그 수정 (`app.js`):**
- **분류규칙 관리 스크롤 위치 초기화**: `renderRulesPanel()` 호출 후 `.rules-table-wrap` scrollTop 0으로 리셋되던 문제
  - `panel.innerHTML` 교체 전 `_prevScroll` 저장 → 교체 후 복원
- **고정지출 완료 항목 체크박스 자동 해제**: 날짜 그룹 내 모든 항목 `status=완료` 시
  - 체크박스 `disabled` + 체크 해제 + opacity 35%
  - `calcFixedCheckedTotal`: 완료 그룹 예정금액에서 자동 제외
- **고정지출 합계 항상 0**: 탭 버튼 클릭(`setupTabs`) 시 `mautoFixedRules` 자동 로드 누락
  - `switchTab`에만 있던 `fetchRulesFromApi("엠오토")` 호출을 `setupTabs` mauto 분기에도 추가

**수정 파일:** `app.js`, `style.css`, `index.html` (버전 `?v=20260615af`)  
**커밋:** `cb23a63`, `950a156`, `2a5fe9d`, `f01bd2f`, `9a6be27`

---

### 2026-06-15: Phase 4-B 고정지출 자동계산 완성

**기능 추가 (`app.js`):**
- `KR_HOLIDAYS`: 한국 공휴일 Set (2025~2026, 대체공휴일 포함)
- `getScheduledPaymentDate(year, month, day)`: 결제예정일(N일) → 주말/공휴일 조정된 실제 영업일 반환 `{ date, dow }`
- `buildFixedFromRules(fixedRules, classifiedRows)` 전면 개편:
  - **자동 월 생성**: 오늘 기준 3개월 전 ~ 6개월 후 자동 생성 (입출금 데이터 없어도 예정표 표시)
  - **거래처 중복 제거**: 같은 `거래처명+결제예정일` 규칙이 여러 개여도 표시는 1행
  - **예정금액 자동계산**: 2개월 이상 실적 있으면 월 평균 → 천원 단위 올림 자동 사용; 데이터 부족 시 규칙 수동값 fallback
  - **완료 플래그**: `allDone` — 전 항목 완료 여부
- `renderMautoFixedAutoView()` 전면 개편:
  - 날짜 오름차순 정렬 + 날짜별 소계행 (예정금액 합 / 실적 합)
  - 분류(이자/카드 등)는 항목 행 옆 작은 태그로만 표시 (그룹 헤더 제거)
  - 완료된 과거 달 → `<details>` 기본 접힘 (`▶`), 미결 과거 달 → 빨간 `● 미결`
  - 조정된 날짜(`*`) 표시 (원래 N일이 주말/공휴일로 밀린 경우)
  - `계산` 파란 뱃지: 자동계산된 예정금액 구분
- `importRulesFromExcel` 버그 수정: `결제예정일`, `고정분류`, `예정금액` 컬럼 누락 (기본 6개만 하드코딩돼 있던 문제)

**분류규칙 seed 파일:**
- `docs/분류규칙_seed_엠오토.xlsx` — 기본 73개 매칭규칙 (항상 먼저 올릴 것)
- `docs/분류규칙_seed_고정지출추가.xlsx` — 36개 고정지출 규칙 (`결제예정일`, `고정분류`, `예정금액` 포함)
- 두 파일 모두 업로드 필요 (순서: 엠오토.xlsx → 고정지출추가.xlsx)

**분류규칙 컬럼 (전체):** `사업체`, `매칭방식`, `매칭키`, `거래처명`, `구분`, `우선순위`, `결제예정일`, `고정분류`, `예정금액`

**수정 파일:** `app.js`, `index.html` (버전 `?v=20260615q`)  
**커밋:** `4d6631e`, `061f1a6`, `14e6e80`, `c3d56b4`, `c675df7`

---

### 2026-06-12 (8): 엠오토 세금계산서 파서 (국세청 전자세금계산서 조회 XLS)

**기능 추가 (`app.js`):**
- `parseMautoTaxInvoiceFile(file)`: 국세청 전자세금계산서 조회 파일 파서
  - 파일 구조: Row 1=제목, Row 7(index 6)=헤더, Row 8+=데이터, 마지막 행=합계(체크섬)
  - 헤더: `구분|종류|작성일자|사업자(주민)번호|종사업장번호|상호|대표자명|공급가액|세액|합계|비고|승인번호|발급일자|발급유형`
  - 거래처 = `상호` (상대방): 매출 파일=매출처(구매자), 매입 파일=매입처(공급자)
  - 체크섬: 마지막 행(`구분=""`, `종류="N건"`) → `공급가액` 합계 대조
  - 잘못된 파일 경고: 모든 거래처명이 "엠오토"인 경우 알림 (타사 파일 업로드 실수 방지)
  - `_row_key` = 승인번호 (없으면 `작성일자_사업자번호_합계`)
- `MAUTO_TAX_SOURCE_KEY = "mauto-tax-source-v1"` — 파일 단위 소스 보관
- `mautoTaxSources`, `mautoTaxInvoices` 전역 변수 + 저장/로드/재빌드 헬퍼
  - `saveMautoTaxSource()`, `loadMautoTaxSource()`, `rebuildMautoTaxInvoices()`
- 엠오토 탭 상단 버튼 추가: `🧾 매출 세금계산서`, `🧾 매입 세금계산서`
  - 파일 선택 → 파싱 → 체크섬 확인 → `mautoTaxSources` 저장 → `rebuildMautoTaxInvoices()`
  - 같은 파일명 재업로드 시 교체 확인
  - 저장 파일 뱃지 표시 (파일명/건수/체크섬 결과/삭제 버튼)
- `renderArRecapView()` 데이터소스 우선순위: `mautoTaxInvoices`(있으면) → `daesaState.taxInvoices`
  - 툴바에 현재 소스 표시 ("엠오토 세금계산서 N건" / "미래 세금계산서 N건")
- 앱 시작 시 `loadMautoTaxSource()` → `rebuildMautoTaxInvoices()` 자동 실행

**파일 구조 확인 (Python xlrd 검증):**
- 헤더 7행(index 6), 데이터 8행부터 (기존 `parseTaxInvoiceFile`과 동일)
- 체크섬 행: 마지막 행, `구분` 빈값, `종류`=`"N건"` 패턴
- 매출/매입 모두 `상호`(col 5)가 거래처 (엠오토는 항상 자기쪽이라 상대방만 표시됨)
- 소형 파일(9건 이하)은 체크섬 행 없음 → `checksumOk=null` (검증 생략)

**수정 파일:** `app.js`, `index.html` (버전 `?v=20260612m`)  
**커밋:** `4beb378`

---

### 2026-06-12 (9): Phase 4-A 거래처 매칭 견고화 — 사업자번호·코드 기반 join

**문제:** `buildArRecap`이 발생(세금계산서 `상호`)과 충당(분류 `거래처명`)을 **문자열로** 묶어 `중원전기(주)` vs `중원전기` 처럼 표기가 조금만 달라도 별개 행으로 갈라짐.

**해법:** 사업자번호 → 거래처코드_norm 기반 키로 전환 + 정규화 상호 fallback.

**매칭 키 우선순위 (발생·충당 동일 기준):**
1. 사업자번호 → 업체마스터 → `거래처코드_norm`
2. 정규화 상호 → 업체마스터 → `거래처코드_norm`
3. 사업자번호만 있고 마스터 없음 → `biz:NNNN`
4. 그 외 → `name:정규화상호`

**기능 추가 (`app.js`):**
- `normalizeVendorName(name)`: 법인 suffix(`(주)`, `주식회사` 등)·공백·괄호 안 영문병기 제거 후 소문자 비교키
  - `중원전기(주)` → `중원전기` / `케이앤에스 이엔지` = `케이앤에스이엔지(K&S ENG)` ✓
- `buildVendorNameMap()`: 업체마스터 거래처명 정규화 → `{ code, name, bizNum }` 맵
- `buildArRecap()` 전면 개편:
  - 내부 `getVendorKey(bizNum, name)` / `getVendorKeyByName(name)` 헬퍼
  - 발생맵·충당맵 모두 `vendorKey\tYYYY-MM` 키로 집계
  - `inexactKeys`: 코드매칭 실패(3·4순위) 키 추적 → `inexactVendors` 목록 반환
- `renderArRecapView()` 업데이트:
  - 툴바에 `🔗 업체마스터 미매칭 N개` 파란 배지 추가
  - 배지 클릭 → 이름으로만 묶인 거래처 목록 펼침 (사업자번호 마스터 등록 유도)
  - 하단 주석에 매칭 로직 설명 추가

**검증:**
- `중원전기(주)` vs `중원전기` → 정규화 후 `중원전기` 동일 → 마스터코드 기준 1행 묶임
- node 인라인 테스트 4케이스 ALL PASS

**수정 파일:** `app.js`, `style.css`, `index.html` (버전 `?v=20260612n`)  
**커밋:** `f6c38c6`

---

### 2026-06-12 (7): Phase 4-A — 미수/미지급 자동계산 (대사 탭 📊 미수/미지급)

**공식:** `잔액 = 발생액(세금계산서) − 충당액(입출금 분류)`

**기능 추가 (`app.js`, `style.css`):**
- `buildArRecap(taxInvoices, classifiedRows, sideType)`: 발생/충당 집계 함수
  - 발생액: `daesaState.taxInvoices` → `거래처 × 작성연월` 집계 (`합계` 기준)
  - 충당액: `mautoClassifiedRows` 중 `구분 ∈ {매출|매입}` + 입출금 방향 확인
    - `parseYearMonthCode(비고)` `status === "ok"` 인 행만 귀속연월로 집계
    - 파싱 실패 행 → `확인필요` 버킷 (버리지 않고 `<details>`로 표시)
  - 잔액 = 발생 − 충당 (거래처×연월 단위)
- `renderArRecapView()`: 거래처별 월별 발생/충당/잔액 표 + 소계/총계 + 확인필요 목록
  - 미수금/미지급 사이드 토글, 연도 필터, 검색 필터 연동
  - 귀속연월 미확인 건수 뱃지 (오파싱 행 경고)
- 대사 탭 툴바 "📊 미수/미지급" 버튼 추가 (정산표·부가세와 상호 배타 토글)
- `daesaState`에 `arRecapView`, `arRecapSide`, `arRecapFilterYear` 상태 추가

**검증 기준 (`phase4_미수미지급_핸드오프.md`):**
- 중원전기 2022-11 미지급 잔액 = 레거시 `수금자료들>ㅁㅇ실험` 동일 값 일치
- `부산743106`·`국민1억(...)` 등 엉뚱 연월 행이 확인필요로 빠지는지
- 급여·이자·인출이 충당에 미포함 (구분=공란 → 필터됨)

**수정 파일:** `app.js`, `style.css`, `index.html` (버전 `?v=20260612l`)  
**커밋:** `fab2a94`

---

### 2026-06-12 (6): 미래 자료업로드 재빌드 모델 + 원장 전표번호 오타 수정

**버그 수정: parseLedgerFile `_row_key` 견표번호 → 전표번호**
- 원인: 실제 Excel 헤더가 `전표번호`인데 코드에서 `row["견표번호"]`로 잘못 접근
- 결과: `일자__거래처코드` (전표번호 빠진 키) → 같은 날 같은 거래처 여러 전표가 1건으로 덮어써짐
- 수정: `row["전표번호"] || row["견표번호"]` (견표번호 fallback 유지)
- 확인 방법: 계정별원장 XLS 바이너리 스캔 → `전표번호` 확인, `견표번호` 없음

**기능 추가 — 파일 단위 교체 + 로컬 재빌드 모델 (`app.js`, `style.css`):**
- `MIRAE_SOURCE_TAX_KEY / LEDGER_KEY / BIZ_KEY` localStorage 3종 소스 보관소 추가
- `getMiraeSectionFiles` / `saveMiraeSectionFile` / `deleteMiraeSectionFile` 헬퍼
- `rebuildDaesaFromSources()`: 보관된 소스 파일 전체 → dedup → `daesaState` 갱신
  - 세금계산서: 승인번호(`_row_key`) dedup, 원장 3종: `ledgerType`별 dedup, 영업현황: dedup
  - 파일 삭제 시 해당 행이 집계에서 자동으로 사라짐 (파워쿼리 refresh 동일 효과)
- 자료업로드 패널:
  - 파일 선택 즉시 소스 저장 + `rebuildDaesaFromSources()` 호출 (클라우드 불필요)
  - 저장된 파일 뱃지 표시 (파일명/건수/✕삭제)
  - 같은 파일명 재업로드 시 기간 교체 확인 모달
  - '구글시트 저장' 유지 (클라우드 백업용)
- 앱 startup: 소스 파일 로드 → 소스 있으면 `daesaState` 즉시 재빌드 (Google Sheets 로드 전 선행)

**수정 파일:** `app.js`, `style.css`, `index.html` (버전 `?v=20260612k`)  
**커밋:** `b90b393`

---

### 2026-06-12 (5): Phase 3 — 부가세 납부세액 집계 대시보드

**기능 추가 (`app.js`, `style.css`):**
- `buildVatSummary(taxInvoices)`: 세금계산서 → 작성연월별 `{ 매출공급, 매출세액, 매입공급, 매입세액 }` 집계 맵 반환
  - 집계 기준: **작성일자(작성연월)** (입출금일·거래일자 아님)
  - 합산 대상: **세액** (합계 아님)
  - 구분 공란(이자·인출 등) 자동 제외
- `buildVatPeriods(year, mode)`: 월간/분기/반기/연간 기간 묶음 생성
- `renderVatView()`: 기간별 매출세액·매입세액·납부(환급)세액 표 렌더링
  - 납부세액 = 매출세액 − 매입세액 (음수 = 환급, 녹색 표시)
  - 연도·기간모드 선택 필터, 합계 행 포함
- 대사 탭 툴바 "🧾 부가세" 버튼 추가 (정산표 버튼과 상호 배타 토글)
- `daesaState`에 `vatView`, `vatMode`, `vatYear` 상태 추가

**사업체별 신고기간 기본값 (Phase 3 스펙):**
- 미래(법인) → 분기 (1~3, 4~6, 7~9, 10~12월)
- 엠오토(개인) → 반기 (1~6, 7~12월)
- 기본값만 다르며 사용자가 모드 변경 가능

**수정 파일:** `app.js`, `style.css`, `index.html` (버전 `?v=20260612j`)  
**커밋:** `761e3e2`

---

### 2026-06-12 (4): Phase 2 저장모델 전환 — 불변/사용자 영역 분리 + 합계행 필터

**저장 모델 분리 (`app.js`):**
- `MAUTO_SOURCE_FILES_KEY = "mauto-source-files-v1"` — 파일 단위 원본 거래 보관 (파일명 기준 교체)
- `MAUTO_USER_EDITS_KEY = "mauto-user-edits-v1"` — 거래키별 사용자 수정 (거래처·구분·제외·오버라이드)
- `rebuildMautoRows()`: source-files 전체 → 날짜정렬 → `classifyBankRow` → user-edits 덮어쓰기
- `migrateLegacyIfNeeded()`: 기존 `mauto-classified-rows-v1` → 새 모델로 자동 1회 전환
- 파일 업로드: 같은 파일명 재업로드 시 교체 확인 → source-files에 통째 교체
- 파일 목록 UI: 파일명/건수/날짜/개별 삭제 버튼

**합계행 필터 수정 (`app.js`, `parseBankSheet`):**
- 기존: `r._debit > 0 || r._credit > 0` — 금액 있는 합계 행 통과됨
- 수정: `/^\d{4}/.test(r._date) && (r._debit > 0 || r._credit > 0)` — 거래일자가 연도(숫자 4자리)로 시작해야 유효

**수정 파일:** `app.js`, `index.html` (버전 `?v=20260612i`)  
**커밋:** `19c6df7`, `9ce94c0`

---

### 2026-06-12 (3): 분류규칙 관리 UI 개선 (sticky 헤더·구분 빈값·아코디언·툴바 한줄) (sticky 헤더·구분 빈값·아코디언·툴바 한줄)

**변경 내용 (`app.js`, `style.css`):**
1. **분류규칙 표 헤더 sticky** — `style.css` `.rules-table-wrap`에 `overflow-y:auto; max-height:420px`, `th`에 `position:sticky; top:0; z-index:2` 추가
2. **구분 빈값 옵션 추가** — `DIV_OPTIONS = ["", "매출", "매입"]` / 렌더 시 빈값은 `(없음)` 표시
3. **아코디언 (기본 접힘, 수동 로드)** — `rulesState.tableOpen = false` 기본값; `▶` 클릭 시 `tableOpen` 토글 후 처음 열릴 때만 `loadRules()`; 패널 열 때 자동 로드 제거
4. **툴바 한 줄** — `renderRulesPanel()` 내 `flex-wrap:nowrap` 단일 `<div>`에 토글버튼·제목·카운트·구분선·필터·버튼 전부 배치; ✕닫기는 우측 끝 `margin-left:auto`

**수정 파일:** `app.js`, `style.css`, `index.html` (버전 `?v=20260612g`)  
**커밋:** `dfc4e07`, `46ee7fb`

---

### 2026-06-12 (1): Phase 2 — 엠오토 입출금 분류 보강 (중복방지·제외토글·규칙학습)

**Phase 2④ — 재업로드 중복 방지 + localStorage 영구저장 (`app.js`):**
- `MAUTO_CLASSIFIED_KEY = "mauto-classified-rows-v1"` localStorage 키 추가
- `saveClassifiedRows()` / `loadClassifiedRows()` 헬퍼 추가, 앱 시작 시 `loadClassifiedRows()` 호출
- `assignTxKeys(bankRows)`: 거래키 부여 — `_time` 있으면 `date|time|credit|debit|memo`, 없으면 배치 내 같은 기본키에 `#N` 시퀀스 번호 추가
- `mergeClassifiedRows(existing, incoming)`: Map 기반 merge — 기존 키 보존(건너뜀), 신규만 추가; 건너뜀 건수 반환
- 파일 업로드 핸들러: `parseBankSheet()` → `assignTxKeys()` → `mergeClassifiedRows()` 순으로 변경
- 저장 대상: 제외·미매칭 행 포함 **모든 행** 저장 (기존엔 `거래처명 있는 행만` 필터)
- 저장 행 shape: `_txKey, date, time, _memo, _memo2, memo, credit, debit, 거래처명, 구분, excluded, 매칭근거`

**Phase 2② — 제외 전체선택/전체해제 마스터 체크박스 3-상태 (`app.js`):**
- `updateMasterCheckbox()`: `전체/일부/없음` → `checked + indeterminate` 3-상태 업데이트
- 마스터 체크박스 클릭 → `전체/일부` 상태면 전체 해제, `없음`이면 전체 체크 (visible 행 기준)
- 저장 후 요약바에 "⚠ 미매칭 N건" 주황 경고 표시 (excluded vs 미매칭 구분 명확화)
- 결과보기(`openMautoClassifyResultView`)에 상태 컬럼 추가 (✓분류/⊘제외/? 미매칭)

**Phase 2③ — 규칙 학습 (오버라이드 행 → 분류규칙 저장) (`app.js`):**
- `isOverride` 플래그: 거래처 셀렉트 변경 시 설정 (구분만 바꾼 경우 포함)
- 규칙 학습 UI: 거래처명 있는 행에 "☑ 분류규칙에 추가" 서브행 표시
  - 매칭방식(키워드/거래처명/계좌) + 매칭키 입력 + 실시간 미리보기 "N건"
  - 매칭키 2자 미만 시 저장 불가 + 경고
  - 키워드: `_memo` 포함, 거래처명: `_memo2` 포함, 계좌: `_memo` 정확일치
- 중복 키 처리: 동일 키·같은 거래처 → 조용히 덮어쓰기, 다른 거래처 → 확인 모달
- 저장 순서: 분류 저장 → `upsertRules` 호출 (규칙 실패 시 분류는 롤백 안 함)
- `applyNewRulesToUnmatched(updatedRules)`: 저장 후 현재 화면 미매칭 행에만 새 규칙 즉시 적용 (분류됨/제외됨 건드리지 않음)

**수정 파일:** `app.js`, `index.html` (버전 `?v=20260612a`~`g`)  
**커밋:** `20935e9`, `814fcb3`, `6eb3a48`, `74b7cd5`

---

### 2026-06-11 (3): 계정별원장 거래처코드 앞자리 0 소실 수정

**증상:** 계정별원장 업로드 시 `05959` → `5959`로 저장됨. 정산표에서 코드 오표시.

**원인:** `parseXlsToRows`에서 `XLSX.utils.sheet_to_json` 기본값(`raw: true`)이 숫자처럼 보이는 셀을 숫자로 변환.

**수정 (`app.js`, `parseXlsToRows`):**  
숫자 셀 읽을 때 `XLSX.utils.encode_cell`로 원본 셀 접근 → `cell.w`(서식 텍스트)가 `0`으로 시작하면 숫자 대신 서식 텍스트 사용.
```javascript
if (typeof val === "number") {
  const cell = ws[XLSX.utils.encode_cell({ r: i, c: j })];
  if (cell && cell.w && /^0\d/.test(cell.w)) val = cell.w;
}
```
적용 범위: `parseLedgerFile`, `parseTaxInvoiceFile`, `parseDailySalesFile` 공통 파서.  
⚠️ 이미 업로드된 데이터는 재업로드 필요.

**수정 파일:** `app.js`  
**커밋:** `780f29b`

---

### 2026-06-11 (2): 분류규칙 관리 패널 Excel 일괄 가져오기 버튼 추가

**기능 추가 (`app.js`):**
- `importRulesFromExcel(file)` — XLSX 파싱 → `upsertRules` 일괄 POST
- 분류규칙 관리 툴바에 `📂 Excel 가져오기` 파일 선택 버튼 추가
- seed 파일: `docs/분류규칙_seed_엠오토.xlsx` (73건)

**수정 파일:** `app.js`, `index.html`  
**커밋:** `ca65df1`

---

### 2026-06-11 (1): Phase 0 + Phase 1 — 분류규칙 관리 UI + 원장 정산표

**Phase 0 — 분류규칙 관리 (`code.gs`, `app.js`, `index.html`):**
- `분류규칙` 시트 (없으면 자동 생성), `_rule_key` = `사업체||매칭방식||매칭키`
- Apps Script: `getRules` (GET) / `upsertRules` / `deleteRule` (POST)
- 마스터 관리 드롭다운 → "분류규칙 관리" 패널 (`rulesPanel`) — 사업체 필터, 추가/편집/삭제

**Phase 1 — 미래 원장 정산표 (`app.js`, `style.css`):**
- `parseYearMonthCode(적요)`: 귀속연월 추출 (yy-mm / yymm / yymmdd 패턴)
- `buildLedgerSettlement(rows, 구분)`: 거래처×귀속연월 집계
  - 매출(108): 차변=발생합계, 대변=충당액
  - 매입(251): 대변=발생합계, 차변=충당액
- 대사 탭 "📊 정산표" 버튼 → `renderSettlementView()` (거래처별 소계/총계 + 확인필요 섹션)

**수정 파일:** `code.gs`, `app.js`, `style.css`, `index.html`  
**커밋:** `bfb0911`

---

### 2026-05-14 (5): pnl.html 서명 대기 뱃지 + 기간 옵션 마커

**기능 추가:**
- `needsViewerAction(entry)` — URL `?email=` 파라미터 기반으로 현재 접속자의 서명 가능 여부 판단
- `updatePendingBadge()` — 서명 대기 건수 계산, 툴바에 빨간 "● N건 서명 대기" 뱃지 표시
- `selPeriod` 옵션에 서명 필요 기간은 "● " 마커 표시
- `render()` → `updatePeriodSelect()` → `updatePendingBadge()` 흐름으로 서명 후 자동 갱신

**수정 파일:** `pnl.html`  
**커밋:** `2696a50`

---

### 2026-05-14 (4): pnl.html 분기 드롭다운 — 기안완료 이상 분기만 표시

**변경:** 분기 모드 `selPeriod`가 1~4분기 전부 표시 → 구글시트에 저장된(기안완료+) 분기만 표시  
- `pnlQuarterlyRows` 기준으로 해당 연도 유효 분기만 옵션 생성
- 현재 선택 분기에 데이터 없으면 마지막 유효 분기로 자동 이동
- 데이터 없는 연도는 "—" 표시

**수정 파일:** `pnl.html`  
**커밋:** `6acc418`

---

### 2026-05-14 (3): pnl.html 분기 보고서 상세 섹션 추가

**변경:** 분기 보고서(`buildQuarterlyReportHtml`)가 KPI + 결재란만 표시 → 월간 보고서 수준으로 보완

추가된 섹션:
- **섹션 1** — 관리기준 영업이익 흐름 (매출 → 원가/제조 차감 → 매출총이익 → 판관비 → 영업이익)
- **섹션 2** — 경영이익 흐름 (영업이익 → 영업외비용 차감 → 경영이익)
- **섹션 3** — 전분기 대비 손익 비교표 (`aggData(prevQY, prevQMonths)` 집계, 월간 데이터 있을 때만 표시)
- KPI 카드에 목표 매출액 달성률 추가

전분기 계산: `prevQ = quarter > 1 ? quarter - 1 : 4` / `prevQY = quarter > 1 ? year : year - 1`

**수정 파일:** `pnl.html`  
**커밋:** `0f09f63`

---

### 2026-05-14 (2): 분기 보고서 구글시트 저장 안 되는 버그 수정

**증상:** 기안 서명 후 pnl.html에서 분기 보고서가 "데이터 없음"으로 표시됨

**원인:** `_saveQtrToSheets`에서 `postSheetWebApp("savePnlData", [{...}])` — 배열을 payload로 직접 넘김.  
`postSheetWebApp` 내부에서 `...payload` 스프레드 시 `{ 0: {...} }` 형태가 돼 `body.rows = undefined` → 저장 0건.

**수정 (`app.js`, `_saveQtrToSheets`):**
```javascript
// 수정 전 (잘못됨)
postSheetWebApp("savePnlData", [{...}])

// 수정 후
postSheetWebApp("savePnlData", { rows: [{...}] })
```

**수정 파일:** `app.js`  
**커밋:** `d5dd52c`

---

### 2026-05-14 (1): 경영손익 보고서 반기/연간 모드 추가

**기능 추가 (`app.js`):**
- `pnlRptHalf` 상태 변수 (1=상반기, 2=하반기)
- `_pnlFlowHtml(entry, c)` / `_pnlMgmtFlowHtml(entry, c)` — 손익 흐름표 HTML 헬퍼
- `renderPnlHalfYearReport(el)` — 상/하반기 집계 + 전년 동기 비교
- `renderPnlAnnualReport(el)` — 연간 12개월 집계 + 전년 비교
- `renderPnlReport()` 분기 추가: `halfyear` / `annual` 분기
- 모드 버튼 HTML에 반기/연간 추가 (월간·분기·반기·연간 4개)

**연도 nav 버그 수정 (`app.js`):**  
`Array.from({length: curY - 2023}, ...)` 로 연도 옵션 생성 시 현재 보는 연도(`pnlRptYear`)가 범위 밖이면 브라우저가 첫 번째(2024)를 선택.  
→ `Math.max(curY + 1, pnlRptYear)`를 maxYear로 사용해 항상 현재 연도가 옵션 안에 포함되게 수정.

**기능 추가 (`pnl.html`):**
- `reportMode`: `"monthly" | "quarterly" | "halfyear" | "annual"`
- `rptHalf` 상태 변수
- 툴바: 월간/분기/반기/연간 버튼 + `selPeriod` 동적 select (연간 모드는 숨김)
- `aggData(yr, months)` — 지정 월 목록 합산 집계
- `updatePeriodSelect()` — 모드별 옵션 자동 전환
- `buildAggReportHtml(yr, label, months, prevYr, prevLabel, prevMonths)` — 반기·연간 공용 렌더러
- `buildHalfYearReportHtml(yr, half)` / `buildAnnualReportHtml(yr)` — 각각 호출

**수정 파일:** `app.js`, `pnl.html`, `index.html` (버전 쿼리 `?v=20260514g`)  
**커밋:** `136dfe4`  
**백업:** `app.js.backup.20260514b.txt`, `pnl.html.backup.20260514.txt`

---

### 2026-05-12 (2): 대사 탭 데이터 로드 속도 개선

**증상:** 대사 탭 "데이터 불러오기" 클릭 시 응답이 느림

**원인:** 6개 시트(세금계산서·원장3종·영업현황·사업부문마스터)를 개별 API 요청으로 각각 호출  
→ HTTP 왕복 6회, Apps Script 기동 6회, `SpreadsheetApp.openById()` 6회 발생

**수정 (`code.gs`):**
- `getSheetRows(sheetName, ss?)` — `ss` 파라미터 추가로 스프레드시트 객체 재사용 가능
- `getDaesaAll` 액션 추가 — 스프레드시트 1회 오픈 후 6개 시트 일괄 반환

**수정 (`app.js`):**
- `fetchDaesaAll()` 함수 추가 — `getDaesaAll` 단일 요청으로 교체
- `loadDaesaData()` — `Promise.all` 6개 요청 → `fetchDaesaAll()` 1회 호출로 변경

**효과:** HTTP 왕복 6회 → 1회 (체감 2~5배 빠름)

**수정 파일:** `code.gs`, `app.js`, `index.html` (버전 쿼리 `?v=20260512c`)  
**커밋:** `3e0b88d`  
**⚠️ Apps Script 재배포 필요**

---

### 2026-05-12 (1): 업체마스터 거래처코드 앞자리 0 소실 버그 수정

**증상:** 코드 `04159`를 업로드하면 구글시트에 `4159`로 저장됨

**원인:** Google Sheets `setValues()`가 숫자처럼 생긴 문자열(`"04159"`)을 자동으로 숫자(`4159`)로 변환  
→ 셀 포맷이 "일반(General)"일 때 발생. JavaScript 쪽 `normalizeVendorCode()`는 정상이었음

**수정 (`code.gs`, `upsertVendorMasterRows`):**  
`거래처코드_norm`, `거래처코드_raw`, `vendor_id` 컬럼에 `setValues()` 전 텍스트 포맷 강제 적용
```javascript
sheet.getRange(...).setNumberFormat("@");
```
빈 시트 분기(신규 등록)와 upsert 분기(기존 업데이트) 양쪽에 모두 추가

**수정 파일:** `code.gs`  
**커밋:** `9f63386`  
**⚠️ Apps Script 재배포 필요**

---

### 2026-05-11: 엠오토 탭 — 다른 컴퓨터에서 구글시트 데이터 불러오기 실패

**증상:** 데이터 입력 컴퓨터에서는 정상 저장, 다른 컴퓨터에서는 엠오토 탭이 항상 0/비어있음

**원인 1 — 브라우저 캐시:**  
`index.html`이 `app.js?v=20260430j`를 참조 중이어서 신규 코드가 배포돼도 다른 컴퓨터가 구버전을 캐시에서 로드.  
→ 버전 쿼리를 `?v=20260511a`로 올려 강제 재다운로드.

**원인 2 — `setupTabs()` 누락:**  
탭 버튼 클릭은 `setupTabs()` 함수가 처리하는데, 여기서 `renderMautoTab()`만 호출하고 `loadMautoDataRemote()`가 없었음.  
`switchTab()` (원격 로드 코드 위치)는 대시보드 카드 클릭·초기화에서만 호출되므로 탭 버튼 클릭과 무관.  
→ `setupTabs()` 내 mauto 분기에 `loadMautoDataRemote()` 호출 추가.

**원인 3 — 응답 형식 미인식:**  
`loadMautoDataRemote()`가 `body.data` 형식만 허용. Apps Script가 다른 형식(`body.rows`, 직접 객체 등)으로 반환 시 항상 `null` 반환.  
→ `body.data` → `body.rows` → `body.funds` 직접 순서로 다중 형식 처리.

**원인 4 — 빈 데이터 덮어쓰기:**  
원격 로드가 `null`이면 로컬(빈값) 데이터를 구글시트에 업로드하여 기존 데이터를 지웠음.  
→ 로컬에 실제 데이터가 있을 때만 업로드하도록 수정.

**수정 파일:** `app.js` (setupTabs, loadMautoDataRemote, switchTab), `index.html` (버전 쿼리)  
**커밋:** `d8dbf47`, `d6cf99c`, `57d9857`, `7d153e7`

---

### 2026-04-30 (2): 미수금/미지급 그룹 행 금액 셀 짙은 색 제거

**증상:** 미수금 그룹 집계 행(말일, 60일 등) 금액 셀이 짙은 녹색/파란색 배경 → 숫자 안 보임

**원인:** `style.css`에 남아있던 구 스타일:
- `.group-summary-cell { background: #1e3a8a }` (진한 파랑 — 미지급)
- `.group-total-cell { background: #172554 }` (거의 검정 — 미지급)
- `.rcv-group-header .group-summary-cell { background: #064e3b }` (짙은 녹색 — 미수금)
- `.rcv-group-header .group-total-cell { background: #022c22 }` (거의 검정 — 미수금)

**수정 (`style.css`):**
- 미지급: `#dbeafe` / `#bfdbfe` (연한 파랑) + `#1e40af` / `#1e3a8a` 텍스트
- 미수금: `#dcfce7` / `#bbf7d0` (연한 녹색) + `#14532d` / `#065f46` 텍스트
- `.group-summary-cell.year-summary-column`: `#bfdbfe` / `#1e3a8a`
- commit: `0269783`

---

### 2026-04-30 (1): 미수금/미지급 표 비주얼 전면 개편 + 버그 수정

**변경 내용:**
1. **sticky 헤더 전면 재작성** — `border-collapse: collapse` → `separate; border-spacing:0`
   - 연도 행: `top:0; height:34px` (CSS 고정, JS 측정 제거)
   - 월 행: `top:34px; height:30px` (CSS 고정)
   - 코너 셀: `z-index:20`, 고정 양방향
   - `fixStickyHeaderOffsets()` 함수 삭제 (JS 측정 방식 제거)
2. **연도 기준 컬럼 음영** — index 기반 → year-index 기반으로 변경
   - `mkcls(mk)` (미수금), `dkcls(dk)` (미지급): 같은 연도 = 같은 색
3. **결제일 배지 미정 버그 수정** — 날짜 설정해도 "미정"으로 표시되던 문제
   - `nextStatus = nextPlanValue ? "예정" : "미정"` (기존: 항상 "미정")
   - `planLabel`: `cellPlanValue` 먼저 체크 후 `isMijeong` 체크 순서로 변경
4. **그룹 헤더 색상 경량화** — 진한 색 → 파스텔 그린/블루 + 좌측 3px 액센트 선
5. **컬럼 줄무늬 제거** → 행 단위 zebra striping으로 통일
- commit: `188c9de`

---

### 2026-04-27 (4): 가용자금 탭 요약 카드 추가 + 구매자금 B2B 참고로 이동

- 가용자금 탭 상단에 요약 카드 그리드 추가 (①계좌 / ②B2B사용가능 / ③전자채권 / ④전자어음 / 합계)
- `grandTotal = ①+②+③+④` (구매자금 제외) — `recalcAvailableFundsSummary` 업데이트
- 구매자금 사용가능 업체 섹션 → ② B2B 대출 섹션 내부 "참고" 로 이동
- 홈 대시보드 가용자금 카드도 grandTotal 표시로 변경
- `style.css`: `.funds-summary-grid`, `.fsc` 계열 스타일 추가

---

### 2026-04-27 (3): 홈 상세 내역 제거 + 탭 순서 변경

- 홈 대시보드에서 "현금 계좌 내역", "대출·채권·어음" 상세 박스 제거
- 탭 순서 변경: 홈 → 가용자금 → 미수금 → 미지급 → 고정지출 → 대사 (`index.html`)

---

### 2026-04-27 (2): 가용자금 B2B/전자채권 합계 0원 표시 수정

**증상:** B2B 대출, 전자채권 섹션에 붙여넣기해도 합계 금액이 0원으로 표시됨

**원인:** `parseFundsPaste`가 헤더명을 완전 일치(===)로만 찾음.
사용자 엑셀의 금액 컬럼명이 "합계" 대신 "실행금액", "대출잔액", "채권금액" 등 다른 이름이면
`idx === -1` → `raw === ""` → `parseNum("") === 0` → 0으로 저장됨.

**수정 내용 (`parseFundsPaste` 함수):**
1. 헤더 매칭에 공백 제거 후 비교 추가 (공백 포함 헤더 허용)
2. 숫자 필드 헤더 매칭 실패 시 → 행의 마지막 미사용 숫자 컬럼으로 자동 fallback
   - B2B 합계, 전자채권 합계 등 금액이 마지막 컬럼인 경우를 자동 처리

**주의:** 엑셀 헤더와 상관없이 마지막 숫자 컬럼이 금액으로 사용됨.
금액 컬럼이 마지막이 아닌 형식이면 헤더를 정확히 맞춰야 함.

---

### 2026-04-27 (1): 미지급 로드 실패 + 가용자금 안 열림 + 느림 수정

**증상:**
- 미지급 탭: `시트 로드 실패: Cannot read properties of undefined (reading 'reduce')`
- 가용자금 탭: 열리지 않음
- 앱 초기 로딩이 매우 느림

**근본 원인:**
`parseAvailableFunds()` 함수가 잘못된 속성명으로 객체를 반환함:
- `purchaseLoans` (틀림) → `purchaseVendors` 이어야 함
- `b2bLoans`, `eNotes` 누락
- 행 객체 키도 `name/value`로 잘못됨 (렌더링은 `bank/balance`, `date/amount` 등 기대)

이 상태에서 `recalcAvailableFundsSummary()` → `availableFunds.b2bLoans.reduce()` 호출 시
`undefined.reduce()` TypeError 발생 → `loadSheetPayables` catch에 잡혀 미지급 에러로 표시됨.

**수정 내용 (`app.js`):**
1. `parseAvailableFunds()`: 올바른 속성명 + 올바른 행 구조로 수정
   - accounts: `{ bank, accountNo, balance }`
   - purchaseVendors: `{ date, name, amount }`
   - eBonds: `{ expiry, client, receiptDate, amount }`
   - b2bLoans: `[]`, eNotes: `[]` 추가
2. `recalcAvailableFundsSummary()`: 모든 배열 접근에 `|| []` 방어 fallback 추가
3. `renderAvailableFunds()`: 모든 `af.*` 배열 접근에 `|| []` 추가
4. `renderDashboard()`: 모든 `availableFunds.*` 배열 접근에 `|| []` 추가
5. `saveAvailableFundsLocal()`: 저장 시 `|| []` fallback 추가
6. `loadSheetPayables()`: **2단계 로딩**으로 성능 개선
   - 1단계: raw 미지급 + 업체마스터만 먼저 로드 → diff 없으면 즉시 화면 표시
   - 2단계: 결제계획 + 이력 백그라운드 로드 → 재적용

---

## ⚠️ 미해결 이슈

### 거래처코드 앞자리 0 소실 문제
- ~~Excel 파싱 시 `00101` → `101`로 저장됨~~
- **2026-06-11 수정 완료**: `parseXlsToRows`에서 `cell.w` 체크로 서식 텍스트 보존 (커밋 `780f29b`)
- **⚠️ 이미 업로드된 계정별원장 데이터는 재업로드 필요**
- code.gs 서버측 `setNumberFormat("@")` 적용은 2026-05-12 완료 상태 유지

### 업체마스터 중복 문제
- Google Sheets `업체마스터` 시트에 동일 업체가 중복으로 쌓이고 있음
- 원인: `vendorMasterImportButton`으로 업로드 시 기존 중복 체크 로직 미흡
- **다음 작업**: 중복 감지(사업자번호/업체명 기준) + 정리 기능 추가 필요

---

## 배포 방법

```bash
# 1. 변경 후 커밋
git add index.html app.js style.css
git commit -m "변경 내용 설명"

# 2. push (GitHub Pages 자동 반영)
git pull origin main --no-edit
git push origin main
```

> ⚠️ OneDrive 저장소 특성상 한글 파일명(.xls 등)이 git에서 "deleted"로 오인식됨.  
> `git restore .` 로 초기화 후 코드 파일만 선택해서 `git add` 할 것.

---

## 공통 스타일 상수

```js
// 색상 변수 (style.css :root)
--bg: #f5f7fb
--surface: #ffffff
--accent: #2563eb   // 파랑
--danger: #dc2626   // 빨강
--success: #16a34a  // 초록
--muted: #6c7a92
```

---

## 주요 전역 함수

| 함수 | 설명 |
|------|------|
| `rerenderAll()` | 전체 재렌더 (홈/미수금/미지급/고정지출/가용자금) |
| `renderAvailableFunds()` | 가용자금 탭 렌더 |
| `applyFundsPaste(sectionId)` | 붙여넣기 데이터 파싱 및 저장 |
| `recalcAvailableFundsSummary()` | 가용자금 합계 재계산 |
| `saveAvailableFundsLocal()` | localStorage 저장 |
| `loadAvailableFundsLocal()` | localStorage 로드 |
| `switchTab(tabId)` | 탭 전환 |
| `formatNumber(n)` | 숫자 → 한국식 콤마 포맷 |
