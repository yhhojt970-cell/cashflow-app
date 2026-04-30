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
- Excel 파싱 시 `00101` → `101`로 저장됨
- **원인**: XLSX.js가 숫자처럼 보이는 셀을 number 타입으로 파싱 → 앞자리 0 소실
- **영향 범위**: `matchVendorEntry()` 코드 매칭, 미수금/미지급/대사 탭 거래처 연결 전반
- **수정 방향**: 파서에서 코드 컬럼은 `String(val).padStart(원본길이)` 또는 `{t:'s'}` 강제 적용
- **다음 작업**: `parseLedgerFile`, `parseTaxInvoiceFile`, `parseVendorMasterFile` 내 코드 컬럼 파싱 수정

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
