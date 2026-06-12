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
