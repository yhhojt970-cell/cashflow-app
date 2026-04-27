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
| `receivables` | 미수금 | Google Sheets `raw` 시트에서 로드 |
| `payables` | 미지급 | Google Sheets `미지급_raw` 시트에서 로드 |
| `fixed` | 고정지출 | Google Sheets `고정지출` 시트에서 로드 |
| `daesa` | 대사 | 입출금 매칭 |
| `funds` | 가용자금 | **Excel 붙여넣기 입력** (localStorage 저장) |

---

## Google Sheets 연동

```js
const SHEET_SPREADSHEET_ID = "1VxYrCD3eZr5PpTORFPCEQPfWM5QSr-tNFNnc_W1C5qM";
const SHEET_APP_SCRIPT_URL = "https://script.google.com/macros/s/..."; // Apps Script URL
```

- 데이터 조회: gviz 방식 우선 → Apps Script fallback
- 업체마스터: Google Sheets `업체마스터` 시트 (⚠️ 중복 문제 미해결 — 아래 참고)

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
