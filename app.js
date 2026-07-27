const partners = [];

let receivables = [];

let payables = [];

const SHEET_SPREADSHEET_ID = "1VxYrCD3eZr5PpTORFPCEQPfWM5QSr-tNFNnc_W1C5qM";
const SHEET_NAME_PAYABLES = "미지급_raw";
const SHEET_APP_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbw9T3kGOQ5xPZ2wwy0Np0LSt-mHoudhvN39Zv2KNimE5ORKKEd_mghZXHua1D_i6LVF/exec"; // Apps Script WebApp URL을 넣으면 시트 데이터를 자동으로 불러옵니다.
const API_TOKEN_STORAGE_KEY = "receivable-payable-webapp.api-token.v1";
const PAYABLES_LOCAL_STATE_KEY = "receivable-payable-webapp.payables-state.v1";
const GROUP_ORDER_KEY = "receivable-payable-webapp.group-order.v1";
const VENDOR_MEMO_KEY = "receivable-payable-webapp.vendor-memos.v1";

let vendorMemos = {}; // { [normalizedCode]: { common: "", payables: "", receivables: "" } }
const MASTER_SHEET_NAME = "업체마스터";
const PLAN_SHEET_NAME = "결제계획";
const HISTORY_SHEET_NAME = "결제이력";
const FIXED_EXPENSES_SHEET_NAME = "고정지출";
const AVAILABLE_FUNDS_SHEET_NAME = "가용자금";
const PAYABLES_SYNC_DEBOUNCE_MS = 700;
const WOORI_TRANSFER_TEMPLATE_PATH = "우리은행 이체 양식.xlsx";
const DEFAULT_SENDER_ACCOUNT_DISPLAY = "미래오토메이션(주)";

// ── 미수금 상수 ─────────────────────────────────────────────
const SHEET_NAME_RECEIVABLES = "raw";
const MANAGER_MASTER_SHEET_NAME = "담당자";

const RECEIVABLE_MANAGER_EMAIL_MAP = {
  "장운기": "jug@mauto.co.kr", "여희정": "yhj@mauto.co.kr", "김도연": "kdy@mauto.co.kr",
  "남예린": "nyr@mauto.co.kr", "오성철": "osc@mauto.co.kr", "장재영": "jjy@mauto.co.kr",
  "김태홍": "kth@mauto.co.kr", "박희선": "phs@mauto.co.kr", "구예솔": "kys@mauto.co.kr",
  "배지혜": "bjh@mauto.co.kr", "임연하": "lyh@mauto.co.kr",
};
const RECEIVABLE_ABSENCE_CHAIN = [
  { name: "박희선", email: "phs@mauto.co.kr" },
  { name: "김도연", email: "kdy@mauto.co.kr" },
  { name: "장운기", email: "jug@mauto.co.kr" },
];
const RECEIVABLE_CC_OPTIONS = [
  { name: "여희정", email: "yhj@mauto.co.kr" }, { name: "구예솔", email: "kys@mauto.co.kr" },
  { name: "김도연", email: "kdy@mauto.co.kr" }, { name: "장운기", email: "jug@mauto.co.kr" },
  { name: "박희선", email: "phs@mauto.co.kr" }, { name: "배지혜", email: "bjh@mauto.co.kr" },
  { name: "임연하", email: "lyh@mauto.co.kr" }, { name: "오성철", email: "osc@mauto.co.kr" },
  { name: "장재영", email: "jjy@mauto.co.kr" }, { name: "김태홍", email: "kth@mauto.co.kr" },
];
const RECEIVABLE_TEST_RECIPIENTS = [
  { name: "여희정", email: "yhj@mauto.co.kr" },
  { name: "구예솔", email: "kys@mauto.co.kr" },
];
const RECEIVABLE_DEPT_HEAD = { name: "김도연", email: "kdy@mauto.co.kr" };
const RECEIVABLE_CEO = { name: "장운기", email: "jug@mauto.co.kr" };

let fixedExpenses = [];
let cashflowTimelineMode = "daily";
const AVAILABLE_FUNDS_LOCAL_KEY = "cashflow-app.available-funds-v2";
const FIXED_LOCAL_KEY = "cashflow-app.fixed-v1";
const B2B_TOTAL_LIMIT = 500000000; // 총대출액 5억 고정
const MAUTO_LOCAL_KEY = "cashflow-app.mauto-v1";
const MAUTO_FIXED_ACCOUNTS = [
  { bankAccount: "국민(415310)", bank: "국민", accountNo: "415310" },
  { bankAccount: "부산(008320)", bank: "부산", accountNo: "008320" },
];

let mautoData = {
  funds: MAUTO_FIXED_ACCOUNTS.map(a => ({ ...a, amount: 0 })),
  receivables: [],
  payables: [],
  fixed: [],
};
let mautoClassifiedRows = []; // Phase 2: 입출금 분류 결과 (재빌드 캐시)
const MAUTO_CLASSIFIED_KEY = "mauto-classified-rows-v1";
// 공유 저장 시 보낼 필드만 추출 (원본 파일 데이터 제외)
const CLASSIFIED_SHARE_FIELDS = ["_txKey","date","_memo","_memo2","debit","credit","거래처명","구분","excluded","매칭근거","savedAt"];
let _classifiedSaveTimer = null;
function saveClassifiedRows() {
  try { localStorage.setItem(MAUTO_CLASSIFIED_KEY, JSON.stringify(mautoClassifiedRows)); } catch (_) {}
  // 구글시트 debounce 저장 (3초)
  if (SHEET_APP_SCRIPT_URL) {
    clearTimeout(_classifiedSaveTimer);
    _classifiedSaveTimer = setTimeout(() => {
      const rows = mautoClassifiedRows
        .filter(r => r._txKey && r.거래처명)
        .map(r => {
          const out = {};
          CLASSIFIED_SHARE_FIELDS.forEach(f => { out[f] = r[f] ?? ""; });
          // 안전장치: 금액이 비었으면 _txKey(date|time|credit|debit|memo)에서 복구해 0금액 서버 오염 방지
          const p = String(r._txKey || "").split("|");
          if (!(Number(out.credit) > 0) && Number(p[2]) > 0) out.credit = Number(p[2]);
          if (!(Number(out.debit)  > 0) && Number(p[3]) > 0) out.debit  = Number(p[3]);
          out.savedAt = out.savedAt || new Date().toISOString().slice(0,19);
          return out;
        });
      if (rows.length) postSheetWebApp("upsertClassifiedRows", { rows }).catch(() => {});
    }, 3000);
  }
}
function loadClassifiedRows() {
  try {
    const raw = localStorage.getItem(MAUTO_CLASSIFIED_KEY);
    mautoClassifiedRows = raw ? JSON.parse(raw) : [];
  } catch (_) { mautoClassifiedRows = []; }
}

// Phase 2 저장모델전환: 불변(원본) / 사용자 영역 분리
const MAUTO_SOURCE_FILES_KEY = "mauto-source-files-v1";  // 파일 단위 원본 거래
const MAUTO_USER_EDITS_KEY   = "mauto-user-edits-v1";   // 거래키별 사용자 수정
let mautoSourceFiles = {}; // { fileKey: { filename, savedAt, isMigration?, rows[] } }
let mautoUserEdits   = {}; // { txKey: { 거래처명, 구분, excluded, isOverride, 매칭근거 } }
let _sourceSaveTimer = null;
const SOURCE_SHARE_FIELDS = ["_txKey","fileKey","filename","date","time","_memo","_memo2","_bank","_account","credit","debit"];
// 은행 파싱 행은 언더바 필드(_date/_time/_credit/_debit)를 쓰므로, 공유 저장 시 언더바 값을 매핑한다.
// (이 매핑이 없으면 credit/debit/date/time이 전부 빈값으로 저장돼 다른 컴퓨터 재빌드 시 금액이 0이 됨)
const SOURCE_FIELD_ALIAS = { date: "_date", time: "_time", credit: "_credit", debit: "_debit" };
function saveSourceFiles() {
  try { localStorage.setItem(MAUTO_SOURCE_FILES_KEY, JSON.stringify(mautoSourceFiles)); } catch (_) {}
  if (SHEET_APP_SCRIPT_URL) {
    clearTimeout(_sourceSaveTimer);
    _sourceSaveTimer = setTimeout(() => {
      const rows = Object.values(mautoSourceFiles).flatMap(f =>
        (f.rows || []).filter(r => r._txKey).map(r => {
          const out = {};
          SOURCE_SHARE_FIELDS.forEach(k => {
            let v = r[k];
            if ((v === undefined || v === null || v === "") && SOURCE_FIELD_ALIAS[k]) v = r[SOURCE_FIELD_ALIAS[k]];
            out[k] = v ?? "";
          });
          // fileKey/filename은 행 객체에 없으므로 파일 그룹명으로 채운다 (안 그러면 다른 컴퓨터에서 "원격 로드"로 표시됨)
          out.fileKey = f.filename || out.fileKey || "";
          out.filename = f.filename || out.filename || "";
          return out;
        })
      );
      if (rows.length) postSheetWebApp("upsertMautoSourceRows", { rows }).catch(() => {});
    }, 3000);
  }
}

// 원격/저장된 소스 행의 언더바 필드(_date/_time/_credit/_debit)를 채운다.
// 값이 비어있으면 _txKey(_date|_time|_credit|_debit|_memo#seq)에서 복구한다.
function normalizeSourceRow(r) {
  const p = String(r._txKey || "").split("|");
  const pick = (u, s, idx) => {
    if (r[u] !== undefined && r[u] !== null && r[u] !== "") return r[u];
    if (r[s] !== undefined && r[s] !== null && r[s] !== "") return r[s];
    return p[idx] !== undefined ? p[idx] : "";
  };
  r._date   = pick("_date", "date", 0) || "";
  r._time   = pick("_time", "time", 1) || "";
  r._credit = Number(pick("_credit", "credit", 2)) || 0;
  r._debit  = Number(pick("_debit",  "debit",  3)) || 0;
  return r;
}
function loadSourceFiles()  { try { const r = localStorage.getItem(MAUTO_SOURCE_FILES_KEY); mautoSourceFiles = r ? JSON.parse(r) : {}; } catch (_) { mautoSourceFiles = {}; } }
function saveUserEdits()    { try { localStorage.setItem(MAUTO_USER_EDITS_KEY, JSON.stringify(mautoUserEdits));   } catch (_) {} }
function loadUserEdits()    { try { const r = localStorage.getItem(MAUTO_USER_EDITS_KEY);   mautoUserEdits   = r ? JSON.parse(r) : {}; } catch (_) { mautoUserEdits = {};   } }

// ── 미래 자료업로드 소스 파일 보관 (파일 단위 교체 + 재빌드 모델) ──
const MIRAE_SOURCE_TAX_KEY    = "mirae-source-tax-v1";
const MIRAE_SOURCE_LEDGER_KEY = "mirae-source-ledger-v1";
const MIRAE_SOURCE_BIZ_KEY    = "mirae-source-biz-v1";
let miraeTaxSources    = {}; // { fileKey: { filename, savedAt, rows[] } }
let miraeLedgerSources = {}; // { fileKey: { filename, savedAt, ledgerType, rows[] } }
let miraeBizSources    = {}; // { fileKey: { filename, savedAt, rows[] } }
function saveMiraeSource(stKey, data) { try { localStorage.setItem(stKey, JSON.stringify(data)); } catch (_) {} }
function loadMiraeSource(stKey) { try { const r = localStorage.getItem(stKey); return r ? JSON.parse(r) : {}; } catch (_) { return {}; } }
function hasMiraeSources() {
  return Object.keys(miraeTaxSources).length > 0 ||
         Object.keys(miraeLedgerSources).length > 0 ||
         Object.keys(miraeBizSources).length > 0;
}

// 섹션키 → 저장된 파일 목록
function getMiraeSectionFiles(key) {
  if (key === "taxInvoice") return Object.values(miraeTaxSources);
  if (key === "dailySales") return Object.values(miraeBizSources);
  const typeMap = { ledgerSales: "매출", ledgerPurchase: "매입", ledgerPayable: "미지급" };
  const lt = typeMap[key];
  return lt ? Object.values(miraeLedgerSources).filter(f => f.ledgerType === lt) : [];
}

// 파싱된 rows를 소스 저장소에 보관
function saveMiraeSectionFile(key, filename, rows) {
  const lt = { ledgerSales: "매출", ledgerPurchase: "매입", ledgerPayable: "미지급" }[key] || null;
  const entry = { filename, savedAt: new Date().toISOString(), rows };
  if (key === "taxInvoice") {
    miraeTaxSources[filename] = entry;
    saveMiraeSource(MIRAE_SOURCE_TAX_KEY, miraeTaxSources);
  } else if (key === "dailySales") {
    miraeBizSources[filename] = entry;
    saveMiraeSource(MIRAE_SOURCE_BIZ_KEY, miraeBizSources);
  } else if (lt) {
    miraeLedgerSources[filename] = { ...entry, ledgerType: lt };
    saveMiraeSource(MIRAE_SOURCE_LEDGER_KEY, miraeLedgerSources);
  }
}

// 소스 파일 삭제
function deleteMiraeSectionFile(key, filename) {
  if (key === "taxInvoice") { delete miraeTaxSources[filename]; saveMiraeSource(MIRAE_SOURCE_TAX_KEY, miraeTaxSources); }
  else if (key === "dailySales") { delete miraeBizSources[filename]; saveMiraeSource(MIRAE_SOURCE_BIZ_KEY, miraeBizSources); }
  else { delete miraeLedgerSources[filename]; saveMiraeSource(MIRAE_SOURCE_LEDGER_KEY, miraeLedgerSources); }
}

// ── 엠오토 미수미지급 제외 거래처 (미수금·미지급 별도 관리) ──
const MAUTO_EXCLUDE_KEY_RCV = "mauto-exclude-vendors-rcv-v1";
const MAUTO_EXCLUDE_KEY_PAY = "mauto-exclude-vendors-pay-v1";
let mautoExcludeVendorsRcv = [];
let mautoExcludeVendorsPay = [];
function saveMautoExcludeVendors(side) {
  try {
    if (side === "rcv") localStorage.setItem(MAUTO_EXCLUDE_KEY_RCV, JSON.stringify(mautoExcludeVendorsRcv));
    else              localStorage.setItem(MAUTO_EXCLUDE_KEY_PAY, JSON.stringify(mautoExcludeVendorsPay));
  } catch (_) {}
  _scheduleMautoRemoteSave();
}
function loadMautoExcludeVendors() {
  try {
    const r = localStorage.getItem(MAUTO_EXCLUDE_KEY_RCV);
    mautoExcludeVendorsRcv = r ? JSON.parse(r) : [];
    const p = localStorage.getItem(MAUTO_EXCLUDE_KEY_PAY);
    mautoExcludeVendorsPay = p ? JSON.parse(p) : [];
    // 구버전 마이그레이션
    const old = localStorage.getItem("mauto-exclude-vendors-v1");
    if (old) {
      const parsed = JSON.parse(old);
      if (parsed.length && !mautoExcludeVendorsRcv.length && !mautoExcludeVendorsPay.length) {
        mautoExcludeVendorsRcv = [...parsed];
        mautoExcludeVendorsPay = [...parsed];
        saveMautoExcludeVendors("rcv");
        saveMautoExcludeVendors("pay");
      }
      localStorage.removeItem("mauto-exclude-vendors-v1");
    }
  } catch (_) { mautoExcludeVendorsRcv = []; mautoExcludeVendorsPay = []; }
}
function isArRecapExcluded(vendorName, side) {
  const list = side === "rcv" ? mautoExcludeVendorsRcv : mautoExcludeVendorsPay;
  if (!list.length) return false;
  const norm = normalizeVendorName(vendorName);
  return list.some(ex => normalizeVendorName(ex) === norm || ex === vendorName);
}

// ── 엠오토 세금계산서 소스 파일 보관 (국세청 양식, 파일 단위 교체 + 재빌드) ──
const MAUTO_TAX_SOURCE_KEY = "mauto-tax-source-v1";
let mautoTaxSources  = {}; // { [filename]: { filename, sideType, savedAt, rows[] } }
let mautoTaxInvoices = []; // 재빌드 캐시
let mautoFixedRules  = null; // null=미로드, []=로드완료(항목없음), [...]=로드완료(항목있음)
const MAUTO_FIXED_CHECKED_KEY = "mauto-fixed-checked-v1";
let mautoFixedChecked = {}; // { "YYYY-MM||YYYY-MM-DD": true/false }
const MAUTO_FIXED_AMOUNT_KEY = "mauto-fixed-amount-overrides-v1";
let mautoFixedAmountOverrides = {}; // { "YYYY-MM||거래처명||예정일": amount }
let mautoVatView = false;
let mautoVatMode = "반기";
let mautoVatYear = new Date().getFullYear();
let mautoPayViewMode = "ym"; // "ym" | "vendor"
let mautoToolsOpen = false; // 엠오토 상단 도구영역(제목~카드 사이) 접기 상태 (기본 접힘)
function loadFixedChecked() {
  try { const s = localStorage.getItem(MAUTO_FIXED_CHECKED_KEY); mautoFixedChecked = s ? JSON.parse(s) : {}; } catch(_) { mautoFixedChecked = {}; }
}
function saveFixedChecked() {
  try { localStorage.setItem(MAUTO_FIXED_CHECKED_KEY, JSON.stringify(mautoFixedChecked)); } catch(_) {}
  _scheduleMautoRemoteSave();
}
function loadFixedAmountOverrides() {
  try { const s = localStorage.getItem(MAUTO_FIXED_AMOUNT_KEY); mautoFixedAmountOverrides = s ? JSON.parse(s) : {}; } catch(_) { mautoFixedAmountOverrides = {}; }
}
function saveFixedAmountOverrides() {
  try { localStorage.setItem(MAUTO_FIXED_AMOUNT_KEY, JSON.stringify(mautoFixedAmountOverrides)); } catch(_) {}
  _scheduleMautoRemoteSave();
}
function applyFixedAmountOverrides(monthData) {
  (monthData || []).forEach(({ ym, items }) => {
    items.forEach(item => {
      const key = `${ym}||${item.거래처명}||${item.예정일 || "0"}`;
      if (mautoFixedAmountOverrides[key] !== undefined) {
        item.예정금액 = mautoFixedAmountOverrides[key];
        item.예정금액출처 = "override";
      }
    });
  });
  return monthData;
}
function calcFixedCheckedTotal(monthData) {
  const today = new Date();
  const todayYM = `${today.getFullYear()}-${String(today.getMonth()+1).padStart(2,"0")}`;
  let total = 0;
  (monthData || []).forEach(({ ym, items }) => {
    const byDate = {};
    items.forEach(item => {
      const d = item.예정결제일?.date || "미정";
      if (!byDate[d]) byDate[d] = { amt: 0, items: [] };
      byDate[d].amt += item.예정금액 || 0;
      byDate[d].items.push(item);
    });
    Object.entries(byDate).forEach(([date, { amt, items: grpItems }]) => {
      // 날짜 그룹 내 모든 항목이 완료(✓)이면 예정금액에서 제외
      if (grpItems.every(i => i.status === "완료")) return;
      const key = `${ym}||${date}`;
      const isChecked = mautoFixedChecked[key] !== undefined ? mautoFixedChecked[key] : (ym === todayYM);
      if (isChecked) total += amt;
    });
  });
  return total;
}
let _mautoTaxSaveTimer = null;
function saveMautoTaxSource() {
  try { localStorage.setItem(MAUTO_TAX_SOURCE_KEY, JSON.stringify(mautoTaxSources)); } catch (_) {}
  // 구글시트 debounce 저장 (3초) — 컴퓨터 간 공유
  if (SHEET_APP_SCRIPT_URL) {
    clearTimeout(_mautoTaxSaveTimer);
    _mautoTaxSaveTimer = setTimeout(() => {
      const rows = mautoTaxInvoices.filter(r => r._row_key);
      if (rows.length) postSheetWebApp("upsertMautoTaxInvoices", { rows }).catch(() => {});
    }, 3000);
  }
}
function loadMautoTaxSource() {
  try { const r = localStorage.getItem(MAUTO_TAX_SOURCE_KEY); mautoTaxSources = r ? JSON.parse(r) : {}; }
  catch (_) { mautoTaxSources = {}; }
}
// 구글시트 엠오토_세금계산서 시트에서 로드 → 로컬에 없는 행 병합 (컴퓨터 간 공유)
async function loadMautoTaxRemote() {
  if (!SHEET_APP_SCRIPT_URL) return;
  try {
    const res = await fetchSheetWebApp({ action: "getMautoTaxInvoices" });
    const remote = (res && (res.rows || res.data)) || [];
    if (!remote.length) {
      // GSheets 비어있고 로컬에 데이터 있으면 → 로컬 → GSheets 자동 업로드 (최초 1회)
      if (mautoTaxInvoices.length) saveMautoTaxSource();
      return;
    }
    const localKeys = new Set(mautoTaxInvoices.map(r => r._row_key).filter(Boolean));
    const newRows = remote.filter(r => r._row_key && !localKeys.has(r._row_key));
    if (!newRows.length) return;
    mautoTaxInvoices = [...mautoTaxInvoices, ...newRows];
    // 가상 소스 항목(원격 로드 표시용)이 없으면 생성
    if (!mautoTaxSources["__remote__"]) {
      mautoTaxSources["__remote__"] = { filename: "원격 로드", sideType: "both", savedAt: new Date().toISOString(), rows: [] };
    }
    mautoTaxSources["__remote__"].rows = newRows;
    mautoTaxSources["__remote__"].savedAt = new Date().toISOString();
    try { localStorage.setItem(MAUTO_TAX_SOURCE_KEY, JSON.stringify(mautoTaxSources)); } catch (_) {}
    renderMautoTab();
  } catch (_) {}
}

// 구글시트 엠오토_소스 시트에서 입출금 원본 행 로드 → 로컬에 없는 파일 그룹 병합 (컴퓨터 간 공유)
async function loadMautoSourceRemote() {
  if (!SHEET_APP_SCRIPT_URL) return;
  try {
    const res = await fetchSheetWebApp({ action: "getMautoSourceRows" });
    const remote = (res && (res.rows || res.data)) || [];
    if (!remote.length) {
      // GSheets 비어있고 로컬에 데이터 있으면 → 로컬 → GSheets 자동 업로드 (최초 1회)
      const localRows = Object.values(mautoSourceFiles).flatMap(f => f.rows || []);
      if (localRows.length) saveSourceFiles();
      return;
    }
    const localKeys = new Set(
      Object.values(mautoSourceFiles).flatMap(f => (f.rows || []).map(r => r._txKey)).filter(Boolean)
    );
    const newRows = remote.filter(r => r._txKey && !localKeys.has(r._txKey)).map(normalizeSourceRow);
    if (!newRows.length) return;
    // 파일 그룹별로 묶어서 mautoSourceFiles에 추가
    const groups = {};
    newRows.forEach(r => {
      const key = r.fileKey || "__remote__";
      if (!groups[key]) groups[key] = { filename: r.fileKey || "원격 로드", savedAt: new Date().toISOString(), rows: [] };
      groups[key].rows.push(r);
    });
    Object.assign(mautoSourceFiles, groups);
    try { localStorage.setItem(MAUTO_SOURCE_FILES_KEY, JSON.stringify(mautoSourceFiles)); } catch (_) {}
    rebuildMautoRows();
    renderMautoTab();
  } catch (_) {}
}

function rebuildMautoTaxInvoices() {
  const seen = new Map();
  for (const src of Object.values(mautoTaxSources)) {
    for (const r of (src.rows || [])) {
      // 신형 포맷(구분 컬럼 없음)은 sideType으로 구분 설정
      if (!r["구분"]) r["구분"] = src.sideType === "매출" ? "매출" : "매입";
      const k = String(r._row_key || "").trim();
      if (!k || !seen.has(k)) seen.set(k || `__rnd_${Math.random()}`, r);
    }
  }
  mautoTaxInvoices = [...seen.values()];
}

let availableFunds = {
  accounts: [],       // [{bank, accountNo, balance}]
  b2bLoans: [],       // [{latestExpiry, execNo, finalExpiry, used}]
  purchaseVendors: [], // [{date, name, amount}]
  eBonds: [],         // [{expiry, client, receiptDate, amount}]
  eNotes: [],         // [{bank, client, receiptDate, expiry, amount}]
  // 대시보드용 요약 (계산값)
  summary: {
    totalAccountBalance: 0,
    totalPurchaseLoanBalance: 0,
    totalEBonds: 0,
    availableTotal: 0
  }
};

const filterState = {
  partner: "",
  year: "",
  month: "",
  status: "",
  search: "",
  groups: null, // null=전체, []=없음, [...]= 선택목록
  groupOrder: [], // 미지급 그룹 드래그 순서
};

const payablesGroupState = {
  collapsed: {},
  groupPaymentDates: {},
};

const rcvSortState = { key: "code", dir: "asc" };
const rcvGroupState = { order: [], filter: null }; // null=전체, []=없음, [...]= 선택

const payablesUiState = {
  lastEdited: null,
};

const paymentPlanUiState = {
  selectedPlanKeys: [],
};

const payablesYearCollapsed = {};

const receivableManagerState = {
  rows: [],
  map: new Map(), // codeNorm → { manager, email }
  lastFileName: "",
};

const payablesSyncState = {
  timeoutId: null,
  inFlight: false,
  pending: false,
  lastError: "",
};

const vendorMasterState = {
  rows: [],
  map: new Map(),
  importedRows: [],
  comparedRows: [],
  stats: null,
  lastFileName: "",
  saving: false,
  lastMessage: "",
};

const paymentHistoryState = {
  rows: [],
};

const payablePlanHistories = {}; // [sourceKey] -> array of history records


const elements = {
  partnerFilter: document.getElementById("partnerFilter"),
  yearFilter: document.getElementById("yearFilter"),
  monthFilter: document.getElementById("monthFilter"),
  statusFilter: document.getElementById("statusFilter"),
  searchInput: document.getElementById("searchInput"),
  groupFilterContainer: document.getElementById("groupFilterContainer"),
  vendorMasterImportButton: document.getElementById("vendorMasterImportButton"),
  vendorMasterFileInput: document.getElementById("vendorMasterFileInput"),
  vendorMasterPanel: document.getElementById("vendorMasterPanel"),
  summaryPanel: document.getElementById("summaryPanel"),
  receivables: document.getElementById("receivables"),
  payables: document.getElementById("payables"),
  fixed: document.getElementById("fixed"),
  mauto: document.getElementById("mauto"),
  tabButtons: [...document.querySelectorAll(".tab-button")],
};

function formatCurrency(value) {
  return formatNumber(value);
}

function formatNumber(value) {
  return Number(value || 0).toLocaleString("ko-KR");
}

function formatPayableCellNumber(value) {
  return Number(value || 0) === 0 ? "" : formatNumber(value);
}

function formatMonthKey(key) {
  if (!key) return "-";
  const parts = key.split("-");
  if (parts.length < 2) return key;
  return `${Number(parts[1])}월`;
}

function formatPlanShortLabel(value) {
  if (!value) return "";
  if (/^\d{4}-\d{2}-\d{2}$/.test(value)) {
    return value.slice(5).replace("-", "/");
  }
  return value;
}

function formatPlanLabel(value) {
  if (!value || value === "미정") return "미정";
  if (value === "보류") return "보류";
  if (value === "제외") return "제외";
  if (/^\d{4}-\d{2}-\d{2}$/.test(value)) {
    return value.slice(5).replace("-", "/");
  }
  return value;
}

function formatPlanLabel(value) {
  if (!value || value === "미정") return "미정";
  if (value === "보류") return "보류";
  if (value === "제외") return "제외";
  if (/^\d{4}-\d{2}-\d{2}$/.test(value)) {
    return value.slice(5).replace("-", "/");
  }
  return value;
}

function getPayableEffectivePaid(item) {
  return Number(item.paidOverride != null ? item.paidOverride : item.paid || 0);
}

function getPayableOutstanding(item) {
  // 앱 내 결제 기록이 없으면 ERP 잔액(balance)을 신뢰, 있으면 합계-paidOverride로 계산
  const appPaymentApplied = Number(item.paidOverride || 0) > Number(item.paid || 0);
  if (!appPaymentApplied && item.balance > 0) return item.balance;
  return Math.max(0, Number(item.purchase || 0) - getPayableEffectivePaid(item));
}

function normalizeVendorCode(value, minLength = 5) {
  const raw = String(value ?? "").trim();
  if (!raw) return "";
  const digitsOnly = raw.replace(/\D/g, "");
  if (!digitsOnly) return raw.toUpperCase();
  return digitsOnly.padStart(minLength, "0");
}

function normalizeBusinessNumber(value) {
  return String(value ?? "").replace(/\D/g, "");
}

// 거래처명 비교용 정규화: 법인 suffix·공백·괄호 안 영문병기 제거
function normalizeVendorName(name) {
  return String(name || "")
    .replace(/주식회사|유한회사|합자회사|합명회사/g, "")
    .replace(/\(주\)|\(유\)|\(합\)/g, "")
    .replace(/\([A-Za-z0-9\s&.]+\)/g, "")
    .replace(/[\s\-_]/g, "")
    .trim()
    .toLowerCase();
}

function normalizeDateValue(value) {
  if (!value) return "";
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    const year = value.getFullYear();
    const month = String(value.getMonth() + 1).padStart(2, "0");
    const day = String(value.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
  }
  const raw = String(value).trim();
  if (!raw) return "";
  const isoMatch = raw.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (isoMatch) {
    return `${isoMatch[1]}-${isoMatch[2]}-${isoMatch[3]}`;
  }
  const slashMatch = raw.match(/^(\d{4})[./](\d{1,2})[./](\d{1,2})$/);
  if (slashMatch) {
    return `${slashMatch[1]}-${String(slashMatch[2]).padStart(2, "0")}-${String(slashMatch[3]).padStart(2, "0")}`;
  }
  const parsed = new Date(raw);
  if (!Number.isNaN(parsed.getTime())) {
    const year = parsed.getFullYear();
    const month = String(parsed.getMonth() + 1).padStart(2, "0");
    const day = String(parsed.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
  }
  return raw;
}

function isFuzzySame(v1, v2, field) {
  const s1 = String(v1 || "").trim();
  const s2 = String(v2 || "").trim();

  // 1. 단순 비교 (둘 다 공백 상태이거나 문자열이 완전히 같은 경우)
  if (s1 === s2) return true;

  // 2. 엑셀 오류값 또는 불완전한 상태 제어: 신규 데이터(s2)가 오류 문자열이면 변경하지 않음 (기존 데이터 보존)
  if (s2.startsWith("#") || s2 === "undefined" || s2 === "null") return true;

  // 3. 필드별 특화 비교
  // 숫자 성격 필드 (사업자번호, 계좌번호, 전화번호, 거래처코드)
  if (field === "사업자번호" || field === "계좌번호" || field === "거래처코드_norm" || field === "전화번호") {
    const n1 = s1.replace(/\D/g, "").replace(/^0+/, "");
    const n2 = s2.replace(/\D/g, "").replace(/^0+/, "");
    if (n1 === n2) return true;  // 둘 다 all-zero("0" vs "0000000000" 등)도 동일로 처리
  }

  // 대표자명/주소 등 일반 텍스트: '-', '0' 인 경우 공백과 동일하게 취급
  if (field === "대표자명" || field === "주소" || field === "예금주") {
    const isBlank1 = !s1 || s1 === "-" || s1 === "0";
    const isBlank2 = !s2 || s2 === "-" || s2 === "0";
    if (isBlank1 && isBlank2) return true;
  }

  return false;
}

function preserveViewport(work) {
  const scrollX = window.scrollX;
  const scrollY = window.scrollY;
  const tableResponsive = elements.payables?.querySelector?.(".table-responsive");
  const tableScrollLeft = tableResponsive?.scrollLeft ?? 0;
  const tableScrollTop = tableResponsive?.scrollTop ?? 0;
  work();
  window.scrollTo(scrollX, scrollY);
  const nextTableResponsive = elements.payables?.querySelector?.(".table-responsive");
  if (nextTableResponsive) {
    nextTableResponsive.scrollLeft = tableScrollLeft;
    nextTableResponsive.scrollTop = tableScrollTop;
  }
}

function getUniqueSortedValues(items, key) {
  return [...new Set(items.map(item => item[key]).filter(Boolean))].sort((a, b) => a - b);
}

function getFilteredItems(items, section) {
  return items.filter(item => {
    if (filterState.partner && section !== "fixed" && item.code !== filterState.partner) {
      return false;
    }
    if (filterState.year && String(item.year) !== filterState.year) {
      return false;
    }
    if (filterState.month && Number(filterState.month) !== 0 && Number(item.month) !== Number(filterState.month)) {
      return false;
    }
    if (section === "payables" && filterState.groups !== null) {
      const dueGroup = getDueGroup(item);
      if (!filterState.groups.includes(dueGroup)) return false;
    }
    if (section === "receivables" && rcvGroupState.filter !== null) {
      if (!rcvGroupState.filter.includes(item.condition || "기타")) return false;
    }
    if (filterState.search) {
      const text = [item.name, item.code, item.memo, item.title, item.bank, item.category]
        .filter(Boolean)
        .join(" ")
        .toLowerCase();
      if (!text.includes(filterState.search.toLowerCase())) {
        return false;
      }
    }
    // Status filtering
    if (filterState.status === "completed") {
      const balance = section === "payables"
        ? getPayableOutstanding(item)
        : section === "receivables"
          ? Number(item.balance || 0)
          : (item.purchase || item.sales || item.amount || 0) - (item.paid || 0);
      if (balance !== 0) return false;
    } else if (filterState.status === "pending") {
      const balance = section === "payables"
        ? getPayableOutstanding(item)
        : section === "receivables"
          ? Number(item.balance || 0)
          : (item.purchase || item.sales || item.amount || 0) - (item.paid || 0);
      if (balance === 0) return false;
    } else if (filterState.status === "excluded") {
      const isExcluded = section === "payables"
        ? item.completionStatus === "제외"
        : section === "receivables"
          ? (item.condition === "제외" || (item.memo && item.memo.includes("제외")))
          : false;
      if (!isExcluded) return false;
    } else if (!filterState.status) {
      // Default view: hide completed items
      const balance = section === "payables"
        ? getPayableOutstanding(item)
        : section === "receivables"
          ? Number(item.balance || 0)
          : (item.purchase || item.sales || item.amount || 0) - (item.paid || 0);
      if (balance === 0 && item.completionStatus === "완료") return false;
    }

    return true;
  });
}

function renderFilterControls() {
  const years = [
    ...new Set([
      ...getUniqueSortedValues(receivables, "year"),
      ...getUniqueSortedValues(payables, "year"),
      ...getUniqueSortedValues(fixedExpenses, "year"),
    ]),
  ];
  const months = [
    ...new Set([
      ...getUniqueSortedValues(receivables, "month"),
      ...getUniqueSortedValues(payables, "month"),
      ...getUniqueSortedValues(fixedExpenses, "month"),
    ]),
  ];

  elements.yearFilter.innerHTML = `<option value="">전체</option>` +
    years.map(year => `<option value="${year}">${year}년</option>`).join("");
  elements.monthFilter.innerHTML = `<option value="">전체</option>` +
    months.map(month => `<option value="${String(month).padStart(2, "0")}">${month}월</option>`).join("");
  elements.statusFilter.innerHTML = `
    <option value="">전체 (완료 제외)</option>
    <option value="pending">미완료 / 지급 대기</option>
    <option value="completed">완료 / 지급 완료</option>
    <option value="excluded">제외 항목</option>
  `;

  elements.partnerFilter.addEventListener("change", event => {
    filterState.partner = event.target.value;
    rerenderAll();
  });
  elements.yearFilter.addEventListener("change", event => {
    filterState.year = event.target.value;
    rerenderAll();
  });
  elements.monthFilter.addEventListener("change", event => {
    filterState.month = event.target.value;
    rerenderAll();
  });
  elements.statusFilter.addEventListener("change", event => {
    filterState.status = event.target.value;
    rerenderAll();
  });
  elements.searchInput.addEventListener("input", event => {
    filterState.search = event.target.value.trim();
    rerenderAll();
  });

  renderGroupFilterControls();
}

function detectPayablesRawDiff(newParsedItems) {
  const savedMap = loadPayablesStateFromLocal();
  if (!Object.keys(savedMap).length) return []; // 처음 로드, diff 없음

  const savedByStableKey = {};
  Object.entries(savedMap).forEach(([srcKey, v]) => {
    if (v.stableKey) savedByStableKey[v.stableKey] = { ...v, srcKey };
  });

  const newStableKeys = new Set(newParsedItems.map(buildPayableStableKey));
  const diff = [];

  // 사라진 항목 (raw에서 제거됨)
  Object.entries(savedByStableKey).forEach(([sk, v]) => {
    if (!newStableKeys.has(sk) && v.completionStatus !== "완료") {
      const [code, year, month, group] = sk.split("||");
      diff.push({
        type: "removed", stableKey: sk, code, year, month, group,
        label: `${year}-${month} ${code} (${group})`, paidOverride: v.paidOverride
      });
    }
  });

  // 금액이 변경된 항목
  newParsedItems.forEach(item => {
    const sk = buildPayableStableKey(item);
    const prev = savedByStableKey[sk];
    if (!prev) return; // 신규
    const prevPurchase = Number(prev.purchase || 0);
    const newPurchase = Number(item.purchase || 0);
    if (prevPurchase && newPurchase && prevPurchase !== newPurchase) {
      diff.push({
        type: "changed", stableKey: sk,
        label: `${item.year}-${String(item.month).padStart(2, "0")} ${item.name} (${getDueGroup(item)})`,
        prevAmount: prevPurchase, newAmount: newPurchase
      });
    }
  });

  return diff;
}

function showPayablesRawDiffDialog(diff, onConfirm) {
  document.querySelector(".raw-diff-overlay")?.remove();
  const overlay = document.createElement("div");
  overlay.className = "raw-diff-overlay";

  const removedItems = diff.filter(d => d.type === "removed");
  const changedItems = diff.filter(d => d.type === "changed");

  overlay.innerHTML = `
    <div class="raw-diff-dialog">
      <div class="raw-diff-header">
        <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;">
          <h3 style="margin:0;">미지급_raw 변경 감지</h3>
          <button type="button" class="diff-email-btn" title="담당자에게 확인 요청 메일 발송"
            style="background:#1e40af;color:white;border:none;border-radius:6px;padding:5px 11px;font-size:13px;cursor:pointer;display:flex;align-items:center;gap:5px;">
            ✉ 이메일 발송
          </button>
        </div>
        <span class="raw-diff-sub">보류/계획이 지정된 항목 원본(구글 시트)에 서식 삭제나 금액 변경이 발생했습니다. 내역을 확인해 주세요.</span>
      </div>
      
      <div class="raw-diff-section">
        <details class="raw-diff-accordion">
          <summary class="raw-diff-accordion-header">
            <strong>📅 ${new Date().toLocaleString("ko-KR", { year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' })} 기준 변경 감지</strong>
            <span style="color:#ef4444; margin-left:8px; font-weight:600;">(총 ${diff.length}건)</span>
          </summary>
          <div class="raw-diff-accordion-body" style="margin-top: 10px; padding-left: 10px; border-left: 2px solid #e5e7eb;">
            ${removedItems.length ? `
              <div class="raw-diff-group" style="margin-bottom: 15px;">
                <div class="raw-diff-section-title removed-title" style="margin-bottom:5px;">사라진 항목 (${removedItems.length}건) — 완료 처리 추천</div>
                ${removedItems.map(d => `
                  <div class="raw-diff-row">
                    <span class="raw-diff-label">${escapeHtml(d.label)}</span>
                    <label class="raw-diff-check">
                      <input type="checkbox" class="diff-complete-chk" data-key="${escapeHtml(d.stableKey)}" checked />
                      완료로 표시
                    </label>
                  </div>`).join("")}
              </div>` : ""}
            
            ${changedItems.length ? `
              <div class="raw-diff-group">
                <div class="raw-diff-section-title changed-title" style="margin-bottom:5px;">금액 변경 항목 (${changedItems.length}건)</div>
                ${changedItems.map(d => `
                  <div class="raw-diff-row">
                    <span class="raw-diff-label">${escapeHtml(d.label)}</span>
                    <span class="raw-diff-amounts">
                      ${formatNumber(d.prevAmount)} → <strong>${formatNumber(d.newAmount)}</strong>
                    </span>
                  </div>`).join("")}
              </div>` : ""}
          </div>
        </details>
      </div>

      <div class="raw-diff-actions">
        <button type="button" class="diff-confirm-btn">확인 후 적용</button>
        <button type="button" class="diff-cancel-btn">상태 무시 (닫기)</button>
      </div>
    </div>
  `;
  document.body.appendChild(overlay);

  overlay.querySelector(".diff-confirm-btn").addEventListener("click", () => {
    const completeKeys = new Set(
      [...overlay.querySelectorAll(".diff-complete-chk:checked")].map(c => c.dataset.key)
    );
    overlay.remove();
    onConfirm(completeKeys);
  });
  overlay.querySelector(".diff-cancel-btn").addEventListener("click", () => overlay.remove());
  overlay.querySelector(".diff-email-btn").addEventListener("click", () => {
    openRawDiffEmailDialog(diff);
  });
}

function openRawDiffEmailDialog(diff) {
  document.querySelector(".diff-email-overlay")?.remove();

  const emailOverlay = document.createElement("div");
  emailOverlay.className = "raw-diff-overlay diff-email-overlay";

  const staffList = RECEIVABLE_CC_OPTIONS;

  emailOverlay.innerHTML = `
    <div class="raw-diff-dialog" style="max-width:440px;">
      <h3 style="margin-top:0;">확인 요청 메일 발송</h3>
      <p style="font-size:13px;color:#555;margin-bottom:12px;">
        미지급 데이터 변경사항 확인을 요청할 담당자를 선택하세요.
      </p>
      <div style="display:flex;flex-direction:column;gap:6px;margin-bottom:14px;">
        ${staffList.map(s => `
          <label style="display:flex;align-items:center;gap:8px;font-size:14px;cursor:pointer;">
            <input type="checkbox" class="diff-email-recipient" value="${escapeHtml(s.email)}"
              data-name="${escapeHtml(s.name)}"
              ${s.name === "김도연" ? "checked" : ""} />
            ${escapeHtml(s.name)} <span style="color:#888;font-size:12px;">${escapeHtml(s.email)}</span>
          </label>`).join("")}
      </div>
      <label style="display:flex;align-items:center;gap:6px;font-size:13px;margin-bottom:14px;">
        <input type="checkbox" id="diffEmailTestMode" />
        테스트 모드 (yhj@mauto.co.kr 로만 발송)
      </label>
      <div style="display:flex;gap:8px;justify-content:flex-end;">
        <button type="button" class="diff-email-cancel btn-secondary" style="padding:7px 16px;">취소</button>
        <button type="button" class="diff-email-send btn-primary" style="padding:7px 16px;">발송</button>
      </div>
      <p class="diff-email-status" style="margin-top:10px;font-size:13px;color:#1e40af;min-height:18px;"></p>
    </div>`;

  document.body.appendChild(emailOverlay);

  emailOverlay.querySelector(".diff-email-cancel").onclick = () => emailOverlay.remove();
  emailOverlay.querySelector(".diff-email-send").onclick = async () => {
    const checked = [...emailOverlay.querySelectorAll(".diff-email-recipient:checked")];
    if (!checked.length) { alert("수신자를 한 명 이상 선택하세요."); return; }
    const recipients = checked.map(c => ({ name: c.dataset.name, email: c.value }));
    const testMode = emailOverlay.querySelector("#diffEmailTestMode").checked;
    const statusEl = emailOverlay.querySelector(".diff-email-status");
    statusEl.textContent = "발송 중...";
    emailOverlay.querySelector(".diff-email-send").disabled = true;
    try {
      await postSheetWebApp("sendRawDiffEmail", { diff, recipients, testMode });
      statusEl.textContent = `${recipients.map(r => r.name).join(", ")} 님께 발송 완료`;
      setTimeout(() => emailOverlay.remove(), 2000);
    } catch (e) {
      statusEl.style.color = "#b71c1c";
      statusEl.textContent = `발송 실패: ${e.message}`;
      emailOverlay.querySelector(".diff-email-send").disabled = false;
    }
  };
}

function openWarningEmailDialog(warnings, reportRows, planKey) {
  document.querySelector(".warning-email-overlay")?.remove();

  const emailOverlay = document.createElement("div");
  emailOverlay.className = "raw-diff-overlay warning-email-overlay";

  const staffList = RECEIVABLE_CC_OPTIONS;
  const planLabel = planKey === "__total__" ? "전체" : formatPlanLabel(planKey);

  // 누락 항목 요약
  const missingList = warnings.map(w => `${w.거래처명}: ${w.missing.join(", ")}`);

  emailOverlay.innerHTML = `
    <div class="raw-diff-dialog" style="max-width:460px;">
      <h3 style="margin-top:0;">은행 업로드 전 확인 요청 메일</h3>
      <p style="font-size:13px;color:#555;margin-bottom:4px;">
        [${planLabel}] 결제 보고서 — 누락 항목 ${warnings.length}건 확인을 요청할 담당자를 선택하세요.
      </p>
      <div style="background:#fff3cd;border-radius:6px;padding:8px 10px;font-size:12px;color:#7c5800;margin-bottom:12px;max-height:80px;overflow-y:auto;">
        ${missingList.map(s => `• ${escapeHtml(s)}`).join("<br>")}
      </div>
      <div style="display:flex;flex-direction:column;gap:6px;margin-bottom:14px;">
        ${staffList.map(s => `
          <label style="display:flex;align-items:center;gap:8px;font-size:14px;cursor:pointer;">
            <input type="checkbox" class="warn-email-recipient" value="${escapeHtml(s.email)}"
              data-name="${escapeHtml(s.name)}"
              ${s.name === "김도연" ? "checked" : ""} />
            ${escapeHtml(s.name)} <span style="color:#888;font-size:12px;">${escapeHtml(s.email)}</span>
          </label>`).join("")}
      </div>
      <label style="display:flex;align-items:center;gap:6px;font-size:13px;margin-bottom:14px;">
        <input type="checkbox" id="warnEmailTestMode" />
        테스트 모드 (yhj@mauto.co.kr 로만 발송)
      </label>
      <div style="display:flex;gap:8px;justify-content:flex-end;">
        <button type="button" class="warn-email-cancel btn-secondary" style="padding:7px 16px;">취소</button>
        <button type="button" class="warn-email-send btn-primary" style="padding:7px 16px;">발송</button>
      </div>
      <p class="warn-email-status" style="margin-top:10px;font-size:13px;color:#1e40af;min-height:18px;"></p>
    </div>`;

  document.body.appendChild(emailOverlay);

  emailOverlay.querySelector(".warn-email-cancel").onclick = () => emailOverlay.remove();
  emailOverlay.querySelector(".warn-email-send").onclick = async () => {
    const checked = [...emailOverlay.querySelectorAll(".warn-email-recipient:checked")];
    if (!checked.length) { alert("수신자를 한 명 이상 선택하세요."); return; }
    const recipients = checked.map(c => ({ name: c.dataset.name, email: c.value }));
    const testMode = emailOverlay.querySelector("#warnEmailTestMode").checked;
    const statusEl = emailOverlay.querySelector(".warn-email-status");
    statusEl.textContent = "발송 중...";
    emailOverlay.querySelector(".warn-email-send").disabled = true;
    try {
      await postSheetWebApp("sendPaymentWarningEmail", {
        warnings, planLabel, recipients, testMode,
      });
      statusEl.textContent = `${recipients.map(r => r.name).join(", ")} 님께 발송 완료`;
      setTimeout(() => emailOverlay.remove(), 2000);
    } catch (e) {
      statusEl.style.color = "#b71c1c";
      statusEl.textContent = `발송 실패: ${e.message}`;
      emailOverlay.querySelector(".warn-email-send").disabled = false;
    }
  };
}

async function loadSheetPayables() {
  try {
    // 1단계: 핵심 데이터(raw + 업체마스터)만 먼저 로드 → 빠르게 화면 표시
    const [vendorRows, rows] = await Promise.all([
      fetchVendorMasterRowsFromApi(),
      SHEET_APP_SCRIPT_URL
        ? fetchSheetWebApp().then(b => Array.isArray(b) ? b : (b && (b.data || b.rows)) || [])
        : fetchPublicSheet(),
    ]);
    setVendorMasterRows(vendorRows);
    // 담당자 마스터가 이미 로드된 경우, 업체마스터 기반 이름→코드 재매핑
    if (receivableManagerState.rows.length) setManagerMasterRows(receivableManagerState.rows);
    if (!rows || !rows.length) {
      elements.payables.innerHTML = `
        <div class="panel">
          <div class="empty-state">시트 데이터를 읽어오지 못했습니다. Apps Script가 데이터를 반환하지 않았습니다.</div>
        </div>
      `;
      console.warn("시트에서 미지급 데이터를 읽어오지 못했습니다.");
      return;
    }

    const newParsedItems = rows.map(parsePayableRow);
    const diff = detectPayablesRawDiff(newParsedItems);

    const applyPayables = (remotePlanRows, remoteHistoryRows, completeStableKeys = new Set()) => {
      payables = applySavedPayablesState(newParsedItems);
      if (completeStableKeys.size) {
        payables = payables.map(item => {
          if (completeStableKeys.has(item.stableKey || buildPayableStableKey(item))) {
            return { ...item, completionStatus: "완료" };
          }
          return item;
        });
      }
      applySavedPaymentPlansFromApi(remotePlanRows);
      applyPaymentHistoryRows(remoteHistoryRows);
      ensureAutoPaymentPlans();
      enrichPayablesWithVendorMaster();
      enrichPayablesWithManagerDays();
      persistPayablesState();
      appendUpdateHistory("payables", diff);
      renderPartnerFilter();
      renderFilterControls();
      rerenderAll();
    };

    // diff 다이얼로그가 없으면 로컬 상태로 먼저 즉시 렌더링
    if (diff.length === 0) {
      applyPayables([], []);
    }

    // 2단계: 보조 데이터(결제계획, 결제이력) 백그라운드 로드 후 재적용
    const [remotePlanRows, remoteHistoryRows] = await Promise.all([
      fetchSavedPaymentPlansFromApi(),
      fetchPaymentHistoryRowsFromApi(),
    ]);

    if (diff.length > 0) {
      showPayablesRawDiffDialog(diff, (completeStableKeys) => applyPayables(remotePlanRows, remoteHistoryRows, completeStableKeys));
    } else {
      applyPayables(remotePlanRows, remoteHistoryRows);
    }
  } catch (error) {
    elements.payables.innerHTML = `
      <div class="panel">
        <div class="empty-state">시트 로드 실패: ${error.message}</div>
      </div>
    `;
    console.warn("Google Sheets 로드 실패:", error);
  }
}

const RECEIVABLES_SNAPSHOT_KEY = "receivable-payable-webapp.receivables-snapshot.v1";

function saveReceivablesSnapshot(rows) {
  try {
    const snap = rows.map(r => ({
      k: (r.code || "") + "||" + (r.year || "") + "||" + (r.month || ""),
      b: r.balance || 0,
      c: r.condition || "",
    }));
    localStorage.setItem(RECEIVABLES_SNAPSHOT_KEY, JSON.stringify(snap));
  } catch (_) { }
}

function detectReceivablesSheetDiff(newRows) {
  try {
    const prev = JSON.parse(localStorage.getItem(RECEIVABLES_SNAPSHOT_KEY) || "[]");
    if (!prev.length) return [];
    const prevMap = Object.fromEntries(prev.map(r => [r.k, r]));
    const diff = [];
    newRows.forEach(r => {
      const k = (r.code || "") + "||" + (r.year || "") + "||" + (r.month || "");
      const p = prevMap[k];
      if (p && Number(p.b) !== Number(r.balance || 0)) {
        diff.push({
          type: "changed", stableKey: k, label: `${r.year}-${String(r.month).padStart(2, "0")} ${r.name}`,
          prevAmount: Number(p.b), newAmount: Number(r.balance || 0),
        });
      }
      if (!p) diff.push({ type: "new", stableKey: k, label: `${r.year}-${String(r.month).padStart(2, "0")} ${r.name}` });
    });
    const newKeys = new Set(newRows.map(r => (r.code || "") + "||" + (r.year || "") + "||" + (r.month || "")));
    prev.forEach(p => {
      if (!newKeys.has(p.k)) diff.push({ type: "removed", stableKey: p.k, label: p.k, prevAmount: Number(p.b) });
    });
    return diff;
  } catch (_) { return []; }
}

async function loadSheetReceivables() {
  try {
    const [rawRows, mgrRows] = await Promise.all([
      fetchReceivablesFromApi(),
      fetchManagerMasterFromApi(),
    ]);
    console.log("[담당자] rows:", mgrRows?.length, "첫행:", JSON.stringify(mgrRows?.[0]));
    setManagerMasterRows(mgrRows);
    const newReceivables = (rawRows || []).map(parseReceivableRow).filter(Boolean);
    const diff = detectReceivablesSheetDiff(newReceivables);
    if (diff.length) appendUpdateHistory("receivables", diff);
    saveReceivablesSnapshot(newReceivables);
    receivables = newReceivables;
    enrichReceivablesWithManager();
    renderReceivables();
    renderSummary(); // 요약 패널 갱신 추가 (로딩 버그 해결)
  } catch (error) {
    console.warn("미수금 데이터 로드 실패:", error);
    receivables = [];
    renderReceivables();
    renderSummary(); // 실패 시에도 갱신
  }
}

function saveFixedLocal() {
  try { localStorage.setItem(FIXED_LOCAL_KEY, JSON.stringify(fixedExpenses)); }
  catch (e) { console.warn("[고정지출] 저장 실패:", e); }
}

function loadFixedLocal() {
  try {
    const raw = localStorage.getItem(FIXED_LOCAL_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch (e) { return null; }
}

function applyFixedPaste(text) {
  const parsed = parseMautoFixedPaste(text);
  if (!parsed.length) { alert("데이터를 읽지 못했습니다.\n헤더를 확인하세요."); return; }
  fixedExpenses = parsed;
  saveFixedLocal();
  renderFixedExpenses();
}

async function loadSheetFixedExpenses() {
  // localStorage 데이터 우선 사용
  const local = loadFixedLocal();
  if (local && local.length) {
    fixedExpenses = local;
    renderFixedExpenses();
    return;
  }
  try {
    let rows = [];
    if (SHEET_APP_SCRIPT_URL) {
      // Apps Script가 있으면 getFixed action 시도, 실패하면 공개 시트로 폴백
      try {
        const url = new URL(SHEET_APP_SCRIPT_URL);
        url.searchParams.set("action", "getFixed");
        const _fxToken = getApiToken();
        if (_fxToken) url.searchParams.set("token", _fxToken);
        const res = await fetch(url.toString());
        if (res.ok) {
          const body = await res.json();
          rows = Array.isArray(body.rows) ? body.rows : (Array.isArray(body) ? body : []);
        }
      } catch (_) { }
    }
    // 공개 시트로 폴백 (또는 기본값)
    if (!rows.length && SHEET_SPREADSHEET_ID) {
      rows = await fetchPublicSheetByName(FIXED_EXPENSES_SHEET_NAME);
    }
    fixedExpenses = (rows || []).map(parseFixedExpenseRow).filter(item => item.year && item.month && item.title);
    renderFixedExpenses();
  } catch (err) {
    console.error("고정지출 로드 실패:", err);
    renderFixedExpenses();
  }
}

function parseFixedExpenseRow(row) {
  const normalized = {};
  Object.keys(row).forEach(key => {
    normalized[normalizeKey(key)] = row[key];
  });

  // 1. 일(일자) 콼럼 직접 추출
  let year = Number(normalized["연도"] || normalized["year"] || 0);
  let month = Number(normalized["월"] || normalized["month"] || 0);
  let day = Number(normalized["일"] || normalized["day"] || 0);

  // 2. 날짜 콼럼에서 연/월/일 보완 (날짜 컬럼이 있다면 무조건 우선)
  const rawDate = normalized["날짜"] || normalized["date"] || "";
  if (rawDate) {
    // Date 오브젝트
    if (rawDate instanceof Date && !isNaN(rawDate)) {
      year = rawDate.getFullYear();
      month = rawDate.getMonth() + 1;
      day = rawDate.getDate();
    } else {
      const dateStr = String(rawDate).trim();
      // gviz API 반환 형식: Date(YYYY,M,D) — 월은 0-indexed
      const gvizMatch = dateStr.match(/^Date\((\d+),(\d+),(\d+)\)/);
      if (gvizMatch) {
        year = parseInt(gvizMatch[1]);
        month = parseInt(gvizMatch[2]) + 1; // 0-indexed 보정
        day = parseInt(gvizMatch[3]);
      } else {
        // YYYY-MM-DD / YYYY.MM.DD / 필요 시 뒤에 요일 등이 붙어도 인식
        const isoMatch = dateStr.match(/(\d{4})[-./](\d{1,2})[-./](\d{1,2})/);
        if (isoMatch) {
          year = parseInt(isoMatch[1]);
          month = parseInt(isoMatch[2]);
          day = parseInt(isoMatch[3]);
        }
      }
    }
  }

  return {
    year,
    month,
    day,
    title: String(normalized["내용"] || normalized["content"] || "").trim(),
    bank: String(normalized["은행"] || normalized["bank"] || "").trim(),
    amount: parseAmt(normalized["결제금액"] || normalized["실결제금액"] || normalized["금액"] || 0),
    raw: row
  };
}

function normalizeKey(key) {
  return String(key || "").trim().replace(/\s+/g, "").replace(/[\u200B-\u200D\uFEFF]/g, "").toLowerCase();
}

function buildPayableSourceKey(item) {
  const normalizedCode = normalizeVendorCode(item.codeNormalized || item.code || item.codeRaw || "");
  const parts = [
    normalizedCode,
    String(item.year || ""),
    String(item.month || "").padStart(2, "0"),
    String(Math.round(Number(item.purchase || 0))),
    normalizeDueGroupLabel(item.dueCategory || ""),
    String(item.payDate || "").trim(),
    String(item.memo || "").trim(),
  ];
  return parts.join("||");
}

function normalizeDueGroupLabel(label) {
  if (!label) return "미정";
  const s = String(label).trim();
  if (s.includes("보류")) return "보류";
  if (s.includes("제외")) return "제외";
  if (s.includes("내일")) return "오늘/내일";
  if (s.includes("오늘")) return "오늘/내일";
  return s;
}

// raw 교체 시에도 살아남는 안정적 식별자 (금액/메모 제외, 납기그룹 포함)
function buildPayableStableKey(item) {
  const code = normalizeVendorCode(item.codeNormalized || item.code || item.codeRaw || "");
  return [
    code,
    String(item.year || ""),
    String(item.month || "").padStart(2, "0"),
    normalizeDueGroupLabel(item.dueCategory || ""),
  ].join("||");
}

function getPayablesStateSnapshot() {
  return payables.reduce((acc, item) => {
    const sourceKey = item.sourceKey || buildPayableSourceKey(item);
    if (!sourceKey) return acc;
    acc[sourceKey] = {
      decisionAmount: Number(item.decisionAmount ?? 0),
      paymentPlan: item.paymentPlan || "",
      selected: Boolean(item.selected),
      paidOverride: Number(item.paidOverride ?? item.paid ?? 0),
      completionStatus: item.completionStatus || "",
      stableKey: item.stableKey || buildPayableStableKey(item),
      updatedAt: new Date().toISOString(),
    };
    return acc;
  }, {});
}

function savePayablesStateToLocal() {
  try {
    window.localStorage.setItem(PAYABLES_LOCAL_STATE_KEY, JSON.stringify(getPayablesStateSnapshot()));
  } catch (error) {
    console.warn("미지급 로컬 저장 실패:", error);
  }
}

function loadPayablesStateFromLocal() {
  try {
    const raw = window.localStorage.getItem(PAYABLES_LOCAL_STATE_KEY);
    return raw ? JSON.parse(raw) : {};
  } catch (error) {
    console.warn("미지급 로컬 상태 복원 실패:", error);
    return {};
  }
}

function saveGroupOrder() {
  try {
    localStorage.setItem(GROUP_ORDER_KEY, JSON.stringify({
      payGroupOrder: filterState.groupOrder,
      payGroups: filterState.groups,
      rcvOrder: rcvGroupState.order,
      rcvFilter: rcvGroupState.filter,
    }));
  } catch (e) { }
}

function loadGroupOrder() {
  try {
    const raw = localStorage.getItem(GROUP_ORDER_KEY);
    if (!raw) return;
    const saved = JSON.parse(raw);
    if (Array.isArray(saved.payGroupOrder)) filterState.groupOrder = saved.payGroupOrder;
    if (saved.payGroups !== undefined) filterState.groups = saved.payGroups;
    if (Array.isArray(saved.rcvOrder)) rcvGroupState.order = saved.rcvOrder;
    if (saved.rcvFilter !== undefined) rcvGroupState.filter = saved.rcvFilter;
  } catch (e) { }
}

function saveVendorMemos() {
  try { localStorage.setItem(VENDOR_MEMO_KEY, JSON.stringify(vendorMemos)); } catch (e) { }
}

function loadVendorMemos() {
  try {
    const raw = localStorage.getItem(VENDOR_MEMO_KEY);
    if (raw) vendorMemos = JSON.parse(raw);
  } catch (e) { }
}

function getVendorMemo(code) {
  return vendorMemos[normalizeVendorCode(code || "")] || { common: "", payables: "", receivables: "" };
}

function buildVendorTooltip(code, rawMemo, section) {
  const vm = getVendorMemo(code);
  const parts = [];
  if (vm.common) parts.push(`[공통] ${vm.common}`);
  if (section === "payables" && vm.payables) parts.push(`[미지급] ${vm.payables}`);
  if (section === "receivables" && vm.receivables) parts.push(`[미수금] ${vm.receivables}`);
  if (rawMemo) parts.push(`[메모] ${rawMemo}`);
  return parts.join("\n");
}

function openVendorMemoEditor(code, name) {
  document.querySelector(".vendor-memo-overlay")?.remove();
  const norm = normalizeVendorCode(code || "");
  const vm = vendorMemos[norm] || { common: "", payables: "", receivables: "" };
  const overlay = document.createElement("div");
  overlay.className = "vendor-memo-overlay";
  overlay.innerHTML = `
    <div class="vendor-memo-popover" role="dialog" aria-modal="true">
      <div class="vendor-memo-header">
        <strong>${escapeHtml(name)}</strong> 업체 메모
        <button type="button" class="vendor-memo-close">✕</button>
      </div>
      <label class="vendor-memo-label">공통 메모 (미수금·미지급 공통)
        <textarea class="vendor-memo-textarea" data-field="common" rows="2">${escapeHtml(vm.common || "")}</textarea>
      </label>
      <label class="vendor-memo-label">미지급 메모
        <textarea class="vendor-memo-textarea" data-field="payables" rows="2">${escapeHtml(vm.payables || "")}</textarea>
      </label>
      <label class="vendor-memo-label">미수금 메모
        <textarea class="vendor-memo-textarea" data-field="receivables" rows="2">${escapeHtml(vm.receivables || "")}</textarea>
      </label>
      <div class="vendor-memo-actions">
        <button type="button" class="vendor-memo-delete">삭제</button>
        <button type="button" class="vendor-memo-save">저장</button>
      </div>
    </div>
  `;
  document.body.appendChild(overlay);

  overlay.querySelector(".vendor-memo-close").addEventListener("click", () => overlay.remove());
  overlay.querySelector(".vendor-memo-delete").addEventListener("click", () => {
    delete vendorMemos[norm];
    saveVendorMemos();
    overlay.remove();
    rerenderAll();
  });
  overlay.querySelector(".vendor-memo-save").addEventListener("click", () => {
    const result = { common: "", payables: "", receivables: "" };
    overlay.querySelectorAll(".vendor-memo-textarea").forEach(ta => {
      result[ta.dataset.field] = ta.value.trim();
    });
    if (result.common || result.payables || result.receivables) {
      vendorMemos[norm] = result;
    } else {
      delete vendorMemos[norm];
    }
    saveVendorMemos();
    overlay.remove();
    rerenderAll();
  });
  overlay.addEventListener("mousedown", e => {
    if (e.target === overlay) overlay.remove();
  });
}

// 사람 이름 여부 판별 (업체명과 대표자명 비교용)
function isPersonName(name) {
  const n = String(name || "").trim().replace(/\s+/g, "");
  if (!n || n.length < 2 || n.length > 5) return false;
  if (/주식회사|\(주\)|\(유\)|\(합\)|상사|시스템|공업|물류|산업|전자|기업|그룹|법인|협회|조합|공단|공사|센터|연구|학원|병원|약국|의원|마트|건설|개발|기계|전기|설비|정비|금속|철강|화학|무역|서비스|솔루션|테크/.test(n)) return false;
  return /^[가-힣]+$/.test(n);
}

function appendUpdateHistory(section, diffItems) {
  if (!diffItems || !diffItems.length || !SHEET_APP_SCRIPT_URL) return;
  const rows = diffItems.map(d => ({
    recorded_at: new Date().toISOString(),
    section,
    action: d.type,
    stable_key: d.stableKey || "",
    label: d.label || "",
    prev_amount: d.prevAmount ?? d.paidOverride ?? "",
    new_amount: d.newAmount ?? "",
    memo: d.type === "removed" ? "raw에서 삭제됨" : "금액 변경",
  }));
  postSheetWebApp("appendUpdateHistory", { rows }).catch(e =>
    console.warn("업데이트이력 저장 실패:", e)
  );
}

// ── API 토큰 관리 ─────────────────────────────────────────────
function getApiToken() {
  return localStorage.getItem(API_TOKEN_STORAGE_KEY) || "miraeautomation2026";
}

function setApiToken(token) {
  if (token) localStorage.setItem(API_TOKEN_STORAGE_KEY, token.trim());
  else localStorage.removeItem(API_TOKEN_STORAGE_KEY);
}

function promptApiToken() {
  return new Promise(resolve => {
    const overlay = document.createElement("div");
    overlay.className = "raw-diff-overlay";
    overlay.innerHTML = `
      <div class="raw-diff-dialog" style="max-width:420px;">
        <h3 style="margin-top:0;">API 인증 토큰 입력</h3>
        <p style="font-size:13px;color:#555;margin-bottom:12px;">
          Apps Script에 설정된 <code>API_TOKEN</code> 값을 입력하세요.<br>
          한 번 입력하면 이 기기에 저장됩니다.
        </p>
        <input id="apiTokenInput" type="password" placeholder="토큰 입력..."
          style="width:100%;box-sizing:border-box;padding:8px 10px;font-size:14px;border:1px solid #ccc;border-radius:6px;margin-bottom:14px;" />
        <div style="display:flex;gap:8px;justify-content:flex-end;">
          <button id="apiTokenCancel" class="btn-secondary" style="padding:7px 16px;">취소</button>
          <button id="apiTokenConfirm" class="btn-primary" style="padding:7px 16px;">저장</button>
        </div>
      </div>`;
    document.body.appendChild(overlay);
    const input = overlay.querySelector("#apiTokenInput");
    const stored = getApiToken();
    if (stored) input.value = stored;
    input.focus();
    overlay.querySelector("#apiTokenCancel").onclick = () => {
      overlay.remove(); resolve(getApiToken());
    };
    overlay.querySelector("#apiTokenConfirm").onclick = () => {
      setApiToken(input.value);
      overlay.remove(); resolve(input.value.trim());
    };
    input.addEventListener("keydown", e => {
      if (e.key === "Enter") { overlay.querySelector("#apiTokenConfirm").click(); }
    });
  });
}

async function postSheetWebApp(action, payload = {}) {
  if (!SHEET_APP_SCRIPT_URL) {
    throw new Error("Apps Script URL이 비어 있습니다.");
  }
  const token = getApiToken();
  const response = await fetch(SHEET_APP_SCRIPT_URL, {
    method: "POST",
    headers: {
      "Content-Type": "text/plain;charset=utf-8",
    },
    body: JSON.stringify({
      action,
      token,
      ...payload,
    }),
  });
  if (!response.ok) {
    throw new Error(`Apps Script 저장 요청 실패: ${response.status}`);
  }
  const body = await response.json();
  if (body && body.error === "인증 실패") {
    const newToken = await promptApiToken();
    if (newToken) return postSheetWebApp(action, payload);
    throw new Error("인증 토큰이 없습니다.");
  }
  if (body && body.error) {
    throw new Error(body.error);
  }
  return body;
}

async function fetchSavedPaymentPlansFromApi() {
  if (!SHEET_APP_SCRIPT_URL) return [];
  try {
    const url = new URL(SHEET_APP_SCRIPT_URL);
    url.searchParams.set("action", "getPaymentPlans");
    const _token1 = getApiToken();
    if (_token1) url.searchParams.set("token", _token1);
    const response = await fetch(url.toString());
    if (!response.ok) {
      throw new Error(`결제계획 조회 실패: ${response.status}`);
    }
    const body = await response.json();
    if (Array.isArray(body)) return body;
    if (Array.isArray(body.rows)) return body.rows;
    if (Array.isArray(body.data)) return body.data;
    return [];
  } catch (error) {
    console.warn("결제계획 원격 조회 실패, 로컬 상태로 유지합니다.", error);
    return [];
  }
}

async function fetchVendorMasterRowsFromApi() {
  if (!SHEET_APP_SCRIPT_URL) return [];
  try {
    const url = new URL(SHEET_APP_SCRIPT_URL);
    url.searchParams.set("action", "getVendorMaster");
    const _token2 = getApiToken();
    if (_token2) url.searchParams.set("token", _token2);
    const response = await fetch(url.toString());
    if (!response.ok) {
      throw new Error(`업체마스터 조회 실패: ${response.status}`);
    }
    const body = await response.json();
    if (Array.isArray(body)) return body;
    if (Array.isArray(body.rows)) return body.rows;
    if (Array.isArray(body.data)) return body.data;
    return [];
  } catch (error) {
    console.warn("업체마스터 원격 조회 실패:", error);
    return [];
  }
}

async function fetchPaymentHistoryRowsFromApi() {
  if (!SHEET_APP_SCRIPT_URL) return [];
  try {
    const url = new URL(SHEET_APP_SCRIPT_URL);
    url.searchParams.set("action", "getPaymentHistory");
    const _token3 = getApiToken();
    if (_token3) url.searchParams.set("token", _token3);
    const response = await fetch(url.toString());
    if (!response.ok) {
      throw new Error(`결제이력 조회 실패: ${response.status}`);
    }
    const body = await response.json();
    if (Array.isArray(body)) return body;
    if (Array.isArray(body.rows)) return body.rows;
    if (Array.isArray(body.data)) return body.data;
    return [];
  } catch (error) {
    console.warn("결제이력 원격 조회 실패:", error);
    return [];
  }
}

async function fetchReceivablesFromApi() {
  if (SHEET_APP_SCRIPT_URL) {
    try {
      const url = new URL(SHEET_APP_SCRIPT_URL);
      url.searchParams.set("action", "getReceivables");
      const _token4 = getApiToken();
      if (_token4) url.searchParams.set("token", _token4);
      const response = await fetch(url.toString());
      if (!response.ok) throw new Error(`미수금 조회 실패: ${response.status}`);
      const body = await response.json();
      if (Array.isArray(body)) return body;
      if (Array.isArray(body.rows)) return body.rows;
      if (Array.isArray(body.data)) return body.data;
    } catch (error) {
      console.warn("미수금 Apps Script 조회 실패, gviz 폴백 시도:", error);
    }
  }
  try {
    return await fetchPublicSheetByName(SHEET_NAME_RECEIVABLES);
  } catch (error) {
    console.warn("미수금 gviz 조회 실패:", error);
    return [];
  }
}

async function fetchManagerMasterFromApi() {
  if (SHEET_APP_SCRIPT_URL) {
    try {
      const url = new URL(SHEET_APP_SCRIPT_URL);
      url.searchParams.set("action", "getManagerMaster");
      const _token5 = getApiToken();
      if (_token5) url.searchParams.set("token", _token5);
      const response = await fetch(url.toString());
      if (!response.ok) throw new Error(`담당자 마스터 조회 실패: ${response.status}`);
      const body = await response.json();
      if (Array.isArray(body)) return body;
      if (Array.isArray(body.rows)) return body.rows;
    } catch (error) {
      console.warn("담당자 마스터 Apps Script 조회 실패, gviz 폴백 시도:", error);
    }
  }
  try {
    return await fetchPublicSheetByName(MANAGER_MASTER_SHEET_NAME);
  } catch (error) {
    console.warn("담당자 마스터 gviz 조회 실패:", error);
    return [];
  }
}

function setManagerMasterRows(rows) {
  if (!Array.isArray(rows) || !rows.length) {
    console.warn("[담당자] 마스터 데이터 없음 또는 비어있음:", rows);
    return;
  }
  const firstRow = rows[0];
  const allKeys = Object.keys(firstRow);
  console.log("[담당자] 시트 컬럼:", allKeys, "첫행:", JSON.stringify(firstRow));

  receivableManagerState.rows = rows;
  receivableManagerState.map = new Map();
  const codeKey  = allKeys.find(k => /코드|code/i.test(k)) || "";
  const nameKey  = allKeys.find(k => /^거래처명$|^업체명$/i.test(k)) || "";
  const mgrKey   = allKeys.find(k => /담당자|manager/i.test(k)) || "";
  const emailKey = allKeys.find(k => /이메일|email/i.test(k)) || "";
  const daysKey  = allKeys.find(k => /^일$|수금조건|납기$/i.test(k)) || "";
  console.log("[담당자] 사용 컬럼 — 코드:", codeKey, "거래처명:", nameKey, "담당자:", mgrKey, "이메일:", emailKey, "일:", daysKey);

  // 업체마스터에서 거래처명 → 코드 역조회 맵 (코드 컬럼 없을 때 fallback)
  const nameToCode = new Map();
  vendorMasterState.rows.forEach(r => {
    const n = String(r["거래처명"] || "").trim().toLowerCase();
    if (n && r["거래처코드_norm"]) nameToCode.set(n, r["거래처코드_norm"]);
  });

  const samples = [];
  rows.forEach((row) => {
    const rawVal  = codeKey ? (row[codeKey] ?? "") : "";
    let code      = normalizeVendorCode(String(rawVal).trim());
    const nameVal = nameKey ? String(row[nameKey] ?? "").trim() : "";
    if (!code && nameVal) {
      // 업체마스터에서 이름으로 코드 조회, 없으면 이름 자체를 키로 사용
      code = nameToCode.get(nameVal.toLowerCase()) || ("__name__" + nameVal.toLowerCase());
    }
    const manager = String(mgrKey   ? (row[mgrKey]   ?? "") : "").trim();
    const email   = String(emailKey ? (row[emailKey] ?? "") : "").trim()
      || RECEIVABLE_MANAGER_EMAIL_MAP[manager] || "";
    const days    = String(daysKey  ? (row[daysKey]  ?? "") : "").trim();
    if (code && manager) {
      receivableManagerState.map.set(code, { manager, email, days });
      if (samples.length < 5) samples.push({ name: nameVal, code });
    }
  });
  console.log(`[담당자] 마스터 로드: ${receivableManagerState.map.size}건`, "샘플:", JSON.stringify(samples));
}

function enrichReceivablesWithManager() {
  if (receivables.length) {
    const sample = receivables[0];
    const mapKeys = [...receivableManagerState.map.keys()].slice(0, 10);
    console.log("[담당자 매칭] 시작", {
      mapKeys: JSON.stringify(mapKeys),
      sampleCode: sample.code,
      sampleMatch: receivableManagerState.map.get(sample.code)
    });
  }
  const todayRcv = new Date(); todayRcv.setHours(0, 0, 0, 0);
  receivables.forEach(item => {
    const mgr = receivableManagerState.map.get(item.code)
      || receivableManagerState.map.get("__name__" + String(item.name || "").trim().toLowerCase());
    if (mgr) {
      item.manager     = mgr.manager || "미지정";
      item.managerEmail = mgr.email  || "";
      item.managerDays  = mgr.days   || "";
    } else {
      item.manager      = "미지정";
      item.managerEmail = "";
      item.managerDays  = "";
    }
    // 담당자 마스터 '일' 값이 있으면 수금조건·납기일 재계산 (우선 적용)
    if (item.managerDays) {
      item.condition = item.managerDays;
      const d = calcReceivableDueDate(item.year, item.month, item.memo, item.managerDays);
      item.dueDate = d ? d.toISOString().slice(0, 10) : item.dueDate;
      item.elapsed = item.dueDate
        ? Math.floor((todayRcv - new Date(item.dueDate + "T00:00:00")) / 86400000)
        : null;
    }
  });
  console.log("[담당자 매칭] 완료", receivables.filter(r => r.manager !== "미지정").length, "건 매칭됨");
  enrichPayablesWithManagerDays();
}

function enrichPayablesWithManagerDays() {
  if (!payables.length) return;
  payables = payables.map(item => {
    const mgr = receivableManagerState.map.get(item.code || item.codeNormalized)
      || receivableManagerState.map.get("__name__" + String(item.name || "").trim().toLowerCase());
    const mgrDays = mgr?.days || "";
    if (!mgrDays) return { ...item, managerDays: "" };
    // 미지급은 ERP 납기값 우선, 없거나 기타일 때만 담당자 '일' 적용
    const cur = item.dueCategory || "";
    return {
      ...item,
      managerDays: mgrDays,
      dueCategory: (!cur || cur === "기타") ? mgrDays : cur,
    };
  });
}

function applyPaymentHistoryRows(rows) {
  paymentHistoryState.rows = Array.isArray(rows) ? rows : [];
  if (!Array.isArray(rows) || !rows.length) return;
  const historyBySourceKey = rows.reduce((acc, row) => {
    const sourceKey = String(row.source_key || row.sourceKey || "").trim();
    if (!sourceKey) return acc;
    if (!acc[sourceKey]) {
      acc[sourceKey] = {
        amount: 0,
        lastDate: "",
        count: 0,
      };
    }
    const item = acc[sourceKey];
    item.amount += Number(row.지급금액 || row.amount || 0);
    item.count += 1;
    const dateValue = normalizeDateValue(row.지급일자 || row.paymentDate || "");
    if (dateValue && (!item.lastDate || dateValue > item.lastDate)) {
      item.lastDate = dateValue;
    }
    return acc;
  }, {});

  payables = payables.map(item => {
    const history = historyBySourceKey[item.sourceKey || ""];
    if (!history) return item;
    // raw 시트 업데이트로 완료가 리셋된 항목은 결제이력 재적용 안함 (raw에 이미 반영)
    if (item._rawResetCompletion) return item;
    // 이미 완료 처리된 항목은 건드리지 않음
    if (item.completionStatus === "완료") return item;
    // raw 지급합을 초과하는 결제이력만 추가 (raw에 포함된 금액 이중계산 방지)
    const rawPaid = Number(item.paid || 0);
    const historyAboveRaw = Math.max(0, history.amount - rawPaid);
    if (historyAboveRaw === 0) return item;
    const nextPaidOverride = Math.min(Number(item.purchase || 0), rawPaid + historyAboveRaw);
    const nextOutstanding = Math.max(0, Number(item.purchase || 0) - nextPaidOverride);
    return {
      ...item,
      paidOverride: nextPaidOverride,
      completionStatus: nextOutstanding === 0 ? "완료" : (history.count > 0 ? "부분결제" : item.completionStatus || ""),
      decisionAmount: nextOutstanding === 0 ? 0 : item.decisionAmount,
      paymentPlan: nextOutstanding === 0 ? "" : item.paymentPlan,
      selected: nextOutstanding === 0 ? false : item.selected,
    };
  });
}

function parseVendorMasterSheetRows(rows) {
  return rows
    .map(row => {
      const vendorCodeRaw = row["거래처코드"] ?? row["거래처코드_raw"] ?? row["거래처코드_norm"] ?? row.vendor_id ?? "";
      const businessNumber = row["사업자(주민)번호"] ?? row["사업자번호"] ?? "";
      return {
        vendor_id: row.vendor_id || normalizeVendorCode(vendorCodeRaw),
        거래처코드_raw: String(vendorCodeRaw || ""),
        거래처코드_norm: normalizeVendorCode(vendorCodeRaw || row["거래처코드_norm"] || ""),
        거래처명: String(row["거래처명"] || ""),
        거래처분류: String(row["거래처분류"] || ""),
        거래처구분: String(row["거래처구분"] || row["거래처구분코드"] || ""),
        대표자명: String(row["대표자명"] || ""),
        사업자번호: normalizeBusinessNumber(businessNumber),
        전화번호: String(row["전화번호"] || ""),
        팩스번호: String(row["팩스번호"] || ""),
        주소: [row["주소"], row["나머지_주소"]].filter(Boolean).join(" ").trim(),
        업태: String(row["업태"] || ""),
        종목: String(row["종목"] || ""),
        홈페이지: String(row["홈페이지"] || ""),
        은행: String(row["은행"] || ""),
        계좌번호: String(row["계좌번호"] || row["계좌"] || ""),
        예금주: String(row["예금주"] || ""),
      };
    })
    .filter(row => row.거래처코드_norm || row.사업자번호 || row.거래처명);
}

function getVendorMatchKey(row) {
  // 사업자번호가 "0"이면 유효한 식별자로 쓰지 않음 (이름 등으로 매칭되게 유도)
  const bizNum = (row.사업자번호 && row.사업자번호 !== "0") ? row.사업자번호 : "";
  return row.거래처코드_norm || bizNum || row.거래처명;
}

function setVendorMasterRows(rows) {
  const parsedRows = parseVendorMasterSheetRows(rows || []);
  vendorMasterState.rows = parsedRows;
  vendorMasterState.map = new Map(parsedRows.map(row => [getVendorMatchKey(row), row]));
}

function getVendorMasterRowForPayable(item) {
  const codeKey = normalizeVendorCode(item.codeNormalized || item.code || item.codeRaw || "");
  const nameKey = String(item.name || "").trim();
  return vendorMasterState.map.get(codeKey)
    || vendorMasterState.map.get(nameKey)
    || null;
}

function enrichPayablesWithVendorMaster() {
  if (!vendorMasterState.rows.length) return;
  payables = payables.map(item => {
    const vendor = getVendorMasterRowForPayable(item);
    if (!vendor) {
      return {
        ...item,
        vendorBank: "",
        vendorAccount: "",
        vendorAccountHolder: "",
        vendorMatched: false,
      };
    }
    return {
      ...item,
      vendorBank: vendor.은행 || "",
      vendorAccount: vendor.계좌번호 || "",
      vendorAccountHolder: vendor.예금주 || vendor.거래처명 || "",
      vendorRepresentative: vendor.대표자명 || "",
      vendorMatched: true,
    };
  });
}

function getPayablesForPlanKey(planKey, items) {
  if (planKey === "__total__") return items;
  return items.filter(item => (item.paymentPlan || "") === planKey);
}

function getPayablesForPlanKeys(planKeys, items) {
  if (!planKeys || !planKeys.length) return items;
  const keySet = new Set(planKeys);
  if (keySet.has("__total__")) return items;
  return items.filter(item => keySet.has(item.paymentPlan || ""));
}

function buildPlannedPaymentReportRows(planKey = "__total__") {
  const filteredPayables = getFilteredItems(payables, "payables");
  const targetItems = getPayablesForPlanKey(planKey, filteredPayables)
    .filter(item => Number(item.decisionAmount || 0) > 0);
  const grouped = new Map();

  targetItems.forEach(item => {
    const key = `${item.code || ""}||${item.name || ""}||${item.paymentPlan || ""}`;
    if (!grouped.has(key)) {
      grouped.set(key, {
        거래처코드: item.code || "",
        거래처명: item.name || "",
        결제예정일: item.paymentPlan || "",
        은행: item.vendorBank || "",
        계좌번호: item.vendorAccount || "",
        예금주: item.vendorAccountHolder || item.name || "",
        대표자명: item.vendorRepresentative || "",
        지급금액: 0,
        연월목록: [],
        연월키목록: [],
        메모목록: [],
      });
    }
    const row = grouped.get(key);
    row.지급금액 += Number(item.decisionAmount || 0);
    const mk = getMonthKey(item);
    row.연월목록.push(mk ? mk.replace(/^20(\d{2})-/, "$1-") : "");
    row.연월키목록.push(mk);
    if (item.memo) {
      row.메모목록.push(item.memo);
    }
  });

  return [...grouped.values()].map(row => ({
    ...row,
    지급금액: Math.round(row.지급금액),
    연월목록: [...new Set(row.연월목록)].join(", "),
    연월키목록: [...new Set(row.연월키목록)],
    메모목록: [...new Set(row.메모목록)].join(", "),
  }));
}

function buildPaymentHistoryRows(planKey = "__total__") {
  const filteredPayables = getFilteredItems(payables, "payables");
  const targetItems = getPayablesForPlanKey(planKey, filteredPayables)
    .filter(item => Number(item.decisionAmount || 0) > 0);
  const stamp = Date.now();
  return targetItems.map((item, index) => ({
    history_id: `${item.sourceKey}||${stamp}||${index}`,
    source_key: item.sourceKey || buildPayableSourceKey(item),
    거래처코드_norm: normalizeVendorCode(item.codeNormalized || item.code || item.codeRaw || ""),
    거래처명: item.name || "",
    지급일자: item.paymentPlan || "",
    지급금액: Number(item.decisionAmount || 0),
    은행: item.vendorBank || "",
    계좌번호: item.vendorAccount || "",
    예금주: item.vendorAccountHolder || item.name || "",
    적요: item.memo || "",
    결과상태: "완료",
    created_at: new Date().toISOString(),
  }));
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function buildPaymentApprovalHtml(planKey = "__total__") {
  const reportRows = buildPlannedPaymentReportRows(planKey);
  const totalAmount = reportRows.reduce((sum, row) => sum + Number(row.지급금액 || 0), 0);
  const reportTitle = planKey === "__total__" ? "전체 예정 결제 보고서" : `${formatPlanLabel(planKey)} 결제 보고서`;
  const generatedAt = new Date().toLocaleString("ko-KR");

  return `
<div style="font-family:'Malgun Gothic','Apple SD Gothic Neo',sans-serif;color:#1f2937;line-height:1.5;">
  <h2 style="margin:0 0 12px;font-size:22px;color:#0f172a;">${escapeHtml(reportTitle)}</h2>
  <table style="width:100%;border-collapse:collapse;margin-bottom:14px;">
    <tr>
      <td style="padding:8px 10px;border:1px solid #dbe3f0;background:#f8fafc;width:160px;font-weight:700;">생성일시</td>
      <td style="padding:8px 10px;border:1px solid #dbe3f0;">${escapeHtml(generatedAt)}</td>
      <td style="padding:8px 10px;border:1px solid #dbe3f0;background:#f8fafc;width:160px;font-weight:700;">총 지급예정액</td>
      <td style="padding:8px 10px;border:1px solid #dbe3f0;font-weight:800;color:#0f172a;">${formatNumber(totalAmount)}원</td>
    </tr>
    <tr>
      <td style="padding:8px 10px;border:1px solid #dbe3f0;background:#f8fafc;font-weight:700;">대상 업체 수</td>
      <td style="padding:8px 10px;border:1px solid #dbe3f0;">${reportRows.length}개</td>
      <td style="padding:8px 10px;border:1px solid #dbe3f0;background:#f8fafc;font-weight:700;">결제 기준</td>
      <td style="padding:8px 10px;border:1px solid #dbe3f0;">${escapeHtml(planKey === "__total__" ? "전체 예정" : formatPlanLabel(planKey))}</td>
    </tr>
  </table>
  <table style="width:100%;border-collapse:collapse;">
    <thead>
      <tr>
        <th style="padding:10px 8px;border:1px solid #cfd8e3;background:#eef4fb;font-size:13px;">업체명</th>
        <th style="padding:10px 8px;border:1px solid #cfd8e3;background:#eef4fb;font-size:13px;">예정일</th>
        <th style="padding:10px 8px;border:1px solid #cfd8e3;background:#eef4fb;font-size:13px;">은행</th>
        <th style="padding:10px 8px;border:1px solid #cfd8e3;background:#eef4fb;font-size:13px;">계좌번호</th>
        <th style="padding:10px 8px;border:1px solid #cfd8e3;background:#eef4fb;font-size:13px;">예금주</th>
        <th style="padding:10px 8px;border:1px solid #cfd8e3;background:#eef4fb;font-size:13px;">대상 연월</th>
        <th style="padding:10px 8px;border:1px solid #cfd8e3;background:#eef4fb;font-size:13px;">금액</th>
      </tr>
    </thead>
    <tbody>
      ${reportRows.map(row => `
      <tr>
        <td style="padding:8px;border:1px solid #dbe3f0;">${escapeHtml(row.거래처명 || "-")}</td>
        <td style="padding:8px;border:1px solid #dbe3f0;">${escapeHtml(formatPlanLabel(row.결제예정일 || ""))}</td>
        <td style="padding:8px;border:1px solid #dbe3f0;">${escapeHtml(row.은행 || "확인 필요")}</td>
        <td style="padding:8px;border:1px solid #dbe3f0;">${escapeHtml(row.계좌번호 || "확인 필요")}</td>
        <td style="padding:8px;border:1px solid #dbe3f0;">${escapeHtml(row.예금주 || "확인 필요")}</td>
        <td style="padding:8px;border:1px solid #dbe3f0;">${escapeHtml(row.연월목록 || "-")}</td>
        <td style="padding:8px;border:1px solid #dbe3f0;text-align:right;font-weight:700;">${formatNumber(row.지급금액 || 0)}원</td>
      </tr>
      `).join("")}
    </tbody>
  </table>
</div>`.trim();
}

async function copyPaymentApprovalHtml(planKey = "__total__") {
  const html = buildPaymentApprovalHtml(planKey);
  await navigator.clipboard.writeText(html);
  return html;
}

function buildCompletedPaymentReportRows() {
  return [...paymentHistoryState.rows]
    .map(row => ({
      거래처명: row["거래처명"] || "",
      지급일자: normalizeDateValue(row["지급일자"] || ""),
      지급금액: Number(row["지급금액"] || 0),
      은행: row["은행"] || "",
      계좌번호: row["계좌번호"] || "",
      예금주: row["예금주"] || "",
      적요: row["적요"] || "",
      결과상태: row["결과상태"] || "",
      created_at: row["created_at"] || "",
    }))
    .sort((a, b) => String(b.created_at || b.지급일자 || "").localeCompare(String(a.created_at || a.지급일자 || "")));
}

function buildCompletedApprovalHtml() {
  const reportRows = buildCompletedPaymentReportRows();
  const totalAmount = reportRows.reduce((sum, row) => sum + Number(row.지급금액 || 0), 0);
  const generatedAt = new Date().toLocaleString("ko-KR");
  return `
<div style="font-family:'Malgun Gothic','Apple SD Gothic Neo',sans-serif;color:#1f2937;line-height:1.5;">
  <h2 style="margin:0 0 12px;font-size:22px;color:#0f172a;">최종 결재 보고서</h2>
  <table style="width:100%;border-collapse:collapse;margin-bottom:14px;">
    <tr>
      <td style="padding:8px 10px;border:1px solid #dbe3f0;background:#f8fafc;width:160px;font-weight:700;">생성일시</td>
      <td style="padding:8px 10px;border:1px solid #dbe3f0;">${escapeHtml(generatedAt)}</td>
      <td style="padding:8px 10px;border:1px solid #dbe3f0;background:#f8fafc;width:160px;font-weight:700;">총 완료금액</td>
      <td style="padding:8px 10px;border:1px solid #dbe3f0;font-weight:800;">${formatNumber(totalAmount)}원</td>
    </tr>
    <tr>
      <td style="padding:8px 10px;border:1px solid #dbe3f0;background:#f8fafc;font-weight:700;">완료 건수</td>
      <td style="padding:8px 10px;border:1px solid #dbe3f0;">${reportRows.length}건</td>
      <td style="padding:8px 10px;border:1px solid #dbe3f0;background:#f8fafc;font-weight:700;">보고 구분</td>
      <td style="padding:8px 10px;border:1px solid #dbe3f0;">결제 완료 내역</td>
    </tr>
  </table>
  <table style="width:100%;border-collapse:collapse;">
    <thead>
      <tr>
        <th style="padding:10px 8px;border:1px solid #cfd8e3;background:#eef4fb;font-size:13px;">업체명</th>
        <th style="padding:10px 8px;border:1px solid #cfd8e3;background:#eef4fb;font-size:13px;">지급일</th>
        <th style="padding:10px 8px;border:1px solid #cfd8e3;background:#eef4fb;font-size:13px;">은행</th>
        <th style="padding:10px 8px;border:1px solid #cfd8e3;background:#eef4fb;font-size:13px;">계좌번호</th>
        <th style="padding:10px 8px;border:1px solid #cfd8e3;background:#eef4fb;font-size:13px;">예금주</th>
        <th style="padding:10px 8px;border:1px solid #cfd8e3;background:#eef4fb;font-size:13px;">적요</th>
        <th style="padding:10px 8px;border:1px solid #cfd8e3;background:#eef4fb;font-size:13px;">금액</th>
      </tr>
    </thead>
    <tbody>
      ${reportRows.map(row => `
      <tr>
        <td style="padding:8px;border:1px solid #dbe3f0;">${escapeHtml(row.거래처명 || "-")}</td>
        <td style="padding:8px;border:1px solid #dbe3f0;">${escapeHtml(formatPlanLabel(row.지급일자 || ""))}</td>
        <td style="padding:8px;border:1px solid #dbe3f0;">${escapeHtml(row.은행 || "-")}</td>
        <td style="padding:8px;border:1px solid #dbe3f0;">${escapeHtml(row.계좌번호 || "-")}</td>
        <td style="padding:8px;border:1px solid #dbe3f0;">${escapeHtml(row.예금주 || "-")}</td>
        <td style="padding:8px;border:1px solid #dbe3f0;">${escapeHtml(row.적요 || "-")}</td>
        <td style="padding:8px;border:1px solid #dbe3f0;text-align:right;font-weight:700;">${formatNumber(row.지급금액 || 0)}원</td>
      </tr>
      `).join("")}
    </tbody>
  </table>
</div>`.trim();
}

function buildCompletedApprovalHtmlForRows(reportRows, titleOverride) {
  const totalAmount = reportRows.reduce((sum, row) => sum + Number(row.지급금액 || 0), 0);
  const generatedAt = new Date().toLocaleString("ko-KR");
  const title = titleOverride || "최종 결재 보고서";
  return `
<div style="font-family:'Malgun Gothic','Apple SD Gothic Neo',sans-serif;color:#1f2937;line-height:1.5;">
  <h2 style="margin:0 0 12px;font-size:22px;color:#0f172a;">${escapeHtml(title)}</h2>
  <table style="width:100%;border-collapse:collapse;margin-bottom:14px;">
    <tr>
      <td style="padding:8px 10px;border:1px solid #dbe3f0;background:#f8fafc;width:160px;font-weight:700;">생성일시</td>
      <td style="padding:8px 10px;border:1px solid #dbe3f0;">${escapeHtml(generatedAt)}</td>
      <td style="padding:8px 10px;border:1px solid #dbe3f0;background:#f8fafc;width:160px;font-weight:700;">총 완료금액</td>
      <td style="padding:8px 10px;border:1px solid #dbe3f0;font-weight:800;">${formatNumber(totalAmount)}원</td>
    </tr>
    <tr>
      <td style="padding:8px 10px;border:1px solid #dbe3f0;background:#f8fafc;font-weight:700;">완료 건수</td>
      <td style="padding:8px 10px;border:1px solid #dbe3f0;">${reportRows.length}건</td>
      <td style="padding:8px 10px;border:1px solid #dbe3f0;background:#f8fafc;font-weight:700;">보고 구분</td>
      <td style="padding:8px 10px;border:1px solid #dbe3f0;">결제 완료 내역</td>
    </tr>
  </table>
  <table style="width:100%;border-collapse:collapse;">
    <thead>
      <tr>
        <th style="padding:10px 8px;border:1px solid #cfd8e3;background:#eef4fb;font-size:13px;">업체명</th>
        <th style="padding:10px 8px;border:1px solid #cfd8e3;background:#eef4fb;font-size:13px;">지급일</th>
        <th style="padding:10px 8px;border:1px solid #cfd8e3;background:#eef4fb;font-size:13px;">은행</th>
        <th style="padding:10px 8px;border:1px solid #cfd8e3;background:#eef4fb;font-size:13px;">계좌번호</th>
        <th style="padding:10px 8px;border:1px solid #cfd8e3;background:#eef4fb;font-size:13px;">예금주</th>
        <th style="padding:10px 8px;border:1px solid #cfd8e3;background:#eef4fb;font-size:13px;">적요</th>
        <th style="padding:10px 8px;border:1px solid #cfd8e3;background:#eef4fb;font-size:13px;">금액</th>
      </tr>
    </thead>
    <tbody>
      ${reportRows.map(row => `
      <tr>
        <td style="padding:8px;border:1px solid #dbe3f0;">${escapeHtml(row.거래처명 || "-")}</td>
        <td style="padding:8px;border:1px solid #dbe3f0;">${escapeHtml(formatPlanLabel(row.지급일자 || ""))}</td>
        <td style="padding:8px;border:1px solid #dbe3f0;">${escapeHtml(row.은행 || "-")}</td>
        <td style="padding:8px;border:1px solid #dbe3f0;">${escapeHtml(row.계좌번호 || "-")}</td>
        <td style="padding:8px;border:1px solid #dbe3f0;">${escapeHtml(row.예금주 || "-")}</td>
        <td style="padding:8px;border:1px solid #dbe3f0;">${escapeHtml(row.적요 || "-")}</td>
        <td style="padding:8px;border:1px solid #dbe3f0;text-align:right;font-weight:700;">${formatNumber(row.지급금액 || 0)}원</td>
      </tr>
      `).join("")}
    </tbody>
  </table>
</div>`.trim();
}

async function copyCompletedApprovalHtml() {
  const html = buildCompletedApprovalHtml();
  await navigator.clipboard.writeText(html);
  return html;
}

function getBankCode(value) {
  const raw = String(value || "").trim();
  if (!raw) return "";
  if (/^\d+$/.test(raw)) {
    return raw.padStart(3, "0");
  }
  const normalized = raw.replace(/\s+/g, "");
  const bankCodeMap = {
    "한국은행": "001",
    "산업은행": "002",
    "기업은행": "003",
    "국민은행": "004",
    "외환은행": "005",
    "수협은행": "007",
    "수출입은행": "008",
    "농협은행": "011",
    "농협": "011",
    "단위농협": "012",
    "지역농축협": "012",
    "우리은행": "020",
    "SC제일은행": "023",
    "씨티은행": "027",
    "대구은행": "031",
    "iM뱅크": "031",
    "부산은행": "032",
    "광주은행": "034",
    "제주은행": "035",
    "전북은행": "037",
    "경남은행": "039",
    "새마을금고": "045",
    "신협": "048",
    "저축은행": "050",
    "산림조합": "064",
    "우체국": "071",
    "하나은행": "081",
    "신한은행": "088",
    "케이뱅크": "089",
    "카카오뱅크": "090",
    "토스뱅크": "092",
  };
  return bankCodeMap[normalized] || "";
}

function getPaymentReportWarnings(rows) {
  const warnings = [];
  rows.forEach(row => {
    const missing = [];
    if (!getBankCode(row.은행)) missing.push("은행코드");
    if (!String(row.계좌번호 || "").trim()) missing.push("계좌번호");
    if (missing.length) {
      warnings.push({
        거래처명: row.거래처명 || "-",
        missing,
      });
    }
  });
  return warnings;
}

async function downloadWooriTransferTemplate(planKey = "__total__") {
  if (typeof XLSX === "undefined") {
    throw new Error("엑셀 라이브러리를 불러오지 못했습니다.");
  }
  const reportRows = buildPlannedPaymentReportRows(planKey);
  if (!reportRows.length) {
    throw new Error("이체할 항목이 없습니다.");
  }
  const warnings = getPaymentReportWarnings(reportRows);
  if (warnings.length) {
    const names = warnings.slice(0, 5).map(w => w.거래처명).join(", ");
    const more = warnings.length > 5 ? ` 외 ${warnings.length - 5}건` : "";
    const ok = window.confirm(`은행코드/계좌번호/예금주 누락 업체 ${warnings.length}건:\n${names}${more}\n\n누락된 칸은 빈칸으로 저장됩니다. 그래도 진행하시겠습니까?`);
    if (!ok) throw new Error("취소됨");
  }

  const workbook = XLSX.utils.book_new();
  const sheetData = reportRows.map((row) => {
    const vendorName = String(row.거래처명 || "");
    const memo = row.메모목록 || row.연월목록 || (planKey === "__total__" ? "전체 예정" : formatPlanLabel(planKey));
    return [
      getBankCode(row.은행),          // A: 은행코드
      String(row.계좌번호 || ""),    // B: 계좌번호
      Number(row.지급금액 || 0),     // C: 금액
      vendorName,                    // D: 예금주 → 업체명 사용
      "", "",                        // E, F: 빈칸
      DEFAULT_SENDER_ACCOUNT_DISPLAY, // G: 출금계좌
      vendorName,                    // H: 받는분통장표시
      String(memo),                  // I: 내통장표시
    ];
  });

  const worksheet = XLSX.utils.aoa_to_sheet(sheetData);
  worksheet["!cols"] = [
    { wch: 6 }, { wch: 18 }, { wch: 14 }, { wch: 16 },
    { wch: 4 }, { wch: 4 }, { wch: 20 }, { wch: 16 }, { wch: 20 },
  ];
  XLSX.utils.book_append_sheet(workbook, worksheet, "이체");
  const downloadName = `우리은행_이체업로드_${planKey === "__total__" ? "전체예정" : formatPlanLabel(planKey).replace("/", "-")}.xlsx`;
  XLSX.writeFile(workbook, downloadName);
}

async function markPlanAsCompleted(planKey = "__total__") {
  const filteredPayables = getFilteredItems(payables, "payables");
  const targetItems = getPayablesForPlanKey(planKey, filteredPayables)
    .filter(item => Number(item.decisionAmount || 0) > 0);
  if (!targetItems.length) return { count: 0 };

  const historyRows = buildPaymentHistoryRows(planKey);
  try {
    await postSheetWebApp("appendPaymentHistory", {
      sheetName: HISTORY_SHEET_NAME,
      rows: historyRows,
    });
  } catch (error) {
    console.warn("결제이력 저장 실패:", error);
    throw new Error(`결제이력 저장 실패: ${error.message}`);
  }

  // 로컬 결제이력에도 즉시 반영 (새로고침 없이 완료 보고서에서 볼 수 있도록)
  paymentHistoryState.rows = [...paymentHistoryState.rows, ...historyRows];

  targetItems.forEach(item => {
    item.paidOverride = getPayableEffectivePaid(item) + Number(item.decisionAmount || 0);
    item.decisionAmount = 0;
    item.selected = false;
    item.completionStatus = "완료";
    item.paymentPlan = "";
  });
  persistPayablesState();
  return { count: targetItems.length };
}

function closePaymentReportModal() {
  const existing = document.querySelector(".payment-report-overlay");
  if (existing) {
    if (typeof existing.cleanup === "function") {
      existing.cleanup();
    }
    existing.remove();
  }
}

function closeCompletedReportModal() {
  const existing = document.querySelector(".completed-report-overlay");
  if (existing) {
    if (typeof existing.cleanup === "function") {
      existing.cleanup();
    }
    existing.remove();
  }
}

function getCompletedBatches() {
  const allRows = buildCompletedPaymentReportRows();
  if (!allRows.length) return [];
  // created_at 초 단위로 잘라 배치 구분 (같은 markPlanAsCompleted 호출은 거의 동시 기록)
  const batchMap = new Map();
  allRows.forEach(row => {
    const batchKey = String(row.created_at || "").slice(0, 19) || "unknown";
    if (!batchMap.has(batchKey)) batchMap.set(batchKey, []);
    batchMap.get(batchKey).push(row);
  });
  const batches = [...batchMap.entries()]
    .sort(([a], [b]) => b.localeCompare(a))
    .map(([key, rows]) => {
      const total = rows.reduce((s, r) => s + Number(r.지급금액 || 0), 0);
      const date = key.slice(0, 10);
      const time = key.length >= 16 ? key.slice(11, 16) : "";
      return { key, rows, date, time, total };
    });
  // 날짜별 회차 수 집계
  const dateCounts = {};
  batches.forEach(b => { dateCounts[b.date] = (dateCounts[b.date] || 0) + 1; });
  const dateSeq = {};
  batches.forEach(b => {
    if (dateCounts[b.date] > 1) {
      dateSeq[b.date] = (dateSeq[b.date] || 0) + 1;
      b.label = `${b.date} ${b.time} (${dateSeq[b.date]}회차 · ${b.rows.length}건 · ${formatNumber(b.total)}원)`;
    } else {
      b.label = `${b.date}${b.time ? " " + b.time : ""} (${b.rows.length}건 · ${formatNumber(b.total)}원)`;
    }
  });
  // 같은 날 여러 회차가 있으면 날짜 전체 합산 항목을 맨 앞에 삽입
  const enriched = [];
  const seenDates = new Set();
  batches.forEach(b => {
    if (!seenDates.has(b.date)) {
      seenDates.add(b.date);
      if (dateCounts[b.date] > 1) {
        const dayBatches = batches.filter(x => x.date === b.date);
        const allRows = dayBatches.flatMap(x => x.rows);
        const total = allRows.reduce((s, r) => s + Number(r.지급금액 || 0), 0);
        enriched.push({
          key: `__date__${b.date}`,
          rows: allRows, date: b.date, time: "", total,
          label: `📋 ${b.date} 전체 합산 (${dateCounts[b.date]}회차 · ${allRows.length}건 · ${formatNumber(total)}원)`,
        });
      }
    }
    enriched.push(b);
  });
  return enriched;
}

function buildCompletedTableHtml(rows) {
  if (!rows.length) return `<tr><td colspan="7" class="empty-state">완료된 결제이력이 없습니다.</td></tr>`;
  return rows.map(row => `
    <tr>
      <td>${row.거래처명 || "-"}</td>
      <td>${formatPlanLabel(row.지급일자 || "")}</td>
      <td>${row.은행 || "-"}</td>
      <td>${row.계좌번호 || "-"}</td>
      <td>${row.예금주 || "-"}</td>
      <td>${row.적요 || "-"}</td>
      <td class="numeric-cell">${formatNumber(row.지급금액 || 0)}</td>
    </tr>
  `).join("");
}

function openCompletedReportModal() {
  closeCompletedReportModal();
  const batches = getCompletedBatches();
  let selectedBatch = batches[0] || null;

  const overlay = document.createElement("div");
  overlay.className = "completed-report-overlay payment-report-overlay";

  function renderContent() {
    const rows = selectedBatch ? selectedBatch.rows : [];
    const total = rows.reduce((s, r) => s + Number(r.지급금액 || 0), 0);
    const summaryText = rows.length ? `${rows.length}건 · ${formatNumber(total)}원` : "완료 이력 없음";
    const batchSelector = batches.length > 1
      ? `<select class="completed-batch-select" style="font-size:13px;padding:3px 6px;border-radius:6px;border:1px solid #cbd5e1;margin-right:6px;">
          ${batches.map((b, i) => `<option value="${i}" ${b.key === selectedBatch?.key ? "selected" : ""}>${b.label}</option>`).join("")}
        </select>`
      : (selectedBatch ? `<span style="font-size:12px;color:#64748b;margin-right:8px;">${selectedBatch.label}</span>` : "");

    overlay.innerHTML = `
      <div class="payment-report-popover" role="dialog" aria-modal="true">
        <div class="payment-report-header">
          <div>
            <h3>최종 결재 보고서</h3>
            <p class="completed-summary-text">${summaryText}</p>
          </div>
          <div class="payment-report-actions">
            ${batchSelector}
            <button type="button" class="completed-html-button">최종 HTML 복사</button>
            <button type="button" class="completed-close-button">닫기</button>
          </div>
        </div>
        <p class="payment-report-note">회차를 선택해 각 완료 처리 시점의 결제 내역을 확인합니다.</p>
        <div class="payment-report-table-wrap">
          <table class="payment-report-table">
            <thead>
              <tr>
                <th>업체명</th><th>지급일</th><th>은행</th><th>계좌번호</th><th>예금주</th><th>적요</th><th class="numeric-header">금액</th>
              </tr>
            </thead>
            <tbody>${buildCompletedTableHtml(rows)}</tbody>
          </table>
        </div>
      </div>
    `;
    attachEvents();
  }

  function attachEvents() {
    const popover = overlay.querySelector(".payment-report-popover");

    const select = overlay.querySelector(".completed-batch-select");
    if (select) {
      select.addEventListener("change", () => {
        selectedBatch = batches[Number(select.value)] || null;
        renderContent();
        positionPopover();
      });
    }

    overlay.querySelector(".completed-close-button").addEventListener("click", closeCompletedReportModal);
    overlay.querySelector(".completed-html-button").addEventListener("click", async () => {
      const button = overlay.querySelector(".completed-html-button");
      try {
        const rows = selectedBatch ? selectedBatch.rows : [];
        const isDateBatch = selectedBatch?.key?.startsWith("__date__");
        const title = isDateBatch
          ? `최종 결재 보고서 (${selectedBatch.date} 전체 합산)`
          : `최종 결재 보고서${selectedBatch?.date ? " (" + selectedBatch.date + (selectedBatch.time ? " " + selectedBatch.time : "") + ")" : ""}`;
        const html = buildCompletedApprovalHtmlForRows(rows, title);
        await navigator.clipboard.writeText(html);
        button.textContent = "HTML 복사 완료";
        window.setTimeout(() => {
          if (document.body.contains(button)) button.textContent = "최종 HTML 복사";
        }, 1600);
      } catch (error) {
        console.warn("최종 보고서 HTML 복사 실패:", error);
        button.textContent = "복사 실패";
        window.setTimeout(() => {
          if (document.body.contains(button)) button.textContent = "최종 HTML 복사";
        }, 1600);
      }
    });
    overlay.addEventListener("mousedown", event => {
      if (!popover.contains(event.target)) closeCompletedReportModal();
    });
  }

  function positionPopover() {
    const popover = overlay.querySelector(".payment-report-popover");
    if (!popover) return;
    const width = Math.min(window.innerWidth - 24, 1080);
    popover.style.width = `${width}px`;
    popover.style.left = `${Math.max(12, (window.innerWidth - width) / 2)}px`;
    popover.style.top = `${Math.max(12, (window.innerHeight - Math.min(window.innerHeight - 24, popover.offsetHeight || 640)) / 2)}px`;
  }

  document.body.appendChild(overlay);
  renderContent();
  const reposition = () => positionPopover();
  window.addEventListener("resize", reposition);
  window.addEventListener("scroll", reposition, true);
  overlay.cleanup = () => {
    window.removeEventListener("resize", reposition);
    window.removeEventListener("scroll", reposition, true);
  };
  positionPopover();
}

function openPaymentReportModal(planKey = "__total__", triggerElement = null) {
  closePaymentReportModal();
  const reportRows = buildPlannedPaymentReportRows(planKey);
  const totalAmount = reportRows.reduce((sum, row) => sum + Number(row.지급금액 || 0), 0);
  const reportWarnings = getPaymentReportWarnings(reportRows);
  const overlay = document.createElement("div");
  overlay.className = "payment-report-overlay";
  overlay.innerHTML = `
    <div class="payment-report-popover" role="dialog" aria-modal="true">
      <div class="payment-report-header">
        <div>
          <h3>${planKey === "__total__" ? "전체 예정 보고서" : `${formatPlanLabel(planKey)} 결제 보고서`}</h3>
          <p>${reportRows.length}개 업체 · ${formatNumber(totalAmount)}원</p>
        </div>
        <div class="payment-report-actions">
          <button type="button" class="report-selected-plan-button" style="display:none;">선택 계획 변경</button>
          <button type="button" class="report-html-button">결재용 HTML 복사</button>
          <button type="button" class="report-completed-button">최종 보고서</button>
          <button type="button" class="report-bank-export-button">우리은행 양식 저장</button>
          <button type="button" class="report-plan-edit-button">일괄 계획 변경</button>
          <button type="button" class="report-complete-button">완료 처리</button>
          <button type="button" class="report-close-button">닫기</button>
        </div>
      </div>
      <p class="payment-report-note">메일플러그 전자결재에는 '결재용 HTML 복사'를, 은행 업로드에는 '우리은행 양식 저장'을 사용하면 됩니다.</p>
      ${reportWarnings.length ? `
        <div class="payment-report-warning-box">
          <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;">
            <strong>은행 업로드 전 확인 필요</strong>
            <span>${reportWarnings.length}개 업체에 은행코드, 계좌번호, 예금주 누락이 있습니다.</span>
            <button type="button" class="report-warning-email-btn"
              style="margin-left:auto;background:#1e40af;color:white;border:none;border-radius:6px;padding:4px 11px;font-size:12px;cursor:pointer;white-space:nowrap;">
              ✉ 확인 요청 메일
            </button>
          </div>
          <div class="payment-report-warning-list">
            ${reportWarnings.slice(0, 8).map(item => `<span>${item.거래처명}: ${item.missing.join(", ")}</span>`).join("")}
            ${reportWarnings.length > 8 ? `<span>외 ${reportWarnings.length - 8}개 업체</span>` : ""}
          </div>
        </div>
      ` : ""}
      <div class="payment-report-table-wrap">
        <table class="payment-report-table">
          <thead>
            <tr>
              <th style="width:32px;"><input type="checkbox" class="report-select-all" title="전체 선택" /></th>
              <th>업체명</th>
              <th>예정일</th>
              <th>은행</th>
              <th>계좌번호</th>
              <th>예금주</th>
              <th>대상 연월</th>
              <th class="numeric-header">금액</th>
            </tr>
          </thead>
          <tbody>
            ${reportRows.length ? reportRows.map((row, idx) => {
    const holderRaw = String(row.예금주 || "").trim();
    const repRaw = String(row.대표자명 || "").trim();
    const holderNorm = holderRaw.replace(/\s+/g, "");
    const repNorm = repRaw.replace(/\s+/g, "");
    const vendorNorm = String(row.거래처명 || "").replace(/\s+/g, "");
    const holderIsPersonName = isPersonName(holderRaw);
    const mismatch = holderIsPersonName && repNorm && holderNorm !== repNorm && holderNorm !== vendorNorm;
    const holderHtml = holderRaw
      ? (mismatch
        ? `<span class="report-holder-mismatch" title="대표자명: ${escapeHtml(repRaw)}">${escapeHtml(holderRaw)}</span>`
        : escapeHtml(holderRaw))
      : '<span class="report-missing">확인 필요</span>';
    const partnerKey = encodeURIComponent(`${row.거래처코드 || ""}||${row.거래처명 || ""}`);
    const firstMonthKey = row.연월키목록?.[0] || "";
    return `
              <tr data-row-idx="${idx}">
                <td><input type="checkbox" class="report-row-check" data-idx="${idx}" /></td>
                <td><button type="button" class="report-vendor-link" data-partner-key="${partnerKey}" data-month-key="${firstMonthKey}">${escapeHtml(row.거래처명 || "-")}</button></td>
                <td>${formatPlanLabel(row.결제예정일 || "")}</td>
                <td>${row.은행 || '<span class="report-missing">확인 필요</span>'}</td>
                <td>${row.계좌번호 || '<span class="report-missing">확인 필요</span>'}</td>
                <td>${holderHtml}</td>
                <td>${escapeHtml(row.연월목록 || "-")}</td>
                <td class="numeric-cell">${formatNumber(row.지급금액 || 0)}</td>
              </tr>`;
  }).join("") : `<tr><td colspan="8" class="empty-state">보고서로 만들 결제 대상이 없습니다.</td></tr>`}
          </tbody>
        </table>
      </div>
    </div>
  `;

  document.body.appendChild(overlay);
  const popover = overlay.querySelector(".payment-report-popover");

  function positionPopover() {
    const width = Math.min(window.innerWidth - 24, 1080);
    popover.style.width = `${width}px`;
    popover.style.left = `${Math.max(12, (window.innerWidth - width) / 2)}px`;
    popover.style.top = `${Math.max(12, (window.innerHeight - Math.min(window.innerHeight - 24, popover.offsetHeight || 640)) / 2)}px`;
  }

  // ── 체크박스 선택 관리
  const selectedIndices = new Set();
  const selectedPlanBtn = overlay.querySelector(".report-selected-plan-button");

  function updateSelectedPlanBtn() {
    selectedPlanBtn.style.display = selectedIndices.size > 0 ? "" : "none";
    selectedPlanBtn.textContent = `선택 계획 변경 (${selectedIndices.size}건)`;
  }

  overlay.querySelector(".report-select-all").addEventListener("change", e => {
    overlay.querySelectorAll(".report-row-check").forEach(cb => {
      cb.checked = e.target.checked;
      const idx = Number(cb.dataset.idx);
      if (e.target.checked) selectedIndices.add(idx);
      else selectedIndices.delete(idx);
    });
    updateSelectedPlanBtn();
  });

  overlay.querySelectorAll(".report-row-check").forEach(cb => {
    cb.addEventListener("change", () => {
      const idx = Number(cb.dataset.idx);
      if (cb.checked) selectedIndices.add(idx);
      else selectedIndices.delete(idx);
      updateSelectedPlanBtn();
    });
  });

  selectedPlanBtn.addEventListener("click", () => {
    const selectedRows = reportRows.filter((_, i) => selectedIndices.has(i));
    const selectedVendorKeys = new Set(selectedRows.map(r => `${r.거래처코드 || ""}||${r.거래처명 || ""}`));
    const items = payables.filter(item => selectedVendorKeys.has(getPartnerGroupKey(item)));
    if (!items.length) return;
    closePaymentReportModal();
    openBatchPlanEditor("선택 업체", items, document.querySelector(".payment-plan-summary-grid") || document.body);
  });

  // ── 업체명 클릭 → 미지급 테이블 해당 행으로 이동
  overlay.querySelectorAll(".report-vendor-link").forEach(btn => {
    btn.addEventListener("click", () => {
      const partnerKey = decodeURIComponent(btn.dataset.partnerKey || "");
      const monthKey = btn.dataset.monthKey || "";
      closePaymentReportModal();
      const vendorItems = payables.filter(item => getPartnerGroupKey(item) === partnerKey);
      vendorItems.forEach(item => {
        const group = getDueGroup(item);
        if (group) payablesGroupState.collapsed[group] = false;
      });
      switchTab("payables");
      rerenderAll();
      window.requestAnimationFrame(() => {
        const encodedKey = encodeURIComponent(partnerKey);
        const targetBtn = elements.payables.querySelector(
          monthKey
            ? `.edit-amount-button[data-partner-key="${encodedKey}"][data-month-key="${monthKey}"]`
            : `.payable-select-checkbox[data-partner-key="${encodedKey}"]`
        );
        if (!targetBtn) return;
        const row = targetBtn.closest("tr");
        if (!row) return;
        const tableResponsive = elements.payables.querySelector(".table-responsive");
        if (tableResponsive) {
          const cellEl = targetBtn.closest("td");
          const containerRect = tableResponsive.getBoundingClientRect();
          if (cellEl) {
            const cellRect = cellEl.getBoundingClientRect();
            tableResponsive.scrollLeft = Math.max(0,
              tableResponsive.scrollLeft + (cellRect.left - containerRect.left) - tableResponsive.clientWidth / 3
            );
          }
          const rowRect = row.getBoundingClientRect();
          tableResponsive.scrollTop = Math.max(0,
            tableResponsive.scrollTop + (rowRect.top - containerRect.top) - tableResponsive.clientHeight / 3
          );
        }
        const tds = Array.from(row.querySelectorAll("td"));
        tds.forEach(td => { td.style.backgroundColor = "#fef08a"; td.style.transition = "background-color 1s ease"; });
        setTimeout(() => {
          tds.forEach(td => { td.style.backgroundColor = ""; });
          setTimeout(() => { tds.forEach(td => { td.style.transition = ""; }); }, 1000);
        }, 4000);
      });
    });
  });

  overlay.querySelector(".report-close-button").addEventListener("click", closePaymentReportModal);

  const warningEmailBtn = overlay.querySelector(".report-warning-email-btn");
  if (warningEmailBtn) {
    warningEmailBtn.addEventListener("click", () => {
      openWarningEmailDialog(reportWarnings, reportRows, planKey);
    });
  }

  overlay.querySelector(".report-completed-button").addEventListener("click", () => {
    openCompletedReportModal();
  });
  overlay.querySelector(".report-html-button").addEventListener("click", async () => {
    const button = overlay.querySelector(".report-html-button");
    try {
      await copyPaymentApprovalHtml(planKey);
      button.textContent = "HTML 복사 완료";
      window.setTimeout(() => {
        if (document.body.contains(button)) {
          button.textContent = "결재용 HTML 복사";
        }
      }, 1600);
    } catch (error) {
      console.warn("결재용 HTML 복사 실패:", error);
      button.textContent = "복사 실패";
      window.setTimeout(() => {
        if (document.body.contains(button)) {
          button.textContent = "결재용 HTML 복사";
        }
      }, 1600);
    }
  });
  overlay.querySelector(".report-bank-export-button").addEventListener("click", async () => {
    const button = overlay.querySelector(".report-bank-export-button");
    try {
      await downloadWooriTransferTemplate(planKey);
      button.textContent = "양식 저장 완료";
      window.setTimeout(() => {
        if (document.body.contains(button)) {
          button.textContent = "우리은행 양식 저장";
        }
      }, 1600);
    } catch (error) {
      console.warn("우리은행 양식 저장 실패:", error);
      button.textContent = "정보 확인 필요";
      window.setTimeout(() => {
        if (document.body.contains(button)) {
          button.textContent = "우리은행 양식 저장";
        }
      }, 1600);
    }
  });
  overlay.querySelector(".report-plan-edit-button").addEventListener("click", () => {
    closePaymentReportModal();
    const filteredPayables = getFilteredItems(payables, "payables");
    openBatchPlanEditor(planKey, getPayablesForPlanKey(planKey, filteredPayables), triggerElement || document.body);
  });
  overlay.querySelector(".report-complete-button").addEventListener("click", async () => {
    const button = overlay.querySelector(".report-complete-button");
    try {
      const result = await markPlanAsCompleted(planKey);
      closePaymentReportModal();
      preserveViewport(() => rerenderAll());
      console.info(`결제 완료 처리: ${result.count}건`);
    } catch (error) {
      console.warn(error);
      button.textContent = "저장 실패";
      window.setTimeout(() => {
        if (document.body.contains(button)) {
          button.textContent = "완료 처리";
        }
      }, 1800);
    }
  });
  overlay.addEventListener("mousedown", event => {
    if (!popover.contains(event.target)) {
      closePaymentReportModal();
    }
  });
  const reposition = () => positionPopover();
  window.addEventListener("resize", reposition);
  window.addEventListener("scroll", reposition, true);
  overlay.cleanup = () => {
    window.removeEventListener("resize", reposition);
    window.removeEventListener("scroll", reposition, true);
  };
  positionPopover();
}

function diffVendorMasterRows(existingRows, importedRows) {
  const existingMap = new Map(existingRows.map(row => [getVendorMatchKey(row), row]));
  const comparedRows = importedRows.map(row => {
    const key = getVendorMatchKey(row);
    const existing = existingMap.get(key);
    if (!existing) {
      return { kind: "new", row, changes: ["신규 업체"] };
    }
    const changeFields = [
      ["거래처명", "거래처명"],
      ["거래처구분", "거래처구분"],
      ["대표자명", "대표자명"],
      ["사업자번호", "사업자번호"],
      ["전화번호", "전화번호"],
      ["주소", "주소"],
      ["은행", "은행"],
      ["계좌번호", "계좌번호"],
      ["예금주", "예금주"],
    ]
      .filter(([field, label]) => !isFuzzySame(existing[field], row[field], field))
      .map(([, label]) => label);

    return {
      kind: changeFields.length ? "updated" : "same",
      row,
      existing,
      changes: changeFields,
    };
  });

  return {
    comparedRows,
    stats: {
      total: importedRows.length,
      added: comparedRows.filter(item => item.kind === "new").length,
      updated: comparedRows.filter(item => item.kind === "updated").length,
      same: comparedRows.filter(item => item.kind === "same").length,
    },
  };
}

function getActivePayableVendorCodeSet() {
  return new Set(
    payables
      .map(item => normalizeVendorCode(item.codeNormalized || item.code || item.codeRaw || ""))
      .filter(Boolean),
  );
}

function getActiveReceivableVendorCodeSet() {
  return new Set(
    receivables
      .map(item => normalizeVendorCode(item.code || item.codeRaw || ""))
      .filter(Boolean),
  );
}

function renderVendorMasterPanel() {
  if (!elements.vendorMasterPanel) return;
  const hasRows = vendorMasterState.comparedRows.length > 0;
  elements.vendorMasterPanel.classList.toggle("hidden", !hasRows);
  if (!hasRows) {
    elements.vendorMasterPanel.innerHTML = "";
    return;
  }
  const stats = vendorMasterState.stats || { total: 0, added: 0, updated: 0, same: 0 };
  const groupedRows = {
    new: vendorMasterState.comparedRows.filter(item => item.kind === "new"),
    updated: vendorMasterState.comparedRows.filter(item => item.kind === "updated"),
    same: vendorMasterState.comparedRows.filter(item => item.kind === "same"),
  };
  const getUpdatedChangeLines = item => {
    const fields = [
      ["거래처명", "거래처명"],
      ["거래처구분", "거래처구분"],
      ["은행", "은행"],
      ["계좌번호", "계좌번호"],
      ["예금주", "예금주"],
      ["사업자번호", "사업자번호"],
      ["대표자명", "대표자명"],
      ["전화번호", "전화번호"],
      ["주소", "주소"],
    ];
    return fields
      .filter(([field]) => !isFuzzySame(item.existing?.[field], item.row?.[field], field))
      .map(([field, label]) => `
        <div class="vendor-master-change-line">
          <span class="field">${label}</span>
          <span class="before">${item.existing?.[field] || "-"}</span>
          <span class="arrow">→</span>
          <span class="after">${item.row?.[field] || "-"}</span>
        </div>
      `)
      .join("");
  };

  const renderGroupedSection = (kind, title, rows) => `
    <details class="vendor-master-section ${kind}">
      <summary>
        <span class="vendor-master-section-title">${title}</span>
        <span class="vendor-master-section-count">${rows.length}건</span>
      </summary>
      <div class="vendor-master-preview">
        ${rows.length ? rows.map(item => `
          <div class="vendor-master-preview-row ${item.kind}">
            <div class="kind">${item.kind === "new" ? "신규" : item.kind === "updated" ? "변경" : "동일"}</div>
            <div>${item.row.거래처명 || "-"}</div>
            <div>${item.row.거래처코드_norm || "-"}</div>
            <div>${item.kind === "updated"
      ? `<div class="vendor-master-change-list">${getUpdatedChangeLines(item) || '<div class="vendor-master-empty">변경 없음</div>'}</div>`
      : (item.changes?.length ? item.changes.join(", ") : "변경 없음")}</div>
          </div>
        `).join("") : `<div class="vendor-master-empty">해당 항목이 없습니다.</div>`}
      </div>
    </details>
  `;
  elements.vendorMasterPanel.innerHTML = `
    <div class="vendor-master-panel-header">
      <div>
        <h3>업체마스터 업로드 결과</h3>
        <div class="vendor-master-panel-meta">${vendorMasterState.lastFileName || "업로드 파일"} 기준 비교 결과입니다. 신규/변경 항목만 시트에 반영합니다.</div>
      </div>
      <div class="vendor-master-actions">
        <button type="button" class="vendor-master-save-button" ${vendorMasterState.saving ? "disabled" : ""}>업체마스터 반영</button>
        <button type="button" class="vendor-master-close-button">닫기</button>
        <span class="vendor-master-status">${vendorMasterState.lastMessage || "신규/변경 항목만 시트에 반영합니다."}</span>
      </div>
    </div>
    <div class="vendor-master-stats">
      <div class="vendor-master-stat"><span>전체</span><strong>${stats.total}</strong></div>
      <div class="vendor-master-stat"><span>신규</span><strong>${stats.added}</strong></div>
      <div class="vendor-master-stat"><span>변경</span><strong>${stats.updated}</strong></div>
      <div class="vendor-master-stat"><span>동일</span><strong>${stats.same}</strong></div>
    </div>
    <div class="vendor-master-sections">
      ${renderGroupedSection("new", "신규", groupedRows.new)}
      ${renderGroupedSection("updated", "변경", groupedRows.updated)}
      ${renderGroupedSection("same", "동일", groupedRows.same)}
    </div>
  `;

  const saveButton = elements.vendorMasterPanel.querySelector(".vendor-master-save-button");
  if (saveButton) {
    saveButton.addEventListener("click", saveVendorMasterRows);
  }
  const closeButton = elements.vendorMasterPanel.querySelector(".vendor-master-close-button");
  if (closeButton) {
    closeButton.addEventListener("click", () => {
      vendorMasterState.comparedRows = [];
      renderVendorMasterPanel();
    });
  }
}

async function saveVendorMasterRows() {
  const targetRows = vendorMasterState.comparedRows
    .filter(item => item.kind === "new" || item.kind === "updated")
    .map(item => {
      const r = item.row;
      return {
        ...r,
        vendor_id: r.vendor_id || "",
        거래처코드_norm: r.거래처코드_norm || "",
        사업자번호: r.사업자번호 || "",
        계좌번호: r.계좌번호 || "",
        active_yn: "Y",
        last_imported_at: new Date().toISOString(),
        last_changed_at: item.kind === "updated" ? new Date().toISOString() : "",
        change_note: item.changes?.join(", ") || "신규 등록",
      };
    });

  if (!targetRows.length) {
    vendorMasterState.lastMessage = "반영할 신규/변경 항목이 없습니다.";
    renderVendorMasterPanel();
    return;
  }

  vendorMasterState.saving = true;
  renderVendorMasterPanel();
  try {
    const BATCH = 1000;
    const PARALLEL = 3;
    const total = targetRows.length;
    const batches = [];
    for (let i = 0; i < total; i += BATCH) batches.push(targetRows.slice(i, i + BATCH));
    let saved = 0;
    for (let i = 0; i < batches.length; i += PARALLEL) {
      const group = batches.slice(i, i + PARALLEL);
      vendorMasterState.lastMessage = `저장 중… ${saved} / ${total}건`;
      renderVendorMasterPanel();
      await Promise.all(group.map(b => postSheetWebApp("upsertVendorMaster", { sheetName: MASTER_SHEET_NAME, rows: b })));
      saved += group.reduce((s, b) => s + b.length, 0);
    }
    setVendorMasterRows([
      ...vendorMasterState.rows.filter(existing => !targetRows.some(next => getVendorMatchKey(next) === getVendorMatchKey(existing))),
      ...targetRows,
    ]);
    enrichPayablesWithVendorMaster();
    enrichPayablesWithManagerDays();
    vendorMasterState.lastMessage = `${total}건을 업체마스터에 반영했습니다.`;
  } catch (error) {
    vendorMasterState.lastMessage = `저장 실패: ${error.message}`;
  } finally {
    vendorMasterState.saving = false;
    renderVendorMasterPanel();
    rerenderAll();
  }
}

async function handleVendorMasterFile(file) {
  if (!file) return;
  if (typeof XLSX === "undefined") {
    vendorMasterState.lastMessage = "엑셀 라이브러리를 불러오지 못했습니다.";
    renderVendorMasterPanel();
    return;
  }

  const arrayBuffer = await file.arrayBuffer();
  const workbook = XLSX.read(arrayBuffer, { type: "array" });

  // ── 업체마스터 (미지급) ─────────────────────────────────
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const rawRows = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
  const allImportedRows = parseVendorMasterSheetRows(rawRows);
  const existingRows = parseVendorMasterSheetRows(await fetchVendorMasterRowsFromApi());
  // 신규 + 기존 모두 upsert 대상 (전체 마스터 갱신)
  const { comparedRows, stats } = diffVendorMasterRows(existingRows, allImportedRows);

  vendorMasterState.importedRows = allImportedRows;
  vendorMasterState.comparedRows = comparedRows;
  vendorMasterState.stats = stats;
  vendorMasterState.lastFileName = file.name;
  vendorMasterState.lastMessage = `파일 ${allImportedRows.length}건 비교 완료 (신규 ${stats.added}건 / 변경 ${stats.updated}건 / 동일 ${stats.same}건).`;
  renderVendorMasterPanel();

  // ── 담당자 마스터 (미수금) ───────────────────────────────
  const mgrSheetName = workbook.SheetNames.find(n => n === "담당자");
  if (mgrSheetName && receivables.length) {
    const mgrSheet = workbook.Sheets[mgrSheetName];
    const mgrRawRows = XLSX.utils.sheet_to_json(mgrSheet, { defval: "" });
    const activeRcvCodes = getActiveReceivableVendorCodeSet();
    const filteredMgrRows = mgrRawRows.filter(row => {
      const code = normalizeVendorCode(String(row["거래처코드"] || row["code"] || "").trim());
      return activeRcvCodes.has(code);
    });
    setManagerMasterRows(filteredMgrRows);
    enrichReceivablesWithManager();
    renderReceivables();
  }
}

function setupVendorMasterImport() {
  if (!elements.vendorMasterImportButton || !elements.vendorMasterFileInput) return;
  elements.vendorMasterImportButton.addEventListener("click", () => {
    elements.vendorMasterFileInput.click();
  });
  elements.vendorMasterFileInput.addEventListener("change", async event => {
    const [file] = event.target.files || [];
    if (!file) return;
    await handleVendorMasterFile(file);
    event.target.value = "";
  });
}

async function importVendorsFromLedger() {
  const btn = document.getElementById("ledgerVendorImportButton");
  const setLabel = t => { if (btn) btn.textContent = t; };
  const setDisabled = v => { if (btn) btn.disabled = v; };

  setDisabled(true);
  setLabel("조회 중…");
  try {
    const [lSales, lPurchase, lPayable] = await Promise.all([
      fetchApiRows("getLedgerSales"),
      fetchApiRows("getLedgerPurchase"),
      fetchApiRows("getLedgerPayable"),
    ]);

    // 계정별원장 전체에서 고유 거래처 추출
    const vendorMap = new Map();
    [...lSales, ...lPurchase, ...lPayable].forEach(r => {
      const rawCode = String(r["거래처코드"] || "").trim();
      const name = String(r["거래처명"] || "").trim();
      if (!rawCode && !name) return;
      const norm = normalizeVendorCode(rawCode);
      if (norm && !vendorMap.has(norm)) {
        vendorMap.set(norm, { 거래처코드_raw: rawCode, 거래처코드_norm: norm, 거래처명: name });
      }
    });

    if (!vendorMap.size) {
      alert("계정별원장에서 거래처를 찾지 못했습니다. 먼저 자료업로드를 해주세요.");
      return;
    }

    // 기존 업체마스터와 비교 — 이미 있는 코드는 건너뜀
    setLabel("비교 중…");
    const existingRows = parseVendorMasterSheetRows(await fetchVendorMasterRowsFromApi());
    const existingCodes = new Set(existingRows.map(r => r.거래처코드_norm).filter(Boolean));
    const newRows = [...vendorMap.values()].filter(r => !existingCodes.has(r.거래처코드_norm));

    if (!newRows.length) {
      alert("추가할 새 거래처가 없습니다. 계정별원장의 거래처가 이미 모두 업체마스터에 있습니다.");
      return;
    }

    const BATCH = 1000;
    const PARALLEL = 3;
    const total = newRows.length;
    const mappedRows = newRows.map(r => ({
      ...r,
      vendor_id: r.거래처코드_norm,
      사업자번호: "",
      active_yn: "Y",
      last_imported_at: new Date().toISOString(),
    }));
    const batches = [];
    for (let i = 0; i < total; i += BATCH) batches.push(mappedRows.slice(i, i + BATCH));
    let saved = 0;
    for (let i = 0; i < batches.length; i += PARALLEL) {
      const group = batches.slice(i, i + PARALLEL);
      setLabel(`저장 중… ${saved}/${total}`);
      await Promise.all(group.map(b => postSheetWebApp("upsertVendorMaster", { sheetName: MASTER_SHEET_NAME, rows: b })));
      saved += group.reduce((s, b) => s + b.length, 0);
    }

    // 메모리 갱신
    setVendorMasterRows([...existingRows, ...newRows]);
    enrichPayablesWithVendorMaster();
    alert(`완료: ${total}개 거래처를 업체마스터에 추가했습니다.`);
  } catch (err) {
    alert(`실패: ${err.message}`);
  } finally {
    setDisabled(false);
    setLabel("원장→업체마스터");
  }
}

function setupLedgerVendorImport() {
  const btn = document.getElementById("ledgerVendorImportButton");
  if (!btn) return;
  btn.addEventListener("click", importVendorsFromLedger);
}

function applySavedPaymentPlansFromApi(rows) {
  if (!Array.isArray(rows) || !rows.length) return;

  // 히스토리 초기화 및 그룹화 (동일 sourceKey에 여러 개의 누적된 기록)
  Object.keys(payablePlanHistories).forEach(k => delete payablePlanHistories[k]);

  const bySourceKey = rows.reduce((acc, row) => {
    const sourceKey = String(row.source_key || row.sourceKey || "").trim();
    if (sourceKey) {
      if (!payablePlanHistories[sourceKey]) payablePlanHistories[sourceKey] = [];
      payablePlanHistories[sourceKey].push(row);

      const existing = acc[sourceKey];
      if (!existing) {
        acc[sourceKey] = row;
      } else {
        // 더 최신 데이터(updated_at 기준)로 덮어쓰기
        const tNew = new Date(row.updated_at || 0).getTime();
        const tOld = new Date(existing.updated_at || 0).getTime();
        if (tNew >= tOld) acc[sourceKey] = row;
      }
    }
    return acc;
  }, {});

  // 각 sourceKey 배열 내에서도 정렬 (최신순)
  Object.values(payablePlanHistories).forEach(arr => {
    arr.sort((a, b) => new Date(b.updated_at || 0).getTime() - new Date(a.updated_at || 0).getTime());
  });

  // 로컬 상태의 타임스탬프와 비교하기 위해 로컬 스냅샷 로드
  const localMap = loadPayablesStateFromLocal();

  payables = payables.map(item => {
    const saved = bySourceKey[item.sourceKey || ""];
    if (!saved) return item;

    // 로컬 상태가 더 최신이면 원격 데이터 무시 (방어 로직)
    const localItem = localMap[item.sourceKey || ""];
    if (localItem && localItem.updatedAt && saved.updated_at) {
      const localTime = new Date(localItem.updatedAt).getTime();
      const remoteTime = new Date(saved.updated_at).getTime();
      if (localTime > remoteTime) return item;
    }

    const rawOutstanding = Math.max(0, Number(item.purchase || 0) - Number(item.paid || 0));
    const savedStatus = saved.plan_status || saved.completionStatus || item.completionStatus || "";
    const prevWasComplete = savedStatus === "완료";
    const rawIsStillOpen = rawOutstanding > 0;
    const effectiveStatus = (prevWasComplete && rawIsStillOpen) ? "" : savedStatus;
    const rawPaid = Number(item.paid || 0);
    const savedPO = saved.paid_override != null ? Number(saved.paid_override) : null;
    const effectivePO = (prevWasComplete && rawIsStillOpen)
      ? item.paidOverride
      : (rawPaid > 0 && savedPO != null && savedPO > rawPaid)
        ? rawPaid  // API에 저장된 이중계산 값 리셋
        : (savedPO ?? item.paidOverride);
    const resetByRaw = prevWasComplete && rawIsStillOpen;
    const isProtectedStatus = effectiveStatus === "보류" || effectiveStatus === "완료" || effectiveStatus === "부분결제";
    const noPayment = rawPaid === 0 && (effectivePO == null || effectivePO === 0);
    const savedDA = saved.decision_amount != null ? Number(saved.decision_amount) : null;
    const shouldResetDA = !resetByRaw && !isProtectedStatus && noPayment &&
      savedDA !== null && savedDA !== item.decisionAmount;
    return {
      ...item,
      decisionAmount: (resetByRaw || shouldResetDA) ? item.decisionAmount : (savedDA ?? item.decisionAmount),
      paymentPlan: saved.payment_plan != null ? normalizeDateValue(saved.payment_plan) : item.paymentPlan,
      selected: saved.selected != null ? String(saved.selected) === "true" || saved.selected === true : item.selected,
      paidOverride: effectivePO,
      completionStatus: effectiveStatus,
      _rawResetCompletion: resetByRaw ? true : item._rawResetCompletion,
    };
  });
}

function buildPaymentPlanRows() {
  return payables.map(item => ({
    source_key: item.sourceKey || buildPayableSourceKey(item),
    거래처코드_norm: normalizeVendorCode(item.codeNormalized || item.code || item.codeRaw || ""),
    거래처명: item.name || "",
    작성연도: Number(item.year || 0),
    작성월: Number(item.month || 0),
    원금액: Number(item.purchase || 0),
    지급합: Number(item.paid || 0),
    잔액: Number(getPayableOutstanding(item)),
    decision_amount: Number(item.decisionAmount ?? 0),
    payment_plan: item.paymentPlan || "",
    plan_status: item.completionStatus || (item.paymentPlan === "보류" ? "보류" : item.paymentPlan ? "예정" : "미정"),
    selected: Boolean(item.selected),
    paid_override: Number(item.paidOverride ?? item.paid ?? 0),
    memo: item.memo || "",
    updated_at: new Date().toISOString(),
  }));
}

async function flushPayablesStateToApi() {
  if (payablesSyncState.inFlight) {
    payablesSyncState.pending = true;
    return;
  }
  payablesSyncState.inFlight = true;
  payablesSyncState.pending = false;
  try {
    await postSheetWebApp("appendPaymentPlans", {
      sheetName: PLAN_SHEET_NAME,
      rows: buildPaymentPlanRows(),
    });
    payablesSyncState.lastError = "";
  } catch (error) {
    payablesSyncState.lastError = error.message;
    console.warn("결제계획 원격 저장 실패, 로컬 저장만 유지합니다.", error);
  } finally {
    payablesSyncState.inFlight = false;
    if (payablesSyncState.pending) {
      payablesSyncState.pending = false;
      schedulePayablesStateSync();
    }
  }
}

function schedulePayablesStateSync() {
  if (payablesSyncState.timeoutId) {
    clearTimeout(payablesSyncState.timeoutId);
  }
  payablesSyncState.timeoutId = window.setTimeout(() => {
    payablesSyncState.timeoutId = null;
    flushPayablesStateToApi();
  }, PAYABLES_SYNC_DEBOUNCE_MS);
}

function persistPayablesState() {
  savePayablesStateToLocal();
  schedulePayablesStateSync();
}

function applySavedPayablesState(items) {
  const savedMap = loadPayablesStateFromLocal();
  // stable_key 역인덱스: raw 교체 후 source_key가 달라져도 계획 복원
  const savedByStableKey = {};
  // 구버전 3파트 stableKey 호환: null = 충돌(동일 거래처+연월 복수 그룹)이므로 무시
  const savedByLegacyKey = {};
  Object.values(savedMap).forEach(v => {
    if (!v.stableKey) return;
    if (!savedByStableKey[v.stableKey]) savedByStableKey[v.stableKey] = v;
    const legacyKey = v.stableKey.split("||").slice(0, 3).join("||");
    if (!(legacyKey in savedByLegacyKey)) savedByLegacyKey[legacyKey] = v;
    else savedByLegacyKey[legacyKey] = null; // 충돌 → 사용 불가
  });
  return items.map(item => {
    const sourceKey = item.sourceKey || buildPayableSourceKey(item);
    const stableKey = buildPayableStableKey(item);
    const legacyKey = stableKey.split("||").slice(0, 3).join("||");
    const saved = savedMap[sourceKey]
      || savedByStableKey[stableKey]
      || savedByLegacyKey[legacyKey]   // 구버전 저장 데이터 마이그레이션 폴백
      || null;
    if (!saved) {
      return { ...item, sourceKey, stableKey };
    }
    // raw 시트 잔액이 있으면 이전 "완료" 상태 무시 (시트 업데이트 반영)
    const rawPaid = Number(item.paid || 0);
    const rawPurchase = Number(item.purchase || 0);
    const rawOutstanding = Math.max(0, rawPurchase - rawPaid);
    const savedStatus = saved.completionStatus || item.completionStatus || "";
    const prevWasComplete = savedStatus === "완료";
    const rawIsStillOpen = rawOutstanding > 0;
    const effectiveStatus = (prevWasComplete && rawIsStillOpen) ? "" : savedStatus;
    // paidOverride는 raw 지급합보다 클 수 없음 (raw 업데이트로 이미 반영된 이중계산 방지)
    // raw.paid > 0이면 saved.paidOverride를 신뢰하지 않고 raw 값으로 리셋 → applyPaymentHistoryRows가 재계산
    const savedPO = saved.paidOverride != null ? Number(saved.paidOverride) : null;
    const effectivePO = (rawPaid > 0 && savedPO != null && savedPO > rawPaid)
      ? rawPaid  // raw 지급합을 초과한 저장값은 이중계산이므로 raw로 리셋
      : (savedPO ?? item.paidOverride);
    const resetByRaw = prevWasComplete && rawIsStillOpen;
    const isProtectedStatus = effectiveStatus === "보류" || effectiveStatus === "완료" || effectiveStatus === "부분결제";
    const erpBalanceChanged = item.balance > 0 && saved.decisionAmount != null && Number(saved.decisionAmount) !== item.balance;
    const noPayment = rawPaid === 0 && (effectivePO == null || effectivePO === 0);
    // ERP 잔액이 달라진 경우: 완료만 보호 (보류·부분결제도 ERP 잔액으로 갱신)
    const shouldResetDA = !resetByRaw && saved.decisionAmount != null && (
      (!isProtectedStatus && noPayment && Number(saved.decisionAmount) !== item.decisionAmount) ||
      (effectiveStatus !== "완료" && erpBalanceChanged)
    );
    return {
      ...item,
      sourceKey,
      stableKey,
      decisionAmount: (resetByRaw || shouldResetDA) ? item.decisionAmount : (saved.decisionAmount != null ? Number(saved.decisionAmount) : item.decisionAmount),
      paymentPlan: saved.paymentPlan != null ? normalizeDateValue(saved.paymentPlan) : item.paymentPlan,
      selected: saved.selected != null ? Boolean(saved.selected) : item.selected,
      paidOverride: resetByRaw ? rawPaid : effectivePO,
      completionStatus: effectiveStatus,
      _rawResetCompletion: resetByRaw ? true : undefined,
    };
  });
}

// ── 미수금 파싱 / 날짜 계산 ─────────────────────────────────

function calcReceivableDueDate(year, month, memo, condition) {
  year = Number(year); month = Number(month);
  if (!year || !month) return null;
  const cond = String(condition || "").replace("전자어음", "").trim();
  const memoStr = String(memo || "").trim();

  if (["바로", "쇼핑몰+", "오토몰"].includes(cond)) {
    const m = memoStr.match(/(\d{6})~\?/);
    if (!m) return null;
    const s = m[1];
    return new Date(2000 + parseInt(s.slice(0, 2)), parseInt(s.slice(2, 4)) - 1, parseInt(s.slice(4, 6)));
  }
  if (cond === "당말일") return rcvLastDay(year, month);
  const cm = cond.match(/^당(\d+)일$/);
  if (cm) { const [ny, nm] = rcvAddMonths(year, month, 1); return new Date(ny, nm - 1, parseInt(cm[1])); }
  if (cond === "25일") { const [ny, nm] = rcvAddMonths(year, month, 1); return new Date(ny, nm - 1, 25); }
  if (cond === "말일") { const [ny, nm] = rcvAddMonths(year, month, 1); return rcvLastDay(ny, nm); }
  if (cond === "60일") { const [ny, nm] = rcvAddMonths(year, month, 2); return rcvLastDay(ny, nm); }
  const dm = cond.match(/^(\d+)일$/);
  if (dm) { const [ny, nm] = rcvAddMonths(year, month, 2); return new Date(ny, nm - 1, parseInt(dm[1])); }
  return null;
}
function rcvLastDay(y, m) { return new Date(y, m, 0); }
function rcvAddMonths(y, m, n) { const t = m + n; return [y + Math.floor((t - 1) / 12), ((t - 1) % 12) + 1]; }

function parseReceivableRow(row) {
  if (!row || typeof row !== "object") return null;
  const year = Number(row["year"] || row["연도"] || row["년"] || row["작성연도"] || 0);
  const month = Number(row["month"] || row["월"] || row["작성월"] || 0);
  const codeRaw = String(row["code"] || row["코드"] || row["거래처코드"] || "").trim();
  const name = String(row["client"] || row["거래처명"] || row["거래처"] || "").trim();
  const memo = String(row["memo"] || row["매출메모"] || row["메모"] || "").trim();
  const condition = String(row["condition"] || row["일"] || row["수금조건"] || "").trim();
  const sales = parseSheetNumber(row["sales"] || row["합계 : 매출금액"] || row["매출금액"] || row["매출"] || 0);
  const collection = parseSheetNumber(row["collection"] || row["합계 : 수금합"] || row["수금합"] || row["수금"] || 0);
  const balance = parseSheetNumber(row["balance"] || row["잔 액"] || row["잔액"] || 0);

  if (!name || (!Number(balance) && condition !== "제외" && !memo.includes("제외"))) return null;

  const code = normalizeVendorCode(codeRaw || "00000");
  const dueDate = calcReceivableDueDate(year, month, memo, condition);
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const elapsed = dueDate ? Math.floor((today - dueDate) / 86400000) : null;

  return {
    year, month, code, codeRaw, name, memo, condition,
    sales, collection, balance,
    dueDate: dueDate ? dueDate.toISOString().slice(0, 10) : "",
    elapsed,
    manager: "",
    managerEmail: "",
  };
}

function parsePayableRow(row) {
  if (!row || typeof row !== "object") return {
    codeRaw: "",
    codeNormalized: "",
    code: "",
    name: "",
    year: 0,
    month: 0,
    purchase: 0,
    paid: 0,
    payDate: "",
    memo: "",
    selected: false,
    decisionAmount: 0,
    paymentPlan: "",
    sourceKey: "",
    paidOverride: 0,
    completionStatus: "",
  };

  const normalized = {};
  Object.keys(row).forEach(key => {
    normalized[normalizeKey(key)] = row[key];
  });

  const codeRaw = normalized["거래처코드"] || normalized["코드번호"] || normalized["코드"] || normalized.code || "";
  const codeNormalized = normalizeVendorCode(codeRaw);
  const code = codeNormalized || String(codeRaw || "");
  const name = normalized["거래처명"] || normalized["거래처"] || normalized.name || "";
  const year = Number(normalized["작성연도"] || normalized["연도"] || normalized.year || 0);
  const month = Number(normalized["작성월"] || normalized["월"] || normalized.month || 0);
  const purchase = parseSheetNumber(normalized["합계"] || normalized["매입금액"] || normalized.purchase || 0);
  const paid = parseSheetNumber(normalized["지급합"] || normalized["지급액"] || normalized.paid || 0);
  const balance = parseSheetNumber(normalized["잔액"] || normalized.balance || purchase - paid);
  const payDate = normalizeDateValue(normalized["지급일"] || normalized.paydate || normalized.paymentdate || "");
  const memo = normalized["메모"] || normalized.memo || "";
  const dueCategory = normalized["납기"] || normalized.due || normalized["구분"] || "";

  const payable = {
    codeRaw: String(codeRaw || ""),
    codeNormalized,
    code,
    name,
    year,
    month,
    purchase,
    paid,
    balance,
    dueCategory: dueCategory || extractDueCategory(payDate, memo),
    payDate,
    memo,
    selected: false,
    decisionAmount: balance,
    paymentPlan: "",
    paidOverride: paid,
    completionStatus: "",
  };
  payable.sourceKey = buildPayableSourceKey(payable);
  return payable;
}

function extractDueCategory(payDate, memo) {
  const text = String(payDate || memo || "").trim();
  const groups = ["60일", "당말일", "말일", "당05일", "05일", "당10일", "10일", "당15일", "15일", "당25일", "25일", "바로", "즉시"];
  const match = groups.find(group => text.includes(group));
  return match || text || "기타";
}

function getDueGroup(item) {
  if (item.dueCategory && item.dueCategory !== "기타") return item.dueCategory;
  return extractDueCategory(item.payDate, item.memo);
}

function getDueGroupRank(group) {
  const ranks = {
    "60일": 3,
    "당말일": 10, "말일": 11,
    "당05일": 20, "05일": 21,
    "당10일": 30, "10일": 31,
    "당15일": 40, "15일": 41,
    "당25일": 50, "25일": 51,
    "즉시": 80, "바로": 81,
    "기타": 99,
  };
  return ranks[group] || 90;
}

function calcPayableDueDate(year, month, group) {
  if (!year || !month) return "";
  let targetYear = Number(year);
  let targetMonth = Number(month);
  let day = 0;

  if (group === "당말일") {
    // 당월 말일 — 월 변경 없음
  } else if (group === "당05일" || group === "당10일" || group === "당15일" || group === "당25일") {
    targetMonth += 1;  // 익월 N일
  } else if (group === "60일") {
    targetMonth += 2;  // 익익월
  } else if (group === "말일") {
    targetMonth += 1;  // 익월 말일
  } else if (group === "05일" || group === "10일" || group === "15일" || group === "25일") {
    targetMonth += 2;  // 익익월 N일 (말일 마감 다음달)
  } else {
    targetMonth += 1;  // 즉시/바로/기타 → 익월
  }
  while (targetMonth > 12) { targetMonth -= 12; targetYear++; }

  if (group.includes("말일") || group === "60일") {
    day = new Date(targetYear, targetMonth, 0).getDate();
  } else if (group.includes("05일")) day = 5;
  else if (group.includes("10일")) day = 10;
  else if (group.includes("15일")) day = 15;
  else if (group.includes("25일")) day = 25;
  else {
    day = new Date(targetYear, targetMonth, 0).getDate();
  }

  return `${targetYear}-${String(targetMonth).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
}

function parseSheetNumber(value) {
  if (value == null || value === "") return 0;
  return Number(String(value).replace(/[^0-9.-]/g, "")) || 0;
}

async function fetchSheetWebApp(params = {}) {
  const url = new URL(SHEET_APP_SCRIPT_URL);
  const token = getApiToken();
  if (token) url.searchParams.set("token", token);
  // 호출부에서 넘긴 action 등 파라미터를 URL 쿼리에 반영 (예: getMautoTaxInvoices)
  Object.entries(params || {}).forEach(([k, v]) => {
    if (v !== undefined && v !== null) url.searchParams.set(k, String(v));
  });
  const response = await fetch(url.toString());
  if (!response.ok) {
    throw new Error(`Apps Script 요청 실패: ${response.status}`);
  }
  const body = await response.json();
  if (body && body.error === "인증 실패") {
    // 토큰 없으면 강제 입력 대신 공개 시트 폴백
    console.warn("Apps Script 인증 실패 → 공개 시트로 폴백합니다.");
    return fetchPublicSheet();
  }
  // 원본 body를 그대로 반환 → action 응답({rows:[...]})·기본 응답({data:[...]})·배열 모두 호출부에서 처리
  return body;
}

async function fetchAvailableFundsJson() {
  if (!SHEET_APP_SCRIPT_URL) return null;
  const url = new URL(SHEET_APP_SCRIPT_URL);
  url.searchParams.set("action", "getAvailableFundsJson");
  const token = getApiToken();
  if (token) url.searchParams.set("token", token);
  const response = await fetch(url.toString());
  if (!response.ok) throw new Error(`가용자금JSON 조회 실패: ${response.status}`);
  const body = await response.json();
  // { updatedAt, data } 형태
  if (body && "updatedAt" in body) return body;
  return null;
}

async function fetchAvailableFundsFromApi() {
  // 1. 구글 시트 직접 조회 (gviz) - 더 정확함
  try {
    const url = `https://docs.google.com/spreadsheets/d/${SHEET_SPREADSHEET_ID}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent(AVAILABLE_FUNDS_SHEET_NAME)}&headers=1`;
    console.log("[가용자금] gviz 호출 시도:", url);
    const data = await fetchPublicSheetByName(AVAILABLE_FUNDS_SHEET_NAME);
    if (data && data.length > 0) {
      console.log("[가용자금] gviz 로드 성공, 데이터 수:", data.length);
      return data;
    }
  } catch (error) {
    console.warn("가용자금 gviz 조회 실패:", error);
  }

  // 2. Apps Script 폴백 (위 방식 실패 시에만 시도)
  if (SHEET_APP_SCRIPT_URL) {
    try {
      const url = new URL(SHEET_APP_SCRIPT_URL);
      url.searchParams.set("action", "getAvailableFunds");
      const _token = getApiToken();
      if (_token) url.searchParams.set("token", _token);
      console.log("[가용자금] Apps Script 호출 시도:", url.toString());
      const response = await fetch(url.toString());
      if (!response.ok) throw new Error(`가용자금 조회 실패: ${response.status}`);
      const body = await response.json();
      const rows = Array.isArray(body) ? body : (body.rows || body.data || []);
      return rows;
    } catch (error) {
      console.warn("가용자금 Apps Script 조회 실패:", error);
    }
  }

  return [];
}

function parseAvailableFunds(rows) {
  const accounts = [];
  const purchaseVendors = [];
  const eBonds = [];

  let totalAccountBalance = 0;
  let totalPurchaseLoanBalance = 0;
  let totalEBonds = 0;

  console.log("[가용자금] 파싱 시작, 데이터 수:", rows.length);
  if (rows.length > 0) console.log("[가용자금] 첫 행 샘플:", JSON.stringify(rows[0]));

  rows.forEach((row, idx) => {
    // 위치 기반 우선 (A=종류, B=은행/만기, D=금액/잔액) + 기존 헤더 명칭 폴백
    const typeStr = String(row["A"] || row["구분"] || row["종류"] || row["type"] || row[""] || "").trim();
    let type = "";
    if (typeStr.includes("계좌") || typeStr.includes("예금") || typeStr.includes("현금") || typeStr.includes("보통")) {
      type = "계좌";
    } else if (typeStr.includes("구매")) {
      type = "구매자금";
    } else if (typeStr.includes("채권")) {
      type = "전자채권";
    }

    const name = String(row["B"] || row["은행"] || row["만기"] || row["거래처"] || row["client"] || row["금융기관"] || "").trim();
    const value = parseSheetNumber(row["D"] || row["가용자금"] || row["금액"] || row["잔액"] || row["잔액(원)"] || row["amount"] || row["balance"] || 0);

    if (idx < 3) console.log(`[가용자금] 행 ${idx} 파싱:`, { typeStr, type, name, value });

    if (type === "계좌") {
      accounts.push({ bank: name, accountNo: "", balance: value });
      totalAccountBalance += value;
    } else if (type === "구매자금") {
      purchaseVendors.push({ date: "", name, amount: value });
      totalPurchaseLoanBalance += value;
    } else if (type === "전자채권") {
      eBonds.push({ expiry: "", client: name, receiptDate: "", amount: value });
      totalEBonds += value;
    }
  });

  return {
    accounts,
    b2bLoans: [],
    purchaseVendors,
    eBonds,
    eNotes: [],
    summary: {
      totalAccountBalance,
      totalPurchaseLoanBalance,
      totalEBonds,
      availableTotal: totalAccountBalance + totalPurchaseLoanBalance + totalEBonds
    }
  };
}

async function loadAvailableFunds() {
  // 1단계: localStorage 즉시 표시
  const local = loadAvailableFundsLocal();
  if (local) {
    availableFunds = local;
    recalcAvailableFundsSummary();
  }

  // 2단계: 구글시트에서 원격 데이터 로드 (백그라운드)
  try {
    const remote = await fetchAvailableFundsJson();
    if (remote && remote.updatedAt && remote.data) {
      const remoteTs = new Date(remote.updatedAt).getTime();
      const localTs = local && local.updatedAt ? new Date(local.updatedAt).getTime() : 0;
      if (remoteTs > localTs) {
        // 원격이 더 최신 → 로컬 덮어쓰기 (API 재저장은 하지 않음)
        availableFunds = {
          accounts: remote.data.accounts || [],
          b2bLoans: remote.data.b2bLoans || [],
          purchaseVendors: remote.data.purchaseVendors || [],
          eBonds: remote.data.eBonds || [],
          eNotes: remote.data.eNotes || [],
          summary: { totalAccountBalance: 0, b2bUsed: 0, b2bAvailable: 0, totalPurchaseLoanBalance: 0, totalEBonds: 0, totalENotes: 0, grandTotal: 0 },
        };
        recalcAvailableFundsSummary();
        // 로컬에도 저장 (updatedAt 포함)
        try {
          localStorage.setItem(AVAILABLE_FUNDS_LOCAL_KEY, JSON.stringify({
            updatedAt: remote.updatedAt,
            accounts: availableFunds.accounts,
            b2bLoans: availableFunds.b2bLoans,
            purchaseVendors: availableFunds.purchaseVendors,
            eBonds: availableFunds.eBonds,
            eNotes: availableFunds.eNotes,
          }));
        } catch (e) { /* 무시 */ }
        renderAvailableFunds();
        renderDashboard();
        console.log("[가용자금] 원격 데이터 적용 (더 최신):", remote.updatedAt);
      } else {
        console.log("[가용자금] 로컬 데이터가 최신 또는 동일:", local?.updatedAt);
      }
    }
  } catch (err) {
    console.warn("[가용자금] 원격 로드 실패 (로컬 유지):", err);
  }

  if (!local) recalcAvailableFundsSummary();
}

// ── 가용자금 localStorage 저장/로드 ─────────────────────────────

function saveAvailableFundsLocal() {
  const updatedAt = new Date().toISOString();
  try {
    localStorage.setItem(AVAILABLE_FUNDS_LOCAL_KEY, JSON.stringify({
      updatedAt,
      accounts: availableFunds.accounts || [],
      b2bLoans: availableFunds.b2bLoans || [],
      purchaseVendors: availableFunds.purchaseVendors || [],
      eBonds: availableFunds.eBonds || [],
      eNotes: availableFunds.eNotes || [],
    }));
  } catch (e) {
    console.warn("[가용자금] localStorage 저장 실패:", e);
  }
  saveAvailableFundsToApi(updatedAt);
}

async function saveAvailableFundsToApi(updatedAt) {
  try {
    await postSheetWebApp("upsertAvailableFunds", {
      updatedAt,
      data: {
        accounts: availableFunds.accounts || [],
        b2bLoans: availableFunds.b2bLoans || [],
        purchaseVendors: availableFunds.purchaseVendors || [],
        eBonds: availableFunds.eBonds || [],
        eNotes: availableFunds.eNotes || [],
      },
    });
    console.log("[가용자금] 구글시트 저장 완료:", updatedAt);
  } catch (e) {
    console.warn("[가용자금] 구글시트 저장 실패:", e);
  }
}

function loadAvailableFundsLocal() {
  try {
    const raw = localStorage.getItem(AVAILABLE_FUNDS_LOCAL_KEY);
    if (!raw) return null;
    const d = JSON.parse(raw);
    return {
      updatedAt: d.updatedAt || null,
      accounts: d.accounts || [],
      b2bLoans: d.b2bLoans || [],
      purchaseVendors: d.purchaseVendors || [],
      eBonds: d.eBonds || [],
      eNotes: d.eNotes || [],
      summary: { totalAccountBalance: 0, b2bUsed: 0, b2bAvailable: 0, totalPurchaseLoanBalance: 0, totalEBonds: 0, totalENotes: 0, grandTotal: 0 },
    };
  } catch (e) {
    return null;
  }
}

function recalcAvailableFundsSummary() {
  const totalAccountBalance = (availableFunds.accounts || []).reduce((s, r) => s + (r.balance || 0), 0);
  const b2bUsed = (availableFunds.b2bLoans || []).reduce((s, r) => s + (r.used || 0), 0);
  const b2bAvailable = Math.max(0, B2B_TOTAL_LIMIT - b2bUsed);
  const totalPurchase = (availableFunds.purchaseVendors || []).reduce((s, r) => s + (r.amount || 0), 0);
  const totalEBonds = (availableFunds.eBonds || []).reduce((s, r) => s + (r.amount || 0), 0);
  const totalENotes = (availableFunds.eNotes || []).reduce((s, r) => s + (r.amount || 0), 0);
  // 가용자금 합계 = ①계좌 + ②B2B사용가능 + ③전자채권 + ④전자어음
  const grandTotal = totalAccountBalance + b2bAvailable + totalEBonds + totalENotes;
  availableFunds.summary = {
    totalAccountBalance,
    b2bUsed,
    b2bAvailable,
    totalPurchaseLoanBalance: totalPurchase,
    totalEBonds,
    totalENotes,
    grandTotal,
    availableTotal: grandTotal,
  };
}

// ── 엑셀 붙여넣기 파서 ────────────────────────────────────────────

function parseNum(v) {
  if (v === null || v === undefined || v === "") return 0;
  return Number(String(v).replace(/[^0-9.\-]/g, "")) || 0;
}

function parseFundsPaste(text, fieldDefs) {
  // fieldDefs: [{col: '은행', key: 'bank', isNum: false}, ...]
  const lines = text.trim().split(/\r?\n/);
  if (lines.length < 2) return [];
  const headers = lines[0].split("\t").map(h => h.trim());

  // 공백 제거 후 비교로 헤더 인덱스 탐색
  const findColIdx = (col) => {
    let idx = headers.findIndex(h => h === col);
    if (idx >= 0) return idx;
    const norm = col.replace(/\s/g, "");
    return headers.findIndex(h => h.replace(/\s/g, "") === norm);
  };

  return lines.slice(1)
    .filter(l => l.trim())
    .map(l => {
      const cols = l.split("\t");
      const row = {};
      const usedIdxs = new Set();
      fieldDefs.forEach(def => {
        let idx = findColIdx(def.col);
        if (idx >= 0) usedIdxs.add(idx);
        // 숫자 필드이고 헤더 매칭 실패 → 마지막 미사용 숫자 컬럼으로 fallback
        if (idx < 0 && def.isNum) {
          for (let i = cols.length - 1; i >= 0; i--) {
            if (!usedIdxs.has(i) && parseNum(cols[i].trim()) !== 0) {
              idx = i;
              usedIdxs.add(i);
              break;
            }
          }
        }
        const raw = idx >= 0 ? (cols[idx] || "").trim() : "";
        row[def.key] = def.isNum ? parseNum(raw) : raw;
      });
      return row;
    })
    .filter(r => Object.values(r).some(v => v !== "" && v !== 0));
}

// ── 가용자금 탭 렌더 ────────────────────────────────────────────

function fundsTableHtml(headers, rows, totalRow) {
  const thCells = headers.map(h => `<th>${h}</th>`).join("");
  const bodyCells = rows.length
    ? rows.map(cells => `<tr>${cells.map((c, i) => {
        const isNum = typeof c === "number";
        return `<td${isNum ? ' class="funds-num"' : ""}>${isNum ? formatNumber(c) : (c || "")}</td>`;
      }).join("")}</tr>`).join("")
    : `<tr><td colspan="${headers.length}" class="funds-empty">데이터 없음 — 엑셀에서 복사 후 붙여넣기 버튼을 사용하세요</td></tr>`;
  const footCells = totalRow.map((c, i) => {
    const isNum = typeof c === "number";
    return `<td${isNum ? ' class="funds-num"' : ""}>${isNum ? formatNumber(c) : (c || "")}</td>`;
  }).join("");
  return `<table class="funds-table">
    <thead><tr>${thCells}</tr></thead>
    <tbody>${bodyCells}</tbody>
    <tfoot><tr>${footCells}</tr></tfoot>
  </table>`;
}

function fundsAccountTableHtml(accounts, totalBalance) {
  const body = accounts.length
    ? accounts.map((r, i) => `<tr>
        <td>${escapeHtml(r.bank || "")}</td>
        <td>${escapeHtml(r.accountNo || "")}</td>
        <td class="funds-num"><input type="text" class="funds-inline-input" data-acc-idx="${i}" value="${formatNumber(r.balance || 0)}" /></td>
      </tr>`).join("")
    : `<tr><td colspan="3" class="funds-empty">데이터 없음 — 엑셀에서 복사 후 붙여넣기 버튼을 사용하세요</td></tr>`;
  return `<table class="funds-table" id="funds-acc-table">
    <thead><tr><th>은행</th><th>계좌번호</th><th>가용자금</th></tr></thead>
    <tbody>${body}</tbody>
    <tfoot><tr><td>합계</td><td></td><td class="funds-num" id="funds-acc-total">${formatNumber(totalBalance)}</td></tr></tfoot>
  </table>`;
}

function fundsSection(id, title, tableHtml, hint) {
  return `<div class="funds-section" id="fs-${id}">
    <div class="funds-sec-header">
      <span class="funds-sec-title">${title}</span>
      <button type="button" class="funds-paste-btn" data-fs="${id}">붙여넣기 입력</button>
    </div>
    <div class="funds-paste-area hidden" id="fpa-${id}">
      <div class="funds-paste-hint">${hint}</div>
      <textarea class="funds-textarea" id="fpt-${id}" placeholder="엑셀에서 복사(Ctrl+C) 후 여기에 붙여넣기(Ctrl+V)"></textarea>
      <div class="funds-paste-actions">
        <button type="button" class="funds-apply-btn" data-fs="${id}">✔ 적용</button>
        <button type="button" class="funds-cancel-btn" data-fs="${id}">취소</button>
      </div>
    </div>
    ${tableHtml}
  </div>`;
}

function renderAvailableFunds() {
  const sec = document.getElementById("funds");
  if (!sec) return;

  recalcAvailableFundsSummary();
  const af = availableFunds;
  const s = af.summary;

  const b2bUsed = s.b2bUsed || 0;
  const b2bAvail = s.b2bAvailable || 0;
  const ebTotal = s.totalEBonds || 0;
  const enTotal = s.totalENotes || 0;
  const grandTotal = s.grandTotal || 0;

  // ① 계좌
  const accTable = fundsAccountTableHtml(af.accounts || [], s.totalAccountBalance);

  // ② B2B 대출
  const b2bTable = fundsTableHtml(
    ["최신만기일", "실행번호", "최종만기", "합계"],
    (af.b2bLoans || []).map(r => [r.latestExpiry, r.execNo, r.finalExpiry, r.used]),
    ["", "", "현사용액", b2bUsed]
  );
  const b2bInfoHtml = `<div class="b2b-summary">
    <span>총한도 <strong>${formatNumber(B2B_TOTAL_LIMIT)}</strong></span>
    <span class="sep">−</span>
    <span>현사용액 <strong class="red">${formatNumber(b2bUsed)}</strong></span>
    <span class="sep">=</span>
    <span>사용가능 <strong class="blue">${formatNumber(b2bAvail)}</strong></span>
  </div>`;

  // ② 구매자금 사용가능 업체 (B2B 참고)
  const pvTotal = s.totalPurchaseLoanBalance || 0;
  const pvTable = fundsTableHtml(
    ["작성일자", "업체명", "금액"],
    (af.purchaseVendors || []).map(r => [r.date, r.name, r.amount]),
    ["합계", "", pvTotal]
  );
  const pvSection = `<div class="funds-ref-section">
    <div class="funds-sec-header funds-ref-header">
      <span class="funds-sec-title" style="color:#64748b;font-weight:600;font-size:12px;">구매자금 사용가능 업체 (참고)</span>
      <button type="button" class="funds-paste-btn" data-fs="pv">붙여넣기 입력</button>
    </div>
    <div class="funds-paste-area hidden" id="fpa-pv">
      <div class="funds-paste-hint">헤더: 작성일자 / 업체명 / 금액</div>
      <textarea class="funds-textarea" id="fpt-pv" placeholder="엑셀에서 복사(Ctrl+C) 후 여기에 붙여넣기(Ctrl+V)"></textarea>
      <div class="funds-paste-actions">
        <button type="button" class="funds-apply-btn" data-fs="pv">✔ 적용</button>
        <button type="button" class="funds-cancel-btn" data-fs="pv">취소</button>
      </div>
    </div>
    ${pvTable}
  </div>`;

  // ③ 전자채권
  const ebTable = fundsTableHtml(
    ["만기일", "거래처명", "수납일", "합계"],
    (af.eBonds || []).map(r => [r.expiry, r.client, r.receiptDate, r.amount]),
    ["합계", "", "", ebTotal]
  );

  // ④ 전자어음
  const enTable = fundsTableHtml(
    ["은행", "거래처명", "수납일", "만기일", "합계"],
    (af.eNotes || []).map(r => [r.bank, r.client, r.receiptDate, r.expiry, r.amount]),
    ["합계", "", "", "", enTotal]
  );

  sec.innerHTML = `<div class="funds-container">
    <div class="funds-top-bar">
      <h2 class="funds-title">가용자금 현황</h2>
      <button type="button" id="fundsClearBtn" class="funds-clear-btn">전체 초기화</button>
    </div>

    <!-- 요약 카드 -->
    <div class="funds-summary-grid">
      <div class="fsc fsc-account">
        <div class="fsc-badge">①</div>
        <div class="fsc-label">계좌</div>
        <div class="fsc-amount">${formatNumber(s.totalAccountBalance)}</div>
      </div>
      <div class="fsc fsc-b2b">
        <div class="fsc-badge">②</div>
        <div class="fsc-label">B2B 사용가능</div>
        <div class="fsc-amount">${formatNumber(b2bAvail)}</div>
        <div class="fsc-sub">한도 ${formatNumber(B2B_TOTAL_LIMIT)}</div>
        <div class="fsc-sub">사용 ${formatNumber(b2bUsed)}</div>
      </div>
      <div class="fsc fsc-bond">
        <div class="fsc-badge">③</div>
        <div class="fsc-label">전자채권</div>
        <div class="fsc-amount">${formatNumber(ebTotal)}</div>
      </div>
      <div class="fsc fsc-note">
        <div class="fsc-badge">④</div>
        <div class="fsc-label">전자어음</div>
        <div class="fsc-amount">${formatNumber(enTotal)}</div>
      </div>
      <div class="fsc fsc-total">
        <div class="fsc-label">합 계</div>
        <div class="fsc-amount">${formatNumber(grandTotal)}</div>
        <div class="fsc-sub">①+②+③+④</div>
      </div>
    </div>

    ${fundsSection("accounts", "① 계좌",
        accTable,
        "헤더: 은행 / 계좌번호 / 가용자금")}
    ${fundsSection("b2b", "② B2B 대출",
        b2bInfoHtml + b2bTable + pvSection,
        "헤더: 최신만기일 / 실행번호 / 최종만기 / 합계")}
    ${fundsSection("eb", "③ 전자채권",
        ebTable,
        "헤더: 만기일 / 거래처명 / 수납일 / 합계")}
    ${fundsSection("en", "④ 전자어음",
        enTable,
        "헤더: 은행 / 거래처명 / 수납일 / 만기일 / 합계")}
  </div>`;

  // 이벤트 바인딩
  sec.querySelectorAll(".funds-paste-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      const id = btn.dataset.fs;
      const area = document.getElementById(`fpa-${id}`);
      if (area) {
        area.classList.toggle("hidden");
        if (!area.classList.contains("hidden")) {
          document.getElementById(`fpt-${id}`)?.focus();
        }
      }
    });
  });

  sec.querySelectorAll(".funds-cancel-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      const area = document.getElementById(`fpa-${btn.dataset.fs}`);
      if (area) area.classList.add("hidden");
    });
  });

  sec.querySelectorAll(".funds-apply-btn").forEach(btn => {
    btn.addEventListener("click", () => applyFundsPaste(btn.dataset.fs));
  });

  // textarea에서 Ctrl+V 후 자동 적용 (붙여넣기 감지)
  sec.querySelectorAll(".funds-textarea").forEach(ta => {
    ta.addEventListener("paste", (e) => {
      const id = ta.id.replace("fpt-", "");
      setTimeout(() => applyFundsPaste(id), 50);
    });
  });

  // 계좌 금액 직접 수정
  sec.querySelectorAll(".funds-inline-input[data-acc-idx]").forEach(input => {
    const idx = parseInt(input.dataset.accIdx, 10);
    input.addEventListener("focus", () => {
      input.value = String(availableFunds.accounts[idx]?.balance || 0);
      input.select();
    });
    input.addEventListener("blur", () => {
      const val = parseNum(input.value);
      if (availableFunds.accounts[idx]) availableFunds.accounts[idx].balance = val;
      input.value = formatNumber(val);
      recalcAvailableFundsSummary();
      const s2 = availableFunds.summary;
      const totalEl = document.getElementById("funds-acc-total");
      if (totalEl) totalEl.textContent = formatNumber(s2.totalAccountBalance);
      const cardEl = sec.querySelector(".fsc-account .fsc-amount");
      if (cardEl) cardEl.textContent = formatNumber(s2.totalAccountBalance);
      const grandEl = sec.querySelector(".fsc-total .fsc-amount");
      if (grandEl) grandEl.textContent = formatNumber(s2.grandTotal);
      saveAvailableFundsLocal();
      renderDashboard();
    });
    input.addEventListener("keydown", e => {
      if (e.key === "Enter") { e.preventDefault(); input.blur(); }
      if (e.key === "Escape") {
        input.value = formatNumber(availableFunds.accounts[idx]?.balance || 0);
        input.blur();
      }
    });
  });

  document.getElementById("fundsClearBtn")?.addEventListener("click", () => {
    if (!confirm("가용자금 데이터를 전체 초기화하시겠습니까?")) return;
    availableFunds.accounts = [];
    availableFunds.b2bLoans = [];
    availableFunds.purchaseVendors = [];
    availableFunds.eBonds = [];
    availableFunds.eNotes = [];
    saveAvailableFundsLocal();
    renderAvailableFunds();
    renderDashboard();
  });
}

function applyFundsPaste(sectionId) {
  const ta = document.getElementById(`fpt-${sectionId}`);
  if (!ta) return;
  const text = ta.value.trim();
  if (!text) return;

  const DEFS = {
    accounts: [
      { col: "은행", key: "bank", isNum: false },
      { col: "계좌번호", key: "accountNo", isNum: false },
      { col: "가용자금", key: "balance", isNum: true },
    ],
    b2b: [
      { col: "최신만기일", key: "latestExpiry", isNum: false },
      { col: "실행번호", key: "execNo", isNum: false },
      { col: "최종만기", key: "finalExpiry", isNum: false },
      { col: "합계", key: "used", isNum: true },
    ],
    pv: [
      { col: "작성일자", key: "date", isNum: false },
      { col: "업체명", key: "name", isNum: false },
      { col: "금액", key: "amount", isNum: true },
    ],
    eb: [
      { col: "만기일", key: "expiry", isNum: false },
      { col: "거래처명", key: "client", isNum: false },
      { col: "수납일", key: "receiptDate", isNum: false },
      { col: "합계", key: "amount", isNum: true },
    ],
    en: [
      { col: "은행", key: "bank", isNum: false },
      { col: "거래처명", key: "client", isNum: false },
      { col: "수납일", key: "receiptDate", isNum: false },
      { col: "만기일", key: "expiry", isNum: false },
      { col: "합계", key: "amount", isNum: true },
    ],
  };

  const defs = DEFS[sectionId];
  if (!defs) return;

  const parsed = parseFundsPaste(text, defs);
  if (!parsed.length) {
    alert("데이터를 읽지 못했습니다.\n엑셀 헤더가 정확한지 확인하세요.");
    return;
  }

  const MAP = {
    accounts: "accounts",
    b2b: "b2bLoans",
    pv: "purchaseVendors",
    eb: "eBonds",
    en: "eNotes",
  };
  availableFunds[MAP[sectionId]] = parsed;
  saveAvailableFundsLocal();
  renderAvailableFunds();
  renderDashboard();
}

function buildCashflowTimeline() {
  const map = {};
  const ensure = d => { if (!map[d]) map[d] = { rcv: 0, rcvOverdue: 0, pay: 0, payHeld: 0, fixed: 0 }; };
  const today = new Date().toISOString().slice(0, 10);

  // 미수금: 납기 경과분은 합계 제외, rcvOverdue로 참고 표시
  let overdueRcv = 0;
  receivables.forEach(r => {
    if (!r.dueDate || !(r.balance > 0)) return;
    if (r.dueDate < today) {
      overdueRcv += r.balance;
    } else {
      ensure(r.dueDate);
      map[r.dueDate].rcv += r.balance;
    }
  });
  if (overdueRcv > 0) {
    const nearestFuture = Object.keys(map).filter(d => d >= today && map[d].rcv > 0).sort()[0] || today;
    ensure(nearestFuture);
    map[nearestFuture].rcvOverdue += overdueRcv;
  }

  // 미지급: 보류/제외는 합계 제외, payHeld로 참고 표시
  let heldPay = 0;
  payables.forEach(p => {
    const outstanding = getPayableOutstanding(p);
    if (outstanding <= 0) return;
    const plan = p.paymentPlan || "";
    const isHeld = plan === "보류" || plan === "제외" || p.completionStatus === "보류";
    const dueDate = /^\d{4}-\d{2}-\d{2}$/.test(plan)
      ? plan
      : calcPayableDueDate(p.year, p.month, getDueGroup(p));
    if (!dueDate) return;
    if (isHeld) {
      heldPay += outstanding;
    } else {
      ensure(dueDate);
      map[dueDate].pay += outstanding;
    }
  });
  if (heldPay > 0) {
    const nearestFuturePay = Object.keys(map).filter(d => d >= today && map[d].pay > 0).sort()[0] || today;
    ensure(nearestFuturePay);
    map[nearestFuturePay].payHeld += heldPay;
  }

  fixedExpenses.forEach(f => {
    if (!f.year || !f.month || !f.day || !(f.amount > 0)) return;
    const d = `${f.year}-${String(f.month).padStart(2,"0")}-${String(f.day).padStart(2,"0")}`;
    ensure(d);
    map[d].fixed += f.amount;
  });

  return map;
}

function buildCashflowTimelineHtml(mode) {
  const dateMap = buildCashflowTimeline();
  const allDates = Object.keys(dateMap).sort();
  if (!allDates.length) {
    return `<div style="margin-top:20px;padding:20px;background:#f9fafb;border-radius:10px;border:1px solid #e5e7eb;text-align:center;color:#9ca3af;font-size:13px;">현금흐름 데이터가 없습니다.</div>`;
  }

  let rows;
  if (mode === "daily") {
    rows = allDates.map(d => {
      const { rcv, rcvOverdue, pay, payHeld, fixed } = dateMap[d];
      return { label: d.slice(5).replace("-", "/"), dateFrom: d, rcv, rcvOverdue: rcvOverdue || 0, pay, payHeld: payHeld || 0, fixed, net: rcv - pay - fixed };
    });
  } else {
    const weekMap = {};
    allDates.forEach(dateStr => {
      const d = new Date(dateStr + "T00:00:00");
      const dow = d.getDay();
      const mon = new Date(d);
      mon.setDate(d.getDate() - (dow === 0 ? 6 : dow - 1));
      const wk = mon.toISOString().slice(0, 10);
      if (!weekMap[wk]) {
        const sun = new Date(mon); sun.setDate(mon.getDate() + 6);
        weekMap[wk] = { rcv: 0, rcvOverdue: 0, pay: 0, payHeld: 0, fixed: 0, end: sun.toISOString().slice(0, 10) };
      }
      weekMap[wk].rcv += dateMap[dateStr].rcv;
      weekMap[wk].rcvOverdue += (dateMap[dateStr].rcvOverdue || 0);
      weekMap[wk].pay += dateMap[dateStr].pay;
      weekMap[wk].payHeld += (dateMap[dateStr].payHeld || 0);
      weekMap[wk].fixed += dateMap[dateStr].fixed;
    });
    rows = Object.keys(weekMap).sort().map(wk => {
      const { rcv, rcvOverdue, pay, payHeld, fixed, end } = weekMap[wk];
      return { label: `${wk.slice(5).replace("-","/")}~${end.slice(5).replace("-","/")}`, dateFrom: wk, rcv, rcvOverdue: rcvOverdue || 0, pay, payHeld: payHeld || 0, fixed, net: rcv - pay - fixed };
    });
  }

  const today = new Date().toISOString().slice(0, 10);
  const TDsty = "padding:8px 14px;border-bottom:1px solid #f0f0f0;";
  const rowHtml = rows.map((r, i) => {
    const isPast = r.dateFrom < today;
    const hasExtra = r.rcvOverdue > 0 || r.payHeld > 0;
    const rowBg = isPast ? "background:#fafafa;" : (i % 2 === 1 ? "background:#f8fbff;" : "");
    const labelStyle = isPast ? "color:#9ca3af;" : "font-weight:600;color:#1f2937;";
    const netStyle = r.net > 0 ? "color:#16a34a;font-weight:bold;" : r.net < 0 ? "color:#dc2626;font-weight:bold;" : "color:#9ca3af;";
    const netStr = r.net === 0 ? "-" : (r.net > 0 ? "+" : "") + formatNumber(r.net);
    const rcvCell = r.rcv > 0 ? `<span style="color:#1565c0;font-weight:500;">${formatNumber(r.rcv)}</span>` : `<span style="color:#d1d5db;">-</span>`;
    const payCell = r.pay > 0 ? `<span style="color:#b71c1c;font-weight:500;">${formatNumber(r.pay)}</span>` : `<span style="color:#d1d5db;">-</span>`;
    const fixedCell = r.fixed > 0 ? `<span style="color:#9a3412;font-weight:500;">${formatNumber(r.fixed)}</span>` : `<span style="color:#d1d5db;">-</span>`;
    const extraIcon = hasExtra ? `<span style="margin-left:5px;font-size:10px;color:#9ca3af;">▾</span>` : "";

    const detailParts = [];
    if (r.rcv > 0) detailParts.push(`<span style="color:#1565c0;">수금 예정 <b>${formatNumber(r.rcv)}</b>원</span>`);
    if (r.rcvOverdue > 0) detailParts.push(`<span style="color:#b45309;">🕐 이월 연체(참고) <b>${formatNumber(r.rcvOverdue)}</b>원</span>`);
    if (r.pay > 0) detailParts.push(`<span style="color:#b71c1c;">지급 예정 <b>${formatNumber(r.pay)}</b>원</span>`);
    if (r.payHeld > 0) detailParts.push(`<span style="color:#6b7280;">⏸ 보류·제외(참고) <b>${formatNumber(r.payHeld)}</b>원</span>`);
    if (r.fixed > 0) detailParts.push(`<span style="color:#9a3412;">고정지출 <b>${formatNumber(r.fixed)}</b>원</span>`);
    const detailRow = hasExtra ? `<tr id="tl-detail-${i}" style="display:none;">
      <td colspan="5" style="padding:5px 22px 8px;border-bottom:1px solid #f0f0f0;background:#fffbf0;">
        <div style="display:flex;flex-wrap:wrap;gap:14px;font-size:12px;">${detailParts.join("")}</div>
      </td>
    </tr>` : "";

    return `<tr data-tl-idx="${i}" style="${rowBg}${isPast ? "opacity:0.65;" : ""}${hasExtra ? "cursor:pointer;" : ""}">
      <td style="${TDsty}${labelStyle}white-space:nowrap;">${r.label}${extraIcon}</td>
      <td style="${TDsty}text-align:right;">${rcvCell}</td>
      <td style="${TDsty}text-align:right;">${payCell}</td>
      <td style="${TDsty}text-align:right;">${fixedCell}</td>
      <td style="${TDsty}text-align:right;${netStyle}">${netStr}</td>
    </tr>${detailRow}`;
  }).join("");

  const dayActive = mode === "daily" ? "background:#2563eb;color:white;" : "background:#f3f4f6;color:#374151;";
  const wkActive  = mode === "weekly" ? "background:#2563eb;color:white;" : "background:#f3f4f6;color:#374151;";

  return `
    <div style="margin-top:20px;background:#fff;border-radius:10px;border:1px solid #e5e7eb;overflow:hidden;">
      <div style="display:flex;align-items:center;justify-content:space-between;padding:12px 18px;border-bottom:1px solid #e5e7eb;background:#f8fafc;">
        <h3 style="margin:0;font-size:14px;font-weight:700;color:#1f2937;">현금흐름 일정</h3>
        <div style="display:flex;gap:3px;">
          <button class="tl-mode-btn" data-mode="daily" style="padding:4px 13px;border-radius:4px;border:none;cursor:pointer;font-size:12px;${dayActive}">일별</button>
          <button class="tl-mode-btn" data-mode="weekly" style="padding:4px 13px;border-radius:4px;border:none;cursor:pointer;font-size:12px;${wkActive}">주별</button>
        </div>
      </div>
      <div style="max-height:420px;overflow:auto;-webkit-overflow-scrolling:touch;">
        <table class="tl-table" style="width:100%;min-width:440px;border-collapse:collapse;font-size:13px;">
          <thead>
            <tr style="background:#f8fafc;position:sticky;top:0;z-index:1;">
              <th style="padding:9px 14px;border-bottom:2px solid #e5e7eb;text-align:left;font-size:12px;color:#6b7280;font-weight:600;">날짜</th>
              <th style="padding:9px 14px;border-bottom:2px solid #e5e7eb;text-align:right;font-size:12px;color:#1565c0;font-weight:600;">미수금 (+)</th>
              <th style="padding:9px 14px;border-bottom:2px solid #e5e7eb;text-align:right;font-size:12px;color:#b71c1c;font-weight:600;">미지급 (-)</th>
              <th style="padding:9px 14px;border-bottom:2px solid #e5e7eb;text-align:right;font-size:12px;color:#9a3412;font-weight:600;">고정지출 (-)</th>
              <th style="padding:9px 14px;border-bottom:2px solid #e5e7eb;text-align:right;font-size:12px;color:#374151;font-weight:600;">증감</th>
            </tr>
          </thead>
          <tbody>${rowHtml}</tbody>
        </table>
      </div>
    </div>`;
}

function renderDashboard() {
  recalcAvailableFundsSummary();
  const summary = calculateSummary();
  const homeSection = document.getElementById("home");
  if (!homeSection) return;

  const s = availableFunds.summary;
  const totalExpected = ((s.grandTotal || s.totalAccountBalance) + summary.totalOutstanding) - (summary.totalUnpaid + summary.totalFixed);

  const b2bUsed = s.b2bUsed || 0;
  const b2bAvail = Math.max(0, B2B_TOTAL_LIMIT - b2bUsed);

  homeSection.innerHTML = `
    <div class="dashboard-container">
      <div class="dashboard-header">
        <h2>금융 통합 대시보드</h2>
        <span class="last-updated">최근 업데이트: ${new Date().toLocaleString()}</span>
      </div>

      <div class="dashboard-summary-cards">
        <div class="dashboard-card funds" data-tab="funds">
          <div class="card-icon"></div>
          <div class="card-label">가용자금 합계</div>
          <div class="card-value">${formatNumber(s.grandTotal || s.totalAccountBalance)}</div>
          <div class="card-footer">B2B 가능 ${formatNumber(b2bAvail)}</div>
        </div>

        <div class="dashboard-card receivables" data-tab="receivables">
          <div class="card-icon"></div>
          <div class="card-label">수금예상(미수금)</div>
          <div class="card-value">${formatNumber(summary.totalOutstanding)}</div>
          <div class="card-footer">${receivables.length}개 업체</div>
        </div>

        <div class="dashboard-card payables" data-tab="payables">
          <div class="card-icon"></div>
          <div class="card-label">외상대지급(미지급)</div>
          <div class="card-value">${formatNumber(summary.totalUnpaid)}</div>
          <div class="card-footer">${payables.length}건</div>
        </div>

        <div class="dashboard-card fixed" data-tab="fixed">
          <div class="card-icon"></div>
          <div class="card-label">고정지출</div>
          <div class="card-value">${formatNumber(summary.totalFixed)}</div>
          <div class="card-footer">이번 달 납부 예정</div>
        </div>

        <div class="dashboard-card expected highlight">
          <div class="card-icon"></div>
          <div class="card-label">예상 잔액</div>
          <div class="card-value">${formatNumber(totalExpected)}</div>
          <div class="card-footer">최종 가용 예상</div>
        </div>
      </div>

      ${buildCashflowTimelineHtml(cashflowTimelineMode)}
    </div>
  `;

  homeSection.querySelectorAll(".dashboard-card[data-tab]").forEach(card => {
    card.addEventListener("click", () => switchTab(card.dataset.tab));
  });
  homeSection.querySelectorAll(".tl-mode-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      cashflowTimelineMode = btn.dataset.mode;
      renderDashboard();
    });
  });
  const tlTable = homeSection.querySelector(".tl-table");
  if (tlTable) {
    tlTable.addEventListener("click", e => {
      const tr = e.target.closest("tr[data-tl-idx]");
      if (!tr) return;
      const detail = document.getElementById(`tl-detail-${tr.dataset.tlIdx}`);
      if (!detail) return;
      detail.style.display = detail.style.display === "none" ? "" : "none";
    });
  }
}

// 검색/연도/월/상태 필터는 미수금·미지급·고정지출·대사에서만 사용 → 그 외 탭에선 숨겨 공간 절약
const FILTER_BAR_TABS = new Set(["receivables", "payables", "fixed", "daesa"]);
function updateFilterBarVisibility(tabId) {
  const fr = document.querySelector(".filter-row");
  if (fr) fr.style.display = FILTER_BAR_TABS.has(tabId) ? "" : "none";
}

function switchTab(tabId) {
  const buttons = document.querySelectorAll(".tab-button");
  const contents = document.querySelectorAll(".tab-content");

  buttons.forEach(btn => {
    btn.classList.toggle("active", btn.dataset.tab === tabId);
  });

  contents.forEach(content => {
    content.classList.toggle("active", content.id === tabId);
  });

  updateFilterBarVisibility(tabId);

  if (tabId === "mauto") {
    renderMautoTab();
    // 고정지출 분류규칙 자동 로드 (탭 첫 진입 시)
    if (mautoFixedRules === null && SHEET_APP_SCRIPT_URL) {
      fetchRulesFromApi("엠오토").then(rules => {
        mautoFixedRules = rules;
        renderMautoTab();
      }).catch(() => { mautoFixedRules = []; });
    }
    // 세금계산서 원격 로드 (로컬에 없는 행 보충 — 컴퓨터 간 공유)
    if (SHEET_APP_SCRIPT_URL) loadMautoTaxRemote();
    // 입출금 원본 행 원격 로드 (로컬에 없는 행 보충 — 컴퓨터 간 공유)
    if (SHEET_APP_SCRIPT_URL) loadMautoSourceRemote();

    // 입출금 분류 원격 로드 → 로컬과 병합 (다른 사람이 저장한 데이터 반영)
    if (SHEET_APP_SCRIPT_URL) {
      fetchSheetWebApp({ action: "getClassifiedRows" }).then(res => {
        const remote = (res && (res.rows || res.data)) || [];
        if (!remote.length) return;
        const localMap = new Map(mautoClassifiedRows.map(r => [r._txKey, r]));
        let added = 0;
        remote.forEach(r => {
          if (!r._txKey) return;
          if (!localMap.has(r._txKey)) { localMap.set(r._txKey, r); added++; }
          else {
            // 원격이 더 최신이면 거래처명/구분만 업데이트 (사용자 수동 수정 보존)
            const local = localMap.get(r._txKey);
            if (!local.savedAt || (r.savedAt && r.savedAt > local.savedAt)) {
              localMap.set(r._txKey, { ...local, 거래처명: r.거래처명, 구분: r.구분, excluded: r.excluded, 매칭근거: r.매칭근거, savedAt: r.savedAt });
            }
          }
        });
        if (added > 0 || remote.length) {
          mautoClassifiedRows = [...localMap.values()].sort((a,b) => (a.date||"") < (b.date||"") ? -1 : 1);
          try { localStorage.setItem(MAUTO_CLASSIFIED_KEY, JSON.stringify(mautoClassifiedRows)); } catch(_) {}
          renderMautoTab();
        }
      }).catch(() => {});
    }
    console.log("[엠오토] 탭 진입 — 원격 로드 시작");
    if (SHEET_APP_SCRIPT_URL) {
      loadMautoDataRemote().then(remote => {
        console.log("[엠오토] 원격 응답:", remote === null ? "null" : JSON.stringify(remote).slice(0, 300));
        if (!remote) {
          // 구글시트에 데이터 없음 → 로컬에 실제 데이터가 있을 때만 업로드 (빈 값으로 덮어쓰기 방지)
          const hasLocalData = mautoData.funds.some(f => (f.amount || 0) > 0) ||
            mautoData.receivables.length > 0 || mautoData.payables.length > 0 || mautoData.fixed.length > 0;
          if (hasLocalData) _scheduleMautoRemoteSave();
          return;
        }
        if (Array.isArray(remote.excludeRcv)) { mautoExcludeVendorsRcv = remote.excludeRcv; try { localStorage.setItem(MAUTO_EXCLUDE_KEY_RCV, JSON.stringify(mautoExcludeVendorsRcv)); } catch (_) {} }
        if (Array.isArray(remote.excludePay)) { mautoExcludeVendorsPay = remote.excludePay; try { localStorage.setItem(MAUTO_EXCLUDE_KEY_PAY, JSON.stringify(mautoExcludeVendorsPay)); } catch (_) {} }
        if (remote.fixedChecked && typeof remote.fixedChecked === "object") { mautoFixedChecked = remote.fixedChecked; try { localStorage.setItem(MAUTO_FIXED_CHECKED_KEY, JSON.stringify(mautoFixedChecked)); } catch (_) {} }
        if (remote.fixedAmountOverrides && typeof remote.fixedAmountOverrides === "object") { mautoFixedAmountOverrides = remote.fixedAmountOverrides; try { localStorage.setItem(MAUTO_FIXED_AMOUNT_KEY, JSON.stringify(mautoFixedAmountOverrides)); } catch (_) {} }
        if (Array.isArray(remote.taxInvoices) && remote.taxInvoices.length) {
          const localKeys = new Set(mautoTaxInvoices.map(r => r._row_key).filter(Boolean));
          const newRows = remote.taxInvoices.filter(r => r._row_key && !localKeys.has(r._row_key));
          if (newRows.length) {
            mautoTaxInvoices = [...mautoTaxInvoices, ...newRows];
            try { localStorage.setItem(MAUTO_TAX_SOURCE_KEY, JSON.stringify(mautoTaxSources)); } catch (_) {}
          }
        }
        mautoData = normalizeMautoData(remote);
        console.log("[엠오토] 정규화 후 funds:", JSON.stringify(mautoData.funds));
        try { localStorage.setItem(MAUTO_LOCAL_KEY, JSON.stringify(mautoData)); } catch (_) {}
        renderMautoTab();
      }).catch(e => console.warn("[엠오토] 원격 로드 실패:", e));
    } else {
      console.warn("[엠오토] SHEET_APP_SCRIPT_URL 없음 — 원격 로드 건너뜀");
    }
  }

  // 대사 탭: 로컬 소스 파일 없고 GSheets 데이터도 안 불러왔으면 자동 로드
  if (tabId === "daesa" && !daesaState.loaded && !hasMiraeSources() && SHEET_APP_SCRIPT_URL) {
    loadDaesaData();
  }

  if (tabId === "pnl") {
    renderPnlTab();
    loadPnlRemote();
  }

  // 탭 이동 시 스크롤 상단으로
  window.scrollTo(0, 0);
}

async function fetchPublicSheetByName(sheetName) {
  const url = `https://docs.google.com/spreadsheets/d/${SHEET_SPREADSHEET_ID}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent(sheetName)}&headers=1`;
  const response = await fetch(url);
  if (!response.ok) throw new Error(`Google Sheets 공개 요청 실패: ${response.status}`);
  const text = await response.text();
  const json = JSON.parse(text.replace(/^.*?\{/, "{").replace(/;$/, ""));
  const cols = json.table.cols.map(col => col.label || "");
  const colTypes = json.table.cols.map(col => col.type || "");
  return json.table.rows.map(row => {
    const item = {};
    row.c.forEach((cell, index) => {
      const colLabel = cols[index] || "";
      const colLetter = String.fromCharCode(65 + index); // A, B, C...
      
      let val = "";
      if (cell) {
        if (colTypes[index] === "date" || colTypes[index] === "datetime") {
          val = cell.f ?? cell.v ?? "";
        } else {
          val = cell.v ?? "";
        }
      }
      
      if (colLabel) item[colLabel] = val;
      item[colLetter] = val; // 무조건 A, B, C... 키도 함께 저장 (더 강력함)
    });
    return item;
  });
}

async function fetchPublicSheet() {
  return fetchPublicSheetByName(SHEET_NAME_PAYABLES);
}

function rerenderAll() {
  renderDashboard();
  renderSummary();
  renderReceivables();
  renderPayables();
  renderFixedExpenses();
  renderAvailableFunds();
}

function calculateSummary() {
  const receivableItems = getFilteredItems(receivables, "receivables");
  const payableItems = getFilteredItems(payables, "payables");
  const fixedItems = getFilteredItems(fixedExpenses, "fixed");

  const totalReceivable = receivableItems.reduce((sum, item) => sum + Number(item.sales || item.balance || 0), 0);
  const totalReceived = receivableItems.reduce((sum, item) => sum + Number(item.collection || item.paid || 0), 0);
  const totalPayable = payableItems.reduce((sum, item) => sum + item.purchase, 0);
  const totalPaid = payableItems.reduce((sum, item) => sum + getPayableEffectivePaid(item), 0);
  const totalFixed = fixedItems.reduce((sum, item) => sum + item.amount, 0);

  return {
    totalReceivable,
    totalOutstanding: totalReceivable - totalReceived,
    totalPayable,
    totalUnpaid: totalPayable - totalPaid,
    totalFixed,
  };
}

function renderSummary() {
  if (elements.summaryPanel) elements.summaryPanel.innerHTML = "";
}

function renderPartnerFilter() {
  const partnerMap = new Map();
  const normalizeCode = code => String(code ?? "").trim();

  partners.forEach(partner => {
    const code = normalizeCode(partner.code);
    if (code) partnerMap.set(code, partner.name);
  });

  payables.forEach(item => {
    const code = normalizeCode(item.code);
    if (code) {
      partnerMap.set(code, item.name || partnerMap.get(code) || code);
    }
  });

  receivables.forEach(item => {
    const code = normalizeCode(item.code);
    if (code) {
      partnerMap.set(code, item.name || partnerMap.get(code) || code);
    }
  });

  elements.partnerFilter.innerHTML = `<option value="">전체 거래처</option>` +
    [...partnerMap.entries()]
      .filter(([code]) => code !== "")
      .sort(([a], [b]) => String(a).localeCompare(String(b)))
      .map(([code, name]) => `
      <option value="${code}">${code} · ${name}</option>
    `).join("");
}

function renderGroupFilterControls() {
  const allGroups = [...new Set(payables.map(getDueGroup).filter(Boolean))]
    .sort((a, b) => {
      const rankDiff = getDueGroupRank(a) - getDueGroupRank(b);
      if (rankDiff !== 0) return rankDiff;
      return String(a).localeCompare(String(b), "ko");
    });

  // 저장된 드래그 순서 적용 (없는 항목은 뒤에 추가)
  const savedOrder = filterState.groupOrder || [];
  const orderedGroups = [
    ...savedOrder.filter(g => allGroups.includes(g)),
    ...allGroups.filter(g => !savedOrder.includes(g)),
  ];

  const isChecked = (group) => filterState.groups === null || (filterState.groups && filterState.groups.includes(group));

  elements.groupFilterContainer.innerHTML = `
    <div class="group-filter-toolbar">
      <button type="button" class="group-manage-link" data-action="select-all">전체 선택</button>
      <button type="button" class="group-manage-link" data-action="clear-all">전체 해제</button>
      <span class="group-filter-guide">드래그로 순서 변경</span>
    </div>
    <div class="group-filter-list compact">
      ${orderedGroups.map(group => {
    const checked = isChecked(group);
    return `
        <label class="group-filter-item ${checked ? "selected" : ""} group-filter-item-draggable" draggable="true" data-group="${group}">
          <span class="group-chip-handle">≡</span>
          <input type="checkbox" value="${group}" ${checked ? "checked" : ""} />
          <span>${group}</span>
        </label>
      `;
  }).join("")}
    </div>
  `;

  elements.groupFilterContainer.querySelectorAll("input[type=checkbox]").forEach(input => {
    input.addEventListener("change", () => {
      const value = input.value;
      const cur = filterState.groups === null ? [...orderedGroups] : [...filterState.groups];
      if (input.checked) {
        if (!cur.includes(value)) cur.push(value);
        filterState.groups = cur.length === allGroups.length ? null : cur;
      } else {
        filterState.groups = cur.filter(group => group !== value);
      }
      // 칩 UI만 제자리에서 토글 (순서 유지)
      const label = input.closest(".group-filter-item");
      if (label) label.classList.toggle("selected", input.checked);
      // 테이블만 다시 렌더링
      preserveViewport(() => {
        renderSummary();
        renderPayables();
      });
    });
  });

  elements.groupFilterContainer.querySelectorAll(".group-manage-link").forEach(button => {
    button.addEventListener("click", event => {
      const action = event.currentTarget.dataset.action;
      if (action === "select-all") {
        filterState.groups = null;
      } else if (action === "clear-all") {
        filterState.groups = [];
      }
      rerenderAll();
    });
  });

  let draggingGroup = "";
  elements.groupFilterContainer.querySelectorAll(".group-filter-item-draggable").forEach(chip => {
    chip.addEventListener("dragstart", event => {
      draggingGroup = event.currentTarget.dataset.group || "";
      event.dataTransfer.effectAllowed = "move";
    });

    chip.addEventListener("dragover", event => {
      event.preventDefault();
      event.dataTransfer.dropEffect = "move";
    });

    chip.addEventListener("drop", event => {
      event.preventDefault();
      const targetGroup = event.currentTarget.dataset.group || "";
      if (!draggingGroup || !targetGroup || draggingGroup === targetGroup) return;

      const nextOrder = [...orderedGroups];
      const fromIndex = nextOrder.indexOf(draggingGroup);
      const toIndex = nextOrder.indexOf(targetGroup);
      if (fromIndex === -1 || toIndex === -1) return;

      nextOrder.splice(fromIndex, 1);
      nextOrder.splice(toIndex, 0, draggingGroup);

      filterState.groupOrder = nextOrder;
      saveGroupOrder();

      // UI 즉시 반영 (전체 리렌더링)
      preserveViewport(() => {
        renderGroupFilterControls(); // 버튼 순서 갱신
        renderPayables(); // 바뀐 순서에 맞춰 미지급 테이블도 갱신
      });
    });
  });
}

function buildGroupChipsHtml(allLabels, selectedFilter, chipClass) {
  // selectedFilter: null=전체, []=없음, [...]= 선택목록
  return allLabels.map(l => {
    const checked = selectedFilter === null || selectedFilter.includes(l);
    return `<span class="group-chip-item ${chipClass} ${checked ? "chip-on" : "chip-off"}" data-group="${escapeHtml(l)}">
      <span class="chip-drag-handle" draggable="true" data-group="${escapeHtml(l)}">⠿</span>
      <label class="chip-label">
        <input type="checkbox" class="chip-cb" data-group="${escapeHtml(l)}" ${checked ? "checked" : ""}/>
        <span>${escapeHtml(l)}</span>
      </label>
    </span>`;
  }).join("");
}

function setupGroupChipEvents(container, allLabels, getFilter, setFilter, setOrder, rerender) {
  // 체크박스
  container.querySelectorAll(".chip-cb").forEach(cb => {
    cb.addEventListener("change", e => {
      e.stopPropagation();
      const g = cb.dataset.group;
      const cur = getFilter() === null ? [...allLabels] : [...getFilter()];
      if (cb.checked) { if (!cur.includes(g)) cur.push(g); }
      else { const i = cur.indexOf(g); if (i !== -1) cur.splice(i, 1); }
      setFilter(cur.length === allLabels.length ? null : cur);
      saveGroupOrder();
      rerender();
    });
  });
  container.querySelector(".chip-select-all")?.addEventListener("click", () => { setFilter(null); saveGroupOrder(); rerender(); });
  container.querySelector(".chip-clear-all")?.addEventListener("click", () => { setFilter([]); saveGroupOrder(); rerender(); });

  // 드래그
  let dragging = "";
  container.querySelectorAll(".chip-drag-handle").forEach(handle => {
    handle.addEventListener("dragstart", e => {
      dragging = handle.dataset.group || "";
      e.dataTransfer.effectAllowed = "move";
      e.stopPropagation();
    });
  });
  container.querySelectorAll(".group-chip-item").forEach(chip => {
    chip.addEventListener("dragover", e => { e.preventDefault(); e.dataTransfer.dropEffect = "move"; });
    chip.addEventListener("drop", e => {
      e.preventDefault();
      const target = chip.dataset.group || "";
      if (!dragging || !target || dragging === target) return;
      const order = allLabels.slice();
      const from = order.indexOf(dragging), to = order.indexOf(target);
      if (from === -1 || to === -1) return;
      order.splice(from, 1); order.splice(to, 0, dragging);
      setOrder(order);
      saveGroupOrder();
      rerender();
    });
  });
}

function renderReceivables() {
  // 칩용: 전체 receivables에서 조건 목록 수집 (필터 전)
  const allCondLabels = (() => {
    const seen = new Map();
    receivables.forEach(i => { const c = i.condition || "기타"; if (!seen.has(c)) seen.set(c, true); });
    const base = [...rcvGroupState.order.filter(l => seen.has(l)),
    ...[...seen.keys()].filter(l => !rcvGroupState.order.includes(l))
      .sort((a, b) => { const r = getDueGroupRank(a) - getDueGroupRank(b); return r !== 0 ? r : a.localeCompare(b, "ko"); })];
    rcvGroupState.order = base;
    return base;
  })();

  const filtered = getFilteredItems(receivables, "receivables");
  const totalBalance = filtered.reduce((s, i) => s + Number(i.balance || 0), 0);
  const monthKeys = [...new Set(filtered.map(i => `${i.year}-${String(i.month).padStart(2, "0")}`))].sort();

  const yearsMap = new Map();
  monthKeys.forEach(mk => {
    const y = mk.split("-")[0];
    if (!yearsMap.has(y)) yearsMap.set(y, []);
    yearsMap.get(y).push(mk);
  });
  const years = [...yearsMap.keys()].sort();

  // 연도별 음영 클래스: 연도가 바뀔 때마다 even/odd 전환 + 첫 컬럼에 굵은 구분선
  const mkToYearIdx = {};
  const mkIsStart = new Set();
  years.forEach((y, yIdx) => {
    const mks = yearsMap.get(y) || [];
    mks.forEach((mk, i) => { mkToYearIdx[mk] = yIdx; if (i === 0) mkIsStart.add(mk); });
  });
  const mkcls = (mk) => {
    const ev = (mkToYearIdx[mk] || 0) % 2 === 0 ? "month-column-even" : "month-column-odd";
    return mkIsStart.has(mk) ? ev + " month-column-year-start" : ev;
  };

  const condGroups = new Map();
  filtered.forEach(item => {
    const cond = item.condition || "기타";
    if (!condGroups.has(cond)) condGroups.set(cond, { label: cond, vendors: new Map() });
    const vendors = condGroups.get(cond).vendors;
    const vKey = item.codeRaw || item.name;
    if (!vendors.has(vKey)) {
      vendors.set(vKey, { name: item.name, codeRaw: item.codeRaw || "", memo: item.memo, manager: item.manager || "", months: {}, total: 0, maxElapsed: null, latestDueDate: "" });
    }
    const v = vendors.get(vKey);
    const mk = `${item.year}-${String(item.month).padStart(2, "0")}`;
    v.months[mk] = (v.months[mk] || 0) + Number(item.balance || 0);
    v.total += Number(item.balance || 0);
    if (item.elapsed !== null && item.elapsed !== undefined && (v.maxElapsed === null || item.elapsed > v.maxElapsed)) {
      v.maxElapsed = item.elapsed; v.latestDueDate = item.dueDate || "";
    }
  });

  const visibleGroups = allCondLabels.filter(l => condGroups.has(l)).map(l => condGroups.get(l));

  const groupsHtml = visibleGroups.map(group => {
    const groupTotal = [...group.vendors.values()].reduce((s, v) => s + v.total, 0);
    const collapsed = Boolean(payablesGroupState.collapsed["rcv_" + group.label]);
    const sortedVendors = [...group.vendors.values()].sort((a, b) => {
      let cmp = 0;
      if (rcvSortState.key === "code") cmp = String(a.codeRaw || "").localeCompare(String(b.codeRaw || ""), undefined, { numeric: true });
      else if (rcvSortState.key === "elapsed") cmp = (a.maxElapsed ?? Infinity) - (b.maxElapsed ?? Infinity);
      else if (rcvSortState.key === "manager") cmp = String(a.manager || "").localeCompare(String(b.manager || ""), "ko");
      return rcvSortState.dir === "asc" ? cmp : -cmp;
    });

    const groupTotalCells = monthKeys.map((mk) => {
      const t = [...group.vendors.values()].reduce((s, v) => s + (v.months[mk] || 0), 0);
      return `<td class="group-summary-cell month-column-cell ${mkcls(mk)}">${t ? formatNumber(t) : ""}</td>`;
    }).join("");

    const itemRowsHtml = collapsed ? "" : sortedVendors.filter(v => {
      if (filterState.status === "excluded") return (v.condition === "제외" || (v.memo && v.memo.includes("제외")));
      return !(v.condition === "제외" || (v.memo && v.memo.includes("제외")));
    }).map((vendor, rowIdx) => {
      const el = vendor.maxElapsed;
      let elapsedHtml = "-", elapsedClass = "";
      if (el !== null && el !== undefined) {
        if (el >= 60) { elapsedHtml = `${el}일`; elapsedClass = "rcv-elapsed-danger"; }
        else if (el >= 30) { elapsedHtml = `${el}일`; elapsedClass = "rcv-elapsed-warn"; }
        else if (el >= 0) { elapsedHtml = `${el}일`; elapsedClass = "rcv-elapsed-ok"; }
        else { elapsedHtml = `D${el}`; elapsedClass = "rcv-elapsed-future"; }
      }
      const monthCells = monthKeys.map((mk) => {
        const val = vendor.months[mk] || 0;
        return `<td class="numeric-cell month-column-cell ${mkcls(mk)}">${val ? formatNumber(val) : ""}</td>`;
      }).join("");

      const vCode = normalizeVendorCode(vendor.codeRaw || vendor.name || "");
      const rcvTooltip = buildVendorTooltip(vCode, vendor.memo, "receivables");
      const memoAttr = rcvTooltip ? ` title="${rcvTooltip.replace(/"/g, "&quot;")}"` : "";
      const hasVMemo = !!(getVendorMemo(vCode).common || getVendorMemo(vCode).receivables);
      const mgrHtml = vendor.manager && vendor.manager !== "미지정" ? `<span class="rcv-manager-badge">${escapeHtml(vendor.manager)}</span>` : "";
      return `<tr class="${rowIdx % 2 === 0 ? "rcv-row-even" : "rcv-row-odd"}">
          <td class="partner-name-cell sticky-col-rcv-name">
            <div class="partner-name-cell-inner">
              <span class="partner-name-button truncate-text ${(vendor.memo || hasVMemo) ? "has-memo" : ""}"${memoAttr}>${escapeHtml(vendor.name)}</span>
              <button type="button" class="vendor-memo-btn" data-code="${escapeHtml(vCode)}" data-name="${escapeHtml(vendor.name)}" title="업체 메모 편집">✎</button>
              ${mgrHtml}
            </div>
          </td>
          <td class="numeric-cell"><span class="rcv-elapsed ${elapsedClass}">${elapsedHtml}</span></td>
          ${monthCells}
          <td class="numeric-cell item-total">${formatNumber(vendor.total)}</td>
        </tr>`;
    }).join("");

    return `<tr class="group-header rcv-group-header">
        <td class="sticky-col-rcv-name">
          <button type="button" class="group-toggle rcv-group-toggle" data-group="${escapeHtml(group.label)}">${collapsed ? "▶" : "▼"}</button>
          <strong>${escapeHtml(group.label)}</strong>
          <span class="group-count">${group.vendors.size}건</span>
        </td>
        <td></td>
        ${groupTotalCells}
        <td class="group-summary-cell group-total-cell">${formatNumber(groupTotal)}</td>
      </tr>
      ${itemRowsHtml}`;
  }).join("");

  const chipsHtml = buildGroupChipsHtml(allCondLabels, rcvGroupState.filter, "rcv-chip");

  // 담당자 마스터 '일' 미설정 업체 안내
  const noDaysVendors = [...new Set(
    receivables.filter(r => r.balance > 0 && !r.managerDays).map(r => r.name)
  )].sort();
  const rcvMgrDaysBanner = noDaysVendors.length
    ? `<div style="margin:4px 0 6px;padding:7px 12px;background:#fffbeb;border:1px solid #fbbf24;border-radius:6px;font-size:12px;color:#92400e;display:flex;gap:8px;align-items:flex-start;flex-wrap:wrap;">
        <span style="flex-shrink:0;">⚠</span>
        <span>담당자 마스터 <strong>"일"</strong> 컬럼 미설정:
          ${noDaysVendors.slice(0, 8).map(n => `<em>${escapeHtml(n)}</em>`).join(", ")}${noDaysVendors.length > 8 ? ` 외 ${noDaysVendors.length - 8}건` : ""}
          <span style="color:#b45309;margin-left:6px;">→ 담당자 마스터 시트에 <strong>"일"</strong> 컬럼을 추가하면 수금조건·납기일이 자동 적용됩니다.</span>
        </span>
      </div>`
    : "";

  const yearHeaders = years.map(y => {
    const count = yearsMap.get(y).length;
    return `<th class="year-header" colspan="${count}">
      <div class="year-header-inner">
        <span>${y}년</span>
      </div>
    </th>`;
  }).join("");

  const monthHeaders = monthKeys.map((mk) => {
    return `<th class="numeric-header month-column-cell ${mkcls(mk)}">${formatMonthKey(mk)}</th>`;
  }).join("");

  elements.receivables.innerHTML = `
    <div class="panel">
      <div class="panel-title-row">
        <div class="panel-title-inline">
          <h3>미수금 목록</h3>
          ${filtered.length ? `<span class="rcv-summary-text">${filtered.length}건 · ${formatNumber(totalBalance)}원</span>` : ""}
          <button type="button" class="rcv-email-btn" title="미수현황 메일 발송">메일 발송</button>
        </div>
        <div style="display:flex;align-items:center;gap:8px;">
          ${(() => { const s = calculateSummary(); return `<div class="tab-mini-stats"><span class="tms-item"><span class="tms-lbl">매출</span><span class="tms-val">${formatNumber(s.totalReceivable)}</span></span><span class="tms-sep">|</span><span class="tms-item tms-green"><span class="tms-lbl">미수금 잔액</span><span class="tms-val">${formatNumber(s.totalOutstanding)}</span></span></div>`; })()}
          <div class="payable-table-actions">
            <button type="button" class="table-action-button subtle rcv-expand-all">전체 펼치기</button>
            <button type="button" class="table-action-button subtle rcv-collapse-all">전체 접기</button>
          </div>
        </div>
      </div>
      ${rcvMgrDaysBanner}
      <div class="rcv-group-chips chips-orderable" id="rcvGroupChips">
        <button type="button" class="group-manage-link chip-select-all">전체 선택</button>
        <button type="button" class="group-manage-link chip-clear-all">전체 해제</button>
        ${chipsHtml}
      </div>
      <div class="table-scrollbar-top" id="rcvTopScrollbar"><div class="table-scrollbar-inner" id="rcvTopScrollbarInner"></div></div>
      <div class="table-responsive">
        <table class="rcv-pivot-table">
          <thead>
            <tr>
              <th rowspan="2" class="rcv-sort-th sticky-col-rcv-name" data-sort="code">거래처명</th>
              <th rowspan="2" class="rcv-sort-th" data-sort="elapsed">경과일수</th>
              ${yearHeaders}
              <th rowspan="2" class="numeric-header">합계</th>
            </tr>
            <tr>
              ${monthHeaders}
            </tr>
          </thead>
          <tbody>
            ${groupsHtml || `<tr><td colspan="${monthKeys.length + 3}" class="empty-state">조건에 맞는 데이터가 없습니다.</td></tr>`}
          </tbody>
        </table>
      </div>
    </div>
  `;

  elements.receivables.querySelector(".rcv-email-btn")?.addEventListener("click", openReceivableEmailDialog);

  elements.receivables.querySelectorAll(".vendor-memo-btn").forEach(btn => {
    btn.addEventListener("click", e => { e.stopPropagation(); openVendorMemoEditor(btn.dataset.code, btn.dataset.name); });
  });

  elements.receivables.querySelectorAll(".rcv-group-toggle").forEach(btn => {
    btn.addEventListener("click", () => {
      const g = btn.dataset.group;
      payablesGroupState.collapsed["rcv_" + g] = !payablesGroupState.collapsed["rcv_" + g];
      renderReceivables();
    });
  });
  elements.receivables.querySelector(".rcv-expand-all")?.addEventListener("click", () => {
    allCondLabels.forEach(l => { payablesGroupState.collapsed["rcv_" + l] = false; });
    renderReceivables();
  });
  elements.receivables.querySelector(".rcv-collapse-all")?.addEventListener("click", () => {
    allCondLabels.forEach(l => { payablesGroupState.collapsed["rcv_" + l] = true; });
    renderReceivables();
  });

  const chipsContainer = document.getElementById("rcvGroupChips");
  if (chipsContainer) {
    setupGroupChipEvents(
      chipsContainer, allCondLabels,
      () => rcvGroupState.filter,
      v => { rcvGroupState.filter = v; },
      order => { rcvGroupState.order = order; },
      renderReceivables
    );
  }

  elements.receivables.querySelectorAll(".rcv-sort-th").forEach(th => {
    th.addEventListener("click", () => {
      const key = th.dataset.sort;
      if (rcvSortState.key === key) rcvSortState.dir = rcvSortState.dir === "asc" ? "desc" : "asc";
      else { rcvSortState.key = key; rcvSortState.dir = "asc"; }
      renderReceivables();
    });
  });
}

// ── 미수금 이메일 발송 ───────────────────────────────────────

function resolveReceivableAbsenceTarget(absentSet) {
  for (const person of RECEIVABLE_ABSENCE_CHAIN) {
    if (!absentSet.has(person.name)) return person;
  }
  return null;
}

function openReceivableEmailDialog() {
  document.querySelector(".rcv-email-overlay")?.remove();

  const managers = [...new Set(receivables
    .map(i => i.manager).filter(m => m && m !== "미지정"))].sort();
  const conditions = [...new Set(receivables.map(i => i.condition).filter(Boolean))].sort();

  // 조건 → 담당자 맵: 어떤 조건에 어떤 담당자가 있는지
  const condToManagers = new Map();
  receivables.forEach(item => {
    if (!item.condition || !item.manager || item.manager === "미지정") return;
    if (!condToManagers.has(item.condition)) condToManagers.set(item.condition, new Set());
    condToManagers.get(item.condition).add(item.manager);
  });

  const overlay = document.createElement("div");
  overlay.className = "rcv-email-overlay";

  const absChainLabel = RECEIVABLE_ABSENCE_CHAIN.map(c => c.name).join(" → ");

  overlay.innerHTML = `
    <div class="rcv-email-dialog">
      <div class="rcv-email-dialog-header">
        <h3>미수금 이메일 발송</h3>
        <button type="button" class="rcv-close-btn">✕</button>
      </div>

      <div class="rcv-email-section">
        <label class="rcv-test-label">
          <input type="checkbox" id="rcvTestMode" checked />
          🧪 테스트 모드
        </label>
        <div class="rcv-test-recipients" id="rcvTestRecips">
          ${RECEIVABLE_TEST_RECIPIENTS.map((r, i) => `
            <label><input type="radio" name="rcvTestRecip" value="${r.email}" ${i === 0 ? "checked" : ""}> ${r.name}</label>
          `).join("")}
        </div>
      </div>

      <div class="rcv-email-section">
        <div class="rcv-section-title">부재자 체인 <span class="rcv-chain-note">${absChainLabel}</span></div>
        <div class="rcv-absence-chain">
          ${RECEIVABLE_ABSENCE_CHAIN.map((c, i) => `
            ${i > 0 ? '<span class="rcv-chain-arrow">→</span>' : ""}
            <label class="rcv-chain-person">
              <input type="checkbox" class="rcv-global-abs-chk" value="${c.name}"> ${c.name}
            </label>
          `).join("")}
        </div>
      </div>

      <div class="rcv-email-section">
        <div class="rcv-section-title">담당자별 발송</div>
        <div class="rcv-manager-list">
          ${managers.map(m => `
            <div class="rcv-mgr-row">
              <label class="rcv-mgr-label">
                <input type="checkbox" class="rcv-mgr-chk" value="${m}" checked> ${m}
              </label>
              <label class="rcv-abs-label">
                <input type="checkbox" class="rcv-abs-chk" data-manager="${m}"> 부재
              </label>
              <span class="rcv-chain-result" id="rcv-cr-${m.replace(/\s/g, "_")}"></span>
            </div>
          `).join("")}
        </div>
      </div>

      <div class="rcv-email-section">
        <div class="rcv-section-title">수금조건 필터</div>
        <div class="rcv-cond-grid">
          ${conditions.map(c => `
            <label class="rcv-cond-item"><input type="checkbox" class="rcv-cond-chk" value="${c}" checked> ${c}</label>
          `).join("")}
        </div>
        <div class="rcv-cond-actions">
          <button type="button" class="rcv-cond-all">전체선택</button>
          <button type="button" class="rcv-cond-none">전체해제</button>
        </div>
      </div>

      <div class="rcv-email-section">
        <label class="rcv-section-title" style="display:flex;align-items:center;gap:6px;cursor:pointer;">
          <input type="checkbox" id="rcvSendSummary" checked>
          전체 현황 보고서 (모든 데이터 통합 표)
        </label>
        <div class="rcv-summary-recipients" id="rcvSummaryRecips" style="margin: 8px 0 12px 24px; display:flex; flex-wrap:wrap; gap:12px; font-size: 0.95rem;">
          <label><input type="checkbox" class="rcv-sum-recip-chk" value="kdy@mauto.co.kr" checked> 김도연</label>
          <label><input type="checkbox" class="rcv-sum-recip-chk" value="jug@mauto.co.kr" checked> 장운기</label>
          <label><input type="checkbox" class="rcv-sum-recip-chk" value="phs@mauto.co.kr"> 박희선</label>
          <label><input type="checkbox" class="rcv-sum-recip-chk" value="yhj@mauto.co.kr"> 여희정</label>
        </div>
        <div class="rcv-summary-opts" id="rcvSummaryOpts">
          <label><input type="radio" name="rcvDOpt" value="include" checked> D- 포함</label>
          <label><input type="radio" name="rcvDOpt" value="exclude"> D- 제외</label>
        </div>
      </div>

      <div class="rcv-email-section">
        <div class="rcv-section-title">참조 (CC)</div>
        <div class="rcv-cc-grid">
          ${RECEIVABLE_CC_OPTIONS.map(c => `
            <label class="rcv-cc-item"><input type="checkbox" class="rcv-cc-chk" value="${c.email}"> ${c.name}</label>
          `).join("")}
        </div>
      </div>

      <div class="rcv-email-section rcv-sender-row">
        <label class="rcv-section-title" style="margin-bottom:4px;">발신자 이름</label>
        <input type="text" id="rcvSenderName" placeholder="예: 홍길동" style="padding:6px 10px;border:1px solid #cbd5e1;border-radius:8px;font-size:0.9rem;width:180px;" />
      </div>
      <div class="rcv-email-actions">
        <button type="button" class="rcv-select-all-btn">전체 선택/해제</button>
        <button type="button" class="rcv-cancel-btn">취소</button>
        <button type="button" class="rcv-send-btn bg-emerald-600 text-white hover:bg-emerald-700">다음: 미리보기</button>
      </div>
    </div>
  `;

  document.body.appendChild(overlay);

  const q = sel => overlay.querySelector(sel);
  const qa = sel => [...overlay.querySelectorAll(sel)];

  q(".rcv-close-btn").addEventListener("click", () => overlay.remove());
  q(".rcv-cancel-btn").addEventListener("click", () => overlay.remove());
  q("#rcvTestMode").addEventListener("change", e => {
    q("#rcvTestRecips").style.display = e.target.checked ? "flex" : "none";
  });
  q("#rcvSendSummary").addEventListener("change", e => {
    q("#rcvSummaryOpts").style.display = e.target.checked ? "flex" : "none";
  });
  function updateManagersByCondition() {
    const checkedConds = new Set(qa(".rcv-cond-chk:checked").map(c => c.value));
    // 선택된 조건에 해당하는 담당자 집합
    const activeMgrs = new Set();
    checkedConds.forEach(cond => {
      (condToManagers.get(cond) || new Set()).forEach(m => activeMgrs.add(m));
    });
    qa(".rcv-mgr-row").forEach(row => {
      const chk = row.querySelector(".rcv-mgr-chk");
      const mgr = chk?.value;
      const hasData = activeMgrs.has(mgr);
      row.style.opacity = hasData ? "1" : "0.35";
      if (chk) chk.disabled = !hasData;
      if (chk && !hasData) chk.checked = false;
    });
  }

  q(".rcv-cond-all").addEventListener("click", () => {
    qa(".rcv-cond-chk").forEach(c => c.checked = true);
    updateManagersByCondition();
  });
  q(".rcv-cond-none").addEventListener("click", () => {
    qa(".rcv-cond-chk").forEach(c => c.checked = false);
    updateManagersByCondition();
  });
  qa(".rcv-cond-chk").forEach(chk => chk.addEventListener("change", updateManagersByCondition));

  updateManagersByCondition(); // 초기화

  q(".rcv-select-all-btn").addEventListener("click", () => {
    const all = qa("input[type=checkbox]");
    const allOn = all.every(b => b.checked);
    all.forEach(b => b.checked = !allOn);
  });

  qa(".rcv-global-abs-chk").forEach(chk => chk.addEventListener("change", () => {
    const globalAbsent = new Set(qa(".rcv-global-abs-chk:checked").map(c => c.value));
    qa(".rcv-abs-chk").forEach(ac => {
      if (RECEIVABLE_ABSENCE_CHAIN.some(c => c.name === ac.dataset.manager)) {
        ac.checked = globalAbsent.has(ac.dataset.manager);
      }
    });
    updateRcvChainResults(overlay);
  }));
  qa(".rcv-abs-chk").forEach(chk => chk.addEventListener("change", () => updateRcvChainResults(overlay)));

  q(".rcv-send-btn").addEventListener("click", () => doSendReceivableEmails(overlay));
}

function updateRcvChainResults(overlay) {
  const absentSet = new Set(
    [...overlay.querySelectorAll(".rcv-abs-chk:checked")].map(c => c.dataset.manager)
  );
  overlay.querySelectorAll(".rcv-abs-chk").forEach(chk => {
    const mgr = chk.dataset.manager;
    const el = overlay.querySelector(`#rcv-cr-${mgr.replace(/\s/g, "_")}`);
    if (!el) return;
    if (chk.checked) {
      const target = resolveReceivableAbsenceTarget(absentSet);
      el.textContent = target ? `→ ${target.name}` : "→ ⚠️ 수신가능자 없음";
      el.style.color = target ? "#1d4ed8" : "#dc2626";
      el.style.display = "inline";
    } else {
      el.style.display = "none";
    }
  });
}

async function doSendReceivableEmails(overlay) {
  const q = sel => overlay.querySelector(sel);
  const testMode = q("#rcvTestMode").checked;
  const testRecipEl = q("input[name=rcvTestRecip]:checked");
  const testRecipient = testMode && testRecipEl ? testRecipEl.value : null;

  const managers = [...overlay.querySelectorAll(".rcv-mgr-chk:checked")].map(c => {
    const absEl = overlay.querySelector(`.rcv-abs-chk[data-manager="${c.value}"]`);
    return { manager: c.value, absent: absEl ? absEl.checked : false };
  });
  const absentChain = [...overlay.querySelectorAll(".rcv-global-abs-chk:checked")].map(c => c.value);
  const conditions = [...overlay.querySelectorAll(".rcv-cond-chk:checked")].map(c => c.value);
  const ccEmails = [...overlay.querySelectorAll(".rcv-cc-chk:checked")].map(c => c.value);
  const sendSummary = q("#rcvSendSummary").checked;
  const summaryRecipients = [...overlay.querySelectorAll(".rcv-sum-recip-chk:checked")].map(c => c.value);
  const excludeMinus = q("input[name=rcvDOpt]:checked")?.value === "exclude";
  const senderName = (q("#rcvSenderName")?.value || "").trim();

  if (!managers.length && !sendSummary) { alert("담당자를 최소 1명 선택하거나 전체 현황 보고서를 선택해주세요."); return; }
  if (sendSummary && !summaryRecipients.length) { alert("전체 현황 보고서를 수신할 사람을 최소 1명 선택해주세요."); return; }

  // 조건이 1개도 선택되지 않은 경우 백엔드(Apps Script)에서는 condSet.size === 0 이 되어 모든 조건의 항목이 포함됩니다.
  // 사용자가 "조건을 모두 체크 해제하면 전체 보고서가 나가는지?" 혼동할 수 있으므로, 0개 선택을 '전체 조건 발송'으로 허용합니다.

  const sendBtn = q(".rcv-send-btn");
  sendBtn.disabled = true;
  sendBtn.textContent = "가져오는 중...";

  const payload = {
    managers, absentChain, ccEmails, conditions,
    testMode, testRecipient, sendSummary, excludeMinus, senderName, summaryRecipients
  };

  try {
    const result = await postSheetWebApp("sendReceivableEmails", { ...payload, previewMode: true });
    overlay.style.display = "none"; // Hide first dialog temporally
    openReceivablePreviewDialog(result.previews || [], payload, overlay);
  } catch (error) {
    alert(`미리보기 생성 실패: ${error.message}`);
    sendBtn.disabled = false;
    sendBtn.textContent = "다음: 미리보기";
  }
}

function openReceivablePreviewDialog(previews, payload, parentOverlay) {
  if (!previews.length) {
    alert("생성된 미리보기 메일이 없습니다. 선택 조건을 확인해주세요.");
    parentOverlay.style.display = "flex";
    const btn = parentOverlay.querySelector(".rcv-send-btn");
    if (btn) { btn.disabled = false; btn.textContent = "다음: 미리보기"; }
    return;
  }

  const overlay = document.createElement("div");
  overlay.className = "rcv-email-overlay";
  overlay.style.zIndex = "3000";

  let activeIdx = 0;

  function render() {
    const p = previews[activeIdx];
    const tabsHtml = previews.map((pr, idx) => `
      <button type="button" class="rcv-preview-tab ${idx === activeIdx ? "active" : ""}" data-idx="${idx}">
        ${escapeHtml(pr.id === "summary" ? "📋 전체 보고서" : (pr.id === "absent" ? "🏢 부재 통합" : "👤 " + pr.id.replace("mgr_", "")))}
      </button>
    `).join("");

    overlay.innerHTML = `
      <div class="modal-content rcv-email-modal" style="max-width:1000px; width:95%; height:90vh; display:flex; flex-direction:column;">
        <div style="display:flex; justify-content:space-between; align-items:center; border-bottom:1px solid #e2e8f0; padding-bottom:10px; margin-bottom:15px;">
          <h2 style="margin:0; font-size:1.2rem; color:#1e293b;">메일 미리보기 및 멘트 작성 (총 ${previews.length}건)</h2>
          <button type="button" class="rcv-prev-close-btn" style="background:none;border:none;font-size:1.5rem;cursor:pointer;">&times;</button>
        </div>
        <div class="rcv-preview-tabs" style="display:flex; flex-wrap:wrap; gap:8px; margin-bottom:15px;">
          ${tabsHtml}
        </div>
        
        <div style="display:flex; flex-direction:column; flex:1; min-height:0;">
          <div style="background:#f8fafc; padding:10px 15px; border:1px solid #e2e8f0; border-radius:6px; margin-bottom:10px; font-size:0.9rem;">
            <div><strong>수신:</strong> ${escapeHtml(p.to)}</div>
            <div style="margin-top:4px;"><strong>제목:</strong> ${escapeHtml(p.subject)}</div>
          </div>
          
          <div style="margin-bottom:10px;">
            <label style="display:block; font-size:0.9rem; font-weight:700; color:#334155; margin-bottom:6px;">추가 멘트 삽입 (선택)</label>
            <textarea id="rcvCustomMsgInput" class="custom-input" rows="3" style="width:100%; resize:vertical;" placeholder="이 메일의 가장 윗부분에 강조되어 들어갈 추가 멘트를 적어주세요.">${escapeHtml(payload.customMessages?.[p.id] || "")}</textarea>
          </div>

          <div style="flex:1; border:1px solid #cbd5e1; border-radius:4px; overflow:auto; background:#fff; padding:20px;">
            ${p.htmlBody}
          </div>
        </div>
        
        <div style="display:flex; justify-content:space-between; align-items:center; margin-top:15px; border-top:1px solid #e2e8f0; padding-top:15px;">
          <button type="button" class="rcv-prev-back-btn button-secondary">이전 (설정)</button>
          <div>
            <button type="button" class="rcv-prev-cancel-btn button-secondary">취소</button>
            <button type="button" class="rcv-prev-send-btn button-primary bg-blue-600 hover:bg-blue-700 font-bold px-6 border-none text-white">✈️ 최종 발송</button>
          </div>
        </div>
      </div>
    `;

    overlay.querySelector(".rcv-prev-close-btn").addEventListener("click", () => { overlay.remove(); parentOverlay.remove(); });
    overlay.querySelector(".rcv-prev-cancel-btn").addEventListener("click", () => { overlay.remove(); parentOverlay.remove(); });
    overlay.querySelector(".rcv-prev-back-btn").addEventListener("click", () => {
      saveCurrentCustomMsg();
      overlay.remove();
      parentOverlay.style.display = "flex";
      const btn = parentOverlay.querySelector(".rcv-send-btn");
      if (btn) { btn.disabled = false; btn.textContent = "다음: 미리보기"; }
    });

    overlay.querySelectorAll(".rcv-preview-tab").forEach(btn => {
      btn.addEventListener("click", (e) => {
        saveCurrentCustomMsg();
        activeIdx = Number(e.currentTarget.dataset.idx);
        render();
      });
    });

    const sendBtn = overlay.querySelector(".rcv-prev-send-btn");
    sendBtn.addEventListener("click", async () => {
      saveCurrentCustomMsg();
      sendBtn.disabled = true;
      sendBtn.textContent = "전송 중...";
      try {
        const finalPayload = { ...payload, previewMode: false };
        const result = await postSheetWebApp("sendReceivableEmails", finalPayload);
        overlay.remove();
        parentOverlay.remove();
        const modeNote = payload.testMode ? `\n※ 테스트: ${payload.testRecipient || ""}으로 발송` : "";
        alert(`발송 완료! ${result.sentCount || ""}건${modeNote}`);
      } catch (error) {
        alert(`발송 실패: ${error.message}`);
        sendBtn.disabled = false;
        sendBtn.textContent = "✈️ 최종 발송";
      }
    });
  }

  function saveCurrentCustomMsg() {
    const input = overlay.querySelector("#rcvCustomMsgInput");
    if (input && previews[activeIdx]) {
      payload.customMessages = payload.customMessages || {};
      payload.customMessages[previews[activeIdx].id] = input.value.trim();
    }
  }

  document.body.appendChild(overlay);
  render();
}

function getMonthKey(item) {
  const year = Number(item.year || 0);
  const month = Number(item.month || 0);
  if (!year || !month) return "";
  return `${year}-${String(month).padStart(2, "0")}`;
}

function getUniqueSortedMonthKeys(items) {
  return [...new Set(items.map(getMonthKey).filter(Boolean))].sort();
}

function calcPayableMonthTotals(filteredPayables, monthKeys) {
  const monthTotals = monthKeys.reduce((totals, key) => {
    totals[key] = 0;
    return totals;
  }, {});
  let total = 0;

  filteredPayables.forEach(item => {
    const decisionValue = item.decisionAmount != null ? item.decisionAmount : getPayableOutstanding(item);
    const key = getMonthKey(item);
    if (key && monthTotals[key] !== undefined) {
      monthTotals[key] += decisionValue;
    }
    total += decisionValue;
  });

  return { monthTotals, total };
}

function calcSelectedMonthTotals(filteredPayables, monthKeys) {
  const monthTotals = monthKeys.reduce((totals, key) => {
    totals[key] = 0;
    return totals;
  }, {});
  let total = 0;
  let count = 0;

  filteredPayables.forEach(item => {
    if (!item.selected) return;
    const decisionValue = item.decisionAmount != null ? item.decisionAmount : getPayableOutstanding(item);
    const key = getMonthKey(item);
    if (key && monthTotals[key] !== undefined) {
      monthTotals[key] += decisionValue;
    }
    total += decisionValue;
    count += 1;
  });

  return { monthTotals, total, count };
}

function calcPaymentPlanSummary(filteredPayables) {
  const buckets = new Map();
  let totalAmount = 0;
  let totalCount = 0;

  filteredPayables.forEach(item => {
    const planKey = item.paymentPlan || "";
    const amount = item.decisionAmount != null ? item.decisionAmount : getPayableOutstanding(item);
    totalAmount += amount;
    totalCount += 1;
    if (!buckets.has(planKey)) {
      buckets.set(planKey, { label: formatPlanLabel(planKey), count: 0, amount: 0, key: planKey });
    }
    const bucket = buckets.get(planKey);
    bucket.count += amount > 0 ? 1 : 0;
    bucket.amount += amount;
  });

  return [
    { label: "전체 예정", count: totalCount, amount: totalAmount, key: "__total__" },
    ...[...buckets.values()].sort((a, b) => {
      if (!a.key && !b.key) return 0;
      if (!a.key) return 1;
      if (!b.key) return -1;
      return a.key.localeCompare(b.key);
    }),
  ];
}

function getItemAutoPayDate(item) {
  return calcPayableDueDate(Number(item.year || 0), Number(item.month || 0), getDueGroup(item));
}

function ensureAutoPaymentPlans() {
  payables.forEach(item => {
    const auto = getItemAutoPayDate(item);
    if (!auto) return; // 계산 불가한 항목은 건드리지 않음
    const status = (item.completionStatus || "").trim();
    // 구 버그: 당N일이 N일로 분류돼 month+2로 계산됐던 날짜 감지 → 자동값으로 재계산
    const group = getDueGroup(item);
    const buggyM = group.match(/^당(\d+)일$/);
    const oldBuggyAuto = buggyM
      ? calcPayableDueDate(Number(item.year), Number(item.month), buggyM[1] + "일")
      : null;
    const isOldBuggyPlan = oldBuggyAuto && item.paymentPlan === oldBuggyAuto;
    // 보류/완료/미정 은 항상 보호. 예정이라도 사용자가 자동값과 다른 날짜를 직접 지정했으면 보호.
    // 단, 구 버그로 계산된 날짜는 수동 설정으로 보지 않음.
    const isManuallySet = status === "보류" || status === "완료" || status === "미정" ||
      (!isOldBuggyPlan && item.paymentPlan && item.paymentPlan !== auto);
    if (!isManuallySet) {
      item.paymentPlan = auto;
    }
    if (!item.sourceKey) {
      item.sourceKey = buildPayableSourceKey(item);
    }
  });
}

function getPartnerGroupKey(item) {
  return `${item.code || ""}||${item.name || ""}`;
}

function getOrderedDueGroups(filteredPayables) {
  const availableGroups = [...new Set(filteredPayables.map(getDueGroup).filter(Boolean))];
  // 커스텀 순서 적용 (드래그로 변경된 순서)
  const customOrdered = filterState.groupOrder.filter(g => availableGroups.includes(g));
  const remaining = availableGroups
    .filter(g => !filterState.groupOrder.includes(g))
    .sort((a, b) => {
      const rankDiff = getDueGroupRank(a) - getDueGroupRank(b);
      if (rankDiff !== 0) return rankDiff;
      return String(a).localeCompare(String(b), "ko");
    });
  return [...customOrdered, ...remaining];
}

function groupPayablesByDue(filteredPayables) {
  const groups = new Map();
  filteredPayables.forEach(item => {
    const key = getDueGroup(item);
    if (!groups.has(key)) {
      groups.set(key, { label: key, items: [] });
    }
    groups.get(key).items.push(item);
  });

  const consolidatedGroups = [...groups.values()].map(group => {
    const aggregated = new Map();
    group.items.forEach(item => {
      const partnerKey = getPartnerGroupKey(item);
      if (!aggregated.has(partnerKey)) {
        aggregated.set(partnerKey, {
          code: item.code,
          name: item.name,
          dueCategory: item.dueCategory,
          memo: item.memo,
          items: [],
          monthTotals: {},
          total: 0,
          selected: false,
        });
      }
      const entry = aggregated.get(partnerKey);
      const monthKey = getMonthKey(item);
      const decisionValue = item.decisionAmount != null ? item.decisionAmount : getPayableOutstanding(item);
      entry.items.push(item);
      entry.monthTotals[monthKey] = (entry.monthTotals[monthKey] || 0) + decisionValue;
      entry.total += decisionValue;
      entry.selected = entry.selected || item.selected;
    });
    const consolidated = [...aggregated.values()];
    consolidated.sort((a, b) => {
      const codeCompare = String(a.code || "").localeCompare(String(b.code || ""));
      if (codeCompare !== 0) return codeCompare;
      return String(a.name || "").localeCompare(String(b.name || ""));
    });
    return { ...group, items: consolidated };
  });

  const order = getOrderedDueGroups(filteredPayables);
  return consolidatedGroups.sort((a, b) => order.indexOf(a.label) - order.indexOf(b.label));
}

function renderPayables() {
  ensureAutoPaymentPlans();
  renderGroupFilterControls();

  // 담당자 마스터 '일' 미설정 미지급 업체 안내
  const payNoDaysVendors = [...new Set(
    payables.filter(p => getPayableOutstanding(p) > 0 && !p.managerDays).map(p => p.name)
  )].sort();
  const payMgrDaysBanner = payNoDaysVendors.length
    ? `<div style="margin:4px 0 6px;padding:7px 12px;background:#fffbeb;border:1px solid #fbbf24;border-radius:6px;font-size:12px;color:#92400e;display:flex;gap:8px;align-items:flex-start;flex-wrap:wrap;">
        <span style="flex-shrink:0;">⚠</span>
        <span>담당자 마스터 <strong>"일"</strong> 컬럼 미설정:
          ${payNoDaysVendors.slice(0, 8).map(n => `<em>${escapeHtml(n)}</em>`).join(", ")}${payNoDaysVendors.length > 8 ? ` 외 ${payNoDaysVendors.length - 8}건` : ""}
          <span style="color:#b45309;margin-left:6px;">→ 담당자 마스터 시트의 <strong>"일"</strong> 컬럼이 없으면 ERP 납기 원본값을 사용합니다.</span>
        </span>
      </div>`
    : "";

  const filteredPayables = getFilteredItems(payables, "payables");
  const matchedVendorCount = [...new Set(filteredPayables.filter(item => item.vendorMatched).map(item => getPartnerGroupKey(item)))].length;
  const unmatchedVendorCount = [...new Set(filteredPayables.filter(item => !item.vendorMatched).map(item => getPartnerGroupKey(item)))].length;
  const monthKeys = getUniqueSortedMonthKeys(filteredPayables);
  // 연도별 그룹 수집
  const yearsMap = new Map();
  monthKeys.forEach(mk => {
    const y = mk.split("-")[0];
    if (!yearsMap.has(y)) yearsMap.set(y, []);
    yearsMap.get(y).push(mk);
  });
  const years = [...yearsMap.keys()].sort();

  // displayKeys: 접힌 연도는 __year__YYYY 단일 키, 펼쳐진 연도는 개별 월 키
  const displayKeys = [];
  years.forEach(y => {
    if (payablesYearCollapsed[y]) {
      displayKeys.push(`__year__${y}`);
    } else {
      yearsMap.get(y).forEach(mk => displayKeys.push(mk));
    }
  });

  const dkToYearIdx = {};
  const dkIsStart = new Set();
  years.forEach((y, yIdx) => {
    if (payablesYearCollapsed[y]) {
      dkToYearIdx[`__year__${y}`] = yIdx;
      dkIsStart.add(`__year__${y}`);
    } else {
      yearsMap.get(y).forEach((mk, i) => {
        dkToYearIdx[mk] = yIdx;
        if (i === 0) dkIsStart.add(mk);
      });
    }
  });
  const dkcls = (dk) => {
    const ev = (dkToYearIdx[dk] || 0) % 2 === 0 ? "month-column-even" : "month-column-odd";
    return dkIsStart.has(dk) ? ev + " month-column-year-start" : ev;
  };

  const groups = groupPayablesByDue(filteredPayables);
  const paymentPlanSummary = calcPaymentPlanSummary(filteredPayables);
  const availablePlanKeys = paymentPlanSummary.map(item => item.key);
  paymentPlanUiState.selectedPlanKeys = paymentPlanUiState.selectedPlanKeys.filter(key => availablePlanKeys.includes(key));
  const hasSelectedCards = paymentPlanUiState.selectedPlanKeys.length > 0;
  const hasSelectedRows = filteredPayables.some(item => item.selected);
  const showBatchButton = hasSelectedCards || hasSelectedRows;

  const rows = groups.map(group => {
    const groupKey = group.label || "기타";
    const collapsed = Boolean(payablesGroupState.collapsed[groupKey]);
    const groupSourceItems = filteredPayables.filter(item => getDueGroup(item) === group.label);
    const groupTotals = calcPayableMonthTotals(groupSourceItems, monthKeys);
    const planCounts = group.items.reduce((acc, entry) => {
      entry.items.forEach(item => {
        const key = item.paymentPlan || "미정";
        acc[key] = (acc[key] || 0) + 1;
      });
      return acc;
    }, {});
    const planSummary = Object.keys(planCounts)
      .filter(key => key !== "미정")
      .sort()
      .map(key => `${key === "보류" ? key : /^\d{4}-\d{2}-\d{2}$/.test(key) ? key.slice(5).replace("-", "/") : key} ${planCounts[key]}건`)
      .join(" · ");

    const groupSummaryCells = displayKeys.map((dk, idx) => {
      let val = 0;
      if (dk.startsWith("__year__")) {
        const y = dk.replace("__year__", "");
        (yearsMap.get(y) || []).forEach(mk => { val += groupTotals.monthTotals[mk] || 0; });
      } else {
        val = groupTotals.monthTotals[dk] || 0;
      }
      return `<td class="group-summary-cell month-column-cell ${dkcls(dk)}">${formatPayableCellNumber(val)}</td>`;
    }).join("");
    const header = `
      <tr class="group-header" data-group="${groupKey}">
        <td colspan="2">
          <button type="button" class="group-toggle" data-group="${groupKey}" aria-expanded="${!collapsed}">
            ${collapsed ? "▶" : "▼"}
          </button>
          <strong>${group.label}</strong>
          <span>${group.items.length}건</span>
          ${planSummary ? `<span class="group-plan-summary">${planSummary}</span>` : ""}
        </td>
        ${groupSummaryCells}
        <td class="group-summary-cell group-total-cell">${formatPayableCellNumber(groupTotals.total)}</td>
      </tr>
    `;

    const itemRows = collapsed ? "" : group.items.filter(entry => {
      // '제외' 필터링
      const firstItem = entry.items[0];
      if (filterState.status === "excluded") return firstItem.completionStatus === "제외";
      return firstItem.completionStatus !== "제외";
    }).map(entry => {
      const checked = entry.selected ? "checked" : "";
      const partnerKey = encodeURIComponent(getPartnerGroupKey(entry.items[0]));

      const monthCells = displayKeys.map((dk, idx) => {
        // 연도 접힘: 해당 연도 전체 합계 표시
        if (dk.startsWith("__year__")) {
          const y = dk.replace("__year__", "");
          const yearMks = yearsMap.get(y) || [];
          const yearVal = yearMks.reduce((s, mk) => s + (entry.monthTotals[mk] || 0), 0);
          return `<td class="editable-amount-cell numeric-cell month-column-cell year-collapsed-cell ${dkcls(dk)}">${yearVal ? formatPayableCellNumber(yearVal) : ""}</td>`;
        }
        const monthKey = dk;
        const decisionValue = entry.monthTotals[monthKey] || 0;
        const monthItems = entry.items.filter(item => getMonthKey(item) === monthKey);
        const originalValue = monthItems.reduce((sum, item) => sum + getPayableOutstanding(item), 0);
        const totalPurchase = monthItems.reduce((sum, item) => sum + Number(item.purchase || 0), 0);
        const totalRawPaid = monthItems.reduce((sum, item) => sum + Number(item.paid || 0), 0);
        const cellPlanValue = monthItems[0]?.paymentPlan || "";
        const autoPlanValue = monthItems[0] ? getItemAutoPayDate(monthItems[0]) : "";
        const isMijeong = monthItems.some(item => item.completionStatus === "미정");
        const planClass = cellPlanValue === "보류" || cellPlanValue === "제외" ? "hold" : cellPlanValue ? "set" : "pending";
        const planLabel = cellPlanValue === "제외" ? "제외" :
                          cellPlanValue ? formatPlanShortLabel(cellPlanValue) :
                          isMijeong ? "미정" :
                          formatPlanShortLabel(autoPlanValue || "");
        const showOriginalValue = originalValue > 0 && decisionValue !== originalValue;
        const showRawBreakdown = totalRawPaid > 0 && totalPurchase > originalValue;
        const isLastEdited = payablesUiState.lastEdited
          && payablesUiState.lastEdited.partnerKey === getPartnerGroupKey(entry.items[0])
          && payablesUiState.lastEdited.monthKey === monthKey;

        if (originalValue === 0) {
          return `<td class="editable-amount-cell numeric-cell month-column-cell ${dkcls(dk)}"></td>`;
        }
        return `
          <td class="editable-amount-cell numeric-cell month-column-cell ${dkcls(dk)} ${isLastEdited ? "recently-edited-cell" : ""}">
            <div class="amount-cell-topline">
              <span class="cell-plan-badge ${planClass}">${planLabel}</span>
              <button
                type="button"
                class="edit-amount-button"
                data-partner-key="${partnerKey}"
                data-month-key="${monthKey}"
              >
                ${formatPayableCellNumber(decisionValue)}
              </button>
            </div>
            ${showRawBreakdown ? `<span class="amount-raw-breakdown" title="합계 ${formatNumber(totalPurchase)} / 지급 ${formatNumber(totalRawPaid)}">합계 ${formatNumber(totalPurchase)} · 지급 ${formatNumber(totalRawPaid)}</span>` : ""}
            ${showOriginalValue && !showRawBreakdown ? `<button type="button" class="amount-original-button" data-partner-key="${partnerKey}" data-month-key="${monthKey}" title="원래 금액으로 되돌리기">원래 ${formatNumber(originalValue)}</button>` : ""}
          </td>
        `;
      }).join("");

      return `
        <tr>
          <td class="sticky-col sticky-col-1"><label><input type="checkbox" class="payable-select-checkbox" data-partner-key="${partnerKey}" ${checked} /></label></td>
          <td class="sticky-col sticky-col-2 partner-name-cell">
            <div class="partner-name-cell-inner">
              ${(() => {
          const pCode = normalizeVendorCode(entry.code || entry.name || "");
          const payTooltip = buildVendorTooltip(pCode, entry.memo, "payables");
          const hasVMemo = !!(getVendorMemo(pCode).common || getVendorMemo(pCode).payables);
          const titleAttr = payTooltip ? ` title="${payTooltip.replace(/"/g, "&quot;")}"` : "";
          return `<span class="partner-name-button ${(entry.memo || hasVMemo) ? "has-memo" : ""}" ${entry.items[0]?.vendorMatched ? "data-vendor-matched=\"true\"" : "data-vendor-matched=\"false\""} ${titleAttr}>${entry.name}</span>
              <button type="button" class="vendor-memo-btn" data-code="${escapeHtml(pCode)}" data-name="${escapeHtml(entry.name)}" title="업체 메모 편집">✎</button>
              ${entry.items[0]?.vendorMatched ? "" : '<span class="vendor-match-chip unmatched">계좌확인</span>'}`;
        })()}
            </div>
          </td>
          ${monthCells}
          <td class="item-total numeric-cell">${formatPayableCellNumber(entry.total)}</td>
        </tr>
      `;
    }).join("");

    return header + itemRows;
  }).join("");

  const yearHeaders = years.map(y => {
    const collapsed = !!payablesYearCollapsed[y];
    const colspan = collapsed ? 1 : yearsMap.get(y).length;
    return `<th class="year-group-header${collapsed ? " collapsed" : ""}" colspan="${colspan}">
        <div class="year-header-inner">
          <button type="button" class="year-toggle-btn" data-year="${y}">${collapsed ? "▶" : "▼"}</button>
          <span>${y}년</span>
        </div>
      </th>`;
  }).join("");

  const monthHeaders = displayKeys.map((dk, idx) => {
    if (dk.startsWith("__year__")) {
      return `<th class="numeric-header year-collapsed-header">합계</th>`;
    }
    return `<th class="numeric-header month-column-cell ${dkcls(dk)}">${formatMonthKey(dk)}</th>`;
  }).join("");

  elements.payables.innerHTML = `
    <div class="panel">
      <div class="payable-controls-row">
        <div class="payable-controls-left">
          <h3 style="margin:0;white-space:nowrap;">미지급 목록</h3>
          ${showBatchButton ? `<button type="button" class="batch-selected-button payment-plan-batch-button">일괄 계획 변경</button>` : ""}
          <div class="rcv-group-chips payable-group-chips" id="payGroupChips" style="margin:0;padding:0;border:none;background:none;">
            <button type="button" class="group-manage-link chip-select-all">전체 선택</button>
            <button type="button" class="group-manage-link chip-clear-all">전체 해제</button>
            ${buildGroupChipsHtml(getOrderedDueGroups(payables), filterState.groups, "pay-chip")}
          </div>
        </div>
        <div style="display:flex;align-items:center;gap:8px;flex-shrink:0;">
          ${(() => { const s = calculateSummary(); return `<div class="tab-mini-stats"><span class="tms-item"><span class="tms-lbl">매입</span><span class="tms-val">${formatNumber(s.totalPayable)}</span></span><span class="tms-sep">|</span><span class="tms-item tms-blue"><span class="tms-lbl">미지급 잔액</span><span class="tms-val">${formatNumber(s.totalUnpaid)}</span></span></div>`; })()}
          <div class="payable-table-actions">
            <button type="button" class="table-action-button subtle" data-action="expand-all">전체 펼치기</button>
            <button type="button" class="table-action-button subtle" data-action="collapse-all">전체 접기</button>
            <button type="button" class="table-action-button subtle" id="completedReportBtn">완료 보고서</button>
          </div>
        </div>
      </div>
      ${payMgrDaysBanner}
      <div class="payment-plan-summary-grid">
        ${paymentPlanSummary.map(item => {
    const encodedKey = encodeURIComponent(item.key);
    const isChecked = paymentPlanUiState.selectedPlanKeys.includes(item.key);
    const cardClass = item.key === "__total__" ? "total" : item.label === "보류" ? "hold" : item.label === "미정" ? "pending" : "";
    return `
            <div class="payment-plan-summary-card ${cardClass} ${isChecked ? "card-selected" : ""}" data-plan-key="${encodedKey}">
              <label class="payment-plan-summary-check" onclick="event.stopPropagation()">
                <input type="checkbox" class="payment-plan-summary-checkbox" data-plan-key="${encodedKey}" ${isChecked ? "checked" : ""} />
              </label>
              <h4>${item.label}</h4>
              <p>${item.amount ? formatNumber(item.amount) : "-"}</p>
              <span>${item.key === "__total__" ? "전체 금액" : `${item.count}건`}</span>
            </div>
          `;
  }).join("")}
      </div>
      <p class="muted" style="margin:2px 0 6px;font-size:0.76rem;">업체마스터 연결: ${matchedVendorCount}개 업체${unmatchedVendorCount > 0 ? ` · 확인 필요: ${unmatchedVendorCount}개` : ""}</p>
      <div class="table-scrollbar-top" id="payablesTopScrollbar"><div class="table-scrollbar-inner" id="payablesTopScrollbarInner"></div></div>
      <div class="table-responsive payables-table-wrap">
        <table>
          <thead>
            <tr>
              <th rowspan="2" class="sticky-col sticky-col-1 payable-header-cell">선택</th>
              <th rowspan="2" class="sticky-col sticky-col-2 payable-header-cell">업체명</th>
              ${yearHeaders}
              <th rowspan="2" class="numeric-header">합계</th>
            </tr>
            <tr>
              ${monthHeaders}
            </tr>
          </thead>
          <tbody>
            ${rows || `<tr><td colspan="${displayKeys.length + 3}" class="empty-state">선택한 거래처에 대한 미지급이 없습니다.</td></tr>`}
          </tbody>
        </table>
      </div>
    </div>
  `;

  document.querySelectorAll(".vendor-memo-btn").forEach(btn => {
    btn.addEventListener("click", e => { e.stopPropagation(); openVendorMemoEditor(btn.dataset.code, btn.dataset.name); });
  });

  document.querySelectorAll(".payable-select-checkbox").forEach(input => {
    input.addEventListener("change", event => {
      const partnerKey = decodeURIComponent(event.target.dataset.partnerKey || "");
      payables.forEach(item => {
        if (getPartnerGroupKey(item) === partnerKey) {
          item.selected = event.target.checked;
        }
      });
      persistPayablesState();
      preserveViewport(() => rerenderAll());
    });
  });

  document.querySelectorAll(".edit-amount-button").forEach(button => {
    button.addEventListener("click", event => {
      const partnerKey = decodeURIComponent(event.currentTarget.dataset.partnerKey || "");
      const monthKey = event.currentTarget.dataset.monthKey;
      openAmountEditor(partnerKey, monthKey, event.currentTarget);
    });
  });

  // 미지급 그룹 칩
  const payGroupChipsEl = document.getElementById("payGroupChips");
  if (payGroupChipsEl) {
    const allPayableGroups = getOrderedDueGroups(payables);
    setupGroupChipEvents(
      payGroupChipsEl, allPayableGroups,
      () => filterState.groups,
      v => { filterState.groups = v; },
      order => { filterState.groupOrder = order; },
      rerenderAll
    );
  }

  document.querySelectorAll(".group-toggle").forEach(button => {
    button.addEventListener("click", event => {
      const groupKey = event.currentTarget.dataset.group;
      payablesGroupState.collapsed[groupKey] = !payablesGroupState.collapsed[groupKey];
      rerenderAll();
    });
  });

  document.querySelectorAll(".table-action-button").forEach(button => {
    button.addEventListener("click", event => {
      const action = event.currentTarget.dataset.action;
      groups.forEach(group => {
        payablesGroupState.collapsed[group.label] = action === "collapse-all";
      });
      rerenderAll();
    });
  });

  document.getElementById("completedReportBtn")?.addEventListener("click", () => {
    openCompletedReportModal();
  });

  document.querySelectorAll(".amount-original-button").forEach(button => {
    button.addEventListener("click", event => {
      event.preventDefault();
      event.stopPropagation();
      const partnerKey = decodeURIComponent(event.currentTarget.dataset.partnerKey || "");
      const targetMonthKey = event.currentTarget.dataset.monthKey;
      payables.forEach(item => {
        if (getPartnerGroupKey(item) === partnerKey && getMonthKey(item) === targetMonthKey) {
          item.decisionAmount = getPayableOutstanding(item);
        }
      });
      payablesUiState.lastEdited = { partnerKey, monthKey: targetMonthKey };
      persistPayablesState();
      preserveViewport(() => rerenderAll());
    });
  });

  // 연도 열 토글
  elements.payables.querySelectorAll(".year-toggle-btn").forEach(btn => {
    btn.addEventListener("click", e => {
      e.stopPropagation();
      const y = btn.dataset.year;
      payablesYearCollapsed[y] = !payablesYearCollapsed[y];
      rerenderAll();
    });
  });

  document.querySelectorAll(".payment-plan-summary-card").forEach(card => {
    card.addEventListener("click", event => {
      if (event.target.closest(".payment-plan-summary-check")) return;
      const planKey = decodeURIComponent(card.dataset.planKey || "");
      openPaymentReportModal(planKey, card);
    });
  });

  document.querySelectorAll(".payment-plan-summary-checkbox").forEach(checkbox => {
    checkbox.addEventListener("change", event => {
      const planKey = decodeURIComponent(event.target.dataset.planKey || "");
      if (event.target.checked) {
        if (!paymentPlanUiState.selectedPlanKeys.includes(planKey)) {
          paymentPlanUiState.selectedPlanKeys.push(planKey);
        }
      } else {
        paymentPlanUiState.selectedPlanKeys = paymentPlanUiState.selectedPlanKeys.filter(k => k !== planKey);
      }
      rerenderAll();
    });
  });

  const batchButton = elements.payables.querySelector(".payment-plan-batch-button");
  if (batchButton) {
    batchButton.addEventListener("click", event => {
      if (paymentPlanUiState.selectedPlanKeys.length > 0) {
        const cardItems = getPayablesForPlanKeys(paymentPlanUiState.selectedPlanKeys, filteredPayables);
        if (!cardItems.length) return;
        const label = paymentPlanUiState.selectedPlanKeys.map(k => formatPlanLabel(k)).join(", ");
        openBatchPlanEditor(label, cardItems, event.currentTarget);
      } else {
        const selectedItems = payables.filter(item => item.selected);
        if (!selectedItems.length) return;
        openBatchPlanEditor("선택 항목", selectedItems, event.currentTarget);
      }
    });
  }

  const topScrollbar = document.getElementById("payablesTopScrollbar");
  const topScrollbarInner = document.getElementById("payablesTopScrollbarInner");
  const tableResponsive = elements.payables.querySelector(".table-responsive");
  const table = tableResponsive?.querySelector("table");
  if (topScrollbar && topScrollbarInner && tableResponsive && table) {
    const updateScrollbarWidth = () => {
      topScrollbarInner.style.width = `${table.scrollWidth}px`;
    };
    updateScrollbarWidth();
    window.requestAnimationFrame(updateScrollbarWidth);
    let syncing = false;
    topScrollbar.addEventListener("scroll", () => {
      if (syncing) return;
      syncing = true;
      tableResponsive.scrollLeft = topScrollbar.scrollLeft;
      syncing = false;
    });
    tableResponsive.addEventListener("scroll", () => {
      if (syncing) return;
      syncing = true;
      topScrollbar.scrollLeft = tableResponsive.scrollLeft;
      syncing = false;
    });
  }

  // 동적 높이: 남은 뷰포트 높이를 table-wrap에 적용 (sticky 헤더 활성화)
  window.requestAnimationFrame(() => {
    const wrap = elements.payables.querySelector(".payables-table-wrap");
    if (wrap) {
      const rect = wrap.getBoundingClientRect();
      const remaining = window.innerHeight - rect.top - 8;
      wrap.style.maxHeight = Math.max(200, remaining) + "px";
    }
  });
}

function showPaymentPlanHistoryDialog(targetItems, partnerKey, monthKey) {
  let combined = [];
  targetItems.forEach(item => {
    const sk = item.sourceKey || "";
    const arr = payablePlanHistories[sk] || [];
    arr.forEach(h => combined.push({ item, row: h }));
  });

  combined.sort((a, b) => new Date(b.row.updated_at || 0).getTime() - new Date(a.row.updated_at || 0).getTime());

  document.querySelector(".history-diff-overlay")?.remove();
  const overlay = document.createElement("div");
  overlay.className = "history-diff-overlay raw-diff-overlay";

  const historyListHtml = combined.length === 0 ? `<div style="padding:20px;text-align:center;color:#666;">과거 원격 저장 이력이 없습니다. (가장 최신의 상태입니다)</div>` :
    combined.map((c, i) => {
      const dt = c.row.updated_at ? new Date(c.row.updated_at).toLocaleString("ko-KR") : "시간 알 수 없음";
      const plan = c.row.payment_plan || "미정";
      const amt = Number(c.row.decision_amount || 0);
      const isLatest = i === 0;
      return `
        <div style="border-bottom:1px solid #eee; padding:15px 0; display:flex; justify-content:space-between; align-items:center;">
          <div>
            <div style="font-size:12px; color:#888;">${dt}</div>
            <div style="font-weight:600; margin-top:4px; font-size:14px;">상태: <span style="color:#2563eb">${plan}</span> / 금액: ${formatNumber(amt)}</div>
            ${c.row.memo ? `<div style="font-size:12px; color:#555; margin-top:4px;">메모: ${c.row.memo}</div>` : ""}
          </div>
          ${!isLatest
          ? `<button type="button" class="btn-restore" style="padding:6px 10px; font-size:13px; cursor:pointer; background:#fff; border:1px solid #ccc; border-radius:4px;" data-index="${i}">이 상태로 복원</button>`
          : `<span style="font-size:13px;color:#10b981;font-weight:600;padding-right:10px;">(현재 상태)</span>`}
        </div>
      `;
    }).join("");

  overlay.innerHTML = `
    <div class="raw-diff-dialog" style="max-height:85vh; overflow-y:auto; width: 450px;">
      <h3 style="margin-top:0; display:flex; align-items:center; gap:8px;">상세 변경 타임라인</h3>
      <p style="font-size:13px; color:#555; margin-bottom:15px;">
        <strong>${targetItems[0]?.name || "알 수 없음"}</strong> (${formatMonthKey(monthKey)}) 건의 상세 변경 이력입니다.
      </p>
      <div style="border-top:2px solid #ddd;">
        ${historyListHtml}
      </div>
      <div style="text-align:right; margin-top:20px;">
        <button type="button" class="btn-close" style="padding:8px 16px; cursor:pointer; background:#e5e7eb; border:none; border-radius:4px; font-weight:600;">닫기</button>
      </div>
    </div>
  `;
  document.body.appendChild(overlay);

  overlay.querySelector(".btn-close").addEventListener("click", () => overlay.remove());

  overlay.querySelectorAll(".btn-restore").forEach(btn => {
    btn.addEventListener("click", (e) => {
      const idx = e.currentTarget.dataset.index;
      const targetState = combined[idx];

      targetItems.forEach(actualItem => {
        if (actualItem.sourceKey === targetState.item.sourceKey) {
          actualItem.paymentPlan = targetState.row.payment_plan || "";
          actualItem.completionStatus = targetState.row.plan_status || (targetState.row.payment_plan === "보류" ? "보류" : targetState.row.payment_plan ? "부분결제" : "미정");
          actualItem.decisionAmount = Number(targetState.row.decision_amount || 0);
          actualItem.memo = targetState.row.memo || "";
          if (targetState.row.paid_override != null && targetState.row.paid_override !== "") {
            actualItem.paidOverride = Number(targetState.row.paid_override);
          } else {
            actualItem.paidOverride = null;
          }
        }
      });

      payablesUiState.lastEdited = { partnerKey, monthKey };
      persistPayablesState();
      overlay.remove();
      rerenderAll();
    });
  });
}

function openAmountEditor(partnerKey, monthKey, triggerElement) {
  const decodedKey = decodeURIComponent(partnerKey || "");
  const monthItems = payables.filter(item => getPartnerGroupKey(item) === decodedKey && getMonthKey(item) === monthKey);
  if (!monthItems.length) return;

  const currentValue = monthItems.reduce((sum, item) => sum + (item.decisionAmount != null ? item.decisionAmount : getPayableOutstanding(item)), 0);
  const totalBalance = monthItems.reduce((sum, item) => sum + getPayableOutstanding(item), 0);
  const partnerName = monthItems[0].name || "";
  const currentPlanValue = monthItems[0].paymentPlan || "";
  const autoPlanValue = getItemAutoPayDate(monthItems[0]);
  const vendorBank = monthItems[0].vendorBank || "";
  const vendorAccount = monthItems[0].vendorAccount || "";
  const vendorAccountHolder = monthItems[0].vendorAccountHolder || "";
  closeCalculator();

  let expression = String(currentValue || 0);
  let replaceOnNextInput = true;
  let calculatorOpen = false;
  let holdPlan = currentPlanValue === "보류";
  let excludePlan = monthItems[0].completionStatus === "제외";

  const overlay = document.createElement("div");
  overlay.className = "calculator-overlay";
  overlay.innerHTML = `
    <div class="editor-popover" role="dialog" aria-modal="true">
      <div class="editor-popover-header">
        <div class="editor-context-title">${partnerName}</div>
        <div class="editor-context-subtitle">${formatMonthKey(monthKey)} 금액 수정</div>
      </div>
      <div class="editor-vendor-meta ${vendorBank || vendorAccount ? "has-vendor" : ""}">
        ${vendorBank || vendorAccount
      ? `<span>${vendorBank || "은행 없음"}</span><span>${vendorAccount || "계좌 없음"}</span><span>${vendorAccountHolder || "예금주 없음"}</span>`
      : `<span>업체마스터에 은행/계좌 정보가 아직 없습니다.</span>`}
      </div>
      <div class="editor-panel">
        <div class="editor-input-row">
          <input
            type="text"
            inputmode="numeric"
            class="editor-input"
            value="${currentValue ? String(currentValue) : ""}"
            autocomplete="off"
            spellcheck="false"
          />
          <button type="button" class="inline-calc-toggle-button" title="계산기 열기" aria-label="계산기 열기">계산기</button>
        </div>
        <div class="editor-preview-label">적용 예정 금액</div>
        <div class="editor-preview-value">${formatNumber(currentValue)}</div>
        <div class="editor-note">금액 입력 중 +를 누르면 000이 붙습니다. Enter는 일반 입력 상태에서 바로 적용됩니다.</div>
        <button type="button" class="editor-original-value-button">원래 금액 ${formatNumber(totalBalance)}</button>
        <div class="editor-plan-row">
          <label class="editor-plan-label">
            결제 예정일
            <input type="date" class="editor-plan-date-input" value="${(holdPlan || excludePlan) ? (autoPlanValue || "") : (/^\d{4}-\d{2}-\d{2}$/.test(currentPlanValue) ? currentPlanValue : (autoPlanValue || ""))}" ${(holdPlan || excludePlan) ? "disabled" : ""} />
          </label>
          <div class="editor-plan-actions">
            <button type="button" class="editor-plan-reset-button">미정</button>
            ${autoPlanValue ? `<button type="button" class="editor-plan-default-button">기본 ${autoPlanValue.replace(/^(\d{4})-(\d{2})-(\d{2})$/, "$2/$3")}</button>` : ""}
            <button type="button" class="editor-plan-hold-button ${holdPlan ? "active" : ""}">보류</button>
            <button type="button" class="editor-plan-exclude-button ${excludePlan ? "active" : ""}">제외</button>
          </div>
        </div>
      </div>
      <div class="editor-actions">
        <button type="button" class="cancel-button">취소</button>
        <button type="button" class="confirm-button">적용</button>
      </div>
    </div>
    <div class="mini-calc-popover hidden">
      <div class="calc-display-wrap">
        <div class="calc-display-label">계산기</div>
        <div class="calc-display">${formatNumber(currentValue)}</div>
      </div>
      <div class="calc-grid calc-grid-simple">
        ${["7", "8", "9", "/", "4", "5", "6", "*", "1", "2", "3", "-", "0", "(", ")", "+", "AC", "⌫", "="].map(value => `
          <button type="button" class="calc-button ${/[/*+\-=]|AC/.test(value) ? "operator" : ""}" data-value="${value}">${value}</button>
        `).join("")}
      </div>
    </div>
  `;

  document.body.appendChild(overlay);
  const editorPopover = overlay.querySelector(".editor-popover");
  const inputField = overlay.querySelector(".editor-input");
  const previewValue = overlay.querySelector(".editor-preview-value");
  const calcPanel = overlay.querySelector(".mini-calc-popover");
  const calcDisplay = overlay.querySelector(".calc-display");
  const calcToggleButton = overlay.querySelector(".inline-calc-toggle-button");
  const originalValueButton = overlay.querySelector(".editor-original-value-button");
  const planDateInput = overlay.querySelector(".editor-plan-date-input");
  const planDefaultButton = overlay.querySelector(".editor-plan-default-button");
  const planHoldButton = overlay.querySelector(".editor-plan-hold-button");

  function syncPlanControls() {
    const planExcludeButton = overlay.querySelector(".editor-plan-exclude-button");
    if (!planDateInput || !planHoldButton || !planExcludeButton) return;
    planDateInput.disabled = holdPlan || excludePlan;
    planHoldButton.classList.toggle("active", holdPlan);
    planExcludeButton.classList.toggle("active", excludePlan);
  }

  function positionPopovers() {
    const rect = triggerElement?.getBoundingClientRect?.() || {
      left: window.innerWidth / 2 - 120,
      top: window.innerHeight / 2 - 40,
      right: window.innerWidth / 2 + 120,
    };
    const editorWidth = editorPopover.offsetWidth || 320;
    const editorHeight = editorPopover.offsetHeight || 320;
    const calcWidth = calcPanel.offsetWidth || 220;
    const calcHeight = calcPanel.offsetHeight || 260;
    const gap = 12;
    const left = Math.min(
      Math.max(16, rect.left),
      Math.max(16, window.innerWidth - editorWidth - calcWidth - gap - 16),
    );
    const top = Math.min(
      Math.max(12, rect.top - 20),
      Math.max(12, window.innerHeight - editorHeight - 12),
    );
    editorPopover.style.left = `${left}px`;
    editorPopover.style.top = `${top}px`;
    calcPanel.style.left = `${Math.max(16, Math.min(window.innerWidth - calcWidth - 16, left + editorWidth + gap))}px`;
    calcPanel.style.top = `${Math.min(top, Math.max(12, window.innerHeight - calcHeight - 12))}px`;
  }

  function sanitizeAmountInput(value) {
    return String(value || "").replace(/[^0-9]/g, "");
  }

  function sanitizeExpression(value) {
    return String(value || "")
      .replace(/,/g, "")
      .replace(/[^0-9+\-*/(). ]/g, "");
  }

  function safeEvaluate(expressionToEvaluate) {
    const sanitized = sanitizeExpression(expressionToEvaluate);
    if (!sanitized.trim()) return 0;
    try {
      const result = Function(`"use strict"; return (${sanitized})`)();
      return Number.isFinite(result) ? result : null;
    } catch {
      return null;
    }
  }

  function updatePreview() {
    const result = calculatorOpen
      ? safeEvaluate(expression)
      : Number(sanitizeAmountInput(inputField.value) || 0);
    const text = result == null ? "계산 불가" : formatNumber(Math.max(0, Math.round(result)));
    previewValue.textContent = text;
    calcDisplay.textContent = text;
  }

  function syncFromInput() {
    expression = calculatorOpen ? (inputField.value.trim() || "0") : (sanitizeAmountInput(inputField.value) || "0");
    updatePreview();
  }

  function insertIntoInput(value) {
    const start = inputField.selectionStart ?? inputField.value.length;
    const end = inputField.selectionEnd ?? inputField.value.length;
    const baseValue = replaceOnNextInput ? "" : inputField.value;
    inputField.value = `${baseValue.slice(0, start)}${value}${baseValue.slice(end)}`;
    const caret = (replaceOnNextInput ? 0 : start) + value.length;
    inputField.setSelectionRange(caret, caret);
    replaceOnNextInput = false;
    syncFromInput();
  }

  function confirmEdit() {
    const evaluated = calculatorOpen ? safeEvaluate(inputField.value) : Number(sanitizeAmountInput(inputField.value) || 0);
    const parsed = evaluated == null ? currentValue : evaluated;
    const newValue = Math.round(Math.min(Math.max(parsed, 0), totalBalance));
    const existingTotal = monthItems.reduce((sum, item) => sum + (item.decisionAmount != null ? item.decisionAmount : getPayableOutstanding(item)), 0);
    monthItems.forEach(item => {
      const current = item.decisionAmount != null ? item.decisionAmount : getPayableOutstanding(item);
      const ratio = existingTotal ? current / existingTotal : 1 / monthItems.length;
      item.decisionAmount = Math.round(newValue * ratio);
      item.selected = item.decisionAmount > 0;
    });
    const remainder = newValue - monthItems.reduce((sum, item) => sum + item.decisionAmount, 0);
    if (remainder !== 0 && monthItems.length > 0) {
      monthItems[0].decisionAmount += remainder;
      monthItems[0].selected = monthItems[0].decisionAmount > 0;
    }
    const nextPlanValue = excludePlan ? "제외" : holdPlan ? "보류" : (planDateInput?.value || "");
    const nextStatus = excludePlan ? "제외" : holdPlan ? "보류" : (nextPlanValue ? "예정" : "미정");

    monthItems.forEach(item => {
      item.paymentPlan = nextPlanValue;
      item.completionStatus = nextStatus;
      if (nextStatus === "제외") {
        item.selected = false;
      }
    });
    payablesUiState.lastEdited = { partnerKey: decodedKey, monthKey };
    persistPayablesState();
    closeCalculator();
    preserveViewport(() => rerenderAll());
  }

  planHoldButton?.addEventListener("click", () => {
    holdPlan = !holdPlan;
    if (holdPlan) excludePlan = false;
    syncPlanControls();
  });

  const planExcludeButton = overlay.querySelector(".editor-plan-exclude-button");
  planExcludeButton?.addEventListener("click", () => {
    excludePlan = !excludePlan;
    if (excludePlan) holdPlan = false;
    syncPlanControls();
  });

  overlay.querySelectorAll(".calc-button").forEach(button => {
    button.addEventListener("click", () => {
      const value = button.dataset.value;
      if (value === "AC") {
        inputField.value = "";
      } else if (value === "⌫") {
        const start = inputField.selectionStart ?? inputField.value.length;
        const end = inputField.selectionEnd ?? inputField.value.length;
        if (start !== end) {
          inputField.value = `${inputField.value.slice(0, start)}${inputField.value.slice(end)}`;
          inputField.setSelectionRange(start, start);
        } else if (start > 0) {
          inputField.value = `${inputField.value.slice(0, start - 1)}${inputField.value.slice(end)}`;
          inputField.setSelectionRange(start - 1, start - 1);
        }
      } else if (value === "=") {
        const result = safeEvaluate(inputField.value);
        if (result != null) {
          inputField.value = String(Math.max(0, Math.round(result)));
        }
      } else {
        insertIntoInput(value);
        inputField.focus();
        return;
      }
      replaceOnNextInput = false;
      syncFromInput();
      inputField.focus();
    });
  });

  inputField.addEventListener("input", event => {
    const previousLength = event.target.value.length;
    const cleaned = calculatorOpen ? sanitizeExpression(event.target.value) : sanitizeAmountInput(event.target.value);
    if (event.target.value !== cleaned) {
      const cursor = event.target.selectionStart ?? cleaned.length;
      const nextCursor = Math.max(0, cursor - (previousLength - cleaned.length));
      event.target.value = cleaned;
      event.target.setSelectionRange(nextCursor, nextCursor);
    }
    replaceOnNextInput = false;
    syncFromInput();
  });

  inputField.addEventListener("keydown", event => {
    if (!calculatorOpen && event.key === "+") {
      event.preventDefault();
      const digits = sanitizeAmountInput(inputField.value);
      inputField.value = digits ? `${digits}000` : "";
      inputField.setSelectionRange(inputField.value.length, inputField.value.length);
      replaceOnNextInput = false;
      syncFromInput();
      return;
    }
    if (event.key === "Enter") {
      event.preventDefault();
      if (calculatorOpen) {
        const result = safeEvaluate(inputField.value);
        if (result != null) {
          inputField.value = String(Math.max(0, Math.round(result)));
        }
      } else {
        syncFromInput();
        confirmEdit();
        return;
      }
      replaceOnNextInput = false;
      syncFromInput();
    }
    if (event.key === "Escape") {
      event.preventDefault();
      closeCalculator();
    }
  });

  inputField.addEventListener("focus", () => {
    inputField.select();
    replaceOnNextInput = true;
  });

  calcToggleButton.addEventListener("click", () => {
    calculatorOpen = !calculatorOpen;
    calcPanel.classList.toggle("hidden", !calculatorOpen);
    calcToggleButton.classList.toggle("active", calculatorOpen);
    if (calculatorOpen) {
      inputField.value = expression || "0";
    } else {
      inputField.value = sanitizeAmountInput(inputField.value);
    }
    replaceOnNextInput = true;
    syncFromInput();
    inputField.focus();
  });

  if (originalValueButton) {
    originalValueButton.addEventListener("click", event => {
      event.preventDefault();
      inputField.value = totalBalance ? String(totalBalance) : "";
      replaceOnNextInput = false;
      syncFromInput();
      inputField.focus();
      const end = inputField.value.length;
      inputField.setSelectionRange(end, end);
    });
  }

  const planResetButton = overlay.querySelector(".editor-plan-reset-button");
  if (planResetButton) {
    planResetButton.addEventListener("click", event => {
      event.preventDefault();
      holdPlan = false;
      excludePlan = false;
      if (planDateInput) planDateInput.value = "";
      syncPlanControls();
    });
  }

  if (planDefaultButton) {
    planDefaultButton.addEventListener("click", event => {
      event.preventDefault();
      holdPlan = false;
      excludePlan = false;
      if (planDateInput) {
        planDateInput.value = autoPlanValue || "";
      }
      syncPlanControls();
    });
  }
  // planHoldButton 리스너는 위(5397)에서 이미 등록됨 — 중복 등록 제거

  overlay.querySelector(".cancel-button").addEventListener("click", closeCalculator);
  overlay.querySelector(".confirm-button").addEventListener("click", confirmEdit);
  overlay.addEventListener("mousedown", event => {
    if (!editorPopover.contains(event.target) && !calcPanel.contains(event.target)) {
      closeCalculator();
    }
  });

  const reposition = () => positionPopovers();
  window.addEventListener("resize", reposition);
  window.addEventListener("scroll", reposition, true);
  overlay.cleanup = () => {
    window.removeEventListener("resize", reposition);
    window.removeEventListener("scroll", reposition, true);
  };

  positionPopovers();
  updatePreview();
  syncPlanControls();
  inputField.focus();
  inputField.select();
}

function openBatchPlanEditor(planKey, targetItems, triggerElement) {
  if (!targetItems.length) return;
  closeBatchPlanEditor();

  const firstDate = targetItems.find(item => /^\d{4}-\d{2}-\d{2}$/.test(item.paymentPlan || ""))?.paymentPlan || "";
  let holdPlan = targetItems.every(item => item.paymentPlan === "보류");
  let excludePlan = targetItems.every(item => item.completionStatus === "제외");

  const overlay = document.createElement("div");
  overlay.className = "batch-plan-overlay";
  overlay.innerHTML = `
    <div class="batch-plan-popover" role="dialog" aria-modal="true">
      <div class="batch-plan-title">${planKey === "__total__" ? "전체 예정 변경" : planKey === "선택 항목" ? `선택 항목 일괄 변경` : `${formatPlanLabel(planKey)} 일괄 변경`}</div>
      <p class="batch-plan-note">${targetItems.length}건에 같은 결제 계획을 적용합니다.</p>
      <label class="editor-plan-label">
        결제 예정일
        <input type="date" class="editor-plan-date-input" value="${(holdPlan || excludePlan) ? "" : firstDate}" ${(holdPlan || excludePlan) ? "disabled" : ""} />
      </label>
      <div class="editor-plan-actions">
        <button type="button" class="editor-plan-reset-button">미정</button>
        <button type="button" class="editor-plan-hold-button ${holdPlan ? "active" : ""}">보류</button>
        <button type="button" class="editor-plan-exclude-button ${excludePlan ? "active" : ""}">제외</button>
      </div>
      <div class="editor-actions compact">
        <button type="button" class="cancel-button">닫기</button>
        <button type="button" class="confirm-button">적용</button>
      </div>
    </div>
  `;

  document.body.appendChild(overlay);
  const popover = overlay.querySelector(".batch-plan-popover");
  const dateInput = overlay.querySelector(".editor-plan-date-input");
  const holdButton = overlay.querySelector(".editor-plan-hold-button");
  const excludeButton = overlay.querySelector(".editor-plan-exclude-button");
  const resetButton = overlay.querySelector(".editor-plan-reset-button");

  function syncState() {
    dateInput.disabled = holdPlan || excludePlan;
    holdButton.classList.toggle("active", holdPlan);
    excludeButton.classList.toggle("active", excludePlan);
  }

  function positionPopover() {
    const rect = triggerElement?.getBoundingClientRect?.() || { left: 24, top: 24 };
    const width = popover.offsetWidth || 260;
    const height = popover.offsetHeight || 220;
    const left = Math.min(Math.max(12, rect.left), Math.max(12, window.innerWidth - width - 12));
    const top = Math.min(Math.max(12, rect.bottom + 8), Math.max(12, window.innerHeight - height - 12));
    popover.style.left = `${left}px`;
    popover.style.top = `${top}px`;
  }

  function applyPlan() {
    const nextStatus = excludePlan ? "제외" : holdPlan ? "보류" : "미정";
    const nextPlanValue = excludePlan ? "제외" : holdPlan ? "보류" : (dateInput.value || "");

    targetItems.forEach(item => {
      item.paymentPlan = nextPlanValue;
      item.completionStatus = nextStatus;
      if (nextStatus === "제외") {
        item.selected = false; // 제외 시 선택 해제
      }
    });
    persistPayablesState();
    closeBatchPlanEditor();
    preserveViewport(() => rerenderAll());
  }

  holdButton.addEventListener("click", () => {
    holdPlan = !holdPlan;
    if (holdPlan) excludePlan = false;
    syncState();
  });

  excludeButton.addEventListener("click", () => {
    excludePlan = !excludePlan;
    if (excludePlan) holdPlan = false;
    syncState();
  });

  resetButton.addEventListener("click", () => {
    holdPlan = false;
    excludePlan = false;
    dateInput.value = "";
    syncState();
  });

  overlay.querySelector(".cancel-button").addEventListener("click", closeBatchPlanEditor);
  overlay.querySelector(".confirm-button").addEventListener("click", applyPlan);
  overlay.addEventListener("mousedown", event => {
    if (!popover.contains(event.target)) {
      closeBatchPlanEditor();
    }
  });

  const reposition = () => positionPopover();
  window.addEventListener("resize", reposition);
  window.addEventListener("scroll", reposition, true);
  overlay.cleanup = () => {
    window.removeEventListener("resize", reposition);
    window.removeEventListener("scroll", reposition, true);
  };

  syncState();
  positionPopover();
}

function closeCalculator() {
  const existing = document.querySelector(".calculator-overlay");
  if (existing) {
    if (typeof existing.cleanup === "function") {
      existing.cleanup();
    }
    existing.remove();
  }
}

function closeBatchPlanEditor() {
  const existing = document.querySelector(".batch-plan-overlay");
  if (existing) {
    if (typeof existing.cleanup === "function") {
      existing.cleanup();
    }
    existing.remove();
  }
}

function renderFixedExpenses() {
  const filteredFixed = getFilteredItems(fixedExpenses, "fixed");

  // 전체 날짜(년-월-일) 기준으로 정렬
  const sortedFixed = [...filteredFixed].sort((a, b) => {
    const da = (a.year || 0) * 10000 + (a.month || 0) * 100 + (a.day || 0);
    const db = (b.year || 0) * 10000 + (b.month || 0) * 100 + (b.day || 0);
    return da - db;
  });

  const selectedYear = filterState.year || new Date().getFullYear();
  const selectedMonth = filterState.month || (new Date().getMonth() + 1);

  const today = new Date();
  const todayKey = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, "0")}-${String(today.getDate()).padStart(2, "0")}`;

  // 추출된 유니크 은행 목록 정렬 (가나다 순)
  const uniqueBanks = [...new Set(sortedFixed.map(item => item.bank).filter(Boolean))].sort();

  // 전체 날짜(YYYY-MM-DD) 기준으로 그룹화
  const dateGroups = {};
  const dateOrder = [];
  sortedFixed.forEach(item => {
    const key = `${item.year || 0}-${String(item.month || 0).padStart(2, "0")}-${String(item.day || 0).padStart(2, "0")}`;
    if (!dateGroups[key]) {
      dateGroups[key] = { items: [], month: item.month, day: item.day, key };
      dateOrder.push(key);
    }
    dateGroups[key].items.push(item);
  });

  // 결제금액(paidAmount) 우선, 없으면 금액(amount) fallback
  const effectiveAmt = i => i.paidAmount || i.amount || 0;
  const grandTotal = sortedFixed.reduce((s, i) => s + effectiveAmt(i), 0);
  const grandBankTotals = {};
  uniqueBanks.forEach(b => grandBankTotals[b] = 0);
  sortedFixed.forEach(item => {
    if (grandBankTotals[item.bank] !== undefined) {
      grandBankTotals[item.bank] += effectiveAmt(item);
    }
  });

  const bankColWidth = uniqueBanks.length > 3 ? 100 : 120; // 은행 수에 따라 너비 자동조절

  // 날짜별 그룹 행 생성
  const groupRows = dateOrder.map(key => {
    const { items, month, day } = dateGroups[key];
    const dayTotal = items.reduce((s, i) => s + effectiveAmt(i), 0);
    const groupId = `fixed-date-${key.replace(/\W/g, "")}`;

    // 해당 날짜의 은행별 합계 계산
    const dayBankTotals = {};
    uniqueBanks.forEach(b => dayBankTotals[b] = 0);
    items.forEach(item => {
      if (dayBankTotals[item.bank] !== undefined) {
        dayBankTotals[item.bank] += effectiveAmt(item);
      }
    });

    const itemRows = items.map((item, idx) => `
      <tr class="fx-item-row" data-group="${groupId}">
        <td class="fx-item-check-cell" style="padding:10px 8px;"></td>
        <td class="fx-item-title" style="padding:10px 12px 10px 24px; color:#475569; font-size:13.5px;">
          <span class="fx-item-dot" style="margin-right:8px;color:#cbd5e1;font-size:12px;">↳</span>
          ${item.title}
        </td>
        ${uniqueBanks.map(b => `
          <td style="text-align:right;font-size:13.5px;padding:10px 12px;">
            ${b === item.bank && effectiveAmt(item) ? `<span style="color:#0f172a;">${formatNumber(effectiveAmt(item))}</span>` : ''}
          </td>
        `).join("")}
        <td style="text-align:right;font-weight:600;color:#1e293b;padding:10px 12px;font-size:14px;">
          ${effectiveAmt(item) ? formatNumber(effectiveAmt(item)) : ''}
        </td>
      </tr>
    `).join("");

    const isToday = key === todayKey;
    const todayBadge = isToday ? `<span style="display:inline-flex;align-items:center;background:#fef2f2;color:#ef4444;border:1px solid #fca5a5;font-size:11px;font-weight:700;padding:2px 6px;border-radius:12px;margin-left:6px;">D-Day</span>` : "";

    return `
      <tr class="fx-date-header ${isToday ? 'fx-today-header' : ''}" data-group="${groupId}" style="${isToday ? 'background-color:#fff1f2;' : ''}">
        <td class="fx-header-check" style="text-align:center;padding:12px 8px;">
          <input type="checkbox" class="fixed-day-check fx-checkbox" data-total="${dayTotal}" checked style="cursor:pointer;width:16px;height:16px;accent-color:#2563eb;">
        </td>
        <td class="fx-header-title" style="padding-top:14px;padding-bottom:14px;">
          <span class="fx-chevron fixed-toggle-btn" style="display:inline-block;transition:all 0.2s;margin-right:8px;font-size:12px;color:#94a3b8;">▼</span>
          <span class="fx-date-badge" style="background:${isToday ? '#ef4444' : '#eff6ff'};color:${isToday ? '#ffffff' : '#1d4ed8'};padding:4px 10px;border-radius:16px;font-size:14px;margin-right:8px;font-weight:700;">${month}/${day}</span>
          <span class="fx-count-pill" style="font-size:12px;color:#3b82f6;background:#dbeafe;padding:3px 8px;border-radius:12px;font-weight:600;">${items.length}건</span>
          ${todayBadge}
        </td>
        ${uniqueBanks.map(b => `
          <td style="text-align:right;font-weight:600;color:#334155;font-size:14px;padding:0 12px;">
            ${dayBankTotals[b] > 0 ? formatNumber(dayBankTotals[b]) : ''}
          </td>
        `).join("")}
        <td class="fx-header-amount" style="text-align:right;font-weight:800;color:#2563eb;font-size:15px;padding:0 12px;">
          ${formatNumber(dayTotal)}
        </td>
      </tr>
      ${itemRows}
    `;
  }).join("");

  elements.fixed.innerHTML = `
    <style>
      .fx-table td, .fx-table th {
        border-bottom: 1px solid #cbd5e1 !important;
      }
      .fx-table tr:last-child td {
        border-bottom: none !important;
      }
      .fx-item-row {
        transition: background-color 0.15s ease;
      }
      .fx-item-row:hover {
        background-color: #f8fafc;
      }
      .fx-date-header {
        background-color: #ffffff;
        cursor: pointer;
        transition: background-color 0.15s ease;
      }
      .fx-date-header:hover {
        background-color: #f1f5f9;
      }
    </style>
    <div class="fx-panel" style="background:#fff;border-radius:16px;box-shadow:0 4px 6px -1px rgba(0,0,0,0.05), 0 2px 4px -2px rgba(0,0,0,0.05);overflow:hidden;border:1px solid #cbd5e1;">
      <div class="fx-panel-header" style="display:flex;justify-content:space-between;align-items:center;padding:20px 24px;border-bottom:1px solid #cbd5e1;background:#f8fafc;">
        <div class="fx-panel-title-group">
          <h3 class="fx-panel-title" style="margin:0;font-size:18px;font-weight:800;color:#0f172a;letter-spacing:-0.02em;">${selectedYear}년 ${selectedMonth}월 고정지출</h3>
          <p class="fx-panel-subtitle" style="margin:6px 0 0;font-size:13.5px;color:#64748b;">날짜별 결제 내역 · ${dateOrder.length}일 · ${sortedFixed.length}건</p>
        </div>
        <div class="fx-panel-controls" style="display:flex;align-items:center;gap:16px;">
          <div class="fx-btn-group" style="display:flex;gap:8px;flex-wrap:wrap;">
            <button id="fixedPasteToggle" class="fx-ctrl-btn" style="padding:6px 12px;font-size:13px;font-weight:600;background:#eff6ff;border:1px solid #bfdbfe;border-radius:6px;cursor:pointer;color:#1d4ed8;">붙여넣기</button>
            <button id="fixedClearBtn" class="fx-ctrl-btn" style="padding:6px 12px;font-size:13px;font-weight:600;background:#fff;border:1px solid #cbd5e1;border-radius:6px;cursor:pointer;color:#ef4444;">초기화</button>
            <button id="fixedSelectAll" class="fx-ctrl-btn" style="padding:6px 12px;font-size:13px;font-weight:600;background:#fff;border:1px solid #cbd5e1;border-radius:6px;cursor:pointer;color:#334155;">전체선택</button>
            <button id="fixedDeselectAll" class="fx-ctrl-btn" style="padding:6px 12px;font-size:13px;font-weight:600;background:#fff;border:1px solid #cbd5e1;border-radius:6px;cursor:pointer;color:#334155;">전체해제</button>
            <button id="fixedExpandAll" class="fx-ctrl-btn" style="padding:6px 12px;font-size:13px;font-weight:600;background:#fff;border:1px solid #cbd5e1;border-radius:6px;cursor:pointer;color:#334155;">전체 펼치기</button>
            <button id="fixedCollapseAll" class="fx-ctrl-btn" style="padding:6px 12px;font-size:13px;font-weight:600;background:#fff;border:1px solid #cbd5e1;border-radius:6px;cursor:pointer;color:#334155;">전체 접기</button>
          </div>
          <div class="fx-total-chip" style="background:#eff6ff;padding:8px 16px;border-radius:10px;border:1px solid #bfdbfe;display:flex;align-items:center;">
            <span class="fx-total-label" style="font-size:13px;color:#1e40af;margin-right:8px;font-weight:700;">선택 합계</span>
            <strong id="fixedCheckedTotal" class="fx-total-value" style="font-size:17px;color:#1d4ed8;letter-spacing:-0.01em;">${formatNumber(grandTotal)}</strong>
          </div>
        </div>
      </div>
      <!-- 붙여넣기 영역 -->
      <div id="fixedPasteArea" style="display:none;padding:12px 24px;background:#f8fafc;border-bottom:1px solid #e2e8f0;">
        <div style="font-size:12px;color:#64748b;margin-bottom:6px;">헤더: 연도 / 월 / 내용 / 일 / 날짜 / 금액 / 은행 / 분류 / 실결제일 / 결제금액</div>
        <textarea id="fixedPasteTextarea" style="width:100%;height:80px;font-size:12px;border:1px solid #cbd5e1;border-radius:6px;padding:6px;box-sizing:border-box;resize:vertical;" placeholder="엑셀에서 복사(Ctrl+C) 후 여기에 붙여넣기(Ctrl+V)"></textarea>
        <div style="display:flex;gap:8px;margin-top:6px;">
          <button id="fixedPasteApply" style="padding:5px 14px;font-size:13px;font-weight:600;background:#2563eb;color:#fff;border:none;border-radius:6px;cursor:pointer;">✔ 적용</button>
          <button id="fixedPasteCancel" style="padding:5px 14px;font-size:13px;font-weight:600;background:#e5e7eb;color:#374151;border:none;border-radius:6px;cursor:pointer;">취소</button>
        </div>
      </div>
      <div class="fx-table-wrap" style="overflow-x:auto;">
        <table class="fx-table" style="width:100%;border-collapse:separate;border-spacing:0;min-width:700px;">
          <thead class="fx-thead" style="background:#f8fafc;">
            <tr>
              <th style="width:40px;padding:14px 8px;"></th>
              <th style="text-align:left;padding:14px 12px;font-weight:700;color:#1e293b;font-size:13px;">내용</th>
              ${uniqueBanks.map(b => `
                <th style="text-align:right;width:${bankColWidth}px;padding:14px 12px;font-weight:700;color:#475569;font-size:13px;">${b}</th>
              `).join("")}
              <th style="text-align:right;width:120px;padding:14px 12px;font-weight:800;color:#0f172a;font-size:13px;">합계</th>
            </tr>
          </thead>
          <tbody>
            ${dateOrder.length ? groupRows : `
            <tr><td colspan="${3 + uniqueBanks.length}" style="text-align:center;padding:60px 0;color:#94a3b8;font-size:14px;">
              📭 해당 월의 고정지출 데이터가 없습니다.
            </td></tr>`}
          </tbody>
          ${dateOrder.length ? `
          <tfoot style="background:#f8fafc;border-top:2px solid #cbd5e1;">
            <tr class="fx-footer-row">
              <td></td>
              <td style="text-align:right;padding:16px 12px;font-weight:700;font-size:14px;color:#334155;">
                총 합계 (전체)
              </td>
              ${uniqueBanks.map(b => `
                <td style="text-align:right;padding:16px 12px;font-weight:600;color:#475569;">
                  ${grandBankTotals[b] > 0 ? formatNumber(grandBankTotals[b]) : ''}
                </td>
              `).join("")}
              <td style="text-align:right;padding:16px 12px;font-weight:800;font-size:16px;color:#1d4ed8;">
                ${formatNumber(grandTotal)}
              </td>
            </tr>
          </tfoot>` : ""}
        </table>
      </div>
    </div>
  `;

  // ── 이벤트 바인딩 ──────────────────────────────────────────

  const setAllCollapsed = (collapsed) => {
    elements.fixed.querySelectorAll(".fx-date-header").forEach(hdr => {
      const gid = hdr.dataset.group;
      const chevron = hdr.querySelector(".fx-chevron");
      elements.fixed.querySelectorAll(`[data-group="${gid}"].fx-item-row, [data-group="${gid}"].fx-summary-row`)
        .forEach(r => r.style.display = collapsed ? "none" : "");
      if (chevron) chevron.style.transform = collapsed ? "rotate(-90deg)" : "";
    });
  };

  document.getElementById("fixedExpandAll")?.addEventListener("click", () => setAllCollapsed(false));
  document.getElementById("fixedCollapseAll")?.addEventListener("click", () => setAllCollapsed(true));

  // 붙여넣기 토글
  document.getElementById("fixedPasteToggle")?.addEventListener("click", () => {
    const area = document.getElementById("fixedPasteArea");
    if (!area) return;
    const hidden = area.style.display === "none";
    area.style.display = hidden ? "" : "none";
    if (hidden) document.getElementById("fixedPasteTextarea")?.focus();
  });
  document.getElementById("fixedPasteCancel")?.addEventListener("click", () => {
    const area = document.getElementById("fixedPasteArea");
    if (area) area.style.display = "none";
  });
  document.getElementById("fixedPasteApply")?.addEventListener("click", () => {
    const ta = document.getElementById("fixedPasteTextarea");
    if (ta) applyFixedPaste(ta.value.trim());
  });
  // 붙여넣기 후 자동 적용
  document.getElementById("fixedPasteTextarea")?.addEventListener("paste", e => {
    setTimeout(() => {
      const ta = document.getElementById("fixedPasteTextarea");
      if (ta) applyFixedPaste(ta.value.trim());
    }, 50);
  });

  // 초기화
  document.getElementById("fixedClearBtn")?.addEventListener("click", () => {
    if (!confirm("고정지출 데이터를 초기화하시겠습니까?")) return;
    fixedExpenses = [];
    saveFixedLocal();
    renderFixedExpenses();
  });

  // 기본 전체 접기
  setAllCollapsed(true);

  elements.fixed.querySelectorAll(".fx-date-header").forEach(hdr => {
    hdr.addEventListener("click", e => {
      if (e.target.classList.contains("fx-checkbox")) return;
      const gid = hdr.dataset.group;
      const chevron = hdr.querySelector(".fx-chevron");
      const collapsed = chevron?.style.transform === "rotate(-90deg)";
      elements.fixed.querySelectorAll(`[data-group="${gid}"].fx-item-row, [data-group="${gid}"].fx-summary-row`)
        .forEach(r => r.style.display = collapsed ? "" : "none");
      if (chevron) chevron.style.transform = collapsed ? "" : "rotate(-90deg)";
    });
  });

  const updateCheckedTotal = () => {
    let sum = 0;
    elements.fixed.querySelectorAll(".fixed-day-check").forEach(cb => {
      if (cb.checked) sum += Number(cb.dataset.total) || 0;
    });
    const el = document.getElementById("fixedCheckedTotal");
    if (el) el.textContent = formatNumber(sum);
  };

  document.getElementById("fixedSelectAll")?.addEventListener("click", () => {
    elements.fixed.querySelectorAll(".fixed-day-check").forEach(cb => cb.checked = true);
    updateCheckedTotal();
  });

  document.getElementById("fixedDeselectAll")?.addEventListener("click", () => {
    elements.fixed.querySelectorAll(".fixed-day-check").forEach(cb => cb.checked = false);
    updateCheckedTotal();
  });

  elements.fixed.querySelectorAll(".fixed-day-check").forEach(cb => {
    cb.addEventListener("change", updateCheckedTotal);
  });
}

// ── 엠오토 탭 ───────────────────────────────────────────────

function createDefaultMautoData() {
  return {
    funds: MAUTO_FIXED_ACCOUNTS.map(a => ({ ...a, amount: 0 })),
    receivables: [],
    payables: [],
    fixed: [],
  };
}

function normalizeMautoYear(value) {
  const n = Math.round(parseNum(value));
  if (!n) return 0;
  return n > 0 && n < 100 ? 2000 + n : n;
}
function normalizeMautoMonth(value) {
  const n = Math.round(parseNum(value));
  return n >= 1 && n <= 12 ? n : 0;
}
function normalizeMautoDay(value) {
  const n = Math.round(parseNum(value));
  return n >= 1 && n <= 31 ? n : 0;
}

function normalizeMautoData(data) {
  const d = data && typeof data === "object" ? data : {};
  const savedFunds = Array.isArray(d.funds) ? d.funds.filter(r => r && typeof r === "object") : [];
  const funds = MAUTO_FIXED_ACCOUNTS.map((account, idx) => {
    const saved = savedFunds.find(r =>
      r.accountNo === account.accountNo ||
      r.bankAccount === account.bankAccount ||
      String(r.bankAccount || "").includes(account.accountNo)
    ) || savedFunds[idx] || {};
    return { ...account, amount: parseNum(saved.amount ?? saved.balance ?? saved.value ?? 0) };
  });
  return {
    funds,
    receivables: Array.isArray(d.receivables) ? d.receivables : [],
    payables: Array.isArray(d.payables) ? d.payables : [],
    fixed: Array.isArray(d.fixed) ? d.fixed : [],
  };
}

function saveMautoDataLocal() {
  try { localStorage.setItem(MAUTO_LOCAL_KEY, JSON.stringify(mautoData)); }
  catch (e) { console.warn("[엠오토] 저장 실패:", e); }
  _scheduleMautoRemoteSave();
}

function loadMautoDataLocal() {
  try {
    const raw = localStorage.getItem(MAUTO_LOCAL_KEY);
    mautoData = raw ? normalizeMautoData(JSON.parse(raw)) : createDefaultMautoData();
  } catch (e) {
    mautoData = createDefaultMautoData();
  }
}

// ── 엠오토 Google Sheets 원격 저장/로드 ──────────────────────
let _mautoSaveTimer = null;
function _scheduleMautoRemoteSave() {
  if (!SHEET_APP_SCRIPT_URL) return;
  clearTimeout(_mautoSaveTimer);
  _mautoSaveTimer = setTimeout(async () => {
    try {
      // ⚠️ taxInvoices는 절대 포함하지 말 것! 엠오토_json은 한 셀(A1)에 저장되는데
      // 세금계산서 146건까지 넣으면 10만 자를 넘겨 Google Sheets 셀 한도(5만 자) 초과 →
      // 저장 전체가 실패해 가용자금 등이 동기화 안 됨. 세금계산서는 엠오토_세금계산서 시트에 별도 저장됨.
      await postSheetWebApp("saveMautoData", {
        data: {
          ...mautoData,
          excludeRcv: mautoExcludeVendorsRcv,
          excludePay: mautoExcludeVendorsPay,
          fixedChecked: mautoFixedChecked,
          fixedAmountOverrides: mautoFixedAmountOverrides,
        }
      });
    } catch (e) {
      console.warn("[엠오토] 구글시트 저장 실패:", e);
    }
  }, 1500);
}

async function loadMautoDataRemote() {
  if (!SHEET_APP_SCRIPT_URL) return null;
  const url = new URL(SHEET_APP_SCRIPT_URL);
  const token = getApiToken();
  if (token) url.searchParams.set("token", token);
  url.searchParams.set("action", "getMautoData");
  const resp = await fetch(url.toString());
  if (!resp.ok) throw new Error(`엠오토 원격 로드 실패: ${resp.status}`);
  const body = await resp.json();
  console.log("[엠오토] raw 응답 키:", body ? Object.keys(body) : body, "/ 전체:", JSON.stringify(body).slice(0, 400));
  if (!body) return null;
  // { data: {...} } 형식
  if (body.data && typeof body.data === "object" && !Array.isArray(body.data)) return body.data;
  // { rows: {...} } 형식
  if (body.rows && typeof body.rows === "object" && !Array.isArray(body.rows)) return body.rows;
  // 직접 객체 형식 (funds 키가 있으면 유효한 mauto 데이터)
  if (body.funds !== undefined) return body;
  console.warn("[엠오토] 원격 응답 형식 미인식:", JSON.stringify(body).slice(0, 200));
  return null;
}

function splitMautoPasteRows(text) {
  return String(text || "").trim().split(/\r?\n/)
    .map(line => line.split("\t").map(cell => cell.trim()))
    .filter(row => row.some(Boolean));
}

function normalizeMautoHeader(value) {
  return normalizeKey(value).replace(/[()]/g, "");
}

function parseMautoTablePaste(text, defs) {
  const rows = splitMautoPasteRows(text);
  if (!rows.length) return [];
  const firstRowHeaders = rows[0].map(normalizeMautoHeader);
  const matchCount = defs.reduce((count, def) => {
    const aliases = (def.aliases || [def.label || def.key]).map(normalizeMautoHeader);
    return count + (aliases.some(a => firstRowHeaders.includes(a)) ? 1 : 0);
  }, 0);
  const hasHeader = matchCount >= Math.min(2, defs.length);
  const dataRows = hasHeader ? rows.slice(1) : rows;
  const findIdx = (def, fallback) => {
    if (!hasHeader) return fallback;
    const aliases = (def.aliases || [def.label || def.key]).map(normalizeMautoHeader);
    return firstRowHeaders.findIndex(h => aliases.includes(h));
  };
  return dataRows.map(cols => {
    const row = {};
    defs.forEach((def, idx) => {
      const ci = findIdx(def, idx);
      const raw = ci >= 0 ? (cols[ci] || "") : "";
      row[def.key] = def.numeric ? parseNum(raw) : raw;
    });
    return row;
  }).filter(row => Object.values(row).some(v => v !== "" && v !== 0));
}

function parseMautoFundsPaste(text) {
  const rows = splitMautoPasteRows(text);
  const values = [];
  rows.forEach(cols => {
    const candidates = cols.flatMap(cell => {
      const matches = String(cell || "").match(/-?[\d,]+(?:\.\d+)?/g) || [];
      return matches.map(m => parseNum(m)).filter(Number.isFinite);
    });
    if (candidates.length) values.push(candidates[candidates.length - 1]);
  });
  if (!values.length) return null;
  return MAUTO_FIXED_ACCOUNTS.map((account, idx) => ({ ...account, amount: values[idx] || 0 }));
}

function parseMautoAccountingPaste(text, kind) {
  const isReceivable = kind === "receivables";
  const defs = [
    { key: "year", aliases: ["작성연도", "연도", "년"], numeric: false },
    { key: "month", aliases: ["작성", "작성월", "월"], numeric: false },
    { key: "company", aliases: ["상호", "거래처", "거래처명", "업체명"], numeric: false },
    { key: "total", aliases: [isReceivable ? "매출합계" : "매입합계", "합계"], numeric: true },
    { key: "supply", aliases: [isReceivable ? "매출공급가액" : "매입공급가액", "공급가액"], numeric: true },
    { key: "tax", aliases: [isReceivable ? "매출세액" : "매입세액", "세액"], numeric: true },
    { key: "inout", aliases: ["입출금", "수금", "지급", "입금", "출금"], numeric: true },
    { key: "balance", aliases: ["잔액", "잔 액"], numeric: true },
  ];
  return parseMautoTablePaste(text, defs).map(row => ({
    year: normalizeMautoYear(row.year),
    month: normalizeMautoMonth(row.month),
    company: String(row.company || "").trim(),
    total: Number(row.total || 0),
    supply: Number(row.supply || 0),
    tax: Number(row.tax || 0),
    inout: Number(row.inout || 0),
    balance: Number(row.balance || 0),
  })).filter(row => row.company || row.total || row.balance);
}

function parseMautoDateParts(value) {
  if (!value) return null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return { year: value.getFullYear(), month: value.getMonth() + 1, day: value.getDate() };
  }
  const raw = String(value).trim();
  const gviz = raw.match(/^Date\((\d+),(\d+),(\d+)\)/);
  if (gviz) return { year: Number(gviz[1]), month: Number(gviz[2]) + 1, day: Number(gviz[3]) };
  const full = raw.match(/(\d{4})[-./년\s]+(\d{1,2})[-./월\s]+(\d{1,2})/);
  if (full) return { year: Number(full[1]), month: Number(full[2]), day: Number(full[3]) };
  const serial = Number(raw);
  if (Number.isFinite(serial) && serial > 20000 && serial < 80000) {
    const d = new Date(Math.round((serial - 25569) * 86400 * 1000));
    return { year: d.getUTCFullYear(), month: d.getUTCMonth() + 1, day: d.getUTCDate() };
  }
  return null;
}

function parseMautoFixedPaste(text) {
  const defs = [
    { key: "year", aliases: ["연도", "년"], numeric: false },
    { key: "month", aliases: ["월"], numeric: false },
    { key: "title", aliases: ["내용", "적요"], numeric: false },
    { key: "day", aliases: ["일", "일자"], numeric: false },
    { key: "date", aliases: ["날짜", "일시", "date"], numeric: false },
    { key: "amount", aliases: ["금액"], numeric: true },
    { key: "bank", aliases: ["은행", "bank"], numeric: false },
    { key: "category", aliases: ["분류", "구분"], numeric: false },
    { key: "actualPayDate", aliases: ["실결제일"], numeric: false },
    { key: "paidAmount", aliases: ["결제금액"], numeric: true },
  ];
  return parseMautoTablePaste(text, defs).map(row => {
    const dp = parseMautoDateParts(row.date);
    // 날짜 컬럼 우선, 없으면 연도/월/일 컬럼으로 보완
    const year = dp?.year || normalizeMautoYear(row.year) || 0;
    const month = dp?.month || normalizeMautoMonth(row.month) || 0;
    const day = dp?.day || normalizeMautoDay(row.day) || 0;
    const rawDate = String(row.date || "").trim();
    const date = year && month && day
      ? `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`
      : rawDate || "";
    return {
      year, month, day, date,
      title: String(row.title || "").trim(),
      amount: Number(row.amount || 0),
      bank: String(row.bank || "").trim() || "미지정",
      category: String(row.category || "").trim(),
      actualPayDate: String(row.actualPayDate || "").trim(),
      paidAmount: Number(row.paidAmount || 0),
    };
  }).filter(row => row.title || row.amount);
}

function mautoNumericCell(value, extraClass = "") {
  return `<td class="mauto-num ${extraClass}">${formatNumber(value)}</td>`;
}

function sumMautoRows(rows) {
  return rows.reduce((acc, row) => {
    acc.total += Number(row.total || 0);
    acc.inout += Number(row.inout || 0);
    acc.balance += Number(row.balance || 0);
    acc.amount += Number(row.amount || 0);
    return acc;
  }, { total: 0, inout: 0, balance: 0, amount: 0 });
}

function renderMautoFundsTable() {
  const rows = normalizeMautoData(mautoData).funds;
  const total = rows.reduce((s, r) => s + Number(r.amount || 0), 0);
  const body = rows.map((r, i) => `
    <tr>
      <td>${escapeHtml(r.bankAccount)}</td>
      <td class="mauto-num"><input type="text" class="mauto-inline-input" data-mauto-acc-idx="${i}" value="${formatNumber(r.amount || 0)}" /></td>
    </tr>`).join("");
  return `<div class="mauto-table-wrap">
    <table class="mauto-table mauto-compact-table" id="mauto-funds-table">
      <thead><tr><th>은행(계좌)</th><th>가용자금</th></tr></thead>
      <tbody>${body}</tbody>
      <tfoot><tr><td>합계</td><td class="mauto-num" id="mauto-funds-total">${formatNumber(total)}</td></tr></tfoot>
    </table>
  </div>`;
}

function renderMautoAccountingTable(rows, kind) {
  const isReceivable = kind === "receivables";
  const totalLabel = isReceivable ? "매출합계" : "매입합계";
  const sorted = [...(rows || [])].sort((a, b) =>
    (a.year || 9999) - (b.year || 9999) ||
    (a.month || 99) - (b.month || 99) ||
    String(a.company || "").localeCompare(String(b.company || ""), "ko")
  );

  if (!sorted.length) {
    return `<div class="mauto-table-wrap">
      <table class="mauto-table">
        <thead><tr><th>작성연도</th><th>작성월</th><th>상호</th><th class="mauto-num">${totalLabel}</th><th class="mauto-num">입출금</th><th class="mauto-num">잔액</th></tr></thead>
        <tbody><tr><td colspan="6" class="mauto-empty">데이터 없음</td></tr></tbody>
      </table>
    </div>`;
  }

  const years = new Map();
  sorted.forEach(row => {
    const yk = row.year || "연도 없음";
    const mk = row.month || "월 없음";
    if (!years.has(yk)) years.set(yk, new Map());
    if (!years.get(yk).has(mk)) years.get(yk).set(mk, []);
    years.get(yk).get(mk).push(row);
  });

  const body = [...years.entries()].map(([year, monthMap]) => {
    const yearRows = [...monthMap.values()].flat();
    const yearSum = sumMautoRows(yearRows);
    const yearKey = `${kind}:${year}`;
    const yearHtml = `<tr class="mauto-year-row mauto-toggle-year" data-mauto-year="${escapeHtml(String(yearKey))}" style="cursor:pointer;">
      <td colspan="3"><span class="mauto-toggle-icon">▼</span> ${escapeHtml(String(year))}${Number(year) ? "년" : ""}</td>
      ${mautoNumericCell(yearSum.total)}
      ${mautoNumericCell(yearSum.inout)}
      ${mautoNumericCell(yearSum.balance, "mauto-balance-cell")}
    </tr>`;
    const monthHtml = [...monthMap.entries()].map(([month, monthRows]) => {
      const monthSum = sumMautoRows(monthRows);
      const monthKey = `${kind}:${year}:${month}`;
      const detailRows = monthRows.map(row => `
        <tr data-mauto-yr="${escapeHtml(String(yearKey))}" data-mauto-mo="${escapeHtml(String(monthKey))}">
          <td></td>
          <td>${escapeHtml(String(row.month || ""))}</td>
          <td>${escapeHtml(row.company || "")}</td>
          ${mautoNumericCell(row.total)}
          ${mautoNumericCell(row.inout)}
          ${mautoNumericCell(row.balance, "mauto-balance-cell")}
        </tr>`).join("");
      const monthRowHtml = `<tr class="mauto-month-row mauto-toggle-month" data-mauto-year="${escapeHtml(String(yearKey))}" data-mauto-month="${escapeHtml(String(monthKey))}" style="cursor:pointer;">
        <td></td>
        <td><span class="mauto-toggle-icon">▼</span> ${escapeHtml(String(month))}${Number(month) ? "월" : ""}</td>
        <td></td>
        ${mautoNumericCell(monthSum.total)}
        ${mautoNumericCell(monthSum.inout)}
        ${mautoNumericCell(monthSum.balance, "mauto-balance-cell")}
      </tr>`;
      return monthRowHtml + detailRows;
    }).join("");
    return yearHtml + monthHtml;
  }).join("");

  const total = sumMautoRows(sorted);
  const inoutLabel = isReceivable ? "입금" : "출금";
  return `<div class="mauto-table-wrap">
    <table class="mauto-table" data-kind="${kind}">
      <thead>
        <tr>
          <th>작성연도</th><th>작성월</th><th>상호</th>
          <th class="mauto-num">${totalLabel}</th>
          <th class="mauto-num">${inoutLabel}</th>
          <th class="mauto-num">잔액</th>
        </tr>
      </thead>
      <tbody>${body}</tbody>
      <tfoot>
        <tr>
          <td colspan="3">총합계</td>
          ${mautoNumericCell(total.total)}
          ${mautoNumericCell(total.inout)}
          ${mautoNumericCell(total.balance, "mauto-balance-cell")}
        </tr>
      </tfoot>
    </table>
  </div>`;
}

// 엠오토 미지급 업체별 보기 (거래처 기준 집계, 잔액 절댓값 내림차순)
function renderMautoPayablesByVendor(payRows) {
  const rows = payRows || [];
  if (!rows.length) {
    return `<div class="mauto-table-wrap">
      <table class="mauto-table">
        <thead><tr><th>상호</th><th>연월</th><th class="mauto-num">매입합계</th><th class="mauto-num">출금</th><th class="mauto-num">잔액</th></tr></thead>
        <tbody><tr><td colspan="5" class="mauto-empty">데이터 없음</td></tr></tbody>
      </table>
    </div>`;
  }

  // 업체별 그룹화
  const vendors = new Map();
  rows.forEach(row => {
    const vk = row.company || "(업체 없음)";
    if (!vendors.has(vk)) vendors.set(vk, []);
    vendors.get(vk).push(row);
  });

  // 잔액 절댓값 내림차순 정렬 (잔액 큰 업체 위로)
  const sorted = [...vendors.entries()].sort((a, b) =>
    Math.abs(sumMautoRows(b[1]).balance) - Math.abs(sumMautoRows(a[1]).balance)
  );

  const body = sorted.map(([company, vRows], idx) => {
    const vSum = sumMautoRows(vRows);
    const vKey = `vendor-pay:${idx}`;
    const sortedVRows = [...vRows].sort((a, b) =>
      (a.year || 9999) - (b.year || 9999) || (a.month || 99) - (b.month || 99)
    );
    const headerRow = `<tr class="mauto-year-row mauto-toggle-year" data-mauto-year="${escapeHtml(vKey)}" style="cursor:pointer;">
      <td><span class="mauto-toggle-icon">▼</span> ${escapeHtml(company)}</td>
      <td></td>
      ${mautoNumericCell(vSum.total)}
      ${mautoNumericCell(vSum.inout)}
      ${mautoNumericCell(vSum.balance, "mauto-balance-cell")}
    </tr>`;
    const detailRows = sortedVRows.map(row => `
      <tr data-mauto-yr="${escapeHtml(vKey)}">
        <td></td>
        <td style="font-size:12px;color:#6b7280;">${row.year || ""}년 ${row.month || ""}월</td>
        ${mautoNumericCell(row.total)}
        ${mautoNumericCell(row.inout)}
        ${mautoNumericCell(row.balance, "mauto-balance-cell")}
      </tr>`).join("");
    return headerRow + detailRows;
  }).join("");

  const total = sumMautoRows(rows);
  return `<div class="mauto-table-wrap">
    <table class="mauto-table" data-kind="payables-vendor">
      <thead><tr>
        <th>상호</th><th>연월</th>
        <th class="mauto-num">매입합계</th>
        <th class="mauto-num">출금</th>
        <th class="mauto-num">잔액</th>
      </tr></thead>
      <tbody>${body}</tbody>
      <tfoot><tr>
        <td colspan="2">총합계 (${sorted.length}개 업체)</td>
        ${mautoNumericCell(total.total)}
        ${mautoNumericCell(total.inout)}
        ${mautoNumericCell(total.balance, "mauto-balance-cell")}
      </tr></tfoot>
    </table>
  </div>`;
}

function getMautoFixedDateLabel(row) {
  if (row.date) return row.date;
  if (row.year && row.month && row.day)
    return `${row.year}-${String(row.month).padStart(2, "0")}-${String(row.day).padStart(2, "0")}`;
  if (row.year && row.month) return `${row.year}-${String(row.month).padStart(2, "0")}`;
  return "날짜 없음";
}

// ── Phase 4-B: 분류규칙 결제예정일 기반 고정지출 자동 계산 ──

// 한국 공휴일 (2025~2026, 대체공휴일 포함)
const KR_HOLIDAYS = new Set([
  // 2025
  "2025-01-01","2025-01-28","2025-01-29","2025-01-30",           // 신정 / 설날연휴
  "2025-03-01","2025-03-03",                                       // 삼일절(토)+대체(월)
  "2025-05-05","2025-05-06",                                       // 어린이날+부처님오신날 대체
  "2025-06-06",                                                    // 현충일(금)
  "2025-08-15",                                                    // 광복절(금)
  "2025-10-03","2025-10-05","2025-10-06","2025-10-07","2025-10-08","2025-10-09", // 개천절/추석연휴+대체/한글날
  "2025-12-25",                                                    // 성탄절(목)
  // 2026
  "2026-01-01",                                                    // 신정(목)
  "2026-02-16","2026-02-17","2026-02-18",                         // 설날연휴(월~수)
  "2026-03-01","2026-03-02",                                       // 삼일절(일)+대체(월)
  "2026-05-05","2026-05-24",                                       // 어린이날(수)/부처님오신날(약)
  "2026-06-06","2026-06-08",                                       // 현충일(토)+대체(월)
  "2026-08-15","2026-08-17",                                       // 광복절(토)+대체(월)
  "2026-09-24","2026-09-25","2026-09-26","2026-09-28",            // 추석연휴(목~토)+대체(월)
  "2026-10-03","2026-10-05",                                       // 개천절(토)+대체(월)
  "2026-10-09",                                                    // 한글날(금)
  "2026-12-25",                                                    // 성탄절(금)
]);

const DAYS_KO = ["일","월","화","수","목","금","토"];

// 결제예정일(N일)을 해당 월의 실제 영업일로 변환 (주말/공휴일 → 다음 영업일)
function getScheduledPaymentDate(year, month, day) {
  const lastDay = new Date(year, month, 0).getDate();
  let d = new Date(year, month - 1, Math.min(day, lastDay));
  for (let i = 0; i < 14; i++) {
    const ds = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
    const dow = d.getDay();
    if (dow !== 0 && dow !== 6 && !KR_HOLIDAYS.has(ds)) return { date: ds, dow: DAYS_KO[dow] };
    d.setDate(d.getDate() + 1);
  }
  const ds = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
  return { date: ds, dow: DAYS_KO[d.getDay()] };
}

function buildFixedFromRules(fixedRules, classifiedRows) {
  const today = new Date();
  const todayYM = `${today.getFullYear()}-${String(today.getMonth()+1).padStart(2,"0")}`;

  // 표시 범위: 전월 ~ 현월 ~ 다음월 (3개월)
  const monthSet = new Set();
  for (let i = -1; i <= 1; i++) {
    const d = new Date(today.getFullYear(), today.getMonth() + i, 1);
    monthSet.add(`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`);
  }

  const months = [...monthSet].sort().reverse(); // 최신 먼저

  // 비활성(활성여부=N) 제외 + 같은 거래처명+결제예정일 중복 제거
  const seenFixed = new Set();
  const dedupedRules = fixedRules.filter(rule => {
    if (String(rule["활성여부"] || "Y").toUpperCase() === "N") return false;
    const k = `${rule["거래처명"]}||${rule["결제예정일"]}`;
    if (seenFixed.has(k)) return false;
    seenFixed.add(k);
    return true;
  });

  // 거래처별 월 실적 평균 계산 → 예정금액 자동 산출 (2개월 이상 데이터 있을 때)
  const vendorActuals = {}; // { vendorName: [amount, ...] }
  months.forEach(ym => {
    dedupedRules.forEach(rule => {
      const name = rule["거래처명"];
      const matched = (classifiedRows || []).filter(r => {
        const d = r._date || r.date || "";
        return d && r.거래처명 === name && d.slice(0, 7) === ym;
      });
      if (matched.length) {
        const total = matched.reduce((s, r) => s + Math.abs(Number(r._debit || r.debit || 0) || Number(r._credit || r.credit || 0)), 0);
        if (total > 0) {
          if (!vendorActuals[name]) vendorActuals[name] = [];
          vendorActuals[name].push(total);
        }
      }
    });
  });
  // 2개월 이상이면 평균 → 천원 단위 올림, 아니면 규칙 입력값 fallback
  const vendorCalcExpected = {};
  Object.entries(vendorActuals).forEach(([name, amounts]) => {
    if (amounts.length >= 2) {
      const avg = amounts.reduce((s, a) => s + a, 0) / amounts.length;
      vendorCalcExpected[name] = Math.ceil(avg / 1000) * 1000;
    }
  });

  return months.map(ym => {
    const [year, month] = ym.split("-");
    const isPast = ym < todayYM;
    const isCurrent = ym === todayYM;
    const monthRules = dedupedRules.filter(rule => {
      const 지급월List = rule["지급월"] ? String(rule["지급월"]).split(",").map(s => s.trim()).filter(Boolean) : [];
      return 지급월List.length === 0 || 지급월List.includes(String(parseInt(month)));
    });
    const items = monthRules.map(rule => {
      const matched = (classifiedRows || []).filter(r => {
        const d = r._date || r.date || "";
        if (!d || !r.거래처명) return false;
        if (d.slice(0, 7) !== ym) return false;
        return r.거래처명 === rule["거래처명"];
      });
      const totalAmount = matched.reduce((s, r) => s + Math.abs(Number(r._debit || r.debit || 0) || Number(r._credit || r.credit || 0)), 0);
      const dates = [...new Set(matched.map(r => (r._date || r.date || "")).filter(Boolean))].sort();
      const dayNum = parseInt(rule["결제예정일"]) || null;
      const 예정결제일 = dayNum ? getScheduledPaymentDate(parseInt(year), parseInt(month), dayNum) : null;
      // 예정금액: 실적 평균(2개월↑) 우선 → 규칙 수동값 fallback
      const calcAmt = vendorCalcExpected[rule["거래처명"]];
      const manualAmt = Number(rule["예정금액"]) || 0;
      const 예정금액 = calcAmt || (manualAmt ? Math.ceil(manualAmt / 1000) * 1000 : 0);
      const 예정금액출처 = calcAmt ? "auto" : "manual";
      return {
        거래처명: rule["거래처명"],
        구분: rule["구분"] || "",
        고정분류: rule["고정분류"] || "",
        예정일: dayNum,
        예정결제일,
        예정금액, 예정금액출처,
        matched, totalAmount, dates,
        status: matched.length > 0 ? "완료" : "예정",
      };
    });
    const monthTotal = items.reduce((s, i) => s + i.totalAmount, 0);
    const allDone = items.length > 0 && items.every(i => i.status === "완료");
    return { year, month, ym, items, monthTotal, isPast, isCurrent, allDone };
  });
}

function renderMautoFixedAutoView(fixedRules, classifiedRows, prebuiltData = null) {
  const items = (fixedRules || []).filter(r => r["결제예정일"]);
  if (!items.length) {
    return `<div style="padding:14px 12px;color:#6b7280;font-size:13px;">📐 분류규칙에 결제예정일이 설정된 항목이 없습니다.<br>분류규칙 관리 → 항목 수정 → <strong>결제예정일</strong>(N일) 입력 후 "불러오기"를 누르세요.</div>`;
  }
  const monthData = prebuiltData || buildFixedFromRules(items, classifiedRows || []);
  if (!monthData.length) {
    return `<div style="padding:14px 12px;color:#6b7280;font-size:13px;">입출금 내역을 업로드하면 월별 실적이 자동으로 채워집니다.</div>`;
  }

  const tdSt = `padding:4px 6px;border-bottom:1px solid #f3f4f6;`;
  const thSt = `padding:4px 6px;border-bottom:1px solid #e5e7eb;color:#6b7280;font-weight:600;font-size:11px;`;
  const CAT_COLOR = { 이자:"#dbeafe", 인출금:"#fef9c3", 카드:"#f3e8ff", 세금:"#fee2e2", 세계:"#d1fae5", 복리:"#ffedd5" };

  // 미완료 항목이 있는 월 or 이번 달/미래 → 기본 펼침 / 완료된 과거 달 → 접힘
  const today2 = new Date();
  const todayYM2 = `${today2.getFullYear()}-${String(today2.getMonth()+1).padStart(2,"0")}`;
  const renderMonth = ({ year, month, ym, items: monthItems, monthTotal, isPast, isCurrent, allDone }) => {
    const today = new Date();
    const todayYM = `${today.getFullYear()}-${String(today.getMonth()+1).padStart(2,"0")}`;
    const collapsed = !isCurrent; // 현재 월만 기본 펼침

    // 날짜 오름차순 정렬 (날짜 없는 항목은 맨 뒤)
    const sorted = [...monthItems].sort((a, b) => {
      const da = a.예정결제일?.date || "9999";
      const db = b.예정결제일?.date || "9999";
      return da < db ? -1 : da > db ? 1 : 0;
    });

    // 날짜별로 그룹화 → 날짜별 소계
    const dateGroups = [];
    sorted.forEach(item => {
      const d = item.예정결제일?.date || "미정";
      const last = dateGroups[dateGroups.length - 1];
      if (last && last.date === d) last.items.push(item);
      else dateGroups.push({ date: d, dow: item.예정결제일?.dow || "", items: [item] });
    });

    const bodyRows = dateGroups.map(({ date, dow, items: grpItems }) => {
      const grpExpected = grpItems.reduce((s, i) => s + (i.예정금액 || 0), 0);
      const grpActual  = grpItems.reduce((s, i) => s + i.totalAmount, 0);
      const dateLabel  = date === "미정" ? "미정" : `${date.slice(5)} (${dow})`;
      const isAdj = grpItems.some(i => i.예정일 && i.예정결제일 && i.예정일 !== parseInt(i.예정결제일.date.slice(8)));
      const chkKey = `${ym}||${date}`;
      const dgKey  = `${ym}||${date}`;
      const allGrpDone = grpItems.every(i => i.status === "완료");
      const defaultChecked = !allGrpDone && (ym === todayYM2);
      const isChecked = allGrpDone ? false : (mautoFixedChecked[chkKey] !== undefined ? mautoFixedChecked[chkKey] : defaultChecked);

      // 날짜 소계 행 (체크박스 + 아코디언 토글)
      const subtotalRow = `<tr class="mauto-fixed-dg-hdr" data-dg="${dgKey}" style="background:#f8fafc;cursor:pointer;">
        <td style="${tdSt}text-align:center;width:28px;" onclick="event.stopPropagation()"><input type="checkbox" class="mauto-fixed-chk" data-chk-key="${chkKey}" data-amt="${grpExpected}" ${isChecked ? "checked" : ""} ${allGrpDone ? "disabled title=\"모든 항목 완료\"" : ""} style="cursor:${allGrpDone ? "default" : "pointer"};accent-color:#2563eb;width:14px;height:14px;${allGrpDone ? "opacity:0.35;" : ""}" /></td>
        <td colspan="2" style="${tdSt}font-weight:700;color:#374151;"><span class="mauto-fixed-dg-icon">▼</span> ${dateLabel}${isAdj ? ' <span style="color:#f59e0b;font-size:10px;">*조정</span>' : ""}</td>
        <td style="${tdSt}text-align:right;font-weight:700;color:#9ca3af;">${grpExpected ? formatNumber(grpExpected) : ""}</td>
        <td style="${tdSt}"></td>
        <td style="${tdSt}text-align:right;font-weight:700;">${grpActual ? formatNumber(grpActual) : ""}</td>
        <td style="${tdSt}"></td>
      </tr>`;

      const itemRows = grpItems.map(item => {
        const paid = item.status === "완료";
        const catBg = CAT_COLOR[item.고정분류] ? `background:${CAT_COLOR[item.고정분류]};` : "";
        const amtKey = `${ym}||${item.거래처명}||${item.예정일 || "0"}`;
        const isOverride = item.예정금액출처 === "override";
        const amtBadge = item.예정금액출처 === "auto"
          ? '<span style="font-size:9px;color:#2563eb;margin-left:2px;">계산</span>'
          : isOverride ? '<span style="font-size:9px;color:#f59e0b;margin-left:2px;">수정</span>'
          : "";
        const amtBorderColor = isOverride ? "#f59e0b" : "#d1d5db";
        return `<tr class="mauto-fixed-dg-item" data-dg="${dgKey}" style="${paid ? "opacity:0.55;" : ""}${catBg}">
          <td style="${tdSt}"></td>
          <td style="${tdSt}padding-left:22px;${paid ? "text-decoration:line-through;color:#9ca3af;" : ""}">${escapeHtml(item.거래처명)}<span style="margin-left:4px;font-size:10px;color:#9ca3af;">${item.고정분류 || ""}</span></td>
          <td style="${tdSt}text-align:center;color:#9ca3af;font-size:11px;">${item.예정일 ? `${item.예정일}일` : "-"}</td>
          <td style="${tdSt}text-align:right;">
            <input type="text" class="mauto-fixed-amt-input"
              data-amt-key="${escapeHtml(amtKey)}"
              data-amt-orig="${item.예정금액 || 0}"
              value="${item.예정금액 ? escapeHtml(formatNumber(item.예정금액)) : ""}"
              style="width:72px;text-align:right;border:none;border-bottom:1px dashed ${amtBorderColor};background:transparent;font-size:inherit;color:inherit;padding:0 2px;cursor:text;"
              title="클릭하여 수정 (빈칸 저장 시 자동계산으로 복원)"
            />${amtBadge}
          </td>
          <td style="${tdSt}text-align:center;color:#6b7280;font-size:11px;">${item.dates.join(", ") || "-"}</td>
          <td style="${tdSt}text-align:right;">${item.totalAmount ? formatNumber(item.totalAmount) : "-"}</td>
          <td style="${tdSt}text-align:center;${paid ? "color:#16a34a;font-weight:700;" : (isPast ? "color:#ef4444;font-weight:700;" : "color:#9ca3af;")}">${paid ? "✓" : (isPast ? "미결" : "-")}</td>
        </tr>`;
      }).join("");

      return subtotalRow + itemRows;
    }).join("");

    const doneBadge = allDone
      ? `<span style="font-size:11px;color:#16a34a;margin-left:6px;">✓ 완료</span>`
      : (isPast ? `<span style="font-size:11px;color:#ef4444;margin-left:6px;">● 미결</span>` : "");
    const headerBorder = isCurrent ? "2px solid #2563eb" : "2px solid #e5e7eb";
    const tableHtml = `<div style="overflow-x:auto;"><table style="width:100%;min-width:480px;border-collapse:collapse;font-size:12px;">
        <thead><tr>
          <th style="${thSt}text-align:center;width:28px;">☑</th>
          <th style="${thSt}text-align:left;">항목</th>
          <th style="${thSt}text-align:center;">기준일</th>
          <th style="${thSt}text-align:right;">예정금액</th>
          <th style="${thSt}text-align:center;">실제결제일</th>
          <th style="${thSt}text-align:right;">실적</th>
          <th style="${thSt}text-align:center;">상태</th>
        </tr></thead>
        <tbody>${bodyRows}</tbody>
      </table></div>`;
    return `<details style="margin-bottom:12px;" ${collapsed ? "" : "open"}>
      <summary style="cursor:pointer;font-weight:700;font-size:13px;color:${isCurrent?"#2563eb":"#374151"};padding:4px 0 6px;border-bottom:${headerBorder};list-style:none;display:flex;align-items:center;gap:6px;">
        <span>${collapsed ? "▶" : "▼"}</span>
        <span>${year}년 ${parseInt(month)}월${doneBadge}</span>
        <span style="margin-left:auto;color:#6b7280;font-weight:400;font-size:12px;">합계 ${formatNumber(monthTotal)}</span>
      </summary>
      <div style="padding-top:4px;">${tableHtml}</div>
    </details>`;
  };

  const html = monthData.map(renderMonth).join("");

  return `<div style="padding:10px 4px;max-height:600px;overflow-y:auto;">
    <div style="display:flex;gap:6px;margin-bottom:8px;">
      <button type="button" id="fixedAutoExpandAll" style="font-size:11px;padding:2px 10px;border:1px solid #d1d5db;border-radius:4px;background:#f9fafb;cursor:pointer;">전체 펼치기</button>
      <button type="button" id="fixedAutoCollapseAll" style="font-size:11px;padding:2px 10px;border:1px solid #d1d5db;border-radius:4px;background:#f9fafb;cursor:pointer;">전체 접기</button>
    </div>
    ${html}
  </div>`;
}

function renderMautoFixedTable(rows) {
  const sorted = [...(rows || [])].sort((a, b) =>
    (a.year || 9999) - (b.year || 9999) || (a.month || 99) - (b.month || 99) ||
    (a.day || 99) - (b.day || 99) || String(a.title || "").localeCompare(String(b.title || ""), "ko")
  );
  const banks = [...new Set(sorted.map(r => r.bank || "미지정"))].sort((a, b) => a.localeCompare(b, "ko"));
  const bankHeaders = banks.length ? banks : ["금액"];

  if (!sorted.length) {
    return `<div class="mauto-table-wrap">
      <table class="mauto-table">
        <thead><tr><th>날짜</th><th>내용</th><th>분류</th><th class="mauto-num">금액</th></tr></thead>
        <tbody><tr><td colspan="4" class="mauto-empty">데이터 없음</td></tr></tbody>
      </table>
    </div>`;
  }

  const groups = new Map();
  sorted.forEach(row => {
    const key = getMautoFixedDateLabel(row);
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(row);
  });

  const body = [...groups.entries()].map(([dateLabel, items]) => {
    const groupTotal = items.reduce((s, r) => s + Number(r.amount || 0), 0);
    const bankTotals = {};
    bankHeaders.forEach(b => { bankTotals[b] = 0; });
    items.forEach(r => { bankTotals[r.bank || "미지정"] = (bankTotals[r.bank || "미지정"] || 0) + Number(r.amount || 0); });
    const dateKey = `mfx:${dateLabel}`;
    const detailRows = items.map(r => `
      <tr data-mauto-fxdate="${escapeHtml(dateKey)}">
        <td></td>
        <td>${escapeHtml(r.title || "")}</td>
        <td>${escapeHtml(r.category || "")}</td>
        ${bankHeaders.map(b => `<td class="mauto-num">${b === (r.bank || "미지정") ? formatNumber(r.amount) : ""}</td>`).join("")}
        <td class="mauto-num mauto-balance-cell">${formatNumber(r.amount)}</td>
      </tr>`).join("");
    return `<tr class="mauto-date-row mauto-toggle-fxdate" data-mauto-fxdate="${escapeHtml(dateKey)}" style="cursor:pointer;">
        <td><span class="mauto-toggle-icon">▼</span> ${escapeHtml(dateLabel)}</td>
        <td>${items.length}건</td>
        <td></td>
        ${bankHeaders.map(b => `<td class="mauto-num">${bankTotals[b] ? formatNumber(bankTotals[b]) : ""}</td>`).join("")}
        <td class="mauto-num mauto-balance-cell">${formatNumber(groupTotal)}</td>
      </tr>
      ${detailRows}`;
  }).join("");

  const grandTotal = sorted.reduce((s, r) => s + Number(r.amount || 0), 0);
  const grandBankTotals = {};
  bankHeaders.forEach(b => { grandBankTotals[b] = 0; });
  sorted.forEach(r => { grandBankTotals[r.bank || "미지정"] = (grandBankTotals[r.bank || "미지정"] || 0) + Number(r.amount || 0); });

  return `<div class="mauto-table-wrap">
    <table class="mauto-table" data-kind="fixed">
      <thead>
        <tr>
          <th>날짜</th><th>내용</th><th>분류</th>
          ${bankHeaders.map(b => `<th class="mauto-num">${escapeHtml(b)}</th>`).join("")}
          <th class="mauto-num">총합계</th>
        </tr>
      </thead>
      <tbody>${body}</tbody>
      <tfoot>
        <tr>
          <td colspan="3">총합계</td>
          ${bankHeaders.map(b => `<td class="mauto-num">${grandBankTotals[b] ? formatNumber(grandBankTotals[b]) : ""}</td>`).join("")}
          <td class="mauto-num mauto-balance-cell">${formatNumber(grandTotal)}</td>
        </tr>
      </tfoot>
    </table>
  </div>`;
}

// buildArRecap 결과 → renderMautoAccountingTable 포맷 변환 (발생 0·잔액 0·제외 거래처 제외)
function arRecapToMautoRows(entries, side) {
  return entries
    .filter(e => e.발생 !== 0)
    .filter(e => e.잔액 !== 0)
    .filter(e => !isArRecapExcluded(e.vendor, side))
    .map(e => ({
      year: e.year, month: e.month, company: e.vendor,
      total: e.발생, inout: e.충당, balance: e.잔액,
    }));
}

function mautoPasteSection(id, title, tableHtml, hint, hasToggle = false, badge = "") {
  const toggleBtns = hasToggle ? `
    <button type="button" class="mauto-ctrl-btn" data-mauto-expand-all="${id}">전체 펼치기</button>
    <button type="button" class="mauto-ctrl-btn" data-mauto-collapse-all="${id}">전체 접기</button>` : "";
  return `<div class="mauto-section" id="mauto-section-${id}" data-kind="${id}">
    <div class="mauto-section-header">
      <div><h3>${escapeHtml(title)}${badge ? ` ${badge}` : ""}</h3></div>
      <div class="mauto-section-actions">
        ${toggleBtns}
        <button type="button" class="mauto-paste-btn" data-mauto-section="${id}">붙여넣기 입력</button>
      </div>
    </div>
    <div class="mauto-paste-area hidden" id="mauto-paste-area-${id}">
      <div class="mauto-paste-hint">${escapeHtml(hint)}</div>
      <textarea class="mauto-textarea" id="mauto-textarea-${id}" placeholder="엑셀에서 복사한 표를 여기에 붙여넣으세요"></textarea>
      <div class="mauto-paste-actions">
        <button type="button" class="mauto-apply-btn" data-mauto-section="${id}">적용</button>
        <button type="button" class="mauto-cancel-btn" data-mauto-section="${id}">취소</button>
      </div>
    </div>
    ${tableHtml}
  </div>`;
}

function renderMautoTab() {
  const sec = elements.mauto || document.getElementById("mauto");
  if (!sec) return;
  mautoData = normalizeMautoData(mautoData);
  const funds = (mautoData.funds || []).reduce((s, r) => s + Number(r.amount || 0), 0);
  let fixed = 0;
  let _fixedMonthDataCache = null;
  if (mautoFixedRules !== null) {
    const ruleItems = (mautoFixedRules || []).filter(r => r["결제예정일"]);
    if (ruleItems.length) {
      _fixedMonthDataCache = buildFixedFromRules(ruleItems, mautoClassifiedRows);
      applyFixedAmountOverrides(_fixedMonthDataCache);
      fixed = calcFixedCheckedTotal(_fixedMonthDataCache);
    }
  } else {
    fixed = (mautoData.fixed || []).reduce((s, r) => s + Number(r.amount || 0), 0);
  }

  // 세금계산서 데이터가 있으면 buildArRecap으로 미수/미지급 자동 계산
  const hasTax = mautoTaxInvoices && mautoTaxInvoices.length > 0;
  let rcvRows, payRows, receivable, payable, rcvBadge, payBadge, rcvWarn = "", payWarn = "";
  if (hasTax) {
    const rcv = buildArRecap(mautoTaxInvoices, mautoClassifiedRows || [], "미수");
    const pay = buildArRecap(mautoTaxInvoices, mautoClassifiedRows || [], "미지급");
    rcvRows = arRecapToMautoRows(rcv.entries, "rcv");
    payRows = arRecapToMautoRows(pay.entries, "pay");
    receivable = rcvRows.reduce((s, r) => s + r.balance, 0);
    payable    = payRows.reduce((s, r) => s + r.balance, 0);
    const taxCnt = mautoTaxInvoices.length;
    const exStyleBase = `font-size:11px;margin-left:8px;padding:1px 7px;border:1px solid #d1d5db;border-radius:10px;background:#f3f4f6;cursor:pointer;`;
    const taxBadge = `<span style="font-size:11px;color:#2563eb;font-weight:600;margin-left:6px;">세금계산서 ${taxCnt}건 기준</span>`;
    const rcvExLabel = mautoExcludeVendorsRcv.length ? `제외 ${mautoExcludeVendorsRcv.length}개` : `제외 설정`;
    const payExLabel = mautoExcludeVendorsPay.length ? `제외 ${mautoExcludeVendorsPay.length}개` : `제외 설정`;
    const payViewToggle = `<button type="button" id="mautoPayViewYm" style="${exStyleBase}${mautoPayViewMode==="ym"?"background:#2563eb;color:#fff;border-color:#2563eb;":""}" title="연월별로 보기">연월별</button>` +
      `<button type="button" id="mautoPayViewVendor" style="${exStyleBase}${mautoPayViewMode==="vendor"?"background:#2563eb;color:#fff;border-color:#2563eb;":""}" title="업체별로 보기">업체별</button>`;
    rcvBadge = taxBadge + `<button type="button" id="mautoExcludeBtnRcv" style="${exStyleBase}" title="미수금 제외 거래처 설정">${rcvExLabel}</button>`;
    payBadge = taxBadge + `<button type="button" id="mautoExcludeBtnPay" style="${exStyleBase}" title="미지급 제외 거래처 설정">${payExLabel}</button>` + payViewToggle;
    if (rcv.확인필요.length) rcvWarn = `<div style="margin:4px 0 6px;padding:6px 10px;background:#fef9c3;border:1px solid #fde68a;border-radius:4px;font-size:12px;color:#92400e;">⚠ 귀속연월 미확인 ${rcv.확인필요.length}건 — 입금 미반영 (입출금 분류 비고에 연월 기재 필요)</div>`;
    if (pay.확인필요.length) payWarn = `<div style="margin:4px 0 6px;padding:6px 10px;background:#fef9c3;border:1px solid #fde68a;border-radius:4px;font-size:12px;color:#92400e;">⚠ 귀속연월 미확인 ${pay.확인필요.length}건 — 출금 미반영 (입출금 분류 비고에 연월 기재 필요)</div>`;
  } else {
    rcvRows = mautoData.receivables || [];
    payRows = mautoData.payables || [];
    receivable = rcvRows.reduce((s, r) => s + Number(r.balance || 0), 0);
    payable    = payRows.reduce((s, r) => s + Number(r.balance || 0), 0);
    rcvBadge = ""; payBadge = "";
  }

  sec.innerHTML = `<div class="mauto-container">
    <div class="mauto-top-bar">
      <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;">
        <h2 style="margin:0;">엠오토</h2>
        <button type="button" id="mautoToolsToggle" title="입력·업로드 도구 펼치기/접기"
          style="font-size:12px;padding:3px 11px;border:1px solid #cbd5e1;border-radius:14px;background:#f1f5f9;color:#475569;cursor:pointer;white-space:nowrap;">
          ${mautoToolsOpen ? "▲ 도구 접기" : "▼ 입력·업로드 도구"}
        </button>
      </div>
    </div>
    <div id="mautoToolsPanel" style="display:${mautoToolsOpen ? "block" : "none"};">
      <p style="color:var(--muted);margin:2px 0 8px;">엠오토 전용 현금흐름</p>
      <div class="mauto-top-actions">
        <label class="header-action-button" style="cursor:pointer;" title="입출금 내역 엑셀 파일을 분류규칙으로 자동 분류">
          입출금 분류
          <input type="file" id="mautoClassifyFileInput" accept=".xls,.xlsx,.xlsm,.csv" hidden />
        </label>
        <label class="header-action-button" style="cursor:pointer;" title="국세청 전자세금계산서 조회 파일 (매출)">
          매출세금계산서
          <input type="file" id="mautoTaxSalesFileInput" accept=".xls,.xlsx" hidden />
        </label>
        <label class="header-action-button" style="cursor:pointer;" title="국세청 전자세금계산서 조회 파일 (매입)">
          매입세금계산서
          <input type="file" id="mautoTaxPurchaseFileInput" accept=".xls,.xlsx" hidden />
        </label>
        <button type="button" id="mautoVatBtn" class="daesa-vat-btn${mautoVatView ? " active" : ""}" title="부가세 납부세액 집계 보고서 (반기/월간/연간)">부가세</button>
        <button type="button" id="mautoClearBtn" class="mauto-clear-btn">전체 초기화</button>
      </div>
    ${(() => {
      const taxEntries = Object.values(mautoTaxSources);
      if (!taxEntries.length) return "";
      return `
    <details data-accordion-id="mauto-tax-files" style="margin:6px 0;border:1px solid #fde68a;border-radius:6px;background:#fefce8;font-size:12px;">
      <summary style="padding:8px 12px;cursor:pointer;font-weight:600;color:#92400e;list-style:none;display:flex;align-items:center;gap:6px;">
        <span>▶</span> 세금계산서 ${taxEntries.length}개 파일 (${mautoTaxInvoices.length}건)
      </summary>
      <div style="padding:4px 12px 8px;">
      ${taxEntries.map(f => `
      <div style="display:flex;align-items:center;gap:8px;padding:3px 0;border-bottom:1px solid #fef9c3;">
        <span style="flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${escapeHtml(f.filename)}</span>
        <span style="color:#b45309;">${f.sideType}</span>
        <span style="color:#6b7280;white-space:nowrap;">${(f.rows||[]).length}건</span>
        ${f.checksumOk === true ? '<span style="color:#16a34a;">✓체크섬</span>' : f.checksumOk === false ? '<span style="color:#dc2626;">✗체크섬</span>' : ""}
        <button type="button" class="mauto-tax-del-btn" data-fname="${encodeURIComponent(f.filename)}" style="font-size:11px;padding:1px 7px;border:1px solid #fecaca;background:#fff;border-radius:3px;cursor:pointer;color:#dc2626;">삭제</button>
      </div>`).join("")}
      </div>
    </details>`;
    })()}
    ${(() => {
      const fileEntries = Object.entries(mautoSourceFiles);
      if (!fileEntries.length) return "";
      return `
    <details data-accordion-id="mauto-bank-files" style="margin:6px 0;border:1px solid #e2e8f0;border-radius:6px;background:#f8fafc;font-size:12px;">
      <summary style="padding:8px 12px;cursor:pointer;font-weight:600;color:#374151;list-style:none;display:flex;align-items:center;gap:6px;">
        <span>▶</span> 업로드된 파일 ${fileEntries.length}개
      </summary>
      <div style="padding:4px 12px 8px;">
      ${fileEntries.map(([key, f]) => `
      <div style="display:flex;align-items:center;gap:8px;padding:3px 0;border-bottom:1px solid #f1f5f9;">
        <span style="flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:${f.isMigration ? '#9ca3af' : '#374151'};">${escapeHtml(f.filename)}${f.isMigration ? ' <em style="color:#d97706;">(마이그레이션)</em>' : ''}</span>
        <span style="color:#6b7280;white-space:nowrap;">${(f.rows||[]).length}건</span>
        <span style="color:#9ca3af;white-space:nowrap;">${(f.savedAt||"").slice(0,10)}</span>
        <button type="button" class="mauto-file-del-btn" data-fkey="${encodeURIComponent(key)}" style="font-size:11px;padding:1px 7px;border:1px solid #fecaca;background:#fff;border-radius:3px;cursor:pointer;color:#dc2626;">삭제</button>
      </div>`).join("")}
      </div>
    </details>`;
    })()}
    ${(() => {
      if (!mautoClassifiedRows.length) return "";
      const active = mautoClassifiedRows.filter(r => !r.excluded && r.거래처명);
      const excl = mautoClassifiedRows.filter(r => r.excluded);
      const unmatched = mautoClassifiedRows.filter(r => !r.excluded && !r.거래처명);
      return `
    <div style="margin:8px 0;padding:10px 14px;background:#f0fdf4;border:1px solid #bbf7d0;border-radius:6px;font-size:13px;display:flex;gap:16px;align-items:center;flex-wrap:wrap;">
      <span style="font-weight:600;color:#166534;">${mautoClassifiedRows.length}건 저장됨</span>
      <span style="color:#374151;">매출 ${active.filter(r=>r.구분==="매출").length}건 / 매입 ${active.filter(r=>r.구분==="매입").length}건</span>
      ${unmatched.length ? `<span style="color:#d97706;font-weight:600;">⚠ 미매칭 ${unmatched.length}건</span>` : ""}
      ${excl.length ? `<span style="color:#9ca3af;">제외 ${excl.length}건</span>` : ""}
      <button type="button" id="mautoClassifyViewBtn" style="font-size:12px;padding:3px 10px;border:1px solid #16a34a;background:white;border-radius:4px;cursor:pointer;color:#15803d;">목록 보기</button>
      <button type="button" id="mautoClassifyClearBtn" style="font-size:12px;padding:3px 10px;border:1px solid #d1d5db;background:white;border-radius:4px;cursor:pointer;color:#6b7280;">지우기</button>
    </div>`;
    })()}
    </div><!-- /mautoToolsPanel (도구 접기 영역 끝) -->
    <div class="mauto-summary-grid">
      <div class="mauto-card card-funds"><span>가용자금</span><strong data-raw="${funds}">${formatNumber(funds)}</strong></div>
      <div class="mauto-card card-receivable"><span>미수금 잔액</span><strong data-raw="${receivable}">${formatNumber(receivable)}</strong></div>
      <div class="mauto-card card-payable"><span>미지급 잔액</span><strong data-raw="${payable}">${formatNumber(payable)}</strong></div>
      <div class="mauto-card card-fixed"><span>고정지출</span><strong id="mauto-fixed-card-total" data-raw="${fixed}">${formatNumber(fixed)}</strong></div>
      ${(() => { const net = funds + receivable - payable - fixed; return `<div class="mauto-card card-net" style="border-top-color:${net>=0?"#16a34a":"#dc2626"};"><span>예상 잔액</span><strong id="mauto-net-total" style="color:${net>=0?"#16a34a":"#dc2626"};">${net>=0?"+":""}${formatNumber(net)}</strong></div>`; })()}
    </div>
    ${mautoVatView ? renderMautoVatView() : ""}
    ${mautoPasteSection("funds", "가용자금",
      renderMautoFundsTable(),
      "금액만 2줄로 붙여넣기: 1행=국민(415310), 2행=부산(008320)", false)}
    ${mautoPasteSection("receivables", "미수금",
      rcvWarn + renderMautoAccountingTable(rcvRows, "receivables"),
      "헤더: 작성연도 / 작성 / 상호 / 매출합계 / 매출공급가액 / 매출세액 / 입금 / 잔액", true, rcvBadge)}
    ${mautoPasteSection("payables", "미지급",
      payWarn + (mautoPayViewMode === "vendor"
        ? renderMautoPayablesByVendor(payRows)
        : renderMautoAccountingTable(payRows, "payables")),
      "헤더: 작성연도 / 작성 / 상호 / 매입합계 / 매입공급가액 / 매입세액 / 출금 / 잔액", true, payBadge)}
    ${mautoPasteSection("fixed", "고정지출",
      mautoFixedRules !== null
        ? renderMautoFixedAutoView(mautoFixedRules, mautoClassifiedRows, _fixedMonthDataCache)
        : renderMautoFixedTable(mautoData.fixed),
      "헤더: 연도 / 월 / 내용 / 일 / 날짜 / 금액 / 은행 / 분류", true,
      `<button type="button" id="mautoFixedRulesBtn" style="font-size:11px;margin-left:8px;padding:1px 8px;border:1px solid #d1d5db;border-radius:10px;background:${mautoFixedRules !== null ? "#dbeafe" : "#f3f4f6"};cursor:pointer;" title="분류규칙 결제예정일 기반 자동계산">${mautoFixedRules !== null ? `규칙 ${(mautoFixedRules||[]).filter(r=>r["결제예정일"]).length}개 적용 중` : "규칙 불러오기"}</button>`)}
  </div>`;

  // 카드 클릭 → 해당 섹션 토글 (기본 숨김)
  const CARD_SECTION_MAP = {
    'card-funds':      'mauto-section-funds',
    'card-receivable': 'mauto-section-receivables',
    'card-payable':    'mauto-section-payables',
    'card-fixed':      'mauto-section-fixed',
  };
  sec.querySelectorAll('.mauto-section').forEach(s => { s.style.display = 'none'; });
  sec.querySelectorAll('.mauto-card').forEach(card => {
    const cls = [...card.classList].find(c => CARD_SECTION_MAP[c]);
    if (!cls) return;
    card.style.cursor = 'pointer';
    card.title = '클릭하여 상세 보기';
    card.addEventListener('click', () => {
      const targetId = CARD_SECTION_MAP[cls];
      const target = document.getElementById(targetId);
      if (!target) return;
      const isOpen = target.style.display !== 'none';
      sec.querySelectorAll('.mauto-section').forEach(s => { s.style.display = 'none'; });
      sec.querySelectorAll('.mauto-card').forEach(c => c.classList.remove('mauto-card-active'));
      if (!isOpen) {
        target.style.display = '';
        card.classList.add('mauto-card-active');
        target.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
      }
    });
  });

  sec.querySelectorAll(".mauto-paste-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      const area = document.getElementById(`mauto-paste-area-${btn.dataset.mautoSection}`);
      if (!area) return;
      area.classList.toggle("hidden");
      if (!area.classList.contains("hidden"))
        document.getElementById(`mauto-textarea-${btn.dataset.mautoSection}`)?.focus();
    });
  });

  sec.querySelectorAll(".mauto-cancel-btn").forEach(btn => {
    btn.addEventListener("click", () =>
      document.getElementById(`mauto-paste-area-${btn.dataset.mautoSection}`)?.classList.add("hidden"));
  });

  sec.querySelectorAll(".mauto-apply-btn").forEach(btn => {
    btn.addEventListener("click", () => applyMautoPaste(btn.dataset.mautoSection));
  });

  // 제외 거래처 설정 버튼 (미수금·미지급 별도)
  const openExcludeDialog = (side) => {
    document.querySelector(".mauto-exclude-overlay")?.remove();
    const iRcv = side === "rcv";
    const label = iRcv ? "미수금" : "미지급";
    const currentList = iRcv ? mautoExcludeVendorsRcv : mautoExcludeVendorsPay;
    const overlay = document.createElement("div");
    overlay.className = "mauto-exclude-overlay";
    overlay.style.cssText = "position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:9000;display:flex;align-items:center;justify-content:center;";
    overlay.innerHTML = `
      <div style="background:#fff;border-radius:8px;padding:24px;width:340px;box-shadow:0 8px 32px rgba(0,0,0,.2);">
        <h4 style="margin:0 0 8px;font-size:15px;">${label} 제외 거래처 설정</h4>
        <p style="margin:0 0 12px;font-size:12px;color:#6b7280;">${label}에 표시하지 않을 거래처를 한 줄에 하나씩 입력하세요.</p>
        <textarea id="mautoExcludeTextarea" rows="7" style="width:100%;box-sizing:border-box;border:1px solid #d1d5db;border-radius:4px;padding:8px;font-size:13px;resize:vertical;">${escapeHtml(currentList.join("\n"))}</textarea>
        <div style="display:flex;gap:8px;justify-content:flex-end;margin-top:12px;">
          <button id="mautoExcludeCancel" style="padding:6px 14px;border:1px solid #d1d5db;border-radius:4px;background:#fff;cursor:pointer;">취소</button>
          <button id="mautoExcludeSave" style="padding:6px 14px;border:none;border-radius:4px;background:#2563eb;color:#fff;cursor:pointer;font-weight:600;">저장</button>
        </div>
      </div>`;
    document.body.appendChild(overlay);
    document.getElementById("mautoExcludeTextarea")?.focus();
    document.getElementById("mautoExcludeCancel").onclick = () => overlay.remove();
    document.getElementById("mautoExcludeSave").onclick = () => {
      const val = document.getElementById("mautoExcludeTextarea").value;
      const parsed = val.split("\n").map(s => s.trim()).filter(Boolean);
      if (iRcv) mautoExcludeVendorsRcv = parsed;
      else      mautoExcludeVendorsPay = parsed;
      saveMautoExcludeVendors(side);
      overlay.remove();
      renderMautoTab();
    };
    overlay.addEventListener("click", e => { if (e.target === overlay) overlay.remove(); });
  };
  document.getElementById("mautoExcludeBtnRcv")?.addEventListener("click", () => openExcludeDialog("rcv"));
  document.getElementById("mautoExcludeBtnPay")?.addEventListener("click", () => openExcludeDialog("pay"));

  // 미지급 연월별/업체별 토글
  const _reopenPayables = () => {
    const ps = document.getElementById("mauto-section-payables");
    if (ps) ps.style.display = "";
    sec.querySelector(".card-payable")?.classList.add("mauto-card-active");
  };
  document.getElementById("mautoPayViewYm")?.addEventListener("click", () => {
    if (mautoPayViewMode === "ym") return;
    mautoPayViewMode = "ym"; renderMautoTab(); _reopenPayables();
  });
  document.getElementById("mautoPayViewVendor")?.addEventListener("click", () => {
    if (mautoPayViewMode === "vendor") return;
    mautoPayViewMode = "vendor"; renderMautoTab(); _reopenPayables();
  });

  // 고정지출 월 전체 펼치기 / 접기
  document.getElementById("fixedAutoExpandAll")?.addEventListener("click", () => {
    sec.querySelectorAll("#mauto-section-fixed details").forEach(d => { d.open = true; });
    sec.querySelectorAll(".mauto-fixed-dg-item").forEach(r => { r.style.display = ""; });
    sec.querySelectorAll(".mauto-fixed-dg-icon").forEach(i => { i.textContent = "▼"; });
  });
  document.getElementById("fixedAutoCollapseAll")?.addEventListener("click", () => {
    sec.querySelectorAll("#mauto-section-fixed details").forEach(d => { d.open = false; });
    sec.querySelectorAll(".mauto-fixed-dg-item").forEach(r => { r.style.display = "none"; });
    sec.querySelectorAll(".mauto-fixed-dg-icon").forEach(i => { i.textContent = "▶"; });
  });

  // 고정지출 날짜 행 클릭 → 세부 항목 아코디언 토글
  sec.querySelectorAll(".mauto-fixed-dg-hdr").forEach(hdr => {
    hdr.addEventListener("click", e => {
      if (e.target.type === "checkbox") return;
      const key = hdr.dataset.dg;
      const rows = sec.querySelectorAll(`.mauto-fixed-dg-item[data-dg="${key}"]`);
      const icon = hdr.querySelector(".mauto-fixed-dg-icon");
      const isOpen = rows[0]?.style.display !== "none";
      rows.forEach(r => r.style.display = isOpen ? "none" : "");
      if (icon) icon.textContent = isOpen ? "▶" : "▼";
    });
  });

  // 고정지출 날짜별 체크박스 → 카드 합계 실시간 업데이트
  sec.querySelectorAll(".mauto-fixed-chk").forEach(chk => {
    chk.addEventListener("change", () => {
      const key = chk.dataset.chkKey;
      mautoFixedChecked[key] = chk.checked;
      saveFixedChecked();
      let fixedTotal = 0;
      sec.querySelectorAll(".mauto-fixed-chk:checked").forEach(c => {
        fixedTotal += parseInt(c.dataset.amt, 10) || 0;
      });
      const cardEl = document.getElementById("mauto-fixed-card-total");
      if (cardEl) { cardEl.textContent = formatNumber(fixedTotal); cardEl.dataset.raw = fixedTotal; }
      // 예상 잔액 재계산
      const netEl = document.getElementById("mauto-net-total");
      const grid = netEl?.closest(".mauto-summary-grid");
      if (netEl && grid) {
        const fundsVal = Number(grid.querySelector(".card-funds strong")?.dataset?.raw || 0);
        const rcvVal   = Number(grid.querySelector(".card-receivable strong")?.dataset?.raw || 0);
        const payVal   = Number(grid.querySelector(".card-payable strong")?.dataset?.raw || 0);
        const net = fundsVal + rcvVal - payVal - fixedTotal;
        netEl.textContent = (net >= 0 ? "+" : "") + formatNumber(net);
        netEl.style.color = net >= 0 ? "#16a34a" : "#dc2626";
        netEl.closest(".mauto-card").style.borderTopColor = net >= 0 ? "#16a34a" : "#dc2626";
      }
    });
  });

  // 고정지출 자동계산 — 분류규칙 결제예정일 기반
  document.getElementById("mautoFixedRulesBtn")?.addEventListener("click", async () => {
    const btn = document.getElementById("mautoFixedRulesBtn");
    if (btn) { btn.textContent = "불러오는 중…"; btn.disabled = true; }
    try {
      mautoFixedRules = await fetchRulesFromApi("엠오토");
    } catch (e) {
      alert("규칙 불러오기 실패: " + e.message);
      mautoFixedRules = null;
    }
    renderMautoTab();
  });

  sec.querySelectorAll(".mauto-textarea").forEach(textarea => {
    textarea.addEventListener("paste", () => {
      const id = textarea.id.replace("mauto-textarea-", "");
      setTimeout(() => applyMautoPaste(id), 50);
    });
  });

  // 가용자금 금액 직접 수정
  sec.querySelectorAll(".mauto-inline-input[data-mauto-acc-idx]").forEach(input => {
    const idx = parseInt(input.dataset.mautoAccIdx, 10);
    input.addEventListener("focus", () => {
      input.value = String(mautoData.funds[idx]?.amount || 0);
      input.select();
    });
    input.addEventListener("blur", () => {
      const val = parseNum(input.value);
      if (mautoData.funds[idx]) mautoData.funds[idx].amount = val;
      input.value = formatNumber(val);
      saveMautoDataLocal();
      // 합계 및 요약 카드 업데이트
      const newTotal = (mautoData.funds || []).reduce((s, r) => s + Number(r.amount || 0), 0);
      const totalEl = document.getElementById("mauto-funds-total");
      if (totalEl) totalEl.textContent = formatNumber(newTotal);
      const cardEl = sec.querySelector(".card-funds strong");
      if (cardEl) cardEl.textContent = formatNumber(newTotal);
    });
    input.addEventListener("keydown", e => {
      if (e.key === "Enter") { e.preventDefault(); input.blur(); }
      if (e.key === "Escape") {
        input.value = formatNumber(mautoData.funds[idx]?.amount || 0);
        input.blur();
      }
    });
  });

  // 고정지출 예정금액 직접 수정
  sec.querySelectorAll(".mauto-fixed-amt-input").forEach(input => {
    input.addEventListener("focus", () => {
      const raw = parseNum(input.value) || 0;
      input.value = raw || "";
      input.select();
      input.style.borderBottomColor = "#2563eb";
    });
    const saveAmt = () => {
      const key = input.dataset.amtKey;
      const amt = parseInt(String(input.value).replace(/[^0-9]/g, ""), 10) || 0;
      if (amt > 0) {
        mautoFixedAmountOverrides[key] = amt;
      } else {
        delete mautoFixedAmountOverrides[key];
      }
      saveFixedAmountOverrides();
      input.value = amt ? formatNumber(amt) : "";
      input.style.borderBottomColor = amt ? "#f59e0b" : "#d1d5db";
      // 소계 행(날짜별 예정금액 합) 실시간 업데이트
      const dgKey = input.closest("[data-dg]")?.dataset.dg;
      if (dgKey) {
        const hdr = [...sec.querySelectorAll(".mauto-fixed-dg-hdr")].find(r => r.dataset.dg === dgKey);
        if (hdr) {
          let grpTotal = 0;
          sec.querySelectorAll(".mauto-fixed-dg-item").forEach(row => {
            if (row.dataset.dg === dgKey) {
              grpTotal += parseNum(row.querySelector(".mauto-fixed-amt-input")?.value || "") || 0;
            }
          });
          const cells = hdr.querySelectorAll("td");
          if (cells[3]) cells[3].textContent = grpTotal ? formatNumber(grpTotal) : "";
          const chk = hdr.querySelector(".mauto-fixed-chk");
          if (chk) chk.dataset.amt = String(grpTotal);
        }
      }
      // 카드 합계 + 예상 잔액 업데이트
      let fixedTotal = 0;
      sec.querySelectorAll(".mauto-fixed-chk:checked").forEach(c => {
        fixedTotal += parseInt(c.dataset.amt, 10) || 0;
      });
      const cardEl = document.getElementById("mauto-fixed-card-total");
      if (cardEl) { cardEl.textContent = formatNumber(fixedTotal); cardEl.dataset.raw = fixedTotal; }
      const netEl = document.getElementById("mauto-net-total");
      const grid = netEl?.closest(".mauto-summary-grid");
      if (netEl && grid) {
        const fundsVal = Number(grid.querySelector(".card-funds strong")?.dataset?.raw || 0);
        const rcvVal   = Number(grid.querySelector(".card-receivable strong")?.dataset?.raw || 0);
        const payVal   = Number(grid.querySelector(".card-payable strong")?.dataset?.raw || 0);
        const net = fundsVal + rcvVal - payVal - fixedTotal;
        netEl.textContent = (net >= 0 ? "+" : "") + formatNumber(net);
        netEl.style.color = net >= 0 ? "#16a34a" : "#dc2626";
        netEl.closest(".mauto-card").style.borderTopColor = net >= 0 ? "#16a34a" : "#dc2626";
      }
    };
    input.addEventListener("blur", saveAmt);
    input.addEventListener("keydown", e => {
      if (e.key === "Enter") { e.preventDefault(); input.blur(); }
      if (e.key === "Escape") {
        input.value = Number(input.dataset.amtOrig) ? formatNumber(Number(input.dataset.amtOrig)) : "";
        input.blur();
      }
    });
  });

  // 파일 개별 삭제
  sec.querySelectorAll(".mauto-file-del-btn").forEach(btn => {
    btn.addEventListener("click", async () => {
      const key = decodeURIComponent(btn.dataset.fkey || "");
      if (!key || !mautoSourceFiles[key]) return;
      const label = mautoSourceFiles[key].filename;
      if (!confirm(`'${label}' 파일을 삭제하고 재빌드하시겠습니까?`)) return;
      const openDetails = [...sec.querySelectorAll("details[data-accordion-id]")]
        .filter(d => d.open).map(d => d.dataset.accordionId);
      delete mautoSourceFiles[key];
      saveSourceFiles();
      if (!rulesState.rows.length) await loadRules();
      rebuildMautoRows();
      renderMautoTab();
      openDetails.forEach(id => {
        const d = sec.querySelector(`details[data-accordion-id="${id}"]`);
        if (d) d.open = true;
      });
    });
  });

  document.getElementById("mautoClearBtn")?.addEventListener("click", () => {
    if (!confirm("엠오토 데이터를 전체 초기화하시겠습니까?")) return;
    mautoData = createDefaultMautoData();
    saveMautoDataLocal();
    renderMautoTab();
  });

  // 입출금 분류 파일 선택
  document.getElementById("mautoClassifyFileInput")?.addEventListener("change", async e => {
    const file = e.target.files?.[0];
    e.target.value = "";
    if (!file) return;
    const ab = await file.arrayBuffer();
    const wb = XLSX.read(ab, { type: "array", cellDates: true });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const sheetData = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
    const bankRows = assignTxKeys(parseBankSheet(sheetData));
    if (!bankRows.length) { alert("입출금 행을 찾을 수 없습니다.\n헤더에 거래일자/입금/출금 컬럼이 있는지 확인해주세요."); return; }

    // 같은 파일명이 이미 있으면 교체 확인
    const fileKey = file.name;
    if (mautoSourceFiles[fileKey]) {
      const ok = confirm(`'${fileKey}' 파일이 이미 저장되어 있습니다.\n교체하시겠습니까?\n(이 파일의 거래 데이터가 새 내용으로 대체됩니다)`);
      if (!ok) return;
    }
    // 불변 영역(원본 거래)에 파일 단위로 저장
    mautoSourceFiles[fileKey] = { filename: fileKey, savedAt: new Date().toISOString(), rows: bankRows };
    saveSourceFiles();

    // 규칙이 없으면 먼저 로드
    if (!rulesState.rows.length) await loadRules();
    openMautoClassifyDialog(bankRows, rulesState.rows);
  });

  // 분류 결과 목록 보기 (재다이얼로그)
  document.getElementById("mautoClassifyViewBtn")?.addEventListener("click", () => {
    if (!rulesState.rows.length) { alert("분류규칙을 먼저 불러오세요."); return; }
    // 저장된 결과를 다시 볼 수 있게 빈 bankRows로 재오픈하되, 기존 결과 표시
    openMautoClassifyResultView(mautoClassifiedRows);
  });

  document.getElementById("mautoClassifyClearBtn")?.addEventListener("click", () => {
    if (!confirm("분류 결과 및 업로드된 파일 전체를 지우시겠습니까?")) return;
    mautoClassifiedRows = [];
    mautoSourceFiles    = {};
    mautoUserEdits      = {};
    saveClassifiedRows();
    saveSourceFiles();
    saveUserEdits();
    renderMautoTab();
  });

  // 세금계산서 파일 업로드 핸들러 (매출/매입 공통)
  async function handleMautoTaxFile(file, sideType) {
    const result = await parseMautoTaxInvoiceFile(file, sideType);
    if (result.error) { alert(`파싱 오류: ${result.error}`); return; }
    if (!result.rows || !result.rows.length) { alert("세금계산서 행을 찾을 수 없습니다.\n헤더가 6번째 행인지 확인해주세요."); return; }

    // 잘못된 파일 경고 (모든 거래처가 엠오토)
    if (result.allMauto) {
      const ok = confirm(`⚠ 모든 거래처가 "엠오토"입니다.\n이 파일은 타사에서 엠오토 앞으로 발행한 세금계산서로 보입니다.\n계속 저장하시겠습니까?`);
      if (!ok) return;
    }

    // 체크섬 결과 안내
    if (result.checksumOk === false) {
      const diff = result.parsedTotal - result.fileTotal;
      alert(`⚠ 체크섬 불일치!\n파일 합계: ${result.fileTotal.toLocaleString()}\n파싱 합계: ${result.parsedTotal.toLocaleString()}\n차이: ${diff.toLocaleString()}\n\n저장은 계속됩니다.`);
    }

    // 같은 파일명이 있으면 교체 확인
    if (mautoTaxSources[file.name]) {
      const ok = confirm(`'${file.name}' 파일이 이미 저장되어 있습니다.\n교체하시겠습니까?`);
      if (!ok) return;
    }

    mautoTaxSources[file.name] = {
      filename: file.name,
      sideType,
      savedAt: new Date().toISOString(),
      rows: result.rows,
      checksumOk: result.checksumOk,
    };
    saveMautoTaxSource();
    rebuildMautoTaxInvoices();
    renderMautoTab();

    const cnt = result.rows.length;
    const chk = result.checksumOk === true ? " ✓체크섬 일치" : result.checksumOk === false ? " ✗체크섬 불일치" : "";
    alert(`${sideType} 세금계산서 ${cnt}건 저장됨${chk}`);
  }

  document.getElementById("mautoTaxSalesFileInput")?.addEventListener("change", async e => {
    const file = e.target.files?.[0]; e.target.value = "";
    if (file) await handleMautoTaxFile(file, "매출");
  });
  document.getElementById("mautoTaxPurchaseFileInput")?.addEventListener("change", async e => {
    const file = e.target.files?.[0]; e.target.value = "";
    if (file) await handleMautoTaxFile(file, "매입");
  });

  // 세금계산서 파일 삭제
  sec.querySelectorAll(".mauto-tax-del-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      const fname = decodeURIComponent(btn.dataset.fname);
      if (!confirm(`'${fname}' 파일을 삭제하시겠습니까?`)) return;
      // 삭제 전 아코디언 열림 상태 저장
      const openDetails = [...sec.querySelectorAll("details[data-accordion-id]")]
        .filter(d => d.open).map(d => d.dataset.accordionId);
      delete mautoTaxSources[fname];
      saveMautoTaxSource();
      rebuildMautoTaxInvoices();
      renderMautoTab();
      // 열림 상태 복원
      openDetails.forEach(id => {
        const d = sec.querySelector(`details[data-accordion-id="${id}"]`);
        if (d) d.open = true;
      });
    });
  });

  document.getElementById("mautoToolsToggle")?.addEventListener("click", e => {
    mautoToolsOpen = !mautoToolsOpen;
    const panel = document.getElementById("mautoToolsPanel");
    if (panel) panel.style.display = mautoToolsOpen ? "block" : "none";
    e.currentTarget.textContent = mautoToolsOpen ? "▲ 도구 접기" : "▼ 입력·업로드 도구";
  });
  document.getElementById("mautoVatBtn")?.addEventListener("click", () => {
    mautoVatView = !mautoVatView;
    renderMautoTab();
  });
  document.getElementById("mautoVatModeFilter")?.addEventListener("change", e => {
    mautoVatMode = e.target.value;
    renderMautoTab();
  });
  document.getElementById("mautoVatYearFilter")?.addEventListener("change", e => {
    mautoVatYear = Number(e.target.value);
    renderMautoTab();
  });

  setupMautoToggleHandlers(sec);

  // 기본 전체 접기
  sec.querySelectorAll(".mauto-toggle-year, .mauto-toggle-fxdate").forEach(row => {
    row.dataset.collapsed = "1";
    const icon = row.querySelector(".mauto-toggle-icon");
    if (icon) icon.textContent = "▶";
  });
  sec.querySelectorAll(".mauto-toggle-month, [data-mauto-yr], [data-mauto-mo], [data-mauto-fxdate]:not(.mauto-toggle-fxdate)").forEach(r => {
    r.style.display = "none";
  });
}

function applyMautoPaste(sectionId) {
  const textarea = document.getElementById(`mauto-textarea-${sectionId}`);
  if (!textarea) return;
  const text = textarea.value.trim();
  if (!text) return;
  let parsed = null;
  if (sectionId === "funds") {
    parsed = parseMautoFundsPaste(text);
    if (!parsed) { alert("가용자금 금액을 읽지 못했습니다.\n금액만 2줄로 붙여넣거나, 은행/금액 표를 붙여넣어 주세요."); return; }
    mautoData.funds = parsed;
  } else if (sectionId === "receivables") {
    parsed = parseMautoAccountingPaste(text, "receivables");
    if (!parsed.length) { alert("미수금 데이터를 읽지 못했습니다.\n헤더를 확인해 주세요."); return; }
    mautoData.receivables = parsed;
  } else if (sectionId === "payables") {
    parsed = parseMautoAccountingPaste(text, "payables");
    if (!parsed.length) { alert("미지급 데이터를 읽지 못했습니다.\n헤더를 확인해 주세요."); return; }
    mautoData.payables = parsed;
  } else if (sectionId === "fixed") {
    parsed = parseMautoFixedPaste(text);
    if (!parsed.length) { alert("고정지출 데이터를 읽지 못했습니다.\n헤더를 확인해 주세요."); return; }
    mautoData.fixed = parsed;
  }
  saveMautoDataLocal();
  renderMautoTab();
}

function setupMautoToggleHandlers(container) {
  container.querySelectorAll("[data-mauto-expand-all]").forEach(btn => {
    btn.addEventListener("click", () => {
      const sec = container.querySelector(`#mauto-section-${btn.dataset.mautoExpandAll}`);
      if (!sec) return;
      sec.querySelectorAll(".mauto-toggle-year, .mauto-toggle-month, .mauto-toggle-fxdate").forEach(row => {
        row.dataset.collapsed = "0";
        const icon = row.querySelector(".mauto-toggle-icon");
        if (icon) icon.textContent = "▼";
      });
      sec.querySelectorAll("[data-mauto-yr], [data-mauto-mo], [data-mauto-fxdate]:not(.mauto-toggle-fxdate)").forEach(r => {
        r.style.display = "";
      });
    });
  });

  container.querySelectorAll("[data-mauto-collapse-all]").forEach(btn => {
    btn.addEventListener("click", () => {
      const sec = container.querySelector(`#mauto-section-${btn.dataset.mautoCollapseAll}`);
      if (!sec) return;
      sec.querySelectorAll(".mauto-toggle-year").forEach(row => {
        row.dataset.collapsed = "1";
        const icon = row.querySelector(".mauto-toggle-icon");
        if (icon) icon.textContent = "▶";
      });
      sec.querySelectorAll(".mauto-toggle-month, [data-mauto-yr], [data-mauto-mo]").forEach(r => {
        r.style.display = "none";
      });
      sec.querySelectorAll(".mauto-toggle-fxdate").forEach(row => {
        row.dataset.collapsed = "1";
        const icon = row.querySelector(".mauto-toggle-icon");
        if (icon) icon.textContent = "▶";
      });
      sec.querySelectorAll("[data-mauto-fxdate]:not(.mauto-toggle-fxdate)").forEach(r => {
        r.style.display = "none";
      });
    });
  });

  container.querySelectorAll(".mauto-toggle-year").forEach(row => {
    row.addEventListener("click", () => {
      const yk = row.dataset.mautoYear;
      const collapsed = row.dataset.collapsed === "1";
      container.querySelectorAll(`[data-mauto-yr="${yk}"], .mauto-toggle-month[data-mauto-year="${yk}"]`)
        .forEach(r => r.style.display = collapsed ? "" : "none");
      row.dataset.collapsed = collapsed ? "0" : "1";
      const icon = row.querySelector(".mauto-toggle-icon");
      if (icon) icon.textContent = collapsed ? "▼" : "▶";
    });
  });

  container.querySelectorAll(".mauto-toggle-month").forEach(row => {
    row.addEventListener("click", e => {
      e.stopPropagation();
      const mk = row.dataset.mautoMonth;
      const collapsed = row.dataset.collapsed === "1";
      container.querySelectorAll(`[data-mauto-mo="${mk}"]`)
        .forEach(r => r.style.display = collapsed ? "" : "none");
      row.dataset.collapsed = collapsed ? "0" : "1";
      const icon = row.querySelector(".mauto-toggle-icon");
      if (icon) icon.textContent = collapsed ? "▼" : "▶";
    });
  });

  container.querySelectorAll(".mauto-toggle-fxdate").forEach(row => {
    row.addEventListener("click", () => {
      const dk = row.dataset.mautoFxdate;
      const collapsed = row.dataset.collapsed === "1";
      container.querySelectorAll(`tr[data-mauto-fxdate="${dk}"]:not(.mauto-toggle-fxdate)`)
        .forEach(r => r.style.display = collapsed ? "" : "none");
      row.dataset.collapsed = collapsed ? "0" : "1";
      const icon = row.querySelector(".mauto-toggle-icon");
      if (icon) icon.textContent = collapsed ? "▼" : "▶";
    });
  });
}

function setupTabs() {
  elements.tabButtons.forEach(button => {
    button.addEventListener("click", () => {
      const target = button.dataset.tab;
      elements.tabButtons.forEach(btn => btn.classList.toggle("active", btn === button));
      document.querySelectorAll(".tab-content").forEach(section => {
        section.classList.toggle("active", section.id === target);
      });
      updateFilterBarVisibility(target);
      if (target === "daesa") renderDaesaTab();
      if (target === "fixed") renderFixedExpenses();
      if (target === "pnl") { renderPnlTab(); loadPnlRemote(); }
      if (target === "mauto") {
        renderMautoTab();
        // 고정지출 분류규칙 자동 로드 (탭 첫 진입 시)
        if (mautoFixedRules === null && SHEET_APP_SCRIPT_URL) {
          fetchRulesFromApi("엠오토").then(rules => {
            mautoFixedRules = rules;
            renderMautoTab();
          }).catch(() => { mautoFixedRules = []; });
        }
        console.log("[엠오토] 탭 진입 — 원격 로드 시작 (setupTabs)");
        if (SHEET_APP_SCRIPT_URL) loadMautoTaxRemote();
        if (SHEET_APP_SCRIPT_URL) loadMautoSourceRemote();
        // 입출금 분류 원격 로드 (setupTabs 경로)
        if (SHEET_APP_SCRIPT_URL) {
          fetchSheetWebApp({ action: "getClassifiedRows" }).then(res => {
            const remote = (res && (res.rows || res.data)) || [];
            if (!remote.length) return;
            const localMap = new Map(mautoClassifiedRows.map(r => [r._txKey, r]));
            let added = 0;
            remote.forEach(r => {
              if (!r._txKey) return;
              if (!localMap.has(r._txKey)) { localMap.set(r._txKey, r); added++; }
              else {
                const local = localMap.get(r._txKey);
                if (!local.savedAt || (r.savedAt && r.savedAt > local.savedAt)) {
                  localMap.set(r._txKey, { ...local, 거래처명: r.거래처명, 구분: r.구분, excluded: r.excluded, 매칭근거: r.매칭근거, savedAt: r.savedAt });
                }
              }
            });
            if (added > 0 || remote.length) {
              mautoClassifiedRows = [...localMap.values()].sort((a,b) => (a.date||"") < (b.date||"") ? -1 : 1);
              try { localStorage.setItem(MAUTO_CLASSIFIED_KEY, JSON.stringify(mautoClassifiedRows)); } catch(_) {}
              renderMautoTab();
            }
          }).catch(() => {});
        }
        if (SHEET_APP_SCRIPT_URL) {
          loadMautoDataRemote().then(remote => {
            console.log("[엠오토] 원격 응답:", remote === null ? "null" : JSON.stringify(remote).slice(0, 300));
            if (!remote) {
              const hasLocalData = mautoData.funds.some(f => (f.amount || 0) > 0) ||
                mautoData.receivables.length > 0 || mautoData.payables.length > 0 || mautoData.fixed.length > 0;
              if (hasLocalData) _scheduleMautoRemoteSave();
              return;
            }
            if (Array.isArray(remote.excludeRcv)) { mautoExcludeVendorsRcv = remote.excludeRcv; try { localStorage.setItem(MAUTO_EXCLUDE_KEY_RCV, JSON.stringify(mautoExcludeVendorsRcv)); } catch (_) {} }
            if (Array.isArray(remote.excludePay)) { mautoExcludeVendorsPay = remote.excludePay; try { localStorage.setItem(MAUTO_EXCLUDE_KEY_PAY, JSON.stringify(mautoExcludeVendorsPay)); } catch (_) {} }
            if (remote.fixedChecked && typeof remote.fixedChecked === "object") { mautoFixedChecked = remote.fixedChecked; try { localStorage.setItem(MAUTO_FIXED_CHECKED_KEY, JSON.stringify(mautoFixedChecked)); } catch (_) {} }
            if (remote.fixedAmountOverrides && typeof remote.fixedAmountOverrides === "object") { mautoFixedAmountOverrides = remote.fixedAmountOverrides; try { localStorage.setItem(MAUTO_FIXED_AMOUNT_KEY, JSON.stringify(mautoFixedAmountOverrides)); } catch (_) {} }
            mautoData = normalizeMautoData(remote);
            console.log("[엠오토] 정규화 후 funds:", JSON.stringify(mautoData.funds));
            try { localStorage.setItem(MAUTO_LOCAL_KEY, JSON.stringify(mautoData)); } catch (_) {}
            renderMautoTab();
          }).catch(e => console.warn("[엠오토] 원격 로드 실패:", e));
        }
      }
    });
  });
}

// ── 4단계: 은행 입출금 매칭 ─────────────────────────────────

function parseBankAmount(val) {
  return Math.abs(Number(String(val ?? "").replace(/[^0-9.-]/g, "")) || 0);
}

function extractYearMonth(text) {
  const s = String(text || "");
  // "26-03", "2603", "260301", "26.03" 패턴
  const m = s.match(/(\d{2})[-./]?(\d{2})/);
  if (!m) return null;
  const y = 2000 + parseInt(m[1]), mo = parseInt(m[2]);
  if (mo < 1 || mo > 12) return null;
  return { year: y, month: mo };
}

function extractPartialFlag(text) {
  const s = String(text || "");
  if (/일부|부분|선금|선불/.test(s)) return "partial";
  if (/나머지|잔금|잔액|완료/.test(s)) return "remainder";
  return "full";
}

function vendorNameSimilarity(a, b) {
  const clean = s => String(s || "").replace(/[\s(주)(유)㈜]/g, "").toLowerCase();
  const ca = clean(a), cb = clean(b);
  if (!ca || !cb) return 0;
  if (ca === cb) return 1;
  if (ca.includes(cb) || cb.includes(ca)) return 0.85;
  // 앞 2글자 공통
  if (ca.slice(0, 2) === cb.slice(0, 2)) return 0.6;
  return 0;
}

function matchBankRowToPayables(bankRow, allPayables) {
  // _memo / _debit 등 정규화 필드 우선, 구버전 원본 필드 폴백
  const memo = bankRow._memo || String(bankRow.memo || bankRow["적요1"] || bankRow["적요"] || bankRow["내용"] || "");
  const memo2 = bankRow._memo2 || String(bankRow["비고"] || bankRow["적요2"] || "");
  const combinedMemo = (memo + " " + memo2).trim();
  const amount = bankRow._debit || parseBankAmount(bankRow["출금"] || bankRow["금액"] || bankRow.amount || 0) ||
    bankRow._credit || parseBankAmount(bankRow["입금"] || 0);
  const ym = extractYearMonth(combinedMemo);
  const partial = extractPartialFlag(combinedMemo);

  const candidates = allPayables
    .filter(p => p.completionStatus !== "완료")
    .map(p => {
      let score = 0;
      const nameSim = vendorNameSimilarity(combinedMemo, p.name);
      score += nameSim * 50;
      if (ym && p.year === ym.year && p.month === ym.month) score += 30;
      else if (ym && (p.year === ym.year || p.month === ym.month)) score += 10;
      const outstanding = getPayableOutstanding(p);
      if (amount && outstanding) {
        const ratio = Math.min(amount, outstanding) / Math.max(amount, outstanding);
        score += ratio * 20;
      }
      return { item: p, score, nameSim, ym, partial, amount };
    })
    .filter(c => c.score > 20)
    .sort((a, b) => b.score - a.score);

  return candidates.slice(0, 3);
}

function parseBankSheet(sheetData) {
  if (!sheetData || sheetData.length < 2) return [];

  // 헤더 행 자동 감지: "거래일자" 또는 "날짜" 또는 "거래일" 포함 행을 찾음
  let headerRowIdx = 0;
  for (let i = 0; i < Math.min(sheetData.length, 10); i++) {
    const row = sheetData[i].map(c => String(c).trim());
    if (row.some(c => /거래일자|거래일|날짜/.test(c))) {
      headerRowIdx = i;
      break;
    }
  }

  const headers = sheetData[headerRowIdx].map(h => String(h).trim());
  return sheetData.slice(headerRowIdx + 1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i] ?? ""; });

    // 정규화: 통일된 필드명으로 접근하기 쉽게
    obj._date = String(obj["거래일자"] || obj["거래일"] || obj["날짜"] || obj["일자"] || "").trim();
    obj._time = String(obj["거래시간"] || obj["시간"] || "").trim();
    obj._bank = String(obj["은행"] || "").trim();
    obj._memo = String(obj["적요1"] || obj["적요"] || obj["내용"] || obj["메모"] || "").trim();
    obj._memo2 = String(obj["비고"] || obj["적요2"] || "").trim();
    obj._credit = parseBankAmount(obj["입금"] || 0);   // 들어온 돈
    obj._debit = parseBankAmount(obj["출금"] || 0);   // 나간 돈
    obj._balance = parseBankAmount(obj["잔액"] || 0);
    obj._account = String(obj["계좌번호"] || "").trim();
    obj._alias = String(obj["계좌별칭"] || "").trim();
    obj._branch = String(obj["취급지점"] || obj["취급점"] || "").trim();
    return obj;
  }).filter(r =>
    /^\d{4}/.test(r._date) &&        // 거래일자가 연도(숫자 4자리)로 시작해야 유효 — 합계/이월/소계 행 제외
    (r._debit > 0 || r._credit > 0)  // 입금·출금 둘 다 0인 행 제외
  );
}

// ────────────────────────────────────────────────────────────
//  Phase 2: 엠오토 입출금 분류 (규칙 기반 3단계 매칭)
// ────────────────────────────────────────────────────────────

// 거래키 부여: _time 있으면 날짜+시간+금액+적요, 없으면 키+배치내 발생순번
function assignTxKeys(bankRows) {
  const seqMap = {};
  return bankRows.map(row => {
    const time = String(row._time || "").trim();
    const base = [row._date, time, row._credit || 0, row._debit || 0, row._memo].join("|");
    // 시간 유무와 관계없이 항상 배치 내 발생순번 부여 → 같은 날 같은 금액 중복 처리 방지
    seqMap[base] = (seqMap[base] || 0) + 1;
    return { ...row, _txKey: `${base}#${seqMap[base]}` };
  });
}

// 기존 저장 데이터와 신규 업로드 merge (거래키 기준, 충돌 시 기존 유지)
function mergeClassifiedRows(existing, incoming) {
  const map = new Map(existing.map(r => [r._txKey, r]));
  let skipped = 0;
  for (const row of incoming) {
    if (map.has(row._txKey)) {
      skipped++;
    } else {
      map.set(row._txKey, row);
    }
  }
  return { merged: [...map.values()], skipped };
}

// ── 전체 재빌드: source-files 전체 → 분류 → 사용자편집 재적용 ──────────────
function rebuildMautoRows() {
  const rules = (rulesState.rows || []).filter(r => String(r["사업체"] || "") === "엠오토");

  // ⚠️ 분류규칙(rulesState.rows)이 아직 로드되지 않은 상태(예: 엠오토 탭만 진입한 모바일)에서
  // 재분류하면 규칙 매칭 행들의 거래처명·구분이 전부 빈값이 되어, 서버에서 불러온
  // 올바른 분류 결과(mautoClassifiedRows)를 통째로 덮어써 미수/미지급/고정지출이 깨진다.
  // 규칙이 하나도 없으면 재빌드를 건너뛰어 기존 분류 결과를 보존한다.
  // (규칙이 없으면 어차피 재분류해도 전부 미분류이므로 손실 없음)
  if (!rules.length) return;

  // 1. 모든 파일 행 합치기 (언더바 금액필드가 비었으면 _txKey에서 복구 → 재빌드 금액 0 방지)
  let allRows = Object.values(mautoSourceFiles).flatMap(f => f.rows || []).map(normalizeSourceRow);

  // 2. 거래키 기준 중복 제거 (첫 번째 발견 우선)
  const seen = new Map();
  for (const row of allRows) {
    if (row._txKey && !seen.has(row._txKey)) seen.set(row._txKey, row);
  }
  allRows = [...seen.values()];

  // 3. 날짜+시간 정렬
  allRows.sort((a, b) => {
    const ka = String(a._date || "") + "|" + String(a._time || "");
    const kb = String(b._date || "") + "|" + String(b._time || "");
    return ka.localeCompare(kb);
  });

  // 4. 자동분류 후 사용자 편집 덮어쓰기
  mautoClassifiedRows = allRows.map(row => {
    const match = classifyBankRow(row, rules);
    const base = {
      _txKey:  row._txKey,
      date:    row._date  || "",
      time:    row._time  || "",
      _memo:   row._memo  || "",
      _memo2:  row._memo2 || "",
      memo:    [row._memo, row._memo2].filter(Boolean).join(" / "),
      credit:  row._credit  || 0,
      debit:   row._debit   || 0,
      거래처명: match.거래처 || "",
      구분:    match.구분   || "",
      excluded:   false,
      매칭근거: match.매칭근거,
      isOverride: false,
    };
    const edit = mautoUserEdits[row._txKey];
    if (edit) {
      if (edit.거래처명 !== undefined) base.거래처명  = edit.거래처명;
      if (edit.구분     !== undefined) base.구분      = edit.구분;
      if (edit.excluded !== undefined) base.excluded  = edit.excluded;
      if (edit.isOverride) { base.isOverride = true; base.매칭근거 = edit.매칭근거 || "수동"; }
      // 비고 수동수정(여러 연월 분배 "25-12=..." 등) → 귀속연월 계산에 쓰는 _memo2 덮어씀
      if (edit.memoOverride !== undefined && edit.memoOverride !== "") {
        base._memo2 = edit.memoOverride;
        base.memo = edit.memoOverride;
      }
    }
    return base;
  });
  saveClassifiedRows();
}

// ── 레거시 마이그레이션: mauto-classified-rows-v1 → source+edits 분리 ────────
function migrateLegacyIfNeeded() {
  if (Object.keys(mautoSourceFiles).length > 0) return;
  const legacyStr = localStorage.getItem(MAUTO_CLASSIFIED_KEY);
  if (!legacyStr) return;
  let legacy;
  try { legacy = JSON.parse(legacyStr); } catch { return; }
  if (!Array.isArray(legacy) || !legacy.length) return;

  // 원본 거래만 추출 → source-files
  const migRows = legacy
    .filter(r => r._txKey && (Number(r.credit) || Number(r.debit)))
    .map(r => ({
      _txKey:   r._txKey,
      _date:    r.date   || "",
      _time:    r.time   || "",
      _memo:    r._memo  || "",
      _memo2:   r._memo2 || "",
      _credit:  Number(r.credit) || 0,
      _debit:   Number(r.debit)  || 0,
      _account: r._account || "",
      _bank:    r._bank    || "",
    }));
  if (migRows.length) {
    mautoSourceFiles["기존데이터(마이그레이션)"] = {
      filename: "기존데이터(마이그레이션)",
      savedAt: new Date().toISOString(),
      isMigration: true,
      rows: migRows,
    };
    saveSourceFiles();
  }

  // 사용자 편집 추출 → user-edits
  for (const r of legacy) {
    if (!r._txKey) continue;
    if (r.excluded || r.거래처명) {
      mautoUserEdits[r._txKey] = {
        거래처명:  r.거래처명 || "",
        구분:     r.구분    || "",
        excluded: !!r.excluded,
        isOverride: !!(r.거래처명),
        매칭근거: r.매칭근거 || "",
      };
    }
  }
  saveUserEdits();
}

function classifyBankRow(row, rules) {
  // 레거시 순서: ①적요1 정확일치(계좌) → ②비고 부분포함(거래처명) → ③적요1 부분포함(키워드)
  const 적요1 = String(row._memo  || "").trim();
  const 비고  = String(row._memo2 || "").trim();
  const contains = (hay, key) => key.length >= 2 && hay.toLowerCase().includes(key.toLowerCase());
  const hit = (r, how) => ({ 거래처: r["거래처명"], 구분: r["구분"] || "", 매칭근거: `${how}:${r["매칭키"]}`, rule: r });

  const 계좌R  = rules.filter(r => String(r["매칭방식"] || "") === "계좌");
  const 비고R  = rules.filter(r => String(r["매칭방식"] || "") === "거래처명");
  const 적요R  = rules.filter(r => String(r["매칭방식"] || "") === "키워드");

  // ① 적요1 정확일치 (빈 적요1은 건너뜀)
  for (const r of 계좌R) {
    const key = String(r["매칭키"] || "").trim();
    if (적요1 && 적요1 === key) return hit(r, "계좌-정확");
  }
  // ② 비고 부분포함
  for (const r of 비고R) if (contains(비고,  String(r["매칭키"] || ""))) return hit(r, "비고-부분");
  // ③ 적요1 부분포함
  for (const r of 적요R) if (contains(적요1, String(r["매칭키"] || ""))) return hit(r, "적요1-부분");

  return { 거래처: null, 구분: "", 매칭근거: "미매칭", rule: null };
}

function openMautoClassifyDialog(bankRows, rules) {
  document.querySelector(".mauto-classify-overlay")?.remove();

  const 엠오토규칙 = rules.filter(r => String(r["사업체"] || "") === "엠오토");
  const vendorNames = [...new Set(엠오토규칙.map(r => String(r["거래처명"] || "")).filter(Boolean))].sort();

  const items = bankRows.map(row => {
    const match = classifyBankRow(row, 엠오토규칙);
    return { row, match, 거래처명: match?.거래처 || "", 구분: match?.구분 || "", excluded: false,
      isOverride: false, ruleAdd: false, ruleMethod: "키워드", ruleKey: "" };
  });

  const overlay = document.createElement("div");
  overlay.className = "mauto-classify-overlay bank-match-overlay";

  function buildTableHtml() {
    return items.map((item, idx) => {
      const row = item.row;
      const isCredit = (row._credit || 0) > 0;
      const amount = row._credit || row._debit || 0;
      const memo = [row._memo, row._memo2].filter(Boolean).join(" / ");
      const dirLabel = isCredit
        ? `<span style="color:#1565c0;font-size:11px;font-weight:bold;">입금</span>`
        : `<span style="color:#b71c1c;font-size:11px;font-weight:bold;">출금</span>`;
      const matched = item.match?.거래처 !== null;
      const matchBadge = matched
        ? `<span style="color:#166534;font-size:10px;">${escapeHtml(item.match.매칭근거)}</span>`
        : `<span style="color:#9ca3af;font-size:10px;">미분류</span>`;
      const rowBg = item.excluded ? "opacity:0.4;" : !matched ? "background:#fff7ed;" : "";
      const vendorOpts = `<option value="">-- 선택 --</option>` +
        vendorNames.map(n => `<option value="${escapeHtml(n)}" ${item.거래처명 === n ? "selected" : ""}>${escapeHtml(n)}</option>`).join("");
      const divOpts = `<option value="">-</option>
        <option value="매출" ${item.구분 === "매출" ? "selected" : ""}>매출</option>
        <option value="매입" ${item.구분 === "매입" ? "selected" : ""}>매입</option>`;
      const ruleRowHidden = item.거래처명 ? "" : "display:none;";
      const ruleDetailHidden = item.ruleAdd ? "" : "display:none;";
      const memoHint = escapeHtml((row._memo || "").slice(0, 25));
      const memo2Hint = row._memo2 ? ` | 비고: ${escapeHtml(row._memo2.slice(0, 15))}` : "";
      return `<tr style="${rowBg}">
        <td style="font-size:12px;white-space:nowrap;">${escapeHtml(row._date)}</td>
        <td>${dirLabel}</td>
        <td style="text-align:right;font-size:12px;font-weight:500;">${formatNumber(amount)}</td>
        <td style="font-size:12px;max-width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="${escapeHtml(memo)}">${escapeHtml(memo.slice(0, 22))}${memo.length > 22 ? "…" : ""}</td>
        <td>${matchBadge}</td>
        <td><select class="mcl-vendor" data-idx="${idx}" style="font-size:12px;max-width:130px;">${vendorOpts}</select></td>
        <td><select class="mcl-div" data-idx="${idx}" style="font-size:12px;">${divOpts}</select></td>
        <td><input type="checkbox" class="mcl-exclude" data-idx="${idx}" ${item.excluded ? "checked" : ""} /></td>
      </tr>
      <tr class="mcl-rule-row" data-idx="${idx}" style="${ruleRowHidden}">
        <td colspan="8" style="padding:2px 12px 6px;background:#f8fafc;border-bottom:1px solid #e2e8f0;">
          <div style="display:flex;align-items:center;gap:8px;font-size:12px;flex-wrap:wrap;">
            <label style="display:flex;align-items:center;gap:4px;cursor:pointer;color:#374151;white-space:nowrap;">
              <input type="checkbox" class="mcl-rule-add" data-idx="${idx}" ${item.ruleAdd ? "checked" : ""}> 분류규칙에 추가
            </label>
            <div class="mcl-rule-detail" data-idx="${idx}" style="${ruleDetailHidden}display:flex;align-items:center;gap:6px;flex-wrap:wrap;">
              <span style="color:#9ca3af;font-size:11px;" title="적요: ${escapeHtml(row._memo || "")} / 비고: ${escapeHtml(row._memo2 || "")}">적요: ${memoHint}${memo2Hint}</span>
              <select class="mcl-rule-method" data-idx="${idx}" style="font-size:12px;">
                <option value="키워드" ${item.ruleMethod==="키워드"?"selected":""}>키워드(적요)</option>
                <option value="거래처명" ${item.ruleMethod==="거래처명"?"selected":""}>거래처명(비고)</option>
                <option value="계좌" ${item.ruleMethod==="계좌"?"selected":""}>계좌(정확)</option>
              </select>
              <input type="text" class="mcl-rule-key" data-idx="${idx}" value="${escapeHtml(item.ruleKey)}" placeholder="매칭키 2자 이상" style="font-size:12px;width:110px;padding:2px 6px;border:1px solid #d1d5db;border-radius:4px;">
              <span class="mcl-rule-preview" data-idx="${idx}" style="font-size:11px;min-width:36px;"></span>
            </div>
          </div>
        </td>
      </tr>`;
    }).join("");
  }

  const matchedCount = items.filter(i => i.match?.거래처 !== null).length;
  overlay.innerHTML = `
    <div class="bank-match-dialog" style="max-width:940px;">
      <div class="bank-match-header">
        <h3>엠오토 입출금 분류</h3>
        <span class="bank-match-sub">${bankRows.length}건 · 자동분류 ${matchedCount}건 / 미분류 ${bankRows.length - matchedCount}건</span>
        <button type="button" class="bank-match-close">✕</button>
      </div>
      <div class="table-responsive bank-match-table-wrap">
        <table class="bank-match-table">
          <thead><tr>
            <th>날짜</th><th>구분</th><th>금액</th><th>적요</th><th>매칭근거</th><th>거래처명</th><th>매출/매입</th><th style="white-space:nowrap;"><input type="checkbox" id="mclExcludeAll" title="전체 제외 토글" style="margin-right:3px;">제외</th>
          </tr></thead>
          <tbody>${buildTableHtml()}</tbody>
        </table>
      </div>
      <div class="bank-match-actions">
        <span id="mclCount" class="bank-match-count"></span>
        <button type="button" class="bank-apply-btn" id="mclApplyBtn">분류 결과 저장</button>
        <button type="button" class="bank-cancel-btn">닫기</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);

  function updateCount() {
    const n = items.filter(i => !i.excluded && i.거래처명).length;
    overlay.querySelector("#mclCount").textContent = `분류 ${n}건`;
  }
  updateCount();

  // 마스터 체크박스 3-상태 갱신
  function updateMasterCheckbox() {
    const master = overlay.querySelector("#mclExcludeAll");
    if (!master) return;
    const exclCount = items.filter(i => i.excluded).length;
    if (exclCount === 0) {
      master.checked = false;
      master.indeterminate = false;
    } else if (exclCount === items.length) {
      master.checked = true;
      master.indeterminate = false;
    } else {
      master.checked = false;
      master.indeterminate = true;
    }
  }

  // 미리보기: 이번 업로드 bankRows 기준으로 매칭방식별 건수 계산
  function updateRulePreview(idx) {
    const item = items[idx];
    const key = (item.ruleKey || "").trim();
    const method = item.ruleMethod || "키워드";
    const previewEl = overlay.querySelector(`.mcl-rule-preview[data-idx="${idx}"]`);
    if (!previewEl) return;
    if (key.length < 2) {
      previewEl.textContent = key.length > 0 ? "2자 이상" : "";
      previewEl.style.color = "#ef4444";
      return;
    }
    const keyLow = key.toLowerCase();
    let count = 0;
    for (const r of bankRows) {
      if (method === "계좌") {
        if (String(r._memo || "").trim() === key) count++;
      } else if (method === "거래처명") {
        if (String(r._memo2 || "").toLowerCase().includes(keyLow)) count++;
      } else {
        if (String(r._memo || "").toLowerCase().includes(keyLow)) count++;
      }
    }
    const isHigh = count > 3 && count / bankRows.length > 0.3;
    previewEl.textContent = `${count}건`;
    previewEl.style.color = isHigh ? "#d97706" : "#166534";
    previewEl.title = isHigh ? "과다매칭 주의" : "";
  }

  overlay.querySelector(".bank-match-close").addEventListener("click", () => overlay.remove());
  overlay.querySelector(".bank-cancel-btn").addEventListener("click", () => overlay.remove());
  overlay.querySelectorAll(".mcl-vendor").forEach(sel =>
    sel.addEventListener("change", () => {
      const idx = +sel.dataset.idx;
      items[idx].거래처명 = sel.value;
      items[idx].isOverride = true;
      const ruleRow = overlay.querySelector(`.mcl-rule-row[data-idx="${idx}"]`);
      if (ruleRow) ruleRow.style.display = sel.value ? "" : "none";
      updateCount();
    }));
  overlay.querySelectorAll(".mcl-div").forEach(sel =>
    sel.addEventListener("change", () => { items[+sel.dataset.idx].구분 = sel.value; }));

  // 규칙 추가 체크박스 → 상세 영역 펼침
  overlay.querySelectorAll(".mcl-rule-add").forEach(chk =>
    chk.addEventListener("change", () => {
      const idx = +chk.dataset.idx;
      items[idx].ruleAdd = chk.checked;
      const detail = overlay.querySelector(`.mcl-rule-detail[data-idx="${idx}"]`);
      if (detail) detail.style.display = chk.checked ? "flex" : "none";
      if (chk.checked) updateRulePreview(idx);
    }));
  overlay.querySelectorAll(".mcl-rule-method").forEach(sel =>
    sel.addEventListener("change", () => {
      const idx = +sel.dataset.idx;
      items[idx].ruleMethod = sel.value;
      updateRulePreview(idx);
    }));
  overlay.querySelectorAll(".mcl-rule-key").forEach(inp =>
    inp.addEventListener("input", () => {
      const idx = +inp.dataset.idx;
      items[idx].ruleKey = inp.value;
      updateRulePreview(idx);
    }));
  overlay.querySelectorAll(".mcl-exclude").forEach(chk =>
    chk.addEventListener("change", () => {
      items[+chk.dataset.idx].excluded = chk.checked;
      updateCount();
      updateMasterCheckbox();
    }));

  // 마스터 체크박스 → 전체 토글 (보이는 행 = items 전체)
  overlay.querySelector("#mclExcludeAll")?.addEventListener("change", e => {
    const val = e.target.checked;
    items.forEach(i => { i.excluded = val; });
    overlay.querySelectorAll(".mcl-exclude").forEach(chk => { chk.checked = val; });
    // indeterminate 해제 (마스터 직접 클릭 시 명확한 상태로)
    e.target.indeterminate = false;
    updateCount();
  });

  overlay.querySelector("#mclApplyBtn").addEventListener("click", async () => {
    // 사용자가 수동으로 바꾼 행(isOverride) 또는 제외한 행만 user-edits에 저장
    for (const i of items) {
      const txKey = i.row._txKey;
      if (!txKey) continue;
      if (i.excluded || i.isOverride) {
        mautoUserEdits[txKey] = {
          거래처명:  i.거래처명 || "",
          구분:     i.구분    || "",
          excluded: !!i.excluded,
          isOverride: !!i.isOverride,
          매칭근거: i.isOverride ? "수동" : (i.match?.매칭근거 || ""),
        };
      }
    }
    saveUserEdits();
    // 불변(source) + 사용자편집 → 전체 재빌드 (규칙 없으면 재빌드가 스킵되므로 먼저 로드)
    if (!rulesState.rows.length) await loadRules();
    rebuildMautoRows();
    overlay.remove();

    // 규칙 학습: ruleAdd 켜져 있고 키 2자 이상인 행만
    const ruleItems = items.filter(i => i.ruleAdd && (i.ruleKey || "").trim().length >= 2 && i.거래처명);
    if (ruleItems.length) {
      const newRules = [];
      for (const item of ruleItems) {
        const key = item.ruleKey.trim();
        const existingRule = rulesState.rows.find(
          r => r["_rule_key"] === buildRuleKey("엠오토", item.ruleMethod, key));
        if (existingRule && existingRule["거래처명"] !== item.거래처명) {
          const ok = confirm(`이미 '${existingRule["거래처명"]}'으로 매핑된 규칙입니다.\n'${item.거래처명}'으로 바꿀까요?`);
          if (!ok) continue;
        }
        newRules.push({ 사업체: "엠오토", 매칭방식: item.ruleMethod, 매칭키: key,
          거래처명: item.거래처명, 구분: item.구분 || "", 우선순위: "" });
      }
      if (newRules.length) {
        try {
          await postSheetWebApp("upsertRules", { rows: newRules });
          await loadRules();
          // 새 규칙으로 전체 재빌드 (미매칭 자동 재분류 포함)
          rebuildMautoRows();
        } catch (_) {
          alert("분류는 저장됨. 규칙 추가는 실패 — 다시 시도해주세요.");
        }
      }
    }

    renderMautoTab();
  });
}

// 새 규칙을 localStorage 보관분 중 미매칭(excluded=false, 거래처명="") 행에 즉시 적용
// 이미 분류/제외된 행은 건드리지 않음 (사용자 결정 보호)
function applyNewRulesToUnmatched(updatedRules) {
  const 엠오토규칙 = updatedRules.filter(r => String(r["사업체"] || "") === "엠오토");
  let changed = 0;
  for (const saved of mautoClassifiedRows) {
    if (saved.excluded || saved.거래처명) continue;
    const fakeRow = { _memo: saved._memo || saved.memo || "", _memo2: saved._memo2 || "" };
    const match = classifyBankRow(fakeRow, 엠오토규칙);
    if (match.거래처) {
      saved.거래처명 = match.거래처;
      saved.구분 = match.구분;
      saved.매칭근거 = `규칙학습:${match.매칭근거}`;
      changed++;
    }
  }
  return changed;
}

function openMautoClassifyResultView(rows) {
  document.querySelector(".mauto-classify-overlay")?.remove();
  const overlay = document.createElement("div");
  overlay.className = "mauto-classify-overlay bank-match-overlay";
  const rowsHtml = rows.map(r => {
    const dir = r.credit > 0
      ? `<span style="color:#1565c0;font-size:11px;">입금</span>`
      : `<span style="color:#b71c1c;font-size:11px;">출금</span>`;
    const isExcl = !!r.excluded;
    const isUnmatched = !r.excluded && !r.거래처명;
    const rowStyle = isExcl ? "opacity:0.4;" : isUnmatched ? "background:#fff7ed;" : "";
    const statusBadge = isExcl
      ? `<span style="font-size:10px;color:#9ca3af;">제외</span>`
      : isUnmatched
        ? `<span style="font-size:10px;color:#d97706;">미매칭</span>`
        : `<span style="font-size:10px;color:#166534;">✓</span>`;
    return `<tr style="${rowStyle}">
      <td style="font-size:12px;white-space:nowrap;">${escapeHtml(r.date)}</td>
      <td>${dir}</td>
      <td style="text-align:right;font-size:12px;">${formatNumber(r.credit || r.debit)}</td>
      <td style="font-size:12px;"><input type="text" class="mcl-memo-edit" data-txkey="${escapeHtml(r._txKey || "")}" value="${escapeHtml(r.memo || "")}" placeholder="비고" title="여러 연월 분배: 25-12=6000000 26-03=10000000" style="width:210px;font-size:12px;padding:3px 5px;border:1px solid #e5e7eb;border-radius:4px;" /></td>
      <td style="font-size:12px;">${escapeHtml(r.거래처명 || "")}</td>
      <td style="font-size:12px;">${escapeHtml(r.구분 || "")}</td>
      <td>${statusBadge}</td>
      <td style="font-size:11px;color:#6b7280;">${escapeHtml(r.매칭근거 || "")}</td>
    </tr>`;
  }).join("");
  overlay.innerHTML = `
    <div class="bank-match-dialog" style="max-width:860px;">
      <div class="bank-match-header">
        <h3>분류 결과 목록</h3>
        <span class="bank-match-sub">${rows.length}건</span>
        <button type="button" class="bank-match-close">✕</button>
      </div>
      <div style="padding:6px 14px;background:#eff6ff;border-bottom:1px solid #dbeafe;font-size:12px;color:#1e40af;">
        💡 한 건을 여러 달로 나눠 충당하려면 적요에 <b>연월=금액</b>을 쓰세요. 예) <code>25-12=6000000 26-03=10000000</code> → 25-12에 600만, 26-03에 1,000만 반영
      </div>
      <div class="table-responsive bank-match-table-wrap">
        <table class="bank-match-table">
          <thead><tr><th>날짜</th><th>구분</th><th>금액</th><th>적요</th><th>거래처명</th><th>매출/매입</th><th>상태</th><th>매칭근거</th></tr></thead>
          <tbody>${rowsHtml}</tbody>
        </table>
      </div>
      <div class="bank-match-actions">
        <button type="button" class="bank-cancel-btn">닫기</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);
  overlay.querySelector(".bank-match-close").addEventListener("click", () => overlay.remove());
  overlay.querySelector(".bank-cancel-btn").addEventListener("click", () => overlay.remove());

  // 적요 편집 → memoOverride 저장(여러 연월 분배 포함) → 재빌드 → 미수/미지급 갱신
  overlay.querySelectorAll(".mcl-memo-edit").forEach(inp => {
    const orig = inp.value;
    inp.addEventListener("blur", () => {
      const txKey = inp.dataset.txkey;
      const val = inp.value.trim();
      if (!txKey || val === orig.trim()) return;
      const prev = mautoUserEdits[txKey] || {};
      mautoUserEdits[txKey] = { ...prev, memoOverride: val };
      saveUserEdits();
      rebuildMautoRows();
      renderMautoTab();
      // 분배 인식 여부 즉시 피드백
      const allocs = parseMemoAllocations(val);
      if (allocs.length) {
        const sum = allocs.reduce((s, a) => s + a.amount, 0);
        inp.style.borderColor = "#16a34a";
        inp.title = `분배 ${allocs.length}건 인식 (합계 ${formatNumber(sum)})`;
      }
    });
    inp.addEventListener("keydown", e => { if (e.key === "Enter") inp.blur(); });
  });
}

function openBankImportDialog(bankRows) {
  document.querySelector(".bank-match-overlay")?.remove();

  const matches = bankRows.map(row => ({
    bankRow: row,
    candidates: matchBankRowToPayables(row, payables),
    selected: null,
    amount: row._debit || row._credit || parseBankAmount(row["출금"] || row["입금"] || 0),
    date: row._date || String(row["거래일자"] || row["날짜"] || row["거래일"] || "").trim(),
    memo: [row._memo, row._memo2].filter(Boolean).join(" / ") ||
      String(row["적요1"] || row["적요"] || row["내용"] || "").trim(),
    isCredit: (row._credit || 0) > 0 && (row._debit || 0) === 0,
    action: "skip",
  }));

  const overlay = document.createElement("div");
  overlay.className = "bank-match-overlay";

  function buildTableHtml() {
    return matches.map((m, idx) => {
      const top = m.candidates[0];
      const autoMatch = top && top.score >= 50;
      if (m.action === "skip" && autoMatch) m.action = "pay";
      if (m.action === "pay" && !m.selected && top) m.selected = top.item.sourceKey;

      const candidateOptions = [
        `<option value="">-- 직접 선택 --</option>`,
        ...m.candidates.map(c =>
          `<option value="${escapeHtml(c.item.sourceKey)}" ${m.selected === c.item.sourceKey ? "selected" : ""}>
            ${escapeHtml(c.item.name)} ${c.item.year}-${String(c.item.month).padStart(2, "0")} (${formatNumber(getPayableOutstanding(c.item))}원) [${Math.round(c.score)}점]
          </option>`
        ),
        ...payables.filter(p => p.completionStatus !== "완료" && !m.candidates.find(c => c.item.sourceKey === p.sourceKey))
          .map(p => `<option value="${escapeHtml(p.sourceKey)}">${escapeHtml(p.name)} ${p.year}-${String(p.month).padStart(2, "0")}</option>`)
      ].join("");

      const dirLabel = m.isCredit
        ? `<span style="color:#1565c0;font-size:11px;">입금</span>`
        : `<span style="color:#b71c1c;font-size:11px;">출금</span>`;
      return `<tr class="bank-match-row ${m.action === "skip" ? "bank-row-skip" : "bank-row-pay"} ${m.isCredit ? "bank-row-credit" : ""}" data-idx="${idx}">
        <td>${escapeHtml(m.date)}</td>
        <td class="bank-memo-cell" title="${escapeHtml(m.memo)}">${escapeHtml(m.memo.slice(0, 24))}${m.memo.length > 24 ? "…" : ""}</td>
        <td class="numeric-cell">${dirLabel} ${formatNumber(m.amount)}</td>
        <td>
          <select class="bank-match-select" data-idx="${idx}">${candidateOptions}</select>
        </td>
        <td>
          <label class="bank-action-toggle">
            <input type="checkbox" class="bank-action-chk" data-idx="${idx}" ${m.action === "pay" ? "checked" : ""} />
            적용
          </label>
        </td>
      </tr>`;
    }).join("");
  }

  overlay.innerHTML = `
    <div class="bank-match-dialog">
      <div class="bank-match-header">
        <h3>입출금 매칭</h3>
        <span class="bank-match-sub">${bankRows.length}건 · 점수 50+ 자동 매칭, 낮은 건은 직접 선택</span>
        <button type="button" class="bank-match-close">✕</button>
      </div>
      <div class="table-responsive bank-match-table-wrap">
        <table class="bank-match-table">
          <thead><tr>
            <th>날짜</th><th>적요</th><th>금액</th><th>매칭 업체</th><th>적용</th>
          </tr></thead>
          <tbody id="bankMatchTbody">${buildTableHtml()}</tbody>
        </table>
      </div>
      <div class="bank-match-actions">
        <span class="bank-match-count" id="bankMatchCount"></span>
        <button type="button" class="bank-apply-btn">선택 항목 지급 처리</button>
        <button type="button" class="bank-cancel-btn">취소</button>
      </div>
    </div>
  `;
  document.body.appendChild(overlay);

  function updateCount() {
    const n = matches.filter(m => m.action === "pay" && m.selected).length;
    overlay.querySelector("#bankMatchCount").textContent = `적용 ${n}건`;
  }
  updateCount();

  overlay.querySelector(".bank-match-close").addEventListener("click", () => overlay.remove());
  overlay.querySelector(".bank-cancel-btn").addEventListener("click", () => overlay.remove());

  overlay.querySelector("#bankMatchTbody").addEventListener("change", e => {
    const idx = Number(e.target.dataset.idx ?? -1);
    if (idx < 0) return;
    if (e.target.classList.contains("bank-match-select")) {
      matches[idx].selected = e.target.value || null;
    }
    if (e.target.classList.contains("bank-action-chk")) {
      matches[idx].action = e.target.checked ? "pay" : "skip";
    }
    updateCount();
    e.target.closest("tr")?.classList.toggle("bank-row-skip", matches[idx].action === "skip");
    e.target.closest("tr")?.classList.toggle("bank-row-pay", matches[idx].action === "pay");
  });

  overlay.querySelector(".bank-apply-btn").addEventListener("click", async () => {
    const toApply = matches.filter(m => m.action === "pay" && m.selected);
    if (!toApply.length) { alert("적용할 항목이 없습니다."); return; }

    // 지급처리: paidOverride 갱신 + 결제이력 append
    const historyRows = [];
    toApply.forEach(m => {
      const item = payables.find(p => p.sourceKey === m.selected);
      if (!item) return;
      const prevPaid = getPayableEffectivePaid(item);
      const newPaid = Math.min(prevPaid + m.amount, Number(item.purchase || 0));
      item.paidOverride = newPaid;
      if (newPaid >= Number(item.purchase || 0)) item.completionStatus = "완료";
      historyRows.push({
        source_key: item.sourceKey,
        거래처코드_norm: item.codeNormalized || item.code || "",
        거래처명: item.name,
        지급일자: m.date,
        지급금액: m.amount,
        적요: m.memo,
        결과상태: item.completionStatus === "완료" ? "완료" : "부분",
        created_at: new Date().toISOString(),
        created_by: "bank_import",
      });
    });

    try {
      if (historyRows.length && SHEET_APP_SCRIPT_URL) {
        await postSheetWebApp("appendPaymentHistory", { rows: historyRows });
        await postSheetWebApp("appendUpdateHistory", {
          rows: historyRows.map(r => ({
            recorded_at: r.created_at, section: "payables", action: "bank_import",
            stable_key: r.source_key, label: r.거래처명, prev_amount: "", new_amount: r.지급금액, memo: r.적요,
          }))
        });
      }
      persistPayablesState();
      overlay.remove();
      rerenderAll();
      alert(`${toApply.length}건 지급 처리 완료`);
    } catch (err) {
      alert(`저장 실패: ${err.message}`);
    }
  });
}

function setupBankImport() {
  const btn = document.getElementById("bankImportButton");
  const fileInput = document.getElementById("bankImportFileInput");
  if (!btn || !fileInput) return;
  btn.addEventListener("click", () => fileInput.click());
  fileInput.addEventListener("change", e => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = ev => {
      const wb = XLSX.read(ev.target.result, { type: "array", cellDates: true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: false });
      const bankRows = parseBankSheet(data);
      fileInput.value = "";
      if (!bankRows.length) { alert("인식된 거래 행이 없습니다. 헤더(날짜/적요/출금/입금) 확인 바랍니다."); return; }
      openBankImportDialog(bankRows);
    };
    reader.readAsArrayBuffer(file);
  });
}

// ── 대사 탭 ─────────────────────────────────────────────────

const daesaState = {
  loaded: false,
  loading: false,
  error: null,
  taxInvoices: [],
  ledgerSales: [],
  ledgerPurchase: [],
  ledgerPayable: [],
  dailySales: [],
  filterYear: new Date().getFullYear(),
  filterMonth: new Date().getMonth() + 1,
  // 정산표 뷰 상태
  settlementView: false,        // true = 정산표 모드
  settlementDiv: "매출",        // "매출" | "매입" | "미지급"
  settlementFilterYear: null,   // null = 전체
  // 부가세 뷰 상태
  vatView: false,               // true = 부가세 대시보드 모드
  vatMode: "분기",              // "월간" | "분기" | "반기" | "연간"
  vatYear: new Date().getFullYear(),
  // 미수/미지급 자동계산 뷰
  arRecapView: false,           // true = 미수/미지급 자동계산 뷰
  arRecapSide: "미지급",        // "미수" | "미지급"
  arRecapFilterYear: null,      // null = 전체
};

// 대사 탭 정렬 상태
const daesaSortState = { key: "name", dir: "asc" };
// 대사 탭 분류별 접힘 상태
const daesaCategoryCollapsed = {};

// 사업부문 마스터
const bizDivisionState = {
  rows: [],
  names: [],       // B열 값 목록 (부분일치용)
  lastFileName: "",
  status: "",
  saving: false,
};

async function fetchApiRows(action) {
  if (!SHEET_APP_SCRIPT_URL) throw new Error("Apps Script URL 없음");
  const url = new URL(SHEET_APP_SCRIPT_URL);
  url.searchParams.set("action", action);
  const _dToken = getApiToken();
  if (_dToken) url.searchParams.set("token", _dToken);
  const res = await fetch(url.toString());
  if (!res.ok) throw new Error(`${action} 조회 실패: ${res.status}`);
  const body = await res.json();
  return Array.isArray(body.rows) ? body.rows : (Array.isArray(body) ? body : []);
}

async function fetchDaesaAll() {
  if (!SHEET_APP_SCRIPT_URL) throw new Error("Apps Script URL 없음");
  const url = new URL(SHEET_APP_SCRIPT_URL);
  url.searchParams.set("action", "getDaesaAll");
  const token = getApiToken();
  if (token) url.searchParams.set("token", token);
  const res = await fetch(url.toString());
  if (!res.ok) throw new Error(`대사 데이터 조회 실패: ${res.status}`);
  const body = await res.json();
  const arr = k => Array.isArray(body[k]) ? body[k] : [];
  return {
    taxInvoices:    arr("taxInvoices"),
    ledgerSales:    arr("ledgerSales"),
    ledgerPurchase: arr("ledgerPurchase"),
    ledgerPayable:  arr("ledgerPayable"),
    dailySales:     arr("dailySales"),
    bizDivision:    arr("bizDivision"),
  };
}

async function loadDaesaData() {
  if (daesaState.loading) return;
  daesaState.loading = true;
  daesaState.error = null;
  renderDaesaTab();
  try {
    const all = await fetchDaesaAll();
    daesaState.taxInvoices = all.taxInvoices;
    daesaState.ledgerSales = all.ledgerSales;
    daesaState.ledgerPurchase = all.ledgerPurchase;
    daesaState.ledgerPayable = all.ledgerPayable;
    daesaState.dailySales = all.dailySales;
    const bizDiv = all.bizDivision;
    // 사업부문 마스터 로드 (시트에 저장된 경우)
    if (Array.isArray(bizDiv) && bizDiv.length) {
      const allHeaders = Object.keys(bizDiv[0] || {});
      const bColHeader = allHeaders[1] || "사업부문/현장명";
      bizDivisionState.rows = bizDiv;
      bizDivisionState.names = [...new Set(
        bizDiv.map(r => String(r[bColHeader] || r["사업부문/현장명"] || "").trim()).filter(Boolean)
      )];
    }
    daesaState.loaded = true;
  } catch (err) {
    daesaState.error = err.message;
  } finally {
    daesaState.loading = false;
    renderDaesaTab();
  }
}

// ────────────────────────────────────────────────────────────
//  미래 소스 파일 재빌드 (파일 단위 교체 → daesaState 갱신)
// ────────────────────────────────────────────────────────────
function rebuildDaesaFromSources() {
  // dedup 헬퍼 — _row_key 기준 1건화
  const dedupByKey = rows => {
    const seen = new Map();
    for (const r of rows) {
      const k = String(r["_row_key"] || "").trim();
      if (!k || !seen.has(k)) seen.set(k || `__rnd_${Math.random()}`, r);
    }
    return [...seen.values()];
  };

  // 세금계산서: 승인번호(_row_key) dedup
  daesaState.taxInvoices = dedupByKey(
    Object.values(miraeTaxSources).flatMap(f => f.rows || [])
  );

  // 원장 3종: ledgerType별 dedup
  const ledgerOf = lt => dedupByKey(
    Object.values(miraeLedgerSources).filter(f => f.ledgerType === lt).flatMap(f => f.rows || [])
  );
  daesaState.ledgerSales    = ledgerOf("매출");
  daesaState.ledgerPurchase = ledgerOf("매입");
  daesaState.ledgerPayable  = ledgerOf("미지급");

  // 영업현황: dedup
  daesaState.dailySales = dedupByKey(
    Object.values(miraeBizSources).flatMap(f => f.rows || [])
  );

  daesaState.loaded = true;
  daesaState.error  = null;

  // 대사 탭 활성 상태면 즉시 재렌더
  const el = document.getElementById("daesa");
  if (el && !el.classList.contains("hidden")) renderDaesaTab();
}

// ════════════════════════════════════════════════════════════
//  Phase 1 — 미래 원장 정산 (거래처×귀속연월 집계)
// ════════════════════════════════════════════════════════════

// 계정별원장 적요에서 귀속연월 추출
// 반환: { 연도, 월, status:"ok"|"확인필요", 원본 }
// 못 알아보면 status="확인필요"로 살림 (절대 null로 버리지 않음)
function parseYearMonthCode(raw) {
  const 원본 = raw == null ? "" : String(raw);
  const s = 원본.trim();
  if (!s) return { 연도: null, 월: null, status: "확인필요", 원본 };

  const ok  = (y, mo, st) => ({ 연도: 2000 + y, 월: mo, status: st, 원본 });
  const okY = (y, mo, st) => ({ 연도: y,         월: mo, status: st, 원본 });
  const valid = mo => mo >= 1 && mo <= 12;
  let m;

  // 1) 4자리연도 yyyy-mm (희귀 → 확인필요)
  if ((m = s.match(/\b(20\d{2})[-/.](\d{1,2})\b/)) && valid(+m[2]))
    return okY(+m[1], +m[2], "확인필요");
  // 2) 슬래시 yy/mm (희귀 → 확인필요)
  if ((m = s.match(/\b(\d{2})\/(\d{1,2})\b/)) && valid(+m[2]))
    return ok(+m[1], +m[2], "확인필요");
  // 3) 표준 yy-mm / yy-m
  if ((m = s.match(/\b(\d{2})-(\d{1,2})\b/)) && valid(+m[2]))
    return ok(+m[1], +m[2], "ok");
  // 4) 구분자 없는 숫자런: yymm(4) / yymmdd(6~8). 앞 4자리가 yymm
  if ((m = s.match(/\b(\d{4,8})\b/))) {
    const d = m[1], mo = +d.slice(2, 4);
    if ((d.length === 4 || d.length >= 6) && valid(mo))
      return ok(+d.slice(0, 2), mo, "ok");
  }
  // 5) 인식 실패 → 확인필요로 살림
  return { 연도: null, 월: null, status: "확인필요", 원본 };
}

// 비고에서 "연월=금액" 명시적 분배를 추출 (한 입금/출금을 여러 귀속연월로 나눠 충당)
// 예: "25-12=6000000 26-03=10000000" → [{ym:"2025-12",amount:6000000},{ym:"2026-03",amount:10000000}]
// 구분자는 = 또는 : , 금액은 콤마 허용. 명시적 구분자(=/:)가 있어야만 인식(오탐 방지).
function parseMemoAllocations(memo) {
  const s = String(memo || "");
  const re = /(\d{2})-(\d{1,2})\s*[=:]\s*([\d,]+)/g;
  const out = [];
  let m;
  while ((m = re.exec(s)) !== null) {
    const yy = +m[1], mo = +m[2];
    if (mo < 1 || mo > 12) continue;
    const amt = Number(String(m[3]).replace(/,/g, "")) || 0;
    if (amt <= 0) continue;
    out.push({ ym: `${2000 + yy}-${String(mo).padStart(2, "0")}`, amount: amt });
  }
  return out;
}

// 원장 행 배열 → 거래처×귀속연월 정산 레코드 집계
// 구분: "매출"(108 외상매출금) | "매입"(251 외상매입금) | "미지급"(미지급금)
//   매출: 차변=발생, 대변=충당
//   매입/미지급: 대변=발생, 차변=충당
function buildLedgerSettlement(ledgerRows, 구분, 사업체 = "미래") {
  const groups = new Map();
  const 확인필요 = [];

  for (const row of ledgerRows) {
    const 거래처코드 = String(row["거래처코드"] ?? "").trim();
    const 거래처명  = String(row["거래처명"]   ?? "").trim();
    const 차변 = Number(String(row["차변"] ?? "").replace(/[^0-9.-]/g, "")) || 0;
    const 대변 = Number(String(row["대변"] ?? "").replace(/[^0-9.-]/g, "")) || 0;
    const ym = parseYearMonthCode(row["적요"]);

    if (ym.status === "확인필요") {
      if (차변 || 대변) {
        확인필요.push({ 거래처코드, 거래처명, 적요: ym.원본, 차변, 대변, 일자: row["일자"] ?? "" });
      }
      continue;
    }

    const key = `${거래처코드}|${ym.연도}|${ym.월}`;
    if (!groups.has(key)) {
      groups.set(key, {
        사업체, 구분, 거래처코드, 거래처명,
        귀속연도: ym.연도, 귀속월: ym.월,
        발생합계: 0, 충당액: 0, 잔액: 0,
      });
    }
    const g = groups.get(key);
    if (구분 === "매출") {
      g.발생합계 += 차변;
      g.충당액   += 대변;
    } else {
      g.발생합계 += 대변;
      g.충당액   += 차변;
    }
  }

  const records = [...groups.values()].map(g => ({ ...g, 잔액: g.발생합계 - g.충당액 }));
  return { records, 확인필요 };
}

// 정산 레코드를 거래처×귀속연월로 집계 (여러 구분 합산 시 사용)
function aggregateSettlement(records) {
  const map = new Map();
  for (const r of records) {
    const key = `${r.거래처코드}|${r.귀속연도}|${r.귀속월}|${r.구분}`;
    if (!map.has(key)) {
      map.set(key, { ...r });
    } else {
      const g = map.get(key);
      g.발생합계 += r.발생합계;
      g.충당액   += r.충당액;
      g.잔액     += r.잔액;
    }
  }
  return [...map.values()];
}

function parseAmt(val) {
  if (typeof val === "number") return val;
  return Number(String(val || "").replace(/[^0-9.-]/g, "")) || 0;
}

function rowToYearMonth(dateStr) {
  const s = String(dateStr || "").trim();
  const m = s.match(/^(\d{4})-(\d{2})/);
  return m ? `${m[1]}-${m[2]}` : null;
}

// 적요 앞부분에서 연월 추출 (예: "25-11", "25.11 소형압연>>", "260101")
function extractYearMonthFromMemo(memo) {
  const s = String(memo || "").trim();
  if (!s) return null;
  // 패턴 1: YY-MM 또는 YY.MM (예: "25-11", "25-11 소형압연>>")
  const m1 = s.match(/^(\d{2})[.\-](\d{2})/);
  if (m1) {
    const year = 2000 + parseInt(m1[1], 10);
    const month = parseInt(m1[2], 10);
    if (month >= 1 && month <= 12) return `${year}-${String(month).padStart(2, "0")}`;
  }
  // 패턴 2: YYMMDD (예: "260101" → 2026-01)
  const m2 = s.match(/^(\d{2})(\d{2})(\d{2})/);
  if (m2) {
    const year = 2000 + parseInt(m2[1], 10);
    const month = parseInt(m2[2], 10);
    const day = parseInt(m2[3], 10);
    if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      return `${year}-${String(month).padStart(2, "0")}`;
    }
  }
  return null;
}

// 적요에서 사업부문 이름 부분일치 추출 (날짜 부분 제거 후)
function extractBizDivision(memoText) {
  if (!bizDivisionState.names.length) return null;
  const s = String(memoText || "").trim();
  if (!s) return null;
  // 날짜 부분 제거: YY-MM 또는 YY.MM 또는 YYMMDD 패턴
  const rest = s.replace(/^\d{2}[.\-]\d{2}\s*/, "").replace(/^\d{6}\s*/, "").trim();
  if (!rest) return null;
  for (const nm of bizDivisionState.names) {
    if (nm && (rest.includes(nm) || nm.includes(rest))) return nm;
  }
  return null;
}

// 차이 표시 셀 생성 (세금계산서 기준)
// 형식: "발급여부 | 세금-원장 차이 | 세금-영업 차이"
function buildDiffCell(tax, ledger, biz) {
  if (tax === 0 && ledger === 0 && biz === 0) return "";
  const hasOther = ledger > 0 || biz > 0;
  // 세금계산서가 0인데 원장이나 영업에 금액이 있으면 X (미발급)
  const s0 = (tax === 0 && hasOther)
    ? '<span class="daesa-err" title="세금계산서 미발급">X</span>'
    : (tax > 0 ? '<span class="daesa-ok">V</span>' : '<span style="color:#94a3b8">—</span>');

  const d1 = tax - ledger;
  const s1 = d1 === 0
    ? '<span class="daesa-ok">V</span>'
    : `<span class="daesa-err" title="원장 차이: ${formatNumber(-d1)}">${d1 < 0 ? "-" : "+"}${formatNumber(Math.abs(d1))}</span>`;

  const d2 = tax - biz;
  const s2 = d2 === 0
    ? '<span class="daesa-ok">V</span>'
    : `<span class="daesa-err" title="영업 차이: ${formatNumber(-d2)}">${d2 < 0 ? "-" : "+"}${formatNumber(Math.abs(d2))}</span>`;

  return `${s0}&nbsp;|&nbsp;${s1}&nbsp;|&nbsp;${s2}`;
}

async function parseBizDivisionFile(file) {
  try {
    const ab = await file.arrayBuffer();
    const { dataRows } = parseXlsToRows(ab, 4); // 5행 = index 4
    if (!dataRows.length) throw new Error("데이터가 없습니다.");
    const allHeaders = Object.keys(dataRows[0]);
    const bColHeader = allHeaders[1] || "사업부문/현장명";
    dataRows.forEach(r => {
      if (!r["_row_key"]) {
        r["_row_key"] = String(r["코드"] || r[bColHeader] || "").trim();
      }
    });
    const rows = dataRows.filter(r => r["_row_key"]);
    bizDivisionState.rows = rows;
    bizDivisionState.names = [...new Set(
      rows.map(r => String(r[bColHeader] || r["사업부문/현장명"] || "").trim()).filter(Boolean)
    )];
    bizDivisionState.lastFileName = file.name;
    return { ok: true, count: bizDivisionState.names.length, rows };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

function getNetOffVendorSet() {
  const s = new Set();
  receivableManagerState.rows.forEach(row => {
    const info = String(row["정보"] || "").trim();
    if (info === "상계") {
      const code = normalizeVendorCode(
        String(row["거래처코드"] || row["code"] || "").trim()
      );
      if (code) s.add(code);
    }
  });
  return s;
}

function buildDaesaMap() {
  const map = new Map();
  const vendorMaps = buildVendorLookupMaps();
  const shoppingDateRegex = /_20\d{6}/; // 최적화: 루프 외부 선언

  const codeToMasterName = {};
  const codeToCategory = {}; // 업체마스터의 '거래처구분' 저장
  vendorMasterState.rows.forEach(v => {
    const code = String(v["거래처코드_norm"] || "").trim();
    if (code) {
      codeToMasterName[code] = String(v["거래처명"] || "");
      codeToCategory[code] = String(v["거래처구분"] || "").trim();
    }
  });

  function ensureEntry(key, name, ym) {
    if (!key || !ym) return null;
    if (!map.has(key)) map.set(key, { name, months: {}, unmatched: key.startsWith("__no__") });
    if (!map.get(key).months[ym]) {
      map.get(key).months[ym] = {
        taxSales: 0, taxPurchase: 0,
        ledgerSales: 0, ledgerCollect: 0,
        ledgerBuy: 0, ledgerPay: 0,
        ledgerPayable: 0, ledgerPayablePay: 0,
        bizSales: 0, bizPurchase: 0, bizCollect: 0, bizPay: 0,
        taxSalesDetail: {}, // { vendorName: amount } - 쇼핑몰 매출용
        divBreakdownLedger: {}, // { divName: { collect:0, pay:0 } }
        divBreakdownBiz: {},    // { divName: { sales:0, purchase:0, collect:0, pay:0 } }
      };
    }
    return map.get(key).months[ym];
  }

  // 세금계산서: 사업자번호(기호 제거) → 업체마스터 거래처코드
  daesaState.taxInvoices.forEach(r => {
    const bn = normalizeBizNum(r["사업자(주민)번호"]);
    const matched = bn ? vendorMaps.byBiz[bn] : null;

    let key, name;
    const cat = matched ? codeToCategory[matched.code] : "";

    // 오토몰인 경우 '◆쇼핑몰매출'로 통합
    if (cat === "오토몰") {
      key = "SHOPPINGMALL_SALES";
      name = "◆쇼핑몰매출";
    } else {
      key = matched ? matched.code : `__no__tax_${bn || r["상호"]}`;
      name = matched ? (codeToMasterName[matched.code] || matched.name)
        : `[마스터없음] ${r["상호"] || bn}`;
    }

    const ym = rowToYearMonth(r["작성일자"]);
    const amt = parseAmt(r["합계"]);
    const type = String(r["구분"] || "").trim();
    const e = ensureEntry(key, name, ym);
    if (!e) return;

    if (type === "매출") {
      e.taxSales += amt;
      // 상세 상호별 합계 기록 (통합 업체용)
      const realName = r["상호"] || "상호불명";
      e.taxSalesDetail[realName] = (e.taxSalesDetail[realName] || 0) + amt;
    }
    else if (type === "매입") e.taxPurchase += amt;
  });

  // 거래처코드 → 업체마스터 거래처코드 (계정별원장·일별영업현황 공통)
  function resolveByCode(rawCode, rawName) {
    const c = String(rawCode || "").trim().replace(/^0+/, "");
    const matched = c ? vendorMaps.byCode[c] : null;

    const cat = matched ? codeToCategory[matched.code] : "";
    if (cat === "오토몰") {
      return { key: "SHOPPINGMALL_SALES", name: "◆쇼핑몰매출" };
    }

    const key = matched ? matched.code : `__no__code_${rawCode || rawName}`;
    const name = matched ? (codeToMasterName[matched.code] || matched.name)
      : `[마스터없음] ${rawName || rawCode}`;
    return { key, name };
  }

  daesaState.ledgerSales.forEach(r => {
    const { key, name } = resolveByCode(r["거래처코드"], r["거래처명"]);
    const ymTrans = rowToYearMonth(r["일자"]);
    const memoText = String(r["적요"] || r["비고"] || "").trim();
    const ymMemo = extractYearMonthFromMemo(memoText) || ymTrans;
    const divName = String(r["사업부분명"] || r["사업부문명"] || "").trim() || extractBizDivision(memoText) || "";

    // 차변(매출 발생): 거래일자 기준 연월에 귀속
    const eSales = ensureEntry(key, name, ymTrans);
    if (eSales) {
      const amt = parseAmt(r["차변"]);
      eSales.ledgerSales += amt;

      if (key === "SHOPPINGMALL_SALES" && amt) {
        // 지연 초기화: 쇼핑몰 업체인 경우에만 객체 생성
        if (!eSales.shoppingBreakdown) {
          eSales.shoppingBreakdown = {
            taxInvoice: { sales: 0, collect: 0 },
            cashReceipt: { sales: 0, collect: 0 },
            unclassified: { sales: 0, collect: 0 }
          };
        }
        if (shoppingDateRegex.test(memoText)) {
          eSales.shoppingBreakdown.taxInvoice.sales += amt;
        } else {
          eSales.shoppingBreakdown.cashReceipt.sales += amt;
        }
      }
    }

    // 대변(수금): 적요에 명시된 연월에 귀속
    const eCollect = ensureEntry(key, name, ymMemo);
    if (eCollect) {
      const amt = parseAmt(r["대변"]);
      eCollect.ledgerCollect += amt;
      if (!eCollect.divBreakdownLedger[divName]) eCollect.divBreakdownLedger[divName] = { collect: 0, pay: 0 };
      eCollect.divBreakdownLedger[divName].collect += amt;

      if (key === "SHOPPINGMALL_SALES" && amt) {
        if (!eCollect.shoppingBreakdown) {
          eCollect.shoppingBreakdown = {
            taxInvoice: { sales: 0, collect: 0 },
            cashReceipt: { sales: 0, collect: 0 },
            unclassified: { sales: 0, collect: 0 }
          };
        }
        if (memoText.includes("세계")) {
          eCollect.shoppingBreakdown.taxInvoice.collect += amt;
        } else if (memoText.includes("현영")) {
          eCollect.shoppingBreakdown.cashReceipt.collect += amt;
        } else {
          eCollect.shoppingBreakdown.unclassified.collect += amt;
        }
      }
    }
  });

  daesaState.ledgerPurchase.forEach(r => {
    const { key, name } = resolveByCode(r["거래처코드"], r["거래처명"]);
    const ymTrans = rowToYearMonth(r["일자"]);
    const memoText = String(r["적요"] || r["비고"] || "").trim();
    const ymMemo = extractYearMonthFromMemo(memoText) || ymTrans;
    const divName = String(r["사업부분명"] || r["사업부문명"] || "").trim() || extractBizDivision(memoText) || "";

    // 대변(매입 발생): 거래일자 기준 연월에 귀속
    const ePurchase = ensureEntry(key, name, ymTrans);
    if (ePurchase) ePurchase.ledgerBuy += parseAmt(r["대변"]);

    // 차변(지급): 적요에 명시된 연월에 귀속
    const ePay = ensureEntry(key, name, ymMemo);
    if (ePay) {
      const amt = parseAmt(r["차변"]);
      ePay.ledgerPay += amt;
      if (!ePay.divBreakdownLedger[divName]) ePay.divBreakdownLedger[divName] = { collect: 0, pay: 0 };
      ePay.divBreakdownLedger[divName].pay += amt;
    }
  });

  daesaState.ledgerPayable.forEach(r => {
    const { key, name } = resolveByCode(r["거래처코드"], r["거래처명"]);
    const ymTrans = rowToYearMonth(r["일자"]);
    const memoText = String(r["적요"] || r["비고"] || "").trim();
    const ymMemo = extractYearMonthFromMemo(memoText) || ymTrans;
    const divName = String(r["사업부분명"] || r["사업부문명"] || "").trim() || extractBizDivision(memoText) || "";

    // 대변(미지급 발생): 거래일자 기준 연월에 귀속
    const ePayable = ensureEntry(key, name, ymTrans);
    if (ePayable) ePayable.ledgerPayable += parseAmt(r["대변"]);

    // 차변(지급): 적요에 명시된 연월에 귀속
    const ePayPay = ensureEntry(key, name, ymMemo);
    if (ePayPay) {
      const amt = parseAmt(r["차변"]);
      ePayPay.ledgerPayablePay += amt;
      if (!ePayPay.divBreakdownLedger[divName]) ePayPay.divBreakdownLedger[divName] = { collect: 0, pay: 0 };
      ePayPay.divBreakdownLedger[divName].pay += amt;
    }
  });

  daesaState.dailySales.forEach(r => {
    const { key, name } = resolveByCode(r["거래처코드"], r["거래처명"]);
    const ymTrans = rowToYearMonth(r["거래일자"]);
    const memoText = String(r["적요"] || r["비고"] || r["메모"] || "").trim();
    const ymMemo = extractYearMonthFromMemo(memoText) || ymTrans;
    const bizDivTitle = String(r["사업부문현장명"] || "").trim();
    const divName = bizDivTitle || extractBizDivision(memoText) || "";

    const gubun = String(r["구분"] || r["g구분"] || "").trim();
    const isTax = String(r["세금계산서"] || r["n세금계산서"] || "").trim() !== "-";

    // 판매·구매금액: 세금계산서가 '-'가 아닌 것만 합산 (거래일자 기준)
    const eTrans = ensureEntry(key, name, ymTrans);
    if (eTrans && isTax) {
      if (gubun.includes("판매")) {
        const sAmt = parseAmt(r["판매금액"]);
        eTrans.bizSales += sAmt;
        if (sAmt) {
          if (!eTrans.divBreakdownBiz[divName]) eTrans.divBreakdownBiz[divName] = { sales: 0, purchase: 0, collect: 0, pay: 0 };
          eTrans.divBreakdownBiz[divName].sales += sAmt;
        }
      }
      if (gubun.includes("구매")) {
        const pAmt = parseAmt(r["구매금액"]);
        eTrans.bizPurchase += pAmt;
        if (pAmt) {
          if (!eTrans.divBreakdownBiz[divName]) eTrans.divBreakdownBiz[divName] = { sales: 0, purchase: 0, collect: 0, pay: 0 };
          eTrans.divBreakdownBiz[divName].purchase += pAmt;
        }
      }
    }

    // 수금·지급액: 세금계산서 발행 여부와 상관없이 항상 합산 (적요 연월 기준)
    const eMemo = ensureEntry(key, name, ymMemo);
    if (eMemo) {
      const cAmt = parseAmt(r["수금액"]);
      const pAmt = parseAmt(r["지급액"]);
      eMemo.bizCollect += cAmt;
      eMemo.bizPay += pAmt;
      if (cAmt || pAmt) {
        if (!eMemo.divBreakdownBiz[divName]) eMemo.divBreakdownBiz[divName] = { sales: 0, purchase: 0, collect: 0, pay: 0 };
        eMemo.divBreakdownBiz[divName].collect += cAmt;
        eMemo.divBreakdownBiz[divName].pay += pAmt;
      }
    }
  });

  return map;
}

function diffLabel(a, b) {
  const d = Math.abs(a - b);
  if (d === 0) return '<span class="daesa-ok">✓</span>';
  if (d <= 1000) return `<span class="daesa-warn">△${formatNumber(d)}</span>`;
  return `<span class="daesa-err">✗${formatNumber(d)}</span>`;
}

function renderDaesaTab() {
  const section = document.getElementById("daesa");
  if (!section) return;

  if (daesaState.loading) {
    section.innerHTML = `<div class="daesa-loading">데이터 불러오는 중…</div>`;
    return;
  }
  if (daesaState.error) {
    section.innerHTML = `<div class="daesa-error">오류: ${escapeHtml(daesaState.error)}
      <button class="daesa-reload-btn">다시 시도</button></div>`;
    section.querySelector(".daesa-reload-btn")?.addEventListener("click", loadDaesaData);
    return;
  }
  if (!daesaState.loaded) {
    section.innerHTML = `<div class="daesa-empty">
      <button class="daesa-load-btn">대사 데이터 불러오기</button>
      <p class="muted">세금계산서·계정별원장·영업현황 시트에서 불러옵니다.</p>
    </div>`;
    section.querySelector(".daesa-load-btn")?.addEventListener("click", loadDaesaData);
    return;
  }

  const ym = `${daesaState.filterYear}-${String(daesaState.filterMonth).padStart(2, "0")}`;
  const daesaMap = buildDaesaMap();
  const netOffSet = getNetOffVendorSet();

  const q = (elements.searchInput?.value || "").toLowerCase().trim();

  // 필터링 + 가공된 데이터
  let vendorEntries = [...daesaMap.entries()]
    .filter(([code, v]) => {
      if (!v.months[ym]) return false;
      if (!q) return true;
      return v.name.toLowerCase().includes(q) || code.toLowerCase().includes(q);
    })
    .filter(([, v]) => {
      // 세금계산서·계정별원장·영업현황 모두 0원이면 숨김
      const d = v.months[ym];
      return d.taxSales !== 0 || d.ledgerSales !== 0 || d.bizSales !== 0 ||
             d.taxPurchase !== 0 || (d.ledgerBuy + (d.ledgerPayable || 0)) !== 0 || d.bizPurchase !== 0;
    });

  // 마스터 미등록 업체 감지
  const unmappedInPeriod = [...daesaMap.entries()]
    .filter(([, v]) => v.name.startsWith("[마스터없음]") && v.months[ym]);
  const unmappedNames = [...new Set(unmappedInPeriod.map(([, v]) => v.name.replace("[마스터없음] ", "")))];


  // 정렬 수행
  vendorEntries.sort((a, b) => {
    const v1 = a[1], v2 = b[1];
    const d1 = v1.months[ym], d2 = v2.months[ym];
    let val1, val2;

    switch (daesaSortState.key) {
      case "name": val1 = v1.name; val2 = v2.name; break;
      case "taxSales": val1 = d1.taxSales; val2 = d2.taxSales; break;
      case "ledgerSales": val1 = d1.ledgerSales; val2 = d2.ledgerSales; break;
      case "bizSales": val1 = d1.bizSales; val2 = d2.bizSales; break;
      case "taxPurchase": val1 = d1.taxPurchase; val2 = d2.taxPurchase; break;
      case "ledgerBuy": val1 = d1.ledgerBuy + d1.ledgerPayable; val2 = d2.ledgerBuy + d2.ledgerPayable; break;
      case "bizPurchase": val1 = d1.bizPurchase; val2 = d2.bizPurchase; break;
      default: val1 = v1.name; val2 = v2.name;
    }

    if (typeof val1 === "string") {
      return daesaSortState.dir === "asc" ? val1.localeCompare(val2, "ko") : val2.localeCompare(val1, "ko");
    }
    return daesaSortState.dir === "asc" ? val1 - val2 : val2 - val1;
  });

  // 연도/월 옵션 생성
  const years = [...new Set([...daesaMap.values()].flatMap(v => Object.keys(v.months).map(k => k.slice(0, 4))))].sort();
  const months = Array.from({ length: 12 }, (_, i) => i + 1);

  const yearOpts = years.map(y =>
    `<option value="${y}" ${y == daesaState.filterYear ? "selected" : ""}>${y}년</option>`
  ).join("");
  const monthOpts = months.map(m =>
    `<option value="${m}" ${m == daesaState.filterMonth ? "selected" : ""}>${m}월</option>`
  ).join("");

  const hasNetOff = vendorEntries.some(([code]) => netOffSet.has(code));
  const colCount = 1 + 4 + 4 + (hasNetOff ? 1 : 0);

  // 업체마스터에서 분류(거래처구분) 조회
  const codeToCategory = {};
  vendorMasterState.rows.forEach(v => {
    const c = String(v["거래처코드_norm"] || "").trim();
    if (c) codeToCategory[c] = String(v["거래처구분"] || "").trim();
  });

  // 분류별 그룹핑
  const categoryMap = new Map();
  vendorEntries.forEach(([code, vendor]) => {
    const cat = codeToCategory[code] || "기타";
    if (!categoryMap.has(cat)) categoryMap.set(cat, []);
    categoryMap.get(cat).push([code, vendor]);
  });
  const sortedCats = [...categoryMap.keys()].sort((a, b) => {
    if (a === "기타") return 1;
    if (b === "기타") return -1;
    return a.localeCompare(b, "ko");
  });

  function makeVendorRow(code, vendor) {
    const d = vendor.months[ym];
    const isNetOff = netOffSet.has(code);
    const ledgerBuyTotal = d.ledgerBuy + d.ledgerPayable;
    const netoffAmt = isNetOff ? Math.min(d.taxSales, d.taxPurchase) : 0;
    const diffS = buildDiffCell(d.taxSales, d.ledgerSales, d.bizSales);
    const diffP = buildDiffCell(d.taxPurchase, ledgerBuyTotal, d.bizPurchase);
    const matchS = d.taxSales === d.ledgerSales && d.taxSales === d.bizSales;
    const matchP = d.taxPurchase === ledgerBuyTotal && d.taxPurchase === d.bizPurchase;
    return `<tr class="${(!matchS || !matchP) ? "daesa-row-mismatch" : ""}">
      <td class="daesa-vendor-cell">
        <button class="daesa-vendor-btn" data-code="${escapeHtml(code)}" data-name="${escapeHtml(vendor.name)}">${escapeHtml(vendor.name)}</button>
        ${isNetOff ? `<span class="daesa-netoff-badge">상계</span>` : ""}
      </td>
      <td class="num col-sales ${d.taxSales !== d.ledgerSales || d.taxSales !== d.bizSales ? "daesa-mismatch-val" : ""}">${formatNumber(d.taxSales)}</td>
      <td class="num col-sales ${d.ledgerSales !== d.taxSales ? "daesa-mismatch-val" : ""}">${formatNumber(d.ledgerSales)}</td>
      <td class="num col-sales ${d.bizSales !== d.taxSales ? "daesa-mismatch-val" : ""}">${formatNumber(d.bizSales)}</td>
      <td class="col-diff col-sales">${diffS}</td>
      <td class="num col-purchase ${d.taxPurchase !== ledgerBuyTotal || d.taxPurchase !== d.bizPurchase ? "daesa-mismatch-val" : ""}">${formatNumber(d.taxPurchase)}</td>
      <td class="num col-purchase ${ledgerBuyTotal !== d.taxPurchase ? "daesa-mismatch-val" : ""}">${formatNumber(ledgerBuyTotal)}</td>
      <td class="num col-purchase ${d.bizPurchase !== d.taxPurchase ? "daesa-mismatch-val" : ""}">${formatNumber(d.bizPurchase)}</td>
      <td class="col-diff col-purchase">${diffP}</td>
      ${hasNetOff ? `<td class="num col-netoff">${formatNumber(netoffAmt)}</td>` : ""}
    </tr>`;
  }

  const rows = sortedCats.map(cat => {
    const entries = categoryMap.get(cat) || [];
    const collapsed = !!daesaCategoryCollapsed[cat];
    const catRow = `<tr class="daesa-cat-header">
      <td colspan="${colCount}">
        <button class="daesa-cat-toggle" data-cat="${escapeHtml(cat)}">${collapsed ? "▶" : "▼"}</button>
        <strong>${escapeHtml(cat)}</strong>
        <span class="daesa-cat-count">${entries.length}개 업체</span>
      </td>
    </tr>`;
    if (collapsed) return catRow;
    return catRow + entries.map(([code, vendor]) => makeVendorRow(code, vendor)).join("");
  }).join("");

  function sortIcon(key) {
    if (daesaSortState.key !== key) return '<span class="sort-arrow">↕</span>';
    return daesaSortState.dir === "asc" ? '<span class="sort-arrow">↑</span>' : '<span class="sort-arrow">↓</span>';
  }

  const unmappedBanner = unmappedNames.length
    ? `<div class="daesa-unmapped-banner">
        <span class="daesa-unmapped-icon">⚠️</span>
        <span><strong>업체마스터 미등록 ${unmappedNames.length}개</strong>: ${unmappedNames.map(n => `<em>${escapeHtml(n)}</em>`).join(", ")}</span>
        <span class="daesa-unmapped-hint">→ <strong>🛠 마스터 관리</strong>에서 업체마스터를 등록하면 같은 업체가 한 줄로 합쳐집니다.</span>
      </div>`
    : "";

  section.innerHTML = `
    ${unmappedBanner}
    <div class="daesa-toolbar">
      <select id="daesaYearFilter">${yearOpts}</select>
      <select id="daesaMonthFilter">${monthOpts}</select>
      <button class="daesa-reload-btn">↺ 새로고침</button>
      <button class="daesa-expand-all-btn">전체 펼치기</button>
      <button class="daesa-collapse-all-btn">전체 접기</button>
      <span class="daesa-count muted">${vendorEntries.length}개 업체 표시중 ${q ? `(검색: ${q})` : ""}</span>
      <span class="daesa-toolbar-sep"></span>
      <button class="daesa-settlement-btn${daesaState.settlementView ? " active" : ""}">정산표</button>
      <button class="daesa-vat-btn${daesaState.vatView ? " active" : ""}">부가세</button>
      <button class="daesa-arrecap-btn${daesaState.arRecapView ? " active" : ""}">미수/미지급</button>
    </div>
    ${daesaState.settlementView ? renderSettlementView() : ""}
    ${daesaState.vatView ? renderVatView() : ""}
    ${daesaState.arRecapView ? renderArRecapView() : ""}
    <div class="table-responsive">
      <table class="daesa-table">
        <thead>
          <tr>
            <th rowspan="2" class="daesa-th-vendor daesa-sort-th" data-key="name">업체명 ${sortIcon("name")}</th>
            <th colspan="4" class="daesa-th-group daesa-th-sales">매출</th>
            <th colspan="4" class="daesa-th-group daesa-th-purchase">매입</th>
            ${hasNetOff ? `<th rowspan="2" class="daesa-th-netoff">상계금액</th>` : ""}
          </tr>
          <tr>
            <th class="daesa-th-sub daesa-th-sub-sales daesa-sort-th" data-key="taxSales">세금계산서 ${sortIcon("taxSales")}</th>
            <th class="daesa-th-sub daesa-th-sub-sales daesa-sort-th" data-key="ledgerSales">계정별원장 ${sortIcon("ledgerSales")}</th>
            <th class="daesa-th-sub daesa-th-sub-sales daesa-sort-th" data-key="bizSales">영업현황 ${sortIcon("bizSales")}</th>
            <th class="daesa-th-sub daesa-th-sub-sales">분석(발|원|영)</th>
            <th class="daesa-th-sub daesa-th-sub-purchase daesa-sort-th" data-key="taxPurchase">세금계산서 ${sortIcon("taxPurchase")}</th>
            <th class="daesa-th-sub daesa-th-sub-purchase daesa-sort-th" data-key="ledgerBuy">계정별원장 ${sortIcon("ledgerBuy")}</th>
            <th class="daesa-th-sub daesa-th-sub-purchase daesa-sort-th" data-key="bizPurchase">영업현황 ${sortIcon("bizPurchase")}</th>
            <th class="daesa-th-sub daesa-th-sub-purchase">분석(발|원|영)</th>
          </tr>
        </thead>
        <tbody>${rows || `<tr><td colspan="10" style="text-align:center;padding:24px;color:#94a3b8;">${ym} 데이터 없음</td></tr>`}</tbody>
      </table>
    </div>
  `;

  section.querySelector("#daesaYearFilter")?.addEventListener("change", e => {
    daesaState.filterYear = Number(e.target.value);
    renderDaesaTab();
  });
  section.querySelector("#daesaMonthFilter")?.addEventListener("change", e => {
    daesaState.filterMonth = Number(e.target.value);
    renderDaesaTab();
  });
  section.querySelector(".daesa-reload-btn")?.addEventListener("click", () => {
    daesaState.loaded = false;
    loadDaesaData();
  });
  section.querySelectorAll(".daesa-vendor-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      openVendorDaesaModal(btn.dataset.code, btn.dataset.name, daesaMap, netOffSet);
    });
  });
  section.querySelectorAll(".daesa-sort-th").forEach(th => {
    th.addEventListener("click", () => {
      const key = th.dataset.key;
      if (daesaSortState.key === key) {
        daesaSortState.dir = (daesaSortState.dir === "asc" ? "desc" : "asc");
      } else {
        daesaSortState.key = key;
        daesaSortState.dir = "asc";
      }
      renderDaesaTab();
    });
  });
  section.querySelectorAll(".daesa-cat-toggle").forEach(btn => {
    btn.addEventListener("click", () => {
      const cat = btn.dataset.cat;
      daesaCategoryCollapsed[cat] = !daesaCategoryCollapsed[cat];
      renderDaesaTab();
    });
  });
  section.querySelector(".daesa-expand-all-btn")?.addEventListener("click", () => {
    sortedCats.forEach(cat => { daesaCategoryCollapsed[cat] = false; });
    renderDaesaTab();
  });
  section.querySelector(".daesa-collapse-all-btn")?.addEventListener("click", () => {
    sortedCats.forEach(cat => { daesaCategoryCollapsed[cat] = true; });
    renderDaesaTab();
  });
  section.querySelector(".daesa-settlement-btn")?.addEventListener("click", () => {
    daesaState.settlementView = !daesaState.settlementView;
    if (daesaState.settlementView) daesaState.vatView = false;
    renderDaesaTab();
  });
  section.querySelector(".daesa-vat-btn")?.addEventListener("click", () => {
    daesaState.vatView = !daesaState.vatView;
    if (daesaState.vatView) { daesaState.settlementView = false; daesaState.arRecapView = false; }
    renderDaesaTab();
  });
  section.querySelector(".daesa-arrecap-btn")?.addEventListener("click", () => {
    daesaState.arRecapView = !daesaState.arRecapView;
    if (daesaState.arRecapView) { daesaState.settlementView = false; daesaState.vatView = false; }
    renderDaesaTab();
  });
  section.querySelector("#arRecapYearFilter")?.addEventListener("change", e => {
    daesaState.arRecapFilterYear = e.target.value === "" ? null : Number(e.target.value);
    renderDaesaTab();
  });
  section.querySelectorAll(".arrecap-side-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      daesaState.arRecapSide = btn.dataset.side;
      renderDaesaTab();
    });
  });
  section.querySelector("#vatModeFilter")?.addEventListener("change", e => {
    daesaState.vatMode = e.target.value;
    renderDaesaTab();
  });
  section.querySelector("#vatYearFilter")?.addEventListener("change", e => {
    daesaState.vatYear = Number(e.target.value);
    renderDaesaTab();
  });
  // 정산표 내부 이벤트
  section.querySelector("#settlementDivFilter")?.addEventListener("change", e => {
    daesaState.settlementDiv = e.target.value;
    renderDaesaTab();
  });
  section.querySelector("#settlementYearFilter")?.addEventListener("change", e => {
    daesaState.settlementFilterYear = e.target.value === "" ? null : Number(e.target.value);
    renderDaesaTab();
  });
}

// ────────────────────────────────────────────────────────────
//  Phase 4-A — 미수/미지급 자동계산 (세금계산서 발생 − 입출금 충당)
// ────────────────────────────────────────────────────────────

// sideType: "미수"(매출) | "미지급"(매입)
// taxInvoices: daesaState.taxInvoices
// classifiedRows: mautoClassifiedRows (입출금 분류 결과)
// taxInvoices: mautoTaxInvoices 또는 daesaState.taxInvoices
// classifiedRows: mautoClassifiedRows (입출금 분류 결과)
// sideType: "미수"(매출) | "미지급"(매입)
// 매칭 키 우선순위: 거래처코드_norm > 사업자번호(biz:) > 정규화상호(name:)
function buildArRecap(taxInvoices, classifiedRows, sideType) {
  const 구분키  = sideType === "미수" ? "매출" : "매입";
  const amtField = sideType === "미수" ? "credit" : "debit";

  // 업체마스터 조회 맵 (사업자번호·이름 양방향)
  const maps    = buildVendorLookupMaps();
  const nameMap = buildVendorNameMap();

  // 주어진 사업자번호+이름으로 표준 매칭 키와 표시명 반환
  // 1순위: 마스터 사업자번호 → 거래처코드_norm
  // 2순위: 마스터 정규화상호 → 거래처코드_norm
  // 3순위: 사업자번호 자체 (biz:NNNN)
  // 4순위: 정규화 상호 (name:…)
  function getVendorKey(bizNum, name) {
    const bn = normalizeBusinessNumber(bizNum || "");
    if (bn && bn !== "0") {
      const v = maps.byBiz[bn];
      if (v?.code) return { key: v.code, displayName: v.name || name, exact: true };
      return { key: `biz:${bn}`, displayName: name, exact: false };
    }
    const norm = normalizeVendorName(name);
    const vn = norm && nameMap[norm];
    if (vn?.code) return { key: vn.code, displayName: vn.name, exact: true };
    return { key: `name:${norm || name}`, displayName: name, exact: false };
  }

  // 거래처명(분류 결과)만으로 키 조회 (충당 측)
  function getVendorKeyByName(name) {
    const norm = normalizeVendorName(name);
    const vn = norm && nameMap[norm];
    if (vn?.code) return { key: vn.code, displayName: vn.name, exact: true };
    return { key: `name:${norm || name}`, displayName: name, exact: false };
  }

  // ─ 발생액: 세금계산서 → 거래처코드_norm×작성연월 집계 ─
  const 발생맵 = new Map(); // key: "vendorKey\tYYYY-MM" → number
  const keyToDisplay = new Map(); // vendorKey → 표시명
  const inexactKeys = new Set();  // 코드매칭 실패한 키 추적

  for (const r of taxInvoices) {
    if (String(r["구분"] || "").trim() !== 구분키) continue;
    const ym = rowToYearMonth(r["작성일자"]);
    if (!ym) continue;
    const bizNum = r["사업자번호"] || r["사업자(주민)번호"] || "";
    const name   = (r["_matched_name"] || r["상호"] || r["거래처명"] || "").trim();
    if (!name && !bizNum) continue;

    const { key, displayName, exact } = getVendorKey(bizNum, name);
    const mapKey = `${key}\t${ym}`;
    발생맵.set(mapKey, (발생맵.get(mapKey) || 0) + parseAmt(r["합계"]));
    if (!keyToDisplay.has(key)) keyToDisplay.set(key, displayName || name);
    if (!exact) inexactKeys.add(key);
  }

  // ─ 충당액: 분류된 입출금 → 거래처코드_norm×귀속연월 집계 ─
  const 충당맵  = new Map();
  const 확인필요 = [];

  for (const r of (classifiedRows || [])) {
    if (r.excluded) continue;
    if (String(r.구분 || "").trim() !== 구분키) continue;
    let amt = Number(r[amtField]) || 0;
    // 방어: 로컬 분류행 금액이 0이면 _txKey(_date|_time|credit|debit|memo)에서 복구 (옛 오염 데이터 자동 보정)
    if (amt <= 0) {
      const p = String(r._txKey || "").split("|");
      amt = Number(p[amtField === "credit" ? 2 : 3]) || 0;
    }
    if (amt <= 0) continue;
    const vendorName = (r.거래처명 || "").trim();
    if (!vendorName) continue;

    const memoSrc = r._memo2 || r._memo || r.memo || "";

    // ── 여러 연월 분배: 비고에 "25-12=6000000 26-03=10000000" 형식이 있으면 그대로 분배 ──
    const allocs = parseMemoAllocations(memoSrc);
    if (allocs.length) {
      const { key, displayName, exact } = getVendorKeyByName(vendorName);
      if (!keyToDisplay.has(key)) keyToDisplay.set(key, displayName);
      if (!exact) inexactKeys.add(key);
      for (const a of allocs) {
        충당맵.set(`${key}\t${a.ym}`, (충당맵.get(`${key}\t${a.ym}`) || 0) + a.amount);
      }
      continue;
    }

    const parsed  = parseYearMonthCode(memoSrc);
    if (parsed.status === "확인필요") {
      확인필요.push({ vendor: vendorName, amt, memo: memoSrc, date: r.date, 매칭근거: r.매칭근거 });
      continue;
    }
    const 귀속ym = `${parsed.연도}-${String(parsed.월).padStart(2, "0")}`;

    const { key, displayName, exact } = getVendorKeyByName(vendorName);
    const mapKey = `${key}\t${귀속ym}`;
    충당맵.set(mapKey, (충당맵.get(mapKey) || 0) + amt);
    if (!keyToDisplay.has(key)) keyToDisplay.set(key, displayName);
    if (!exact) inexactKeys.add(key);
  }

  // ─ 발생 기준 머지 → 잔액 계산 (세금계산서 없는 월의 충당은 제외) ─
  const allKeys = new Set([...발생맵.keys()]);
  const entries = [];
  for (const mapKey of allKeys) {
    const tabIdx = mapKey.lastIndexOf("\t");
    const vendorKey = mapKey.slice(0, tabIdx);
    const ym  = mapKey.slice(tabIdx + 1);
    const [year, month] = ym.split("-").map(Number);
    const 발생 = 발생맵.get(mapKey) || 0;
    const 충당 = 충당맵.get(mapKey) || 0;
    const vendor = keyToDisplay.get(vendorKey) || vendorKey;
    const inexact = inexactKeys.has(vendorKey);
    entries.push({ vendor, vendorKey, ym, year, month, 발생, 충당, 잔액: 발생 - 충당, inexact });
  }

  // 거래처명 → 연도 → 월 순 정렬
  entries.sort((a, b) => {
    const nc = a.vendor.localeCompare(b.vendor, "ko");
    if (nc !== 0) return nc;
    if (a.year !== b.year) return a.year - b.year;
    return a.month - b.month;
  });

  // 매칭 불가 거래처 목록 (업체마스터에 없어서 이름으로만 묶인 것)
  const inexactVendors = [...inexactKeys].map(k => keyToDisplay.get(k) || k);

  return { entries, 확인필요, inexactVendors };
}

function renderArRecapView() {
  const side       = daesaState.arRecapSide;
  const filterYear = daesaState.arRecapFilterYear;
  const fn = n => formatNumber(n);
  const jCls = n => n > 0 ? "arrecap-pos" : n < 0 ? "arrecap-neg" : "";

  // 엠오토 세금계산서가 있으면 우선 사용, 없으면 미래 세금계산서 fallback
  const taxSrc = (mautoTaxInvoices && mautoTaxInvoices.length) ? mautoTaxInvoices : daesaState.taxInvoices;
  const { entries, 확인필요, inexactVendors } = buildArRecap(
    taxSrc,
    typeof mautoClassifiedRows !== "undefined" ? mautoClassifiedRows : [],
    side
  );

  // 연도 필터 옵션
  const dataYears = [...new Set(entries.map(e => e.year).filter(Boolean))].sort((a, b) => b - a);
  const yearOpts = [`<option value="">전체</option>`,
    ...dataYears.map(y => `<option value="${y}" ${y === filterYear ? "selected" : ""}>${y}년</option>`)
  ].join("");

  const filtered = filterYear ? entries.filter(e => e.year === filterYear) : entries;

  // 검색어 필터
  const q = (document.getElementById("searchInput")?.value || "").toLowerCase().trim();
  const searched = q ? filtered.filter(e => e.vendor.toLowerCase().includes(q)) : filtered;

  // 거래처별 그룹
  const groups = new Map();
  for (const e of searched) {
    if (!groups.has(e.vendor)) groups.set(e.vendor, []);
    groups.get(e.vendor).push(e);
  }

  let bodyHtml = "";
  let gDev = 0, gChung = 0, gJan = 0;

  for (const [vendor, rows] of groups) {
    const subDev   = rows.reduce((s, r) => s + r.발생, 0);
    const subChung = rows.reduce((s, r) => s + r.충당, 0);
    const subJan   = rows.reduce((s, r) => s + r.잔액, 0);
    gDev += subDev; gChung += subChung; gJan += subJan;

    rows.forEach((r, i) => {
      bodyHtml += `<tr class="arrecap-row${i === 0 ? " arrecap-first" : ""}">
        <td class="arrecap-vendor-cell">${i === 0 ? escapeHtml(vendor) : ""}</td>
        <td class="arrecap-ym-cell">${r.ym}</td>
        <td class="arrecap-num-cell">${r.발생 ? fn(r.발생) : "-"}</td>
        <td class="arrecap-num-cell">${r.충당 ? fn(r.충당) : "-"}</td>
        <td class="arrecap-num-cell arrecap-jan-cell ${jCls(r.잔액)}">${fn(r.잔액)}</td>
      </tr>`;
    });
    if (rows.length > 1) {
      bodyHtml += `<tr class="arrecap-sub-row">
        <td colspan="2" style="text-align:right;font-size:12px;color:#374151;">↳ ${escapeHtml(vendor)} 소계</td>
        <td class="arrecap-num-cell">${fn(subDev)}</td>
        <td class="arrecap-num-cell">${fn(subChung)}</td>
        <td class="arrecap-num-cell arrecap-jan-cell ${jCls(subJan)}"><strong>${fn(subJan)}</strong></td>
      </tr>`;
    }
  }

  const pendingBadge = 확인필요.length
    ? `<span class="arrecap-pending-badge">⚠ 귀속연월 미확인 ${확인필요.length}건</span>`
    : "";
  const inexactBadge = inexactVendors.length
    ? `<span class="arrecap-inexact-badge" title="${inexactVendors.slice(0,10).join(', ')}${inexactVendors.length > 10 ? ' 외 ' + (inexactVendors.length-10) + '개' : ''}">업체마스터 미매칭 ${inexactVendors.length}개</span>`
    : "";

  const pendingHtml = 확인필요.length ? `
    <details class="arrecap-pending-wrap">
      <summary class="arrecap-pending-summary">⚠ 귀속연월 미확인 ${확인필요.length}건 — 비고에서 연월을 인식 못한 입출금 행 (충당 미반영)</summary>
      <table class="arrecap-pending-table">
        <thead><tr><th>거래처</th><th>금액</th><th>비고(원본)</th><th>거래일자</th><th>매칭근거</th></tr></thead>
        <tbody>${확인필요.map(r => `<tr>
          <td>${escapeHtml(r.vendor)}</td>
          <td class="arrecap-num-cell">${fn(r.amt)}</td>
          <td class="arrecap-memo">${escapeHtml(r.memo)}</td>
          <td>${escapeHtml(r.date || "")}</td>
          <td class="muted">${escapeHtml(r.매칭근거 || "")}</td>
        </tr>`).join("")}</tbody>
      </table>
    </details>` : "";

  const inexactHtml = inexactVendors.length ? `
    <details class="arrecap-pending-wrap">
      <summary class="arrecap-pending-summary">업체마스터 미매칭 ${inexactVendors.length}개 — 사업자번호 없어 이름으로만 묶임. 마스터에 사업자번호를 추가하면 교차매칭 정확도가 높아집니다.</summary>
      <div style="padding:8px 12px;font-size:12px;color:#374151;">${inexactVendors.map(n => `<span style="display:inline-block;margin:2px 4px;padding:1px 8px;background:#fef9c3;border:1px solid #fde68a;border-radius:12px;">${escapeHtml(n)}</span>`).join("")}</div>
    </details>` : "";

  const emptyMsg = `<tr><td colspan="5" style="text-align:center;padding:20px;color:#9ca3af;">
    데이터 없음 — 세금계산서를 자료업로드로 불러오고, 엠오토 탭에서 입출금을 분류하세요.
  </td></tr>`;

  return `
    <div class="arrecap-wrap">
      <div class="arrecap-toolbar">
        <strong>${side} 자동계산</strong>
        <div class="arrecap-side-toggle">
          <button class="arrecap-side-btn${side === "미수" ? " active" : ""}" data-side="미수">미수금</button>
          <button class="arrecap-side-btn${side === "미지급" ? " active" : ""}" data-side="미지급">미지급</button>
        </div>
        <span class="daesa-toolbar-sep"></span>
        <label>연도 <select id="arRecapYearFilter">${yearOpts}</select></label>
        ${pendingBadge}${inexactBadge}
        <span class="muted" style="font-size:12px;">발생: ${(mautoTaxInvoices && mautoTaxInvoices.length) ? `엠오토 세금계산서 ${mautoTaxInvoices.length}건` : `미래 세금계산서 ${daesaState.taxInvoices.length}건`} | 충당: 엠오토 입출금 분류(비고→귀속연월)</span>
      </div>
      <div class="table-responsive">
        <table class="arrecap-table">
          <thead>
            <tr>
              <th class="arrecap-th-vendor">거래처</th>
              <th class="arrecap-th-ym">연월</th>
              <th class="arrecap-th-num">발생액 (세금계산서)</th>
              <th class="arrecap-th-num">충당액 (입출금)</th>
              <th class="arrecap-th-jan">잔액</th>
            </tr>
          </thead>
          <tbody>${searched.length ? bodyHtml : emptyMsg}</tbody>
          ${searched.length ? `<tfoot><tr class="arrecap-total-row">
            <td colspan="2"><strong>총계</strong></td>
            <td class="arrecap-num-cell">${fn(gDev)}</td>
            <td class="arrecap-num-cell">${fn(gChung)}</td>
            <td class="arrecap-num-cell arrecap-jan-cell ${jCls(gJan)}"><strong>${fn(gJan)}</strong></td>
          </tr></tfoot>` : ""}
        </table>
      </div>
      ${pendingHtml}
      ${inexactHtml}
      <p class="arrecap-note muted">※ 발생: 작성일자(작성연월) 기준. 충당: 입출금 비고의 귀속연월(yy-mm) 파싱. 구분=매출/매입 분류된 행만 포함. 거래처 매칭: 업체마스터 사업자번호 → 거래처코드_norm → 정규화 상호 순.</p>
    </div>
  `;
}

// ────────────────────────────────────────────────────────────
//  부가세 대시보드 (Phase 3 — 세금계산서 작성일자 기준)
// ────────────────────────────────────────────────────────────

// taxInvoices 배열 → 연월별 집계 맵 반환
// 반환: Map(ym → { 매출공급: number, 매출세액: number, 매입공급: number, 매입세액: number })
function buildVatSummary(taxInvoices) {
  const map = new Map();
  const ensure = ym => {
    if (!map.has(ym)) map.set(ym, { 매출공급: 0, 매출세액: 0, 매입공급: 0, 매입세액: 0 });
    return map.get(ym);
  };
  for (const r of taxInvoices) {
    const ym = rowToYearMonth(r["작성일자"]);
    if (!ym) continue;
    const type = String(r["구분"] || "").trim();
    if (type !== "매출" && type !== "매입") continue;
    const supply = parseAmt(r["공급가액"]);
    const tax    = parseAmt(r["세액"]);
    const e = ensure(ym);
    if (type === "매출") { e.매출공급 += supply; e.매출세액 += tax; }
    else                  { e.매입공급 += supply; e.매입세액 += tax; }
  }
  return map;
}

// mode별 기간 묶음 생성: [{ label, months: ["YYYY-MM", ...] }, ...]
function buildVatPeriods(year, mode) {
  const y = String(year);
  if (mode === "월간") {
    return Array.from({ length: 12 }, (_, i) => {
      const m = String(i + 1).padStart(2, "0");
      return { label: `${i + 1}월`, months: [`${y}-${m}`] };
    });
  }
  if (mode === "분기") {
    return [
      { label: "1분기 (1~3월)",  months: [`${y}-01`, `${y}-02`, `${y}-03`] },
      { label: "2분기 (4~6월)",  months: [`${y}-04`, `${y}-05`, `${y}-06`] },
      { label: "3분기 (7~9월)",  months: [`${y}-07`, `${y}-08`, `${y}-09`] },
      { label: "4분기 (10~12월)", months: [`${y}-10`, `${y}-11`, `${y}-12`] },
    ];
  }
  if (mode === "반기") {
    return [
      { label: "1기 (1~6월)",  months: [`${y}-01`,`${y}-02`,`${y}-03`,`${y}-04`,`${y}-05`,`${y}-06`] },
      { label: "2기 (7~12월)", months: [`${y}-07`,`${y}-08`,`${y}-09`,`${y}-10`,`${y}-11`,`${y}-12`] },
    ];
  }
  // 연간
  return [{ label: `${year}년 합계`, months: Array.from({ length: 12 }, (_, i) => `${y}-${String(i+1).padStart(2,"0")}`) }];
}

function renderVatView() {
  const vatMap = buildVatSummary(daesaState.taxInvoices);
  const mode = daesaState.vatMode;
  const year = daesaState.vatYear;

  // 연도 옵션: 데이터 있는 연도 + 현재 연도
  const dataYears = [...new Set([...vatMap.keys()].map(ym => Number(ym.slice(0, 4))))].sort((a, b) => b - a);
  if (!dataYears.includes(year)) dataYears.unshift(year);
  const yearOpts = dataYears.map(y => `<option value="${y}" ${y === year ? "selected" : ""}>${y}년</option>`).join("");

  const periods = buildVatPeriods(year, mode);

  // 합산: period별
  const rows = periods.map(p => {
    const agg = { 매출공급: 0, 매출세액: 0, 매입공급: 0, 매입세액: 0 };
    for (const ym of p.months) {
      const e = vatMap.get(ym);
      if (e) { agg.매출공급 += e.매출공급; agg.매출세액 += e.매출세액; agg.매입공급 += e.매입공급; agg.매입세액 += e.매입세액; }
    }
    agg.납부세액 = agg.매출세액 - agg.매입세액;
    return { label: p.label, ...agg };
  });

  // 총계
  const total = rows.reduce((s, r) => {
    s.매출공급 += r.매출공급; s.매출세액 += r.매출세액;
    s.매입공급 += r.매입공급; s.매입세액 += r.매입세액;
    s.납부세액 += r.납부세액;
    return s;
  }, { 매출공급: 0, 매출세액: 0, 매입공급: 0, 매입세액: 0, 납부세액: 0 });

  const fn = n => formatNumber(n);
  const납부cls = n => n < 0 ? "vat-refund" : n > 0 ? "vat-pay" : "";

  const bodyRows = rows.map(r => `
    <tr>
      <td class="vat-period-cell">${r.label}</td>
      <td class="vat-num-cell">${fn(r.매출공급)}</td>
      <td class="vat-num-cell vat-tax-cell">${fn(r.매출세액)}</td>
      <td class="vat-num-cell">${fn(r.매입공급)}</td>
      <td class="vat-num-cell vat-tax-cell">${fn(r.매입세액)}</td>
      <td class="vat-num-cell vat-result-cell ${납부cls(r.납부세액)}">${r.납부세액 < 0 ? "▲ " + fn(-r.납부세액) : fn(r.납부세액)}</td>
    </tr>
  `).join("");

  const totalRow = `
    <tr class="vat-total-row">
      <td class="vat-period-cell"><strong>합계</strong></td>
      <td class="vat-num-cell">${fn(total.매출공급)}</td>
      <td class="vat-num-cell vat-tax-cell">${fn(total.매출세액)}</td>
      <td class="vat-num-cell">${fn(total.매입공급)}</td>
      <td class="vat-num-cell vat-tax-cell">${fn(total.매입세액)}</td>
      <td class="vat-num-cell vat-result-cell ${납부cls(total.납부세액)}">${total.납부세액 < 0 ? "▲ " + fn(-total.납부세액) : fn(total.납부세액)}</td>
    </tr>
  `;

  const dataCount = daesaState.taxInvoices.filter(r => {
    const t = String(r["구분"] || "").trim();
    return (t === "매출" || t === "매입") && rowToYearMonth(r["작성일자"])?.startsWith(String(year));
  }).length;

  return `
    <div class="vat-view-wrap">
      <div class="vat-toolbar">
        <strong>부가세 납부세액 집계</strong>
        <span class="daesa-toolbar-sep"></span>
        <label>연도 <select id="vatYearFilter">${yearOpts}</select></label>
        <label>기간 <select id="vatModeFilter">
          <option value="월간" ${mode === "월간" ? "selected" : ""}>월간</option>
          <option value="분기" ${mode === "분기" ? "selected" : ""}>분기</option>
          <option value="반기" ${mode === "반기" ? "selected" : ""}>반기</option>
          <option value="연간" ${mode === "연간" ? "selected" : ""}>연간</option>
        </select></label>
        <span class="muted" style="font-size:12px;">${year}년 세금계산서 ${dataCount}건 기준</span>
      </div>
      <div class="table-responsive">
        <table class="vat-table">
          <thead>
            <tr>
              <th rowspan="2" class="vat-th-period">기간</th>
              <th colspan="2" class="vat-th-group vat-th-sales">매출</th>
              <th colspan="2" class="vat-th-group vat-th-purchase">매입</th>
              <th rowspan="2" class="vat-th-result">납부(환급)세액</th>
            </tr>
            <tr>
              <th class="vat-th-sub">공급가액</th>
              <th class="vat-th-sub vat-th-tax">세액</th>
              <th class="vat-th-sub">공급가액</th>
              <th class="vat-th-sub vat-th-tax">세액</th>
            </tr>
          </thead>
          <tbody>${bodyRows}</tbody>
          <tfoot>${totalRow}</tfoot>
        </table>
      </div>
      <p class="vat-note muted">※ 집계 기준: 세금계산서 작성일자(작성연월). 구분이 비어있는 행(이자·인출 등)은 제외됩니다.</p>
    </div>
  `;
}

// 엠오토 부가세 보고서 (mautoTaxInvoices 기반, 반기 기본)
function renderMautoVatView() {
  if (!mautoTaxInvoices.length) {
    return `<div class="vat-view-wrap" id="mauto-vat-view" style="padding:20px 16px;text-align:center;color:#9ca3af;">
      세금계산서 데이터가 없습니다. 상단 <strong>🧾 매출세금계산서</strong> / <strong>🧾 매입세금계산서</strong> 버튼으로 파일을 업로드하세요.
    </div>`;
  }

  const vatMap = buildVatSummary(mautoTaxInvoices);
  const mode   = mautoVatMode;
  const year   = mautoVatYear;

  // 연도 옵션: 데이터 있는 연도 + 현재 연도
  const dataYears = [...new Set([...vatMap.keys()].map(ym => Number(ym.slice(0, 4))))].sort((a, b) => b - a);
  if (!dataYears.includes(year)) dataYears.unshift(year);
  const yearOpts = dataYears.map(y => `<option value="${y}" ${y === year ? "selected" : ""}>${y}년</option>`).join("");

  const periods = buildVatPeriods(year, mode);

  const rows = periods.map(p => {
    const agg = { 매출공급: 0, 매출세액: 0, 매입공급: 0, 매입세액: 0 };
    for (const ym of p.months) {
      const e = vatMap.get(ym);
      if (e) { agg.매출공급 += e.매출공급; agg.매출세액 += e.매출세액; agg.매입공급 += e.매입공급; agg.매입세액 += e.매입세액; }
    }
    agg.납부세액 = agg.매출세액 - agg.매입세액;
    return { label: p.label, ...agg };
  });

  const total = rows.reduce((s, r) => {
    s.매출공급 += r.매출공급; s.매출세액 += r.매출세액;
    s.매입공급 += r.매입공급; s.매입세액 += r.매입세액;
    s.납부세액 += r.납부세액;
    return s;
  }, { 매출공급: 0, 매출세액: 0, 매입공급: 0, 매입세액: 0, 납부세액: 0 });

  const fn = n => formatNumber(n);
  const 납부cls = n => n < 0 ? "vat-refund" : n > 0 ? "vat-pay" : "";

  const bodyRows = rows.map(r => `
    <tr>
      <td class="vat-period-cell">${r.label}</td>
      <td class="vat-num-cell">${fn(r.매출공급)}</td>
      <td class="vat-num-cell vat-tax-cell">${fn(r.매출세액)}</td>
      <td class="vat-num-cell">${fn(r.매입공급)}</td>
      <td class="vat-num-cell vat-tax-cell">${fn(r.매입세액)}</td>
      <td class="vat-num-cell vat-result-cell ${납부cls(r.납부세액)}">${r.납부세액 < 0 ? "▲ " + fn(-r.납부세액) : fn(r.납부세액)}</td>
    </tr>
  `).join("");

  const totalRow = `
    <tr class="vat-total-row">
      <td class="vat-period-cell"><strong>합계</strong></td>
      <td class="vat-num-cell">${fn(total.매출공급)}</td>
      <td class="vat-num-cell vat-tax-cell">${fn(total.매출세액)}</td>
      <td class="vat-num-cell">${fn(total.매입공급)}</td>
      <td class="vat-num-cell vat-tax-cell">${fn(total.매입세액)}</td>
      <td class="vat-num-cell vat-result-cell ${납부cls(total.납부세액)}">${total.납부세액 < 0 ? "▲ " + fn(-total.납부세액) : fn(total.납부세액)}</td>
    </tr>
  `;

  const dataCount = mautoTaxInvoices.filter(r => {
    const t = String(r["구분"] || "").trim();
    return (t === "매출" || t === "매입") && rowToYearMonth(r["작성일자"])?.startsWith(String(year));
  }).length;

  return `
    <div class="vat-view-wrap" id="mauto-vat-view">
      <div class="vat-toolbar">
        <strong>부가세 납부세액 집계 (엠오토)</strong>
        <span class="daesa-toolbar-sep"></span>
        <label>연도 <select id="mautoVatYearFilter">${yearOpts}</select></label>
        <label>기간 <select id="mautoVatModeFilter">
          <option value="월간" ${mode === "월간" ? "selected" : ""}>월간</option>
          <option value="반기" ${mode === "반기" ? "selected" : ""}>반기 (기본)</option>
          <option value="연간" ${mode === "연간" ? "selected" : ""}>연간</option>
        </select></label>
        <span class="muted" style="font-size:12px;">${year}년 ${dataCount}건 / 세금계산서 총 ${mautoTaxInvoices.length}건</span>
      </div>
      <div class="table-responsive">
        <table class="vat-table">
          <thead>
            <tr>
              <th rowspan="2" class="vat-th-period">기간</th>
              <th colspan="2" class="vat-th-group vat-th-sales">매출</th>
              <th colspan="2" class="vat-th-group vat-th-purchase">매입</th>
              <th rowspan="2" class="vat-th-result">납부(환급)세액</th>
            </tr>
            <tr>
              <th class="vat-th-sub">공급가액</th>
              <th class="vat-th-sub vat-th-tax">세액</th>
              <th class="vat-th-sub">공급가액</th>
              <th class="vat-th-sub vat-th-tax">세액</th>
            </tr>
          </thead>
          <tbody>${bodyRows}</tbody>
          <tfoot>${totalRow}</tfoot>
        </table>
      </div>
      <p class="vat-note muted">※ 집계 기준: 세금계산서 작성일자(작성연월). 구분이 비어있는 행은 제외됩니다.<br>
      ※ 엠오토(개인사업자) 기본 신고기간: 반기 (1기 1~6월 / 2기 7~12월)</p>
    </div>
  `;
}

// ────────────────────────────────────────────────────────────
//  정산표 뷰 (Phase 1 — 미래 원장 기반 거래처×귀속연월)
// ────────────────────────────────────────────────────────────
function renderSettlementView() {
  const div = daesaState.settlementDiv;   // "매출" | "매입" | "미지급"
  const filterYear = daesaState.settlementFilterYear;

  // 원장 선택
  const ledgerMap = { 매출: daesaState.ledgerSales, 매입: daesaState.ledgerPurchase, 미지급: daesaState.ledgerPayable };
  const rows = ledgerMap[div] || [];

  const { records, 확인필요 } = buildLedgerSettlement(rows, div, "미래");

  // 연도 필터 옵션 생성
  const years = [...new Set(records.map(r => r.귀속연도).filter(Boolean))].sort((a, b) => b - a);
  const yearOpts = [
    `<option value="">전체</option>`,
    ...years.map(y => `<option value="${y}" ${y === filterYear ? "selected" : ""}>${y}년</option>`),
  ].join("");

  // 필터 적용
  const filtered = filterYear ? records.filter(r => r.귀속연도 === filterYear) : records;

  // 검색어 필터
  const q = (document.getElementById("searchInput")?.value || "").toLowerCase().trim();
  const searched = q
    ? filtered.filter(r => r.거래처명.toLowerCase().includes(q) || r.거래처코드.includes(q))
    : filtered;

  // 정렬: 거래처명 → 귀속연도 → 귀속월
  searched.sort((a, b) => {
    if (a.거래처명 !== b.거래처명) return a.거래처명.localeCompare(b.거래처명, "ko");
    if (a.귀속연도 !== b.귀속연도) return a.귀속연도 - b.귀속연도;
    return a.귀속월 - b.귀속월;
  });

  // 잔액 소계 행 (거래처별)
  const clientGroups = new Map();
  for (const r of searched) {
    if (!clientGroups.has(r.거래처명)) clientGroups.set(r.거래처명, []);
    clientGroups.get(r.거래처명).push(r);
  }

  let rowsHtml = "";
  let grandDev = 0, grandChung = 0, grandJan = 0;

  for (const [name, grp] of clientGroups) {
    const subDev   = grp.reduce((s, r) => s + r.발생합계, 0);
    const subChung = grp.reduce((s, r) => s + r.충당액,   0);
    const subJan   = grp.reduce((s, r) => s + r.잔액,     0);
    grandDev   += subDev;
    grandChung += subChung;
    grandJan   += subJan;

    grp.forEach((r, i) => {
      rowsHtml += `<tr class="stl-row${i === 0 ? " stl-first-in-group" : ""}">
        <td class="stl-name">${i === 0 ? escapeHtml(name) : ""}</td>
        <td class="stl-code">${i === 0 ? escapeHtml(r.거래처코드) : ""}</td>
        <td class="stl-ym">${r.귀속연도}-${String(r.귀속월).padStart(2, "0")}</td>
        <td class="num stl-dev">${formatNumber(r.발생합계)}</td>
        <td class="num stl-chung">${formatNumber(r.충당액)}</td>
        <td class="num stl-jan ${r.잔액 > 0 ? "stl-jan-pos" : r.잔액 < 0 ? "stl-jan-neg" : ""}">${formatNumber(r.잔액)}</td>
      </tr>`;
    });

    if (grp.length > 1) {
      rowsHtml += `<tr class="stl-subtotal">
        <td colspan="3" style="text-align:right;font-size:12px;color:#374151;">↳ ${escapeHtml(name)} 소계</td>
        <td class="num">${formatNumber(subDev)}</td>
        <td class="num">${formatNumber(subChung)}</td>
        <td class="num ${subJan > 0 ? "stl-jan-pos" : subJan < 0 ? "stl-jan-neg" : ""}">${formatNumber(subJan)}</td>
      </tr>`;
    }
  }

  const emptyRow = `<tr><td colspan="6" style="text-align:center;padding:20px;color:#9ca3af;">데이터 없음</td></tr>`;

  // 확인필요 섹션
  const needsBadge = 확인필요.length
    ? `<span class="stl-need-badge">확인 필요 ${확인필요.length}건</span>` : "";

  const needsHtml = 확인필요.length ? `
    <details class="stl-needs-wrap">
      <summary class="stl-needs-summary">⚠ 확인 필요 ${확인필요.length}건 — 적요에서 귀속연월을 인식하지 못한 행</summary>
      <table class="stl-needs-table">
        <thead><tr><th>거래처명</th><th>거래처코드</th><th>적요 원본</th><th class="num">차변</th><th class="num">대변</th><th>일자</th></tr></thead>
        <tbody>${확인필요.map(r => `<tr>
          <td>${escapeHtml(r.거래처명)}</td>
          <td>${escapeHtml(r.거래처코드)}</td>
          <td class="stl-raw">${escapeHtml(r.적요)}</td>
          <td class="num">${formatNumber(r.차변)}</td>
          <td class="num">${formatNumber(r.대변)}</td>
          <td>${escapeHtml(String(r.일자))}</td>
        </tr>`).join("")}</tbody>
      </table>
    </details>` : "";

  return `
    <div class="stl-wrap">
      <div class="stl-toolbar">
        <strong>미래 원장 정산표</strong>
        <select id="settlementDivFilter">
          <option value="매출" ${div === "매출" ? "selected" : ""}>매출 (외상매출금 108)</option>
          <option value="매입" ${div === "매입" ? "selected" : ""}>매입 (외상매입금 251)</option>
          <option value="미지급" ${div === "미지급" ? "selected" : ""}>미지급금</option>
        </select>
        <select id="settlementYearFilter">${yearOpts}</select>
        <span class="stl-count muted">${searched.length}건 ${needsBadge}</span>
      </div>
      <div class="table-responsive">
        <table class="stl-table">
          <thead>
            <tr>
              <th class="stl-th-name">거래처명</th>
              <th class="stl-th-code">코드</th>
              <th class="stl-th-ym">귀속연월</th>
              <th class="stl-th-num">발생합계</th>
              <th class="stl-th-num">충당액</th>
              <th class="stl-th-num">잔액</th>
            </tr>
          </thead>
          <tbody>${searched.length ? rowsHtml : emptyRow}</tbody>
          ${searched.length ? `
          <tfoot>
            <tr class="stl-grand">
              <td colspan="3" style="text-align:right;font-weight:700;">합 계</td>
              <td class="num">${formatNumber(grandDev)}</td>
              <td class="num">${formatNumber(grandChung)}</td>
              <td class="num ${grandJan > 0 ? "stl-jan-pos" : grandJan < 0 ? "stl-jan-neg" : ""}">${formatNumber(grandJan)}</td>
            </tr>
          </tfoot>` : ""}
        </table>
      </div>
      ${needsHtml}
    </div>`;
}

function openVendorDaesaModal(code, name, daesaMap, netOffSet) {
  const vendor = daesaMap.get(code);
  if (!vendor) return;
  const isNetOff = netOffSet.has(code);
  const allMonths = Object.keys(vendor.months).sort((a, b) => b.localeCompare(a)); // 최신순

  // 1. 전체 합계 계산 (섹션 노출 여부 결정용)
  const totals = allMonths.reduce((acc, ym) => {
    const d = vendor.months[ym];
    const ledgerBuyTotal = d.ledgerBuy + d.ledgerPayable;
    const ledgerPayTotal = d.ledgerPay + d.ledgerPayablePay;
    const netoffAmt = isNetOff ? Math.min(d.taxSales, d.taxPurchase) : 0;
    // 잔액 기준: 세금계산서 → 원장 → 영업 순 fallback
    const effSales = d.taxSales || d.ledgerSales || d.bizSales;
    const effBuy   = d.taxPurchase || ledgerBuyTotal || d.bizPurchase;
    acc.taxSales += d.taxSales;
    acc.ledgerSales += d.ledgerSales;
    acc.bizSales += d.bizSales;
    acc.collect += d.ledgerCollect || d.bizCollect;
    acc.netoff += netoffAmt;
    acc.taxBuy += d.taxPurchase;
    acc.ledgerBuy += ledgerBuyTotal;
    acc.bizBuy += d.bizPurchase;
    acc.pay += ledgerPayTotal || d.bizPay;
    acc.balanceSales += effSales - netoffAmt - (d.ledgerCollect || d.bizCollect);
    acc.balanceBuy   += effBuy  - netoffAmt - (ledgerPayTotal || d.bizPay);
    return acc;
  }, { taxSales: 0, ledgerSales: 0, bizSales: 0, collect: 0, netoff: 0, taxBuy: 0, ledgerBuy: 0, bizBuy: 0, pay: 0, balanceSales: 0, balanceBuy: 0 });

  const showSales = (totals.taxSales || totals.ledgerSales || totals.bizSales || totals.collect);
  const showBuy = (totals.taxBuy || totals.ledgerBuy || totals.bizBuy || totals.pay);
  const groupColspan = isNetOff ? 6 : 5;
  const netSalesTotal = totals.taxSales - totals.netoff;
  const netBuyTotal = totals.taxBuy - totals.netoff;
  const netOffCols = isNetOff ? '<th>상계</th>' : '';

  // 2. 행 데이터 생성 (Template literals를 조각화하여 구문 오류 방지)
  const rowsHtml = allMonths.map(ym => {
    const d = vendor.months[ym];
    const ledgerBuyTotal = d.ledgerBuy + d.ledgerPayable;
    const ledgerPayTotal = d.ledgerPay + d.ledgerPayablePay;
    const netoffAmt = isNetOff ? Math.min(d.taxSales, d.taxPurchase) : 0;
    const netSales = d.taxSales - netoffAmt;
    const netBuy = d.taxPurchase - netoffAmt;
    const [y, m] = ym.split("-");
    const label = `${y.slice(2)}/${Number(m)}`;
    const useLedger = (d.ledgerCollect || d.ledgerPay || d.ledgerPayablePay);

    const bd = {};
    Object.entries(d.divBreakdownBiz || {}).forEach(([div, info]) => {
      if (!bd[div]) bd[div] = { sales: 0, purchase: 0, collect: 0, pay: 0 };
      bd[div].sales += (info.sales || 0);
      bd[div].purchase += (info.purchase || 0);
    });
    const sourceBD = useLedger ? d.divBreakdownLedger : d.divBreakdownBiz;
    Object.entries(sourceBD || {}).forEach(([div, info]) => {
      if (!bd[div]) bd[div] = { sales: 0, purchase: 0, collect: 0, pay: 0 };
      bd[div].collect += (info.collect || 0);
      bd[div].pay += (info.pay || 0);
    });

    const hasBreakdown = Object.keys(bd).some(div => div !== "");
    const showTaxDetail = (code === "SHOPPINGMALL_SALES") && (Object.keys(d.taxSalesDetail || {}).length > 0);
    const showShoppingLedgerDetail = (code === "SHOPPINGMALL_SALES");
    const hasDetail = hasBreakdown || showTaxDetail || showShoppingLedgerDetail;

    let detailRows = "";
    if (hasDetail) {
      const detailColspan = (showSales ? groupColspan : 0) + (showBuy ? groupColspan : 0) + 1;

      let shoppingLedgerTable = "";
      if (showShoppingLedgerDetail) {
        const sb = d.shoppingBreakdown || { taxInvoice: { sales: 0, collect: 0 }, cashReceipt: { sales: 0, collect: 0 }, unclassified: { sales: 0, collect: 0 } };
        shoppingLedgerTable = `
          <div class="daesa-detail-section">
            <strong>원장 상세 (세금계산서 vs 현금영수증)</strong>
            <table class="daesa-subtable">
              <thead><tr><th>분류</th><th>매출(차변)</th><th>수금(대변)</th><th>잔액</th></tr></thead>
              <tbody>
                <tr><td>세금계산서</td><td class="num">${formatNumber(sb.taxInvoice.sales)}</td><td class="num">${formatNumber(sb.taxInvoice.collect)}</td><td class="num">${formatNumber(sb.taxInvoice.sales - sb.taxInvoice.collect)}</td></tr>
                <tr><td>현금영수증</td><td class="num">${formatNumber(sb.cashReceipt.sales)}</td><td class="num">${formatNumber(sb.cashReceipt.collect)}</td><td class="num">${formatNumber(sb.cashReceipt.sales - sb.cashReceipt.collect)}</td></tr>
                <tr><td>미분류</td><td class="num">-</td><td class="num">${formatNumber(sb.unclassified.collect)}</td><td class="num">${formatNumber(-sb.unclassified.collect)}</td></tr>
              </tbody>
              <tfoot style="background:#f8fafc; font-weight:bold;">
                <tr>
                  <td>합계</td>
                  <td class="num">${formatNumber(sb.taxInvoice.sales + sb.cashReceipt.sales)}</td>
                  <td class="num">${formatNumber(sb.taxInvoice.collect + sb.cashReceipt.collect + sb.unclassified.collect)}</td>
                  <td class="num">${formatNumber((sb.taxInvoice.sales + sb.cashReceipt.sales) - (sb.taxInvoice.collect + sb.cashReceipt.collect + sb.unclassified.collect))}</td>
                </tr>
              </tfoot>
            </table>
          </div>`;
      }

      let breakdownTable = "";
      if (hasBreakdown) {
        const hasBuyBD = Object.values(bd).some(info => info.purchase || info.pay);
        const bdRows = Object.entries(bd).map(([div, info]) => `
          <tr>
            <td>${escapeHtml(div)}</td>
            <td class="num">${formatNumber(info.sales)}</td>
            <td class="num">${formatNumber(info.collect)}</td>
            <td class="num">${formatNumber(info.sales - info.collect)}</td>
            ${hasBuyBD ? `
            <td class="num">${formatNumber(info.purchase)}</td>
            <td class="num">${formatNumber(info.pay)}</td>
            <td class="num">${formatNumber(info.purchase - info.pay)}</td>` : ""}
          </tr>`).join("");
        breakdownTable = `
          <div class="daesa-detail-section">
            <strong>사업부문/현장별 실적</strong>
            <table class="daesa-subtable">
              <thead><tr>
                <th>사업부문명</th>
                <th>매출</th><th>수금</th><th>잔액(매출)</th>
                ${hasBuyBD ? `<th>매입</th><th>지급</th><th>잔액(매입)</th>` : ""}
              </tr></thead>
              <tbody>${bdRows}</tbody>
            </table>
          </div>`;
      }

      let taxDetailTable = "";
      if (showTaxDetail) {
        const tdRows = Object.entries(d.taxSalesDetail).map(([nm, amt]) => `<tr><td>${escapeHtml(nm)}</td><td>${formatNumber(amt)}</td></tr>`).join("");
        taxDetailTable = `
          <div class="daesa-detail-section">
            <strong>세금계산서 상세 (상공업체별)</strong>
            <table class="daesa-subtable">
              <thead><tr><th>업체명</th><th>금액</th></tr></thead>
              <tbody>${tdRows}</tbody>
            </table>
          </div>`;
      }

      detailRows = `
        <tr class="daesa-modal-detail-row hidden" data-ym-detail="${ym}">
          <td colspan="${detailColspan}" class="daesa-modal-detail-cell">
            <div style="display:flex; gap:20px; flex-wrap:wrap;">
              ${shoppingLedgerTable}
              ${taxDetailTable}
              ${breakdownTable}
            </div>
          </td>
        </tr>`;
    }

    const effSalesRow = d.taxSales || d.ledgerSales || d.bizSales;
    const effBuyRow   = d.taxPurchase || ledgerBuyTotal || d.bizPurchase;
    const salesCols = showSales ? `
      <td class="num">${formatNumber(d.taxSales)}</td>
      <td class="num">${formatNumber(d.ledgerSales)}</td>
      <td class="num">${formatNumber(d.bizSales)}</td>
      <td class="num daesa-collect">${formatNumber(d.ledgerCollect || d.bizCollect)}</td>
      ${isNetOff ? `<td class="num daesa-netoff-amt">${formatNumber(netoffAmt)}</td>` : ""}
      <td class="num daesa-balance">${formatNumber(effSalesRow - netoffAmt - (d.ledgerCollect || d.bizCollect))}</td>
    ` : "";

    const buyCols = showBuy ? `
      <td class="num">${formatNumber(d.taxPurchase)}</td>
      <td class="num">${formatNumber(ledgerBuyTotal)}</td>
      <td class="num">${formatNumber(d.bizPurchase)}</td>
      <td class="num daesa-pay">${formatNumber(ledgerPayTotal || d.bizPay)}</td>
      ${isNetOff ? `<td class="num daesa-netoff-amt">${formatNumber(netoffAmt)}</td>` : ""}
      <td class="num daesa-balance">${formatNumber(effBuyRow - netoffAmt - (ledgerPayTotal || d.bizPay))}</td>
    ` : "";

    return `
      <tr class="daesa-modal-ym-row ${hasDetail ? "has-detail" : ""}" data-ym="${ym}">
        <td class="daesa-modal-ym">${label}${hasDetail ? '<span class="daesa-expand-icon">▼</span>' : ""}</td>
        ${salesCols}
        ${buyCols}
      </tr>
      ${detailRows}
    `;
  }).join("");

  // 3. 모달 생성 및 이벤트 연결
  const overlay = document.createElement("div");
  overlay.className = "daesa-modal-overlay";
  overlay.innerHTML = `
    <div class="daesa-modal">
      <div class="daesa-modal-header">
        <h3>${escapeHtml(name)} — 누적 대사 현황</h3>
        ${isNetOff ? '<span class="daesa-netoff-badge">상계 업체</span>' : ''}
        <button class="daesa-modal-close">✕</button>
      </div>
      <div class="daesa-modal-body">
        <p class="di-desc" style="margin-bottom:10px;">* 각 행을 클릭하면 상공업체 상세 및 사업부문별 내역을 볼 수 있습니다.</p>
        <div class="table-responsive">
          <table class="daesa-modal-table">
            <thead>
              <tr>
                <th rowspan="2">년/월</th>
                ${showSales ? `<th colspan="${groupColspan}" class="daesa-th-group daesa-th-sales">매출</th>` : ""}
                ${showBuy ? `<th colspan="${groupColspan}" class="daesa-th-group daesa-th-purchase">매입</th>` : ""}
              </tr>
              <tr>
                ${showSales ? `<th>세금계산서</th><th>원장</th><th>영업</th><th>수금</th>${netOffCols}<th>잔액</th>` : ""}
                ${showBuy ? `<th>세금계산서</th><th>원장</th><th>영업</th><th>지급</th>${netOffCols}<th>잔액</th>` : ""}
              </tr>
            </thead>
            <tbody>${rowsHtml}</tbody>
            <tfoot>
              <tr class="daesa-total-row">
                <td>합계</td>
                ${showSales ? `
                  <td class="num">${formatNumber(totals.taxSales)}</td>
                  <td class="num">${formatNumber(totals.ledgerSales)}</td>
                  <td class="num">${formatNumber(totals.bizSales)}</td>
                  <td class="num daesa-collect">${formatNumber(totals.collect)}</td>
                  ${isNetOff ? `<td class="num daesa-netoff-amt">${formatNumber(totals.netoff)}</td>` : ""}
                  <td class="num daesa-balance">${formatNumber(totals.balanceSales)}</td>
                ` : ""}
                ${showBuy ? `
                  <td class="num">${formatNumber(totals.taxBuy)}</td>
                  <td class="num">${formatNumber(totals.ledgerBuy)}</td>
                  <td class="num">${formatNumber(totals.bizBuy)}</td>
                  <td class="num daesa-pay">${formatNumber(totals.pay)}</td>
                  ${isNetOff ? `<td class="num daesa-netoff-amt">${formatNumber(totals.netoff)}</td>` : ""}
                  <td class="num daesa-balance">${formatNumber(totals.balanceBuy)}</td>
                ` : ""}
              </tr>
            </tfoot>
          </table>
        </div>
      </div>
      <div class="daesa-modal-footer">
        <button class="daesa-modal-print">인쇄 / PDF</button>
        <button class="daesa-modal-close">닫기</button>
      </div>
    </div>
  `;
  document.body.appendChild(overlay);

  overlay.querySelectorAll(".daesa-modal-ym-row.has-detail").forEach(row => {
    row.addEventListener("click", () => {
      const ym = row.dataset.ym;
      const detail = overlay.querySelector(`tr[data-ym-detail="${ym}"]`);
      detail?.classList.toggle("hidden");
      row.querySelector(".daesa-expand-icon").textContent = detail?.classList.contains("hidden") ? "▼" : "▲";
    });
  });

  overlay.querySelectorAll(".daesa-modal-close").forEach(b =>
    b.addEventListener("click", () => overlay.remove())
  );
  overlay.addEventListener("mousedown", e => { if (e.target === overlay) overlay.remove(); });
  overlay.querySelector(".daesa-modal-print")?.addEventListener("click", () => {
    // 업체마스터에서 기본정보 조회
    const vmRow = vendorMasterState.rows.find(v =>
      String(v["거래처코드_norm"] || v["거래처코드_raw"] || "").trim() === code
    );
    const bizNo    = vmRow ? String(vmRow["사업자번호"] || "").trim() : "";
    const category = vmRow ? String(vmRow["거래처구분"] || "").trim() : "";
    const bank     = vmRow ? String(vmRow["은행"] || "").trim() : "";
    const account  = vmRow ? String(vmRow["계좌번호"] || "").trim() : "";

    // 미수금에서 담당자 조회
    const rcvItem = receivables.find(r =>
      (r.code === code || r.codeNormalized === code) && r.manager && r.manager !== "미지정"
    );
    const managerName = rcvItem?.manager || "";

    const today = new Date();
    const dateStr = `${today.getFullYear()}년 ${today.getMonth()+1}월 ${today.getDate()}일`;

    const printWin = window.open("", "_blank", "width=1100,height=750");
    const tableHtml = overlay.querySelector("table").cloneNode(true);
    tableHtml.querySelectorAll(".daesa-modal-detail-row").forEach(r => r.classList.remove("hidden"));

    const infoItems = [
      ["거래처명", name],
      ["거래처코드", code],
      bizNo    ? ["사업자번호", bizNo]    : null,
      category ? ["거래처구분", category] : null,
      bank     ? ["은행",     bank]       : null,
      account  ? ["계좌번호", account]    : null,
      managerName ? ["담당자", managerName] : null,
    ].filter(Boolean);

    const infoHtml = infoItems.map(([label, val]) =>
      `<div class="vi-item"><span class="vi-label">${label}</span>${escapeHtml(String(val))}</div>`
    ).join("");

    printWin.document.write(`<!DOCTYPE html><html><head><meta charset="UTF-8">
      <title>${escapeHtml(name)} 대사 현황</title>
      <style>
        *{box-sizing:border-box;}
        body{font-family:'맑은 고딕',sans-serif;font-size:13px;margin:18px 22px;color:#1a1a1a;}
        .rpt-header{display:flex;justify-content:space-between;align-items:flex-end;padding-bottom:10px;border-bottom:2px solid #1e3a5f;margin-bottom:12px;}
        .rpt-title{font-size:17px;font-weight:bold;color:#1e3a5f;}
        .rpt-company{font-size:11px;color:#6b7280;margin-bottom:3px;}
        .rpt-date{font-size:12px;color:#6b7280;text-align:right;}
        .vi-box{display:flex;flex-wrap:wrap;gap:4px 20px;padding:9px 14px;border:1px solid #d1d5db;border-radius:4px;background:#f8fafc;margin-bottom:14px;font-size:12px;}
        .vi-item{white-space:nowrap;}
        .vi-label{font-weight:600;color:#374151;margin-right:4px;}
        table{border-collapse:collapse;width:100%;font-size:12px;}
        th,td{border:1px solid #d1d5db;padding:5px 8px;text-align:right;}
        th{background:#f1f5f9;text-align:center;font-weight:600;}
        td:first-child{text-align:center;white-space:nowrap;}
        .daesa-th-group{font-size:13px;}
        .daesa-th-sales{background:#dbeafe;}
        .daesa-th-purchase{background:#fee2e2;}
        .daesa-balance{background:#fef9c3;font-weight:600;}
        .daesa-collect{color:#1565c0;}
        .daesa-pay{color:#b71c1c;}
        .daesa-modal-detail-row{background:#fafafa;}
        .daesa-modal-detail-cell{text-align:left;padding:5px 12px;}
        tfoot tr{background:#e8edf8;font-weight:bold;}
        .daesa-expand-icon{display:none;}
        .rpt-footer{margin-top:16px;padding-top:6px;border-top:1px solid #e5e7eb;font-size:11px;color:#9ca3af;text-align:center;}
        @media print{@page{size:A4 landscape;margin:10mm 8mm;}}
      </style></head><body>
      <div class="rpt-header">
        <div>
          <div class="rpt-company">미래오토메이션(주)</div>
          <div class="rpt-title">거래처 대사 현황 보고서${isNetOff ? " · 상계 업체" : ""}</div>
        </div>
        <div class="rpt-date">기준일: ${dateStr}</div>
      </div>
      <div class="vi-box">${infoHtml}</div>
      ${tableHtml.outerHTML}
      <div class="rpt-footer">미래오토메이션(주) · 출력일: ${dateStr}</div>
      </body></html>`);
    printWin.document.close();
    printWin.focus();
    setTimeout(() => printWin.print(), 500);
  });
}

// ── 자료 업로드 ─────────────────────────────────────────────

let dataImportState = {
  visible: false,
  taxInvoice: { parsed: null, status: "", saving: false },
  ledgerSales: { parsed: null, status: "", saving: false },
  ledgerPurchase: { parsed: null, status: "", saving: false },
  ledgerPayable: { parsed: null, status: "", saving: false },
  dailySales: { parsed: null, status: "", saving: false },
};

// 마스터 관리 메뉴 설정
function setupMasterMenu() {
  const menuBtn = document.getElementById("masterMenuButton");
  const menu = document.getElementById("masterDropdownMenu");
  const bizBtn = document.getElementById("bizDivImportButton");
  const bizInput = document.getElementById("bizDivMasterFileInput");

  if (!menuBtn || !menu) return;

  // 메뉴 토글
  menuBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    menu.classList.toggle("visible");
  });

  // 바깥 클릭 시 메뉴 닫기
  document.addEventListener("click", () => menu.classList.remove("visible"));
  menu.addEventListener("click", (e) => e.stopPropagation());

  // 사업부문 마스터 업로드 트리거
  if (bizBtn && bizInput) {
    bizBtn.addEventListener("click", () => {
      menu.classList.remove("visible");
      bizInput.click();
    });

    bizInput.addEventListener("change", async (e) => {
      const file = e.target.files[0];
      if (!file) return;

      showToast(`${file.name} 분석 중…`);
      const result = await parseBizDivisionFile(file);

      if (!result.ok) {
        alert(`분석 실패: ${result.error}`);
        return;
      }

      if (confirm(`${result.count}건의 사업부문을 찾았습니다. 구글 시트에 저장하시겠습니까?`)) {
        try {
          showToast("시트에 저장 중…");
          await postSheetWebApp("upsertBizDivision", { rows: result.rows });
          alert("✓ 사업부문 마스터 저장 완료");
          // 로컬 상태 업데이트는 loadDaesaData 등에서 자동 처리되도록 유도하거나 직접 업데이트
          bizDivisionState.rows = result.rows;
        } catch (err) {
          alert(`저장 실패: ${err.message}`);
        }
      }
      bizInput.value = ""; // 초기화
    });
  }
}

// ════════════════════════════════════════════════════════════
//  분류규칙 관리 (Phase 0)
// ════════════════════════════════════════════════════════════

const rulesState = {
  rows: [],           // 로드된 규칙 배열
  loading: false,
  saving: false,
  msg: "",
  bizFilter: "전체",  // "전체" | "엠오토" | "미래"
  editKey: null,      // 현재 편집 중인 _rule_key (null=신규 추가 폼)
  addingNew: false,   // 신규 추가 폼 표시 여부
  tableOpen: false,   // 아코디언 — 기본 접힘
};

function buildRuleKey(사업체, 매칭방식, 매칭키) {
  return `${String(사업체||"").trim()}||${String(매칭방식||"").trim()}||${String(매칭키||"").trim()}`;
}

async function fetchRulesFromApi(bizFilter) {
  if (!SHEET_APP_SCRIPT_URL) throw new Error("Apps Script URL 없음");
  const url = new URL(SHEET_APP_SCRIPT_URL);
  url.searchParams.set("action", "getRules");
  const token = getApiToken();
  if (token) url.searchParams.set("token", token);
  if (bizFilter && bizFilter !== "전체") url.searchParams.set("사업체", bizFilter);
  const res = await fetch(url.toString());
  if (!res.ok) throw new Error(`분류규칙 조회 실패: ${res.status}`);
  const body = await res.json();
  return Array.isArray(body.rows) ? body.rows : [];
}

async function loadRules() {
  rulesState.loading = true;
  rulesState.msg = "불러오는 중…";
  renderRulesPanel();
  try {
    rulesState.rows = await fetchRulesFromApi(rulesState.bizFilter);
    rulesState.msg = `${rulesState.rows.length}건 로드됨`;
  } catch (e) {
    rulesState.msg = `조회 실패: ${e.message}`;
  } finally {
    rulesState.loading = false;
    renderRulesPanel();
  }
}

async function saveRule(ruleObj) {
  rulesState.saving = true;
  rulesState.msg = "저장 중…";
  renderRulesPanel();
  try {
    const key = buildRuleKey(ruleObj["사업체"], ruleObj["매칭방식"], ruleObj["매칭키"]);
    const row = { ...ruleObj, _rule_key: key };
    await postSheetWebApp("upsertRules", { rows: [row] });
    // 로컬 상태 갱신
    const idx = rulesState.rows.findIndex(r => r["_rule_key"] === key);
    if (idx >= 0) rulesState.rows[idx] = row;
    else rulesState.rows.push(row);
    rulesState.msg = "저장 완료";
    rulesState.editKey = null;
    rulesState.addingNew = false;
  } catch (e) {
    rulesState.msg = `저장 실패: ${e.message}`;
  } finally {
    rulesState.saving = false;
    renderRulesPanel();
  }
}

async function deleteRule(key) {
  if (!confirm(`규칙을 삭제하시겠습니까?\n${key}`)) return;
  rulesState.saving = true;
  rulesState.msg = "삭제 중…";
  renderRulesPanel();
  try {
    await postSheetWebApp("deleteRule", { key });
    rulesState.rows = rulesState.rows.filter(r => r["_rule_key"] !== key);
    rulesState.msg = "삭제 완료";
    rulesState.editKey = null;
  } catch (e) {
    rulesState.msg = `삭제 실패: ${e.message}`;
  } finally {
    rulesState.saving = false;
    renderRulesPanel();
  }
}

function renderRulesPanel() {
  const panel = document.getElementById("rulesPanel");
  if (!panel) return;

  const BIZ_OPTIONS = ["전체", "엠오토", "미래"];
  const METHOD_OPTIONS = ["계좌", "키워드", "거래처명"];
  const DIV_OPTIONS = ["", "매출", "매입"];

  const filtered = rulesState.bizFilter === "전체"
    ? rulesState.rows
    : rulesState.rows.filter(r => String(r["사업체"] || "").trim() === rulesState.bizFilter);

  // 인라인 편집 폼 HTML
  function editForm(data = {}) {
    const dis = rulesState.saving ? "disabled" : "";
    const sKey = data["_rule_key"] || "";
    return `
      <tr class="rules-edit-row" data-edit-key="${escapeAttr(sKey)}">
        <td><select class="rules-inp" name="사업체" ${dis}>
          ${["엠오토","미래"].map(v => `<option${data["사업체"]===v?" selected":""}>${v}</option>`).join("")}
        </select></td>
        <td><select class="rules-inp" name="매칭방식" ${dis}>
          ${METHOD_OPTIONS.map(v => `<option${data["매칭방식"]===v?" selected":""}>${v}</option>`).join("")}
        </select></td>
        <td><input class="rules-inp" name="매칭키" value="${escapeAttr(data["매칭키"]||"")}" placeholder="계좌번호 또는 키워드" ${dis} /></td>
        <td><input class="rules-inp" name="거래처명" value="${escapeAttr(data["거래처명"]||"")}" placeholder="거래처명" ${dis} /></td>
        <td><select class="rules-inp" name="구분" ${dis}>
          ${DIV_OPTIONS.map(v => `<option value="${v}"${data["구분"]===v?" selected":""}>${v || "(없음)"}</option>`).join("")}
        </select></td>
        <td><input class="rules-inp rules-inp-sm" name="결제예정일" type="number" min="1" max="31" placeholder="없음" value="${escapeAttr(String(data["결제예정일"]||""))}" ${dis} style="width:52px;" title="매월 N일 결제 (고정지출 자동계산용)" /></td>
        <td><input class="rules-inp" name="지급월" placeholder="1,7 또는 3,6,9,12" value="${escapeAttr(data["지급월"]||"")}" ${dis} style="width:90px;" title="빈칸=매월 / 1,7=반기 / 3,6,9,12=분기" /></td>
        <td><select class="rules-inp" name="활성여부" ${dis} style="width:48px;">
          <option value="Y"${String(data["활성여부"]||"Y").toUpperCase()!=="N"?" selected":""}>Y</option>
          <option value="N"${String(data["활성여부"]||"Y").toUpperCase()==="N"?" selected":""}>N</option>
        </select></td>
        <td><input class="rules-inp" name="고정분류" placeholder="이자/인출금/카드" value="${escapeAttr(data["고정분류"]||"")}" ${dis} style="width:80px;" /></td>
        <td><input class="rules-inp rules-inp-sm" name="예정금액" type="number" min="0" placeholder="평균금액" value="${escapeAttr(String(data["예정금액"]||""))}" ${dis} style="width:80px;" /></td>
        <td><input class="rules-inp rules-inp-sm" name="우선순위" type="number" min="1" value="${escapeAttr(String(data["우선순위"]||"10"))}" ${dis} /></td>
        <td>
          <button class="rules-btn rules-save-btn" data-key="${escapeAttr(sKey)}" ${dis}>저장</button>
          <button class="rules-btn rules-cancel-btn" data-key="${escapeAttr(sKey)}" ${dis}>취소</button>
        </td>
      </tr>`;
  }

  function escapeAttr(v) {
    return String(v).replace(/&/g,"&amp;").replace(/"/g,"&quot;").replace(/</g,"&lt;");
  }

  const rowsHtml = filtered.map(r => {
    const key = r["_rule_key"] || buildRuleKey(r["사업체"],r["매칭방식"],r["매칭키"]);
    if (rulesState.editKey === key) return editForm({ ...r, _rule_key: key });
    const dis = rulesState.saving ? "disabled" : "";
    return `
      <tr>
        <td>${escapeAttr(r["사업체"]||"")}</td>
        <td>${escapeAttr(r["매칭방식"]||"")}</td>
        <td class="rules-key-cell" title="${escapeAttr(r["매칭키"]||"")}">${escapeAttr(r["매칭키"]||"")}</td>
        <td>${escapeAttr(r["거래처명"]||"")}</td>
        <td>${escapeAttr(r["구분"]||"")}</td>
        <td style="text-align:right;color:${r["결제예정일"]?"#2563eb":"#9ca3af"};">${r["결제예정일"] ? `매월 ${r["결제예정일"]}일` : ""}</td>
        <td style="color:#6b7280;font-size:11px;">${escapeAttr(r["지급월"]||"")}</td>
        <td style="text-align:center;">${String(r["활성여부"]||"Y").toUpperCase()==="N" ? '<span style="color:#ef4444;font-size:11px;">N</span>' : '<span style="color:#16a34a;font-size:11px;">Y</span>'}</td>
        <td style="color:#6b7280;">${escapeAttr(r["고정분류"]||"")}</td>
        <td style="text-align:right;color:#6b7280;">${r["예정금액"] ? formatNumber(Number(r["예정금액"])) : ""}</td>
        <td style="text-align:right;">${escapeAttr(String(r["우선순위"]||""))}</td>
        <td>
          <button class="rules-btn rules-edit-btn" data-key="${escapeAttr(key)}" ${dis}>수정</button>
          <button class="rules-btn rules-del-btn" data-key="${escapeAttr(key)}" ${dis}>삭제</button>
        </td>
      </tr>`;
  }).join("");

  const newRowHtml = rulesState.addingNew ? editForm({}) : "";

  const toggleIcon = rulesState.tableOpen ? "▼" : "▶";
  const countBadge = rulesState.rows.length
    ? `<span style="font-size:12px;color:#6b7280;margin-left:6px;">${filtered.length}건${rulesState.bizFilter !== "전체" ? ` (${rulesState.bizFilter})` : ""}</span>`
    : "";

  // 스크롤 위치 보존 (재렌더 후 복원)
  const _prevScroll = panel.querySelector(".rules-table-wrap")?.scrollTop || 0;

  panel.innerHTML = `
    <div class="rules-panel-inner">
      <div class="rules-toolbar">
        <div style="display:flex;align-items:center;gap:6px;flex-wrap:nowrap;">
          <button id="rulesToggleBtn" class="rules-btn" style="min-width:28px;">${toggleIcon}</button>
          <strong style="font-size:14px;white-space:nowrap;">분류규칙 관리</strong>
          ${countBadge}
          ${rulesState.tableOpen ? `
          <span style="color:#d1d5db;">|</span>
          ${BIZ_OPTIONS.map(b =>
            `<button class="rules-biz-btn${rulesState.bizFilter===b?" active":""}" data-biz="${b}">${b}</button>`
          ).join("")}
          <button class="rules-btn" id="rulesReloadBtn" ${rulesState.loading?"disabled":""}>새로고침</button>
          <button class="rules-btn rules-add-btn" id="rulesAddBtn" ${rulesState.addingNew||rulesState.saving?"disabled":""}>+ 추가</button>
          <label class="rules-btn rules-import-btn" title="Excel 파일에서 규칙 일괄 가져오기" style="cursor:pointer;">
            Excel 가져오기
            <input type="file" id="rulesImportFileInput" accept=".xls,.xlsx" hidden />
          </label>` : ""}
        </div>
        <button class="rules-btn rules-close-btn" id="rulesPanelClose">✕ 닫기</button>
      </div>
      ${rulesState.tableOpen ? `
        ${rulesState.msg ? `<div class="rules-msg">${escapeAttr(rulesState.msg)}</div>` : ""}
        ${rulesState.loading ? `<div class="rules-msg">불러오는 중…</div>` : `
        <div class="rules-table-wrap">
          <table class="rules-table">
            <thead><tr>
              <th>사업체</th><th>매칭방식</th><th>매칭키</th><th>거래처명</th><th>구분</th><th>결제예정일</th><th>지급월</th><th>활성</th><th>고정분류</th><th>예정금액</th><th>우선순위</th><th>액션</th>
            </tr></thead>
            <tbody>
              ${rowsHtml}
              ${newRowHtml}
              ${!filtered.length && !rulesState.addingNew ? `<tr><td colspan="12" style="text-align:center;color:#9ca3af;padding:16px;">규칙 없음 — "+ 추가" 버튼으로 추가하세요</td></tr>` : ""}
            </tbody>
          </table>
        </div>`}` : ""}
    </div>`;

  // 스크롤 위치 복원
  if (_prevScroll) {
    const wrap = panel.querySelector(".rules-table-wrap");
    if (wrap) wrap.scrollTop = _prevScroll;
  }

  // 이벤트 바인딩
  panel.querySelector("#rulesToggleBtn")?.addEventListener("click", () => {
    rulesState.tableOpen = !rulesState.tableOpen;
    if (rulesState.tableOpen && !rulesState.rows.length && !rulesState.loading) {
      loadRules();
    } else {
      renderRulesPanel();
    }
  });

  panel.querySelector("#rulesPanelClose")?.addEventListener("click", () => {
    panel.classList.add("hidden");
    rulesState.editKey = null;
    rulesState.addingNew = false;
  });

  panel.querySelector("#rulesReloadBtn")?.addEventListener("click", loadRules);

  panel.querySelector("#rulesImportFileInput")?.addEventListener("change", e => {
    const file = e.target.files?.[0];
    if (file) importRulesFromExcel(file);
    e.target.value = "";
  });

  panel.querySelector("#rulesAddBtn")?.addEventListener("click", () => {
    rulesState.addingNew = true;
    rulesState.editKey = null;
    rulesState.tableOpen = true;
    renderRulesPanel();
    panel.querySelector(".rules-edit-row input")?.focus();
  });

  panel.querySelectorAll(".rules-biz-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      rulesState.bizFilter = btn.dataset.biz;
      renderRulesPanel();
    });
  });

  panel.querySelectorAll(".rules-edit-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      rulesState.editKey = btn.dataset.key;
      rulesState.addingNew = false;
      renderRulesPanel();
      panel.querySelector(".rules-edit-row input")?.focus();
    });
  });

  panel.querySelectorAll(".rules-del-btn").forEach(btn => {
    btn.addEventListener("click", () => deleteRule(btn.dataset.key));
  });

  panel.querySelectorAll(".rules-cancel-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      rulesState.editKey = null;
      rulesState.addingNew = false;
      renderRulesPanel();
    });
  });

  panel.querySelectorAll(".rules-save-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      const row = btn.closest("tr.rules-edit-row");
      if (!row) return;
      const get = name => row.querySelector(`[name="${name}"]`)?.value?.trim() || "";
      const ruleObj = {
        사업체: get("사업체"),
        매칭방식: get("매칭방식"),
        매칭키: get("매칭키"),
        거래처명: get("거래처명"),
        구분: get("구분"),
        결제예정일: get("결제예정일") || "",
        지급월: get("지급월") || "",
        활성여부: get("활성여부") || "Y",
        고정분류: get("고정분류") || "",
        예정금액: get("예정금액") || "",
        우선순위: get("우선순위") || "10",
      };
      if (!ruleObj.매칭키) { alert("매칭키를 입력해주세요."); return; }
      if (!ruleObj.거래처명) { alert("거래처명을 입력해주세요."); return; }
      saveRule(ruleObj);
    });
  });
}

async function importRulesFromExcel(file) {
  rulesState.saving = true;
  rulesState.msg = "Excel 파싱 중…";
  renderRulesPanel();
  try {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const raw = XLSX.utils.sheet_to_json(ws, { defval: "" });
    if (!raw.length) throw new Error("데이터 없음");

    const rows = raw.map(r => {
      const ruleObj = {
        사업체:     String(r["사업체"]     ?? "").trim(),
        매칭방식:   String(r["매칭방식"]   ?? "").trim(),
        매칭키:     String(r["매칭키"]     ?? "").trim(),
        거래처명:   String(r["거래처명"]   ?? "").trim(),
        구분:       String(r["구분"]       ?? "").trim(),
        우선순위:   String(r["우선순위"]   ?? "10").trim() || "10",
        결제예정일: String(r["결제예정일"] ?? "").trim(),
        고정분류:   String(r["고정분류"]   ?? "").trim(),
        예정금액:   String(r["예정금액"]   ?? "").trim(),
        지급월:     String(r["지급월"]     ?? "").trim(),
        활성여부:   String(r["활성여부"]   ?? "Y").trim() || "Y",
      };
      ruleObj._rule_key = buildRuleKey(ruleObj["사업체"], ruleObj["매칭방식"], ruleObj["매칭키"]);
      return ruleObj;
    }).filter(r => r.매칭키 && r.거래처명);

    if (!rows.length) throw new Error("유효한 행 없음 (매칭키·거래처명 필수)");

    await postSheetWebApp("upsertRules", { rows });
    rulesState.msg = `${rows.length}건 가져오기 완료`;
    await loadRules();
    // 새 규칙으로 저장된 입출금 전체 재분류 (거래처명 매칭 반영)
    if (typeof rebuildMautoRows === "function") rebuildMautoRows();
  } catch (e) {
    rulesState.msg = `가져오기 실패: ${e.message}`;
  } finally {
    rulesState.saving = false;
    renderRulesPanel();
  }
}

function setupRulesPanel() {
  const btn = document.getElementById("rulesManageButton");
  const menu = document.getElementById("masterDropdownMenu");
  const panel = document.getElementById("rulesPanel");
  if (!btn || !panel) return;

  btn.addEventListener("click", () => {
    menu?.classList.remove("visible");
    const isOpen = !panel.classList.contains("hidden");
    if (isOpen) {
      panel.classList.add("hidden");
    } else {
      panel.classList.remove("hidden");
      renderRulesPanel(); // 접힌 상태로 표시, 로드는 ▶ 펼치기 시
    }
  });
}

function formatExcelDateToStr(val) {
  if (val instanceof Date && !isNaN(val)) {
    const y = val.getFullYear();
    const m = String(val.getMonth() + 1).padStart(2, "0");
    const d = String(val.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }
  if (typeof val === "number" && val > 10000) {
    const d = new Date(Math.round((val - 25569) * 86400000));
    const y = d.getUTCFullYear();
    const mo = String(d.getUTCMonth() + 1).padStart(2, "0");
    const da = String(d.getUTCDate()).padStart(2, "0");
    return `${y}-${mo}-${da}`;
  }
  return String(val || "").trim();
}

function normalizeBizNum(bn) {
  return String(bn || "").replace(/[^0-9]/g, "");
}

function buildVendorLookupMaps() {
  const byBiz = {};
  const byCode = {};
  vendorMasterState.rows.forEach(v => {
    const bn = normalizeBizNum(v["사업자번호"] || v["사업자(주민)번호"] || v.businessNumber || "");
    const code = String(v["거래처코드_norm"] || v["거래처코드_raw"] || "").trim().replace(/^0+/, "");
    const entry = { code: v["거래처코드_norm"] || "", name: v["거래처명"] || "" };
    if (bn) byBiz[bn] = entry;
    if (code) byCode[code] = entry;
  });
  return { byBiz, byCode };
}

function matchVendorEntry(bizNum, code, maps) {
  const bn = normalizeBizNum(bizNum);
  if (bn && maps.byBiz[bn]) return maps.byBiz[bn];
  const c = String(code || "").trim().replace(/^0+/, "");
  if (c && maps.byCode[c]) return maps.byCode[c];
  return null;
}

// 거래처명(정규화) → { code, name, bizNum } 맵 생성
function buildVendorNameMap() {
  const byName = {}; // normalizedName → { code, name, bizNum }
  vendorMasterState.rows.forEach(v => {
    const name = String(v["거래처명"] || "").trim();
    const norm = normalizeVendorName(name);
    if (norm) byName[norm] = {
      code: v["거래처코드_norm"] || "",
      name,
      bizNum: normalizeBusinessNumber(v["사업자번호"] || v["사업자(주민)번호"] || ""),
    };
  });
  return byName;
}

function parseXlsToRows(arrayBuffer, headerRowIndex) {
  const wb = XLSX.read(arrayBuffer, { type: "array", cellDates: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const allRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (allRows.length <= headerRowIndex) throw new Error("헤더 행이 없습니다.");
  const headers = allRows[headerRowIndex].map(h => String(h).trim());
  const dataRows = [];
  for (let i = headerRowIndex + 1; i < allRows.length; i++) {
    const raw = allRows[i];
    if (raw.every(v => v === "" || v == null)) continue;
    const row = {};
    headers.forEach((h, j) => {
      let val = raw[j] ?? "";
      // 숫자이지만 셀 서식에 앞자리 0이 있으면(거래처코드 등) 서식 텍스트를 사용
      if (typeof val === "number") {
        const cellAddr = XLSX.utils.encode_cell({ r: i, c: j });
        const cell = ws[cellAddr];
        if (cell && cell.w && /^0\d/.test(cell.w)) val = cell.w;
      }
      row[h] = val;
    });
    dataRows.push(row);
  }
  return { headers, dataRows };
}

async function parseTaxInvoiceFile(file) {
  try {
    const ab = await file.arrayBuffer();
    const { dataRows } = parseXlsToRows(ab, 6); // 7행 헤더
    const maps = buildVendorLookupMaps();
    return {
      rows: dataRows.map(row => {
        if (row["작성일자"]) row["작성일자"] = formatExcelDateToStr(row["작성일자"]);
        if (row["발급일자"]) row["발급일자"] = formatExcelDateToStr(row["발급일자"]);
        const v = matchVendorEntry(row["사업자(주민)번호"], "", maps);
        row["_matched_code"] = v?.code || "";
        row["_matched_name"] = v?.name || "";
        const approvalNum = String(row["승인번호"] || "").trim();
        row["_row_key"] = approvalNum ||
          `${row["작성일자"]}_${normalizeBizNum(row["사업자(주민)번호"])}_${row["합계"]}`;
        return row;
      }),
      error: null,
    };
  } catch (err) {
    return { rows: null, error: err.message };
  }
}

async function parseLedgerFile(file) {
  try {
    const ab = await file.arrayBuffer();
    const { dataRows } = parseXlsToRows(ab, 0); // 1행 헤더
    const maps = buildVendorLookupMaps();
    return {
      rows: dataRows.map(row => {
        if (row["일자"]) row["일자"] = formatExcelDateToStr(row["일자"]);
        const v = matchVendorEntry(row["사업자(주민)번호"], row["거래처코드"], maps);
        row["_matched_code"] = v?.code || "";
        row["_matched_name"] = v?.name || "";
        row["_row_key"] =
          `${row["일자"]}_${String(row["전표번호"] || row["견표번호"] || "").trim()}_${String(row["거래처코드"] || "").trim()}`;
        return row;
      }),
      error: null,
    };
  } catch (err) {
    return { rows: null, error: err.message };
  }
}

async function parseDailySalesFile(file) {
  try {
    const ab = await file.arrayBuffer();
    const { dataRows } = parseXlsToRows(ab, 7); // 8행 헤더
    const maps = buildVendorLookupMaps();
    return {
      rows: dataRows.map(row => {
        if (row["거래일자"]) row["거래일자"] = formatExcelDateToStr(row["거래일자"]);
        const v = matchVendorEntry("", row["거래처코드"], maps);
        row["_matched_code"] = v?.code || "";
        row["_matched_name"] = v?.name || "";
        const txNum = String(row["전표번호"] || row["전포번호"] || "").trim();
        row["_row_key"] = txNum
          ? `${row["거래일자"]}_${txNum}`
          : `${row["거래일자"]}_${String(row["거래처코드"] || "").trim()}_${row["판매금액"]}_${row["구매금액"]}`;
        return row;
      }),
      error: null,
    };
  } catch (err) {
    return { rows: null, error: err.message };
  }
}

// ── 엠오토 세금계산서 파서 (국세청 전자세금계산서 조회 XLS 양식) ──
// 헤더: 6행(index 5), 상호(col)=거래처(상대방), 마지막 행=합계(체크섬)
// sideType: "매출"|"매입" — 중복 상호 컬럼 중 어느 쪽을 거래처로 볼지 결정
//   매입: 공급자(첫번째 상호) = 매입처, 매출: 공급받는자(마지막 상호) = 매출처
async function parseMautoTaxInvoiceFile(file, sideType) {
  try {
    const ab = await file.arrayBuffer();
    const wb = XLSX.read(ab, { type: "array", cellDates: true });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const allRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });

    const HEADER_IDX = 5; // 6행
    if (allRows.length <= HEADER_IDX) throw new Error("헤더 행이 없습니다.");
    const headers = allRows[HEADER_IDX].map(h => String(h).trim());

    // 중복 컬럼 처리: 매입=공급자(첫번째), 매출=공급받는자(마지막)
    const allIdxOf = (name) => headers.reduce((acc, h, i) => h === name ? [...acc, i] : acc, []);
    const pickIdx  = (idxs) => sideType === "매입" ? (idxs[0] ?? -1) : (idxs[idxs.length - 1] ?? -1);
    const vendorColIdx = pickIdx(allIdxOf("상호"));
    const bizColIdx    = pickIdx(allIdxOf("사업자(주민)번호"));

    const hasDiv = headers.includes("구분");
    const filteredRows = [];
    let totalsRow = null;

    for (let i = HEADER_IDX + 1; i < allRows.length; i++) {
      const raw = allRows[i];
      if (raw.every(v => v === "" || v == null)) continue;

      // 표준 row 객체 (중복 헤더는 마지막 값이 됨)
      const row = {};
      headers.forEach((h, j) => {
        let val = raw[j] ?? "";
        if (typeof val === "number") {
          const cell = ws[XLSX.utils.encode_cell({ r: i, c: j })];
          if (cell && cell.w && /^0\d/.test(cell.w)) val = cell.w;
        }
        row[h] = val;
      });

      // 상대방 거래처 컬럼을 올바른 위치(첫번째/마지막)로 덮어씀
      if (vendorColIdx >= 0) row["상호"]              = String(raw[vendorColIdx] || "").trim();
      if (bizColIdx    >= 0) row["사업자(주민)번호"] = String(raw[bizColIdx]    || "").trim();

      const div     = String(row["구분"] || "").trim();
      // 합계행 감지 ①: 구형 포맷
      if (div === "" && /^\d+건/.test(String(row["종류"] || "").trim())) { totalsRow = row; continue; }
      // 합계행 감지 ②: 신형 포맷 (작성일자 없음)
      const dateVal = String(row["작성일자"] || "").trim();
      if (!hasDiv && !dateVal) { totalsRow = row; continue; }
      if (hasDiv && div !== "매출" && div !== "매입") continue;

      if (row["작성일자"]) row["작성일자"] = formatExcelDateToStr(row["작성일자"]);
      if (row["발급일자"]) row["발급일자"] = formatExcelDateToStr(row["발급일자"]);

      row["거래처명"]   = String(row["상호"] || "").trim();
      row["사업자번호"] = normalizeBizNum(String(row["사업자(주민)번호"] || "").trim());
      row["공급가액"]   = Number(row["공급가액"] || 0);
      row["세액"]       = Number(row["세액"]     || 0);
      row["합계"]       = Number(row["합계"] || row["합계금액"] || 0);

      const apprNo = String(row["승인번호"] || row["승인번호(발급)"] || "").trim();
      row["_row_key"] = apprNo || `${row["작성일자"]}_${row["사업자번호"]}_${row["합계"]}`;

      filteredRows.push(row);
    }

    const parsedTotal = filteredRows.reduce((s, r) => s + r["공급가액"], 0);
    let checksumOk = null, fileTotal = null;
    if (totalsRow) {
      fileTotal = Number(totalsRow["공급가액"] || totalsRow["합계금액"] || totalsRow["합계"] || 0);
      checksumOk = Math.abs(parsedTotal - fileTotal) < 1;
    }

    const names = [...new Set(filteredRows.map(r => r["거래처명"]))];
    const allMauto = names.length > 0 && names.every(n => /엠오토|M오토|EM오토/i.test(n));

    return { rows: filteredRows, checksumOk, parsedTotal, fileTotal, allMauto, error: null };
  } catch (err) {
    return { rows: null, checksumOk: null, parsedTotal: null, fileTotal: null, allMauto: false, error: err.message };
  }
}

const DATA_IMPORT_LABELS = {
  taxInvoice: "세금계산서 (매출/매입 통합)",
  ledgerSales: "계정별원장 — 외상매출금",
  ledgerPurchase: "계정별원장 — 외상매입금",
  ledgerPayable: "계정별원장 — 미지급금",
  dailySales: "영업현황 (일별)",
};

const DATA_IMPORT_ACTIONS = {
  taxInvoice: { action: "upsertTaxInvoices" },
  ledgerSales: { action: "upsertLedger", ledgerType: "매출" },
  ledgerPurchase: { action: "upsertLedger", ledgerType: "매입" },
  ledgerPayable: { action: "upsertLedger", ledgerType: "미지급" },
  dailySales: { action: "upsertDailySales" },
};

function renderDataImportPanel() {
  const panel = document.getElementById("dataImportPanel");
  if (!panel) return;
  panel.classList.toggle("hidden", !dataImportState.visible);
  if (!dataImportState.visible) return;

  const sections = Object.keys(DATA_IMPORT_LABELS).map(key => {
    const sec = dataImportState[key];
    const parsed = sec.parsed;
    const matchedCount = parsed ? parsed.filter(r => r._matched_code).length : 0;
    const unmatchedCount = parsed ? parsed.length - matchedCount : 0;

    // 저장된 소스 파일 목록
    const storedFiles = getMiraeSectionFiles(key);
    const storedFilesHtml = storedFiles.length > 0 ? `
      <div class="di-source-files">
        ${storedFiles.map(f => `
          <span class="di-source-file">
            📄 ${escapeHtml(f.filename)} <span class="muted">(${f.rows?.length || 0}건)</span>
            <button type="button" class="di-source-del-btn" data-key="${key}" data-filename="${escapeHtml(f.filename)}" title="삭제">✕</button>
          </span>
        `).join("")}
      </div>
    ` : "";

    return `
      <div class="di-section">
        <div class="di-section-header">
          <span class="di-section-label">${DATA_IMPORT_LABELS[key]}</span>
          <label class="di-file-btn">
            파일 선택
            <input type="file" class="di-file-input" data-key="${key}" accept=".xls,.xlsx" hidden />
          </label>
          ${parsed ? `
            <span class="di-count">
              ${parsed.length}행
              · <span class="di-match-ok">${matchedCount}건 매칭</span>
              ${unmatchedCount > 0 ? `· <span class="di-match-fail">${unmatchedCount}건 미매칭</span>` : ""}
            </span>
            <button type="button" class="di-save-btn" data-key="${key}" ${sec.saving ? "disabled" : ""}>
              ${sec.saving ? sec.status || "저장 중…" : `구글시트 저장 (${parsed.length}건)`}
            </button>
          ` : ""}
        </div>
        ${storedFilesHtml}
        ${sec.status ? `<div class="di-status ${sec.status.startsWith("✓") ? "di-status-ok" : sec.status.startsWith("저장") ? "" : "di-status-err"}">${sec.status}</div>` : ""}
      </div>
    `;
  }).join("");

  panel.innerHTML = `
    <div class="di-header">
      <h3>자료 업로드</h3>
      <p class="di-desc muted">파일을 선택하면 <strong>로컬에 즉시 저장</strong>되고 대사 탭이 재빌드됩니다. 같은 기간 재업로드 시 파일 단위로 교체됩니다.<br>'구글시트 저장'은 클라우드 백업용입니다 (중복 행 자동 덮어쓰기).</p>
      <button type="button" class="di-close-btn" id="dataImportCloseBtn">✕ 닫기</button>
    </div>
    <div class="di-sections">${sections}</div>
  `;

  panel.querySelector("#dataImportCloseBtn").addEventListener("click", () => {
    dataImportState.visible = false;
    panel.classList.add("hidden");
  });

  panel.querySelectorAll(".di-file-input").forEach(input => {
    input.addEventListener("change", async e => {
      const file = e.target.files[0];
      if (!file) return;
      const key = input.dataset.key;

      // 같은 파일명이 이미 저장된 경우 교체 확인
      const existing = getMiraeSectionFiles(key).find(f => f.filename === file.name);
      if (existing) {
        if (!confirm(`"${file.name}" 파일이 이미 저장되어 있습니다.\n같은 기간 데이터를 통째 교체하겠습니까?`)) {
          input.value = "";
          return;
        }
      }

      const sec = dataImportState[key];
      sec.status = "파싱 중…";
      sec.parsed = null;
      renderDataImportPanel();

      let result;
      if (key === "taxInvoice") result = await parseTaxInvoiceFile(file);
      else if (key === "dailySales") result = await parseDailySalesFile(file);
      else result = await parseLedgerFile(file);

      if (result.error) {
        sec.status = `오류: ${result.error}`;
      } else {
        sec.parsed = result.rows;
        sec.status = "";
        // 소스 파일 저장 + 재빌드
        saveMiraeSectionFile(key, file.name, result.rows);
        rebuildDaesaFromSources();
      }
      renderDataImportPanel();
    });
  });

  // 소스 파일 삭제 버튼
  panel.querySelectorAll(".di-source-del-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      const { key, filename } = btn.dataset;
      if (!confirm(`"${filename}" 파일을 로컬 저장소에서 삭제하겠습니까?`)) return;
      deleteMiraeSectionFile(key, filename);
      rebuildDaesaFromSources();
      renderDataImportPanel();
    });
  });

  panel.querySelectorAll(".di-save-btn").forEach(btn => {
    btn.addEventListener("click", async () => {
      const key = btn.dataset.key;
      const sec = dataImportState[key];
      if (!sec.parsed?.length) return;
      const rowsToSave = sec.parsed;
      if (!rowsToSave.length) return;
      sec.saving = true;
      renderDataImportPanel();
      try {
        const { action, ...extra } = DATA_IMPORT_ACTIONS[key];
        const BATCH = 1000;
        const PARALLEL = 3;
        const total = rowsToSave.length;
        const batches = [];
        for (let i = 0; i < total; i += BATCH) batches.push(rowsToSave.slice(i, i + BATCH));
        let saved = 0;
        for (let i = 0; i < batches.length; i += PARALLEL) {
          const group = batches.slice(i, i + PARALLEL);
          sec.status = `저장 중… ${saved} / ${total}건`;
          renderDataImportPanel();
          await Promise.all(group.map(b => postSheetWebApp(action, { rows: b, ...extra })));
          saved += group.reduce((sum, b) => sum + b.length, 0);
        }
        const matchedCount = rowsToSave.filter(r => r._matched_code).length;
        sec.status = `✓ ${total}건 저장 완료 (매칭 ${matchedCount}건 / 미매칭 ${total - matchedCount}건)`;
      } catch (err) {
        sec.status = `저장 실패: ${err.message}`;
      } finally {
        sec.saving = false;
        renderDataImportPanel();
      }
    });
  });
}

function setupDataImport() {
  const btn = document.getElementById("dataImportButton");
  const panel = document.getElementById("dataImportPanel");
  if (!btn || !panel) return;
  btn.addEventListener("click", () => {
    dataImportState.visible = !dataImportState.visible;
    panel.classList.toggle("hidden", !dataImportState.visible);
    if (dataImportState.visible) renderDataImportPanel();
  });
}

function setupApiTokenButton() {
  const btn = document.getElementById("apiTokenButton");
  if (!btn) return;
  btn.onclick = () => promptApiToken();
  const stored = getApiToken();
  if (stored) btn.title = "API 토큰 설정됨 (클릭하여 변경)";
}

async function init() {
  loadGroupOrder();
  loadVendorMemos();
  loadMautoDataLocal();
  loadPnlLocal();

  renderPartnerFilter();
  renderFilterControls();
  renderVendorMasterPanel();
  setupTabs();
  updateFilterBarVisibility(document.querySelector(".tab-button.active")?.dataset.tab || "home");
  setupMasterMenu();
  setupVendorMasterImport();
  setupLedgerVendorImport();
  setupBankImport();
  setupDataImport();
  setupRulesPanel();
  setupApiTokenButton();

  const sheetLinkBtn = document.getElementById("sheetLinkButton");
  if (sheetLinkBtn && typeof SHEET_SPREADSHEET_ID !== "undefined" && SHEET_SPREADSHEET_ID) {
    sheetLinkBtn.href = `https://docs.google.com/spreadsheets/d/${SHEET_SPREADSHEET_ID}`;
  }

  // 검색창 → 대사 탭 연동
  elements.searchInput?.addEventListener("input", () => {
    const daesaEl = document.getElementById("daesa");
    if (daesaEl && !daesaEl.classList.contains("hidden") && daesaState.loaded) {
      renderDaesaTab();
    }
  });

  // 초기 로딩 시 홈 탭 활성화
  loadClassifiedRows();
  loadSourceFiles();
  loadUserEdits();
  migrateLegacyIfNeeded(); // 기존 분류 데이터를 불변/사용자 영역으로 분리
  // 미래 자료업로드 소스 파일 로드 → 소스 있으면 daesaState 즉시 재빌드
  miraeTaxSources    = loadMiraeSource(MIRAE_SOURCE_TAX_KEY);
  miraeLedgerSources = loadMiraeSource(MIRAE_SOURCE_LEDGER_KEY);
  miraeBizSources    = loadMiraeSource(MIRAE_SOURCE_BIZ_KEY);
  if (hasMiraeSources()) rebuildDaesaFromSources();
  // 엠오토 세금계산서 소스 로드 + 제외 거래처 로드
  loadMautoTaxSource();
  loadFixedChecked();
  loadFixedAmountOverrides();
  loadMautoExcludeVendors();
  if (Object.keys(mautoTaxSources).length) rebuildMautoTaxInvoices();
  switchTab("home");

  await Promise.all([
    loadSheetPayables(),
    loadSheetReceivables(),
    loadSheetFixedExpenses(),
    loadAvailableFunds()
  ]);

  rerenderAll(); // 모든 데이터 로드 후 최종 갱신 보장
}

// ══════════════════════════════════════════════════════════════
// 경영손익 (P&L) 탭
// ══════════════════════════════════════════════════════════════

const PNL_LOCAL_KEY   = "cashflow-app.pnl-v1";
const PNL_SHEET_NAME  = "월간손익";
const PNL_STATUS_ORDER = ["draft", "기안", "합의1", "합의2", "결재완료"];
const PNL_META = {
  companyName: "미래오토메이션(주)",
  department:  "관리부",
  author:    { name: "여희정", title: "차장" },
  approvers: [
    { name: "김도연", title: "부장", dept: "영업부" },
    { name: "오성철", title: "부장", dept: "시스템사업부" },
  ],
  ceo: { name: "장운기", title: "대표이사" },
};
const PNL_APPROVAL_STEPS = [
  { dateKey: "draftDate",  nextStatus: "기안",    role: "기안",    name: PNL_META.author.name,      title: PNL_META.author.title,    stampCls: "pnl-stamp-signed"   },
  { dateKey: "agree1Date", nextStatus: "합의1",   role: "합의",    name: PNL_META.approvers[0].name, title: PNL_META.approvers[0].title, stampCls: "pnl-stamp-signed" },
  { dateKey: "agree2Date", nextStatus: "합의2",   role: "합의",    name: PNL_META.approvers[1].name, title: PNL_META.approvers[1].title, stampCls: "pnl-stamp-signed" },
  { dateKey: "ceoDate",    nextStatus: "결재완료", role: "최종결재", name: PNL_META.ceo.name,         title: PNL_META.ceo.title,       stampCls: "pnl-stamp-approved" },
];
const PNL_SEED = [
  { year:2026, month:1,  revenue:502848410, targetRevenue:600000000, cogs:422295376, mfg:18154934, sga:66248997, interest:9192216,
    approvalStatus:"결재완료", draftDate:"2026. 02. 24", agree1Date:"2026. 02. 25", agree2Date:"2026. 02. 26", ceoDate:"2026. 03. 03",
    docNo:"2026-03-000092", ceoComment:"전체적 계획 후 정리 합시다." },
  { year:2026, month:2,  revenue:283944347, targetRevenue:400000000, cogs:254640693, mfg:18471639, sga:97841624, interest:2215544,
    approvalStatus:"결재완료", draftDate:"2026. 03. 17", agree1Date:"2026. 03. 20", agree2Date:"2026. 03. 18", ceoDate:"2026. 04. 07",
    docNo:"2026-05-000231", ceoComment:"3월에는 흑자로 갑시다" },
];

let pnlData     = [];
let pnlSubTab   = "input";
let pnlInputYear  = new Date().getFullYear();
let pnlInputMonth = new Date().getMonth() + 1;

// ── 경영손익 결재 Firebase 알림 ────────────────────────────────
const PNL_STEP_EMAILS = [
  "yhj@mauto.co.kr",  // 기안 (여희정)
  "kdy@mauto.co.kr",  // 합의1 (김도연)
  "osc@mauto.co.kr",  // 합의2 (오성철)
  "jug@mauto.co.kr",  // 최종결재 (장운기)
];

let _pnlFireDb = null;
function _initPnlFirebase() {
  if (_pnlFireDb) return _pnlFireDb;
  try {
    const cfg = {
      apiKey: "AIzaSyC7kPGuahfBGk5Z-A7tNMVFzm14HCV_R-8",
      authDomain: "staff-directory-app-9e17b.firebaseapp.com",
      databaseURL: "https://staff-directory-app-9e17b-default-rtdb.asia-southeast1.firebasedatabase.app",
      projectId: "staff-directory-app-9e17b",
    };
    const existing = (typeof firebase !== "undefined" && firebase.apps || []).find(a => a.name === "pnl-notif");
    const app = existing || firebase.initializeApp(cfg, "pnl-notif");
    _pnlFireDb = app.database();
  } catch(e) { console.warn("[pnl-notif] Firebase init 실패:", e); }
  return _pnlFireDb;
}

function _pnlEk(email) { return email.toLowerCase().replace(/\./g, ",").replace(/@/g, "|"); }

function writePnlPendingToFirebase() {
  const db = _initPnlFirebase();
  if (!db) return;
  const nonDraft = pnlData.filter(e => e.approvalStatus && e.approvalStatus !== "draft");
  const qRows    = Object.values(pnlQuarterApproval || {}).filter(q => q && q.approvalStatus && q.approvalStatus !== "draft");

  PNL_STEP_EMAILS.forEach((email, idx) => {
    let total = 0;
    if (idx === 0) {
      // 기안자: 결재 진행 중 건수 (본인이 기안했지만 아직 결재완료 아닌 것)
      const si = s => PNL_STATUS_ORDER.indexOf(s || "draft");
      total = nonDraft.filter(e => { const s = si(e.approvalStatus); return s >= 1 && s < 4; }).length
            + qRows.filter(e   => { const s = si(e.approvalStatus); return s >= 1 && s < 4; }).length;
    } else {
      const step = PNL_APPROVAL_STEPS[idx];
      const check = e => {
        const s = PNL_STATUS_ORDER.indexOf(e.approvalStatus || "draft");
        if (idx === 1 || idx === 2) return s >= 1 && !e[step.dateKey];
        if (idx === 3) return !!e.agree1Date && !!e.agree2Date && !e.ceoDate;
        return false;
      };
      total = nonDraft.filter(check).length + qRows.filter(check).length;
    }
    db.ref("pnlPending/" + _pnlEk(email)).set(total > 0 ? total : null).catch(() => {});
  });
}
let pnlRptYear    = new Date().getFullYear();
let pnlRptMonth   = new Date().getMonth() + 1;
let pnlRptMode    = "monthly";   // "monthly" | "quarterly" | "halfyear" | "annual"
let pnlRptQuarter = Math.ceil((new Date().getMonth() + 1) / 3);
let pnlRptHalf    = new Date().getMonth() < 6 ? 1 : 2; // 반기 모드: 1=상반기, 2=하반기
let pnlDashYear = new Date().getFullYear();
let pnlDashPeriod = "monthly";
let pnlInvYear  = new Date().getFullYear();
let _pnlCharts  = {};
const _pnlQtrSyncedKeys = new Set(); // 세션 중 구글시트 동기화 완료한 분기 키
let _pnlImportIncome = null;  // 손익계산서 파싱 결과 {month: {revenue,cogs,sga,interest}}
let _pnlImportCost   = null;  // 원가명세서 파싱 결과 {month: {mfg}}
let _pnlImportYear   = new Date().getFullYear();
const PNL_Q_APPROVAL_KEY = "cashflow-app.pnl-quarter-approval-v1";
let pnlQuarterApproval = {};  // { "2026_Q1": { approvalStatus, draftDate, ... } }

// ── 로컬 스토리지 ─────────────────────────────────────────────
function loadPnlLocal() {
  try {
    const raw = localStorage.getItem(PNL_LOCAL_KEY);
    pnlData = raw ? JSON.parse(raw) : [];
  } catch (_) { pnlData = []; }
  if (!pnlData.length) {
    pnlData = PNL_SEED.map(d => ({ ...d }));
    savePnlLocal();
  }
  loadPnlQuarterApprovalLocal();
}

function savePnlLocal() {
  try { localStorage.setItem(PNL_LOCAL_KEY, JSON.stringify(pnlData)); } catch (_) {}
}
function loadPnlQuarterApprovalLocal() {
  try {
    const raw = localStorage.getItem(PNL_Q_APPROVAL_KEY);
    pnlQuarterApproval = raw ? JSON.parse(raw) : {};
  } catch (_) { pnlQuarterApproval = {}; }
}
function savePnlQuarterApprovalLocal() {
  try { localStorage.setItem(PNL_Q_APPROVAL_KEY, JSON.stringify(pnlQuarterApproval)); } catch (_) {}
}

function _saveQtrToSheets(year, quarter, qKey, qApproval) {
  const months = [1, 2, 3].map(m => m + (quarter - 1) * 3);
  const fins = (pnlData || []).filter(d => d.year === year && months.includes(d.month))
    .reduce((a, d) => ({
      revenue:       a.revenue       + (d.revenue       || 0),
      targetRevenue: a.targetRevenue + (d.targetRevenue || 0),
      cogs:          a.cogs          + (d.cogs          || 0),
      mfg:           a.mfg           + (d.mfg           || 0),
      sga:           a.sga           + (d.sga           || 0),
      interest:      a.interest      + (d.interest      || 0),
    }), { revenue: 0, targetRevenue: 0, cogs: 0, mfg: 0, sga: 0, interest: 0 });
  postSheetWebApp("savePnlData", { rows: [{
    _key: qKey,
    year, quarter, month: 0,
    ...fins,
    approvalStatus: qApproval.approvalStatus,
    draftDate:      qApproval.draftDate  || "",
    agree1Date:     qApproval.agree1Date || "",
    agree2Date:     qApproval.agree2Date || "",
    ceoDate:        qApproval.ceoDate    || "",
    docNo:          qApproval.docNo      || "",
  }]}).catch(e => console.warn("[PNL-Q] Sheets 저장 실패:", e));
}

function getPnlEntry(year, month) {
  return pnlData.find(d => d.year === year && d.month === month) ?? null;
}

function upsertPnlEntry(entry) {
  const idx = pnlData.findIndex(d => d.year === entry.year && d.month === entry.month);
  if (idx >= 0) pnlData[idx] = { ...pnlData[idx], ...entry };
  else pnlData.push({ ...entry });
  pnlData.sort((a, b) => a.year !== b.year ? a.year - b.year : a.month - b.month);
  savePnlLocal();
  _schedulePnlSave(entry);
}

function deletePnlEntry(year, month) {
  pnlData = pnlData.filter(d => !(d.year === year && d.month === month));
  savePnlLocal();
}

let _pnlSaveTimer = null;
function _schedulePnlSave(entry) {
  clearTimeout(_pnlSaveTimer);
  _pnlSaveTimer = setTimeout(async () => {
    if (!SHEET_APP_SCRIPT_URL) return;
    try {
      await postSheetWebApp("savePnlData", {
        row: {
          ...entry,
          corrections: JSON.stringify(entry.corrections || []),
          _key: `${entry.year}_${String(entry.month).padStart(2, "0")}`,
        },
      });
    } catch (e) { console.warn("[손익] 구글시트 저장 실패:", e); }
  }, 800);
}

async function loadPnlRemote() {
  if (!SHEET_APP_SCRIPT_URL) return;
  try {
    const url = new URL(SHEET_APP_SCRIPT_URL);
    const token = getApiToken();
    if (token) url.searchParams.set("token", token);
    url.searchParams.set("action", "getPnlData");
    const resp = await fetch(url.toString());
    if (!resp.ok) return;
    const body = await resp.json();
    const rows = Array.isArray(body?.rows) ? body.rows : [];
    if (!rows.length) return;
    rows.forEach(r => {
      const y = Number(r.year), m = Number(r.month);
      if (!y || !m) return;
      const entry = {
        year: y, month: m,
        revenue: Number(r.revenue || 0), targetRevenue: Number(r.targetRevenue || 0),
        cogs: Number(r.cogs || 0), mfg: Number(r.mfg || 0),
        sga: Number(r.sga || 0), interest: Number(r.interest || 0),
        purchaseAmount: Number(r.purchaseAmount || 0),
        approvalStatus: r.approvalStatus || "draft",
        draftDate: r.draftDate || "", agree1Date: r.agree1Date || "",
        agree2Date: r.agree2Date || "", ceoDate: r.ceoDate || "",
        docNo: r.docNo || "", ceoComment: r.ceoComment || "",
        corrections: (() => { try { return JSON.parse(r.corrections || "[]"); } catch (_) { return []; } })(),
        ...(r.beginInventory != null && r.beginInventory !== "" ? { beginInventory: Number(r.beginInventory) } : {}),
        ...(r.endInventory   != null && r.endInventory   !== "" ? { endInventory:   Number(r.endInventory)   } : {}),
      };
      const idx = pnlData.findIndex(d => d.year === y && d.month === m);
      if (idx >= 0) pnlData[idx] = entry; else pnlData.push(entry);
    });
    pnlData.sort((a, b) => a.year !== b.year ? a.year - b.year : a.month - b.month);
    savePnlLocal();
    renderPnlTab();

    // localStorage에만 있는 분기 서명 데이터를 구글시트에 동기화 (세션당 1회)
    Object.entries(pnlQuarterApproval).forEach(([qKey, qApproval]) => {
      if (qApproval.approvalStatus === "draft") return;
      if (_pnlQtrSyncedKeys.has(qKey)) return;
      const m = qKey.match(/^(\d{4})_Q([1-4])$/);
      if (!m) return;
      _pnlQtrSyncedKeys.add(qKey);
      _saveQtrToSheets(+m[1], +m[2], qKey, qApproval);
    });
  } catch (e) { console.warn("[손익] 원격 로드 실패:", e); }
}

// ── 계산 유틸 ─────────────────────────────────────────────────
function calcPnl(d) {
  // 기초/기말 재고 수동 입력 시: 매출원가 = 기초재고 + 당기상품매입 − 기말재고
  const hasManualInv = d.beginInventory !== undefined && d.endInventory !== undefined;
  const effectiveCogs = hasManualInv
    ? (d.beginInventory || 0) + (d.purchaseAmount || 0) - (d.endInventory || 0)
    : (d.cogs || 0);
  const gross = (d.revenue || 0) - (effectiveCogs + (d.mfg || 0));
  const op    = gross - (d.sga || 0);
  const mgmt  = op - (d.interest || 0);
  const gmRate  = d.revenue > 0 ? gross / d.revenue * 100 : 0;
  const opRate  = d.revenue > 0 ? op    / d.revenue * 100 : 0;
  const targetAchieve = d.targetRevenue > 0 ? d.revenue / d.targetRevenue * 100 : null;
  return { gross, op, mgmt, gmRate, opRate, targetAchieve, cogs: effectiveCogs };
}
const _pf  = n => Math.abs(n).toLocaleString("ko-KR");
const _ps  = n => n >= 0 ? _pf(n) : `△${_pf(n)}`;
const _pc  = n => n >= 0 ? "pnl-pos" : "pnl-neg";
function _todayKor() {
  const d = new Date();
  return `${d.getFullYear()}. ${String(d.getMonth()+1).padStart(2,"0")}. ${String(d.getDate()).padStart(2,"0")}`;
}
function _pnlStatusIdx(status) { return PNL_STATUS_ORDER.indexOf(status || "draft"); }

function pnlToast(msg) {
  let el = document.getElementById("pnl-toast");
  if (!el) {
    el = document.createElement("div");
    el.id = "pnl-toast";
    el.className = "pnl-toast";
    document.body.appendChild(el);
  }
  el.textContent = msg;
  el.classList.add("pnl-toast-show");
  setTimeout(() => el.classList.remove("pnl-toast-show"), 2200);
}

// ── 메인 렌더러 ───────────────────────────────────────────────
function renderPnlTab() {
  const sec = document.getElementById("pnl");
  if (!sec || !sec.classList.contains("active")) return;

  if (!document.getElementById("pnl-fonts")) {
    const lk = document.createElement("link");
    lk.id   = "pnl-fonts"; lk.rel = "stylesheet";
    lk.href = "https://fonts.googleapis.com/css2?family=Noto+Serif+KR:wght@400;600;700&family=DM+Mono:wght@400;500&display=swap";
    document.head.appendChild(lk);
  }

  sec.innerHTML = `
    <div class="pnl-container">
      <div class="pnl-sub-tabs">
        <button class="pnl-sub-btn${pnlSubTab==="input"?" active":""}" data-pnl-tab="input">입력</button>
        <button class="pnl-sub-btn${pnlSubTab==="report"?" active":""}" data-pnl-tab="report">보고서</button>
        <button class="pnl-sub-btn${pnlSubTab==="dashboard"?" active":""}" data-pnl-tab="dashboard">대시보드</button>
        <button class="pnl-sub-btn${pnlSubTab==="inventory"?" active":""}" data-pnl-tab="inventory">재고</button>
        <button class="pnl-sub-btn pnl-sync-btn" id="pnlSyncAll" title="로컬 데이터를 구글시트로 일괄 전송">동기화</button>
      </div>
      <div id="pnl-sub-content"></div>
    </div>`;

  sec.querySelectorAll(".pnl-sub-btn[data-pnl-tab]").forEach(btn => {
    btn.addEventListener("click", () => {
      pnlSubTab = btn.dataset.pnlTab;
      renderPnlTab();
    });
  });
  document.getElementById("pnlSyncAll")?.addEventListener("click", async () => {
    if (!pnlData.length) { pnlToast("동기화할 데이터가 없습니다"); return; }
    const btn = document.getElementById("pnlSyncAll");
    btn.textContent = "⏳ 전송중..."; btn.disabled = true;
    try {
      const rows = pnlData.map(e => ({ ...e, _key: `${e.year}_${String(e.month).padStart(2,"0")}` }));
      await postSheetWebApp("savePnlData", { rows });
      pnlToast(`구글시트 동기화 완료 (${rows.length}건)`);
    } catch(e) {
      pnlToast("동기화 실패: " + e.message);
    } finally {
      btn.textContent = "☁️ 동기화"; btn.disabled = false;
    }
  });

  const content = document.getElementById("pnl-sub-content");
  if (pnlSubTab === "input")          renderPnlInput(content);
  else if (pnlSubTab === "report")    renderPnlReport(content);
  else if (pnlSubTab === "inventory") renderPnlInventory(content);
  else                                renderPnlDashboard(content);
}

// ── 재고 관리 탭 ─────────────────────────────────────────────
function renderPnlInventory(el) {
  const curY = new Date().getFullYear();
  const yearOpts = Array.from({length: curY - 2023}, (_, i) => 2024 + i)
    .map(y => `<option value="${y}" ${y === pnlInvYear ? "selected" : ""}>${y}년</option>`).join("");

  const inc = _pnlImportIncome || {};

  function buildRows() {
    return Array.from({length: 12}, (_, i) => {
      const m = i + 1;
      const entry = getPnlEntry(pnlInvYear, m) || {};
      const purchase = (inc[m] && inc[m].purchaseAmount) || entry.purchaseAmount || 0;
      const begin = entry.beginInventory;
      const end   = entry.endInventory;
      const hasManual = begin !== undefined && end !== undefined;
      const calcCogs  = hasManual
        ? (begin || 0) + purchase - (end || 0)
        : (entry.cogs || 0);
      return { m, begin, end, purchase, calcCogs, hasManual, hasPurchase: !!purchase };
    });
  }

  function render() {
    const rows = buildRows();
    el.innerHTML = `
      <div class="pnl-inv-wrap">
        <div class="pnl-inv-top">
          <span class="pnl-inv-title">기초 / 기말 재고 관리</span>
          <select id="pnlInvYearSel">${yearOpts}</select>
        </div>
        <div class="pnl-inv-hint">
          기초·기말 재고를 입력하면 <b>상품매출원가 = 기초재고 + 당기상품매입 − 기말재고</b>로 자동 계산됩니다.<br/>
          당기상품매입은 손익계산서 Excel 업로드 후 자동 추출됩니다. 비어있으면 직접 입력하세요.
        </div>
        <div class="pnl-inv-table-wrap">
          <table class="pnl-inv-table">
            <thead><tr>
              <th>월</th>
              <th>기초재고</th>
              <th>당기상품매입</th>
              <th>기말재고</th>
              <th>계산 매출원가</th>
              <th>저장</th>
            </tr></thead>
            <tbody>
              ${rows.map(r => `
                <tr data-month="${r.m}">
                  <td class="pnl-inv-m">${r.m}월</td>
                  <td><input type="text" class="pnl-inv-inp" data-f="begin" value="${r.begin !== undefined ? _pf(r.begin) : ""}" placeholder="0" inputmode="numeric" /></td>
                  <td><input type="text" class="pnl-inv-inp pnl-inv-purchase-inp" data-f="purchase" value="${r.purchase ? _pf(r.purchase) : ""}" placeholder="0" inputmode="numeric" /></td>
                  <td><input type="text" class="pnl-inv-inp" data-f="end" value="${r.end !== undefined ? _pf(r.end) : ""}" placeholder="0" inputmode="numeric" /></td>
                  <td class="pnl-inv-calc${r.hasManual ? " pnl-inv-active" : ""}">${r.hasManual ? _pf(r.calcCogs) : "—"}</td>
                  <td><button class="pnl-inv-save-btn" data-month="${r.m}">저장</button></td>
                </tr>`).join("")}
            </tbody>
          </table>
        </div>
        <div class="pnl-inv-footer">
          <button class="pnl-btn pnl-btn-ghost" id="pnlInvClearYear">전체 초기화</button>
          <button class="pnl-btn pnl-btn-primary" id="pnlInvSaveAll">전체 저장</button>
        </div>
      </div>`;

    el.querySelector("#pnlInvYearSel").addEventListener("change", e => {
      pnlInvYear = parseInt(e.target.value);
      render();
    });

    function parseN(s) { return Number(String(s).replace(/[^0-9]/g, "")) || 0; }

    function refreshCalcCell(tr) {
      const beginVal   = tr.querySelector("[data-f=begin]").value.trim();
      const endVal     = tr.querySelector("[data-f=end]").value.trim();
      const purchaseVal = tr.querySelector("[data-f=purchase]").value.trim();
      const calcTd     = tr.querySelector(".pnl-inv-calc");
      if (beginVal !== "" || endVal !== "") {
        const calc = parseN(beginVal) + parseN(purchaseVal) - parseN(endVal);
        calcTd.textContent = _pf(calc);
        calcTd.classList.add("pnl-inv-active");
      } else {
        calcTd.textContent = "—";
        calcTd.classList.remove("pnl-inv-active");
      }
    }

    el.querySelectorAll(".pnl-inv-inp").forEach(inp => {
      inp.addEventListener("input", () => {
        const raw = parseN(inp.value);
        inp.value = raw ? _pf(raw) : "";
        refreshCalcCell(inp.closest("tr"));
      });
    });

    function saveRow(tr) {
      const m = parseInt(tr.dataset.month);
      const beginVal   = tr.querySelector("[data-f=begin]").value.trim();
      const endVal     = tr.querySelector("[data-f=end]").value.trim();
      const purchaseVal = parseN(tr.querySelector("[data-f=purchase]").value);
      const existing   = getPnlEntry(pnlInvYear, m) || {};
      const update     = { ...existing, year: pnlInvYear, month: m, purchaseAmount: purchaseVal };

      if (beginVal !== "" || endVal !== "") {
        update.beginInventory = parseN(beginVal);
        update.endInventory   = parseN(endVal);
        update.cogs = update.beginInventory + purchaseVal - update.endInventory;
      } else {
        delete update.beginInventory;
        delete update.endInventory;
      }
      upsertPnlEntry(update);
    }

    el.querySelectorAll(".pnl-inv-save-btn").forEach(btn => {
      btn.addEventListener("click", () => {
        const m  = parseInt(btn.dataset.month);
        const tr = el.querySelector(`tr[data-month="${m}"]`);
        saveRow(tr);
        btn.textContent = "✓";
        setTimeout(() => { btn.textContent = "저장"; }, 1500);
      });
    });

    el.querySelector("#pnlInvSaveAll").addEventListener("click", () => {
      el.querySelectorAll("tbody tr").forEach(tr => saveRow(tr));
      pnlToast("재고 전체 저장 완료");
      render();
    });

    el.querySelector("#pnlInvClearYear").addEventListener("click", () => {
      if (!confirm(`${pnlInvYear}년 재고 입력값을 모두 초기화하겠습니까?`)) return;
      Array.from({length: 12}, (_, i) => i + 1).forEach(m => {
        const existing = getPnlEntry(pnlInvYear, m);
        if (!existing) return;
        const { beginInventory, endInventory, ...rest } = existing;
        upsertPnlEntry(rest);
      });
      pnlToast("초기화 완료");
      render();
    });
  }

  render();
}

// ── Excel 일괄입력 파서 ────────────────────────────────────────

// 셀 하나에서 월(1~12) 추출. "1월"/"01월"/Date 객체/숫자 1~12 모두 처리
function _cellToMonth(c) {
  if (c instanceof Date && !isNaN(c)) return c.getMonth() + 1;
  if (typeof c === "number" && Number.isFinite(c) && c >= 1 && c <= 12 && c === Math.floor(c)) return c;
  // ERP XLS 파일은 문자열 끝에 \x00(null byte)를 붙이는 경우가 있어 제거 후 매칭
  const s = String(c).replace(/[\x00\s]/g, "");
  const m = s.match(/^0?(\d{1,2})월$/);
  if (m) { const n = parseInt(m[1]); if (n >= 1 && n <= 12) return n; }
  return null;
}

// ERP 행 레이블 정규화: null byte·공백·점 제거, 전각 → 반각 변환, 숫자 접미사 제거
function _normLabel(s) {
  return String(s ?? "").replace(/\x00/g, "").trim()
    .replace(/[Ⅰ]/g, "I").replace(/[Ⅱ]/g, "II").replace(/[Ⅲ]/g, "III")
    .replace(/[Ⅳ]/g, "IV").replace(/[Ⅴ]/g, "V").replace(/[Ⅵ]/g, "VI")
    .replace(/[Ⅶ]/g, "VII").replace(/[Ⅷ]/g, "VIII").replace(/[Ⅸ]/g, "IX")
    .replace(/[·\s.]/g, "")  // 공백·점 제거
    .replace(/\d+$/, "");     // 후행 숫자(ERP 코드) 제거
}

function _parsePnlMonthSheet(wb, rowFinders) {
  const sheetName = wb.SheetNames.find(n =>
    n.includes("손익") || n.includes("원가") || n.includes("계산서") || n.includes("명세서")
  ) || wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  const wsRange = ws["!ref"] ? XLSX.utils.decode_range(ws["!ref"]) : null;

  // 셀의 표시 텍스트(w)와 원시값(v) 모두에서 월 번호를 추출하는 헬퍼
  function _cellMonthFromWs(ri, ci) {
    if (wsRange) {
      const addr = XLSX.utils.encode_cell({ r: ri, c: ci });
      const cell = ws[addr];
      if (!cell) return null;
      // cell.w: 표시 텍스트 (예: "1월", "01월") — 날짜 시리얼이라도 이 값은 정상
      return _cellToMonth(cell.w) || _cellToMonth(cell.v);
    }
    return _cellToMonth(raw[ri] && raw[ri][ci]);
  }

  let detectedYear = null;
  let headerRowIdx = -1;
  let monthCols = {};
  let labelCol  = 0;

  for (let ri = 0; ri < Math.min(30, raw.length); ri++) {
    const row = raw[ri];
    const rowText = row.map(c => String(c)).join(" ");
    const ym = rowText.match(/(\d{4})년/);
    if (ym && !detectedYear) detectedYear = parseInt(ym[1]);

    const maxCol = wsRange ? wsRange.e.c : row.length - 1;
    const tmpMonthCols = {};
    for (let ci = 0; ci <= maxCol; ci++) {
      const m = _cellMonthFromWs(ri, ci);
      if (m) tmpMonthCols[m] = ci;
    }
    if (Object.keys(tmpMonthCols).length >= 3) {
      headerRowIdx = ri;
      monthCols = tmpMonthCols;
      for (let ci = 0; ci <= maxCol; ci++) {
        const addr = wsRange ? XLSX.utils.encode_cell({ r: ri, c: ci }) : null;
        const cellTxt = addr
          ? String((ws[addr] && (ws[addr].w || ws[addr].v)) || "")
          : String(row[ci] || "");
        if (cellTxt.replace(/[\s\x00]/g, "") === "과목") { labelCol = ci; break; }
      }
      break;
    }
  }

  // 진단 로그 (F12 Console에서 확인)
  console.log("[PNL] sheet:", sheetName, "headerRow:", headerRowIdx, "monthCols:", monthCols, "labelCol:", labelCol, "year:", detectedYear);
  if (raw.length > 5 && wsRange) {
    const r5 = [];
    for (let ci = 0; ci <= Math.min(15, wsRange.e.c); ci++) {
      const addr = XLSX.utils.encode_cell({ r: 5, c: ci });
      const cell = ws[addr];
      r5.push(cell ? `w=${cell.w} v=${cell.v} t=${cell.t}` : "-");
    }
    console.log("[PNL] row[5]:", r5);
  }
  if (headerRowIdx < 0) return { data: null, year: detectedYear };

  const result = {};
  for (let m = 1; m <= 12; m++) result[m] = {};

  const found = new Set();
  const finderEntries = Object.entries(rowFinders);

  for (let ri = headerRowIdx + 1; ri < raw.length; ri++) {
    const row = raw[ri];
    const rawLabel = String(row[labelCol] ?? "").replace(/\x00/g, "").trim().replace(/\s+/g, " ");
    if (!rawLabel) continue;
    const normLabel = _normLabel(rawLabel);
    for (const [field, matcher] of finderEntries) {
      if (found.has(field)) continue;
      if (matcher(rawLabel, normLabel)) {
        found.add(field);
        for (let m = 1; m <= 12; m++) {
          const ci = monthCols[m];
          if (ci == null) continue;
          const val = row[ci];
          result[m][field] = (typeof val === "number") ? val
            : parseFloat(String(val ?? "").replace(/[^0-9.-]/g, "")) || 0;
        }
        break;
      }
    }
    if (found.size === finderEntries.length) break;
  }
  console.log("[PNL] found:", [...found], "missing:", finderEntries.map(([k])=>k).filter(k=>!found.has(k)));
  return { data: result, year: detectedYear };
}

// matcher(rawLabel, normLabel) 형태로 호출됨
function parsePnlIncomeStatement(wb) {
  return _parsePnlMonthSheet(wb, {
    // 최상위 매출액 행: "I매출액", "1매출액", "매출액...", 기타 접두어 포함 형식 모두 허용
    // 하위 항목(상품매출액·제품매출액)은 "상품"/"제품"으로 시작하므로 자동 제외됨
    revenue: (raw, norm) =>
      norm.startsWith("I매출액") || norm.startsWith("1매출액") ||
      norm.startsWith("매출액") ||
      (norm.includes("매출액") && !norm.startsWith("상품") && !norm.startsWith("제품")),
    // "II.매출원가" / "Ⅱ.매출원가" / "매출원가(상품·제품)" 등
    cogs: (raw, norm) =>
      norm.startsWith("II매출원가") || norm.startsWith("매출원가"),
    // "IV.판매비와관리비" 등
    sga: (raw, norm) =>
      norm.startsWith("IV") && (norm.includes("판매비") || norm.includes("판관비")),
    // "V.영업외비용" 합계 행 — 숫자/Ⅴ 접두어 포함 형식 모두 허용
    // fallback: 영업외비용 합계 행이 없으면 이자비용 행 사용
    interest: (raw, norm) =>
      norm.startsWith("V영업외비용") || norm.startsWith("5영업외비용") ||
      norm === "영업외비용" || norm.startsWith("영업외비용합계") ||
      (norm === "이자비용" && !raw.includes("수익")),
    // "당기상품매입액" — 매출원가 계산의 재고 공식용
    purchaseAmount: (raw, norm) =>
      norm === "당기상품매입액" || norm.endsWith("당기상품매입액"),
  });
}

function parsePnlCostStatement(wb) {
  return _parsePnlMonthSheet(wb, {
    // "11.당기제품제조원가"
    mfg: (raw, norm) =>
      norm.startsWith("11당기제품제조원가") || norm === "당기제품제조원가",
  });
}

function openPnlImportDialog() {
  const inc  = _pnlImportIncome || {};
  const cst  = _pnlImportCost   || {};
  const yr   = _pnlImportYear;
  const curY = new Date().getFullYear();
  const curM = new Date().getMonth() + 1;
  const yearOpts = Array.from({length: curY - 2023}, (_, i) => 2024 + i)
    .map(y => `<option value="${y}" ${y === yr ? "selected" : ""}>${y}년</option>`).join("");

  const months = Array.from({length: 12}, (_, i) => i + 1);
  const rows = months.map(m => {
    const id       = getPnlEntry(yr, m);
    const approved = id?.approvalStatus === "결재완료";
    // 엑셀 값 우선, 없으면 저장된 값 fallback (결재완료 월 수정 대응)
    const rev  = (inc[m] && inc[m].revenue)  || (id && id.revenue)  || 0;
    const cogs = (inc[m] && inc[m].cogs)     || (id && id.cogs)     || 0;
    const sga  = (inc[m] && inc[m].sga)      || (id && id.sga)      || 0;
    const mfg  = (cst[m] && cst[m].mfg)      || (id && id.mfg)      || 0;
    const intr = (inc[m] && inc[m].interest) || (id && id.interest) || 0;
    const tgt  = id ? (id.targetRevenue || 0) : 0;
    const hasData = !!(rev || cogs || sga || mfg || intr);
    // 결재완료 월: 기본 체크 해제 (실수 덮어쓰기 방지), 체크는 가능
    const isDefaultChecked = hasData && !approved && (yr === curY ? m === curM : false);
    return { m, rev, cogs, sga, mfg, intr, tgt, approved, hasData, isDefaultChecked };
  });

  // 기본 체크된 행이 없으면 마지막 데이터 행을 체크
  if (!rows.some(r => r.isDefaultChecked)) {
    const lastWithData = [...rows].reverse().find(r => r.hasData && !r.approved);
    if (lastWithData) lastWithData.isDefaultChecked = true;
  }

  const fv = v => v ? _pf(v) : "";
  const tableRows = rows.map(r => `
    <tr data-month="${r.m}" class="${r.approved ? "pnl-id-row-locked" : ""}${!r.hasData ? " pnl-id-row-empty" : ""}">
      <td class="pnl-id-chk-cell">
        <input type="checkbox" class="pnl-id-chk" ${r.isDefaultChecked ? "checked" : ""} />
      </td>
      <td class="pnl-id-m">${r.m}월${r.approved ? " 🔒" : ""}</td>
      <td><input type="text" class="pnl-id-inp" data-f="revenue"       value="${fv(r.rev)}"  placeholder="0" inputmode="numeric" /></td>
      <td><input type="text" class="pnl-id-inp" data-f="cogs"          value="${fv(r.cogs)}" placeholder="0" inputmode="numeric" /></td>
      <td><input type="text" class="pnl-id-inp" data-f="sga"           value="${fv(r.sga)}"  placeholder="0" inputmode="numeric" /></td>
      <td><input type="text" class="pnl-id-inp" data-f="mfg"           value="${fv(r.mfg)}"  placeholder="0" inputmode="numeric" /></td>
      <td><input type="text" class="pnl-id-inp" data-f="interest"      value="${fv(r.intr)}" placeholder="0" inputmode="numeric" /></td>
      <td><input type="text" class="pnl-id-inp pnl-id-manual" data-f="targetRevenue" value="${fv(r.tgt)}" placeholder="수동입력" inputmode="numeric" /></td>
      <td class="pnl-id-calc" data-calc="gross">—</td>
      <td class="pnl-id-calc" data-calc="op">—</td>
      <td class="pnl-id-calc" data-calc="mgmt">—</td>
    </tr>`).join("");

  const overlay = document.createElement("div");
  overlay.id = "pnlImportOverlay";
  overlay.className = "pnl-import-overlay";
  overlay.innerHTML = `
    <div class="pnl-import-dialog">
      <div class="pnl-id-header">
        <span class="pnl-id-title">Excel 일괄 입력 — 미리보기</span>
        <select id="pnlIdYearSel">${yearOpts}</select>
        <button class="pnl-id-close" id="pnlIdClose">✕</button>
      </div>
      <div class="pnl-id-hint">
        ☑ 체크한 월만 저장됩니다. 🔒 결재완료 월은 기본 체크 해제 — 체크 후 수정 가능합니다.<br/>
        회색 셀은 Excel에서 자동 추출된 값이며, <span class="pnl-id-manual-hint">목표매출</span>만 직접 입력하세요.
      </div>
      <div class="pnl-id-sel-btns">
        <button class="pnl-id-sel-btn" id="pnlIdSelAll">전체선택</button>
        <button class="pnl-id-sel-btn" id="pnlIdSelNone">전체해제</button>
        <button class="pnl-id-sel-btn pnl-id-sel-cur" id="pnlIdSelCur">이번달만</button>
      </div>
      <div class="pnl-id-table-wrap">
        <table class="pnl-id-table">
          <thead><tr>
            <th class="pnl-id-chk-head">저장</th>
            <th>월</th><th>매출액</th><th>매출원가</th><th>판관비</th>
            <th>제조원가</th><th>영업외비용</th><th class="pnl-id-manual-col">목표매출</th>
            <th class="pnl-id-calc-col">매출총이익</th>
            <th class="pnl-id-calc-col">영업이익</th>
            <th class="pnl-id-calc-col">경영이익</th>
          </tr></thead>
          <tbody>${tableRows}</tbody>
        </table>
      </div>
      <div class="pnl-id-footer">
        <span id="pnlIdMsg" class="pnl-id-msg"></span>
        <button class="pnl-btn pnl-btn-ghost" id="pnlIdCancel">취소</button>
        <button class="pnl-btn pnl-btn-primary" id="pnlIdSave">선택 저장</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);

  function parseN(s) { return Number(String(s).replace(/[^0-9]/g, "")) || 0; }
  function getRowVals(tr) {
    const v = {};
    tr.querySelectorAll(".pnl-id-inp").forEach(inp => { v[inp.dataset.f] = parseN(inp.value); });
    return v;
  }
  function refreshCalcCells() {
    overlay.querySelectorAll("tbody tr").forEach(tr => {
      const v = getRowVals(tr);
      const c = calcPnl(v);
      const hasVals = !!(v.revenue || v.cogs || v.mfg || v.sga || v.interest);
      [
        { key: "gross", val: c.gross },
        { key: "op",    val: c.op    },
        { key: "mgmt",  val: c.mgmt  },
      ].forEach(({ key, val }) => {
        const td = tr.querySelector(`[data-calc=${key}]`);
        if (!td) return;
        if (hasVals) {
          td.textContent = _ps(val) + " 원";
          td.className = "pnl-id-calc " + _pc(val);
        } else {
          td.textContent = "—";
          td.className = "pnl-id-calc";
        }
      });
    });
  }

  overlay.querySelectorAll(".pnl-id-inp").forEach(inp => {
    inp.addEventListener("input", () => {
      const raw = parseN(inp.value);
      inp.value = raw ? _pf(raw) : "";
      refreshCalcCells();
    });
  });
  refreshCalcCells();

  // 체크박스 → 행 dimming
  function updateRowDim(tr) {
    const chk = tr.querySelector(".pnl-id-chk");
    if (!chk || chk.disabled) return;
    tr.classList.toggle("pnl-id-row-unchecked", !chk.checked);
  }
  overlay.querySelectorAll("tbody tr").forEach(tr => {
    updateRowDim(tr);
    const chk = tr.querySelector(".pnl-id-chk");
    if (chk) chk.addEventListener("change", () => updateRowDim(tr));
  });

  // 전체선택 / 전체해제 / 이번달만
  overlay.querySelector("#pnlIdSelAll").addEventListener("click", () => {
    overlay.querySelectorAll(".pnl-id-chk").forEach(c => { c.checked = true; updateRowDim(c.closest("tr")); });
  });
  overlay.querySelector("#pnlIdSelNone").addEventListener("click", () => {
    overlay.querySelectorAll(".pnl-id-chk").forEach(c => { c.checked = false; updateRowDim(c.closest("tr")); });
  });
  overlay.querySelector("#pnlIdSelCur").addEventListener("click", () => {
    overlay.querySelectorAll(".pnl-id-chk").forEach(c => {
      const m = parseInt(c.closest("tr").dataset.month);
      c.checked = (m === curM);
      updateRowDim(c.closest("tr"));
    });
  });

  document.getElementById("pnlIdYearSel").addEventListener("change", e => {
    _pnlImportYear = parseInt(e.target.value);
    overlay.remove();
    openPnlImportDialog();
  });
  document.getElementById("pnlIdClose").addEventListener("click", () => overlay.remove());
  document.getElementById("pnlIdCancel").addEventListener("click", () => overlay.remove());

  document.getElementById("pnlIdSave").addEventListener("click", () => {
    const yr2 = parseInt(document.getElementById("pnlIdYearSel").value);
    let saved = 0;
    overlay.querySelectorAll("tbody tr").forEach(tr => {
      const chk = tr.querySelector(".pnl-id-chk");
      if (!chk || !chk.checked) return; // 체크 안 된 행 skip
      const m = parseInt(tr.dataset.month);
      const v = getRowVals(tr);
      if (!v.revenue && !v.cogs && !v.sga && !v.mfg && !v.interest) return; // 빈 행 skip
      const existing = getPnlEntry(yr2, m) || {};
      upsertPnlEntry({
        ...existing,
        year: yr2, month: m,
        revenue: v.revenue, cogs: v.cogs, sga: v.sga,
        mfg: v.mfg, interest: v.interest,
        targetRevenue: v.targetRevenue || existing.targetRevenue || 0,
        approvalStatus: existing.approvalStatus || "draft",
        purchaseAmount: (inc[m] && inc[m].purchaseAmount) || existing.purchaseAmount || 0,
      });
      saved++;
    });
    document.getElementById("pnlIdMsg").textContent = `${saved}개월 저장 완료`;
    setTimeout(() => { overlay.remove(); renderPnlTab(); }, 800);
  });
}

// ── 입력 탭 ──────────────────────────────────────────────────
// ── 결재완료 후 수정 사유 다이얼로그 ──────────────────────────
function openPnlCorrectionDialog(changes, onConfirm) {
  document.querySelector(".pnl-correction-overlay")?.remove();
  const overlay = document.createElement("div");
  overlay.className = "raw-diff-overlay pnl-correction-overlay";
  overlay.innerHTML = `
    <div class="raw-diff-dialog pnl-correction-dialog">
      <div class="raw-diff-header">
        <h3>결재 진행 중인 보고서 수정</h3>
        <span class="raw-diff-sub">이미 서명이 진행된 보고서입니다. 수정 사유를 입력하면 저장과 동시에 결재가 취소되어 기안부터 다시 서명해야 합니다.</span>
      </div>
      <div class="raw-diff-section">
        <div class="raw-diff-section-title changed-title">변경 항목</div>
        ${changes.map(c => `
          <div class="raw-diff-row">
            <span class="raw-diff-label">${escapeHtml(c.label)}</span>
            <span class="raw-diff-amounts">${_pf(c.oldValue)} → <strong>${_pf(c.newValue)}</strong></span>
          </div>`).join("")}
      </div>
      <div class="pnl-correction-reason-wrap">
        <label class="pnl-correction-reason-label">수정 사유 <span style="color:#dc2626">*</span></label>
        <textarea id="pnlCorrectionReason" class="pnl-correction-reason-input" rows="3" placeholder="예: 판관비 과대 입력 정정 (중복 반영된 항목 제외)"></textarea>
      </div>
      <div class="raw-diff-actions">
        <button type="button" class="diff-cancel-btn" id="pnlCorrectionCancelBtn">취소</button>
        <button type="button" class="diff-confirm-btn" id="pnlCorrectionConfirmBtn">사유 저장 후 수정 반영</button>
      </div>
    </div>`;
  document.body.appendChild(overlay);
  const reasonInput = overlay.querySelector("#pnlCorrectionReason");
  reasonInput.focus();
  overlay.querySelector("#pnlCorrectionConfirmBtn").addEventListener("click", () => {
    const reason = reasonInput.value.trim();
    if (!reason) {
      reasonInput.style.borderColor = "#dc2626";
      reasonInput.placeholder = "수정 사유를 입력해야 저장할 수 있습니다";
      return;
    }
    overlay.remove();
    onConfirm(reason);
  });
  overlay.querySelector("#pnlCorrectionCancelBtn").addEventListener("click", () => overlay.remove());
}

function renderPnlInput(el) {
  const entry = getPnlEntry(pnlInputYear, pnlInputMonth) || {
    year: pnlInputYear, month: pnlInputMonth,
    revenue:0, targetRevenue:0, cogs:0, mfg:0, sga:0, interest:0,
    approvalStatus:"draft", draftDate:"", agree1Date:"", agree2Date:"",
    ceoDate:"", docNo:"", ceoComment:"",
  };
  const c = calcPnl(entry);
  const hasManualInv = entry.beginInventory !== undefined && entry.endInventory !== undefined;
  const statusLabels = { draft:"작성중", 기안:"기안완료", 합의1:"합의①완료", 합의2:"합의②완료", 결재완료:"결재완료" };
  const statusColors = { draft:"#6b7280", 기안:"#2563eb", 합의1:"#7c3aed", 합의2:"#d97706", 결재완료:"#16a34a" };
  const st = entry.approvalStatus || "draft";
  const curY = new Date().getFullYear();
  const yearOpts = Array.from({length: curY - 2023}, (_,i) => 2024 + i).map(y =>
    `<option value="${y}" ${y===pnlInputYear?"selected":""}>${y}년</option>`).join("");
  const monOpts  = Array.from({length:12},(_,i) =>
    `<option value="${i+1}" ${i+1===pnlInputMonth?"selected":""}>${i+1}월</option>`).join("");

  const fields = [
    { key:"revenue",       id:"pnlRev",    label:"매출액",       displayVal: entry.revenue       },
    { key:"targetRevenue", id:"pnlTgt",    label:"목표매출액",   displayVal: entry.targetRevenue  },
    // 재고 탭 수동 입력 시: effectiveCogs(c.cogs) 표시 + 안내 뱃지
    { key:"cogs",          id:"pnlCogs",   label:"상품매출원가", displayVal: c.cogs,
      note: hasManualInv ? "📦 재고탭 적용중" : null },
    { key:"mfg",           id:"pnlMfg",    label:"당기총제조비용", displayVal: entry.mfg         },
    { key:"sga",           id:"pnlSga",    label:"판매관리비",   displayVal: entry.sga           },
    { key:"interest",      id:"pnlInt",    label:"영업외비용",   displayVal: entry.interest      },
  ];

  const incLabel = _pnlImportIncome ? "손익계산서" : "손익계산서 업로드";
  const cstLabel = _pnlImportCost   ? "원가명세서" : "원가명세서 업로드";
  const canPreview = _pnlImportIncome || _pnlImportCost;

  el.innerHTML = `
    <div class="pnl-input-wrap">

      <div class="pnl-import-section">
        <span class="pnl-import-label">Excel 일괄 입력</span>
        <button class="pnl-import-btn${_pnlImportIncome?" pnl-import-done":""}" id="pnlIncomeUploadBtn">${incLabel}</button>
        <button class="pnl-import-btn${_pnlImportCost?" pnl-import-done":""}" id="pnlCostUploadBtn">${cstLabel}</button>
        ${canPreview ? `<button class="pnl-btn pnl-btn-primary pnl-import-preview-btn" id="pnlImportPreviewBtn">미리보기 / 일괄저장 →</button>` : ""}
        <input type="file" id="pnlIncomeFileInput" accept=".xls,.xlsx" hidden />
        <input type="file" id="pnlCostFileInput"   accept=".xls,.xlsx" hidden />
      </div>

      <div class="pnl-input-nav">
        <button class="pnl-nav-btn" id="pnlNavPrev">◀</button>
        <select id="pnlSelYear">${yearOpts}</select>
        <select id="pnlSelMonth">${monOpts}</select>
        <button class="pnl-nav-btn" id="pnlNavNext">▶</button>
        <span class="pnl-status-badge" style="background:${statusColors[st]}18;color:${statusColors[st]};border:1px solid ${statusColors[st]}40">${statusLabels[st]||st}</span>
      </div>

      <div class="pnl-annual-dist-row no-print">
        <span class="pnl-annual-dist-label">연간 목표 배분</span>
        <input type="text" id="pnlAnnualTgt" class="pnl-annual-tgt-input" placeholder="연간 목표 총액 입력" inputmode="numeric" />
        <span class="pnl-field-unit">원</span>
        <button id="pnlAnnualDistBtn" class="pnl-btn pnl-btn-ghost pnl-annual-dist-btn">÷12 → 전월 배분</button>
      </div>

      <div class="pnl-form-body">
        <div class="pnl-form-card">
          <div class="pnl-form-title">입력 항목</div>
          ${fields.map(f => `
            <div class="pnl-field-row">
              <label class="pnl-field-label">${f.label}${f.note ? `<span class="pnl-inv-badge">${f.note}</span>` : ""}</label>
              <input type="text" id="${f.id}" class="pnl-field-input${f.note ? " pnl-field-inv" : ""}" value="${f.displayVal>0?_pf(f.displayVal):""}" placeholder="0" inputmode="numeric" />
              <span class="pnl-field-unit">원</span>
            </div>`).join("")}
        </div>

        <div class="pnl-preview-card">
          <div class="pnl-form-title">자동 계산</div>
          <div class="pnl-prev-row"><span>매출총이익</span><strong id="prvGross" class="${_pc(c.gross)}">${_ps(c.gross)} 원</strong></div>
          <div class="pnl-prev-row pnl-muted"><span>총이익률</span><strong id="prvGmRate">${c.gmRate.toFixed(1)}%</strong></div>
          <div class="pnl-prev-row ${c.targetAchieve!==null?(c.targetAchieve>=100?"pnl-pos":"pnl-neg"):"pnl-muted"}">
            <span>목표 달성률</span>
            <strong id="prvTarget">${c.targetAchieve!==null?c.targetAchieve.toFixed(1)+"%":"—"}</strong>
          </div>
          <div class="pnl-prev-divider"></div>
          <div class="pnl-prev-row"><span>영업이익</span><strong id="prvOp" class="${_pc(c.op)}">${_ps(c.op)} 원</strong></div>
          <div class="pnl-prev-row pnl-muted"><span>영업이익률</span><strong id="prvOpRate">${c.opRate.toFixed(1)}%</strong></div>
          <div class="pnl-prev-divider"></div>
          <div class="pnl-prev-row pnl-prev-highlight"><span>경영이익</span><strong id="prvMgmt" class="${_pc(c.mgmt)}">${_ps(c.mgmt)} 원</strong></div>
        </div>
      </div>

      <div class="pnl-input-actions">
        <button id="pnlSaveBtn" class="pnl-btn pnl-btn-primary">저장</button>
        <button id="pnlDeleteBtn" class="pnl-btn pnl-btn-danger">삭제</button>
        <button id="pnlToReportBtn" class="pnl-btn pnl-btn-ghost">보고서로 →</button>
      </div>

      ${(entry.corrections && entry.corrections.length) ? `
      <div class="pnl-correction-history">
        <div class="pnl-correction-history-title">📝 수정이력 (${entry.corrections.length}건)</div>
        ${entry.corrections.slice().reverse().map(c => `
          <div class="pnl-correction-item">
            <div class="pnl-correction-item-head"><span>${c.date}</span><span class="pnl-correction-from">${c.fromStatus} 상태에서 수정</span></div>
            <div class="pnl-correction-reason">${escapeHtml(c.reason)}</div>
            <div class="pnl-correction-changes">${c.changes.map(ch => `${ch.label} ${_pf(ch.oldValue)}원 → ${_pf(ch.newValue)}원`).join(" · ")}</div>
          </div>`).join("")}
      </div>` : ""}
    </div>`;

  // 월 이동
  document.getElementById("pnlNavPrev").addEventListener("click", () => {
    pnlInputMonth--; if (pnlInputMonth < 1) { pnlInputMonth = 12; pnlInputYear--; }
    renderPnlTab();
  });
  document.getElementById("pnlNavNext").addEventListener("click", () => {
    pnlInputMonth++; if (pnlInputMonth > 12) { pnlInputMonth = 1; pnlInputYear++; }
    renderPnlTab();
  });
  document.getElementById("pnlSelYear").addEventListener("change", e => { pnlInputYear  = +e.target.value; renderPnlTab(); });
  document.getElementById("pnlSelMonth").addEventListener("change", e => { pnlInputMonth = +e.target.value; renderPnlTab(); });

  function parseN(s) { return Number(String(s).replace(/[^0-9]/g,"")) || 0; }
  function getVals() {
    const v = {};
    fields.forEach(f => { v[f.key] = parseN(document.getElementById(f.id)?.value || ""); });
    return v;
  }
  function refreshPreview() {
    // 저장 시와 동일하게 entry(beginInventory 등 포함)에 입력값 병합 후 계산
    const cv = calcPnl({ ...entry, ...getVals() });
    const g = document.getElementById("prvGross");   if (g) { g.textContent = `${_ps(cv.gross)} 원`; g.className = _pc(cv.gross); }
    const gm = document.getElementById("prvGmRate"); if (gm) gm.textContent = `${cv.gmRate.toFixed(1)}%`;
    const tg = document.getElementById("prvTarget"); if (tg) {
      tg.textContent = cv.targetAchieve !== null ? `${cv.targetAchieve.toFixed(1)}%` : "—";
      tg.className   = cv.targetAchieve !== null ? (cv.targetAchieve >= 100 ? "pnl-pos" : "pnl-neg") : "";
    }
    const op = document.getElementById("prvOp");   if (op) { op.textContent = `${_ps(cv.op)} 원`; op.className = _pc(cv.op); }
    const or = document.getElementById("prvOpRate"); if (or) or.textContent = `${cv.opRate.toFixed(1)}%`;
    const mg = document.getElementById("prvMgmt"); if (mg) { mg.textContent = `${_ps(cv.mgmt)} 원`; mg.className = _pc(cv.mgmt); }
  }

  fields.forEach(f => {
    document.getElementById(f.id)?.addEventListener("input", e => {
      const raw = parseN(e.target.value);
      e.target.value = raw ? _pf(raw) : "";
      refreshPreview();
    });
  });

  document.getElementById("pnlAnnualDistBtn")?.addEventListener("click", () => {
    const annual = parseN(document.getElementById("pnlAnnualTgt")?.value || "");
    if (!annual) { pnlToast("연간 목표 총액을 입력하세요"); return; }
    const monthly = Math.round(annual / 12);
    for (let m = 1; m <= 12; m++) {
      const existing = getPnlEntry(pnlInputYear, m) || {
        year: pnlInputYear, month: m, revenue: 0, targetRevenue: 0,
        cogs: 0, mfg: 0, sga: 0, interest: 0,
        approvalStatus: "draft", draftDate: "", agree1Date: "", agree2Date: "",
        ceoDate: "", docNo: "", ceoComment: "",
      };
      upsertPnlEntry({ ...existing, targetRevenue: monthly });
    }
    const tgtEl = document.getElementById("pnlTgt");
    if (tgtEl) { tgtEl.value = _pf(monthly); refreshPreview(); }
    pnlToast(`${pnlInputYear}년 12개월 목표 배분 완료 (월 ${_pf(monthly)}원)`);
  });

  document.getElementById("pnlSaveBtn").addEventListener("click", () => {
    const newVals = getVals();
    const wasSigned = (entry.approvalStatus || "draft") !== "draft";
    if (!wasSigned) {
      const newEntry = { ...entry, ...newVals };
      if (!newEntry.approvalStatus) newEntry.approvalStatus = "draft";
      upsertPnlEntry(newEntry);
      pnlToast("저장되었습니다.");
      renderPnlTab();
      return;
    }
    // 이미 서명이 진행된 보고서 — 수정 사유 입력 필수
    const changes = fields
      .filter(f => (Number(entry[f.key]) || 0) !== newVals[f.key])
      .map(f => ({ label: f.label, oldValue: Number(entry[f.key]) || 0, newValue: newVals[f.key] }));
    if (!changes.length) { pnlToast("변경된 값이 없습니다."); return; }
    openPnlCorrectionDialog(changes, reason => {
      const newEntry = { ...entry, ...newVals };
      PNL_APPROVAL_STEPS.forEach(step => { newEntry[step.dateKey] = ""; });
      newEntry.approvalStatus = "draft";
      newEntry.docNo = "";
      newEntry.corrections = [...(entry.corrections || []), {
        date: _todayKor(), reason, fromStatus: entry.approvalStatus, changes,
      }];
      upsertPnlEntry(newEntry);
      writePnlPendingToFirebase();
      pnlToast("수정사유가 저장되었습니다. 결재가 취소되어 기안부터 다시 서명해야 합니다.");
      renderPnlTab();
    });
  });
  document.getElementById("pnlDeleteBtn").addEventListener("click", () => {
    if (!getPnlEntry(pnlInputYear, pnlInputMonth)) { pnlToast("저장된 데이터가 없습니다."); return; }
    if (!confirm(`${pnlInputYear}년 ${pnlInputMonth}월 데이터를 삭제하시겠습니까?`)) return;
    deletePnlEntry(pnlInputYear, pnlInputMonth);
    pnlToast("삭제되었습니다.");
    renderPnlTab();
  });
  document.getElementById("pnlToReportBtn").addEventListener("click", () => {
    pnlRptYear = pnlInputYear; pnlRptMonth = pnlInputMonth;
    pnlSubTab = "report"; renderPnlTab();
  });

  // ── Excel 일괄입력 파일 핸들러 ──
  function handlePnlFile(inputId, type) {
    const inp = document.getElementById(inputId);
    if (!inp) return;
    inp.value = "";
    inp.onchange = e => {
      const file = e.target.files[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = ev => {
        try {
          const wb = XLSX.read(new Uint8Array(ev.target.result), { type: "array", cellDates: true });
          if (type === "income") {
            const { data, year } = parsePnlIncomeStatement(wb);
            if (!data) { pnlToast("손익계산서 파싱 실패 — 시트 구조를 확인하세요."); return; }
            _pnlImportIncome = data;
            if (year) _pnlImportYear = year;
          } else {
            const { data, year } = parsePnlCostStatement(wb);
            if (!data) { pnlToast("원가명세서 파싱 실패 — 시트 구조를 확인하세요."); return; }
            _pnlImportCost = data;
            if (year) _pnlImportYear = year;
          }
          pnlToast(`${type === "income" ? "손익계산서" : "원가명세서"} 파싱 완료 (${_pnlImportYear}년)`);
          renderPnlTab();
        } catch (err) {
          pnlToast("파일 읽기 오류: " + (err.message || err));
        }
      };
      reader.readAsArrayBuffer(file);
    };
    inp.click();
  }

  document.getElementById("pnlIncomeUploadBtn").addEventListener("click", () => handlePnlFile("pnlIncomeFileInput", "income"));
  document.getElementById("pnlCostUploadBtn").addEventListener("click",   () => handlePnlFile("pnlCostFileInput",   "cost"));
  document.getElementById("pnlImportPreviewBtn")?.addEventListener("click", openPnlImportDialog);
}

// ── 보고서 탭 ─────────────────────────────────────────────────
function renderPnlReport(el) {
  if (pnlRptMode === "quarterly") { renderPnlQuarterlyReport(el); return; }
  if (pnlRptMode === "halfyear")  { renderPnlHalfYearReport(el);  return; }
  if (pnlRptMode === "annual")    { renderPnlAnnualReport(el);    return; }
  const curY = new Date().getFullYear();
  const maxYear = Math.max(curY + 1, pnlRptYear);
  const yearOpts = Array.from({length: maxYear - 2023}, (_,i) => 2024+i).map(y =>
    `<option value="${y}" ${y===pnlRptYear?"selected":""}>${y}년</option>`).join("");
  const monOpts  = Array.from({length:12},(_,i) =>
    `<option value="${i+1}" ${i+1===pnlRptMonth?"selected":""}>${i+1}월</option>`).join("");

  const entry = getPnlEntry(pnlRptYear, pnlRptMonth);
  const c = entry ? calcPnl(entry) : null;
  const prevM = pnlRptMonth > 1 ? pnlRptMonth - 1 : 12;
  const prevY = pnlRptMonth > 1 ? pnlRptYear       : pnlRptYear - 1;
  const prev  = getPnlEntry(prevY, prevM);
  const pc    = prev ? calcPnl(prev) : null;

  const statusIdx = entry ? _pnlStatusIdx(entry.approvalStatus) : 0;

  function approvalBox(stepIdx) {
    const step = PNL_APPROVAL_STEPS[stepIdx];
    const done = statusIdx > stepIdx;
    const current = statusIdx === stepIdx && !!entry;
    const date = entry?.[step.dateKey] || "";
    return `
      <div class="pnl-ap-box">
        <div class="pnl-ap-role">${step.role}</div>
        <div class="pnl-ap-name">${step.name} ${step.title}</div>
        <div class="pnl-ap-date">${done && date ? date : "&nbsp;"}</div>
        ${done
          ? `<div class="pnl-stamp ${step.stampCls}">서명
               <button class="pnl-revoke-btn" data-step="${stepIdx}" title="결재 취소">↩</button>
             </div>`
          : current
            ? `<button class="pnl-sign-btn" data-step="${stepIdx}">서명<br>하기</button>`
            : `<div class="pnl-stamp pnl-stamp-empty"></div>`
        }
      </div>`;
  }

  function cmpRow(label, prev, curr, isSub) {
    if (!pc || !c) return "";
    const diff = curr - prev;
    const diffCls = diff > 0 ? "pnl-pos" : diff < 0 ? "pnl-neg" : "";
    const trCls = label === "경영이익(손실)" ? "pnl-tr-total" : isSub ? "pnl-tr-sub" : "";
    return `<tr class="${trCls}">
      <td>${label}</td>
      <td>${_ps(prev)}</td>
      <td>${_ps(curr)}</td>
      <td class="${diffCls}">${diff >= 0 ? "▲ " : "▼ "}${_pf(diff)}</td>
    </tr>`;
  }

  const noDataHtml = `<div class="pnl-no-data">이 달의 데이터가 없습니다. 입력 탭에서 먼저 저장해 주세요.</div>`;

  el.innerHTML = `
    <div class="pnl-report-wrap">
      <!-- 툴바 (인쇄 시 숨김) -->
      <div class="pnl-report-toolbar no-print">
        <div class="pnl-rpt-mode-tabs">
          <button class="pnl-mode-btn active" data-rpt-mode="monthly">월간</button>
          <button class="pnl-mode-btn" data-rpt-mode="quarterly">분기</button>
          <button class="pnl-mode-btn" data-rpt-mode="halfyear">반기</button>
          <button class="pnl-mode-btn" data-rpt-mode="annual">연간</button>
        </div>
        <button class="pnl-nav-btn" id="pnlRptPrev">◀</button>
        <select id="pnlRptYear">${yearOpts}</select>
        <select id="pnlRptMonth">${monOpts}</select>
        <button class="pnl-nav-btn" id="pnlRptNext">▶</button>
        <button class="pnl-btn pnl-btn-print" id="pnlPrintBtn">인쇄 / PDF</button>
      </div>

      <!-- 보고서 본체 -->
      <div class="pnl-page" id="pnlReportPage">
        <!-- 헤더 -->
        <div class="pnl-doc-header">
          <div class="pnl-company-badge">MIRAE AUTOMATION CO., LTD</div>
          <div class="pnl-doc-title">${pnlRptYear}년 ${pnlRptMonth}월 &mdash; 월간 경영손익 보고서</div>
          <div class="pnl-doc-sub">관리기준 손익 및 경영이익 산출내역</div>
          <div class="pnl-meta-row">
            <div class="pnl-meta-item"><span class="pnl-meta-lbl">기안부서</span><span class="pnl-meta-val">${PNL_META.department}</span></div>
            <div class="pnl-meta-item"><span class="pnl-meta-lbl">기안자</span><span class="pnl-meta-val">${PNL_META.author.name} ${PNL_META.author.title}</span></div>
            <div class="pnl-meta-item"><span class="pnl-meta-lbl">기안일</span><span class="pnl-meta-val">${entry?.draftDate || "—"}</span></div>
            <div class="pnl-meta-item"><span class="pnl-meta-lbl">문서번호</span>
              <span class="pnl-meta-val" id="pnlDocNoVal" contenteditable="${!!entry}" style="outline:none;cursor:${entry?"text":"default"}">${entry?.docNo || "—"}</span>
            </div>
          </div>
        </div>

        <div class="pnl-doc-body">
          ${!entry ? noDataHtml : `
          <!-- KPI 카드 -->
          <div class="pnl-kpi-grid">
            <div class="pnl-kpi-card">
              <div class="pnl-kpi-lbl">매출액</div>
              <div class="pnl-kpi-val">${_pf(entry.revenue)}</div>
              <div class="pnl-kpi-unit">원${entry.targetRevenue>0?" / 목표 "+_pf(entry.targetRevenue)+"원":""}</div>
              ${c.targetAchieve!==null?`<div class="pnl-kpi-achieve ${c.targetAchieve>=100?"pnl-pos":"pnl-neg"}">달성률 ${c.targetAchieve.toFixed(1)}%</div>`:""}
            </div>
            <div class="pnl-kpi-card">
              <div class="pnl-kpi-lbl">매출총이익</div>
              <div class="pnl-kpi-val ${_pc(c.gross)}">${_ps(c.gross)}</div>
              <div class="pnl-kpi-unit">원 / 총이익률 ${c.gmRate.toFixed(1)}%</div>
            </div>
            <div class="pnl-kpi-card pnl-kpi-highlight">
              <div class="pnl-kpi-lbl">경영이익(손실)</div>
              <div class="pnl-kpi-val ${_pc(c.mgmt)}">${_ps(c.mgmt)}</div>
              <div class="pnl-kpi-unit">금융비용 반영 기준</div>
            </div>
          </div>

          <!-- 섹션1 영업이익 -->
          <div class="pnl-section">
            <div class="pnl-sec-title"><span class="pnl-sec-num">1</span>관리기준 영업이익 <small>실질 기준</small></div>
            <div class="pnl-flow">
              <div class="pnl-flow-row"><span class="pnl-flow-lbl"><span class="pnl-tag">ㄱ</span> 매출액</span><span class="pnl-flow-val">${_pf(entry.revenue)} 원</span></div>
              <div class="pnl-flow-divider">차감</div>
              <div class="pnl-flow-row pnl-indent"><span class="pnl-flow-lbl"><span class="pnl-minus">−</span><span class="pnl-tag">ㄴ</span> 상품매출원가</span><span class="pnl-flow-val pnl-neg">(${_pf(c.cogs)}) 원</span></div>
              <div class="pnl-flow-row pnl-indent"><span class="pnl-flow-lbl"><span class="pnl-minus">−</span><span class="pnl-tag">ㄷ</span> 당기총제조비용</span><span class="pnl-flow-val pnl-neg">(${_pf(entry.mfg)}) 원</span></div>
              <div class="pnl-flow-row pnl-flow-sub"><span class="pnl-flow-lbl">① 매출총이익 <small>[ㄱ−(ㄴ+ㄷ)]</small></span><span class="pnl-flow-val ${_pc(c.gross)}">${_ps(c.gross)} 원</span></div>
              <div class="pnl-flow-divider">차감</div>
              <div class="pnl-flow-row pnl-indent"><span class="pnl-flow-lbl"><span class="pnl-minus">−</span> 판매관리비</span><span class="pnl-flow-val pnl-neg">(${_pf(entry.sga)}) 원</span></div>
              <div class="pnl-flow-row pnl-flow-total"><span class="pnl-flow-lbl">② 관리기준 영업이익 <small>[①−판관비]</small></span><span class="pnl-flow-val ${_pc(c.op)}">${_ps(c.op)} 원</span></div>
            </div>
          </div>

          <!-- 섹션2 경영이익 -->
          <div class="pnl-section">
            <div class="pnl-sec-title"><span class="pnl-sec-num">2</span>경영이익 <small>금융비용 반영 기준</small></div>
            <div class="pnl-flow">
              <div class="pnl-flow-row"><span class="pnl-flow-lbl">관리기준 영업이익 (②)</span><span class="pnl-flow-val ${_pc(c.op)}">${_ps(c.op)} 원</span></div>
              <div class="pnl-flow-divider">차감</div>
              <div class="pnl-flow-row pnl-indent"><span class="pnl-flow-lbl"><span class="pnl-minus">−</span> 영업외비용</span><span class="pnl-flow-val pnl-neg">(${_pf(entry.interest)}) 원</span></div>
              <div class="pnl-flow-row pnl-flow-total"><span class="pnl-flow-lbl">③ 경영이익(손실) <small>[②−영업외비용]</small></span><span class="pnl-flow-val ${_pc(c.mgmt)}">${_ps(c.mgmt)} 원</span></div>
            </div>
            <div class="pnl-remark">※ 비고: 영업외비용(이자비용 등) 반영 시 실제 경영성과를 함께 확인할 수 있도록 별도 표시하였습니다.</div>
          </div>

          <!-- 섹션3 전월 비교 -->
          ${prev && pc ? `
          <div class="pnl-section">
            <div class="pnl-sec-title"><span class="pnl-sec-num">3</span>전월 대비 손익 비교</div>
            <table class="pnl-cmp-table">
              <thead><tr><th>항목</th><th>${prevY}년 ${prevM}월</th><th>${pnlRptYear}년 ${pnlRptMonth}월</th><th>증감액</th></tr></thead>
              <tbody>
                ${cmpRow("매출액",          prev.revenue,       entry.revenue,    false)}
                ${cmpRow("상품매출원가",    pc.cogs,            c.cogs,           false)}
                ${cmpRow("당기총제조비용",  prev.mfg,           entry.mfg,        false)}
                ${cmpRow("매출총이익",      pc.gross,           c.gross,          true)}
                ${cmpRow("판관비",          prev.sga,           entry.sga,        false)}
                ${cmpRow("관리기준 영업이익", pc.op,            c.op,             true)}
                ${cmpRow("영업외비용",       prev.interest,      entry.interest,   false)}
                ${cmpRow("경영이익(손실)",  pc.mgmt,            c.mgmt,           true)}
              </tbody>
            </table>
            ${entry.ceoComment ? `<div class="pnl-remark">대표이사 의견: "<em>${escapeHtml(entry.ceoComment)}</em>"</div>` : ""}
          </div>` : ""}

          <!-- 결재란 -->
          <div class="pnl-section">
            <div class="pnl-sec-title"><span class="pnl-sec-num" style="font-size:11px">✓</span>결재</div>
            <div class="pnl-ap-grid">
              ${PNL_APPROVAL_STEPS.map((_,i) => approvalBox(i)).join("")}
            </div>
            <div class="pnl-ceo-comment-row no-print">
              <label>대표이사 의견:</label>
              <input type="text" id="pnlCeoComment" class="pnl-ceo-input" value="${escapeHtml(entry?.ceoComment||"")}" placeholder="의견을 입력하세요" />
            </div>
          </div>

          ${entry?.corrections?.length ? `
          <div class="pnl-section">
            <div class="pnl-sec-title"><span class="pnl-sec-num" style="font-size:11px">✎</span>수정이력</div>
            ${entry.corrections.map(c => `
              <div class="pnl-correction-item">
                <div class="pnl-correction-item-head"><span>${c.date}</span><span class="pnl-correction-from">${c.fromStatus} → 재기안</span></div>
                <div class="pnl-correction-reason">${escapeHtml(c.reason)}</div>
                <div class="pnl-correction-changes">${c.changes.map(ch => `${ch.label} ${_pf(ch.oldValue)}원 → ${_pf(ch.newValue)}원`).join(" · ")}</div>
              </div>`).join("")}
          </div>` : ""}
          `}
        </div><!-- /pnl-doc-body -->
        <div class="pnl-doc-footer">${PNL_META.companyName} · ${PNL_META.department} · 대외비</div>
      </div><!-- /pnl-page -->
    </div>`;

  // 툴바 이벤트
  document.getElementById("pnlRptPrev")?.addEventListener("click", () => {
    pnlRptMonth--; if (pnlRptMonth < 1) { pnlRptMonth = 12; pnlRptYear--; } renderPnlReport(el);
  });
  document.getElementById("pnlRptNext")?.addEventListener("click", () => {
    pnlRptMonth++; if (pnlRptMonth > 12) { pnlRptMonth = 1; pnlRptYear++; } renderPnlReport(el);
  });
  document.getElementById("pnlRptYear")?.addEventListener("change", e => { pnlRptYear  = +e.target.value; renderPnlReport(el); });
  document.getElementById("pnlRptMonth")?.addEventListener("change", e => { pnlRptMonth = +e.target.value; renderPnlReport(el); });
  document.getElementById("pnlPrintBtn")?.addEventListener("click", () => window.print());

  // 문서번호 인라인 편집
  document.getElementById("pnlDocNoVal")?.addEventListener("blur", e => {
    if (!entry) return;
    const val = e.target.textContent.trim();
    if (val === "—") return;
    entry.docNo = val;
    upsertPnlEntry(entry);
  });

  // 대표이사 의견 저장
  document.getElementById("pnlCeoComment")?.addEventListener("change", e => {
    if (!entry) return;
    entry.ceoComment = e.target.value.trim();
    upsertPnlEntry(entry);
  });

  // 서명 버튼
  el.querySelectorAll(".pnl-sign-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      const stepIdx = +btn.dataset.step;
      const step = PNL_APPROVAL_STEPS[stepIdx];
      if (!entry) return;
      entry[step.dateKey]  = _todayKor();
      entry.approvalStatus = step.nextStatus;
      if (stepIdx === 0 && (!entry.docNo || entry.docNo === "—")) {
        entry.docNo = `MA-PNL-${entry.year}${String(entry.month).padStart(2,"0")}-001`;
      }
      upsertPnlEntry(entry);
      writePnlPendingToFirebase();
      pnlToast(`${step.role} 서명 완료`);
      renderPnlReport(el);
    });
  });

  // 취소(revoke) 버튼
  el.querySelectorAll(".pnl-revoke-btn").forEach(btn => {
    btn.addEventListener("click", e => {
      e.stopPropagation();
      if (!confirm("결재를 취소하시겠습니까?")) return;
      const stepIdx = +btn.dataset.step;
      if (!entry) return;
      for (let i = stepIdx; i < PNL_APPROVAL_STEPS.length; i++) {
        entry[PNL_APPROVAL_STEPS[i].dateKey] = "";
      }
      entry.approvalStatus = stepIdx > 0 ? PNL_APPROVAL_STEPS[stepIdx-1].nextStatus : "draft";
      if (stepIdx === 0) entry.docNo = "";
      upsertPnlEntry(entry);
      writePnlPendingToFirebase();
      renderPnlReport(el);
    });
  });

  // 월간/분기 모드 토글
  el.querySelectorAll("[data-rpt-mode]").forEach(btn => {
    btn.addEventListener("click", () => { pnlRptMode = btn.dataset.rptMode; renderPnlReport(el); });
  });
}

// ── 분기 집계 헬퍼 ────────────────────────────────────────────
function _quarterMonths(q) {
  return [[1,2,3],[4,5,6],[7,8,9],[10,11,12]][q - 1];
}

function _aggregateMonths(year, months) {
  const rows = months.map(m => getPnlEntry(year, m)).filter(Boolean);
  if (!rows.length) return null;
  const sum = rows.reduce((acc, d) => {
    acc.revenue       += d.revenue       || 0;
    acc.targetRevenue += d.targetRevenue || 0;
    acc.cogs          += d.cogs          || 0;
    acc.mfg           += d.mfg          || 0;
    acc.sga           += d.sga          || 0;
    acc.interest      += d.interest      || 0;
    acc.purchaseAmount = (acc.purchaseAmount || 0) + (d.purchaseAmount || 0);
    return acc;
  }, { revenue:0, targetRevenue:0, cogs:0, mfg:0, sga:0, interest:0 });
  // 재고 수동입력 있으면 분기 기초(첫 월) / 기말(마지막 월) 적용
  const withInv = rows.filter(r => r.beginInventory !== undefined && r.endInventory !== undefined);
  if (withInv.length === rows.length) {
    sum.beginInventory = rows[0].beginInventory || 0;
    sum.endInventory   = rows[rows.length - 1].endInventory || 0;
  }
  return sum;
}

// ── 분기 보고서 ───────────────────────────────────────────────
function renderPnlQuarterlyReport(el) {
  const curY = new Date().getFullYear();
  const maxYear = Math.max(curY + 1, pnlRptYear);
  const yearOpts = Array.from({length: maxYear - 2023}, (_,i) => 2024+i).map(y =>
    `<option value="${y}" ${y===pnlRptYear?"selected":""}>${y}년</option>`).join("");
  const qOpts = [1,2,3,4].map(q =>
    `<option value="${q}" ${q===pnlRptQuarter?"selected":""}>${q}분기</option>`).join("");

  const months  = _quarterMonths(pnlRptQuarter);
  const entry   = _aggregateMonths(pnlRptYear, months);
  const c       = entry ? calcPnl(entry) : null;

  // 분기 결재 상태
  const qKey = `${pnlRptYear}_Q${pnlRptQuarter}`;
  const qApproval = pnlQuarterApproval[qKey] || { approvalStatus:"draft", draftDate:"", agree1Date:"", agree2Date:"", ceoDate:"", docNo:"" };
  const qStatusIdx = _pnlStatusIdx(qApproval.approvalStatus);

  // 기존 서명 데이터가 있으면 세션 중 1회 구글시트 자동 동기화
  if (qApproval.approvalStatus !== "draft" && !_pnlQtrSyncedKeys.has(qKey)) {
    _pnlQtrSyncedKeys.add(qKey);
    _saveQtrToSheets(pnlRptYear, pnlRptQuarter, qKey, qApproval);
  }

  function approvalBoxQ(stepIdx) {
    const step = PNL_APPROVAL_STEPS[stepIdx];
    const done    = qStatusIdx > stepIdx;
    const current = qStatusIdx === stepIdx && !!entry;
    const date    = qApproval[step.dateKey] || "";
    return `
      <div class="pnl-ap-box">
        <div class="pnl-ap-role">${step.role}</div>
        <div class="pnl-ap-name">${step.name} ${step.title}</div>
        <div class="pnl-ap-date">${done && date ? date : "&nbsp;"}</div>
        ${done
          ? `<div class="pnl-stamp ${step.stampCls}">서명
               <button class="pnl-revoke-btn" data-step="${stepIdx}" title="결재 취소">↩</button>
             </div>`
          : current
            ? `<button class="pnl-sign-btn" data-step="${stepIdx}">서명<br>하기</button>`
            : `<div class="pnl-stamp pnl-stamp-empty"></div>`
        }
      </div>`;
  }

  const prevQ   = pnlRptQuarter > 1 ? pnlRptQuarter - 1 : 4;
  const prevQY  = pnlRptQuarter > 1 ? pnlRptYear : pnlRptYear - 1;
  const prev    = _aggregateMonths(prevQY, _quarterMonths(prevQ));
  const pc      = prev ? calcPnl(prev) : null;

  function cmpRow(label, pv, cv, isSub) {
    if (!pc || !c) return "";
    const diff = cv - pv;
    const diffCls = diff > 0 ? "pnl-pos" : diff < 0 ? "pnl-neg" : "";
    const trCls = label === "경영이익(손실)" ? "pnl-tr-total" : isSub ? "pnl-tr-sub" : "";
    return `<tr class="${trCls}">
      <td>${label}</td>
      <td>${_ps(pv)}</td>
      <td>${_ps(cv)}</td>
      <td class="${diffCls}">${diff >= 0 ? "▲ " : "▼ "}${_pf(Math.abs(diff))}</td>
    </tr>`;
  }

  const noDataHtml = `<div class="pnl-no-data">이 분기의 데이터가 없습니다. 입력 탭에서 월별 데이터를 먼저 저장해 주세요.</div>`;

  el.innerHTML = `
    <div class="pnl-report-wrap">
      <div class="pnl-report-toolbar no-print">
        <div class="pnl-rpt-mode-tabs">
          <button class="pnl-mode-btn" data-rpt-mode="monthly">월간</button>
          <button class="pnl-mode-btn active" data-rpt-mode="quarterly">분기</button>
          <button class="pnl-mode-btn" data-rpt-mode="halfyear">반기</button>
          <button class="pnl-mode-btn" data-rpt-mode="annual">연간</button>
        </div>
        <button class="pnl-nav-btn" id="pnlRptPrev">◀</button>
        <select id="pnlRptYear">${yearOpts}</select>
        <select id="pnlRptQtr">${qOpts}</select>
        <button class="pnl-nav-btn" id="pnlRptNext">▶</button>
        <button class="pnl-btn pnl-btn-print" id="pnlPrintBtn">인쇄 / PDF</button>
      </div>

      <div class="pnl-page" id="pnlReportPage">
        <div class="pnl-doc-header">
          <div class="pnl-company-badge">MIRAE AUTOMATION CO., LTD</div>
          <div class="pnl-doc-title">${pnlRptYear}년 Q${pnlRptQuarter} &mdash; 분기 경영손익 보고서</div>
          <div class="pnl-doc-sub">${months[0]}월 ~ ${months[months.length-1]}월 합산 관리기준 손익</div>
          <div class="pnl-meta-row">
            <div class="pnl-meta-item"><span class="pnl-meta-lbl">기안부서</span><span class="pnl-meta-val">${PNL_META.department}</span></div>
            <div class="pnl-meta-item"><span class="pnl-meta-lbl">기안자</span><span class="pnl-meta-val">${PNL_META.author.name} ${PNL_META.author.title}</span></div>
            <div class="pnl-meta-item"><span class="pnl-meta-lbl">기안일</span><span class="pnl-meta-val">${qApproval.draftDate||"—"}</span></div>
            <div class="pnl-meta-item"><span class="pnl-meta-lbl">문서번호</span>
              <span class="pnl-meta-val" id="pnlQDocNoVal" contenteditable="${!!entry}" style="outline:none;cursor:${entry?"text":"default"}">${qApproval.docNo||"—"}</span>
            </div>
          </div>
        </div>

        <div class="pnl-doc-body">
          ${!entry ? noDataHtml : `
          <div class="pnl-kpi-grid">
            <div class="pnl-kpi-card">
              <div class="pnl-kpi-lbl">매출액 (Q${pnlRptQuarter})</div>
              <div class="pnl-kpi-val">${_pf(entry.revenue)}</div>
              <div class="pnl-kpi-unit">원${entry.targetRevenue>0?" / 목표 "+_pf(entry.targetRevenue)+"원":""}</div>
              ${c.targetAchieve!==null?`<div class="pnl-kpi-achieve ${c.targetAchieve>=100?"pnl-pos":"pnl-neg"}">달성률 ${c.targetAchieve.toFixed(1)}%</div>`:""}
            </div>
            <div class="pnl-kpi-card">
              <div class="pnl-kpi-lbl">매출총이익</div>
              <div class="pnl-kpi-val ${_pc(c.gross)}">${_ps(c.gross)}</div>
              <div class="pnl-kpi-unit">원 / 총이익률 ${c.gmRate.toFixed(1)}%</div>
            </div>
            <div class="pnl-kpi-card pnl-kpi-highlight">
              <div class="pnl-kpi-lbl">경영이익(손실)</div>
              <div class="pnl-kpi-val ${_pc(c.mgmt)}">${_ps(c.mgmt)}</div>
              <div class="pnl-kpi-unit">금융비용 반영 기준</div>
            </div>
            <div class="pnl-kpi-card">
              <div class="pnl-kpi-lbl">영업이익률</div>
              <div class="pnl-kpi-val ${_pc(c.op)}">${c.opRate.toFixed(1)}%</div>
              <div class="pnl-kpi-unit">영업이익 ${_ps(c.op)}</div>
            </div>
          </div>

          <div class="pnl-section">
            <div class="pnl-sec-title"><span class="pnl-sec-num">1</span>관리기준 영업이익 <small>실질 기준</small></div>
            <div class="pnl-flow">
              <div class="pnl-flow-row"><span class="pnl-flow-lbl"><span class="pnl-tag">ㄱ</span> 매출액</span><span class="pnl-flow-val">${_pf(entry.revenue)} 원</span></div>
              <div class="pnl-flow-divider">차감</div>
              <div class="pnl-flow-row pnl-indent"><span class="pnl-flow-lbl"><span class="pnl-minus">−</span><span class="pnl-tag">ㄴ</span> 상품매출원가</span><span class="pnl-flow-val pnl-neg">(${_pf(c.cogs)}) 원</span></div>
              <div class="pnl-flow-row pnl-indent"><span class="pnl-flow-lbl"><span class="pnl-minus">−</span><span class="pnl-tag">ㄷ</span> 당기총제조비용</span><span class="pnl-flow-val pnl-neg">(${_pf(entry.mfg)}) 원</span></div>
              <div class="pnl-flow-row pnl-flow-sub"><span class="pnl-flow-lbl">① 매출총이익 <small>[ㄱ−(ㄴ+ㄷ)]</small></span><span class="pnl-flow-val ${_pc(c.gross)}">${_ps(c.gross)} 원</span></div>
              <div class="pnl-flow-divider">차감</div>
              <div class="pnl-flow-row pnl-indent"><span class="pnl-flow-lbl"><span class="pnl-minus">−</span> 판매관리비</span><span class="pnl-flow-val pnl-neg">(${_pf(entry.sga)}) 원</span></div>
              <div class="pnl-flow-row pnl-flow-total"><span class="pnl-flow-lbl">② 관리기준 영업이익 <small>[①−판관비]</small></span><span class="pnl-flow-val ${_pc(c.op)}">${_ps(c.op)} 원</span></div>
            </div>
          </div>

          <div class="pnl-section">
            <div class="pnl-sec-title"><span class="pnl-sec-num">2</span>경영이익 <small>금융비용 반영 기준</small></div>
            <div class="pnl-flow">
              <div class="pnl-flow-row"><span class="pnl-flow-lbl">관리기준 영업이익 (②)</span><span class="pnl-flow-val ${_pc(c.op)}">${_ps(c.op)} 원</span></div>
              <div class="pnl-flow-divider">차감</div>
              <div class="pnl-flow-row pnl-indent"><span class="pnl-flow-lbl"><span class="pnl-minus">−</span> 영업외비용</span><span class="pnl-flow-val pnl-neg">(${_pf(entry.interest)}) 원</span></div>
              <div class="pnl-flow-row pnl-flow-total"><span class="pnl-flow-lbl">③ 경영이익(손실) <small>[②−영업외비용]</small></span><span class="pnl-flow-val ${_pc(c.mgmt)}">${_ps(c.mgmt)} 원</span></div>
            </div>
            <div class="pnl-remark">※ 비고: 영업외비용(이자비용 등) 반영 시 실제 경영성과를 함께 확인할 수 있도록 별도 표시하였습니다.</div>
          </div>

          <!-- 전분기 비교 -->
          ${prev && pc ? `
          <div class="pnl-section">
            <div class="pnl-sec-title"><span class="pnl-sec-num">3</span>전분기 대비 손익 비교 <small>${prevQY}년 Q${prevQ} 대비</small></div>
            <table class="pnl-cmp-table">
              <thead><tr><th>항목</th><th>${prevQY}년 Q${prevQ}</th><th>${pnlRptYear}년 Q${pnlRptQuarter}</th><th>증감액</th></tr></thead>
              <tbody>
                ${cmpRow("매출액",          prev.revenue,       entry.revenue,      false)}
                ${cmpRow("상품매출원가",    pc.cogs,            c.cogs,             false)}
                ${cmpRow("당기총제조비용",  prev.mfg,           entry.mfg,          false)}
                ${cmpRow("매출총이익",      pc.gross,           c.gross,            true)}
                ${cmpRow("판관비",          prev.sga,           entry.sga,          false)}
                ${cmpRow("관리기준 영업이익", pc.op,            c.op,               true)}
                ${cmpRow("영업외비용",      prev.interest,      entry.interest,     false)}
                ${cmpRow("경영이익(손실)",  pc.mgmt,            c.mgmt,             true)}
              </tbody>
            </table>
          </div>` : ""}

          <!-- 월별 내역 -->
          <div class="pnl-section">
            <div class="pnl-sec-title"><span class="pnl-sec-num">4</span>월별 내역</div>
            <table class="pnl-cmp-table">
              <thead><tr><th>월</th><th>매출액</th><th>매출총이익</th><th>영업이익</th><th>경영이익</th></tr></thead>
              <tbody>
                ${months.map(m => {
                  const me = getPnlEntry(pnlRptYear, m);
                  const mc = me ? calcPnl(me) : null;
                  return `<tr>
                    <td>${m}월</td>
                    <td>${me ? _pf(me.revenue) : "—"}</td>
                    <td class="${mc?_pc(mc.gross):""}">${mc ? _ps(mc.gross) : "—"}</td>
                    <td class="${mc?_pc(mc.op):""}">${mc ? _ps(mc.op) : "—"}</td>
                    <td class="${mc?_pc(mc.mgmt):""}">${mc ? _ps(mc.mgmt) : "—"}</td>
                  </tr>`;
                }).join("")}
              </tbody>
            </table>
          </div>

          <!-- 결재란 -->
          <div class="pnl-section">
            <div class="pnl-sec-title"><span class="pnl-sec-num" style="font-size:11px">✓</span>결재</div>
            <div class="pnl-ap-grid">
              ${PNL_APPROVAL_STEPS.map((_,i) => approvalBoxQ(i)).join("")}
            </div>
          </div>
          `}
        </div>
        <div class="pnl-doc-footer">${PNL_META.companyName} · ${PNL_META.department} · 대외비</div>
      </div>
    </div>`;

  el.querySelectorAll("[data-rpt-mode]").forEach(btn => {
    btn.addEventListener("click", () => { pnlRptMode = btn.dataset.rptMode; renderPnlReport(el); });
  });
  document.getElementById("pnlRptPrev")?.addEventListener("click", () => {
    pnlRptQuarter--; if (pnlRptQuarter < 1) { pnlRptQuarter = 4; pnlRptYear--; }
    renderPnlReport(el);
  });
  document.getElementById("pnlRptNext")?.addEventListener("click", () => {
    pnlRptQuarter++; if (pnlRptQuarter > 4) { pnlRptQuarter = 1; pnlRptYear++; }
    renderPnlReport(el);
  });
  document.getElementById("pnlRptYear")?.addEventListener("change", e => { pnlRptYear = +e.target.value; renderPnlReport(el); });
  document.getElementById("pnlRptQtr")?.addEventListener("change",  e => { pnlRptQuarter = +e.target.value; renderPnlReport(el); });
  document.getElementById("pnlPrintBtn")?.addEventListener("click", () => window.print());

  // 분기 문서번호 인라인 편집
  document.getElementById("pnlQDocNoVal")?.addEventListener("blur", e => {
    if (!entry) return;
    const val = e.target.textContent.trim();
    if (val === "—") return;
    qApproval.docNo = val;
    pnlQuarterApproval[qKey] = qApproval;
    savePnlQuarterApprovalLocal();
  });

  // 분기 서명 버튼
  el.querySelectorAll(".pnl-sign-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      const stepIdx = +btn.dataset.step;
      const step = PNL_APPROVAL_STEPS[stepIdx];
      if (!entry) return;
      qApproval[step.dateKey] = _todayKor();
      qApproval.approvalStatus = step.nextStatus;
      if (stepIdx === 0 && (!qApproval.docNo || qApproval.docNo === "—")) {
        qApproval.docNo = `MA-PNL-${pnlRptYear}Q${pnlRptQuarter}-001`;
      }
      pnlQuarterApproval[qKey] = qApproval;
      savePnlQuarterApprovalLocal();
      _saveQtrToSheets(pnlRptYear, pnlRptQuarter, qKey, qApproval);
      writePnlPendingToFirebase();
      pnlToast(`${step.role} 서명 완료`);
      renderPnlReport(el);
    });
  });

  // 분기 취소(revoke) 버튼
  el.querySelectorAll(".pnl-revoke-btn").forEach(btn => {
    btn.addEventListener("click", e => {
      e.stopPropagation();
      if (!confirm("결재를 취소하시겠습니까?")) return;
      const stepIdx = +btn.dataset.step;
      for (let i = stepIdx; i < PNL_APPROVAL_STEPS.length; i++) {
        qApproval[PNL_APPROVAL_STEPS[i].dateKey] = "";
      }
      qApproval.approvalStatus = stepIdx > 0 ? PNL_APPROVAL_STEPS[stepIdx-1].nextStatus : "draft";
      if (stepIdx === 0) qApproval.docNo = "";
      pnlQuarterApproval[qKey] = qApproval;
      savePnlQuarterApprovalLocal();
      _saveQtrToSheets(pnlRptYear, pnlRptQuarter, qKey, qApproval);
      writePnlPendingToFirebase();
      renderPnlReport(el);
    });
  });
}

// ── 반기 / 연간 공통 흐름표 HTML 헬퍼 ────────────────────────
function _pnlFlowHtml(entry, c) {
  return `<div class="pnl-flow">
    <div class="pnl-flow-row"><span class="pnl-flow-lbl"><span class="pnl-tag">ㄱ</span> 매출액</span><span class="pnl-flow-val">${_pf(entry.revenue)} 원</span></div>
    <div class="pnl-flow-divider">차감</div>
    <div class="pnl-flow-row pnl-indent"><span class="pnl-flow-lbl"><span class="pnl-minus">−</span><span class="pnl-tag">ㄴ</span> 상품매출원가</span><span class="pnl-flow-val pnl-neg">(${_pf(c.cogs)}) 원</span></div>
    <div class="pnl-flow-row pnl-indent"><span class="pnl-flow-lbl"><span class="pnl-minus">−</span><span class="pnl-tag">ㄷ</span> 당기총제조비용</span><span class="pnl-flow-val pnl-neg">(${_pf(entry.mfg)}) 원</span></div>
    <div class="pnl-flow-row pnl-flow-sub"><span class="pnl-flow-lbl">① 매출총이익 <small>[ㄱ−(ㄴ+ㄷ)]</small></span><span class="pnl-flow-val ${_pc(c.gross)}">${_ps(c.gross)} 원</span></div>
    <div class="pnl-flow-divider">차감</div>
    <div class="pnl-flow-row pnl-indent"><span class="pnl-flow-lbl"><span class="pnl-minus">−</span> 판매관리비</span><span class="pnl-flow-val pnl-neg">(${_pf(entry.sga)}) 원</span></div>
    <div class="pnl-flow-row pnl-flow-total"><span class="pnl-flow-lbl">② 관리기준 영업이익 <small>[①−판관비]</small></span><span class="pnl-flow-val ${_pc(c.op)}">${_ps(c.op)} 원</span></div>
  </div>`;
}
function _pnlMgmtFlowHtml(entry, c) {
  return `<div class="pnl-flow">
    <div class="pnl-flow-row"><span class="pnl-flow-lbl">관리기준 영업이익 (②)</span><span class="pnl-flow-val ${_pc(c.op)}">${_ps(c.op)} 원</span></div>
    <div class="pnl-flow-divider">차감</div>
    <div class="pnl-flow-row pnl-indent"><span class="pnl-flow-lbl"><span class="pnl-minus">−</span> 영업외비용</span><span class="pnl-flow-val pnl-neg">(${_pf(entry.interest)}) 원</span></div>
    <div class="pnl-flow-row pnl-flow-total"><span class="pnl-flow-lbl">③ 경영이익(손실) <small>[②−영업외비용]</small></span><span class="pnl-flow-val ${_pc(c.mgmt)}">${_ps(c.mgmt)} 원</span></div>
  </div>`;
}

// ── 반기 / 연간 보고서 ────────────────────────────────────────
function renderPnlHalfYearReport(el) {
  const curY = new Date().getFullYear();
  const maxYear = Math.max(curY + 1, pnlRptYear);
  const yearOpts = Array.from({length: maxYear - 2023}, (_,i) => 2024+i)
    .map(y => `<option value="${y}" ${y===pnlRptYear?"selected":""}>${y}년</option>`).join("");
  const halfOpts = [1,2].map(h =>
    `<option value="${h}" ${h===pnlRptHalf?"selected":""}>${h===1?"상반기":"하반기"}</option>`).join("");

  const months  = pnlRptHalf === 1 ? [1,2,3,4,5,6] : [7,8,9,10,11,12];
  const entry   = _aggregateMonths(pnlRptYear, months);
  const c       = entry ? calcPnl(entry) : null;
  const prevH   = _aggregateMonths(pnlRptYear - 1, months);
  const pc      = prevH ? calcPnl(prevH) : null;
  const halfLabel = pnlRptHalf === 1 ? "상반기" : "하반기";

  function cmpRow(label, pv, cv, isSub) {
    const diff = cv - pv;
    const diffCls = diff > 0 ? "pnl-pos" : diff < 0 ? "pnl-neg" : "";
    const rowCls = label.includes("경영이익") ? "pnl-cmp-total" : isSub ? "pnl-cmp-sub" : "";
    return `<tr class="${rowCls}">
      <td>${label}</td><td>${_ps(pv)}</td><td>${_ps(cv)}</td>
      <td class="${diffCls}">${diff >= 0 ? "▲ " : "▼ "}${_pf(Math.abs(diff))}</td>
    </tr>`;
  }

  el.innerHTML = `
    <div class="pnl-report-wrap">
      <div class="pnl-report-toolbar no-print">
        <div class="pnl-rpt-mode-tabs">
          <button class="pnl-mode-btn" data-rpt-mode="monthly">월간</button>
          <button class="pnl-mode-btn" data-rpt-mode="quarterly">분기</button>
          <button class="pnl-mode-btn active" data-rpt-mode="halfyear">반기</button>
          <button class="pnl-mode-btn" data-rpt-mode="annual">연간</button>
        </div>
        <button class="pnl-nav-btn" id="pnlRptPrev">◀</button>
        <select id="pnlRptYear">${yearOpts}</select>
        <select id="pnlRptHalf">${halfOpts}</select>
        <button class="pnl-nav-btn" id="pnlRptNext">▶</button>
        <button class="pnl-btn pnl-btn-print" id="pnlPrintBtn">인쇄 / PDF</button>
      </div>
      <div class="pnl-page" id="pnlReportPage">
        <div class="pnl-doc-header">
          <div class="pnl-company-badge">MIRAE AUTOMATION CO., LTD</div>
          <div class="pnl-doc-title">${pnlRptYear}년 ${halfLabel} &mdash; 반기 경영손익 보고서</div>
          <div class="pnl-doc-sub">${months[0]}월 ~ ${months[months.length-1]}월 합산 관리기준 손익</div>
        </div>
        <div class="pnl-doc-body">
        ${!entry ? `<div class="pnl-no-data">이 반기의 데이터가 없습니다. 월별 데이터를 먼저 저장해 주세요.</div>` : `
        <div class="pnl-kpi-grid">
          <div class="pnl-kpi-card">
            <div class="pnl-kpi-lbl">매출액</div>
            <div class="pnl-kpi-val">${_pf(entry.revenue)}</div>
            <div class="pnl-kpi-unit">원</div>
            ${c.targetAchieve !== null ? `<div class="pnl-kpi-achieve ${c.targetAchieve>=100?"pnl-pos":"pnl-neg"}">달성률 ${c.targetAchieve.toFixed(1)}%</div>` : ""}
          </div>
          <div class="pnl-kpi-card">
            <div class="pnl-kpi-lbl">매출총이익</div>
            <div class="pnl-kpi-val ${_pc(c.gross)}">${_ps(c.gross)}</div>
            <div class="pnl-kpi-unit">원 / 총이익률 ${c.gmRate.toFixed(1)}%</div>
          </div>
          <div class="pnl-kpi-card pnl-kpi-highlight">
            <div class="pnl-kpi-lbl">경영이익(손실)</div>
            <div class="pnl-kpi-val ${_pc(c.mgmt)}">${_ps(c.mgmt)}</div>
            <div class="pnl-kpi-unit">영업이익률 ${c.opRate.toFixed(1)}%</div>
          </div>
        </div>
        <div class="pnl-section">
          <div class="pnl-sec-title"><span class="pnl-sec-num">1</span>관리기준 영업이익</div>
          ${_pnlFlowHtml(entry, c)}
        </div>
        <div class="pnl-section">
          <div class="pnl-sec-title"><span class="pnl-sec-num">2</span>경영이익 <small>금융비용 반영</small></div>
          ${_pnlMgmtFlowHtml(entry, c)}
        </div>
        ${prevH && pc ? `
        <div class="pnl-section">
          <div class="pnl-sec-title"><span class="pnl-sec-num">3</span>전년 동기 대비 <small>${pnlRptYear-1}년 ${halfLabel}</small></div>
          <table class="pnl-cmp-table">
            <thead><tr><th>항목</th><th>${pnlRptYear-1}년 ${halfLabel}</th><th>${pnlRptYear}년 ${halfLabel}</th><th>증감액</th></tr></thead>
            <tbody>
              ${cmpRow("매출액",            prevH.revenue,  entry.revenue,  false)}
              ${cmpRow("상품매출원가",      prevH.cogs,     entry.cogs,     false)}
              ${cmpRow("당기총제조비용",    prevH.mfg,      entry.mfg,      false)}
              ${cmpRow("매출총이익",        pc.gross,       c.gross,        true)}
              ${cmpRow("판관비",            prevH.sga,      entry.sga,      false)}
              ${cmpRow("관리기준 영업이익", pc.op,          c.op,           true)}
              ${cmpRow("영업외비용",         prevH.interest, entry.interest, false)}
              ${cmpRow("경영이익(손실)",    pc.mgmt,        c.mgmt,         true)}
            </tbody>
          </table>
        </div>` : ""}
        `}
        </div><!-- /pnl-doc-body -->
        <div class="pnl-doc-footer">${PNL_META.companyName} · ${PNL_META.department} · 대외비</div>
      </div>
    </div>`;

  el.querySelectorAll("[data-rpt-mode]").forEach(btn => {
    btn.addEventListener("click", () => { pnlRptMode = btn.dataset.rptMode; renderPnlReport(el); });
  });
  document.getElementById("pnlRptPrev")?.addEventListener("click", () => {
    pnlRptHalf--; if (pnlRptHalf < 1) { pnlRptHalf = 2; pnlRptYear--; } renderPnlReport(el);
  });
  document.getElementById("pnlRptNext")?.addEventListener("click", () => {
    pnlRptHalf++; if (pnlRptHalf > 2) { pnlRptHalf = 1; pnlRptYear++; } renderPnlReport(el);
  });
  document.getElementById("pnlRptYear")?.addEventListener("change", e => { pnlRptYear = +e.target.value; renderPnlReport(el); });
  document.getElementById("pnlRptHalf")?.addEventListener("change", e => { pnlRptHalf = +e.target.value; renderPnlReport(el); });
  document.getElementById("pnlPrintBtn")?.addEventListener("click", () => window.print());
}

function renderPnlAnnualReport(el) {
  const curY = new Date().getFullYear();
  const maxYear = Math.max(curY + 1, pnlRptYear);
  const yearOpts = Array.from({length: maxYear - 2023}, (_,i) => 2024+i)
    .map(y => `<option value="${y}" ${y===pnlRptYear?"selected":""}>${y}년</option>`).join("");

  const months  = [1,2,3,4,5,6,7,8,9,10,11,12];
  const entry   = _aggregateMonths(pnlRptYear, months);
  const c       = entry ? calcPnl(entry) : null;
  const prevY   = _aggregateMonths(pnlRptYear - 1, months);
  const pc      = prevY ? calcPnl(prevY) : null;

  function cmpRow(label, pv, cv, isSub) {
    const diff = cv - pv;
    const diffCls = diff > 0 ? "pnl-pos" : diff < 0 ? "pnl-neg" : "";
    const rowCls = label.includes("경영이익") ? "pnl-cmp-total" : isSub ? "pnl-cmp-sub" : "";
    return `<tr class="${rowCls}">
      <td>${label}</td><td>${_ps(pv)}</td><td>${_ps(cv)}</td>
      <td class="${diffCls}">${diff >= 0 ? "▲ " : "▼ "}${_pf(Math.abs(diff))}</td>
    </tr>`;
  }

  el.innerHTML = `
    <div class="pnl-report-wrap">
      <div class="pnl-report-toolbar no-print">
        <div class="pnl-rpt-mode-tabs">
          <button class="pnl-mode-btn" data-rpt-mode="monthly">월간</button>
          <button class="pnl-mode-btn" data-rpt-mode="quarterly">분기</button>
          <button class="pnl-mode-btn" data-rpt-mode="halfyear">반기</button>
          <button class="pnl-mode-btn active" data-rpt-mode="annual">연간</button>
        </div>
        <button class="pnl-nav-btn" id="pnlRptPrev">◀</button>
        <select id="pnlRptYear">${yearOpts}</select>
        <button class="pnl-nav-btn" id="pnlRptNext">▶</button>
        <button class="pnl-btn pnl-btn-print" id="pnlPrintBtn">인쇄 / PDF</button>
      </div>
      <div class="pnl-page" id="pnlReportPage">
        <div class="pnl-doc-header">
          <div class="pnl-company-badge">MIRAE AUTOMATION CO., LTD</div>
          <div class="pnl-doc-title">${pnlRptYear}년 &mdash; 연간 경영손익 보고서</div>
          <div class="pnl-doc-sub">1월 ~ 12월 합산 관리기준 손익</div>
        </div>
        <div class="pnl-doc-body">
        ${!entry ? `<div class="pnl-no-data">이 연도의 데이터가 없습니다. 월별 데이터를 먼저 저장해 주세요.</div>` : `
        <div class="pnl-kpi-grid">
          <div class="pnl-kpi-card">
            <div class="pnl-kpi-lbl">매출액</div>
            <div class="pnl-kpi-val">${_pf(entry.revenue)}</div>
            <div class="pnl-kpi-unit">원</div>
            ${c.targetAchieve !== null ? `<div class="pnl-kpi-achieve ${c.targetAchieve>=100?"pnl-pos":"pnl-neg"}">달성률 ${c.targetAchieve.toFixed(1)}%</div>` : ""}
          </div>
          <div class="pnl-kpi-card">
            <div class="pnl-kpi-lbl">매출총이익</div>
            <div class="pnl-kpi-val ${_pc(c.gross)}">${_ps(c.gross)}</div>
            <div class="pnl-kpi-unit">원 / 총이익률 ${c.gmRate.toFixed(1)}%</div>
          </div>
          <div class="pnl-kpi-card pnl-kpi-highlight">
            <div class="pnl-kpi-lbl">경영이익(손실)</div>
            <div class="pnl-kpi-val ${_pc(c.mgmt)}">${_ps(c.mgmt)}</div>
            <div class="pnl-kpi-unit">영업이익률 ${c.opRate.toFixed(1)}%</div>
          </div>
        </div>
        <div class="pnl-section">
          <div class="pnl-sec-title"><span class="pnl-sec-num">1</span>관리기준 영업이익</div>
          ${_pnlFlowHtml(entry, c)}
        </div>
        <div class="pnl-section">
          <div class="pnl-sec-title"><span class="pnl-sec-num">2</span>경영이익 <small>금융비용 반영</small></div>
          ${_pnlMgmtFlowHtml(entry, c)}
        </div>
        ${prevY && pc ? `
        <div class="pnl-section">
          <div class="pnl-sec-title"><span class="pnl-sec-num">3</span>전년 대비 <small>${pnlRptYear-1}년 연간</small></div>
          <table class="pnl-cmp-table">
            <thead><tr><th>항목</th><th>${pnlRptYear-1}년 연간</th><th>${pnlRptYear}년 연간</th><th>증감액</th></tr></thead>
            <tbody>
              ${cmpRow("매출액",            prevY.revenue,  entry.revenue,  false)}
              ${cmpRow("상품매출원가",      prevY.cogs,     entry.cogs,     false)}
              ${cmpRow("당기총제조비용",    prevY.mfg,      entry.mfg,      false)}
              ${cmpRow("매출총이익",        pc.gross,       c.gross,        true)}
              ${cmpRow("판관비",            prevY.sga,      entry.sga,      false)}
              ${cmpRow("관리기준 영업이익", pc.op,          c.op,           true)}
              ${cmpRow("영업외비용",         prevY.interest, entry.interest, false)}
              ${cmpRow("경영이익(손실)",    pc.mgmt,        c.mgmt,         true)}
            </tbody>
          </table>
        </div>` : ""}
        `}
        </div><!-- /pnl-doc-body -->
        <div class="pnl-doc-footer">${PNL_META.companyName} · ${PNL_META.department} · 대외비</div>
      </div>
    </div>`;

  el.querySelectorAll("[data-rpt-mode]").forEach(btn => {
    btn.addEventListener("click", () => { pnlRptMode = btn.dataset.rptMode; renderPnlReport(el); });
  });
  document.getElementById("pnlRptPrev")?.addEventListener("click", () => { pnlRptYear--; renderPnlReport(el); });
  document.getElementById("pnlRptNext")?.addEventListener("click", () => { pnlRptYear++; renderPnlReport(el); });
  document.getElementById("pnlRptYear")?.addEventListener("change", e => { pnlRptYear = +e.target.value; renderPnlReport(el); });
  document.getElementById("pnlPrintBtn")?.addEventListener("click", () => window.print());
}

// ── 대시보드 탭 ───────────────────────────────────────────────
async function _loadChartJs() {
  if (window.Chart) return;
  await new Promise((res, rej) => {
    const s = document.createElement("script");
    s.src = "https://cdn.jsdelivr.net/npm/chart.js@4.4.2/dist/chart.umd.min.js";
    s.onload = res; s.onerror = rej;
    document.head.appendChild(s);
  });
}

function _pnlAggregate(year, period) {
  const rows = pnlData.filter(d => d.year === year);
  if (period === "monthly") {
    return rows.map(d => {
      const c = calcPnl(d);
      return { label: `${d.month}월`, ...d, ...c };
    });
  }
  const groups = period === "quarterly"
    ? [[1,2,3],[4,5,6],[7,8,9],[10,11,12]]
    : period === "halfyear"
      ? [[1,2,3,4,5,6],[7,8,9,10,11,12]]
      : [[1,2,3,4,5,6,7,8,9,10,11,12]];
  const labels = period === "quarterly" ? ["Q1","Q2","Q3","Q4"]
    : period === "halfyear" ? ["상반기","하반기"] : [`${year}년`];

  return groups.map((months, gi) => {
    const grpRows = rows.filter(d => months.includes(d.month));
    if (!grpRows.length) return null;
    const sum = grpRows.reduce((acc, d) => {
      acc.revenue += d.revenue; acc.targetRevenue += d.targetRevenue;
      acc.cogs += d.cogs; acc.mfg += d.mfg; acc.sga += d.sga; acc.interest += d.interest;
      return acc;
    }, { revenue:0, targetRevenue:0, cogs:0, mfg:0, sga:0, interest:0 });
    const c = calcPnl(sum);
    return { label: labels[gi], ...sum, ...c };
  }).filter(Boolean);
}

async function renderPnlDashboard(el) {
  const curY = new Date().getFullYear();
  const allYears = [...new Set(pnlData.map(d => d.year))].sort((a,b) => b - a);
  if (!allYears.includes(pnlDashYear) && allYears.length) pnlDashYear = allYears[0];
  const yearOpts = (allYears.length ? allYears : [curY]).map(y =>
    `<option value="${y}" ${y===pnlDashYear?"selected":""}>${y}년</option>`).join("");

  const agg = _pnlAggregate(pnlDashYear, pnlDashPeriod);
  const prevYearAgg = _pnlAggregate(pnlDashYear - 1, "monthly");

  // KPI 합산 (전체 or 선택 기간)
  const totals = agg.reduce((acc, d) => {
    acc.revenue += d.revenue; acc.targetRevenue += d.targetRevenue;
    acc.cogs += d.cogs; acc.mfg += d.mfg; acc.sga += d.sga; acc.interest += d.interest;
    return acc;
  }, { revenue:0, targetRevenue:0, cogs:0, mfg:0, sga:0, interest:0 });
  const tc = calcPnl(totals);

  el.innerHTML = `
    <div class="pnl-dash-wrap">
      <div class="pnl-dash-toolbar">
        <select id="pnlDashYear">${yearOpts}</select>
        <div class="pnl-period-tabs">
          ${["monthly","quarterly","halfyear","annual"].map((p,i) =>
            `<button class="pnl-period-btn${pnlDashPeriod===p?" active":""}" data-period="${p}">${["월별","분기","반기","연간"][i]}</button>`
          ).join("")}
        </div>
      </div>

      <!-- KPI 카드 -->
      <div class="pnl-dash-kpis">
        ${[
          { label:"매출액",    val:totals.revenue,      sub: totals.targetRevenue>0?`목표 달성률 ${tc.targetAchieve.toFixed(1)}%`:"",  cls:"" },
          { label:"매출총이익", val:tc.gross,             sub:`총이익률 ${tc.gmRate.toFixed(1)}%`,  cls:_pc(tc.gross) },
          { label:"영업이익",  val:tc.op,               sub:`영업이익률 ${tc.opRate.toFixed(1)}%`, cls:_pc(tc.op) },
          { label:"경영이익",  val:tc.mgmt,             sub:"영업외비용 반영",                     cls:_pc(tc.mgmt) },
          { label:"판관비 합계", val:totals.sga,          sub:"",                                  cls:"" },
          { label:"영업외비용 합계", val:totals.interest,  sub:"",                                  cls:"" },
        ].map(k => `
          <div class="pnl-dash-kpi">
            <div class="pnl-dash-kpi-lbl">${k.label}</div>
            <div class="pnl-dash-kpi-val ${k.cls}">${_ps(k.val)}</div>
            ${k.sub ? `<div class="pnl-dash-kpi-sub">${k.sub}</div>` : ""}
          </div>`).join("")}
      </div>

      <!-- 차트 영역 -->
      <div class="pnl-charts-grid">
        <div class="pnl-chart-box pnl-chart-wide">
          <div class="pnl-chart-title">매출액 · 원가 구성</div>
          <canvas id="pnlChart1"></canvas>
        </div>
        <div class="pnl-chart-box">
          <div class="pnl-chart-title">손익 추이</div>
          <canvas id="pnlChart2"></canvas>
        </div>
        <div class="pnl-chart-box">
          <div class="pnl-chart-title">비용 구조</div>
          <canvas id="pnlChart3"></canvas>
        </div>
      </div>

      <!-- 상세 테이블 -->
      <div class="pnl-dash-table-wrap">
        <table class="pnl-dash-table">
          <thead><tr><th>기간</th><th>매출액</th><th>목표</th><th>달성률</th><th>매출총이익</th><th>총이익률</th><th>영업이익</th><th>경영이익</th><th>판관비</th><th>영업외비용</th></tr></thead>
          <tbody>
            ${agg.map(d => `
              <tr class="${d.label.includes("Q")||d.label.includes("반기")||d.label.includes("년")?"pnl-tr-sub":""}">
                <td>${d.label}</td>
                <td>${_pf(d.revenue)}</td>
                <td>${d.targetRevenue>0?_pf(d.targetRevenue):"—"}</td>
                <td class="${d.targetAchieve!==null?(d.targetAchieve>=100?"pnl-pos":"pnl-neg"):""}">${d.targetAchieve!==null?d.targetAchieve.toFixed(1)+"%":"—"}</td>
                <td class="${_pc(d.gross)}">${_ps(d.gross)}</td>
                <td>${d.gmRate.toFixed(1)}%</td>
                <td class="${_pc(d.op)}">${_ps(d.op)}</td>
                <td class="${_pc(d.mgmt)}">${_ps(d.mgmt)}</td>
                <td>${_pf(d.sga)}</td>
                <td>${_pf(d.interest)}</td>
              </tr>`).join("")}
          </tbody>
        </table>
      </div>

      <!-- 전년 동월 비교 -->
      ${prevYearAgg.length && pnlDashPeriod==="monthly" ? `
      <div class="pnl-section" style="margin-top:24px">
        <div class="pnl-sec-title" style="font-size:14px;padding-bottom:8px;margin-bottom:12px">전년(${pnlDashYear-1}년) 동월 비교</div>
        <table class="pnl-dash-table">
          <thead><tr><th>월</th><th>${pnlDashYear-1}년 매출</th><th>${pnlDashYear}년 매출</th><th>매출 증감률</th><th>${pnlDashYear-1}년 경영이익</th><th>${pnlDashYear}년 경영이익</th><th>손익 변화</th></tr></thead>
          <tbody>
            ${Array.from({length:12},(_,i)=>i+1).map(mo => {
              const cur  = getPnlEntry(pnlDashYear,   mo);
              const prv  = getPnlEntry(pnlDashYear-1, mo);
              if (!cur && !prv) return "";
              const curC = cur ? calcPnl(cur) : null;
              const prvC = prv ? calcPnl(prv) : null;
              const revGrowth = prv && cur ? (cur.revenue - prv.revenue) / prv.revenue * 100 : null;
              return `<tr>
                <td>${mo}월</td>
                <td>${prv ? _pf(prv.revenue) : "—"}</td>
                <td>${cur ? _pf(cur.revenue) : "—"}</td>
                <td class="${revGrowth!==null?(revGrowth>=0?"pnl-pos":"pnl-neg"):""}">${revGrowth!==null?(revGrowth>=0?"▲":"▼")+Math.abs(revGrowth).toFixed(1)+"%":"—"}</td>
                <td class="${prvC?_pc(prvC.mgmt):""}">${prvC?_ps(prvC.mgmt):"—"}</td>
                <td class="${curC?_pc(curC.mgmt):""}">${curC?_ps(curC.mgmt):"—"}</td>
                <td class="${curC&&prvC?(curC.mgmt>prvC.mgmt?"pnl-pos":"pnl-neg"):""}">${curC&&prvC?(curC.mgmt>prvC.mgmt?"개선":"악화"):"—"}</td>
              </tr>`;
            }).join("")}
          </tbody>
        </table>
      </div>` : ""}
    </div>`;

  // 기간 버튼 이벤트
  el.querySelectorAll(".pnl-period-btn").forEach(btn => {
    btn.addEventListener("click", () => { pnlDashPeriod = btn.dataset.period; renderPnlDashboard(el); });
  });
  document.getElementById("pnlDashYear")?.addEventListener("change", e => { pnlDashYear = +e.target.value; renderPnlDashboard(el); });

  // Chart.js 차트 렌더
  if (!agg.length) return;
  try {
    await _loadChartJs();
  } catch (e) {
    console.warn("[손익] Chart.js 로드 실패:", e);
    el.querySelectorAll(".pnl-chart-box canvas").forEach(cv => {
      cv.insertAdjacentHTML("afterend", `<p style="color:#dc2626;font-size:12px;padding:16px;margin:0;text-align:center">차트 로드 실패<br><small>CDN 차단 또는 네트워크 오류</small></p>`);
      cv.style.display = "none";
    });
    return;
  }

  const labels = agg.map(d => d.label);
  const chartDefaults = { font: { family: "'Noto Sans KR', sans-serif", size: 11 } };
  Chart.defaults.font = chartDefaults.font;

  // 기존 차트 파기
  ["pnlChart1","pnlChart2","pnlChart3"].forEach(id => {
    if (_pnlCharts[id]) { _pnlCharts[id].destroy(); delete _pnlCharts[id]; }
  });

  // canvas가 DOM에서 사라진 경우(탭 이동 중 race) 중단
  const c1 = document.getElementById("pnlChart1");
  const c2 = document.getElementById("pnlChart2");
  const c3 = document.getElementById("pnlChart3");
  if (!c1 || !c2 || !c3) return;

  try {
  _pnlCharts.pnlChart1 = new Chart(c1, {
    type: "bar",
    data: {
      labels,
      datasets: [
        { label:"매출액",         data:agg.map(d=>d.revenue),  backgroundColor:"#3b82f620", borderColor:"#3b82f6", borderWidth:2 },
        { label:"상품매출원가",    data:agg.map(d=>d.cogs),     backgroundColor:"#94a3b8" },
        { label:"당기총제조비용",  data:agg.map(d=>d.mfg),      backgroundColor:"#cbd5e1" },
      ],
    },
    options: { responsive:true, plugins:{legend:{position:"top"}}, scales:{y:{ticks:{callback:v=>_pf(v)+"원"}}} },
  });

  _pnlCharts.pnlChart2 = new Chart(c2, {
    type: "line",
    data: {
      labels,
      datasets: [
        { label:"매출총이익", data:agg.map(d=>d.gross),  borderColor:"#16a34a", backgroundColor:"#16a34a18", tension:.3, fill:false },
        { label:"영업이익",  data:agg.map(d=>d.op),    borderColor:"#2563eb", backgroundColor:"#2563eb18", borderDash:[4,3], tension:.3, fill:false },
        { label:"경영이익",  data:agg.map(d=>d.mgmt),  borderColor:"#dc2626", backgroundColor:"#dc262618", borderDash:[2,2], tension:.3, fill:false },
      ],
    },
    options: { responsive:true, plugins:{legend:{position:"top"}}, scales:{y:{ticks:{callback:v=>_pf(v)+"원"}}} },
  });

  _pnlCharts.pnlChart3 = new Chart(c3, {
    type: "bar",
    data: {
      labels,
      datasets: [
        { label:"원가+제조비", data:agg.map(d=>d.cogs+d.mfg), backgroundColor:"#3b82f6", stack:"cost" },
        { label:"판관비",      data:agg.map(d=>d.sga),         backgroundColor:"#f97316", stack:"cost" },
        { label:"영업외비용",  data:agg.map(d=>d.interest),    backgroundColor:"#ef4444", stack:"cost" },
      ],
    },
    options: { responsive:true, plugins:{legend:{position:"top"}}, scales:{x:{stacked:true},y:{stacked:true,ticks:{callback:v=>_pf(v)+"원"}}} },
  });
  } catch(e) { console.warn("[손익] 차트 생성 실패:", e); }
}

init();

