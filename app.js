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

function getPayableEffectivePaid(item) {
  return Number(item.paidOverride != null ? item.paidOverride : item.paid || 0);
}

function getPayableOutstanding(item) {
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
    if (n1 === n2 && n1 !== "") return true;
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
    if (section === "payables" && !filterState.status) {
      if (item.completionStatus === "완료") return false;
    }
    if (filterState.status) {
      const balance = section === "payables"
        ? getPayableOutstanding(item)
        : section === "receivables"
          ? Number(item.balance || 0)
          : (item.purchase || item.sales || item.amount || 0) - (item.paid || 0);
      const isPaid = section === "fixed" ? Boolean(item.paid) : balance === 0;
      if (filterState.status === "completed" && !isPaid) return false;
      if (filterState.status === "pending" && isPaid) return false;
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
    <option value="">전체</option>
    <option value="pending">미완료 / 지급 대기</option>
    <option value="completed">완료 / 지급 완료</option>
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
                <div class="raw-diff-section-title removed-title" style="margin-bottom:5px;">🗑 사라진 항목 (${removedItems.length}건) — 완료 처리 추천</div>
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
                <div class="raw-diff-section-title changed-title" style="margin-bottom:5px;">✏️ 금액 변경 항목 (${changedItems.length}건)</div>
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
      statusEl.textContent = `✅ ${recipients.map(r => r.name).join(", ")} 님께 발송 완료`;
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
      statusEl.textContent = `✅ ${recipients.map(r => r.name).join(", ")} 님께 발송 완료`;
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
    const [vendorRows, rows, remotePlanRows, remoteHistoryRows] = await Promise.all([
      fetchVendorMasterRowsFromApi(),
      SHEET_APP_SCRIPT_URL ? fetchSheetWebApp() : fetchPublicSheet(),
      fetchSavedPaymentPlansFromApi(),
      fetchPaymentHistoryRowsFromApi(),
    ]);
    setVendorMasterRows(vendorRows);
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

    const applyPayables = (completeStableKeys = new Set()) => {
      payables = applySavedPayablesState(newParsedItems);
      // diff에서 완료 체크된 항목 처리
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
      ensureAutoPaymentPlans(); // 자동계산이 항상 마지막에 적용 (미확정 항목 한정)
      enrichPayablesWithVendorMaster();
      persistPayablesState();
      appendUpdateHistory("payables", diff); // 2단계: 변경 이력 기록
      renderPartnerFilter();
      renderFilterControls();
      rerenderAll();
    };

    if (diff.length > 0) {
      showPayablesRawDiffDialog(diff, applyPayables);
    } else {
      applyPayables();
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

async function loadSheetFixedExpenses() {
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

// raw 교체 시에도 살아남는 안정적 식별자 (금액/메모 제외)
function buildPayableStableKey(item) {
  const code = normalizeVendorCode(item.codeNormalized || item.code || item.codeRaw || "");
  return [
    code,
    String(item.year || ""),
    String(item.month || "").padStart(2, "0"),
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
  return localStorage.getItem(API_TOKEN_STORAGE_KEY) || "";
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
  const codeKey = allKeys.find(k => /코드|code/i.test(k)) || "";
  const mgrKey = allKeys.find(k => /담당자|manager/i.test(k)) || "";
  const emailKey = allKeys.find(k => /이메일|email/i.test(k)) || "";
  console.log("[담당자] 사용 컬럼 — 코드:", codeKey, "담당자:", mgrKey, "이메일:", emailKey);
  rows.forEach(row => {
    const code = normalizeVendorCode(String(codeKey ? (row[codeKey] ?? "") : "").trim());
    const manager = String(mgrKey ? (row[mgrKey] ?? "") : "").trim();
    const email = String(emailKey ? (row[emailKey] ?? "") : "").trim()
      || RECEIVABLE_MANAGER_EMAIL_MAP[manager] || "";
    if (code && manager) receivableManagerState.map.set(code, { manager, email });
  });
  console.log(`[담당자] 마스터 로드: ${receivableManagerState.map.size}건`, [...receivableManagerState.map.entries()].slice(0, 5));
}

function enrichReceivablesWithManager() {
  if (receivables.length) {
    const sample = receivables[0];
    const mapKeys = [...receivableManagerState.map.keys()].slice(0, 5);
    console.log("[담당자 매칭]", "맵크기:", receivableManagerState.map.size, "샘플코드:", sample.code, "맵키샘플:", mapKeys, "매칭:", receivableManagerState.map.get(sample.code));
  }
  receivables.forEach(item => {
    const mgr = receivableManagerState.map.get(item.code);
    if (mgr) {
      item.manager = mgr.manager || "미지정";
      item.managerEmail = mgr.email || "";
    } else {
      item.manager = "미지정";
      item.managerEmail = "";
    }
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
        거래처구분: String(row["거래처구분"] || ""),
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
        메모목록: [],
      });
    }
    const row = grouped.get(key);
    row.지급금액 += Number(item.decisionAmount || 0);
    row.연월목록.push(formatMonthKey(getMonthKey(item)));
    if (item.memo) {
      row.메모목록.push(item.memo);
    }
  });

  return [...grouped.values()].map(row => ({
    ...row,
    지급금액: Math.round(row.지급금액),
    연월목록: [...new Set(row.연월목록)].join(", "),
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

function buildCompletedApprovalHtmlForRows(reportRows) {
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
  // 같은 날 여러 배치면 회차 표시
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
  return batches;
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
        const html = buildCompletedApprovalHtmlForRows(rows);
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
            ${reportRows.length ? reportRows.map(row => {
    const holderRaw = String(row.예금주 || "").trim();
    const repRaw = String(row.대표자명 || "").trim();
    const holderNorm = holderRaw.replace(/\s+/g, "");
    const repNorm = repRaw.replace(/\s+/g, "");
    const holderIsPersonName = isPersonName(holderRaw);
    const mismatch = holderIsPersonName && repNorm && holderNorm !== repNorm;
    const holderHtml = holderRaw
      ? (mismatch
        ? `<span class="report-holder-mismatch" title="대표자명: ${escapeHtml(repRaw)}">${escapeHtml(holderRaw)}</span>`
        : escapeHtml(holderRaw))
      : '<span class="report-missing">확인 필요</span>';
    return `
              <tr>
                <td>${escapeHtml(row.거래처명 || "-")}</td>
                <td>${formatPlanLabel(row.결제예정일 || "")}</td>
                <td>${row.은행 || '<span class="report-missing">확인 필요</span>'}</td>
                <td>${row.계좌번호 || '<span class="report-missing">확인 필요</span>'}</td>
                <td>${holderHtml}</td>
                <td>${escapeHtml(row.연월목록 || "-")}</td>
                <td class="numeric-cell">${formatNumber(row.지급금액 || 0)}</td>
              </tr>`;
  }).join("") : `<tr><td colspan="7" class="empty-state">보고서로 만들 결제 대상이 없습니다.</td></tr>`}
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
        // 중요 식별자 및 숫자 필드에 ' 접두사 추가하여 시트에서 텍스트로 보존 (0 누락 방지)
        vendor_id: r.vendor_id ? (r.vendor_id.startsWith("'") ? r.vendor_id : "'" + r.vendor_id) : "",
        거래처코드_norm: r.거래처코드_norm ? (r.거래처코드_norm.startsWith("'") ? r.거래처코드_norm : "'" + r.거래처코드_norm) : "",
        사업자번호: r.사업자번호 ? (r.사업자번호.startsWith("'") ? r.사업자번호 : "'" + r.사업자번호) : "",
        계좌번호: r.계좌번호 ? (r.계좌번호.startsWith("'") ? r.계좌번호 : "'" + r.계좌번호) : "",
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
  vendorMasterState.lastMessage = "업체마스터 저장 중...";
  renderVendorMasterPanel();
  try {
    await postSheetWebApp("upsertVendorMaster", {
      sheetName: MASTER_SHEET_NAME,
      rows: targetRows,
    });
    setVendorMasterRows([
      ...vendorMasterState.rows.filter(existing => !targetRows.some(next => getVendorMatchKey(next) === getVendorMatchKey(existing))),
      ...targetRows,
    ]);
    enrichPayablesWithVendorMaster();
    vendorMasterState.lastMessage = `${targetRows.length}건을 업체마스터에 반영했습니다.`;
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
  const rawRows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  const allImportedRows = parseVendorMasterSheetRows(rawRows);
  const existingRows = parseVendorMasterSheetRows(await fetchVendorMasterRowsFromApi());
  const existingCodes = new Set(existingRows.map(r => r.거래처코드_norm).filter(Boolean));
  // 이미 마스터에 등록된 코드만 업데이트 대상으로
  const importedRows = allImportedRows.filter(r => existingCodes.has(r.거래처코드_norm));
  const { comparedRows, stats } = diffVendorMasterRows(existingRows, importedRows);

  vendorMasterState.importedRows = importedRows;
  vendorMasterState.comparedRows = comparedRows;
  vendorMasterState.stats = stats;
  vendorMasterState.lastFileName = file.name;
  vendorMasterState.lastMessage = `파일 ${allImportedRows.length}건 중 마스터 등록 ${importedRows.length}건 비교 완료.`;
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

    // 200건씩 배치 저장
    const BATCH = 200;
    const total = newRows.length;
    for (let i = 0; i < total; i += BATCH) {
      setLabel(`저장 중… ${Math.min(i + BATCH, total)}/${total}`);
      const batch = newRows.slice(i, i + BATCH).map(r => ({
        ...r,
        vendor_id: r.거래처코드_norm,
        사업자번호: "",
        active_yn: "Y",
        last_imported_at: new Date().toISOString(),
      }));
      await postSheetWebApp("upsertVendorMaster", { sheetName: MASTER_SHEET_NAME, rows: batch });
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
  Object.values(savedMap).forEach(v => {
    if (v.stableKey && !savedByStableKey[v.stableKey]) savedByStableKey[v.stableKey] = v;
  });
  return items.map(item => {
    const sourceKey = item.sourceKey || buildPayableSourceKey(item);
    const stableKey = buildPayableStableKey(item);
    const saved = savedMap[sourceKey] || savedByStableKey[stableKey] || null;
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
    // 결제 없고 보류/완료/부분결제가 아닌 항목은 stale decisionAmount 무시 → raw 잔액으로
    const isProtectedStatus = effectiveStatus === "보류" || effectiveStatus === "완료" || effectiveStatus === "부분결제";
    const noPayment = rawPaid === 0 && (effectivePO == null || effectivePO === 0);
    const shouldResetDA = !resetByRaw && !isProtectedStatus && noPayment &&
      saved.decisionAmount != null && Number(saved.decisionAmount) !== item.decisionAmount;
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

  if (!name || !Number(balance)) return null;
  if (condition === "제외") return null;
  if (memo.includes("제외")) return null;

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
  const groups = ["60일", "말일", "당말일", "05일", "당05일", "10일", "당10일", "15일", "당15일", "25일", "바로", "즉시"];
  const match = groups.find(group => text.includes(group));
  return match || text || "기타";
}

function parseSheetNumber(value) {
  if (value == null || value === "") return 0;
  return Number(String(value).replace(/[^0-9.-]/g, "")) || 0;
}

async function fetchSheetWebApp() {
  const url = new URL(SHEET_APP_SCRIPT_URL);
  const token = getApiToken();
  if (token) url.searchParams.set("token", token);
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
  if (Array.isArray(body)) return body;
  if (body.data && Array.isArray(body.data)) return body.data;
  throw new Error("Apps Script 응답 형식이 올바르지 않습니다.");
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
      if (!cell) { item[cols[index]] = ""; return; }
      // 날짜 셀: v는 "Date(2026,3,21)" 형태 → 사람이 읽을 수 있는 f값 우선 사용
      if (colTypes[index] === "date" || colTypes[index] === "datetime") {
        item[cols[index]] = cell.f ?? cell.v ?? "";
      } else {
        item[cols[index]] = cell.v ?? "";
      }
    });
    return item;
  });
}

async function fetchPublicSheet() {
  return fetchPublicSheetByName(SHEET_NAME_PAYABLES);
}

function rerenderAll() {
  renderSummary();
  renderReceivables();
  renderPayables();
  renderFixedExpenses();
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
  const summary = calculateSummary();
  elements.summaryPanel.innerHTML = `
    <!-- 그룹 1: 매출 현황 (미수금) -->
    <div class="summary-group green">
      <div class="summary-card">
        <h2>매출</h2>
        <p>${formatNumber(summary.totalReceivable)}</p>
      </div>
      <div class="summary-card highlight">
        <h2>미수금 잔액</h2>
        <p>${formatNumber(summary.totalOutstanding)}</p>
      </div>
    </div>

    <!-- 그룹 2: 매입 현황 (미지급) -->
    <div class="summary-group blue">
      <div class="summary-card">
        <h2>매입</h2>
        <p>${formatNumber(summary.totalPayable)}</p>
      </div>
      <div class="summary-card highlight">
        <h2>미지급 잔액</h2>
        <p>${formatNumber(summary.totalUnpaid)}</p>
      </div>
    </div>

    <!-- 그룹 3: 기타 현황 -->
    <div class="summary-group slate">
      <div class="summary-card">
        <h2>고정지출 합계</h2>
        <p>${formatNumber(summary.totalFixed)}</p>
      </div>
    </div>
  `;
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
      `;}).join("")}
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

    const groupTotalCells = monthKeys.map((mk, idx) => {
      const t = [...group.vendors.values()].reduce((s, v) => s + (v.months[mk] || 0), 0);
      return `<td class="group-summary-cell month-column-cell ${idx % 2 === 0 ? "month-column-even" : "month-column-odd"}">${t ? formatNumber(t) : ""}</td>`;
    }).join("");

    const itemRowsHtml = collapsed ? "" : sortedVendors.map((vendor, rowIdx) => {
      const el = vendor.maxElapsed;
      let elapsedHtml = "-", elapsedClass = "";
      if (el !== null && el !== undefined) {
        if (el >= 60) { elapsedHtml = `${el}일`; elapsedClass = "rcv-elapsed-danger"; }
        else if (el >= 30) { elapsedHtml = `${el}일`; elapsedClass = "rcv-elapsed-warn"; }
        else if (el >= 0) { elapsedHtml = `${el}일`; elapsedClass = "rcv-elapsed-ok"; }
        else { elapsedHtml = `D${el}`; elapsedClass = "rcv-elapsed-future"; }
      }
      const monthCells = monthKeys.map((mk, idx) => {
        const val = vendor.months[mk] || 0;
        return `<td class="numeric-cell month-column-cell ${idx % 2 === 0 ? "month-column-even" : "month-column-odd"}">${val ? formatNumber(val) : ""}</td>`;
      }).join("");
      const vCode = normalizeVendorCode(vendor.codeRaw || vendor.name || "");
      const rcvTooltip = buildVendorTooltip(vCode, vendor.memo, "receivables");
      const memoAttr = rcvTooltip ? ` title="${rcvTooltip.replace(/"/g, "&quot;")}"` : "";
      const hasVMemo = !!(getVendorMemo(vCode).common || getVendorMemo(vCode).receivables);
      const mgrHtml = vendor.manager && vendor.manager !== "미지정" ? `<span class="rcv-manager-badge">${escapeHtml(vendor.manager)}</span>` : "";
      return `<tr class="${rowIdx % 2 === 0 ? "rcv-row-even" : "rcv-row-odd"}">
          <td class="partner-name-cell">
            <div class="partner-name-cell-inner">
              <span class="partner-name-button ${(vendor.memo || hasVMemo) ? "has-memo" : ""}"${memoAttr}>${escapeHtml(vendor.name)}</span>
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
        <td>
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

  elements.receivables.innerHTML = `
    <div class="panel">
      <div class="panel-title-row">
        <div class="panel-title-inline">
          <h3>미수금 목록</h3>
          ${filtered.length ? `<span class="rcv-summary-text">${filtered.length}건 · ${formatNumber(totalBalance)}원</span>` : ""}
        </div>
        <div class="payable-table-actions">
          <button type="button" class="table-action-button subtle rcv-expand-all">전체 펼치기</button>
          <button type="button" class="table-action-button subtle rcv-collapse-all">전체 접기</button>
          <button type="button" class="header-action-button" id="receivableEmailButton">이메일 발송</button>
        </div>
      </div>
      <div class="rcv-group-chips" id="rcvGroupChips">
        <button type="button" class="group-manage-link chip-select-all">전체 선택</button>
        <button type="button" class="group-manage-link chip-clear-all">전체 해제</button>
        ${chipsHtml}
      </div>
      <div class="table-responsive">
        <table class="rcv-pivot-table">
          <thead>
            <tr>
              <th class="rcv-sort-th" data-sort="code">거래처명 ${rcvSortState.key === "code" ? (rcvSortState.dir === "asc" ? "▲" : "▼") : "⇅"}</th>
              <th class="numeric-header rcv-sort-th" data-sort="elapsed">경과일수 ${rcvSortState.key === "elapsed" ? (rcvSortState.dir === "asc" ? "▲" : "▼") : "⇅"}</th>
              ${monthKeys.map((mk, idx) => `<th class="numeric-header month-column-cell ${idx % 2 === 0 ? "month-column-even" : "month-column-odd"}">${mk.slice(2)}</th>`).join("")}
              <th class="numeric-header">합계</th>
            </tr>
          </thead>
          <tbody>
            ${visibleGroups.length ? groupsHtml : `<tr><td colspan="${3 + monthKeys.length}" class="empty-state">${receivables.length ? "표시할 미수금 데이터가 없습니다." : "미수금 데이터를 불러오는 중입니다."}</td></tr>`}
          </tbody>
        </table>
      </div>
    </div>
  `;

  document.getElementById("receivableEmailButton")?.addEventListener("click", openReceivableEmailDialog);

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
        <button type="button" class="rcv-send-btn">발송</button>
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

  if (!conditions.length) { alert("수금조건을 최소 1개 선택해주세요."); return; }
  if (!managers.length && !sendSummary) { alert("담당자를 최소 1명 선택해주세요."); return; }
  if (sendSummary && !summaryRecipients.length) { alert("전체 현황 보고서를 수신할 사람을 최소 1명 선택해주세요."); return; }

  const sendBtn = q(".rcv-send-btn");
  sendBtn.disabled = true;
  sendBtn.textContent = "발송 중...";

  try {
    const result = await postSheetWebApp("sendReceivableEmails", {
      managers, absentChain, ccEmails, conditions,
      testMode, testRecipient, sendSummary, excludeMinus, senderName, summaryRecipients
    });
    overlay.remove();
    const modeNote = testMode ? `\n※ 테스트: ${testRecipient || ""}으로 발송` : "";
    alert(`발송 완료! ${result.sentCount || ""}건${modeNote}`);
  } catch (error) {
    alert(`발송 실패: ${error.message}`);
    sendBtn.disabled = false;
    sendBtn.textContent = "발송";
  }
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

function getPayablesForPlanKey(planKey, sourceItems) {
  if (planKey === "__total__") {
    return [...sourceItems];
  }
  return sourceItems.filter(item => (item.paymentPlan || "") === planKey);
}

function getPayablesForPlanKeys(planKeys, sourceItems) {
  const uniqueKeys = [...new Set((planKeys || []).filter(Boolean))];
  if (!uniqueKeys.length) return [];
  if (uniqueKeys.includes("__total__")) {
    return [...sourceItems];
  }
  const seen = new Set();
  const result = [];
  uniqueKeys.forEach(planKey => {
    getPayablesForPlanKey(planKey, sourceItems).forEach(item => {
      const key = item.sourceKey || buildPayableSourceKey(item);
      if (!seen.has(key)) {
        seen.add(key);
        result.push(item);
      }
    });
  });
  return result;
}

function formatMonthKey(key) {
  const [year, month] = key.split("-");
  return `${String(year).slice(2)}-${month}`;
}

function formatPlanLabel(key) {
  if (!key) return "미정";
  if (key === "보류") return "보류";
  const normalized = normalizeDateValue(key);
  const match = normalized.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (match) return `${Number(match[2])}/${Number(match[3])}`;
  return normalized;
}

function formatPlanShortLabel(key) {
  if (!key) return "미정";
  if (key === "보류") return "보류";
  const normalized = normalizeDateValue(key);
  const match = normalized.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (match) return `${Number(match[2])}/${Number(match[3])}`;
  return normalized;
}

function normalizeDueGroupLabel(label) {
  return String(label || "").replace(/\s+/g, "").trim();
}

// ── 한국 공휴일 + 평일 보정 ───────────────────────────────────
const KR_HOLIDAYS = new Set([
  // 2025
  "2025-01-01", "2025-01-28", "2025-01-29", "2025-01-30",
  "2025-03-01", "2025-03-03",         // 삼일절(토)→월 대체
  "2025-05-05", "2025-05-06",         // 어린이날+부처님오신날 대체공휴일
  "2025-06-06",
  "2025-08-15",
  "2025-10-03", "2025-10-06", "2025-10-07", "2025-10-08", "2025-10-09",
  "2025-12-25",
  // 2026
  "2026-01-01",
  "2026-02-16", "2026-02-17", "2026-02-18",
  "2026-03-01", "2026-03-02",         // 삼일절(일)→월 대체
  "2026-05-05",                      // 어린이날
  "2026-05-24", "2026-05-25",         // 부처님오신날(일)→월 대체
  "2026-06-06",
  "2026-08-15",
  "2026-09-24", "2026-09-25", "2026-09-26",
  "2026-10-03", "2026-10-09",
  "2026-12-25",
]);

function nextBusinessDay(date) {
  const d = new Date(date);
  d.setHours(0, 0, 0, 0);
  while (true) {
    const dow = d.getDay();
    const ymd = toYMD(d);
    if (dow !== 0 && dow !== 6 && !KR_HOLIDAYS.has(ymd)) break;
    d.setDate(d.getDate() + 1);
  }
  return d;
}

function prevBusinessDay(date) {
  const d = new Date(date);
  d.setHours(0, 0, 0, 0);
  while (true) {
    const dow = d.getDay();
    const ymd = toYMD(d);
    if (dow !== 0 && dow !== 6 && !KR_HOLIDAYS.has(ymd)) break;
    d.setDate(d.getDate() - 1);
  }
  return d;
}

function toYMD(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function getDueGroupRank(label) {
  const normalized = normalizeDueGroupLabel(label);
  const rankMap = {
    "당말일": 0,
    "당05일": 1,
    "당10일": 2,
    "당15일": 3,
    "25일": 4,
    "말일": 5,
    "60일": 6,
    "05일": 7,
    "10일": 8,
    "15일": 9,
    "바로": 10,
    "즉시": 10,
  };
  return rankMap[normalized] ?? 99;
}

function getDueGroup(item) {
  return item.dueCategory || item.payDate || item.memo || "기타";
}

const PAYABLE_DUE_RULES = {
  "당말일": [0, "last"],
  "말일": [1, "last"],
  "60일": [2, "last"],
  "25일": [1, 25],
  "당05일": [1, 5],
  "05일": [2, 5],
  "당10일": [1, 10],
  "10일": [2, 10],
  "당15일": [1, 15],
  "15일": [2, 15],
};

function calcPayableDueDate(year, month, groupLabel) {
  const label = normalizeDueGroupLabel(groupLabel);
  const rule = PAYABLE_DUE_RULES[label];
  if (!rule || !year || !month) return "";
  const [addMonths, day] = rule;
  const rawMonth = month + addMonths;
  const targetYear = year + Math.floor((rawMonth - 1) / 12);
  const targetMonth = ((rawMonth - 1) % 12) + 1;
  let raw;
  if (day === "last") {
    raw = new Date(targetYear, targetMonth, 0);
  } else {
    raw = new Date(targetYear, targetMonth - 1, day);
  }
  return toYMD(nextBusinessDay(raw));
}

function getGroupAutoPayDate(groupLabel, consolidatedItems) {
  let maxYear = 0, maxMonth = 0;
  consolidatedItems.forEach(entry => {
    (entry.items || []).forEach(item => {
      const yr = Number(item.year || 0);
      const mo = Number(item.month || 0);
      if (yr > maxYear || (yr === maxYear && mo > maxMonth)) {
        maxYear = yr; maxMonth = mo;
      }
    });
  });
  if (!maxYear || !maxMonth) return "";
  return calcPayableDueDate(maxYear, maxMonth, groupLabel);
}

function getItemAutoPayDate(item) {
  return calcPayableDueDate(Number(item.year || 0), Number(item.month || 0), getDueGroup(item));
}

function ensureAutoPaymentPlans() {
  payables.forEach(item => {
    const auto = getItemAutoPayDate(item);
    if (!auto) return; // 계산 불가한 항목은 건드리지 않음
    // 사용자가 날짜를 직접 변경한 경우만 보호 (보류=특정날짜지정, 완료=결제완료)
    // "예정"은 자동계산 날짜 위에 눌린 것이므로 재계산 적용
    const status = (item.completionStatus || "").trim();
    const isManuallySet = status === "보류" || status === "완료" || status === "미정";
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
  const filteredPayables = getFilteredItems(payables, "payables");
  const matchedVendorCount = [...new Set(filteredPayables.filter(item => item.vendorMatched).map(item => getPartnerGroupKey(item)))].length;
  const unmatchedVendorCount = [...new Set(filteredPayables.filter(item => !item.vendorMatched).map(item => getPartnerGroupKey(item)))].length;
  const monthKeys = getUniqueSortedMonthKeys(filteredPayables);
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
    const groupSummaryCells = monthKeys.map((key, index) => `
      <td class="group-summary-cell month-column-cell ${index % 2 === 0 ? "month-column-even" : "month-column-odd"}">${formatPayableCellNumber(groupTotals.monthTotals[key] || 0)}</td>
    `).join("");
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

    const itemRows = collapsed ? "" : group.items.map(entry => {
      const checked = entry.selected ? "checked" : "";
      const partnerKey = encodeURIComponent(getPartnerGroupKey(entry.items[0]));

      const monthCells = monthKeys.map((monthKey, index) => {
        const decisionValue = entry.monthTotals[monthKey] || 0;
        const monthItems = entry.items.filter(item => getMonthKey(item) === monthKey);
        const originalValue = monthItems.reduce((sum, item) => sum + getPayableOutstanding(item), 0);
        const totalPurchase = monthItems.reduce((sum, item) => sum + Number(item.purchase || 0), 0);
        const totalRawPaid = monthItems.reduce((sum, item) => sum + Number(item.paid || 0), 0);
        const cellPlanValue = monthItems[0]?.paymentPlan || "";
        const autoPlanValue = monthItems[0] ? getItemAutoPayDate(monthItems[0]) : "";
        const isMijeong = monthItems.some(item => item.completionStatus === "미정");
        const planClass = cellPlanValue === "보류" ? "hold" : cellPlanValue ? "set" : "pending";
        const planLabel = isMijeong ? "미정" : formatPlanShortLabel(cellPlanValue || autoPlanValue || "");
        const showOriginalValue = originalValue > 0 && decisionValue !== originalValue;
        // raw 지급합이 있으면 합계/지급 정보 표시 (버튼 아님)
        const showRawBreakdown = totalRawPaid > 0 && totalPurchase > originalValue;
        const isLastEdited = payablesUiState.lastEdited
          && payablesUiState.lastEdited.partnerKey === getPartnerGroupKey(entry.items[0])
          && payablesUiState.lastEdited.monthKey === monthKey;
        if (originalValue === 0) {
          return `<td class="editable-amount-cell numeric-cell month-column-cell ${index % 2 === 0 ? "month-column-even" : "month-column-odd"}"></td>`;
        }
        return `
          <td class="editable-amount-cell numeric-cell month-column-cell ${index % 2 === 0 ? "month-column-even" : "month-column-odd"} ${isLastEdited ? "recently-edited-cell" : ""}">
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
              <button class="history-payable-button" type="button" title="과거 이력 및 롤백" data-partner-key="${partnerKey}" data-month-key="${monthKey}" style="border:none;background:transparent;cursor:pointer;font-size:12px;opacity:0.6;padding:0 2px;">🕒</button>
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
              <span class="vendor-match-chip ${entry.items[0]?.vendorMatched ? "matched" : "unmatched"}">${entry.items[0]?.vendorMatched ? "계좌연결" : "계좌확인"}</span>`;
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

  elements.payables.innerHTML = `
    <div class="panel">
      <div class="panel-title-row">
        <div class="panel-title-inline">
          <h3>미지급 목록</h3>
          ${showBatchButton ? `<button type="button" class="batch-selected-button payment-plan-batch-button">일괄 계획 변경</button>` : ""}
        </div>
        <div class="payable-table-actions">
          <button type="button" class="table-action-button subtle" data-action="expand-all">전체 펼치기</button>
          <button type="button" class="table-action-button subtle" data-action="collapse-all">전체 접기</button>
        </div>
      </div>
      <div class="rcv-group-chips payable-group-chips" id="payGroupChips">
        <button type="button" class="group-manage-link chip-select-all">전체 선택</button>
        <button type="button" class="group-manage-link chip-clear-all">전체 해제</button>
        ${buildGroupChipsHtml(getOrderedDueGroups(payables), filterState.groups, "pay-chip")}
      </div>
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
      <div class="table-responsive">
        <table>
          <thead>
            <tr>
              <th class="sticky-col sticky-col-1 payable-header-cell">선택</th>
              <th class="sticky-col sticky-col-2 payable-header-cell">업체명</th>
              ${monthKeys.map((key, index) => `<th class="numeric-header month-column-cell ${index % 2 === 0 ? "month-column-even" : "month-column-odd"}">${formatMonthKey(key)}</th>`).join("")}
              <th class="numeric-header">합계</th>
            </tr>
          </thead>
          <tbody>
            ${rows || `<tr><td colspan="${3 + monthKeys.length}" class="empty-state">선택한 거래처에 대한 미지급이 없습니다.</td></tr>`}
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

  document.querySelectorAll(".history-payable-button").forEach(button => {
    button.addEventListener("click", event => {
      event.preventDefault();
      event.stopPropagation();
      const partnerKey = decodeURIComponent(event.currentTarget.dataset.partnerKey || "");
      const monthKey = event.currentTarget.dataset.monthKey;
      
      const targetItems = payables.filter(item => getPartnerGroupKey(item) === partnerKey && getMonthKey(item) === monthKey);
      if (!targetItems.length) return;
      
      showPaymentPlanHistoryDialog(targetItems, partnerKey, monthKey);
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
    topScrollbarInner.style.width = `${table.scrollWidth}px`;
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
      <h3 style="margin-top:0; display:flex; align-items:center; gap:8px;">🕒 상세 변경 타임라인</h3>
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
            <input type="date" class="editor-plan-date-input" value="${holdPlan ? (autoPlanValue || "") : (/^\d{4}-\d{2}-\d{2}$/.test(currentPlanValue) ? currentPlanValue : (autoPlanValue || ""))}" ${holdPlan ? "disabled" : ""} />
          </label>
          <div class="editor-plan-actions">
            <button type="button" class="editor-plan-reset-button">미정</button>
            ${autoPlanValue ? `<button type="button" class="editor-plan-default-button">기본 ${autoPlanValue.replace(/^(\d{4})-(\d{2})-(\d{2})$/, "$2/$3")}</button>` : ""}
            <button type="button" class="editor-plan-hold-button ${holdPlan ? "active" : ""}">보류</button>
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
    if (!planDateInput || !planHoldButton) return;
    planDateInput.disabled = holdPlan;
    planHoldButton.classList.toggle("active", holdPlan);
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
    const nextPlanValue = holdPlan ? "보류" : (planDateInput?.value || "");
    monthItems.forEach(item => {
      item.paymentPlan = nextPlanValue;
      item.completionStatus = nextPlanValue ? "보류" : "미정";
    });
    payablesUiState.lastEdited = { partnerKey: decodedKey, monthKey };
    persistPayablesState();
    closeCalculator();
    preserveViewport(() => rerenderAll());
  }

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
      if (planDateInput) planDateInput.value = "";
      syncPlanControls();
    });
  }

  if (planDefaultButton) {
    planDefaultButton.addEventListener("click", event => {
      event.preventDefault();
      holdPlan = false;
      if (planDateInput) {
        planDateInput.value = autoPlanValue || "";
      }
      syncPlanControls();
    });
  }

  if (planHoldButton) {
    planHoldButton.addEventListener("click", event => {
      event.preventDefault();
      holdPlan = !holdPlan;
      syncPlanControls();
    });
  }

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

  const overlay = document.createElement("div");
  overlay.className = "batch-plan-overlay";
  overlay.innerHTML = `
    <div class="batch-plan-popover" role="dialog" aria-modal="true">
      <div class="batch-plan-title">${planKey === "__total__" ? "전체 예정 변경" : planKey === "선택 항목" ? `선택 항목 일괄 변경` : `${formatPlanLabel(planKey)} 일괄 변경`}</div>
      <p class="batch-plan-note">${targetItems.length}건에 같은 결제 계획을 적용합니다.</p>
      <label class="editor-plan-label">
        결제 예정일
        <input type="date" class="editor-plan-date-input" value="${holdPlan ? "" : firstDate}" ${holdPlan ? "disabled" : ""} />
      </label>
      <div class="editor-plan-actions">
        <button type="button" class="editor-plan-reset-button">미정</button>
        <button type="button" class="editor-plan-hold-button ${holdPlan ? "active" : ""}">보류</button>
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
  const resetButton = overlay.querySelector(".editor-plan-reset-button");

  function syncState() {
    dateInput.disabled = holdPlan;
    holdButton.classList.toggle("active", holdPlan);
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
    const nextPlanValue = holdPlan ? "보류" : (dateInput.value || "");
    targetItems.forEach(item => {
      item.paymentPlan = nextPlanValue;
      item.completionStatus = nextPlanValue ? "보류" : "미정";
    });
    persistPayablesState();
    closeBatchPlanEditor();
    preserveViewport(() => rerenderAll());
  }

  holdButton.addEventListener("click", () => {
    holdPlan = !holdPlan;
    syncState();
  });

  resetButton.addEventListener("click", () => {
    holdPlan = false;
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

  const grandTotal = sortedFixed.reduce((s, i) => s + (i.amount || 0), 0);
  const grandBankTotals = {};
  uniqueBanks.forEach(b => grandBankTotals[b] = 0);
  sortedFixed.forEach(item => {
    if (grandBankTotals[item.bank] !== undefined) {
      grandBankTotals[item.bank] += (item.amount || 0);
    }
  });

  const bankColWidth = uniqueBanks.length > 3 ? 100 : 120; // 은행 수에 따라 너비 자동조절

  // 날짜별 그룹 행 생성
  const groupRows = dateOrder.map(key => {
    const { items, month, day } = dateGroups[key];
    const dayTotal = items.reduce((s, i) => s + (i.amount || 0), 0);
    const groupId = `fixed-date-${key.replace(/\W/g, "")}`;

    // 해당 날짜의 은행별 합계 계산
    const dayBankTotals = {};
    uniqueBanks.forEach(b => dayBankTotals[b] = 0);
    items.forEach(item => {
      if (dayBankTotals[item.bank] !== undefined) {
        dayBankTotals[item.bank] += (item.amount || 0);
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
            ${b === item.bank && item.amount ? `<span style="color:#0f172a;">${formatNumber(item.amount)}</span>` : ''}
          </td>
        `).join("")}
        <td style="text-align:right;font-weight:600;color:#1e293b;padding:10px 12px;font-size:14px;">
          ${item.amount ? formatNumber(item.amount) : ''}
        </td>
      </tr>
    `).join("");

    return `
      <tr class="fx-date-header" data-group="${groupId}">
        <td class="fx-header-check" style="text-align:center;padding:12px 8px;">
          <input type="checkbox" class="fixed-day-check fx-checkbox" data-total="${dayTotal}" checked style="cursor:pointer;width:16px;height:16px;accent-color:#2563eb;">
        </td>
        <td class="fx-header-title" style="padding-top:14px;padding-bottom:14px;">
          <span class="fx-chevron fixed-toggle-btn" style="display:inline-block;transition:all 0.2s;margin-right:8px;font-size:12px;color:#94a3b8;">▼</span>
          <span class="fx-date-badge" style="background:#eff6ff;color:#1d4ed8;padding:4px 10px;border-radius:16px;font-size:14px;margin-right:8px;font-weight:700;">${month}/${day}</span>
          <span class="fx-count-pill" style="font-size:12px;color:#3b82f6;background:#dbeafe;padding:3px 8px;border-radius:12px;font-weight:600;">${items.length}건</span>
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
          <div class="fx-btn-group" style="display:flex;gap:8px;">
            <button id="fixedExpandAll" class="fx-ctrl-btn" style="padding:6px 12px;font-size:13px;font-weight:600;background:#fff;border:1px solid #cbd5e1;border-radius:6px;cursor:pointer;color:#334155;box-shadow:0 1px 2px rgba(0,0,0,0.05);transition:all 0.15s;" onmouseover="this.style.backgroundColor='#f8fafc'" onmouseout="this.style.backgroundColor='#fff'">🔽 전체 펼치기</button>
            <button id="fixedCollapseAll" class="fx-ctrl-btn" style="padding:6px 12px;font-size:13px;font-weight:600;background:#fff;border:1px solid #cbd5e1;border-radius:6px;cursor:pointer;color:#334155;box-shadow:0 1px 2px rgba(0,0,0,0.05);transition:all 0.15s;" onmouseover="this.style.backgroundColor='#f8fafc'" onmouseout="this.style.backgroundColor='#fff'">🔼 전체 접기</button>
          </div>
          <div class="fx-total-chip" style="background:#eff6ff;padding:8px 16px;border-radius:10px;border:1px solid #bfdbfe;display:flex;align-items:center;">
            <span class="fx-total-label" style="font-size:13px;color:#1e40af;margin-right:8px;font-weight:700;">선택 합계</span>
            <strong id="fixedCheckedTotal" class="fx-total-value" style="font-size:17px;color:#1d4ed8;letter-spacing:-0.01em;">${formatNumber(grandTotal)}</strong>
          </div>
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

  elements.fixed.querySelectorAll(".fixed-day-check").forEach(cb => {
    cb.addEventListener("change", updateCheckedTotal);
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
      if (target === "daesa") renderDaesaTab();
      if (target === "fixed") renderFixedExpenses();
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
  }).filter(r => r._debit > 0 || r._credit > 0);  // 금액 없는 행 제외
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
};

// 대사 탭 정렬 상태
const daesaSortState = { key: "name", dir: "asc" };

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

async function loadDaesaData() {
  if (daesaState.loading) return;
  daesaState.loading = true;
  daesaState.error = null;
  renderDaesaTab();
  try {
    const [tax, lSales, lPurchase, lPayable, daily, bizDiv] = await Promise.all([
      fetchApiRows("getTaxInvoices"),
      fetchApiRows("getLedgerSales"),
      fetchApiRows("getLedgerPurchase"),
      fetchApiRows("getLedgerPayable"),
      fetchApiRows("getDailySales"),
      fetchApiRows("getBizDivision").catch(() => []),  // 없으면 빈 배열
    ]);
    daesaState.taxInvoices = tax;
    daesaState.ledgerSales = lSales;
    daesaState.ledgerPurchase = lPurchase;
    daesaState.ledgerPayable = lPayable;
    daesaState.dailySales = daily;
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
    });

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

  const rows = vendorEntries.map(([code, vendor]) => {
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
      ${vendorEntries.some(([c]) => netOffSet.has(c)) ? `<td class="num col-netoff">${formatNumber(netoffAmt)}</td>` : ""}
    </tr>`;
  }).join("");

  const hasNetOff = vendorEntries.some(([code]) => netOffSet.has(code));

  function sortIcon(key) {
    if (daesaSortState.key !== key) return '<span class="sort-arrow">↕</span>';
    return daesaSortState.dir === "asc" ? '<span class="sort-arrow">↑</span>' : '<span class="sort-arrow">↓</span>';
  }

  section.innerHTML = `
    <div class="daesa-toolbar">
      <select id="daesaYearFilter">${yearOpts}</select>
      <select id="daesaMonthFilter">${monthOpts}</select>
      <button class="daesa-reload-btn">↺ 새로고침</button>
      <span class="daesa-count muted">${vendorEntries.length}개 업체 표시중 ${q ? `(검색: ${q})` : ""}</span>
    </div>
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
    acc.taxSales += d.taxSales;
    acc.ledgerSales += d.ledgerSales;
    acc.bizSales += d.bizSales;
    acc.collect += d.ledgerCollect || d.bizCollect;
    acc.netoff += netoffAmt;
    acc.taxBuy += d.taxPurchase;
    acc.ledgerBuy += ledgerBuyTotal;
    acc.bizBuy += d.bizPurchase;
    acc.pay += ledgerPayTotal || d.bizPay;
    return acc;
  }, { taxSales: 0, ledgerSales: 0, bizSales: 0, collect: 0, netoff: 0, taxBuy: 0, ledgerBuy: 0, bizBuy: 0, pay: 0 });

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
        const bdRows = Object.entries(bd).map(([div, info]) => `
          <tr>
            <td>${escapeHtml(div)}</td>
            <td class="num">${formatNumber(info.sales)}</td>
            <td class="num">${formatNumber(info.collect)}</td>
            <td class="num">${formatNumber(info.sales - info.collect)}</td>
          </tr>`).join("");
        breakdownTable = `
          <div class="daesa-detail-section">
            <strong>사업부문/현장별 실적</strong>
            <table class="daesa-subtable">
              <thead><tr><th>사업부문명</th><th>매출</th><th>수금</th><th>잔액</th></tr></thead>
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

    const salesCols = showSales ? `
      <td class="num">${formatNumber(d.taxSales)}</td>
      <td class="num">${formatNumber(d.ledgerSales)}</td>
      <td class="num">${formatNumber(d.bizSales)}</td>
      <td class="num daesa-collect">${formatNumber(d.ledgerCollect || d.bizCollect)}</td>
      ${isNetOff ? `<td class="num daesa-netoff-amt">${formatNumber(netoffAmt)}</td>` : ""}
      <td class="num daesa-balance">${formatNumber(netSales - (d.ledgerCollect || d.bizCollect))}</td>
    ` : "";

    const buyCols = showBuy ? `
      <td class="num">${formatNumber(d.taxPurchase)}</td>
      <td class="num">${formatNumber(ledgerBuyTotal)}</td>
      <td class="num">${formatNumber(d.bizPurchase)}</td>
      <td class="num daesa-pay">${formatNumber(ledgerPayTotal || d.bizPay)}</td>
      ${isNetOff ? `<td class="num daesa-netoff-amt">${formatNumber(netoffAmt)}</td>` : ""}
      <td class="num daesa-balance">${formatNumber(netBuy - (ledgerPayTotal || d.bizPay))}</td>
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
                  <td class="num daesa-balance">${formatNumber(netSalesTotal - totals.collect)}</td>
                ` : ""}
                ${showBuy ? `
                  <td class="num">${formatNumber(totals.taxBuy)}</td>
                  <td class="num">${formatNumber(totals.ledgerBuy)}</td>
                  <td class="num">${formatNumber(totals.bizBuy)}</td>
                  <td class="num daesa-pay">${formatNumber(totals.pay)}</td>
                  ${isNetOff ? `<td class="num daesa-netoff-amt">${formatNumber(totals.netoff)}</td>` : ""}
                  <td class="num daesa-balance">${formatNumber(netBuyTotal - totals.pay)}</td>
                ` : ""}
              </tr>
            </tfoot>
          </table>
        </div>
      </div>
      <div class="daesa-modal-footer">
        <button class="daesa-modal-print">🖨 인쇄 / PDF</button>
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
    const printWin = window.open("", "_blank", "width=1100,height=750");
    const tableHtml = overlay.querySelector("table").cloneNode(true);
    tableHtml.querySelectorAll(".daesa-modal-detail-row").forEach(r => r.classList.remove("hidden"));

    printWin.document.write(`<!DOCTYPE html><html><head><meta charset="UTF-8">
      <title>${name} 대사 현황</title>
      <style>
        body{font-family:'맑은 고딕',sans-serif;font-size:11px;margin:16px;}
        table{border-collapse:collapse;width:100%;}
        th,td{border:1px solid #ccc;padding:4px 6px;text-align:right;}
        th{background:#f1f5f9;text-align:center;}
        .num{text-align:right;}
        .daesa-modal-detail-row { background: #fafafa; }
        .daesa-modal-detail-cell { padding: 4px 12px; text-align: left; }
        tfoot tr{background:#f0f4ff;font-weight:bold;}
        h2{margin:0 0 12px;}
        .daesa-expand-icon { display: none; }
      </style></head><body>
      <h2>${escapeHtml(name)}${isNetOff ? " (상계업체)" : ""} — 누적 대사 현황</h2>
      ${tableHtml.outerHTML}
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
    headers.forEach((h, j) => { row[h] = raw[j] ?? ""; });
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
          `${row["일자"]}_${String(row["견표번호"] || "").trim()}_${String(row["거래처코드"] || "").trim()}`;
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
              ${sec.saving ? "저장 중…" : "구글시트 저장"}
            </button>
          ` : ""}
        </div>
        ${sec.status ? `<div class="di-status ${sec.status.startsWith("✓") ? "di-status-ok" : sec.status.startsWith("저장") ? "" : "di-status-err"}">${sec.status}</div>` : ""}
      </div>
    `;
  }).join("");

  panel.innerHTML = `
    <div class="di-header">
      <h3>자료 업로드</h3>
      <p class="di-desc muted">파일을 선택하면 파싱 결과를 미리 보여줍니다. '구글시트 저장'을 눌러야 반영됩니다.<br>전체 기간을 항상 올려도 괜찮습니다 — 중복 행은 자동으로 덮어씁니다.</p>
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
      const sec = dataImportState[key];
      sec.status = "파싱 중…";
      sec.parsed = null;
      renderDataImportPanel();

      let result;
      if (key === "taxInvoice") result = await parseTaxInvoiceFile(file);
      else if (key === "dailySales") result = await parseDailySalesFile(file);
      else result = await parseLedgerFile(file);

      sec.parsed = result.error ? null : result.rows;
      sec.status = result.error ? `오류: ${result.error}` : "";
      renderDataImportPanel();
    });
  });

  panel.querySelectorAll(".di-save-btn").forEach(btn => {
    btn.addEventListener("click", async () => {
      const key = btn.dataset.key;
      const sec = dataImportState[key];
      if (!sec.parsed?.length) return;
      sec.saving = true;
      sec.status = "저장 중…";
      renderDataImportPanel();
      try {
        const { action, ...extra } = DATA_IMPORT_ACTIONS[key];
        await postSheetWebApp(action, { rows: sec.parsed, ...extra });
        sec.status = `✓ ${sec.parsed.length}건 저장 완료`;
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
  renderPartnerFilter();
  renderFilterControls();
  renderVendorMasterPanel();
  setupTabs();
  setupMasterMenu();
  setupVendorMasterImport();
  setupLedgerVendorImport();
  setupBankImport();
  setupDataImport();
  setupApiTokenButton();
  // 검색창 → 대사 탭 연동
  elements.searchInput?.addEventListener("input", () => {
    const daesaEl = document.getElementById("daesa");
    if (daesaEl && !daesaEl.classList.contains("hidden") && daesaState.loaded) {
      renderDaesaTab();
    }
  });
  await Promise.all([loadSheetPayables(), loadSheetReceivables(), loadSheetFixedExpenses()]);
  rerenderAll(); // 모든 데이터 로드 후 최종 갱신 보장
}

init();

