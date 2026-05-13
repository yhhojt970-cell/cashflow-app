// ════════════════════════════════════════════════════════════
//  미래오토메이션(주) — 현금흐름 관리 Apps Script
//  스프레드시트: https://docs.google.com/spreadsheets/d/1VxYrCD3eZr5PpTORFPCEQPfWM5QSr-tNFNnc_W1C5qM
// ════════════════════════════════════════════════════════════

// ── 스프레드시트 & 인증 ──────────────────────────────────────
const SHEET_ID  = "1VxYrCD3eZr5PpTORFPCEQPfWM5QSr-tNFNnc_W1C5qM";
const API_TOKEN = "miraeautomation2026";

// ── 시트명 ───────────────────────────────────────────────────
const PAYABLES_SHEET       = "미지급_raw";
const RECEIVABLES_SHEET    = "raw";
const MANAGER_SHEET        = "담당자";
const PLAN_SHEET           = "결제계획";
const MASTER_SHEET         = "업체마스터";
const HISTORY_SHEET        = "결제이력";
const UPDATE_HISTORY_SHEET = "업데이트이력";
const TAX_INVOICE_SHEET    = "세금계산서_raw";
const LEDGER_SALES_SHEET   = "계정별원장_매출_raw";
const LEDGER_BUY_SHEET     = "계정별원장_매입_raw";
const LEDGER_PAY_SHEET     = "계정별원장_미지급_raw";
const DAILY_SALES_SHEET    = "영업현황_raw";
const BIZ_DIVISION_SHEET   = "사업부문마스터";
const FIXED_SHEET          = "고정지출";
const PNL_SHEET            = "경영손익_data";

// ── 미수금 이메일 설정 ───────────────────────────────────────
const RCV_MANAGER_EMAIL_MAP = {
  "장운기":"jug@mauto.co.kr", "여희정":"yhj@mauto.co.kr", "김도연":"kdy@mauto.co.kr",
  "남예린":"nyr@mauto.co.kr", "오성철":"osc@mauto.co.kr", "장재영":"jjy@mauto.co.kr",
  "김태홍":"kth@mauto.co.kr", "박희선":"phs@mauto.co.kr", "구예솔":"kys@mauto.co.kr",
  "배지혜":"bjh@mauto.co.kr", "임연하":"lyh@mauto.co.kr",
};
const RCV_ABSENCE_CHAIN = [
  { name:"박희선", email:"phs@mauto.co.kr" },
  { name:"김도연", email:"kdy@mauto.co.kr" },
  { name:"장운기", email:"jug@mauto.co.kr" },
];
const RCV_DEPT_HEAD = { name:"김도연", email:"kdy@mauto.co.kr" };
const RCV_CEO       = { name:"장운기", email:"jug@mauto.co.kr" };

// ════════════════════════════════════════════════════════════
//  인증
// ════════════════════════════════════════════════════════════
function checkAuth(tokenValue) {
  return String(tokenValue || "").trim() === API_TOKEN;
}

// ════════════════════════════════════════════════════════════
//  GET 라우터
// ════════════════════════════════════════════════════════════
function doGet(e) {
  const params = (e && e.parameter) || {};
  if (!checkAuth(params.token)) return jsonOutput({ error: "인증 실패" });
  const action = String(params.action || "").trim();

  // 단순 시트 읽기 (action → 시트명 매핑)
  const GET_SHEET_MAP = {
    getPaymentPlans:  PLAN_SHEET,
    getVendorMaster:  MASTER_SHEET,
    getPaymentHistory:HISTORY_SHEET,
    getReceivables:   RECEIVABLES_SHEET,
    getManagerMaster: MANAGER_SHEET,
    getTaxInvoices:   TAX_INVOICE_SHEET,
    getLedgerSales:   LEDGER_SALES_SHEET,
    getLedgerPurchase:LEDGER_BUY_SHEET,
    getLedgerPayable: LEDGER_PAY_SHEET,
    getDailySales:    DAILY_SALES_SHEET,
    getBizDivision:   BIZ_DIVISION_SHEET,
    getFixed:         FIXED_SHEET,
    getPnlData:       PNL_SHEET,
  };
  if (GET_SHEET_MAP[action]) return jsonOutput({ rows: getSheetRows(GET_SHEET_MAP[action]) });

  // 대사 데이터 일괄 (스프레드시트 1회 오픈으로 속도 최적화)
  if (action === "getDaesaAll") {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    return jsonOutput({
      taxInvoices:    getSheetRows(TAX_INVOICE_SHEET,  ss),
      ledgerSales:    getSheetRows(LEDGER_SALES_SHEET, ss),
      ledgerPurchase: getSheetRows(LEDGER_BUY_SHEET,   ss),
      ledgerPayable:  getSheetRows(LEDGER_PAY_SHEET,   ss),
      dailySales:     getSheetRows(DAILY_SALES_SHEET,  ss),
      bizDivision:    getSheetRows(BIZ_DIVISION_SHEET, ss),
    });
  }

  // 엠오토 JSON (단일 셀)
  if (action === "getMautoData") {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName("엠오토_json");
    const val = sh && sh.getRange("A1").getValue();
    return jsonOutput({ data: val ? JSON.parse(val) : null });
  }

  // 가용자금 JSON (단일 셀)
  if (action === "getAvailableFundsJson") {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName("가용자금_json");
    const val = sh && sh.getRange("B1").getValue();
    if (!val) return jsonOutput({ updatedAt: null, data: null });
    return jsonOutput({
      updatedAt: String(sh.getRange("A1").getValue()),
      data:      JSON.parse(val),
    });
  }

  // action 없음 → 미지급 raw (기본값)
  return jsonOutput({ data: getSheetRows(PAYABLES_SHEET) });
}

// ════════════════════════════════════════════════════════════
//  POST 라우터
// ════════════════════════════════════════════════════════════
function doPost(e) {
  const body   = JSON.parse((e && e.postData && e.postData.contents) || "{}");
  if (!checkAuth(body.token)) return jsonOutput({ error: "인증 실패" });
  const action = String(body.action || "").trim();
  const rows   = Array.isArray(body.rows) ? body.rows : [];

  // ── append 계열 ──
  if (action === "appendPaymentPlans")  { appendRows(PLAN_SHEET,           rows); return jsonOutput({ ok: true, count: rows.length }); }
  if (action === "appendPaymentHistory"){ appendRows(HISTORY_SHEET,        rows); return jsonOutput({ ok: true, count: rows.length }); }
  if (action === "appendUpdateHistory") { appendRows(UPDATE_HISTORY_SHEET, rows); return jsonOutput({ ok: true, count: rows.length }); }

  // ── 업체마스터 (전용 upsert — 거래처코드 앞자리 0 보존) ──
  if (action === "upsertVendorMaster") {
    upsertVendorMasterRows(MASTER_SHEET, rows);
    return jsonOutput({ ok: true, count: rows.length });
  }

  // ── 담당자마스터 ──
  if (action === "upsertManagerMaster") {
    upsertRowsByKey(MANAGER_SHEET, "거래처코드", rows);
    return jsonOutput({ ok: true, count: rows.length });
  }

  // ── 대사 데이터 upsert ──
  if (action === "upsertTaxInvoices") {
    upsertRowsByKey(TAX_INVOICE_SHEET, "_row_key", rows);
    return jsonOutput({ ok: true, count: rows.length });
  }
  if (action === "upsertLedger") {
    const sheetMap = { 매출: LEDGER_SALES_SHEET, 매입: LEDGER_BUY_SHEET, 미지급: LEDGER_PAY_SHEET };
    const sn = sheetMap[body.ledgerType];
    if (!sn) return jsonOutput({ error: "잘못된 ledgerType" });
    upsertRowsByKey(sn, "_row_key", rows);
    return jsonOutput({ ok: true, count: rows.length });
  }
  if (action === "upsertDailySales") {
    upsertRowsByKey(DAILY_SALES_SHEET, "_row_key", rows);
    return jsonOutput({ ok: true, count: rows.length });
  }
  if (action === "upsertBizDivision") {
    upsertRowsByKey(BIZ_DIVISION_SHEET, "_row_key", rows);
    return jsonOutput({ ok: true, count: rows.length });
  }

  // ── 경영손익 저장 (rows 배열 또는 단건 row) ──
  if (action === "savePnlData") {
    const pnlRows = rows.length ? rows : (body.row ? [body.row] : []);
    if (!pnlRows.length) return jsonOutput({ ok: true, count: 0 });
    upsertRowsByKey(PNL_SHEET, "_key", pnlRows);
    return jsonOutput({ ok: true, count: pnlRows.length });
  }

  // ── 엠오토 JSON 저장 ──
  if (action === "saveMautoData") {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sh = ss.getSheetByName("엠오토_json");
    if (!sh) sh = ss.insertSheet("엠오토_json");
    sh.getRange("A1").setValue(JSON.stringify(body.data));
    return jsonOutput({ ok: true });
  }

  // ── 가용자금 JSON 저장 ──
  if (action === "upsertAvailableFunds") {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sh = ss.getSheetByName("가용자금_json");
    if (!sh) sh = ss.insertSheet("가용자금_json");
    sh.getRange("A1").setValue(body.updatedAt || new Date().toISOString());
    sh.getRange("B1").setValue(JSON.stringify(body.data));
    return jsonOutput({ ok: true });
  }

  // ── 이메일 발송 ──
  if (action === "sendReceivableEmails")    return jsonOutput(handleSendReceivableEmails(body));
  if (action === "sendRawDiffEmail")        return jsonOutput(handleSendRawDiffEmail(body));
  if (action === "sendPaymentWarningEmail") return jsonOutput(handleSendPaymentWarningEmail(body));

  return jsonOutput({ error: "지원하지 않는 action 입니다." });
}

// ════════════════════════════════════════════════════════════
//  시트 읽기 헬퍼
// ════════════════════════════════════════════════════════════

/**
 * 시트의 모든 행을 객체 배열로 반환.
 * ss를 전달하면 openById 재호출을 생략할 수 있어 일괄 요청 시 빠름(getDaesaAll 등).
 */
function getSheetRows(sheetName, ss) {
  ss = ss || SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) return [];

  // 헤더 행 자동 감지 (최대 10행)
  const HEADER_KEYWORDS = [
    "거래처코드","코드","code","담당자","거래처명","client",
    "년도","연도","year","vendor_id",
  ];
  let headerIdx = 0;
  for (let i = 0; i < Math.min(10, values.length); i++) {
    const rowStr = values[i].map(v => String(v).trim()).join("");
    if (HEADER_KEYWORDS.some(k => rowStr.includes(k))) { headerIdx = i; break; }
  }

  const headers = values[headerIdx].map(h => String(h).trim());
  return values.slice(headerIdx + 1).map(row => {
    const obj = {};
    row.forEach((val, i) => {
      obj[headers[i]] = (val instanceof Date && !isNaN(val.getTime()))
        ? Utilities.formatDate(val, "Asia/Seoul", "yyyy-MM-dd")
        : val;
    });
    return obj;
  });
}

// ════════════════════════════════════════════════════════════
//  시트 쓰기 헬퍼
// ════════════════════════════════════════════════════════════

/**
 * keyField 기준으로 기존 행은 업데이트, 신규 행은 추가.
 * 전체 데이터를 메모리에 올린 뒤 setValues 2회(기존/신규)로 처리해 API 호출 최소화.
 */
function upsertRowsByKey(sheetName, keyField, rows) {
  if (!rows.length) return;
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  const lastRow = sheet.getLastRow();

  // 앞자리 0 및 아포스트로피 제거 후 비교 (101 / 00101 / '00101 → 동일)
  function normKey(k) {
    return String(k ?? "").trim().replace(/^'+/, "").replace(/^0+(\d)/, "$1");
  }

  // 빈 시트: 헤더 + 전체 데이터 한 번에 쓰기
  if (lastRow === 0) {
    const headers = Object.keys(rows[0]);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(2, 1, rows.length, headers.length)
         .setValues(rows.map(r => headers.map(h => r[h] ?? "")));
    return;
  }

  // 기존 헤더 읽기 + 신규 컬럼 확장
  const lastCol    = sheet.getLastColumn();
  const curHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v).trim());
  const headers    = [...curHeaders];
  Object.keys(rows[0]).forEach(h => { if (!headers.includes(h)) headers.push(h); });
  if (!headers.includes(keyField)) headers.push(keyField);
  if (headers.length !== curHeaders.length)
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // 기존 데이터 메모리 로드
  const curKeyIdx   = curHeaders.indexOf(keyField);
  const existingRaw = lastRow > 1
    ? sheet.getRange(2, 1, lastRow - 1, curHeaders.length).getValues()
    : [];
  const existingData = existingRaw.map(r => headers.map((_, i) => r[i] !== undefined ? r[i] : ""));

  const existingMap = {};
  existingData.forEach((row, idx) => {
    const key = normKey(row[curKeyIdx >= 0 ? curKeyIdx : headers.indexOf(keyField)]);
    if (key) existingMap[key] = idx;
  });

  // 메모리 upsert
  const newRows = [];
  rows.forEach(row => {
    const key = normKey(row[keyField]);
    if (!key) return;
    const vals = headers.map(h => row[h] ?? "");
    if (existingMap[key] !== undefined) {
      existingData[existingMap[key]] = vals;
    } else {
      existingMap[key] = -1;
      newRows.push(vals);
    }
  });

  // 일괄 쓰기
  if (existingData.length)
    sheet.getRange(2, 1, existingData.length, headers.length).setValues(existingData);
  if (newRows.length)
    sheet.getRange(lastRow + 1, 1, newRows.length, headers.length).setValues(newRows);
}

/**
 * 업체마스터 전용 upsert.
 * 복합키: 거래처코드_norm → 사업자번호 → 거래처명.
 * 코드/사업자번호/계좌번호 컬럼에 텍스트 포맷 강제 적용 (앞자리 0 보존).
 */
function upsertVendorMasterRows(sheetName, newRows) {
  if (!newRows.length) return;
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  const lastRow = sheet.getLastRow();

  const TEXT_FORMAT_COLS = ["거래처코드_norm","거래처코드_raw","vendor_id","사업자번호","계좌번호"];

  function normCode(v) { return String(v ?? "").trim().replace(/^0+(\d)/, "$1"); }
  function normBiz(v)  { return String(v ?? "").trim().replace(/[^0-9]/g, ""); }
  function getKey(row) {
    return normCode(row["거래처코드_norm"] || "")
        || normBiz(row["사업자번호"] || "")
        || String(row["거래처명"] || "").trim()
        || "";
  }

  function applyTextFormat(headers, dataLen, startRow) {
    TEXT_FORMAT_COLS.forEach(col => {
      const ci = headers.indexOf(col);
      if (ci >= 0 && dataLen > 0)
        sheet.getRange(startRow, ci + 1, dataLen, 1).setNumberFormat("@");
    });
  }

  // 빈 시트
  if (lastRow === 0) {
    const headers = Object.keys(newRows[0]);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    applyTextFormat(headers, newRows.length, 2);
    sheet.getRange(2, 1, newRows.length, headers.length)
         .setValues(newRows.map(r => headers.map(h => r[h] ?? "")));
    return;
  }

  const lastCol    = sheet.getLastColumn();
  const curHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v).trim());
  const headers    = [...curHeaders];
  Object.keys(newRows[0]).forEach(h => { if (!headers.includes(h)) headers.push(h); });
  if (headers.length !== curHeaders.length)
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const existingRaw  = lastRow > 1
    ? sheet.getRange(2, 1, lastRow - 1, curHeaders.length).getValues()
    : [];
  const existingData = existingRaw.map(r => headers.map((_, i) => r[i] !== undefined ? r[i] : ""));

  const keyMap = new Map();
  existingData.forEach((row, idx) => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    const key = getKey(obj);
    if (key && !keyMap.has(key)) keyMap.set(key, idx);
  });

  const toAppend = [];
  newRows.forEach(row => {
    const key = getKey(row);
    if (!key) return;
    const vals = headers.map(h => row[h] ?? "");
    if (keyMap.has(key)) {
      existingData[keyMap.get(key)] = vals;
    } else {
      keyMap.set(key, -1);
      toAppend.push(vals);
    }
  });

  applyTextFormat(headers, existingData.length, 2);
  applyTextFormat(headers, toAppend.length, lastRow + 1);

  if (existingData.length)
    sheet.getRange(2, 1, existingData.length, headers.length).setValues(existingData);
  if (toAppend.length)
    sheet.getRange(lastRow + 1, 1, toAppend.length, headers.length).setValues(toAppend);
}

/** 시트 맨 아래에 행 추가 */
function appendRows(sheetName, rows) {
  if (!rows.length) return;
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  const lastRow = sheet.getLastRow();

  if (lastRow === 0) {
    const headers = Object.keys(rows[0]);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(2, 1, rows.length, headers.length)
         .setValues(rows.map(r => headers.map(h => r[h] ?? "")));
    return;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(v => String(v).trim());
  sheet.getRange(lastRow + 1, 1, rows.length, headers.length)
       .setValues(rows.map(r => headers.map(h => r[h] ?? "")));
}

/** JSON ContentService 응답 생성 */
function jsonOutput(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════
//  이메일 핸들러: 미수금 발송
// ════════════════════════════════════════════════════════════
function handleSendReceivableEmails(params) {
  const {
    managers=[], absentChain=[], ccEmails=[], conditions=[],
    testMode=false, testRecipient=null, sendSummary=true,
    excludeMinus=false, senderName="",
    previewMode=false, customMessages={},
  } = params;

  const rawData = getSheetRows(RECEIVABLES_SHEET);
  const mgrData = getSheetRows(MANAGER_SHEET);
  const today   = new Date(); today.setHours(0,0,0,0);

  // 담당자 코드 맵 구성
  const mgrMap = {};
  mgrData.forEach(r => {
    const code = String(r["코드"]||r["거래처코드"]||r["code"]||"")
      .replace(/[^0-9]/g,"").replace(/^0+/,"");
    if (code) mgrMap[code] = { manager: r["담당자"]||"", email: r["이메일"]||"" };
  });

  // 수금예정일 계산
  function calcDueDate(year, month, memo, condition) {
    const cond = String(condition||"").replace("전자어음","").trim();
    const ms   = String(memo||"").trim();
    year=Number(year); month=Number(month);
    if (!year||!month) return null;
    if (["바로","쇼핑몰+","오토몰"].includes(cond)) {
      const m=ms.match(/(\d{6})~\?/); if(!m) return null;
      const s=m[1];
      return new Date(2000+parseInt(s.slice(0,2)), parseInt(s.slice(2,4))-1, parseInt(s.slice(4,6)));
    }
    function lastDay(y,m){ return new Date(y,m,0); }
    function add(y,m,n){ const t=m+n; return [y+Math.floor((t-1)/12), ((t-1)%12)+1]; }
    if (cond==="당말일") return lastDay(year,month);
    const cm=cond.match(/^당(\d+)일$/); if(cm){const[ny,nm]=add(year,month,1);return new Date(ny,nm-1,parseInt(cm[1]));}
    if (cond==="25일") {const[ny,nm]=add(year,month,1);return new Date(ny,nm-1,25);}
    if (cond==="말일") {const[ny,nm]=add(year,month,1);return lastDay(ny,nm);}
    if (cond==="60일") {const[ny,nm]=add(year,month,2);return lastDay(ny,nm);}
    const dm=cond.match(/^(\d+)일$/); if(dm){const[ny,nm]=add(year,month,2);return new Date(ny,nm-1,parseInt(dm[1]));}
    return null;
  }

  // 담당자별 그룹 구성
  const condSet = new Set(conditions);
  const groups  = {};
  rawData.forEach(row => {
    const year      = Number(row["연도"]||row["year"]||row["작성연도"]||0);
    const month     = Number(row["월"]||row["month"]||row["작성월"]||0);
    const code      = String(row["코드"]||row["거래처코드"]||row["code"]||"").trim()
                        .replace(/[^0-9]/g,"").replace(/^0+/,"");
    const name      = String(row["거래처명"]||row["client"]||"").trim();
    const memo      = String(row["매출메모"]||row["메모"]||row["memo"]||"").trim();
    const condition = String(row["수금조건"]||row["일"]||row["condition"]||"").trim();
    const balance   = Number(String(row["잔 액"]??row["잔액"]??row["balance"]??0)
                        .replace(/[^0-9.-]/g,"")) || 0;

    if (!name||!balance||condition==="제외"||memo.includes("제외")) return;
    if (condSet.size && !condSet.has(condition)) return;

    const mgr  = mgrMap[code] || { manager:"미지정", email:"" };
    const email= mgr.email || RCV_MANAGER_EMAIL_MAP[mgr.manager] || "";
    if (!email && mgr.manager !== "미지정") return;

    const dueDate = calcDueDate(year, month, memo, condition);
    const elapsed = dueDate ? Math.floor((today-dueDate)/86400000) : null;
    const ym      = year && month ? `${String(year).slice(2)}-${String(month).padStart(2,"0")}` : "";

    if (!groups[mgr.manager]) groups[mgr.manager] = { manager:mgr.manager, email, rows:[] };
    groups[mgr.manager].rows.push({
      name, condition, ym,
      dueDate: dueDate ? Utilities.formatDate(dueDate,"Asia/Seoul","yyyy-MM-dd") : "",
      elapsed, balance, memo, manager: mgr.manager,
    });
  });

  // 부재 연락 체인
  const absentSet = new Set(absentChain||[]);
  function resolveChain() {
    for (const p of RCV_ABSENCE_CHAIN) { if (!absentSet.has(p.name)) return p; }
    return null;
  }

  const cc      = (ccEmails||[]).join(",");
  const testTo  = testRecipient || "yhj@mauto.co.kr";
  const dateStr = Utilities.formatDate(new Date(),"Asia/Seoul","yyyy년 MM월 dd일");
  const TD = "padding:7px 10px;border:1px solid #ddd;white-space:nowrap;";
  const TH = "padding:8px 10px;border:1px solid #1565c0;white-space:nowrap;";

  const senderLine = senderName
    ? `<strong>미래오토메이션(주) 관리부</strong> · ${senderName}`
    : `<strong>미래오토메이션(주) 관리부</strong>`;

  function wrapEmail(innerHtml, customMsg) {
    const box = customMsg
      ? `<div style="margin-bottom:15px;padding:12px;background:#fffde7;border-left:4px solid #fbc02d;color:#333;font-size:14px;white-space:pre-wrap;line-height:1.5;">${customMsg}</div>`
      : "";
    return `<div style="font-family:'맑은 고딕',sans-serif;max-width:900px;margin:0 auto;color:#333;">
      ${box}${innerHtml}
      <p style="margin-top:20px;font-size:12px;color:#888;">본 메일은 자동 발송됩니다.</p>
      <br><p>감사합니다.<br>${senderLine}</p></div>`;
  }

  function buildTableRows(rowList, showManager) {
    let html="", total=0;
    rowList.forEach(r => {
      const el=r.elapsed;
      let bg="", elStyle="color:#333;";
      if      (el>=60){bg="background:#fff0f0;";elStyle="color:#d32f2f;font-weight:bold;";}
      else if (el>=30){bg="background:#fffde7;";elStyle="color:#f57f17;font-weight:bold;";}
      const elLabel = el<0
        ? `<span style="color:#1565c0;">D${el}</span>`
        : `<span style="${elStyle}">${el}일</span>`;
      total += r.balance;
      html += `<tr style="${bg}">
        <td style="${TD}text-align:center;">${r.ym}</td>
        ${showManager ? `<td style="${TD}text-align:center;">${r.manager}</td>` : ""}
        <td style="${TD}">${r.name}</td>
        <td style="${TD}text-align:center;">${r.condition}</td>
        <td style="${TD}text-align:center;">${r.dueDate||"-"}</td>
        <td style="${TD}text-align:center;">${elLabel}</td>
        <td style="${TD}text-align:right;">${r.balance.toLocaleString()}원</td>
        <td style="${TD}font-size:12px;color:#666;">${r.memo}</td>
      </tr>`;
    });
    return { html, total };
  }

  let sentCount=0;
  const previewsOut=[];

  function fireEmail(id, toList, subject, innerHtml) {
    const html = wrapEmail(innerHtml, customMessages[id]||"");
    if (previewMode) {
      previewsOut.push({ id, to: Array.isArray(toList)?toList.join(", "):toList, subject, htmlBody: html });
      return;
    }
    const opts = { htmlBody: html, name:"미래오토메이션(주) 관리부" };
    if (cc) opts.cc = cc;
    (Array.isArray(toList)?toList:[toList]).forEach(t => {
      if (!t) return;
      GmailApp.sendEmail(t, subject, "HTML 형식", opts);
      sentCount++;
    });
  }

  // 담당자별 개별 발송
  managers.filter(m=>!m.absent).forEach(({ manager }) => {
    const group = groups[manager];
    if (!group?.rows.length) return;
    const { html, total } = buildTableRows(group.rows, false);
    const to = testMode ? testTo : (group.email || RCV_MANAGER_EMAIL_MAP[manager] || "");
    if (!to) return;
    const subject = (testMode?"[테스트] ":"") + `[미래오토메이션] ${manager} 담당자 미수금 현황 안내`;
    fireEmail("mgr_"+manager, to, subject, `
      <p>${dateStr} 기준 담당 미수금 현황을 안내드립니다.</p>
      <table style="border-collapse:collapse;width:100%;font-size:13px;margin-top:12px;">
        <thead><tr style="background:#1565c0;color:white;">
          <th style="${TH}">매출연월</th><th style="${TH}text-align:left;">거래처명</th>
          <th style="${TH}">수금조건</th><th style="${TH}">수금예정일</th>
          <th style="${TH}">경과일수</th><th style="${TH}">잔액</th><th style="${TH}">메모</th>
        </tr></thead><tbody>${html}</tbody>
        <tfoot><tr style="background:#e3f2fd;font-weight:bold;">
          <td colspan="5" style="padding:8px 10px;border:1px solid #ddd;text-align:right;">합 계</td>
          <td style="padding:8px 10px;border:1px solid #ddd;text-align:right;">${total.toLocaleString()}원</td>
          <td style="padding:8px 10px;border:1px solid #ddd;"></td>
        </tr></tfoot>
      </table>`);
  });

  // 부재자 대리 통합 발송
  const absentMgrs  = managers.filter(m=>m.absent);
  const chainTarget = resolveChain();
  if (absentMgrs.length && chainTarget) {
    let combinedHtml="", combinedTotal=0;
    absentMgrs.forEach(({ manager }) => {
      if (!groups[manager]?.rows.length) return;
      const { html, total } = buildTableRows(groups[manager].rows, true);
      combinedHtml += `
        <h4 style="margin-top:20px;margin-bottom:8px;color:#1565c0;border-bottom:2px solid #1565c0;padding-bottom:4px;font-size:14px;">👤 담당자: ${manager}</h4>
        <table style="border-collapse:collapse;width:100%;font-size:13px;">
          <thead><tr style="background:#1565c0;color:white;">
            <th style="${TH}">매출연월</th><th style="${TH}">담당자</th><th style="${TH}text-align:left;">거래처명</th>
            <th style="${TH}">수금조건</th><th style="${TH}">수금예정일</th>
            <th style="${TH}">경과일수</th><th style="${TH}">잔액</th><th style="${TH}">메모</th>
          </tr></thead><tbody>${html}</tbody>
          <tfoot><tr style="background:#e3f2fd;font-weight:bold;">
            <td colspan="6" style="padding:8px 10px;border:1px solid #ddd;text-align:right;">${manager} 합계</td>
            <td style="padding:8px 10px;border:1px solid #ddd;text-align:right;">${total.toLocaleString()}원</td>
            <td style="padding:8px 10px;border:1px solid #ddd;"></td>
          </tr></tfoot>
        </table>`;
      combinedTotal += total;
    });
    if (combinedHtml) {
      const mgrLabel = absentMgrs.length===1
        ? absentMgrs[0].manager
        : `${absentMgrs[0].manager} 외 ${absentMgrs.length-1}명`;
      fireEmail("absent", testMode?testTo:chainTarget.email,
        (testMode?"[테스트] ":"") + `[미래오토메이션] ${mgrLabel} 담당자 미수금 현황 안내 (부재 대리 수신)`,
        `<p style="color:#7b1fa2;background:#f3e5f5;padding:10px 14px;border-left:4px solid #7b1fa2;">
          ※ 부재 담당자(${mgrLabel}) 대리 수신 — ${chainTarget.name}님께 통합 발송</p>
        <p>${dateStr} 기준 부재 담당자의 미수금 현황을 안내드립니다.</p>
        ${combinedHtml}
        <div style="margin-top:20px;padding:12px;background:#e3f2fd;border:1px solid #90caf9;font-weight:bold;text-align:right;font-size:15px;color:#0d47a1;">
          부재자 총 합계: ${combinedTotal.toLocaleString()}원
        </div>`);
    }
  }

  // 전체 현황 보고서
  if (sendSummary) {
    const allRows  = Object.values(groups).flatMap(g=>g.rows);
    const filtered = excludeMinus ? allRows.filter(r=>(r.elapsed||0)>=0) : allRows;
    filtered.sort((a,b)=>(a.dueDate||"").localeCompare(b.dueDate||""));
    const { html, total } = buildTableRows(filtered, true);
    const note    = excludeMinus ? " (D- 제외)" : "";
    const toList  = testMode ? [testTo] : (params.summaryRecipients || [RCV_DEPT_HEAD.email, RCV_CEO.email]);
    fireEmail("summary", toList,
      (testMode?"[테스트] ":"") + "[미래오토메이션] 미수금 현황 보고",
      `<p>${dateStr} 기준 전체 미수금 현황을 보고드립니다.${note}</p>
      <table style="border-collapse:collapse;width:100%;font-size:13px;margin-top:12px;">
        <thead><tr style="background:#1565c0;color:white;">
          <th style="${TH}">매출연월</th><th style="${TH}">담당자</th><th style="${TH}text-align:left;">거래처명</th>
          <th style="${TH}">수금조건</th><th style="${TH}">수금예정일</th>
          <th style="${TH}">경과일수</th><th style="${TH}">잔액</th><th style="${TH}">메모</th>
        </tr></thead><tbody>${html}</tbody>
        <tfoot><tr style="background:#e3f2fd;font-weight:bold;">
          <td colspan="6" style="padding:8px 10px;border:1px solid #ddd;text-align:right;">총 합계${note}</td>
          <td style="padding:8px 10px;border:1px solid #ddd;text-align:right;">${total.toLocaleString()}원</td>
          <td style="padding:8px 10px;border:1px solid #ddd;"></td>
        </tr></tfoot>
      </table>`);
  }

  return previewMode ? { ok:true, previews:previewsOut } : { ok:true, sentCount };
}

// ════════════════════════════════════════════════════════════
//  이메일 핸들러: 결제 경고 (은행정보 누락)
// ════════════════════════════════════════════════════════════
function handleSendPaymentWarningEmail(params) {
  const { warnings=[], planLabel="", recipients=[], testMode=false } = params;
  if (!warnings.length || !recipients.length) return { ok:false, error:"데이터 없음" };

  const testTo  = "yhj@mauto.co.kr";
  const dateStr = Utilities.formatDate(new Date(),"Asia/Seoul","yyyy년 MM월 dd일 HH:mm");
  const TD = "padding:7px 10px;border:1px solid #ddd;";
  const TH = "padding:8px 10px;border:1px solid #b45309;color:white;text-align:left;";

  const bodyHtml = `
    <div style="font-family:'맑은 고딕',sans-serif;max-width:700px;margin:0 auto;color:#333;">
      <p style="background:#fff3cd;border-left:4px solid #f59e0b;padding:10px 14px;font-size:14px;">
        ⚠️ [${planLabel}] 확인이 필요한 항목이 발견되었습니다.<br>
        아래 업체의 은행정보를 ERP '거래처정보 관리'에 등록해주세요.
      </p>
      <p style="color:#555;font-size:13px;">${dateStr} 기준</p>
      <table style="border-collapse:collapse;width:100%;font-size:13px;margin-top:8px;">
        <thead><tr style="background:#b45309;"><th style="${TH}">거래처명</th><th style="${TH}">누락 항목</th></tr></thead>
        <tbody>${warnings.map(w=>`<tr>
          <td style="${TD}">${w.거래처명||""}</td>
          <td style="${TD}color:#b45309;font-weight:bold;">${(w.missing||[]).join(", ")}</td>
        </tr>`).join("")}</tbody>
      </table>
      <br>
      <p style="font-size:12px;color:#888;">본 메일은 현금흐름 관리 앱에서 자동 발송됩니다.</p>
      <p>감사합니다.<br><strong>미래오토메이션(주) 관리부</strong></p>
    </div>`;

  const subject = (testMode?"[테스트] ":"") +
    `[미래오토메이션] ${planLabel} 결제 보고서 — 은행정보 확인 요청 (${warnings.length}건)`;
  let sentCount=0;
  recipients.forEach(r => {
    const to = testMode ? testTo : r.email;
    if (!to) return;
    GmailApp.sendEmail(to, subject, "HTML 형식 메일입니다.", { htmlBody:bodyHtml, name:"미래오토메이션(주) 관리부" });
    sentCount++;
  });
  return { ok:true, sentCount };
}

// ════════════════════════════════════════════════════════════
//  이메일 핸들러: 미지급 변경 감지
// ════════════════════════════════════════════════════════════
function handleSendRawDiffEmail(params) {
  const { diff=[], recipients=[], testMode=false } = params;
  if (!diff.length || !recipients.length) return { ok:false, error:"데이터 없음" };

  const testTo  = "yhj@mauto.co.kr";
  const dateStr = Utilities.formatDate(new Date(),"Asia/Seoul","yyyy년 MM월 dd일 HH:mm");
  const TD = "padding:7px 10px;border:1px solid #ddd;";
  const TH = "padding:8px 10px;border:1px solid #1e40af;color:white;";

  function buildSection(title, color, rows) {
    if (!rows.length) return "";
    return `
      <h4 style="color:${color};margin:16px 0 6px;">${title} (${rows.length}건)</h4>
      <table style="border-collapse:collapse;width:100%;font-size:13px;">
        <thead><tr style="background:${color};">
          <th style="${TH}text-align:left;">항목</th>
          <th style="${TH}">이전 금액</th>
          <th style="${TH}">변경 금액</th>
        </tr></thead>
        <tbody>${rows.map(d=>`<tr>
          <td style="${TD}">${d.label||d.stableKey||""}</td>
          <td style="${TD}text-align:right;">${d.prevAmount!=null?Number(d.prevAmount).toLocaleString()+"원":"-"}</td>
          <td style="${TD}text-align:right;font-weight:bold;">${d.newAmount!=null?Number(d.newAmount).toLocaleString()+"원":"-"}</td>
        </tr>`).join("")}</tbody>
      </table>`;
  }

  const bodyHtml = `
    <div style="font-family:'맑은 고딕',sans-serif;max-width:800px;margin:0 auto;color:#333;">
      <p style="background:#fff3cd;border-left:4px solid #f59e0b;padding:10px 14px;font-size:14px;">
        ⚠️ 미지급_raw 시트 업데이트 시 기존 결제 계획과 충돌이 발생한 항목이 있습니다.<br>
        내용을 확인하고 앱에서 <strong>확인 후 적용</strong> 버튼을 눌러주세요.
      </p>
      <p style="color:#555;font-size:13px;">${dateStr} 기준 감지된 변경사항입니다.</p>
      ${buildSection("🗑 사라진 항목 (완료 처리 권장)", "#7f1d1d", diff.filter(d=>d.type==="removed"))}
      ${buildSection("✏️ 금액 변경 항목", "#1e3a8a", diff.filter(d=>d.type==="changed"))}
      <br>
      <p style="font-size:12px;color:#888;">본 메일은 현금흐름 관리 앱에서 자동 발송됩니다.</p>
      <p>감사합니다.<br><strong>미래오토메이션(주) 관리부</strong></p>
    </div>`;

  const subject = (testMode?"[테스트] ":"") + "[미래오토메이션] 미지급 데이터 변경 확인 요청";
  let sentCount=0;
  recipients.forEach(r => {
    const to = testMode ? testTo : r.email;
    if (!to) return;
    GmailApp.sendEmail(to, subject, "HTML 형식 메일입니다.", { htmlBody:bodyHtml, name:"미래오토메이션(주) 관리부" });
    sentCount++;
  });
  return { ok:true, sentCount };
}
