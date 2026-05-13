const SHEET_ID = "1VxYrCD3eZr5PpTORFPCEQPfWM5QSr-tNFNnc_W1C5qM";
const API_TOKEN = "miraeautomation2026";
const PAYABLES_SHEET    = "미지급_raw";
const RECEIVABLES_SHEET = "raw";
const MANAGER_SHEET     = "담당자";
const PLAN_SHEET        = "결제계획";
const MASTER_SHEET      = "업체마스터";
const HISTORY_SHEET     = "결제이력";
const UPDATE_HISTORY_SHEET = "업데이트이력";
const TAX_INVOICE_SHEET  = "세금계산서_raw";
const LEDGER_SALES_SHEET = "계정별원장_매출_raw";
const LEDGER_BUY_SHEET   = "계정별원장_매입_raw";
const LEDGER_PAY_SHEET   = "계정별원장_미지급_raw";
const DAILY_SALES_SHEET  = "영업현황_raw";
const BIZ_DIVISION_SHEET = "사업부문마스터";
const FIXED_SHEET        = "고정지출";
const PNL_SHEET          = "경영손익_data";

function checkAuth(tokenValue) {
  return String(tokenValue || "").trim() === API_TOKEN;
}

// ── 미수금 이메일 설정 ──────────────────────────────────────
const RCV_MANAGER_EMAIL_MAP = {
  "장운기":"jug@mauto.co.kr","여희정":"yhj@mauto.co.kr","김도연":"kdy@mauto.co.kr",
  "남예린":"nyr@mauto.co.kr","오성철":"osc@mauto.co.kr","장재영":"jjy@mauto.co.kr",
  "김태홍":"kth@mauto.co.kr","박희선":"phs@mauto.co.kr","구예솔":"kys@mauto.co.kr",
  "배지혜":"bjh@mauto.co.kr","임연하":"lyh@mauto.co.kr",
};
const RCV_ABSENCE_CHAIN = [
  { name:"박희선", email:"phs@mauto.co.kr" },
  { name:"김도연", email:"kdy@mauto.co.kr" },
  { name:"장운기", email:"jug@mauto.co.kr" },
];
const RCV_DEPT_HEAD = { name:"김도연", email:"kdy@mauto.co.kr" };
const RCV_CEO       = { name:"장운기", email:"jug@mauto.co.kr" };

function doGet(e) {
  const params = (e && e.parameter) || {};
  if (!checkAuth(params.token)) return jsonOutput({ error: "인증 실패" });
  const action = String(params.action || "").trim();
  if (action === "getPaymentPlans")  return jsonOutput({ rows: getSheetRows(PLAN_SHEET) });
  if (action === "getVendorMaster")  return jsonOutput({ rows: getSheetRows(MASTER_SHEET) });
  if (action === "getPaymentHistory")return jsonOutput({ rows: getSheetRows(HISTORY_SHEET) });
  if (action === "getReceivables")   return jsonOutput({ rows: getSheetRows(RECEIVABLES_SHEET) });
  if (action === "getManagerMaster") return jsonOutput({ rows: getSheetRows(MANAGER_SHEET) });
  if (action === "getTaxInvoices")    return jsonOutput({ rows: getSheetRows(TAX_INVOICE_SHEET) });
  if (action === "getLedgerSales")    return jsonOutput({ rows: getSheetRows(LEDGER_SALES_SHEET) });
  if (action === "getLedgerPurchase") return jsonOutput({ rows: getSheetRows(LEDGER_BUY_SHEET) });
  if (action === "getLedgerPayable")  return jsonOutput({ rows: getSheetRows(LEDGER_PAY_SHEET) });
  if (action === "getDailySales")     return jsonOutput({ rows: getSheetRows(DAILY_SALES_SHEET) });
  if (action === "getBizDivision")    return jsonOutput({ rows: getSheetRows(BIZ_DIVISION_SHEET) });
  if (action === "getFixed")          return jsonOutput({ rows: getSheetRows(FIXED_SHEET) });
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
  if (action === "getPnlData") {
    return jsonOutput({ rows: getSheetRows(PNL_SHEET) });
  }
  if (action === "getMautoData") {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName("엠오토_json");
    if (!sh || !sh.getRange("A1").getValue())
      return jsonOutput({ data: null });
    return jsonOutput({ data: JSON.parse(sh.getRange("A1").getValue()) });
  }
  if (action === "getAvailableFundsJson") {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName("가용자금_json");
    if (!sh || !sh.getRange("B1").getValue())
      return jsonOutput({ updatedAt: null, data: null });
    return jsonOutput({
      updatedAt: String(sh.getRange("A1").getValue()),
      data: JSON.parse(sh.getRange("B1").getValue())
    });
  }
  return jsonOutput({ data: getSheetRows(PAYABLES_SHEET) });
}

function doPost(e) {
  const body = JSON.parse((e && e.postData && e.postData.contents) || "{}");
  if (!checkAuth(body.token)) return jsonOutput({ error: "인증 실패" });
  const action = String(body.action || "").trim();

  if (action === "appendPaymentPlans") {
    appendRows(PLAN_SHEET, Array.isArray(body.rows) ? body.rows : []);
    return jsonOutput({ ok: true, count: (body.rows||[]).length });
  }
  if (action === "upsertVendorMaster") {
    upsertVendorMasterRows(MASTER_SHEET, Array.isArray(body.rows) ? body.rows : []);
    return jsonOutput({ ok: true, count: (body.rows||[]).length });
  }
  if (action === "appendPaymentHistory") {
    appendRows(HISTORY_SHEET, Array.isArray(body.rows) ? body.rows : []);
    return jsonOutput({ ok: true, count: (body.rows||[]).length });
  }
  if (action === "upsertManagerMaster") {
    upsertRowsByKey(MANAGER_SHEET, "거래처코드", Array.isArray(body.rows) ? body.rows : []);
    return jsonOutput({ ok: true, count: (body.rows||[]).length });
  }
  if (action === "appendUpdateHistory") {
    appendRows(UPDATE_HISTORY_SHEET, Array.isArray(body.rows) ? body.rows : []);
    return jsonOutput({ ok: true, count: (body.rows||[]).length });
  }
  if (action === "sendReceivableEmails") {
    return jsonOutput(handleSendReceivableEmails(body));
  }
  if (action === "sendRawDiffEmail") {
    return jsonOutput(handleSendRawDiffEmail(body));
  }
  if (action === "sendPaymentWarningEmail") {
    return jsonOutput(handleSendPaymentWarningEmail(body));
  }
  if (action === "upsertTaxInvoices") {
    upsertRowsByKey(TAX_INVOICE_SHEET, "_row_key", Array.isArray(body.rows) ? body.rows : []);
    return jsonOutput({ ok: true, count: (body.rows||[]).length });
  }
  if (action === "upsertLedger") {
    const sheetMap = { 매출: LEDGER_SALES_SHEET, 매입: LEDGER_BUY_SHEET, 미지급: LEDGER_PAY_SHEET };
    const sn = sheetMap[body.ledgerType];
    if (!sn) return jsonOutput({ error: "잘못된 ledgerType" });
    upsertRowsByKey(sn, "_row_key", Array.isArray(body.rows) ? body.rows : []);
    return jsonOutput({ ok: true, count: (body.rows||[]).length });
  }
  if (action === "upsertDailySales") {
    upsertRowsByKey(DAILY_SALES_SHEET, "_row_key", Array.isArray(body.rows) ? body.rows : []);
    return jsonOutput({ ok: true, count: (body.rows||[]).length });
  }
  if (action === "upsertBizDivision") {
    upsertRowsByKey(BIZ_DIVISION_SHEET, "_row_key", Array.isArray(body.rows) ? body.rows : []);
    return jsonOutput({ ok: true, count: (body.rows||[]).length });
  }
  if (action === "savePnlData") {
    const rows = Array.isArray(body.rows) ? body.rows : (body.row ? [body.row] : []);
    if (!rows.length) return jsonOutput({ ok: false, error: "no rows" });
    upsertRowsByKey(PNL_SHEET, "_key", rows);
    return jsonOutput({ ok: true, count: rows.length });
  }
  if (action === "saveMautoData") {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sh = ss.getSheetByName("엠오토_json");
    if (!sh) sh = ss.insertSheet("엠오토_json");
    sh.getRange("A1").setValue(JSON.stringify(body.data));
    return jsonOutput({ ok: true });
  }
  if (action === "upsertAvailableFunds") {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sh = ss.getSheetByName("가용자금_json");
    if (!sh) sh = ss.insertSheet("가용자금_json");
    sh.getRange("A1").setValue(body.updatedAt || new Date().toISOString());
    sh.getRange("B1").setValue(JSON.stringify(body.data));
    return jsonOutput({ success: true });
  }
  return jsonOutput({ error: "지원하지 않는 action 입니다." });
}

// ── 미수금 이메일 발송 핸들러 ───────────────────────────────
function handleSendReceivableEmails(params) {
  const { managers=[], absentChain=[], ccEmails=[], conditions=[],
          testMode=false, testRecipient=null, sendSummary=true, excludeMinus=false,
          senderName="" } = params;

  const rawData  = getSheetRows(RECEIVABLES_SHEET);
  const mgrData  = getSheetRows(MANAGER_SHEET);
  const today    = new Date(); today.setHours(0,0,0,0);

  // 담당자 맵 구성
  const mgrMap = {};
  mgrData.forEach(r => {
    const codeRaw = String(r["코드"]||r["거래처코드"]||r["code"]||"");
    const code = codeRaw.replace(/[^0-9]/g, "").replace(/^0+/,"");
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
      const s=m[1]; return new Date(2000+parseInt(s.slice(0,2)),parseInt(s.slice(2,4))-1,parseInt(s.slice(4,6)));
    }
    function lastDay(y,m){return new Date(y,m,0);}
    function add(y,m,n){const t=m+n;return[y+Math.floor((t-1)/12),((t-1)%12)+1];}
    if(cond==="당말일") return lastDay(year,month);
    const cm=cond.match(/^당(\d+)일$/); if(cm){const[ny,nm]=add(year,month,1);return new Date(ny,nm-1,parseInt(cm[1]));}
    if(cond==="25일"){const[ny,nm]=add(year,month,1);return new Date(ny,nm-1,25);}
    if(cond==="말일"){const[ny,nm]=add(year,month,1);return lastDay(ny,nm);}
    if(cond==="60일"){const[ny,nm]=add(year,month,2);return lastDay(ny,nm);}
    const dm=cond.match(/^(\d+)일$/); if(dm){const[ny,nm]=add(year,month,2);return new Date(ny,nm-1,parseInt(dm[1]));}
    return null;
  }

  // 조건 필터 셋
  const condSet = new Set(conditions);

  // 담당자별 rows 구성
  const groups = {};
  rawData.forEach(row => {
    const year=Number(row["연도"]||row["year"]||row["작성연도"]||0);
    const month=Number(row["월"]||row["month"]||row["작성월"]||0);
    const codeRaw=String(row["코드"]||row["거래처코드"]||row["code"]||"").trim();
    const code=codeRaw.replace(/[^0-9]/g, "").replace(/^0+/,"");
    const name=String(row["거래처명"]||row["client"]||"").trim();
    const memo=String(row["매출메모"]||row["메모"]||row["memo"]||"").trim();
    const condition=String(row["수금조건"]||row["일"]||row["condition"]||"").trim();
    const balanceRaw=row["잔 액"]??row["잔액"]??row["balance"]??0;
    const balance=Number(String(balanceRaw).replace(/[^0-9.-]/g,""))||0;
    if (!name || !balance || condition==="제외" || memo.includes("제외")) return;
    if (condSet.size && !condSet.has(condition)) return;
    const mgr = mgrMap[code] || { manager:"미지정", email:"" };
    const email = mgr.email || RCV_MANAGER_EMAIL_MAP[mgr.manager] || "";
    if (!email && mgr.manager !== "미지정") return;
    const dueDate = calcDueDate(year, month, memo, condition);
    const elapsed = dueDate ? Math.floor((today-dueDate)/86400000) : null;
    const dueDateStr = dueDate ? Utilities.formatDate(dueDate,"Asia/Seoul","yyyy-MM-dd") : "";
    const ym = year && month ? `${String(year).slice(2)}-${String(month).padStart(2,"0")}` : "";
    if (!groups[mgr.manager]) groups[mgr.manager] = { manager:mgr.manager, email, rows:[] };
    groups[mgr.manager].rows.push({ name, condition, ym, dueDate:dueDateStr, elapsed, balance, memo, manager: mgr.manager });
  });

  const absentSet = new Set(absentChain||[]);
  function resolveChain() {
    for (const p of RCV_ABSENCE_CHAIN) { if (!absentSet.has(p.name)) return p; }
    return null;
  }

  const cc       = (ccEmails||[]).join(",");
  const testTo   = testRecipient || "yhj@mauto.co.kr";
  const td       = "padding:7px 10px;border:1px solid #ddd;white-space:nowrap;";
  const th       = "padding:8px 10px;border:1px solid #1565c0;white-space:nowrap;";
  const dateStr  = Utilities.formatDate(new Date(),"Asia/Seoul","yyyy년 MM월 dd일");
  let sentCount  = 0;

  function buildRows(rowList, showManager = false) {
    let html="", total=0;
    rowList.forEach(r => {
      const el=r.elapsed;
      let bg="", elStyle="color:#333;";
      if(el>=60){bg="background:#fff0f0;";elStyle="color:#d32f2f;font-weight:bold;";}
      else if(el>=30){bg="background:#fffde7;";elStyle="color:#f57f17;font-weight:bold;";}
      const elLabel = el<0 ? `<span style="color:#1565c0;">D${el}</span>`
                            : `<span style="${elStyle}">${el}일</span>`;
      total+=r.balance;
      html+=`<tr style="${bg}">
        <td style="${td}text-align:center;">${r.ym}</td>
        ${showManager ? `<td style="${td}text-align:center;">${r.manager}</td>` : ""}
        <td style="${td}">${r.name}</td>
        <td style="${td}text-align:center;">${r.condition}</td>
        <td style="${td}text-align:center;">${r.dueDate||"-"}</td>
        <td style="${td}text-align:center;">${elLabel}</td>
        <td style="${td}text-align:right;">${r.balance.toLocaleString()}원</td>
        <td style="${td}font-size:12px;color:#666;">${r.memo}</td>
      </tr>`;
    });
    return { html, total };
  }

  const senderLine = senderName ? `<strong>미래오토메이션(주) 관리부</strong> · ${senderName}` : `<strong>미래오토메이션(주) 관리부</strong>`;
  function wrapEmail(body, customMessage = "") {
    const customHtml = customMessage ? `<div style="margin-bottom:15px;padding:12px;background:#fffde7;border-left:4px solid #fbc02d;color:#333;font-size:14px;white-space:pre-wrap;line-height:1.5;">${customMessage}</div>` : "";
    return `<div style="font-family:'맑은 고딕',sans-serif;max-width:900px;margin:0 auto;color:#333;">
      ${customHtml}
      ${body}
      <p style="margin-top:20px;font-size:12px;color:#888;">본 메일은 자동 발송됩니다.</p>
      <br><p>감사합니다.<br>${senderLine}</p></div>`;
  }

  const previewMode = !!params.previewMode;
  const customMsgs = params.customMessages || {};
  const previewsOut = [];

  function fireEmail(id, toList, subject, innerBody) {
    const customMsg = customMsgs[id] || "";
    const finalHtml = wrapEmail(innerBody, customMsg);
    if (previewMode) {
      previewsOut.push({ id, to: Array.isArray(toList) ? toList.join(", ") : toList, subject, htmlBody: finalHtml });
    } else {
      const opts = { htmlBody: finalHtml, name:"미래오토메이션(주) 관리부" };
      if (cc) opts.cc = cc;
      const targets = Array.isArray(toList) ? toList : [toList];
      targets.forEach(t => {
        if (!t) return;
        GmailApp.sendEmail(t, subject, "HTML 형식", opts);
        sentCount++;
      });
    }
  }

  // 담당자별 발송
  const normalManagers = managers.filter(m => !m.absent);
  const absentManagers = managers.filter(m =>  m.absent);

  normalManagers.forEach(({ manager }) => {
    const group = groups[manager]; if (!group || !group.rows.length) return;
    const { html, total } = buildRows(group.rows);
    const to = testMode ? testTo : (group.email || RCV_MANAGER_EMAIL_MAP[manager] || "");
    if (!to) return;
    const subject = (testMode?"[테스트] ":"") + `[미래오토메이션] ${manager} 담당자 미수금 현황 안내`;
    const innerBody = `<p>${dateStr} 기준 담당 미수금 현황을 안내드립니다.</p>
      <table style="border-collapse:collapse;width:100%;font-size:13px;margin-top:12px;">
        <thead><tr style="background:#1565c0;color:white;">
          <th style="${th}">매출연월</th><th style="${th}text-align:left;">거래처명</th>
          <th style="${th}">수금조건</th><th style="${th}">수금예정일</th>
          <th style="${th}">경과일수</th><th style="${th}">잔액</th><th style="${th}">메모</th>
        </tr></thead><tbody>${html}</tbody>
        <tfoot><tr style="background:#e3f2fd;font-weight:bold;">
          <td colspan="5" style="padding:8px 10px;border:1px solid #ddd;text-align:right;">합 계</td>
          <td style="padding:8px 10px;border:1px solid #ddd;text-align:right;">${total.toLocaleString()}원</td>
          <td style="padding:8px 10px;border:1px solid #ddd;"></td>
        </tr></tfoot>
      </table>`;
    fireEmail("mgr_" + manager, to, subject, innerBody);
  });

  // 부재자 통합 발송
  const chainTarget = resolveChain();
  if (absentManagers.length && chainTarget) {
    let combinedHtml = "";
    let combinedTotal = 0;

    absentManagers.forEach(({ manager }) => {
      if (groups[manager] && groups[manager].rows.length) {
        const { html, total } = buildRows(groups[manager].rows, true);
        combinedHtml += `
          <h4 style="margin-top:20px;margin-bottom:8px;color:#1565c0;border-bottom:2px solid #1565c0;padding-bottom:4px;font-size:14px;">👤 담당자: ${manager}</h4>
          <table style="border-collapse:collapse;width:100%;font-size:13px;">
          <thead><tr style="background:#1565c0;color:white;">
            <th style="${th}">매출연월</th><th style="${th}">담당자</th><th style="${th}text-align:left;">거래처명</th>
            <th style="${th}">수금조건</th><th style="${th}">수금예정일</th>
            <th style="${th}">경과일수</th><th style="${th}">잔액</th><th style="${th}">메모</th>
          </tr></thead><tbody>${html}</tbody>
          <tfoot><tr style="background:#e3f2fd;font-weight:bold;">
            <td colspan="6" style="padding:8px 10px;border:1px solid #ddd;text-align:right;">${manager} 합계</td>
            <td style="padding:8px 10px;border:1px solid #ddd;text-align:right;">${total.toLocaleString()}원</td>
            <td style="padding:8px 10px;border:1px solid #ddd;"></td>
          </tr></tfoot>
          </table>
        `;
        combinedTotal += total;
      }
    });

    if (combinedHtml) {
      const mgrLabel = absentManagers.length===1 ? absentManagers[0].manager
        : `${absentManagers[0].manager} 외 ${absentManagers.length-1}명`;
      const to = testMode ? testTo : chainTarget.email;
      const subject = (testMode?"[테스트] ":"") +
        `[미래오토메이션] ${mgrLabel} 담당자 미수금 현황 안내 (부재 대리 수신)`;
      const innerBody = `
        <p style="color:#7b1fa2;background:#f3e5f5;padding:10px 14px;border-left:4px solid #7b1fa2;">
          ※ 부재 담당자(${mgrLabel}) 대리 수신 — ${chainTarget.name}님께 통합 발송</p>
        <p>${dateStr} 기준 부재 담당자의 미수금 현황을 안내드립니다.</p>
        ${combinedHtml}
        <div style="margin-top:20px;padding:12px;background:#e3f2fd;border:1px solid #90caf9;font-weight:bold;text-align:right;font-size:15px;color:#0d47a1;">
          부재자 총 합계: ${combinedTotal.toLocaleString()}원
        </div>`;
      fireEmail("absent", to, subject, innerBody);
    }
  }

  // 전체 현황 보고서
  if (sendSummary) {
    const allRows = Object.values(groups).flatMap(g => g.rows);
    const filtered = excludeMinus ? allRows.filter(r => (r.elapsed||0) >= 0) : allRows;
    filtered.sort((a,b) => (a.dueDate||"").localeCompare(b.dueDate||""));
    const { html, total } = buildRows(filtered, true);

    const subject = (testMode?"[테스트] ":"") + "[미래오토메이션] 미수금 현황 보고";
    const excludeNote = excludeMinus ? " (D- 제외)" : "";
    const innerBody = `<p>${dateStr} 기준 전체 미수금 현황을 보고드립니다.${excludeNote}</p>
      <table style="border-collapse:collapse;width:100%;font-size:13px;margin-top:12px;">
        <thead><tr style="background:#1565c0;color:white;">
          <th style="${th}">매출연월</th><th style="${th}">담당자</th><th style="${th}text-align:left;">거래처명</th>
          <th style="${th}">수금조건</th><th style="${th}">수금예정일</th>
          <th style="${th}">경과일수</th><th style="${th}">잔액</th><th style="${th}">메모</th>
        </tr></thead><tbody>${html}</tbody>
        <tfoot><tr style="background:#e3f2fd;font-weight:bold;">
          <td colspan="6" style="padding:8px 10px;border:1px solid #ddd;text-align:right;">총 합계${excludeNote}</td>
          <td style="padding:8px 10px;border:1px solid #ddd;text-align:right;">${total.toLocaleString()}원</td>
          <td style="padding:8px 10px;border:1px solid #ddd;"></td>
        </tr></tfoot>
      </table>`;

    const toList = testMode ? [testTo] : (params.summaryRecipients || [RCV_DEPT_HEAD.email, RCV_CEO.email]);
    fireEmail("summary", toList, subject, innerBody);
  }

  if (previewMode) {
    return { ok: true, previews: previewsOut };
  }

  return { ok: true, sentCount };
}

// ── 은행 업로드 전 확인 요청 이메일 ─────────────────────────
function handleSendPaymentWarningEmail(params) {
  const { warnings=[], planLabel="", recipients=[], testMode=false } = params;
  if (!warnings.length || !recipients.length) return { ok: false, error: "데이터 없음" };

  const testTo = "yhj@mauto.co.kr";
  const dateStr = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy년 MM월 dd일 HH:mm");
  const td = "padding:7px 10px;border:1px solid #ddd;";
  const th = "padding:8px 10px;border:1px solid #b45309;color:white;text-align:left;";

  const rows = warnings.map(w =>
    `<tr>
      <td style="${td}">${w.거래처명 || ""}</td>
      <td style="${td}color:#b45309;font-weight:bold;">${(w.missing||[]).join(", ")}</td>
    </tr>`
  ).join("");

  const bodyHtml = `
    <div style="font-family:'맑은 고딕',sans-serif;max-width:700px;margin:0 auto;color:#333;">
      <p style="background:#fff3cd;border-left:4px solid #f59e0b;padding:10px 14px;font-size:14px;">
        ⚠️ [${planLabel}] 확인이 필요한 항목이 발견되었습니다.<br>
        아래 업체의 은행정보를 ERP '거래처정보 관리'에 등록해주세요.
      </p>
      <p style="color:#555;font-size:13px;">${dateStr} 기준</p>
      <table style="border-collapse:collapse;width:100%;font-size:13px;margin-top:8px;">
        <thead><tr style="background:#b45309;">
          <th style="${th}">거래처명</th>
          <th style="${th}">누락 항목</th>
        </tr></thead>
        <tbody>${rows}</tbody>
      </table>
      <br>
      <p style="font-size:12px;color:#888;">본 메일은 현금흐름 관리 앱에서 자동 발송됩니다.</p>
      <p>감사합니다.<br><strong>미래오토메이션(주) 관리부</strong></p>
    </div>`;

  const subject = (testMode ? "[테스트] " : "") +
    `[미래오토메이션] ${planLabel} 결제 보고서 — 은행정보 확인 요청 (${warnings.length}건)`;
  let sentCount = 0;
  recipients.forEach(r => {
    const to = testMode ? testTo : r.email;
    if (!to) return;
    GmailApp.sendEmail(to, subject, "HTML 형식 메일입니다.", {
      htmlBody: bodyHtml,
      name: "미래오토메이션(주) 관리부",
    });
    sentCount++;
  });
  return { ok: true, sentCount };
}

// ── 미지급 변경 확인 요청 이메일 ────────────────────────────
function handleSendRawDiffEmail(params) {
  const { diff=[], recipients=[], testMode=false } = params;
  if (!diff.length || !recipients.length) return { ok: false, error: "데이터 없음" };

  const testTo = "yhj@mauto.co.kr";
  const dateStr = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy년 MM월 dd일 HH:mm");
  const td = "padding:7px 10px;border:1px solid #ddd;";
  const th = "padding:8px 10px;border:1px solid #1e40af;color:white;";

  const removedItems = diff.filter(d => d.type === "removed");
  const changedItems = diff.filter(d => d.type === "changed");

  function buildSection(title, color, rows) {
    if (!rows.length) return "";
    return `
      <h4 style="color:${color};margin:16px 0 6px;">${title} (${rows.length}건)</h4>
      <table style="border-collapse:collapse;width:100%;font-size:13px;">
        <thead><tr style="background:${color};">
          <th style="${th}text-align:left;">항목</th>
          <th style="${th}">이전 금액</th>
          <th style="${th}">변경 금액</th>
        </tr></thead>
        <tbody>
          ${rows.map(d => `<tr>
            <td style="${td}">${d.label || d.stableKey || ""}</td>
            <td style="${td}text-align:right;">${d.prevAmount != null ? Number(d.prevAmount).toLocaleString()+"원" : "-"}</td>
            <td style="${td}text-align:right;font-weight:bold;">${d.newAmount != null ? Number(d.newAmount).toLocaleString()+"원" : "-"}</td>
          </tr>`).join("")}
        </tbody>
      </table>`;
  }

  const bodyHtml = `
    <div style="font-family:'맑은 고딕',sans-serif;max-width:800px;margin:0 auto;color:#333;">
      <p style="background:#fff3cd;border-left:4px solid #f59e0b;padding:10px 14px;font-size:14px;">
        ⚠️ 미지급_raw 시트 업데이트 시 기존 결제 계획과 충돌이 발생한 항목이 있습니다.<br>
        내용을 확인하고 앱에서 <strong>확인 후 적용</strong> 버튼을 눌러주세요.
      </p>
      <p style="color:#555;font-size:13px;">${dateStr} 기준 감지된 변경사항입니다.</p>
      ${buildSection("🗑 사라진 항목 (완료 처리 권장)", "#7f1d1d", removedItems)}
      ${buildSection("✏️ 금액 변경 항목", "#1e3a8a", changedItems)}
      <br>
      <p style="font-size:12px;color:#888;">본 메일은 현금흐름 관리 앱에서 자동 발송됩니다.</p>
      <p>감사합니다.<br><strong>미래오토메이션(주) 관리부</strong></p>
    </div>`;

  const subject = (testMode ? "[테스트] " : "") + "[미래오토메이션] 미지급 데이터 변경 확인 요청";
  let sentCount = 0;
  recipients.forEach(r => {
    const to = testMode ? testTo : r.email;
    if (!to) return;
    GmailApp.sendEmail(to, subject, "HTML 형식 메일입니다.", {
      htmlBody: bodyHtml,
      name: "미래오토메이션(주) 관리부",
    });
    sentCount++;
  });

  return { ok: true, sentCount };
}

function getSheetRows(sheetName, ss) {
  ss = ss || SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) return [];
  let headerIdx = 0;
  for (let i = 0; i < Math.min(10, values.length); i++) {
    const rowStr = values[i].map(v => String(v).trim()).join("");
    if (rowStr.includes("거래처코드") || rowStr.includes("코드") || rowStr.includes("code") || rowStr.includes("담당자") || rowStr.includes("거래처명") || rowStr.includes("client") || rowStr.includes("년도") || rowStr.includes("연도") || rowStr.includes("year") || rowStr.includes("vendor_id")) {
      headerIdx = i;
      break;
    }
  }
  const headers = values[headerIdx].map(header => String(header).trim());
  return values.slice(headerIdx + 1).map(row => {
    const item = {};
    row.forEach((value, index) => {
      if (value instanceof Date && !isNaN(value.getTime())) {
        item[headers[index]] = Utilities.formatDate(value, "Asia/Seoul", "yyyy-MM-dd");
      } else {
        item[headers[index]] = value;
      }
    });
    return item;
  });
}

function upsertRowsByKey(sheetName, keyField, rows) {
  if (!rows.length) return;
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  const lastRow = sheet.getLastRow();

  // 아포스트로피 제거 + 앞자리 0 제거 후 비교 (101 / 00101 / '00101 모두 동일하게 처리)
  function normKey(k) {
    return String(k ?? "").trim().replace(/^'+/, "").replace(/^0+(\d)/, "$1");
  }

  // 시트가 비어있으면 헤더+전체 데이터 한 번에 쓰기
  if (lastRow === 0) {
    const headers = Object.keys(rows[0]);
    const body = rows.map(row => headers.map(h => row[h] ?? ""));
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(2, 1, body.length, headers.length).setValues(body);
    return;
  }

  const lastCol = sheet.getLastColumn();
  const currentHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v).trim());
  const incomingHeaders = Object.keys(rows[0]);
  const headers = [...currentHeaders];

  incomingHeaders.forEach(h => { if (!headers.includes(h)) headers.push(h); });

  if (headers.length !== currentHeaders.length) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  if (headers.indexOf(keyField) === -1) {
    headers.push(keyField);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  // 기존 데이터를 메모리에 한 번에 로드
  const curKeyIdx = currentHeaders.indexOf(keyField);
  const existingRaw = lastRow > 1
    ? sheet.getRange(2, 1, lastRow - 1, currentHeaders.length).getValues()
    : [];

  // headers 길이로 패딩하면서 Map 구성
  const existingData = existingRaw.map(r => headers.map((_, i) => (r[i] !== undefined ? r[i] : "")));
  const existingMap = {};
  existingData.forEach((row, idx) => {
    const key = normKey(row[curKeyIdx !== -1 ? curKeyIdx : headers.indexOf(keyField)]);
    if (key) existingMap[key] = idx;
  });

  // 메모리에서 upsert: API 호출 없이 배열만 수정
  const newRows = [];
  rows.forEach(row => {
    const key = normKey(row[keyField]);
    if (!key) return;
    const values = headers.map(h => row[h] ?? "");
    if (existingMap[key] !== undefined) {
      existingData[existingMap[key]] = values;  // 기존 행 메모리 업데이트
    } else {
      newRows.push(values);
    }
  });

  // 일괄 쓰기: 기존 행 전체를 setValues 한 번으로 처리
  if (existingData.length > 0) {
    sheet.getRange(2, 1, existingData.length, headers.length).setValues(existingData);
  }
  // 일괄 쓰기: 새 행 추가도 setValues 한 번으로 처리
  if (newRows.length > 0) {
    sheet.getRange(lastRow + 1, 1, newRows.length, headers.length).setValues(newRows);
  }
}

// 업체마스터 전용 upsert: 거래처코드_norm → 사업자번호 → 거래처명 복합키
function upsertVendorMasterRows(sheetName, newRows) {
  if (!newRows.length) return;
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  const lastRow = sheet.getLastRow();

  function normCode(v) { return String(v ?? "").trim().replace(/^0+(\d)/, "$1"); }
  function normBiz(v)  { return String(v ?? "").trim().replace(/[^0-9]/g, ""); }
  function getKey(row) {
    const code = normCode(row["거래처코드_norm"] || "");
    const biz  = normBiz(row["사업자번호"] || "");
    const name = String(row["거래처명"] || "").trim();
    return code || biz || name || "";
  }

  if (lastRow === 0) {
    const headers = Object.keys(newRows[0]);
    const body = newRows.map(row => headers.map(h => row[h] ?? ""));
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    ["거래처코드_norm", "거래처코드_raw", "vendor_id", "사업자번호", "계좌번호"].forEach(col => {
      const ci = headers.indexOf(col);
      if (ci >= 0) sheet.getRange(2, ci + 1, body.length, 1).setNumberFormat("@");
    });
    sheet.getRange(2, 1, body.length, headers.length).setValues(body);
    return;
  }

  const lastCol = sheet.getLastColumn();
  const currentHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v).trim());
  const incomingHeaders = Object.keys(newRows[0]);
  const headers = [...currentHeaders];
  incomingHeaders.forEach(h => { if (!headers.includes(h)) headers.push(h); });
  if (headers.length !== currentHeaders.length)
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const existingRaw = lastRow > 1
    ? sheet.getRange(2, 1, lastRow - 1, currentHeaders.length).getValues()
    : [];
  const existingData = existingRaw.map(r => headers.map((_, i) => r[i] !== undefined ? r[i] : ""));

  // 기존 행을 복합키로 인덱싱
  const existingKeyMap = new Map();
  existingData.forEach((row, idx) => {
    const rowObj = {};
    headers.forEach((h, i) => { rowObj[h] = row[i]; });
    const key = getKey(rowObj);
    if (key && !existingKeyMap.has(key)) existingKeyMap.set(key, idx);
  });

  const toAppend = [];
  newRows.forEach(row => {
    const key = getKey(row);
    if (!key) return;
    const values = headers.map(h => row[h] ?? "");
    if (existingKeyMap.has(key)) {
      existingData[existingKeyMap.get(key)] = values;
    } else {
      existingKeyMap.set(key, -1);
      toAppend.push(values);
    }
  });

  ["거래처코드_norm", "거래처코드_raw", "vendor_id", "사업자번호", "계좌번호"].forEach(col => {
    const ci = headers.indexOf(col);
    if (ci >= 0) {
      if (existingData.length > 0) sheet.getRange(2, ci + 1, existingData.length, 1).setNumberFormat("@");
      if (toAppend.length > 0) sheet.getRange(lastRow + 1, ci + 1, toAppend.length, 1).setNumberFormat("@");
    }
  });
  if (existingData.length > 0)
    sheet.getRange(2, 1, existingData.length, headers.length).setValues(existingData);
  if (toAppend.length > 0)
    sheet.getRange(lastRow + 1, 1, toAppend.length, headers.length).setValues(toAppend);
}

function appendRows(sheetName, rows) {
  if (!rows.length) return;
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  const lastRow = sheet.getLastRow();

  if (lastRow === 0) {
    const headers = Object.keys(rows[0]);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const body = rows.map(row => headers.map(header => row[header] ?? ""));
    sheet.getRange(2, 1, body.length, headers.length).setValues(body);
    return;
  }

  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v).trim());
  const body = rows.map(row => headers.map(header => row[header] ?? ""));
  sheet.getRange(lastRow + 1, 1, body.length, headers.length).setValues(body);
}

function jsonOutput(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
