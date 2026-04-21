# Google Apps Script 연동 가이드

## 1. 현재 기준
- 스프레드시트 ID: `1VxYrCD3eZr5PpTORFPCEQPfWM5QSr-tNFNnc_W1C5qM`
- 원본 시트: `미지급_raw`
- 앱에서 읽는 URL: `app.js`의 `SHEET_APP_SCRIPT_URL`

현재 프론트는 `미지급_raw`를 읽어 와서 화면에 보여주고,  
이번 1단계부터는 아래 3개 시트를 추가로 쓰는 구조를 기준으로 합니다.

- `업체마스터`
- `결제계획`
- `결제이력`

## 2. 시트 구성

### `미지급_raw`
원본 데이터를 그대로 유지합니다.

권장 컬럼:
- `거래처코드`
- `거래처명`
- `작성연도`
- `작성월`
- `합계`
- `지급합`
- `잔액`
- `지급일`
- `메모`

### `업체마스터`
거래처 엑셀 업로드 대상입니다.

권장 컬럼:
- `vendor_id`
- `거래처코드_raw`
- `거래처코드_norm`
- `거래처명`
- `거래처분류`
- `거래처구분`
- `대표자명`
- `사업자번호`
- `전화번호`
- `팩스번호`
- `주소`
- `업태`
- `종목`
- `홈페이지`
- `은행`
- `계좌번호`
- `예금주`
- `active_yn`
- `last_imported_at`
- `last_changed_at`
- `change_note`

업로드 동작 기준:
- 거래처 파일 전체를 읽습니다.
- 하지만 프론트에서는 `현재 미지급_raw`에 있는 `거래처코드_norm`만 추려서 비교합니다.
- 즉, 파일에 4000건이 있어도 현재 미지급 대상 업체만 `신규 / 변경 / 동일` 비교 후 저장합니다.
- 이렇게 해야 실제 결제 대상 업체만 빠르게 검토하고 반영할 수 있습니다.

### `결제계획`
새로고침해도 유지되어야 하는 계획 저장용입니다.

권장 컬럼:
- `source_key`
- `거래처코드_norm`
- `거래처명`
- `작성연도`
- `작성월`
- `원금액`
- `잔액`
- `decision_amount`
- `payment_plan`
- `plan_status`
- `memo`
- `updated_at`
- `updated_by`

`payment_plan` 예시:
- `2026-04-30`
- `보류`
- 빈값(`미정`)

`plan_status` 예시:
- `미정`
- `예정`
- `보류`
- `부분결제`
- `완료`

### `결제이력`
실제 이체 후 누적 저장합니다.

권장 컬럼:
- `history_id`
- `source_key`
- `거래처코드_norm`
- `거래처명`
- `지급일자`
- `지급금액`
- `은행`
- `계좌번호`
- `예금주`
- `적요`
- `결과상태`
- `created_at`
- `created_by`

## 3. 핵심 키 규칙

### 거래처코드 정규화
엑셀에서는 `00101`, 구글시트에서는 `101`처럼 보일 수 있으므로  
비교용 코드는 항상 문자열로 정규화합니다.

예:
- `101` -> `00101`
- `00101` -> `00101`

프론트에서는 `normalizeVendorCode()`로 처리합니다.

### source_key
원본 미지급 한 줄을 식별하는 키입니다.

현재 프론트 기준 조합:
- `거래처코드_norm`
- `작성연도`
- `작성월`
- `합계`
- `납기그룹`
- `지급일`
- `메모`

즉:
```text
거래처코드_norm||작성연도||작성월||합계||납기그룹||지급일||메모
```

이 키로 `결제계획`, `결제이력`, 향후 `부분결제`를 연결합니다.

## 4. Apps Script 기본 코드
아래 코드를 Google Sheets의 `Extensions > Apps Script`에 넣고 시작하면 됩니다.

```javascript
const SHEET_ID = "1VxYrCD3eZr5PpTORFPCEQPfWM5QSr-tNFNnc_W1C5qM";
const API_TOKEN = "miraeautomation2026"; // 예: "mj9Kp2xQrL8wZn4vBt7hYcAs"
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
const BIZ_DIVISION_SHEET = "사업부문마스터"; // 이 줄을 추가하세요.
const FIXED_SHEET        = "고정지출";        // ← 고정지출 추가



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
  // 아래 줄을 추가하세요.
  if (action === "getBizDivision")    return jsonOutput({ rows: getSheetRows(BIZ_DIVISION_SHEET) });
  if (action === "getFixed")          return jsonOutput({ rows: getSheetRows(FIXED_SHEET) });   // ← 고정지출 추가


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
    upsertRowsByKey(MASTER_SHEET, "거래처코드_norm", Array.isArray(body.rows) ? body.rows : []);
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
  
  // 아래 이 부분을 추가하세요.
  if (action === "upsertBizDivision") {
    upsertRowsByKey(BIZ_DIVISION_SHEET, "_row_key", Array.isArray(body.rows) ? body.rows : []);
    return jsonOutput({ ok: true, count: (body.rows||[]).length });
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
    const code = String(r["거래처코드"]||"").trim().replace(/^0+/,"");
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
    const code=codeRaw.replace(/^0+/,"");
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
    groups[mgr.manager].rows.push({ name, condition, ym, dueDate:dueDateStr, elapsed, balance, memo });
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

  function buildRows(rowList) {
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
  function wrapEmail(body) {
    return `<div style="font-family:'맑은 고딕',sans-serif;max-width:900px;margin:0 auto;color:#333;">${body}
      <p style="font-size:12px;color:#888;">본 메일은 자동 발송됩니다.</p>
      <br><p>감사합니다.<br>${senderLine}</p></div>`;
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
    const body = wrapEmail(`<p>${dateStr} 기준 담당 미수금 현황을 안내드립니다.</p>
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
      </table>`);
    const opts = { htmlBody:body, name:"미래오토메이션(주) 관리부" };
    if (cc) opts.cc = cc;
    GmailApp.sendEmail(to, subject, "본 메일은 HTML 형식입니다.", opts);
    sentCount++;
  });

  // 부재자 통합 발송
  const chainTarget = resolveChain();
  if (absentManagers.length && chainTarget) {
    let combinedHtml = "";
    let combinedTotal = 0;
    
    absentManagers.forEach(({ manager }) => {
      if (groups[manager] && groups[manager].rows.length) {
        const { html, total } = buildRows(groups[manager].rows);
        combinedHtml += `
          <h4 style="margin-top:20px;margin-bottom:8px;color:#1565c0;border-bottom:2px solid #1565c0;padding-bottom:4px;font-size:14px;">👤 담당자: ${manager}</h4>
          <table style="border-collapse:collapse;width:100%;font-size:13px;">
          <thead><tr style="background:#1565c0;color:white;">
            <th style="${th}">매출연월</th><th style="${th}text-align:left;">거래처명</th>
            <th style="${th}">수금조건</th><th style="${th}">수금예정일</th>
            <th style="${th}">경과일수</th><th style="${th}">잔액</th><th style="${th}">메모</th>
          </tr></thead><tbody>${html}</tbody>
          <tfoot><tr style="background:#e3f2fd;font-weight:bold;">
            <td colspan="5" style="padding:8px 10px;border:1px solid #ddd;text-align:right;">${manager} 합계</td>
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
      const body = wrapEmail(`
        <p style="color:#7b1fa2;background:#f3e5f5;padding:10px 14px;border-left:4px solid #7b1fa2;">
          ※ 부재 담당자(${mgrLabel}) 대리 수신 — ${chainTarget.name}님께 통합 발송</p>
        <p>${dateStr} 기준 부재 담당자의 미수금 현황을 안내드립니다.</p>
        ${combinedHtml}
        <div style="margin-top:20px;padding:12px;background:#e3f2fd;border:1px solid #90caf9;font-weight:bold;text-align:right;font-size:15px;color:#0d47a1;">
          부재자 총 합계: ${combinedTotal.toLocaleString()}원
        </div>`);
      const opts = { htmlBody:body, name:"미래오토메이션(주) 관리부" };
      if (cc) opts.cc = cc;
      GmailApp.sendEmail(to, subject, "본 메일은 HTML 형식입니다.", opts);
      sentCount++;
    }
  }

  // 전체 현황 보고서
  if (sendSummary) {
    let combinedHtml = "";
    let grandTotal = 0;
    
    // 담당자 이름순으로 정렬해서 개별 표 생성
    const allManagers = Object.keys(groups).sort();
    
    allManagers.forEach(manager => {
      let mgrRows = groups[manager].rows;
      if (excludeMinus) mgrRows = mgrRows.filter(r => (r.elapsed||0) >= 0);
      if (mgrRows.length === 0) return;
      
      mgrRows.sort((a,b) => (a.dueDate||"").localeCompare(b.dueDate||""));
      const { html, total } = buildRows(mgrRows);
      
      combinedHtml += `
          <h4 style="margin-top:24px;margin-bottom:8px;color:#0d47a1;border-bottom:2px solid #0d47a1;padding-bottom:4px;font-size:14px;">👤 담당자: ${manager}</h4>
          <table style="border-collapse:collapse;width:100%;font-size:13px;">
          <thead><tr style="background:#1565c0;color:white;">
            <th style="${th}">매출연월</th><th style="${th}text-align:left;">거래처명</th>
            <th style="${th}">수금조건</th><th style="${th}">수금예정일</th>
            <th style="${th}">경과일수</th><th style="${th}">잔액</th><th style="${th}">메모</th>
          </tr></thead><tbody>${html}</tbody>
          <tfoot><tr style="background:#e3f2fd;font-weight:bold;">
            <td colspan="5" style="padding:8px 10px;border:1px solid #ddd;text-align:right;">${manager} 합계</td>
            <td style="padding:8px 10px;border:1px solid #ddd;text-align:right;">${total.toLocaleString()}원</td>
            <td style="padding:8px 10px;border:1px solid #ddd;"></td>
          </tr></tfoot>
          </table>
      `;
      grandTotal += total;
    });

    if (combinedHtml) {
      const subject = (testMode?"[테스트] ":"") + "[미래오토메이션] 미수금 현황 보고";
      const excludeNote = excludeMinus ? " (D- 제외)" : "";
      const body = wrapEmail(`
        <p style="font-size:14px;color:#333;">${dateStr} 기준 전체 미수금 현황을 보고드립니다.${excludeNote}</p>
        ${combinedHtml}
        <div style="margin-top:24px;padding:15px;background:#e8eaf6;border:2px solid #3f51b5;font-weight:bold;text-align:right;font-size:16px;color:#1a237e;">
          총 합계${excludeNote}: ${grandTotal.toLocaleString()}원
        </div>
      `);
      const opts = { htmlBody:body, name:"미래오토메이션(주) 관리부" };
      if (cc) opts.cc = cc;
      GmailApp.sendEmail(testMode ? testTo : RCV_DEPT_HEAD.email, subject, "HTML 형식", opts);
      sentCount++;
      if (!testMode) {
        GmailApp.sendEmail(RCV_CEO.email, subject, "HTML 형식", opts);
        sentCount++;
      }
    }
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

function getSheetRows(sheetName) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) return [];

  const headers = values[0].map(header => String(header).trim());
  return values.slice(1).map(row => {
    const item = {};
    row.forEach((value, index) => {
      // ★ Date 객체를 KST 기준 YYYY-MM-DD 문자열로 변환
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

  if (lastRow === 0) {
    const headers = Object.keys(rows[0]);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const body = rows.map(row => headers.map(header => row[header] ?? ""));
    sheet.getRange(2, 1, body.length, headers.length).setValues(body);
    return;
  }

  const lastCol = sheet.getLastColumn();
  const currentHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v).trim());
  const incomingHeaders = Object.keys(rows[0]);
  const headers = [...currentHeaders];

  incomingHeaders.forEach(header => {
    if (!headers.includes(header)) headers.push(header);
  });

  if (headers.length !== currentHeaders.length) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  const keyIndex = headers.indexOf(keyField);
  if (keyIndex === -1) {
    headers.push(keyField);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  const existingData = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, currentHeaders.length).getValues() : [];
  const existingMap = {};
  existingData.forEach((row, index) => {
    const rowObject = {};
    currentHeaders.forEach((header, i) => { rowObject[header] = row[i]; });
    const key = String(rowObject[keyField] || "").trim();
    if (key) existingMap[key] = index + 2;
  });

  rows.forEach(row => {
    const key = String(row[keyField] || "").trim();
    if (!key) return;
    const values = headers.map(header => row[header] ?? "");
    if (existingMap[key]) {
      sheet.getRange(existingMap[key], 1, 1, headers.length).setValues([values]);
    } else {
      sheet.appendRow(values);
    }
  });
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
```

## 5. 다음 단계에서 추가할 Apps Script 함수
1단계를 더 완성하려면 아래 함수들을 이어서 추가하면 됩니다.

- `getVendorMaster()`
- `upsertVendorMaster(rows)`
- `appendPaymentHistory(rows)`
- `upsertPaymentPlans(rows)`
- `appendPaymentHistory(rows)`
- `getSavedPaymentPlans()`
- `getPaymentHistory()`

추천 방식:
- `doGet`: 조회
- `doPost`: 저장

예시 payload:

```json
{
  "action": "upsertPaymentPlans",
  "rows": [
    {
      "source_key": "00101||2026||03||30000000||말일||말일||26-02-?",
      "거래처코드_norm": "00101",
      "거래처명": "예시업체",
      "decision_amount": 12000000,
      "payment_plan": "2026-04-30",
      "plan_status": "예정"
    }
  ]
}
```

## 6. 배포 방법

### Apps Script 배포
1. Apps Script에서 `Deploy > New deployment`
2. `Web app` 선택
3. `Execute as`: `Me`
4. `Who has access`: `Anyone` (토큰으로 보호됨 — Google 로그인 불필요)
5. 배포 후 Web App URL 복사
6. `app.js`의 `SHEET_APP_SCRIPT_URL`에 반영

### GitHub Pages 배포 (다중 사용자 접근)

1. **GitHub 계정 준비** — github.com에서 계정 생성 (없는 경우)

2. **저장소 생성**
   - `New repository` 클릭
   - 이름 예: `cashflow-app` (Private 권장)
   - `Add a README file` 체크 후 생성

3. **파일 업로드**
   ```
   index.html
   app.js
   style.css
   ```
   저장소 > `Add file > Upload files`로 3개 파일 업로드

4. **GitHub Pages 활성화**
   - 저장소 > `Settings > Pages`
   - `Source`: `Deploy from a branch`
   - `Branch`: `main` / `/ (root)` 선택 후 `Save`
   - 몇 분 후 `https://[계정명].github.io/cashflow-app/` 주소 생성

5. **앱 접속 및 토큰 입력**
   - GitHub Pages 주소로 접속
   - 헤더의 **🔑 토큰** 버튼 클릭
   - Apps Script의 `API_TOKEN`과 동일한 값 입력 → 저장
   - 이후 해당 기기에서는 자동 사용됨

6. **다른 직원 공유**
   - GitHub Pages URL을 공유
   - 각자 기기에서 토큰 한 번 입력하면 사용 가능
   - Private 저장소여도 Pages는 공개 접근 가능 (URL 아는 사람만 접속)

> **보안 참고**: 토큰은 각 기기의 localStorage에 저장됩니다.  
> 토큰을 모르는 사람은 시트 데이터에 접근할 수 없습니다.  
> 토큰 유출 시 Apps Script의 `API_TOKEN`만 바꾸고 재배포하면 됩니다.

## 7. 주의사항
- 프론트는 현재 로컬 저장도 같이 사용합니다.
- 로컬 저장은 임시 안전장치이고, 최종 기준 데이터는 시트 저장이어야 합니다.
- `결제계획`과 `결제이력`은 원본 `미지급_raw`를 직접 수정하지 말고 별도 시트로 누적하는 방식이 안전합니다.
