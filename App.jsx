import { useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import emailjs from "@emailjs/browser";
import { sendPasswordResetEmail } from "firebase/auth";
import { auth } from "./firebase";
import {
  subscribeAccounts, addAccount, updateAccount,
  deleteAccount, deleteAllAccounts, writeLog, subscribeLogs, subscribeFavorites,
  setFavorite, seedAdminProfile, checkIsAdmin, subscribeAdmins,
  subscribeCategories, addCategory, updateCategory, deleteCategory,
  recordUserLogin, subscribeUserProfiles, subscribeUserProfile, addUserProfile, deleteUserProfile,
  updateUserProfile, addBulkUserProfiles, subscribeUserPermissions,
  subscribeAllUserPermissions, setUserPermissions as saveUserPermissions,
  subscribeDeptPermissions, setDeptPermissions, emailToKey, departmentToKey,
  addPasswordHistory, subscribePasswordHistory,
  subscribeUserColVis, saveUserColVis,
  subscribeCashAccess, subscribeAllCashAccess, setCashAccess as saveCashAccess,
  subscribeContacts, addContact, updateContact, deleteContact,
  subscribeNotices, addNotice, updateNotice, deleteNotice,
  subscribeCompanyInfo, saveCompanyInfo, uploadCompanyAttachment, deleteCompanyAttachment,
} from "./db";
import { loginWithCompanyId, logout, subscribeAuth } from "./auth";
import CashManager from "./CashManager";

const FF = "'Pretendard','Nanum Gothic',sans-serif";
const ADMIN_EMAILS = ["yhj@mauto.co.kr"];
const EMAILJS_SERVICE_ID = "service_w360xqf";
const EMAILJS_TEMPLATE_ID = "template_22pj15i";
const EMAILJS_PUBLIC_KEY = "g0OBqy80ODWftj-V2";
const DEFAULT_COMPANY_INFO = {
  name: "미래오토메이션(주)",
  businessNo: "606-86-34210",
  intro: "",
  address: "(46721) 부산광역시 강서구 유통단지1로 50, 214동 110,111,210,211호 (대저2동, 부산티플렉스)",
  addressEn: "214-110, 50, Yutongdanji 1-ro, Gangseo-gu, Busan, 46721, Republic of Korea",
  factoryAddress: "(46720) 부산광역시 강서구 공항로767번다길 17 (대저2동)",
  phone: "051-322-1765",
  fax: "051-326-1765",
  faxAdmin: "051-326-1766",
  email: "mrat@mauto.co.kr",
  emailTax: "tax@mauto.co.kr",
  emailMall: "mall@mauto.co.kr",
  emailAdmin: "admin@mauto.co.kr",
  accountMain: "우리)1005-204-413062 / 국민) 927437-01-013302 / 신한) 100-035-572826",
  accountSub: "",
  accountMall: "국민) 927437-01-016415",
  cargoKyungdong: "부산강서대저2동3153A지점",
  cargoDaesin: "서부산유통영업소",
  erpInternal: "192.168.0.25 채널2",
  erpExternal: "121.175.14.212",
  serpInstallUrl: "https://serp2.webcash.co.kr/board/ns_fil_list2.jsp?spage=&subSelect=&FILE_PATH=&FILE_NM=",
  serpKeyboardUrl: "https://serp2.webcash.co.kr/board/ns_man_list2.jsp?spage=&FILE_PATH=&FILE_NM=&clsf_cd1=2&blbr_id=72&pst_srno=&subSelect=#none",
  fileBizLicense: [
    "사업자등록증_미래오토메이션(주)_메모★.png",
    "사업자등록증_미래오토메이션(주)_메모★.pdf",
    "사업자등록증_미래오토메이션_주_원본★.pdf",
  ].join("\n"),
  fileBankbook: [
    "미래오토메이션(주)_통장사본_우리1005-204-413062.pdf",
    "미래오토메이션(주)_통장사본_국민927437-01-013302.pdf",
    "미래오토메이션(주)_통장사본_신한100-035-572826.pdf",
    "미래오토메이션(주)_통장사본_FA몰_국민927437-01-016415.pdf",
  ].join("\n"),
  homepage: "",
  hours: "",
};

emailjs.init({
  publicKey: EMAILJS_PUBLIC_KEY,
});



const CAT_STYLE = {
  "은행": { color: "#075985", bg: "#e0f2fe", border: "#7dd3fc" },
  "세금/회계": { color: "#9a3412", bg: "#ffedd5", border: "#fdba74" },
  "쇼핑몰": { color: "#831843", bg: "#fce7f3", border: "#f9a8d4" },
  "정부/공공": { color: "#166534", bg: "#dcfce7", border: "#86efac" },
  "물류": { color: "#5b21b6", bg: "#ede9fe", border: "#c4b5fd" },
  "광고/마케팅": { color: "#b45309", bg: "#fef3c7", border: "#fcd34d" },
  "기타": { color: "#334155", bg: "#f1f5f9", border: "#cbd5e1" },
};

function CategoriesModal({ categories, onClose }) {
  const [newName, setNewName] = useState("");
  const [editId, setEditId] = useState(null);
  const [editName, setEditName] = useState("");
  const [saving, setSaving] = useState(false);

  async function handleAdd() {
    if (!newName.trim()) return;
    setSaving(true);
    try {
      await addCategory(newName.trim());
      setNewName("");
    } finally {
      setSaving(false);
    }
  }

  async function handleUpdate(id) {
    if (!editName.trim()) return;
    setSaving(true);
    try {
      await updateCategory(id, editName.trim());
      setEditId(null);
      setEditName("");
    } finally {
      setSaving(false);
    }
  }

  async function handleDelete(id, name) {
    if (!window.confirm(`"${name}" 카테고리를 삭제할까요?\n해당 카테고리의 계정은 "기타"로 유지됩니다.`)) return;
    await deleteCategory(id);
  }

  return (
    <Modal title="카테고리 관리" onClose={onClose}>
      {/* 추가 */}
      <div style={{ display: "flex", gap: 8, marginBottom: 20 }}>
        <input
          value={newName}
          onChange={e => setNewName(e.target.value)}
          onKeyDown={e => e.key === "Enter" && handleAdd()}
          placeholder="새 카테고리 이름"
          style={{ ...base.input, flex: 1 }}
        />
        <button
          onClick={handleAdd}
          disabled={saving || !newName.trim()}
          style={{
            padding: "10px 16px", borderRadius: 10, border: "none",
            background: newName.trim() ? "#2563eb" : "#e2e8f0",
            color: newName.trim() ? "#fff" : "#94a3b8",
            fontWeight: 800, cursor: newName.trim() ? "pointer" : "not-allowed",
            fontFamily: FF, whiteSpace: "nowrap",
          }}
        >
          + 추가
        </button>
      </div>

      {/* 목록 */}
      <div style={{ display: "grid", gap: 8 }}>
        {categories.length === 0 && (
          <div style={{ color: "#cbd5e1", fontSize: 14, padding: "12px 0", textAlign: "center" }}>
            카테고리가 없습니다. 위에서 추가해주세요.
          </div>
        )}
        {categories.map((cat, idx) => (
          <div key={cat.id} style={{
            border: "1px solid #f1f5f9", borderRadius: 12, padding: "12px 14px",
            background: "#fafafa", display: "flex", alignItems: "center", gap: 10,
          }}>
            <span style={{ fontSize: 13, color: "#cbd5e1", fontWeight: 700, width: 20 }}>{idx + 1}</span>

            {editId === cat.id ? (
              <>
                <input
                  value={editName}
                  onChange={e => setEditName(e.target.value)}
                  onKeyDown={e => e.key === "Enter" && handleUpdate(cat.id)}
                  autoFocus
                  style={{ ...base.input, flex: 1, padding: "7px 10px", fontSize: 13 }}
                />
                <button
                  onClick={() => handleUpdate(cat.id)}
                  disabled={saving}
                  style={{ padding: "7px 12px", borderRadius: 8, border: "none", background: "#2563eb", color: "#fff", fontWeight: 800, cursor: "pointer", fontFamily: FF, fontSize: 13 }}
                >
                  저장
                </button>
                <button
                  onClick={() => { setEditId(null); setEditName(""); }}
                  style={{ padding: "7px 12px", borderRadius: 8, border: "1.5px solid #e2e8f0", background: "#fff", fontWeight: 700, cursor: "pointer", fontFamily: FF, fontSize: 13 }}
                >
                  취소
                </button>
              </>
            ) : (
              <>
                <span style={{ flex: 1, fontWeight: 700, fontSize: 14, color: "#0f172a" }}>{cat.name}</span>
                <button
                  onClick={() => { setEditId(cat.id); setEditName(cat.name); }}
                  style={{ padding: "6px 12px", borderRadius: 8, border: "1.5px solid #3b82f6", background: "#eff6ff", color: "#2563eb", fontWeight: 700, cursor: "pointer", fontFamily: FF, fontSize: 13 }}
                >
                  수정
                </button>
                <button
                  onClick={() => handleDelete(cat.id, cat.name)}
                  style={{ padding: "6px 12px", borderRadius: 8, border: "1.5px solid #fca5a5", background: "#fef2f2", color: "#dc2626", fontWeight: 700, cursor: "pointer", fontFamily: FF, fontSize: 13 }}
                >
                  삭제
                </button>
              </>
            )}
          </div>
        ))}
      </div>

      <div style={{ marginTop: 16, padding: "12px 14px", background: "#f8fafc", border: "1px solid #f1f5f9", borderRadius: 10, fontSize: 12, color: "#94a3b8", lineHeight: 1.6 }}>
        카테고리를 추가하면 계정 등록 시 바로 사용할 수 있습니다.<br />
        삭제해도 기존 계정의 카테고리 값은 유지됩니다.
      </div>
    </Modal>
  );
}

const getCS = (cat) => CAT_STYLE[cat] || CAT_STYLE["기타"];

const base = {
  input: {
    width: "100%", boxSizing: "border-box", border: "1.5px solid #e2e8f0",
    borderRadius: 12, padding: "11px 14px", fontSize: 14, outline: "none",
    background: "#fff", fontFamily: FF, color: "#0f172a",
  },
};

/* ─── CSV ─── */
function downloadCsv(filename, rows) {
  const csv = "\ufeff" + rows.map(r =>
    r.map(v => `"${String(v ?? "").replace(/"/g, '""')}"`).join(",")
  ).join("\n");
  const a = Object.assign(document.createElement("a"), {
    href: URL.createObjectURL(new Blob([csv], { type: "text/csv;charset=utf-8;" })),
    download: filename,
  });
  a.click();
}

/* ─── 엑셀 양식 ─── */
function downloadTemplate() {
  const ws = XLSX.utils.aoa_to_sheet([
    ["사이트명", "사이트URL", "아이디", "비밀번호", "카테고리", "담당자", "비고"],
    ["예시_홈택스", "https://www.hometax.go.kr", "sample_id", "sample_pw", "세금/회계", "홍길동", ""],
  ]);
  ws["!cols"] = Array(7).fill({ wch: 20 });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "업무계정양식");
  XLSX.writeFile(wb, "업무계정등록양식.xlsx");
}

function downloadAccountsExcel(accounts) {
  const rows = [
    ["사이트명", "사이트URL", "아이디", "비밀번호", "카테고리", "담당자", "비고"],
    ...accounts.map(a => [a.siteName || "", a.siteUrl || "", a.loginId || "", a.password || "", a.category || "", a.owner || "", a.note || ""]),
  ];
  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws["!cols"] = Array(7).fill({ wch: 20 });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "업무계정");
  XLSX.writeFile(wb, `업무계정_${new Date().toISOString().slice(0, 10)}.xlsx`);
}

/* ══════════════ 공통 컴포넌트 ══════════════ */

function Toast({ toast }) {
  if (!toast) return null;
  const isErr = toast.type === "error";
  return (
    <div style={{
      position: "fixed", bottom: 28, left: "50%", transform: "translateX(-50%)",
      zIndex: 9999, padding: "11px 20px", borderRadius: 14, fontWeight: 700,
      fontSize: 13, fontFamily: FF, maxWidth: "calc(100vw - 40px)", textAlign: "center",
      boxShadow: "0 8px 32px rgba(0,0,0,0.14)",
      background: isErr ? "#fef2f2" : "#f0fdf4",
      color: isErr ? "#dc2626" : "#15803d",
      border: `1px solid ${isErr ? "#fca5a5" : "#86efac"}`,
    }}>
      {toast.msg}
    </div>
  );
}

function Tag({ label, style }) {
  return (
    <span style={{
      display: "inline-flex", alignItems: "center", fontSize: 11,
      fontWeight: 800, padding: "3px 9px", borderRadius: 999, ...style,
    }}>
      {label}
    </span>
  );
}

function Divider({ label }) {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 10, margin: "24px 0 12px" }}>
      <div style={{ fontWeight: 900, fontSize: 15, color: "#0f172a" }}>{label}</div>
      <div style={{ flex: 1, height: 1, background: "#f1f5f9" }} />
    </div>
  );
}

/* ══════════════ 로그인 ══════════════ */

function LoginScreen({ onLogin, loading, error }) {
  const [id, setId] = useState("");
  const [pw, setPw] = useState("");
  const [resetEmail, setResetEmail] = useState("");
  const [resetMode, setResetMode] = useState(false);
  const [resetSending, setResetSending] = useState(false);
  const [resetMsg, setResetMsg] = useState("");

  async function handleReset() {
    if (!resetEmail.trim()) { setResetMsg("아이디를 입력해주세요."); return; }
    setResetSending(true);
    try {
      const email = resetEmail.trim().includes("@")
        ? resetEmail.trim()
        : `${resetEmail.trim()}@mauto.co.kr`;
      await sendPasswordResetEmail(auth, email);
      setResetMsg("비밀번호 재설정 링크가 메일로 발송되었습니다.");
    } catch {
      setResetMsg("메일 발송에 실패했습니다. 아이디를 다시 확인해주세요.");
    } finally {
      setResetSending(false);
    }
  }

  return (
    <div style={{
      minHeight: "100vh",
      background: "linear-gradient(135deg,#1e3a5f 0%,#1e293b 55%,#0f172a 100%)",
      display: "flex", alignItems: "center", justifyContent: "center",
      padding: 16, fontFamily: FF,
    }}>
      <div style={{
        width: "100%", maxWidth: 400, background: "#fff",
        borderRadius: 24, padding: "36px 28px",
        boxShadow: "0 32px 80px rgba(0,0,0,0.32)",
      }}>
        <div style={{ textAlign: "center", marginBottom: 28 }}>
          <div style={{ fontSize: 42, marginBottom: 8 }}>🔐</div>
          <div style={{ fontWeight: 900, fontSize: 22, color: "#0f172a" }}>업무 계정 관리</div>
          <div style={{ fontSize: 13, color: "#94a3b8", marginTop: 6 }}>회사 계정으로 로그인하세요</div>
        </div>

        {!resetMode ? (
          <>
            <div style={{ marginBottom: 12 }}>
              <label style={{ display: "block", fontSize: 12, fontWeight: 700, color: "#64748b", marginBottom: 6 }}>회사 아이디</label>
              <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                <input value={id} onChange={e => setId(e.target.value)}
                  onKeyDown={e => e.key === "Enter" && onLogin(id, pw)}
                  placeholder="예: yhj"
                  style={{ ...base.input, flex: 1 }} />
                <span style={{ fontSize: 13, color: "#94a3b8", whiteSpace: "nowrap" }}>@mauto.co.kr</span>
              </div>
            </div>

            <div style={{ marginBottom: 16 }}>
              <label style={{ display: "block", fontSize: 12, fontWeight: 700, color: "#64748b", marginBottom: 6 }}>비밀번호</label>
              <input type="password" value={pw} onChange={e => setPw(e.target.value)}
                onKeyDown={e => e.key === "Enter" && onLogin(id, pw)}
                style={base.input} />
            </div>

            {error && (
              <div style={{ background: "#fef2f2", border: "1px solid #fecaca", color: "#dc2626", borderRadius: 10, padding: "10px 12px", fontSize: 13, fontWeight: 700, marginBottom: 14 }}>
                {error}
              </div>
            )}

            <button onClick={() => onLogin(id, pw)} disabled={loading}
              style={{
                width: "100%", padding: "13px", borderRadius: 12, border: "none",
                background: loading ? "#94a3b8" : "#2563eb", color: "#fff",
                fontSize: 15, fontWeight: 800, cursor: loading ? "not-allowed" : "pointer", fontFamily: FF,
              }}>
              {loading ? "로그인 중..." : "로그인"}
            </button>

            <button onClick={() => { setResetMode(true); setResetMsg(""); }}
              style={{
                width: "100%", marginTop: 10, padding: "11px", borderRadius: 12,
                border: "1.5px solid #e2e8f0", background: "#f8fafc", color: "#64748b",
                fontSize: 13, fontWeight: 700, cursor: "pointer", fontFamily: FF,
              }}>
              🔑 비밀번호 초기화 요청
            </button>
          </>
        ) : (
          <>
            <div style={{ marginBottom: 12 }}>
              <label style={{ display: "block", fontSize: 12, fontWeight: 700, color: "#64748b", marginBottom: 6 }}>회사 아이디 입력</label>
              <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                <input value={resetEmail} onChange={e => setResetEmail(e.target.value)}
                  placeholder="예: yhj"
                  style={{ ...base.input, flex: 1 }} />
                <span style={{ fontSize: 13, color: "#94a3b8", whiteSpace: "nowrap" }}>@mauto.co.kr</span>
              </div>
            </div>

            <div style={{
              padding: 12, background: "#eff6ff", border: "1px solid #bfdbfe",
              borderRadius: 10, fontSize: 13, color: "#1e3a8a", lineHeight: 1.65, marginBottom: 14,
            }}>
              아이디를 입력하면 등록된 이메일로 비밀번호 변경 링크가 발송됩니다.
            </div>

            {resetMsg && (
              <div style={{
                padding: "10px 12px", borderRadius: 10, fontSize: 13, fontWeight: 700, marginBottom: 12,
                background: resetMsg.includes("발송") ? "#f0fdf4" : "#fef2f2",
                color: resetMsg.includes("발송") ? "#15803d" : "#dc2626",
                border: `1px solid ${resetMsg.includes("발송") ? "#86efac" : "#fca5a5"}`,
              }}>
                {resetMsg}
              </div>
            )}

            <button onClick={handleReset} disabled={resetSending}
              style={{
                width: "100%", padding: "13px", borderRadius: 12, border: "none",
                background: resetSending ? "#94a3b8" : "#2563eb", color: "#fff",
                fontSize: 14, fontWeight: 800, cursor: resetSending ? "not-allowed" : "pointer", fontFamily: FF,
              }}>
              {resetSending ? "발송 중..." : "비밀번호 재설정 메일 받기"}
            </button>

            <button onClick={() => { setResetMode(false); setResetMsg(""); }}
              style={{
                width: "100%", marginTop: 10, padding: "11px", borderRadius: 12,
                border: "1.5px solid #e2e8f0", background: "#fff", color: "#64748b",
                fontSize: 13, fontWeight: 700, cursor: "pointer", fontFamily: FF,
              }}>
              ← 로그인으로 돌아가기
            </button>
          </>
        )}
      </div>
    </div>
  );
}

/* ══════════════ 헤더 ══════════════ */

function HeaderOld({ user, isAdmin, onLogout, detailMode, onBack,
  onAdd, onImportExcel, onDownloadAccounts, onDownloadTemplate, onShowLogs, onShowUsers,
  onShowCategories, onDeleteAll }) {
  return (
    <div style={{
      position: "sticky", top: 0, zIndex: 30,
      background: "rgba(255,255,255,0.94)", backdropFilter: "blur(12px)",
      borderBottom: "1px solid #f1f5f9",
      boxShadow: "0 1px 4px rgba(15,23,42,0.06)",
    }}>
      <div style={{
        maxWidth: 960, margin: "0 auto", padding: "12px 16px",
        display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap",
      }}>
        {detailMode
          ? <button onClick={onBack} style={{ background: "none", border: "none", cursor: "pointer", color: "#64748b", fontWeight: 800, fontSize: 14, fontFamily: FF, display: "flex", alignItems: "center", gap: 4 }}>← 목록</button>
          : <span style={{ fontWeight: 900, fontSize: 17, color: "#0f172a" }}>🔐 업무 계정 관리</span>
        }

        <Tag label={isAdmin ? "관리자" : "열람 전용"} style={{
          color: isAdmin ? "#7c3aed" : "#0369a1",
          background: isAdmin ? "#f5f3ff" : "#f0f9ff",
          border: `1px solid ${isAdmin ? "#c4b5fd" : "#bae6fd"}`,
        }} />

        <div style={{ flex: 1 }} />
        <span style={{ fontSize: 12, color: "#94a3b8" }}>{user?.email}</span>

        {isAdmin && !detailMode && <>
          <HBtn onClick={onDownloadTemplate}>📋 양식</HBtn>
          <HBtn onClick={onDownloadAccounts}>⬇ 다운로드</HBtn>
          <HBtn onClick={onImportExcel} variant="green">📥 업로드</HBtn>
          <HBtn onClick={onDeleteAll} variant="red">🗑 전체삭제</HBtn>
          <HBtn onClick={onShowCategories}>카테고리</HBtn>
          <HBtn onClick={onShowLogs}>로그</HBtn>
          <HBtn onClick={onShowUsers}>사용자</HBtn>
          <HBtn onClick={onAdd} variant="blue">+ 추가</HBtn>
        </>}

        <HBtn onClick={onLogout}>로그아웃</HBtn>
      </div>
    </div>
  );
}

function HBtn({ children, onClick, variant }) {
  const v = {
    blue: { background: "#2563eb", color: "#fff", border: "none" },
    green: { background: "#f0fdf4", color: "#16a34a", border: "1.5px solid #86efac" },
    red: { background: "#fef2f2", color: "#dc2626", border: "1.5px solid #fca5a5" },
    default: { background: "#fff", color: "#475569", border: "1.5px solid #e2e8f0" },
  };
  const s = v[variant] || v.default;
  return (
    <button onClick={onClick} style={{
      ...s, borderRadius: 10, padding: "8px 12px",
      fontWeight: 700, fontSize: 13, cursor: "pointer", fontFamily: FF,
    }}>
      {children}
    </button>
  );
}

/* ══════════════ 계정 행 (리스트) ══════════════ */

function HeaderUnused({ user, isAdmin, onLogout, detailMode, onBack,
  onAdd, onImportExcel, onDownloadAccounts, onDownloadTemplate, onShowLogs, onShowUsers,
  onShowCategories, onDeleteAll }) {
  return (
    <div style={{
      position: "sticky", top: 0, zIndex: 30,
      background: "rgba(255,255,255,0.94)", backdropFilter: "blur(12px)",
      borderBottom: "1px solid #f1f5f9",
      boxShadow: "0 1px 4px rgba(15,23,42,0.06)",
    }}>
      <div style={{
        maxWidth: 960, margin: "0 auto", padding: "10px 16px",
        display: "grid", gap: 10,
      }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
          {detailMode && <HBtn onClick={onBack}>목록</HBtn>}
          <Tag label={isAdmin ? "관리자" : "조회 전용"} style={{
            color: isAdmin ? "#7c3aed" : "#0369a1",
            background: isAdmin ? "#f5f3ff" : "#f0f9ff",
            border: `1px solid ${isAdmin ? "#c4b5fd" : "#bae6fd"}`,
          }} />
          <div style={{ flex: 1 }} />
          <span style={{ fontSize: 12, color: "#94a3b8" }}>{user?.email}</span>
          <HBtn onClick={onLogout}>로그아웃</HBtn>
        </div>

        {isAdmin && !detailMode && (
          <div style={{ display: "grid", gap: 8 }}>
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
              <HBtn onClick={onAdd} variant="blue">추가</HBtn>
              <HBtn onClick={onDownloadTemplate}>양식</HBtn>
              <HBtn onClick={onImportExcel} variant="green">업로드</HBtn>
            </div>
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
              <HBtn onClick={onDeleteAll} variant="red">전체삭제</HBtn>
              <HBtn onClick={onDownloadAccounts}>다운로드</HBtn>
              <HBtn onClick={onShowCategories}>카테고리</HBtn>
              <HBtn onClick={onShowLogs}>로그</HBtn>
              <HBtn onClick={onShowUsers}>사용자</HBtn>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

function HeaderLegacy({ user, isAdmin, onLogout, detailMode, onBack,
  onAdd, onImportExcel, onDownloadAccounts, onDownloadTemplate, onShowLogs, onShowUsers,
  onShowCategories, onDeleteAll }) {
  return (
    <div style={{
      position: "sticky", top: 0, zIndex: 30,
      background: "rgba(255,255,255,0.94)", backdropFilter: "blur(12px)",
      borderBottom: "1px solid #f1f5f9",
      boxShadow: "0 1px 4px rgba(15,23,42,0.06)",
    }}>
      <div style={{
        maxWidth: 960, margin: "0 auto", padding: "10px 16px",
        display: "flex", flexDirection: "column", gap: 10,
      }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap", width: "100%" }}>
          {detailMode && (
            <button onClick={onBack} style={{
              background: "#fff", border: "1.5px solid #e2e8f0", borderRadius: 10,
              cursor: "pointer", color: "#475569", fontWeight: 800, fontSize: 13,
              fontFamily: FF, display: "flex", alignItems: "center", gap: 4, padding: "8px 12px",
            }}>목록으로</button>
          )}
          <Tag label={isAdmin ? "관리자" : "조회 전용"} style={{
            color: isAdmin ? "#7c3aed" : "#0369a1",
            background: isAdmin ? "#f5f3ff" : "#f0f9ff",
            border: `1px solid ${isAdmin ? "#c4b5fd" : "#bae6fd"}`,
          }} />
          <span style={{ fontSize: 12, color: "#94a3b8", flex: 1, minWidth: 180 }}>{user?.email}</span>
          <HBtn onClick={onLogout}>로그아웃</HBtn>
        </div>

        {isAdmin && !detailMode && (
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap", width: "100%" }}>
            <HBtn onClick={onAdd} variant="blue">+ 계정 추가</HBtn>
            <HBtn onClick={onImportExcel} variant="green">엑셀 업로드</HBtn>
            <HBtn onClick={onDownloadAccounts}>전체 다운로드</HBtn>
            <HBtn onClick={onDownloadTemplate}>양식</HBtn>
            <HBtn onClick={onShowUsers}>사용자</HBtn>
            <HBtn onClick={onShowCategories}>카테고리</HBtn>
            <HBtn onClick={onShowLogs}>로그</HBtn>
            <HBtn onClick={onDeleteAll} variant="red">전체 삭제</HBtn>
          </div>
        )}
      </div>
    </div>
  );
}

function ToolbarHeader({ user, isAdmin, onLogout, detailMode, onBack,
  onAdd, onImportExcel, onDownloadAccounts, onDownloadTemplate, onShowLogs, onShowUsers,
  onShowCategories, onDeleteAll }) {
  return (
    <div style={{
      position: "sticky", top: 0, zIndex: 30,
      background: "rgba(255,255,255,0.94)", backdropFilter: "blur(12px)",
      borderBottom: "1px solid #f1f5f9",
      boxShadow: "0 1px 4px rgba(15,23,42,0.06)",
    }}>
      <div style={{
        maxWidth: 960, margin: "0 auto", padding: "10px 16px",
        display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap",
      }}>
        {detailMode && <HBtn onClick={onBack}>목록</HBtn>}

        {isAdmin && !detailMode && (
          <>
            <HBtn onClick={onAdd} variant="blue">+ 계정 추가</HBtn>
            <HBtn onClick={onDownloadTemplate}>업로드 양식</HBtn>
            <HBtn onClick={onImportExcel} variant="green">엑셀 업로드</HBtn>
            <span style={{ width: 1, height: 26, background: "#e2e8f0", margin: "0 2px" }} />
            <HBtn onClick={onDeleteAll} variant="red">전체 삭제</HBtn>
            <span style={{ width: 1, height: 26, background: "#e2e8f0", margin: "0 2px" }} />
            <HBtn onClick={onDownloadAccounts}>계정 다운로드</HBtn>
            <HBtn onClick={onShowCategories}>카테고리</HBtn>
            <HBtn onClick={onShowLogs}>로그</HBtn>
            <HBtn onClick={onShowUsers}>사용자</HBtn>
          </>
        )}

        <div style={{ flex: 1 }} />
        <span style={{ fontSize: 12, color: "#94a3b8" }}>{user?.email}</span>
        <Tag label={isAdmin ? "관리자" : "조회 전용"} style={{
          color: isAdmin ? "#7c3aed" : "#0369a1",
          background: isAdmin ? "#f5f3ff" : "#f0f9ff",
          border: `1px solid ${isAdmin ? "#c4b5fd" : "#bae6fd"}`,
        }} />
        <HBtn onClick={onLogout}>로그아웃</HBtn>
      </div>
    </div>
  );
}

function CompactToolbarHeader({ user, isAdmin, onLogout, detailMode, onBack,
  onAdd, onImportExcel, onDownloadAccounts, onDownloadTemplate, onShowLogs, onShowUsers,
  onShowCategories, onDeleteAll }) {
  return (
    <div style={{
      position: "sticky", top: 0, zIndex: 30,
      background: "rgba(255,255,255,0.94)", backdropFilter: "blur(12px)",
      borderBottom: "1px solid #f1f5f9",
      boxShadow: "0 1px 4px rgba(15,23,42,0.06)",
    }}>
      <div style={{
        maxWidth: 960, margin: "0 auto", padding: "10px 16px",
        display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap",
      }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap", flex: 1, minWidth: 0 }}>
          {detailMode && <HBtn onClick={onBack}>목록</HBtn>}

          {isAdmin && !detailMode && <>
            <HBtn onClick={onAdd} variant="blue">+ 계정 추가</HBtn>
            <HBtn onClick={onDownloadTemplate}>업로드 양식</HBtn>
            <HBtn onClick={onImportExcel} variant="green">엑셀 업로드</HBtn>
            <span style={{ width: 1, height: 26, background: "#e2e8f0", margin: "0 2px" }} />
            <HBtn onClick={onDeleteAll} variant="red">전체 삭제</HBtn>
            <span style={{ width: 1, height: 26, background: "#e2e8f0", margin: "0 2px" }} />
            <HBtn onClick={onDownloadAccounts}>계정 다운로드</HBtn>
            <HBtn onClick={onShowCategories}>카테고리</HBtn>
            <HBtn onClick={onShowLogs}>로그</HBtn>
            <HBtn onClick={onShowUsers}>사용자</HBtn>
          </>}
        </div>

        <div style={{ display: "flex", alignItems: "center", gap: 8, whiteSpace: "nowrap", flexShrink: 0 }}>
          <span style={{ fontSize: 12, color: "#94a3b8" }}>{user?.email}</span>
          <Tag label={isAdmin ? "관리자" : "조회 전용"} style={{
            color: isAdmin ? "#7c3aed" : "#0369a1",
            background: isAdmin ? "#f5f3ff" : "#f0f9ff",
            border: `1px solid ${isAdmin ? "#c4b5fd" : "#bae6fd"}`,
          }} />
          <HBtn onClick={onLogout}>로그아웃</HBtn>
        </div>
      </div>
    </div>
  );
}

/* ══════════════ 새 헤더 (모바일/PC 통합) ══════════════ */

function AppHeader({ user, userProfile, isAdmin, onLogout, detailMode, onBack,
  mainTab, onTabChange, hasCashAccess,
  onAdd, onImportExcel, onDownloadAccounts, onDownloadTemplate,
  onShowLogs, onShowUsers, onShowCategories, onDeleteAll }) {
  const [showAdmin, setShowAdmin] = useState(false);
  const isDesktop = useMediaQuery("(min-width: 900px)");
  const dept = String(userProfile?.department || "").trim();
  const pos = String(userProfile?.position || "").trim();
  const isMgmtDept = dept.includes("관리");
  const isDeptHead = !!userProfile?.isDeptHead;
  const isCEO = pos.includes("대표");
  const showCashflowLink = isMgmtDept || isCEO;

  const ICON_LINKS = [
    ...(isMgmtDept ? [{
      key: "mgmt",
      title: "관리부 업무 바로가기",
      icon: "🗂️",
      href: "https://script.google.com/macros/s/AKfycbwqAhvVZ-QtoyKeE4ESUKIuD9Gr5cYBKlf5hIyb6suFuEQSTKQI5caL_MogqeW7JoU-/exec",
    }] : []),
    ...(isDeptHead ? [{
      key: "head",
      title: "부서장 업무 바로가기",
      icon: "👔",
      href: "https://script.google.com/macros/s/AKfycbwYHAmXGxwmnPTRWftG3ujya8sT3657M6uZb8b4ATh05MMQX3Ich3N99_z2I8oKqvKx/exec",
    }] : []),
    ...(showCashflowLink ? [{
      key: "cashflow",
      title: "시재/현금흐름 바로가기",
      icon: "💸",
      href: "https://yhhojt970-cell.github.io/cashflow-app/",
    }] : []),
    {
      key: "stock",
      title: "전직원 공통 (stock-mirae)",
      icon: "📈",
      href: "https://stock-mirae.web.app/",
    },
  ];

  return (
    <div style={{
      position: "sticky", top: 0, zIndex: 30,
      background: "#fff",
      boxShadow: "0 1px 3px rgba(15,23,42,0.08)",
    }}>
      {/* ── 상단 바 ── */}
      <div style={{
        maxWidth: 960, margin: "0 auto",
        padding: isDesktop ? "0 16px" : "8px 12px",
        minHeight: 52,
        display: "flex",
        alignItems: "center",
        gap: 10,
        flexWrap: isDesktop ? "nowrap" : "wrap",
      }}>
        {detailMode ? (
          <button onClick={onBack} style={{
            background: "none", border: "none", cursor: "pointer",
            color: "#2563eb", fontWeight: 800, fontSize: 14,
            fontFamily: FF, display: "flex", alignItems: "center", gap: 4, padding: 0,
          }}>← 목록</button>
        ) : (
          <span style={{ fontWeight: 900, fontSize: 16, color: "#0f172a" }}>미래오토메이션(주)</span>
        )}

        <div style={{ flex: isDesktop ? 1 : 0 }} />

        {/* 역할/부서별 바로가기 */}
        <div style={{
          display: "flex",
          alignItems: "center",
          gap: 6,
          flexShrink: 0,
          maxWidth: isDesktop ? "none" : "100%",
          overflowX: isDesktop ? "visible" : "auto",
          paddingBottom: isDesktop ? 0 : 2,
        }}>
          {ICON_LINKS.map((l) => (
            <a
              key={l.key}
              href={l.href}
              target="_blank"
              rel="noopener noreferrer"
              title={l.title}
              style={{
                width: isDesktop ? 34 : 32,
                height: isDesktop ? 34 : 32,
                display: "inline-flex",
                alignItems: "center",
                justifyContent: "center",
                borderRadius: 11,
                border: "1.5px solid #e2e8f0",
                background: "#fff",
                color: "#0f172a",
                textDecoration: "none",
                cursor: "pointer",
                fontSize: isDesktop ? 16 : 15,
                lineHeight: 1,
              }}
            >
              {l.icon}
            </a>
          ))}
        </div>

        {/* 사용자 이메일 (짧게) */}
        <span style={{
          fontSize: 12, color: "#94a3b8",
          maxWidth: 130, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap",
          display: isDesktop ? "inline" : "none",
        }} className="email-label">{user?.email}</span>

        <div style={{
          display: "flex",
          alignItems: "center",
          gap: 6,
          marginLeft: isDesktop ? 0 : "auto",
          flexShrink: 0,
        }}>
          {isDesktop && (
            <Tag label={isAdmin ? "관리자" : "열람"} style={{
              color: isAdmin ? "#7c3aed" : "#0369a1",
              background: isAdmin ? "#f5f3ff" : "#f0f9ff",
              border: `1px solid ${isAdmin ? "#c4b5fd" : "#bae6fd"}`,
              fontSize: 11,
            }} />
          )}

          {/* 관리 도구 토글 (관리자 + 업무계정 탭만) */}
          {isAdmin && !detailMode && mainTab === "accounts" && (
            <button onClick={() => setShowAdmin(s => !s)} style={{
              padding: isDesktop ? "6px 11px" : "5px 9px",
              borderRadius: 9,
              border: showAdmin ? "1.5px solid #c4b5fd" : "1.5px solid #e2e8f0",
              background: showAdmin ? "#f5f3ff" : "#fff",
              color: showAdmin ? "#7c3aed" : "#475569",
              fontWeight: 700,
              fontSize: isDesktop ? 12 : 11,
              cursor: "pointer",
              fontFamily: FF,
              display: "flex",
              alignItems: "center",
              gap: 4,
            }}>
              ⚙ 관리{showAdmin ? " ▲" : " ▼"}
            </button>
          )}

          <button onClick={onLogout} style={{
            padding: isDesktop ? "6px 12px" : "5px 9px",
            borderRadius: 9,
            border: "1.5px solid #e2e8f0", background: "#fff",
            color: "#475569", fontWeight: 700, fontSize: isDesktop ? 12 : 11,
            cursor: "pointer", fontFamily: FF,
          }}>로그아웃</button>
        </div>
      </div>

      {/* ── 관리 도구 (접힘/펼침) ── */}
      {isAdmin && showAdmin && !detailMode && mainTab === "accounts" && (
        <div style={{
          borderTop: "1px solid #f1f5f9",
          padding: "10px 16px",
          display: "flex", flexWrap: "wrap", gap: 6,
          maxWidth: 960, margin: "0 auto", boxSizing: "border-box",
          background: "#fafbff",
        }}>
          <HBtn onClick={() => { onAdd(); setShowAdmin(false); }} variant="blue">+ 계정 추가</HBtn>
          <HBtn onClick={onDownloadTemplate}>업로드 양식</HBtn>
          <HBtn onClick={() => { onImportExcel(); setShowAdmin(false); }} variant="green">엑셀 업로드</HBtn>
          <HBtn onClick={onDownloadAccounts}>다운로드</HBtn>
          <HBtn onClick={() => { onShowCategories(); setShowAdmin(false); }}>카테고리</HBtn>
          <HBtn onClick={() => { onShowLogs(); setShowAdmin(false); }}>로그</HBtn>
          <HBtn onClick={() => { onShowUsers(); setShowAdmin(false); }}>사용자</HBtn>
          <HBtn onClick={() => { onDeleteAll(); setShowAdmin(false); }} variant="red">전체 삭제</HBtn>
        </div>
      )}

      {/* ── 탭 바 ── */}
      {!detailMode && (
        <div style={{
          borderTop: "1px solid #f1f5f9",
          display: "flex",
          overflowX: isDesktop ? "visible" : "auto",
          maxWidth: 960, margin: "0 auto",
        }}>
          {[
            ["accounts", "🔐 업무계정"],
            ["contacts", "👥 직원"],
            ...(hasCashAccess ? [["cash", "💰 시재관리"]] : []),
            ["notice", "📢 공지사항"],
            ["company", "🏢 회사 정보"],
          ].map(([id, label]) => (
            <button key={id} onClick={() => onTabChange(id)} style={{
              flex: isDesktop ? 1 : "0 0 auto",
              minWidth: isDesktop ? 0 : 122,
              padding: "10px 8px",
              fontFamily: FF, fontWeight: 700, fontSize: 13,
              border: "none", background: "none", cursor: "pointer",
              borderBottom: mainTab === id ? "2.5px solid #2563eb" : "2.5px solid transparent",
              color: mainTab === id ? "#2563eb" : "#64748b",
              transition: "color 0.15s",
              whiteSpace: "nowrap",
            }}>{label}</button>
          ))}
        </div>
      )}
    </div>
  );
}

/* ══════════════ 직원 연락처 ══════════════ */

function ContactModal({ contact, onSave, onClose }) {
  const [form, setForm] = useState({
    name: contact?.name || "",
    dept: contact?.dept || "",
    position: contact?.position || "",
    phone: contact?.phone || "",
    ext: contact?.ext || "",
    direct: contact?.direct || "",
  });
  const [saving, setSaving] = useState(false);

  function up(k, v) { setForm(f => ({ ...f, [k]: v })); }

  async function handleSave() {
    if (!form.name.trim()) { alert("이름을 입력하세요."); return; }
    setSaving(true);
    try { await onSave(form); onClose(); }
    finally { setSaving(false); }
  }

  const row = (label, key, placeholder, type = "text") => (
    <div style={{ marginBottom: 12 }}>
      <label style={{ display: "block", fontSize: 12, fontWeight: 700, color: "#64748b", marginBottom: 5 }}>{label}</label>
      <input
        type={type} value={form[key]}
        onChange={e => up(key, e.target.value)}
        placeholder={placeholder}
        style={{ ...base.input }}
      />
    </div>
  );

  return (
    <Modal title={contact ? "연락처 수정" : "연락처 추가"} onClose={onClose}>
      {row("이름 *", "name", "홍길동")}
      {row("부서", "dept", "관리부")}
      {row("직위", "position", "차장")}
      {row("전화번호", "phone", "010-1234-5678", "tel")}
      {row("내선번호", "ext", "1234")}
      {row("직통번호", "direct", "031-123-4567", "tel")}
      <button
        onClick={handleSave} disabled={saving}
        style={{
          width: "100%", padding: "12px", borderRadius: 12, border: "none",
          background: saving ? "#94a3b8" : "#2563eb", color: "#fff",
          fontWeight: 800, fontSize: 14, cursor: saving ? "not-allowed" : "pointer", fontFamily: FF,
        }}
      >{saving ? "저장 중..." : "저장"}</button>
    </Modal>
  );
}

function formatPhone(v) {
  if (!v) return v;
  const d = v.replace(/\D/g, "");
  if (d.startsWith("02")) {
    if (d.length === 9) return d.replace(/^(02)(\d{3})(\d{4})$/, "$1-$2-$3");
    if (d.length === 10) return d.replace(/^(02)(\d{4})(\d{4})$/, "$1-$2-$3");
  }
  if (d.length === 10) return d.replace(/^(\d{3})(\d{3})(\d{4})$/, "$1-$2-$3");
  if (d.length === 11) return d.replace(/^(\d{3})(\d{4})(\d{4})$/, "$1-$2-$3");
  return v;
}

function ContactCard({ contact, isAdmin, onEdit, onDelete }) {
  const [showPhone, setShowPhone] = useState(false);
  const [countdown, setCountdown] = useState(0);
  const timerRef = useRef(null);

  useEffect(() => () => clearInterval(timerRef.current), []);

  function revealPhone() {
    if (showPhone) { clearInterval(timerRef.current); setShowPhone(false); setCountdown(0); return; }
    setShowPhone(true); setCountdown(10);
    clearInterval(timerRef.current);
    timerRef.current = setInterval(() => {
      setCountdown(prev => {
        if (prev <= 1) { clearInterval(timerRef.current); setShowPhone(false); return 0; }
        return prev - 1;
      });
    }, 1000);
  }

  const initials = (contact.name || "?").slice(0, 1);
  const deptColors = ["#dbeafe", "#dcfce7", "#fce7f3", "#fef3c7", "#ede9fe", "#ffedd5", "#f1f5f9"];
  const colorIdx = (contact.dept || "").split("").reduce((s, c) => s + c.charCodeAt(0), 0) % deptColors.length;

  return (
    <div style={{
      background: "#fff", borderRadius: 16, padding: "18px 16px 14px",
      border: "1px solid #e8edf3",
      boxShadow: "0 2px 8px rgba(15,23,42,0.07)",
      marginBottom: 10, position: "relative",
    }}>
      {/* 관리자 버튼 - 우상단 고정 */}
      {isAdmin && (
        <div style={{ position: "absolute", top: 10, right: 10, display: "flex", gap: 5 }}>
          <button onClick={() => onEdit(contact)} style={{
            padding: "3px 9px", borderRadius: 7, border: "1.5px solid #bfdbfe",
            background: "#eff6ff", color: "#2563eb", fontWeight: 700, fontSize: 11,
            cursor: "pointer", fontFamily: FF,
          }}>수정</button>
          <button onClick={() => onDelete(contact)} style={{
            padding: "3px 9px", borderRadius: 7, border: "1.5px solid #fca5a5",
            background: "#fef2f2", color: "#dc2626", fontWeight: 700, fontSize: 11,
            cursor: "pointer", fontFamily: FF,
          }}>삭제</button>
        </div>
      )}

      {/* 상단: 아바타 + 이름 + 직위 중앙 정렬 */}
      <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 5, marginBottom: 14, paddingTop: 4 }}>
        <div style={{
          width: 50, height: 50, borderRadius: "50%",
          background: deptColors[colorIdx],
          display: "flex", alignItems: "center", justifyContent: "center",
          fontWeight: 900, fontSize: 20, color: "#374151",
        }}>{initials}</div>
        <div style={{ fontWeight: 800, fontSize: 16, color: "#0f172a", textAlign: "center" }}>{contact.name}</div>
        {contact.position && (
          <div style={{ fontSize: 12, color: "#64748b", textAlign: "center" }}>{contact.position}</div>
        )}
      </div>

      {/* 구분선 */}
      <div style={{ borderTop: "1px solid #f1f5f9", marginBottom: 12 }} />

      {/* 연락처 정보 */}
      <div style={{ display: "grid", gap: 8 }}>
        {/* 전화번호 (보안) */}
        {(contact.phone || true) && (
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <span style={{ fontSize: 11, fontWeight: 700, color: "#94a3b8", width: 46, flexShrink: 0, textAlign: "right" }}>전화</span>
            <span style={{
              flex: 1, fontSize: 13, fontWeight: 600,
              color: showPhone ? "#0f172a" : "#94a3b8",
              letterSpacing: showPhone ? 0 : 1,
            }}>
              {showPhone ? formatPhone(contact.phone) || "-" : (contact.phone ? "•••-••••-••••" : "-")}
            </span>
            {contact.phone && (
              <button onClick={revealPhone} style={{
                padding: "3px 10px", borderRadius: 6,
                border: showPhone ? "1.5px solid #475569" : "1.5px solid #2563eb",
                background: showPhone ? "#f8fafc" : "#eff6ff",
                color: showPhone ? "#475569" : "#2563eb",
                fontWeight: 700, fontSize: 11, cursor: "pointer", fontFamily: FF,
                whiteSpace: "nowrap",
              }}>
                {showPhone ? `숨기기 ${countdown}s` : "보기"}
              </button>
            )}
          </div>
        )}
        {contact.ext && (
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <span style={{ fontSize: 11, fontWeight: 700, color: "#94a3b8", width: 46, flexShrink: 0, textAlign: "right" }}>내선</span>
            <span style={{ fontSize: 13, color: "#374151" }}>{contact.ext}</span>
          </div>
        )}
        {contact.direct && (
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <span style={{ fontSize: 11, fontWeight: 700, color: "#94a3b8", width: 46, flexShrink: 0, textAlign: "right" }}>직통</span>
            <span style={{ fontSize: 13, color: "#374151" }}>{formatPhone(contact.direct)}</span>
          </div>
        )}
      </div>
    </div>
  );
}

function ContactsView({ contacts, isAdmin, showToast }) {
  const [search, setSearch] = useState("");
  const [expandedDepts, setExpandedDepts] = useState({});
  const [modalContact, setModalContact] = useState(undefined); // undefined=닫힘, null=신규, {...}=수정

  const filtered = useMemo(() => {
    const q = search.trim().toLowerCase();
    if (!q) return contacts;
    return contacts.filter(c =>
      [c.name, c.dept, c.position, c.ext, c.direct].join(" ").toLowerCase().includes(q)
    );
  }, [contacts, search]);

  const grouped = useMemo(() => {
    const g = {};
    filtered.forEach(c => {
      const dept = c.dept || "기타";
      if (!g[dept]) g[dept] = [];
      g[dept].push(c);
    });
    return g;
  }, [filtered]);

  const depts = Object.keys(grouped).sort((a, b) => a.localeCompare(b, "ko"));

  async function handleSave(form) {
    try {
      if (modalContact?.id) {
        await updateContact(modalContact.id, form);
        showToast("수정되었습니다.");
      } else {
        await addContact(form);
        showToast("추가되었습니다.");
      }
    } catch (err) {
      console.error("연락처 저장 오류:", err);
      showToast(`저장 오류: ${err?.message || err}`, "error");
      throw err;
    }
  }

  async function handleDelete(contact) {
    if (!window.confirm(`"${contact.name}"을(를) 삭제할까요?`)) return;
    try { await deleteContact(contact.id); showToast("삭제되었습니다."); }
    catch { showToast("삭제 중 오류가 발생했습니다.", "error"); }
  }

  return (
    <div style={{ maxWidth: 960, margin: "0 auto", padding: "20px 16px 80px" }}>
      {/* 연락처 모달 */}
      {modalContact !== undefined && (
        <ContactModal
          contact={modalContact}
          onSave={handleSave}
          onClose={() => setModalContact(undefined)}
        />
      )}

      {/* 검색 + 추가 버튼 */}
      <div style={{ display: "flex", gap: 10, marginBottom: 20, alignItems: "center" }}>
        <div style={{
          flex: 1, background: "#fff", border: "1.5px solid #f1f5f9", borderRadius: 14,
          padding: "4px 14px", boxShadow: "0 2px 8px rgba(15,23,42,0.05)",
          display: "flex", alignItems: "center", gap: 8,
        }}>
          <span style={{ color: "#cbd5e1", fontSize: 15 }}>🔍</span>
          <input value={search} onChange={e => setSearch(e.target.value)}
            placeholder="이름, 부서, 직위로 검색"
            style={{ ...base.input, border: "none", padding: "8px 4px", boxShadow: "none", flex: 1 }} />
          {search && (
            <button onClick={() => setSearch("")} style={{
              background: "none", border: "none", cursor: "pointer", color: "#cbd5e1", fontSize: 14,
            }}>✕</button>
          )}
        </div>
        {isAdmin && (
          <button onClick={() => setModalContact(null)} style={{
            padding: "10px 16px", borderRadius: 12, border: "none",
            background: "#2563eb", color: "#fff", fontWeight: 800, fontSize: 13,
            cursor: "pointer", fontFamily: FF, whiteSpace: "nowrap",
          }}>+ 추가</button>
        )}
      </div>

      {/* 총 인원 */}
      <div style={{ fontSize: 13, color: "#cbd5e1", fontWeight: 600, marginBottom: 16 }}>
        총 {contacts.length}명{search ? ` · 검색 결과 ${filtered.length}명` : ""}
      </div>

      {/* 부서별 섹션 */}
      {depts.length === 0 && (
        <div style={{ textAlign: "center", padding: "60px 0", color: "#cbd5e1" }}>
          <div style={{ fontSize: 36, marginBottom: 12 }}>👥</div>
          <div style={{ fontSize: 15 }}>
            {search ? "검색 결과가 없습니다." : "등록된 연락처가 없습니다."}
          </div>
          {isAdmin && !search && (
            <button onClick={() => setModalContact(null)} style={{
              marginTop: 16, padding: "10px 20px", borderRadius: 12, border: "none",
              background: "#2563eb", color: "#fff", fontWeight: 800, fontSize: 13,
              cursor: "pointer", fontFamily: FF,
            }}>+ 첫 연락처 추가</button>
          )}
        </div>
      )}

      {depts.map(dept => {
        const isExpanded = expandedDepts[dept] !== false;
        const items = grouped[dept];
        return (
          <div key={dept} style={{ marginBottom: 6 }}>
            {/* 부서 헤더 */}
            <div
              onClick={() => setExpandedDepts(p => ({ ...p, [dept]: !isExpanded }))}
              style={{
                display: "flex", alignItems: "center", gap: 10,
                padding: "10px 14px", borderRadius: 10, background: "#f8fafc",
                cursor: "pointer", marginBottom: isExpanded ? 8 : 0,
                userSelect: "none",
              }}
            >
              <span style={{ fontWeight: 800, fontSize: 14, color: "#0f172a" }}>{dept}</span>
              <span style={{ fontSize: 12, color: "#94a3b8", fontWeight: 600 }}>{items.length}명</span>
              <div style={{ flex: 1 }} />
              <span style={{
                fontSize: 12, color: "#94a3b8",
                transform: isExpanded ? "rotate(0deg)" : "rotate(-90deg)",
                transition: "transform 0.2s", display: "inline-block",
              }}>▼</span>
            </div>
            {/* 카드 목록 */}
            {isExpanded && items.map(c => (
              <ContactCard
                key={c.id} contact={c} isAdmin={isAdmin}
                onEdit={ct => setModalContact(ct)}
                onDelete={handleDelete}
              />
            ))}
          </div>
        );
      })}
    </div>
  );
}

function compactPhone(v) {
  return v ? v.replace(/\D/g, "") : v;
}

function useMediaQuery(query) {
  const getMatches = () => {
    if (typeof window === "undefined") return false;
    return window.matchMedia(query).matches;
  };

  const [matches, setMatches] = useState(getMatches);

  useEffect(() => {
    if (typeof window === "undefined") return undefined;
    const mediaQuery = window.matchMedia(query);
    const onChange = (event) => setMatches(event.matches);

    setMatches(mediaQuery.matches);
    mediaQuery.addEventListener("change", onChange);
    return () => mediaQuery.removeEventListener("change", onChange);
  }, [query]);

  return matches;
}

function DesktopContactCard({ contact, isAdmin, isDesktop, onEdit, onDelete }) {
  const [showPhone, setShowPhone] = useState(false);
  const [countdown, setCountdown] = useState(0);
  const timerRef = useRef(null);

  useEffect(() => () => clearInterval(timerRef.current), []);

  function revealPhone() {
    if (showPhone) {
      clearInterval(timerRef.current);
      setShowPhone(false);
      setCountdown(0);
      return;
    }

    setShowPhone(true);
    setCountdown(10);
    clearInterval(timerRef.current);
    timerRef.current = setInterval(() => {
      setCountdown((prev) => {
        if (prev <= 1) {
          clearInterval(timerRef.current);
          setShowPhone(false);
          return 0;
        }
        return prev - 1;
      });
    }, 1000);
  }

  const infoItem = (label, value, accent = false, href) => {
    if (!value) return null;

    return (
      <div style={{
        padding: "12px 14px",
        borderRadius: 12,
        background: accent ? "#eff6ff" : "#f8fafc",
        border: `1px solid ${accent ? "#bfdbfe" : "#e2e8f0"}`,
      }}>
        <div style={{ fontSize: 11, fontWeight: 800, color: accent ? "#2563eb" : "#94a3b8", marginBottom: 5 }}>
          {label}
        </div>
        <div style={{ fontSize: 14, lineHeight: 1.35, fontWeight: accent ? 800 : 700, color: "#0f172a" }}>
          {href ? <a href={href} style={{ color: "#0f172a", textDecoration: "none" }}>{value}</a> : value}
        </div>
      </div>
    );
  };

  const visiblePhone = showPhone
    ? formatPhone(contact.phone) || "-"
    : (contact.phone ? "•••-••••-••••" : "-");

  return (
    <div style={{
      background: "#fff",
      borderRadius: isDesktop ? 20 : 16,
      padding: isDesktop ? "20px 22px" : "16px 15px",
      border: "1px solid #e8edf3",
      boxShadow: isDesktop ? "0 10px 28px rgba(15,23,42,0.08)" : "0 2px 8px rgba(15,23,42,0.07)",
      marginBottom: 10,
      position: "relative",
    }}>
      {isAdmin && (
        <div style={{ position: "absolute", top: 12, right: 12, display: "flex", gap: 6 }}>
          <button onClick={() => onEdit(contact)} style={{
            padding: "5px 10px", borderRadius: 8, border: "1.5px solid #bfdbfe",
            background: "#eff6ff", color: "#2563eb", fontWeight: 700, fontSize: 11,
            cursor: "pointer", fontFamily: FF,
          }}>수정</button>
          <button onClick={() => onDelete(contact)} style={{
            padding: "5px 10px", borderRadius: 8, border: "1.5px solid #fca5a5",
            background: "#fef2f2", color: "#dc2626", fontWeight: 700, fontSize: 11,
            cursor: "pointer", fontFamily: FF,
          }}>삭제</button>
        </div>
      )}

      <div style={{
        display: "flex",
        alignItems: isDesktop ? "center" : "flex-start",
        justifyContent: "space-between",
        gap: 12,
        marginBottom: 16,
        paddingRight: isAdmin ? 96 : 0,
      }}>
        <div>
          <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap", marginBottom: 8 }}>
            <div style={{ fontWeight: 900, fontSize: isDesktop ? 21 : 18, color: "#0f172a", lineHeight: 1.1 }}>
              {contact.name || "-"}
            </div>
            {contact.position && (
              <span style={{
                display: "inline-flex",
                alignItems: "center",
                padding: "5px 10px",
                borderRadius: 999,
                background: "#f8fafc",
                border: "1px solid #e2e8f0",
                color: "#475569",
                fontSize: 12,
                fontWeight: 700,
              }}>
                {contact.position}
              </span>
            )}
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
            <span style={{
              display: "inline-flex",
              alignItems: "center",
              padding: "5px 10px",
              borderRadius: 999,
              background: "#ecfeff",
              border: "1px solid #bae6fd",
              color: "#0f766e",
              fontSize: 12,
              fontWeight: 800,
            }}>
              {contact.dept || "기타"}
            </span>
            {(contact.ext || contact.direct) && (
              <span style={{ fontSize: 12, color: "#94a3b8", fontWeight: 700 }}>
                {contact.ext ? `내선 ${contact.ext}` : ""}
                {contact.ext && contact.direct ? " · " : ""}
                {contact.direct ? `직통 ${formatPhone(contact.direct)}` : ""}
              </span>
            )}
          </div>
        </div>

        {contact.phone && (
          <button onClick={revealPhone} style={{
            padding: isDesktop ? "8px 14px" : "6px 12px",
            borderRadius: 10,
            border: showPhone ? "1.5px solid #475569" : "1.5px solid #2563eb",
            background: showPhone ? "#f8fafc" : "#eff6ff",
            color: showPhone ? "#475569" : "#2563eb",
            fontWeight: 800,
            fontSize: 12,
            cursor: "pointer",
            fontFamily: FF,
            whiteSpace: "nowrap",
            flexShrink: 0,
          }}>
            {showPhone ? `숨기기 ${countdown}s` : "휴대전화 보기"}
          </button>
        )}
      </div>

      <div style={{
        display: "grid",
        gridTemplateColumns: isDesktop ? "minmax(180px, 1.3fr) repeat(2, minmax(120px, 1fr))" : "1fr",
        gap: 10,
      }}>
        {infoItem("휴대전화", visiblePhone, true)}
        {infoItem("내선번호", contact.ext)}
        {infoItem("직통번호", contact.direct ? formatPhone(contact.direct) : null, false, contact.direct ? `tel:${compactPhone(contact.direct)}` : undefined)}
      </div>
    </div>
  );
}

function ResponsiveContactsView({ contacts, isAdmin, showToast }) {
  const [search, setSearch] = useState("");
  const [expandedDepts, setExpandedDepts] = useState({});
  const [modalContact, setModalContact] = useState(undefined);
  const isDesktop = useMediaQuery("(min-width: 900px)");

  const filtered = useMemo(() => {
    const q = search.trim().toLowerCase();
    if (!q) return contacts;
    return contacts.filter((c) =>
      [c.name, c.dept, c.position, c.ext, c.direct].join(" ").toLowerCase().includes(q)
    );
  }, [contacts, search]);

  const grouped = useMemo(() => {
    const g = {};
    filtered.forEach((c) => {
      const dept = c.dept || "기타";
      if (!g[dept]) g[dept] = [];
      g[dept].push(c);
    });
    return g;
  }, [filtered]);

  const depts = Object.keys(grouped).sort((a, b) => a.localeCompare(b, "ko"));

  function expandAll() {
    setExpandedDepts(Object.fromEntries(depts.map((dept) => [dept, true])));
  }

  function collapseAll() {
    setExpandedDepts(Object.fromEntries(depts.map((dept) => [dept, false])));
  }

  async function handleSave(form) {
    try {
      if (modalContact?.id) {
        await updateContact(modalContact.id, form);
        showToast("수정되었습니다.");
      } else {
        await addContact(form);
        showToast("추가되었습니다.");
      }
    } catch (err) {
      console.error("연락처 저장 오류:", err);
      showToast(`저장 오류: ${err?.message || err}`, "error");
      throw err;
    }
  }

  async function handleDelete(contact) {
    if (!window.confirm(`"${contact.name}"을(를) 삭제할까요?`)) return;
    try {
      await deleteContact(contact.id);
      showToast("삭제되었습니다.");
    } catch {
      showToast("삭제 중 오류가 발생했습니다.", "error");
    }
  }

  return (
    <div style={{ maxWidth: 1180, margin: "0 auto", padding: isDesktop ? "28px 24px 88px" : "20px 16px 80px" }}>
      {modalContact !== undefined && (
        <ContactModal
          contact={modalContact}
          onSave={handleSave}
          onClose={() => setModalContact(undefined)}
        />
      )}

      <div style={{
        display: "flex",
        flexDirection: isDesktop ? "row" : "column",
        gap: 14,
        marginBottom: 22,
        alignItems: isDesktop ? "center" : "stretch",
      }}>

        <div style={{ fontSize: 13, color: "#64748b", lineHeight: 1.6 }}>
          PC에서는 더 정돈된 카드형으로, 모바일에서는 빠르게 훑기 좋게 보이도록 구성했습니다.
        </div>


        <div style={{
          flex: isDesktop ? "0 1 420px" : 1,
          width: isDesktop ? "auto" : "100%",
          background: "#fff",
          border: "1.5px solid #f1f5f9",
          borderRadius: 16,
          padding: "6px 14px",
          boxShadow: "0 8px 24px rgba(15,23,42,0.05)",
          display: "flex",
          alignItems: "center",
          gap: 8,
        }}>
          <span style={{ color: "#cbd5e1", fontSize: 15 }}>🔍</span>
          <input
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="이름, 부서, 직위로 검색"
            style={{ ...base.input, border: "none", padding: "8px 4px", boxShadow: "none", flex: 1 }}
          />
          {search && (
            <button onClick={() => setSearch("")} style={{
              background: "none", border: "none", cursor: "pointer", color: "#cbd5e1", fontSize: 14,
            }}>✕</button>
          )}
        </div>

        {isAdmin && (
          <button onClick={() => setModalContact(null)} style={{
            padding: "12px 18px", borderRadius: 14, border: "none",
            background: "#2563eb", color: "#fff", fontWeight: 800, fontSize: 13,
            cursor: "pointer", fontFamily: FF, whiteSpace: "nowrap",
          }}>+ 추가</button>
        )}
      </div>

      <div style={{ fontSize: 13, color: "#64748b", fontWeight: 700, marginBottom: 18 }}>
        총 {contacts.length}명{search ? ` · 검색 결과 ${filtered.length}명` : ""}
      </div>

      {depts.length > 0 && (
        <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", marginBottom: 14 }}>
          <button
            onClick={expandAll}
            style={{
              padding: "8px 12px",
              borderRadius: 10,
              border: "1px solid #cbd5e1",
              background: "#fff",
              color: "#475569",
              fontWeight: 700,
              fontSize: 12,
              cursor: "pointer",
              fontFamily: FF,
            }}
          >
            전체 펼치기
          </button>
          <button
            onClick={collapseAll}
            style={{
              padding: "8px 12px",
              borderRadius: 10,
              border: "1px solid #cbd5e1",
              background: "#fff",
              color: "#475569",
              fontWeight: 700,
              fontSize: 12,
              cursor: "pointer",
              fontFamily: FF,
            }}
          >
            전체 접기
          </button>
        </div>
      )}

      {depts.length === 0 && (
        <div style={{ textAlign: "center", padding: "60px 0", color: "#cbd5e1" }}>
          <div style={{ fontSize: 36, marginBottom: 12 }}>📇</div>
          <div style={{ fontSize: 15 }}>
            {search ? "검색 결과가 없습니다." : "등록된 연락처가 없습니다."}
          </div>
          {isAdmin && !search && (
            <button onClick={() => setModalContact(null)} style={{
              marginTop: 16, padding: "10px 20px", borderRadius: 12, border: "none",
              background: "#2563eb", color: "#fff", fontWeight: 800, fontSize: 13,
              cursor: "pointer", fontFamily: FF,
            }}>+ 첫 연락처 추가</button>
          )}
        </div>
      )}

      {depts.map((dept) => {
        const isExpanded = expandedDepts[dept] !== false;
        const items = grouped[dept];
        return (
          <div key={dept} style={{
            marginBottom: 18,
            background: "#fff",
            border: "1px solid #e5edf5",
            borderRadius: isDesktop ? 22 : 14,
            padding: isDesktop ? "14px" : "10px",
            boxShadow: isDesktop ? "0 14px 36px rgba(15,23,42,0.05)" : "none",
          }}>
            <div
              onClick={() => setExpandedDepts((p) => ({ ...p, [dept]: !isExpanded }))}
              style={{
                display: "flex",
                alignItems: "center",
                gap: 10,
                padding: isDesktop ? "14px 16px" : "10px 12px",
                borderRadius: 14,
                background: "#f8fafc",
                cursor: "pointer",
                marginBottom: isExpanded ? 8 : 0,
                userSelect: "none",
              }}
            >
              <span style={{ fontWeight: 900, fontSize: isDesktop ? 18 : 14, color: "#0f172a" }}>{dept}</span>
              <span style={{ fontSize: 12, color: "#94a3b8", fontWeight: 700 }}>{items.length}명</span>
              <div style={{ flex: 1 }} />
              <span style={{
                fontSize: 12,
                color: "#94a3b8",
                transform: isExpanded ? "rotate(0deg)" : "rotate(-90deg)",
                transition: "transform 0.2s",
                display: "inline-block",
              }}>⌄</span>
            </div>

            {isExpanded && (
              <div style={{
                display: "grid",
                gridTemplateColumns: isDesktop ? "repeat(auto-fit, minmax(360px, 1fr))" : "1fr",
                gap: isDesktop ? 12 : 0,
                paddingTop: 4,
              }}>
                {items.map((c) => (
                  <SimpleContactCard
                    key={c.id}
                    contact={c}
                    isAdmin={isAdmin}
                    isDesktop={isDesktop}
                    onEdit={(ct) => setModalContact(ct)}
                    onDelete={handleDelete}
                  />
                ))}
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}

function SimpleContactCard({ contact, isAdmin, isDesktop, onEdit, onDelete }) {
  const [showPhone, setShowPhone] = useState(false);
  const [countdown, setCountdown] = useState(0);
  const timerRef = useRef(null);
  const hasExt = !!String(contact.ext || "").trim();
  const hasDirect = !!String(contact.direct || "").trim();

  useEffect(() => () => clearInterval(timerRef.current), []);

  function revealPhone() {
    if (showPhone) {
      clearInterval(timerRef.current);
      setShowPhone(false);
      setCountdown(0);
      return;
    }

    setShowPhone(true);
    setCountdown(10);
    clearInterval(timerRef.current);
    timerRef.current = setInterval(() => {
      setCountdown((prev) => {
        if (prev <= 1) {
          clearInterval(timerRef.current);
          setShowPhone(false);
          return 0;
        }
        return prev - 1;
      });
    }, 1000);
  }

  const visiblePhone = showPhone
    ? formatPhone(contact.phone) || "-"
    : (contact.phone ? "•••-••••-••••" : "-");

  return (
    <div style={{
      background: "#fff",
      borderRadius: isDesktop ? 18 : 14,
      padding: isDesktop ? "18px 20px" : "14px 14px",
      border: "1px solid #e8edf3",
      boxShadow: isDesktop ? "0 8px 24px rgba(15,23,42,0.06)" : "0 2px 8px rgba(15,23,42,0.06)",
      position: "relative",
    }}>
      {isAdmin && (
        <div style={{ position: "absolute", top: 12, right: 12, display: "flex", gap: 6 }}>
          <button onClick={() => onEdit(contact)} style={{
            padding: "5px 10px", borderRadius: 8, border: "1.5px solid #bfdbfe",
            background: "#eff6ff", color: "#2563eb", fontWeight: 700, fontSize: 11,
            cursor: "pointer", fontFamily: FF,
          }}>수정</button>
          <button onClick={() => onDelete(contact)} style={{
            padding: "5px 10px", borderRadius: 8, border: "1.5px solid #fca5a5",
            background: "#fef2f2", color: "#dc2626", fontWeight: 700, fontSize: 11,
            cursor: "pointer", fontFamily: FF,
          }}>삭제</button>
        </div>
      )}

      <div style={{ display: "grid", gap: 10, paddingRight: isAdmin ? 96 : 0 }}>
        <div style={{ display: "flex", alignItems: "baseline", gap: 8, flexWrap: "wrap" }}>
          <div style={{ fontWeight: 900, fontSize: isDesktop ? 20 : 17, color: "#0f172a", lineHeight: 1.1 }}>
            {contact.name || "-"}
          </div>
          {contact.dept && (
            <span style={{ fontSize: 14, fontWeight: 700, color: "#0891b2" }}>
              {contact.dept}
            </span>
          )}
          {contact.position && (
            <span style={{ fontSize: 14, fontWeight: 700, color: "#64748b" }}>
              {contact.position}
            </span>
          )}
        </div>

        <div style={{
          display: "flex",
          flexDirection: isDesktop ? "row" : "column",
          alignItems: isDesktop ? "center" : "flex-start",
          gap: isDesktop ? 14 : 8,
          fontSize: 14,
          fontWeight: 700,
          color: "#334155",
        }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
            <span style={{ color: "#64748b" }}>휴대전화</span>
            {contact.phone && (
              <button onClick={revealPhone} style={{
                padding: "4px 12px",
                borderRadius: 10,
                border: showPhone ? "1.5px solid #475569" : "1.5px solid #2563eb",
                background: "#fff",
                color: showPhone ? "#475569" : "#2563eb",
                fontWeight: 800,
                fontSize: 12,
                cursor: "pointer",
                fontFamily: FF,
                whiteSpace: "nowrap",
              }}>
                {showPhone ? `숨기기 ${countdown}s` : "보기"}
              </button>
            )}
            <span style={{ color: "#0f172a", fontWeight: 800 }}>{visiblePhone}</span>
          </div>

          {hasExt && (
            <>
              <span style={{ color: "#cbd5e1", display: isDesktop ? "inline" : "none" }}>|</span>
              <span><span style={{ color: "#64748b" }}>내선번호</span>{` ${contact.ext}`}</span>
            </>
          )}

          {hasDirect && (
            <>
              <span style={{ color: "#cbd5e1", display: isDesktop ? "inline" : "none" }}>|</span>
              <span>
                <span style={{ color: "#64748b" }}>직통번호</span>
                {" "}
                <a href={`tel:${compactPhone(contact.direct)}`} style={{ color: "#0f172a", textDecoration: "none", fontWeight: 800 }}>
                  {formatPhone(contact.direct)}
                </a>
              </span>
            </>
          )}
        </div>

        {historySchedules.length > 0 && (
          <div style={{ display: "grid", gap: 10 }}>
            <button
              onClick={() => setHistoryOpen((prev) => !prev)}
              style={{
                justifySelf: "flex-start",
                padding: "6px 12px",
                borderRadius: 10,
                border: "1px solid #cbd5e1",
                background: "#fff",
                color: "#475569",
                fontWeight: 700,
                fontSize: 12,
                cursor: "pointer",
                fontFamily: FF,
              }}
            >
              {historyOpen ? `일정 히스토리 닫기` : `일정 히스토리 ${historySchedules.length}건`}
            </button>

            {historyOpen && (
              <div style={{
                display: "grid",
                gap: 8,
                padding: "12px 14px",
                borderRadius: 12,
                background: "#f8fafc",
                border: "1px solid #e2e8f0",
              }}>
                {historySchedules.map((schedule) => {
                  const meta = SCHEDULE_META[schedule.type] || { label: schedule.type, icon: "📌" };
                  const status = getScheduleStatus(schedule, todayYmd);
                  const statusLabel = status === "active" ? "현재" : status === "upcoming" ? "예정" : "종료";
                  const statusColor = status === "active" ? "#2563eb" : status === "upcoming" ? "#7c3aed" : "#94a3b8";

                  return (
                    <div
                      key={schedule.id}
                      style={{
                        display: "flex",
                        flexDirection: isDesktop ? "row" : "column",
                        flexWrap: "nowrap",
                        alignItems: isDesktop ? "center" : "flex-start",
                        gap: 8,
                        padding: "10px 12px",
                        borderRadius: 10,
                        background: "#fff",
                        border: "1px solid #e2e8f0",
                      }}
                    >
                      <span style={{ fontSize: 13, fontWeight: 800, color: "#0f172a", whiteSpace: "nowrap" }}>
                        {meta.icon} {meta.label}
                      </span>
                      <span style={{ fontSize: 12, fontWeight: 800, color: statusColor, whiteSpace: "nowrap" }}>
                        {statusLabel}
                      </span>
                      <span style={{ fontSize: 12, color: "#475569", fontWeight: 700, whiteSpace: "nowrap" }}>
                        {formatSchedulePeriod(schedule)}
                      </span>
                      {schedule.note && (
                        <span style={{ fontSize: 12, color: "#64748b" }}>
                          메모: {schedule.note}
                        </span>
                      )}
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}

const SCHEDULE_META = {
  maternity_leave: { label: "출산휴가", icon: "👶" },
  parental_leave: { label: "육아휴직", icon: "👶" },
  leave_of_absence: { label: "휴직", icon: "🗓️" },
  annual_leave: { label: "연차", icon: "✈️" },
  half_day_am: { label: "오전 반차", icon: "✈️" },
  half_day_pm: { label: "오후 반차", icon: "✈️" },
};

SCHEDULE_META.compensatory_leave = { label: "대체휴무", icon: "✈️" };
const EXTRA_SCHEDULE_TYPE_MAP = {
  "출산전후휴가": "maternity_leave",
  "대체휴무": "compensatory_leave",
  // Attendance (근태) - B~I 업로드 양식의 "휴가종류" 칸에 넣으면 근태로 저장됩니다.
  "업무": "work_task",
  "외근": "field_work",
  "출장": "business_trip",
  "연장": "overtime",
  "재택": "remote_work",
};

// Attendance (근태) types - used for employee self-entry in calendar.
SCHEDULE_META.work_task = { label: "\uC5C5\uBB34", icon: "📝" };
SCHEDULE_META.field_work = { label: "\uC678\uADFC", icon: "🚗" };
SCHEDULE_META.business_trip = { label: "\uCD9C\uC7A5", icon: "🚀" };
SCHEDULE_META.overtime = { label: "\uC5F0\uC7A5", icon: "⏱️" };
SCHEDULE_META.remote_work = { label: "\uC7AC\uD0DD", icon: "🏠" };

function getTodayYmd() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const day = String(now.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function normalizeDateYmd(value) {
  if (value == null) return "";
  const raw = String(value).trim();
  if (!raw) return "";

  if (/^\d{5}$/.test(raw)) {
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    excelEpoch.setUTCDate(excelEpoch.getUTCDate() + Number(raw));
    return excelEpoch.toISOString().slice(0, 10);
  }

  const normalized = raw.replace(/[./]/g, "-").replace(/\s+/g, "");
  if (/^\d{4}-\d{1,2}-\d{1,2}$/.test(normalized)) {
    const [year, month, day] = normalized.split("-");
    return `${year}-${month.padStart(2, "0")}-${day.padStart(2, "0")}`;
  }

  return "";
}

function normalizeScheduleType(value) {
  const text = String(value || "").trim().replace(/\s+/g, "");
  const map = {
    "출산휴가": "maternity_leave",
    "육아휴직": "parental_leave",
    "휴직": "leave_of_absence",
    "연차": "annual_leave",
    "오전반차": "half_day_am",
    "오후반차": "half_day_pm",
  };
  return map[text] || EXTRA_SCHEDULE_TYPE_MAP[text] || "";
}

function normalizeAnnualLeaveByTime(value) {
  const text = String(value || "").trim().replace(/\s+/g, "");
  if (!text) return "annual_leave";
  if (["오전", "반차오전", "오전반차"].includes(text)) return "half_day_am";
  if (["오후", "반차오후", "오후반차"].includes(text)) return "half_day_pm";
  if (["1일", "1일+a", "1일+a)", "1일+a(", "1일+a연차", "연차-1일+a"].includes(text)) return "annual_leave";
  return "annual_leave";
}

function tokenizeScheduleLine(line) {
  const raw = String(line || "").trim();
  if (!raw) return [];
  if (raw.includes("\t")) return raw.split("\t").map((cell) => cell.trim());
  return raw.split(/\s+/).map((cell) => cell.trim()).filter(Boolean);
}

function parseSchedulePasteRow(cells) {
  let name = "";
  let dept = "";
  let position = "";
  let type = "";
  let startDate = "";
  let endDate = "";
  let detail = "";

  const trimmed = cells.map((cell) => String(cell || "").trim());

  if (trimmed.length >= 8) {
    const [
      startRaw,
      endRaw,
      requestTime,
      daysText,
      personName,
      positionText,
      deptText,
      leaveTypeText,
      detailText = "",
    ] = trimmed;

    const normalizedType = normalizeScheduleType(leaveTypeText);
    if (personName && deptText && (normalizedType || leaveTypeText === "연차")) {
      name = personName;
      dept = deptText;
      position = positionText;
      startDate = normalizeDateYmd(startRaw);
      endDate = normalizeDateYmd(endRaw) || startDate;
      detail = String(detailText || "").trim();

      if (leaveTypeText === "연차") type = normalizeAnnualLeaveByTime(requestTime);
      else type = normalizedType;

      // If it's an attendance type and no detail is provided, fallback to requestTime (e.g. 오전/오후).
      if (!detail && type && getScheduleCategoryByType(type) === "attendance") {
        detail = String(requestTime || "").trim();
      }

      return { name, dept, position, type, startDate, endDate, detail };
    }
  }

  let typeText = "";
  let startRaw = "";
  let endRaw = "";
  let detailText = "";

  if (normalizeScheduleType(trimmed[1])) {
    [name, typeText, startRaw, endRaw = "", detailText = ""] = trimmed;
  } else {
    [name, dept = "", typeText, startRaw, endRaw = "", detailText = ""] = trimmed;
  }

  type = normalizeScheduleType(typeText);
  startDate = normalizeDateYmd(startRaw);
  endDate = normalizeDateYmd(endRaw) || startDate;
  detail = String(detailText || "").trim();

  return { name, dept, position, type, startDate, endDate, detail };
}

function formatSchedulePeriod(schedule) {
  if (!schedule?.startDate) return "";
  if (!schedule.endDate || schedule.endDate === schedule.startDate) return schedule.startDate;
  return `${schedule.startDate} ~ ${schedule.endDate}`;
}

function cleanSchedulePayload(schedule) {
  return {
    id: schedule?.id || `${Date.now()}_${Math.random().toString(36).slice(2, 8)}`,
    type: String(schedule?.type || "").trim(),
    detail: String(schedule?.detail || "").trim(),
    startDate: normalizeDateYmd(schedule?.startDate),
    endDate: normalizeDateYmd(schedule?.endDate) || normalizeDateYmd(schedule?.startDate),
    createdAt: schedule?.createdAt || new Date().toISOString(),
  };
}

function isSameSchedule(a, b) {
  const leaveSet = new Set([
    "annual_leave",
    "half_day_am",
    "half_day_pm",
    "compensatory_leave",
    "maternity_leave",
    "parental_leave",
    "leave_of_absence",
  ]);
  const aType = String(a?.type || "");
  const bType = String(b?.type || "");
  const isAttendance = !leaveSet.has(aType) && !leaveSet.has(bType);
  const sameDetail = !isAttendance || String(a?.detail || "") === String(b?.detail || "");

  return (
    aType === bType &&
    String(a?.startDate || "") === String(b?.startDate || "") &&
    String(a?.endDate || a?.startDate || "") === String(b?.endDate || b?.startDate || "") &&
    sameDetail
  );
}

function dedupeSchedules(list) {
  const next = [];
  for (const item of Array.isArray(list) ? list : []) {
    const cleaned = cleanSchedulePayload(item);
    if (!cleaned.type || !cleaned.startDate) continue;
    if (next.some((saved) => isSameSchedule(saved, cleaned))) continue;
    next.push(cleaned);
  }
  return next;
}

function getActiveSchedules(contact, todayYmd) {
  const list = Array.isArray(contact?.schedules) ? contact.schedules : [];
  return list
    .filter((schedule) => {
      if (!schedule?.startDate) return false;
      const endDate = schedule.endDate || schedule.startDate;
      return schedule.startDate <= todayYmd && todayYmd <= endDate;
    })
    .sort((a, b) => String(a.startDate).localeCompare(String(b.startDate)));
}

function getScheduleStatus(schedule, todayYmd) {
  if (!schedule?.startDate) return "unknown";
  const endDate = schedule.endDate || schedule.startDate;
  if (todayYmd < schedule.startDate) return "upcoming";
  if (todayYmd > endDate) return "ended";
  return "active";
}

function getScheduleHistory(contact) {
  const list = Array.isArray(contact?.schedules) ? contact.schedules : [];
  return [...list].sort((a, b) => {
    const aDate = `${a.startDate || ""}_${a.endDate || ""}`;
    const bDate = `${b.startDate || ""}_${b.endDate || ""}`;
    return bDate.localeCompare(aDate);
  });
}

function getDeptBadge(dept) {
  const d = String(dept || "");
  if (d.includes("관리")) return { label: "\uAD00", bg: "#eff6ff", fg: "#2563eb", bd: "#bfdbfe" };
  if (d.includes("시스템")) return { label: "\uC2DC", bg: "#ecfeff", fg: "#0e7490", bd: "#a5f3fc" };
  if (d.includes("영업")) return { label: "\uC601", bg: "#fff7ed", fg: "#c2410c", bd: "#fed7aa" };
  const label = d.trim() ? d.trim().slice(0, 1) : "\uAE30";
  return { label, bg: "#f1f5f9", fg: "#334155", bd: "#cbd5e1" };
}

function getScheduleCategoryByType(type) {
  const t = String(type || "");
  const leaveSet = new Set([
    "annual_leave",
    "half_day_am",
    "half_day_pm",
    "compensatory_leave",
    "maternity_leave",
    "parental_leave",
    "leave_of_absence",
  ]);
  return leaveSet.has(t) ? "leave" : "attendance";
}

function isOwnContactForUser(contact, userProfile) {
  const targetName = String(userProfile?.name || "").trim();
  const targetDept = String(userProfile?.department || "").trim();
  if (!targetName) return false;
  const isSameName = String(contact?.name || "").trim() === targetName;
  const isSameDept = !targetDept || String(contact?.dept || "").trim() === targetDept;
  return isSameName && isSameDept;
}

function canEditSchedule({ isAdmin, contact, schedule, userProfile }) {
  if (!contact || !schedule) return false;
  if (isAdmin) return true;
  return isOwnContactForUser(contact, userProfile) && getScheduleCategoryByType(schedule.type) === "attendance";
}

function DayDeptEventsModal({ ymd, dept, items, onOpenEmployee, onClose }) {
  return (
    <Modal
      title={`${dept || ""} ${ymd}`}
      onClose={onClose}
      extra={<span style={{ fontSize: 12, color: "#64748b", fontWeight: 800 }}>{`${items.length}\uBA85`}</span>}
    >
      <div style={{ display: "grid", gap: 8 }}>
        {items.length === 0 && (
          <div style={{ color: "#64748b", fontWeight: 700, padding: "10px 2px" }}>
            {`\uB4F1\uB85D\uB41C \uB0B4\uC5ED\uC774 \uC5C6\uC2B5\uB2C8\uB2E4.`}
          </div>
        )}
        {items.map((row) => (
          <button
            key={row.contact.id}
            onClick={() => onOpenEmployee(row.contact, row.schedules)}
            style={{
              textAlign: "left",
              border: "1px solid #e2e8f0",
              background: "#fff",
              borderRadius: 12,
              padding: "10px 12px",
              cursor: "pointer",
              fontFamily: FF,
            }}
          >
            <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
              <span style={{ fontWeight: 900, color: "#0f172a" }}>{row.contact.name}</span>
              {row.contact.position && <span style={{ fontSize: 12, fontWeight: 800, color: "#64748b" }}>{row.contact.position}</span>}
              <span style={{ fontSize: 12, fontWeight: 800, color: "#94a3b8" }}>{`${row.schedules.length}\uAC74`}</span>
              <div style={{ flex: 1 }} />
              <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                {row.schedules.slice(0, 6).map((s) => {
                  const meta = SCHEDULE_META[s.type] || { icon: "📌" };
                  return <span key={s.id} style={{ fontSize: 14, lineHeight: 1 }}>{meta.icon || "📌"}</span>;
                })}
                {row.schedules.length > 6 && (
                  <span style={{ fontSize: 11, fontWeight: 900, color: "#64748b" }}>{`+${row.schedules.length - 6}`}</span>
                )}
              </div>
            </div>
          </button>
        ))}
      </div>
    </Modal>
  );
}

function DayEmployeeEventsModal({ ymd, contact, schedules, isAdmin, userProfile, onUpdateSchedule, onDeleteSchedule, onClose }) {
  const [editingSchedule, setEditingSchedule] = useState(null);
  const editableScheduleTypes = useMemo(
    () => Object.entries(SCHEDULE_META).filter(([value]) => isAdmin || getScheduleCategoryByType(value) === "attendance"),
    [isAdmin]
  );

  return (
    <Modal
      title={`${contact?.name || ""} ${ymd}`}
      onClose={onClose}
    >
      {editingSchedule && (
        <ContactScheduleEditModal
          schedule={editingSchedule}
          typeOptions={editableScheduleTypes}
          onSave={async (nextSchedule) => {
            await onUpdateSchedule(contact, editingSchedule.id, nextSchedule);
            setEditingSchedule(null);
          }}
          onClose={() => setEditingSchedule(null)}
        />
      )}
      <div style={{ display: "grid", gap: 8 }}>
        {schedules.length === 0 && (
          <div style={{ color: "#64748b", fontWeight: 700, padding: "10px 2px" }}>
            {`\uB4F1\uB85D\uB41C \uB0B4\uC5ED\uC774 \uC5C6\uC2B5\uB2C8\uB2E4.`}
          </div>
        )}
        {schedules.map((s) => {
          const meta = SCHEDULE_META[s.type] || { label: s.type, icon: "📌" };
          const canEdit = canEditSchedule({ isAdmin, contact, schedule: s, userProfile });
          return (
            <div
              key={s.id}
              style={{
                padding: "10px 12px",
                borderRadius: 12,
                background: "#fff",
                border: "1px solid #e2e8f0",
                fontFamily: FF,
                textAlign: "left",
              }}
            >
              <div style={{ display: "flex", alignItems: "baseline", gap: 10, minWidth: 0, justifyContent: "space-between", flexWrap: "wrap" }}>
                <div style={{ display: "flex", alignItems: "baseline", gap: 10, minWidth: 0, flexWrap: "wrap" }}>
                  <span style={{ fontWeight: 900, flexShrink: 0 }}>{meta.icon || "📌"}</span>
                  <span style={{ fontWeight: 900, color: "#0f172a", flexShrink: 0 }}>{meta.label || s.type}</span>
                  <span style={{ color: "#64748b", fontWeight: 800, fontSize: 12, minWidth: 0, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                    {formatSchedulePeriod(s)}
                  </span>
                </div>
                {canEdit && (
                  <div style={{ display: "inline-flex", alignItems: "center", gap: 6, flexWrap: "nowrap", whiteSpace: "nowrap" }}>
                    <button
                      onClick={() => setEditingSchedule(s)}
                      style={{ padding: "4px 8px", borderRadius: 8, border: "1px solid #bfdbfe", background: "#eff6ff", color: "#2563eb", fontWeight: 700, fontSize: 11, cursor: "pointer", fontFamily: FF, whiteSpace: "nowrap", flexShrink: 0 }}
                    >
                      {"수정"}
                    </button>
                    {isAdmin && (
                      <button
                        onClick={async () => {
                          if (!window.confirm("이 일정을 삭제할까요?")) return;
                          await onDeleteSchedule(contact, s.id);
                        }}
                        style={{ padding: "4px 8px", borderRadius: 8, border: "1px solid #fecaca", background: "#fff1f2", color: "#dc2626", fontWeight: 700, fontSize: 11, cursor: "pointer", fontFamily: FF, whiteSpace: "nowrap", flexShrink: 0 }}
                      >
                        {"삭제"}
                      </button>
                    )}
                  </div>
                )}
              </div>
              {!!String(s.detail || "").trim() && (
                <div style={{ marginTop: 4, fontSize: 12, color: "#475569", fontWeight: 700, whiteSpace: "pre-wrap", overflowWrap: "anywhere" }}>
                  {s.detail}
                </div>
              )}
            </div>
          );
        })}
      </div>
    </Modal>
  );
}

function DayAllEventsModal({ ymd, items, meContact, onOpenEmployee, onAddForMe, onClose }) {
  const attendance = items.filter((x) => x.category === "attendance");
  const leave = items.filter((x) => x.category === "leave");

  function Row({ item }) {
    const meta = SCHEDULE_META[item.schedule.type] || { label: item.schedule.type, icon: "📌" };
    const badge = getDeptBadge(item.dept);
    return (
      <button
        onClick={() => onOpenEmployee(item.contact, item.contactSchedules)}
        style={{
          width: "100%",
          textAlign: "left",
          border: "1px solid #e2e8f0",
          background: "#fff",
          borderRadius: 12,
          padding: "10px 12px",
          cursor: "pointer",
          fontFamily: FF,
        }}
      >
        <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
          <span style={{ fontSize: 16, lineHeight: 1, flexShrink: 0 }}>{meta.icon || "📌"}</span>
          <span style={{ fontWeight: 900, color: "#0f172a", flexShrink: 0 }}>{item.contact.name}</span>
          <span
            title={item.dept}
            style={{
              border: `1px solid ${badge.bd}`,
              background: badge.bg,
              color: badge.fg,
              borderRadius: 10,
              padding: "2px 7px",
              fontWeight: 900,
              fontSize: 12,
              lineHeight: 1.1,
              whiteSpace: "nowrap",
              flexShrink: 0,
            }}
          >
            {badge.label}
          </span>
          <span style={{ fontWeight: 900, color: "#334155" }}>{meta.label || item.schedule.type}</span>
        </div>
        {!!String(item.schedule.detail || "").trim() && (
          <div style={{ marginTop: 4, fontSize: 12, color: "#475569", fontWeight: 700, whiteSpace: "pre-wrap", overflowWrap: "anywhere" }}>
            {item.schedule.detail}
          </div>
        )}
      </button>
    );
  }

  return (
    <Modal
      title={`${ymd}`}
      onClose={onClose}
      extra={<span style={{ fontSize: 12, color: "#64748b", fontWeight: 800 }}>{`${items.length}\uAC74`}</span>}
    >
      <div style={{ display: "grid", gap: 12 }}>
        {meContact && (
          <button
            onClick={onAddForMe}
            style={{
              justifySelf: "flex-start",
              padding: "10px 12px",
              borderRadius: 12,
              border: "1px solid #bfdbfe",
              background: "#eff6ff",
              color: "#2563eb",
              fontWeight: 900,
              cursor: "pointer",
              fontFamily: FF,
              whiteSpace: "nowrap",
            }}
          >
            {`\uB0B4 \uADFC\uD0DC \uC785\uB825`}
          </button>
        )}
        <div style={{ display: "grid", gap: 8 }}>
          <div style={{ fontWeight: 900, color: "#0f172a" }}>{`[\uADFC\uD0DC]`}</div>
          {attendance.length === 0 ? (
            <div style={{ color: "#64748b", fontWeight: 700, padding: "6px 2px" }}>{`\uC5C6\uC74C`}</div>
          ) : (
            attendance.map((item) => <Row key={item.key} item={item} />)
          )}
        </div>

        <div style={{ display: "grid", gap: 8 }}>
          <div style={{ fontWeight: 900, color: "#0f172a" }}>{`[\uD734\uAC00]`}</div>
          {leave.length === 0 ? (
            <div style={{ color: "#64748b", fontWeight: 700, padding: "6px 2px" }}>{`\uC5C6\uC74C`}</div>
          ) : (
            leave.map((item) => <Row key={item.key} item={item} />)
          )}
        </div>

        <div style={{ color: "#64748b", fontWeight: 700, fontSize: 12 }}>
          {`\uD56D\uBAA9\uC744 \uB204\uB974\uBA74 \uD574\uB2F9 \uC9C1\uC6D0\uC758 \uC0C1\uC138 \uB0B4\uC5ED\uC744 \uD655\uC778\uD569\uB2C8\uB2E4.`}
        </div>
      </div>
    </Modal>
  );
}

function ContactScheduleCalendarDept({ depts, grouped, isDesktop, todayYmd, isAdmin, userProfile, onAddSchedule, onUpdateSchedule, onDeleteSchedule }) {
  const [monthCursor, setMonthCursor] = useState(() => {
    const d = new Date();
    d.setHours(0, 0, 0, 0);
    d.setDate(1);
    return d;
  });
  const [selectedDept, setSelectedDept] = useState(null);
  const [selectedYmd, setSelectedYmd] = useState("");
  const [dayModalYmd, setDayModalYmd] = useState("");
  const [employeeModal, setEmployeeModal] = useState(null); // { ymd, contact, schedules }
  const [attendanceAdd, setAttendanceAdd] = useState(null); // { ymd, contact, returnTo?: { type: "dayAll", ymd } }

  const meContact = useMemo(() => {
    const targetName = String(userProfile?.name || "").trim();
    const targetDept = String(userProfile?.department || "").trim();
    if (!targetName) return null;

    if (targetDept && grouped?.[targetDept]) {
      const candidates = (grouped[targetDept] || []).filter((c) => String(c.name || "").trim() === targetName);
      if (candidates.length === 1) return candidates[0];
    }

    const matches = [];
    for (const d of depts || []) {
      for (const c of grouped?.[d] || []) {
        if (String(c.name || "").trim() === targetName) matches.push(c);
      }
    }
    return matches.length === 1 ? matches[0] : null;
  }, [depts, grouped, userProfile]);

  function dateToYmd(dt) {
    const y = dt.getFullYear();
    const m = String(dt.getMonth() + 1).padStart(2, "0");
    const d = String(dt.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }

  function ymdToDate(ymd) {
    if (!ymd) return null;
    const [y, m, d] = String(ymd).split("-").map((x) => Number(x));
    if (!y || !m || !d) return null;
    const dt = new Date(y, m - 1, d);
    dt.setHours(0, 0, 0, 0);
    return dt;
  }

  function addMonths(dt, delta) {
    const next = new Date(dt);
    next.setMonth(next.getMonth() + delta);
    next.setDate(1);
    next.setHours(0, 0, 0, 0);
    return next;
  }

  const monthLabel = useMemo(() => {
    const y = monthCursor.getFullYear();
    const m = monthCursor.getMonth() + 1;
    return `${y}-${String(m).padStart(2, "0")}`;
  }, [monthCursor]);

  const calendarCells = useMemo(() => {
    const first = new Date(monthCursor);
    first.setDate(1);
    const last = new Date(monthCursor);
    last.setMonth(last.getMonth() + 1);
    last.setDate(0);

    const firstDow = (first.getDay() + 6) % 7; // Mon-start
    const start = new Date(first);
    start.setDate(first.getDate() - firstDow);

    const cells = [];
    for (let i = 0; i < 42; i++) {
      const d = new Date(start);
      d.setDate(start.getDate() + i);
      const ymd = dateToYmd(d);
      cells.push({ date: d, ymd, inMonth: d.getMonth() === monthCursor.getMonth() });
    }

    return { first, last, cells };
  }, [monthCursor]);

  const deptDayIndex = useMemo(() => {
    const monthStartYmd = dateToYmd(calendarCells.first);
    const monthEndYmd = dateToYmd(calendarCells.last);
    const map = new Map(); // ymd -> dept -> contactId -> {contact, schedules}

    for (const dept of depts) {
      const contacts = grouped?.[dept] || [];
      for (const contact of contacts) {
        const schedules = Array.isArray(contact?.schedules) ? contact.schedules : [];
        for (const s of schedules) {
          if (!s?.startDate) continue;
          const start = s.startDate;
          const end = s.endDate || s.startDate;
          if (end < monthStartYmd || start > monthEndYmd) continue;

          const cursor = ymdToDate(start);
          const endDt = ymdToDate(end);
          if (!cursor || !endDt) continue;

          for (let dt = new Date(cursor); dt <= endDt; dt.setDate(dt.getDate() + 1)) {
            const ymd = dateToYmd(dt);
            if (ymd < monthStartYmd || ymd > monthEndYmd) continue;
            if (!map.has(ymd)) map.set(ymd, new Map());
            const deptMap = map.get(ymd);
            if (!deptMap.has(dept)) deptMap.set(dept, new Map());
            const contactMap = deptMap.get(dept);
            if (!contactMap.has(contact.id)) contactMap.set(contact.id, { contact, schedules: [] });
            contactMap.get(contact.id).schedules.push(s);
          }
        }
      }
    }

    // sort schedules within each contact bucket
    for (const deptMap of map.values()) {
      for (const contactMap of deptMap.values()) {
        for (const bucket of contactMap.values()) {
          bucket.schedules.sort((a, b) => String(a.type || "").localeCompare(String(b.type || "")));
        }
      }
    }

    return map;
  }, [calendarCells.first, calendarCells.last, depts, grouped]);

  const openDeptItems = useMemo(() => {
    if (!selectedDept || !selectedYmd) return [];
    const deptMap = deptDayIndex.get(selectedYmd);
    const contactMap = deptMap?.get(selectedDept);
    if (!contactMap) return [];
    const items = [...contactMap.values()];
    items.sort((a, b) => String(a.contact?.name || "").localeCompare(String(b.contact?.name || ""), "ko"));
    return items;
  }, [deptDayIndex, selectedDept, selectedYmd]);

  const dayAllItems = useMemo(() => {
    if (!dayModalYmd) return [];
    const deptMap = deptDayIndex.get(dayModalYmd);
    if (!deptMap) return [];

    const rows = [];
    for (const [dept, contactMap] of deptMap.entries()) {
      for (const bucket of contactMap.values()) {
        const contactSchedules = bucket.schedules || [];
        for (const s of contactSchedules) {
          const meta = SCHEDULE_META[s.type] || { label: s.type, icon: "📌" };
          const category = getScheduleCategoryByType(s.type);
          rows.push({
            key: `${dept}_${bucket.contact.id}_${s.id}`,
            dept,
            contact: bucket.contact,
            schedule: s,
            category,
            sortKey: `${category}_${dept}_${bucket.contact.name}_${meta.label || ""}`,
            contactSchedules,
          });
        }
      }
    }

    rows.sort((a, b) => String(a.sortKey).localeCompare(String(b.sortKey), "ko"));
    return rows;
  }, [dayModalYmd, deptDayIndex]);

  return (
    <>
      {dayModalYmd && (
        <DayAllEventsModal
          ymd={dayModalYmd}
          items={dayAllItems}
          meContact={meContact}
          onAddForMe={() => {
            if (!meContact) return;
            const ymd = dayModalYmd;
            // Close the viewing modal first, then open the input modal on top.
            setDayModalYmd("");
            setSelectedDept(null);
            setSelectedYmd("");
            setEmployeeModal(null);
            setAttendanceAdd({ ymd, contact: meContact, returnTo: { type: "dayAll", ymd } });
          }}
          onOpenEmployee={(contact, schedules) => setEmployeeModal({ ymd: dayModalYmd, contact, schedules })}
          onClose={() => setDayModalYmd("")}
        />
      )}
      {selectedDept && selectedYmd && (
        <DayDeptEventsModal
          ymd={selectedYmd}
          dept={selectedDept}
          items={openDeptItems}
          onOpenEmployee={(contact, schedules) => setEmployeeModal({ ymd: selectedYmd, contact, schedules })}
          onClose={() => { setSelectedDept(null); setSelectedYmd(""); }}
        />
      )}
      {employeeModal && (
        <DayEmployeeEventsModal
          ymd={employeeModal.ymd}
          contact={employeeModal.contact}
          schedules={employeeModal.schedules}
          isAdmin={isAdmin}
          userProfile={userProfile}
          onUpdateSchedule={onUpdateSchedule}
          onDeleteSchedule={onDeleteSchedule}
          onClose={() => setEmployeeModal(null)}
        />
      )}

      {attendanceAdd && (
        <ContactAttendanceAddModal
          ymd={attendanceAdd.ymd}
          contact={attendanceAdd.contact}
          onBack={attendanceAdd.returnTo
            ? () => {
              const back = attendanceAdd.returnTo;
              setAttendanceAdd(null);
              if (back.type === "dayAll") setDayModalYmd(back.ymd);
            }
            : null}
          onCloseAll={() => setAttendanceAdd(null)}
          onSave={async (nextSchedule) => {
            await onAddSchedule(attendanceAdd.contact, nextSchedule);
            const back = attendanceAdd.returnTo;
            setAttendanceAdd(null);
            if (back?.type === "dayAll") setDayModalYmd(back.ymd);
          }}
        />
      )}

      <div style={{
        background: "#fff",
        border: "1px solid #e5edf5",
        borderRadius: isDesktop ? 22 : 14,
        padding: isDesktop ? "16px" : "12px",
        boxShadow: isDesktop ? "0 14px 36px rgba(15,23,42,0.05)" : "none",
      }}>
        <div style={{ display: "flex", gap: 8, alignItems: "center", justifyContent: "flex-end", flexWrap: "wrap", marginBottom: 12 }}>
          <button onClick={() => setMonthCursor((d) => addMonths(d, -1))} style={{ padding: "8px 12px", borderRadius: 10, border: "1px solid #cbd5e1", background: "#fff", cursor: "pointer", fontFamily: FF, fontWeight: 800 }}>
            {`\uC774\uC804`}
          </button>
          <div style={{ fontWeight: 900, color: "#0f172a", minWidth: 92, textAlign: "center" }}>{monthLabel}</div>
          <button onClick={() => setMonthCursor((d) => addMonths(d, 1))} style={{ padding: "8px 12px", borderRadius: 10, border: "1px solid #cbd5e1", background: "#fff", cursor: "pointer", fontFamily: FF, fontWeight: 800 }}>
            {`\uB2E4\uC74C`}
          </button>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "repeat(7, minmax(0, 1fr))", gap: isDesktop ? 8 : 4, marginBottom: 12 }}>
          {["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"].map((d) => (
            <div key={d} style={{ fontSize: 12, fontWeight: 900, color: "#64748b", textAlign: "center", minWidth: 0 }}>{d}</div>
          ))}
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "repeat(7, minmax(0, 1fr))", gap: isDesktop ? 8 : 4 }}>
          {calendarCells.cells.map((cell) => {
            const isToday = cell.ymd === todayYmd;
            const deptMap = deptDayIndex.get(cell.ymd) || new Map();
            const hasAny = deptMap.size > 0;
            return (
              <div
                key={cell.ymd}
                onClick={() => {
                  // Toggle: click same day again closes (like pressing "닫기")
                  setSelectedDept(null);
                  setSelectedYmd("");
                  setEmployeeModal(null);
                  setDayModalYmd((prev) => (prev === cell.ymd ? "" : cell.ymd));
                }}
                style={{
                  padding: isDesktop ? "10px 10px 8px" : "7px 6px 6px",
                  borderRadius: 14,
                  border: isToday ? "2px solid #2563eb" : "1px solid #e2e8f0",
                  background: isToday ? "#eff6ff" : (cell.inMonth ? "#fff" : "#f8fafc"),
                  minHeight: isDesktop ? 92 : 64,
                  minWidth: 0,
                  boxShadow: isToday ? "0 12px 26px rgba(37,99,235,0.20)" : "none",
                  cursor: "pointer",
                  overflow: "hidden",
                }}
              >
                <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 4 }}>
                  <span style={{ fontWeight: 900, fontSize: isDesktop ? 16 : 13, color: cell.inMonth ? "#0f172a" : "#94a3b8" }}>{cell.date.getDate()}</span>
                  {isToday && isDesktop && (
                    <span style={{
                      fontSize: 11,
                      fontWeight: 900,
                      color: "#fff",
                      background: "#2563eb",
                      padding: "2px 7px",
                      borderRadius: 999,
                      lineHeight: 1.2,
                      flexShrink: 0,
                    }}>{`\uC624\uB298`}</span>
                  )}
                </div>

                <div style={{ display: "flex", gap: 4, flexWrap: "wrap", marginTop: isDesktop ? 8 : 5, opacity: hasAny ? 1 : 0.35 }}>
                  {depts.map((dept) => {
                    if (!deptMap.has(dept)) return null;
                    const badge = getDeptBadge(dept);
                    return (
                      <button
                        key={dept}
                        onClick={(e) => {
                          e.stopPropagation();
                          setDayModalYmd("");
                          setEmployeeModal(null);
                          if (selectedDept === dept && selectedYmd === cell.ymd) {
                            setSelectedDept(null);
                            setSelectedYmd("");
                          } else {
                            setSelectedDept(dept);
                            setSelectedYmd(cell.ymd);
                          }
                        }}
                        style={{
                          border: `1px solid ${badge.bd}`,
                          background: badge.bg,
                          color: badge.fg,
                          borderRadius: 10,
                          padding: "3px 6px",
                          fontWeight: 900,
                          fontSize: isDesktop ? 12 : 11,
                          cursor: "pointer",
                          fontFamily: FF,
                          lineHeight: 1.1,
                          whiteSpace: "nowrap",
                          maxWidth: "100%",
                        }}
                        title={dept}
                      >
                        {badge.label}
                      </button>
                    );
                  })}
                </div>
              </div>
            );
          })}
        </div>
      </div>
    </>
  );
}

function ContactScheduleCalendar({ depts, grouped, isDesktop, todayYmd, userProfile }) {
  const [dept, setDept] = useState(() => depts[0] || "");
  const [contactId, setContactId] = useState("");
  const [category, setCategory] = useState("all"); // all | attendance | leave
  const [monthCursor, setMonthCursor] = useState(() => {
    const d = new Date();
    d.setHours(0, 0, 0, 0);
    d.setDate(1);
    return d;
  });
  const [selectedYmd, setSelectedYmd] = useState("");

  useEffect(() => {
    if (!dept && depts[0]) setDept(depts[0]);
  }, [dept, depts]);

  useEffect(() => {
    const targetName = String(userProfile?.name || "").trim();
    const targetDept = String(userProfile?.department || "").trim();
    if (!targetName || !depts.length) return;

    // Prefer exact dept+name match if possible.
    if (targetDept && grouped?.[targetDept]) {
      const candidates = (grouped[targetDept] || []).filter((c) => String(c.name || "").trim() === targetName);
      if (candidates.length === 1) {
        setDept(targetDept);
        setContactId(candidates[0].id);
        return;
      }
    }

    // Fallback: unique name match across all depts.
    const matches = [];
    for (const d of depts) {
      for (const c of grouped?.[d] || []) {
        if (String(c.name || "").trim() === targetName) matches.push({ dept: d, contact: c });
      }
    }
    if (matches.length === 1) {
      setDept(matches[0].dept);
      setContactId(matches[0].contact.id);
    }
  }, [depts, grouped, userProfile]);

  useEffect(() => {
    setContactId("");
    setSelectedYmd("");
  }, [dept]);

  const contactsInDept = useMemo(() => {
    const items = grouped?.[dept] || [];
    return [...items].sort((a, b) => String(a.name || "").localeCompare(String(b.name || ""), "ko"));
  }, [dept, grouped]);

  const selectedContact = useMemo(() => {
    if (!contactId) return null;
    return contactsInDept.find((c) => c.id === contactId) || null;
  }, [contactId, contactsInDept]);

  function ymdToDate(ymd) {
    if (!ymd) return null;
    const [y, m, d] = String(ymd).split("-").map((x) => Number(x));
    if (!y || !m || !d) return null;
    const dt = new Date(y, m - 1, d);
    dt.setHours(0, 0, 0, 0);
    return dt;
  }

  function dateToYmd(dt) {
    const y = dt.getFullYear();
    const m = String(dt.getMonth() + 1).padStart(2, "0");
    const d = String(dt.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }

  function addMonths(dt, delta) {
    const next = new Date(dt);
    next.setMonth(next.getMonth() + delta);
    next.setDate(1);
    next.setHours(0, 0, 0, 0);
    return next;
  }

  function getScheduleCategory(type) {
    const t = String(type || "");
    const leaveSet = new Set([
      "annual_leave",
      "half_day_am",
      "half_day_pm",
      "compensatory_leave",
      "maternity_leave",
      "parental_leave",
      "leave_of_absence",
    ]);
    if (leaveSet.has(t)) return "leave";
    return "attendance";
  }

  const monthLabel = useMemo(() => {
    const y = monthCursor.getFullYear();
    const m = monthCursor.getMonth() + 1;
    return `${y}-${String(m).padStart(2, "0")}`;
  }, [monthCursor]);

  const calendarCells = useMemo(() => {
    const first = new Date(monthCursor);
    first.setDate(1);
    const last = new Date(monthCursor);
    last.setMonth(last.getMonth() + 1);
    last.setDate(0);

    // Monday-start calendar
    const firstDow = (first.getDay() + 6) % 7; // 0..6 (Mon..Sun)
    const start = new Date(first);
    start.setDate(first.getDate() - firstDow);

    const cells = [];
    for (let i = 0; i < 42; i++) {
      const d = new Date(start);
      d.setDate(start.getDate() + i);
      const ymd = dateToYmd(d);
      cells.push({
        date: d,
        ymd,
        inMonth: d.getMonth() === monthCursor.getMonth(),
      });
    }
    return { first, last, cells };
  }, [monthCursor]);

  const schedules = useMemo(() => {
    const list = Array.isArray(selectedContact?.schedules) ? selectedContact.schedules : [];
    return list.filter((s) => s?.startDate);
  }, [selectedContact]);

  const eventsByDay = useMemo(() => {
    const map = new Map();
    if (!selectedContact) return map;

    const monthStartYmd = dateToYmd(calendarCells.first);
    const monthEndYmd = dateToYmd(calendarCells.last);

    for (const s of schedules) {
      const start = s.startDate;
      const end = s.endDate || s.startDate;
      const cat = getScheduleCategory(s.type);
      if (category !== "all" && cat !== category) continue;

      // quick reject: no overlap with visible month range
      if (end < monthStartYmd || start > monthEndYmd) continue;

      const cursor = ymdToDate(start);
      const endDt = ymdToDate(end);
      if (!cursor || !endDt) continue;

      for (let dt = new Date(cursor); dt <= endDt; dt.setDate(dt.getDate() + 1)) {
        const ymd = dateToYmd(dt);
        if (ymd < monthStartYmd || ymd > monthEndYmd) continue;
        if (!map.has(ymd)) map.set(ymd, []);
        map.get(ymd).push(s);
      }
    }

    // stable ordering
    for (const [k, v] of map.entries()) {
      v.sort((a, b) => String(a.type || "").localeCompare(String(b.type || "")));
      map.set(k, v);
    }
    return map;
  }, [calendarCells.first, calendarCells.last, category, schedules, selectedContact]);

  const selectedDayEvents = useMemo(() => {
    if (!selectedYmd) return [];
    return eventsByDay.get(selectedYmd) || [];
  }, [eventsByDay, selectedYmd]);

  return (
    <div style={{
      background: "#fff",
      border: "1px solid #e5edf5",
      borderRadius: isDesktop ? 22 : 14,
      padding: isDesktop ? "16px" : "12px",
      boxShadow: isDesktop ? "0 14px 36px rgba(15,23,42,0.05)" : "none",
    }}>
      <div style={{
        display: "flex",
        flexDirection: isDesktop ? "row" : "column",
        alignItems: isDesktop ? "center" : "stretch",
        gap: 10,
        marginBottom: 12,
      }}>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
          <select value={dept} onChange={(e) => setDept(e.target.value)} style={{ ...base.input, width: isDesktop ? 180 : "100%" }}>
            {depts.map((d) => <option key={d} value={d}>{d}</option>)}
          </select>
          <select value={contactId} onChange={(e) => setContactId(e.target.value)} style={{ ...base.input, width: isDesktop ? 180 : "100%" }}>
            <option value="">{`\uC9C1\uC6D0 \uC120\uD0DD`}</option>
            {contactsInDept.map((c) => (
              <option key={c.id} value={c.id}>{c.name}</option>
            ))}
          </select>
          <select value={category} onChange={(e) => setCategory(e.target.value)} style={{ ...base.input, width: isDesktop ? 160 : "100%" }}>
            <option value="all">{`\uC804\uCCB4`}</option>
            <option value="attendance">{`\uADFC\uD0DC`}</option>
            <option value="leave">{`\uD734\uAC00`}</option>
          </select>
        </div>

        <div style={{ flex: 1 }} />

        <div style={{ display: "flex", gap: 8, alignItems: "center", justifyContent: "flex-end", flexWrap: "wrap" }}>
          <button onClick={() => setMonthCursor((d) => addMonths(d, -1))} style={{ padding: "8px 12px", borderRadius: 10, border: "1px solid #cbd5e1", background: "#fff", cursor: "pointer", fontFamily: FF, fontWeight: 800 }}>
            {`\uC774\uC804`}
          </button>
          <div style={{ fontWeight: 900, color: "#0f172a", minWidth: 92, textAlign: "center" }}>{monthLabel}</div>
          <button onClick={() => setMonthCursor((d) => addMonths(d, 1))} style={{ padding: "8px 12px", borderRadius: 10, border: "1px solid #cbd5e1", background: "#fff", cursor: "pointer", fontFamily: FF, fontWeight: 800 }}>
            {`\uB2E4\uC74C`}
          </button>
        </div>
      </div>

      {!selectedContact && (
        <div style={{ padding: "22px 10px", color: "#64748b", fontWeight: 700 }}>
          {`\uBD80\uC11C\uC640 \uC9C1\uC6D0\uC744 \uC120\uD0DD\uD558\uBA74 \uB2EC\uB825\uC5D0 \uC77C\uC815\uC774 \uD45C\uC2DC\uB429\uB2C8\uB2E4.`}
        </div>
      )}

      {selectedContact && (
        <>
          <div style={{
            display: "grid",
            gridTemplateColumns: "repeat(7, minmax(0, 1fr))",
            gap: isDesktop ? 8 : 4,
            marginBottom: 12,
          }}>
            {["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"].map((d) => (
              <div key={d} style={{ fontSize: 12, fontWeight: 900, color: "#64748b", textAlign: "center", minWidth: 0 }}>{d}</div>
            ))}
          </div>

          <div style={{ display: "grid", gridTemplateColumns: "repeat(7, minmax(0, 1fr))", gap: isDesktop ? 8 : 4 }}>
            {calendarCells.cells.map((cell) => {
              const isToday = cell.ymd === todayYmd;
              const dayEvents = eventsByDay.get(cell.ymd) || [];
              const show = dayEvents.slice(0, 3);
              const more = Math.max(0, dayEvents.length - show.length);
              return (
                <button
                  key={cell.ymd}
                  onClick={() => setSelectedYmd((prev) => (prev === cell.ymd ? "" : cell.ymd))}
                  style={{
                    textAlign: "left",
                    padding: isDesktop ? "10px 10px 8px" : "7px 6px 6px",
                    borderRadius: 14,
                    border: isToday ? "2px solid #2563eb" : (selectedYmd === cell.ymd ? "2px solid #2563eb" : "1px solid #e2e8f0"),
                    background: isToday ? "#eff6ff" : (cell.inMonth ? "#fff" : "#f8fafc"),
                    cursor: "pointer",
                    fontFamily: FF,
                    minHeight: isDesktop ? 92 : 64,
                    minWidth: 0,
                    boxShadow: isToday ? "0 12px 26px rgba(37,99,235,0.20)" : "none",
                    overflow: "hidden",
                  }}
                >
                  <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 4 }}>
                    <span style={{ fontWeight: 900, fontSize: isDesktop ? 16 : 13, color: cell.inMonth ? "#0f172a" : "#94a3b8" }}>{cell.date.getDate()}</span>
                    {isToday && isDesktop && (
                      <span style={{
                        fontSize: 11,
                        fontWeight: 900,
                        color: "#fff",
                        background: "#2563eb",
                        padding: "2px 7px",
                        borderRadius: 999,
                        lineHeight: 1.2,
                        flexShrink: 0,
                      }}>{`\uC624\uB298`}</span>
                    )}
                  </div>
                  <div style={{ display: "flex", gap: 4, flexWrap: "wrap", marginTop: isDesktop ? 8 : 5 }}>
                    {show.map((s) => {
                      const meta = SCHEDULE_META[s.type] || { icon: "📌" };
                      return (
                        <span key={s.id} style={{ fontSize: 14, lineHeight: 1, flexShrink: 0 }} title={meta.label || s.type}>
                          {meta.icon || "📌"}
                        </span>
                      );
                    })}
                    {more > 0 && (
                      <span style={{ fontSize: 11, fontWeight: 900, color: "#64748b" }}>{`+${more}`}</span>
                    )}
                  </div>
                </button>
              );
            })}
          </div>

          <div style={{ marginTop: 14, padding: "12px 14px", borderRadius: 14, border: "1px solid #e2e8f0", background: "#f8fafc" }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}>
              <div style={{ fontWeight: 900, color: "#0f172a" }}>{selectedYmd || monthLabel}</div>
              <div style={{ color: "#94a3b8", fontWeight: 800, fontSize: 12 }}>{selectedContact.name}</div>
            </div>
            {selectedYmd && selectedDayEvents.length === 0 && (
              <div style={{ color: "#64748b", fontWeight: 700, fontSize: 13 }}>{`\uC77C\uC815 \uC5C6\uC74C`}</div>
            )}
            {selectedYmd && selectedDayEvents.length > 0 && (
              <div style={{ display: "grid", gap: 8 }}>
                {selectedDayEvents.map((s) => {
                  const meta = SCHEDULE_META[s.type] || { label: s.type, icon: "📌" };
                  return (
                    <div key={s.id} style={{ padding: "10px 12px", borderRadius: 12, background: "#fff", border: "1px solid #e2e8f0" }}>
                      <div style={{ display: "flex", alignItems: "baseline", gap: 10, minWidth: 0 }}>
                        <span style={{ fontWeight: 900, flexShrink: 0 }}>{meta.icon || "📌"}</span>
                        <span style={{ fontWeight: 900, color: "#0f172a", flexShrink: 0 }}>{meta.label || s.type}</span>
                        <span style={{ color: "#64748b", fontWeight: 800, fontSize: 12, minWidth: 0, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                          {formatSchedulePeriod(s)}
                        </span>
                      </div>
                      {!!String(s.detail || "").trim() && (
                        <div style={{ marginTop: 4, fontSize: 12, color: "#475569", fontWeight: 700, whiteSpace: "pre-wrap", overflowWrap: "anywhere" }}>
                          {s.detail}
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        </>
      )}
    </div>
  );
}

function ContactScheduleBatchModal({ contacts, onSave, onClose }) {
  const [text, setText] = useState("");
  const [saving, setSaving] = useState(false);
  const excelRef = useRef(null);

  function handleExcelTemplateDownload() {
    // Template that matches the supported B~I paste format:
    // 시작일 / 종료일 / 신청시간 / 신청일수 / 이름 / 직위 / 소속 / 휴가종류 / 내용(선택)
    const rows = [
      ["시작일", "종료일", "신청시간", "신청일수", "이름", "직위", "소속", "휴가종류", "내용(선택)"],
      ["2026-04-21", "2026-04-21", "오전", "0.5", "여희정", "차장", "관리부", "연차", ""],
      ["2026-04-22", "2026-04-22", "오전", "", "오성철", "부장", "시스템사업부", "외근", "상가 방문"],
      ["2026-04-22", "2026-04-22", "", "", "오성철", "부장", "시스템사업부", "출장", "[강한솔루션 한라산소주] MES 설비 미팅"],
    ];
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws["!cols"] = [12, 12, 10, 10, 14, 12, 14, 14, 36].map((wch) => ({ wch }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Template");
    XLSX.writeFile(wb, "일정_업로드_양식.xlsx");
  }

  function cellsFromExcelRow(row) {
    const r = Array.isArray(row) ? row : [];

    // Excel에서 "B~I" 또는 "B~J" 범위를 그대로 복사/업로드하면 첫 칸이 시작일(날짜)입니다.
    // 반대로 전체 시트를 업로드하면 A열(문서번호 등) 뒤에 B열이 시작일이 됩니다.
    const firstIsDate = !!normalizeDateYmd(r[0]);
    const secondIsDate = !!normalizeDateYmd(r[1]);

    // Range-only: B~J(9칸) / B~I(8칸)
    if (firstIsDate) {
      if (r.length >= 9) return r.slice(0, 9);
      if (r.length >= 8) return r.slice(0, 8);
      return [];
    }

    // Full sheet: A(문서번호 등) + B~J / B~I
    if (secondIsDate) {
      if (r.length >= 10) return r.slice(1, 10); // B~J
      if (r.length >= 9) return r.slice(1, 9); // B~I
      return [];
    }

    return [];
  }

  function appendLines(lines) {
    const next = lines.filter(Boolean).join("\n").trim();
    if (!next) return;
    setText((prev) => (prev ? `${prev}\n${next}` : next));
  }

  async function handleExcelUpload(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    e.target.value = "";

    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
        const lines = aoa
          .map((row) => cellsFromExcelRow(row))
          .filter((cells) => cells.length > 0 && cells.some((c) => String(c || "").trim()))
          .map((cells) =>
            cells
              .map((c) => String(c ?? "").replace(/\r?\n/g, " / ").trim())
              .join("\t")
          );
        appendLines(lines);
      } catch (err) {
        console.error(err);
        alert("\uC5D1\uC140 \uD30C\uC77C\uC744 \uC77D\uC9C0 \uBABB\uD588\uC2B5\uB2C8\uB2E4.");
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function handleExcelDownload() {
    const rows = [
      ["\uBD80\uC11C", "\uC774\uB984", "\uC9C1\uC704", "\uAD6C\uBD84", "\uC77C\uC815\uC885\uB958", "\uC2DC\uC791\uC77C", "\uC885\uB8CC\uC77C", "\uB0B4\uC6A9"],
    ];

    for (const c of contacts || []) {
      const list = Array.isArray(c.schedules) ? c.schedules : [];
      for (const s of list) {
        const meta = SCHEDULE_META[s.type] || { label: s.type };
        const category = getScheduleCategoryByType(s.type) === "leave" ? "\uD734\uAC00" : "\uADFC\uD0DC";
        rows.push([
          c.dept || "",
          c.name || "",
          c.position || "",
          category,
          meta.label || s.type || "",
          s.startDate || "",
          s.endDate || s.startDate || "",
          s.detail || "",
        ]);
      }
    }

    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws["!cols"] = [18, 14, 12, 10, 14, 12, 12, 26].map((wch) => ({ wch }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Schedules");
    XLSX.writeFile(wb, "schedules.xlsx");
  }

  async function handleSave() {
    const rows = text
      .split(/\r?\n/)
      .map((line) => tokenizeScheduleLine(line))
      .filter((cells) => cells.length > 0);

    if (!rows.length) {
      alert("붙여넣을 일정 데이터가 없습니다.");
      return;
    }

    const isHeaderRow = rows[0].some((cell) => /이름|부서|일정|시작|종료|메모|내용/.test(cell));
    const fallbackHeaderRow = rows[0].some((cell) => /이름|부서|소속|일정|휴가|시작|종료|메모|내용|신청시간|신청일수/.test(cell));
    const dataRows = (isHeaderRow || fallbackHeaderRow) ? rows.slice(1) : rows;
    const errors = [];
    const grouped = new Map();
    const pendingCreates = new Map();

    for (const [index, cells] of dataRows.entries()) {
      if (!cells.some(Boolean)) continue;

      const { name, dept, position, type, startDate, endDate, detail } = parseSchedulePasteRow(cells);

      if (!name || !type || !startDate) {
        errors.push(`${index + 1}행: 이름, 일정종류, 시작일은 필수입니다.`);
        continue;
      }

      let matched = contacts.filter((contact) => {
        const sameName = String(contact.name || "").trim() === name;
        const sameDept = !dept || String(contact.dept || "").trim() === dept;
        return sameName && sameDept;
      });
      if (matched.length === 0 && (name || dept || position)) {
        const createKey = `${name}__${dept}`;
        if (pendingCreates.has(createKey)) {
          matched = [{ id: createKey, ...pendingCreates.get(createKey) }];
        } else {
          const draftContact = {
            name,
            dept,
            position: position || "",
            phone: "",
            ext: "",
            direct: "",
            schedules: [],
          };
          pendingCreates.set(createKey, draftContact);
          matched = [{ id: createKey, ...draftContact }];
        }
      }

      if (matched.length === 0) {
        errors.push(`${index + 1}행: "${name}" 연락처를 찾지 못했습니다.`);
        continue;
      }

      if (matched.length > 1) {
        errors.push(`${index + 1}행: "${name}" 이름이 중복됩니다. 부서 열도 같이 넣어주세요.`);
        continue;
      }

      const contact = matched[0];
      const nextList = grouped.get(contact.id) || [...(Array.isArray(contact.schedules) ? contact.schedules : [])];
      nextList.push({
        id: `${Date.now()}_${index}_${Math.random().toString(36).slice(2, 8)}`,
        type,
        startDate,
        endDate,
        detail: String(detail || "").trim(),
        createdAt: new Date().toISOString(),
      });
      const dedupedList = dedupeSchedules(nextList);
      if (pendingCreates.has(contact.id)) {
        pendingCreates.get(contact.id).schedules = dedupedList;
      } else {
        grouped.set(contact.id, dedupedList);
      }
    }

    if (errors.length) {
      alert(errors.slice(0, 8).join("\n"));
      return;
    }

    setSaving(true);
    try {
      if (pendingCreates.size > 0) {
        await Promise.all([...pendingCreates.values()].map((item) => addContact(item)));
      }
      await onSave(grouped);
      onClose();
    } finally {
      setSaving(false);
    }
  }

  return (
    <Modal title="일정 일괄등록" onClose={onClose}>
      <input
        ref={excelRef}
        type="file"
        accept=".xlsx,.xls"
        style={{ display: "none" }}
        onChange={handleExcelUpload}
      />

      <div style={{ display: "flex", gap: 8, flexWrap: "wrap", justifyContent: "flex-end", marginBottom: 12 }}>
        <button onClick={handleExcelTemplateDownload} style={{
          padding: "9px 12px",
          borderRadius: 10,
          border: "1.5px solid #e2e8f0",
          background: "#fff",
          color: "#475569",
          fontWeight: 800,
          fontSize: 12,
          cursor: "pointer",
          fontFamily: FF,
          whiteSpace: "nowrap",
        }}>📋 업로드 양식</button>
        <button onClick={() => excelRef.current?.click()} style={{
          padding: "9px 12px",
          borderRadius: 10,
          border: "1.5px solid #86efac",
          background: "#f0fdf4",
          color: "#16a34a",
          fontWeight: 800,
          fontSize: 12,
          cursor: "pointer",
          fontFamily: FF,
          whiteSpace: "nowrap",
        }}>📥 엑셀 업로드</button>
        <button onClick={handleExcelDownload} style={{
          padding: "9px 12px",
          borderRadius: 10,
          border: "1.5px solid #bfdbfe",
          background: "#eff6ff",
          color: "#2563eb",
          fontWeight: 800,
          fontSize: 12,
          cursor: "pointer",
          fontFamily: FF,
          whiteSpace: "nowrap",
        }}>⬇ 일정 다운로드</button>
      </div>

      <div style={{
        padding: 12,
        marginBottom: 14,
        background: "#f8fafc",
        border: "1px solid #e2e8f0",
        borderRadius: 12,
        fontSize: 12,
        color: "#475569",
        lineHeight: 1.65,
      }}>
        엑셀에서 여러 셀을 그대로 복사해서 붙여넣으면 됩니다.<br />
        형식 1: 이름 / 일정종류 / 시작일 / 종료일 / 내용(선택)<br />
        형식 2: 이름 / 부서 / 일정종류 / 시작일 / 종료일 / 내용(선택)<br />
        예시: `여희정	출산휴가	2026-01-01	2026-05-31`<br />
        예시: `오성철	시스템사업부	외근	2026-04-21	2026-04-21	오전: 상가 / 오후: 남경테크윈`
      </div>
      <div style={{
        padding: 12,
        marginBottom: 14,
        background: "#eff6ff",
        border: "1px solid #bfdbfe",
        borderRadius: 12,
        fontSize: 12,
        color: "#1d4ed8",
        lineHeight: 1.65,
      }}>
        B~I 또는 B~J 열 그대로 붙여넣기/업로드도 지원합니다.<br />
        B~I(8칸): 시작일 / 종료일 / 신청시간 / 신청일수 / 이름 / 직위 / 소속 / 휴가종류<br />
        B~J(9칸): 위 + 내용(선택) (업무/외근/출장은 내용 입력 권장)<br />
        `휴가종류=연차`인 경우 `신청시간`으로 자동 변환됩니다.
        `1일`, `1일+a`는 연차 / `오전`은 오전 반차 / `오후`는 오후 반차로 등록됩니다.
        <br />
        근태(업무/외근/출장/연장/재택)는 `휴가종류` 칸에 그대로 입력하면 근태로 등록됩니다.
        근태의 `내용`이 비어 있으면 `신청시간`(오전/오후)을 내용으로 자동 저장합니다.<br />
        외근을 오전+오후 둘 다 쓰면 같은 날짜로 2줄로 올리거나, 내용에 `오전: ... / 오후: ...` 형태로 입력하세요. (셀 줄바꿈은 붙여넣기에서 행이 깨질 수 있어 `/` 등 구분자 추천)
      </div>

      <textarea
        value={text}
        onChange={(e) => setText(e.target.value)}
        placeholder={"이름\t일정종류\t시작일\t종료일\t내용(선택)"}
        style={{
          ...base.input,
          minHeight: 220,
          resize: "vertical",
          lineHeight: 1.5,
          marginBottom: 14,
        }}
      />

      <button
        onClick={handleSave}
        disabled={saving}
        style={{
          width: "100%",
          padding: "12px",
          borderRadius: 12,
          border: "none",
          background: saving ? "#94a3b8" : "#2563eb",
          color: "#fff",
          fontWeight: 800,
          fontSize: 14,
          cursor: saving ? "not-allowed" : "pointer",
          fontFamily: FF,
        }}
      >
        {saving ? "등록 중.." : "일정 등록"}
      </button>
    </Modal>
  );
}

function ScheduledContactCard({ contact, isAdmin, isDesktop, onEdit, onDelete, todayYmd }) {
  const [showPhone, setShowPhone] = useState(false);
  const [countdown, setCountdown] = useState(0);
  const [historyOpen, setHistoryOpen] = useState(false);
  const timerRef = useRef(null);
  const hasExt = !!String(contact.ext || "").trim();
  const hasDirect = !!String(contact.direct || "").trim();
  const activeSchedules = getActiveSchedules(contact, todayYmd);
  const historySchedules = getScheduleHistory(contact);

  useEffect(() => () => clearInterval(timerRef.current), []);

  function revealPhone() {
    if (showPhone) {
      clearInterval(timerRef.current);
      setShowPhone(false);
      setCountdown(0);
      return;
    }

    setShowPhone(true);
    setCountdown(10);
    clearInterval(timerRef.current);
    timerRef.current = setInterval(() => {
      setCountdown((prev) => {
        if (prev <= 1) {
          clearInterval(timerRef.current);
          setShowPhone(false);
          return 0;
        }
        return prev - 1;
      });
    }, 1000);
  }

  const visiblePhone = showPhone
    ? formatPhone(contact.phone) || "-"
    : (contact.phone ? "•••-••••-••••" : "-");

  return (
    <div style={{
      background: "#fff",
      borderRadius: isDesktop ? 18 : 14,
      padding: isDesktop ? "18px 20px" : "14px 14px",
      border: "1px solid #e8edf3",
      boxShadow: isDesktop ? "0 8px 24px rgba(15,23,42,0.06)" : "0 2px 8px rgba(15,23,42,0.06)",
      position: "relative",
    }}>
      {isAdmin && (
        <div style={{ position: "absolute", top: 12, right: 12, display: "flex", gap: 6 }}>
          <button onClick={() => onEdit(contact)} style={{
            padding: "5px 10px", borderRadius: 8, border: "1.5px solid #bfdbfe",
            background: "#eff6ff", color: "#2563eb", fontWeight: 700, fontSize: 11,
            cursor: "pointer", fontFamily: FF,
          }}>수정</button>
          <button onClick={() => onDelete(contact)} style={{
            padding: "5px 10px", borderRadius: 8, border: "1.5px solid #fca5a5",
            background: "#fef2f2", color: "#dc2626", fontWeight: 700, fontSize: 11,
            cursor: "pointer", fontFamily: FF,
          }}>삭제</button>
        </div>
      )}

      <div style={{ display: "grid", gap: 10, paddingRight: isAdmin ? 96 : 0 }}>
        <div style={{ display: "flex", alignItems: "baseline", gap: 8, flexWrap: "wrap" }}>
          <div style={{ fontWeight: 900, fontSize: isDesktop ? 20 : 17, color: "#0f172a", lineHeight: 1.1 }}>
            {contact.name || "-"}
          </div>
          {contact.dept && (
            <span style={{ fontSize: 14, fontWeight: 700, color: "#0891b2" }}>
              {contact.dept}
            </span>
          )}
          {contact.position && (
            <span style={{ fontSize: 14, fontWeight: 700, color: "#64748b" }}>
              {contact.position}
            </span>
          )}
          {activeSchedules.map((schedule) => {
            const meta = SCHEDULE_META[schedule.type] || { label: schedule.type, icon: "📌" };
            return (
              <span
                key={schedule.id}
                style={{
                  fontSize: 13,
                  fontWeight: 800,
                  color: meta.icon === "👶" ? "#be185d" : "#2563eb",
                }}
              >
                {meta.icon} {meta.label} {formatSchedulePeriod(schedule)}
              </span>
            );
          })}
        </div>

        <div style={{
          display: "flex",
          flexDirection: isDesktop ? "row" : "column",
          alignItems: isDesktop ? "center" : "flex-start",
          gap: isDesktop ? 14 : 8,
          fontSize: 14,
          fontWeight: 700,
          color: "#334155",
        }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
            <span style={{ color: "#64748b" }}>휴대전화</span>
            {contact.phone && (
              <button onClick={revealPhone} style={{
                padding: "4px 12px",
                borderRadius: 10,
                border: showPhone ? "1.5px solid #475569" : "1.5px solid #2563eb",
                background: "#fff",
                color: showPhone ? "#475569" : "#2563eb",
                fontWeight: 800,
                fontSize: 12,
                cursor: "pointer",
                fontFamily: FF,
                whiteSpace: "nowrap",
              }}>
                {showPhone ? `숨기기 ${countdown}s` : "보기"}
              </button>
            )}
            <span style={{ color: "#0f172a", fontWeight: 800 }}>{visiblePhone}</span>
          </div>

          {hasExt && (
            <>
              <span style={{ color: "#cbd5e1", display: isDesktop ? "inline" : "none" }}>|</span>
              <span><span style={{ color: "#64748b" }}>내선번호</span>{` ${contact.ext}`}</span>
            </>
          )}

          {hasDirect && (
            <>
              <span style={{ color: "#cbd5e1", display: isDesktop ? "inline" : "none" }}>|</span>
              <span>
                <span style={{ color: "#64748b" }}>직통번호</span>
                {" "}
                <a href={`tel:${compactPhone(contact.direct)}`} style={{ color: "#0f172a", textDecoration: "none", fontWeight: 800 }}>
                  {formatPhone(contact.direct)}
                </a>
              </span>
            </>
          )}
        </div>
      </div>
    </div>
  );
}

function ScheduledContactsView({ contacts, isAdmin, showToast, userProfile, user }) {
  const [search, setSearch] = useState("");
  const [expandedDepts, setExpandedDepts] = useState({});
  const [modalContact, setModalContact] = useState(undefined);
  const [showScheduleBatch, setShowScheduleBatch] = useState(false);
  const [viewMode, setViewMode] = useState("calendar"); // list | calendar
  const isDesktop = useMediaQuery("(min-width: 900px)");
  const todayYmd = getTodayYmd();

  const filtered = useMemo(() => {
    const q = search.trim().toLowerCase();
    if (!q) return contacts;
    return contacts.filter((c) =>
      [
        c.name,
        c.dept,
        c.position,
        c.ext,
        c.direct,
        ...(Array.isArray(c.schedules) ? c.schedules.map((schedule) => {
          const meta = SCHEDULE_META[schedule.type];
          return `${meta?.label || ""} ${schedule.startDate || ""} ${schedule.endDate || ""}`;
        }) : []),
      ].join(" ").toLowerCase().includes(q)
    );
  }, [contacts, search]);

  const grouped = useMemo(() => {
    const g = {};
    filtered.forEach((c) => {
      const dept = c.dept || "기타";
      if (!g[dept]) g[dept] = [];
      g[dept].push(c);
    });
    return g;
  }, [filtered]);

  const depts = Object.keys(grouped).sort((a, b) => a.localeCompare(b, "ko"));

  function expandAll() {
    setExpandedDepts(Object.fromEntries(depts.map((dept) => [dept, true])));
  }

  function collapseAll() {
    setExpandedDepts(Object.fromEntries(depts.map((dept) => [dept, false])));
  }

  async function handleSave(form) {
    try {
      if (modalContact?.id) {
        await updateContact(modalContact.id, form);
        showToast("수정되었습니다.");
      } else {
        await addContact(form);
        showToast("추가되었습니다.");
      }
    } catch (err) {
      console.error("연락처 저장 오류:", err);
      showToast(`저장 오류: ${err?.message || err}`, "error");
      throw err;
    }
  }

  async function handleDelete(contact) {
    if (!window.confirm(`"${contact.name}"을(를) 삭제할까요?`)) return;
    try {
      await deleteContact(contact.id);
      showToast("삭제되었습니다.");
    } catch {
      showToast("삭제 중 오류가 발생했습니다.", "error");
    }
  }

  async function handleSaveSchedules(groupedSchedules) {
    const updates = [...groupedSchedules.entries()].map(([contactId, schedules]) => {
      const existing = contacts.find((contact) => contact.id === contactId);
      if (!existing) return Promise.resolve();
      return updateContact(contactId, { ...existing, schedules: dedupeSchedules(schedules) });
    });

    await Promise.all(updates);
    showToast("일정을 등록했습니다.");
  }

  async function handleAddSchedule(contact, schedule) {
    if (!contact?.id) return;

    // Security-in-depth: UI에서 막더라도, 비관리자는 "본인"만 입력 가능
    if (!isAdmin) {
      const targetName = String(userProfile?.name || "").trim();
      if (!targetName) { showToast("본인 일정 입력을 위해 사용자 이름이 필요합니다.", "error"); return; }

      const isSameName = isOwnContactForUser(contact, userProfile);
      if (!isSameName) { showToast("본인 일정만 입력할 수 있습니다.", "error"); return; }
    }

    const nextSchedules = dedupeSchedules([
      ...(Array.isArray(contact.schedules) ? contact.schedules : []),
      schedule,
    ]);
    await updateContact(contact.id, { ...contact, schedules: nextSchedules });
    showToast("일정을 추가했습니다.");
  }

  async function handleUpdateSchedule(contact, scheduleId, updates) {
    if (!contact?.id || !scheduleId) return;
    const currentSchedules = Array.isArray(contact.schedules) ? contact.schedules : [];
    const currentSchedule = currentSchedules.find((schedule) => schedule.id === scheduleId);
    if (!currentSchedule) return;

    if (!canEditSchedule({ isAdmin, contact, schedule: currentSchedule, userProfile })) {
      showToast("수정할 수 없는 일정입니다.", "error");
      return;
    }

    const nextSchedule = { ...currentSchedule, ...updates };
    if (!isAdmin && getScheduleCategoryByType(nextSchedule.type) !== "attendance") {
      showToast("근태 일정만 수정할 수 있습니다.", "error");
      return;
    }
    const nextSchedules = dedupeSchedules(
      (Array.isArray(contact.schedules) ? contact.schedules : []).map((schedule) =>
        schedule.id === scheduleId ? nextSchedule : schedule
      )
    );
    await updateContact(contact.id, { ...contact, schedules: nextSchedules });
    try {
      const prevMeta = SCHEDULE_META[currentSchedule.type];
      const nextMeta = SCHEDULE_META[nextSchedule.type];
      await writeLog({
        action: "일정 수정",
        email: user?.email || "",
        targetId: contact.id,
        targetName: `${contact.name || ""} · ${(prevMeta?.label || currentSchedule.type || "").trim()} -> ${(nextMeta?.label || nextSchedule.type || "").trim()}`.trim(),
      });
    } catch { }
    showToast("일정을 수정했습니다.");
  }

  async function handleDeleteSchedule(contact, scheduleId) {
    if (!contact?.id || !scheduleId) return;
    const currentSchedules = Array.isArray(contact.schedules) ? contact.schedules : [];
    const currentSchedule = currentSchedules.find((schedule) => schedule.id === scheduleId);
    if (!currentSchedule) return;
    if (!isAdmin) {
      showToast("관리자만 일정을 삭제할 수 있습니다.", "error");
      return;
    }

    const nextSchedules = (Array.isArray(contact.schedules) ? contact.schedules : []).filter(
      (schedule) => schedule.id !== scheduleId
    );
    await updateContact(contact.id, { ...contact, schedules: nextSchedules });
    try {
      const meta = SCHEDULE_META[currentSchedule.type];
      await writeLog({
        action: "일정 삭제",
        email: user?.email || "",
        targetId: contact.id,
        targetName: `${contact.name || ""} · ${(meta?.label || currentSchedule.type || "").trim()}`.trim(),
      });
    } catch { }
    showToast("일정을 삭제했습니다.");
  }

  return (
    <div style={{ maxWidth: 1180, margin: "0 auto", padding: isDesktop ? "28px 24px 88px" : "20px 16px 80px" }}>
      {modalContact !== undefined && (
        <ContactModal
          contact={modalContact}
          onSave={handleSave}
          onClose={() => setModalContact(undefined)}
        />
      )}
      {showScheduleBatch && (
        <ContactScheduleBatchModal
          contacts={contacts}
          onSave={handleSaveSchedules}
          onClose={() => setShowScheduleBatch(false)}
        />
      )}

      {/* ── 상단 툴바: 연락처(검색/+추가) | 일정(일괄등록) ── */}
      <div style={{
        display: "flex",
        flexDirection: isDesktop ? "row" : "column",
        gap: 12,
        marginBottom: 22,
        alignItems: "stretch",
      }}>
        {/* 연락처 그룹 */}
        <div style={{
          flex: 1,
          width: "100%",
          background: "#fff",
          border: "1.5px solid #e2e8f0",
          borderRadius: 18,
          padding: "10px 12px",
          boxShadow: "0 10px 26px rgba(15,23,42,0.06)",
          display: "flex",
          alignItems: "center",
          gap: 10,
          minWidth: 0,
        }}>
          <button
            onClick={() => setViewMode("list")}
            style={{
              padding: "10px 14px",
              borderRadius: 14,
              border: viewMode === "list" ? "1.5px solid #2563eb" : "1.5px solid #cbd5e1",
              background: viewMode === "list" ? "#eff6ff" : "#fff",
              color: viewMode === "list" ? "#2563eb" : "#475569",
              fontWeight: 900,
              fontSize: 13,
              cursor: "pointer",
              fontFamily: FF,
              whiteSpace: "nowrap",
              flexShrink: 0,
            }}
          >
            {`\uC5F0\uB77D\uCC98`}
          </button>

          <span style={{ width: 1, height: 26, background: "#e2e8f0", flexShrink: 0 }} />

          <div style={{
            display: "flex",
            alignItems: "center",
            gap: 8,
            flex: 1,
            minWidth: 0,
          }}>
            <span style={{ color: "#cbd5e1", fontSize: 15, flexShrink: 0 }}>🔍</span>
            <input
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              placeholder="이름, 부서, 직위, 일정으로 검색"
              style={{ ...base.input, border: "none", padding: "8px 4px", boxShadow: "none", flex: 1, minWidth: 0 }}
            />
            {search && (
              <button onClick={() => setSearch("")} style={{
                background: "none", border: "none", cursor: "pointer",
                color: "#cbd5e1", fontSize: 14, padding: "0 4px",
                flexShrink: 0,
              }}>✕</button>
            )}
          </div>

          {isAdmin && (
            <button onClick={() => setModalContact(null)} style={{
              padding: "10px 14px",
              borderRadius: 14,
              border: "none",
              background: "#2563eb",
              color: "#fff",
              fontWeight: 900,
              fontSize: 13,
              cursor: "pointer",
              fontFamily: FF,
              whiteSpace: "nowrap",
              flexShrink: 0,
            }}>+ 추가</button>
          )}
        </div>

        {/* 일정 그룹 */}
        <div style={{
          flex: isDesktop ? "0 0 auto" : 0,
          width: isDesktop ? "auto" : "100%",
          background: "#fff",
          border: "1.5px solid #e2e8f0",
          borderRadius: 18,
          padding: "10px 12px",
          boxShadow: "0 10px 26px rgba(15,23,42,0.06)",
          display: "flex",
          alignItems: "center",
          gap: 10,
          justifyContent: "space-between",
        }}>
          <button
            onClick={() => setViewMode("calendar")}
            style={{
              padding: "10px 14px",
              borderRadius: 14,
              border: viewMode === "calendar" ? "1.5px solid #2563eb" : "1.5px solid #cbd5e1",
              background: viewMode === "calendar" ? "#eff6ff" : "#fff",
              color: viewMode === "calendar" ? "#2563eb" : "#475569",
              fontWeight: 900,
              fontSize: 13,
              cursor: "pointer",
              fontFamily: FF,
              whiteSpace: "nowrap",
            }}
          >
            {`\uC77C\uC815`}
          </button>

          {isAdmin && (
            <button onClick={() => setShowScheduleBatch(true)} style={{
              padding: "10px 14px",
              borderRadius: 14,
              border: "1.5px solid #bfdbfe",
              background: "#eff6ff",
              color: "#2563eb",
              fontWeight: 900,
              fontSize: 13,
              cursor: "pointer",
              fontFamily: FF,
              whiteSpace: "nowrap",
            }}>일정 일괄등록</button>
          )}
        </div>
      </div>

      <div style={{ fontSize: 13, color: "#64748b", fontWeight: 700, marginBottom: 18 }}>
        총 {contacts.length}명{search ? ` · 검색 결과 ${filtered.length}명` : ""}
      </div>

      {viewMode === "calendar" && (
        <ContactScheduleCalendarDept
          depts={depts}
          grouped={grouped}
          isDesktop={isDesktop}
          todayYmd={todayYmd}
          isAdmin={isAdmin}
          userProfile={userProfile}
          onAddSchedule={handleAddSchedule}
          onUpdateSchedule={handleUpdateSchedule}
          onDeleteSchedule={handleDeleteSchedule}
        />
      )}

      {viewMode === "list" && depts.length > 0 && (
        <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", marginBottom: 14 }}>
          <button onClick={expandAll} style={{
            padding: "8px 12px", borderRadius: 10, border: "1px solid #cbd5e1",
            background: "#fff", color: "#475569", fontWeight: 700, fontSize: 12,
            cursor: "pointer", fontFamily: FF,
          }}>전체 펼치기</button>
          <button onClick={collapseAll} style={{
            padding: "8px 12px", borderRadius: 10, border: "1px solid #cbd5e1",
            background: "#fff", color: "#475569", fontWeight: 700, fontSize: 12,
            cursor: "pointer", fontFamily: FF,
          }}>전체 접기</button>
        </div>
      )}

      {viewMode === "list" && depts.length === 0 && (
        <div style={{ textAlign: "center", padding: "60px 0", color: "#cbd5e1" }}>
          <div style={{ fontSize: 36, marginBottom: 12 }}>📇</div>
          <div style={{ fontSize: 15 }}>
            {search ? "검색 결과가 없습니다." : "등록된 연락처가 없습니다."}
          </div>
        </div>
      )}

      {viewMode === "list" && depts.map((dept) => {
        const isExpanded = expandedDepts[dept] !== false;
        const items = grouped[dept];
        return (
          <div key={dept} style={{
            marginBottom: 18,
            background: "#fff",
            border: "1px solid #e5edf5",
            borderRadius: isDesktop ? 22 : 14,
            padding: isDesktop ? "14px" : "10px",
            boxShadow: isDesktop ? "0 14px 36px rgba(15,23,42,0.05)" : "none",
          }}>
            <div
              onClick={() => setExpandedDepts((p) => ({ ...p, [dept]: !isExpanded }))}
              style={{
                display: "flex",
                alignItems: "center",
                gap: 10,
                padding: isDesktop ? "14px 16px" : "10px 12px",
                borderRadius: 14,
                background: "#f8fafc",
                cursor: "pointer",
                marginBottom: isExpanded ? 8 : 0,
                userSelect: "none",
              }}
            >
              <span style={{ fontWeight: 900, fontSize: isDesktop ? 18 : 14, color: "#0f172a" }}>{dept}</span>
              <span style={{ fontSize: 12, color: "#94a3b8", fontWeight: 700 }}>{items.length}명</span>
              <div style={{ flex: 1 }} />
              <span style={{
                fontSize: 12,
                color: "#94a3b8",
                transform: isExpanded ? "rotate(0deg)" : "rotate(-90deg)",
                transition: "transform 0.2s",
                display: "inline-block",
              }}>⌄</span>
            </div>

            {isExpanded && (
              <div style={{
                display: "grid",
                gridTemplateColumns: isDesktop ? "repeat(auto-fit, minmax(360px, 1fr))" : "1fr",
                gap: isDesktop ? 12 : 0,
                paddingTop: 4,
              }}>
                {items.map((contact) => (
                  <ScheduledContactCardV3
                    key={contact.id}
                    contact={contact}
                    isAdmin={isAdmin}
                    userProfile={userProfile}
                    isDesktop={isDesktop}
                    onEdit={(item) => setModalContact(item)}
                    onDelete={handleDelete}
                    onUpdateSchedule={handleUpdateSchedule}
                    onDeleteSchedule={handleDeleteSchedule}
                    todayYmd={todayYmd}
                  />
                ))}
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}

function ScheduledContactCardV2({ contact, isAdmin, isDesktop, onEdit, onDelete, todayYmd }) {
  const [showPhone, setShowPhone] = useState(false);
  const [countdown, setCountdown] = useState(0);
  const [historyOpen, setHistoryOpen] = useState(false);
  const timerRef = useRef(null);
  const phoneDigits = compactPhone(contact.phone || "");
  const directDigits = compactPhone(contact.direct || "");
  const hasExt = !!String(contact.ext || "").trim();
  const hasDirect = directDigits.length >= 8;
  const activeSchedules = getActiveSchedules(contact, todayYmd);
  const historySchedules = getScheduleHistory(contact);

  useEffect(() => () => clearInterval(timerRef.current), []);

  function revealPhone() {
    if (showPhone) {
      clearInterval(timerRef.current);
      setShowPhone(false);
      setCountdown(0);
      return;
    }

    setShowPhone(true);
    setCountdown(10);
    clearInterval(timerRef.current);
    timerRef.current = setInterval(() => {
      setCountdown((prev) => {
        if (prev <= 1) {
          clearInterval(timerRef.current);
          setShowPhone(false);
          return 0;
        }
        return prev - 1;
      });
    }, 1000);
  }

  const visiblePhone = showPhone
    ? formatPhone(contact.phone) || "-"
    : (phoneDigits ? "•••-••••-••••" : "-");

  return (
    <div style={{
      background: "#fff",
      borderRadius: isDesktop ? 18 : 14,
      padding: isDesktop ? "18px 20px" : "14px 14px",
      border: "1px solid #e8edf3",
      boxShadow: isDesktop ? "0 8px 24px rgba(15,23,42,0.06)" : "0 2px 8px rgba(15,23,42,0.06)",
      position: "relative",
    }}>
      {isAdmin && (
        <div style={{ position: "absolute", top: 12, right: 12, display: "flex", gap: 6 }}>
          <button onClick={() => onEdit(contact)} style={{
            padding: "5px 10px", borderRadius: 8, border: "1.5px solid #bfdbfe",
            background: "#eff6ff", color: "#2563eb", fontWeight: 700, fontSize: 11,
            cursor: "pointer", fontFamily: FF,
          }}>수정</button>
          <button onClick={() => onDelete(contact)} style={{
            padding: "5px 10px", borderRadius: 8, border: "1.5px solid #fca5a5",
            background: "#fef2f2", color: "#dc2626", fontWeight: 700, fontSize: 11,
            cursor: "pointer", fontFamily: FF,
          }}>삭제</button>
        </div>
      )}

      <div style={{ display: "grid", gap: 10, paddingRight: isAdmin ? 96 : 0 }}>
        <div style={{ display: "flex", alignItems: "baseline", gap: 8, flexWrap: "wrap" }}>
          <div style={{ fontWeight: 900, fontSize: isDesktop ? 20 : 17, color: "#0f172a", lineHeight: 1.1 }}>
            {contact.name || "-"}
          </div>
          {contact.dept && (
            <span style={{ fontSize: 14, fontWeight: 700, color: "#0891b2" }}>
              {contact.dept}
            </span>
          )}
          {contact.position && (
            <span style={{ fontSize: 14, fontWeight: 700, color: "#64748b" }}>
              {contact.position}
            </span>
          )}
          {activeSchedules.map((schedule) => {
            const meta = SCHEDULE_META[schedule.type] || { label: schedule.type, icon: "📌" };
            return (
              <span
                key={schedule.id}
                style={{
                  fontSize: 13,
                  fontWeight: 800,
                  color: meta.icon === "👶" ? "#be185d" : "#2563eb",
                }}
              >
                {meta.icon} {meta.label} {formatSchedulePeriod(schedule)}
              </span>
            );
          })}
        </div>

        <div style={{
          display: "flex",
          flexDirection: isDesktop ? "row" : "column",
          alignItems: isDesktop ? "center" : "flex-start",
          flexWrap: "nowrap",
          gap: isDesktop ? 12 : 8,
          fontSize: 14,
          fontWeight: 700,
          color: "#334155",
          overflowX: "auto",
        }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "nowrap", whiteSpace: "nowrap", flexShrink: 0 }}>
            <span style={{ color: "#64748b" }}>휴대전화</span>
            {phoneDigits && (
              <button onClick={revealPhone} style={{
                padding: "4px 12px",
                borderRadius: 10,
                border: showPhone ? "1.5px solid #475569" : "1.5px solid #2563eb",
                background: "#fff",
                color: showPhone ? "#475569" : "#2563eb",
                fontWeight: 800,
                fontSize: 12,
                cursor: "pointer",
                fontFamily: FF,
                whiteSpace: "nowrap",
              }}>
                {showPhone ? `숨기기 ${countdown}s` : "보기"}
              </button>
            )}
            <span style={{ color: "#0f172a", fontWeight: 800, whiteSpace: "nowrap" }}>{visiblePhone}</span>
          </div>

          {hasExt && (
            <>
              <span style={{ color: "#cbd5e1", display: isDesktop ? "inline" : "none", flexShrink: 0 }}>|</span>
              <span style={{ whiteSpace: "nowrap", flexShrink: 0 }}>
                <span style={{ color: "#64748b" }}>내선번호</span>{` ${contact.ext}`}
              </span>
            </>
          )}

          {hasDirect && (
            <>
              <span style={{ color: "#cbd5e1", display: isDesktop ? "inline" : "none", flexShrink: 0 }}>|</span>
              <span style={{ whiteSpace: "nowrap", flexShrink: 0 }}>
                <span style={{ color: "#64748b" }}>직통번호</span>
                {" "}
                <a href={`tel:${directDigits}`} style={{ color: "#0f172a", textDecoration: "none", fontWeight: 800 }}>
                  {formatPhone(contact.direct)}
                </a>
              </span>
            </>
          )}
        </div>

        {historySchedules.length > 0 && (
          <div style={{ display: "grid", gap: 10 }}>
            <button
              onClick={() => setHistoryOpen((prev) => !prev)}
              style={{
                justifySelf: "flex-start",
                padding: "6px 12px",
                borderRadius: 10,
                border: "1px solid #cbd5e1",
                background: "#fff",
                color: "#475569",
                fontWeight: 700,
                fontSize: 12,
                cursor: "pointer",
                fontFamily: FF,
              }}
            >
              {historyOpen ? "일정 히스토리 닫기" : `일정 히스토리 ${historySchedules.length}건`}
            </button>

            {historyOpen && (
              <div style={{
                display: "grid",
                gap: 8,
                padding: "12px 14px",
                borderRadius: 12,
                background: "#f8fafc",
                border: "1px solid #e2e8f0",
              }}>
                {historySchedules.map((schedule) => {
                  const meta = SCHEDULE_META[schedule.type] || { label: schedule.type, icon: "📌" };
                  const status = getScheduleStatus(schedule, todayYmd);
                  const statusLabel = status === "active" ? "현재" : status === "upcoming" ? "예정" : "종료";
                  const statusColor = status === "active" ? "#2563eb" : status === "upcoming" ? "#7c3aed" : "#94a3b8";

                  return (
                    <div
                      key={schedule.id}
                      style={{
                        display: "flex",
                        flexDirection: isDesktop ? "row" : "column",
                        flexWrap: "nowrap",
                        alignItems: isDesktop ? "center" : "flex-start",
                        gap: 8,
                        padding: "10px 12px",
                        borderRadius: 10,
                        background: "#fff",
                        border: "1px solid #e2e8f0",
                      }}
                    >
                      <span style={{ fontSize: 13, fontWeight: 800, color: "#0f172a", whiteSpace: "nowrap" }}>
                        {meta.icon} {meta.label}
                      </span>
                      <span style={{ fontSize: 12, fontWeight: 800, color: statusColor, whiteSpace: "nowrap" }}>
                        {statusLabel}
                      </span>
                      <span style={{ fontSize: 12, color: "#475569", fontWeight: 700, whiteSpace: "nowrap" }}>
                        {formatSchedulePeriod(schedule)}
                      </span>
                      {schedule.note && (
                        <span style={{ fontSize: 12, color: "#64748b" }}>
                          메모: {schedule.note}
                        </span>
                      )}
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}

function ContactScheduleEditModal({ schedule, onSave, onClose, typeOptions }) {
  const [form, setForm] = useState(() => ({
    type: schedule?.type || "annual_leave",
    startDate: schedule?.startDate || "",
    endDate: schedule?.endDate || schedule?.startDate || "",
    detail: schedule?.detail || "",
  }));

  function submit() {
    const next = cleanSchedulePayload({ ...schedule, ...form });
    if (!next.type || !next.startDate) {
      alert("일정 종류와 시작일을 입력해주세요.");
      return;
    }
    onSave(next);
  }

  return (
    <Modal title={"일정 수정"} onClose={onClose}>
      <div style={{ display: "grid", gap: 12 }}>
        <select
          value={form.type}
          onChange={(e) => setForm((prev) => ({ ...prev, type: e.target.value }))}
          style={base.input}
        >
          {(typeOptions || Object.entries(SCHEDULE_META)).map(([value, meta]) => (
            <option key={value} value={value}>{meta.label}</option>
          ))}
        </select>
        <input
          type="date"
          value={form.startDate}
          onChange={(e) => setForm((prev) => ({ ...prev, startDate: e.target.value, endDate: prev.endDate || e.target.value }))}
          style={base.input}
        />
        <input
          type="date"
          value={form.endDate}
          onChange={(e) => setForm((prev) => ({ ...prev, endDate: e.target.value }))}
          style={base.input}
        />
        <input
          value={form.detail}
          onChange={(e) => setForm((prev) => ({ ...prev, detail: e.target.value }))}
          placeholder={"내용 (업무/외근/출장 상세 등) - 선택"}
          style={base.input}
        />
        <div style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}>
          <button onClick={onClose} style={{ padding: "10px 14px", borderRadius: 10, border: "1px solid #cbd5e1", background: "#fff", color: "#475569", fontWeight: 700, cursor: "pointer", fontFamily: FF }}>{"취소"}</button>
          <button onClick={submit} style={{ padding: "10px 14px", borderRadius: 10, border: "none", background: "#2563eb", color: "#fff", fontWeight: 800, cursor: "pointer", fontFamily: FF }}>{"저장"}</button>
        </div>
      </div>
    </Modal>
  );
}

function ContactAttendanceAddModal({ ymd, contact, onSave, onBack, onCloseAll }) {
  const typeOptions = [
    { value: "work_task", label: "\uC5C5\uBB34" },
    { value: "field_work", label: "\uC678\uADFC" },
    { value: "business_trip", label: "\uCD9C\uC7A5" },
    { value: "overtime", label: "\uC5F0\uC7A5" },
    { value: "remote_work", label: "\uC7AC\uD0DD" },
  ];

  const [form, setForm] = useState(() => ({
    type: "work_task",
    startDate: ymd || "",
    endDate: ymd || "",
    detail: "",
  }));

  function submit() {
    const next = cleanSchedulePayload({ ...form });
    if (!next.type || !next.startDate) {
      alert("\uADFC\uD0DC \uC885\uB958\uC640 \uB0A0\uC9DC\uB97C \uC785\uB825\uD574\uC8FC\uC138\uC694.");
      return;
    }
    onSave(next);
  }

  return (
    <Modal
      title={`${contact?.name || ""} \uADFC\uD0DC \uC785\uB825`}
      onClose={onBack || onCloseAll}
      extra={<span style={{ fontSize: 12, color: "#64748b", fontWeight: 800 }}>{ymd || ""}</span>}
    >
      <div style={{ display: "grid", gap: 12 }}>
        <select
          value={form.type}
          onChange={(e) => setForm((p) => ({ ...p, type: e.target.value }))}
          style={base.input}
        >
          {typeOptions.map((o) => (
            <option key={o.value} value={o.value}>{o.label}</option>
          ))}
        </select>

        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
          <input
            type="date"
            value={form.startDate}
            onChange={(e) => setForm((p) => ({ ...p, startDate: e.target.value, endDate: p.endDate || e.target.value }))}
            style={base.input}
          />
          <input
            type="date"
            value={form.endDate}
            onChange={(e) => setForm((p) => ({ ...p, endDate: e.target.value }))}
            style={base.input}
          />
        </div>

        <input
          value={form.detail}
          onChange={(e) => setForm((p) => ({ ...p, detail: e.target.value }))}
          placeholder={"\uB0B4\uC6A9 (\uC608: \uC5C5\uBB34, \uC678\uADFC \uC7A5\uC18C \uB4F1)"}
          style={base.input}
        />

        <div style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}>
          {onBack && (
            <button onClick={onBack} style={{
              padding: "10px 14px",
              borderRadius: 10,
              border: "1px solid #cbd5e1",
              background: "#fff",
              color: "#475569",
              fontWeight: 800,
              cursor: "pointer",
              fontFamily: FF,
              whiteSpace: "nowrap",
            }}>
              {`\uC774\uC804\uC73C\uB85C`}
            </button>
          )}
          <button onClick={onCloseAll} style={{
            padding: "10px 14px",
            borderRadius: 10,
            border: "1px solid #cbd5e1",
            background: "#fff",
            color: "#475569",
            fontWeight: 800,
            cursor: "pointer",
            fontFamily: FF,
            whiteSpace: "nowrap",
          }}>
            {`\uB2EB\uAE30`}
          </button>
          <button onClick={submit} style={{ padding: "10px 14px", borderRadius: 10, border: "none", background: "#2563eb", color: "#fff", fontWeight: 800, cursor: "pointer", fontFamily: FF }}>
            {`\uC800\uC7A5`}
          </button>
        </div>
      </div>
    </Modal>
  );
}

function ScheduledContactCardV3({ contact, isAdmin, isDesktop, onEdit, onDelete, onUpdateSchedule, onDeleteSchedule, todayYmd, userProfile }) {
  const [showPhone, setShowPhone] = useState(false);
  const [countdown, setCountdown] = useState(0);
  const [historyOpen, setHistoryOpen] = useState(false);
  const initialHistoryCount = isDesktop ? 3 : 2;
  const [historyLimit, setHistoryLimit] = useState(initialHistoryCount);
  const [editingSchedule, setEditingSchedule] = useState(null);
  const timerRef = useRef(null);
  const phoneDigits = compactPhone(contact.phone || "");
  const directDigits = compactPhone(contact.direct || "");
  const hasExt = !!String(contact.ext || "").trim();
  const hasDirect = directDigits.length >= 8;
  const splitDeskLine = isDesktop && hasExt && hasDirect;
  const activeSchedules = getActiveSchedules(contact, todayYmd);
  const historySchedules = getScheduleHistory(contact);
  const editableScheduleTypes = useMemo(
    () => Object.entries(SCHEDULE_META).filter(([value]) => isAdmin || getScheduleCategoryByType(value) === "attendance"),
    [isAdmin]
  );

  useEffect(() => () => clearInterval(timerRef.current), []);

  function revealPhone() {
    if (showPhone) {
      clearInterval(timerRef.current);
      setShowPhone(false);
      setCountdown(0);
      return;
    }

    setShowPhone(true);
    setCountdown(10);
    clearInterval(timerRef.current);
    timerRef.current = setInterval(() => {
      setCountdown((prev) => {
        if (prev <= 1) {
          clearInterval(timerRef.current);
          setShowPhone(false);
          return 0;
        }
        return prev - 1;
      });
    }, 1000);
  }

  const visiblePhone = showPhone
    ? formatPhone(contact.phone) || "-"
    : (phoneDigits ? "·····" : "-");

  return (
    <>
      {editingSchedule && (
        <ContactScheduleEditModal
          schedule={editingSchedule}
          typeOptions={editableScheduleTypes}
          onSave={async (nextSchedule) => {
            await onUpdateSchedule(contact, editingSchedule.id, nextSchedule);
            setEditingSchedule(null);
          }}
          onClose={() => setEditingSchedule(null)}
        />
      )}
      <div style={{
        background: "#fff",
        borderRadius: isDesktop ? 18 : 14,
        padding: isDesktop ? "18px 20px" : "14px 14px",
        border: "1px solid #e8edf3",
        boxShadow: isDesktop ? "0 8px 24px rgba(15,23,42,0.06)" : "0 2px 8px rgba(15,23,42,0.06)",
        position: "relative",
      }}>
        {isAdmin && (
          <div style={{ position: "absolute", top: 12, right: 12, display: "flex", gap: 6 }}>
            <button onClick={() => onEdit(contact)} style={{
              padding: "5px 10px", borderRadius: 8, border: "1.5px solid #bfdbfe",
              background: "#eff6ff", color: "#2563eb", fontWeight: 700, fontSize: 11,
              cursor: "pointer", fontFamily: FF,
            }}>{"수정"}</button>
            <button onClick={() => onDelete(contact)} style={{
              padding: "5px 10px", borderRadius: 8, border: "1.5px solid #fca5a5",
              background: "#fef2f2", color: "#dc2626", fontWeight: 700, fontSize: 11,
              cursor: "pointer", fontFamily: FF,
            }}>{"삭제"}</button>
          </div>
        )}

        <div style={{ display: "grid", gap: 10, paddingRight: isAdmin ? 96 : 0 }}>
          <div style={{ display: "flex", alignItems: "baseline", gap: 8, flexWrap: "wrap" }}>
            <div style={{ fontWeight: 900, fontSize: isDesktop ? 20 : 17, color: "#0f172a", lineHeight: 1.1 }}>
              {contact.name || "-"}
            </div>
            {contact.dept && <span style={{ fontSize: 14, fontWeight: 700, color: "#0891b2" }}>{contact.dept}</span>}
            {contact.position && <span style={{ fontSize: 14, fontWeight: 700, color: "#64748b" }}>{contact.position}</span>}
            {activeSchedules.map((schedule) => {
              const meta = SCHEDULE_META[schedule.type] || { label: schedule.type, icon: "??" };
              return (
                <span key={schedule.id} style={{ fontSize: 13, fontWeight: 800, color: meta.icon === "??" ? "#be185d" : "#2563eb" }}>
                  {meta.icon} {meta.label} {formatSchedulePeriod(schedule)}
                </span>
              );
            })}
          </div>

          <div style={{ display: "grid", gap: 8, fontSize: 14, fontWeight: 700, color: "#334155" }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
              <span style={{ color: "#64748b" }}>{"휴대전화"}</span>
              {phoneDigits && (
                <button onClick={revealPhone} style={{
                  padding: "4px 12px",
                  borderRadius: 10,
                  border: showPhone ? "1.5px solid #475569" : "1.5px solid #2563eb",
                  background: "#fff",
                  color: showPhone ? "#475569" : "#2563eb",
                  fontWeight: 800,
                  fontSize: 12,
                  cursor: "pointer",
                  fontFamily: FF,
                  whiteSpace: "nowrap",
                }}>
                  {showPhone ? "숨기기 " + countdown + "s" : "보기"}
                </button>
              )}
              <span style={{ color: "#0f172a", fontWeight: 800, whiteSpace: "nowrap" }}>{visiblePhone}</span>

              {!splitDeskLine && hasExt && (
                <>
                  <span style={{ color: "#cbd5e1" }}>|</span>
                  <span style={{ whiteSpace: "nowrap" }}><span style={{ color: "#64748b" }}>{"내선번호"}</span>{" " + contact.ext}</span>
                </>
              )}

              {!splitDeskLine && hasDirect && (
                <>
                  <span style={{ color: "#cbd5e1" }}>|</span>
                  <span style={{ whiteSpace: "nowrap" }}>
                    <span style={{ color: "#64748b" }}>{"직통번호"}</span>{" "}
                    <a href={"tel:" + directDigits} style={{ color: "#0f172a", textDecoration: "none", fontWeight: 800 }}>
                      {formatPhone(contact.direct)}
                    </a>
                  </span>
                </>
              )}
            </div>

            {splitDeskLine && (
              <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
                <span style={{ whiteSpace: "nowrap" }}><span style={{ color: "#64748b" }}>{"내선번호"}</span>{" " + contact.ext}</span>
                <span style={{ color: "#cbd5e1" }}>|</span>
                <span style={{ whiteSpace: "nowrap" }}>
                  <span style={{ color: "#64748b" }}>{"직통번호"}</span>{" "}
                  <a href={"tel:" + directDigits} style={{ color: "#0f172a", textDecoration: "none", fontWeight: 800 }}>
                    {formatPhone(contact.direct)}
                  </a>
                </span>
              </div>
            )}
          </div>

          {historySchedules.length > 0 && (
            <div style={{ display: "grid", gap: 10 }}>
              <button
                onClick={() =>
                  setHistoryOpen((prev) => {
                    const next = !prev;
                    if (next) setHistoryLimit(initialHistoryCount);
                    return next;
                  })
                }
                style={{
                  justifySelf: "flex-start",
                  padding: "6px 12px",
                  borderRadius: 10,
                  border: "1px solid #cbd5e1",
                  background: "#fff",
                  color: "#475569",
                  fontWeight: 700,
                  fontSize: 12,
                  cursor: "pointer",
                  fontFamily: FF,
                }}
              >
                {historyOpen ? "일정 히스토리 닫기" : "일정 히스토리 " + historySchedules.length + "건"}
              </button>

              {historyOpen && (
                <div style={{
                  display: "grid",
                  gap: 8,
                  padding: "12px 14px",
                  borderRadius: 12,
                  background: "#f8fafc",
                  border: "1px solid #e2e8f0",
                }}>
                  {historySchedules.slice(0, historyLimit).map((schedule) => {
                    const meta = SCHEDULE_META[schedule.type] || { label: schedule.type, icon: "??" };
                    const status = getScheduleStatus(schedule, todayYmd);
                    const statusLabel = status === "active" ? "현재" : status === "upcoming" ? "예정" : "종료";
                    const statusColor = status === "active" ? "#2563eb" : status === "upcoming" ? "#7c3aed" : "#94a3b8";
                    return (
                      <div key={schedule.id} style={{
                        display: "grid",
                        gridTemplateColumns: isDesktop ? "1fr auto" : "1fr",
                        gap: 8,
                        padding: "10px 12px",
                        borderRadius: 10,
                        background: "#fff",
                        border: "1px solid #e2e8f0",
                      }}>
                        <div style={{ display: "flex", alignItems: "center", gap: 8, minWidth: 0, flexWrap: "wrap" }}>
                          <span style={{ fontSize: 13, fontWeight: 800, color: "#0f172a", whiteSpace: "nowrap", flexShrink: 0 }}>{meta.icon} {meta.label}</span>
                          <span style={{ fontSize: 12, fontWeight: 800, color: statusColor, whiteSpace: "nowrap", flexShrink: 0 }}>{statusLabel}</span>
                          <span style={{ fontSize: 12, color: "#475569", fontWeight: 700, whiteSpace: "nowrap", minWidth: 0, overflow: "hidden", textOverflow: "ellipsis" }}>{formatSchedulePeriod(schedule)}</span>
                        </div>
                        {canEditSchedule({ isAdmin, contact, schedule, userProfile }) && (
                          <div style={{ display: "inline-flex", alignItems: "center", gap: 6, justifySelf: isDesktop ? "end" : "start", flexWrap: "nowrap", whiteSpace: "nowrap" }}>
                            <button
                              onClick={() => setEditingSchedule(schedule)}
                              style={{ padding: "4px 8px", borderRadius: 8, border: "1px solid #bfdbfe", background: "#eff6ff", color: "#2563eb", fontWeight: 700, fontSize: 11, cursor: "pointer", fontFamily: FF, whiteSpace: "nowrap", flexShrink: 0 }}
                            >
                              {"수정"}
                            </button>
                            {isAdmin && (
                              <button
                                onClick={async () => {
                                  if (!window.confirm("이 일정을 삭제할까요?")) return;
                                  await onDeleteSchedule(contact, schedule.id);
                                }}
                                style={{ padding: "4px 8px", borderRadius: 8, border: "1px solid #fecaca", background: "#fff1f2", color: "#dc2626", fontWeight: 700, fontSize: 11, cursor: "pointer", fontFamily: FF, whiteSpace: "nowrap", flexShrink: 0 }}
                              >
                                {"삭제"}
                              </button>
                            )}
                          </div>
                        )}
                        {!!String(schedule.detail || "").trim() && (
                          <div style={{ gridColumn: "1 / -1", fontSize: 12, color: "#475569", fontWeight: 700, whiteSpace: "pre-wrap", overflowWrap: "anywhere" }}>
                            {schedule.detail}
                          </div>
                        )}
                      </div>
                    );
                  })}

                  {historyLimit < historySchedules.length && (
                    <button
                      onClick={() => setHistoryLimit((prev) => Math.min(historySchedules.length, prev + 5))}
                      style={{
                        justifySelf: "flex-start",
                        padding: "8px 12px",
                        borderRadius: 10,
                        border: "1px solid #cbd5e1",
                        background: "#fff",
                        color: "#0f172a",
                        fontWeight: 800,
                        fontSize: 12,
                        cursor: "pointer",
                        fontFamily: FF,
                      }}
                    >
                      {`\uB354\uBCF4\uAE30 (+${Math.min(5, historySchedules.length - historyLimit)})`}
                    </button>
                  )}
                </div>
              )}
            </div>
          )}
        </div>
      </div>
    </>
  );
}

function NoticeBoardView({ notices, isAdmin, user, onAddNotice, onUpdateNotice, onDeleteNotice, showToast }) {
  const [form, setForm] = useState({ title: "", content: "", pinned: false, isHtml: false });
  const [editingId, setEditingId] = useState(null);

  const editNotice = useMemo(
    () => (editingId ? notices.find((n) => n.id === editingId) || null : null),
    [editingId, notices]
  );

  useEffect(() => {
    if (!editNotice) return;
    setForm({
      title: String(editNotice.title || ""),
      content: String(editNotice.content || ""),
      pinned: !!editNotice.pinned,
      isHtml: !!editNotice.isHtml,
    });
  }, [editNotice]);

  function resetForm() {
    setEditingId(null);
    setForm({ title: "", content: "", pinned: false, isHtml: false });
  }

  async function handleSubmit() {
    if (!form.title.trim() || !form.content.trim()) {
      showToast("제목과 내용을 입력해주세요.", "error");
      return;
    }
    if (editingId) {
      await onUpdateNotice(editingId, form);
      showToast("공지사항이 수정되었습니다.");
    } else {
      await onAddNotice({
        ...form,
        author: String(user?.email || "").trim(),
      });
      showToast("공지사항이 등록되었습니다.");
    }
    resetForm();
  }

  return (
    <div style={{ maxWidth: 960, margin: "0 auto", padding: "20px 16px 36px", display: "grid", gap: 14 }}>
      <div style={{
        borderRadius: 16,
        border: "1px solid #e2e8f0",
        background: "linear-gradient(135deg, #eff6ff 0%, #f8fafc 100%)",
        padding: "16px 18px",
      }}>
        <div style={{ fontSize: 18, fontWeight: 900, color: "#0f172a", marginBottom: 4 }}>공지사항</div>
        <div style={{ color: "#64748b", fontWeight: 700, fontSize: 13 }}>
          관리자 작성 · 전 직원 열람
        </div>
      </div>

      {isAdmin && (
        <div style={{
          borderRadius: 16,
          border: "1px solid #e2e8f0",
          background: "#fff",
          padding: 14,
          display: "grid",
          gap: 10,
        }}>
          <input
            value={form.title}
            onChange={(e) => setForm((prev) => ({ ...prev, title: e.target.value }))}
            placeholder="공지 제목"
            style={base.input}
          />
          <textarea
            value={form.content}
            onChange={(e) => setForm((prev) => ({ ...prev, content: e.target.value }))}
            placeholder="공지 내용"
            style={{ ...base.input, minHeight: 100, resize: "vertical", lineHeight: 1.5 }}
          />
          <label style={{ display: "flex", alignItems: "center", gap: 8, color: "#475569", fontWeight: 700, fontSize: 13 }}>
            <input
              type="checkbox"
              checked={form.pinned}
              onChange={(e) => setForm((prev) => ({ ...prev, pinned: e.target.checked }))}
            />
            상단 고정
          </label>
          <label style={{ display: "flex", alignItems: "center", gap: 8, color: "#475569", fontWeight: 700, fontSize: 13 }}>
            <input
              type="checkbox"
              checked={form.isHtml}
              onChange={(e) => setForm((prev) => ({ ...prev, isHtml: e.target.checked }))}
            />
            HTML로 작성
          </label>
          <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", flexWrap: "wrap" }}>
            {editingId && (
              <button onClick={resetForm} style={{
                padding: "8px 12px", borderRadius: 10, border: "1px solid #cbd5e1",
                background: "#fff", color: "#475569", fontWeight: 700, cursor: "pointer", fontFamily: FF,
              }}>취소</button>
            )}
            <button onClick={handleSubmit} style={{
              padding: "8px 14px", borderRadius: 10, border: "none",
              background: "#2563eb", color: "#fff", fontWeight: 800, cursor: "pointer", fontFamily: FF,
            }}>
              {editingId ? "공지 수정" : "공지 등록"}
            </button>
          </div>
        </div>
      )}

      <div style={{ display: "grid", gap: 10 }}>
        {notices.length === 0 && (
          <div style={{ color: "#94a3b8", fontWeight: 700, textAlign: "center", padding: "24px 0" }}>
            등록된 공지사항이 없습니다.
          </div>
        )}
        {notices.map((notice) => (
          <article key={notice.id} style={{
            borderRadius: 14,
            border: "1px solid #e2e8f0",
            background: "#fff",
            padding: "14px 16px",
            boxShadow: "0 3px 12px rgba(15,23,42,0.05)",
          }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
              {notice.pinned && (
                <span style={{
                  fontSize: 11, fontWeight: 900, color: "#b45309",
                  border: "1px solid #fdba74", background: "#ffedd5",
                  borderRadius: 999, padding: "2px 8px",
                }}>
                  상단고정
                </span>
              )}
              <h3 style={{ margin: 0, fontSize: 16, color: "#0f172a", fontWeight: 900 }}>{notice.title}</h3>
            </div>
            {notice.isHtml ? (
              <div
                style={{ margin: "8px 0 0", color: "#334155", fontWeight: 600, lineHeight: 1.55 }}
                dangerouslySetInnerHTML={{ __html: notice.content || "" }}
              />
            ) : (
              <p style={{ margin: "8px 0 0", color: "#334155", fontWeight: 600, whiteSpace: "pre-wrap", lineHeight: 1.55 }}>
                {notice.content}
              </p>
            )}
            <div style={{ marginTop: 10, fontSize: 12, color: "#94a3b8", fontWeight: 700, display: "flex", gap: 10, flexWrap: "wrap" }}>
              <span>{notice.author || "-"}</span>
              <span>{notice.createdAt ? new Date(notice.createdAt).toLocaleString("ko-KR") : ""}</span>
            </div>
            {isAdmin && (
              <div style={{ marginTop: 10, display: "flex", gap: 6, justifyContent: "flex-end" }}>
                <button onClick={() => setEditingId(notice.id)} style={{
                  padding: "5px 10px", borderRadius: 8, border: "1.5px solid #bfdbfe",
                  background: "#eff6ff", color: "#2563eb", fontWeight: 700, fontSize: 12, cursor: "pointer", fontFamily: FF,
                }}>수정</button>
                <button onClick={async () => {
                  if (!window.confirm("이 공지사항을 삭제할까요?")) return;
                  await onDeleteNotice(notice.id);
                  showToast("공지사항이 삭제되었습니다.");
                  if (editingId === notice.id) resetForm();
                }} style={{
                  padding: "5px 10px", borderRadius: 8, border: "1.5px solid #fecaca",
                  background: "#fff1f2", color: "#dc2626", fontWeight: 700, fontSize: 12, cursor: "pointer", fontFamily: FF,
                }}>삭제</button>
              </div>
            )}
          </article>
        ))}
      </div>
    </div>
  );
}

function CompanyInfoView({ companyInfo, isAdmin, user, onSaveCompanyInfo, onUploadAttachment, onDeleteAttachment, showToast }) {
  const [editing, setEditing] = useState(false);
  const [uploading, setUploading] = useState("");
  const [viewTab, setViewTab] = useState("address");
  const [form, setForm] = useState(DEFAULT_COMPANY_INFO);
  const isDesktop = useMediaQuery("(min-width: 900px)");
  const bizFileRef = useRef(null);
  const bankFileRef = useRef(null);
  const etcFileRef = useRef(null);
  const [previewState, setPreviewState] = useState(null);

  useEffect(() => {
    setForm({
      ...DEFAULT_COMPANY_INFO,
      ...(companyInfo || {}),
    });
  }, [companyInfo]);

  const rawAttachments = Array.isArray(companyInfo?.attachments)
    ? companyInfo.attachments
    : (companyInfo?.attachments && typeof companyInfo.attachments === "object")
      ? Object.values(companyInfo.attachments)
      : [];
  const attachments = rawAttachments
    .filter((x) => x && typeof x === "object")
    .map((x) => ({
      ...x,
      category: String(x?.category || "etc").toLowerCase(),
      name: String(x?.name || "attachment"),
      url: String(x?.url || ""),
    }));
  const bizFiles = attachments.filter((x) => x.category === "biz");
  const bankFiles = attachments.filter((x) => x.category === "bank");
  const etcFiles = attachments.filter((x) => x.category === "etc");
  const uncategorizedFiles = attachments.filter((x) => !["biz", "bank", "etc"].includes(x.category));
  const allFiles = [...bizFiles, ...bankFiles, ...etcFiles, ...uncategorizedFiles];

  useEffect(() => {
    return () => {
      if (previewState?.objectUrl) URL.revokeObjectURL(previewState.objectUrl);
    };
  }, [previewState]);

  function formatBytes(bytes) {
    const n = Number(bytes || 0);
    if (!n) return "-";
    if (n >= 1024 * 1024) return `${(n / (1024 * 1024)).toFixed(1)} MB`;
    if (n >= 1024) return `${Math.round(n / 1024)} KB`;
    return `${n} B`;
  }

  function dataUrlToBlob(dataUrl) {
    try {
      const [meta, base64] = String(dataUrl || "").split(",");
      if (!meta || !base64) return null;
      const mime = (meta.match(/data:([^;]+);base64/i) || [])[1] || "application/octet-stream";
      const binary = atob(base64);
      const bytes = new Uint8Array(binary.length);
      for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
      return new Blob([bytes], { type: mime });
    } catch {
      return null;
    }
  }

  function dataUrlToText(dataUrl) {
    try {
      const blob = dataUrlToBlob(dataUrl);
      if (!blob) return "";
      return blob.text();
    } catch {
      return Promise.resolve("");
    }
  }

  function getPreviewKind(item) {
    const type = String(item?.type || "").toLowerCase();
    const name = String(item?.name || "").toLowerCase();
    if (type.startsWith("image/") || /\.(png|jpe?g|gif|webp|bmp|svg)$/i.test(name)) return "image";
    if (type === "application/pdf" || /\.pdf$/i.test(name)) return "pdf";
    if (type.startsWith("text/") || /\.(txt|md|csv)$/i.test(name)) return "text";
    return "unsupported";
  }

  function canPreview(item) {
    return getPreviewKind(item) !== "unsupported";
  }

  async function openPreview(item) {
    const href = String(item?.url || "").trim();
    if (!href) {
      showToast("미리보기 링크가 없습니다.", "error");
      return;
    }
    if (!canPreview(item)) {
      showToast("이 형식은 미리보기를 지원하지 않습니다. 다운로드를 이용해주세요.", "error");
      return;
    }
    if (previewState?.objectUrl) URL.revokeObjectURL(previewState.objectUrl);

    const kind = getPreviewKind(item);
    if (href.startsWith("data:")) {
      if (kind === "text") {
        const text = await dataUrlToText(href);
        setPreviewState({ item, kind, src: "", text, objectUrl: "" });
        return;
      }
      const blob = dataUrlToBlob(href);
      if (!blob) {
        showToast("미리보기 변환에 실패했습니다.", "error");
        return;
      }
      const objectUrl = URL.createObjectURL(blob);
      setPreviewState({ item, kind, src: objectUrl, text: "", objectUrl });
      return;
    }
    setPreviewState({ item, kind, src: href, text: "", objectUrl: "" });
  }

  function downloadAttachment(item) {
    const href = String(item?.url || "").trim();
    if (!href) {
      showToast("다운로드 링크가 없습니다.", "error");
      return;
    }
    if (href.startsWith("data:")) {
      const blob = dataUrlToBlob(href);
      if (!blob) {
        showToast("다운로드 변환에 실패했습니다.", "error");
        return;
      }
      const objectUrl = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = objectUrl;
      a.download = item?.name || "attachment";
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(objectUrl), 10000);
      return;
    }
    const a = document.createElement("a");
    a.href = href;
    a.download = item?.name || "attachment";
    a.target = "_blank";
    a.rel = "noopener noreferrer";
    document.body.appendChild(a);
    a.click();
    a.remove();
  }

  async function handleSave() {
    await onSaveCompanyInfo(form);
    setEditing(false);
    showToast("회사 정보가 저장되었습니다.");
  }

  async function handleFileUpload(file, category) {
    if (!file) return;
    if (file.size > 3 * 1024 * 1024) {
      showToast("Storage 미사용 정책으로 3MB 이하 파일만 첨부할 수 있습니다.", "error");
      return;
    }
    setUploading(category);
    try {
      await onUploadAttachment(file, category);
      showToast("파일이 첨부되었습니다.");
    } catch (error) {
      const message = String(error?.message || error?.code || "");
      showToast(
        message ? `파일 첨부 실패: ${message}` : "파일 첨부에 실패했습니다. 권한 설정을 확인해주세요.",
        "error"
      );
    } finally {
      setUploading("");
    }
  }

  const row = (label, key, placeholder) => (
    <div style={{ display: "grid", gap: 5 }}>
      <label style={{ fontSize: 12, fontWeight: 800, color: "#64748b" }}>{label}</label>
      <input
        value={form[key]}
        onChange={(e) => setForm((prev) => ({ ...prev, [key]: e.target.value }))}
        placeholder={placeholder}
        style={base.input}
      />
    </div>
  );

  const section = (title, children) => (
    <section style={{
      borderRadius: 14,
      border: "1px solid #e2e8f0",
      background: "#fff",
      padding: "14px 16px",
      display: "grid",
      gap: 10,
    }}>
      <div style={{ fontSize: 14, fontWeight: 900, color: "#0f172a" }}>{title}</div>
      {children}
    </section>
  );
  const panel = (children) => (
    <section style={{
      borderRadius: 14,
      border: "1px solid #e2e8f0",
      background: "#fff",
      padding: "14px 16px",
      display: "grid",
      gap: 10,
      textAlign: "left",
    }}>
      {children}
    </section>
  );

  const fileList = (items, category) => (
    <div style={{ display: "grid", gap: 8 }}>
      {items.length === 0 && <div style={{ color: "#94a3b8", fontSize: 12, fontWeight: 700 }}>등록된 파일 없음</div>}
      {items.map((item) => {
        const href = String(item?.url || "").trim();
        return (
          <div key={item.id || `${item.name}_${item.uploadedAt || ""}`} style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap", padding: "8px 10px", border: "1px solid #e2e8f0", borderRadius: 10, background: "#fff" }}>
            <span style={{ color: "#334155", fontWeight: 700, flex: 1, minWidth: 160 }}>{item.name}</span>
            <span style={{ color: "#94a3b8", fontWeight: 700, fontSize: 12 }}>{formatBytes(item.size)}</span>
            {href ? (
              <>
                <button
                  onClick={() => openPreview(item)}
                  style={{ border: "none", background: "transparent", color: "#2563eb", fontWeight: 800, cursor: "pointer", padding: 0, fontFamily: FF }}
                >
                  미리보기
                </button>
                <button
                  onClick={() => downloadAttachment(item)}
                  style={{ border: "none", background: "transparent", color: "#0f766e", fontWeight: 800, cursor: "pointer", padding: 0, fontFamily: FF }}
                >
                  다운로드
                </button>
              </>
            ) : (
              <span style={{ color: "#94a3b8", fontWeight: 700, fontSize: 12 }}>링크 없음</span>
            )}
            {isAdmin && (
              <button onClick={async () => {
                if (!window.confirm("이 파일을 삭제할까요?")) return;
                await onDeleteAttachment(item.id);
                showToast("파일이 삭제되었습니다.");
              }} style={{
                padding: "4px 8px", borderRadius: 8, border: "1px solid #fecaca",
                background: "#fff1f2", color: "#dc2626", fontWeight: 700, fontSize: 12, cursor: "pointer", fontFamily: FF,
              }}>
                삭제
              </button>
            )}
          </div>
        )
      })}
      {isAdmin && (
        <button
          onClick={() => {
            if (category === "biz") bizFileRef.current?.click();
            if (category === "bank") bankFileRef.current?.click();
            if (category === "etc") etcFileRef.current?.click();
          }}
          disabled={uploading === category}
          style={{
            justifySelf: "flex-start",
            padding: "7px 12px", borderRadius: 10, border: "1px solid #bfdbfe",
            background: "#eff6ff", color: "#2563eb", fontWeight: 800, cursor: "pointer", fontFamily: FF,
          }}
        >
          {uploading === category ? "업로드 중..." : "파일 첨부"}
        </button>
      )}
    </div>
  );

  const mainAccountParts = String(form.accountMain || "")
    .split("/")
    .map((v) => v.trim())
    .filter(Boolean);
  const primaryAccount = mainAccountParts[0] || "-";
  const secondaryAccounts = mainAccountParts.slice(1);

  return (
    <div style={{ maxWidth: 960, margin: "0 auto", padding: "20px 16px 36px", display: "grid", gap: 14 }}>
      <input ref={bizFileRef} type="file" accept="image/*,.pdf,.doc,.docx,.xls,.xlsx,.ppt,.pptx,.hwp,.hwpx,.txt,.zip" style={{ display: "none" }} onChange={(e) => { handleFileUpload(e.target.files?.[0], "biz"); e.target.value = ""; }} />
      <input ref={bankFileRef} type="file" accept="image/*,.pdf,.doc,.docx,.xls,.xlsx,.ppt,.pptx,.hwp,.hwpx,.txt,.zip" style={{ display: "none" }} onChange={(e) => { handleFileUpload(e.target.files?.[0], "bank"); e.target.value = ""; }} />
      <input ref={etcFileRef} type="file" accept="image/*,.pdf,.doc,.docx,.xls,.xlsx,.ppt,.pptx,.hwp,.hwpx,.txt,.zip" style={{ display: "none" }} onChange={(e) => { handleFileUpload(e.target.files?.[0], "etc"); e.target.value = ""; }} />

      {previewState && (
        <Modal
          title={previewState.item?.name || "파일 미리보기"}
          onClose={() => {
            if (previewState?.objectUrl) URL.revokeObjectURL(previewState.objectUrl);
            setPreviewState(null);
          }}
          extra={
            <button
              onClick={() => downloadAttachment(previewState.item)}
              style={{ border: "none", background: "transparent", color: "#0f766e", fontWeight: 800, cursor: "pointer", padding: 0, fontFamily: FF }}
            >
              다운로드
            </button>
          }
        >
          {previewState.kind === "image" && (
            <img
              src={previewState.src}
              alt={previewState.item?.name || "preview"}
              style={{ width: "100%", maxHeight: "70vh", objectFit: "contain", borderRadius: 12, background: "#f8fafc" }}
            />
          )}
          {previewState.kind === "pdf" && (
            <iframe
              src={previewState.src}
              title={previewState.item?.name || "pdf-preview"}
              style={{ width: "100%", height: "70vh", border: "1px solid #e2e8f0", borderRadius: 12, background: "#fff" }}
            />
          )}
          {previewState.kind === "text" && (
            <pre style={{ margin: 0, maxHeight: "70vh", overflow: "auto", padding: 14, borderRadius: 12, border: "1px solid #e2e8f0", background: "#f8fafc", whiteSpace: "pre-wrap", overflowWrap: "anywhere", color: "#334155", fontFamily: "'Pretendard','Nanum Gothic',sans-serif" }}>
              {previewState.text || "(내용 없음)"}
            </pre>
          )}
        </Modal>
      )}

      <div style={{
        borderRadius: 20,
        border: "1px solid #dbeafe",
        background: "linear-gradient(130deg, #e0f2fe 0%, #eef2ff 45%, #f8fafc 100%)",
        padding: isDesktop ? "22px 24px" : "16px 16px",
        boxShadow: "0 16px 36px rgba(15,23,42,0.08)",
      }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap" }}>
          <div style={{ textAlign: "left" }}>
            <div style={{ fontSize: isDesktop ? 24 : 20, fontWeight: 900, color: "#0f172a", letterSpacing: "-0.02em" }}>
              {form.name || "미래오토메이션(주)"}
            </div>
            <div style={{
              fontWeight: 900,
              fontSize: 13,
              marginTop: 6,
              display: "block",
            }}>
              <span className="info-line business" style={{ color: "#0f766e" }}>
                사업자등록번호 {form.businessNo || "-"}
              </span>
              <span className="info-line hours" style={{ color: "#334155" }}>
                운영시간 {form.hours || "-"}
              </span>
            </div>
          </div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
            <span style={{ border: "1px solid #ccfbf1", background: "#f0fdfa", color: "#0f766e", borderRadius: 999, padding: "4px 10px", fontSize: 12, fontWeight: 800 }}>
              첨부파일 {allFiles.length}건
            </span>
            {isAdmin && (
              <button onClick={() => setEditing(true)} style={{
                padding: "5px 10px", borderRadius: 999, border: "1px solid #bfdbfe",
                background: "#eff6ff", color: "#2563eb", fontWeight: 800, cursor: "pointer", fontFamily: FF, fontSize: 12,
              }}>
                회사 정보 수정
              </button>
            )}
          </div>
        </div>
      </div>

      {editing ? (
        <div style={{ display: "grid", gap: 10 }}>
          {section("기본", <>
            {row("회사명", "name", "미래오토메이션(주)")}
            {row("사업자등록번호", "businessNo", "000-00-00000")}
            {row("회사 소개", "intro", "회사 소개")}
          </>)}
          {section("주소", <>
            {row("국문 주소", "address", "주소")}
            {row("영문 주소", "addressEn", "영문 주소")}
            {row("공장 주소", "factoryAddress", "공장 주소")}
          </>)}
          {section("연락처", <>
            {row("전화", "phone", "전화번호")}
            {row("팩스", "fax", "팩스번호")}
            {row("관리부 팩스", "faxAdmin", "관리부 팩스")}
            {row("대표 메일", "email", "mrat@mauto.co.kr")}
            {row("세금계산서 메일", "emailTax", "tax@mauto.co.kr")}
            {row("쇼핑몰 메일", "emailMall", "mall@mauto.co.kr")}
            {row("관리부 메일", "emailAdmin", "admin@mauto.co.kr")}
          </>)}
          {section("업무 정보", <>
            {row("주계좌", "accountMain", "계좌정보")}
            {row("보조 계좌", "accountSub", "계좌정보")}
            {row("쇼핑몰 계좌", "accountMall", "계좌정보")}
            {row("화물지점(경동)", "cargoKyungdong", "지점명")}
            {row("화물지점(대신)", "cargoDaesin", "지점명")}
            {row("ERP 내부", "erpInternal", "192.168.0.25 채널2")}
            {row("ERP 외부", "erpExternal", "121.175.14.212")}
            {row("SERP 재설치 URL", "serpInstallUrl", "https://")}
            {row("SERP 키보드보안 URL", "serpKeyboardUrl", "https://")}
            {row("홈페이지", "homepage", "https://")}
            {row("운영시간", "hours", "예: 평일 09:00~18:00")}
          </>)}
          <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", flexWrap: "wrap" }}>
            <button onClick={() => setEditing(false)} style={{ padding: "8px 12px", borderRadius: 10, border: "1px solid #cbd5e1", background: "#fff", color: "#475569", fontWeight: 700, cursor: "pointer", fontFamily: FF }}>취소</button>
            <button onClick={handleSave} style={{ padding: "8px 14px", borderRadius: 10, border: "none", background: "#2563eb", color: "#fff", fontWeight: 800, cursor: "pointer", fontFamily: FF }}>저장</button>
          </div>
        </div>
      ) : (
        <div style={{ display: "grid", gap: 10 }}>
          {(() => {
            const tabConfig = {
              address: { label: "주소", bg: "#e0f2fe", fg: "#0369a1" },
              contact: { label: "연락처", bg: "#fff7ed", fg: "#c2410c" },
              business: { label: "업무 정보", bg: "#ede9fe", fg: "#6d28d9" },
              files: { label: "첨부 파일", bg: "#dcfce7", fg: "#15803d" },
            };
            const tabs = ["address", "contact", "business", "files"];
            return (
              <div style={{
                display: "flex",
                gap: 8,
                flexWrap: "wrap",
                background: "#f8fafc",
                border: "1px solid #e2e8f0",
                borderRadius: 16,
                padding: 8,
              }}>
                {tabs.map((id) => {
                  const meta = tabConfig[id];
                  const active = viewTab === id;
                  return (
                    <button
                      key={id}
                      onClick={() => setViewTab(id)}
                      style={{
                        display: "inline-flex",
                        alignItems: "center",
                        gap: 6,
                        padding: "9px 13px",
                        borderRadius: 11,
                        border: active ? `1.5px solid ${meta.fg}` : "1.5px solid #e2e8f0",
                        background: active ? meta.bg : "#fff",
                        color: active ? meta.fg : "#475569",
                        fontWeight: 800,
                        cursor: "pointer",
                        fontFamily: FF,
                        boxShadow: active ? "0 6px 14px rgba(15,23,42,0.08)" : "none",
                      }}
                    >
                      {meta.label}
                    </button>
                  );
                })}
              </div>
            );
          })()}
          {viewTab === "address" && panel(<div style={{ display: "grid", gap: 8, color: "#334155", fontWeight: 700 }}>
            <div style={{ padding: "11px 12px", border: "1px solid #bae6fd", borderRadius: 12, background: "#f0f9ff", lineHeight: 1.6 }}>
              <div style={{ color: "#0369a1", fontWeight: 900, fontSize: 12, marginBottom: 4 }}>국문</div>
              <div style={{ fontSize: 14 }}>{form.address || "-"}</div>
            </div>
            <div style={{ padding: "11px 12px", border: "1px solid #bfdbfe", borderRadius: 12, background: "#f8fbff", lineHeight: 1.6 }}>
              <div style={{ color: "#2563eb", fontWeight: 900, fontSize: 12, marginBottom: 4 }}>영문</div>
              <div style={{ fontSize: 14 }}>{form.addressEn || "-"}</div>
            </div>
            <div style={{ padding: "11px 12px", border: "1px solid #c7d2fe", borderRadius: 12, background: "#f8faff", lineHeight: 1.6 }}>
              <div style={{ color: "#4338ca", fontWeight: 900, fontSize: 12, marginBottom: 4 }}>공장</div>
              <div style={{ fontSize: 14 }}>{form.factoryAddress || "-"}</div>
            </div>
          </div>)}

          {viewTab === "contact" && panel(<div style={{ display: "grid", gridTemplateColumns: isDesktop ? "1fr 1fr 1fr" : "1fr", gap: 10, color: "#334155", fontWeight: 700 }}>
            <div style={{ padding: "12px", border: "1px solid #fdba74", borderRadius: 12, background: "#fff7ed", display: "grid", gap: 7, alignContent: "start", alignItems: "start", textAlign: "left" }}>
              <div style={{ color: "#c2410c", fontWeight: 900, fontSize: 13 }}>전화</div>
              <div style={{ fontSize: 14, fontWeight: 900 }}>대표: {form.phone || "-"}</div>
            </div>
            <div style={{ padding: "12px", border: "1px solid #bae6fd", borderRadius: 12, background: "#f0f9ff", display: "grid", gap: 7, alignContent: "start", alignItems: "start", textAlign: "left" }}>
              <div style={{ color: "#0369a1", fontWeight: 900, fontSize: 13 }}>팩스</div>
              <div style={{ fontSize: 14, fontWeight: 900 }}>대표: {form.fax || "-"}</div>
              <div style={{ fontSize: 14 }}>관리부: {form.faxAdmin || "-"}</div>
            </div>
            <div style={{ padding: "12px", border: "1px solid #ddd6fe", borderRadius: 12, background: "#f5f3ff", display: "grid", gap: 7, alignContent: "start", alignItems: "start", textAlign: "left" }}>
              <div style={{ color: "#6d28d9", fontWeight: 900, fontSize: 13 }}>메일</div>
              <div style={{ fontSize: 14, fontWeight: 900 }}>대표: {form.email || "-"}</div>
              <div style={{ fontSize: 14 }}>세금계산서: {form.emailTax || "-"}</div>
              <div style={{ fontSize: 14 }}>쇼핑몰: {form.emailMall || "-"}</div>
              <div style={{ fontSize: 14 }}>관리부: {form.emailAdmin || "-"}</div>
            </div>
          </div>)}

          {viewTab === "business" && panel(<div style={{ display: "grid", gridTemplateColumns: isDesktop ? "1fr 1fr 1fr" : "1fr", gap: 10 }}>
            <div style={{ border: "1px solid #ddd6fe", borderRadius: 12, background: "#faf5ff", padding: "12px", color: "#334155", fontWeight: 700, display: "grid", gap: 7, textAlign: "left", alignContent: "start", alignItems: "start" }}>
              <div style={{ color: "#6d28d9", fontWeight: 900, fontSize: 13 }}>계좌번호</div>
              <div style={{ fontSize: 14, fontWeight: 900, color: "#111827" }}>{primaryAccount}</div>
              {secondaryAccounts.map((acc) => (
                <div key={acc} style={{ fontSize: 14 }}>{acc}</div>
              ))}
              {form.accountSub && <div style={{ fontSize: 14 }}>{form.accountSub}</div>}
              <div style={{ fontSize: 14 }}>쇼핑몰: {form.accountMall || "-"}</div>
            </div>
            <div style={{ border: "1px solid #bfdbfe", borderRadius: 12, background: "#eff6ff", padding: "12px", color: "#334155", fontWeight: 700, display: "grid", gap: 7, textAlign: "left", alignContent: "start", alignItems: "start" }}>
              <div style={{ color: "#1d4ed8", fontWeight: 900, fontSize: 13 }}>화물택배지점</div>
              <div style={{ fontSize: 14 }}>경동: {form.cargoKyungdong || "-"}</div>
              <div style={{ fontSize: 14 }}>대신: {form.cargoDaesin || "-"}</div>
            </div>
            <div style={{ border: "1px solid #a7f3d0", borderRadius: 12, background: "#ecfdf5", padding: "12px", color: "#334155", fontWeight: 700, display: "grid", gap: 7, textAlign: "left", alignContent: "start", alignItems: "start" }}>
              <div style={{ color: "#047857", fontWeight: 900, fontSize: 13 }}>ERP 관련</div>
              <div style={{ fontSize: 14 }}>ERP 내부: {form.erpInternal || "-"}</div>
              <div style={{ fontSize: 14 }}>ERP 외부: {form.erpExternal || "-"}</div>
              <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginTop: 2 }}>
                {form.serpInstallUrl && <a href={form.serpInstallUrl} target="_blank" rel="noopener noreferrer" style={{ color: "#2563eb", fontWeight: 800, textDecoration: "none", fontSize: 13, border: "1px solid #bfdbfe", background: "#eff6ff", borderRadius: 999, padding: "4px 9px" }}>SERP 재설치</a>}
                {form.serpKeyboardUrl && <a href={form.serpKeyboardUrl} target="_blank" rel="noopener noreferrer" style={{ color: "#2563eb", fontWeight: 800, textDecoration: "none", fontSize: 13, border: "1px solid #bfdbfe", background: "#eff6ff", borderRadius: 999, padding: "4px 9px" }}>키보드보안</a>}
                {form.homepage && <a href={form.homepage} target="_blank" rel="noopener noreferrer" style={{ color: "#2563eb", fontWeight: 800, textDecoration: "none", fontSize: 13, border: "1px solid #bfdbfe", background: "#eff6ff", borderRadius: 999, padding: "4px 9px" }}>홈페이지</a>}
              </div>
            </div>
          </div>)}

          {viewTab === "files" && panel(<div style={{ display: "grid", gap: 6 }}>
            <div style={{ fontSize: 12, color: "#64748b", fontWeight: 700 }}>
              사업자등록증/통장사본/기타 파일을 한 곳에서 관리합니다.
            </div>
            {fileList(allFiles, "etc")}
          </div>)}
        </div>
      )}
    </div>
  );
}

function AccountRow({ item, isFavorite, onOpen, onToggleFavorite, onToggleLock, isAdmin }) {
  const c = getCS(item.category);
  const [hovered, setHovered] = useState(false);
  const showLockControl = !!item.locked || hovered;
  return (
    <div
      onClick={onOpen}
      onMouseEnter={() => setHovered(true)}
      onMouseLeave={() => setHovered(false)}
      style={{
        display: "flex", alignItems: "center", gap: 10,
        padding: "8px 12px", cursor: "pointer",
        background: hovered ? "#f8fafc" : "transparent",
        borderBottom: "1px solid #f1f5f9",
        transition: "background 0.12s",
      }}>
      {/* 잠금 아이콘 */}
      {isAdmin && showLockControl && (
        <button onClick={e => { e.stopPropagation(); onToggleLock(); }}
          title={item.locked ? "잠금 해제" : "잠금 (관리자만 열람)"}
          style={{ background: "none", border: "none", cursor: "pointer", fontSize: 14, padding: "2px 2px", flexShrink: 0, color: item.locked ? "#dc2626" : "#94a3b8", transition: "color 0.15s" }}>
          {item.locked ? "🔒" : "🔓"}
        </button>
      )}
      {!isAdmin && item.locked && (
        <span style={{ fontSize: 14, color: "#dc2626", flexShrink: 0 }}>🔒</span>
      )}
      {/* 사이트명 */}
      <span style={{ fontWeight: 700, fontSize: 14, color: "#0f172a", minWidth: 120, maxWidth: 180, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{item.siteName}</span>
      {/* 카테고리 태그 */}
      <Tag label={item.category} style={{ color: c.color, background: c.bg, border: `1px solid ${c.border}`, flexShrink: 0 }} />
      {/* 아이디 */}
      <span style={{ fontSize: 13, color: "#64748b", flex: 1, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{item.loginId || ""}</span>
      {/* 담당자 */}
      {item.owner && <span style={{ fontSize: 12, color: "#94a3b8", flexShrink: 0, whiteSpace: "nowrap" }}>{item.owner}</span>}
      {/* 즐겨찾기 */}
      <button onClick={e => { e.stopPropagation(); onToggleFavorite(); }}
        style={{ background: "none", border: "none", cursor: "pointer", fontSize: 16, padding: "2px 4px", flexShrink: 0, color: isFavorite ? "#f59e0b" : "#e2e8f0", transition: "color 0.15s" }}>
        {isFavorite ? "★" : "☆"}
      </button>
    </div>
  );
}

/* ══════════════ 카테고리 섹션 ══════════════ */

function CategorySection({ title, items, expanded, onToggle, favoritesMap, onToggleFavorite, onOpen, onToggleLock, isAdmin }) {
  if (!items.length) return null;
  return (
    <div style={{ marginBottom: 4 }}>
      {/* 헤더 — 클릭으로 토글 */}
      <div onClick={onToggle} style={{ display: "flex", alignItems: "center", gap: 8, padding: "8px 12px", cursor: "pointer", borderRadius: 8, background: "#f8fafc", marginBottom: 2 }}>
        <span style={{ fontWeight: 800, fontSize: 14, color: "#0f172a" }}>{title}</span>
        <span style={{ fontSize: 12, color: "#94a3b8", fontWeight: 600 }}>{items.length}개</span>
        <div style={{ flex: 1 }} />
        <span style={{ fontSize: 12, color: "#94a3b8", transition: "transform 0.2s", display: "inline-block", transform: expanded ? "rotate(0deg)" : "rotate(-90deg)" }}>▼</span>
      </div>
      {/* 목록 */}
      {expanded && (
        <div style={{ background: "#fff", border: "1px solid #f1f5f9", borderRadius: 8, overflow: "hidden", marginBottom: 4 }}>
          {items.map(item => (
            <AccountRow key={item.id} item={item} isAdmin={isAdmin}
              isFavorite={!!favoritesMap[item.id]}
              onOpen={() => onOpen(item)}
              onToggleFavorite={() => onToggleFavorite(item)}
              onToggleLock={() => onToggleLock(item)} />
          ))}
        </div>
      )}
    </div>
  );
}

/* ══════════════ 미니 섹션 (즐겨찾기 / 최근) ══════════════ */

function MiniSection({ title, items, favoritesMap, onOpen, onToggleFavorite, onToggleLock, isAdmin }) {
  if (!items.length) return null;
  return (
    <>
      <Divider label={title} />
      <div style={{ background: "#fff", border: "1px solid #f1f5f9", borderRadius: 8, overflow: "hidden", marginBottom: 8 }}>
        {items.slice(0, 3).map(item => (
          <AccountRow key={item.id} item={item} isAdmin={isAdmin}
            isFavorite={!!favoritesMap[item.id]}
            onOpen={() => onOpen(item)}
            onToggleFavorite={() => onToggleFavorite(item)}
            onToggleLock={() => onToggleLock(item)} />
        ))}
      </div>
    </>
  );
}

/* ══════════════ 상세 화면 ══════════════ */

function DetailView({ item, user, isFavorite, onToggleFavorite, onBack, showToast, isAdmin, onEdit, onDelete, passwordHistory = [], colVis, onToggleCol }) {
  const [showPw, setShowPw] = useState(false);
  const [countdown, setCountdown] = useState(0);
  const [requestMsg, setRequestMsg] = useState("");
  const [sending, setSending] = useState(false);
  const [expandedHistoryId, setExpandedHistoryId] = useState(null);
  const timerRef = useRef(null);
  const [showColSettings, setShowColSettings] = useState(false);

  useEffect(() => () => clearInterval(timerRef.current), []);

  function handleReveal() {
    if (!showPw) {
      setShowPw(true);
      setCountdown(10);
      clearInterval(timerRef.current);
      timerRef.current = setInterval(() => {
        setCountdown(prev => {
          if (prev <= 1) { clearInterval(timerRef.current); setShowPw(false); return 0; }
          return prev - 1;
        });
      }, 1000);
      writeLog({ action: "비밀번호 열람", email: user.email, targetId: item.id, targetName: item.siteName });
    } else {
      clearInterval(timerRef.current);
      setShowPw(false);
      setCountdown(0);
    }
  }

  async function handleRequestUpdate() {
    if (!requestMsg.trim()) { showToast("수정 요청 내용을 입력해주세요.", "error"); return; }
    setSending(true);
    try {
      await writeLog({
        action: "수정 사항 요청", email: user.email,
        targetId: item.id, targetName: item.siteName, message: requestMsg.trim(),
      });
      await emailjs.send(
        EMAILJS_SERVICE_ID,
        EMAILJS_TEMPLATE_ID,
        {
          user_email: user.email,
          site_name: item.siteName,
          message: requestMsg.trim(),
        },
        { publicKey: EMAILJS_PUBLIC_KEY }
      );
      setRequestMsg("");
      showToast("수정 요청이 관리자에게 전달되었습니다.");
    } catch (err) {
      console.error("EmailJS error:", err);
      showToast("전송에 실패했습니다. EmailJS 설정을 확인해주세요.", "error");
    } finally {
      setSending(false);
    }
  }

  const c = getCS(item.category);
  const hasPw = !!item.password;

  return (
    <div style={{ maxWidth: 720, margin: "0 auto", padding: "24px 16px 40px" }}>
      <div style={{
        background: "#fff", borderRadius: 20, padding: "24px 22px",
        boxShadow: "0 4px 24px rgba(15,23,42,0.07)", border: "1px solid #f1f5f9",
      }}>
        {/* 헤더 */}
        <div style={{ display: "flex", alignItems: "flex-start", gap: 12, marginBottom: 22 }}>
          <div style={{ flex: 1 }}>
            <div style={{ fontWeight: 900, fontSize: 22, color: "#0f172a", lineHeight: 1.3 }}>{item.siteName}</div>
            <div style={{ marginTop: 8, display: "flex", gap: 6, flexWrap: "wrap" }}>
              <Tag label={item.category} style={{ color: c.color, background: c.bg, border: `1px solid ${c.border}` }} />
              {colVis.owner && item.owner && <Tag label={`담당자 ${item.owner}`} style={{ color: "#64748b", background: "#f8fafc", border: "1px solid #e2e8f0" }} />}
            </div>
          </div>
          <div style={{ display: "flex", flexDirection: "column", alignItems: "flex-end", gap: 6, flexShrink: 0 }}>
            <button onClick={onToggleFavorite}
              style={{ background: "none", border: "none", cursor: "pointer", fontSize: 24, color: isFavorite ? "#f59e0b" : "#e2e8f0" }}>
              {isFavorite ? "★" : "☆"}
            </button>
            <button
              onClick={() => setShowColSettings(s => !s)}
              title="표시 항목 설정"
              style={{ background: "none", border: "1px solid #e2e8f0", borderRadius: 8, cursor: "pointer", fontSize: 12, padding: "3px 8px", color: showColSettings ? "#2563eb" : "#94a3b8", fontFamily: FF }}>
              ⚙ 표시설정
            </button>
          </div>
        </div>

        {/* 표시 항목 설정 패널 */}
        {showColSettings && (
          <div style={{ marginBottom: 16, padding: "12px 14px", background: "#f8fafc", border: "1px solid #e2e8f0", borderRadius: 12 }}>
            <div style={{ fontSize: 12, fontWeight: 800, color: "#64748b", marginBottom: 8 }}>표시할 항목 선택</div>
            <div style={{ display: "flex", flexWrap: "wrap", gap: "8px 18px" }}>
              {[["siteUrl", "URL"], ["loginId", "아이디"], ["password", "비밀번호"], ["owner", "담당자"], ["phone", "전화번호"], ["fax", "팩스번호"]].map(([k, label]) => (
                <label key={k} style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 13, cursor: "pointer", color: "#374151", userSelect: "none" }}>
                  <input type="checkbox" checked={!!colVis[k]} onChange={() => onToggleCol(k)} style={{ cursor: "pointer" }} />
                  {label}
                </label>
              ))}
            </div>
          </div>
        )}

        {/* 정보 박스들 */}
        <div style={{ display: "grid", gap: 10, marginBottom: 18 }}>
          {colVis.siteUrl && item.siteUrl && <IBox label="사이트 주소" value={item.siteUrl} isUrl />}
          {colVis.loginId && item.loginId && <IBox label="아이디" value={item.loginId} />}
          {colVis.password && hasPw && (
            <IBox
              label="비밀번호"
              value={showPw ? item.password : "••••••••••"}
            />
          )}
          {colVis.owner && item.owner && <IBox label="담당자" value={item.owner} />}
          {colVis.phone && item.phone && <IBox label="전화번호" value={item.phone} />}
          {colVis.fax && item.fax && <IBox label="팩스번호" value={item.fax} />}
          {item.note && <IBox label="비고" value={item.note} />}
        </div>

        {/* 비밀번호 섹션 */}
        {colVis.password && hasPw && (
          <>
            <div style={{
              padding: "12px 14px", background: "#eff6ff", border: "1px solid #bfdbfe",
              borderRadius: 12, fontSize: 13, color: "#1e40af", lineHeight: 1.7, marginBottom: 14,
            }}>
              🔒 비밀번호는 보안을 위해 숨겨져 있습니다.
              {showPw && countdown > 0 && <span style={{ color: "#dc2626", fontWeight: 800 }}> {countdown}초 후 자동 숨김</span>}
            </div>
            <button onClick={handleReveal}
              style={{
                width: "100%", padding: "12px", borderRadius: 12, border: "none",
                background: showPw ? "#475569" : "#2563eb", color: "#fff",
                fontSize: 14, fontWeight: 800, cursor: "pointer", fontFamily: FF, marginBottom: 10,
              }}>
              {showPw ? "🙈 비밀번호 숨기기" : "👁 비밀번호 보기"}
            </button>

            {/* 비밀번호 이력 */}
            {passwordHistory && passwordHistory.length > 0 && (
              <div style={{ marginTop: 14, padding: "12px 14px", background: "#f8fafc", border: "1px solid #e2e8f0", borderRadius: 10 }}>
                <div style={{ fontSize: 12, fontWeight: 800, color: "#64748b", marginBottom: 8 }}>📋 변경 이력 ({passwordHistory.length}건)</div>
                <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                  {passwordHistory.map((h, idx) => (
                    <div key={idx} style={{ padding: "8px 10px", background: "#fff", border: "1px solid #f1f5f9", borderRadius: 8, fontSize: 12 }}>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 4 }}>
                        <span style={{ color: "#0f172a", fontWeight: 700 }}>
                          {new Date(h.changedAt).toLocaleString('ko-KR')}
                        </span>
                        <span style={{ color: "#94a3b8", fontSize: 11 }}>{h.changedBy}</span>
                      </div>
                      <div style={{
                        padding: "6px 8px", background: "#f8fafc", borderRadius: 6,
                        fontSize: 11, color: "#0f172a", fontFamily: "monospace",
                        userSelect: "none", cursor: "pointer",
                        border: expandedHistoryId === idx ? "1px solid #cbd5e1" : "1px solid #e2e8f0",
                      }}
                        onClick={() => setExpandedHistoryId(expandedHistoryId === idx ? null : idx)}>
                        {expandedHistoryId === idx ? h.password : "•••••• (클릭하면 표시)"}
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </>
        )}

        {/* 수정 사항 요청 */}
        <div style={{
          border: "1px solid #f1f5f9", borderRadius: 14, padding: "16px 16px 14px",
          background: "#fafafa", marginBottom: 10,
        }}>
          <div style={{ fontWeight: 800, fontSize: 14, color: "#0f172a", marginBottom: 10 }}>📧 수정 사항 요청</div>
          <div style={{ fontSize: 13, color: "#94a3b8", marginBottom: 10, lineHeight: 1.6 }}>
            계정 정보 수정이 필요한 경우 내용을 적어 전송해주세요. 관리자에게 바로 전달됩니다.
          </div>
          <textarea
            value={requestMsg}
            onChange={e => setRequestMsg(e.target.value)}
            placeholder="예: 아이디 변경 요청, 담당자 변경, 비고 추가 등"
            style={{ ...base.input, minHeight: 90, resize: "vertical", marginBottom: 10 }}
          />
          <button onClick={handleRequestUpdate} disabled={sending}
            style={{
              width: "100%", padding: "11px", borderRadius: 10, border: "none",
              background: sending ? "#94a3b8" : "#16a34a", color: "#fff",
              fontSize: 14, fontWeight: 800, cursor: sending ? "not-allowed" : "pointer", fontFamily: FF,
            }}>
            {sending ? "전송 중..." : "전송하기"}
          </button>
        </div>

        {/* 관리자 버튼 */}
        {isAdmin && (
          <div style={{ display: "flex", gap: 8, marginBottom: 10 }}>
            <button onClick={onEdit}
              style={{ flex: 1, padding: "11px", borderRadius: 10, border: "1.5px solid #3b82f6", background: "#eff6ff", color: "#2563eb", fontWeight: 800, cursor: "pointer", fontFamily: FF }}>
              수정
            </button>
            <button onClick={onDelete}
              style={{ flex: 1, padding: "11px", borderRadius: 10, border: "1.5px solid #fca5a5", background: "#fef2f2", color: "#dc2626", fontWeight: 800, cursor: "pointer", fontFamily: FF }}>
              삭제
            </button>
          </div>
        )}

        <button onClick={onBack}
          style={{
            width: "100%", padding: "11px", borderRadius: 10,
            border: "1.5px solid #e2e8f0", background: "#fff", color: "#64748b",
            fontWeight: 700, cursor: "pointer", fontFamily: FF,
          }}>
          목록으로 돌아가기
        </button>
      </div>
    </div>
  );
}

function IBox({ label, value, isUrl, dim }) {
  return (
    <div style={{ border: "1px solid #f1f5f9", background: "#fafafa", borderRadius: 12, padding: "11px 14px" }}>
      <div style={{ fontSize: 11, fontWeight: 800, color: "#cbd5e1", textTransform: "uppercase", marginBottom: 4, letterSpacing: "0.06em" }}>{label}</div>
      <div style={{
        fontSize: 14, fontWeight: 600, color: dim ? "#cbd5e1" : "#0f172a",
        wordBreak: "break-all", whiteSpace: "pre-line",
        ...(isUrl ? { color: "#2563eb", textDecoration: "underline", cursor: "pointer" } : {}),
      }}
        onClick={isUrl ? () => window.open(value, "_blank", "noopener") : undefined}
      >
        {value}
      </div>
    </div>
  );
}

/* ══════════════ 계정 폼 ══════════════ */

function AccountForm({ form, setForm, onSave, onCancel, editMode, categories }) {
  const f = key => e => setForm(prev => ({ ...prev, [key]: e.target.value }));
  return (
    <div style={{ maxWidth: 720, margin: "0 auto", padding: "24px 16px" }}>
      <div style={{ background: "#fff", borderRadius: 18, padding: "24px 22px", boxShadow: "0 4px 24px rgba(15,23,42,0.07)" }}>
        <div style={{ fontWeight: 900, fontSize: 19, color: "#0f172a", marginBottom: 20 }}>
          {editMode ? "✏️ 계정 수정" : "➕ 계정 추가"}
        </div>
        <div style={{ display: "grid", gap: 12 }}>
          {[["siteName", "사이트명 *"], ["siteUrl", "사이트 URL"], ["loginId", "아이디"], ["password", "비밀번호"], ["owner", "담당자"]].map(([k, p]) => (
            <div key={k}>
              <div style={{ fontSize: 12, fontWeight: 700, color: "#64748b", marginBottom: 5 }}>{p}</div>
              <input placeholder={p} value={form[k]} onChange={f(k)} style={base.input} />
            </div>
          ))}
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <div>
              <div style={{ fontSize: 12, fontWeight: 700, color: "#64748b", marginBottom: 5 }}>전화번호</div>
              <input placeholder="전화번호" value={form.phone} onChange={f("phone")} style={base.input} />
            </div>
            <div>
              <div style={{ fontSize: 12, fontWeight: 700, color: "#64748b", marginBottom: 5 }}>팩스번호</div>
              <input placeholder="팩스번호" value={form.fax} onChange={f("fax")} style={base.input} />
            </div>
          </div>
          <div>
            <div style={{ fontSize: 12, fontWeight: 700, color: "#64748b", marginBottom: 5 }}>카테고리</div>
            <select value={form.category} onChange={f("category")} style={{ ...base.input }}>
              {categories.map(c => <option key={c.id} value={c.name}>{c.name}</option>)}
            </select>
          </div>
          <div>
            <div style={{ fontSize: 12, fontWeight: 700, color: "#64748b", marginBottom: 5 }}>비고</div>
            <textarea placeholder="비고" value={form.note} onChange={f("note")}
              style={{ ...base.input, minHeight: 80, resize: "vertical" }} />
          </div>
        </div>
        <div style={{ display: "flex", gap: 10, marginTop: 20 }}>
          <button onClick={onCancel}
            style={{ flex: 1, padding: "12px", borderRadius: 10, border: "1.5px solid #e2e8f0", background: "#fff", fontWeight: 700, cursor: "pointer", fontFamily: FF }}>
            취소
          </button>
          <button onClick={onSave}
            style={{ flex: 1, padding: "12px", borderRadius: 10, border: "none", background: "#2563eb", color: "#fff", fontWeight: 800, cursor: "pointer", fontFamily: FF }}>
            저장
          </button>
        </div>
      </div>
    </div>
  );
}

/* ══════════════ 모달 공통 ══════════════ */

function Modal({ title, children, onClose, extra }) {
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", zIndex: 50, display: "flex", alignItems: "center", justifyContent: "center", padding: 16 }}>
      <div style={{ background: "#fff", width: "100%", maxWidth: 860, maxHeight: "84vh", overflow: "auto", borderRadius: 18, padding: "22px 24px" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 18, flexWrap: "wrap" }}>
          <span style={{ fontWeight: 900, fontSize: 17, color: "#0f172a" }}>{title}</span>
          <div style={{ flex: 1 }} />
          {extra}
          <HBtn onClick={onClose}>닫기</HBtn>
        </div>
        {children}
      </div>
    </div>
  );
}

function LogModal({ logs, onClose, onDownload }) {
  return (
    <Modal title="사용 로그" onClose={onClose} extra={<HBtn onClick={onDownload}>CSV 다운로드</HBtn>}>
      <div style={{ display: "grid", gap: 8 }}>
        {logs.map(log => (
          <div key={log.id} style={{ border: "1px solid #f1f5f9", borderRadius: 10, padding: "10px 14px", background: "#fafafa" }}>
            <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
              <span style={{ fontWeight: 800, fontSize: 13, color: "#0f172a" }}>{log.action}</span>
              <span style={{ fontSize: 12, color: "#64748b" }}>{log.email}</span>
              {log.targetName && <span style={{ fontSize: 12, color: "#94a3b8" }}>· {log.targetName}</span>}
              <span style={{ fontSize: 11, color: "#cbd5e1", marginLeft: "auto" }}>
                {log.createdAt ? new Date(log.createdAt).toLocaleString("ko-KR") : ""}
              </span>
            </div>
          </div>
        ))}
      </div>
    </Modal>
  );
}

function CatTab({ categories, onAdd, onUpdate, onDelete }) {
  const [newName, setNewName] = useState("");
  const [editId, setEditId] = useState(null);
  const [editName, setEditName] = useState("");
  const [saving, setSaving] = useState(false);

  async function handleAdd() {
    if (!newName.trim()) return;
    setSaving(true);
    try { await onAdd(newName.trim()); setNewName(""); } finally { setSaving(false); }
  }

  async function handleUpdate(id) {
    if (!editName.trim()) return;
    setSaving(true);
    try { await onUpdate(id, editName.trim()); setEditId(null); setEditName(""); } finally { setSaving(false); }
  }

  return (
    <div style={{ display: "grid", gap: 8 }}>
      <div style={{ display: "flex", gap: 8, marginBottom: 4 }}>
        <input value={newName} onChange={e => setNewName(e.target.value)}
          onKeyDown={e => e.key === "Enter" && handleAdd()}
          placeholder="새 카테고리 이름"
          style={{ ...base.input, flex: 1 }} />
        <button onClick={handleAdd} disabled={saving || !newName.trim()} style={{
          padding: "10px 16px", borderRadius: 10, border: "none", fontFamily: FF,
          background: newName.trim() ? "#2563eb" : "#e2e8f0",
          color: newName.trim() ? "#fff" : "#94a3b8",
          fontWeight: 800, cursor: newName.trim() ? "pointer" : "not-allowed", whiteSpace: "nowrap",
        }}>+ 추가</button>
      </div>

      {categories.length === 0 && (
        <div style={{ color: "#cbd5e1", fontSize: 14, padding: "12px 0", textAlign: "center" }}>
          카테고리가 없습니다. 위에서 추가해주세요.
        </div>
      )}

      {categories.map((cat, idx) => (
        <div key={cat.id} style={{ border: "1px solid #f1f5f9", borderRadius: 12, padding: "10px 14px", background: "#fafafa", display: "flex", alignItems: "center", gap: 10 }}>
          <span style={{ fontSize: 13, color: "#cbd5e1", fontWeight: 700, width: 20 }}>{idx + 1}</span>
          {editId === cat.id ? (
            <>
              <input value={editName} onChange={e => setEditName(e.target.value)}
                onKeyDown={e => e.key === "Enter" && handleUpdate(cat.id)}
                autoFocus style={{ ...base.input, flex: 1, padding: "7px 10px", fontSize: 13 }} />
              <button onClick={() => handleUpdate(cat.id)} disabled={saving}
                style={{ padding: "7px 12px", borderRadius: 8, border: "none", background: "#2563eb", color: "#fff", fontWeight: 800, cursor: "pointer", fontFamily: FF, fontSize: 13 }}>저장</button>
              <button onClick={() => { setEditId(null); setEditName(""); }}
                style={{ padding: "7px 12px", borderRadius: 8, border: "1.5px solid #e2e8f0", background: "#fff", fontWeight: 700, cursor: "pointer", fontFamily: FF, fontSize: 13 }}>취소</button>
            </>
          ) : (
            <>
              <span style={{ flex: 1, fontWeight: 700, fontSize: 14, color: "#0f172a" }}>{cat.name}</span>
              <button onClick={() => { setEditId(cat.id); setEditName(cat.name); }}
                style={{ padding: "6px 12px", borderRadius: 8, border: "1.5px solid #3b82f6", background: "#eff6ff", color: "#2563eb", fontWeight: 700, cursor: "pointer", fontFamily: FF, fontSize: 13 }}>수정</button>
              <button onClick={() => { if (window.confirm(`"${cat.name}" 카테고리를 삭제할까요?`)) onDelete(cat.id); }}
                style={{ padding: "6px 12px", borderRadius: 8, border: "1.5px solid #fca5a5", background: "#fef2f2", color: "#dc2626", fontWeight: 700, cursor: "pointer", fontFamily: FF, fontSize: 13 }}>삭제</button>
            </>
          )}
        </div>
      ))}
    </div>
  );
}

function UsersModal({ admins, users, allUserPerms, deptPerms, categories, allCashAccess, onAddUser, onDeleteUser, onSaveUserProfile, onSaveUserPerm, onSaveDeptPerm, onBulkAddUsers, onSaveCashAccess, onAddCategory, onUpdateCategory, onDeleteCategory, onClose }) {
  const [activeTab, setActiveTab] = useState("users");
  const [newEmail, setNewEmail] = useState("");
  const [addingUser, setAddingUser] = useState(false);
  const [editingEmail, setEditingEmail] = useState(null);
  const [editProfile, setEditProfile] = useState({ name: "", department: "", position: "", isDeptHead: false });
  const [editPermType, setEditPermType] = useState("all");
  const [editPerms, setEditPerms] = useState([]);
  const [editCashRole, setEditCashRole] = useState(null); // null | "sales" | "finance" | "admin"
  const [editingDept, setEditingDept] = useState(null);
  const [editDeptPermType, setEditDeptPermType] = useState("all");
  const [editDeptPerms, setEditDeptPerms] = useState([]);
  const [saving, setSaving] = useState(false);
  const [addError, setAddError] = useState("");
  const [bulkMsg, setBulkMsg] = useState("");
  const userFileRef = useRef(null);

  function downloadUserTemplate() {
    const ws = XLSX.utils.aoa_to_sheet([
      ["이메일", "이름", "부서", "직위"],
      ["yhj@mauto.co.kr", "여현진", "영업팀", "과장"],
    ]);
    ws["!cols"] = Array(4).fill({ wch: 20 });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "사용자등록양식");
    XLSX.writeFile(wb, "사용자등록양식.xlsx");
  }

  async function handleUserExcel(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    setBulkMsg("");
    const reader = new FileReader();
    reader.onload = async ev => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "array" });
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "" });
        const items = rows
          .filter(r => r["이메일"])
          .map(r => ({
            email: String(r["이메일"]).trim(),
            name: String(r["이름"] || "").trim(),
            department: String(r["부서"] || "").trim(),
            position: String(r["직위"] || "").trim(),
          }));
        if (!items.length) { setBulkMsg("등록할 데이터가 없습니다."); return; }
        await onBulkAddUsers(items);
        setBulkMsg(`✓ ${items.length}명 등록 완료`);
      } catch { setBulkMsg("파일 처리 중 오류가 발생했습니다."); }
      finally { e.target.value = ""; }
    };
    reader.readAsArrayBuffer(file);
  }

  const catNames = categories.map(c => c.name);
  const departments = [...new Set(users.map(u => u.department).filter(Boolean))].sort();

  // editPermType: "dept" | "all" | "custom"
  function startEditUser(u) {
    setEditingEmail(u.email);
    setEditProfile({ name: u.name || "", department: u.department || "", position: u.position || "", isDeptHead: !!u.isDeptHead });
    const raw = allUserPerms[emailToKey(u.email)] ?? null;
    if (raw === null) { setEditPermType("dept"); setEditPerms([]); }
    else if (raw === "__ALL__") { setEditPermType("all"); setEditPerms([]); }
    else { setEditPermType("custom"); setEditPerms(raw); }
    // 시재 권한 초기화
    const cashEntry = (allCashAccess || {})[emailToKey(u.email)];
    setEditCashRole(cashEntry?.role || null);
  }

  async function saveUser(email) {
    setSaving(true);
    try {
      await onSaveUserProfile(email, editProfile);
      const mode = editPermType === "dept" ? "dept"
        : editPermType === "all" ? "all"
          : editPerms;
      await onSaveUserPerm(email, mode);
      await onSaveCashAccess(email, editCashRole ? { role: editCashRole } : null);
      setEditingEmail(null);
    } finally { setSaving(false); }
  }

  async function handleAddUser() {
    if (!newEmail.trim()) return;
    setAddingUser(true);
    setAddError("");
    try {
      const email = newEmail.trim().includes("@") ? newEmail.trim() : `${newEmail.trim()}@mauto.co.kr`;
      await onAddUser(email);
      setNewEmail("");
    } catch (e) {
      setAddError("추가 실패: " + (e?.message || "Firebase 권한을 확인해주세요."));
    } finally { setAddingUser(false); }
  }

  function startEditDept(dept) {
    setEditingDept(dept);
    const perms = deptPerms[departmentToKey(dept)] ?? null;
    setEditDeptPermType(perms ? "custom" : "all");
    setEditDeptPerms(perms || []);
  }

  async function saveDept(dept) {
    setSaving(true);
    try {
      await onSaveDeptPerm(dept, editDeptPermType === "all" ? null : editDeptPerms);
      setEditingDept(null);
    } finally { setSaving(false); }
  }

  function toggleCat(cat, setList) {
    setList(prev => prev.includes(cat) ? prev.filter(c => c !== cat) : [...prev, cat]);
  }

  const tabBtn = (id, label) => (
    <button onClick={() => setActiveTab(id)} style={{
      padding: "8px 16px", borderRadius: 8, border: "none", cursor: "pointer", fontFamily: FF,
      fontWeight: 700, fontSize: 13,
      background: activeTab === id ? "#2563eb" : "#f1f5f9",
      color: activeTab === id ? "#fff" : "#64748b",
    }}>{label}</button>
  );

  const permBadge = (raw) => {
    if (raw === null) return <span style={{ fontSize: 12, color: "#64748b", background: "#f1f5f9", border: "1px solid #e2e8f0", borderRadius: 6, padding: "2px 8px", fontWeight: 700 }}>부서별</span>;
    if (raw === "__ALL__") return <span style={{ fontSize: 12, color: "#15803d", background: "#f0fdf4", border: "1px solid #86efac", borderRadius: 6, padding: "2px 8px", fontWeight: 700 }}>개인-전체허용</span>;
    if (raw.length === 0) return <span style={{ fontSize: 12, color: "#dc2626", background: "#fef2f2", border: "1px solid #fecaca", borderRadius: 6, padding: "2px 8px", fontWeight: 700 }}>접근 없음</span>;
    return <span style={{ fontSize: 12, color: "#7c3aed", background: "#f5f3ff", border: "1px solid #c4b5fd", borderRadius: 6, padding: "2px 8px", fontWeight: 700 }}>{raw.join(", ")}</span>;
  };

  const catCheckboxes = (checkedList, setList) => catNames.length === 0
    ? <div style={{ fontSize: 12, color: "#f59e0b", background: "#fffbeb", border: "1px solid #fde68a", borderRadius: 8, padding: "8px 12px", fontWeight: 600 }}>
      카테고리가 없습니다. "카테고리 관리" 탭에서 먼저 추가하세요.
    </div>
    : <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
      {catNames.map(cat => (
        <label key={cat} style={{ display: "flex", alignItems: "center", gap: 6, cursor: "pointer", fontSize: 13, fontWeight: 600, color: "#334155" }}>
          <input type="checkbox" checked={checkedList.includes(cat)}
            onChange={() => toggleCat(cat, setList)}
            style={{ width: 15, height: 15, cursor: "pointer" }} />
          {cat}
        </label>
      ))}
    </div>;

  return (
    <Modal title="사용자 관리" onClose={onClose}>
      {/* 탭 */}
      <div style={{ display: "flex", gap: 8, marginBottom: 16, flexWrap: "wrap" }}>
        {tabBtn("users", "사용자별 권한")}
        {tabBtn("depts", "부서별 권한")}
        {tabBtn("cats", "카테고리 관리")}
        {tabBtn("admins", "관리자 목록")}
      </div>

      {/* ── 사용자별 권한 탭 ── */}
      {activeTab === "users" && (
        <div style={{ display: "grid", gap: 8 }}>
          {/* 엑셀 일괄 업로드 */}
          <input ref={userFileRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={handleUserExcel} />
          <div style={{ display: "flex", gap: 8, marginBottom: 4 }}>
            <button onClick={downloadUserTemplate} style={{ padding: "9px 14px", borderRadius: 10, border: "1.5px solid #e2e8f0", background: "#fff", fontWeight: 700, fontSize: 13, cursor: "pointer", fontFamily: FF }}>📋 양식 다운로드</button>
            <button onClick={() => userFileRef.current?.click()} style={{ padding: "9px 14px", borderRadius: 10, border: "1.5px solid #86efac", background: "#f0fdf4", color: "#16a34a", fontWeight: 700, fontSize: 13, cursor: "pointer", fontFamily: FF }}>📥 엑셀 업로드</button>
            {bulkMsg && <span style={{ fontSize: 13, color: "#15803d", fontWeight: 700, alignSelf: "center" }}>{bulkMsg}</span>}
          </div>

          {/* 사용자 개별 추가 */}
          <div style={{ display: "flex", gap: 8, marginBottom: 4 }}>
            <input
              value={newEmail}
              onChange={e => { setNewEmail(e.target.value); setAddError(""); }}
              onKeyDown={e => e.key === "Enter" && handleAddUser()}
              placeholder="이메일 또는 아이디 (예: yhj)"
              style={{ ...base.input, flex: 1 }}
            />
            <button onClick={handleAddUser} disabled={addingUser || !newEmail.trim()} style={{
              padding: "10px 16px", borderRadius: 10, border: "none", fontFamily: FF,
              background: newEmail.trim() ? "#2563eb" : "#e2e8f0",
              color: newEmail.trim() ? "#fff" : "#94a3b8",
              fontWeight: 800, cursor: newEmail.trim() ? "pointer" : "not-allowed", whiteSpace: "nowrap",
            }}>
              {addingUser ? "추가 중..." : "+ 추가"}
            </button>
          </div>
          {addError && (
            <div style={{ background: "#fef2f2", border: "1px solid #fecaca", color: "#dc2626", borderRadius: 8, padding: "8px 12px", fontSize: 12, fontWeight: 700 }}>
              {addError}
            </div>
          )}
          <div style={{ fontSize: 12, color: "#94a3b8", marginBottom: 8 }}>
            @mauto.co.kr 생략 가능 · Firebase에 등록된 이메일만 로그인 가능
          </div>

          {users.length === 0 && (
            <div style={{ color: "#cbd5e1", fontSize: 14, padding: "12px 0", textAlign: "center" }}>
              위에서 사용자를 추가해주세요.
            </div>
          )}
          {users.map(u => (
            <div key={u.email} style={{ border: "1px solid #f1f5f9", borderRadius: 12, background: "#fafafa", overflow: "hidden" }}>
              {editingEmail === u.email ? (
                /* 편집 모드 */
                <div style={{ padding: "14px 16px", display: "grid", gap: 12 }}>
                  <div style={{ fontWeight: 800, fontSize: 13, color: "#0f172a" }}>{u.email}</div>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8 }}>
                    {[["name", "이름"], ["department", "부서"], ["position", "직위"]].map(([k, label]) => (
                      <div key={k}>
                        <div style={{ fontSize: 11, fontWeight: 700, color: "#94a3b8", marginBottom: 4 }}>{label}</div>
                        <input value={editProfile[k]}
                          onChange={e => setEditProfile(prev => ({ ...prev, [k]: e.target.value }))}
                          style={{ ...base.input, padding: "7px 10px", fontSize: 13 }} />
                      </div>
                    ))}
                  </div>
                  <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13, fontWeight: 700, color: "#334155" }}>
                    <input
                      type="checkbox"
                      checked={!!editProfile.isDeptHead}
                      onChange={(e) => setEditProfile((p) => ({ ...p, isDeptHead: e.target.checked }))}
                      style={{ width: 16, height: 16, cursor: "pointer" }}
                    />
                    부서장
                  </label>
                  <div>
                    <div style={{ fontSize: 11, fontWeight: 700, color: "#94a3b8", marginBottom: 8 }}>카테고리 접근 권한</div>
                    <div style={{ display: "flex", gap: 8, marginBottom: 10 }}>
                      {[["dept", "부서별"], ["all", "개인-전체허용"], ["custom", "개인-직접설정"]].map(([t, label]) => (
                        <button key={t} onClick={() => setEditPermType(t)} style={{
                          padding: "6px 12px", borderRadius: 8, border: "none", cursor: "pointer",
                          fontFamily: FF, fontWeight: 700, fontSize: 12,
                          background: editPermType === t ? "#2563eb" : "#f1f5f9",
                          color: editPermType === t ? "#fff" : "#64748b",
                        }}>{label}</button>
                      ))}
                    </div>
                    {editPermType === "dept" && (
                      <div style={{ fontSize: 12, color: "#64748b", background: "#f8fafc", borderRadius: 8, padding: "8px 12px" }}>
                        부서별 권한 탭에서 설정한 권한이 적용됩니다.
                        {editProfile.department
                          ? ` (현재 부서: ${editProfile.department})`
                          : " (부서 미설정 시 전체 허용)"}
                      </div>
                    )}
                    {editPermType === "custom" && catCheckboxes(editPerms, setEditPerms)}
                  </div>
                  {/* 시재 관리 접근 권한 */}
                  <div>
                    <div style={{ fontSize: 11, fontWeight: 700, color: "#94a3b8", marginBottom: 8 }}>시재 관리 접근 권한</div>
                    <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                      {[
                        [null, "접근 없음"],
                        ["sales", "영업부"],
                        ["finance", "관리부"],
                        ["admin", "관리자"],
                      ].map(([val, label]) => (
                        <button key={String(val)} onClick={() => setEditCashRole(val)} style={{
                          padding: "6px 14px", borderRadius: 8, border: "none", cursor: "pointer",
                          fontFamily: FF, fontWeight: 700, fontSize: 12,
                          background: editCashRole === val ? "#2563eb" : "#f1f5f9",
                          color: editCashRole === val ? "#fff" : "#64748b",
                        }}>{label}</button>
                      ))}
                    </div>
                  </div>
                  <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
                    <button onClick={() => setEditingEmail(null)} style={{ padding: "8px 16px", borderRadius: 8, border: "1.5px solid #e2e8f0", background: "#fff", fontWeight: 700, cursor: "pointer", fontFamily: FF, fontSize: 13 }}>취소</button>
                    <button onClick={() => saveUser(u.email)} disabled={saving} style={{ padding: "8px 16px", borderRadius: 8, border: "none", background: "#2563eb", color: "#fff", fontWeight: 800, cursor: "pointer", fontFamily: FF, fontSize: 13 }}>
                      {saving ? "저장 중..." : "저장"}
                    </button>
                  </div>
                </div>
              ) : (
                /* 표시 모드 */
                <div style={{ padding: "12px 14px", display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ fontWeight: 700, fontSize: 14, color: "#0f172a" }}>
                      {u.name || <span style={{ color: "#cbd5e1" }}>(이름 미입력)</span>}
                      {u.position && <span style={{ fontSize: 12, color: "#94a3b8", marginLeft: 6 }}>{u.position}</span>}
                    </div>
                    <div style={{ fontSize: 12, color: "#64748b", marginTop: 2 }}>
                      {u.email}
                      {u.department && <span style={{ marginLeft: 8, color: "#7c3aed", fontWeight: 700 }}>{u.department}</span>}
                    </div>
                    <div style={{ marginTop: 5, display: "flex", gap: 6, flexWrap: "wrap" }}>
                      {permBadge(allUserPerms[emailToKey(u.email)] ?? null)}
                      {(() => {
                        const cr = (allCashAccess || {})[emailToKey(u.email)]?.role;
                        if (!cr) return null;
                        const roleLabel = { sales: "영업부", finance: "관리부", admin: "관리자" }[cr] || cr;
                        return <span style={{ fontSize: 12, color: "#0369a1", background: "#e0f2fe", border: "1px solid #bae6fd", borderRadius: 6, padding: "2px 8px", fontWeight: 700 }}>시재-{roleLabel}</span>;
                      })()}
                    </div>
                  </div>
                  <button onClick={() => startEditUser(u)} style={{ padding: "7px 12px", borderRadius: 8, border: "1.5px solid #3b82f6", background: "#eff6ff", color: "#2563eb", fontWeight: 700, cursor: "pointer", fontFamily: FF, fontSize: 13 }}>편집</button>
                  <button onClick={() => { if (window.confirm(`"${u.email}" 사용자를 삭제할까요?`)) onDeleteUser(u.email); }}
                    style={{ padding: "7px 12px", borderRadius: 8, border: "1.5px solid #fca5a5", background: "#fef2f2", color: "#dc2626", fontWeight: 700, cursor: "pointer", fontFamily: FF, fontSize: 13 }}>삭제</button>
                </div>
              )}
            </div>
          ))}
        </div>
      )}

      {/* ── 부서별 권한 탭 ── */}
      {activeTab === "depts" && (
        <div style={{ display: "grid", gap: 8 }}>
          {departments.length === 0 && (
            <div style={{ padding: "20px 0", textAlign: "center", color: "#94a3b8", fontSize: 13, lineHeight: 1.7 }}>
              사용자 프로필에 부서를 입력하면<br />여기서 부서별 카테고리 권한을 설정할 수 있습니다.
            </div>
          )}
          {departments.map(dept => (
            <div key={dept} style={{ border: "1px solid #f1f5f9", borderRadius: 12, background: "#fafafa", overflow: "hidden" }}>
              {editingDept === dept ? (
                <div style={{ padding: "14px 16px", display: "grid", gap: 12 }}>
                  <div style={{ fontWeight: 800, fontSize: 14, color: "#0f172a" }}>{dept}</div>
                  <div>
                    <div style={{ fontSize: 11, fontWeight: 700, color: "#94a3b8", marginBottom: 8 }}>카테고리 접근 권한</div>
                    <div style={{ display: "flex", gap: 8, marginBottom: 10 }}>
                      {["all", "custom"].map(t => (
                        <button key={t} onClick={() => setEditDeptPermType(t)} style={{
                          padding: "6px 14px", borderRadius: 8, border: "none", cursor: "pointer",
                          fontFamily: FF, fontWeight: 700, fontSize: 12,
                          background: editDeptPermType === t ? "#2563eb" : "#f1f5f9",
                          color: editDeptPermType === t ? "#fff" : "#64748b",
                        }}>{t === "all" ? "전체 허용" : "직접 설정"}</button>
                      ))}
                    </div>
                    {editDeptPermType === "custom" && catCheckboxes(editDeptPerms, setEditDeptPerms)}
                  </div>
                  <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
                    <button onClick={() => setEditingDept(null)} style={{ padding: "8px 16px", borderRadius: 8, border: "1.5px solid #e2e8f0", background: "#fff", fontWeight: 700, cursor: "pointer", fontFamily: FF, fontSize: 13 }}>취소</button>
                    <button onClick={() => saveDept(dept)} disabled={saving} style={{ padding: "8px 16px", borderRadius: 8, border: "none", background: "#2563eb", color: "#fff", fontWeight: 800, cursor: "pointer", fontFamily: FF, fontSize: 13 }}>
                      {saving ? "저장 중..." : "저장"}
                    </button>
                  </div>
                </div>
              ) : (
                <div style={{ padding: "12px 14px", display: "flex", alignItems: "center", gap: 10 }}>
                  <div style={{ flex: 1 }}>
                    <div style={{ fontWeight: 700, fontSize: 14, color: "#0f172a" }}>{dept}</div>
                    <div style={{ fontSize: 12, color: "#94a3b8", marginTop: 2 }}>{users.filter(u => u.department === dept).length}명</div>
                    <div style={{ marginTop: 5 }}>{permBadge(deptPerms[departmentToKey(dept)] ?? null)}</div>
                  </div>
                  <button onClick={() => startEditDept(dept)} style={{ padding: "7px 14px", borderRadius: 8, border: "1.5px solid #3b82f6", background: "#eff6ff", color: "#2563eb", fontWeight: 700, cursor: "pointer", fontFamily: FF, fontSize: 13 }}>권한 설정</button>
                </div>
              )}
            </div>
          ))}
          <div style={{ marginTop: 8, padding: "12px 14px", background: "#f8fafc", border: "1px solid #f1f5f9", borderRadius: 10, fontSize: 12, color: "#94a3b8", lineHeight: 1.6 }}>
            개인 권한이 설정된 사용자는 부서 권한보다 개인 권한이 우선 적용됩니다.
          </div>
        </div>
      )}

      {/* ── 카테고리 관리 탭 ── */}
      {activeTab === "cats" && (
        <CatTab categories={categories} onAdd={onAddCategory} onUpdate={onUpdateCategory} onDelete={onDeleteCategory} />
      )}

      {/* ── 관리자 목록 탭 ── */}
      {activeTab === "admins" && (
        <div style={{ display: "grid", gap: 8 }}>
          <div style={{ padding: "14px 16px", background: "#fff7ed", border: "1px solid #fdba74", borderRadius: 12, fontSize: 13, color: "#9a3412", marginBottom: 4, lineHeight: 1.7 }}>
            직원 계정 생성 · 비밀번호 초기화는 <strong>Firebase Authentication 콘솔</strong>에서 진행해주세요.
            <div style={{ marginTop: 10 }}>
              <a href="https://console.firebase.google.com/project/staff-directory-app-9e17b/authentication/users" target="_blank" rel="noopener noreferrer"
                style={{ display: "inline-flex", alignItems: "center", gap: 6, padding: "8px 14px", borderRadius: 10, background: "#f97316", color: "#fff", fontWeight: 800, fontSize: 13, textDecoration: "none" }}>
                🔗 Firebase 콘솔 바로가기
              </a>
            </div>
          </div>
          {admins.length === 0
            ? <div style={{ color: "#cbd5e1", fontSize: 14, padding: "12px 0" }}>등록된 관리자가 없습니다.</div>
            : admins.map(a => (
              <div key={a.uid} style={{ border: "1px solid #f1f5f9", borderRadius: 10, padding: "10px 14px", background: "#fafafa" }}>
                <div style={{ fontWeight: 700, fontSize: 14 }}>{a.email}</div>
                <div style={{ fontSize: 12, color: "#94a3b8", marginTop: 2 }}>
                  등록일 {a.createdAt ? new Date(a.createdAt).toLocaleString("ko-KR") : "-"}
                </div>
              </div>
            ))
          }
        </div>
      )}
    </Modal>
  );
}

/* ══════════════ 메인 App ══════════════ */

const INIT_FORM = { siteName: "", siteUrl: "", loginId: "", password: "", category: "기타", owner: "", phone: "", fax: "", note: "" };

const COL_VIS_DEFAULT = { siteUrl: true, loginId: true, password: true, owner: true, phone: true, fax: true };

export default function App() {
  const [user, setUser] = useState(null);
  const [isAdmin, setIsAdmin] = useState(false);
  const [authLoading, setAuthLoading] = useState(true);
  const [loginLoading, setLoginLoading] = useState(false);
  const [loginError, setLoginError] = useState("");

  const [accounts, setAccounts] = useState([]);
  const [logs, setLogs] = useState([]);
  const [admins, setAdmins] = useState([]);
  const [favoritesMap, setFavoritesMap] = useState({});

  // 사용자 프로필 & 권한
  const [users, setUsers] = useState([]);       // 전체 사용자 목록 (관리자용)
  const [userProfile, setUserProfile] = useState(null);     // 현재 로그인 사용자 프로필
  const [userPermissions, setUserPermissions] = useState(null); // 개인별 권한 (null=전체허용)
  const [allUserPerms, setAllUserPerms] = useState({});       // 전체 사용자 권한 (관리자용)
  const [deptPerms, setDeptPerms] = useState({});       // 부서별 권한

  const [view, setView] = useState("list");
  const [selected, setSelected] = useState(null);
  const [form, setForm] = useState(INIT_FORM);
  const [editId, setEditId] = useState(null);

  const [search, setSearch] = useState("");
  const [expandedCats, setExpandedCats] = useState({});
  const [toast, setToast] = useState(null);
  const [showLogs, setShowLogs] = useState(false);
  const [showUsers, setShowUsers] = useState(false);

  const fileRef = useRef(null);

  const [categories, setCategories] = useState([]);
  const [showCategories, setShowCategories] = useState(false);
  const [passwordHistory, setPasswordHistory] = useState([]);
  const [colVis, setColVis] = useState(COL_VIS_DEFAULT);

  // 시재 관리
  const [mainTab, setMainTab] = useState("accounts");
  const [myCashAccess, setMyCashAccess] = useState(null);   // 현재 사용자 시재 권한
  const [allCashAccess, setAllCashAccess] = useState({});   // 전체 사용자 시재 권한 (관리자용)
  const [contacts, setContacts] = useState([]);     // 직원 연락처
  const [notices, setNotices] = useState([]);     // 공지사항
  const [companyInfo, setCompanyInfo] = useState(null);   // 회사 정보

  function showToast(msg, type = "success") {
    setToast({ msg, type });
    window.clearTimeout(showToast._t);
    showToast._t = setTimeout(() => setToast(null), 3000);
  }

  /* ── Auth ── */
  useEffect(() => {
    return subscribeAuth(async (u) => {
      setUser(u || null);
      if (u) {
        if (ADMIN_EMAILS.includes(u.email)) await seedAdminProfile(u.uid, u.email);
        setIsAdmin(await checkIsAdmin(u.uid));
        try { await recordUserLogin(u.uid, u.email); } catch { }
      } else {
        setIsAdmin(false);
        setUserProfile(null);
        setUserPermissions(null);
      }
      setAuthLoading(false);
    });
  }, []);

  /* ── 데이터 구독 ── */
  useEffect(() => {
    if (!user) return;
    const u1 = subscribeAccounts(list =>
      setAccounts([...list].sort((a, b) => new Date(b.updatedAt || 0) - new Date(a.updatedAt || 0)))
    );
    const u2 = subscribeLogs(setLogs);
    const u3 = subscribeFavorites(user.uid, setFavoritesMap);
    const u4 = subscribeAdmins(setAdmins);
    const u5 = subscribeCategories(async (list) => {
      if (list.length === 0) {
        // Firebase에 카테고리가 없으면 CAT_STYLE 기본값 시딩
        const defaults = Object.keys(CAT_STYLE);
        for (const name of defaults) { try { await addCategory(name); } catch { } }
      } else {
        setCategories(list);
      }
    });
    // 현재 사용자 프로필: 이메일 키로 직접 구독 (전체 목록 읽기 권한 없어도 동작)
    const u6 = subscribeUserProfile(user.email, setUserProfile);
    // 관리자용 전체 사용자 목록
    const u6b = subscribeUserProfiles(setUsers);
    const u7 = subscribeUserPermissions(user.email, setUserPermissions);
    // 관리자용: 전체 권한 목록
    const u8 = subscribeAllUserPermissions(setAllUserPerms);
    const u9 = subscribeDeptPermissions(setDeptPerms);
    const u10 = subscribeUserColVis(user.email, (vis) => {
      setColVis(vis ? { ...COL_VIS_DEFAULT, ...vis } : COL_VIS_DEFAULT);
    });
    const uC1 = subscribeCashAccess(user.email, setMyCashAccess);
    const uC2 = subscribeAllCashAccess(setAllCashAccess);
    const uK1 = subscribeContacts(setContacts);
    const uN1 = subscribeNotices(setNotices);
    const uN2 = subscribeCompanyInfo(setCompanyInfo);
    return () => { u1(); u2(); u3(); u4(); u5(); u6(); u6b(); u7(); u8(); u9(); u10(); uC1(); uC2(); uK1(); uN1(); uN2(); };
  }, [user]);

  /* ── 비밀번호 이력 구독 ── */
  useEffect(() => {
    if (!selected?.id) { setPasswordHistory([]); return; }
    const unsubscribe = subscribePasswordHistory(selected.id, setPasswordHistory);
    return unsubscribe;
  }, [selected?.id]);

  /* ── 유효 권한 계산: 개인 권한 > 부서 권한 > 전체 허용 ── */
  const effectivePermissions = useMemo(() => {
    if (isAdmin) return null;
    if (userPermissions === "__ALL__") return null;  // 개인 전체 허용
    if (Array.isArray(userPermissions)) return userPermissions; // 개인 직접 설정

    // 개인 설정 없음 → 부서 권한
    const dept = String(userProfile?.department || "").trim();
    const deptKey = departmentToKey(dept);

    if (deptKey && deptPerms[deptKey]) return deptPerms[deptKey];

    // 혹시 deptPerms가 원래 부서명으로 들어오는 경우도 대비
    if (dept && deptPerms[dept]) return deptPerms[dept];

    return null;
  }, [isAdmin, userPermissions, userProfile, deptPerms]);

  /* ── 권한 적용된 계정/카테고리 ── */
  const visibleAccounts = useMemo(() => {
    let list = effectivePermissions ? accounts.filter(a => effectivePermissions.includes(a.category)) : accounts;
    if (!isAdmin) list = list.filter(a => !a.locked);  // 비관리자는 잠금 계정 제외
    return list;
  }, [accounts, effectivePermissions, isAdmin]);

  const visibleCategories = useMemo(() => {
    if (!effectivePermissions) return categories;
    return categories.filter(c => effectivePermissions.includes(c.name));
  }, [categories, effectivePermissions]);

  /* ── 검색 필터 ── */
  const filtered = useMemo(() => {
    const q = search.trim().toLowerCase();
    if (!q) return visibleAccounts;
    return visibleAccounts.filter(a =>
      [a.siteName, a.loginId, a.owner, a.note, a.category].join(" ").toLowerCase().includes(q)
    );
  }, [visibleAccounts, search]);

  // Firebase 카테고리 + 계정에 실제 사용된 카테고리 병합 (카테고리 미설정시에도 표시)
  const displayCategories = useMemo(() => {
    const known = new Set(visibleCategories.map(c => c.name));
    const extra = [...new Set(visibleAccounts.map(a => a.category).filter(c => c && !known.has(c)))];
    return [...visibleCategories.map(c => c.name), ...extra];
  }, [visibleCategories, visibleAccounts]);

  const grouped = useMemo(() => {
    const g = {};
    displayCategories.forEach(name => { g[name] = filtered.filter(i => i.category === name); });
    return g;
  }, [filtered, displayCategories]);

  /* ── 즐겨찾기 목록 ── */
  const favoriteItems = useMemo(() =>
    visibleAccounts.filter(a => favoritesMap[a.id]), [visibleAccounts, favoritesMap]);

  /* ── 최근 많이 본 계정 (내 기준 상위 3) ── */
  const recentItems = useMemo(() => {
    if (!user) return [];
    const countMap = {};
    for (const log of logs) {
      if (log.email !== user.email) continue;
      if (!["상세 열람", "비밀번호 열람"].includes(log.action)) continue;
      if (log.targetId) countMap[log.targetId] = (countMap[log.targetId] || 0) + 1;
    }
    return Object.entries(countMap)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 3)
      .map(([id]) => visibleAccounts.find(a => a.id === id))
      .filter(Boolean);
  }, [logs, visibleAccounts, user]);

  /* ── 로그인/로그아웃 ── */
  async function handleLogin(companyId, password) {
    if (!companyId.trim()) { setLoginError("회사 아이디를 입력하세요."); return; }
    try {
      setLoginLoading(true); setLoginError("");
      const result = await loginWithCompanyId(companyId, password);
      await writeLog({ action: "로그인", email: result.user.email });
    } catch {
      setLoginError("아이디 또는 비밀번호를 확인하세요.");
    } finally {
      setLoginLoading(false);
    }
  }

  async function handleLogout() {
    try { await writeLog({ action: "로그아웃", email: user.email }); } catch { }
    await logout(); setSelected(null); setView("list");
  }

  /* ── 즐겨찾기 토글 ── */
  async function toggleFavorite(item) {
    const next = !favoritesMap[item.id];
    try {
      await setFavorite(user.uid, item.id, next);
      await writeLog({ action: next ? "즐겨찾기 추가" : "즐겨찾기 해제", email: user.email, targetId: item.id, targetName: item.siteName });
      showToast(next ? "즐겨찾기에 추가했습니다." : "즐겨찾기에서 제거했습니다.");
    } catch (e) {
      console.error("favorite error:", e);
      showToast("즐겨찾기 처리 중 오류가 발생했습니다.", "error");
    }
  }

  /* ── 잠금 토글 ── */
  async function toggleLock(item) {
    const next = !item.locked;
    try {
      await updateAccount(item.id, { ...item, locked: next });
      await writeLog({ action: next ? "계정 잠금" : "계정 잠금 해제", email: user.email, targetId: item.id, targetName: item.siteName });
      showToast(next ? "잠금 설정되었습니다. (관리자만 열람)" : "잠금이 해제되었습니다.");
    } catch {
      showToast("처리 중 오류가 발생했습니다.", "error");
    }
  }

  /* ── 상세 열기 ── */
  async function openDetail(item) {
    setSelected(item); setView("detail");
    await writeLog({ action: "상세 열람", email: user.email, targetId: item.id, targetName: item.siteName });
  }

  /* ── 계정 저장 ── */
  async function saveItem() {
    if (!form.siteName.trim()) { showToast("사이트명을 입력하세요.", "error"); return; }
    try {
      if (editId) {
        // 기존 계정 수정 시 비밀번호 변경 확인
        const existingItems = accounts.filter(a => a.id === editId);
        const existingItem = existingItems.length > 0 ? existingItems[0] : null;
        const passwordChanged = existingItem && form.password !== existingItem.password && form.password;

        await updateAccount(editId, form);

        // 비밀번호가 변경되었으면 이력 기록
        if (passwordChanged) {
          await addPasswordHistory(editId, {
            password: form.password,
            changedBy: user.email
          });
        }

        await writeLog({ action: "계정 수정", email: user.email, targetId: editId, targetName: form.siteName });
        showToast("수정되었습니다.");
      } else {
        const id = await addAccount(form);
        await writeLog({ action: "계정 추가", email: user.email, targetId: id, targetName: form.siteName });
        showToast("등록되었습니다.");
      }
      setForm(INIT_FORM); setEditId(null); setView("list");
    } catch {
      showToast("저장 중 오류가 발생했습니다.", "error");
    }
  }

  /* ── 계정 삭제 ── */
  async function removeItem(item) {
    if (!window.confirm(`"${item.siteName}" 계정을 삭제할까요?`)) return;
    await deleteAccount(item.id);
    await writeLog({ action: "계정 삭제", email: user.email, targetId: item.id, targetName: item.siteName });
    showToast("삭제되었습니다."); setSelected(null); setView("list");
  }

  async function handleAddNotice(data) {
    await addNotice(data);
    await writeLog({ action: "공지사항 등록", email: user.email, targetName: data?.title || "" });
  }

  async function handleUpdateNotice(id, data) {
    await updateNotice(id, data);
    await writeLog({ action: "공지사항 수정", email: user.email, targetId: id, targetName: data?.title || "" });
  }

  async function handleDeleteNotice(id) {
    const target = notices.find((n) => n.id === id);
    await deleteNotice(id);
    await writeLog({ action: "공지사항 삭제", email: user.email, targetId: id, targetName: target?.title || "" });
  }

  async function handleSaveCompanyInfo(data) {
    await saveCompanyInfo(data);
    await writeLog({ action: "회사 정보 수정", email: user.email, targetName: data?.name || "회사 정보" });
  }

  async function handleUploadCompanyAttachment(file, category) {
    await uploadCompanyAttachment(file, category, user?.email || "");
    await writeLog({ action: "회사 파일 첨부", email: user.email, targetName: file?.name || "" });
  }

  async function handleDeleteCompanyAttachment(itemId) {
    await deleteCompanyAttachment(itemId);
    await writeLog({ action: "회사 파일 삭제", email: user.email, targetId: itemId, targetName: "회사 첨부파일" });
  }

  /* ── 전체 삭제 ── */
  async function handleDeleteAll() {
    if (!window.confirm(`전체 계정 ${accounts.length}개를 모두 삭제합니다. 계속하시겠습니까?`)) return;
    try {
      await deleteAllAccounts();
      await writeLog({ action: "전체 삭제", email: user.email, targetName: `${accounts.length}건` });
      showToast("전체 계정이 삭제되었습니다.");
    } catch (err) {
      console.error("전체 삭제 오류:", err);
      showToast("삭제 중 오류가 발생했습니다.", "error");
    }
  }

  /* ── 엑셀 업로드 ── */
  function handleExcel(e) {
    console.log("handleExcel 실행됨");
    const file = e.target.files?.[0];
    if (!file) return;
    console.log("파일:", file.name);
    const reader = new FileReader();
    reader.onload = async ev => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "binary" });
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "" });
        const existingCatNames = categories.map(c => c.name);
        // 새 카테고리 자동 추가
        const newCats = [...new Set(rows.map(r => String(r["카테고리"] || "").trim()).filter(c => c && !existingCatNames.includes(c)))];
        for (const catName of newCats) await addCategory(catName);
        const allCatNames = [...existingCatNames, ...newCats];
        const existingMap = new Map(accounts.map(a => [`${a.siteName}__${a.loginId}`, a.id]));
        const allItems = rows.filter(r => r["사이트명"]).map(r => ({
          siteName: String(r["사이트명"] || ""), siteUrl: String(r["사이트URL"] || ""),
          loginId: String(r["아이디"] || ""), password: String(r["비밀번호"] || ""),
          category: allCatNames.includes(String(r["카테고리"] || "").trim()) ? String(r["카테고리"]).trim() : "기타",
          owner: String(r["담당자"] || ""), note: String(r["비고"] || ""),
        }));
        if (!allItems.length) { showToast("등록할 데이터가 없습니다.", "error"); return; }
        let added = 0, updated = 0;
        await Promise.all(allItems.map(item => {
          const key = `${item.siteName}__${item.loginId}`;
          const existingId = existingMap.get(key);
          if (existingId) { updated++; return updateAccount(existingId, item); }
          else { added++; return addAccount(item); }
        }));
        await writeLog({ action: "엑셀 업로드", email: user.email, targetName: `추가 ${added}건, 수정 ${updated}건` });
        const newCatMsg = newCats.length ? ` (새 카테고리 ${newCats.length}개 추가)` : "";
        showToast(`추가 ${added}건, 수정 ${updated}건 반영되었습니다.${newCatMsg}`);
      } catch (err) {
        console.error("엑셀 업로드 오류:", err);
        showToast("엑셀 처리 중 오류가 발생했습니다.", "error");
      } finally { e.target.value = ""; }
    };
    reader.readAsBinaryString(file);
  }

  /* ── 렌더링 ── */
  if (authLoading) return (
    <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: FF, color: "#94a3b8" }}>
      불러오는 중...
    </div>
  );

  if (!user) return <><LoginScreen onLogin={handleLogin} loading={loginLoading} error={loginError} /><Toast toast={toast} /></>;

  return (
    <div style={{ minHeight: "100vh", background: "#f8fafc", fontFamily: FF, color: "#111827" }}>
      <Toast toast={toast} />
      <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={handleExcel} />

      {showCategories && isAdmin && (
        <CategoriesModal categories={categories} onClose={() => setShowCategories(false)} />
      )}
      {showLogs && isAdmin && (
        <LogModal logs={logs} onClose={() => setShowLogs(false)}
          onDownload={() => downloadCsv(`업무계정로그_${new Date().toISOString().slice(0, 10)}.csv`, [
            ["시간", "이메일", "액션", "대상"],
            ...logs.map(l => [l.createdAt ? new Date(l.createdAt).toLocaleString("ko-KR") : "", l.email || "", l.action || "", l.targetName || ""])
          ])} />
      )}
      {showUsers && isAdmin && (
        <UsersModal
          admins={admins}
          users={users}
          allUserPerms={allUserPerms}
          deptPerms={deptPerms}
          categories={categories}
          allCashAccess={allCashAccess}
          onAddUser={addUserProfile}
          onDeleteUser={deleteUserProfile}
          onSaveUserProfile={updateUserProfile}
          onSaveUserPerm={saveUserPermissions}
          onSaveDeptPerm={setDeptPermissions}
          onBulkAddUsers={addBulkUserProfiles}
          onSaveCashAccess={saveCashAccess}
          onAddCategory={addCategory}
          onUpdateCategory={updateCategory}
          onDeleteCategory={deleteCategory}
          onClose={() => setShowUsers(false)}
        />
      )}

      {/* ── 통합 헤더 (탭 포함) ── */}
      <AppHeader
        user={user} userProfile={userProfile} isAdmin={isAdmin} onLogout={handleLogout}
        detailMode={mainTab === "accounts" && (view === "detail" || view === "form")}
        onBack={() => { setView("list"); setSelected(null); setForm(INIT_FORM); setEditId(null); }}
        mainTab={mainTab} onTabChange={setMainTab}
        hasCashAccess={!!(isAdmin || myCashAccess?.role)}
        onAdd={() => { setForm(INIT_FORM); setEditId(null); setView("form"); }}
        onImportExcel={() => fileRef.current?.click()}
        onDownloadAccounts={() => downloadAccountsExcel(accounts)}
        onDownloadTemplate={downloadTemplate}
        onDeleteAll={handleDeleteAll}
        onShowLogs={() => setShowLogs(true)}
        onShowUsers={() => setShowUsers(true)}
        onShowCategories={() => setShowCategories(true)}
      />

      {/* ── 직원 연락처 ── */}
      {mainTab === "contacts" && (
        <ScheduledContactsView contacts={contacts} isAdmin={isAdmin} showToast={showToast} userProfile={userProfile} user={user} />
      )}

      {/* ── 시재 관리 ── */}
      {mainTab === "cash" && (() => {
        const effectiveCashRole = isAdmin ? "admin" : myCashAccess?.role || null;
        return effectiveCashRole
          ? <CashManager user={user} cashRole={effectiveCashRole} showToast={showToast} isAdmin={isAdmin} />
          : null;
      })()}

      {mainTab === "notice" && (
        <NoticeBoardView
          notices={notices}
          isAdmin={isAdmin}
          user={user}
          onAddNotice={handleAddNotice}
          onUpdateNotice={handleUpdateNotice}
          onDeleteNotice={handleDeleteNotice}
          showToast={showToast}
        />
      )}

      {mainTab === "company" && (
        <CompanyInfoView
          companyInfo={companyInfo}
          isAdmin={isAdmin}
          user={user}
          onSaveCompanyInfo={handleSaveCompanyInfo}
          onUploadAttachment={handleUploadCompanyAttachment}
          onDeleteAttachment={handleDeleteCompanyAttachment}
          showToast={showToast}
        />
      )}

      {/* ── 업무계정 관리 ── */}
      {mainTab === "accounts" && view === "form" && (
        <AccountForm
          form={form} setForm={setForm} editMode={!!editId} onSave={saveItem}
          categories={categories}
          onCancel={() => { setView("list"); setForm(INIT_FORM); setEditId(null); }}
        />
      )}

      {mainTab === "accounts" && view === "detail" && selected && (
        <DetailView item={selected} user={user} isAdmin={isAdmin}
          isFavorite={!!favoritesMap[selected.id]}
          passwordHistory={passwordHistory}
          onToggleFavorite={() => toggleFavorite(selected)}
          onBack={() => { setSelected(null); setView("list"); }}
          showToast={showToast}
          colVis={colVis}
          onToggleCol={(key) => {
            const next = { ...colVis, [key]: !colVis[key] };
            setColVis(next);
            saveUserColVis(user.email, next);
          }}
          onEdit={() => {
            setForm({ siteName: selected.siteName || "", siteUrl: selected.siteUrl || "", loginId: selected.loginId || "", password: selected.password || "", category: selected.category || "기타", owner: selected.owner || "", phone: selected.phone || "", fax: selected.fax || "", note: selected.note || "" });
            setEditId(selected.id); setView("form");
          }}
          onDelete={() => removeItem(selected)} />
      )}

      {mainTab === "accounts" && view === "list" && (
        <div style={{ maxWidth: 960, margin: "0 auto", padding: "20px 16px 40px" }}>

          {/* 검색창 */}
          <div style={{
            background: "#fff", border: "1.5px solid #f1f5f9", borderRadius: 16,
            padding: "4px 14px", marginBottom: 20,
            boxShadow: "0 2px 8px rgba(15,23,42,0.05)",
            display: "flex", alignItems: "center", gap: 10,
          }}>
            <span style={{ color: "#cbd5e1", fontSize: 16 }}>🔍</span>
            <input value={search} onChange={e => setSearch(e.target.value)}
              placeholder="사이트명, 아이디, 담당자로 검색"
              style={{ ...base.input, border: "none", padding: "10px 4px", boxShadow: "none" }} />
            {search && (
              <button onClick={() => setSearch("")}
                style={{ background: "none", border: "none", cursor: "pointer", color: "#cbd5e1", fontSize: 16, padding: "0 4px" }}>✕</button>
            )}
          </div>

          <div style={{ fontSize: 13, color: "#cbd5e1", fontWeight: 600, marginBottom: 8 }}>
            총 {visibleAccounts.length}개 계정{search ? ` · 검색 결과 ${filtered.length}개` : ""}
          </div>

          {/* 즐겨찾기 */}
          {!search && <MiniSection title="⭐ 즐겨찾기" items={favoriteItems} favoritesMap={favoritesMap} onOpen={openDetail} onToggleFavorite={toggleFavorite} onToggleLock={toggleLock} isAdmin={isAdmin} />}

          {/* 많이 본 계정 */}
          {!search && recentItems.length > 0 && <MiniSection title="🕐 자주 본 계정" items={recentItems} favoritesMap={favoritesMap} onOpen={openDetail} onToggleFavorite={toggleFavorite} onToggleLock={toggleLock} isAdmin={isAdmin} />}

          {/* 카테고리별 */}
          {!search && <Divider label="전체 계정" />}
          {displayCategories.filter(name => grouped[name]?.length > 0).map(name => (
            <div key={name} style={{ marginBottom: 4 }}>
              <CategorySection
                title={name}
                items={grouped[name] || []}
                expanded={!!expandedCats[name]}
                onToggle={() => setExpandedCats(prev => ({ ...prev, [name]: !prev[name] }))}
                favoritesMap={favoritesMap}
                onOpen={openDetail}
                onToggleFavorite={toggleFavorite}
                onToggleLock={toggleLock}
                isAdmin={isAdmin}
              />
            </div>
          ))}

          {filtered.length === 0 && (
            <div style={{ textAlign: "center", padding: "60px 0", color: "#cbd5e1" }}>
              <div style={{ fontSize: 40, marginBottom: 12 }}>🔍</div>
              <div style={{ fontSize: 15 }}>{search ? "검색 결과가 없습니다." : "등록된 계정이 없습니다."}</div>
            </div>
          )}
        </div>
      )}
    </div>
  );
}
