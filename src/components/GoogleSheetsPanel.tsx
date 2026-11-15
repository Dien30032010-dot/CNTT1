import { useCallback, useEffect, useMemo, useState } from "react";
import type { User } from "firebase/auth";
import {
  AlertCircle,
  CheckCircle2,
  ExternalLink,
  FilePlus2,
  FileSpreadsheet,
  FolderOpen,
  LogOut,
  RefreshCw,
  Table2,
  UploadCloud,
  UserPlus,
} from "lucide-react";
import {
  CLASS_SHEET_TABS,
  createClassSpreadsheet,
  extractSpreadsheetId,
  getAccessToken,
  getSpreadsheetMetadata,
  googleSignIn,
  googleSignOut,
  initGoogleAuth,
  listSpreadsheets,
  readSheetValues,
  syncDatabaseToSpreadsheet,
  type DriveSpreadsheetFile,
  type SpreadsheetMetadata,
} from "../lib/googleWorkspace";
import type { ClassDatabase, GradeRecord, SessionState, Student } from "../types";
import { GoogleSignInButton } from "./GoogleSignInButton";
import type { ModalConfig } from "./ModalDialog";

interface GoogleSheetsPanelProps {
  db: ClassDatabase;
  session: SessionState;
  onUpdateDb: (updater: (prev: ClassDatabase) => ClassDatabase) => void;
  onToast: (msg: string) => void;
  onOpenModal: (modal: ModalConfig) => void;
}

function normalizeHeader(cell: string): string {
  return cell
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

function parseStudentsFromRows(
  rows: string[][],
  existingStudents: Student[]
): { parsed: Student[]; skippedDuplicates: number; invalidCount: number } {
  if (rows.length === 0) {
    return { parsed: [], skippedDuplicates: 0, invalidCount: 0 };
  }

  const firstRow = rows[0].map(normalizeHeader);
  const idIdxByHeader = firstRow.findIndex(
    (h) =>
      h.includes("ma sinh vien") ||
      h.includes("ma sv") ||
      h === "masv" ||
      h === "msv" ||
      h === "id" ||
      h === "sid"
  );
  const nameIdxByHeader = firstRow.findIndex(
    (h) =>
      h.includes("ho va ten") ||
      h.includes("ho ten") ||
      h === "ten" ||
      h === "name" ||
      h.includes("sinh vien")
  );

  const hasHeader = idIdxByHeader !== -1 || nameIdxByHeader !== -1;
  const dataRows = hasHeader ? rows.slice(1) : rows;

  const idIdx =
    idIdxByHeader !== -1
      ? idIdxByHeader
      : rows[0].length >= 3 && /^\d+$/.test((rows[1]?.[0] || "").trim())
      ? 1
      : 0;
  const nameIdx =
    nameIdxByHeader !== -1
      ? nameIdxByHeader
      : idIdx === 1 && rows[0].length >= 3
      ? 2
      : 1;

  const existingIds = new Set(existingStudents.map((s) => s.id.toUpperCase()));
  const seenInFile = new Set<string>();
  const parsed: Student[] = [];
  let skippedDuplicates = 0;
  let invalidCount = 0;

  for (const row of dataRows) {
    const rawId = (row[idIdx] || "").trim();
    const rawName = (row[nameIdx] || "").trim();
    if (!rawId && !rawName) continue;
    if (!rawId || !rawName) {
      invalidCount++;
      continue;
    }
    const upperId = rawId.toUpperCase();
    if (existingIds.has(upperId) || seenInFile.has(upperId)) {
      skippedDuplicates++;
      continue;
    }
    seenInFile.add(upperId);
    parsed.push({ id: rawId, name: rawName });
  }

  return { parsed, skippedDuplicates, invalidCount };
}

function parseGradesFromRows(
  rows: string[][],
  students: Student[],
  session: SessionState
): { parsed: GradeRecord[]; invalidCount: number } {
  if (rows.length === 0) {
    return { parsed: [], invalidCount: 0 };
  }

  const firstRow = rows[0].map(normalizeHeader);
  const sidIdx = firstRow.findIndex(
    (h) =>
      h.includes("ma sinh vien") ||
      h.includes("ma sv") ||
      h === "masv" ||
      h === "msv" ||
      h === "sid" ||
      h === "id"
  );
  const subjectIdx = firstRow.findIndex(
    (h) => h.includes("mon hoc") || h === "mon" || h === "subject"
  );
  const scoreIdx = firstRow.findIndex(
    (h) =>
      h.includes("diem") || h === "score" || h === "value" || h === "grade"
  );
  const termIdx = firstRow.findIndex(
    (h) => h.includes("hoc ky") || h.includes("dot") || h === "term"
  );
  const noteIdx = firstRow.findIndex(
    (h) => h.includes("ghi chu") || h === "note"
  );

  const hasHeader = sidIdx !== -1 || subjectIdx !== -1 || scoreIdx !== -1;
  const dataRows = hasHeader ? rows.slice(1) : rows;

  const colSid = sidIdx !== -1 ? sidIdx : 0;
  const colSubject = subjectIdx !== -1 ? subjectIdx : 1;
  const colScore = scoreIdx !== -1 ? scoreIdx : 2;

  const studentIds = new Set(students.map((s) => s.id.toUpperCase()));
  const studentCanonicalId = new Map(
    students.map((s) => [s.id.toUpperCase(), s.id])
  );

  const parsed: GradeRecord[] = [];
  let invalidCount = 0;
  const now = new Date().toISOString();
  const byName =
    session.role === "teacher"
      ? session.currentTeacher || "Giáo viên"
      : session.role === "admin"
      ? "Quản lý"
      : session.currentOfficer || "Ban cán sự";

  for (const row of dataRows) {
    const rawSid = (row[colSid] || "").trim();
    const rawSubject = (row[colSubject] || "").trim();
    const rawScore = (row[colScore] || "").trim().replace(",", ".");
    if (!rawSid && !rawSubject && !rawScore) continue;

    const canonicalSid = studentCanonicalId.get(rawSid.toUpperCase());
    const numScore = Number(rawScore);
    if (
      !canonicalSid ||
      !studentIds.has(rawSid.toUpperCase()) ||
      !rawSubject ||
      rawScore === "" ||
      !Number.isFinite(numScore) ||
      numScore < 0 ||
      numScore > 10
    ) {
      invalidCount++;
      continue;
    }

    parsed.push({
      id: crypto.randomUUID(),
      sid: canonicalSid,
      subject: rawSubject,
      value: Math.round(numScore * 100) / 100,
      term: termIdx !== -1 ? (row[termIdx] || "").trim() : "",
      note: noteIdx !== -1 ? (row[noteIdx] || "").trim() : "Nhập từ Google Sheets",
      by: session.role,
      byName,
      at: now,
    });
  }

  return { parsed, invalidCount };
}

export function GoogleSheetsPanel({
  db,
  session,
  onUpdateDb,
  onToast,
  onOpenModal,
}: GoogleSheetsPanelProps) {
  const isAdmin = session.role === "admin";
  const canImportGrades = isAdmin || session.role === "teacher";

  const [googleUser, setGoogleUser] = useState<User | null>(null);
  const [needsAuth, setNeedsAuth] = useState(true);
  const [isSigningIn, setIsSigningIn] = useState(false);

  const [sheetInput, setSheetInput] = useState(db.linkedSpreadsheetId || "");
  const [driveFiles, setDriveFiles] = useState<DriveSpreadsheetFile[]>([]);
  const [loadingDrive, setLoadingDrive] = useState(false);

  const [meta, setMeta] = useState<SpreadsheetMetadata | null>(null);
  const [loadingMeta, setLoadingMeta] = useState(false);
  const [selectedTab, setSelectedTab] = useState("");
  const [sheetRows, setSheetRows] = useState<string[][]>([]);
  const [loadingRows, setLoadingRows] = useState(false);
  const [syncing, setSyncing] = useState(false);
  const [creating, setCreating] = useState(false);
  const [errorMsg, setErrorMsg] = useState("");

  useEffect(() => {
    const unsubscribe = initGoogleAuth(
      (user) => {
        setGoogleUser(user);
        setNeedsAuth(false);
        setErrorMsg("");
      },
      () => {
        setGoogleUser(null);
        setNeedsAuth(true);
      }
    );
    return unsubscribe;
  }, []);

  const requireToken = useCallback((): string | null => {
    const token = getAccessToken();
    if (!token) {
      setNeedsAuth(true);
      setErrorMsg("Phiên Google đã hết hạn. Vui lòng đăng nhập lại Google.");
      return null;
    }
    return token;
  }, []);

  const loadTabValues = useCallback(
    async (spreadsheetId: string, tabTitle: string) => {
      const token = requireToken();
      if (!token || !tabTitle) return;
      setLoadingRows(true);
      setErrorMsg("");
      try {
        const rows = await readSheetValues(token, spreadsheetId, tabTitle);
        setSheetRows(rows);
      } catch (err) {
        const msg =
          err instanceof Error ? err.message : "Không thể đọc dữ liệu tab.";
        setErrorMsg(msg);
        if (!getAccessToken()) setNeedsAuth(true);
      } finally {
        setLoadingRows(false);
      }
    },
    [requireToken]
  );

  const openSpreadsheetById = useCallback(
    async (rawIdOrUrl: string, persistLink = true) => {
      const token = requireToken();
      if (!token) return;
      const spreadsheetId = extractSpreadsheetId(rawIdOrUrl);
      if (!spreadsheetId) {
        onToast("Vui lòng nhập ID hoặc liên kết Google Sheets.");
        return;
      }

      setLoadingMeta(true);
      setErrorMsg("");
      try {
        const nextMeta = await getSpreadsheetMetadata(token, spreadsheetId);
        setMeta(nextMeta);
        setSheetInput(nextMeta.spreadsheetId);
        const firstTab = nextMeta.sheets[0]?.title || "";
        setSelectedTab(firstTab);
        if (persistLink) {
          onUpdateDb((prev) => ({
            ...prev,
            linkedSpreadsheetId: nextMeta.spreadsheetId,
            linkedSpreadsheetName: nextMeta.title,
          }));
        }
        if (firstTab) {
          await loadTabValues(nextMeta.spreadsheetId, firstTab);
        } else {
          setSheetRows([]);
        }
        onToast(`Đã kết nối bảng tính "${nextMeta.title}".`);
      } catch (err) {
        const msg =
          err instanceof Error
            ? err.message
            : "Không thể mở bảng tính Google Sheets.";
        setErrorMsg(msg);
        if (!getAccessToken()) setNeedsAuth(true);
      } finally {
        setLoadingMeta(false);
      }
    },
    [loadTabValues, onToast, onUpdateDb, requireToken]
  );

  const handleFetchDriveFiles = useCallback(async () => {
    const token = requireToken();
    if (!token) return;
    setLoadingDrive(true);
    setErrorMsg("");
    try {
      const files = await listSpreadsheets(token);
      setDriveFiles(files);
      if (files.length === 0) {
        onToast("Không tìm thấy tệp Google Sheets nào trên Drive.");
      }
    } catch (err) {
      const msg =
        err instanceof Error
          ? err.message
          : "Không thể tải danh sách tệp từ Google Drive.";
      setErrorMsg(msg);
      if (!getAccessToken()) setNeedsAuth(true);
    } finally {
      setLoadingDrive(false);
    }
  }, [onToast, requireToken]);

  // Automatically load Drive list & linked spreadsheet once signed in
  useEffect(() => {
    if (!needsAuth && googleUser && getAccessToken()) {
      handleFetchDriveFiles();
      if (db.linkedSpreadsheetId && !meta) {
        openSpreadsheetById(db.linkedSpreadsheetId, false);
      }
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [needsAuth, googleUser]);

  const handleGoogleSignIn = async () => {
    setIsSigningIn(true);
    setErrorMsg("");
    try {
      const { user } = await googleSignIn();
      setGoogleUser(user);
      setNeedsAuth(false);
      onToast("Đã kết nối tài khoản Google.");
    } catch (err) {
      const msg =
        err instanceof Error
          ? err.message
          : "Đăng nhập Google không thành công.";
      setErrorMsg(msg);
      onToast("Đăng nhập Google thất bại.");
    } finally {
      setIsSigningIn(false);
    }
  };

  const handleGoogleSignOut = async () => {
    await googleSignOut();
    setGoogleUser(null);
    setNeedsAuth(true);
    setDriveFiles([]);
    setMeta(null);
    setSheetRows([]);
    setSelectedTab("");
    onToast("Đã ngắt kết nối Google.");
  };

  const handleCreateSpreadsheet = () => {
    const token = requireToken();
    if (!token) return;
    const defaultTitle = `${db.className} - Quản lý lớp học`;
    onOpenModal({
      type: "confirm",
      title: "Tạo Google Sheet mới trên Drive",
      message: `Hệ thống sẽ tạo một bảng tính mới tên "${defaultTitle}" trên Google Drive của bạn với 6 trang: ${CLASS_SHEET_TABS.join(
        ", "
      )}. Tiếp tục?`,
      confirmLabel: "Tạo bảng tính",
      onConfirm: async () => {
        const freshToken = requireToken();
        if (!freshToken) return;
        setCreating(true);
        setErrorMsg("");
        try {
          const created = await createClassSpreadsheet(freshToken, defaultTitle);
          setMeta(created);
          setSheetInput(created.spreadsheetId);
          setSelectedTab(created.sheets[0]?.title || "");
          onUpdateDb((prev) => ({
            ...prev,
            linkedSpreadsheetId: created.spreadsheetId,
            linkedSpreadsheetName: created.title,
          }));
          onToast(`Đã tạo bảng tính "${created.title}".`);
          await handleFetchDriveFiles();
        } catch (err) {
          const msg =
            err instanceof Error
              ? err.message
              : "Không thể tạo bảng tính Google Sheets.";
          setErrorMsg(msg);
          if (!getAccessToken()) setNeedsAuth(true);
        } finally {
          setCreating(false);
        }
      },
    });
  };

  const handleSyncToSheet = () => {
    if (!meta) {
      onToast("Vui lòng chọn hoặc tạo bảng tính trước.");
      return;
    }
    const token = requireToken();
    if (!token) return;

    onOpenModal({
      type: "confirm",
      title: "Đồng bộ dữ liệu lên Google Sheets",
      message: `Hành động này sẽ xóa dữ liệu cũ và ghi đè toàn bộ 6 trang (${CLASS_SHEET_TABS.join(
        ", "
      )}) trong bảng tính "${meta.title}" bằng dữ liệu hiện tại của lớp (${
        db.students.length
      } sinh viên, ${db.grades.length} điểm, ${
        db.evaluations.length
      } đánh giá, ${db.violations.length} vi phạm, ${
        db.fundRecords.length
      } thu/chi, ${db.announcements.length} thông báo).`,
      confirmLabel: "Đồng bộ ngay",
      onConfirm: async () => {
        const freshToken = requireToken();
        if (!freshToken) return;
        setSyncing(true);
        setErrorMsg("");
        try {
          const updatedMeta = await syncDatabaseToSpreadsheet(
            freshToken,
            meta.spreadsheetId,
            db
          );
          setMeta(updatedMeta);
          const nextTab =
            selectedTab &&
            updatedMeta.sheets.some((s) => s.title === selectedTab)
              ? selectedTab
              : updatedMeta.sheets[0]?.title || "";
          setSelectedTab(nextTab);
          if (nextTab) {
            await loadTabValues(updatedMeta.spreadsheetId, nextTab);
          }
          onToast("Đã đồng bộ dữ liệu lớp lên Google Sheets.");
        } catch (err) {
          const msg =
            err instanceof Error
              ? err.message
              : "Đồng bộ Google Sheets thất bại.";
          setErrorMsg(msg);
          if (!getAccessToken()) setNeedsAuth(true);
        } finally {
          setSyncing(false);
        }
      },
    });
  };

  const handleImportStudents = () => {
    if (!isAdmin) return;
    const { parsed, skippedDuplicates, invalidCount } = parseStudentsFromRows(
      sheetRows,
      db.students
    );
    if (parsed.length === 0) {
      onToast(
        skippedDuplicates > 0
          ? `Tất cả ${skippedDuplicates} sinh viên trong tab đã tồn tại.`
          : "Không tìm thấy dòng sinh viên hợp lệ (cần cột Mã sinh viên và Họ và tên)."
      );
      return;
    }
    onOpenModal({
      type: "confirm",
      title: "Nhập danh sách sinh viên từ Google Sheet",
      message: `Sẽ thêm ${parsed.length} sinh viên mới từ tab "${selectedTab}" vào danh sách lớp.${
        skippedDuplicates > 0
          ? ` Bỏ qua ${skippedDuplicates} mã đã trùng.`
          : ""
      }${invalidCount > 0 ? ` Bỏ qua ${invalidCount} dòng thiếu dữ liệu.` : ""}`,
      confirmLabel: `Thêm ${parsed.length} sinh viên`,
      onConfirm: () => {
        onUpdateDb((prev) => ({
          ...prev,
          students: [...prev.students, ...parsed],
        }));
        onToast(`Đã nhập thêm ${parsed.length} sinh viên.`);
      },
    });
  };

  const handleImportGrades = () => {
    if (!canImportGrades) return;
    const { parsed, invalidCount } = parseGradesFromRows(
      sheetRows,
      db.students,
      session
    );
    if (parsed.length === 0) {
      onToast(
        "Không tìm thấy điểm hợp lệ (cần Mã sinh viên có trong lớp, Môn học, Điểm 0–10)."
      );
      return;
    }
    onOpenModal({
      type: "confirm",
      title: "Nhập điểm số từ Google Sheet",
      message: `Sẽ nhập ${parsed.length} bản ghi điểm hợp lệ từ tab "${selectedTab}" vào hệ thống.${
        invalidCount > 0
          ? ` Bỏ qua ${invalidCount} dòng không hợp lệ hoặc sai mã sinh viên.`
          : ""
      }`,
      confirmLabel: `Nhập ${parsed.length} điểm`,
      onConfirm: () => {
        onUpdateDb((prev) => ({
          ...prev,
          grades: [...prev.grades, ...parsed],
        }));
        onToast(`Đã nhập ${parsed.length} bản ghi điểm.`);
      },
    });
  };

  const previewRows = useMemo(() => sheetRows.slice(0, 50), [sheetRows]);
  const maxCols = useMemo(
    () =>
      previewRows.reduce((max, r) => (r.length > max ? r.length : max), 0),
    [previewRows]
  );

  const cardClass =
    "bg-[var(--panel)] border border-[var(--line)] rounded-2xl p-5 shadow-xs";
  const inputClass =
    "w-full px-3 py-2.5 rounded-xl border border-[var(--input-border)] bg-[var(--input-bg)] text-[var(--ink)] text-sm focus:outline-none focus:border-[var(--primary)] focus:ring-2 focus:ring-[var(--primary)]/20";
  const btnPrimary =
    "inline-flex items-center justify-center gap-2 px-4 py-2.5 rounded-xl bg-[var(--primary)] hover:bg-[var(--primary-hover)] text-white text-sm font-semibold transition-colors cursor-pointer whitespace-nowrap disabled:opacity-60 disabled:cursor-not-allowed";
  const btnSecondary =
    "inline-flex items-center justify-center gap-1.5 px-3 py-2 rounded-xl bg-[var(--primary-soft)] text-[var(--primary-soft-ink)] hover:opacity-90 text-xs font-semibold transition-opacity cursor-pointer whitespace-nowrap disabled:opacity-60 disabled:cursor-not-allowed";

  return (
    <div className="space-y-5">
      <div className="flex flex-wrap items-center justify-between gap-3">
        <div>
          <h2 className="text-2xl font-bold tracking-tight">
            Kết nối Google Sheets
          </h2>
          <p className="text-sm text-[var(--muted)] mt-0.5">
            Đồng bộ dữ liệu lớp {db.className} lên Google Sheets hoặc đọc dữ liệu từ bảng tính trên Google Drive.
          </p>
        </div>
      </div>

      {errorMsg && (
        <div className="flex items-start gap-2.5 p-3.5 rounded-xl bg-rose-50 dark:bg-rose-950/40 text-rose-800 dark:text-rose-200 border border-rose-200 dark:border-rose-800/70 text-sm">
          <AlertCircle className="w-4 h-4 shrink-0 mt-0.5" />
          <span>{errorMsg}</span>
        </div>
      )}

      {needsAuth || !googleUser ? (
        <div className={cardClass}>
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-5">
            <div className="space-y-2 max-w-2xl">
              <div className="flex items-center gap-2 text-[var(--primary)]">
                <FileSpreadsheet className="w-5 h-5" />
                <h3 className="text-base font-bold text-[var(--ink)]">
                  Đăng nhập Google để sử dụng Google Sheets
                </h3>
              </div>
              <p className="text-sm text-[var(--muted)] leading-relaxed">
                Sau khi đăng nhập, bạn có thể chọn bảng tính Google Sheets từ Google Drive, tạo mới bảng tính chuẩn 6 trang cho lớp, đồng bộ danh sách sinh viên, điểm số, đánh giá, vi phạm, quỹ lớp và thông báo, hoặc xem trước và nhập dữ liệu trực tiếp từ sheet.
              </p>
            </div>
            <GoogleSignInButton
              onClick={handleGoogleSignIn}
              disabled={isSigningIn}
            />
          </div>
        </div>
      ) : (
        <>
          {/* Account bar */}
          <div className={cardClass}>
            <div className="flex flex-wrap items-center justify-between gap-4">
              <div className="flex items-center gap-3 min-w-0">
                {googleUser.photoURL ? (
                  <img
                    src={googleUser.photoURL}
                    alt={googleUser.displayName || "Google User"}
                    className="w-10 h-10 rounded-full border border-[var(--line)] shrink-0 object-cover"
                    referrerPolicy="no-referrer"
                  />
                ) : (
                  <div className="w-10 h-10 rounded-full bg-[var(--primary-soft)] text-[var(--primary-soft-ink)] grid place-items-center font-bold text-sm shrink-0">
                    G
                  </div>
                )}
                <div className="min-w-0">
                  <div className="flex items-center gap-1.5 text-sm font-bold truncate">
                    <span>{googleUser.displayName || "Tài khoản Google"}</span>
                    <CheckCircle2 className="w-4 h-4 text-emerald-600 dark:text-emerald-400 shrink-0" />
                  </div>
                  <div className="text-xs text-[var(--muted)] truncate">
                    {googleUser.email}
                  </div>
                </div>
              </div>

              <div className="flex flex-wrap items-center gap-2">
                <button
                  type="button"
                  onClick={handleCreateSpreadsheet}
                  disabled={creating}
                  className={btnPrimary}
                >
                  <FilePlus2 className="w-4 h-4" />
                  <span>
                    {creating
                      ? "Đang tạo..."
                      : "Tạo Google Sheet mới cho lớp"}
                  </span>
                </button>
                <button
                  type="button"
                  onClick={handleGoogleSignOut}
                  className="inline-flex items-center gap-1.5 px-3 py-2.5 rounded-xl border border-[var(--line)] text-xs font-semibold text-[var(--ink)] hover:bg-[var(--panel-subtle)] transition-colors cursor-pointer whitespace-nowrap"
                >
                  <LogOut className="w-4 h-4" />
                  <span>Đăng xuất Google</span>
                </button>
              </div>
            </div>
          </div>

          {/* Open by ID/URL or Pick from Drive */}
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
            <div className={cardClass}>
              <div className="flex items-center gap-2 mb-3">
                <FolderOpen className="w-4 h-4 text-[var(--primary)]" />
                <h3 className="text-base font-bold">
                  Mở bảng tính theo ID hoặc liên kết
                </h3>
              </div>
              <form
                onSubmit={(e) => {
                  e.preventDefault();
                  openSpreadsheetById(sheetInput);
                }}
                className="flex flex-col sm:flex-row gap-2.5"
              >
                <input
                  type="text"
                  value={sheetInput}
                  onChange={(e) => setSheetInput(e.target.value)}
                  placeholder="Dán link Google Sheets hoặc Spreadsheet ID..."
                  className={inputClass}
                />
                <button
                  type="submit"
                  disabled={loadingMeta}
                  className={btnPrimary}
                >
                  {loadingMeta ? "Đang mở..." : "Mở bảng tính"}
                </button>
              </form>

              {meta && (
                <div className="mt-4 pt-4 border-t border-[var(--line)] flex flex-wrap items-center justify-between gap-3">
                  <div className="min-w-0">
                    <div className="text-xs text-[var(--muted)]">
                      Bảng tính đang kết nối
                    </div>
                    <div className="text-sm font-bold truncate">
                      {meta.title}
                    </div>
                    <div className="text-xs font-mono-num text-[var(--muted)] truncate">
                      {meta.spreadsheetId}
                    </div>
                  </div>
                  {meta.spreadsheetUrl && (
                    <a
                      href={meta.spreadsheetUrl}
                      target="_blank"
                      rel="noopener noreferrer"
                      className={btnSecondary}
                    >
                      <span>Mở trên Google Sheets</span>
                      <ExternalLink className="w-3.5 h-3.5" />
                    </a>
                  )}
                </div>
              )}
            </div>

            <div className={cardClass}>
              <div className="flex items-center justify-between gap-2 mb-3">
                <div className="flex items-center gap-2">
                  <FileSpreadsheet className="w-4 h-4 text-[var(--primary)]" />
                  <h3 className="text-base font-bold">
                    Google Sheets trên Drive của bạn
                  </h3>
                </div>
                <button
                  type="button"
                  onClick={handleFetchDriveFiles}
                  disabled={loadingDrive}
                  className={btnSecondary}
                >
                  <RefreshCw
                    className={`w-3.5 h-3.5 ${loadingDrive ? "animate-spin" : ""}`}
                  />
                  <span>{loadingDrive ? "Đang tải..." : "Tải lại"}</span>
                </button>
              </div>

              {driveFiles.length > 0 ? (
                <div className="max-h-52 overflow-y-auto divide-y divide-[var(--line)] border border-[var(--line)] rounded-xl">
                  {driveFiles.map((file) => {
                    const active = meta?.spreadsheetId === file.id;
                    return (
                      <button
                        key={file.id}
                        type="button"
                        onClick={() => openSpreadsheetById(file.id)}
                        className={`w-full text-left px-3 py-2.5 flex items-center justify-between gap-3 text-sm transition-colors cursor-pointer ${
                          active
                            ? "bg-[var(--primary-soft)] text-[var(--primary-soft-ink)] font-semibold"
                            : "hover:bg-[var(--panel-subtle)]"
                        }`}
                      >
                        <span className="truncate">{file.name}</span>
                        {file.modifiedTime && (
                          <span className="text-xs text-[var(--muted)] tabular-nums shrink-0">
                            {new Date(file.modifiedTime).toLocaleDateString(
                              "vi-VN"
                            )}
                          </span>
                        )}
                      </button>
                    );
                  })}
                </div>
              ) : (
                <div className="py-8 text-center text-sm text-[var(--muted)] border border-dashed border-[var(--line)] rounded-xl">
                  {loadingDrive
                    ? "Đang tìm tệp Google Sheets trên Drive..."
                    : "Chưa có tệp Google Sheets nào được hiển thị."}
                </div>
              )}
            </div>
          </div>

          {/* Sync card */}
          {meta && (
            <div className={cardClass}>
              <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
                <div className="space-y-1">
                  <div className="flex items-center gap-2">
                    <UploadCloud className="w-4 h-4 text-[var(--primary)]" />
                    <h3 className="text-base font-bold">
                      Đồng bộ dữ liệu lớp lên Google Sheet
                    </h3>
                  </div>
                  <p className="text-sm text-[var(--muted)]">
                    Ghi dữ liệu hiện tại (
                    <span className="tabular-nums font-semibold text-[var(--ink)]">
                      {db.students.length}
                    </span>{" "}
                    sinh viên,{" "}
                    <span className="tabular-nums font-semibold text-[var(--ink)]">
                      {db.grades.length}
                    </span>{" "}
                    điểm,{" "}
                    <span className="tabular-nums font-semibold text-[var(--ink)]">
                      {db.evaluations.length}
                    </span>{" "}
                    đánh giá,{" "}
                    <span className="tabular-nums font-semibold text-[var(--ink)]">
                      {db.violations.length}
                    </span>{" "}
                    vi phạm,{" "}
                    <span className="tabular-nums font-semibold text-[var(--ink)]">
                      {db.fundRecords.length}
                    </span>{" "}
                    thu/chi,{" "}
                    <span className="tabular-nums font-semibold text-[var(--ink)]">
                      {db.announcements.length}
                    </span>{" "}
                    thông báo) vào bảng tính{" "}
                    <span className="font-semibold text-[var(--ink)]">
                      {meta.title}
                    </span>
                    .
                  </p>
                </div>

                <button
                  type="button"
                  onClick={handleSyncToSheet}
                  disabled={syncing}
                  className={btnPrimary}
                >
                  <UploadCloud className="w-4 h-4" />
                  <span>
                    {syncing
                      ? "Đang đồng bộ..."
                      : "Đồng bộ dữ liệu lớp lên Google Sheet"}
                  </span>
                </button>
              </div>
            </div>
          )}

          {/* Sheet tab preview & import */}
          {meta && (
            <div className={cardClass}>
              <div className="flex flex-wrap items-center justify-between gap-3 mb-4">
                <div className="flex items-center gap-2">
                  <Table2 className="w-4 h-4 text-[var(--primary)]" />
                  <h3 className="text-base font-bold">
                    Xem trước & nhập dữ liệu từ trang tính
                  </h3>
                </div>

                <div className="flex flex-wrap items-center gap-2">
                  <select
                    value={selectedTab}
                    onChange={(e) => {
                      const nextTab = e.target.value;
                      setSelectedTab(nextTab);
                      loadTabValues(meta.spreadsheetId, nextTab);
                    }}
                    aria-label="Chọn trang tính Google Sheet"
                    className="px-3 py-2 rounded-xl border border-[var(--input-border)] bg-[var(--input-bg)] text-[var(--ink)] text-sm font-medium focus:outline-none focus:border-[var(--primary)]"
                  >
                    {meta.sheets.map((tab) => (
                      <option key={tab.sheetId} value={tab.title}>
                        {tab.title}
                      </option>
                    ))}
                  </select>

                  <button
                    type="button"
                    onClick={() =>
                      selectedTab &&
                      loadTabValues(meta.spreadsheetId, selectedTab)
                    }
                    disabled={loadingRows || !selectedTab}
                    className={btnSecondary}
                  >
                    <RefreshCw
                      className={`w-3.5 h-3.5 ${
                        loadingRows ? "animate-spin" : ""
                      }`}
                    />
                    <span>{loadingRows ? "Đang đọc..." : "Đọc dữ liệu tab"}</span>
                  </button>

                  {isAdmin && sheetRows.length > 0 && (
                    <button
                      type="button"
                      onClick={handleImportStudents}
                      className={btnSecondary}
                    >
                      <UserPlus className="w-3.5 h-3.5" />
                      <span>Nhập sinh viên từ tab</span>
                    </button>
                  )}

                  {canImportGrades && sheetRows.length > 0 && (
                    <button
                      type="button"
                      onClick={handleImportGrades}
                      className={btnSecondary}
                    >
                      <UploadCloud className="w-3.5 h-3.5" />
                      <span>Nhập điểm từ tab</span>
                    </button>
                  )}
                </div>
              </div>

              {loadingRows ? (
                <div className="py-10 text-center text-sm text-[var(--muted)]">
                  Đang đọc dữ liệu từ trang "{selectedTab}"...
                </div>
              ) : previewRows.length > 0 ? (
                <>
                  <div className="overflow-x-auto border border-[var(--line)] rounded-xl">
                    <table className="w-full border-collapse text-sm">
                      <thead>
                        <tr className="border-b border-[var(--line)] bg-[var(--panel-subtle)] text-left text-xs font-semibold text-[var(--muted)]">
                          <th className="py-2 px-3 w-12 tabular-nums">#</th>
                          {Array.from({ length: maxCols }).map((_, colIdx) => (
                            <th key={colIdx} className="py-2 px-3">
                              {previewRows[0]?.[colIdx] || `Cột ${colIdx + 1}`}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {previewRows.slice(1).length > 0 ? (
                          previewRows.slice(1).map((row, rowIdx) => (
                            <tr
                              key={rowIdx}
                              className="border-b border-[var(--line)] last:border-b-0"
                            >
                              <td className="py-2 px-3 text-xs text-[var(--muted)] tabular-nums">
                                {rowIdx + 1}
                              </td>
                              {Array.from({ length: maxCols }).map(
                                (_, colIdx) => (
                                  <td
                                    key={colIdx}
                                    className="py-2 px-3 tabular-nums"
                                  >
                                    {row[colIdx] ?? ""}
                                  </td>
                                )
                              )}
                            </tr>
                          ))
                        ) : (
                          <tr>
                            <td
                              colSpan={maxCols + 1}
                              className="py-6 text-center text-[var(--muted)]"
                            >
                              Trang tính chỉ có dòng tiêu đề.
                            </td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                  {sheetRows.length > 50 && (
                    <p className="text-xs text-[var(--muted)] mt-2 tabular-nums">
                      Đang hiển thị 50 dòng đầu tiên trên tổng số{" "}
                      {sheetRows.length} dòng.
                    </p>
                  )}
                </>
              ) : (
                <div className="py-10 text-center text-sm text-[var(--muted)] border border-dashed border-[var(--line)] rounded-xl">
                  Trang "{selectedTab}" chưa có dữ liệu.
                </div>
              )}
            </div>
          )}
        </>
      )}
    </div>
  );
}
