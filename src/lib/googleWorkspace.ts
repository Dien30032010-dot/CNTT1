import { initializeApp, getApps } from "firebase/app";
import {
  getAuth,
  GoogleAuthProvider,
  onAuthStateChanged,
  signInWithPopup,
  type User,
} from "firebase/auth";
import firebaseConfig from "../../firebase-applet-config.json";
import type { ClassDatabase } from "../types";

export const WORKSPACE_SCOPES = [
  "https://www.googleapis.com/auth/drive",
  "https://www.googleapis.com/auth/drive.file",
  "https://www.googleapis.com/auth/drive.readonly",
  "https://www.googleapis.com/auth/spreadsheets",
  "https://www.googleapis.com/auth/spreadsheets.readonly",
] as const;

const app =
  getApps().length > 0 ? getApps()[0] : initializeApp(firebaseConfig);
const auth = getAuth(app);

const provider = new GoogleAuthProvider();
for (const scope of WORKSPACE_SCOPES) {
  provider.addScope(scope);
}

// Strict in-memory token cache (never persisted to localStorage or sessionStorage)
let isSigningIn = false;
let cachedAccessToken: string | null = null;

export function initGoogleAuth(
  onAuthSuccess: (user: User, token: string) => void,
  onAuthFailure: () => void
) {
  return onAuthStateChanged(auth, (user) => {
    if (user) {
      if (cachedAccessToken) {
        onAuthSuccess(user, cachedAccessToken);
      } else if (!isSigningIn) {
        cachedAccessToken = null;
        onAuthFailure();
      }
    } else {
      cachedAccessToken = null;
      onAuthFailure();
    }
  });
}

export async function googleSignIn(): Promise<{
  user: User;
  accessToken: string;
}> {
  try {
    isSigningIn = true;
    const result = await signInWithPopup(auth, provider);
    const credential = GoogleAuthProvider.credentialFromResult(result);
    if (!credential?.accessToken) {
      throw new Error("Không nhận được quyền truy cập Google Workspace.");
    }
    cachedAccessToken = credential.accessToken;
    return { user: result.user, accessToken: cachedAccessToken };
  } finally {
    isSigningIn = false;
  }
}

export function getAccessToken(): string | null {
  return cachedAccessToken;
}

export async function googleSignOut(): Promise<void> {
  await auth.signOut();
  cachedAccessToken = null;
}

export interface DriveSpreadsheetFile {
  id: string;
  name: string;
  modifiedTime?: string;
  webViewLink?: string;
}

export interface SheetTabMeta {
  sheetId: number;
  title: string;
}

export interface SpreadsheetMetadata {
  spreadsheetId: string;
  title: string;
  spreadsheetUrl?: string;
  sheets: SheetTabMeta[];
}

export function extractSpreadsheetId(input: string): string {
  const trimmed = input.trim();
  const match = trimmed.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (match?.[1]) return match[1];
  return trimmed;
}

async function googleFetch<T>(
  url: string,
  token: string,
  init?: RequestInit
): Promise<T> {
  const res = await fetch(url, {
    ...init,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      ...(init?.headers ?? {}),
    },
  });

  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    const errMsg =
      (data as { error?: { message?: string } })?.error?.message ||
      `Lỗi Google API (${res.status})`;
    if (res.status === 401) {
      cachedAccessToken = null;
    }
    throw new Error(errMsg);
  }
  return data as T;
}

export async function listSpreadsheets(
  token: string
): Promise<DriveSpreadsheetFile[]> {
  const query = encodeURIComponent(
    "mimeType='application/vnd.google-apps.spreadsheet' and trashed=false"
  );
  const fields = encodeURIComponent(
    "files(id,name,modifiedTime,webViewLink)"
  );
  const data = await googleFetch<{ files?: DriveSpreadsheetFile[] }>(
    `https://www.googleapis.com/drive/v3/files?q=${query}&fields=${fields}&orderBy=modifiedTime desc&pageSize=30`,
    token
  );
  return data.files ?? [];
}

export async function getSpreadsheetMetadata(
  token: string,
  spreadsheetId: string
): Promise<SpreadsheetMetadata> {
  const fields = encodeURIComponent(
    "spreadsheetId,spreadsheetUrl,properties.title,sheets.properties(sheetId,title)"
  );
  const data = await googleFetch<{
    spreadsheetId: string;
    spreadsheetUrl?: string;
    properties?: { title?: string };
    sheets?: { properties?: { sheetId?: number; title?: string } }[];
  }>(
    `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(
      spreadsheetId
    )}?fields=${fields}`,
    token
  );

  return {
    spreadsheetId: data.spreadsheetId,
    title: data.properties?.title || "Bảng tính Google Sheets",
    spreadsheetUrl: data.spreadsheetUrl,
    sheets: (data.sheets ?? []).map((s) => ({
      sheetId: s.properties?.sheetId ?? 0,
      title: s.properties?.title || "Sheet",
    })),
  };
}

export async function readSheetValues(
  token: string,
  spreadsheetId: string,
  sheetTitle: string
): Promise<string[][]> {
  const escapedTitle = `'${sheetTitle.replace(/'/g, "''")}'`;
  const range = encodeURIComponent(`${escapedTitle}!A1:Z500`);
  const data = await googleFetch<{ values?: string[][] }>(
    `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(
      spreadsheetId
    )}/values/${range}`,
    token
  );
  return data.values ?? [];
}

export const CLASS_SHEET_TABS = [
  "Sinh viên",
  "Điểm số",
  "Đánh giá",
  "Vi phạm",
  "Quỹ lớp",
  "Thông báo",
] as const;

export async function createClassSpreadsheet(
  token: string,
  title: string
): Promise<SpreadsheetMetadata> {
  const data = await googleFetch<{
    spreadsheetId: string;
    spreadsheetUrl?: string;
    properties?: { title?: string };
    sheets?: { properties?: { sheetId?: number; title?: string } }[];
  }>("https://sheets.googleapis.com/v4/spreadsheets", token, {
    method: "POST",
    body: JSON.stringify({
      properties: { title },
      sheets: CLASS_SHEET_TABS.map((tabTitle) => ({
        properties: { title: tabTitle },
      })),
    }),
  });

  return {
    spreadsheetId: data.spreadsheetId,
    title: data.properties?.title || title,
    spreadsheetUrl: data.spreadsheetUrl,
    sheets: (data.sheets ?? []).map((s) => ({
      sheetId: s.properties?.sheetId ?? 0,
      title: s.properties?.title || "Sheet",
    })),
  };
}

export async function syncDatabaseToSpreadsheet(
  token: string,
  spreadsheetId: string,
  db: ClassDatabase
): Promise<SpreadsheetMetadata> {
  const meta = await getSpreadsheetMetadata(token, spreadsheetId);
  const existingTitles = new Set(meta.sheets.map((s) => s.title));
  const missingTabs = CLASS_SHEET_TABS.filter((t) => !existingTitles.has(t));

  if (missingTabs.length > 0) {
    await googleFetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(
        spreadsheetId
      )}:batchUpdate`,
      token,
      {
        method: "POST",
        body: JSON.stringify({
          requests: missingTabs.map((title) => ({
            addSheet: { properties: { title } },
          })),
        }),
      }
    );
  }

  const studentName = (sid: string) =>
    db.students.find((s) => s.id === sid)?.name || sid;

  const clearRanges = CLASS_SHEET_TABS.map(
    (t) => `'${t.replace(/'/g, "''")}'!A1:Z1000`
  );
  await googleFetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(
      spreadsheetId
    )}/values:batchClear`,
    token,
    {
      method: "POST",
      body: JSON.stringify({ ranges: clearRanges }),
    }
  );

  const studentRows: (string | number)[][] = [
    ["STT", "Mã sinh viên", "Họ và tên"],
    ...db.students.map((s, idx) => [idx + 1, s.id, s.name]),
  ];

  const gradeRows: (string | number)[][] = [
    ["Mã sinh viên", "Họ và tên", "Môn học", "Điểm số", "Học kỳ", "Giáo viên", "Ghi chú", "Thời gian"],
    ...db.grades.map((g) => [
      g.sid,
      studentName(g.sid),
      g.subject,
      g.value,
      g.term || "",
      g.byName || g.by,
      g.note || "",
      new Date(g.at).toLocaleString("vi-VN"),
    ]),
  ];

  const evalRows: (string | number)[][] = [
    ["Mã sinh viên", "Họ và tên", "Mức đánh giá", "Nội dung", "Giáo viên", "Thời gian"],
    ...db.evaluations.map((e) => [
      e.sid,
      studentName(e.sid),
      e.level,
      e.note,
      e.byName || e.by,
      new Date(e.at).toLocaleString("vi-VN"),
    ]),
  ];

  const violationRows: (string | number)[][] = [
    ["Mã sinh viên", "Họ và tên", "Loại vi phạm", "Chi tiết", "Người ghi", "Thời gian"],
    ...db.violations.map((v) => [
      v.sid,
      studentName(v.sid),
      v.type,
      v.detail || "",
      v.byName || v.by,
      new Date(v.at).toLocaleString("vi-VN"),
    ]),
  ];

  const fundRows: (string | number)[][] = [
    ["Ngày", "Loại", "Nội dung", "Số tiền (VNĐ)", "Người ghi", "Chức vụ", "Ghi chú"],
    ...db.fundRecords.map((f) => [
      f.date,
      f.kind === "income" ? "Khoản thu" : "Khoản chi",
      f.title,
      f.amount,
      f.byName,
      f.byPosition || "",
      f.note || "",
    ]),
  ];

  const announcementRows: (string | number)[][] = [
    ["Tiêu đề", "Nội dung", "Người đăng", "Vai trò", "Thời gian"],
    ...db.announcements.map((a) => [
      a.title,
      a.content,
      a.authorName,
      a.authorDetail ? `${a.authorRole} (${a.authorDetail})` : a.authorRole,
      new Date(a.at).toLocaleString("vi-VN"),
    ]),
  ];

  await googleFetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(
      spreadsheetId
    )}/values:batchUpdate`,
    token,
    {
      method: "POST",
      body: JSON.stringify({
        valueInputOption: "USER_ENTERED",
        data: [
          { range: "'Sinh viên'!A1", values: studentRows },
          { range: "'Điểm số'!A1", values: gradeRows },
          { range: "'Đánh giá'!A1", values: evalRows },
          { range: "'Vi phạm'!A1", values: violationRows },
          { range: "'Quỹ lớp'!A1", values: fundRows },
          { range: "'Thông báo'!A1", values: announcementRows },
        ],
      }),
    }
  );

  return getSpreadsheetMetadata(token, spreadsheetId);
}
