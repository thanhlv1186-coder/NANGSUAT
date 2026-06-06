import { readSheet } from "read-excel-file/browser";
import {
  KHO_LABEL as FALLBACK_KHO_LABEL,
  META as FALLBACK_META,
  VUNG_INFO as FALLBACK_VUNG_INFO,
} from "./data.js";

export const EXCEL_POLL_INTERVAL_MS = 30_000;

export function getExcelFileName(month) {
  return month ? `NANG SUAT T${month}.xlsx` : "NANG SUAT.xlsx";
}

function getExcelFileUrl(month) {
  const fileName = getExcelFileName(month);
  return `${import.meta.env.BASE_URL}${encodeURIComponent(fileName)}`;
}
const MS_PER_DAY = 24 * 60 * 60 * 1000;

const COLORS_BY_VUNG = {
  DBSH: "#00d4ff",
  DTB: "#a78bfa",
};

const TEXT_BY_VUNG = {
  DBSH: "Đồng Bằng Sông Hồng",
  DTB: "Đông Tây Bắc",
};

const COL = {
  vung: 0,
  khoMaCode: 1,
  khoName: 2,
  uid: 3,
  name: 4,
  loai: 5,
  trangThai: 6,
  tongQD: 7,
  ml: 8,
  spk: 9,
};

function cacheBustedUrl(month) {
  return `${getExcelFileUrl(month)}?v=${Date.now()}`;
}

function text(value) {
  return String(value ?? "").trim();
}

function normalizeKey(value) {
  return text(value)
    .replace(/đ/g, "d")
    .replace(/Đ/g, "D")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase();
}

function isEmpty(value) {
  return value === null || value === undefined || text(value) === "";
}

function toNumber(value) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (typeof value === "string") {
    const normalized = value.replace(/\s+/g, "").replace(",", ".");
    const parsed = Number(normalized);
    return Number.isFinite(parsed) ? parsed : 0;
  }
  return 0;
}

function roundHalfEven(value) {
  const number = toNumber(value);
  const sign = Math.sign(number) || 1;
  const abs = Math.abs(number);
  const floor = Math.floor(abs);
  const diff = abs - floor;

  if (diff > 0.5) return sign * (floor + 1);
  if (diff < 0.5) return sign * floor;
  return sign * (floor % 2 === 0 ? floor : floor + 1);
}

function valueToDate(value) {
  if (isEmpty(value)) return null;

  if (value instanceof Date && !Number.isNaN(value.valueOf())) return value;

  if (typeof value === "number" && Number.isFinite(value)) {
    return new Date((value - 25569) * MS_PER_DAY);
  }

  const parsed = new Date(value);
  return Number.isNaN(parsed.valueOf()) ? null : parsed;
}

function formatMonthLabel(date) {
  if (!date) return FALLBACK_META.label;
  const month = String(date.getUTCMonth() + 1).padStart(2, "0");
  return `Tháng ${month}/${date.getUTCFullYear()}`;
}

function codeFromVung(value) {
  const normalized = normalizeKey(value);
  if (normalized.includes("DONG BANG SONG HONG")) return "DBSH";
  if (normalized.includes("DONG TAY BAC")) return "DTB";
  return normalized
    .replace(/^VUNG\s+/, "")
    .split(/\s+/)
    .map((part) => part[0])
    .join("")
    .slice(0, 6) || "OTHER";
}

function khoFromName(value) {
  const raw = text(value);
  const codeMatch = raw.match(/(?:^|[_\s-])([A-Z]{2,4}(?:-[A-Z]{2,4})?)\s*-/);
  const code = codeMatch?.[1] || raw || "UNKNOWN";
  const generatedLabel = raw
    .split(" - ")
    .pop()
    ?.replace(/^Kho\s+(CN\s+)?Tận Tâm\s*/i, "")
    .trim();

  return {
    code,
    label: FALLBACK_KHO_LABEL[code] || generatedLabel || code,
  };
}

function normalizeLoai(value) {
  const raw = text(value);
  const normalized = normalizeKey(raw);
  if (normalized.includes("CONG TAC VIEN")) return "Cộng Tác Viên";
  if (normalized.includes("NHAN VIEN")) return "Nhân Viên";
  if (normalized.includes("DOI TAC")) return "8 - ĐỐI TÁC GHLĐ";
  if (normalized.includes("TAI XE")) return "Tài xế";
  return raw;
}

function normalizeTrangThai(value) {
  const raw = text(value);
  const normalized = normalizeKey(raw);
  if (normalized.includes("TAM KHOA")) return "Tạm khóa";
  if (normalized.includes("HOAT DONG")) return "Hoạt động";
  return raw || "Hoạt động";
}

function normalizeUid(value) {
  const raw = text(value);
  if (/^\d+$/.test(raw)) return Number(raw);
  return raw;
}

function buildDailyColumns(headerRow, subHeaderRow) {
  const cols = [];

  for (let col = 0; col < Math.max(headerRow.length, subHeaderRow.length); col += 1) {
    if (text(subHeaderRow[col]).toUpperCase() !== "ML") continue;

    const date = valueToDate(headerRow[col]);
    if (!date) continue;

    const spkCol = text(subHeaderRow[col + 1]).toUpperCase() === "SPK" ? col + 1 : -1;
    if (spkCol === -1) continue;

    cols.push({
      date,
      day: date.getUTCDate(),
      mlCol: col,
      spkCol,
    });
  }

  return cols;
}

function buildMeta(dailyColumns) {
  const firstDate = dailyColumns[0]?.date;

  return {
    month: firstDate ? firstDate.getUTCMonth() + 1 : FALLBACK_META.month,
    year: firstDate ? firstDate.getUTCFullYear() : FALLBACK_META.year,
    nDays: dailyColumns.length || FALLBACK_META.nDays,
    days: dailyColumns.length ? dailyColumns.map((col) => col.day) : undefined,
    label: formatMonthLabel(firstDate),
  };
}

const GENERIC_FILE_NAME = "NANG SUAT.xlsx";
const GENERIC_FILE_URL  = `${import.meta.env.BASE_URL}${encodeURIComponent(GENERIC_FILE_NAME)}`;

export async function loadExcelDashboardData(month) {
  const fileName = getExcelFileName(month);

  // Thử load file theo tháng trước, nếu không có thì dùng file gốc
  let response = await fetch(cacheBustedUrl(month), { cache: "no-store" });
  let usedFileName = fileName;

  if (!response.ok) {
    response = await fetch(`${GENERIC_FILE_URL}?v=${Date.now()}`, { cache: "no-store" });
    usedFileName = GENERIC_FILE_NAME;
    if (!response.ok) {
      throw new Error(`Không tìm thấy ${fileName} (${response.status})`);
    }
  }

  const buffer = await response.arrayBuffer();
  let rows;

  try {
    rows = await readSheet(buffer, "Don hang");
  } catch {
    rows = await readSheet(buffer, 1);
  }

  if (!Array.isArray(rows) || rows.length < 3) {
    throw new Error(`${usedFileName} không có đủ dữ liệu để dựng dashboard`);
  }

  const [headerRow = [], subHeaderRow = [], ...dataRows] = rows;
  const dailyColumns = buildDailyColumns(headerRow, subHeaderRow);

  if (!dailyColumns.length) {
    throw new Error(`${usedFileName} thiếu cặp cột ngày ML/SPK`);
  }

  const khoLabel = { ...FALLBACK_KHO_LABEL };
  const vungInfo = {
    ...FALLBACK_VUNG_INFO,
    ALL: FALLBACK_VUNG_INFO.ALL || { name: "Tất cả vùng", color: "#00d4ff" },
  };
  const raw = [];
  const daily = [];

  for (const row of dataRows) {
    if (!row || isEmpty(row[COL.uid]) || isEmpty(row[COL.name])) continue;

    const { code: khoCodeFallback, label: khoName } = khoFromName(row[COL.khoName]);
    const maKho = text(row[COL.khoMaCode]);
    const khoCode = maKho || khoCodeFallback;
    const vungCode = codeFromVung(row[COL.vung]);
    const uid = normalizeUid(row[COL.uid]);
    const mlByDay = dailyColumns.map((col) => roundHalfEven(row[col.mlCol]));
    const spkByDay = dailyColumns.map((col) => roundHalfEven(row[col.spkCol]));

    khoLabel[khoCode] = khoLabel[khoCode] || khoName;
    vungInfo[vungCode] = vungInfo[vungCode] || {
      name: TEXT_BY_VUNG[vungCode] || text(row[COL.vung]) || vungCode,
      color: COLORS_BY_VUNG[vungCode] || "#22c55e",
    };

    raw.push([
      uid,
      text(row[COL.name]),
      normalizeLoai(row[COL.loai]),
      vungCode,
      khoCode,
      roundHalfEven(row[COL.ml]),
      roundHalfEven(row[COL.spk]),
      roundHalfEven(row[COL.tongQD]),
      normalizeTrangThai(row[COL.trangThai]),
    ]);

    daily.push([uid, khoCode, mlByDay, spkByDay]);
  }

  if (!raw.length) {
    throw new Error(`${usedFileName} không có dòng nhân sự hợp lệ`);
  }

  return {
    raw,
    daily,
    khoLabel,
    vungInfo,
    meta: buildMeta(dailyColumns),
    source: usedFileName,
    loadedAt: new Date().toISOString(),
  };
}
