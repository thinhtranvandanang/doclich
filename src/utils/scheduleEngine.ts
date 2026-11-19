import ExcelJS from 'exceljs';
import JSZip from 'jszip';

export interface ScheduleMetadata {
  term: string;
  dateBegin: Date;
  dateEnd: Date;
  dateBeginStr: string;
  dateEndStr: string;
  className: string;
  room: string;
  shiftLabel: string;
  totalWeeks: number;
  totalBlocks: number;
}

export interface SessionRecord {
  term: string;
  dateObj: Date;
  dateStr: string;
  slot: number;
  room: string;
  subject: string;
  className: string;
  lecturer: string;
  sessionNo: number;
}

export interface IgnoredNote {
  text: string;
  dateStr: string;
  slot: number;
}

export interface BlockSummary {
  label: string;
  weeksCount: number;
  daysCount: number;
  slotsRange: string;
  sessionsFound: number;
}

export interface ParseResult {
  metadata: ScheduleMetadata;
  records: SessionRecord[];
  subjects: string[];
  ignoredNotes: IgnoredNote[];
  blocks: BlockSummary[];
}

const VIETNAMESE_DIACRITICS_RE =
  /[àáạảãâầấậẩẫăằắặẳẵèéẹẻẽêềếệểễìíịỉĩòóọỏõôồốộổỗơờớợởỡùúụủũưừứựửữỳýỵỷỹđÀÁẠẢÃÂẦẤẬẨẪĂẰẮẶẲẴÈÉẸẺẼÊỀẾỆỂỄÌÍỊỈĨÒÓỌỎÕÔỒỐỘỔỖƠỜỚỢỞỠÙÚỤỦŨƯỪỨỰỬỮỲÝỴỶỸĐ]/;

const NON_SUBJECT_KEYWORDS = new Set([
  'NGHI',
  'LE',
  'HOLIDAY',
  'OFF',
  'BREAK',
  'EVENT',
  'ORIENTATION',
  'MEETING',
  'EXAM',
  'TEST',
  'FINAL',
  'MIDTERM',
  'CEREMONY',
  'OPENING',
  'CLOSING',
  'NOTE',
]);

const WEEK_BLOCK_HEADER_RE = /^Tuần\s+(\d+)(?:\s*[-–]\s*(\d+))?$/i;
const SLOT_LABEL_RE = /^Tiết\s*(\d+)$/i;
const WEEKDAY_MARKER_RE =
  /(thứ\s*(?:[2-7]|hai|ba|tư|năm|sáu|bảy)|chủ\s*nhật|^cn$|^t[2-7]$|mon|tue|wed|thu|fri|sat|sun)/i;

function pad2(n: number): string {
  return String(n).padStart(2, '0');
}

export function formatDateDDMMYYYY(d: Date): string {
  return `${pad2(d.getDate())}/${pad2(d.getMonth() + 1)}/${d.getFullYear()}`;
}

function parseDateDDMMYYYY(str: string): Date {
  const parts = str.trim().split('/');
  if (parts.length !== 3) {
    throw new Error(`Định dạng ngày không hợp lệ: ${str}`);
  }
  const day = Number(parts[0]);
  const month = Number(parts[1]);
  const year = Number(parts[2]);
  const dt = new Date(year, month - 1, day);
  if (
    Number.isNaN(dt.getTime()) ||
    dt.getFullYear() !== year ||
    dt.getMonth() !== month - 1 ||
    dt.getDate() !== day
  ) {
    throw new Error(`Ngày không hợp lệ: ${str}`);
  }
  return dt;
}

function addDays(d: Date, days: number): Date {
  const next = new Date(d.getFullYear(), d.getMonth(), d.getDate() + days);
  return next;
}

function cellText(cell: ExcelJS.Cell): string {
  const val = cell.value;
  if (val === null || val === undefined) return '';
  if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') {
    return String(val).trim();
  }
  if (val instanceof Date) {
    return formatDateDDMMYYYY(val);
  }
  if (typeof val === 'object') {
    if ('richText' in val && Array.isArray(val.richText)) {
      return val.richText.map((part) => part.text).join('').trim();
    }
    if ('result' in val && val.result !== undefined) {
      return String(val.result).trim();
    }
    if ('text' in val && typeof val.text === 'string') {
      return val.text.trim();
    }
  }
  return String(val).trim();
}

export function normalizeAndValidateSubject(raw: string): string | null {
  const text = raw.replace(/\s+/g, ' ').trim();
  if (!text) return null;
  if (VIETNAMESE_DIACRITICS_RE.test(text)) return null;
  if (text.includes('/') || text.length > 20) return null;

  const tokens = text.toUpperCase().split(/[\s_-]+/);
  if (tokens.some((tok) => NON_SUBJECT_KEYWORDS.has(tok))) return null;

  const compact = text.replace(/\s+/g, '').toUpperCase();
  if (/^[A-Z]{2,8}[-_]?[A-Z0-9]*\d[A-Z0-9]*$/.test(compact)) {
    return compact;
  }
  return null;
}

function parseHeaderMetadata(ws: ExcelJS.Worksheet): ScheduleMetadata {
  const headerLines: string[] = [];
  const maxCheckRow = Math.min(5, ws.rowCount);
  const maxCheckCol = Math.min(10, ws.columnCount);

  for (let r = 1; r <= maxCheckRow; r++) {
    const colAText = cellText(ws.getCell(r, 1));
    if (WEEK_BLOCK_HEADER_RE.test(colAText)) {
      break;
    }
    const rowTexts: string[] = [];
    for (let c = 1; c <= maxCheckCol; c++) {
      const txt = cellText(ws.getCell(r, c));
      if (txt) rowTexts.push(txt);
    }
    if (rowTexts.length > 0) {
      headerLines.push(rowTexts.join(' '));
    }
  }

  const fullHeader = headerLines.join('\n').trim();
  if (!fullHeader) {
    throw new Error('Không tìm thấy nội dung tiêu đề ở dòng đầu tiên (ô A1).');
  }

  const termMatch = fullHeader.match(/HỌC\s*KỲ\s*:?\s*([A-Za-z0-9\s_-]+?)(?=\s*\()/i);
  if (!termMatch) {
    throw new Error('Không thể bóc tách Học kỳ (ví dụ: HỌC KỲ FALL 2026) từ tiêu đề.');
  }
  const term = termMatch[1].replace(/\s+/g, ' ').trim().toUpperCase();

  const dateMatch = fullHeader.match(/\(?\s*(\d{1,2}\/\d{1,2}\/\d{4})\s*[-–]\s*(\d{1,2}\/\d{1,2}\/\d{4})\s*\)?/);
  if (!dateMatch) {
    throw new Error('Không thể bóc tách khoảng ngày (DD/MM/YYYY - DD/MM/YYYY) từ tiêu đề.');
  }
  const dateBegin = parseDateDDMMYYYY(dateMatch[1]);
  const dateEnd = parseDateDDMMYYYY(dateMatch[2]);
  if (dateEnd.getTime() < dateBegin.getTime()) {
    throw new Error('Ngày kết thúc nhỏ hơn ngày bắt đầu.');
  }

  const classMatch = fullHeader.match(/Lớp\s*:\s*([A-Za-z0-9._-]+)/i);
  if (!classMatch) {
    throw new Error('Không thể bóc tách Mã lớp (ví dụ: Lớp: PC2204) từ tiêu đề.');
  }
  const className = classMatch[1].trim().toUpperCase();

  const roomMatch = fullHeader.match(/Phòng\s*(?:học)?\s*:\s*([A-Za-z0-9._-]+)/i);
  if (!roomMatch) {
    throw new Error('Không thể bóc tách Phòng học (ví dụ: Phòng học: 107) từ tiêu đề.');
  }
  const room = roomMatch[1].trim();

  return {
    term,
    dateBegin,
    dateEnd,
    dateBeginStr: formatDateDDMMYYYY(dateBegin),
    dateEndStr: formatDateDDMMYYYY(dateEnd),
    className,
    room,
    shiftLabel: 'Chưa xác định',
    totalWeeks: 0,
    totalBlocks: 0,
  };
}

function detectBlockDayColumns(
  ws: ExcelJS.Worksheet,
  weekdayRow: number,
  blockHeaderText: string
): number {
  let countFromWeekdays = 0;
  for (let col = 2; col <= ws.columnCount + 1; col++) {
    const val = cellText(ws.getCell(weekdayRow, col));
    if (val && WEEKDAY_MARKER_RE.test(val)) {
      countFromWeekdays += 1;
    } else if (countFromWeekdays > 0) {
      const nextVal = col + 1 <= ws.columnCount ? cellText(ws.getCell(weekdayRow, col + 1)) : '';
      if (!nextVal) break;
    }
  }
  if (countFromWeekdays >= 7) return countFromWeekdays;

  const m = blockHeaderText.match(WEEK_BLOCK_HEADER_RE);
  if (m) {
    const startW = Number(m[1]);
    const endW = m[2] ? Number(m[2]) : startW;
    if (endW >= startW) return (endW - startW + 1) * 7;
  }

  const dateRow = weekdayRow + 1;
  let countFromDates = 0;
  for (let col = 2; col <= ws.columnCount; col++) {
    if (cellText(ws.getCell(dateRow, col))) {
      countFromDates += 1;
    } else if (countFromDates > 0) {
      break;
    }
  }
  return countFromDates;
}

export async function parseScheduleBuffer(buffer: ArrayBuffer): Promise<ParseResult> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  const ws = wb.worksheets[0];
  if (!ws) {
    throw new Error('File Excel không có sheet dữ liệu nào.');
  }

  const metadata = parseHeaderMetadata(ws);
  const records: SessionRecord[] = [];
  const ignoredNotes: IgnoredNote[] = [];
  const blocks: BlockSummary[] = [];
  const detectedSlots = new Set<number>();

  let currentDate = new Date(metadata.dateBegin.getTime());
  let totalDaysScanned = 0;
  let totalBlocks = 0;

  let r = 1;
  const maxRow = ws.rowCount;
  while (r <= maxRow) {
    const colA = cellText(ws.getCell(r, 1));
    const blockMatch = colA.match(WEEK_BLOCK_HEADER_RE);
    if (!blockMatch) {
      r += 1;
      continue;
    }

    let foundSlotStart: number | null = null;
    for (let probe = r + 1; probe <= Math.min(r + 6, maxRow); probe++) {
      if (SLOT_LABEL_RE.test(cellText(ws.getCell(probe, 1)))) {
        foundSlotStart = probe;
        break;
      }
    }

    if (foundSlotStart === null) {
      r += 1;
      continue;
    }

    const slotRows: Array<{ rowIdx: number; slotNum: number }> = [];
    let scanR = foundSlotStart;
    while (scanR <= maxRow) {
      const labelA = cellText(ws.getCell(scanR, 1));
      const sm = labelA.match(SLOT_LABEL_RE);
      if (!sm) break;
      const slotNum = Number(sm[1]);
      slotRows.push({ rowIdx: scanR, slotNum });
      detectedSlots.add(slotNum);
      scanR += 1;
    }

    if (slotRows.length === 0) {
      r += 1;
      continue;
    }

    const weekdayRow = Math.max(r + 1, foundSlotStart - 2);
    const numDayCols = detectBlockDayColumns(ws, weekdayRow, colA);
    if (numDayCols <= 0) {
      r = scanR;
      continue;
    }

    totalBlocks += 1;
    let blockSessions = 0;

    for (let offset = 0; offset < numDayCols; offset++) {
      const colIdx = 2 + offset;
      const cellDate = new Date(currentDate.getTime());
      const dateStr = formatDateDDMMYYYY(cellDate);

      for (const { rowIdx, slotNum } of slotRows) {
        const rawText = cellText(ws.getCell(rowIdx, colIdx));
        if (!rawText) continue;

        const subjectCode = normalizeAndValidateSubject(rawText);
        if (!subjectCode) {
          ignoredNotes.push({ text: rawText, dateStr, slot: slotNum });
          continue;
        }

        records.push({
          term: metadata.term,
          dateObj: cellDate,
          dateStr,
          slot: slotNum,
          room: metadata.room,
          subject: subjectCode,
          className: metadata.className,
          lecturer: '',
          sessionNo: 0,
        });
        blockSessions += 1;
      }

      currentDate = addDays(currentDate, 1);
      totalDaysScanned += 1;
    }

    const minBlockSlot = Math.min(...slotRows.map((s) => s.slotNum));
    const maxBlockSlot = Math.max(...slotRows.map((s) => s.slotNum));
    blocks.push({
      label: colA,
      weeksCount: Math.max(1, Math.round(numDayCols / 7)),
      daysCount: numDayCols,
      slotsRange: `Tiết ${minBlockSlot}–${maxBlockSlot}`,
      sessionsFound: blockSessions,
    });

    r = scanR;
  }

  if (totalBlocks === 0 || records.length === 0) {
    throw new Error(
      'Không quét được khối thời khóa biểu hoặc mã môn học hợp lệ nào. Vui lòng kiểm tra cấu trúc Schedule.xlsx.'
    );
  }

  if (detectedSlots.size > 0) {
    const minSlot = Math.min(...detectedSlots);
    const maxSlot = Math.max(...detectedSlots);
    if (minSlot >= 6) {
      metadata.shiftLabel = `Ca Chiều (Tiết ${minSlot} - ${maxSlot})`;
    } else if (maxSlot <= 5) {
      metadata.shiftLabel = `Ca Sáng (Tiết ${minSlot} - ${maxSlot})`;
    } else {
      metadata.shiftLabel = `Cả ngày (Tiết ${minSlot} - ${maxSlot})`;
    }
  }

  metadata.totalWeeks = Math.max(1, Math.round(totalDaysScanned / 7));
  metadata.totalBlocks = totalBlocks;

  const subjects = Array.from(new Set(records.map((rec) => rec.subject))).sort();

  return {
    metadata,
    records,
    subjects,
    ignoredNotes,
    blocks,
  };
}

export function getSubjectSessions(
  parseResult: ParseResult,
  subject: string,
  lecturerName: string
): SessionRecord[] {
  return parseResult.records
    .filter((r) => r.subject === subject)
    .sort((a, b) => a.dateObj.getTime() - b.dateObj.getTime() || a.slot - b.slot)
    .map((r, index) => ({
      ...r,
      lecturer: lecturerName.trim(),
      sessionNo: index + 1,
    }));
}

export async function exportSubjectWorkbookBuffer(
  parseResult: ParseResult,
  subject: string,
  lecturerName: string
): Promise<{ filename: string; buffer: ArrayBuffer }> {
  const sessions = getSubjectSessions(parseResult, subject, lecturerName);
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(subject.slice(0, 31));

  const headers = ['Term', 'Date', 'Slot', 'Room', 'Lecturer', 'Subject', 'Session No.', 'Class'];
  const headerRow = ws.addRow(headers);
  headerRow.height = 28;

  const thinBorder: Partial<ExcelJS.Borders> = {
    top: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    right: { style: 'thin', color: { argb: 'FFCBD5E1' } },
  };

  headerRow.eachCell((cell) => {
    cell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FF0F172A' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCE6F1' } };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = thinBorder;
  });

  sessions.forEach((s, idx) => {
    const row = ws.addRow([
      s.term,
      s.dateStr,
      s.slot,
      s.room,
      s.lecturer,
      s.subject,
      s.sessionNo,
      s.className,
    ]);
    row.height = 22;
    row.eachCell((cell) => {
      cell.font = { name: 'Calibri', size: 11, color: { argb: 'FF1E293B' } };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.border = thinBorder;
      if (idx % 2 === 1) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
      }
    });
  });

  ws.views = [{ state: 'frozen', ySplit: 1 }];

  headers.forEach((header, colIdx) => {
    let maxLen = header.length;
    sessions.forEach((s) => {
      const vals = [
        s.term,
        s.dateStr,
        String(s.slot),
        s.room,
        s.lecturer,
        s.subject,
        String(s.sessionNo),
        s.className,
      ];
      if (vals[colIdx].length > maxLen) {
        maxLen = vals[colIdx].length;
      }
    });
    ws.getColumn(colIdx + 1).width = Math.max(maxLen + 6, 12);
  });

  const buffer = await wb.xlsx.writeBuffer();
  return {
    filename: `${parseResult.metadata.className}_${subject}.xlsx`,
    buffer: buffer as ArrayBuffer,
  };
}

export async function exportAllSubjectsZip(
  parseResult: ParseResult,
  lecturerMap: Record<string, string>
): Promise<{ zipFilename: string; blob: Blob }> {
  const zip = new JSZip();
  for (const subject of parseResult.subjects) {
    const { filename, buffer } = await exportSubjectWorkbookBuffer(
      parseResult,
      subject,
      lecturerMap[subject] || ''
    );
    zip.file(filename, buffer);
  }
  const blob = await zip.generateAsync({ type: 'blob' });
  return {
    zipFilename: `${parseResult.metadata.className}_${parseResult.metadata.term.replace(/\s+/g, '_')}_Schedules.zip`,
    blob,
  };
}

export async function generateSampleScheduleBuffer(options: {
  totalWeeks: 17 | 21;
  isAfternoon: boolean;
  className: string;
  room: string;
  term: string;
  startDate: Date;
}): Promise<ArrayBuffer> {
  const { totalWeeks, isAfternoon, className, room, term, startDate } = options;
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Schedule');

  const endDate = addDays(startDate, totalWeeks * 7 - 1);
  const headerText = `THỜI KHÓA BIỂU - HỌC KỲ ${term} (${formatDateDDMMYYYY(startDate)} - ${formatDateDDMMYYYY(endDate)})\nLớp: ${className}     Phòng học: ${room}`;

  ws.getCell('A1').value = headerText;
  ws.getCell('A1').font = { name: 'Calibri', size: 12, bold: true, color: { argb: 'FF0F172A' } };
  ws.getCell('A1').alignment = { vertical: 'middle', wrapText: true };
  ws.getRow(1).height = 40;

  const weekdays = ['Thứ hai', 'Thứ ba', 'Thứ tư', 'Thứ năm', 'Thứ sáu', 'Thứ bảy', 'CN'];
  const slots = isAfternoon ? [6, 7, 8, 9, 10] : [1, 2, 3, 4, 5];

  const weeklyTemplate: Record<string, string> = isAfternoon
    ? {
        '0-0': 'MAT001',
        '0-1': 'MAT001',
        '1-0': 'PRO192',
        '1-1': 'PRO192',
        '1-2': 'ENG101',
        '2-0': 'CSI104',
        '2-1': 'CSI104',
        '3-0': 'PRO192',
        '3-1': 'PRO192',
        '3-3': 'WEB201',
        '4-0': 'MAT001',
        '4-2': 'ENG101',
        '4-3': 'ENG101',
        '5-0': 'WEB201',
        '5-1': 'WEB201',
      }
    : {
        '0-0': 'PRF192',
        '0-1': 'PRF192',
        '1-0': 'MAE101',
        '1-1': 'MAE101',
        '2-0': 'CEA201',
        '2-1': 'CEA201',
        '3-0': 'SSL101',
        '3-2': 'PRF192',
        '3-3': 'PRF192',
        '4-0': 'MAE101',
        '4-1': 'NWC203',
        '4-2': 'NWC203',
      };

  const thinBorder: Partial<ExcelJS.Borders> = {
    top: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    right: { style: 'thin', color: { argb: 'FFCBD5E1' } },
  };

  let currentRow = 3;
  let currentDate = new Date(startDate.getTime());
  let weekCursor = 1;

  while (weekCursor <= totalWeeks) {
    const weeksInBlock = Math.min(3, totalWeeks - weekCursor + 1);
    const endWeekInBlock = weekCursor + weeksInBlock - 1;
    const label =
      weeksInBlock > 1 ? `Tuần ${weekCursor}-${endWeekInBlock}` : `Tuần ${weekCursor}`;

    const blockCell = ws.getCell(currentRow, 1);
    blockCell.value = label;
    blockCell.font = { bold: true };
    blockCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };

    for (let w = 0; w < weeksInBlock; w++) {
      const colStart = 2 + w * 7;
      const subCell = ws.getCell(currentRow + 1, colStart);
      subCell.value = `Tuần ${weekCursor + w}`;
      subCell.font = { bold: true };
    }

    const totalDays = weeksInBlock * 7;
    for (let d = 0; d < totalDays; d++) {
      const col = 2 + d;
      const dayDate = addDays(currentDate, d);
      const dayOfWeek = d % 7;

      const wdCell = ws.getCell(currentRow + 2, col);
      wdCell.value = weekdays[dayOfWeek];
      wdCell.font = { bold: true, size: 9 };
      wdCell.alignment = { horizontal: 'center' };
      wdCell.border = thinBorder;

      const dtCell = ws.getCell(currentRow + 3, col);
      dtCell.value = dayDate.getDate();
      dtCell.alignment = { horizontal: 'center' };
      dtCell.border = thinBorder;
    }

    slots.forEach((slotNum, sIdx) => {
      const rSlot = currentRow + 4 + sIdx;
      const slotLabelCell = ws.getCell(rSlot, 1);
      slotLabelCell.value = `Tiết ${slotNum}`;
      slotLabelCell.font = { bold: true };
      slotLabelCell.border = thinBorder;

      for (let d = 0; d < totalDays; d++) {
        const col = 2 + d;
        const dayDate = addDays(currentDate, d);
        const dayOfWeek = d % 7;
        const cell = ws.getCell(rSlot, col);
        cell.alignment = { horizontal: 'center' };
        cell.border = thinBorder;

        if (dayDate.getMonth() === 8 && dayDate.getDate() === 2 && sIdx <= 1) {
          cell.value = 'Nghỉ lễ 2/9';
        } else if (dayDate.getMonth() === 3 && dayDate.getDate() === 30 && sIdx <= 1) {
          cell.value = 'Nghỉ lễ 30/4';
        } else if (weekCursor === 1 && dayOfWeek === 0 && sIdx === 2) {
          cell.value = 'Khai giảng K22';
        } else {
          const key = `${dayOfWeek}-${sIdx}`;
          if (weeklyTemplate[key]) {
            cell.value = weeklyTemplate[key];
          }
        }
      }
    });

    currentDate = addDays(currentDate, totalDays);
    weekCursor += weeksInBlock;
    currentRow += 10;
  }

  ws.getColumn(1).width = 14;
  const buf = await wb.xlsx.writeBuffer();
  return buf as ArrayBuffer;
}
