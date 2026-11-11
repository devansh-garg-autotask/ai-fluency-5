/**
 * Attendance Report Extractor - Multi-Format Engine
 *
 * Supports:
 * 1. Microsoft Excel 97-2003 Worksheet (.xls, BIFF8 binary format).
 * 2. Modern Microsoft Excel Workbooks (.xlsx).
 * 3. Legacy HTML attendance reports exported with .xls extension.
 * 4. Multi-sheet workbooks and multi-table layouts.
 * 5. Dynamic subject identification with code extracted from brackets.
 * 6. Dynamic column indexing for roll numbers, names, and attendance metrics.
 * 7. Deduplication by roll number, selecting highest attendance.
 */

import * as XLSX from 'xlsx';

export const BASE_COLUMNS = 6;
export const ATTENDANCE_COLUMNS = 3;

export interface ExtractedSubject {
  name: string;
  code: string;
  teachingType: string | null;
}

export interface TableSubject extends ExtractedSubject {
  columnIndex: number;
  tableIndex: number;
  dCol?: number;
  aCol?: number;
  pctCol?: number;
}

export interface SubjectAttendanceRecord {
  subjectName: string;
  subjectCode: string;
  teachingType: string | null;
  conducted: number | null;
  attended: number | null;
  percentage: number | null;
  className?: string;
  tableIndex?: number;
}

export interface ExtractedStudent {
  rollNumber: string;
  name: string;
  section: string;
  uniqueId: string;
  className: string;
  attendanceRecords: Record<string, SubjectAttendanceRecord[]>;
}

export interface ExtractionDiagnostics {
  totalTablesFound: number;
  validTablesProcessed: number;
  totalUniqueSubjects: number;
  totalUniqueStudents: number;
  totalStudentRecords: number;
  duplicateRollNumbers: string[];
  fileFormatDetected?: string;
  tables: Array<{
    tableIndex: number;
    className: string;
    subjectCount: number;
    studentCount: number;
    warnings: string[];
  }>;
}

export interface ExtractionResult {
  subjects: ExtractedSubject[];
  students: ExtractedStudent[];
  diagnostics: ExtractionDiagnostics;
}

export interface BestSubjectAttendance {
  subjectCode: string;
  subjectName: string;
  teachingType: string | null;
  conducted: number | null;
  attended: number | null;
  percentage: number | null;
  className?: string;
  sourceTable?: number;
}

export interface BestAttendanceStudent {
  rollNumber: string;
  name: string;
  section: string;
  uniqueId?: string;
  attendance: Record<string, BestSubjectAttendance | null>;
}

export function cleanText(value: any): string {
  return String(value ?? '')
    .replace(/\s+/g, ' ')
    .trim();
}

export function parseNumber(value: any): number | null {
  const cleaned = cleanText(value)
    .replace('%', '')
    .replace(/,/g, '');

  if (
    !cleaned ||
    cleaned === '-' ||
    cleaned === '—' ||
    cleaned.toLowerCase() === 'na' ||
    cleaned.toLowerCase() === 'n/a'
  ) {
    return null;
  }

  const number = Number(cleaned);
  return Number.isFinite(number) ? number : null;
}

function getCellText(cell: Element | null): string {
  return cleanText(cell?.textContent);
}

function getDirectCells(row: HTMLTableRowElement | Element): Element[] {
  return Array.from(row.children).filter(
    (element) => element.tagName === 'TD' || element.tagName === 'TH'
  );
}

function getTableRows(table: HTMLTableElement): HTMLTableRowElement[] {
  return Array.from(table.rows);
}

function getColSpan(cell: Element): number {
  const value = Number(cell.getAttribute('colspan') || 1);
  return Number.isInteger(value) && value > 0 ? value : 1;
}

function getRowSpan(cell: Element): number {
  const value = Number(cell.getAttribute('rowspan') || 1);
  return Number.isInteger(value) && value > 0 ? value : 1;
}

function getHeaderLines(cell: Element): string[] {
  const clone = cell.cloneNode(true) as Element;
  clone.querySelectorAll('br').forEach((br) => {
    br.replaceWith('\n');
  });

  return (clone.textContent || '')
    .split(/\n/)
    .map(cleanText)
    .filter(Boolean);
}

/**
 * Robust Subject and Course Code parser:
 * 1. "Data Structures (CS201)" -> name: "Data Structures", code: "CS201"
 * 2. "Algorithms (CS202) [PP]" -> name: "Algorithms", code: "CS202", teachingType: "PP"
 * 3. "(MATH101) Calculus" -> name: "Calculus", code: "MATH101"
 * 4. "CS301 - Operating Systems" -> name: "Operating Systems", code: "CS301"
 * 5. Multiline cells with code on its own line
 */
export function parseSubjectInfo(text: string): ExtractedSubject | null {
  const cleaned = cleanText(text);
  if (!cleaned) return null;

  // Ignore general header labels
  if (
    /^(sr\.?|s\.?\s*no\.?|sl\.?\s*no\.?|roll|roll\s*no\.?|roll\s*number|rollno|name|student\s*name|student|section|sec|uid|unique\s*id|class|total|average|attendance|division)$/i.test(
      cleaned
    )
  ) {
    return null;
  }

  // Pattern 1: Name (CODE) [PP|PR] or Name (CODE)
  const bracketMatch = cleaned.match(
    /^(.*?)\s*[\(\[]\s*([A-Za-z0-9_-]+?)\s*[\)\]](?:\s*\[\s*(PP|PR)\s*\])?$/i
  );
  if (bracketMatch) {
    const rawName = cleanText(bracketMatch[1]);
    const code = cleanText(bracketMatch[2]);
    const teachingType = bracketMatch[3] ? bracketMatch[3].toUpperCase() : null;
    return { name: rawName || code, code, teachingType };
  }

  // Pattern 2: [PP|PR] Name (CODE)
  const prefixBracketMatch = cleaned.match(
    /^(?:\[\s*(PP|PR)\s*\]\s*)?(.*?)\s*[\(\[]\s*([A-Za-z0-9_-]+?)\s*[\)\]]$/i
  );
  if (prefixBracketMatch) {
    const teachingType = prefixBracketMatch[1] ? prefixBracketMatch[1].toUpperCase() : null;
    const rawName = cleanText(prefixBracketMatch[2]);
    const code = cleanText(prefixBracketMatch[3]);
    return { name: rawName || code, code, teachingType };
  }

  // Pattern 3: (CODE) Name
  const leadingBracketMatch = cleaned.match(
    /^[\(\[]\s*([A-Za-z0-9_-]+?)\s*[\)\]]\s*(.*?)(?:\s*\[\s*(PP|PR)\s*\])?$/i
  );
  if (leadingBracketMatch) {
    const code = cleanText(leadingBracketMatch[1]);
    const rawName = cleanText(leadingBracketMatch[2]);
    const teachingType = leadingBracketMatch[3] ? leadingBracketMatch[3].toUpperCase() : null;
    return { name: rawName || code, code, teachingType };
  }

  // Pattern 4: CODE - Name
  const prefixMatch = cleaned.match(/^([A-Za-z0-9_-]{3,})\s*[-:]\s*(.*?)(?:\s*\[\s*(PP|PR)\s*\])?$/i);
  if (prefixMatch) {
    const code = cleanText(prefixMatch[1]);
    const rawName = cleanText(prefixMatch[2]);
    const teachingType = prefixMatch[3] ? prefixMatch[3].toUpperCase() : null;
    return { name: rawName || code, code, teachingType };
  }

  // Pattern 5: Multiline cell (from ERP HTML tables)
  const lines = String(text ?? '')
    .split(/\n|<br\s*\/?>/i)
    .map(cleanText)
    .filter(Boolean);

  if (lines.length > 1) {
    const codeIndex = lines.findIndex(
      (line) =>
        /^[A-Z]{2,}[A-Z0-9]*\d+[A-Z0-9]*$/i.test(line) ||
        /^ns_cec26_hs101$/i.test(line)
    );

    if (codeIndex !== -1) {
      const code = lines[codeIndex];
      const teachingType =
        lines.find((line) => /^(PP|PR)$/i.test(line))?.toUpperCase() || null;

      const name = lines
        .filter((line, index) => {
          if (index === codeIndex) return false;
          if (/^(PP|PR)$/i.test(line)) return false;
          if (/^(Core|Generic|Generic Elective)$/i.test(line)) return false;
          return true;
        })
        .join(' ')
        .replace(/[()]/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();

      return {
        name: name || code,
        code,
        teachingType,
      };
    }
  }

  // Pattern 6: Standard Course Code standalone (e.g. CS201, MATH101, PHY102)
  if (/^[A-Z]{2,}[A-Z0-9_-]*\d+[A-Z0-9_-]*$/i.test(cleaned)) {
    return { name: cleaned, code: cleaned, teachingType: null };
  }

  return null;
}

export function isStudentRollNumber(value: any): boolean {
  const cleaned = cleanText(value);
  if (!cleaned) return false;
  if (/^(total|grand\s*total|average|sr\.?|roll|name|na|n\/a|absent|present)$/i.test(cleaned)) {
    return false;
  }
  // Allow numeric (e.g. 2026001) or alphanumeric codes (e.g. CS-2026-001, 21BCE045, CS01)
  return /^[A-Za-z0-9_-]{2,}$/.test(cleaned) && /[0-9]/.test(cleaned);
}

function findAttendanceHeaderIndex(rows: HTMLTableRowElement[]): number {
  return rows.findIndex((row) => {
    const cells = getDirectCells(row);
    return cells.some((cell) => {
      const text = getCellText(cell).toLowerCase();
      return (
        text === 'sr.' ||
        text === 'sr' ||
        text === 's.no' ||
        text === 's.no.' ||
        text === 'roll no' ||
        text === 'roll number' ||
        text === 'rollno' ||
        text === 'roll' ||
        text === 'student name' ||
        text === 'identifier'
      );
    });
  });
}

function extractClassName(rows: HTMLTableRowElement[], headerIndex: number): string {
  const metadata = rows
    .slice(0, headerIndex)
    .map((row) => getCellText(row))
    .join(' ');

  const match =
    metadata.match(/Class Name:\s*(.*?)(?=Division:|$)/i) ||
    metadata.match(/Class:\s*(.*?)(?=$)/i);
  return match ? cleanText(match[1]) : '';
}

function validateAttendanceHeaders(subheaderCells: Element[], subjectOffset: number): boolean {
  const labels = subheaderCells
    .slice(subjectOffset, subjectOffset + 3)
    .map((cell) => getCellText(cell).toUpperCase());

  if (labels.length !== 3) {
    return false;
  }

  const isDelivered = /^D\b/.test(labels[0]) || /DELIVERED|TOTAL|CONDUCTED/.test(labels[0]);
  const isAttended = /^A\b/.test(labels[1]) || /ATTENDED|PRESENT/.test(labels[1]);
  const isPercentage = /^%/.test(labels[2]) || /PCT|PERCENT/.test(labels[2]);

  return isDelivered && isAttended && isPercentage;
}

/**
 * Extracts attendance data from an HTML <table> element.
 */
function extractHtmlTable(table: HTMLTableElement, tableIndex: number) {
  const rows = getTableRows(table);
  const headerIndex = findAttendanceHeaderIndex(rows);

  if (headerIndex === -1) {
    return null;
  }

  const subjectHeaderRow = rows[headerIndex];
  const subjectHeaderCells = getDirectCells(subjectHeaderRow);

  const subheaderRow = rows[headerIndex + 1];
  const subheaderCells = subheaderRow ? getDirectCells(subheaderRow) : [];

  // Determine metadata column positions
  let rollColIndex = -1;
  let nameColIndex = -1;
  let secColIndex = -1;
  let uidColIndex = -1;

  let currentOffset = 0;
  subjectHeaderCells.forEach((cell, idx) => {
    const text = getCellText(cell).toLowerCase();
    const colspan = getColSpan(cell);
    if (/^roll\s*(no|number)?$/i.test(text) || text === 'rollno' || text === 'identifier') {
      rollColIndex = currentOffset;
    } else if (/^(student\s*name|name|full\s*name)$/i.test(text)) {
      nameColIndex = currentOffset;
    } else if (/^(section|sec)$/i.test(text)) {
      secColIndex = currentOffset;
    } else if (/^(uid|unique\s*id)$/i.test(text)) {
      uidColIndex = currentOffset;
    }
    currentOffset += colspan;
  });

  const subjects: TableSubject[] = [];
  const warnings: string[] = [];

  let attendanceColumnIndex = 0;
  let subjectOffset = 0;

  for (const cell of subjectHeaderCells) {
    const colspan = getColSpan(cell);
    const rowspan = getRowSpan(cell);
    const headerText = getCellText(cell);

    if (rowspan > 1 && /^(sr\.?|roll|name|sec|uid)/i.test(headerText)) {
      attendanceColumnIndex += colspan;
      subjectOffset += colspan;
      continue;
    }

    if (/^total$/i.test(headerText)) {
      break;
    }

    const details = parseSubjectInfo(headerText) || parseSubjectInfo(getHeaderLines(cell).join('\n'));

    if (details && colspan === ATTENDANCE_COLUMNS) {
      const headersValid = subheaderCells.length > 0
        ? validateAttendanceHeaders(subheaderCells, subjectOffset)
        : true;

      subjects.push({
        ...details,
        columnIndex: attendanceColumnIndex,
        tableIndex,
        dCol: attendanceColumnIndex,
        aCol: attendanceColumnIndex + 1,
        pctCol: attendanceColumnIndex + 2,
      });

      if (!headersValid) {
        warnings.push(`D/A/% subheaders varied for ${details.code}`);
      }
    } else if (details && colspan === 1) {
      subjects.push({
        ...details,
        columnIndex: attendanceColumnIndex,
        tableIndex,
        pctCol: attendanceColumnIndex,
      });
    }

    attendanceColumnIndex += colspan;
    subjectOffset += colspan;
  }

  if (!subjects.length) {
    return null;
  }

  const className = extractClassName(rows, headerIndex) || `Table ${tableIndex + 1}`;
  const dataStartRow = subheaderCells.length > 0 ? headerIndex + 2 : headerIndex + 1;

  // Fallback metadata column indices if not found in header text
  if (rollColIndex === -1) {
    rollColIndex = 3; // traditional ERP offset
  }
  if (nameColIndex === -1) {
    nameColIndex = rollColIndex + 1;
  }

  const students: Array<{
    rollNumber: string;
    name: string;
    section: string;
    uniqueId: string;
    className: string;
    attendance: Record<string, SubjectAttendanceRecord>;
  }> = [];

  for (let i = dataStartRow; i < rows.length; i++) {
    const cells = getDirectCells(rows[i]);
    if (cells.length < 2) continue;

    const values = cells.map(getCellText);
    let rollNumber = values[rollColIndex];

    // Fallback if rollColIndex didn't match a roll number
    if (!isStudentRollNumber(rollNumber)) {
      const possibleRollIdx = values.findIndex((val) => isStudentRollNumber(val));
      if (possibleRollIdx !== -1) {
        rollNumber = values[possibleRollIdx];
        if (nameColIndex === rollColIndex + 1) {
          nameColIndex = possibleRollIdx + 1;
        }
      } else {
        continue;
      }
    }

    const student = {
      rollNumber: cleanText(rollNumber),
      name: nameColIndex >= 0 && nameColIndex < values.length ? values[nameColIndex] : '',
      section: secColIndex >= 0 && secColIndex < values.length ? values[secColIndex] : '',
      uniqueId: uidColIndex >= 0 && uidColIndex < values.length ? values[uidColIndex] : '',
      className,
      attendance: {} as Record<string, SubjectAttendanceRecord>,
    };

    for (const subject of subjects) {
      let conducted: number | null = null;
      let attended: number | null = null;
      let percentage: number | null = null;

      if (subject.dCol !== undefined && subject.aCol !== undefined) {
        conducted = parseNumber(values[subject.dCol]);
        attended = parseNumber(values[subject.aCol]);
        percentage = parseNumber(values[subject.pctCol ?? subject.columnIndex + 2]);
      } else if (subject.pctCol !== undefined) {
        percentage = parseNumber(values[subject.pctCol]);
      }

      if (percentage === null && conducted !== null && attended !== null && conducted > 0) {
        percentage = Math.round((attended / conducted) * 100);
      }

      student.attendance[subject.code] = {
        subjectName: subject.name,
        subjectCode: subject.code,
        teachingType: subject.teachingType,
        conducted,
        attended,
        percentage,
        className,
        tableIndex,
      };
    }

    students.push(student);
  }

  return {
    tableIndex,
    className,
    subjects,
    students,
    warnings,
  };
}

/**
 * Extracts attendance data directly from 2D array sheet rows (Microsoft Excel 97-2003 .xls & .xlsx).
 */
function extractFromSheetRows(rows: any[][], sheetName: string, tableIndex: number) {
  if (!rows || rows.length < 2) return null;

  // Extract Class Name if present in leading rows
  let className = sheetName || '';
  for (let r = 0; r < Math.min(rows.length, 5); r++) {
    const rowText = (rows[r] || []).map(cleanText).join(' ');
    const match =
      rowText.match(/Class Name:\s*(.*?)(?=Division:|$)/i) ||
      rowText.match(/Class:\s*(.*?)(?=$)/i);
    if (match) {
      className = cleanText(match[1]);
      break;
    }
  }

  // Find header row containing Roll / Roll No / Sr. / Student Name
  let headerIndex = -1;
  for (let r = 0; r < Math.min(rows.length, 15); r++) {
    const row = rows[r] || [];
    const hasHeaderKeyword = row.some((cell) => {
      const txt = cleanText(cell).toLowerCase();
      return (
        txt === 'sr.' ||
        txt === 'sr' ||
        txt === 's.no' ||
        txt === 's.no.' ||
        txt === 'roll no' ||
        txt === 'roll number' ||
        txt === 'rollno' ||
        txt === 'student name' ||
        txt === 'roll' ||
        txt === 'identifier'
      );
    });
    if (hasHeaderKeyword) {
      headerIndex = r;
      break;
    }
  }

  if (headerIndex === -1) {
    // If not found with keyword, check if row contains subject names and brackets
    for (let r = 0; r < Math.min(rows.length, 10); r++) {
      const row = rows[r] || [];
      const hasSubject = row.some((cell) => Boolean(parseSubjectInfo(cell)));
      if (hasSubject) {
        headerIndex = r;
        break;
      }
    }
  }

  if (headerIndex === -1) return null;

  const headerRow = rows[headerIndex] || [];
  const nextRow = rows[headerIndex + 1] || [];

  const hasSubheaders = nextRow.some((cell) => {
    const t = cleanText(cell).toUpperCase();
    return (
      t === 'D' ||
      t === 'A' ||
      t === '%' ||
      t === 'DELIVERED' ||
      t === 'ATTENDED' ||
      t === 'TOTAL' ||
      t === 'PRESENT'
    );
  });

  const dataStartIndex = hasSubheaders ? headerIndex + 2 : headerIndex + 1;

  let rollCol = -1;
  let nameCol = -1;
  let secCol = -1;
  let uidCol = -1;

  headerRow.forEach((cell, idx) => {
    const txt = cleanText(cell).toLowerCase();
    if (/^roll\s*(no|number)?$/i.test(txt) || txt === 'rollno' || txt === 'identifier') rollCol = idx;
    else if (/^(student\s*name|name|full\s*name)$/i.test(txt)) nameCol = idx;
    else if (/^(section|sec)$/i.test(txt)) secCol = idx;
    else if (/^(uid|unique\s*id)$/i.test(txt)) uidCol = idx;
  });

  // Fallbacks for rollCol if not explicitly named
  if (rollCol === -1) {
    if (headerRow.length > 3 && isStudentRollNumber(rows[dataStartIndex]?.[3])) rollCol = 3;
    else if (headerRow.length > 1 && isStudentRollNumber(rows[dataStartIndex]?.[1])) rollCol = 1;
    else if (isStudentRollNumber(rows[dataStartIndex]?.[0])) rollCol = 0;
  }
  if (nameCol === -1 && rollCol !== -1) {
    if (rows[dataStartIndex]?.[rollCol + 1]) nameCol = rollCol + 1;
  }

  // Detect subjects
  const subjects: TableSubject[] = [];
  let col = 0;
  while (col < headerRow.length) {
    if (col === rollCol || col === nameCol || col === secCol || col === uidCol) {
      col++;
      continue;
    }

    const cellText = cleanText(headerRow[col]);
    const subInfo = parseSubjectInfo(cellText);
    if (subInfo) {
      if (hasSubheaders) {
        subjects.push({
          ...subInfo,
          columnIndex: col,
          tableIndex,
          dCol: col,
          aCol: col + 1,
          pctCol: col + 2,
        });
        col += 3;
      } else {
        subjects.push({
          ...subInfo,
          columnIndex: col,
          tableIndex,
          dCol: -1,
          aCol: -1,
          pctCol: col,
        });
        col++;
      }
    } else {
      col++;
    }
  }

  if (subjects.length === 0 || rollCol === -1) return null;

  const students: Array<{
    rollNumber: string;
    name: string;
    section: string;
    uniqueId: string;
    className: string;
    attendance: Record<string, SubjectAttendanceRecord>;
  }> = [];

  for (let r = dataStartIndex; r < rows.length; r++) {
    const row = rows[r] || [];
    const rollNumber = cleanText(row[rollCol]);
    if (!isStudentRollNumber(rollNumber)) continue;

    const studentName = nameCol !== -1 ? cleanText(row[nameCol]) : '';
    const section = secCol !== -1 ? cleanText(row[secCol]) : '';
    const uniqueId = uidCol !== -1 ? cleanText(row[uidCol]) : '';

    const attendance: Record<string, SubjectAttendanceRecord> = {};
    for (const sub of subjects) {
      let conducted = sub.dCol !== undefined && sub.dCol !== -1 ? parseNumber(row[sub.dCol]) : null;
      let attended = sub.aCol !== undefined && sub.aCol !== -1 ? parseNumber(row[sub.aCol]) : null;
      let percentage = sub.pctCol !== undefined ? parseNumber(row[sub.pctCol]) : null;

      // Handle ratio format e.g. "18/20"
      if (percentage === null && sub.pctCol !== undefined) {
        const rawPct = cleanText(row[sub.pctCol]);
        const slashMatch = rawPct.match(/^(\d+)\s*\/\s*(\d+)$/);
        if (slashMatch) {
          attended = Number(slashMatch[1]);
          conducted = Number(slashMatch[2]);
          if (conducted > 0) percentage = Math.round((attended / conducted) * 100);
        }
      }

      if (percentage === null && conducted !== null && attended !== null && conducted > 0) {
        percentage = Math.round((attended / conducted) * 100);
      }

      attendance[sub.code] = {
        subjectName: sub.name,
        subjectCode: sub.code,
        teachingType: sub.teachingType,
        conducted,
        attended,
        percentage,
        className,
        tableIndex,
      };
    }

    students.push({
      rollNumber,
      name: studentName,
      section,
      uniqueId,
      className,
      attendance,
    });
  }

  return {
    tableIndex,
    className,
    subjects,
    students,
    warnings: [],
  };
}

/**
 * Universal attendance extractor:
 * Automatically detects and processes:
 * - Microsoft Excel 97-2003 Worksheet (.xls, BIFF8 format)
 * - Modern Excel (.xlsx)
 * - Legacy HTML attendance reports with .xls extension
 */
export async function extractAttendance(file: File): Promise<ExtractionResult> {
  if (!file) {
    throw new Error('Please upload an attendance file.');
  }

  const arrayBuffer = await file.arrayBuffer();
  if (!arrayBuffer || arrayBuffer.byteLength === 0) {
    throw new Error('The uploaded file is empty.');
  }

  const bytes = new Uint8Array(arrayBuffer);
  // Magic bytes check
  const isBiff8 =
    bytes.length >= 8 &&
    bytes[0] === 0xd0 &&
    bytes[1] === 0xcf &&
    bytes[2] === 0x11 &&
    bytes[3] === 0xe0 &&
    bytes[4] === 0xa1 &&
    bytes[5] === 0xb1 &&
    bytes[6] === 0x1a &&
    bytes[7] === 0xe1;

  const isZipXlsx =
    bytes.length >= 4 &&
    bytes[0] === 0x50 &&
    bytes[1] === 0x4b &&
    bytes[2] === 0x03 &&
    bytes[3] === 0x04;

  const isBinaryExcel = isBiff8 || isZipXlsx;
  const extractedTables: any[] = [];
  let fileFormatDetected = isBiff8
    ? 'Microsoft Excel 97-2003 Worksheet (.xls)'
    : isZipXlsx
    ? 'Excel Workbook (.xlsx)'
    : 'HTML / Text (.xls)';

  if (isBinaryExcel) {
    // Process binary Excel workbook (BIFF8 / XLSX)
    try {
      const workbook = XLSX.read(bytes, { type: 'array' });
      workbook.SheetNames.forEach((sheetName, index) => {
        const worksheet = workbook.Sheets[sheetName];
        if (!worksheet) return;

        // Try direct 2D array parsing
        const rawRows: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
        const sheetResult = extractFromSheetRows(rawRows, sheetName, index);
        if (sheetResult && sheetResult.students.length > 0) {
          extractedTables.push(sheetResult);
        } else {
          // Fallback to HTML table conversion if merged cells are complex
          try {
            const sheetHtml = XLSX.utils.sheet_to_html(worksheet);
            const parser = new DOMParser();
            const doc = parser.parseFromString(sheetHtml, 'text/html');
            const tableElem = doc.querySelector('table');
            if (tableElem) {
              const htmlResult = extractHtmlTable(tableElem as HTMLTableElement, index);
              if (htmlResult && htmlResult.students.length > 0) {
                extractedTables.push(htmlResult);
              }
            }
          } catch {
            // ignore fallback error
          }
        }
      });
    } catch (excelErr: any) {
      console.warn('XLSX binary parse warning:', excelErr);
    }
  } else {
    // Check if it's HTML text
    let textContent = '';
    try {
      const decoder = new TextDecoder('utf-8');
      textContent = decoder.decode(bytes);
    } catch {
      textContent = '';
    }

    const looksLikeHtml =
      /<!doctype|<html|<table/i.test(textContent.slice(0, 500)) ||
      textContent.includes('<table');

    if (looksLikeHtml && typeof DOMParser !== 'undefined') {
      try {
        const parser = new DOMParser();
        const document = parser.parseFromString(textContent, 'text/html');

        const tables = Array.from(document.querySelectorAll('table.mytable')) as HTMLTableElement[];
        if (!tables.length) {
          const fallbackTables = Array.from(document.querySelectorAll('table')) as HTMLTableElement[];
          tables.push(...fallbackTables);
        }

        tables.forEach((table, index) => {
          const result = extractHtmlTable(table, index);
          if (result && result.students.length > 0) {
            extractedTables.push(result);
          }
        });
      } catch (domErr) {
        console.warn('DOMParser warning:', domErr);
      }
    }

    // If HTML parsing found nothing or text isn't HTML, try SheetJS XLSX parser as secondary
    if (extractedTables.length === 0) {
      try {
        const workbook = XLSX.read(bytes, { type: 'array' });
        workbook.SheetNames.forEach((sheetName, index) => {
          const worksheet = workbook.Sheets[sheetName];
          if (!worksheet) return;
          const rawRows: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
          const sheetResult = extractFromSheetRows(rawRows, sheetName, index);
          if (sheetResult && sheetResult.students.length > 0) {
            extractedTables.push(sheetResult);
          }
        });
      } catch {
        // ignore secondary error
      }
    }
  }

  if (!extractedTables.length) {
    throw new Error(
      `Unable to extract attendance records from this file (${fileFormatDetected}). Please ensure the sheet includes Roll Number, Student Name, and Subject columns (e.g., "Mathematics (MATH101)" with D/A/% attendance metrics).`
    );
  }

  // Merge subjects by subject code
  const subjectMap = new Map<string, ExtractedSubject>();
  extractedTables.forEach((table) => {
    table.subjects.forEach((subject: TableSubject) => {
      if (!subjectMap.has(subject.code)) {
        subjectMap.set(subject.code, {
          code: subject.code,
          name: subject.name,
          teachingType: subject.teachingType,
        });
      }
    });
  });

  // Merge students by roll number
  const studentMap = new Map<string, ExtractedStudent>();
  extractedTables.forEach((table) => {
    table.students.forEach((student: any) => {
      const roll = student.rollNumber;
      if (!studentMap.has(roll)) {
        studentMap.set(roll, {
          rollNumber: roll,
          name: student.name,
          section: student.section,
          uniqueId: student.uniqueId,
          className: student.className,
          attendanceRecords: {},
        });
      }

      const existing = studentMap.get(roll)!;
      if (!existing.name && student.name) {
        existing.name = student.name;
      }
      if (!existing.section && student.section) {
        existing.section = student.section;
      }

      Object.entries(student.attendance).forEach(([subjectCode, attendance]: [string, any]) => {
        if (!existing.attendanceRecords[subjectCode]) {
          existing.attendanceRecords[subjectCode] = [];
        }
        existing.attendanceRecords[subjectCode].push({
          ...attendance,
          className: student.className,
          tableIndex: table.tableIndex,
        });
      });
    });
  });

  const students = Array.from(studentMap.values());
  const totalRecords = extractedTables.reduce((sum, table) => sum + table.students.length, 0);
  const duplicateRollNumbers = Array.from(studentMap.entries())
    .filter(([, student]) =>
      Object.values(student.attendanceRecords).some((records) => records.length > 1)
    )
    .map(([roll]) => roll);

  const diagnostics: ExtractionDiagnostics = {
    totalTablesFound: extractedTables.length,
    validTablesProcessed: extractedTables.length,
    totalUniqueSubjects: subjectMap.size,
    totalUniqueStudents: students.length,
    totalStudentRecords: totalRecords,
    duplicateRollNumbers,
    fileFormatDetected,
    tables: extractedTables.map((table) => ({
      tableIndex: table.tableIndex,
      className: table.className,
      subjectCount: table.subjects.length,
      studentCount: table.students.length,
      warnings: table.warnings || [],
    })),
  };

  return {
    subjects: Array.from(subjectMap.values()),
    students,
    diagnostics,
  };
}

/**
 * Deduplicates attendance records per student for selected subject codes,
 * selecting the record with highest attendance percentage and highest conducted classes.
 */
export function selectBestAttendance(
  data: { students: ExtractedStudent[] },
  selectedCodes: string[]
): BestAttendanceStudent[] {
  if (!data || !Array.isArray(data.students)) {
    throw new Error('Invalid attendance data.');
  }

  if (!Array.isArray(selectedCodes)) {
    throw new Error('Selected subject codes must be an array.');
  }

  return data.students.map((student) => {
    const result: BestAttendanceStudent = {
      rollNumber: student.rollNumber,
      name: student.name,
      section: student.section,
      uniqueId: student.uniqueId,
      attendance: {},
    };

    selectedCodes.forEach((code) => {
      const records = student.attendanceRecords[code] || [];
      const validRecords = records.filter(
        (record) => typeof record.percentage === 'number' && Number.isFinite(record.percentage)
      );

      if (!validRecords.length) {
        result.attendance[code] = null;
        return;
      }

      const bestRecord = validRecords.reduce((best, current) => {
        if ((current.percentage ?? 0) > (best.percentage ?? 0)) {
          return current;
        }
        if (
          current.percentage === best.percentage &&
          (current.conducted ?? -1) > (best.conducted ?? -1)
        ) {
          return current;
        }
        return best;
      });

      result.attendance[code] = {
        subjectCode: code,
        subjectName: bestRecord.subjectName,
        teachingType: bestRecord.teachingType,
        conducted: bestRecord.conducted,
        attended: bestRecord.attended,
        percentage: bestRecord.percentage,
        className: bestRecord.className,
        sourceTable: bestRecord.tableIndex,
      };
    });

    return result;
  });
}
