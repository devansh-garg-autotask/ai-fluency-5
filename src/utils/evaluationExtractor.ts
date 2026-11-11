/**
 * Evaluation Record Extractor Utility
 *
 * Supports:
 * - Microsoft Excel 97-2003 Worksheet (.xls, BIFF8)
 * - Modern Excel (.xlsx)
 * - CSV files
 *
 * Logic:
 * - Identifies Student Name and Roll Number columns.
 * - Extracts evaluation/assignment columns with max marks from parentheses/brackets:
 *   e.g. "Assignment 1 (20)" -> name: "Assignment 1", maxMarks: 20
 * - Marks empty/blank cells or 'A'/'AB'/'Absent' as ABSENT.
 * - Computes percentages per assessment and tracks multi-file evaluations.
 */

import * as XLSX from 'xlsx';

export interface EvaluationColumn {
  id: string;
  name: string;
  maxMarks: number;
  rawHeader: string;
  columnIndex: number;
  subjectName?: string;
  subjectCode?: string;
  batchId?: string;
}

export interface StudentEvaluationItem {
  evaluationId: string;
  evaluationName: string;
  maxMarks: number;
  marksObtained: number | null;
  isAbsent: boolean;
  percentage: number | null;
  subjectName?: string;
  subjectCode?: string;
}

export interface ExtractedEvaluationStudent {
  rollNumber: string;
  name: string;
  section?: string;
  evaluations: Record<string, StudentEvaluationItem>; // keyed by evaluationId
}

export interface EvaluationExtractionResult {
  subjectName: string;
  subjectCode: string;
  evaluations: EvaluationColumn[];
  students: ExtractedEvaluationStudent[];
  fileFormatDetected: string;
  totalRows: number;
}

export function cleanText(val: any): string {
  return String(val ?? '')
    .replace(/\s+/g, ' ')
    .trim();
}

export function isStudentRollNumber(val: any): boolean {
  const cleaned = cleanText(val);
  if (!cleaned) return false;
  if (/^(total|average|sr\.?|roll|name|na|n\/a|absent|present)$/i.test(cleaned)) {
    return false;
  }
  return /^[A-Za-z0-9_-]{2,}$/.test(cleaned) && /[0-9]/.test(cleaned);
}

/**
 * Extracts evaluation title and maximum marks from column header:
 * e.g. "Assignment 1 (20)" -> { name: "Assignment 1", maxMarks: 20 }
 * e.g. "Quiz 1 [10]" -> { name: "Quiz 1", maxMarks: 10 }
 * e.g. "Midterm (50.5)" -> { name: "Midterm", maxMarks: 50.5 }
 */
export function parseEvaluationHeader(header: string): { name: string; maxMarks: number } | null {
  const cleaned = cleanText(header);
  if (!cleaned) return null;

  // Ignore student identity headers
  if (
    /^(sr\.?|s\.?\s*no\.?|roll|roll\s*no\.?|roll\s*number|rollno|name|student\s*name|student|section|sec|uid|unique\s*id|email|phone|contact)$/i.test(
      cleaned
    )
  ) {
    return null;
  }

  // Bracket extraction: e.g. "Assignment 1 (20)" or "Quiz [10]" or "Midterm Exam (50)"
  const bracketMatch = cleaned.match(/^(.*?)\s*[\(\[]\s*(\d+(?:\.\d+)?)\s*(?:marks|pts|m)?\s*[\)\]]$/i);
  if (bracketMatch) {
    const rawName = cleanText(bracketMatch[1]);
    const maxMarks = Number(bracketMatch[2]);
    if (Number.isFinite(maxMarks) && maxMarks > 0) {
      return {
        name: rawName || 'Assessment',
        maxMarks,
      };
    }
  }

  // Fallback: header ending with numbers or /20 or :20
  const slashMatch = cleaned.match(/^(.*?)\s*[/:]\s*(\d+(?:\.\d+)?)$/);
  if (slashMatch) {
    const rawName = cleanText(slashMatch[1]);
    const maxMarks = Number(slashMatch[2]);
    if (Number.isFinite(maxMarks) && maxMarks > 0) {
      return {
        name: rawName || 'Assessment',
        maxMarks,
      };
    }
  }

  return null;
}

/**
 * Parses cell value for evaluation score.
 * If empty, blank, or text like 'A', 'AB', 'Absent', returns isAbsent: true.
 */
export function parseEvaluationMarks(
  cellValue: any,
  maxMarks: number
): { marks: number | null; isAbsent: boolean; percentage: number | null } {
  const cleaned = cleanText(cellValue);

  // If empty or explicitly absent
  if (
    !cleaned ||
    cleaned === '-' ||
    cleaned === '—' ||
    /^(a|ab|abs|absent|na|n\/a)$/i.test(cleaned)
  ) {
    return {
      marks: null,
      isAbsent: true,
      percentage: null,
    };
  }

  const num = Number(cleaned.replace(/,/g, ''));
  if (Number.isFinite(num)) {
    const marks = Math.max(0, num);
    const percentage = maxMarks > 0 ? Math.round((marks / maxMarks) * 100 * 10) / 10 : 0;
    return {
      marks,
      isAbsent: false,
      percentage,
    };
  }

  // Unparseable non-numeric value treated as absent
  return {
    marks: null,
    isAbsent: true,
    percentage: null,
  };
}

/**
 * Main parser for Evaluation Record Excel files.
 */
export async function extractEvaluationRecord(
  file: File,
  subjectName: string,
  subjectCode: string,
  batchId = `batch_${Date.now()}`
): Promise<EvaluationExtractionResult> {
  if (!file) {
    throw new Error('Please select an Evaluation Record file.');
  }

  const arrayBuffer = await file.arrayBuffer();
  if (!arrayBuffer || arrayBuffer.byteLength === 0) {
    throw new Error('Uploaded Evaluation file is empty.');
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

  const fileFormatDetected = isBiff8
    ? 'Microsoft Excel 97-2003 Worksheet (.xls)'
    : isZipXlsx
    ? 'Excel Workbook (.xlsx)'
    : 'Spreadsheet / CSV';

  const workbook = XLSX.read(bytes, { type: 'array' });
  if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
    throw new Error('No worksheets found in the Excel workbook.');
  }

  // Look through sheets to find one with evaluation columns
  let bestSheetRows: any[][] = [];
  let bestHeaderIndex = -1;
  let bestRollCol = -1;
  let bestNameCol = -1;
  let detectedEvaluations: EvaluationColumn[] = [];

  for (const sheetName of workbook.SheetNames) {
    const ws = workbook.Sheets[sheetName];
    if (!ws) continue;

    const rows: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    if (!rows || rows.length < 2) continue;

    // Find header row containing Roll Number / Student Name
    for (let r = 0; r < Math.min(rows.length, 10); r++) {
      const row = rows[r] || [];
      let rollCol = -1;
      let nameCol = -1;
      const evals: EvaluationColumn[] = [];

      row.forEach((cell, colIdx) => {
        const text = cleanText(cell);
        if (/^roll\s*(no|number)?$/i.test(text) || text.toLowerCase() === 'rollno' || text.toLowerCase() === 'identifier') {
          rollCol = colIdx;
        } else if (/^(student\s*name|name|full\s*name)$/i.test(text)) {
          nameCol = colIdx;
        } else {
          const parsed = parseEvaluationHeader(text);
          if (parsed) {
            evals.push({
              id: `${batchId}_eval_${colIdx}_${parsed.name.replace(/[^a-zA-Z0-9]/g, '_')}`,
              name: parsed.name,
              maxMarks: parsed.maxMarks,
              rawHeader: text,
              columnIndex: colIdx,
              subjectName,
              subjectCode,
              batchId,
            });
          }
        }
      });

      // If we found roll number or name plus at least one evaluation column
      if (evals.length > 0) {
        if (rollCol === -1) {
          if (isStudentRollNumber(rows[r + 1]?.[1])) rollCol = 1;
          else if (isStudentRollNumber(rows[r + 1]?.[0])) rollCol = 0;
        }
        if (nameCol === -1 && rollCol !== -1) {
          nameCol = rollCol === 0 ? 1 : 0;
        }

        bestSheetRows = rows;
        bestHeaderIndex = r;
        bestRollCol = rollCol;
        bestNameCol = nameCol;
        detectedEvaluations = evals;
        break;
      }
    }

    if (detectedEvaluations.length > 0) break;
  }

  if (detectedEvaluations.length === 0) {
    throw new Error(
      `No evaluation columns detected in "${file.name}". Please ensure evaluation headers include maximum marks in parentheses, e.g. "Assignment 1 (20)" or "Midterm Exam (50)".`
    );
  }

  if (bestRollCol === -1) {
    throw new Error(
      `Could not locate the "Roll Number" column in "${file.name}". Please include a column labeled "Roll Number" or "Roll No".`
    );
  }

  // Extract student evaluation marks
  const studentMap = new Map<string, ExtractedEvaluationStudent>();

  for (let r = bestHeaderIndex + 1; r < bestSheetRows.length; r++) {
    const row = bestSheetRows[r] || [];
    const rollNumber = cleanText(row[bestRollCol]);
    if (!isStudentRollNumber(rollNumber)) continue;

    const studentName = bestNameCol !== -1 ? cleanText(row[bestNameCol]) : '';

    if (!studentMap.has(rollNumber)) {
      studentMap.set(rollNumber, {
        rollNumber,
        name: studentName,
        evaluations: {},
      });
    }

    const student = studentMap.get(rollNumber)!;
    if (!student.name && studentName) student.name = studentName;

    detectedEvaluations.forEach((evalCol) => {
      const cellVal = row[evalCol.columnIndex];
      const parsedMarks = parseEvaluationMarks(cellVal, evalCol.maxMarks);

      student.evaluations[evalCol.id] = {
        evaluationId: evalCol.id,
        evaluationName: evalCol.name,
        maxMarks: evalCol.maxMarks,
        marksObtained: parsedMarks.marks,
        isAbsent: parsedMarks.isAbsent,
        percentage: parsedMarks.percentage,
        subjectName,
        subjectCode,
      };
    });
  }

  const students = Array.from(studentMap.values());
  if (students.length === 0) {
    throw new Error(
      `Evaluation columns were detected, but no valid student rows were found below row ${bestHeaderIndex + 1}. Check that roll numbers are present.`
    );
  }

  return {
    subjectName,
    subjectCode,
    evaluations: detectedEvaluations,
    students,
    fileFormatDetected,
    totalRows: students.length,
  };
}
