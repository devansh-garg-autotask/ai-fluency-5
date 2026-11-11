/**
 * Evaluation Record Schema Template Generator
 *
 * Generates downloadable Excel templates (.xlsx and .xls) for Student Evaluation Records.
 * Columns:
 * - Student Name
 * - Roll Number
 * - Assessment columns with max marks in brackets: e.g. "Assignment 1 (20)", "Quiz 1 (10)", "Midterm Exam (50)"
 * Empty cells denote ABSENT student.
 */

import * as XLSX from 'xlsx';
import { StudentMaster } from '../types/index.ts';

export function createEvaluationTemplateWorkbook(
  students: StudentMaster[] = [],
  subjectName = 'Data Structures & Algorithms',
  subjectCode = 'CS201'
): XLSX.WorkBook {
  const sampleData =
    students.length > 0
      ? students.slice(0, 15).map((s, idx) => ({
          'Student Name': s.name,
          'Roll Number': s.rollNumber,
          'Assignment 1 (20)': idx === 2 ? '' : 18 - (idx % 4),
          'Quiz 1 (10)': idx === 3 ? '' : 9 - (idx % 2),
          'Midterm Exam (50)': idx === 4 ? '' : 44 - (idx % 6),
          'Lab Assessment (20)': 19 - (idx % 3),
        }))
      : [
          {
            'Student Name': 'Alex Rivera',
            'Roll Number': 'CS-2026-001',
            'Assignment 1 (20)': 18,
            'Quiz 1 (10)': 9,
            'Midterm Exam (50)': 45,
            'Lab Assessment (20)': 19,
          },
          {
            'Student Name': 'Emily Watson',
            'Roll Number': 'CS-2026-002',
            'Assignment 1 (20)': 20,
            'Quiz 1 (10)': 10,
            'Midterm Exam (50)': 48,
            'Lab Assessment (20)': 20,
          },
          {
            'Student Name': 'Marcus Chen',
            'Roll Number': 'CS-2026-003',
            'Assignment 1 (20)': 16,
            'Quiz 1 (10)': '', // Empty -> Absent
            'Midterm Exam (50)': 40,
            'Lab Assessment (20)': 17,
          },
          {
            'Student Name': 'Sophia Martinez',
            'Roll Number': 'CS-2026-004',
            'Assignment 1 (20)': 19,
            'Quiz 1 (10)': 8,
            'Midterm Exam (50)': '', // Empty -> Absent
            'Lab Assessment (20)': 18,
          },
        ];

  const guideData = [
    {
      'Column / Rule': 'Student Name',
      'Status': 'MANDATORY',
      'Format & Instructions': 'Full name of the student',
    },
    {
      'Column / Rule': 'Roll Number',
      'Status': 'MANDATORY (Identifier)',
      'Format & Instructions': 'Unique student roll number/identifier (e.g. CS-2026-001). Verified against Master DB.',
    },
    {
      'Column / Rule': 'Assessment Headers',
      'Status': 'DYNAMIC',
      'Format & Instructions':
        'Name of evaluation followed by maximum marks in parentheses, e.g. "Assignment 1 (20)" or "Midterm Exam (50)". You can add multiple evaluation columns.',
    },
    {
      'Column / Rule': 'Absent Students',
      'Status': 'RULE',
      'Format & Instructions':
        'LEAVE THE CELL BLANK / EMPTY (or write "A" / "Absent") if a student was absent for that evaluation.',
    },
    {
      'Column / Rule': 'File Compatibility',
      'Status': 'SUPPORTED',
      'Format & Instructions':
        'Supports Microsoft Excel 97-2003 Worksheet (.xls), Modern Excel (.xlsx), and CSV.',
    },
  ];

  const evalSheet = XLSX.utils.json_to_sheet(sampleData);
  evalSheet['!cols'] = [
    { wch: 22 }, // Student Name
    { wch: 18 }, // Roll Number
    { wch: 20 }, // Assignment 1 (20)
    { wch: 16 }, // Quiz 1 (10)
    { wch: 22 }, // Midterm Exam (50)
    { wch: 22 }, // Lab Assessment (20)
  ];

  const guideSheet = XLSX.utils.json_to_sheet(guideData);
  guideSheet['!cols'] = [{ wch: 24 }, { wch: 24 }, { wch: 75 }];

  const workbook = XLSX.utils.book_new();
  const safeSheetName = `${subjectCode.replace(/[^a-zA-Z0-9]/g, '_')}_Marks`.slice(0, 31);
  XLSX.utils.book_append_sheet(workbook, evalSheet, safeSheetName || 'Evaluation_Record');
  XLSX.utils.book_append_sheet(workbook, guideSheet, 'Instructions');

  return workbook;
}

/**
 * Direct client-side download utility for Evaluation Record template.
 */
export function downloadEvaluationTemplate(
  format: 'xlsx' | 'xls' = 'xlsx',
  subjectName = 'Data Structures & Algorithms',
  subjectCode = 'CS201',
  students: StudentMaster[] = []
): void {
  const workbook = createEvaluationTemplateWorkbook(students, subjectName, subjectCode);
  const codeSlug = (subjectCode || 'evaluation').toLowerCase().replace(/[^a-z0-9]/g, '_');
  const filename = `evaluation_record_template_${codeSlug}.${format}`;

  if (format === 'xls') {
    XLSX.writeFile(workbook, filename, { bookType: 'biff8' });
  } else {
    XLSX.writeFile(workbook, filename, { bookType: 'xlsx' });
  }
}
