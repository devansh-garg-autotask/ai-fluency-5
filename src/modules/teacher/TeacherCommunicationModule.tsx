import React, { useState, useEffect, useMemo, useRef } from 'react';
import {
  Upload,
  FileSpreadsheet,
  Calendar,
  CheckSquare,
  Square,
  AlertTriangle,
  AlertCircle,
  CheckCircle2,
  XCircle,
  Search,
  Filter,
  Download,
  Trash2,
  Plus,
  RefreshCw,
  Edit3,
  Layers,
  HelpCircle,
  X,
  Phone,
  Mail,
  Award,
  BookOpen,
  Send,
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { StudentMaster } from '../../types/index.ts';
import * as api from '../../services/api.ts';
import { useAuth } from '../../context/AuthContext.tsx';
import { BulkCommunicationModule } from './BulkCommunicationModule.tsx';
import {
  extractAttendance,
  selectBestAttendance,
  ExtractedSubject,
  ExtractionResult,
} from '../../utils/attendanceExtractor.ts';
import {
  EvaluationColumn,
  extractEvaluationRecord,
  EvaluationExtractionResult,
} from '../../utils/evaluationExtractor.ts';
import { downloadEvaluationTemplate } from '../../utils/evaluationTemplate.ts';

// Type definitions for stacked batches
export type UploadFileType = 'attendance' | 'evaluation';

export interface AttendanceBatch {
  id: string;
  type: 'attendance';
  fileName: string;
  dateRange: {
    startDate: string;
    endDate: string;
    label: string;
  };
  subjects: ExtractedSubject[];
  uploadedAt: string;
}

export interface EvaluationBatchData {
  id: string;
  type: 'evaluation';
  fileName: string;
  subjectName: string;
  subjectCode: string;
  evaluations: EvaluationColumn[];
  uploadedAt: string;
}

export type StackedBatch = AttendanceBatch | EvaluationBatchData;

export interface UnifiedAttendanceCell {
  conducted: number | null;
  attended: number | null;
  percentage: number | null;
}

export interface UnifiedEvaluationCell {
  evaluationId: string;
  evaluationName: string;
  maxMarks: number;
  marksObtained: number | null;
  isAbsent: boolean;
  percentage: number | null;
}

export interface UnifiedStudentRow {
  rollNumber: string;
  name: string;
  section?: string;
  uniqueId?: string;
  // Key: `${batchId}_${subjectCode}`
  attendance: Record<string, UnifiedAttendanceCell>;
  // Key: `${batchId}_${evalId}`
  evaluations: Record<string, UnifiedEvaluationCell>;
}

export interface StudentMasterMatch {
  inMaster: boolean;
  hasStudentContact: boolean;
  hasParentContact: boolean;
  master: StudentMaster | null;
}

type EditingTarget =
  | {
      type: 'attendance';
      rollNumber: string;
      batchId: string;
      subjectCode: string;
      field: 'conducted' | 'attended' | 'percentage';
    }
  | {
      type: 'evaluation';
      rollNumber: string;
      batchId: string;
      evalId: string;
      maxMarks: number;
    };

export const TeacherCommunicationModule: React.FC = () => {
  const { user } = useAuth();

  // Master database records for cross-referencing
  const [masterStudents, setMasterStudents] = useState<StudentMaster[]>([]);
  const [isLoadingMaster, setIsLoadingMaster] = useState(true);

  // Stacked batches (both attendance & evaluation) and unified student rows
  const [batches, setBatches] = useState<StackedBatch[]>([]);
  const [students, setStudents] = useState<UnifiedStudentRow[]>([]);

  // Bulk Communication Mode and recipient selection state
  const [viewMode, setViewMode] = useState<'stack-report' | 'email-report'>('stack-report');
  const [isBulkModalOpen, setIsBulkModalOpen] = useState(false);
  const [sendToStudent, setSendToStudent] = useState(true);
  const [sendToParent, setSendToParent] = useState(true);

  // Search & Filter controls
  const [searchQuery, setSearchQuery] = useState('');
  const [statusFilter, setStatusFilter] = useState<
    'all' | 'missing-master' | 'missing-contact' | 'absent' | 'verified'
  >('all');

  // Unified Upload Modal State
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [selectedFileType, setSelectedFileType] = useState<UploadFileType>('attendance');

  // Attendance Wizard State
  const [startDate, setStartDate] = useState<string>(() => {
    const d = new Date();
    d.setDate(1);
    return d.toISOString().split('T')[0];
  });
  const [endDate, setEndDate] = useState<string>(() => new Date().toISOString().split('T')[0]);
  const [parsedAttendance, setParsedAttendance] = useState<ExtractionResult | null>(null);
  const [selectedSubjectCodes, setSelectedSubjectCodes] = useState<string[]>([]);

  // Evaluation Wizard State
  const [subjectName, setSubjectName] = useState('Data Structures & Algorithms');
  const [subjectCode, setSubjectCode] = useState('CS201');
  const [parsedEvaluation, setParsedEvaluation] = useState<EvaluationExtractionResult | null>(null);

  // Common upload state
  const [uploadedFile, setUploadedFile] = useState<File | null>(null);
  const [isParsing, setIsParsing] = useState(false);
  const [parseError, setParseError] = useState<string | null>(null);

  // Inline cell editing state
  const [editingTarget, setEditingTarget] = useState<EditingTarget | null>(null);
  const [editValue, setEditValue] = useState<string>('');

  const fileInputRef = useRef<HTMLInputElement>(null);

  // Load Master Student Database
  const loadMasterStudents = async () => {
    setIsLoadingMaster(true);
    try {
      const data = await api.getStudents();
      setMasterStudents(data);
    } catch (err) {
      console.error('Failed to load master students in communication tab:', err);
    } finally {
      setIsLoadingMaster(false);
    }
  };

  useEffect(() => {
    loadMasterStudents();
  }, []);

  // Quick lookup map for Master Students by Roll Number
  const masterStudentMap = useMemo(() => {
    const map = new Map<string, StudentMaster>();
    masterStudents.forEach((st) => {
      if (st.rollNumber) {
        map.set(st.rollNumber.trim().toLowerCase(), st);
      }
    });
    return map;
  }, [masterStudents]);

  // List of course names for quick subject selection in evaluation wizard
  const availableCourses = useMemo(() => {
    const set = new Set<string>();
    masterStudents.forEach((st) => {
      if (st.course) set.add(st.course.trim());
    });
    return Array.from(set);
  }, [masterStudents]);

  // Check student status against Master DB:
  // - inMaster: whether roll number exists
  // - hasStudentContact: whether phone or email is available
  // - hasParentContact: whether parent name, phone or email is available
  const getStudentStatus = (rollNumber: string): StudentMasterMatch => {
    const cleanRoll = rollNumber.trim().toLowerCase();
    const master = masterStudentMap.get(cleanRoll) || null;

    if (!master) {
      return {
        inMaster: false,
        hasStudentContact: false,
        hasParentContact: false,
        master: null,
      };
    }

    const hasStudentContact =
      (master.phoneNumbers && master.phoneNumbers.length > 0) ||
      (master.emailIds && master.emailIds.length > 0);

    const hasParentContact =
      Boolean(master.parentName && master.parentName.trim()) ||
      (master.parentPhoneNumbers && master.parentPhoneNumbers.length > 0) ||
      (master.parentEmailIds && master.parentEmailIds.length > 0);

    return {
      inMaster: true,
      hasStudentContact,
      hasParentContact,
      master,
    };
  };

  // Open upload wizard modal
  const handleOpenUpload = (defaultType?: UploadFileType) => {
    if (defaultType) setSelectedFileType(defaultType);
    setUploadedFile(null);
    setParseError(null);
    setParsedAttendance(null);
    setParsedEvaluation(null);
    setSelectedSubjectCodes([]);
    setIsModalOpen(true);
  };

  // Handle file selection and parsing based on selectedFileType
  const handleFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setUploadedFile(file);
    setIsParsing(true);
    setParseError(null);

    try {
      if (selectedFileType === 'attendance') {
        const result = await extractAttendance(file);
        if (!result.subjects || result.subjects.length === 0) {
          throw new Error('No attendance subjects found in this report.');
        }
        setParsedAttendance(result);
        setSelectedSubjectCodes(result.subjects.map((s) => s.code));
      } else {
        if (!subjectName.trim() || !subjectCode.trim()) {
          throw new Error('Please enter Subject Name and Subject Code before uploading evaluation file.');
        }
        const batchId = `eval_${Date.now()}`;
        const result = await extractEvaluationRecord(file, subjectName.trim(), subjectCode.trim(), batchId);
        setParsedEvaluation(result);
      }
    } catch (err: any) {
      console.error('File parsing error:', err);
      setParseError(err.message || 'Failed to parse file.');
    } finally {
      setIsParsing(false);
    }
  };

  // Finalize attendance batch and stack columns
  const handleFinalizeAttendanceBatch = () => {
    if (!parsedAttendance || selectedSubjectCodes.length === 0 || !uploadedFile) {
      setParseError('Please select at least one subject to extract.');
      return;
    }

    const batchId = `att_batch_${Date.now()}`;
    const dateLabel = startDate && endDate ? `${startDate} to ${endDate}` : 'Current Range';

    const selectedSubjects = parsedAttendance.subjects.filter((s) =>
      selectedSubjectCodes.includes(s.code)
    );

    const newBatch: AttendanceBatch = {
      id: batchId,
      type: 'attendance',
      fileName: uploadedFile.name,
      dateRange: {
        startDate,
        endDate,
        label: dateLabel,
      },
      subjects: selectedSubjects,
      uploadedAt: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
    };

    const bestData = selectBestAttendance(parsedAttendance, selectedSubjectCodes);

    setBatches((prev) => [...prev, newBatch]);

    // Merge student records by roll number
    setStudents((prev) => {
      const studentMap = new Map<string, UnifiedStudentRow>();

      prev.forEach((st) => {
        studentMap.set(st.rollNumber, {
          ...st,
          attendance: { ...st.attendance },
          evaluations: { ...st.evaluations },
        });
      });

      bestData.forEach((item) => {
        const roll = item.rollNumber;
        if (!studentMap.has(roll)) {
          studentMap.set(roll, {
            rollNumber: roll,
            name: item.name,
            section: item.section,
            uniqueId: item.uniqueId,
            attendance: {},
            evaluations: {},
          });
        }

        const student = studentMap.get(roll)!;
        if (!student.name && item.name) student.name = item.name;
        if (!student.section && item.section) student.section = item.section;

        selectedSubjectCodes.forEach((subCode) => {
          const rec = item.attendance[subCode];
          student.attendance[`${batchId}_${subCode}`] = {
            conducted: rec?.conducted ?? null,
            attended: rec?.attended ?? null,
            percentage: rec?.percentage ?? null,
          };
        });
      });

      return Array.from(studentMap.values());
    });

    setIsModalOpen(false);
    setUploadedFile(null);
    setParsedAttendance(null);
  };

  // Finalize evaluation batch and stack columns
  const handleFinalizeEvaluationBatch = () => {
    if (!parsedEvaluation || !uploadedFile) return;

    const batchId = `eval_batch_${Date.now()}`;
    const newBatch: EvaluationBatchData = {
      id: batchId,
      type: 'evaluation',
      fileName: uploadedFile.name,
      subjectName: parsedEvaluation.subjectName,
      subjectCode: parsedEvaluation.subjectCode,
      uploadedAt: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
      evaluations: parsedEvaluation.evaluations.map((ev) => ({
        ...ev,
        batchId,
      })),
    };

    setBatches((prev) => [...prev, newBatch]);

    // Merge student records by roll number
    setStudents((prev) => {
      const studentMap = new Map<string, UnifiedStudentRow>();

      prev.forEach((st) => {
        studentMap.set(st.rollNumber, {
          ...st,
          attendance: { ...st.attendance },
          evaluations: { ...st.evaluations },
        });
      });

      parsedEvaluation.students.forEach((item) => {
        const roll = item.rollNumber;
        if (!studentMap.has(roll)) {
          studentMap.set(roll, {
            rollNumber: roll,
            name: item.name,
            section: item.section,
            attendance: {},
            evaluations: {},
          });
        }

        const student = studentMap.get(roll)!;
        if (!student.name && item.name) student.name = item.name;
        if (!student.section && item.section) student.section = item.section;

        Object.entries(item.evaluations).forEach(([evalId, evalItem]) => {
          student.evaluations[`${batchId}_${evalId}`] = {
            evaluationId: evalItem.evaluationId,
            evaluationName: evalItem.evaluationName,
            maxMarks: evalItem.maxMarks,
            marksObtained: evalItem.marksObtained,
            isAbsent: evalItem.isAbsent,
            percentage: evalItem.percentage,
          };
        });
      });

      return Array.from(studentMap.values());
    });

    setIsModalOpen(false);
    setUploadedFile(null);
    setParsedEvaluation(null);
  };

  // Remove a batch and its columns
  const handleRemoveBatch = (batchId: string) => {
    setBatches((prev) => prev.filter((b) => b.id !== batchId));
    setStudents((prev) =>
      prev.map((s) => {
        const updatedAtt = { ...s.attendance };
        const updatedEval = { ...s.evaluations };

        Object.keys(updatedAtt).forEach((k) => {
          if (k.startsWith(`${batchId}_`)) delete updatedAtt[k];
        });
        Object.keys(updatedEval).forEach((k) => {
          if (k.startsWith(`${batchId}_`)) delete updatedEval[k];
        });

        return {
          ...s,
          attendance: updatedAtt,
          evaluations: updatedEval,
        };
      })
    );
  };

  // Clear all stacked batches
  const handleClearAll = () => {
    if (window.confirm('Clear all uploaded attendance and evaluation records?')) {
      setBatches([]);
      setStudents([]);
    }
  };

  // Inline Cell Editing Handlers
  const handleStartEditAttendance = (
    rollNumber: string,
    batchId: string,
    subjectCode: string,
    field: 'conducted' | 'attended' | 'percentage',
    currentVal: number | null
  ) => {
    setEditingTarget({
      type: 'attendance',
      rollNumber,
      batchId,
      subjectCode,
      field,
    });
    setEditValue(currentVal !== null ? String(currentVal) : '');
  };

  const handleSaveEditAttendance = () => {
    if (!editingTarget || editingTarget.type !== 'attendance') return;

    const { rollNumber, batchId, subjectCode, field } = editingTarget;
    const num = editValue.trim() === '' ? null : Number(editValue.trim());

    setStudents((prev) =>
      prev.map((s) => {
        if (s.rollNumber !== rollNumber) return s;
        const key = `${batchId}_${subjectCode}`;
        const current = s.attendance[key] || { conducted: null, attended: null, percentage: null };

        const updated = {
          ...current,
          [field]: Number.isFinite(num) ? num : null,
        };

        // If conducted and attended are present, auto recalculate percentage
        if (field !== 'percentage' && updated.conducted && updated.attended !== null && updated.conducted > 0) {
          updated.percentage = Math.round((updated.attended / updated.conducted) * 100);
        }

        return {
          ...s,
          attendance: {
            ...s.attendance,
            [key]: updated,
          },
        };
      })
    );

    setEditingTarget(null);
  };

  const handleStartEditEvaluation = (
    rollNumber: string,
    batchId: string,
    evalId: string,
    maxMarks: number,
    currentVal: number | null,
    isAbsent: boolean
  ) => {
    setEditingTarget({
      type: 'evaluation',
      rollNumber,
      batchId,
      evalId,
      maxMarks,
    });
    setEditValue(isAbsent ? 'A' : currentVal !== null ? String(currentVal) : '');
  };

  const handleSaveEditEvaluation = () => {
    if (!editingTarget || editingTarget.type !== 'evaluation') return;

    const { rollNumber, batchId, evalId, maxMarks } = editingTarget;
    const trimmed = editValue.trim().toUpperCase();
    const cellKey = `${batchId}_${evalId}`;

    setStudents((prev) =>
      prev.map((s) => {
        if (s.rollNumber !== rollNumber) return s;
        const current = s.evaluations[cellKey];

        if (!trimmed || trimmed === 'A' || trimmed === 'AB' || trimmed === 'ABSENT' || trimmed === '-') {
          return {
            ...s,
            evaluations: {
              ...s.evaluations,
              [cellKey]: {
                evaluationId: evalId,
                evaluationName: current?.evaluationName || 'Assessment',
                maxMarks,
                isAbsent: true,
                marksObtained: null,
                percentage: null,
              },
            },
          };
        }

        const num = Number(trimmed);
        if (Number.isFinite(num)) {
          const clamped = Math.max(0, num);
          const percentage = maxMarks > 0 ? Math.round((clamped / maxMarks) * 100 * 10) / 10 : 0;
          return {
            ...s,
            evaluations: {
              ...s.evaluations,
              [cellKey]: {
                evaluationId: evalId,
                evaluationName: current?.evaluationName || 'Assessment',
                maxMarks,
                isAbsent: false,
                marksObtained: clamped,
                percentage,
              },
            },
          };
        }

        return s;
      })
    );

    setEditingTarget(null);
  };

  // Filter students based on search and status
  const filteredStudents = useMemo(() => {
    return students.filter((s) => {
      const q = searchQuery.toLowerCase().trim();
      const matchesSearch =
        !q ||
        s.name.toLowerCase().includes(q) ||
        s.rollNumber.toLowerCase().includes(q);

      if (!matchesSearch) return false;

      const status = getStudentStatus(s.rollNumber);

      if (statusFilter === 'missing-master') {
        return !status.inMaster;
      }
      if (statusFilter === 'missing-contact') {
        return status.inMaster && (!status.hasStudentContact || !status.hasParentContact);
      }
      if (statusFilter === 'verified') {
        return status.inMaster && status.hasStudentContact && status.hasParentContact;
      }
      if (statusFilter === 'absent') {
        return Object.values(s.evaluations).some((ev) => ev.isAbsent);
      }

      return true;
    });
  }, [students, searchQuery, statusFilter, masterStudentMap]);

  // Overall statistics
  const summary = useMemo(() => {
    let missingMaster = 0;
    let missingContact = 0;
    let verified = 0;
    let absentCount = 0;

    students.forEach((s) => {
      const status = getStudentStatus(s.rollNumber);
      if (!status.inMaster) missingMaster++;
      else if (!status.hasStudentContact || !status.hasParentContact) missingContact++;
      else verified++;

      if (Object.values(s.evaluations).some((ev) => ev.isAbsent)) {
        absentCount++;
      }
    });

    const totalAttendanceSubjects = batches.reduce(
      (acc, b) => (b.type === 'attendance' ? acc + b.subjects.length : acc),
      0
    );
    const totalEvaluationAssessments = batches.reduce(
      (acc, b) => (b.type === 'evaluation' ? acc + b.evaluations.length : acc),
      0
    );

    return {
      total: students.length,
      missingMaster,
      missingContact,
      verified,
      absentCount,
      totalAttendanceSubjects,
      totalEvaluationAssessments,
    };
  }, [students, batches, masterStudentMap]);

  // Export combined stacked matrix to Excel
  const handleExportExcel = () => {
    if (students.length === 0) return;

    const rows = students.map((s) => {
      const status = getStudentStatus(s.rollNumber);
      const rowObj: any = {
        'Roll Number': s.rollNumber,
        'Student Name': s.name,
        'Master DB Status': status.inMaster ? 'In Master DB' : 'Missing from Master DB',
        'Student Contact': status.hasStudentContact ? 'Available' : 'Missing',
        'Parent Contact': status.hasParentContact ? 'Available' : 'Missing',
      };

      batches.forEach((b) => {
        if (b.type === 'attendance') {
          b.subjects.forEach((sub) => {
            const key = `${b.id}_${sub.code}`;
            const rec = s.attendance[key];
            const pfx = `[ATT ${b.dateRange.label}] ${sub.name} (${sub.code})`;
            rowObj[`${pfx} Total (D)`] = rec?.conducted ?? '—';
            rowObj[`${pfx} Attended (A)`] = rec?.attended ?? '—';
            rowObj[`${pfx} %`] = rec?.percentage !== null ? `${rec?.percentage}%` : '—';
          });
        } else {
          b.evaluations.forEach((ev) => {
            const cellKey = `${b.id}_${ev.id}`;
            const rec = s.evaluations[cellKey];
            const pfx = `[EVAL ${b.subjectCode}] ${ev.name} (${ev.maxMarks})`;
            if (!rec || rec.isAbsent) {
              rowObj[pfx] = 'ABSENT';
              rowObj[`${pfx} %`] = '0%';
            } else {
              rowObj[pfx] = rec.marksObtained;
              rowObj[`${pfx} %`] = `${rec.percentage}%`;
            }
          });
        }
      });

      return rowObj;
    });

    const worksheet = XLSX.utils.json_to_sheet(rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Academic_Records');
    XLSX.writeFile(
      workbook,
      `Student_Academic_Report_${new Date().toISOString().split('T')[0]}.xlsx`
    );
  };

  if (viewMode === 'email-report') {
    return (
      <BulkCommunicationModule
        batches={batches}
        students={students}
        masterStudentMap={masterStudentMap}
        teacherUser={user}
        recipientOptions={{
          sendToStudent,
          sendToParent,
        }}
        onBack={() => setViewMode('stack-report')}
      />
    );
  }

  return (
    <div className="space-y-6">
      {/* Top Banner & Control Bar */}
      <div className="bg-gradient-to-r from-slate-900 via-indigo-950/30 to-slate-900 border border-slate-800 rounded-2xl p-5 flex flex-col md:flex-row md:items-center justify-between gap-4 shadow-lg">
        <div className="flex items-center space-x-3.5">
          <div className="w-12 h-12 rounded-xl bg-indigo-950/80 border border-indigo-800/60 text-indigo-300 flex items-center justify-center shrink-0 shadow-lg shadow-indigo-950/50">
            <FileSpreadsheet className="w-6 h-6" />
          </div>
          <div>
            <div className="flex items-center space-x-2.5">
              <h2 className="text-lg font-bold text-white tracking-tight">Communication Portal</h2>
              <span className="text-[10px] px-2 py-0.5 rounded-full font-semibold bg-indigo-950 text-indigo-300 border border-indigo-800/60">
                Attendance &amp; Evaluation Stack
              </span>
            </div>
            <p className="text-xs text-slate-400 mt-0.5">
              Multi-file academic matrix: upload attendance &amp; evaluation records to stack dynamic columns, cross-verified with Student Master DB.
            </p>
          </div>
        </div>

        <div className="flex flex-wrap items-center gap-2.5">
          {/* Template Download Utility */}
          <div className="flex items-center space-x-1 bg-slate-950 p-1 rounded-xl border border-slate-800 text-xs">
            <button
              onClick={() => downloadEvaluationTemplate('xlsx', subjectName, subjectCode, masterStudents)}
              className="px-2.5 py-1.5 rounded-lg text-slate-300 hover:text-white hover:bg-slate-800 transition-colors flex items-center space-x-1.5 cursor-pointer"
              title="Download Evaluation template (.xlsx)"
            >
              <Download className="w-3.5 h-3.5 text-purple-400" />
              <span>Eval Template</span>
            </button>
            <button
              onClick={() => downloadEvaluationTemplate('xls', subjectName, subjectCode, masterStudents)}
              className="px-2 py-1.5 rounded-lg text-slate-400 hover:text-white hover:bg-slate-800 transition-colors font-mono text-[11px] cursor-pointer"
              title="Download Excel 97-2003 (.xls)"
            >
              .xls
            </button>
          </div>

          {/* Export to Excel */}
          {students.length > 0 && (
            <button
              onClick={handleExportExcel}
              className="px-3.5 py-2 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-700 text-xs font-semibold flex items-center space-x-1.5 transition-all cursor-pointer"
              title="Download consolidated stacked report as Excel"
            >
              <Download className="w-3.5 h-3.5 text-emerald-400" />
              <span>Export Grid</span>
            </button>
          )}

          {/* Initiate Bulk Communication */}
          {students.length > 0 && (
            <button
              onClick={() => setIsBulkModalOpen(true)}
              className="px-4 py-2.5 rounded-xl bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-500 hover:to-teal-500 text-white text-xs font-bold flex items-center space-x-2 shadow-lg shadow-emerald-600/30 transition-all cursor-pointer ring-1 ring-emerald-400/40"
              title="Initiate bulk customized academic communication to students and parents"
            >
              <Send className="w-4 h-4" />
              <span>Initiate Bulk Communication</span>
            </button>
          )}

          {/* Upload Button */}
          <button
            onClick={() => handleOpenUpload()}
            className="px-4 py-2.5 rounded-xl bg-indigo-600 hover:bg-indigo-500 text-white text-xs font-bold flex items-center space-x-2 shadow-lg shadow-indigo-600/30 transition-all cursor-pointer"
          >
            <Upload className="w-4 h-4" />
            <span>Upload File (Attendance / Eval)</span>
          </button>

          {/* Clear All */}
          {batches.length > 0 && (
            <button
              onClick={handleClearAll}
              className="p-2 rounded-xl bg-slate-950 hover:bg-rose-950/60 border border-slate-800 hover:border-rose-800/60 text-slate-400 hover:text-rose-300 transition-colors cursor-pointer"
              title="Clear all batches"
            >
              <Trash2 className="w-4 h-4" />
            </button>
          )}
        </div>
      </div>

      {/* Summary Metrics & Row Status Guidelines */}
      <div className="grid grid-cols-2 sm:grid-cols-5 gap-3">
        <div className="bg-slate-900/60 border border-slate-800 p-3.5 rounded-xl">
          <span className="text-[11px] text-slate-400 block mb-0.5">Total Students</span>
          <span className="text-xl font-bold text-slate-100">{summary.total}</span>
        </div>
        <div className="bg-slate-900/60 border border-slate-800 p-3.5 rounded-xl">
          <span className="text-[11px] text-slate-400 block mb-0.5">Attendance Subjects</span>
          <span className="text-xl font-bold text-indigo-400">{summary.totalAttendanceSubjects}</span>
        </div>
        <div className="bg-slate-900/60 border border-slate-800 p-3.5 rounded-xl">
          <span className="text-[11px] text-slate-400 block mb-0.5">Evaluation Assessments</span>
          <span className="text-xl font-bold text-purple-400">{summary.totalEvaluationAssessments}</span>
        </div>
        <div className="bg-rose-950/30 border border-rose-800/50 p-3.5 rounded-xl">
          <span className="text-[11px] text-rose-300 block mb-0.5">Not in Master (Red)</span>
          <span className="text-xl font-bold text-rose-400">{summary.missingMaster}</span>
        </div>
        <div className="bg-amber-950/30 border border-amber-800/50 p-3.5 rounded-xl">
          <span className="text-[11px] text-amber-300 block mb-0.5">Missing Contacts (Yellow)</span>
          <span className="text-xl font-bold text-amber-400">{summary.missingContact}</span>
        </div>
      </div>

      {/* Active Stacked Batches Badges (Shows both attendance & evaluation files) */}
      {batches.length > 0 && (
        <div className="p-3 bg-slate-900/60 border border-slate-800 rounded-xl flex flex-wrap items-center gap-2 text-xs">
          <span className="text-slate-400 text-[11px] font-semibold flex items-center space-x-1.5 pr-2">
            <Layers className="w-3.5 h-3.5 text-indigo-400" />
            <span>Stacked Reports ({batches.length}):</span>
          </span>

          {batches.map((b) => (
            <div
              key={b.id}
              className={`flex items-center space-x-2 px-3 py-1.5 rounded-lg border text-xs shadow-sm ${
                b.type === 'attendance'
                  ? 'bg-indigo-950/40 border-indigo-800/50 text-indigo-200'
                  : 'bg-purple-950/40 border-purple-800/50 text-purple-200'
              }`}
            >
              <div
                className={`w-2 h-2 rounded-full ${
                  b.type === 'attendance' ? 'bg-indigo-400' : 'bg-purple-400'
                }`}
              />
              <div className="flex flex-col">
                <span className="font-semibold text-[11px] truncate max-w-[200px]">
                  {b.type === 'attendance' ? `📅 [Att] ${b.fileName}` : `📝 [${b.subjectCode}] ${b.fileName}`}
                </span>
                <span className="text-[10px] text-slate-400 font-mono">
                  {b.type === 'attendance'
                    ? `${b.dateRange.label} • ${b.subjects.length} sub`
                    : `${b.subjectName} • ${b.evaluations.length} evals`}
                </span>
              </div>
              <button
                onClick={() => handleRemoveBatch(b.id)}
                title="Remove this batch and its columns"
                className="text-slate-500 hover:text-rose-400 p-1 rounded transition-colors cursor-pointer"
              >
                <X className="w-3.5 h-3.5" />
              </button>
            </div>
          ))}

          <button
            onClick={() => handleOpenUpload()}
            className="ml-auto flex items-center space-x-1 text-xs text-indigo-400 hover:text-indigo-300 font-medium px-2.5 py-1 rounded-lg border border-indigo-900/50 bg-indigo-950/30 cursor-pointer"
          >
            <Plus className="w-3.5 h-3.5" />
            <span>Upload More Files</span>
          </button>
        </div>
      )}

      {/* Legend & Filter Controls */}
      {students.length > 0 && (
        <div className="bg-slate-900/40 border border-slate-800 p-4 rounded-xl flex flex-col md:flex-row items-start md:items-center justify-between gap-4 text-xs">
          {/* Row Indicators Legend */}
          <div className="flex flex-wrap items-center gap-3">
            <span className="text-slate-400 font-semibold text-[11px]">Row Highlights:</span>
            <div className="flex items-center space-x-1.5 px-2.5 py-1 rounded-lg bg-rose-950/70 text-rose-300 border border-rose-800/60">
              <span className="w-2 h-2 rounded-full bg-rose-500 shrink-0" />
              <span>Red: Missing from Master DB ({summary.missingMaster})</span>
            </div>
            <div className="flex items-center space-x-1.5 px-2.5 py-1 rounded-lg bg-amber-950/70 text-amber-300 border border-amber-800/60">
              <span className="w-2 h-2 rounded-full bg-amber-500 shrink-0" />
              <span>Yellow: Contact Missing ({summary.missingContact})</span>
            </div>
            <div className="flex items-center space-x-1.5 px-2.5 py-1 rounded-lg bg-emerald-950/70 text-emerald-300 border border-emerald-800/60">
              <span className="w-2 h-2 rounded-full bg-emerald-500 shrink-0" />
              <span>Clean: Verified ({summary.verified})</span>
            </div>
          </div>

          {/* Search & Filter Chips */}
          <div className="flex flex-wrap items-center gap-2.5 w-full md:w-auto">
            <div className="relative flex-1 md:w-60">
              <Search className="w-3.5 h-3.5 text-slate-500 absolute left-3 top-1/2 -translate-y-1/2" />
              <input
                type="text"
                placeholder="Search name or roll..."
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                className="w-full pl-8 pr-3 py-1.5 bg-slate-950 border border-slate-800 rounded-xl text-xs text-slate-200 placeholder-slate-500 focus:outline-none focus:border-indigo-500"
              />
            </div>

            <div className="flex items-center space-x-1 bg-slate-950 p-1 rounded-xl border border-slate-800 text-xs">
              <button
                onClick={() => setStatusFilter('all')}
                className={`px-2.5 py-1 rounded-lg transition-all ${
                  statusFilter === 'all'
                    ? 'bg-indigo-600 text-white font-semibold'
                    : 'text-slate-400 hover:text-slate-200'
                }`}
              >
                All ({students.length})
              </button>
              <button
                onClick={() => setStatusFilter('missing-master')}
                className={`px-2 py-1 rounded-lg transition-all ${
                  statusFilter === 'missing-master'
                    ? 'bg-rose-600 text-white font-semibold'
                    : 'text-rose-400 hover:bg-rose-950/40'
                }`}
              >
                Red ({summary.missingMaster})
              </button>
              <button
                onClick={() => setStatusFilter('missing-contact')}
                className={`px-2 py-1 rounded-lg transition-all ${
                  statusFilter === 'missing-contact'
                    ? 'bg-amber-600 text-white font-semibold'
                    : 'text-amber-400 hover:bg-amber-950/40'
                }`}
              >
                Yellow ({summary.missingContact})
              </button>
              <button
                onClick={() => setStatusFilter('absent')}
                className={`px-2 py-1 rounded-lg transition-all ${
                  statusFilter === 'absent'
                    ? 'bg-purple-600 text-white font-semibold'
                    : 'text-purple-400 hover:bg-purple-950/40'
                }`}
              >
                Absent ({summary.absentCount})
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Main Unified Table (Stacked Side-by-Side) */}
      {students.length === 0 ? (
        <div className="bg-slate-900/40 border border-dashed border-slate-800 rounded-2xl p-12 text-center space-y-4">
          <div className="w-16 h-16 rounded-2xl bg-indigo-950/50 border border-indigo-800/40 text-indigo-400 flex items-center justify-center mx-auto">
            <FileSpreadsheet className="w-8 h-8" />
          </div>
          <div className="max-w-md mx-auto space-y-1.5">
            <h3 className="text-base font-bold text-slate-100">No Academic Records Uploaded Yet</h3>
            <p className="text-xs text-slate-400 leading-relaxed">
              Upload institutional attendance reports (.xls / HTML) or evaluation marks (.xlsx / .xls). Records are stacked alongside each other in one unified table, cross-verified with the Student Master Database.
            </p>
          </div>
          <div className="flex flex-wrap items-center justify-center gap-3 pt-2">
            <button
              onClick={() => handleOpenUpload('attendance')}
              className="px-5 py-2.5 rounded-xl bg-indigo-600 hover:bg-indigo-500 text-white text-xs font-semibold inline-flex items-center space-x-2 shadow-lg shadow-indigo-600/30 transition-all cursor-pointer"
            >
              <FileSpreadsheet className="w-4 h-4" />
              <span>Upload Attendance Record</span>
            </button>
            <button
              onClick={() => handleOpenUpload('evaluation')}
              className="px-5 py-2.5 rounded-xl bg-purple-600 hover:bg-purple-500 text-white text-xs font-semibold inline-flex items-center space-x-2 shadow-lg shadow-purple-600/30 transition-all cursor-pointer"
            >
              <Award className="w-4 h-4" />
              <span>Upload Evaluation Record</span>
            </button>
          </div>
        </div>
      ) : (
        <div className="bg-slate-900/60 border border-slate-800 rounded-2xl overflow-hidden shadow-xl">
          <div className="overflow-x-auto">
            <table className="w-full text-left text-xs border-collapse">
              <thead>
                {/* Super Header Row: Identity Columns + Stacked Batch Superheaders */}
                <tr className="bg-slate-950 border-b border-slate-800 text-slate-300">
                  <th
                    rowSpan={2}
                    className="px-4 py-3 border-r border-slate-800 font-bold uppercase text-[11px] min-w-[130px] sticky left-0 bg-slate-950 z-20"
                  >
                    Roll Number
                  </th>
                  <th
                    rowSpan={2}
                    className="px-4 py-3 border-r border-slate-800 font-bold uppercase text-[11px] min-w-[170px] sticky left-[130px] bg-slate-950 z-20"
                  >
                    Student Name
                  </th>
                  <th
                    rowSpan={2}
                    className="px-4 py-3 border-r border-slate-800 font-bold uppercase text-[11px] min-w-[210px]"
                  >
                    Master DB &amp; Contact Status
                  </th>

                  {/* Dynamic Stacked Batches: Attendance groups & Evaluation groups */}
                  {batches.map((b) => {
                    if (b.type === 'attendance') {
                      return b.subjects.map((sub) => (
                        <th
                          key={`${b.id}_${sub.code}`}
                          colSpan={3}
                          className="px-3 py-2 text-center border-r border-slate-800 bg-indigo-950/40 text-indigo-200 font-semibold text-xs border-b border-slate-800"
                        >
                          <div className="flex flex-col items-center justify-center space-y-0.5">
                            <div className="flex items-center space-x-1.5">
                              <span className="text-[10px] px-1.5 py-0.2 rounded bg-indigo-900/90 text-indigo-300 font-mono">
                                📅 {b.dateRange.label}
                              </span>
                            </div>
                            <span className="font-bold text-slate-100 truncate max-w-[200px]">
                              {sub.name} ({sub.code})
                            </span>
                            {sub.teachingType && (
                              <span className="text-[9px] px-1.5 rounded bg-indigo-900/60 text-indigo-300 font-mono">
                                Type: {sub.teachingType}
                              </span>
                            )}
                          </div>
                        </th>
                      ));
                    } else {
                      return (
                        <th
                          key={b.id}
                          colSpan={b.evaluations.length}
                          className="px-3 py-2 text-center border-r border-slate-800 bg-purple-950/40 text-purple-200 font-semibold text-xs border-b border-slate-800"
                        >
                          <div className="flex flex-col items-center justify-center space-y-0.5">
                            <div className="flex items-center space-x-1.5">
                              <span className="text-[10px] px-1.5 py-0.2 rounded bg-purple-900 text-purple-300 font-mono font-bold">
                                📝 [{b.subjectCode}]
                              </span>
                              <span className="font-bold text-slate-100 truncate max-w-[200px]">
                                {b.subjectName}
                              </span>
                            </div>
                            <span className="text-[10px] text-purple-400 font-mono">
                              {b.evaluations.length} assessment{b.evaluations.length > 1 ? 's' : ''}
                            </span>
                          </div>
                        </th>
                      );
                    }
                  })}
                </tr>

                {/* Second Header Row: Specific Columns per Batch */}
                <tr className="bg-slate-950/90 border-b border-slate-800 text-[10px] text-slate-400 uppercase font-mono">
                  {batches.map((b) => {
                    if (b.type === 'attendance') {
                      return b.subjects.map((sub) => (
                        <React.Fragment key={`${b.id}_${sub.code}_sub`}>
                          <th className="px-2 py-2 text-center border-r border-slate-800/60 min-w-[65px]" title="Total Days / Delivered">
                            Total (D)
                          </th>
                          <th className="px-2 py-2 text-center border-r border-slate-800/60 min-w-[65px]" title="Days Attended">
                            Attended (A)
                          </th>
                          <th className="px-2 py-2 text-center border-r border-slate-800 min-w-[70px]" title="Percentage">
                            %
                          </th>
                        </React.Fragment>
                      ));
                    } else {
                      return b.evaluations.map((ev) => (
                        <th
                          key={ev.id}
                          className="px-3 py-2 text-center border-r border-slate-800 min-w-[110px]"
                          title={ev.rawHeader}
                        >
                          <div className="font-semibold text-slate-200 normal-case text-xs">{ev.name}</div>
                          <div className="text-[9px] text-purple-400 font-mono">
                            Max: {ev.maxMarks}
                          </div>
                        </th>
                      ));
                    }
                  })}
                </tr>
              </thead>

              <tbody className="divide-y divide-slate-800/70">
                {filteredStudents.length === 0 ? (
                  <tr>
                    <td
                      colSpan={
                        3 +
                        batches.reduce(
                          (acc, b) =>
                            b.type === 'attendance'
                              ? acc + b.subjects.length * 3
                              : acc + b.evaluations.length,
                          0
                        )
                      }
                      className="px-6 py-10 text-center text-slate-500"
                    >
                      No student records match the active search or status filter.
                    </td>
                  </tr>
                ) : (
                  filteredStudents.map((st) => {
                    const status = getStudentStatus(st.rollNumber);

                    // Row color coding:
                    // 🔴 Red if not in Master DB
                    // 🟡 Yellow if contact details incomplete
                    // Normal otherwise
                    let rowStyle = 'hover:bg-slate-800/40 transition-colors';
                    if (!status.inMaster) {
                      rowStyle =
                        'bg-rose-950/30 border-l-4 border-l-rose-500 hover:bg-rose-950/45 text-rose-100 transition-colors';
                    } else if (!status.hasStudentContact || !status.hasParentContact) {
                      rowStyle =
                        'bg-amber-950/25 border-l-4 border-l-amber-500 hover:bg-amber-950/40 text-amber-100 transition-colors';
                    }

                    return (
                      <tr key={st.rollNumber} className={rowStyle}>
                        {/* Roll Number */}
                        <td className="px-4 py-3 font-mono font-bold text-indigo-300 border-r border-slate-800/60 sticky left-0 bg-slate-950/95 z-10 whitespace-nowrap">
                          {st.rollNumber}
                        </td>

                        {/* Student Name */}
                        <td className="px-4 py-3 font-medium text-slate-200 border-r border-slate-800/60 sticky left-[130px] bg-slate-950/95 z-10 whitespace-nowrap">
                          {st.name || status.master?.name || '—'}
                        </td>

                        {/* Master DB & Contact Status */}
                        <td className="px-4 py-3 border-r border-slate-800/60 align-top">
                          <div className="space-y-1.5">
                            {status.inMaster ? (
                              <div className="flex items-center space-x-1.5 text-[11px] text-emerald-400 font-medium">
                                <CheckCircle2 className="w-3.5 h-3.5 text-emerald-400 shrink-0" />
                                <span>Master DB Verified</span>
                              </div>
                            ) : (
                              <div className="flex items-center space-x-1.5 text-[11px] text-rose-400 font-semibold bg-rose-950/60 px-2 py-0.5 rounded border border-rose-800/70 w-fit">
                                <XCircle className="w-3.5 h-3.5 text-rose-400 shrink-0" />
                                <span>Missing from Master DB</span>
                              </div>
                            )}

                            <div className="flex items-center gap-2 text-[10px]">
                              <span
                                className={`px-1.5 py-0.5 rounded font-mono ${
                                  status.hasStudentContact
                                    ? 'bg-slate-800 text-slate-300'
                                    : 'bg-amber-950 text-amber-300 border border-amber-800/60 font-semibold'
                                }`}
                              >
                                Student: {status.hasStudentContact ? 'Contact OK' : 'No Contact'}
                              </span>
                              <span
                                className={`px-1.5 py-0.5 rounded font-mono ${
                                  status.hasParentContact
                                    ? 'bg-slate-800 text-slate-300'
                                    : 'bg-amber-950 text-amber-300 border border-amber-800/60 font-semibold'
                                }`}
                              >
                                Parent: {status.hasParentContact ? 'Contact OK' : 'No Contact'}
                              </span>
                            </div>
                          </div>
                        </td>

                        {/* Stacked Columns per Batch */}
                        {batches.map((b) => {
                          if (b.type === 'attendance') {
                            return b.subjects.map((sub) => {
                              const key = `${b.id}_${sub.code}`;
                              const record = st.attendance[key];

                              const isEditingCond =
                                editingTarget?.type === 'attendance' &&
                                editingTarget.rollNumber === st.rollNumber &&
                                editingTarget.batchId === b.id &&
                                editingTarget.subjectCode === sub.code &&
                                editingTarget.field === 'conducted';

                              const isEditingAtt =
                                editingTarget?.type === 'attendance' &&
                                editingTarget.rollNumber === st.rollNumber &&
                                editingTarget.batchId === b.id &&
                                editingTarget.subjectCode === sub.code &&
                                editingTarget.field === 'attended';

                              const isEditingPct =
                                editingTarget?.type === 'attendance' &&
                                editingTarget.rollNumber === st.rollNumber &&
                                editingTarget.batchId === b.id &&
                                editingTarget.subjectCode === sub.code &&
                                editingTarget.field === 'percentage';

                              return (
                                <React.Fragment key={key}>
                                  {/* Total Conducted (D) */}
                                  <td
                                    className="px-2 py-2 text-center border-r border-slate-800/60 font-mono text-xs cursor-pointer hover:bg-slate-800/40"
                                    onClick={() =>
                                      handleStartEditAttendance(
                                        st.rollNumber,
                                        b.id,
                                        sub.code,
                                        'conducted',
                                        record?.conducted ?? null
                                      )
                                    }
                                  >
                                    {isEditingCond ? (
                                      <input
                                        type="number"
                                        autoFocus
                                        value={editValue}
                                        onChange={(e) => setEditValue(e.target.value)}
                                        onKeyDown={(e) => {
                                          if (e.key === 'Enter') handleSaveEditAttendance();
                                          else if (e.key === 'Escape') setEditingTarget(null);
                                        }}
                                        onBlur={handleSaveEditAttendance}
                                        className="w-12 px-1 py-0.5 bg-slate-900 border border-indigo-500 rounded text-center text-xs text-white"
                                      />
                                    ) : (
                                      <span className="text-slate-300">
                                        {record?.conducted !== null && record?.conducted !== undefined
                                          ? record.conducted
                                          : '—'}
                                      </span>
                                    )}
                                  </td>

                                  {/* Attended (A) */}
                                  <td
                                    className="px-2 py-2 text-center border-r border-slate-800/60 font-mono text-xs cursor-pointer hover:bg-slate-800/40"
                                    onClick={() =>
                                      handleStartEditAttendance(
                                        st.rollNumber,
                                        b.id,
                                        sub.code,
                                        'attended',
                                        record?.attended ?? null
                                      )
                                    }
                                  >
                                    {isEditingAtt ? (
                                      <input
                                        type="number"
                                        autoFocus
                                        value={editValue}
                                        onChange={(e) => setEditValue(e.target.value)}
                                        onKeyDown={(e) => {
                                          if (e.key === 'Enter') handleSaveEditAttendance();
                                          else if (e.key === 'Escape') setEditingTarget(null);
                                        }}
                                        onBlur={handleSaveEditAttendance}
                                        className="w-12 px-1 py-0.5 bg-slate-900 border border-indigo-500 rounded text-center text-xs text-white"
                                      />
                                    ) : (
                                      <span className="text-slate-200 font-semibold">
                                        {record?.attended !== null && record?.attended !== undefined
                                          ? record.attended
                                          : '—'}
                                      </span>
                                    )}
                                  </td>

                                  {/* Percentage (%) */}
                                  <td
                                    className="px-2 py-2 text-center border-r border-slate-800 font-mono text-xs cursor-pointer hover:bg-slate-800/40"
                                    onClick={() =>
                                      handleStartEditAttendance(
                                        st.rollNumber,
                                        b.id,
                                        sub.code,
                                        'percentage',
                                        record?.percentage ?? null
                                      )
                                    }
                                  >
                                    {isEditingPct ? (
                                      <input
                                        type="number"
                                        autoFocus
                                        value={editValue}
                                        onChange={(e) => setEditValue(e.target.value)}
                                        onKeyDown={(e) => {
                                          if (e.key === 'Enter') handleSaveEditAttendance();
                                          else if (e.key === 'Escape') setEditingTarget(null);
                                        }}
                                        onBlur={handleSaveEditAttendance}
                                        className="w-12 px-1 py-0.5 bg-slate-900 border border-indigo-500 rounded text-center text-xs text-white"
                                      />
                                    ) : (
                                      <span
                                        className={
                                          record?.percentage !== null && record?.percentage !== undefined
                                            ? record.percentage >= 75
                                              ? 'text-emerald-400 font-semibold'
                                              : record.percentage >= 60
                                              ? 'text-amber-400'
                                              : 'text-rose-400 font-semibold'
                                            : 'text-slate-500'
                                        }
                                      >
                                        {record?.percentage !== null && record?.percentage !== undefined
                                          ? `${record.percentage}%`
                                          : '—'}
                                      </span>
                                    )}
                                  </td>
                                </React.Fragment>
                              );
                            });
                          } else {
                            // Evaluation Batch Columns
                            return b.evaluations.map((ev) => {
                              const cellKey = `${b.id}_${ev.id}`;
                              const evalItem = st.evaluations[cellKey];
                              const isEditing =
                                editingTarget?.type === 'evaluation' &&
                                editingTarget.rollNumber === st.rollNumber &&
                                editingTarget.batchId === b.id &&
                                editingTarget.evalId === ev.id;

                              if (isEditing) {
                                return (
                                  <td
                                    key={cellKey}
                                    className="px-2 py-2 text-center border-r border-slate-800 bg-purple-950/50"
                                  >
                                    <input
                                      type="text"
                                      autoFocus
                                      value={editValue}
                                      onChange={(e) => setEditValue(e.target.value)}
                                      onKeyDown={(e) => {
                                        if (e.key === 'Enter') handleSaveEditEvaluation();
                                        else if (e.key === 'Escape') setEditingTarget(null);
                                      }}
                                      onBlur={handleSaveEditEvaluation}
                                      className="w-16 px-1.5 py-1 bg-slate-900 border border-purple-500 rounded text-center text-xs text-white font-mono"
                                      placeholder="0 or A"
                                    />
                                    <span className="text-[9px] text-slate-400 block mt-0.5">
                                      'A' for Absent
                                    </span>
                                  </td>
                                );
                              }

                              if (!evalItem || evalItem.isAbsent) {
                                return (
                                  <td
                                    key={cellKey}
                                    onClick={() =>
                                      handleStartEditEvaluation(
                                        st.rollNumber,
                                        b.id,
                                        ev.id,
                                        ev.maxMarks,
                                        evalItem?.marksObtained ?? null,
                                        true
                                      )
                                    }
                                    className="px-2 py-2 text-center border-r border-slate-800 group cursor-pointer hover:bg-slate-800/50"
                                    title="Click to enter marks or toggle absent"
                                  >
                                    <div className="inline-flex items-center justify-center space-x-1 px-2 py-0.5 rounded bg-rose-950/80 text-rose-300 border border-rose-800/70 font-mono text-[10px] font-bold">
                                      <span>ABSENT</span>
                                      <Edit3 className="w-2.5 h-2.5 opacity-0 group-hover:opacity-100 text-rose-400 transition-opacity" />
                                    </div>
                                  </td>
                                );
                              }

                              const pct = evalItem.percentage ?? 0;
                              let markColor = 'text-slate-200';
                              if (pct >= 85) markColor = 'text-emerald-400 font-semibold';
                              else if (pct < 50) markColor = 'text-rose-400 font-semibold';

                              return (
                                <td
                                  key={cellKey}
                                  onClick={() =>
                                    handleStartEditEvaluation(
                                      st.rollNumber,
                                      b.id,
                                      ev.id,
                                      ev.maxMarks,
                                      evalItem.marksObtained,
                                      false
                                    )
                                  }
                                  className="px-3 py-2 text-center border-r border-slate-800 font-mono group cursor-pointer hover:bg-slate-800/50"
                                  title="Click to edit assessment marks"
                                >
                                  <div className="flex items-center justify-center space-x-1">
                                    <span className={`text-xs ${markColor}`}>
                                      {evalItem.marksObtained}
                                    </span>
                                    <span className="text-[10px] text-slate-500">/ {ev.maxMarks}</span>
                                    <Edit3 className="w-2.5 h-2.5 opacity-0 group-hover:opacity-100 text-purple-400 transition-opacity" />
                                  </div>
                                  <div className="text-[10px] text-slate-400">{pct}%</div>
                                </td>
                              );
                            });
                          }
                        })}
                      </tr>
                    );
                  })
                )}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* Unified Upload Wizard Modal */}
      {isModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-950/80 backdrop-blur-sm overflow-y-auto">
          <div className="bg-slate-900 border border-slate-800 rounded-2xl w-full max-w-2xl shadow-2xl overflow-hidden my-8">
            {/* Modal Header */}
            <div className="px-6 py-4 border-b border-slate-800 flex items-center justify-between bg-slate-900/60">
              <div className="flex items-center space-x-3">
                <div className="p-2 rounded-xl bg-indigo-950 border border-indigo-800/60 text-indigo-400">
                  <Upload className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="font-bold text-slate-100 text-sm">Upload Student Academic Record</h3>
                  <p className="text-xs text-slate-400">
                    Stack attendance or evaluation records alongside existing columns in the unified table.
                  </p>
                </div>
              </div>
              <button
                onClick={() => setIsModalOpen(false)}
                className="p-1.5 text-slate-400 hover:text-slate-100 rounded-lg hover:bg-slate-800 transition-colors cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Modal Body */}
            <div className="p-6 space-y-5">
              {parseError && (
                <div className="p-3.5 rounded-xl bg-rose-950/60 border border-rose-800/80 text-rose-300 text-xs flex items-start space-x-2.5">
                  <AlertCircle className="w-4 h-4 text-rose-400 shrink-0 mt-0.5" />
                  <span>{parseError}</span>
                </div>
              )}

              {/* Step 1: Select Type of File */}
              <div className="space-y-3">
                <label className="block text-xs font-bold text-slate-200">
                  1. Select Academic Record Classification *
                </label>
                <div className="grid grid-cols-2 gap-3">
                  <button
                    type="button"
                    onClick={() => {
                      setSelectedFileType('attendance');
                      setUploadedFile(null);
                      setParseError(null);
                    }}
                    className={`p-3.5 rounded-xl border text-left transition-all cursor-pointer ${
                      selectedFileType === 'attendance'
                        ? 'bg-indigo-950/60 border-indigo-600 text-white shadow-md shadow-indigo-600/20'
                        : 'bg-slate-950/60 border-slate-800 text-slate-400 hover:border-slate-700'
                    }`}
                  >
                    <div className="flex items-center space-x-2 font-semibold text-xs mb-1">
                      <FileSpreadsheet className="w-4 h-4 text-indigo-400" />
                      <span>Attendance Record</span>
                    </div>
                    <p className="text-[11px] text-slate-400 leading-tight">
                      Extract multi-table delivered &amp; attended counts per subject.
                    </p>
                  </button>

                  <button
                    type="button"
                    onClick={() => {
                      setSelectedFileType('evaluation');
                      setUploadedFile(null);
                      setParseError(null);
                    }}
                    className={`p-3.5 rounded-xl border text-left transition-all cursor-pointer ${
                      selectedFileType === 'evaluation'
                        ? 'bg-purple-950/60 border-purple-600 text-white shadow-md shadow-purple-600/20'
                        : 'bg-slate-950/60 border-slate-800 text-slate-400 hover:border-slate-700'
                    }`}
                  >
                    <div className="flex items-center justify-between mb-1">
                      <div className="flex items-center space-x-2 font-semibold text-xs text-purple-300">
                        <Award className="w-4 h-4 text-purple-400" />
                        <span>Evaluation Record</span>
                      </div>
                      <span className="text-[9px] px-1.5 py-0.5 rounded bg-purple-950 text-purple-300 border border-purple-800/60 font-mono">
                        Active
                      </span>
                    </div>
                    <p className="text-[11px] text-slate-400 leading-tight">
                      Assignment &amp; assessment marks, max marks from brackets, and absent tracking.
                    </p>
                  </button>
                </div>
              </div>

              {/* Step 2A: If Attendance Record Selected -> Date Range Configuration */}
              {selectedFileType === 'attendance' && (
                <div className="space-y-3 bg-slate-950/70 p-4 rounded-xl border border-slate-800 text-xs">
                  <div className="flex items-center justify-between">
                    <label className="font-bold text-slate-200 flex items-center space-x-1.5">
                      <Calendar className="w-3.5 h-3.5 text-indigo-400" />
                      <span>2. Attendance Period / Date Range *</span>
                    </label>
                    <span className="text-[10px] text-slate-400">Displayed on column group header</span>
                  </div>

                  <div className="grid grid-cols-2 gap-3">
                    <div>
                      <label className="text-[11px] text-slate-400 block mb-1">Start Date</label>
                      <input
                        type="date"
                        value={startDate}
                        onChange={(e) => setStartDate(e.target.value)}
                        className="w-full px-3 py-2 bg-slate-900 border border-slate-800 rounded-xl text-xs text-slate-100 font-mono focus:outline-none focus:border-indigo-500"
                      />
                    </div>
                    <div>
                      <label className="text-[11px] text-slate-400 block mb-1">End Date</label>
                      <input
                        type="date"
                        value={endDate}
                        onChange={(e) => setEndDate(e.target.value)}
                        className="w-full px-3 py-2 bg-slate-900 border border-slate-800 rounded-xl text-xs text-slate-100 font-mono focus:outline-none focus:border-indigo-500"
                      />
                    </div>
                  </div>
                </div>
              )}

              {/* Step 2B: If Evaluation Record Selected -> Subject details & Template Download */}
              {selectedFileType === 'evaluation' && (
                <div className="space-y-3 bg-slate-950/70 p-4 rounded-xl border border-slate-800 text-xs">
                  <label className="font-bold text-slate-200 block">
                    2. Subject Details &amp; Schema Template *
                  </label>

                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                    <div>
                      <label className="text-[11px] text-slate-400 block mb-1">Subject Name</label>
                      <input
                        type="text"
                        value={subjectName}
                        onChange={(e) => setSubjectName(e.target.value)}
                        placeholder="e.g. Data Structures"
                        className="w-full px-3 py-2 bg-slate-900 border border-slate-800 rounded-xl text-xs text-slate-100 focus:outline-none focus:border-purple-500"
                      />
                    </div>
                    <div>
                      <label className="text-[11px] text-slate-400 block mb-1">Subject Code</label>
                      <input
                        type="text"
                        value={subjectCode}
                        onChange={(e) => setSubjectCode(e.target.value)}
                        placeholder="e.g. CS201"
                        className="w-full px-3 py-2 bg-slate-900 border border-slate-800 rounded-xl text-xs text-slate-100 font-mono focus:outline-none focus:border-purple-500"
                      />
                    </div>
                  </div>

                  {availableCourses.length > 0 && (
                    <div className="pt-1">
                      <span className="text-[10px] text-slate-400 block mb-1">
                        Quick select from institutional programs:
                      </span>
                      <div className="flex flex-wrap gap-1.5">
                        {availableCourses.map((c) => (
                          <button
                            key={c}
                            type="button"
                            onClick={() => {
                              setSubjectName(c);
                              setSubjectCode(c.slice(0, 8).toUpperCase().replace(/[^A-Z0-9]/g, ''));
                            }}
                            className="px-2 py-0.5 rounded-lg bg-slate-900 hover:bg-slate-800 border border-slate-800 text-[10px] text-slate-300 transition-colors cursor-pointer"
                          >
                            {c}
                          </button>
                        ))}
                      </div>
                    </div>
                  )}

                  {/* Template Download Utility */}
                  <div className="pt-2 border-t border-slate-800 flex items-center justify-between">
                    <div>
                      <span className="font-semibold text-purple-200 block">Need the standard format?</span>
                      <span className="text-[10px] text-purple-300/80">
                        Pre-formatted with Student Name, Roll No, and assessment brackets.
                      </span>
                    </div>
                    <div className="flex items-center space-x-1.5">
                      <button
                        type="button"
                        onClick={() => downloadEvaluationTemplate('xlsx', subjectName, subjectCode, masterStudents)}
                        className="px-3 py-1.5 rounded-lg bg-purple-600 hover:bg-purple-500 text-white text-xs font-semibold inline-flex items-center space-x-1.5 cursor-pointer"
                      >
                        <Download className="w-3.5 h-3.5" />
                        <span>.xlsx Template</span>
                      </button>
                      <button
                        type="button"
                        onClick={() => downloadEvaluationTemplate('xls', subjectName, subjectCode, masterStudents)}
                        className="px-2.5 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-200 text-xs font-mono cursor-pointer"
                        title="Excel 97-2003 (.xls)"
                      >
                        .xls (97-2003)
                      </button>
                    </div>
                  </div>
                </div>
              )}

              {/* Step 3: Choose File */}
              <div className="space-y-2">
                <label className="block text-xs font-bold text-slate-200">
                  3. Choose {selectedFileType === 'attendance' ? 'Attendance' : 'Evaluation'} File (Excel .xlsx / Microsoft Excel 97-2003 .xls / HTML) *
                </label>

                <div
                  onClick={() => fileInputRef.current?.click()}
                  className="border-2 border-dashed border-slate-800 hover:border-indigo-500/70 rounded-xl p-5 text-center cursor-pointer bg-slate-950/50 hover:bg-slate-950/80 transition-all space-y-2"
                >
                  <input
                    ref={fileInputRef}
                    type="file"
                    accept=".xls,.xlsx,.html,.htm,.txt,.csv,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    onChange={handleFileChange}
                    className="hidden"
                  />
                  <div className="w-10 h-10 rounded-xl bg-slate-900 border border-slate-800 text-indigo-400 flex items-center justify-center mx-auto">
                    {isParsing ? (
                      <RefreshCw className="w-5 h-5 animate-spin text-indigo-400" />
                    ) : (
                      <Upload className="w-5 h-5" />
                    )}
                  </div>
                  <div>
                    <span className="text-xs font-semibold text-slate-200 block">
                      {uploadedFile ? uploadedFile.name : `Click to browse ${selectedFileType} file`}
                    </span>
                    <span className="text-[11px] text-slate-400">
                      Supports Microsoft Excel 97-2003 Worksheet (.xls), Modern Excel (.xlsx), and CSV
                    </span>
                    {selectedFileType === 'evaluation' && (
                      <span className="text-[10px] text-purple-300 block mt-1">
                        Note: Empty or blank marks cells are automatically registered as ABSENT
                      </span>
                    )}
                  </div>
                </div>
              </div>

              {/* Step 4A: Attendance Subject Selection */}
              {selectedFileType === 'attendance' && parsedAttendance && parsedAttendance.subjects.length > 0 && (
                <div className="space-y-3 pt-2 border-t border-slate-800">
                  <div className="flex items-center justify-between">
                    <div>
                      <span className="text-xs font-bold text-slate-200 block">
                        4. Select Subject(s) to Stack into Table
                      </span>
                      <span className="text-[11px] text-slate-400">
                        Discovered {parsedAttendance.subjects.length} unique subject code{parsedAttendance.subjects.length > 1 ? 's' : ''}
                      </span>
                    </div>

                    <div className="flex items-center space-x-2 text-[11px]">
                      <button
                        type="button"
                        onClick={() => setSelectedSubjectCodes(parsedAttendance.subjects.map((s) => s.code))}
                        className="text-indigo-400 hover:text-indigo-300 cursor-pointer font-medium"
                      >
                        Select All
                      </button>
                      <span className="text-slate-600">|</span>
                      <button
                        type="button"
                        onClick={() => setSelectedSubjectCodes([])}
                        className="text-slate-400 hover:text-slate-300 cursor-pointer"
                      >
                        Deselect All
                      </button>
                    </div>
                  </div>

                  <div className="max-h-48 overflow-y-auto space-y-1.5 pr-1 border border-slate-800/80 rounded-xl p-2 bg-slate-950/60">
                    {parsedAttendance.subjects.map((sub) => {
                      const isSelected = selectedSubjectCodes.includes(sub.code);
                      return (
                        <div
                          key={sub.code}
                          onClick={() =>
                            setSelectedSubjectCodes((prev) =>
                              prev.includes(sub.code) ? prev.filter((c) => c !== sub.code) : [...prev, sub.code]
                            )
                          }
                          className={`flex items-center justify-between p-2 rounded-lg border cursor-pointer text-xs transition-colors ${
                            isSelected
                              ? 'bg-indigo-950/50 border-indigo-700/80 text-white'
                              : 'bg-slate-900/40 border-slate-800 text-slate-400 hover:border-slate-700'
                          }`}
                        >
                          <div className="flex items-center space-x-2.5">
                            {isSelected ? (
                              <CheckSquare className="w-4 h-4 text-indigo-400" />
                            ) : (
                              <Square className="w-4 h-4 text-slate-600" />
                            )}
                            <span className="font-medium text-slate-200">{sub.name}</span>
                          </div>
                          <span className="px-2 py-0.5 rounded bg-slate-800 text-indigo-300 font-mono text-[10px]">
                            {sub.code}
                          </span>
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}

              {/* Step 4B: Evaluation Columns Preview */}
              {selectedFileType === 'evaluation' && parsedEvaluation && (
                <div className="bg-slate-950/70 p-4 rounded-xl border border-slate-800 space-y-2.5">
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-bold text-slate-200">
                      Discovered Assessment Columns ({parsedEvaluation.evaluations.length}):
                    </span>
                    <span className="text-[11px] text-purple-300 font-mono">
                      {parsedEvaluation.totalRows} student rows found
                    </span>
                  </div>

                  <div className="flex flex-wrap gap-1.5">
                    {parsedEvaluation.evaluations.map((ev) => (
                      <span
                        key={ev.id}
                        className="px-2.5 py-1 rounded-lg bg-purple-950/70 border border-purple-800/60 text-purple-300 text-xs font-mono"
                      >
                        {ev.name} (Max: {ev.maxMarks})
                      </span>
                    ))}
                  </div>
                </div>
              )}
            </div>

            {/* Modal Actions */}
            <div className="px-6 py-3.5 bg-slate-950 border-t border-slate-800 flex items-center justify-between">
              <button
                type="button"
                onClick={() => setIsModalOpen(false)}
                className="px-4 py-2 rounded-xl text-xs font-medium text-slate-400 hover:text-slate-200 hover:bg-slate-800 transition-colors cursor-pointer"
              >
                Cancel
              </button>

              {selectedFileType === 'attendance' ? (
                <button
                  type="button"
                  disabled={!parsedAttendance || selectedSubjectCodes.length === 0 || isParsing}
                  onClick={handleFinalizeAttendanceBatch}
                  className="px-5 py-2 rounded-xl bg-indigo-600 hover:bg-indigo-500 text-white text-xs font-semibold shadow-md shadow-indigo-600/20 disabled:opacity-50 cursor-pointer flex items-center space-x-1.5"
                >
                  {isParsing && <RefreshCw className="w-3.5 h-3.5 animate-spin" />}
                  <span>
                    {parsedAttendance ? `Stack ${selectedSubjectCodes.length} Attendance Subject(s)` : 'Extract Report'}
                  </span>
                </button>
              ) : (
                <button
                  type="button"
                  disabled={!parsedEvaluation || isParsing}
                  onClick={handleFinalizeEvaluationBatch}
                  className="px-5 py-2 rounded-xl bg-purple-600 hover:bg-purple-500 text-white text-xs font-semibold shadow-md shadow-purple-600/20 disabled:opacity-50 cursor-pointer flex items-center space-x-1.5"
                >
                  {isParsing && <RefreshCw className="w-3.5 h-3.5 animate-spin" />}
                  <span>
                    {parsedEvaluation
                      ? `Stack ${parsedEvaluation.evaluations.length} Evaluation Column(s)`
                      : 'Extract Evaluation'}
                  </span>
                </button>
              )}
            </div>
          </div>
        </div>
      )}

      {/* Recipient Selection Modal for Initiate Bulk Communication */}
      {isBulkModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/75 backdrop-blur-sm animate-in fade-in duration-200">
          <div className="bg-slate-900 border border-slate-700 rounded-2xl w-full max-w-lg shadow-2xl overflow-hidden flex flex-col">
            {/* Header */}
            <div className="px-6 py-4 border-b border-slate-800 flex items-center justify-between bg-slate-950/70">
              <div className="flex items-center space-x-3">
                <div className="w-10 h-10 rounded-xl bg-emerald-950/80 border border-emerald-800/60 text-emerald-400 flex items-center justify-center shrink-0 shadow-sm">
                  <Send className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="font-bold text-slate-100 text-base">Initiate Bulk Communication</h3>
                  <p className="text-xs text-slate-400">Select target recipient channels to generate email templates</p>
                </div>
              </div>
              <button
                onClick={() => setIsBulkModalOpen(false)}
                className="text-slate-400 hover:text-white p-1 rounded-lg hover:bg-slate-800 transition-colors cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Body */}
            <div className="p-6 space-y-5 text-sm">
              <div className="bg-slate-950/60 rounded-xl p-3.5 border border-slate-800 text-xs text-slate-300 space-y-1.5">
                <p className="font-semibold text-slate-200 flex items-center space-x-1.5">
                  <CheckCircle2 className="w-4 h-4 text-emerald-400 shrink-0" />
                  <span>Stack Report Loaded ({students.length} Students)</span>
                </p>
                <p className="text-slate-400 text-[11px] leading-relaxed">
                  The system will generate a consolidated Email Report retaining all stacked attendance and evaluation columns, plus customized email draft columns for selected recipients.
                </p>
              </div>

              <div>
                <label className="text-xs font-bold text-slate-300 block mb-2.5 uppercase tracking-wider">
                  Selection of Recipients
                </label>
                <div className="space-y-3">
                  {/* Student Checkbox */}
                  <label
                    onClick={() => setSendToStudent(!sendToStudent)}
                    className={`flex items-start space-x-3 p-3.5 rounded-xl border transition-all cursor-pointer select-none ${
                      sendToStudent
                        ? 'bg-emerald-950/30 border-emerald-500/50 text-emerald-100'
                        : 'bg-slate-950/40 border-slate-800 text-slate-400 hover:border-slate-700'
                    }`}
                  >
                    <div className="pt-0.5">
                      {sendToStudent ? (
                        <CheckSquare className="w-5 h-5 text-emerald-400" />
                      ) : (
                        <Square className="w-5 h-5 text-slate-500" />
                      )}
                    </div>
                    <div className="flex-1">
                      <span className="font-semibold text-sm text-slate-100 block">Send to Student Email ID</span>
                      <span className="text-xs text-slate-400 block mt-0.5">
                        Generates a <span className="text-emerald-300 font-medium">Student Email Template</span> column with informative, encouraging, advisory tone addressed directly to the student.
                      </span>
                    </div>
                  </label>

                  {/* Parent Checkbox */}
                  <label
                    onClick={() => setSendToParent(!sendToParent)}
                    className={`flex items-start space-x-3 p-3.5 rounded-xl border transition-all cursor-pointer select-none ${
                      sendToParent
                        ? 'bg-indigo-950/30 border-indigo-500/50 text-indigo-100'
                        : 'bg-slate-950/40 border-slate-800 text-slate-400 hover:border-slate-700'
                    }`}
                  >
                    <div className="pt-0.5">
                      {sendToParent ? (
                        <CheckSquare className="w-5 h-5 text-indigo-400" />
                      ) : (
                        <Square className="w-5 h-5 text-slate-500" />
                      )}
                    </div>
                    <div className="flex-1">
                      <span className="font-semibold text-sm text-slate-100 block">Send to Parent Email ID</span>
                      <span className="text-xs text-slate-400 block mt-0.5">
                        Generates a <span className="text-indigo-300 font-medium">Parent Email Template</span> column with formal, collaborative tone addressed to the parent/guardian.
                      </span>
                    </div>
                  </label>
                </div>
              </div>

              {!sendToStudent && !sendToParent && (
                <div className="text-xs text-amber-400 flex items-center space-x-2 bg-amber-950/30 border border-amber-800/40 p-3 rounded-xl">
                  <AlertTriangle className="w-4 h-4 shrink-0" />
                  <span>Please select at least one recipient option (Student or Parent) to proceed.</span>
                </div>
              )}
            </div>

            {/* Footer */}
            <div className="px-6 py-4 bg-slate-950 border-t border-slate-800 flex items-center justify-between">
              <button
                type="button"
                onClick={() => setIsBulkModalOpen(false)}
                className="px-4 py-2 rounded-xl text-xs font-medium text-slate-400 hover:text-slate-200 hover:bg-slate-800 transition-colors cursor-pointer"
              >
                Cancel
              </button>

              <button
                type="button"
                disabled={!sendToStudent && !sendToParent}
                onClick={() => {
                  setIsBulkModalOpen(false);
                  setViewMode('email-report');
                }}
                className="px-5 py-2.5 rounded-xl bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-500 hover:to-teal-500 text-white text-xs font-bold shadow-lg shadow-emerald-600/30 disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer flex items-center space-x-2"
              >
                <Send className="w-4 h-4" />
                <span>Initiate</span>
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
