import React, { useState, useEffect, useMemo, useRef } from 'react';
import {
  Upload,
  FileSpreadsheet,
  Download,
  AlertTriangle,
  AlertCircle,
  CheckCircle2,
  XCircle,
  Search,
  Filter,
  Plus,
  Trash2,
  RefreshCw,
  Edit3,
  Layers,
  HelpCircle,
  X,
  Phone,
  Mail,
  Award,
  BookOpen,
  Calendar,
  Sparkles,
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { StudentMaster } from '../../types/index.ts';
import * as api from '../../services/api.ts';
import {
  EvaluationColumn,
  ExtractedEvaluationStudent,
  EvaluationExtractionResult,
  extractEvaluationRecord,
} from '../../utils/evaluationExtractor.ts';
import { downloadEvaluationTemplate } from '../../utils/evaluationTemplate.ts';

export interface EvaluationBatch {
  id: string;
  fileName: string;
  subjectName: string;
  subjectCode: string;
  uploadedAt: string;
  evaluations: EvaluationColumn[];
}

export interface StudentMasterEvaluationMatch {
  inMaster: boolean;
  hasStudentContact: boolean;
  hasParentContact: boolean;
  master: StudentMaster | null;
}

export const TeacherEvaluationModule: React.FC = () => {
  // Master database state
  const [masterStudents, setMasterStudents] = useState<StudentMaster[]>([]);
  const [isLoadingMaster, setIsLoadingMaster] = useState<boolean>(true);
  const [masterLoadError, setMasterLoadError] = useState<string | null>(null);

  // Uploaded batches and extracted student records
  const [batches, setBatches] = useState<EvaluationBatch[]>([]);
  const [students, setStudents] = useState<ExtractedEvaluationStudent[]>([]);

  // Search & Filtering
  const [searchQuery, setSearchQuery] = useState('');
  const [statusFilter, setStatusFilter] = useState<'all' | 'red' | 'yellow' | 'absent'>('all');

  // Upload Wizard Modal State
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [subjectName, setSubjectName] = useState('Data Structures & Algorithms');
  const [subjectCode, setSubjectCode] = useState('CS201');
  const [uploadedFile, setUploadedFile] = useState<File | null>(null);
  const [isParsing, setIsParsing] = useState(false);
  const [parseError, setParseError] = useState<string | null>(null);
  const [parsedResult, setParsedResult] = useState<EvaluationExtractionResult | null>(null);

  // Inline editing state: { studentRoll, evalId }
  const [editingCell, setEditingCell] = useState<{ rollNumber: string; evalId: string } | null>(null);
  const [editValue, setEditValue] = useState<string>('');

  const fileInputRef = useRef<HTMLInputElement | null>(null);

  // Load persistent student master database for cross-referencing
  useEffect(() => {
    const fetchMaster = async () => {
      setIsLoadingMaster(true);
      try {
        const data = await api.getStudents();
        setMasterStudents(data);
      } catch (err: any) {
        console.error('Failed to load Student Master for evaluation check:', err);
        setMasterLoadError('Failed to synchronize with Student Master Database.');
      } finally {
        setIsLoadingMaster(false);
      }
    };
    fetchMaster();
  }, []);

  // Map master students by roll number for fast lookup
  const masterMap = useMemo(() => {
    const map = new Map<string, StudentMaster>();
    masterStudents.forEach((st) => {
      if (st.rollNumber) {
        map.set(st.rollNumber.trim().toLowerCase(), st);
      }
    });
    return map;
  }, [masterStudents]);

  // Extract unique subjects from master database for easy selection dropdown
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
  const getStudentStatus = (rollNumber: string): StudentMasterEvaluationMatch => {
    const cleanRoll = rollNumber.trim().toLowerCase();
    const master = masterMap.get(cleanRoll) || null;

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

  // Open Upload Modal
  const handleOpenUpload = () => {
    setUploadedFile(null);
    setParseError(null);
    setParsedResult(null);
    setIsModalOpen(true);
  };

  // Handle file selection in modal
  const handleFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    if (!subjectName.trim() || !subjectCode.trim()) {
      setParseError('Please specify both Subject Name and Subject Code before uploading.');
      return;
    }

    setUploadedFile(file);
    setIsParsing(true);
    setParseError(null);

    try {
      const batchId = `eval_batch_${Date.now()}`;
      const result = await extractEvaluationRecord(file, subjectName.trim(), subjectCode.trim(), batchId);
      setParsedResult(result);
    } catch (err: any) {
      console.error('Evaluation file extraction error:', err);
      setParseError(err.message || 'Failed to extract evaluation records.');
    } finally {
      setIsParsing(false);
    }
  };

  // Commit parsed batch to state
  const handleCommitBatch = () => {
    if (!parsedResult || !uploadedFile) return;

    const batchId = parsedResult.evaluations[0]?.batchId || `eval_batch_${Date.now()}`;
    const newBatch: EvaluationBatch = {
      id: batchId,
      fileName: uploadedFile.name,
      subjectName: parsedResult.subjectName,
      subjectCode: parsedResult.subjectCode,
      uploadedAt: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
      evaluations: parsedResult.evaluations,
    };

    setBatches((prev) => [...prev, newBatch]);

    // Merge student records by rollNumber
    setStudents((prev) => {
      const studentMap = new Map<string, ExtractedEvaluationStudent>();

      // Preserve existing
      prev.forEach((s) => {
        studentMap.set(s.rollNumber, {
          ...s,
          evaluations: { ...s.evaluations },
        });
      });

      // Merge new
      parsedResult.students.forEach((newS) => {
        if (!studentMap.has(newS.rollNumber)) {
          studentMap.set(newS.rollNumber, {
            rollNumber: newS.rollNumber,
            name: newS.name,
            section: newS.section,
            evaluations: {},
          });
        }

        const existing = studentMap.get(newS.rollNumber)!;
        if (!existing.name && newS.name) existing.name = newS.name;
        if (!existing.section && newS.section) existing.section = newS.section;

        Object.entries(newS.evaluations).forEach(([evalId, evalItem]) => {
          existing.evaluations[evalId] = evalItem;
        });
      });

      return Array.from(studentMap.values());
    });

    setIsModalOpen(false);
    setUploadedFile(null);
    setParsedResult(null);
  };

  // Remove a batch
  const handleRemoveBatch = (batchId: string) => {
    const batchToRemove = batches.find((b) => b.id === batchId);
    if (!batchToRemove) return;

    const evalIdsToRemove = new Set(batchToRemove.evaluations.map((e) => e.id));

    setBatches((prev) => prev.filter((b) => b.id !== batchId));
    setStudents((prev) =>
      prev.map((s) => {
        const updatedEvals = { ...s.evaluations };
        evalIdsToRemove.forEach((id) => delete updatedEvals[id]);
        return {
          ...s,
          evaluations: updatedEvals,
        };
      })
    );
  };

  // Clear all evaluation batches
  const handleClearAll = () => {
    if (window.confirm('Clear all uploaded evaluation records and columns?')) {
      setBatches([]);
      setStudents([]);
    }
  };

  // Inline Cell Editing
  const handleStartEdit = (rollNumber: string, evalId: string, currentVal: number | null, isAbsent: boolean) => {
    setEditingCell({ rollNumber, evalId });
    setEditValue(isAbsent ? 'A' : currentVal !== null ? String(currentVal) : '');
  };

  const handleSaveEdit = (rollNumber: string, evalId: string, maxMarks: number) => {
    const trimmed = editValue.trim().toUpperCase();

    setStudents((prev) =>
      prev.map((s) => {
        if (s.rollNumber !== rollNumber) return s;

        const currentEval = s.evaluations[evalId];
        if (!currentEval) return s;

        if (!trimmed || trimmed === 'A' || trimmed === 'AB' || trimmed === 'ABSENT' || trimmed === '-') {
          return {
            ...s,
            evaluations: {
              ...s.evaluations,
              [evalId]: {
                ...currentEval,
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
              [evalId]: {
                ...currentEval,
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

    setEditingCell(null);
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

      if (statusFilter === 'red') {
        return !status.inMaster;
      }
      if (statusFilter === 'yellow') {
        return status.inMaster && (!status.hasStudentContact || !status.hasParentContact);
      }
      if (statusFilter === 'absent') {
        return Object.values(s.evaluations).some((ev) => ev.isAbsent);
      }

      return true;
    });
  }, [students, searchQuery, statusFilter, masterMap]);

  // Overall statistics
  const stats = useMemo(() => {
    let redCount = 0;
    let yellowCount = 0;
    let absentInstances = 0;

    students.forEach((s) => {
      const status = getStudentStatus(s.rollNumber);
      if (!status.inMaster) redCount++;
      else if (!status.hasStudentContact || !status.hasParentContact) yellowCount++;

      Object.values(s.evaluations).forEach((ev) => {
        if (ev.isAbsent) absentInstances++;
      });
    });

    const totalEvals = batches.reduce((acc, b) => acc + b.evaluations.length, 0);

    return {
      totalStudents: students.length,
      redCount,
      yellowCount,
      totalEvals,
      absentInstances,
    };
  }, [students, batches, masterMap]);

  // Export current evaluation grid to Excel
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
        b.evaluations.forEach((ev) => {
          const record = s.evaluations[ev.id];
          const colHeader = `${b.subjectCode} - ${ev.name} (${ev.maxMarks})`;
          if (!record || record.isAbsent) {
            rowObj[colHeader] = 'ABSENT';
            rowObj[`${colHeader} %`] = '0%';
          } else {
            rowObj[colHeader] = record.marksObtained;
            rowObj[`${colHeader} %`] = `${record.percentage}%`;
          }
        });
      });

      return rowObj;
    });

    const worksheet = XLSX.utils.json_to_sheet(rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Evaluation_Records');
    XLSX.writeFile(
      workbook,
      `Student_Evaluation_Report_${new Date().toISOString().split('T')[0]}.xlsx`
    );
  };

  return (
    <div className="space-y-6">
      {/* Top Action Header */}
      <div className="bg-slate-900/60 border border-slate-800 rounded-2xl p-5 flex flex-col md:flex-row md:items-center justify-between gap-4 shadow-lg">
        <div>
          <div className="flex items-center space-x-2">
            <div className="p-2 rounded-xl bg-purple-950/80 border border-purple-800/60 text-purple-400">
              <Award className="w-5 h-5" />
            </div>
            <div>
              <h2 className="text-base font-bold text-slate-100">Student Evaluation Records</h2>
              <p className="text-xs text-slate-400">
                Track assessment assignments, maximum marks, absent entries, and verify against Master Database.
              </p>
            </div>
          </div>
        </div>

        <div className="flex flex-wrap items-center gap-2.5">
          {/* Download Template Buttons */}
          <div className="flex items-center space-x-1 bg-slate-950 p-1 rounded-xl border border-slate-800 text-xs">
            <button
              onClick={() => downloadEvaluationTemplate('xlsx', subjectName, subjectCode, masterStudents)}
              className="px-2.5 py-1.5 rounded-lg text-slate-300 hover:text-white hover:bg-slate-800 transition-colors flex items-center space-x-1.5 cursor-pointer"
              title="Download modern Excel template (.xlsx)"
            >
              <Download className="w-3.5 h-3.5 text-indigo-400" />
              <span>Template (.xlsx)</span>
            </button>
            <button
              onClick={() => downloadEvaluationTemplate('xls', subjectName, subjectCode, masterStudents)}
              className="px-2 py-1.5 rounded-lg text-slate-400 hover:text-white hover:bg-slate-800 transition-colors font-mono text-[11px] cursor-pointer"
              title="Download Microsoft Excel 97-2003 Worksheet (.xls)"
            >
              .xls (97-2003)
            </button>
          </div>

          {/* Upload Button */}
          <button
            onClick={handleOpenUpload}
            className="px-4 py-2 rounded-xl bg-purple-600 hover:bg-purple-500 text-white text-xs font-semibold flex items-center space-x-2 shadow-md shadow-purple-600/20 transition-all cursor-pointer"
          >
            <Plus className="w-4 h-4" />
            <span>Upload Evaluation Record</span>
          </button>

          {/* Export to Excel */}
          {students.length > 0 && (
            <button
              onClick={handleExportExcel}
              className="px-3.5 py-2 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-200 text-xs font-medium flex items-center space-x-1.5 transition-colors cursor-pointer"
            >
              <Download className="w-3.5 h-3.5 text-emerald-400" />
              <span>Export Grid</span>
            </button>
          )}

          {/* Clear All */}
          {batches.length > 0 && (
            <button
              onClick={handleClearAll}
              className="p-2 rounded-xl bg-slate-950 hover:bg-rose-950/60 border border-slate-800 hover:border-rose-800/60 text-slate-400 hover:text-rose-300 transition-colors cursor-pointer"
              title="Clear all evaluation records"
            >
              <Trash2 className="w-4 h-4" />
            </button>
          )}
        </div>
      </div>

      {/* Metric Cards Banner */}
      <div className="grid grid-cols-2 sm:grid-cols-5 gap-3">
        <div className="bg-slate-900/60 border border-slate-800 p-3.5 rounded-xl">
          <span className="text-[11px] text-slate-400 block mb-0.5">Total Students</span>
          <span className="text-xl font-bold text-slate-100">{stats.totalStudents}</span>
        </div>
        <div className="bg-slate-900/60 border border-slate-800 p-3.5 rounded-xl">
          <span className="text-[11px] text-slate-400 block mb-0.5">Evaluation Columns</span>
          <span className="text-xl font-bold text-purple-400">{stats.totalEvals}</span>
        </div>
        <div className="bg-slate-900/60 border border-slate-800 p-3.5 rounded-xl">
          <span className="text-[11px] text-slate-400 block mb-0.5">Absent Instances</span>
          <span className="text-xl font-bold text-amber-400">{stats.absentInstances}</span>
        </div>
        <div className="bg-rose-950/30 border border-rose-800/50 p-3.5 rounded-xl">
          <span className="text-[11px] text-rose-300 block mb-0.5">Not in Master (Red)</span>
          <span className="text-xl font-bold text-rose-400">{stats.redCount}</span>
        </div>
        <div className="bg-amber-950/30 border border-amber-800/50 p-3.5 rounded-xl">
          <span className="text-[11px] text-amber-300 block mb-0.5">Missing Contacts (Yellow)</span>
          <span className="text-xl font-bold text-amber-400">{stats.yellowCount}</span>
        </div>
      </div>

      {/* Batch Ribbon (Stacked Files) */}
      {batches.length > 0 && (
        <div className="bg-slate-900/40 border border-slate-800 p-3 rounded-xl flex flex-wrap items-center gap-2">
          <span className="text-xs text-slate-400 flex items-center space-x-1.5 font-medium mr-1">
            <Layers className="w-3.5 h-3.5 text-purple-400" />
            <span>Active Batches ({batches.length}):</span>
          </span>
          {batches.map((b) => (
            <div
              key={b.id}
              className="inline-flex items-center space-x-2 px-3 py-1 rounded-lg bg-slate-950 border border-slate-800 text-xs text-slate-300 shadow-sm"
            >
              <div className="w-2 h-2 rounded-full bg-purple-500 animate-pulse" />
              <span className="font-semibold text-purple-300 font-mono">[{b.subjectCode}]</span>
              <span className="truncate max-w-[140px] text-slate-200">{b.fileName}</span>
              <span className="text-[10px] text-slate-500">({b.evaluations.length} evals)</span>
              <button
                onClick={() => handleRemoveBatch(b.id)}
                className="text-slate-500 hover:text-rose-400 ml-1 cursor-pointer"
                title="Remove this batch"
              >
                <X className="w-3.5 h-3.5" />
              </button>
            </div>
          ))}
        </div>
      )}

      {/* Legend & Filter Controls */}
      {students.length > 0 && (
        <div className="bg-slate-900/40 border border-slate-800 p-4 rounded-xl flex flex-col md:flex-row items-start md:items-center justify-between gap-4">
          {/* Color Code Legend */}
          <div className="flex flex-wrap items-center gap-4 text-xs">
            <span className="text-slate-400 font-semibold">Row Indicators:</span>
            <div className="flex items-center space-x-1.5">
              <span className="w-3 h-3 rounded-sm bg-rose-500 inline-block" />
              <span className="text-slate-300">Red: Roll Number missing from Master DB</span>
            </div>
            <div className="flex items-center space-x-1.5">
              <span className="w-3 h-3 rounded-sm bg-amber-500 inline-block" />
              <span className="text-slate-300">Yellow: Contact details incomplete</span>
            </div>
            <div className="flex items-center space-x-1.5">
              <span className="w-3 h-3 rounded-sm bg-slate-700 inline-block" />
              <span className="text-slate-300">Normal: Complete in Master DB</span>
            </div>
          </div>

          {/* Search & Filter Bar */}
          <div className="flex flex-wrap items-center gap-2.5 w-full md:w-auto">
            <div className="relative flex-1 md:w-64">
              <Search className="w-3.5 h-3.5 text-slate-500 absolute left-3 top-1/2 -translate-y-1/2" />
              <input
                type="text"
                placeholder="Search name or roll..."
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                className="w-full pl-8 pr-3 py-1.5 bg-slate-950 border border-slate-800 rounded-xl text-xs text-slate-200 placeholder-slate-500 focus:outline-none focus:border-purple-500"
              />
            </div>

            <div className="flex items-center space-x-1 bg-slate-950 p-1 rounded-xl border border-slate-800 text-xs">
              <button
                onClick={() => setStatusFilter('all')}
                className={`px-2.5 py-1 rounded-lg transition-all ${
                  statusFilter === 'all'
                    ? 'bg-purple-600 text-white font-semibold'
                    : 'text-slate-400 hover:text-slate-200'
                }`}
              >
                All ({students.length})
              </button>
              <button
                onClick={() => setStatusFilter('red')}
                className={`px-2 py-1 rounded-lg transition-all ${
                  statusFilter === 'red'
                    ? 'bg-rose-600 text-white font-semibold'
                    : 'text-rose-400 hover:bg-rose-950/40'
                }`}
              >
                Red ({stats.redCount})
              </button>
              <button
                onClick={() => setStatusFilter('yellow')}
                className={`px-2 py-1 rounded-lg transition-all ${
                  statusFilter === 'yellow'
                    ? 'bg-amber-600 text-white font-semibold'
                    : 'text-amber-400 hover:bg-amber-950/40'
                }`}
              >
                Yellow ({stats.yellowCount})
              </button>
              <button
                onClick={() => setStatusFilter('absent')}
                className={`px-2 py-1 rounded-lg transition-all ${
                  statusFilter === 'absent'
                    ? 'bg-slate-700 text-white font-semibold'
                    : 'text-slate-400 hover:text-slate-200'
                }`}
              >
                Absent ({stats.absentInstances})
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Main Table or Empty State */}
      {students.length === 0 ? (
        <div className="bg-slate-900/40 border border-dashed border-slate-800 rounded-2xl p-12 text-center space-y-4">
          <div className="w-16 h-16 rounded-2xl bg-purple-950/50 border border-purple-800/40 text-purple-400 flex items-center justify-center mx-auto">
            <Award className="w-8 h-8" />
          </div>
          <div className="max-w-md mx-auto space-y-1.5">
            <h3 className="text-base font-bold text-slate-100">No Evaluation Records Uploaded</h3>
            <p className="text-xs text-slate-400 leading-relaxed">
              Upload evaluation marks sheets (.xlsx, Microsoft Excel 97-2003 .xls, or CSV). The system will extract assessment titles, maximum marks from brackets, identify absent students, and verify student identity in the Master Database.
            </p>
          </div>
          <div className="flex items-center justify-center gap-3 pt-2">
            <button
              onClick={() => downloadEvaluationTemplate('xlsx', subjectName, subjectCode, masterStudents)}
              className="px-4 py-2 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-200 text-xs font-semibold inline-flex items-center space-x-2 transition-colors cursor-pointer"
            >
              <Download className="w-4 h-4 text-indigo-400" />
              <span>Download Excel Template</span>
            </button>
            <button
              onClick={handleOpenUpload}
              className="px-5 py-2 rounded-xl bg-purple-600 hover:bg-purple-500 text-white text-xs font-semibold inline-flex items-center space-x-2 shadow-lg shadow-purple-600/30 transition-all cursor-pointer"
            >
              <Upload className="w-4 h-4" />
              <span>Upload Evaluation File</span>
            </button>
          </div>
        </div>
      ) : (
        <div className="bg-slate-900/60 border border-slate-800 rounded-2xl overflow-hidden shadow-xl">
          <div className="overflow-x-auto">
            <table className="w-full text-left text-xs border-collapse">
              <thead>
                {/* Super Header Row: Student Information + Subject Batch Groups */}
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

                  {/* Dynamic Evaluation Column Superheaders */}
                  {batches.map((b) => (
                    <th
                      key={b.id}
                      colSpan={b.evaluations.length}
                      className="px-4 py-2.5 text-center border-r border-slate-800 bg-purple-950/40 text-purple-200 font-semibold text-xs border-b border-slate-800"
                    >
                      <div className="flex items-center justify-center space-x-2">
                        <span className="px-2 py-0.5 rounded bg-purple-900/80 text-purple-300 font-mono font-bold text-[10px]">
                          {b.subjectCode}
                        </span>
                        <span className="truncate max-w-[220px] font-medium">{b.subjectName}</span>
                      </div>
                    </th>
                  ))}
                </tr>

                {/* Second Header Row: Individual Evaluation Names & Max Marks */}
                <tr className="bg-slate-950/90 border-b border-slate-800 text-[11px] text-slate-400 font-mono">
                  {batches.map((b) =>
                    b.evaluations.map((ev) => (
                      <th
                        key={ev.id}
                        className="px-3 py-2.5 text-center border-r border-slate-800/70 min-w-[110px]"
                        title={ev.rawHeader}
                      >
                        <div className="font-semibold text-slate-200">{ev.name}</div>
                        <div className="text-[10px] text-purple-400 font-mono">
                          Max: {ev.maxMarks} pts
                        </div>
                      </th>
                    ))
                  )}
                </tr>
              </thead>

              <tbody className="divide-y divide-slate-800/70">
                {filteredStudents.length === 0 ? (
                  <tr>
                    <td
                      colSpan={3 + batches.reduce((acc, b) => acc + b.evaluations.length, 0)}
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

                        {/* Master DB & Contact Verification Status */}
                        <td className="px-4 py-3 border-r border-slate-800/60 align-top">
                          <div className="space-y-1.5">
                            {/* DB Existence Status */}
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

                            {/* Contact Verification details */}
                            <div className="flex items-center gap-2 text-[10px]">
                              {/* Student Contact */}
                              <span
                                className={`px-1.5 py-0.5 rounded font-mono ${
                                  status.hasStudentContact
                                    ? 'bg-slate-800 text-slate-300'
                                    : 'bg-amber-950 text-amber-300 border border-amber-800/60 font-semibold'
                                }`}
                              >
                                Student: {status.hasStudentContact ? 'Contact OK' : 'No Contact'}
                              </span>

                              {/* Parent Contact */}
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

                        {/* Individual Evaluation Mark Cells */}
                        {batches.map((b) =>
                          b.evaluations.map((ev) => {
                            const evalItem = st.evaluations[ev.id];
                            const isEditing =
                              editingCell?.rollNumber === st.rollNumber &&
                              editingCell?.evalId === ev.id;

                            if (isEditing) {
                              return (
                                <td
                                  key={ev.id}
                                  className="px-2 py-2 text-center border-r border-slate-800/60 bg-purple-950/50"
                                >
                                  <div className="flex items-center justify-center space-x-1">
                                    <input
                                      type="text"
                                      autoFocus
                                      value={editValue}
                                      onChange={(e) => setEditValue(e.target.value)}
                                      onKeyDown={(e) => {
                                        if (e.key === 'Enter') {
                                          handleSaveEdit(st.rollNumber, ev.id, ev.maxMarks);
                                        } else if (e.key === 'Escape') {
                                          setEditingCell(null);
                                        }
                                      }}
                                      onBlur={() => handleSaveEdit(st.rollNumber, ev.id, ev.maxMarks)}
                                      className="w-16 px-1.5 py-1 bg-slate-900 border border-purple-500 rounded text-center text-xs text-white font-mono focus:outline-none"
                                      placeholder="0 or A"
                                    />
                                  </div>
                                  <span className="text-[9px] text-slate-400 block mt-0.5">
                                    Type 'A' for Absent
                                  </span>
                                </td>
                              );
                            }

                            // Absent representation
                            if (!evalItem || evalItem.isAbsent) {
                              return (
                                <td
                                  key={ev.id}
                                  onClick={() =>
                                    handleStartEdit(
                                      st.rollNumber,
                                      ev.id,
                                      evalItem?.marksObtained ?? null,
                                      true
                                    )
                                  }
                                  className="px-2 py-2 text-center border-r border-slate-800/60 group cursor-pointer hover:bg-slate-800/50"
                                  title="Click to edit mark or remove absent status"
                                >
                                  <div className="inline-flex items-center justify-center space-x-1 px-2 py-0.5 rounded bg-rose-950/80 text-rose-300 border border-rose-800/70 font-mono text-[10px] font-bold">
                                    <span>ABSENT</span>
                                    <Edit3 className="w-2.5 h-2.5 opacity-0 group-hover:opacity-100 text-rose-400 transition-opacity" />
                                  </div>
                                </td>
                              );
                            }

                            // Present with score
                            const pct = evalItem.percentage ?? 0;
                            let scoreBadgeColor = 'text-slate-200';
                            if (pct >= 85) scoreBadgeColor = 'text-emerald-400 font-semibold';
                            else if (pct < 50) scoreBadgeColor = 'text-rose-400 font-semibold';

                            return (
                              <td
                                key={ev.id}
                                onClick={() =>
                                  handleStartEdit(
                                    st.rollNumber,
                                    ev.id,
                                    evalItem.marksObtained,
                                    false
                                  )
                                }
                                className="px-3 py-2 text-center border-r border-slate-800/60 font-mono group cursor-pointer hover:bg-slate-800/50"
                                title="Click to edit mark"
                              >
                                <div className="flex items-center justify-center space-x-1">
                                  <span className={`text-xs ${scoreBadgeColor}`}>
                                    {evalItem.marksObtained}
                                  </span>
                                  <span className="text-[10px] text-slate-500">/ {ev.maxMarks}</span>
                                  <Edit3 className="w-2.5 h-2.5 opacity-0 group-hover:opacity-100 text-purple-400 transition-opacity" />
                                </div>
                                <div className="text-[10px] text-slate-400">{pct}%</div>
                              </td>
                            );
                          })
                        )}
                      </tr>
                    );
                  })
                )}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* Upload Wizard Modal */}
      {isModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-950/80 backdrop-blur-sm overflow-y-auto">
          <div className="bg-slate-900 border border-slate-800 rounded-2xl w-full max-w-xl shadow-2xl overflow-hidden my-8">
            {/* Modal Header */}
            <div className="px-6 py-4 border-b border-slate-800 flex items-center justify-between bg-slate-900/60">
              <div className="flex items-center space-x-3">
                <div className="p-2 rounded-xl bg-purple-950 border border-purple-800/60 text-purple-400">
                  <Award className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="font-bold text-slate-100 text-sm">Upload Student Evaluation Record</h3>
                  <p className="text-xs text-slate-400">
                    Configure subject details, download template if needed, and import assessment marks.
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

              {/* Step 1: Subject Selection & Subject Code */}
              <div className="space-y-3 bg-slate-950/60 p-4 rounded-xl border border-slate-800 text-xs">
                <label className="font-bold text-slate-200 block">
                  1. Subject Name &amp; Subject Code *
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
              </div>

              {/* Step 2: Download Standard Template */}
              <div className="bg-purple-950/30 border border-purple-800/40 rounded-xl p-3.5 flex items-center justify-between gap-3 text-xs">
                <div>
                  <span className="font-semibold text-purple-200 block">Download Template First?</span>
                  <span className="text-[11px] text-purple-300/80">
                    Pre-formatted with Student Name, Roll Number, and sample assessment columns.
                  </span>
                </div>
                <div className="flex items-center space-x-1.5 shrink-0">
                  <button
                    type="button"
                    onClick={() => downloadEvaluationTemplate('xlsx', subjectName, subjectCode, masterStudents)}
                    className="px-3 py-1.5 rounded-lg bg-purple-600 hover:bg-purple-500 text-white text-xs font-semibold inline-flex items-center space-x-1.5 shadow-sm transition-colors cursor-pointer"
                  >
                    <Download className="w-3.5 h-3.5" />
                    <span>.xlsx</span>
                  </button>
                  <button
                    type="button"
                    onClick={() => downloadEvaluationTemplate('xls', subjectName, subjectCode, masterStudents)}
                    className="px-2.5 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-200 text-xs font-mono transition-colors cursor-pointer"
                    title="Excel 97-2003 (.xls)"
                  >
                    .xls
                  </button>
                </div>
              </div>

              {/* Step 3: Choose File to Upload */}
              <div className="space-y-2">
                <label className="block text-xs font-bold text-slate-200">
                  2. Upload Evaluation File (Excel .xlsx / Microsoft Excel 97-2003 .xls / CSV) *
                </label>

                <div
                  onClick={() => fileInputRef.current?.click()}
                  className="border-2 border-dashed border-slate-800 hover:border-purple-500/70 rounded-xl p-5 text-center cursor-pointer bg-slate-950/50 hover:bg-slate-950/80 transition-all space-y-2"
                >
                  <input
                    ref={fileInputRef}
                    type="file"
                    accept=".xlsx,.xls,.csv,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    onChange={handleFileChange}
                    className="hidden"
                  />
                  <div className="w-10 h-10 rounded-xl bg-slate-900 border border-slate-800 text-purple-400 flex items-center justify-center mx-auto">
                    {isParsing ? (
                      <RefreshCw className="w-5 h-5 animate-spin text-purple-400" />
                    ) : (
                      <Upload className="w-5 h-5" />
                    )}
                  </div>
                  <div>
                    <span className="text-xs font-semibold text-slate-200 block">
                      {uploadedFile ? uploadedFile.name : 'Click to select evaluation workbook'}
                    </span>
                    <span className="text-[11px] text-slate-400">
                      Empty cells or blank entries will automatically be registered as ABSENT
                    </span>
                    {parsedResult?.fileFormatDetected && (
                      <div className="mt-2">
                        <span className="inline-flex items-center space-x-1 px-2.5 py-0.5 rounded-full text-[10px] font-mono bg-emerald-950/80 text-emerald-300 border border-emerald-800/60">
                          <CheckCircle2 className="w-3 h-3 text-emerald-400" />
                          <span>Detected: {parsedResult.fileFormatDetected}</span>
                        </span>
                      </div>
                    )}
                  </div>
                </div>
              </div>

              {/* Step 4: Preview Extracted Columns */}
              {parsedResult && (
                <div className="bg-slate-950/70 p-4 rounded-xl border border-slate-800 space-y-2.5">
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-bold text-slate-200">
                      Extracted Assessment Columns ({parsedResult.evaluations.length}):
                    </span>
                    <span className="text-[11px] text-purple-300 font-mono">
                      {parsedResult.totalRows} student rows found
                    </span>
                  </div>

                  <div className="flex flex-wrap gap-1.5">
                    {parsedResult.evaluations.map((ev) => (
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

              <button
                type="button"
                disabled={!parsedResult}
                onClick={handleCommitBatch}
                className="px-5 py-2 rounded-xl bg-purple-600 hover:bg-purple-500 text-white text-xs font-semibold shadow-md shadow-purple-600/20 transition-all disabled:opacity-40 cursor-pointer"
              >
                Add Columns to Grid
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
