import React, { useState, useMemo, useEffect, useRef } from 'react';
import {
  ArrowLeft,
  Mail,
  Send,
  Download,
  Search,
  Filter,
  CheckCircle2,
  AlertTriangle,
  AlertCircle,
  Eye,
  Edit3,
  Copy,
  Check,
  X,
  FileSpreadsheet,
  Users,
  ShieldCheck,
  RefreshCw,
  Clock,
  Sparkles,
  XCircle,
  AlertOctagon,
  Lock,
  ChevronDown,
  ChevronUp,
  RotateCw,
  Info,
} from 'lucide-react';
import * as XLSX from 'xlsx';
import {
  StudentMaster,
  User,
  DispatchOverallStatus,
  RecipientDispatchResult,
  StudentDispatchRecord,
  DispatchJob,
} from '../../types/index.ts';
import {
  StackedBatch,
  UnifiedStudentRow,
  StudentMasterMatch,
} from './TeacherCommunicationModule.tsx';
import {
  AttendanceSubjectItem,
  EvaluationAssessmentItem,
  TeacherSenderInfo,
  StudentRecipientData,
  GeneratedEmail,
  generateStudentEmail,
  generateParentEmail,
} from '../../utils/communicationEmailGenerator.ts';
import {
  getOutlookStatus,
  initiateBulkDispatch,
  getBulkDispatchProgress,
  getLatestBulkDispatch,
  retryBulkDispatch,
  stopBulkDispatch,
} from '../../services/outlookApi.ts';

interface BulkCommunicationModuleProps {
  batches: StackedBatch[];
  students: UnifiedStudentRow[];
  masterStudentMap: Map<string, StudentMaster>;
  teacherUser: User | null;
  recipientOptions: {
    sendToStudent: boolean;
    sendToParent: boolean;
  };
  onBack: () => void;
}

export const BulkCommunicationModule: React.FC<BulkCommunicationModuleProps> = ({
  batches,
  students,
  masterStudentMap,
  teacherUser,
  recipientOptions,
  onBack,
}) => {
  const { sendToStudent, sendToParent } = recipientOptions;

  // Search and filter state
  const [searchQuery, setSearchQuery] = useState('');
  const [filterMode, setFilterMode] = useState<
    'all' | 'ready' | 'missing-email' | 'no-data' | 'success' | 'partial' | 'failed'
  >('all');

  // Review / Edit Email Modal State
  const [activeReviewEmail, setActiveReviewEmail] = useState<{
    student: UnifiedStudentRow;
    recipientType: 'student' | 'parent';
    email: GeneratedEmail;
  } | null>(null);

  // Status Details Modal State for a specific student's recipient breakdown
  const [selectedStudentStatusDetail, setSelectedStudentStatusDetail] = useState<{
    student: UnifiedStudentRow;
    master: StudentMaster | null;
    statusRecord?: StudentDispatchRecord;
    computedOverallStatus: DispatchOverallStatus;
  } | null>(null);

  // Customized email overrides: key `${rollNumber}_${recipientType}` -> custom text
  const [customizedBodies, setCustomizedBodies] = useState<Record<string, string>>({});
  const [copiedKey, setCopiedKey] = useState<string | null>(null);

  // Outlook Configuration Notice Modal (when teacher account is not configured/authenticated)
  const [showOutlookNoticeModal, setShowOutlookNoticeModal] = useState(false);

  // Dispatch progress state & locking
  const [isDispatching, setIsDispatching] = useState(false);
  const [isStopping, setIsStopping] = useState(false);
  const [activeDispatchJob, setActiveDispatchJob] = useState<DispatchJob | null>(null);
  const [dispatchError, setDispatchError] = useState<string | null>(null);

  // Expanded student status breakdown accordion in table: rollNumber -> boolean
  const [expandedBreakdownRolls, setExpandedBreakdownRolls] = useState<Record<string, boolean>>({});

  const pollingTimerRef = useRef<any>(null);

  // Load persistent latest dispatch on initial mount
  useEffect(() => {
    let isMounted = true;
    const fetchLatest = async () => {
      if (!teacherUser?.id) return;
      try {
        const res = await getLatestBulkDispatch(teacherUser.id);
        if (isMounted && res.success && res.job) {
          setActiveDispatchJob(res.job);
        }
      } catch (e) {
        // quiet fallback
      }
    };
    fetchLatest();
    return () => {
      isMounted = false;
      if (pollingTimerRef.current) clearInterval(pollingTimerRef.current);
    };
  }, [teacherUser?.id]);

  // Application Closure Warning when bulk dispatch is in progress (Requirement 4)
  useEffect(() => {
    if (!isDispatching) return;

    const handleBeforeUnload = (e: BeforeUnloadEvent) => {
      e.preventDefault();
      const message =
        'Please do not close or exit the application while bulk email communication is in progress. Closing the application may interrupt the dispatch process, resulting in incomplete work and loss of unsaved data. Kindly wait until the entire process is completed.';
      e.returnValue = message;
      return message;
    };

    window.addEventListener('beforeunload', handleBeforeUnload);
    return () => {
      window.removeEventListener('beforeunload', handleBeforeUnload);
    };
  }, [isDispatching]);

  // Teacher sender information
  const teacherSender: TeacherSenderInfo = useMemo(() => {
    const dept = (teacherUser as any)?.department;
    return {
      name: teacherUser?.name || 'Faculty Member',
      designation: dept
        ? `Faculty, Department of ${dept}`
        : teacherUser?.employeeId
        ? `Faculty (Emp ID: ${teacherUser.employeeId})`
        : 'Faculty Member',
      department: dept,
      email: teacherUser?.emailId,
    };
  }, [teacherUser]);

  // Compute generated emails for all students
  const generatedReports = useMemo(() => {
    return students.map((st) => {
      const cleanRoll = st.rollNumber.trim().toLowerCase();
      const master = masterStudentMap.get(cleanRoll) || null;

      // Collect student attendance data across all attendance batches
      const attendanceItems: AttendanceSubjectItem[] = [];
      batches.forEach((b) => {
        if (b.type === 'attendance') {
          b.subjects.forEach((sub) => {
            const key = `${b.id}_${sub.code}`;
            const rec = st.attendance[key];
            if (rec && (rec.conducted !== null || rec.attended !== null || rec.percentage !== null)) {
              attendanceItems.push({
                subjectName: sub.name,
                subjectCode: sub.code,
                period: b.dateRange.label,
                conducted: rec.conducted,
                attended: rec.attended,
                percentage: rec.percentage,
              });
            }
          });
        }
      });

      // Collect student evaluation data across all evaluation batches
      const evaluationItems: EvaluationAssessmentItem[] = [];
      batches.forEach((b) => {
        if (b.type === 'evaluation') {
          b.evaluations.forEach((ev) => {
            const cellKey = `${b.id}_${ev.id}`;
            const rec = st.evaluations[cellKey];
            if (rec) {
              evaluationItems.push({
                subjectName: b.subjectName,
                subjectCode: b.subjectCode,
                componentName: ev.name,
                marksObtained: rec.marksObtained,
                maxMarks: ev.maxMarks,
                isAbsent: rec.isAbsent,
                percentage: rec.percentage,
              });
            }
          });
        }
      });

      const hasAcademicRecord = attendanceItems.length > 0 || evaluationItems.length > 0;

      const recipientData: StudentRecipientData = {
        rollNumber: st.rollNumber,
        name: st.name || master?.name || 'Student',
        section: st.section || master?.section,
        studentEmails: master?.emailIds || [],
        parentName: master?.parentName,
        parentEmails: master?.parentEmailIds || [],
      };

      const studentEmail = sendToStudent
        ? generateStudentEmail(recipientData, attendanceItems, evaluationItems, teacherSender)
        : null;

      const parentEmail = sendToParent
        ? generateParentEmail(recipientData, attendanceItems, evaluationItems, teacherSender)
        : null;

      const studentEmailIds = master?.emailIds || [];
      const parentEmailIds = master?.parentEmailIds || [];

      const hasStudentEmailAddress = studentEmailIds.length > 0;
      const hasParentEmailAddress = parentEmailIds.length > 0;

      return {
        student: st,
        master,
        hasAcademicRecord,
        attendanceItems,
        evaluationItems,
        studentEmail,
        parentEmail,
        studentEmailIds,
        parentEmailIds,
        hasStudentEmailAddress,
        hasParentEmailAddress,
      };
    });
  }, [students, batches, masterStudentMap, teacherSender, sendToStudent, sendToParent]);

  // Lookup map for persistent dispatch records by roll number
  const dispatchRecordsMap = useMemo(() => {
    const map = new Map<string, StudentDispatchRecord>();
    if (activeDispatchJob?.records) {
      activeDispatchJob.records.forEach((r) => {
        map.set(r.rollNumber.trim().toLowerCase(), r);
      });
    }
    return map;
  }, [activeDispatchJob]);

  // Helper to determine exact dynamic status for a student row
  const getStudentDispatchStatus = (
    rollNumber: string,
    hasAcademicRecord: boolean,
    studentEmailIds: string[],
    parentEmailIds: string[]
  ): {
    overallStatus: DispatchOverallStatus | 'Ready to Dispatch';
    record?: StudentDispatchRecord;
    missingReason?: string;
  } => {
    const cleanRoll = rollNumber.trim().toLowerCase();
    const record = dispatchRecordsMap.get(cleanRoll);

    // If an active persistent dispatch record exists, use its updated state
    if (record) {
      return {
        overallStatus: record.overallStatus,
        record,
      };
    }

    // Otherwise, compute dynamic status based on data availability
    if (!hasAcademicRecord) {
      return {
        overallStatus: 'Record Missing in Persistent Database',
        missingReason: 'Neither attendance nor evaluation records exist for this student.',
      };
    }

    if (sendToStudent && studentEmailIds.length === 0) {
      return {
        overallStatus: 'Student Email Missing',
        missingReason: 'No student email address registered in the Student Master Database.',
      };
    }

    if (sendToParent && parentEmailIds.length === 0) {
      return {
        overallStatus: 'Parent Email Missing',
        missingReason: 'No parent/guardian email address registered in the Student Master Database.',
      };
    }

    return {
      overallStatus: 'Ready to Dispatch',
    };
  };

  // Overall statistics for communication
  const stats = useMemo(() => {
    let studentEmailsReady = 0;
    let parentEmailsReady = 0;
    let missingStudentEmail = 0;
    let missingParentEmail = 0;
    let noAcademicDataCount = 0;
    let successCount = 0;
    let partialCount = 0;
    let failedCount = 0;

    generatedReports.forEach((r) => {
      const statusInfo = getStudentDispatchStatus(
        r.student.rollNumber,
        r.hasAcademicRecord,
        r.studentEmailIds,
        r.parentEmailIds
      );

      if (statusInfo.overallStatus === 'Success') successCount++;
      else if (statusInfo.overallStatus === 'Partial Success') partialCount++;
      else if (statusInfo.overallStatus === 'Failed') failedCount++;

      if (!r.hasAcademicRecord) {
        noAcademicDataCount++;
      } else {
        if (sendToStudent) {
          if (r.studentEmail && r.hasStudentEmailAddress) studentEmailsReady++;
          else if (r.studentEmail && !r.hasStudentEmailAddress) missingStudentEmail++;
        }
        if (sendToParent) {
          if (r.parentEmail && r.hasParentEmailAddress) parentEmailsReady++;
          else if (r.parentEmail && !r.hasParentEmailAddress) missingParentEmail++;
        }
      }
    });

    return {
      totalStudents: students.length,
      studentEmailsReady,
      parentEmailsReady,
      missingStudentEmail,
      missingParentEmail,
      noAcademicDataCount,
      successCount,
      partialCount,
      failedCount,
    };
  }, [generatedReports, students, sendToStudent, sendToParent, dispatchRecordsMap]);

  // Filtered rows for display
  const filteredReports = useMemo(() => {
    return generatedReports.filter((r) => {
      const q = searchQuery.toLowerCase().trim();
      const matchesSearch =
        !q ||
        r.student.name.toLowerCase().includes(q) ||
        r.student.rollNumber.toLowerCase().includes(q);

      if (!matchesSearch) return false;

      const statusInfo = getStudentDispatchStatus(
        r.student.rollNumber,
        r.hasAcademicRecord,
        r.studentEmailIds,
        r.parentEmailIds
      );

      if (filterMode === 'ready') {
        return statusInfo.overallStatus === 'Ready to Dispatch' || statusInfo.overallStatus === 'Success';
      }

      if (filterMode === 'missing-email') {
        return (
          statusInfo.overallStatus === 'Student Email Missing' ||
          statusInfo.overallStatus === 'Parent Email Missing'
        );
      }

      if (filterMode === 'no-data') {
        return statusInfo.overallStatus === 'Record Missing in Persistent Database';
      }

      if (filterMode === 'success') {
        return statusInfo.overallStatus === 'Success';
      }

      if (filterMode === 'partial') {
        return statusInfo.overallStatus === 'Partial Success';
      }

      if (filterMode === 'failed') {
        return statusInfo.overallStatus === 'Failed';
      }

      return true;
    });
  }, [generatedReports, searchQuery, filterMode, sendToStudent, sendToParent, dispatchRecordsMap]);

  // Copy text helper
  const handleCopy = (text: string, key: string) => {
    navigator.clipboard.writeText(text);
    setCopiedKey(key);
    setTimeout(() => setCopiedKey(null), 2000);
  };

  // Toggle row accordion breakdown
  const toggleRowBreakdown = (roll: string) => {
    setExpandedBreakdownRolls((prev) => ({
      ...prev,
      [roll]: !prev[roll],
    }));
  };

  // Save manual text override from review modal
  const handleSaveCustomOverride = () => {
    if (!activeReviewEmail) return;
    const key = `${activeReviewEmail.student.rollNumber}_${activeReviewEmail.recipientType}`;
    setCustomizedBodies((prev) => ({
      ...prev,
      [key]: activeReviewEmail.email.bodyText,
    }));
    setActiveReviewEmail(null);
  };

  // Poll background job progress until complete
  const startProgressPolling = (dispatchId: string) => {
    if (pollingTimerRef.current) clearInterval(pollingTimerRef.current);

    pollingTimerRef.current = setInterval(async () => {
      try {
        const res = await getBulkDispatchProgress(dispatchId);
        if (res.success && res.job) {
          setActiveDispatchJob(res.job);
          if (res.job.status === 'completed' || res.job.status === 'failed') {
            clearInterval(pollingTimerRef.current);
            pollingTimerRef.current = null;
            setIsDispatching(false);
          }
        }
      } catch (err) {
        console.warn('Polling progress error:', err);
      }
    }, 700);
  };

  // Initiate Bulk Email Dispatch (Requirement 1 & 4)
  const handleInitiateDispatch = async () => {
    if (!teacherUser?.id) {
      alert('Teacher authentication session not found. Please log in again.');
      return;
    }

    setDispatchError(null);

    // 1. Verify Microsoft Outlook Account authentication status
    try {
      const outlookStatus = await getOutlookStatus(teacherUser.id);
      if (!outlookStatus.connected || outlookStatus.status !== 'CONNECTED') {
        setShowOutlookNoticeModal(true);
        return;
      }
    } catch (err) {
      setShowOutlookNoticeModal(true);
      return;
    }

    // 2. Prepare student payload with multiple emails and customized content
    const studentsPayload = generatedReports.map((r) => {
      const studentKey = `${r.student.rollNumber}_student`;
      const parentKey = `${r.student.rollNumber}_parent`;

      const studentBody = customizedBodies[studentKey] || r.studentEmail?.bodyHtml || '';
      const parentBody = customizedBodies[parentKey] || r.parentEmail?.bodyHtml || '';

      return {
        rollNumber: r.student.rollNumber,
        name: r.student.name || r.master?.name || 'Student',
        hasAcademicRecord: r.hasAcademicRecord,
        studentEmailIds: r.studentEmailIds,
        parentEmailIds: r.parentEmailIds,
        studentTemplate: studentBody,
        parentTemplate: parentBody,
      };
    });

    try {
      setIsDispatching(true);
      const res = await initiateBulkDispatch({
        teacherId: teacherUser.id,
        teacherName: teacherUser.name,
        recipientOptions: {
          sendToStudent,
          sendToParent,
        },
        students: studentsPayload,
      });

      if (res.success && res.job) {
        setActiveDispatchJob(res.job);
        startProgressPolling(res.dispatchId);
      }
    } catch (err: any) {
      setIsDispatching(false);
      if (err.requiresConfiguration) {
        setShowOutlookNoticeModal(true);
      } else {
        setDispatchError(err.message || 'Failed to initiate bulk dispatch');
      }
    }
  };

  // Selective Retry for Failed Emails (Requirement 5)
  const handleRetryFailed = async () => {
    if (!activeDispatchJob?.id || !teacherUser?.id) return;

    // Verify Outlook account first
    try {
      const outlookStatus = await getOutlookStatus(teacherUser.id);
      if (!outlookStatus.connected || outlookStatus.status !== 'CONNECTED') {
        setShowOutlookNoticeModal(true);
        return;
      }
    } catch (err) {
      setShowOutlookNoticeModal(true);
      return;
    }

    try {
      setIsDispatching(true);
      const res = await retryBulkDispatch({
        dispatchId: activeDispatchJob.id,
        teacherId: teacherUser.id,
      });

      if (res.success && res.job) {
        setActiveDispatchJob(res.job);
        startProgressPolling(activeDispatchJob.id);
      }
    } catch (err: any) {
      setIsDispatching(false);
      if (err.requiresConfiguration) {
        setShowOutlookNoticeModal(true);
      } else {
        alert(err.message || 'Failed to retry failed emails.');
      }
    }
  };

  // Hard stop / Interrupt Dispatch Process
  const handleStopDispatch = async () => {
    if (!activeDispatchJob?.id) {
      setIsDispatching(false);
      return;
    }

    if (
      !window.confirm(
        'Are you sure you want to hard stop the bulk email dispatch? All remaining pending emails will be immediately cancelled. Emails already sent will remain recorded as Success.'
      )
    ) {
      return;
    }

    setIsStopping(true);
    try {
      if (pollingTimerRef.current) {
        clearInterval(pollingTimerRef.current);
        pollingTimerRef.current = null;
      }

      const res = await stopBulkDispatch({
        dispatchId: activeDispatchJob.id,
        teacherId: teacherUser?.id || '',
      });

      if (res.job) {
        setActiveDispatchJob(res.job);
      }
    } catch (err: any) {
      console.error('Failed to stop dispatch:', err);
    } finally {
      setIsStopping(false);
      setIsDispatching(false);
    }
  };


  // Export full Email Report to Excel (.xlsx) including Status and Email Templates
  const handleExportEmailReport = () => {
    const rows = generatedReports.map((r) => {
      const statusInfo = getStudentDispatchStatus(
        r.student.rollNumber,
        r.hasAcademicRecord,
        r.studentEmailIds,
        r.parentEmailIds
      );

      const studentKey = `${r.student.rollNumber}_student`;
      const parentKey = `${r.student.rollNumber}_parent`;

      const studentText =
        customizedBodies[studentKey] ||
        r.studentEmail?.bodyText ||
        (r.hasAcademicRecord ? '[Student Email Generated]' : '[No Data]');

      const parentText =
        customizedBodies[parentKey] ||
        r.parentEmail?.bodyText ||
        (r.hasAcademicRecord ? '[Parent Email Generated]' : '[No Data]');

      // Build recipient breakdown text for Excel
      let recipientBreakdown = '';
      if (statusInfo.record) {
        const parts: string[] = [];
        if (statusInfo.record.studentEmails.length > 0) {
          parts.push(
            'STUDENT EMAILS: ' +
              statusInfo.record.studentEmails
                .map((se) => `${se.email} (${se.status}${se.reason ? ' - ' + se.reason : ''})`)
                .join(', ')
          );
        }
        if (statusInfo.record.parentEmails.length > 0) {
          parts.push(
            'PARENT EMAILS: ' +
              statusInfo.record.parentEmails
                .map((pe) => `${pe.email} (${pe.status}${pe.reason ? ' - ' + pe.reason : ''})`)
                .join(', ')
          );
        }
        recipientBreakdown = parts.join(' | ');
      } else if (statusInfo.missingReason) {
        recipientBreakdown = statusInfo.missingReason;
      }

      const rowObj: Record<string, any> = {
        'Roll Number': r.student.rollNumber,
        'Student Name': r.student.name || r.master?.name || '—',
        'In Master DB': r.master ? 'VERIFIED' : 'MISSING',
        'Student Email ID(s)': r.studentEmailIds.join(', ') || 'N/A',
        'Parent Email ID(s)': r.parentEmailIds.join(', ') || 'N/A',
        'Dispatch Status': statusInfo.overallStatus,
        'Recipient Status Details': recipientBreakdown,
      };

      // Stacked attendance & evaluation columns
      batches.forEach((b) => {
        if (b.type === 'attendance') {
          b.subjects.forEach((sub) => {
            const cellKey = `${b.id}_${sub.code}`;
            const rec = r.student.attendance[cellKey];
            const pfx = `[ATT ${sub.code}]`;
            rowObj[`${pfx} Total Days`] = rec?.conducted ?? '—';
            rowObj[`${pfx} Days Attended`] = rec?.attended ?? '—';
            rowObj[`${pfx} Attendance %`] =
              rec?.percentage !== null && rec?.percentage !== undefined ? `${rec.percentage}%` : '—';
          });
        } else {
          b.evaluations.forEach((ev) => {
            const cellKey = `${b.id}_${ev.id}`;
            const rec = r.student.evaluations[cellKey];
            const pfx = `[EVAL ${b.subjectCode}] ${ev.name} (${ev.maxMarks})`;
            if (!rec || rec.isAbsent) {
              rowObj[pfx] = 'ABSENT';
            } else {
              rowObj[pfx] = rec.marksObtained;
            }
          });
        }
      });

      if (sendToStudent) {
        rowObj['Student Email Template'] = studentText;
      }

      if (sendToParent) {
        rowObj['Parent Email Template'] = parentText;
      }

      return rowObj;
    });

    const worksheet = XLSX.utils.json_to_sheet(rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Bulk_Email_Report');
    XLSX.writeFile(
      workbook,
      `Bulk_Email_Communication_Report_${new Date().toISOString().split('T')[0]}.xlsx`
    );
  };

  // Helper badge color & icon for overall dispatch status
  const renderStatusBadge = (status: DispatchOverallStatus | 'Ready to Dispatch') => {
    switch (status) {
      case 'Success':
        return (
          <span className="inline-flex items-center space-x-1 px-2.5 py-1 rounded-full text-[11px] font-bold bg-emerald-950/80 text-emerald-300 border border-emerald-700/60 shadow-sm">
            <CheckCircle2 className="w-3.5 h-3.5 text-emerald-400 shrink-0" />
            <span>Success</span>
          </span>
        );
      case 'Partial Success':
        return (
          <span className="inline-flex items-center space-x-1 px-2.5 py-1 rounded-full text-[11px] font-bold bg-amber-950/80 text-amber-300 border border-amber-700/60 shadow-sm">
            <AlertTriangle className="w-3.5 h-3.5 text-amber-400 shrink-0" />
            <span>Partial Success</span>
          </span>
        );
      case 'Failed':
        return (
          <span className="inline-flex items-center space-x-1 px-2.5 py-1 rounded-full text-[11px] font-bold bg-rose-950/80 text-rose-300 border border-rose-700/60 shadow-sm">
            <XCircle className="w-3.5 h-3.5 text-rose-400 shrink-0" />
            <span>Failed</span>
          </span>
        );
      case 'Student Email Missing':
        return (
          <span className="inline-flex items-center space-x-1 px-2.5 py-1 rounded-full text-[11px] font-semibold bg-amber-950/60 text-amber-400 border border-amber-800/60 shadow-sm">
            <AlertCircle className="w-3.5 h-3.5 shrink-0" />
            <span>Student Email Missing</span>
          </span>
        );
      case 'Parent Email Missing':
        return (
          <span className="inline-flex items-center space-x-1 px-2.5 py-1 rounded-full text-[11px] font-semibold bg-amber-950/60 text-amber-400 border border-amber-800/60 shadow-sm">
            <AlertCircle className="w-3.5 h-3.5 shrink-0" />
            <span>Parent Email Missing</span>
          </span>
        );
      case 'Record Missing in Persistent Database':
        return (
          <span className="inline-flex items-center space-x-1 px-2.5 py-1 rounded-full text-[11px] font-semibold bg-purple-950/70 text-purple-300 border border-purple-800/60 shadow-sm">
            <AlertOctagon className="w-3.5 h-3.5 shrink-0" />
            <span>Record Missing in Persistent Database</span>
          </span>
        );
      case 'Ready to Dispatch':
      default:
        return (
          <span className="inline-flex items-center space-x-1 px-2.5 py-1 rounded-full text-[11px] font-medium bg-slate-800 text-slate-300 border border-slate-700 shadow-sm">
            <Clock className="w-3.5 h-3.5 text-slate-400 shrink-0" />
            <span>Ready to Dispatch</span>
          </span>
        );
    }
  };

  const hasFailedEmails = (activeDispatchJob?.failedEmails || 0) > 0;

  return (
    <div className="space-y-6">
      {/* Top Header & Navigation Bar */}
      <div className="bg-gradient-to-r from-slate-900 via-indigo-950/40 to-slate-900 border border-slate-800 rounded-2xl p-5 flex flex-col md:flex-row md:items-center justify-between gap-4 shadow-xl">
        <div className="flex items-center space-x-3.5">
          <button
            onClick={onBack}
            className="p-2.5 rounded-xl bg-slate-950 border border-slate-800 text-slate-300 hover:text-white hover:border-slate-700 transition-colors cursor-pointer"
            title="Return to Stack Report"
          >
            <ArrowLeft className="w-5 h-5" />
          </button>
          <div>
            <div className="flex items-center space-x-2">
              <h2 className="text-base font-bold text-white tracking-tight">
                Bulk Communication Email Report
              </h2>
              <span className="text-[10px] px-2.5 py-0.5 rounded-full font-semibold bg-indigo-950 text-indigo-300 border border-indigo-800/60">
                Email Dispatch &amp; Status Tracking
              </span>
            </div>
            <p className="text-xs text-slate-400 mt-0.5">
              Target Channels:{' '}
              <strong className="text-slate-200">
                {sendToStudent && sendToParent
                  ? 'Student & Parent Email Addresses'
                  : sendToStudent
                  ? 'Student Email Addresses Only'
                  : 'Parent Email Addresses Only'}
              </strong>
              . Status and individual recipient delivery are tracked persistently in the database.
            </p>
          </div>
        </div>

        <div className="flex flex-wrap items-center gap-2.5">
          {/* Export Email Report */}
          <button
            onClick={handleExportEmailReport}
            className="px-3.5 py-2 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-700 text-xs font-semibold flex items-center space-x-1.5 transition-all cursor-pointer"
            title="Download complete email report with status and customized templates"
          >
            <Download className="w-4 h-4 text-emerald-400" />
            <span>Export Email Report (.xlsx)</span>
          </button>

          {/* Selective Retry Failed Emails Button (Requirement 5) */}
          {hasFailedEmails && (
            <button
              onClick={handleRetryFailed}
              disabled={isDispatching}
              className="px-4 py-2 rounded-xl bg-rose-600 hover:bg-rose-500 text-white text-xs font-bold flex items-center space-x-2 shadow-lg shadow-rose-600/30 transition-all cursor-pointer disabled:opacity-50"
              title="Selectively retry only failed emails without resending successful ones"
            >
              <RotateCw className="w-4 h-4" />
              <span>Retry Failed Emails ({activeDispatchJob?.failedEmails})</span>
            </button>
          )}

          {/* Initiate Bulk Email Dispatch Button (Requirement 1) */}
          <button
            onClick={handleInitiateDispatch}
            disabled={isDispatching}
            className="px-4 py-2 rounded-xl bg-gradient-to-r from-indigo-600 to-purple-600 hover:from-indigo-500 hover:to-purple-500 text-white text-xs font-bold flex items-center space-x-2 shadow-lg shadow-indigo-600/30 transition-all cursor-pointer disabled:opacity-50 ring-1 ring-indigo-400/40"
            title="Send customized emails to student and parent email IDs via Microsoft Outlook"
          >
            <Send className="w-4 h-4" />
            <span>Initiate Bulk Email Dispatch</span>
          </button>
        </div>
      </div>

      {/* Top Banner Alert for Dispatch Errors */}
      {dispatchError && (
        <div className="p-4 rounded-xl bg-rose-950/50 border border-rose-800 text-rose-200 text-xs flex items-center space-x-2.5">
          <AlertCircle className="w-4 h-4 shrink-0 text-rose-400" />
          <span>{dispatchError}</span>
        </div>
      )}

      {/* Persistent Dispatch Summary Card (Requirement 5) */}
      {activeDispatchJob && (
        <div className="bg-slate-900/60 border border-slate-800 rounded-2xl p-4 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3 text-xs">
          <div className="flex items-center space-x-3">
            <div className="w-9 h-9 rounded-xl bg-indigo-950 border border-indigo-800 text-indigo-400 flex items-center justify-center shrink-0">
              <ShieldCheck className="w-5 h-5 text-indigo-300" />
            </div>
            <div>
              <div className="flex items-center space-x-2">
                <span className="font-bold text-slate-100">Persistent Dispatch Session</span>
                <span
                  className={`text-[10px] px-2 py-0.5 rounded-full font-bold uppercase tracking-wider ${
                    activeDispatchJob.status === 'completed'
                      ? 'bg-emerald-950 text-emerald-300 border border-emerald-800'
                      : activeDispatchJob.status === 'in_progress'
                      ? 'bg-amber-950 text-amber-300 border border-amber-800 animate-pulse'
                      : 'bg-rose-950 text-rose-300 border border-rose-800'
                  }`}
                >
                  {activeDispatchJob.status}
                </span>
              </div>
              <p className="text-[11px] text-slate-400 mt-0.5">
                Last updated:{' '}
                <strong className="text-slate-300">
                  {new Date(activeDispatchJob.updatedAt).toLocaleTimeString()}
                </strong>{' '}
                • {activeDispatchJob.currentProcessingStatus}
              </p>
            </div>
          </div>

          <div className="flex items-center space-x-4 self-stretch sm:self-auto justify-between sm:justify-end border-t sm:border-t-0 border-slate-800 pt-2 sm:pt-0">
            <div className="text-center">
              <span className="text-[10px] text-slate-400 block">Total</span>
              <span className="font-bold text-slate-200">{activeDispatchJob.totalEmails}</span>
            </div>
            <div className="text-center">
              <span className="text-[10px] text-emerald-400 block">Sent</span>
              <span className="font-bold text-emerald-400">{activeDispatchJob.sentEmails}</span>
            </div>
            <div className="text-center">
              <span className="text-[10px] text-rose-400 block">Failed</span>
              <span className="font-bold text-rose-400">{activeDispatchJob.failedEmails}</span>
            </div>
            <div className="text-center">
              <span className="text-[10px] text-indigo-400 block">Remaining</span>
              <span className="font-bold text-indigo-400">{activeDispatchJob.remainingEmails}</span>
            </div>
          </div>
        </div>
      )}

      {/* Metrics Banner */}
      <div className="grid grid-cols-2 sm:grid-cols-6 gap-3">
        <div className="bg-slate-900/60 border border-slate-800 p-3.5 rounded-xl">
          <span className="text-[11px] text-slate-400 block mb-0.5">Total Students</span>
          <span className="text-xl font-bold text-slate-100">{stats.totalStudents}</span>
        </div>
        <div className="bg-emerald-950/30 border border-emerald-800/50 p-3.5 rounded-xl">
          <span className="text-[11px] text-emerald-300 block mb-0.5">Dispatched: Success</span>
          <span className="text-xl font-bold text-emerald-400">{stats.successCount}</span>
        </div>
        <div className="bg-amber-950/30 border border-amber-800/50 p-3.5 rounded-xl">
          <span className="text-[11px] text-amber-300 block mb-0.5">Partial Success</span>
          <span className="text-xl font-bold text-amber-400">{stats.partialCount}</span>
        </div>
        <div className="bg-rose-950/30 border border-rose-800/50 p-3.5 rounded-xl">
          <span className="text-[11px] text-rose-300 block mb-0.5">Failed Dispatches</span>
          <span className="text-xl font-bold text-rose-400">{stats.failedCount}</span>
        </div>
        <div className="bg-amber-950/20 border border-amber-900/40 p-3.5 rounded-xl">
          <span className="text-[11px] text-amber-400 block mb-0.5">Missing Email ID</span>
          <span className="text-xl font-bold text-amber-400">
            {stats.missingStudentEmail + stats.missingParentEmail}
          </span>
        </div>
        <div className="bg-slate-900/60 border border-slate-800 p-3.5 rounded-xl">
          <span className="text-[11px] text-slate-400 block mb-0.5">No Academic Data</span>
          <span className="text-xl font-bold text-slate-400">{stats.noAcademicDataCount}</span>
        </div>
      </div>

      {/* Search and Filters Bar */}
      <div className="bg-slate-900/40 border border-slate-800 p-4 rounded-xl flex flex-col md:flex-row items-start md:items-center justify-between gap-4 text-xs">
        <div className="relative flex-1 w-full md:max-w-md">
          <Search className="w-3.5 h-3.5 text-slate-500 absolute left-3 top-1/2 -translate-y-1/2" />
          <input
            type="text"
            placeholder="Search student by name or roll number..."
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
            className="w-full pl-8 pr-3 py-1.5 bg-slate-950 border border-slate-800 rounded-xl text-xs text-slate-200 placeholder-slate-500 focus:outline-none focus:border-indigo-500"
          />
        </div>

        <div className="flex flex-wrap items-center gap-1.5">
          <button
            onClick={() => setFilterMode('all')}
            className={`px-3 py-1.5 rounded-lg transition-all cursor-pointer ${
              filterMode === 'all'
                ? 'bg-indigo-600 text-white font-semibold'
                : 'text-slate-400 hover:text-slate-200 bg-slate-950'
            }`}
          >
            All ({generatedReports.length})
          </button>
          <button
            onClick={() => setFilterMode('success')}
            className={`px-3 py-1.5 rounded-lg transition-all cursor-pointer ${
              filterMode === 'success'
                ? 'bg-emerald-600 text-white font-semibold'
                : 'text-emerald-400 hover:text-emerald-300 bg-slate-950'
            }`}
          >
            Success ({stats.successCount})
          </button>
          <button
            onClick={() => setFilterMode('partial')}
            className={`px-3 py-1.5 rounded-lg transition-all cursor-pointer ${
              filterMode === 'partial'
                ? 'bg-amber-600 text-white font-semibold'
                : 'text-amber-400 hover:text-amber-300 bg-slate-950'
            }`}
          >
            Partial Success ({stats.partialCount})
          </button>
          <button
            onClick={() => setFilterMode('failed')}
            className={`px-3 py-1.5 rounded-lg transition-all cursor-pointer ${
              filterMode === 'failed'
                ? 'bg-rose-600 text-white font-semibold'
                : 'text-rose-400 hover:text-rose-300 bg-slate-950'
            }`}
          >
            Failed ({stats.failedCount})
          </button>
          <button
            onClick={() => setFilterMode('missing-email')}
            className={`px-3 py-1.5 rounded-lg transition-all cursor-pointer ${
              filterMode === 'missing-email'
                ? 'bg-amber-700 text-white font-semibold'
                : 'text-amber-400 hover:text-amber-300 bg-slate-950'
            }`}
          >
            Missing Email ({stats.missingStudentEmail + stats.missingParentEmail})
          </button>
          <button
            onClick={() => setFilterMode('no-data')}
            className={`px-3 py-1.5 rounded-lg transition-all cursor-pointer ${
              filterMode === 'no-data'
                ? 'bg-slate-700 text-white font-semibold'
                : 'text-slate-400 hover:text-slate-200 bg-slate-950'
            }`}
          >
            No Academic Data ({stats.noAcademicDataCount})
          </button>
        </div>
      </div>

      {/* Main Consolidated Email Report Grid */}
      <div className="bg-slate-900 border border-slate-800 rounded-2xl shadow-xl overflow-hidden">
        <div className="overflow-x-auto max-h-[640px]">
          <table className="w-full text-left text-xs border-collapse">
            <thead className="bg-slate-950 text-slate-300 sticky top-0 z-20 border-b border-slate-800 text-[11px] uppercase tracking-wider select-none">
              {/* Primary Group Header Row */}
              <tr>
                <th
                  rowSpan={2}
                  className="px-4 py-3 font-bold border-r border-slate-800 sticky left-0 bg-slate-950 z-30 min-w-[130px]"
                >
                  Roll Number
                </th>
                <th
                  rowSpan={2}
                  className="px-4 py-3 font-bold border-r border-slate-800 sticky left-[130px] bg-slate-950 z-30 min-w-[170px]"
                >
                  Student Name
                </th>
                {/* Status Column Header (Requirement 3) */}
                <th
                  rowSpan={2}
                  className="px-4 py-3 font-bold border-r border-slate-800 bg-slate-950 z-20 min-w-[220px]"
                >
                  Status
                </th>

                {/* Stacked Academic Batch Headers */}
                {batches.map((b) => {
                  if (b.type === 'attendance') {
                    const totalCols = b.subjects.length * 3;
                    return (
                      <th
                        key={b.id}
                        colSpan={totalCols}
                        className="px-3 py-2 text-center font-bold border-r border-slate-800 bg-indigo-950/60 text-indigo-200"
                      >
                        📅 Attendance: {b.dateRange.label} ({b.subjects.length} Subjects)
                      </th>
                    );
                  } else {
                    return (
                      <th
                        key={b.id}
                        colSpan={b.evaluations.length}
                        className="px-3 py-2 text-center font-bold border-r border-slate-800 bg-purple-950/60 text-purple-200"
                      >
                        📝 Eval: [{b.subjectCode}] {b.subjectName} ({b.evaluations.length} Components)
                      </th>
                    );
                  }
                })}

                {/* Template Columns */}
                {sendToStudent && (
                  <th
                    rowSpan={2}
                    className="px-4 py-3 font-bold text-center border-r border-slate-800 bg-indigo-950/40 text-indigo-300 min-w-[260px]"
                  >
                    Student Email Template
                  </th>
                )}
                {sendToParent && (
                  <th
                    rowSpan={2}
                    className="px-4 py-3 font-bold text-center border-r border-slate-800 bg-purple-950/40 text-purple-300 min-w-[260px]"
                  >
                    Parent Email Template
                  </th>
                )}
              </tr>

              {/* Secondary Sub-headers for Stacked Batches */}
              <tr>
                {batches.map((b) => {
                  if (b.type === 'attendance') {
                    return b.subjects.map((sub) => (
                      <React.Fragment key={sub.code}>
                        <th
                          className="px-2 py-2 text-center border-r border-slate-800 min-w-[70px]"
                          title={`${sub.name} (${sub.code}) Total Working Days`}
                        >
                          <div className="truncate max-w-[70px] font-mono text-[10px]">{sub.code} Tot</div>
                        </th>
                        <th
                          className="px-2 py-2 text-center border-r border-slate-800 min-w-[70px]"
                          title={`${sub.name} (${sub.code}) Days Attended`}
                        >
                          <div className="truncate max-w-[70px] font-mono text-[10px] text-emerald-400">
                            {sub.code} Att
                          </div>
                        </th>
                        <th
                          className="px-2 py-2 text-center border-r border-slate-800 min-w-[65px]"
                          title={`${sub.name} (${sub.code}) Percentage`}
                        >
                          <div className="truncate max-w-[65px] font-mono text-[10px] text-indigo-400">
                            {sub.code} %
                          </div>
                        </th>
                      </React.Fragment>
                    ));
                  } else {
                    return b.evaluations.map((ev) => (
                      <th
                        key={ev.id}
                        className="px-2 py-2 text-center border-r border-slate-800 min-w-[80px]"
                      >
                        <div className="truncate max-w-[80px]">{ev.name}</div>
                      </th>
                    ));
                  }
                })}
              </tr>
            </thead>

            <tbody className="divide-y divide-slate-800/70">
              {filteredReports.length === 0 ? (
                <tr>
                  <td
                    colSpan={
                      3 +
                      (sendToStudent ? 1 : 0) +
                      (sendToParent ? 1 : 0) +
                      batches.reduce(
                        (acc, b) =>
                          b.type === 'attendance'
                            ? acc + b.subjects.length * 3
                            : acc + b.evaluations.length,
                        0
                      )
                    }
                    className="px-6 py-12 text-center text-slate-500"
                  >
                    No students found matching the selected filter.
                  </td>
                </tr>
              ) : (
                filteredReports.map(
                  ({
                    student,
                    master,
                    hasAcademicRecord,
                    studentEmail,
                    parentEmail,
                    studentEmailIds,
                    parentEmailIds,
                  }) => {
                    const statusInfo = getStudentDispatchStatus(
                      student.rollNumber,
                      hasAcademicRecord,
                      studentEmailIds,
                      parentEmailIds
                    );

                    const inMaster = Boolean(master);
                    let rowStyle = 'hover:bg-slate-800/40 transition-colors';
                    if (!inMaster) {
                      rowStyle = 'bg-rose-950/20 border-l-4 border-l-rose-500 hover:bg-rose-950/35 text-rose-100';
                    } else if (
                      statusInfo.overallStatus === 'Student Email Missing' ||
                      statusInfo.overallStatus === 'Parent Email Missing'
                    ) {
                      rowStyle = 'bg-amber-950/20 border-l-4 border-l-amber-500 hover:bg-amber-950/35 text-amber-100';
                    }

                    const studentOverrideKey = `${student.rollNumber}_student`;
                    const parentOverrideKey = `${student.rollNumber}_parent`;
                    const isExpanded = Boolean(expandedBreakdownRolls[student.rollNumber]);

                    return (
                      <React.Fragment key={student.rollNumber}>
                        <tr className={rowStyle}>
                          {/* Roll Number */}
                          <td className="px-4 py-3 font-mono font-bold text-indigo-300 border-r border-slate-800/60 sticky left-0 bg-slate-950/95 z-10 whitespace-nowrap">
                            {student.rollNumber}
                          </td>

                          {/* Student Name */}
                          <td className="px-4 py-3 font-medium text-slate-200 border-r border-slate-800/60 sticky left-[130px] bg-slate-950/95 z-10 whitespace-nowrap">
                            <div className="flex items-center space-x-1.5">
                              <span>{student.name || master?.name || '—'}</span>
                              {!inMaster && (
                                <span className="text-[10px] px-1.5 py-0.2 rounded bg-rose-950 text-rose-400 border border-rose-800 font-mono">
                                  Not in Master
                                </span>
                              )}
                            </div>
                          </td>

                          {/* Status Column (Requirement 3: Dynamic Outcome & Recipient Breakdown) */}
                          <td className="px-3.5 py-3 border-r border-slate-800/60 text-xs">
                            <div className="space-y-1.5">
                              <div className="flex items-center justify-between gap-1.5">
                                {renderStatusBadge(statusInfo.overallStatus)}
                                {statusInfo.record && (
                                  <button
                                    onClick={() => toggleRowBreakdown(student.rollNumber)}
                                    className="p-1 rounded hover:bg-slate-800 text-slate-400 hover:text-slate-200 transition-colors cursor-pointer text-[10px] flex items-center space-x-0.5"
                                    title="Toggle recipient breakdown"
                                  >
                                    <span>Details</span>
                                    {isExpanded ? (
                                      <ChevronUp className="w-3 h-3" />
                                    ) : (
                                      <ChevronDown className="w-3 h-3" />
                                    )}
                                  </button>
                                )}
                              </div>

                              {/* Missing records / emails explanatory message */}
                              {statusInfo.missingReason && (
                                <p className="text-[11px] text-amber-300/80 leading-tight">
                                  {statusInfo.missingReason}
                                </p>
                              )}

                              {/* Direct visible recipient outcome breakdown */}
                              {statusInfo.record && (
                                <div className="space-y-1 pt-1 text-[11px]">
                                  {statusInfo.record.studentEmails.length > 0 && (
                                    <div>
                                      <span className="text-[10px] font-bold text-indigo-400 uppercase tracking-wider block">
                                        Student Emails ({statusInfo.record.studentEmails.length}):
                                      </span>
                                      {statusInfo.record.studentEmails.map((se, idx) => (
                                        <div key={idx} className="font-mono text-[10px] leading-tight pl-1">
                                          <span className="text-slate-300">{se.email}: </span>
                                          {se.status === 'Success' ? (
                                            <span className="text-emerald-400 font-bold">Success</span>
                                          ) : (
                                            <span className="text-rose-400 font-bold">
                                              Failed{' '}
                                              {se.reason && (
                                                <span className="text-rose-300/90 font-sans block text-[10px]">
                                                  Reason: {se.reason}
                                                </span>
                                              )}
                                            </span>
                                          )}
                                        </div>
                                      ))}
                                    </div>
                                  )}

                                  {statusInfo.record.parentEmails.length > 0 && (
                                    <div className="pt-0.5">
                                      <span className="text-[10px] font-bold text-purple-400 uppercase tracking-wider block">
                                        Parent Emails ({statusInfo.record.parentEmails.length}):
                                      </span>
                                      {statusInfo.record.parentEmails.map((pe, idx) => (
                                        <div key={idx} className="font-mono text-[10px] leading-tight pl-1">
                                          <span className="text-slate-300">{pe.email}: </span>
                                          {pe.status === 'Success' ? (
                                            <span className="text-emerald-400 font-bold">Success</span>
                                          ) : (
                                            <span className="text-rose-400 font-bold">
                                              Failed{' '}
                                              {pe.reason && (
                                                <span className="text-rose-300/90 font-sans block text-[10px]">
                                                  Reason: {pe.reason}
                                                </span>
                                              )}
                                            </span>
                                          )}
                                        </div>
                                      ))}
                                    </div>
                                  )}
                                </div>
                              )}
                            </div>
                          </td>

                          {/* Stacked Academic Data Cells */}
                          {batches.map((b) => {
                            if (b.type === 'attendance') {
                              return b.subjects.map((sub) => {
                                const key = `${b.id}_${sub.code}`;
                                const rec = student.attendance[key];
                                return (
                                  <React.Fragment key={key}>
                                    <td className="px-2 py-2 text-center border-r border-slate-800/60 font-mono text-xs">
                                      {rec?.conducted ?? '—'}
                                    </td>
                                    <td className="px-2 py-2 text-center border-r border-slate-800/60 font-mono text-xs font-semibold">
                                      {rec?.attended ?? '—'}
                                    </td>
                                    <td className="px-2 py-2 text-center border-r border-slate-800 font-mono text-xs">
                                      {rec?.percentage !== null && rec?.percentage !== undefined
                                        ? `${rec.percentage}%`
                                        : '—'}
                                    </td>
                                  </React.Fragment>
                                );
                              });
                            } else {
                              return b.evaluations.map((ev) => {
                                const cellKey = `${b.id}_${ev.id}`;
                                const rec = student.evaluations[cellKey];
                                return (
                                  <td
                                    key={cellKey}
                                    className="px-2 py-2 text-center border-r border-slate-800 font-mono text-xs"
                                  >
                                    {rec?.isAbsent ? (
                                      <span className="text-rose-400 font-bold text-[10px]">ABS</span>
                                    ) : rec?.marksObtained !== null && rec?.marksObtained !== undefined ? (
                                      <span>{rec.marksObtained}</span>
                                    ) : (
                                      '—'
                                    )}
                                  </td>
                                );
                              });
                            }
                          })}

                          {/* Student Email Template Column (Requirement 2 & 6) */}
                          {sendToStudent && (
                            <td className="px-3 py-2 border-r border-slate-800/60 align-top max-w-[280px]">
                              {studentEmail ? (
                                <div className="space-y-1.5">
                                  <div className="flex items-center justify-between text-[11px]">
                                    <span className="font-semibold text-indigo-300 truncate max-w-[170px]">
                                      {studentEmail.subject}
                                    </span>
                                    <button
                                      onClick={() =>
                                        setActiveReviewEmail({
                                          student,
                                          recipientType: 'student',
                                          email: {
                                            ...studentEmail,
                                            bodyText:
                                              customizedBodies[studentOverrideKey] || studentEmail.bodyText,
                                          },
                                        })
                                      }
                                      className="px-2 py-0.5 rounded bg-indigo-950 hover:bg-indigo-900 border border-indigo-800/80 text-indigo-300 text-[10px] font-medium flex items-center space-x-1 cursor-pointer transition-colors"
                                    >
                                      <Eye className="w-3 h-3" />
                                      <span>Review &amp; Edit</span>
                                    </button>
                                  </div>
                                  <div className="bg-slate-950/80 p-2 rounded-lg border border-slate-800 text-[10px] font-mono text-slate-300 line-clamp-3 leading-relaxed">
                                    {customizedBodies[studentOverrideKey] || studentEmail.bodyText}
                                  </div>
                                  <div className="flex items-center justify-between text-[10px] text-slate-400">
                                    <span>
                                      {studentEmailIds.length > 0 ? (
                                        <span className="text-emerald-400 font-mono">
                                          {studentEmailIds.length} Email(s)
                                        </span>
                                      ) : (
                                        <span className="text-amber-400">Missing Email ID</span>
                                      )}
                                    </span>
                                    <button
                                      onClick={() =>
                                        handleCopy(
                                          customizedBodies[studentOverrideKey] || studentEmail.bodyText,
                                          studentOverrideKey
                                        )
                                      }
                                      className="text-slate-400 hover:text-slate-200 flex items-center space-x-1 cursor-pointer"
                                    >
                                      {copiedKey === studentOverrideKey ? (
                                        <>
                                          <Check className="w-3 h-3 text-emerald-400" />
                                          <span className="text-emerald-400">Copied</span>
                                        </>
                                      ) : (
                                        <>
                                          <Copy className="w-3 h-3" />
                                          <span>Copy</span>
                                        </>
                                      )}
                                    </button>
                                  </div>
                                </div>
                              ) : (
                                <div className="py-4 text-center text-slate-500 italic text-[11px]">
                                  No academic record available to generate student email.
                                </div>
                              )}
                            </td>
                          )}

                          {/* Parent Email Template Column (Requirement 2 & 6) */}
                          {sendToParent && (
                            <td className="px-3 py-2 border-r border-slate-800/60 align-top max-w-[280px]">
                              {parentEmail ? (
                                <div className="space-y-1.5">
                                  <div className="flex items-center justify-between text-[11px]">
                                    <span className="font-semibold text-purple-300 truncate max-w-[170px]">
                                      {parentEmail.subject}
                                    </span>
                                    <button
                                      onClick={() =>
                                        setActiveReviewEmail({
                                          student,
                                          recipientType: 'parent',
                                          email: {
                                            ...parentEmail,
                                            bodyText:
                                              customizedBodies[parentOverrideKey] || parentEmail.bodyText,
                                          },
                                        })
                                      }
                                      className="px-2 py-0.5 rounded bg-purple-950 hover:bg-purple-900 border border-purple-800/80 text-purple-300 text-[10px] font-medium flex items-center space-x-1 cursor-pointer transition-colors"
                                    >
                                      <Eye className="w-3 h-3" />
                                      <span>Review &amp; Edit</span>
                                    </button>
                                  </div>
                                  <div className="bg-slate-950/80 p-2 rounded-lg border border-slate-800 text-[10px] font-mono text-slate-300 line-clamp-3 leading-relaxed">
                                    {customizedBodies[parentOverrideKey] || parentEmail.bodyText}
                                  </div>
                                  <div className="flex items-center justify-between text-[10px] text-slate-400">
                                    <span>
                                      {parentEmailIds.length > 0 ? (
                                        <span className="text-purple-400 font-mono">
                                          {parentEmailIds.length} Parent Email(s)
                                        </span>
                                      ) : (
                                        <span className="text-amber-400">Missing Parent Email</span>
                                      )}
                                    </span>
                                    <button
                                      onClick={() =>
                                        handleCopy(
                                          customizedBodies[parentOverrideKey] || parentEmail.bodyText,
                                          parentOverrideKey
                                        )
                                      }
                                      className="text-slate-400 hover:text-slate-200 flex items-center space-x-1 cursor-pointer"
                                    >
                                      {copiedKey === parentOverrideKey ? (
                                        <>
                                          <Check className="w-3 h-3 text-emerald-400" />
                                          <span className="text-emerald-400">Copied</span>
                                        </>
                                      ) : (
                                        <>
                                          <Copy className="w-3 h-3" />
                                          <span>Copy</span>
                                        </>
                                      )}
                                    </button>
                                  </div>
                                </div>
                              ) : (
                                <div className="py-4 text-center text-slate-500 italic text-[11px]">
                                  No academic record available to generate parent email.
                                </div>
                              )}
                            </td>
                          )}
                        </tr>

                        {/* Expandable Recipient Detail Audit Row */}
                        {isExpanded && statusInfo.record && (
                          <tr className="bg-slate-950/90 border-b border-slate-800">
                            <td
                              colSpan={
                                3 +
                                (sendToStudent ? 1 : 0) +
                                (sendToParent ? 1 : 0) +
                                batches.reduce(
                                  (acc, b) =>
                                    b.type === 'attendance'
                                      ? acc + b.subjects.length * 3
                                      : acc + b.evaluations.length,
                                  0
                                )
                              }
                              className="px-6 py-3"
                            >
                              <div className="bg-slate-900 border border-slate-800 rounded-xl p-3.5 space-y-3">
                                <div className="flex items-center justify-between">
                                  <span className="font-bold text-slate-200 flex items-center space-x-2">
                                    <Info className="w-4 h-4 text-indigo-400" />
                                    <span>
                                      Recipient-Wise Dispatch Breakdown for {student.name} ({student.rollNumber})
                                    </span>
                                  </span>
                                  <span className="text-xs">{renderStatusBadge(statusInfo.overallStatus)}</span>
                                </div>

                                <div className="grid grid-cols-1 md:grid-cols-2 gap-4 text-xs font-mono">
                                  {/* Student Recipients */}
                                  <div className="bg-slate-950 p-3 rounded-lg border border-slate-800 space-y-1.5">
                                    <span className="font-bold text-indigo-300 block font-sans text-xs">
                                      Student Email Recipients:
                                    </span>
                                    {statusInfo.record.studentEmails.length === 0 ? (
                                      <p className="text-slate-500 font-sans italic text-[11px]">
                                        No registered student email address.
                                      </p>
                                    ) : (
                                      statusInfo.record.studentEmails.map((se, i) => (
                                        <div
                                          key={i}
                                          className="p-2 rounded bg-slate-900 border border-slate-800 space-y-0.5"
                                        >
                                          <div className="flex items-center justify-between">
                                            <span className="text-slate-200">{se.email}</span>
                                            {se.status === 'Success' ? (
                                              <span className="text-emerald-400 font-bold">Success</span>
                                            ) : (
                                              <span className="text-rose-400 font-bold">Failed</span>
                                            )}
                                          </div>
                                          {se.reason && (
                                            <p className="text-rose-400 text-[10px] font-sans">
                                              Reason: {se.reason}
                                            </p>
                                          )}
                                          {se.sentAt && (
                                            <p className="text-slate-500 text-[9px] font-sans">
                                              Dispatched at: {new Date(se.sentAt).toLocaleTimeString()}
                                            </p>
                                          )}
                                        </div>
                                      ))
                                    )}
                                  </div>

                                  {/* Parent Recipients */}
                                  <div className="bg-slate-950 p-3 rounded-lg border border-slate-800 space-y-1.5">
                                    <span className="font-bold text-purple-300 block font-sans text-xs">
                                      Parent Email Recipients:
                                    </span>
                                    {statusInfo.record.parentEmails.length === 0 ? (
                                      <p className="text-slate-500 font-sans italic text-[11px]">
                                        No registered parent/guardian email address.
                                      </p>
                                    ) : (
                                      statusInfo.record.parentEmails.map((pe, i) => (
                                        <div
                                          key={i}
                                          className="p-2 rounded bg-slate-900 border border-slate-800 space-y-0.5"
                                        >
                                          <div className="flex items-center justify-between">
                                            <span className="text-slate-200">{pe.email}</span>
                                            {pe.status === 'Success' ? (
                                              <span className="text-emerald-400 font-bold">Success</span>
                                            ) : (
                                              <span className="text-rose-400 font-bold">Failed</span>
                                            )}
                                          </div>
                                          {pe.reason && (
                                            <p className="text-rose-400 text-[10px] font-sans">
                                              Reason: {pe.reason}
                                            </p>
                                          )}
                                          {pe.sentAt && (
                                            <p className="text-slate-500 text-[9px] font-sans">
                                              Dispatched at: {new Date(pe.sentAt).toLocaleTimeString()}
                                            </p>
                                          )}
                                        </div>
                                      ))
                                    )}
                                  </div>
                                </div>
                              </div>
                            </td>
                          </tr>
                        )}
                      </React.Fragment>
                    );
                  }
                )
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* Review & Edit Email Modal */}
      {activeReviewEmail && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/80 backdrop-blur-sm animate-in fade-in duration-200">
          <div className="bg-slate-900 border border-slate-700 rounded-2xl w-full max-w-2xl max-h-[90vh] shadow-2xl flex flex-col overflow-hidden">
            {/* Header */}
            <div className="px-6 py-4 border-b border-slate-800 flex items-center justify-between bg-slate-950/70">
              <div className="flex items-center space-x-3">
                <div
                  className={`p-2 rounded-xl ${
                    activeReviewEmail.recipientType === 'student'
                      ? 'bg-indigo-950 border border-indigo-800 text-indigo-300'
                      : 'bg-purple-950 border border-purple-800 text-purple-300'
                  }`}
                >
                  <Mail className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="font-bold text-slate-100 text-base">
                    Review {activeReviewEmail.recipientType === 'student' ? 'Student' : 'Parent'} Email
                  </h3>
                  <p className="text-xs text-slate-400">
                    Recipient: <strong className="text-slate-200">{activeReviewEmail.student.name}</strong> (
                    <span className="font-mono text-indigo-300">
                      {activeReviewEmail.student.rollNumber}
                    </span>
                    )
                  </p>
                </div>
              </div>
              <button
                onClick={() => setActiveReviewEmail(null)}
                className="text-slate-400 hover:text-white p-1 rounded-lg hover:bg-slate-800 cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Subject Preview */}
            <div className="px-6 py-3 bg-slate-950/40 border-b border-slate-800 flex items-center justify-between text-xs">
              <span className="text-slate-400">
                Subject:{' '}
                <strong className="text-slate-200">{activeReviewEmail.email.subject}</strong>
              </span>
              <span className="text-slate-400">
                Tone:{' '}
                <strong className="text-slate-200">
                  {activeReviewEmail.recipientType === 'student'
                    ? 'Informative & Encouraging'
                    : 'Formal & Collaborative'}
                </strong>
              </span>
            </div>

            {/* Content Body (Editable Text Area) */}
            <div className="p-6 overflow-y-auto space-y-4 flex-1">
              <div>
                <div className="flex items-center justify-between mb-1.5">
                  <label className="text-xs font-bold text-slate-200 flex items-center space-x-1.5">
                    <Edit3 className="w-3.5 h-3.5 text-indigo-400" />
                    <span>Customized Email Content (Editable per Student)</span>
                  </label>
                  <button
                    onClick={() => handleCopy(activeReviewEmail.email.bodyText, 'modal_copy')}
                    className="text-xs text-indigo-400 hover:text-indigo-300 flex items-center space-x-1 cursor-pointer font-medium"
                  >
                    {copiedKey === 'modal_copy' ? (
                      <>
                        <Check className="w-3 h-3 text-emerald-400" />
                        <span className="text-emerald-400">Copied!</span>
                      </>
                    ) : (
                      <>
                        <Copy className="w-3 h-3" />
                        <span>Copy Email Text</span>
                      </>
                    )}
                  </button>
                </div>

                <textarea
                  rows={14}
                  value={activeReviewEmail.email.bodyText}
                  onChange={(e) =>
                    setActiveReviewEmail({
                      ...activeReviewEmail,
                      email: {
                        ...activeReviewEmail.email,
                        bodyText: e.target.value,
                      },
                    })
                  }
                  className="w-full p-4 bg-slate-950 border border-slate-800 rounded-xl text-xs text-slate-100 font-mono leading-relaxed focus:outline-none focus:border-indigo-500 shadow-inner"
                />
              </div>

              {/* Dynamic availability notice */}
              <div className="p-3 bg-slate-950/60 rounded-xl border border-slate-800 flex items-center justify-between text-xs text-slate-400">
                <span>
                  Included Records:{' '}
                  <strong className="text-slate-200">
                    {activeReviewEmail.email.hasAttendance && activeReviewEmail.email.hasEvaluation
                      ? 'Attendance + Evaluation'
                      : activeReviewEmail.email.hasAttendance
                      ? 'Attendance Only'
                      : 'Evaluation Only'}
                  </strong>
                </span>
                <span className="text-slate-500 text-[11px]">
                  Signature verified: KR Mangalam University
                </span>
              </div>
            </div>

            {/* Footer */}
            <div className="px-6 py-4 bg-slate-950 border-t border-slate-800 flex items-center justify-between">
              <button
                type="button"
                onClick={() => setActiveReviewEmail(null)}
                className="px-4 py-2 rounded-xl text-xs font-medium text-slate-400 hover:text-slate-200 hover:bg-slate-800 cursor-pointer"
              >
                Close
              </button>

              <button
                type="button"
                onClick={handleSaveCustomOverride}
                className="px-5 py-2 rounded-xl bg-indigo-600 hover:bg-indigo-500 text-white text-xs font-bold shadow-md shadow-indigo-600/30 cursor-pointer"
              >
                Save Customized Changes
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Microsoft Outlook Configuration Notice Modal (Requirement 1) */}
      {showOutlookNoticeModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/80 backdrop-blur-sm animate-in fade-in duration-200">
          <div className="bg-slate-900 border border-amber-500/70 rounded-2xl w-full max-w-lg shadow-2xl p-6 space-y-4">
            <div className="flex items-start space-x-3.5">
              <div className="w-11 h-11 rounded-xl bg-amber-950/90 border border-amber-700/70 text-amber-400 flex items-center justify-center shrink-0 shadow-lg">
                <AlertTriangle className="w-6 h-6" />
              </div>
              <div className="space-y-2">
                <h3 className="font-bold text-slate-100 text-base">
                  Microsoft Outlook Account Required
                </h3>
                {/* Exact mandatory text specified in Requirement 1 */}
                <p className="text-xs text-amber-200/90 font-medium leading-relaxed bg-amber-950/40 p-3 rounded-xl border border-amber-800/40">
                  Please configure your Microsoft Outlook email account with your application account
                  before initiating bulk email communication. Once configured, please return to this
                  section and initiate the bulk email dispatch.
                </p>
              </div>
            </div>

            <div className="pt-2 flex justify-end space-x-3">
              <button
                onClick={() => setShowOutlookNoticeModal(false)}
                className="px-5 py-2 rounded-xl text-xs font-semibold bg-slate-800 hover:bg-slate-700 text-slate-200 transition-colors cursor-pointer"
              >
                Understood &amp; Close
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Application Locking and Real-Time Dispatch Progress Modal (Requirement 4) */}
      {isDispatching && (
        <div
          className="fixed inset-0 z-[9999] bg-slate-950/95 backdrop-blur-md flex items-center justify-center p-4 select-none pointer-events-auto"
          role="dialog"
          aria-modal="true"
        >
          <div className="bg-slate-900 border border-slate-700 rounded-3xl w-full max-w-xl shadow-2xl p-6 sm:p-8 space-y-6">
            {/* Header with animated pulse */}
            <div className="flex items-center space-x-4">
              <div className="w-12 h-12 rounded-2xl bg-indigo-600/20 border border-indigo-500/50 text-indigo-400 flex items-center justify-center shadow-lg shadow-indigo-600/30">
                <RefreshCw className="w-6 h-6 animate-spin text-indigo-400" />
              </div>
              <div>
                <div className="flex items-center space-x-2">
                  <h3 className="font-bold text-slate-100 text-lg">Bulk Email Dispatch in Progress</h3>
                  <span className="text-[10px] px-2 py-0.5 rounded-full font-bold bg-indigo-950 text-indigo-300 border border-indigo-800 flex items-center space-x-1">
                    <Lock className="w-2.5 h-2.5" />
                    <span>App Locked</span>
                  </span>
                </div>
                <p className="text-xs text-slate-400 mt-0.5">
                  Controlled server-side dispatch executing via Microsoft Outlook.
                </p>
              </div>
            </div>

            {/* Real-time Dispatch Progress Metrics Grid (Requirement 4) */}
            <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 text-center">
              <div className="bg-slate-950 p-3 rounded-2xl border border-slate-800">
                <span className="text-[11px] text-slate-400 block mb-0.5">Total Emails</span>
                <span className="text-xl font-bold text-slate-100">
                  {activeDispatchJob?.totalEmails || 0}
                </span>
              </div>
              <div className="bg-emerald-950/40 p-3 rounded-2xl border border-emerald-800/60">
                <span className="text-[11px] text-emerald-300 block mb-0.5">Successfully Sent</span>
                <span className="text-xl font-bold text-emerald-400">
                  {activeDispatchJob?.sentEmails || 0}
                </span>
              </div>
              <div className="bg-rose-950/40 p-3 rounded-2xl border border-rose-800/60">
                <span className="text-[11px] text-rose-300 block mb-0.5">Failed Emails</span>
                <span className="text-xl font-bold text-rose-400">
                  {activeDispatchJob?.failedEmails || 0}
                </span>
              </div>
              <div className="bg-indigo-950/40 p-3 rounded-2xl border border-indigo-800/60">
                <span className="text-[11px] text-indigo-300 block mb-0.5">Remaining</span>
                <span className="text-xl font-bold text-indigo-400">
                  {activeDispatchJob?.remainingEmails || 0}
                </span>
              </div>
            </div>

            {/* Progress Bar */}
            <div className="space-y-1.5">
              <div className="flex justify-between text-xs text-slate-400">
                <span>Dispatch Progress</span>
                <span className="font-mono text-indigo-400 font-bold">
                  {activeDispatchJob?.totalEmails
                    ? Math.round(
                        (((activeDispatchJob.sentEmails || 0) + (activeDispatchJob.failedEmails || 0)) /
                          activeDispatchJob.totalEmails) *
                          100
                      )
                    : 0}
                  %
                </span>
              </div>
              <div className="w-full bg-slate-950 rounded-full h-3 overflow-hidden border border-slate-800 p-0.5">
                <div
                  className="bg-gradient-to-r from-indigo-500 via-purple-500 to-emerald-500 h-2 rounded-full transition-all duration-300 shadow-sm"
                  style={{
                    width: `${
                      activeDispatchJob?.totalEmails
                        ? (((activeDispatchJob.sentEmails || 0) + (activeDispatchJob.failedEmails || 0)) /
                            activeDispatchJob.totalEmails) *
                          100
                        : 0
                    }%`,
                  }}
                />
              </div>
            </div>

            {/* Current Processing Status (Requirement 4) */}
            <div className="bg-slate-950 p-3.5 rounded-xl border border-slate-800 space-y-1">
              <span className="text-[10px] text-slate-500 uppercase tracking-wider block font-bold">
                Current Activity
              </span>
              <p className="text-xs text-slate-200 font-mono truncate">
                {activeDispatchJob?.currentProcessingStatus || 'Processing dispatch queue...'}
              </p>
            </div>

            {/* Mandatory Application Closure Warning (Requirement 4) */}
            <div className="p-4 rounded-2xl bg-amber-950/60 border border-amber-700/70 text-amber-200 space-y-2">
              <div className="flex items-center space-x-2 text-amber-300 font-bold text-xs uppercase tracking-wide">
                <AlertTriangle className="w-4 h-4 shrink-0" />
                <span>Application Closure Warning</span>
              </div>
              <p className="text-xs text-amber-100/90 leading-relaxed">
                Please do not close or exit the application while bulk email communication is in
                progress. Closing the application may interrupt the dispatch process, resulting in
                incomplete work and loss of unsaved data. Kindly wait until the entire process is
                completed.
              </p>
            </div>

            {/* Hard Stop / Interrupt Process Button */}
            <div className="pt-3 border-t border-slate-800 flex flex-col sm:flex-row items-center justify-between gap-3">
              <div className="text-left">
                <span className="text-xs font-semibold text-slate-300 block">Emergency Dispatch Control</span>
                <span className="text-[11px] text-slate-500 block">
                  Immediately halts sending remaining emails and unlocks the application.
                </span>
              </div>
              <button
                type="button"
                onClick={handleStopDispatch}
                disabled={isStopping}
                className="w-full sm:w-auto px-5 py-2.5 rounded-xl bg-gradient-to-r from-rose-600 to-red-700 hover:from-rose-500 hover:to-red-600 text-white font-bold text-xs flex items-center justify-center space-x-2 shadow-xl shadow-rose-900/40 ring-1 ring-rose-400/50 transition-all cursor-pointer disabled:opacity-50"
              >
                {isStopping ? (
                  <>
                    <RefreshCw className="w-4 h-4 animate-spin" />
                    <span>Interrupting Dispatch...</span>
                  </>
                ) : (
                  <>
                    <XCircle className="w-4 h-4 text-white" />
                    <span>Hard Stop / Interrupt Dispatch</span>
                  </>
                )}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
