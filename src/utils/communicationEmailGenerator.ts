/**
 * Bulk Communication Email Generator Utility
 *
 * Implements Institutional Functional Requirements:
 * - Dynamic generation based on data availability (Both, Attendance Only, Evaluation Only, Neither).
 * - Differentiated language and tone for Student vs Parent emails:
 *   - Student Email: Informative, encouraging, and advisory.
 *   - Parent Email: Formal, informative, and collaborative.
 * - Granular Attendance Breakdown: Subject Name, Subject Code, Period, Total Days, Days Attended, Percentage.
 * - Granular Evaluation Breakdown: Subject Name, Subject Code, Component Name, Marks Obtained / Absent, Max Marks.
 * - Mandatory Teacher Signature:
 *     Regards,
 *     [Teacher's Name]
 *     [Teacher's Designation/Role]
 *     KR Mangalam University
 */

export interface AttendanceSubjectItem {
  subjectName: string;
  subjectCode: string;
  period: string;
  conducted: number | null;
  attended: number | null;
  percentage: number | null;
}

export interface EvaluationAssessmentItem {
  subjectName: string;
  subjectCode: string;
  componentName: string;
  marksObtained: number | null;
  maxMarks: number;
  isAbsent: boolean;
  percentage: number | null;
}

export interface TeacherSenderInfo {
  name: string;
  designation?: string;
  department?: string;
  email?: string;
}

export interface StudentRecipientData {
  rollNumber: string;
  name: string;
  section?: string;
  studentEmails: string[];
  parentName?: string;
  parentEmails: string[];
}

export interface GeneratedEmail {
  recipientType: 'student' | 'parent';
  recipientName: string;
  recipientEmails: string[];
  subject: string;
  bodyText: string;
  bodyHtml: string;
  hasAttendance: boolean;
  hasEvaluation: boolean;
  isAvailable: boolean;
}

export function formatAttendanceSectionText(items: AttendanceSubjectItem[]): string {
  if (items.length === 0) return '';

  const rows = items
    .map((item, idx) => {
      const conductedStr = item.conducted !== null ? `${item.conducted} days` : 'N/A';
      const attendedStr = item.attended !== null ? `${item.attended} days` : 'N/A';
      const pctStr = item.percentage !== null ? `${item.percentage}%` : 'N/A';

      return `  ${idx + 1}. ${item.subjectName} (${item.subjectCode})
     • Attendance Period: ${item.period || 'Current Semester'}
     • Total Working Days: ${conductedStr}
     • Number of Days Attended: ${attendedStr}
     • Attendance Percentage: ${pctStr}`;
    })
    .join('\n\n');

  return `--- ATTENDANCE RECORD SUMMARY ---\n${rows}`;
}

export function formatEvaluationSectionText(items: EvaluationAssessmentItem[]): string {
  if (items.length === 0) return '';

  const rows = items
    .map((item, idx) => {
      const scoreStr = item.isAbsent
        ? 'ABSENT'
        : item.marksObtained !== null
        ? `${item.marksObtained} / ${item.maxMarks}`
        : 'N/A';
      const pctStr = item.isAbsent
        ? '0%'
        : item.percentage !== null
        ? `${item.percentage}%`
        : 'N/A';

      return `  ${idx + 1}. ${item.subjectName} (${item.subjectCode})
     • Evaluation Component: ${item.componentName}
     • Marks Obtained: ${scoreStr}
     • Maximum Marks: ${item.maxMarks}
     • Component Score: ${pctStr}`;
    })
    .join('\n\n');

  return `--- EVALUATION & ASSESSMENT SUMMARY ---\n${rows}`;
}

export function buildTeacherSignature(teacher: TeacherSenderInfo): string {
  const teacherName = teacher.name?.trim() || 'Faculty Member';
  const designation = teacher.department
    ? `Faculty, Department of ${teacher.department}`
    : teacher.designation?.trim() || 'Faculty Member';

  return `Regards,\n${teacherName}\n${designation}\nKR Mangalam University`;
}

/**
 * Generates an individualized Student email.
 * Tone: Informative, encouraging, and advisory.
 */
export function generateStudentEmail(
  student: StudentRecipientData,
  attendanceItems: AttendanceSubjectItem[],
  evaluationItems: EvaluationAssessmentItem[],
  teacher: TeacherSenderInfo
): GeneratedEmail | null {
  const hasAttendance = attendanceItems.length > 0;
  const hasEvaluation = evaluationItems.length > 0;

  // Rule: If neither record is available, no email shall be generated
  if (!hasAttendance && !hasEvaluation) {
    return null;
  }

  const studentName = student.name || 'Student';
  const roll = student.rollNumber || 'N/A';
  const subjectLine = `Academic Performance & Attendance Update | Roll No: ${roll} | KR Mangalam University`;

  // Student specific advisory commentary
  let attendanceCommentary = '';
  if (hasAttendance) {
    const hasLowAttendance = attendanceItems.some(
      (item) => item.percentage !== null && item.percentage < 75
    );
    if (hasLowAttendance) {
      attendanceCommentary =
        '\nAdvisory Note on Attendance:\nPlease be mindful that KR Mangalam University mandates a minimum of 75% attendance in each registered course. One or more of your subjects currently falls below this required threshold. We strongly urge you to attend all upcoming lectures, laboratory sessions, and tutorials diligently to avoid academic penalties or examination debarment.\n';
    } else {
      attendanceCommentary =
        '\nCommendation on Attendance:\nCommendable effort maintaining your attendance above the mandatory 75% threshold. Consistent classroom engagement is fundamental to your ongoing academic success.\n';
    }
  }

  let evaluationCommentary = '';
  if (hasEvaluation) {
    const hasAbsent = evaluationItems.some((e) => e.isAbsent);
    const hasLowMarks = evaluationItems.some(
      (e) => !e.isAbsent && e.percentage !== null && e.percentage < 50
    );

    if (hasAbsent || hasLowMarks) {
      evaluationCommentary =
        '\nAcademic Advisory & Support:\nReview the marks obtained in the assessments listed below. If you missed any scheduled evaluation or need clarification on specific concepts, please schedule time during faculty office hours. We are here to support your learning curve and assist you in excelling in forthcoming midterms and finals.\n';
    } else {
      evaluationCommentary =
        '\nPerformance Feedback:\nYour performance across the evaluated components reflects consistent effort and preparation. Continue this momentum in the upcoming curricular assessments.\n';
    }
  }

  const sections: string[] = [];
  sections.push(`Dear ${studentName} (Roll No: ${roll}),`);
  sections.push(
    `I am writing to share your official academic performance and institutional attendance report at KR Mangalam University. Please review the breakdown below:`
  );

  if (hasAttendance) {
    sections.push(formatAttendanceSectionText(attendanceItems));
    if (attendanceCommentary) sections.push(attendanceCommentary.trim());
  }

  if (hasEvaluation) {
    sections.push(formatEvaluationSectionText(evaluationItems));
    if (evaluationCommentary) sections.push(evaluationCommentary.trim());
  }

  sections.push(
    `Should you require academic counseling, assignment review, or topic reinforcement, do not hesitate to contact me or visit the departmental office.`
  );
  sections.push(buildTeacherSignature(teacher));

  const bodyText = sections.join('\n\n');

  // HTML format for rich view / Outlook sendMail
  const bodyHtml = `
<div style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; line-height: 1.6; color: #1e293b; max-width: 680px; margin: 0 auto; padding: 20px; border: 1px solid #e2e8f0; border-radius: 12px; background-color: #ffffff;">
  <div style="border-bottom: 2px solid #4f46e5; padding-bottom: 12px; margin-bottom: 20px;">
    <h2 style="color: #1e1b4b; margin: 0 0 4px 0; font-size: 20px;">KR Mangalam University</h2>
    <p style="color: #64748b; margin: 0; font-size: 13px;">Academic Performance & Attendance Communication Portal</p>
  </div>

  <p style="font-size: 15px; font-weight: 600; color: #0f172a;">Dear ${studentName} <span style="font-family: monospace; color: #4338ca;">(Roll No: ${roll})</span>,</p>
  <p style="font-size: 14px; color: #334155;">I am writing to provide you with your official academic performance and institutional attendance summary. Please review the details below carefully:</p>

  ${
    hasAttendance
      ? `
  <div style="margin: 20px 0; padding: 16px; background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px;">
    <h3 style="margin: 0 0 12px 0; font-size: 14px; color: #1e293b; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 1px solid #cbd5e1; padding-bottom: 6px;">📅 Attendance Record Summary</h3>
    <table style="width: 100%; border-collapse: collapse; font-size: 13px;">
      <thead>
        <tr style="background-color: #e2e8f0; color: #475569; text-align: left;">
          <th style="padding: 8px 10px;">Subject & Code</th>
          <th style="padding: 8px 10px;">Period</th>
          <th style="padding: 8px 10px; text-align: center;">Total (D)</th>
          <th style="padding: 8px 10px; text-align: center;">Attended (A)</th>
          <th style="padding: 8px 10px; text-align: center;">%</th>
        </tr>
      </thead>
      <tbody>
        ${attendanceItems
          .map((item) => {
            const pct = item.percentage ?? 0;
            const pctColor = pct >= 75 ? '#16a34a' : pct >= 60 ? '#d97706' : '#dc2626';
            return `
          <tr style="border-bottom: 1px solid #e2e8f0;">
            <td style="padding: 8px 10px; font-weight: 500;">${item.subjectName} <span style="color: #64748b; font-family: monospace;">(${item.subjectCode})</span></td>
            <td style="padding: 8px 10px; color: #64748b; font-size: 12px;">${item.period}</td>
            <td style="padding: 8px 10px; text-align: center; font-family: monospace;">${item.conducted ?? '—'}</td>
            <td style="padding: 8px 10px; text-align: center; font-family: monospace; font-weight: 600;">${item.attended ?? '—'}</td>
            <td style="padding: 8px 10px; text-align: center; font-family: monospace; font-weight: 700; color: ${pctColor};">${item.percentage !== null ? item.percentage + '%' : '—'}</td>
          </tr>`;
          })
          .join('')}
      </tbody>
    </table>
    ${attendanceCommentary ? `<p style="margin: 12px 0 0 0; font-size: 12px; color: #475569; line-height: 1.5;">${attendanceCommentary.replace(/\n/g, '<br/>')}</p>` : ''}
  </div>`
      : ''
  }

  ${
    hasEvaluation
      ? `
  <div style="margin: 20px 0; padding: 16px; background-color: #faf5ff; border: 1px solid #f3e8ff; border-radius: 8px;">
    <h3 style="margin: 0 0 12px 0; font-size: 14px; color: #581c87; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 1px solid #e9d5ff; padding-bottom: 6px;">📝 Evaluation & Assessment Summary</h3>
    <table style="width: 100%; border-collapse: collapse; font-size: 13px;">
      <thead>
        <tr style="background-color: #f3e8ff; color: #6b21a8; text-align: left;">
          <th style="padding: 8px 10px;">Subject</th>
          <th style="padding: 8px 10px;">Component</th>
          <th style="padding: 8px 10px; text-align: center;">Marks Obtained</th>
          <th style="padding: 8px 10px; text-align: center;">Max Marks</th>
          <th style="padding: 8px 10px; text-align: center;">%</th>
        </tr>
      </thead>
      <tbody>
        ${evaluationItems
          .map((item) => {
            const isAb = item.isAbsent;
            return `
          <tr style="border-bottom: 1px solid #f3e8ff;">
            <td style="padding: 8px 10px; font-weight: 500;">${item.subjectName} <span style="color: #6b21a8; font-family: monospace;">(${item.subjectCode})</span></td>
            <td style="padding: 8px 10px; color: #334155;">${item.componentName}</td>
            <td style="padding: 8px 10px; text-align: center; font-family: monospace;">
              ${isAb ? '<span style="color: #dc2626; font-weight: bold; background: #fee2e2; padding: 2px 6px; border-radius: 4px;">ABSENT</span>' : `<strong>${item.marksObtained}</strong>`}
            </td>
            <td style="padding: 8px 10px; text-align: center; font-family: monospace; color: #64748b;">${item.maxMarks}</td>
            <td style="padding: 8px 10px; text-align: center; font-family: monospace; font-weight: 600;">${isAb ? '0%' : (item.percentage !== null ? item.percentage + '%' : '—')}</td>
          </tr>`;
          })
          .join('')}
      </tbody>
    </table>
    ${evaluationCommentary ? `<p style="margin: 12px 0 0 0; font-size: 12px; color: #581c87; line-height: 1.5;">${evaluationCommentary.replace(/\n/g, '<br/>')}</p>` : ''}
  </div>`
      : ''
  }

  <p style="font-size: 13px; color: #475569; margin: 24px 0 16px 0;">Should you require academic guidance or assistance with your coursework, feel free to visit during office hours.</p>

  <div style="margin-top: 24px; padding-top: 16px; border-top: 1px solid #e2e8f0; font-size: 13px; color: #1e293b;">
    <p style="margin: 0; line-height: 1.4;">Regards,<br/>
    <strong>${teacher.name?.trim() || 'Faculty Member'}</strong><br/>
    <span style="color: #64748b;">${teacher.department ? `Faculty, Department of ${teacher.department}` : teacher.designation?.trim() || 'Faculty Member'}</span><br/>
    <strong style="color: #1e1b4b;">KR Mangalam University</strong></p>
  </div>
</div>`;

  return {
    recipientType: 'student',
    recipientName: studentName,
    recipientEmails: student.studentEmails,
    subject: subjectLine,
    bodyText,
    bodyHtml,
    hasAttendance,
    hasEvaluation,
    isAvailable: true,
  };
}

/**
 * Generates an individualized Parent/Guardian email.
 * Tone: Formal, informative, and collaborative.
 */
export function generateParentEmail(
  student: StudentRecipientData,
  attendanceItems: AttendanceSubjectItem[],
  evaluationItems: EvaluationAssessmentItem[],
  teacher: TeacherSenderInfo
): GeneratedEmail | null {
  const hasAttendance = attendanceItems.length > 0;
  const hasEvaluation = evaluationItems.length > 0;

  // Rule: If neither record is available, no email shall be generated
  if (!hasAttendance && !hasEvaluation) {
    return null;
  }

  const studentName = student.name || 'Student';
  const roll = student.rollNumber || 'N/A';
  const parentGreeting = student.parentName
    ? `Dear ${student.parentName} (Parent/Guardian of ${studentName}, Roll No: ${roll}),`
    : `Dear Parent / Guardian of ${studentName} (Roll No: ${roll}),`;

  const subjectLine = `Academic Progress & Attendance Report: ${studentName} (Roll No: ${roll}) | KR Mangalam University`;

  // Parent specific formal advisory commentary
  let attendanceCommentary = '';
  if (hasAttendance) {
    const lowAttSubjects = attendanceItems.filter(
      (item) => item.percentage !== null && item.percentage < 75
    );
    if (lowAttSubjects.length > 0) {
      attendanceCommentary =
        '\nArea Requiring Parental Attention:\nPlease note that KR Mangalam University enforces a statutory policy requiring a minimum of 75% attendance in each academic subject. Your ward currently holds attendance below 75% in the following subject(s): ' +
        lowAttSubjects.map((s) => `${s.subjectName} (${s.percentage}%)`).join(', ') +
        '. We kindly request your intervention in counseling your ward to ensure uninterrupted attendance in upcoming classes, thereby averting academic debarment.\n';
    } else {
      attendanceCommentary =
        '\nAttendance Status:\nYour ward is maintaining commendable attendance that satisfies all institutional eligibility standards across courses.\n';
    }
  }

  let evaluationCommentary = '';
  if (hasEvaluation) {
    const absentAssessments = evaluationItems.filter((e) => e.isAbsent);
    const lowMarkAssessments = evaluationItems.filter(
      (e) => !e.isAbsent && e.percentage !== null && e.percentage < 50
    );

    if (absentAssessments.length > 0 || lowMarkAssessments.length > 0) {
      evaluationCommentary =
        '\nAcademic Evaluation Observations:\nThe performance record indicates components where your ward was either absent or secured scores requiring academic improvement. We recommend discussing these assessments with your ward. The department offers faculty consultations and remedial guidance to support your ward in strengthening these subject competencies.\n';
    } else {
      evaluationCommentary =
        '\nPerformance Overview:\nYour ward has demonstrated satisfactory progress in the recorded evaluation assessments. We encourage continued dedication throughout the semester.\n';
    }
  }

  const sections: string[] = [];
  sections.push(parentGreeting);
  sections.push(
    `I am writing on behalf of KR Mangalam University to formally communicate the current academic progress and institutional attendance report for your ward, ${studentName} (Roll No: ${roll}).`
  );

  if (hasAttendance) {
    sections.push(formatAttendanceSectionText(attendanceItems));
    if (attendanceCommentary) sections.push(attendanceCommentary.trim());
  }

  if (hasEvaluation) {
    sections.push(formatEvaluationSectionText(evaluationItems));
    if (evaluationCommentary) sections.push(evaluationCommentary.trim());
  }

  sections.push(
    `We value your partnership in guiding your ward towards academic discipline and comprehensive excellence. For any queries or faculty consultations, please feel free to reach out to the university.`
  );
  sections.push(buildTeacherSignature(teacher));

  const bodyText = sections.join('\n\n');

  // HTML format for rich preview and dispatch
  const bodyHtml = `
<div style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; line-height: 1.6; color: #1e293b; max-width: 680px; margin: 0 auto; padding: 20px; border: 1px solid #e2e8f0; border-radius: 12px; background-color: #ffffff;">
  <div style="border-bottom: 2px solid #0284c7; padding-bottom: 12px; margin-bottom: 20px;">
    <h2 style="color: #0c4a6e; margin: 0 0 4px 0; font-size: 20px;">KR Mangalam University</h2>
    <p style="color: #64748b; margin: 0; font-size: 13px;">Official Academic Progress & Attendance Report to Parents</p>
  </div>

  <p style="font-size: 15px; font-weight: 600; color: #0f172a;">${parentGreeting}</p>
  <p style="font-size: 14px; color: #334155;">I am writing to formally present the institutional attendance and academic assessment summary for your ward, <strong>${studentName}</strong> (Roll No: <span style="font-family: monospace;">${roll}</span>):</p>

  ${
    hasAttendance
      ? `
  <div style="margin: 20px 0; padding: 16px; background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px;">
    <h3 style="margin: 0 0 12px 0; font-size: 14px; color: #0f172a; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 1px solid #cbd5e1; padding-bottom: 6px;">📅 Institutional Attendance Summary</h3>
    <table style="width: 100%; border-collapse: collapse; font-size: 13px;">
      <thead>
        <tr style="background-color: #e2e8f0; color: #475569; text-align: left;">
          <th style="padding: 8px 10px;">Course & Code</th>
          <th style="padding: 8px 10px;">Attendance Period</th>
          <th style="padding: 8px 10px; text-align: center;">Total Days</th>
          <th style="padding: 8px 10px; text-align: center;">Days Attended</th>
          <th style="padding: 8px 10px; text-align: center;">Attendance %</th>
        </tr>
      </thead>
      <tbody>
        ${attendanceItems
          .map((item) => {
            const pct = item.percentage ?? 0;
            const pctColor = pct >= 75 ? '#16a34a' : pct >= 60 ? '#d97706' : '#dc2626';
            return `
          <tr style="border-bottom: 1px solid #e2e8f0;">
            <td style="padding: 8px 10px; font-weight: 500;">${item.subjectName} <span style="color: #64748b; font-family: monospace;">(${item.subjectCode})</span></td>
            <td style="padding: 8px 10px; color: #64748b; font-size: 12px;">${item.period}</td>
            <td style="padding: 8px 10px; text-align: center; font-family: monospace;">${item.conducted ?? '—'}</td>
            <td style="padding: 8px 10px; text-align: center; font-family: monospace; font-weight: 600;">${item.attended ?? '—'}</td>
            <td style="padding: 8px 10px; text-align: center; font-family: monospace; font-weight: 700; color: ${pctColor};">${item.percentage !== null ? item.percentage + '%' : '—'}</td>
          </tr>`;
          })
          .join('')}
      </tbody>
    </table>
    ${attendanceCommentary ? `<p style="margin: 12px 0 0 0; font-size: 12px; color: #334155; line-height: 1.5;">${attendanceCommentary.replace(/\n/g, '<br/>')}</p>` : ''}
  </div>`
      : ''
  }

  ${
    hasEvaluation
      ? `
  <div style="margin: 20px 0; padding: 16px; background-color: #f0fdf4; border: 1px solid #dcfce7; border-radius: 8px;">
    <h3 style="margin: 0 0 12px 0; font-size: 14px; color: #166534; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 1px solid #bbf7d0; padding-bottom: 6px;">📝 Academic Evaluation Breakdown</h3>
    <table style="width: 100%; border-collapse: collapse; font-size: 13px;">
      <thead>
        <tr style="background-color: #dcfce7; color: #14532d; text-align: left;">
          <th style="padding: 8px 10px;">Subject</th>
          <th style="padding: 8px 10px;">Assessment Title</th>
          <th style="padding: 8px 10px; text-align: center;">Marks Secured</th>
          <th style="padding: 8px 10px; text-align: center;">Max Marks</th>
          <th style="padding: 8px 10px; text-align: center;">%</th>
        </tr>
      </thead>
      <tbody>
        ${evaluationItems
          .map((item) => {
            const isAb = item.isAbsent;
            return `
          <tr style="border-bottom: 1px solid #dcfce7;">
            <td style="padding: 8px 10px; font-weight: 500;">${item.subjectName} <span style="color: #166534; font-family: monospace;">(${item.subjectCode})</span></td>
            <td style="padding: 8px 10px; color: #334155;">${item.componentName}</td>
            <td style="padding: 8px 10px; text-align: center; font-family: monospace;">
              ${isAb ? '<span style="color: #dc2626; font-weight: bold; background: #fee2e2; padding: 2px 6px; border-radius: 4px;">ABSENT</span>' : `<strong>${item.marksObtained}</strong>`}
            </td>
            <td style="padding: 8px 10px; text-align: center; font-family: monospace; color: #64748b;">${item.maxMarks}</td>
            <td style="padding: 8px 10px; text-align: center; font-family: monospace; font-weight: 600;">${isAb ? '0%' : (item.percentage !== null ? item.percentage + '%' : '—')}</td>
          </tr>`;
          })
          .join('')}
      </tbody>
    </table>
    ${evaluationCommentary ? `<p style="margin: 12px 0 0 0; font-size: 12px; color: #166534; line-height: 1.5;">${evaluationCommentary.replace(/\n/g, '<br/>')}</p>` : ''}
  </div>`
      : ''
  }

  <p style="font-size: 13px; color: #475569; margin: 24px 0 16px 0;">We appreciate your proactive support in fostering your ward's academic discipline. Please reach out for any faculty consultations.</p>

  <div style="margin-top: 24px; padding-top: 16px; border-top: 1px solid #e2e8f0; font-size: 13px; color: #1e293b;">
    <p style="margin: 0; line-height: 1.4;">Regards,<br/>
    <strong>${teacher.name?.trim() || 'Faculty Member'}</strong><br/>
    <span style="color: #64748b;">${teacher.department ? `Faculty, Department of ${teacher.department}` : teacher.designation?.trim() || 'Faculty Member'}</span><br/>
    <strong style="color: #1e1b4b;">KR Mangalam University</strong></p>
  </div>
</div>`;

  return {
    recipientType: 'parent',
    recipientName: student.parentName || `Parent of ${studentName}`,
    recipientEmails: student.parentEmails,
    subject: subjectLine,
    bodyText,
    bodyHtml,
    hasAttendance,
    hasEvaluation,
    isAvailable: true,
  };
}
