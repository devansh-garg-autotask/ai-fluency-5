import { doc, getDoc, setDoc, getDocs, collection, query, where, orderBy, limit } from 'firebase/firestore';
import { firestoreDb } from '../../db/firebase.ts';
import { OutlookService } from './outlook.service.ts';
import type {
  DispatchJob,
  StudentDispatchRecord,
  RecipientDispatchResult,
  DispatchOverallStatus,
} from '../../../src/types/index.ts';

const DISPATCH_COLLECTION = 'communication_dispatches';

// In-memory cache for fast polling updates and resilient tracking
const inMemoryDispatches = new Map<string, DispatchJob>();

// Active background running tasks: dispatchId -> boolean
const activeRunningJobs = new Set<string>();

// Hard stop / cancellation flags: dispatchId -> boolean
const cancelledJobIds = new Set<string>();

// Helper to sanitize Firestore documents
function sanitizeForFirestore<T extends Record<string, any>>(obj: T): Record<string, any> {
  const result: Record<string, any> = {};
  for (const [key, value] of Object.entries(obj)) {
    if (value !== undefined) {
      if (value !== null && typeof value === 'object' && !Array.isArray(value) && !(value instanceof Date)) {
        result[key] = sanitizeForFirestore(value);
      } else if (Array.isArray(value)) {
        result[key] = value
          .filter((v) => v !== undefined)
          .map((item) => (typeof item === 'object' && item !== null ? sanitizeForFirestore(item) : item));
      } else {
        result[key] = value;
      }
    }
  }
  return result;
}

export class DispatchService {
  /**
   * Retrieves a dispatch job by ID from cache or Firestore.
   */
  public static async getDispatchJob(dispatchId: string): Promise<DispatchJob | null> {
    if (inMemoryDispatches.has(dispatchId)) {
      return inMemoryDispatches.get(dispatchId)!;
    }

    try {
      const snap = await getDoc(doc(firestoreDb, DISPATCH_COLLECTION, dispatchId));
      if (snap.exists()) {
        const job = snap.data() as DispatchJob;
        inMemoryDispatches.set(dispatchId, job);
        return job;
      }
    } catch (err) {
      console.warn(`[DispatchService] Firestore read failed for ${dispatchId}:`, err);
    }

    return null;
  }

  /**
   * Retrieves the latest dispatch job for a specific teacher.
   */
  public static async getLatestDispatch(teacherId: string): Promise<DispatchJob | null> {
    // Check in-memory first for latest matching teacherId
    let latest: DispatchJob | null = null;
    for (const job of inMemoryDispatches.values()) {
      if (job.teacherId === teacherId) {
        if (!latest || new Date(job.createdAt).getTime() > new Date(latest.createdAt).getTime()) {
          latest = job;
        }
      }
    }
    if (latest) return latest;

    let dbLatest: DispatchJob | null = null;
    try {
      const snap = await getDocs(collection(firestoreDb, DISPATCH_COLLECTION));
      snap.forEach((d) => {
        const data = d.data() as DispatchJob;
        if (data.teacherId === teacherId) {
          if (!dbLatest || new Date(data.createdAt).getTime() > new Date(dbLatest.createdAt).getTime()) {
            dbLatest = data;
          }
        }
      });
      if (dbLatest) {
        inMemoryDispatches.set((dbLatest as DispatchJob).id, dbLatest);
      }
      return dbLatest;
    } catch (err) {
      console.warn(`[DispatchService] Error finding latest dispatch for ${teacherId}:`, err);
    }

    return null;
  }

  /**
   * Persists a dispatch job to Firestore and memory.
   */
  public static async saveDispatchJob(job: DispatchJob): Promise<void> {
    job.updatedAt = new Date().toISOString();
    inMemoryDispatches.set(job.id, job);

    try {
      const cleanData = sanitizeForFirestore(job);
      await setDoc(doc(firestoreDb, DISPATCH_COLLECTION, job.id), cleanData);
    } catch (err) {
      console.error(`[DispatchService] Failed to persist job ${job.id} to Firestore:`, err);
    }
  }

  /**
   * Initializes a new bulk dispatch job.
   */
  public static async createDispatchJob(params: {
    teacherId: string;
    teacherName?: string;
    recipientOptions: {
      sendToStudent: boolean;
      sendToParent: boolean;
    };
    students: {
      rollNumber: string;
      name: string;
      hasAcademicRecord: boolean;
      studentEmailIds: string[];
      parentEmailIds: string[];
      studentTemplate?: string;
      parentTemplate?: string;
    }[];
  }): Promise<DispatchJob> {
    const dispatchId = `disp_${Date.now()}_${Math.random().toString(36).substring(2, 8)}`;

    const records: StudentDispatchRecord[] = params.students.map((st) => {
      const studentEmails: RecipientDispatchResult[] = [];
      const parentEmails: RecipientDispatchResult[] = [];

      if (params.recipientOptions.sendToStudent) {
        const emails = Array.isArray(st.studentEmailIds) ? st.studentEmailIds : [];
        emails.forEach((em) => {
          const clean = (em || '').trim();
          if (clean) {
            studentEmails.push({
              email: clean,
              recipientType: 'student',
              status: 'Pending',
            });
          }
        });
      }

      if (params.recipientOptions.sendToParent) {
        const emails = Array.isArray(st.parentEmailIds) ? st.parentEmailIds : [];
        emails.forEach((em) => {
          const clean = (em || '').trim();
          if (clean) {
            parentEmails.push({
              email: clean,
              recipientType: 'parent',
              status: 'Pending',
            });
          }
        });
      }

      // Determine initial status based on exact specification
      let overallStatus: DispatchOverallStatus = 'Failed';
      if (!st.hasAcademicRecord) {
        overallStatus = 'Record Missing in Persistent Database';
      } else if (params.recipientOptions.sendToStudent && studentEmails.length === 0) {
        overallStatus = 'Student Email Missing';
      } else if (params.recipientOptions.sendToParent && parentEmails.length === 0) {
        overallStatus = 'Parent Email Missing';
      } else {
        overallStatus = 'Failed'; // Default before sending
      }

      return {
        rollNumber: st.rollNumber,
        name: st.name,
        hasAcademicRecord: st.hasAcademicRecord,
        studentEmails,
        parentEmails,
        studentTemplate: st.studentTemplate,
        parentTemplate: st.parentTemplate,
        overallStatus,
      };
    });

    // Count total pending emails
    let totalEmails = 0;
    records.forEach((rec) => {
      if (rec.hasAcademicRecord) {
        totalEmails += rec.studentEmails.length + rec.parentEmails.length;
      }
    });

    const job: DispatchJob = {
      id: dispatchId,
      teacherId: params.teacherId,
      teacherName: params.teacherName || 'Faculty',
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      status: 'in_progress',
      recipientOptions: params.recipientOptions,
      totalEmails,
      sentEmails: 0,
      failedEmails: 0,
      remainingEmails: totalEmails,
      currentProcessingStatus: 'Initiating Microsoft Outlook email dispatch pipeline...',
      records,
    };

    await this.saveDispatchJob(job);
    return job;
  }

  /**
   * Recalculates student overall status based on recipient email outcomes.
   */
  public static computeStudentOverallStatus(
    rec: StudentDispatchRecord,
    options: { sendToStudent: boolean; sendToParent: boolean }
  ): DispatchOverallStatus {
    if (!rec.hasAcademicRecord) {
      return 'Record Missing in Persistent Database';
    }

    if (options.sendToStudent && rec.studentEmails.length === 0) {
      return 'Student Email Missing';
    }

    if (options.sendToParent && rec.parentEmails.length === 0) {
      return 'Parent Email Missing';
    }

    const allApplicable: RecipientDispatchResult[] = [];
    if (options.sendToStudent) allApplicable.push(...rec.studentEmails);
    if (options.sendToParent) allApplicable.push(...rec.parentEmails);

    if (allApplicable.length === 0) {
      return 'Failed';
    }

    const allSuccess = allApplicable.every((item) => item.status === 'Success');
    if (allSuccess) return 'Success';

    const allFailed = allApplicable.every((item) => item.status === 'Failed');
    if (allFailed) return 'Failed';

    const hasAnySuccess = allApplicable.some((item) => item.status === 'Success');
    if (hasAnySuccess) {
      return 'Partial Success';
    }

    return 'Failed';
  }

  /**
   * Immediately stops and interrupts an active dispatch job (Hard Stop).
   */
  public static async stopDispatch(dispatchId: string): Promise<{ success: boolean; job: DispatchJob | null }> {
    cancelledJobIds.add(dispatchId);
    activeRunningJobs.delete(dispatchId);

    const job = await this.getDispatchJob(dispatchId);
    if (job) {
      job.status = 'interrupted';
      job.currentProcessingStatus = `Dispatch was manually interrupted and halted by faculty. Sent: ${job.sentEmails}, Failed: ${job.failedEmails}, Remaining: ${job.remainingEmails}.`;
      job.updatedAt = new Date().toISOString();
      await this.saveDispatchJob(job);
    }
    return { success: true, job };
  }

  /**
   * Runs the background dispatch processing loop.
   * Sends emails sequentially through Microsoft Outlook via Microsoft Graph API.
   * If retryFailedOnly is true, only retries recipients whose status is 'Failed'.
   */
  public static async executeDispatch(
    dispatchId: string,
    teacherId: string,
    retryFailedOnly = false
  ): Promise<void> {
    if (activeRunningJobs.has(dispatchId)) {
      console.log(`[DispatchService] Dispatch ${dispatchId} is already running.`);
      return;
    }

    activeRunningJobs.add(dispatchId);
    cancelledJobIds.delete(dispatchId);

    try {
      const job = await this.getDispatchJob(dispatchId);
      if (!job) {
        throw new Error(`Dispatch job ${dispatchId} not found`);
      }

      job.status = 'in_progress';
      job.currentProcessingStatus = 'Authenticating Microsoft Outlook connection...';
      await this.saveDispatchJob(job);

      // Verify teacher Outlook access token
      const tokenResult = await OutlookService.getValidAccessToken(teacherId);
      if (!tokenResult.token) {
        job.status = 'failed';
        job.currentProcessingStatus =
          tokenResult.error ||
          'Please configure your Microsoft Outlook email account with your application account before initiating bulk email communication.';
        await this.saveDispatchJob(job);
        activeRunningJobs.delete(dispatchId);
        return;
      }

      // Calculate queue
      const queue: {
        recordIndex: number;
        recipientType: 'student' | 'parent';
        emailIndex: number;
        email: string;
        subject: string;
        bodyHtml: string;
      }[] = [];

      job.records.forEach((rec, recIdx) => {
        if (!rec.hasAcademicRecord) return;

        if (job.recipientOptions.sendToStudent) {
          rec.studentEmails.forEach((se, seIdx) => {
            if (retryFailedOnly) {
              if (se.status === 'Failed') {
                queue.push({
                  recordIndex: recIdx,
                  recipientType: 'student',
                  emailIndex: seIdx,
                  email: se.email,
                  subject: `Academic Performance & Attendance Report - ${rec.name} (${rec.rollNumber})`,
                  bodyHtml: rec.studentTemplate || '',
                });
              }
            } else {
              if (se.status !== 'Success') {
                queue.push({
                  recordIndex: recIdx,
                  recipientType: 'student',
                  emailIndex: seIdx,
                  email: se.email,
                  subject: `Academic Performance & Attendance Report - ${rec.name} (${rec.rollNumber})`,
                  bodyHtml: rec.studentTemplate || '',
                });
              }
            }
          });
        }

        if (job.recipientOptions.sendToParent) {
          rec.parentEmails.forEach((pe, peIdx) => {
            if (retryFailedOnly) {
              if (pe.status === 'Failed') {
                queue.push({
                  recordIndex: recIdx,
                  recipientType: 'parent',
                  emailIndex: peIdx,
                  email: pe.email,
                  subject: `Student Academic Performance & Attendance Notification: ${rec.name} (${rec.rollNumber})`,
                  bodyHtml: rec.parentTemplate || '',
                });
              }
            } else {
              if (pe.status !== 'Success') {
                queue.push({
                  recordIndex: recIdx,
                  recipientType: 'parent',
                  emailIndex: peIdx,
                  email: pe.email,
                  subject: `Student Academic Performance & Attendance Notification: ${rec.name} (${rec.rollNumber})`,
                  bodyHtml: rec.parentTemplate || '',
                });
              }
            }
          });
        }
      });

      if (retryFailedOnly) {
        job.totalEmails = queue.length;
        job.remainingEmails = queue.length;
        job.failedEmails = 0;
      }

      // Sequential dispatch loop with throttling and hard stop checks
      for (let i = 0; i < queue.length; i++) {
        // Immediate hard stop / interrupt check
        if (cancelledJobIds.has(dispatchId)) {
          console.log(`[DispatchService] Hard stop triggered on dispatch ${dispatchId} at email ${i + 1}/${queue.length}.`);
          job.status = 'interrupted';
          job.currentProcessingStatus = `Dispatch was manually interrupted. Sent: ${job.sentEmails}, Failed: ${job.failedEmails}, Remaining: ${job.remainingEmails}.`;
          job.updatedAt = new Date().toISOString();
          await this.saveDispatchJob(job);
          cancelledJobIds.delete(dispatchId);
          activeRunningJobs.delete(dispatchId);
          return;
        }

        const item = queue[i];
        const student = job.records[item.recordIndex];

        job.currentProcessingStatus = `Sending email (${i + 1}/${queue.length}) to ${item.email} for student ${student.name}...`;
        job.updatedAt = new Date().toISOString();
        await this.saveDispatchJob(job);

        // Send via Microsoft Graph API
        const sendResult = await OutlookService.sendMail(teacherId, {
          to: item.email,
          subject: item.subject,
          bodyHtml: item.bodyHtml,
        });

        const targetList =
          item.recipientType === 'student'
            ? student.studentEmails
            : student.parentEmails;

        if (sendResult.success) {
          targetList[item.emailIndex].status = 'Success';
          targetList[item.emailIndex].sentAt = new Date().toISOString();
          delete targetList[item.emailIndex].reason;
          job.sentEmails++;
        } else {
          targetList[item.emailIndex].status = 'Failed';
          targetList[item.emailIndex].reason = sendResult.error || 'Recipient address rejected';
          job.failedEmails++;
        }

        if (job.remainingEmails > 0) {
          job.remainingEmails--;
        }

        // Recompute student overall status
        student.overallStatus = this.computeStudentOverallStatus(
          student,
          job.recipientOptions
        );

        // Save progress after each email
        await this.saveDispatchJob(job);

        // Friendly rate limiting delay with sub-interval interrupt checks
        if (i < queue.length - 1) {
          for (let step = 0; step < 5; step++) {
            if (cancelledJobIds.has(dispatchId)) break;
            await new Promise((resolve) => setTimeout(resolve, 50));
          }
        }
      }

      job.status = 'completed';
      job.currentProcessingStatus = `Bulk email communication completed. Sent: ${job.sentEmails}, Failed: ${job.failedEmails}.`;
      job.remainingEmails = 0;
      await this.saveDispatchJob(job);
    } catch (err: any) {
      console.error(`[DispatchService] Unhandled error during dispatch ${dispatchId}:`, err);
      const job = inMemoryDispatches.get(dispatchId);
      if (job) {
        job.status = 'failed';
        job.currentProcessingStatus = `Dispatch encountered an error: ${err.message}`;
        await this.saveDispatchJob(job);
      }
    } finally {
      activeRunningJobs.delete(dispatchId);
      cancelledJobIds.delete(dispatchId);
    }
  }
}
