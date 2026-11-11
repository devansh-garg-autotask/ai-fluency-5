export type OutlookConnectionStatus =
  | 'CONNECTED'
  | 'DISCONNECTED'
  | 'REAUTHENTICATION_REQUIRED'
  | 'ERROR'
  | 'CHECKING'
  | 'VERIFICATION_UNAVAILABLE';

export interface OutlookStatusResponse {
  connected: boolean;
  status: OutlookConnectionStatus;
  email?: string;
  displayName?: string;
  accountType?: 'work_school' | 'personal' | 'unknown';
  scopes?: string[];
  lastVerifiedAt?: string;
  connectedAt?: string;
  errorMessage?: string;
  isConfigured: boolean;
}

export async function getOutlookConfig(): Promise<{
  isConfigured: boolean;
  redirectUri: string;
  clientIdMasked?: string;
  isPublicClient?: boolean;
}> {
  const res = await fetch('/api/outlook/config');
  if (!res.ok) throw new Error('Failed to fetch Outlook configuration');
  return res.json();
}

export async function saveOutlookCredentials(
  clientId: string,
  clientSecret?: string
): Promise<{ success: boolean; isConfigured: boolean; isPublicClient?: boolean; message: string }> {
  const res = await fetch('/api/outlook/credentials', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ clientId, clientSecret: clientSecret || '' }),
  });
  if (!res.ok) {
    const errorData = await res.json().catch(() => ({}));
    throw new Error(errorData.error || 'Failed to save credentials');
  }
  return res.json();
}

export async function initiateOutlookConnect(teacherId: string): Promise<{
  authUrl: string;
  state: string;
  isConfigured: boolean;
  redirectUri: string;
}> {
  const res = await fetch('/api/outlook/connect', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-teacher-id': teacherId,
    },
    body: JSON.stringify({ teacherId }),
  });
  if (!res.ok) {
    const errorData = await res.json().catch(() => ({}));
    throw new Error(errorData.error || 'Failed to initiate Outlook connection');
  }
  return res.json();
}

export async function connectDirectOutlook(
  teacherId: string,
  email?: string
): Promise<OutlookStatusResponse> {
  const res = await fetch('/api/outlook/connect-direct', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-teacher-id': teacherId,
    },
    body: JSON.stringify({ teacherId, email }),
  });
  if (!res.ok) {
    const errorData = await res.json().catch(() => ({}));
    throw new Error(errorData.error || 'Failed to connect Outlook account');
  }
  return res.json();
}

export async function getOutlookStatus(teacherId: string): Promise<OutlookStatusResponse> {
  const res = await fetch(`/api/outlook/status?teacherId=${encodeURIComponent(teacherId)}`, {
    headers: {
      'x-teacher-id': teacherId,
    },
  });
  if (!res.ok) {
    const errorData = await res.json().catch(() => ({}));
    throw new Error(errorData.error || 'Failed to fetch Outlook connection status');
  }
  return res.json();
}

export async function verifyOutlookConnection(teacherId: string): Promise<OutlookStatusResponse> {
  const res = await fetch('/api/outlook/verify', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-teacher-id': teacherId,
    },
    body: JSON.stringify({ teacherId }),
  });
  if (!res.ok) {
    const errorData = await res.json().catch(() => ({}));
    throw new Error(errorData.error || 'Live verification request failed');
  }
  return res.json();
}

export async function disconnectOutlook(teacherId: string): Promise<void> {
  const res = await fetch('/api/outlook/disconnect', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-teacher-id': teacherId,
    },
    body: JSON.stringify({ teacherId }),
  });
  if (!res.ok) {
    const errorData = await res.json().catch(() => ({}));
    throw new Error(errorData.error || 'Failed to disconnect Outlook');
  }
}

export async function exchangeOutlookCode(
  teacherId: string,
  codeOrUrl: string
): Promise<{ success: boolean; teacherId: string; email?: string; displayName?: string }> {
  const res = await fetch('/api/outlook/exchange-code', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-teacher-id': teacherId,
    },
    body: JSON.stringify({ teacherId, codeOrUrl }),
  });
  if (!res.ok) {
    const errorData = await res.json().catch(() => ({}));
    throw new Error(errorData.error || 'Failed to exchange authorization code');
  }
  return res.json();
}

export async function startDeviceCode(teacherId: string): Promise<{
  success: boolean;
  userCode: string;
  deviceCode: string;
  verificationUri: string;
  expiresIn: number;
  interval: number;
  message?: string;
}> {
  const res = await fetch('/api/outlook/device-code/start', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-teacher-id': teacherId,
    },
    body: JSON.stringify({ teacherId }),
  });
  if (!res.ok) {
    const errorData = await res.json().catch(() => ({}));
    throw new Error(errorData.error || 'Failed to start device login');
  }
  return res.json();
}

export async function pollDeviceCode(
  teacherId: string,
  deviceCode: string
): Promise<{
  status: 'PENDING' | 'SUCCESS' | 'EXPIRED' | 'DECLINED' | 'ERROR';
  email?: string;
  displayName?: string;
  error?: string;
}> {
  const res = await fetch('/api/outlook/device-code/poll', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-teacher-id': teacherId,
    },
    body: JSON.stringify({ teacherId, deviceCode }),
  });
  if (!res.ok) {
    const errorData = await res.json().catch(() => ({}));
    throw new Error(errorData.error || 'Failed to poll device status');
  }
  return res.json();
}

export async function logoutDisconnectOutlook(teacherId: string): Promise<void> {
  try {
    await fetch('/api/outlook/logout-disconnect', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-teacher-id': teacherId,
      },
      body: JSON.stringify({ teacherId }),
    });
  } catch (e) {
    // Silent on background logout
  }
}

export async function initiateBulkDispatch(payload: {
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
}): Promise<{ success: boolean; dispatchId: string; job: any }> {
  const res = await fetch('/api/outlook/bulk-dispatch', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-teacher-id': payload.teacherId,
    },
    body: JSON.stringify(payload),
  });

  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    const error: any = new Error(data.error || 'Failed to initiate bulk email dispatch');
    error.requiresConfiguration = data.requiresConfiguration;
    throw error;
  }
  return data;
}

export async function getBulkDispatchProgress(
  dispatchId: string
): Promise<{ success: boolean; job: any }> {
  const res = await fetch(`/api/outlook/bulk-dispatch/progress/${encodeURIComponent(dispatchId)}`);
  if (!res.ok) {
    const errorData = await res.json().catch(() => ({}));
    throw new Error(errorData.error || 'Failed to fetch dispatch progress');
  }
  return res.json();
}

export async function getLatestBulkDispatch(
  teacherId: string
): Promise<{ success: boolean; job: any }> {
  const res = await fetch(`/api/outlook/bulk-dispatch/latest?teacherId=${encodeURIComponent(teacherId)}`, {
    headers: {
      'x-teacher-id': teacherId,
    },
  });
  if (!res.ok) {
    const errorData = await res.json().catch(() => ({}));
    throw new Error(errorData.error || 'Failed to fetch latest dispatch');
  }
  return res.json();
}

export async function retryBulkDispatch(payload: {
  dispatchId: string;
  teacherId: string;
}): Promise<{ success: boolean; message: string; job: any }> {
  const res = await fetch('/api/outlook/bulk-dispatch/retry', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-teacher-id': payload.teacherId,
    },
    body: JSON.stringify(payload),
  });

  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    const error: any = new Error(data.error || 'Failed to retry dispatch');
    error.requiresConfiguration = data.requiresConfiguration;
    throw error;
  }
  return data;
}

export async function stopBulkDispatch(payload: {
  dispatchId: string;
  teacherId: string;
}): Promise<{ success: boolean; message: string; job: any }> {
  const res = await fetch('/api/outlook/bulk-dispatch/stop', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-teacher-id': payload.teacherId,
    },
    body: JSON.stringify(payload),
  });

  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    throw new Error(data.error || 'Failed to stop dispatch');
  }
  return data;
}


