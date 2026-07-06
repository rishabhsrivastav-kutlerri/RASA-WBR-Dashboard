// Client-side API helpers with JWT-bearer auth.
// Token is stored in localStorage under 'wbr_token'.
import { upload } from '@vercel/blob/client';

const TOKEN_KEY = 'wbr_token';
const USER_KEY  = 'wbr_user';

function token() {
  if (typeof window === 'undefined') return '';
  return localStorage.getItem(TOKEN_KEY) || '';
}

function authHeaders() {
  return { 'Content-Type': 'application/json', Authorization: `Bearer ${token()}` };
}

export function getStoredUser() {
  if (typeof window === 'undefined') return null;
  try { return JSON.parse(localStorage.getItem(USER_KEY) || 'null'); } catch { return null; }
}

function handle401(res) {
  if (res.status === 401 && typeof window !== 'undefined') {
    localStorage.removeItem(TOKEN_KEY);
    window.location.href = '/login';
  }
}

export async function login(email, password) {
  const res = await fetch('/api/auth/login', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ email, password }),
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data.error || 'Login failed');
  // Persist user info (including role) so the dashboard can read it without decoding JWT.
  if (data.user && typeof window !== 'undefined') {
    localStorage.setItem(USER_KEY, JSON.stringify(data.user));
  }
  return data;
}

// ── Admin API ────────────────────────────────────────────────────────────────

export async function fetchAdminStatus() {
  const res = await fetch('/api/admin/status', { headers: authHeaders() });
  handle401(res);
  if (!res.ok) { const e = await res.json().catch(() => ({})); throw new Error(e.error || 'Failed'); }
  return res.json();
}

export async function adminUploadData(weekName, fileType, file) {
  // Upload directly from browser to Vercel Blob (bypasses Vercel's 4.5MB function limit),
  // then send the blob URL to the server which pulls it and commits to GitHub.
  const blob = await upload(`wbr-upload/${weekName}/${fileType}.xlsx`, file, {
    access: 'public',
    handleUploadUrl: '/api/admin/upload/blob-auth',
    clientPayload: token(),
  });

  const res = await fetch('/api/admin/upload/data', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${token()}` },
    body: JSON.stringify({ weekName, fileType, blobUrl: blob.url }),
  });
  handle401(res);
  if (!res.ok) {
    const body = await res.text().catch(() => '');
    let msg = `HTTP ${res.status}`;
    try { const e = JSON.parse(body); msg = e.error || msg; } catch { if (body) msg += ': ' + body.slice(0, 200); }
    throw new Error(msg);
  }
  return res.json();
}

// Downloads a file through the admin API (auth header required) and triggers
// a browser save via a temporary object URL.
async function downloadViaAdminApi(url) {
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token()}` } });
  handle401(res);
  if (!res.ok) { const e = await res.json().catch(() => ({})); throw new Error(e.error || 'Download failed'); }
  const disposition = res.headers.get('Content-Disposition') || '';
  const match = disposition.match(/filename="([^"]+)"/);
  const filename = match ? match[1] : 'download.xlsx';
  const blob = await res.blob();
  const objectUrl = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = objectUrl;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(objectUrl);
}

export function adminDownloadData(weekName, fileType) {
  const qs = `weekName=${encodeURIComponent(weekName)}&fileType=${encodeURIComponent(fileType)}`;
  return downloadViaAdminApi('/api/admin/file/data?' + qs);
}

export function adminDownloadScorecard(granularity, filename) {
  const qs = `granularity=${encodeURIComponent(granularity)}&filename=${encodeURIComponent(filename)}`;
  return downloadViaAdminApi('/api/admin/file/scorecard?' + qs);
}

export async function adminUploadScorecard(granularity, file) {
  const form = new FormData();
  form.append('granularity', granularity);
  form.append('file', file);
  const res = await fetch('/api/admin/upload/scorecard', {
    method: 'POST',
    headers: { Authorization: `Bearer ${token()}` },
    body: form,
  });
  handle401(res);
  if (!res.ok) { const e = await res.json().catch(() => ({})); throw new Error(e.error || 'Upload failed'); }
  return res.json();
}

export async function adminDeleteData(weekName, fileType) {
  const res = await fetch('/api/admin/delete/data', {
    method: 'DELETE',
    headers: authHeaders(),
    body: JSON.stringify({ weekName, fileType }),
  });
  handle401(res);
  if (!res.ok) { const e = await res.json().catch(() => ({})); throw new Error(e.error || 'Delete failed'); }
  return res.json();
}

export async function adminDeleteScorecard(granularity, filename) {
  const res = await fetch('/api/admin/delete/scorecard', {
    method: 'DELETE',
    headers: authHeaders(),
    body: JSON.stringify({ granularity, filename }),
  });
  handle401(res);
  if (!res.ok) { const e = await res.json().catch(() => ({})); throw new Error(e.error || 'Delete failed'); }
  return res.json();
}

export async function fetchSheets() {
  const res = await fetch('/api/sheets', { headers: authHeaders() });
  handle401(res);
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.error || 'Failed to load sheet list');
  }
  const data = await res.json();
  return data.sheets || [];
}

export async function fetchWeekData(week) {
  const res = await fetch('/api/data/' + encodeURIComponent(week), { headers: authHeaders() });
  handle401(res);
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.error || 'Failed to load week data');
  }
  return await res.json();
}

// Leadership Scorecard: index of available weekly/period/quarter selections.
export async function fetchScorecardIndex() {
  const res = await fetch('/api/scorecard', { headers: authHeaders() });
  handle401(res);
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.error || 'Failed to load scorecard list');
  }
  return await res.json();
}

// Parsed dashboard + scoring matrix for one selection.
export async function fetchScorecard(granularity, item) {
  const qs = `granularity=${encodeURIComponent(granularity)}&item=${encodeURIComponent(item)}`;
  const res = await fetch('/api/scorecard?' + qs, { headers: authHeaders() });
  handle401(res);
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.error || 'Failed to load scorecard');
  }
  return await res.json();
}
