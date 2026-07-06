// Client-side API helpers with JWT-bearer auth.
// Token is stored in localStorage under 'wbr_token'.

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

// Safe ArrayBuffer → base64 that doesn't blow the call stack on large files.
function arrayBufferToBase64(buffer) {
  const bytes = new Uint8Array(buffer);
  let binary = '';
  const chunk = 8192;
  for (let i = 0; i < bytes.length; i += chunk) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunk));
  }
  return btoa(binary);
}

// Commit a file directly to GitHub from the browser using the Git Data API.
// This bypasses Vercel's 4.5 MB function payload limit entirely.
async function githubDirectUpload(weekName, fileType, file) {
  // 1. Get GitHub credentials from the server (admin-only endpoint).
  const credsRes = await fetch('/api/admin/github-upload-creds', {
    headers: { Authorization: `Bearer ${token()}` },
  });
  if (!credsRes.ok) {
    const e = await credsRes.json().catch(() => ({}));
    throw new Error(e.error || 'Could not get upload credentials');
  }
  const { ghToken, owner, repo, branch } = await credsRes.json();

  const gh = (path, opts = {}) => fetch(`https://api.github.com${path}`, {
    ...opts,
    headers: {
      Authorization: `Bearer ${ghToken}`,
      Accept: 'application/vnd.github+json',
      'X-GitHub-Api-Version': '2022-11-28',
      'Content-Type': 'application/json',
      ...(opts.headers || {}),
    },
  });

  const filePath = `data/${weekName}/${fileType}.xlsx`;

  // 2. Base64-encode the file content.
  const base64Content = arrayBufferToBase64(await file.arrayBuffer());

  // 3. Create a Git blob with the file content.
  const blobRes = await gh(`/repos/${owner}/${repo}/git/blobs`, {
    method: 'POST',
    body: JSON.stringify({ content: base64Content, encoding: 'base64' }),
  });
  if (!blobRes.ok) throw new Error(`GitHub blob failed: ${blobRes.status} ${await blobRes.text()}`);
  const { sha: blobSha } = await blobRes.json();

  // 4. Get the current HEAD commit SHA.
  const refRes = await gh(`/repos/${owner}/${repo}/git/ref/heads/${branch}`);
  if (!refRes.ok) throw new Error(`GitHub ref failed: ${refRes.status}`);
  const { object: { sha: headSha } } = await refRes.json();

  // 5. Get the current tree SHA from the HEAD commit.
  const commitRes = await gh(`/repos/${owner}/${repo}/git/commits/${headSha}`);
  if (!commitRes.ok) throw new Error(`GitHub commit fetch failed: ${commitRes.status}`);
  const { tree: { sha: treeSha } } = await commitRes.json();

  // 6. Create a new tree that adds/replaces this file.
  const newTreeRes = await gh(`/repos/${owner}/${repo}/git/trees`, {
    method: 'POST',
    body: JSON.stringify({
      base_tree: treeSha,
      tree: [{ path: filePath, mode: '100644', type: 'blob', sha: blobSha }],
    }),
  });
  if (!newTreeRes.ok) throw new Error(`GitHub tree failed: ${newTreeRes.status}`);
  const { sha: newTreeSha } = await newTreeRes.json();

  // 7. Create a commit.
  const newCommitRes = await gh(`/repos/${owner}/${repo}/git/commits`, {
    method: 'POST',
    body: JSON.stringify({
      message: `Admin: upload ${fileType} for ${weekName}`,
      tree: newTreeSha,
      parents: [headSha],
    }),
  });
  if (!newCommitRes.ok) throw new Error(`GitHub commit failed: ${newCommitRes.status}`);
  const { sha: newCommitSha } = await newCommitRes.json();

  // 8. Update the branch to point to the new commit.
  const updateRes = await gh(`/repos/${owner}/${repo}/git/refs/heads/${branch}`, {
    method: 'PATCH',
    body: JSON.stringify({ sha: newCommitSha }),
  });
  if (!updateRes.ok) throw new Error(`GitHub ref update failed: ${updateRes.status}`);

  return { ok: true, weekName, fileType };
}

// Threshold above which we bypass Vercel and upload directly to GitHub from the browser.
const LARGE_FILE_THRESHOLD = 3.5 * 1024 * 1024; // 3.5 MB

export async function adminUploadData(weekName, fileType, file) {
  if (file.size > LARGE_FILE_THRESHOLD) {
    return githubDirectUpload(weekName, fileType, file);
  }

  // Small file: send through the Vercel function as before.
  const form = new FormData();
  form.append('weekName', weekName);
  form.append('fileType', fileType);
  form.append('file', file);
  const res = await fetch('/api/admin/upload/data', {
    method: 'POST',
    headers: { Authorization: `Bearer ${token()}` },
    body: form,
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
  if (file.size > LARGE_FILE_THRESHOLD) {
    return githubDirectUploadScorecard(granularity, file);
  }
  const form = new FormData();
  form.append('granularity', granularity);
  form.append('file', file);
  const res = await fetch('/api/admin/upload/scorecard', {
    method: 'POST',
    headers: { Authorization: `Bearer ${token()}` },
    body: form,
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

async function githubDirectUploadScorecard(granularity, file) {
  const credsRes = await fetch('/api/admin/github-upload-creds', {
    headers: { Authorization: `Bearer ${token()}` },
  });
  if (!credsRes.ok) {
    const e = await credsRes.json().catch(() => ({}));
    throw new Error(e.error || 'Could not get upload credentials');
  }
  const { ghToken, owner, repo, branch } = await credsRes.json();

  const gh = (path, opts = {}) => fetch(`https://api.github.com${path}`, {
    ...opts,
    headers: {
      Authorization: `Bearer ${ghToken}`,
      Accept: 'application/vnd.github+json',
      'X-GitHub-Api-Version': '2022-11-28',
      'Content-Type': 'application/json',
      ...(opts.headers || {}),
    },
  });

  const filePath = `scorecard/${granularity}/${file.name}`;
  const base64Content = arrayBufferToBase64(await file.arrayBuffer());

  const blobRes = await gh(`/repos/${owner}/${repo}/git/blobs`, {
    method: 'POST',
    body: JSON.stringify({ content: base64Content, encoding: 'base64' }),
  });
  if (!blobRes.ok) throw new Error(`GitHub blob failed: ${blobRes.status}`);
  const { sha: blobSha } = await blobRes.json();

  const refRes = await gh(`/repos/${owner}/${repo}/git/ref/heads/${branch}`);
  if (!refRes.ok) throw new Error(`GitHub ref failed: ${refRes.status}`);
  const { object: { sha: headSha } } = await refRes.json();

  const commitRes = await gh(`/repos/${owner}/${repo}/git/commits/${headSha}`);
  if (!commitRes.ok) throw new Error(`GitHub commit fetch failed: ${commitRes.status}`);
  const { tree: { sha: treeSha } } = await commitRes.json();

  const newTreeRes = await gh(`/repos/${owner}/${repo}/git/trees`, {
    method: 'POST',
    body: JSON.stringify({
      base_tree: treeSha,
      tree: [{ path: filePath, mode: '100644', type: 'blob', sha: blobSha }],
    }),
  });
  if (!newTreeRes.ok) throw new Error(`GitHub tree failed: ${newTreeRes.status}`);
  const { sha: newTreeSha } = await newTreeRes.json();

  const newCommitRes = await gh(`/repos/${owner}/${repo}/git/commits`, {
    method: 'POST',
    body: JSON.stringify({
      message: `Admin: upload scorecard ${granularity}/${file.name}`,
      tree: newTreeSha,
      parents: [headSha],
    }),
  });
  if (!newCommitRes.ok) throw new Error(`GitHub commit failed: ${newCommitRes.status}`);
  const { sha: newCommitSha } = await newCommitRes.json();

  const updateRes = await gh(`/repos/${owner}/${repo}/git/refs/heads/${branch}`, {
    method: 'PATCH',
    body: JSON.stringify({ sha: newCommitSha }),
  });
  if (!updateRes.ok) throw new Error(`GitHub ref update failed: ${updateRes.status}`);

  return { ok: true, granularity, filename: file.name };
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

export async function fetchScorecardIndex() {
  const res = await fetch('/api/scorecard', { headers: authHeaders() });
  handle401(res);
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.error || 'Failed to load scorecard list');
  }
  return await res.json();
}

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
