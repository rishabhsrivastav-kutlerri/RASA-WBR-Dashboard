// GitHub Contents API wrapper for admin uploads/deletes.
//
// Every write here is a real commit pushed to GITHUB_BRANCH — Vercel's
// GitHub integration then rebuilds and redeploys automatically, which is
// how uploads end up shipped to production (no Vercel Blob involved).

const OWNER  = process.env.GITHUB_OWNER;
const REPO   = process.env.GITHUB_REPO;
const BRANCH = process.env.GITHUB_BRANCH || 'main';
const TOKEN  = process.env.GITHUB_TOKEN;

const API = 'https://api.github.com';

function encodeGitPath(p) {
  return p.split('/').map(encodeURIComponent).join('/');
}

async function gh(path, options = {}) {
  const url = `${API}${path}`;
  const headers = {
    Authorization: `Bearer ${TOKEN}`,
    Accept: 'application/vnd.github+json',
    'X-GitHub-Api-Version': '2022-11-28',
    ...(options.headers || {}),
  };
  // Retry transient GitHub 5xx errors (502 Unicorn, 503, 504) with backoff.
  const RETRYABLE = new Set([502, 503, 504]);
  let delay = 500;
  for (let attempt = 0; attempt < 3; attempt++) {
    const res = await fetch(url, { ...options, headers });
    if (!RETRYABLE.has(res.status) || attempt === 2) return res;
    await new Promise(r => setTimeout(r, delay));
    delay *= 2;
  }
}

async function ghErrText(res) {
  const ct = res.headers.get('content-type') || '';
  if (ct.includes('text/html')) return res.status.toString();
  const t = await res.text().catch(() => '');
  return `${res.status} ${t}`.trim();
}

// Raw GET of a file or directory listing. Returns null if the path doesn't exist.
async function getContents(repoPath) {
  const res = await gh(`/repos/${OWNER}/${REPO}/contents/${encodeGitPath(repoPath)}?ref=${BRANCH}`);
  if (res.status === 404) return null;
  if (!res.ok) throw new Error(`GitHub GET ${repoPath} failed: ${await ghErrText(res)}`);
  return res.json();
}

// List entries in a directory. Returns [] if the directory doesn't exist.
async function listDir(repoPath) {
  const data = await getContents(repoPath);
  if (!data) return [];
  return Array.isArray(data) ? data : [data];
}

// Download a file's content as a Buffer. Returns null if it doesn't exist.
// For files > 1 MB the GitHub Contents API returns empty content with a download_url instead.
async function getFileBuffer(repoPath) {
  const data = await getContents(repoPath);
  if (!data || Array.isArray(data)) return null;
  if (!data.content || data.content.trim() === '') {
    if (!data.download_url) return null;
    const res = await fetch(data.download_url, {
      headers: { Authorization: `Bearer ${TOKEN}` },
    });
    if (!res.ok) throw new Error(`Failed to download ${repoPath}: ${res.status}`);
    return Buffer.from(await res.arrayBuffer());
  }
  return Buffer.from(data.content, 'base64');
}

// Create or update a file (commits to BRANCH).
async function putFile(repoPath, buffer, message) {
  const existing = await getContents(repoPath);
  const body = {
    message,
    content: buffer.toString('base64'),
    branch: BRANCH,
    ...(existing && !Array.isArray(existing) ? { sha: existing.sha } : {}),
  };
  const res = await gh(`/repos/${OWNER}/${REPO}/contents/${encodeGitPath(repoPath)}`, {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  if (!res.ok) throw new Error(`GitHub PUT ${repoPath} failed: ${await ghErrText(res)}`);
  return res.json();
}

// Delete a file (commits to BRANCH). No-op if it doesn't exist.
async function deleteFile(repoPath, message) {
  const existing = await getContents(repoPath);
  if (!existing || Array.isArray(existing)) return { skipped: true };
  const res = await gh(`/repos/${OWNER}/${REPO}/contents/${encodeGitPath(repoPath)}`, {
    method: 'DELETE',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ message, sha: existing.sha, branch: BRANCH }),
  });
  if (!res.ok) throw new Error(`GitHub DELETE ${repoPath} failed: ${await ghErrText(res)}`);
  return res.json();
}

// ── WBR Data (data/<weekName>/*.xlsx) ───────────────────────────────────────

const DATA_ROOT = 'data';
// PCR workbook (Costs-tab Labor/COGS actuals) lives in a separate root,
// mirroring the same "<weekName>" sub-folder convention: PCR/<weekName>/*.xlsx.
const PCR_ROOT = 'PCR';

function detectFileType(filename) {
  const n = filename.toLowerCase();
  if (n.includes('loyalty')) return 'loyalty';
  if (n.includes('catering') || n.includes('internal purpose')) return 'catering';
  if (n.includes('weekly review') || n.includes('powered by kutlerri') || n.includes('wbr')) return 'wbr';
  return null;
}

export async function listWeeks() {
  const entries = await listDir(DATA_ROOT);
  return entries.filter(e => e.type === 'dir').map(e => e.name).sort();
}

// { present, paths, shas } for one week. shas used as cache fingerprint.
export async function getWeekStatus(weekName) {
  const files = await listDir(`${DATA_ROOT}/${weekName}`);
  const present = { wbr: false, loyalty: false, catering: false };
  const paths   = { wbr: null,  loyalty: null,  catering: null  };
  const shas    = { wbr: null,  loyalty: null,  catering: null  };
  for (const f of files) {
    if (f.type !== 'file') continue;
    const type = detectFileType(f.name);
    if (type) { present[type] = true; paths[type] = f.path; shas[type] = f.sha; }
  }
  return { present, paths, shas };
}

// { present, path, sha } for one week's PCR workbook, if one has been provided.
export async function getPcrFileStatus(weekName) {
  const files = await listDir(`${PCR_ROOT}/${weekName}`);
  const file = files.find(f => f.type === 'file' && f.name.toLowerCase().includes('pcr'));
  return file ? { present: true, path: file.path, sha: file.sha } : { present: false, path: null, sha: null };
}

// Download a file at a known repo path. Used by the data API to fetch xlsx files at runtime.
export async function downloadFileAtPath(repoPath) {
  return getFileBuffer(repoPath);
}

export async function getAllWeeks() {
  const weekNames = await listWeeks();
  const weeks = [];
  for (const weekName of weekNames) {
    const { present } = await getWeekStatus(weekName);
    const { present: pcr } = await getPcrFileStatus(weekName);
    weeks.push({ weekName, ...present, pcr });
  }
  return weeks;
}

// PCR workbook lives under PCR_ROOT (separate from the wbr/loyalty/catering
// files under DATA_ROOT) — same admin actions (upload/replace/download/delete),
// own storage location.
export async function uploadPcrFile(weekName, buffer) {
  await deletePcrFile(weekName);
  const path = `${PCR_ROOT}/${weekName}/PCR.xlsx`;
  await putFile(path, buffer, `Admin: upload PCR for ${weekName}`);
  return path;
}

export async function downloadPcrFile(weekName) {
  const { path } = await getPcrFileStatus(weekName);
  if (!path) return null;
  const buffer = await getFileBuffer(path);
  return buffer ? { buffer, filename: path.split('/').pop() } : null;
}

export async function deletePcrFile(weekName) {
  const { path } = await getPcrFileStatus(weekName);
  if (!path) return { skipped: true };
  return deleteFile(path, `Admin: delete PCR for ${weekName}`);
}

export async function uploadWeekFile(weekName, fileType, buffer) {
  // Delete any existing file for this type first (may have a non-standard filename).
  await deleteWeekFile(weekName, fileType);
  const path = `${DATA_ROOT}/${weekName}/${fileType}.xlsx`;
  await putFile(path, buffer, `Admin: upload ${fileType} for ${weekName}`);
  return path;
}

export async function downloadWeekFile(weekName, fileType) {
  const { paths } = await getWeekStatus(weekName);
  const path = paths[fileType];
  if (!path) return null;
  const buffer = await getFileBuffer(path);
  return buffer ? { buffer, filename: path.split('/').pop() } : null;
}

export async function deleteWeekFile(weekName, fileType) {
  const { paths } = await getWeekStatus(weekName);
  const path = paths[fileType];
  if (!path) return { skipped: true };
  return deleteFile(path, `Admin: delete ${fileType} for ${weekName}`);
}

export async function deleteWeek(weekName) {
  const files = await listDir(`${DATA_ROOT}/${weekName}`);
  for (const f of files) {
    if (f.type === 'file') await deleteFile(f.path, `Admin: delete week ${weekName}`);
  }
}

// ── Scorecard (scorecard/<granularity>/*.xlsx) ──────────────────────────────

const SCORECARD_ROOT = 'scorecard';
const GRANULARITIES = ['weekly', 'period', 'quarter'];

export async function listScorecardFiles() {
  const result = { weekly: [], period: [], quarter: [] };
  for (const g of GRANULARITIES) {
    const files = await listDir(`${SCORECARD_ROOT}/${g}`);
    result[g] = files.filter(f => f.type === 'file').map(f => ({ filename: f.name }));
  }
  return result;
}

export async function uploadScorecardFile(granularity, buffer, filename) {
  const path = `${SCORECARD_ROOT}/${granularity}/${filename}`;
  await putFile(path, buffer, `Admin: upload scorecard ${granularity}/${filename}`);
  return path;
}

export async function downloadScorecardFile(granularity, filename) {
  const path = `${SCORECARD_ROOT}/${granularity}/${filename}`;
  const buffer = await getFileBuffer(path);
  return buffer ? { buffer, filename } : null;
}

export async function deleteScorecardFile(granularity, filename) {
  const path = `${SCORECARD_ROOT}/${granularity}/${filename}`;
  return deleteFile(path, `Admin: delete scorecard ${granularity}/${filename}`);
}

// ── User config (config/users.json) ─────────────────────────────────────────

const USERS_CONFIG_PATH = 'config/users.json';

// Returns parsed users object from GitHub, or null if not found.
export async function getUsersConfig() {
  const data = await getContents(USERS_CONFIG_PATH);
  if (!data || Array.isArray(data)) return null;
  try {
    return JSON.parse(Buffer.from(data.content, 'base64').toString('utf8'));
  } catch { return null; }
}

// Writes the users object to config/users.json in GitHub.
export async function saveUsersConfig(usersObj) {
  const buffer = Buffer.from(JSON.stringify(usersObj, null, 2));
  await putFile(USERS_CONFIG_PATH, buffer, 'Admin: update users');
}

// ── Locations config (config/locations.json) ─────────────────────────────────

const LOCATIONS_CONFIG_PATH = 'config/locations.json';

// Returns parsed locations object from GitHub, or null if not found.
export async function getLocationsConfig() {
  const data = await getContents(LOCATIONS_CONFIG_PATH);
  if (!data || Array.isArray(data)) return null;
  try {
    return JSON.parse(Buffer.from(data.content, 'base64').toString('utf8'));
  } catch { return null; }
}

// Writes the locations object to config/locations.json in GitHub.
export async function saveLocationsConfig(locationsObj) {
  const buffer = Buffer.from(JSON.stringify(locationsObj, null, 2));
  await putFile(LOCATIONS_CONFIG_PATH, buffer, 'Admin: update locations');
}
