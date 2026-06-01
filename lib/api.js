// Client-side API helpers with JWT-bearer auth.
// Token is stored in localStorage under 'wbr_token'.

const TOKEN_KEY = 'wbr_token';

function token() {
  if (typeof window === 'undefined') return '';
  return localStorage.getItem(TOKEN_KEY) || '';
}

function authHeaders() {
  return { 'Content-Type': 'application/json', Authorization: `Bearer ${token()}` };
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
  return data;
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
