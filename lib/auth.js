import jwt from 'jsonwebtoken';

// Verify a Bearer token from the Authorization header. Returns the decoded
// payload (e.g. { email }) on success, or null if missing/invalid/expired.
export function verifyAuth(request) {
  const header = request.headers.get('authorization') || '';
  const token = header.startsWith('Bearer ') ? header.slice(7) : null;
  if (!token) return null;
  try {
    return jwt.verify(token, process.env.JWT_SECRET);
  } catch {
    return null;
  }
}

// Returns the user map { email: bcrypt_hash } parsed from USERS_JSON env.
// Empty object if the env var is missing or malformed.
export function getUsers() {
  try {
    return JSON.parse(process.env.USERS_JSON || '{}');
  } catch {
    return {};
  }
}
