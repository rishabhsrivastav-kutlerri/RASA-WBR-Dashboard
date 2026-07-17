// Reads precomputed JSON written by scripts/precompute.mjs at build time.
// Returns null if the file is missing so callers can fall back to live parsing.
import fs from 'fs';
import path from 'path';

const GEN_DIR = path.join(process.cwd(), 'generated');

export function readGenerated(relPath) {
  try {
    return JSON.parse(fs.readFileSync(path.join(GEN_DIR, relPath), 'utf8'));
  } catch {
    return null;
  }
}
