/** @type {import('next').NextConfig} */
const nextConfig = {
  experimental: {
    // Allow large file uploads via Server Actions
    serverActions: { bodySizeLimit: '20mb' },
  },
  // Bundle the build-time precomputed JSON into the serverless functions that
  // read it, so fs.readFileSync('generated/...') works at runtime on Vercel.
  outputFileTracingIncludes: {
    '/api/data/[week]': ['./generated/**'],
    '/api/sheets': ['./generated/**'],
    '/api/scorecard': ['./generated/**'],
  },
  // xlsxParser.js's dynamic fs.readdirSync/path.join calls make Next's file
  // tracer overcautious, so it drags the entire checked-in data/ and PCR/
  // workbook trees (200MB+, growing with every admin upload) into this
  // function even though its runtime fallback only ever reads from GitHub
  // downloads under /tmp — never from these repo paths. That pushed the
  // function past Vercel's 250MB uncompressed limit.
  outputFileTracingExcludes: {
    '/api/data/[week]': ['./data/**', './PCR/**'],
  },
};

module.exports = nextConfig;
