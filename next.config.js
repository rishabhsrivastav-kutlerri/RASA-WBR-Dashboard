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
  // NOTE: Turbopack (Next's default build engine as of this version) silently
  // ignores outputFileTracingIncludes/Excludes — it still traced in the full
  // data/PCR/scorecard trees with these excludes in place. package.json's
  // "build" script runs `next build --webpack` specifically so this config
  // actually takes effect. Don't drop that flag without re-verifying tracing.
  outputFileTracingExcludes: {
    '/api/data/[week]': ['./data/**', './PCR/**'],
    // Same issue via lib/scorecard.js's dynamic fs.readdirSync — the runtime
    // fallback in app/api/scorecard/route.js now downloads from GitHub instead
    // of reading scorecard/<granularity>/ locally, so this is safe to drop too.
    '/api/scorecard': ['./scorecard/**'],
  },
};

module.exports = nextConfig;
