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
};

module.exports = nextConfig;
