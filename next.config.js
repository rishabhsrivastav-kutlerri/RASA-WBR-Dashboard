/** @type {import('next').NextConfig} */
const nextConfig = {
  experimental: {
    // Allow large file uploads via Server Actions
    serverActions: { bodySizeLimit: '20mb' },
  },
};

module.exports = nextConfig;
