/** @type {import('next').NextConfig} */
const nextConfig = {
  output: 'export',
  trailingSlash: true,
  skipTrailingSlashRedirect: true,
  transpilePackages: ['recharts'],
  images: {
    unoptimized: true,
  },
};

module.exports = nextConfig;
