/** @type {import('next').NextConfig} */
const nextConfig = {
  reactStrictMode: true,
  serverExternalPackages: ['mbbank', 'xlsx'],
};

module.exports = nextConfig;
