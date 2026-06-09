/** @type {import('next').NextConfig} */
const nextConfig = {
  reactStrictMode: true,

  // Proxy /api/orders to the external Vercel app
  // Keep /api/mb/* handled locally by Next.js API routes
  async rewrites() {
    return [
      {
        source: '/api/orders',
        destination: 'https://dat-com-ivory.vercel.app/api/orders',
      },
    ];
  },
};

module.exports = nextConfig;
