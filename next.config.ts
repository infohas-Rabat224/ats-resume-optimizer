import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  output: "standalone",
  /* config options here */
  typescript: {
    ignoreBuildErrors: true,
  },
  reactStrictMode: false,
  // Exclude pdf.js from server bundle since it's client-side only
  serverExternalPackages: ['pdfjs-dist', 'pdf-parse'],
};

export default nextConfig;
