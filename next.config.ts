import type { NextConfig } from "next";

const isProduction = process.env.NODE_ENV === "production";
const noStoreHeaders = [
  {
    key: "Cache-Control",
    value: "no-store, no-cache, max-age=0, must-revalidate",
  },
  { key: "Pragma", value: "no-cache" },
  { key: "Expires", value: "0" },
];

const nextConfig: NextConfig = {
  assetPrefix: isProduction ? "/infowb-assets" : undefined,
  async rewrites() {
    return [
      {
        source: "/infowb-assets/_next/:path*",
        destination: "/_next/:path*",
      },
    ];
  },
  async headers() {
    return [
      { source: "/", headers: noStoreHeaders },
      { source: "/compare", headers: noStoreHeaders },
      { source: "/login", headers: noStoreHeaders },
    ];
  },
};

export default nextConfig;
