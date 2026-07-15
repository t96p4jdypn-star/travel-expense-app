import type { Metadata, Viewport } from "next";
import "./globals.css";
import "./extras.css";

export const metadata: Metadata = {
  title: "出張旅費申請書作成アプリ",
  description: "端末内だけで予定を整理し、出張旅費申請書を作成します。",
  applicationName: "出張旅費申請書作成アプリ",
  manifest: "/manifest.webmanifest",
  icons: { icon: "/favicon.svg", apple: "/favicon.svg" },
};

export const viewport: Viewport = {
  width: "device-width",
  initialScale: 1,
  themeColor: "#173f3a",
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return <html lang="ja"><body>{children}</body></html>;
}
