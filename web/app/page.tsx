import type { Metadata } from "next";
import { headers } from "next/headers";
import { TravelExpenseApp } from "./travel-expense-app";

export async function generateMetadata(): Promise<Metadata> {
  const requestHeaders = await headers();
  const host = requestHeaders.get("x-forwarded-host") || requestHeaders.get("host") || "localhost:3000";
  const protocol = requestHeaders.get("x-forwarded-proto") || (host.startsWith("localhost") ? "http" : "https");
  const image = `${protocol}://${host}/og.png`;
  return {
    title: "出張旅費申請書作成アプリ",
    description: "予定から1日の移動経路を組み立て、定期券控除後の旅費申請を作成します。",
    openGraph: { title: "出張旅費申請書作成アプリ", description: "予定をつなげて、申請できる移動だけを整理", images: [image] },
    twitter: { card: "summary_large_image", images: [image] },
  };
}

export default function Home() {
  return <TravelExpenseApp />;
}
