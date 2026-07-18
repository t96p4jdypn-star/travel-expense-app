import { NextResponse } from "next/server";

const array = <T,>(value: T | T[] | undefined): T[] => value == null ? [] : Array.isArray(value) ? value : [value];

export async function POST(request: Request) {
  const key = process.env.EKISPERT_API_KEY;
  if (!key) return NextResponse.json({ code: "NOT_CONFIGURED", message: "自動運賃サービスは準備中です" }, { status: 503 });
  try {
    const body = await request.json() as { origin?: string; arrival?: string; date?: string; time?: string };
    const origin = body.origin?.trim(); const arrival = body.arrival?.trim();
    if (!origin || !arrival) return NextResponse.json({ message: "出発駅と到着駅が必要です" }, { status: 400 });
    const params = new URLSearchParams({
      key, viaList: `${origin}:${arrival}`, date: (body.date || "").replaceAll("-", ""),
      time: (body.time || "0900").replace(":", ""), searchType: "departure", sort: "price", answerCount: "3",
      conditionDetail: "T3221233232319:F3321121120000:A23121141:",
    });
    const response = await fetch(`https://api.ekispert.jp/v1/json/search/course/extreme?${params}`, { signal: AbortSignal.timeout(12000) });
    const data = await response.json() as any;
    if (!response.ok || data.ResultSet?.Error) return NextResponse.json({ code: "UPSTREAM_ERROR", message: "最新運賃を取得できませんでした" }, { status: 502 });
    const courses = array<any>(data.ResultSet?.Course);
    const candidate = courses.map((course) => {
      const prices = array<any>(course.Price);
      const ic = prices.find((price) => price.Type === "FareICCard" && price.selected !== "false");
      const route = course.Route ?? {}; const points = array<any>(route.Point);
      const lines = array<any>(route.Line).map((line) => line.Name).filter(Boolean);
      return {
        fare: Number(ic?.Oneway || 0), revisionStatus: ic?.RevisionStatus || "unknown",
        route: lines.join(" → "), paidSection: `${points[0]?.Station?.Name || points[0]?.Name || origin}→${points.at(-1)?.Station?.Name || points.at(-1)?.Name || arrival}`,
      };
    }).find((item) => item.fare > 0);
    if (!candidate) return NextResponse.json({ code: "NO_FARE", message: "IC運賃を特定できませんでした" }, { status: 422 });
    return NextResponse.json({ ...candidate, checkedAt: new Date().toISOString(), provider: "駅すぱあと API" });
  } catch {
    return NextResponse.json({ code: "FARE_LOOKUP_FAILED", message: "最新運賃を取得できませんでした" }, { status: 502 });
  }
}
