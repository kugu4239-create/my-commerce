// 자사몰(merryon.co.kr) 검색 페이지 프록시 — CORS 우회 + 공용 CORS 프록시 불안정 해결
// 호출: GET /api/merryon-search?keyword=에이플
// 응답: 200 + merryon 검색 결과 HTML (text/html; charset=utf-8)
//
// Vercel Serverless (Node) — fetch 내장 사용 (Node 18+)

export default async function handler(req, res) {
  const keyword = String(req.query?.keyword || "").trim();
  if (!keyword) {
    res.status(400).json({ error: "keyword required" });
    return;
  }
  const url = `https://merryon.co.kr/product/search.html?banner_action=&keyword=${encodeURIComponent(keyword)}`;
  try {
    const r = await fetch(url, {
      headers: {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.8",
      },
    });
    if (!r.ok) {
      res.status(502).json({ error: `merryon HTTP ${r.status}` });
      return;
    }
    const html = await r.text();
    res.setHeader("Content-Type", "text/html; charset=utf-8");
    // 5분 edge 캐시 + 30초 stale-while-revalidate — 같은 키워드 반복 요청 시 비용 절감
    res.setHeader("Cache-Control", "public, s-maxage=300, stale-while-revalidate=30");
    res.status(200).send(html);
  } catch (err) {
    res.status(500).json({ error: String(err?.message || err) });
  }
}
