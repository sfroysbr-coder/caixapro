// Netlify Function — busca RSS do Google News no servidor (sem CORS)
const https = require("https");
const http  = require("http");

const fetchUrl = (url) =>
  new Promise((resolve, reject) => {
    const lib = url.startsWith("https") ? https : http;
    const req = lib.get(url, { headers: { "User-Agent": "Mozilla/5.0 CaixaProBot/1.0" } }, (res) => {
      // Seguir redirect
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        return fetchUrl(res.headers.location).then(resolve).catch(reject);
      }
      let data = "";
      res.on("data", (chunk) => (data += chunk));
      res.on("end", () => resolve(data));
    });
    req.on("error", reject);
    req.setTimeout(8000, () => { req.destroy(); reject(new Error("Timeout")); });
  });

const parseRSS = (xml) => {
  const items = [];
  const itemMatches = xml.match(/<item>([\s\S]*?)<\/item>/g) || [];

  for (const item of itemMatches.slice(0, 15)) {
    const get = (tag) => {
      const m = item.match(new RegExp(`<${tag}(?:[^>]*)?><!\\[CDATA\\[([\\s\\S]*?)\\]\\]><\\/${tag}>|<${tag}(?:[^>]*)?>([\\s\\S]*?)<\\/${tag}>`));
      return m ? (m[1] || m[2] || "").trim() : "";
    };

    const title  = get("title").replace(/<[^>]*>/g, "").replace(/\s+-\s+.+$/, "").trim();
    const link   = get("link") || get("guid");
    const date   = get("pubDate");
    const source = get("source") || "";
    const desc   = get("description").replace(/<[^>]*>/g, "").slice(0, 250).trim();

    if (title && link) {
      items.push({ title, link, date, source, description: desc });
    }
  }
  return items;
};

const QUERIES = {
  tirzepatida:  "tirzepatida",
  mounjaro:     "mounjaro+tirzepatida+injetável",
  emagrecimento:"emagrecimento+medicamento+tratamento",
  glp1:         "ozempic+semaglutida+GLP-1+emagrecimento",
  saude:        "saúde+alimentação+dieta+emagrecer",
};

exports.handler = async (event) => {
  const headers = {
    "Access-Control-Allow-Origin":  "*",
    "Access-Control-Allow-Headers": "Content-Type",
    "Content-Type": "application/json",
  };

  if (event.httpMethod === "OPTIONS") return { statusCode: 200, headers };

  const category = event.queryStringParameters?.category || "tirzepatida";
  const query    = QUERIES[category] || QUERIES.tirzepatida;

  try {
    const rssUrl = `https://news.google.com/rss/search?q=${encodeURIComponent(query)}&hl=pt-BR&gl=BR&ceid=BR:pt`;
    const xml    = await fetchUrl(rssUrl);

    if (!xml || xml.length < 100) {
      return { statusCode: 200, headers, body: JSON.stringify({ ok: false, error: "Feed vazio", items: [] }) };
    }

    const items = parseRSS(xml);

    return {
      statusCode: 200,
      headers,
      body: JSON.stringify({ ok: true, category, query, count: items.length, items }),
    };
  } catch (err) {
    return {
      statusCode: 200,
      headers,
      body: JSON.stringify({ ok: false, error: err.message, items: [] }),
    };
  }
};
