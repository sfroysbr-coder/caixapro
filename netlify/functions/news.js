const https = require("https");

function get(url) {
  return new Promise((resolve, reject) => {
    const req = https.get(url, {
      headers: {
        "User-Agent": "Mozilla/5.0 (compatible; CaixaProBot/1.0)",
        "Accept": "application/rss+xml, application/xml, text/xml, */*",
      },
    }, (res) => {
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        return get(res.headers.location).then(resolve).catch(reject);
      }
      if (res.statusCode !== 200) {
        return reject(new Error("HTTP " + res.statusCode));
      }
      const chunks = [];
      res.on("data", c => chunks.push(c));
      res.on("end", () => resolve(Buffer.concat(chunks).toString("utf8")));
    });
    req.on("error", reject);
    req.setTimeout(10000, () => { req.destroy(); reject(new Error("Timeout")); });
  });
}

function tag(xml, name) {
  const re = new RegExp(`<${name}[^>]*>(?:<!\\[CDATA\\[)?([\\s\\S]*?)(?:\\]\\]>)?<\\/${name}>`, "i");
  const m = xml.match(re);
  return m ? m[1].replace(/<[^>]*>/g, "").trim() : "";
}

function parseRSS(xml) {
  const blocks = xml.match(/<item[\s\S]*?<\/item>/g) || [];
  return blocks.slice(0, 18).map(b => {
    const title  = tag(b, "title").replace(/\s+-\s+[\w\s]+$/, "").trim();
    const link   = tag(b, "link") || tag(b, "guid");
    const date   = tag(b, "pubDate");
    const source = tag(b, "source") || "";
    const desc   = tag(b, "description").slice(0, 220).trim();
    return { title, link, date, source, description: desc };
  }).filter(i => i.title && i.link);
}

const QUERIES = {
  tirzepatida:   "tirzepatida+medicamento",
  mounjaro:      "mounjaro+tirzepatida",
  emagrecimento: "emagrecimento+medicamento+tratamento",
  glp1:          "ozempic+semaglutida+GLP-1",
  saude:         "saude+alimentacao+dieta+emagrecer",
};

exports.handler = async (event) => {
  const H = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Headers": "Content-Type",
    "Content-Type": "application/json",
  };
  if (event.httpMethod === "OPTIONS") return { statusCode: 200, headers: H };

  const cat   = (event.queryStringParameters || {}).category || "tirzepatida";
  const query = QUERIES[cat] || QUERIES.tirzepatida;

  const urls = [
    `https://news.google.com/rss/search?q=${encodeURIComponent(query)}&hl=pt-BR&gl=BR&ceid=BR:pt`,
    `https://news.google.com/rss/search?q=${encodeURIComponent(query)}&hl=pt`,
    `https://news.google.com/rss/search?q=${encodeURIComponent(query)}`,
  ];

  let items = [];
  let lastErr = "";

  for (const url of urls) {
    try {
      const xml = await get(url);
      if (xml && xml.includes("<item>")) {
        items = parseRSS(xml);
        if (items.length > 0) break;
      }
    } catch (e) {
      lastErr = e.message;
    }
  }

  if (items.length === 0) {
    return {
      statusCode: 200,
      headers: H,
      body: JSON.stringify({ ok: false, error: lastErr || "Sem resultados", items: [] }),
    };
  }

  return {
    statusCode: 200,
    headers: H,
    body: JSON.stringify({ ok: true, cat, count: items.length, items }),
  };
};
