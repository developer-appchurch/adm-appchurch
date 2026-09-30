import app from "./app";

// Handler para Vercel Serverless Functions
export default function handler(req: any, res: any) {
  // Se o Vercel reescreveu a URL interna (ex: para /api ou /api/index),
  // restaura o caminho original a partir dos headers de roteamento do Vercel
  const matchedPath = req.headers["x-matched-path"] || req.headers["x-forwarded-url"] || req.headers["x-now-route-matches"];
  if (typeof matchedPath === "string" && (req.url === "/api" || req.url.startsWith("/api/index") || req.url === "/" || req.url.startsWith("/?"))) {
    req.url = matchedPath;
  }

  // Garante que requisições que chegam sem /api sejam mapeadas
  if (typeof req.url === "string" && req.url.startsWith("/sharepoint/")) {
    req.url = "/api" + req.url;
  }

  return app(req, res);
}
