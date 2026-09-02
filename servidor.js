/*
 * Escala do Pátio — servidor
 *
 * Node puro, sem nenhuma biblioteca externa. Faz duas coisas:
 *   1. serve os arquivos do app (pasta dist/)
 *   2. guarda os dados compartilhados em dados/quadro.json
 *
 * Rodar:  node servidor.js
 * Abrir:  http://localhost:8080   (ou o IP da máquina, na TV)
 */

const http = require("http");
const fs = require("fs");
const path = require("path");

const PORTA = process.env.PORTA || 8080;
const RAIZ = __dirname;
const DIST = path.join(RAIZ, "dist");
const ARQUIVO = path.join(RAIZ, "dados", "quadro.json");
const BACKUPS = path.join(RAIZ, "dados", "backups");

/* ------------------------------------------------------------------ */
/*  Armazenamento                                                      */
/* ------------------------------------------------------------------ */

let dados = {};

function carregar() {
  try {
    dados = JSON.parse(fs.readFileSync(ARQUIVO, "utf8"));
    console.log(`Dados carregados: ${Object.keys(dados).length} registros.`);
  } catch {
    dados = {};
    console.log("Começando com a base vazia.");
  }
}

/* Grava em arquivo temporário e só então substitui, para nunca ficar
   com um JSON pela metade se a máquina desligar no meio. */
let gravacaoPendente = null;
function gravar() {
  if (gravacaoPendente) return;
  gravacaoPendente = setTimeout(() => {
    gravacaoPendente = null;
    try {
      fs.mkdirSync(path.dirname(ARQUIVO), { recursive: true });
      const tmp = `${ARQUIVO}.tmp`;
      fs.writeFileSync(tmp, JSON.stringify(dados));
      fs.renameSync(tmp, ARQUIVO);
    } catch (e) {
      console.error("Falha ao gravar:", e.message);
    }
  }, 300);
}

/* Uma cópia por dia, mantendo as 30 últimas. */
function backupDiario() {
  try {
    if (!fs.existsSync(ARQUIVO)) return;
    fs.mkdirSync(BACKUPS, { recursive: true });
    const hoje = new Date().toISOString().slice(0, 10);
    const alvo = path.join(BACKUPS, `quadro-${hoje}.json`);
    if (!fs.existsSync(alvo)) fs.copyFileSync(ARQUIVO, alvo);
    const antigos = fs.readdirSync(BACKUPS).sort();
    while (antigos.length > 30) fs.unlinkSync(path.join(BACKUPS, antigos.shift()));
  } catch (e) {
    console.error("Falha no backup:", e.message);
  }
}

/* ------------------------------------------------------------------ */
/*  API                                                                */
/* ------------------------------------------------------------------ */

const responder = (res, codigo, corpo) => {
  res.writeHead(codigo, { "Content-Type": "application/json; charset=utf-8" });
  res.end(JSON.stringify(corpo));
};

function api(req, res, url) {
  const chave = decodeURIComponent(url.searchParams.get("chave") || "");

  if (req.method === "GET" && url.pathname === "/api/valor") {
    if (!(chave in dados)) return responder(res, 404, { erro: "não encontrado" });
    return responder(res, 200, { chave, valor: dados[chave] });
  }

  if (req.method === "GET" && url.pathname === "/api/lista") {
    const prefixo = url.searchParams.get("prefixo") || "";
    return responder(res, 200, { chaves: Object.keys(dados).filter((k) => k.startsWith(prefixo)) });
  }

  if (req.method === "DELETE" && url.pathname === "/api/valor") {
    delete dados[chave];
    gravar();
    return responder(res, 200, { chave, apagado: true });
  }

  if (req.method === "PUT" && url.pathname === "/api/valor") {
    let corpo = "";
    req.on("data", (p) => {
      corpo += p;
      if (corpo.length > 6e6) req.destroy(); // limite de 6 MB por registro
    });
    req.on("end", () => {
      try {
        dados[chave] = JSON.parse(corpo).valor;
        gravar();
        responder(res, 200, { chave, ok: true });
      } catch {
        responder(res, 400, { erro: "corpo inválido" });
      }
    });
    return;
  }

  responder(res, 404, { erro: "rota desconhecida" });
}

/* ------------------------------------------------------------------ */
/*  Arquivos do app                                                    */
/* ------------------------------------------------------------------ */

const TIPOS = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".svg": "image/svg+xml",
  ".png": "image/png",
  ".ico": "image/x-icon",
  ".woff2": "font/woff2",
};

function estatico(req, res, url) {
  let rel = url.pathname === "/" ? "/index.html" : url.pathname;
  try { rel = decodeURIComponent(rel); } catch { return responder(res, 400, { erro: "caminho inválido" }); }
  const alvoBruto = path.resolve(DIST, `.${path.posix.normalize(rel)}`);
  /* resolve() elimina qualquer ".." — se ainda assim saiu de dist/, recusa */
  if (alvoBruto !== DIST && !alvoBruto.startsWith(DIST + path.sep))
    return responder(res, 403, { erro: "caminho inválido" });
  let alvo = alvoBruto;
  if (!fs.existsSync(alvo) || fs.statSync(alvo).isDirectory()) alvo = path.join(DIST, "index.html");
  if (!fs.existsSync(alvo)) {
    res.writeHead(500, { "Content-Type": "text/plain; charset=utf-8" });
    return res.end("A pasta dist/ não existe. Rode antes: npm install && npm run build");
  }
  res.writeHead(200, { "Content-Type": TIPOS[path.extname(alvo)] || "application/octet-stream" });
  fs.createReadStream(alvo).pipe(res);
}

/* ------------------------------------------------------------------ */

carregar();
backupDiario();
setInterval(backupDiario, 6 * 60 * 60 * 1000);

const servidor = http.createServer((req, res) => {
  /* Uma requisição malformada não pode derrubar o quadro no meio do turno. */
  try {
    const url = new URL(req.url, `http://${req.headers.host || "localhost"}`);
    if (url.pathname.startsWith("/api/")) return api(req, res, url);
    estatico(req, res, url);
  } catch (e) {
    console.error("Requisição inválida:", req.method, req.url, "-", e.message);
    try { responder(res, 400, { erro: "requisição inválida" }); } catch {}
  }
});

servidor
  .listen(PORTA, () => {
    console.log(`\nEscala do Pátio rodando.`);
    console.log(`  Nesta máquina:  http://localhost:${PORTA}`);
    for (const [nome, addrs] of Object.entries(require("os").networkInterfaces()))
      for (const a of addrs || [])
        if (a.family === "IPv4" && !a.internal)
          console.log(`  Na rede (${nome}):  http://${a.address}:${PORTA}`);
    console.log(`\nDados em: ${ARQUIVO}`);
    console.log("Para parar, aperte Ctrl+C.\n");
  });

/* Último anteparo: registra e segue, em vez de encerrar o processo. */
process.on("uncaughtException", (e) => console.error("Erro não tratado:", e));
process.on("unhandledRejection", (e) => console.error("Promessa rejeitada:", e));

process.on("SIGINT", () => {
  if (gravacaoPendente) clearTimeout(gravacaoPendente);
  try {
    fs.mkdirSync(path.dirname(ARQUIVO), { recursive: true });
    fs.writeFileSync(ARQUIVO, JSON.stringify(dados));
  } catch {}
  console.log("\nDados salvos. Até logo.");
  process.exit(0);
});
