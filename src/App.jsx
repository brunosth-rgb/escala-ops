import React, { useState, useEffect, useMemo, useRef, useCallback } from "react";
import * as XLSX from "xlsx";

/* ================================================================== */
/*  MODELO                                                             */
/* ================================================================== */

const TURNOS_6H = [
  { id: "01X07", label: "01·07" },
  { id: "07X13", label: "07·13" },
  { id: "13X19", label: "13·19" },
  { id: "19X01", label: "19·01" },
];
const TURNOS_12H = [
  { id: "07X19", label: "07·19" },
  { id: "19X07", label: "19·07" },
];
const TODOS_TURNOS = [...TURNOS_6H, ...TURNOS_12H];
const eh12h = (t) => t === "07X19" || t === "19X07";
const rotuloTurno = (t) => TODOS_TURNOS.find((x) => x.id === t)?.label.replace("·", " x ") || t;

const FUNCOES = {
  III: { nome: "Operador III", curto: "III", fams: ["RS"] },
  IV: { nome: "Operador IV", curto: "IV", fams: ["STS", "MHC"] },
  TT: { nome: "Operador TT", curto: "TT", fams: ["CAM"] },
};

/** Cada função só opera o seu equipamento. Vale para clique, arrasto e automações. */
const funcaoDe = (op) => (FUNCOES[op?.funcao] ? op.funcao : "III"); /* TPAs antigos vinham sem função */
const podeOperar = (op, eq) => !!op && !!eq && FUNCOES[funcaoDe(op)].fams.includes(eq.fam);
const equipamentosDaFuncao = (fn) => (FUNCOES[fn]?.fams || []).map((f) => FAMILIAS[f].nome).join(" e ");

function funcaoDoArquivo(nome) {
  const n = nome.toUpperCase();
  if (/OP[_\s-]*IV/.test(n)) return "IV";
  if (/OP[_\s-]*III/.test(n)) return "III";
  if (/OP[_\s-]*TT/.test(n)) return "TT";
  return "TT";
}

/* Paleta JBS Terminais: marinho 2E2C75, verde 93C83D, azul 285EAC, ciano 4AC9F1 */
const FAMILIAS = {
  RS: { nome: "Reach Stacker", curto: "RS", cor: "#4AC9F1", prefere: ["III"] },
  CAM: { nome: "Caminhão", curto: "CAM", cor: "#93C83D", prefere: ["TT"] },
  STS: { nome: "STS", curto: "STS", cor: "#6B9FE0", prefere: ["IV"] },
  MHC: { nome: "MHC", curto: "MHC", cor: "#A79BF0", prefere: ["IV"] },
};
const ORDEM_FAM = ["RS", "CAM", "STS", "MHC"];

/* Como a lateral se organiza: uma seção por função, na cor do equipamento dela */
const GRUPOS_LATERAL = [
  { fn: "III", titulo: "Reach Stacker", sub: "Operador III", cor: "#4AC9F1" },
  { fn: "TT", titulo: "Caminhão", sub: "Operador TT", cor: "#93C83D" },
  { fn: "IV", titulo: "Guindaste", sub: "Operador IV", cor: "#6B9FE0" },
];

const CORES_FRENTE = ["#8891C4", "#4AC9F1", "#93C83D", "#6B9FE0", "#A79BF0", "#5FD3B4"];
const FRENTE_PADRAO = [{ id: "patio", nome: "Pátio", terno: "", cor: CORES_FRENTE[0] }];
const MOTIVOS_PARADA = ["Em manutenção", "Avaria", "Preventiva", "Abastecendo", "Sem operador"];
const intervaloPadrao = (iso) => (fimDeSemana(iso) ? 60 : 15);
const INTERVALO_OGMO = 30; /* TPAs do OGMO: 30 minutos, domingo a domingo */

function equipamentosPadrao() {
  const l = [];
  for (let i = 1; i <= 12; i++) l.push({ id: `RS${i}`, fam: "RS", nome: `RS ${String(i).padStart(2, "0")}` });
  for (let i = 1; i <= 25; i++) l.push({ id: `CAM${i}`, fam: "CAM", nome: `CAM ${String(i).padStart(2, "0")}` });
  l.push({ id: "STS1", fam: "STS", nome: "STS 01" });
  l.push({ id: "STS2", fam: "STS", nome: "STS 02" });
  l.push({ id: "MHCL", fam: "MHC", nome: "MHC Liebherr" });
  l.push({ id: "MHCK1", fam: "MHC", nome: "MHC KONE 01" });
  l.push({ id: "MHCK2", fam: "MHC", nome: "MHC KONE 02" });
  return l;
}

const STATUS = {
  TRABALHA: { rot: "Escalado", cor: "#93C83D" },
  DOZE: { rot: "Turno de 12h", cor: "#4AC9F1" },
  FOLGA: { rot: "Folga", cor: "#7A80AC" },
  FH: { rot: "FH", cor: "#A79BF0" },
  ATESTADO: { rot: "Atestado", cor: "#E8B24D" },
  FALTA: { rot: "Falta", cor: "#E05A63" },
  FERIAS: { rot: "Férias", cor: "#5FD3B4" },
  DESLIGADO: { rot: "Desligado", cor: "#C4494F" },
  VAZIO: { rot: "Sem marcação", cor: "#4A4F80" },
};
const AJUSTES = [
  { id: "TRABALHA", rot: "Presente" },
  { id: "ATESTADO", rot: "Atestado" },
  { id: "FALTA", rot: "Falta" },
  { id: "FOLGA", rot: "Folga" },
  { id: "FERIAS", rot: "Férias" },
];

function classificarCodigo(v) {
  const c = String(v ?? "").trim().toUpperCase();
  if (!c) return "VAZIO";
  if (/^\d{2}\s*X\s*\d{2}$/.test(c)) return "DOZE";
  if (/^6([,.]0+)?$/.test(c)) return "TRABALHA";
  if (c.startsWith("DESLIG")) return "DESLIGADO";
  if (c.startsWith("FÉRIAS") || c.startsWith("FERIAS")) return "FERIAS";
  if (c.startsWith("ATES")) return "ATESTADO";
  if (c === "FH") return "FH";
  if (c.startsWith("FOLGA")) return "FOLGA";
  if (c.startsWith("FALTA")) return "FALTA";
  return "VAZIO";
}
const normTurno = (v) => String(v ?? "").trim().toUpperCase().replace(/\s/g, "");

/**
 * O turno 01x07 é a ponta final do 19x07 da véspera: quando a planilha não
 * marca nada no dia (segunda depois de fim de semana de 12h) buscamos ontem.
 */
function situacao(op, data, turno, ajustes, dataAnterior) {
  const aj = ajustes?.[op.id];
  const d = op.dias?.[data] || {};
  const c6 = d.c6 ?? null;
  const c12 = d.c12 ?? d.c6 ?? null;

  let status, disponivel, codigo, herdado = false;
  if (eh12h(turno)) {
    codigo = c12;
    status = classificarCodigo(c12);
    disponivel = status === "DOZE" && normTurno(c12) === turno;
  } else {
    codigo = c6;
    status = classificarCodigo(c6);
    disponivel = status === "TRABALHA" && op.turno === turno;
    if (!disponivel && turno === "01X07" && dataAnterior) {
      const ant = op.dias?.[dataAnterior] || {};
      const cAnt = ant.c12 ?? ant.c6;
      if (normTurno(cAnt) === "19X07") { disponivel = true; status = "DOZE"; codigo = cAnt; herdado = true; }
    }
  }
  if (op.avulso) { status = "TRABALHA"; disponivel = true; codigo = null; }
  if (aj) { status = aj.status; disponivel = aj.status === "TRABALHA"; }
  return { status, disponivel, codigo, herdado, ajuste: aj || null };
}

/* ================================================================== */
/*  LEITURA DO EXCEL                                                   */
/* ================================================================== */

const isoDe = (d) =>
  `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;

function lerPlanilha(buffer, nomeArquivo) {
  const wb = XLSX.read(buffer, { type: "array", cellDates: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  const funcao = funcaoDoArquivo(nomeArquivo);

  const meses = [];
  for (let r = 0; r < Math.min(6, rows.length); r++)
    for (const cel of rows[r] || [])
      if (cel instanceof Date) meses.push(new Date(cel.getFullYear(), cel.getMonth(), 1));
  meses.sort((a, b) => a - b);
  if (!meses.length) { const h = new Date(); meses.push(new Date(h.getFullYear(), h.getMonth(), 1)); }
  const mesEm = (i) => {
    if (meses[i]) return meses[i];
    const u = meses[meses.length - 1];
    return new Date(u.getFullYear(), u.getMonth() + i - meses.length + 1, 1);
  };

  let colNome = 1, colMat = 2, colRS = -1;
  for (const row of rows.slice(0, 8)) {
    if (!row) continue;
    row.forEach((v, i) => {
      const t = String(v || "").trim().toUpperCase();
      if (t === "NOME") colNome = i;
      if (t === "MAT") colMat = i;
      if (t === "RS") colRS = i;
    });
  }

  const reTurno = /^(\d{2})\s*[Xx]\s*(\d{2})$/;
  const operadores = [];
  let periodo = { de: null, ate: null };

  for (let r = 0; r < rows.length; r++) {
    const row = rows[r] || [];
    let colTurno = -1, turnoId = null;
    for (let c = 0; c < row.length; c++) {
      const m = reTurno.exec(String(row[c] || "").trim());
      if (m) { colTurno = c; turnoId = `${m[1]}X${m[2]}`; break; }
    }
    if (colTurno < 0 || !TURNOS_6H.some((t) => t.id === turnoId)) continue;

    const linhaDias = rows[r + 1] || [];
    const mapa = [];
    let mIdx = 0, ultimo = 0;
    const vistos = new Set();
    for (let c = colTurno + 1; c < linhaDias.length; c++) {
      const v = linhaDias[c];
      const dia = typeof v === "number" ? v : parseInt(String(v || "").trim(), 10);
      if (!dia || dia < 1 || dia > 31) continue;
      if (dia < ultimo) mIdx++;
      ultimo = dia;
      const base = mesEm(mIdx);
      const iso = isoDe(new Date(base.getFullYear(), base.getMonth(), dia));
      mapa.push({ col: c, iso, bloco: vistos.has(iso) ? 12 : 6 });
      vistos.add(iso);
      if (!periodo.de || iso < periodo.de) periodo.de = iso;
      if (!periodo.ate || iso > periodo.ate) periodo.ate = iso;
    }

    let equipe = null;
    for (let rr = r + 2; rr < rows.length; rr++) {
      const lin = rows[rr] || [];
      const nome = String(lin[colNome] || "").trim();
      if (!nome) break;
      if (String(lin[0] || "").trim()) equipe = String(lin[0]).trim().replace(/^EQUIPE\s+/i, "");
      const dias = {};
      for (const { col, iso, bloco } of mapa) {
        const val = lin[col];
        if (val === null || val === undefined || val === "") continue;
        dias[iso] = dias[iso] || {};
        if (bloco === 12) dias[iso].c12 = String(val).trim();
        else dias[iso].c6 = String(val).trim();
      }
      const mat = String(lin[colMat] ?? "").trim();
      operadores.push({
        id: `${funcao}-${mat || nome}`,
        nome: nome.replace(/\s+/g, " "),
        mat, funcao, equipe, turno: turnoId, origem: nomeArquivo, dias,
        rs: colRS >= 0 && lin[colRS] != null ? String(lin[colRS]).trim() : null,
      });
      r = rr;
    }
  }
  return { operadores, periodo };
}

/* ================================================================== */
/*  UTILIDADES                                                         */
/* ================================================================== */

const DIAS_SEM = ["Domingo", "Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado"];
const dataDe = (iso) => { const [a, m, d] = String(iso).split("-").map(Number); return new Date(a, m - 1, d); };
const rotuloData = (iso) => {
  if (!iso) return "";
  const dt = dataDe(iso);
  return `${DIAS_SEM[dt.getDay()]}, ${String(dt.getDate()).padStart(2, "0")}/${String(dt.getMonth() + 1).padStart(2, "0")}`;
};
const fimDeSemana = (iso) => { const g = dataDe(iso).getDay(); return g === 0 || g === 6; };
const somarDias = (iso, n) => { const d = dataDe(iso); d.setDate(d.getDate() + n); return isoDe(d); };

function dataOperacional() {
  const a = new Date();
  if (a.getHours() < 1) a.setDate(a.getDate() - 1);
  return isoDe(a);
}
function turnoAutomatico(iso) {
  const h = new Date().getHours();
  if (fimDeSemana(iso)) return h >= 7 && h < 19 ? "07X19" : "19X07";
  if (h >= 1 && h < 7) return "01X07";
  if (h >= 7 && h < 13) return "07X13";
  if (h >= 13 && h < 19) return "13X19";
  return "19X01";
}
const horaMin = (d) => d.toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit" });
const nomeFrente = (f) => (f ? `${f.nome}${f.terno ? ` · T${f.terno}` : ""}` : "");

/* ================================================================== */
/*  COFRE — os dados compartilhados vão cifrados para o armazenamento  */
/* ================================================================== */
/*
 * Sem isto, qualquer pessoa com o link publicado leria nomes e matrículas.
 * Uma chave aleatória (K) cifra tudo. Cada pessoa autorizada tem essa chave
 * guardada dentro de um envelope que só a senha dela abre, então dá para ter
 * várias senhas diferentes sem servidor nenhum.
 */

const PREFIXO_CIFRA = "enc1:";
const PREFIXOS_DADOS = ["q:", "aj:", "h:", "escala:", "usuarios"];
const te = new TextEncoder(), td = new TextDecoder();
const b64 = (buf) => btoa(String.fromCharCode(...new Uint8Array(buf)));
const deb64 = (t) => Uint8Array.from(atob(t), (c) => c.charCodeAt(0));
const temCripto = () => typeof crypto !== "undefined" && !!crypto.subtle;

let CHAVE = null; /* chave de dados aberta, só na memória desta aba */

async function derivar(senha, salt) {
  const base = await crypto.subtle.importKey("raw", te.encode(senha), "PBKDF2", false, ["deriveKey"]);
  return crypto.subtle.deriveKey(
    { name: "PBKDF2", salt: deb64(salt), iterations: 250000, hash: "SHA-256" },
    base, { name: "AES-GCM", length: 256 }, false, ["encrypt", "decrypt"]);
}
async function cifrarTexto(texto, chave) {
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const buf = await crypto.subtle.encrypt({ name: "AES-GCM", iv }, chave, te.encode(texto));
  return `${PREFIXO_CIFRA}${b64(iv)}.${b64(buf)}`;
}
async function decifrarTexto(txt, chave) {
  const [iv, dado] = txt.slice(PREFIXO_CIFRA.length).split(".");
  const buf = await crypto.subtle.decrypt({ name: "AES-GCM", iv: deb64(iv) }, chave, deb64(dado));
  return td.decode(buf);
}

const K_COFRE = "cofre:meta";
const K_LEMBRADO = "cofre:lembrado";

const lerCofre = async () => {
  try { const r = await window.storage.get(K_COFRE, true); return r ? JSON.parse(r.value) : null; }
  catch { return null; }
};
const gravarCofre = (meta) => window.storage.set(K_COFRE, JSON.stringify(meta), true);

async function envelopar(meta, K, nome, senha) {
  const salt = b64(crypto.getRandomValues(new Uint8Array(16)));
  const kd = await derivar(senha, salt);
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const raw = await crypto.subtle.exportKey("raw", K);
  const env = await crypto.subtle.encrypt({ name: "AES-GCM", iv }, kd, raw);
  meta.usuarios[nome.trim().toLowerCase()] = { nome: nome.trim(), salt, iv: b64(iv), env: b64(env) };
  return meta;
}
async function abrirEnvelope(meta, nome, senha) {
  const u = meta?.usuarios?.[String(nome).trim().toLowerCase()];
  if (!u) return null;
  try {
    const kd = await derivar(senha, u.salt);
    const raw = await crypto.subtle.decrypt({ name: "AES-GCM", iv: deb64(u.iv) }, kd, deb64(u.env));
    return { chave: await crypto.subtle.importKey("raw", raw, { name: "AES-GCM" }, true, ["encrypt", "decrypt"]), nome: u.nome };
  } catch { return null; }
}

/** Cifra o que já estava gravado em texto puro, ao ligar a proteção. */
async function cifrarOQueJaExiste() {
  let n = 0;
  for (const pre of PREFIXOS_DADOS) {
    let r;
    try { r = await window.storage.list(pre, true); } catch { continue; }
    for (const bruta of r?.keys || []) {
      const k = typeof bruta === "string" ? bruta : bruta?.key;
      if (!k || k.startsWith("cofre:")) continue;
      try {
        const item = await window.storage.get(k, true);
        if (!item || String(item.value).startsWith(PREFIXO_CIFRA)) continue;
        await window.storage.set(k, await cifrarTexto(item.value, CHAVE), true);
        n++;
      } catch { /* uma chave ilegível não impede as outras */ }
    }
  }
  return n;
}

async function salvar(chave, valor, compartilhado = true) {
  try {
    const bruto = JSON.stringify(valor);
    const conteudo = CHAVE && compartilhado && !chave.startsWith("cofre:")
      ? await cifrarTexto(bruto, CHAVE) : bruto;
    await window.storage.set(chave, conteudo, compartilhado);
    return true;
  } catch (e) { console.error("salvar", chave, e); return false; }
}
async function carregar(chave, compartilhado = true) {
  try {
    const r = await window.storage.get(chave, compartilhado);
    if (!r) return null;
    let txt = String(r.value);
    if (txt.startsWith(PREFIXO_CIFRA)) {
      if (!CHAVE) return null;
      txt = await decifrarTexto(txt, CHAVE);
    }
    return JSON.parse(txt);
  } catch { return null; }
}

const quadroVazio = () => ({
  alocacoes: {}, extras: [], obs: "", rev: 0,
  intervalos: {}, frentes: FRENTE_PADRAO, locais: {}, fora: {},
});
const kQuadro = (d, t) => `q:${d}:${t}`;
const kAjustes = (d) => `aj:${d}`;
const kHist = (d) => `h:${d}`;

function lerOgmo(texto) {
  const out = [];
  for (const bruta of texto.split(/\r?\n/)) {
    const linha = bruta.trim();
    if (!linha) continue;
    if (/^(função|funcao|legenda|relação|requisi|navio|período)/i.test(linha)) continue;
    let campos = linha.split(/\t+/).map((s) => s.trim()).filter(Boolean);
    if (campos.length < 4) campos = linha.split(/\s{2,}/).map((s) => s.trim()).filter(Boolean);
    if (campos.length < 4) continue;
    const nums = campos.filter((c) => /^\d+$/.test(c));
    const textos = campos.filter((c) => !/^\d+$/.test(c) && !/^(OK|X|NR35)$/i.test(c));
    const nome = textos.length ? textos[textos.length - 1] : "";
    if (!nome || nome.length < 4) continue;
    const ft = textos.length > 1 ? textos[0] : "";
    /* sem função reconhecida assumimos Operador III, que é o caso mais comum no
       pátio; a lista é editável antes de salvar */
    const fn = /\bIV\b/.test(ft) ? "IV" : /\bIII\b/.test(ft) ? "III" : "III";
    out.push({
      id: `ogmo-${nums[nums.length - 1] || nome.replace(/\s/g, "")}`,
      nome: nome.toUpperCase(),
      mat: nums.length >= 3 ? nums[nums.length - 2] : nums[0] || "",
      funcao: fn, funcaoTxt: ft || "TPA", ogmo: true, avulso: true, dias: {},
      inferida: !/\bI(II|V)\b/.test(ft),
    });
  }
  return out;
}

/* ================================================================== */
/*  ESTILOS — identidade JBS Terminais                                 */
/* ================================================================== */

const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Poppins:wght@500;600;700&family=Inter:wght@400;500;600;700&display=swap');
.jbs{
  --bg:#101230; --sur:#191C40; --sur2:#232752; --sur3:#2C3163; --lin:#343A6B;
  --txt:#EDEFF9; --dim:#9BA1CC;
  --navy:#2E2C75; --cyan:#4AC9F1; --blue:#285EAC; --green:#93C83D; --red:#E05A63;
  background:var(--bg); color:var(--txt); font-family:Inter,system-ui,sans-serif;
  min-height:100vh; font-size:14px; -webkit-font-smoothing:antialiased;
}
.jbs *{box-sizing:border-box;}
.jbs h1,.jbs h2,.jbs h3{font-family:Poppins,Inter,sans-serif;font-weight:600;margin:0;}

/* --- barra de contexto (nível 1) --- */
.topo{position:sticky;top:0;z-index:30;}
.n1{display:flex;gap:14px;align-items:center;flex-wrap:wrap;padding:10px 16px;
 background:linear-gradient(180deg,var(--navy) 0%,#252A66 100%);
 border-bottom:1px solid #3A4080;}
.marca{font-family:Poppins;font-weight:700;font-size:16px;letter-spacing:-.01em;
 padding-right:14px;border-right:1px solid rgba(255,255,255,.18);}
.marca span{color:var(--cyan);}
.nav{display:flex;align-items:center;gap:4px;}
.nav .dia{min-width:196px;text-align:center;font-family:Poppins;font-weight:700;font-size:19px;
 line-height:1.15;}
.nav .dia small{display:block;font-family:Inter;font-weight:500;font-size:11px;color:#AEB5E6;
 letter-spacing:.01em;}
.relogio{font-family:Poppins,sans-serif;font-weight:700;font-size:26px;line-height:1;
 font-variant-numeric:tabular-nums;}

/* --- barra de ações (nível 2): passos numerados --- */
.n2{display:flex;gap:8px;align-items:center;flex-wrap:wrap;padding:10px 16px;
 background:var(--sur);border-bottom:2px solid var(--navy);}
.sep{width:1px;height:24px;background:var(--lin);margin:0 5px;}
.passos{display:flex;gap:5px;align-items:center;flex-wrap:wrap;}
.passo{display:flex;align-items:center;gap:8px;background:var(--sur2);border:1px solid var(--lin);
 border-radius:8px;padding:6px 13px 6px 7px;font-family:inherit;font-size:13px;font-weight:500;
 color:var(--txt);cursor:pointer;transition:background .12s,border-color .12s;}
.passo:hover{background:var(--sur3);border-color:#4A5192;}
.passo:focus-visible{outline:2px solid var(--cyan);outline-offset:2px;}
.passo .num{width:21px;height:21px;border-radius:50%;background:var(--lin);color:#AEB4DC;
 display:grid;place-items:center;font-size:11px;font-weight:700;flex:none;}
.passo.ok{background:#1D4327;border-color:var(--green);color:#DEF2CE;}
.passo.ok .num{background:var(--green);color:#12132B;}
.passo .qtd{font-size:11px;color:inherit;opacity:.75;}
.seta{color:#5B6199;font-size:15px;user-select:none;}

/* --- cabeçalhos de seção da lateral --- */
.sech{display:flex;align-items:baseline;gap:7px;margin:16px 0 6px;padding-bottom:5px;
 border-bottom:1px solid var(--lin);}
.sech b{font-family:Poppins,sans-serif;font-size:13px;font-weight:700;letter-spacing:.02em;}
.sech .sub{font-size:11px;color:var(--dim);}
.sech .qt2{margin-left:auto;font-size:11px;color:var(--dim);}

.b{background:var(--sur2);color:var(--txt);border:1px solid var(--lin);border-radius:7px;
 padding:7px 13px;font-size:13px;font-weight:500;cursor:pointer;font-family:inherit;
 transition:background .12s,border-color .12s;}
.b:hover:not(:disabled){background:var(--sur3);border-color:#4A5192;}
.b:focus-visible{outline:2px solid var(--cyan);outline-offset:2px;}
.b:disabled{opacity:.4;cursor:default;}
.b.pri{background:var(--green);border-color:var(--green);color:#12132B;font-weight:600;}
.b.pri:hover:not(:disabled){background:#A4D653;border-color:#A4D653;}
.b.perigo{background:#4A2230;border-color:#7A3644;color:#F2B8BF;}
.b.ghost{background:transparent;border-color:transparent;color:var(--dim);}
.b.ghost:hover{background:rgba(255,255,255,.06);color:var(--txt);}
.b.sm{padding:4px 9px;font-size:12px;border-radius:6px;}
.b.on{background:var(--cyan);border-color:var(--cyan);color:#0C1A22;font-weight:600;}

.i{background:#0D0F28;color:var(--txt);border:1px solid var(--lin);border-radius:7px;
 padding:7px 10px;font-size:13px;font-family:inherit;}
.i:focus{outline:2px solid var(--cyan);outline-offset:-1px;}
.i::placeholder{color:#6F76A6;}

.seg{display:flex;background:rgba(0,0,0,.22);border:1px solid rgba(255,255,255,.14);
 border-radius:8px;overflow:hidden;}
.seg button{background:transparent;border:0;color:#C3C8EA;padding:6px 14px;cursor:pointer;
 font-family:Poppins,sans-serif;font-size:14px;font-weight:600;letter-spacing:.02em;}
.seg button + button{border-left:1px solid rgba(255,255,255,.12);}
.seg button:hover{background:rgba(255,255,255,.08);}
.seg button[aria-pressed="true"]{background:var(--cyan);color:#0C1A22;}
.seg.doze button[aria-pressed="true"]{background:var(--green);color:#12132B;}

/* --- layout --- */
.wrap{display:grid;grid-template-columns:312px 1fr;align-items:start;}
@media(max-width:920px){.wrap{grid-template-columns:1fr;}}
.lat{background:var(--sur);border-right:1px solid var(--lin);padding:12px;min-height:74vh;
 position:sticky;top:var(--topo,96px);max-height:calc(100vh - var(--topo,96px));overflow:auto;}
@media(max-width:920px){.lat{position:static;max-height:none;}}
.main{padding:14px 16px 96px;}

/* --- grupos e cartões --- */
.grp{margin-bottom:26px;}
.grph{display:flex;align-items:center;gap:10px;margin-bottom:10px;padding:6px 0 7px 11px;
 border-left:4px solid;border-bottom:1px solid var(--lin);}
.grph h3{font-size:16px;letter-spacing:.02em;font-weight:700;}
.grph .cnt{font-size:11.5px;color:var(--dim);background:var(--sur);border-radius:99px;padding:2px 9px;}
.barra{flex:1;height:1px;background:var(--lin);}
.slots{display:grid;gap:7px;grid-template-columns:repeat(auto-fill,minmax(172px,1fr));}

.card{position:relative;background:var(--sur);border:1px solid var(--lin);border-radius:8px;
 padding:8px 10px 7px;cursor:pointer;transition:background .12s,box-shadow .12s;
 border-top:3px solid var(--lin);}
.card:hover{background:var(--sur2);}
.card:focus-visible{outline:2px solid var(--cyan);outline-offset:1px;}
.card.sel{box-shadow:0 0 0 2px var(--cyan);background:var(--sur2);}
.card.parado{background:#2A1626;border-color:#5E2A3C;}
.card .eqn{font-family:Poppins,sans-serif;font-weight:600;font-size:13px;letter-spacing:.03em;
 padding-right:22px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.card .opn{font-size:13.5px;font-weight:600;line-height:1.25;margin-top:3px;
 overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
.card.vazio .opn{color:#7A81BA;font-weight:400;}
.card.parado .opn{color:#F0A0AC;}
.meta{display:flex;gap:8px;align-items:center;margin-top:4px;font-size:11px;min-height:15px;}
.meta .int{color:var(--cyan);font-weight:600;font-variant-numeric:tabular-nums;}
.chk{position:absolute;top:6px;right:6px;width:17px;height:17px;border-radius:5px;
 border:1px solid #4A5192;background:rgba(0,0,0,.3);cursor:pointer;display:flex;
 align-items:center;justify-content:center;padding:0;color:#0C1A22;font-size:11px;font-weight:700;}
.chk:hover{border-color:var(--cyan);}
.chk[data-on="1"]{background:var(--cyan);border-color:var(--cyan);}

/* --- lista de operadores --- */
.oc{display:flex;align-items:center;gap:8px;padding:6px 8px;border-radius:7px;width:100%;
 text-align:left;background:transparent;color:inherit;font-family:inherit;border:1px solid transparent;}
.oc:hover{background:var(--sur2);}
.oc[data-sel="1"]{background:var(--sur3);border-color:var(--cyan);}
.dot{width:7px;height:7px;border-radius:50%;flex:none;}
.tag{font-size:10.5px;padding:2px 7px;border-radius:99px;font-weight:600;white-space:nowrap;}

/* --- barra flutuante de seleção --- */
.flut{position:fixed;left:50%;transform:translateX(-50%);bottom:16px;z-index:40;
 background:var(--sur2);border:1px solid #4A5192;border-radius:12px;padding:10px 13px;
 box-shadow:0 12px 32px rgba(0,0,0,.55);display:flex;gap:8px;align-items:center;
 flex-wrap:wrap;max-width:min(940px,94vw);}
.flut .cont{font-family:Poppins;font-weight:600;font-size:14px;color:var(--cyan);white-space:nowrap;}

/* --- modais --- */
.modal{position:fixed;inset:0;background:rgba(8,10,26,.82);display:flex;align-items:center;
 justify-content:center;padding:18px;z-index:60;}
.dlg{background:var(--sur);border:1px solid var(--lin);border-radius:12px;max-width:600px;
 width:100%;max-height:88vh;overflow:auto;padding:20px;}
.dlgh{display:flex;align-items:flex-start;gap:12px;margin-bottom:4px;}
.dlgh h3{font-size:20px;flex:1;}
.campo{margin-top:16px;}
.campo > label{display:block;font-size:11.5px;color:var(--dim);margin-bottom:5px;font-weight:500;}

/* --- arrastar e soltar --- */
.card[draggable="true"]{cursor:grab;}
.card[draggable="true"]:active{cursor:grabbing;}
.oc[draggable="true"]{cursor:grab;}
.card[data-alvo="1"]{box-shadow:0 0 0 2px var(--green);background:var(--sur2);}
.fantasma{opacity:.4;}

/* --- doca das frentes de trabalho --- */
.dock{display:flex;gap:6px;flex-wrap:wrap;align-items:center;margin-bottom:20px;padding:10px 11px;
 background:var(--sur);border:1px solid var(--lin);border-left:4px solid #C7A2F0;border-radius:9px;}
.dock .rot{font-size:11.5px;color:var(--dim);margin-right:2px;}
.dz{display:flex;align-items:center;gap:7px;padding:6px 11px;border-radius:7px;
 border:1px dashed var(--lin);font-size:12.5px;cursor:pointer;background:transparent;
 color:inherit;font-family:inherit;transition:background .12s,border-color .12s;}
.dz:hover{background:var(--sur2);}
.dz:focus-visible{outline:2px solid var(--cyan);outline-offset:2px;}
.dz[data-alvo="1"]{border-style:solid;background:var(--sur3);}
.dz .qt{font-size:11px;font-weight:700;background:rgba(0,0,0,.3);border-radius:99px;padding:1px 7px;}

.dim{color:var(--dim);} .hint{font-size:12px;color:var(--dim);line-height:1.55;}
.row{display:flex;gap:8px;align-items:center;}
.chips{display:flex;gap:6px;flex-wrap:wrap;}

/* --- TV --- */
.tv{background:#0B0D24;min-height:100vh;padding:18px 22px;}
.tvh{display:flex;justify-content:space-between;align-items:flex-end;gap:16px;
 padding-bottom:12px;border-bottom:2px solid var(--navy);margin-bottom:16px;}
.tvcols{display:grid;gap:14px 22px;grid-template-columns:repeat(auto-fill,minmax(430px,1fr));
 align-items:start;}
.tvbloco{background:#141838;border:1px solid #262B5C;border-radius:10px;padding:10px 13px 6px;}
.tvbh{display:flex;align-items:baseline;gap:10px;padding-bottom:7px;margin-bottom:4px;
 border-bottom:2px solid;}
.tvbh h3{font-size:21px;}
.tvl{display:flex;align-items:baseline;gap:12px;padding:5px 0;border-bottom:1px solid #21264F;}
.tvl:last-child{border-bottom:0;}
.tvl-eq{width:118px;flex:none;font-family:Poppins,sans-serif;font-size:16px;font-weight:600;
 letter-spacing:.03em;}
.tvl-op{flex:1;min-width:0;font-size:20px;font-weight:600;overflow:hidden;text-overflow:ellipsis;
 white-space:nowrap;}
.tvl-fr{font-size:15px;white-space:nowrap;}
.tvl-hr{width:52px;flex:none;text-align:right;font-size:17px;font-weight:600;color:var(--cyan);
 font-variant-numeric:tabular-nums;}
@media (prefers-reduced-motion: reduce){.jbs *{transition:none !important;}}
`;

/* ================================================================== */
/*  APP                                                                */
/* ================================================================== */

export default function EscalaPatio() {
  const [usuario, setUsuario] = useState(null);
  const [cofre, setCofre] = useState(null);
  const [modo, setModo] = useState("operacao");
  const [operadores, setOperadores] = useState([]);
  const [equipamentos] = useState(equipamentosPadrao);
  const [data, setData] = useState(dataOperacional);
  const [turno, setTurno] = useState(() => turnoAutomatico(dataOperacional()));
  const [quadro, setQuadro] = useState(quadroVazio);
  const [ajustes, setAjustes] = useState({});
  const [historico, setHistorico] = useState([]);
  const [carregando, setCarregando] = useState(true);
  const [opSelecionado, setOpSelecionado] = useState(null);
  const [selecao, setSelecao] = useState([]);
  const [ficha, setFicha] = useState(null);
  const [painel, setPainel] = useState(null);
  const [buscaOp, setBuscaOp] = useState("");
  const [aviso, setAviso] = useState(null);
  const [relogio, setRelogio] = useState(new Date());
  const [tvPorFrente, setTvPorFrente] = useState(true);
  const [arrasto, setArrasto] = useState(null);
  const [alvo, setAlvo] = useState(null);
  const inputRef = useRef(null);
  const topoRef = useRef(null);
  const filaRef = useRef(Promise.resolve());
  const revRef = useRef(0);

  /* a lateral gruda logo abaixo do topo, que muda de altura ao quebrar linha */
  useEffect(() => {
    const el = topoRef.current;
    if (!el || typeof ResizeObserver === "undefined") return;
    const medir = () =>
      el.closest(".jbs")?.style.setProperty("--topo", `${el.offsetHeight}px`);
    medir();
    const ro = new ResizeObserver(medir);
    ro.observe(el);
    return () => ro.disconnect();
  }, [modo, usuario, operadores.length]);

  /* ---------- sessão e cofre ---------- */
  useEffect(() => {
    (async () => {
      const meta = await lerCofre();
      setCofre(meta);

      if (meta) {
        /* chave guardada neste computador dispensa digitar a senha de novo */
        try {
          const lembrado = await window.storage.get(K_LEMBRADO, false);
          if (lembrado) {
            const { raw, nome } = JSON.parse(lembrado.value);
            CHAVE = await crypto.subtle.importKey("raw", deb64(raw), { name: "AES-GCM" }, true,
              ["encrypt", "decrypt"]);
            setUsuario({ nome, desde: Date.now() });
          }
        } catch { CHAVE = null; }
      } else {
        const u = await carregar("sessao", false);
        if (u) setUsuario(u);
      }

      if (!meta || CHAVE) {
        const ops = await carregar("escala:operadores");
        if (ops) setOperadores(ops);
      }
      setCarregando(false);
    })();
  }, []);

  /** Entrada sem proteção: só o nome, para o histórico. */
  async function entrar(nome) {
    const u = { nome: nome.trim(), desde: Date.now() };
    setUsuario(u);
    await salvar("sessao", u, false);
    const lista = (await carregar("usuarios")) || [];
    if (!lista.some((x) => x.toLowerCase() === u.nome.toLowerCase()))
      await salvar("usuarios", [...lista, u.nome]);
  }

  /** Entrada protegida: a senha abre o envelope e libera a chave dos dados. */
  async function destrancar(nome, senha, lembrar) {
    const r = await abrirEnvelope(cofre, nome, senha);
    if (!r) return false;
    CHAVE = r.chave;
    if (lembrar) {
      const raw = b64(await crypto.subtle.exportKey("raw", r.chave));
      await window.storage.set(K_LEMBRADO, JSON.stringify({ raw, nome: r.nome }), false);
    }
    setUsuario({ nome: r.nome, desde: Date.now() });
    const ops = await carregar("escala:operadores");
    setOperadores(ops || []);
    return true;
  }

  async function sair() {
    CHAVE = null;
    setUsuario(null);
    setOperadores([]);
    setQuadro(quadroVazio());
    try { await window.storage.delete(K_LEMBRADO, false); } catch { /* pode não existir */ }
    if (!cofre) await salvar("sessao", null, false);
  }

  /** Liga a proteção: cria a chave, envelopa para o primeiro usuário e cifra o que já existe. */
  async function protegerComSenha(nome, senha) {
    if (!temCripto()) return { erro: "Este navegador não oferece as funções de criptografia necessárias." };
    const K = await crypto.subtle.generateKey({ name: "AES-GCM", length: 256 }, true, ["encrypt", "decrypt"]);
    const meta = await envelopar({ versao: 1, usuarios: {} }, K, nome, senha);
    CHAVE = K;
    const n = await cifrarOQueJaExiste();
    await gravarCofre(meta);
    setCofre(meta);
    setUsuario({ nome: nome.trim(), desde: Date.now() });
    registrar(`Ativou a proteção por senha (${n} registros cifrados)`);
    return { ok: `Proteção ativada. ${n} registros já existentes foram cifrados.` };
  }

  async function adicionarAcesso(nome, senha) {
    if (!CHAVE || !cofre) return { erro: "Entre com a sua senha primeiro." };
    const meta = await envelopar({ ...cofre, usuarios: { ...cofre.usuarios } }, CHAVE, nome, senha);
    await gravarCofre(meta);
    setCofre(meta);
    registrar(`Liberou acesso para ${nome.trim()}`);
    return { ok: `${nome.trim()} já pode entrar com a senha dele.` };
  }

  async function removerAcesso(chaveNome) {
    const usuarios = { ...cofre.usuarios };
    const nome = usuarios[chaveNome]?.nome;
    delete usuarios[chaveNome];
    const meta = { ...cofre, usuarios };
    await gravarCofre(meta);
    setCofre(meta);
    registrar(`Removeu o acesso de ${nome}`);
  }

  /* ---------- sincronização ---------- */
  const puxar = useCallback(async () => {
    const [q, a] = await Promise.all([carregar(kQuadro(data, turno)), carregar(kAjustes(data))]);
    if (q && (q.rev || 0) >= revRef.current) {
      revRef.current = q.rev || 0;
      setQuadro({ ...quadroVazio(), ...q, frentes: q.frentes?.length ? q.frentes : FRENTE_PADRAO });
    } else if (!q && revRef.current === 0) setQuadro(quadroVazio());
    setAjustes(a || {});
  }, [data, turno]);

  useEffect(() => {
    revRef.current = 0;
    setQuadro(quadroVazio());
    setSelecao([]);
    puxar();
  }, [data, turno, puxar]);

  useEffect(() => {
    const ms = modo === "tv" ? 4000 : 8000;
    const t = setInterval(() => { setRelogio(new Date()); puxar(); }, ms);
    return () => clearInterval(t);
  }, [modo, puxar]);

  useEffect(() => {
    if (modo !== "tv") return;
    const d = dataOperacional();
    const t = turnoAutomatico(d);
    if (d !== data) setData(d);
    if (t !== turno) setTurno(t);
  }, [modo, relogio]); // eslint-disable-line

  useEffect(() => {
    const esc = (e) => {
      if (e.key !== "Escape") return;
      if (modo === "tv") { setModo("operacao"); return; }
      if (ficha || painel) { setFicha(null); setPainel(null); }
      else if (selecao.length) setSelecao([]);
      else if (opSelecionado) setOpSelecionado(null);
    };
    window.addEventListener("keydown", esc);
    return () => window.removeEventListener("keydown", esc);
  }, [modo, ficha, painel, selecao.length, opSelecionado]);

  function registrar(acao) {
    filaRef.current = filaRef.current.then(async () => {
      const h = (await carregar(kHist(data))) || [];
      h.push({ ts: Date.now(), por: usuario?.nome || "?", turno, acao });
      await salvar(kHist(data), h.slice(-300));
    });
  }

  function gravarQuadro(novo, acao) {
    const comRev = { ...novo, rev: Date.now(), por: usuario?.nome || "?" };
    revRef.current = comRev.rev;
    setQuadro(comRev);
    filaRef.current = filaRef.current
      .then(() => salvar(kQuadro(data, turno), comRev))
      .then((ok) => { if (!ok) setAviso("Não consegui gravar. Refaça a alteração."); });
    if (acao) registrar(acao);
  }

  function gravarAjuste(opId, status, motivo) {
    const novo = { ...ajustes };
    if (!status) delete novo[opId];
    else novo[opId] = { status, motivo: motivo || "", por: usuario?.nome || "?", em: Date.now() };
    setAjustes(novo);
    filaRef.current = filaRef.current.then(() => salvar(kAjustes(data), novo));
    const op = operadores.find((o) => o.id === opId) || (quadro.extras || []).find((o) => o.id === opId);
    registrar(status
      ? `${op?.nome || opId} marcado como ${STATUS[status]?.rot}${motivo ? ` (${motivo})` : ""}`
      : `${op?.nome || opId} voltou à marcação da planilha`);
    if (!status || status !== "TRABALHA") {
      const a = { ...(quadro.alocacoes || {}) };
      let mudou = false;
      for (const k of Object.keys(a)) if (a[k] === opId) { delete a[k]; mudou = true; }
      if (mudou) gravarQuadro({ ...quadro, alocacoes: a });
    }
  }

  async function importar(arquivos) {
    const novos = [], erros = [];
    for (const arq of arquivos) {
      try {
        const { operadores: ops } = lerPlanilha(new Uint8Array(await arq.arrayBuffer()), arq.name);
        if (!ops.length) erros.push(`${arq.name}: nenhum operador encontrado.`);
        novos.push(...ops);
      } catch (e) { erros.push(`${arq.name}: ${e.message}`); }
    }
    if (!novos.length) { setAviso(erros.join(" ") || "Nada foi lido."); return; }
    const mapa = new Map(operadores.map((o) => [o.id, o]));
    novos.forEach((o) => {
      const ja = mapa.get(o.id);
      if (ja && ja.origem === o.origem &&
          Object.keys(ja.dias || {}).length > Object.keys(o.dias || {}).length) return;
      mapa.set(o.id, o);
    });
    const lista = [...mapa.values()];
    setOperadores(lista);
    await salvar("escala:operadores", lista);
    registrar(`Importou ${novos.length} operadores`);
    setAviso(`${novos.length} operadores importados.${erros.length ? " " + erros.join(" ") : ""}`);
  }

  /* ---------- derivações ---------- */
  const extras = quadro.extras || [];
  const dataAnterior = useMemo(() => somarDias(data, -1), [data]);

  const elenco = useMemo(() => {
    const f = (o) => ({ ...o, ...situacao(o, data, turno, ajustes, dataAnterior) });
    return [...operadores.map(f), ...extras.map(f)];
  }, [operadores, extras, data, turno, ajustes, dataAnterior]);

  const disponiveis = useMemo(
    () => elenco.filter((o) => o.disponivel).sort((a, b) => a.nome.localeCompare(b.nome, "pt-BR")),
    [elenco]);
  const foraDaOperacao = useMemo(
    () => elenco.filter((o) => !o.disponivel && (o.turno === turno || o.ajuste || (eh12h(turno) && o.codigo)))
      .sort((a, b) => a.nome.localeCompare(b.nome, "pt-BR")),
    [elenco, turno]);

  const porId = useMemo(() => { const m = new Map(); elenco.forEach((o) => m.set(o.id, o)); return m; }, [elenco]);
  const usados = useMemo(() => new Set(Object.values(quadro.alocacoes || {})), [quadro.alocacoes]);
  const frentes = quadro.frentes || FRENTE_PADRAO;
  const frentePorId = useMemo(() => { const m = new Map(); frentes.forEach((f) => m.set(f.id, f)); return m; }, [frentes]);

  const porFamilia = useMemo(() => {
    const g = {};
    ORDEM_FAM.forEach((f) => (g[f] = []));
    equipamentos.forEach((e) => g[e.fam].push(e));
    return g;
  }, [equipamentos]);

  const listaLateral = useMemo(() => {
    const q = buscaOp.trim().toLowerCase();
    if (!q) return disponiveis;
    return disponiveis.filter((o) => o.nome.toLowerCase().includes(q) || String(o.mat).includes(q));
  }, [disponiveis, buscaOp]);

  /* ---------- ações do quadro ---------- */
  function alocar(slotId, opId) {
    if (quadro.fora?.[slotId]) { setAviso("Este equipamento está fora de operação."); return; }
    const a = { ...(quadro.alocacoes || {}) };
    const eq = equipamentos.find((e) => e.id === slotId);
    const cand = opId ? porId.get(opId) : null;
    if (cand && !podeOperar(cand, eq)) {
      setAviso(`${cand.nome} é ${FUNCOES[funcaoDe(cand)].nome} e opera ${equipamentosDaFuncao(funcaoDe(cand))}. Não pode ir para ${eq.nome}.`);
      setOpSelecionado(null);
      return;
    }
    for (const k of Object.keys(a)) if (a[k] === opId) delete a[k];
    if (opId) a[slotId] = opId; else delete a[slotId];
    gravarQuadro({ ...quadro, alocacoes: a },
      opId ? `${eq?.nome}: ${porId.get(opId)?.nome}` : `${eq?.nome}: liberado`);
    setOpSelecionado(null);
  }

  function definirIntervalo(slotId, hora) {
    const i = { ...(quadro.intervalos || {}) };
    if (hora) i[slotId] = hora; else delete i[slotId];
    const eq = equipamentos.find((e) => e.id === slotId);
    gravarQuadro({ ...quadro, intervalos: i }, `${eq?.nome}: intervalo ${hora || "removido"}`);
  }

  /** Aplica uma frente a um conjunto de equipamentos de uma vez. */
  function definirLocalEmLote(ids, frenteId) {
    const l = { ...(quadro.locais || {}) };
    ids.forEach((id) => { if (frenteId) l[id] = frenteId; else delete l[id]; });
    const fr = frentePorId.get(frenteId);
    const nomes = ids.map((id) => equipamentos.find((e) => e.id === id)?.nome).join(", ");
    gravarQuadro({ ...quadro, locais: l },
      frenteId ? `${nomeFrente(fr)}: ${nomes}` : `Sem frente: ${nomes}`);
  }

  function definirParada(slotId, motivo) {
    const f = { ...(quadro.fora || {}) };
    const a = { ...(quadro.alocacoes || {}) };
    const eq = equipamentos.find((e) => e.id === slotId);
    if (motivo) { f[slotId] = { motivo, por: usuario?.nome || "?", em: Date.now() }; delete a[slotId]; }
    else delete f[slotId];
    gravarQuadro({ ...quadro, fora: f, alocacoes: a },
      motivo ? `${eq?.nome} fora de operação: ${motivo}` : `${eq?.nome} voltou à operação`);
  }

  /**
   * Escalona os intervalos. Cada rodada avança pela maior duração do grupo,
   * porque o TPA do OGMO tem 30 minutos e o pessoal próprio 15 ou 60.
   */
  function preencherIntervalos({ familias, inicio, duracao, duracaoOgmo, porVez, sobrescrever }) {
    const i = { ...(quadro.intervalos || {}) };
    const [h, m] = inicio.split(":").map(Number);
    let minutos = h * 60 + m, n = 0;
    let grupo = [];
    equipamentos
      .filter((e) => familias.includes(e.fam))
      .filter((e) => quadro.alocacoes?.[e.id] && !quadro.fora?.[e.id])
      .filter((e) => sobrescrever || !i[e.id])
      .forEach((e) => {
        i[e.id] = `${String(Math.floor((minutos / 60) % 24)).padStart(2, "0")}:${String(minutos % 60).padStart(2, "0")}`;
        n++;
        grupo.push(porId.get(quadro.alocacoes[e.id])?.ogmo ? duracaoOgmo : duracao);
        if (grupo.length >= porVez) { minutos += Math.max(...grupo); grupo = []; }
      });
    if (!n) { setAviso("Nenhum equipamento com operador nessa seleção."); return; }
    gravarQuadro({ ...quadro, intervalos: i },
      `Preencheu ${n} intervalos a partir de ${inicio} (${duracao} min, OGMO ${duracaoOgmo} min, ${porVez} por vez)`);
    setAviso(`${n} intervalos preenchidos.`);
  }

  function montarRS() {
    const a = { ...(quadro.alocacoes || {}) };
    const rsSlots = equipamentos.filter((e) => e.fam === "RS" && !quadro.fora?.[e.id]);
    let fixos = 0, ogmo = 0;
    disponiveis.filter((o) => o.funcao === "III" && o.rs && !Object.values(a).includes(o.id))
      .forEach((o) => {
        const alvo = `RS${parseInt(o.rs, 10)}`;
        if (rsSlots.some((e) => e.id === alvo) && !a[alvo]) { a[alvo] = o.id; fixos++; }
      });
    disponiveis.filter((o) => o.ogmo && o.funcao === "III" && !Object.values(a).includes(o.id))
      .forEach((o) => {
        const alvo = rsSlots.find((e) => !a[e.id]);
        if (alvo) { a[alvo.id] = o.id; ogmo++; }
      });
    gravarQuadro({ ...quadro, alocacoes: a }, `Montou os RS (${fixos} da escala, ${ogmo} do OGMO)`);
    setAviso(fixos + ogmo === 0
      ? "Os RS já estão preenchidos ou não há Operador III disponível."
      : `${fixos} na máquina de referência, ${ogmo} do OGMO nos RS restantes.`);
  }

  /* ---------- arrastar e soltar ---------- */
  /* Um único tipo MIME com JSON dentro é o que todos os navegadores aceitam. */
  function iniciarArrasto(ev, carga) {
    ev.dataTransfer.effectAllowed = "move";
    ev.dataTransfer.setData("text/plain", JSON.stringify(carga));
    setArrasto(carga);
  }
  const encerrarArrasto = () => { setArrasto(null); setAlvo(null); };

  function lerCarga(ev) {
    try { return JSON.parse(ev.dataTransfer.getData("text/plain")); } catch { return null; }
  }

  /** Troca os operadores de dois equipamentos, ou apenas move quando o destino está livre. */
  function trocarOperadores(origem, destino) {
    if (origem === destino) return;
    const a = { ...(quadro.alocacoes || {}) };
    const oa = a[origem], ob = a[destino];
    const eqA = equipamentos.find((e) => e.id === origem);
    const eqB = equipamentos.find((e) => e.id === destino);
    if ((oa && !podeOperar(porId.get(oa), eqB)) || (ob && !podeOperar(porId.get(ob), eqA))) {
      setAviso(`${eqA.nome} e ${eqB.nome} são de tipos diferentes: os operadores não se substituem.`);
      return;
    }
    if (ob) a[origem] = ob; else delete a[origem];
    if (oa) a[destino] = oa; else delete a[destino];
    const nA = equipamentos.find((e) => e.id === origem)?.nome;
    const nB = equipamentos.find((e) => e.id === destino)?.nome;
    gravarQuadro({ ...quadro, alocacoes: a },
      ob ? `Trocou os operadores de ${nA} e ${nB}` : `${porId.get(oa)?.nome}: ${nA} → ${nB}`);
  }

  function soltarNoEquipamento(ev, slotId) {
    ev.preventDefault();
    const carga = lerCarga(ev);
    encerrarArrasto();
    if (!carga) return;
    if (carga.tipo === "frente") return definirLocalEmLote([slotId], carga.id || null);
    if (quadro.fora?.[slotId]) { setAviso("Este equipamento está fora de operação."); return; }
    if (carga.tipo === "op") return alocar(slotId, carga.id);
    if (carga.tipo === "equips" && carga.ids.length === 1) return trocarOperadores(carga.ids[0], slotId);
    if (carga.tipo === "equips") setAviso("Solte um equipamento por vez para trocar operadores.");
  }

  function soltarNaFrente(ev, frenteId) {
    ev.preventDefault();
    const carga = lerCarga(ev);
    encerrarArrasto();
    if (!carga) return;
    if (carga.tipo !== "equips") { setAviso("Arraste equipamentos para a frente, não operadores."); return; }
    definirLocalEmLote(carga.ids, frenteId);
    setSelecao([]);
  }

  /** Soltar um navio sobre um operador manda a frente para o equipamento dele. */
  function soltarNoOperador(ev, opId) {
    ev.preventDefault();
    const carga = lerCarga(ev);
    encerrarArrasto();
    if (!carga || carga.tipo !== "frente") return;
    const slot = Object.keys(quadro.alocacoes || {}).find((k) => quadro.alocacoes[k] === opId);
    if (!slot) {
      setAviso(`${porId.get(opId)?.nome} ainda não está em nenhum equipamento. Aloque primeiro.`);
      return;
    }
    definirLocalEmLote([slot], carga.id || null);
  }

  /** Quando o cartão faz parte da seleção, arrastar leva a seleção inteira. */
  const idsDoArrasto = (id) => (selecao.includes(id) && selecao.length > 1 ? selecao : [id]);

  /* ---------- zerar ---------- */
  async function zerar(opcoes) {
    const feito = [];
    if (opcoes.quadro) {
      gravarQuadro({ ...quadroVazio(), frentes: opcoes.frentes ? FRENTE_PADRAO : quadro.frentes },
        "Zerou o quadro do turno");
      feito.push("quadro do turno");
    }
    if (opcoes.ajustes) {
      setAjustes({});
      await salvar(kAjustes(data), {});
      feito.push("ajustes do dia");
    }
    if (opcoes.escala) {
      setOperadores([]);
      await salvar("escala:operadores", []);
      feito.push("escala importada");
    }
    if (opcoes.historico) {
      await salvar(kHist(data), []);
      setHistorico([]);
      feito.push("histórico do dia");
    }
    setSelecao([]);
    setOpSelecionado(null);
    if (!opcoes.historico) registrar(`Zerou: ${feito.join(", ")}`);
    setAviso(`Zerado: ${feito.join(", ")}.`);
  }

  function copiarResumo() {
    const l = [`ESCALA PÁTIO — ${rotuloData(data)} — Turno ${rotuloTurno(turno)}`, ""];
    ORDEM_FAM.forEach((fam) => {
      const p = porFamilia[fam].filter((e) => quadro.alocacoes?.[e.id] || quadro.fora?.[e.id]);
      if (!p.length) return;
      l.push(FAMILIAS[fam].nome);
      p.forEach((e) => {
        const parado = quadro.fora?.[e.id];
        if (parado) { l.push(`  ${e.nome}: FORA — ${parado.motivo}`); return; }
        const hora = quadro.intervalos?.[e.id];
        const fr = frentePorId.get(quadro.locais?.[e.id]);
        const extra = [fr ? nomeFrente(fr) : null, hora ? `int. ${hora}` : null].filter(Boolean).join(" · ");
        l.push(`  ${e.nome}: ${porId.get(quadro.alocacoes[e.id])?.nome || "?"}${extra ? ` (${extra})` : ""}`);
      });
      l.push("");
    });
    if (quadro.obs) l.push(`Obs: ${quadro.obs}`);
    navigator.clipboard?.writeText(l.join("\n")).then(
      () => setAviso("Resumo copiado."),
      () => setAviso("Não consegui copiar automaticamente."));
  }

  async function abrirHistorico() {
    setHistorico(((await carregar(kHist(data))) || []).slice().reverse());
    setPainel("historico");
  }

  /** Carrega os quadros dos seis turnos do dia para a aba de escalados. */
  const carregarDia = useCallback(async (iso) => {
    const qs = await Promise.all(TODOS_TURNOS.map((t) => carregar(kQuadro(iso, t.id))));
    return TODOS_TURNOS.map((t, i) => ({ turno: t.id, quadro: qs[i] })).filter((x) => x.quadro);
  }, []);

  /* ---------- seleção múltipla ---------- */
  const alternarSelecao = (id) =>
    setSelecao((s) => (s.includes(id) ? s.filter((x) => x !== id) : [...s, id]));
  const selecionarFamilia = (fam) => {
    const ids = porFamilia[fam].map((e) => e.id);
    const todos = ids.every((i) => selecao.includes(i));
    setSelecao((s) => (todos ? s.filter((x) => !ids.includes(x)) : [...new Set([...s, ...ids])]));
  };

  /* ================================================================ */

  if (carregando)
    return <div className="jbs"><style>{CSS}</style><div style={{ padding: 40 }} className="dim">Carregando…</div></div>;

  if (!usuario)
    return cofre
      ? <TelaSenha cofre={cofre} onDestrancar={destrancar} />
      : <TelaEntrada onEntrar={entrar} />;

  if (modo === "tv")
    return (
      <TelaTV
        {...{ quadro, equipamentos, porFamilia, porId, frentes, frentePorId, data, turno, relogio, tvPorFrente }}
        onAlternarVista={() => setTvPorFrente((v) => !v)}
        onSair={() => setModo("operacao")}
      />
    );

  const fds = fimDeSemana(data);
  const preenchidos = Object.keys(quadro.alocacoes || {}).length;
  const parados = Object.keys(quadro.fora || {}).length;
  const comIntervalo = Object.keys(quadro.intervalos || {}).length;
  const comFrente = Object.keys(quadro.locais || {}).length;

  /* Roteiro do turno. Cada passo fica verde assim que aquilo já está no quadro. */
  const passos = [
    { n: 1, rot: "Importar planilha", feito: operadores.length > 0,
      qtd: operadores.length ? `${operadores.length}` : "", acao: () => inputRef.current?.click() },
    { n: 2, rot: "TPAs do OGMO", feito: extras.length > 0,
      qtd: extras.length ? `${extras.length}` : "", acao: () => setPainel("ogmo") },
    { n: 3, rot: "Montar", feito: preenchidos > 0,
      qtd: preenchidos ? `${preenchidos}` : "", acao: montarRS },
    { n: 4, rot: "Intervalo", feito: comIntervalo > 0,
      qtd: comIntervalo ? `${comIntervalo}` : "", acao: () => setPainel("intervalos") },
    { n: 5, rot: "Frente", feito: comFrente > 0,
      qtd: comFrente ? `${comFrente}` : "", acao: () => setPainel("frentes") },
  ];

  return (
    <div className="jbs"><style>{CSS}</style>

      <div className="topo" ref={topoRef}>
      {/* ---------- nível 1: contexto ---------- */}
      <div className="n1">
        <div className="marca">JBS <span>Terminais</span></div>

        <div className="nav">
          <button className="b ghost" title="Dia anterior" onClick={() => setData(somarDias(data, -1))}>‹</button>
          <div className="dia">
            {rotuloData(data)}
            <small>{fds ? "fim de semana · turnos de 12h" : "dia útil · turnos de 6h"}</small>
          </div>
          <button className="b ghost" title="Próximo dia" onClick={() => setData(somarDias(data, 1))}>›</button>
        </div>

        <input className="i" type="date" value={data} onChange={(e) => setData(e.target.value)} aria-label="Data" />

        <div className="seg" role="group" aria-label="Turnos de 6 horas">
          {TURNOS_6H.map((t) => (
            <button key={t.id} aria-pressed={turno === t.id} onClick={() => setTurno(t.id)}>{t.label}</button>
          ))}
        </div>
        <div className="seg doze" role="group" aria-label="Turnos de 12 horas">
          {TURNOS_12H.map((t) => (
            <button key={t.id} aria-pressed={turno === t.id} onClick={() => setTurno(t.id)}>{t.label}</button>
          ))}
        </div>

        <button className="b ghost" title="Ir para o turno em andamento"
          onClick={() => { const d = dataOperacional(); setData(d); setTurno(turnoAutomatico(d)); }}>
          agora
        </button>

        <div style={{ flex: 1 }} />
        <div style={{ textAlign: "right", fontSize: 12, color: "#C3C8EA",
          paddingRight: 14, borderRight: "1px solid rgba(255,255,255,.18)" }}>
          <div className="relogio">{horaMin(relogio)}</div>
          <div style={{ marginTop: 2 }}>
            {preenchidos} de {equipamentos.length} com operador{parados ? ` · ${parados} fora` : ""}
          </div>
        </div>
        <button className="b pri" onClick={() => setModo("tv")}>Modo TV</button>
      </div>

      {/* ---------- nível 2: ações ---------- */}
      <div className="n2">
        <input ref={inputRef} type="file" accept=".xlsx,.xls" multiple style={{ display: "none" }}
          onChange={(e) => { importar([...e.target.files]); e.target.value = ""; }} />

        <div className="passos">
          {passos.map((p, idx) => (
            <React.Fragment key={p.n}>
              {idx > 0 && <span className="seta" aria-hidden="true">›</span>}
              <button className={`passo${p.feito ? " ok" : ""}`} onClick={p.acao}
                title={p.feito ? `${p.rot} — já feito` : p.rot}>
                <span className="num">{p.feito ? "✓" : p.n}</span>
                {p.rot}
                {p.qtd ? <span className="qtd">{p.qtd}</span> : null}
              </button>
            </React.Fragment>
          ))}
        </div>

        <div className="sep" />
        <button className="b" onClick={abrirHistorico}>Histórico</button>
        <button className="b" onClick={copiarResumo}>Copiar resumo</button>
        <button className="b perigo" onClick={() => setPainel("zerar")}>Zerar</button>
        <div style={{ flex: 1 }} />
        <span className="dim" style={{ fontSize: 12 }}>{usuario.nome}</span>
        <button className={`b sm${cofre ? " pri" : ""}`} onClick={() => setPainel("acesso")}
          title={cofre ? "Quadro protegido por senha" : "Quadro sem senha"}>
          {cofre ? "protegido" : "sem senha"}
        </button>
        <button className="b ghost sm" onClick={sair}>sair</button>
      </div>
      </div>

      {aviso && (
        <div style={{ padding: "9px 16px", background: "#1D3320", borderBottom: "1px solid #3F7A33", display: "flex", gap: 12 }}>
          <span style={{ flex: 1, fontSize: 13 }}>{aviso}</span>
          <button className="b sm" onClick={() => setAviso(null)}>Fechar</button>
        </div>
      )}

      {operadores.length === 0 ? (
        <div style={{ padding: "48px 24px", maxWidth: 640 }}>
          <h2 style={{ fontSize: 25, marginBottom: 10 }}>Comece importando as escalas do mês</h2>
          <p className="hint">
            Envie de uma vez os três arquivos: ESCALA_OP_TT, ESCALA_OP_III e ESCALA_OP_IV. O app identifica
            a função pelo nome do arquivo e lê equipes, matrículas, o RS de referência e os códigos de cada dia,
            incluindo os turnos de 12 horas do fim de semana.
          </p>
          <button className="b pri" style={{ marginTop: 14 }} onClick={() => inputRef.current?.click()}>
            Escolher arquivos
          </button>
        </div>
      ) : (
        <div className="wrap">
          {/* ---------- lateral ---------- */}
          <aside className="lat">
            <div className="row" style={{ marginBottom: 8 }}>
              <h2 style={{ fontSize: 16, flex: 1 }}>Disponíveis</h2>
              <span className="tag" style={{ background: "var(--sur2)", color: "var(--dim)" }}>
                {disponiveis.length - usados.size} livres
              </span>
            </div>

            <input className="i" style={{ width: "100%", marginBottom: 8 }} placeholder="Buscar nome ou matrícula"
              value={buscaOp} onChange={(e) => setBuscaOp(e.target.value)} />

            <div className="sech">
              <b style={{ color: "#C7A2F0" }}>Navio e pátio</b>
              <span className="sub">arraste até o equipamento</span>
              <span className="qt2">{frentes.length}</span>
            </div>
            <div className="chips">
              {frentes.map((f) => (
                <span
                  key={f.id}
                  className="dz"
                  draggable
                  onDragStart={(ev) => iniciarArrasto(ev, { tipo: "frente", id: f.id })}
                  onDragEnd={encerrarArrasto}
                  style={{ borderColor: f.cor, color: f.cor, cursor: "grab", padding: "5px 9px" }}
                  title={`${nomeFrente(f)} — arraste para um equipamento ou para um operador`}
                >
                  {nomeFrente(f)}
                  <span className="qt">
                    {Object.values(quadro.locais || {}).filter((v) => v === f.id).length}
                  </span>
                </span>
              ))}
              <button className="b ghost sm" onClick={() => setPainel("frentes")}>+ navio</button>
            </div>

            {opSelecionado && (
              <div style={{ background: "var(--sur3)", border: "1px solid var(--cyan)", borderRadius: 8,
                padding: "8px 10px", margin: "12px 0 0", fontSize: 12.5 }}>
                <strong>{porId.get(opSelecionado)?.nome}</strong> selecionado.
                Clique num equipamento para alocar.
                <button className="b ghost sm" style={{ marginLeft: 6 }} onClick={() => setOpSelecionado(null)}>
                  cancelar
                </button>
              </div>
            )}

            {listaLateral.length === 0 && (
              <p className="hint" style={{ marginTop: 14 }}>
                {buscaOp ? "Ninguém com esse nome entre os disponíveis." :
                  "Ninguém disponível neste turno. Confira a data ou traga TPAs do OGMO."}
              </p>
            )}

            {GRUPOS_LATERAL.map((g) => {
              const gente = listaLateral.filter((o) => funcaoDe(o) === g.fn);
              if (!gente.length) return null;
              const livres = gente.filter((o) => !usados.has(o.id)).length;
              return (
                <div key={g.fn}>
                  <div className="sech">
                    <b style={{ color: g.cor }}>{g.titulo}</b>
                    <span className="sub">{g.sub}</span>
                    <span className="qt2">{livres} de {gente.length} livres</span>
                  </div>
                  {gente.map((o) => (
                    <div
                      key={o.id}
                      className={`oc${arrasto?.tipo === "op" && arrasto.id === o.id ? " fantasma" : ""}`}
                      data-sel={opSelecionado === o.id ? "1" : "0"}
                      draggable
                      onDragStart={(ev) => iniciarArrasto(ev, { tipo: "op", id: o.id })}
                      onDragEnd={encerrarArrasto}
                      onDragOver={(ev) => { if (arrasto?.tipo === "frente") ev.preventDefault(); }}
                      onDrop={(ev) => soltarNoOperador(ev, o.id)}
                      title="Arraste até um equipamento, ou solte um navio aqui"
                    >
                      <span className="dot" style={{ background: usados.has(o.id) ? "#3F4472" : STATUS[o.status]?.cor }} />
                      <button onClick={() => setOpSelecionado(opSelecionado === o.id ? null : o.id)}
                        style={{ flex: 1, minWidth: 0, background: "none", border: 0, color: "inherit",
                          textAlign: "left", cursor: "pointer", font: "inherit", padding: 0 }}>
                        <span style={{ display: "block", fontWeight: 600, fontSize: 13,
                          opacity: usados.has(o.id) ? 0.42 : 1, whiteSpace: "nowrap",
                          overflow: "hidden", textOverflow: "ellipsis" }}>{o.nome}</span>
                        <span className="dim" style={{ fontSize: 11 }}>
                          {o.mat && `${o.mat}`}{o.equipe && ` · Eq. ${o.equipe}`}{o.rs && ` · RS ${o.rs}`}
                          {o.ogmo && " · OGMO"}{o.herdado && " · da véspera"}
                        </span>
                      </button>
                      <button className="b ghost sm" title="Marcar atestado, falta ou folga"
                        onClick={() => setPainel(`ajuste:${o.id}`)}>•••</button>
                    </div>
                  ))}
                </div>
              );
            })}

            {foraDaOperacao.length > 0 && (
              <details style={{ marginTop: 14 }}>
                <summary style={{ cursor: "pointer", color: "var(--dim)", fontSize: 12.5 }}>
                  Fora da operação ({foraDaOperacao.length})
                </summary>
                <div style={{ marginTop: 6 }}>
                  {foraDaOperacao.map((o) => (
                    <div key={o.id} className="oc">
                      <span className="dot" style={{ background: STATUS[o.status]?.cor }} />
                      <span style={{ flex: 1, minWidth: 0 }}>
                        <span style={{ display: "block", fontSize: 12.5, whiteSpace: "nowrap",
                          overflow: "hidden", textOverflow: "ellipsis" }}>{o.nome}</span>
                        <span className="dim" style={{ fontSize: 11 }}>
                          {STATUS[o.status]?.rot}
                          {o.ajuste ? ` · por ${o.ajuste.por}` : o.codigo ? ` · ${o.codigo}` : ""}
                          {o.ajuste?.motivo ? ` · ${o.ajuste.motivo}` : ""}
                        </span>
                      </span>
                      <button className="b ghost sm" onClick={() => setPainel(`ajuste:${o.id}`)}>•••</button>
                    </div>
                  ))}
                </div>
              </details>
            )}
          </aside>

          {/* ---------- quadro ---------- */}
          <main className="main">
            <input className="i" style={{ width: "100%", marginBottom: 12 }}
              placeholder="Observação do turno (aparece na TV)"
              value={quadro.obs || ""}
              onChange={(e) => setQuadro({ ...quadro, obs: e.target.value })}
              onBlur={() => gravarQuadro(quadro, "Atualizou a observação do turno")} />

            <div className="dock">
              <span className="rot">
                {selecao.length
                  ? `Mandar ${selecao.length} equipamento${selecao.length > 1 ? "s" : ""} para`
                  : "Frentes de trabalho — arraste equipamentos para cá"}
              </span>
              {[...frentes, { id: "", nome: "Sem frente", terno: "", cor: "#6E75AC" }].map((f) => {
                const n = f.id
                  ? Object.values(quadro.locais || {}).filter((v) => v === f.id).length
                  : equipamentos.length - Object.keys(quadro.locais || {}).length;
                return (
                  <button
                    key={f.id || "sem"}
                    className="dz"
                    data-alvo={alvo === `f:${f.id}` ? "1" : "0"}
                    style={{ borderColor: f.cor, color: f.cor }}
                    onDragOver={(ev) => { ev.preventDefault(); setAlvo(`f:${f.id}`); }}
                    onDragLeave={() => setAlvo(null)}
                    onDrop={(ev) => soltarNaFrente(ev, f.id || null)}
                    onClick={() => {
                      if (!selecao.length) { setPainel("frentes"); return; }
                      definirLocalEmLote(selecao, f.id || null);
                      setSelecao([]);
                    }}
                  >
                    {nomeFrente(f) || f.nome}
                    <span className="qt">{n}</span>
                  </button>
                );
              })}
              <button className="b ghost sm" onClick={() => setPainel("frentes")}>+ navio</button>
            </div>

            {ORDEM_FAM.map((fam) => {
              const lista = porFamilia[fam];
              const cor = FAMILIAS[fam].cor;
              const n = lista.filter((e) => quadro.alocacoes?.[e.id]).length;
              const fora = lista.filter((e) => quadro.fora?.[e.id]).length;
              return (
                <section className="grp" key={fam}>
                  <div className="grph" style={{ borderLeftColor: cor }}>
                    <h3 style={{ color: cor }}>{FAMILIAS[fam].nome}</h3>
                    <span className="cnt">
                      {n} de {lista.length}{fora ? ` · ${fora} fora` : ""}
                    </span>
                    <div className="barra" />
                    <button className="b ghost sm" onClick={() => selecionarFamilia(fam)}>selecionar todos</button>
                  </div>
                  <div className="slots">
                    {lista.map((e) => (
                      <CartaoEquipamento
                        key={e.id}
                        eq={e}
                        cor={cor}
                        op={porId.get(quadro.alocacoes?.[e.id])}
                        parado={quadro.fora?.[e.id]}
                        intervalo={quadro.intervalos?.[e.id]}
                        frente={frentePorId.get(quadro.locais?.[e.id])}
                        selecionado={selecao.includes(e.id)}
                        onToggle={() => alternarSelecao(e.id)}
                        onAbrir={() => (opSelecionado ? alocar(e.id, opSelecionado) : setFicha(e.id))}
                        alvo={alvo === `e:${e.id}`}
                        fantasma={arrasto?.tipo === "equips" && arrasto.ids.includes(e.id)}
                        onDragStart={(ev) => iniciarArrasto(ev, { tipo: "equips", ids: idsDoArrasto(e.id) })}
                        onDragEnd={encerrarArrasto}
                        onDragOver={(ev) => { ev.preventDefault(); setAlvo(`e:${e.id}`); }}
                        onDragLeave={() => setAlvo(null)}
                        onDrop={(ev) => soltarNoEquipamento(ev, e.id)}
                      />
                    ))}
                  </div>
                </section>
              );
            })}
          </main>
        </div>
      )}

      {/* ---------- barra flutuante de seleção ---------- */}
      {selecao.length > 0 && (
        <div className="flut">
          <span className="cont">{selecao.length} selecionado{selecao.length > 1 ? "s" : ""}</span>
          <span className="dim" style={{ fontSize: 12 }}>enviar para</span>
          <div className="chips">
            {frentes.map((f) => (
              <button key={f.id} className="b sm"
                style={{ borderColor: f.cor, color: f.cor }}
                onClick={() => { definirLocalEmLote(selecao, f.id); setSelecao([]); }}>
                {nomeFrente(f)}
              </button>
            ))}
            <button className="b sm" onClick={() => { definirLocalEmLote(selecao, null); setSelecao([]); }}>
              sem frente
            </button>
          </div>
          <div className="sep" />
          <button className="b sm" onClick={() => setPainel("frentes")}>nova frente</button>
          <button className="b ghost sm" onClick={() => setSelecao([])}>limpar</button>
        </div>
      )}

      {/* ---------- ficha do equipamento ---------- */}
      {ficha && equipamentos.some((e) => e.id === ficha) && (
        <FichaEquipamento
          eq={equipamentos.find((e) => e.id === ficha)}
          op={porId.get(quadro.alocacoes?.[ficha])}
          parado={quadro.fora?.[ficha]}
          intervalo={quadro.intervalos?.[ficha] || ""}
          frenteId={quadro.locais?.[ficha] || ""}
          frentes={frentes}
          disponiveis={disponiveis}
          usados={usados}
          onFechar={() => setFicha(null)}
          onAlocar={(opId) => alocar(ficha, opId)}
          onIntervalo={(h) => definirIntervalo(ficha, h)}
          onFrente={(id) => definirLocalEmLote([ficha], id)}
          onParada={(m) => definirParada(ficha, m)}
        />
      )}

      {painel === "intervalos" && (
        <PainelIntervalos data={data} equipamentos={equipamentos} quadro={quadro} porId={porId}
          onFechar={() => setPainel(null)}
          onAplicar={(c) => { preencherIntervalos(c); setPainel(null); }} />
      )}

      {painel === "frentes" && (
        <PainelFrentes frentes={frentes} locais={quadro.locais || {}}
          onFechar={() => setPainel(null)}
          onSalvar={(fr, lo) => {
            gravarQuadro({ ...quadro, frentes: fr, locais: lo }, `Atualizou as frentes (${fr.length})`);
            setPainel(null);
          }} />
      )}

      {painel === "ogmo" && (
        <PainelOgmo extras={extras} onFechar={() => setPainel(null)}
          onSalvar={(l) => { gravarQuadro({ ...quadro, extras: l }, `Atualizou os TPAs do OGMO (${l.length})`); setPainel(null); }} />
      )}

      {painel?.startsWith("ajuste:") && (() => {
        const op = porId.get(painel.slice(7));
        if (!op) return null;
        return <PainelAjuste op={op} data={data} onFechar={() => setPainel(null)}
          onAplicar={(s, m) => { gravarAjuste(op.id, s, m); setPainel(null); }} />;
      })()}

      {painel === "historico" && (
        <PainelHistorico
          data={data} historico={historico} equipamentos={equipamentos}
          operadores={operadores} carregarDia={carregarDia} onFechar={() => setPainel(null)} />
      )}

      {painel === "acesso" && (
        <PainelAcesso
          cofre={cofre} usuario={usuario}
          onFechar={() => setPainel(null)}
          onProteger={protegerComSenha}
          onAdicionar={adicionarAcesso}
          onRemover={removerAcesso}
          onAviso={setAviso}
        />
      )}

      {painel === "zerar" && (
        <PainelZerar
          data={data} turno={turno}
          resumo={{
            alocados: preenchidos,
            parados,
            extras: extras.length,
            frentes: frentes.length,
            ajustes: Object.keys(ajustes).length,
            operadores: operadores.length,
          }}
          onFechar={() => setPainel(null)}
          onZerar={(op) => { zerar(op); setPainel(null); }}
        />
      )}
    </div>
  );
}

/* ================================================================== */
/*  CARTÃO DE EQUIPAMENTO                                              */
/* ================================================================== */

function CartaoEquipamento({
  eq, cor, op, parado, intervalo, frente, selecionado, alvo, fantasma,
  onToggle, onAbrir, onDragStart, onDragEnd, onDragOver, onDragLeave, onDrop,
}) {
  const classe = ["card", parado ? "parado" : op ? "" : "vazio", selecionado ? "sel" : "", fantasma ? "fantasma" : ""]
    .filter(Boolean).join(" ");
  return (
    <div
      className={classe}
      data-alvo={alvo ? "1" : "0"}
      style={{ borderTopColor: parado ? "var(--red)" : op ? cor : "var(--lin)" }}
      role="button"
      tabIndex={0}
      draggable={!parado && (!!op || selecionado)}
      onDragStart={onDragStart}
      onDragEnd={onDragEnd}
      onDragOver={onDragOver}
      onDragLeave={onDragLeave}
      onDrop={onDrop}
      onClick={onAbrir}
      onKeyDown={(e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); onAbrir(); } }}
      title={op ? "Arraste para outro equipamento ou para uma frente" : "Clique para escolher o operador"}
    >
      <button
        className="chk"
        data-on={selecionado ? "1" : "0"}
        aria-label={`Selecionar ${eq.nome}`}
        aria-pressed={selecionado}
        onClick={(e) => { e.stopPropagation(); onToggle(); }}
      >
        {selecionado ? "✓" : ""}
      </button>

      <div className="eqn" style={{ color: parado ? "var(--red)" : op ? cor : "#6E75AC" }}>{eq.nome}</div>
      <div className="opn" title={parado ? parado.motivo : op?.nome}>
        {parado ? parado.motivo : op ? op.nome : "livre"}
      </div>
      <div className="meta">
        {!parado && intervalo && <span className="int">{intervalo}</span>}
        {!parado && frente && (
          <span style={{ color: frente.cor, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
            {nomeFrente(frente)}
          </span>
        )}
      </div>
    </div>
  );
}

/* ================================================================== */
/*  FICHA DO EQUIPAMENTO                                               */
/* ================================================================== */

function FichaEquipamento({
  eq, op, parado, intervalo, frenteId, frentes, disponiveis, usados,
  onFechar, onAlocar, onIntervalo, onFrente, onParada,
}) {
  const [busca, setBusca] = useState("");
  const [trocando, setTrocando] = useState(!op && !parado);
  const [motivo, setMotivo] = useState(parado?.motivo || "");

  const q = busca.trim().toLowerCase();
  /* só aparecem os operadores habilitados para este tipo de equipamento */
  const lista = disponiveis
    .filter((o) => podeOperar(o, eq))
    .filter((o) => !q || o.nome.toLowerCase().includes(q) || String(o.mat).includes(q));
  const livres = lista.filter((o) => !usados.has(o.id));
  const ocupados = lista.filter((o) => usados.has(o.id));

  const Linha = (o) => (
    <button key={o.id} className="oc" onClick={() => { onAlocar(o.id); onFechar(); }}>
      <span className="dot" style={{ background: usados.has(o.id) ? "#3F4472" : STATUS[o.status]?.cor }} />
      <span style={{ flex: 1 }}>
        <span style={{ display: "block", fontWeight: 600, fontSize: 13 }}>{o.nome}</span>
        <span className="dim" style={{ fontSize: 11 }}>
          {FUNCOES[o.funcao]?.nome || o.funcaoTxt}{o.mat && ` · ${o.mat}`}{o.rs && ` · RS ${o.rs}`}
          {usados.has(o.id) && " · já está em outro equipamento"}
        </span>
      </span>
    </button>
  );

  return (
    <div className="modal" onClick={onFechar}>
      <div className="dlg" style={{ maxWidth: 540 }} onClick={(e) => e.stopPropagation()}>
        <div className="dlgh">
          <div style={{ flex: 1 }}>
            <h3 style={{ color: FAMILIAS[eq.fam].cor }}>{eq.nome}</h3>
            <span className="hint">
              Somente {FAMILIAS[eq.fam].prefere.map((f) => FUNCOES[f].nome).join(" ou ")}
            </span>
          </div>
          <button className="b ghost sm" onClick={onFechar}>fechar</button>
        </div>

        {parado ? (
          <>
            <p className="hint" style={{ marginTop: 0 }}>
              Fora de operação desde {horaMin(new Date(parado.em))}, por {parado.por}. Motivo: {parado.motivo}.
            </p>
            <button className="b pri" style={{ marginTop: 12 }} onClick={() => { onParada(null); onFechar(); }}>
              Voltar à operação
            </button>
          </>
        ) : (
          <>
            <div className="campo">
              <label>Operador</label>
              {op && !trocando ? (
                <div className="row">
                  <strong style={{ flex: 1, fontSize: 15 }}>{op.nome}</strong>
                  <button className="b sm" onClick={() => setTrocando(true)}>trocar</button>
                  <button className="b sm" onClick={() => { onAlocar(null); onFechar(); }}>remover</button>
                </div>
              ) : (
                <>
                  <input className="i" style={{ width: "100%" }} autoFocus
                    placeholder="Buscar nome ou matrícula"
                    value={busca} onChange={(e) => setBusca(e.target.value)} />
                  <div style={{ maxHeight: 210, overflow: "auto", marginTop: 6 }}>
                    {livres.map(Linha)}
                    {ocupados.length > 0 && (
                      <div className="hint" style={{ margin: "8px 0 2px" }}>Já estão em outro equipamento</div>
                    )}
                    {ocupados.map(Linha)}
                    {lista.length === 0 && (
                      <p className="hint">
                        {q ? "Ninguém com esse nome." :
                          `Nenhum ${FAMILIAS[eq.fam].prefere.map((f) => FUNCOES[f].nome).join(" ou ")} disponível neste turno.`}
                      </p>
                    )}
                  </div>
                </>
              )}
            </div>

            <div className="campo">
              <label>Intervalo</label>
              <div className="row">
                <input className="i" type="time" value={intervalo} onChange={(e) => onIntervalo(e.target.value)} />
                {intervalo && <button className="b ghost sm" onClick={() => onIntervalo("")}>limpar</button>}
              </div>
            </div>

            <div className="campo">
              <label>Frente de trabalho</label>
              <div className="chips">
                {frentes.map((f) => (
                  <button key={f.id} className="b sm"
                    style={frenteId === f.id
                      ? { background: f.cor, borderColor: f.cor, color: "#10122B", fontWeight: 600 }
                      : { borderColor: f.cor, color: f.cor }}
                    onClick={() => onFrente(f.id)}>{nomeFrente(f)}</button>
                ))}
                {frenteId && <button className="b ghost sm" onClick={() => onFrente(null)}>tirar</button>}
              </div>
            </div>

            <div className="campo">
              <label>Tirar de operação</label>
              <div className="chips">
                {MOTIVOS_PARADA.map((m) => (
                  <button key={m} className="b sm" style={motivo === m ? { background: "#4A2230", borderColor: "#7A3644" } : null}
                    onClick={() => setMotivo(m)}>{m}</button>
                ))}
              </div>
              <div className="row" style={{ marginTop: 8 }}>
                <input className="i" style={{ flex: 1 }} placeholder="ou escreva o motivo"
                  value={motivo} onChange={(e) => setMotivo(e.target.value)} />
                <button className="b perigo" disabled={!motivo.trim()}
                  onClick={() => { onParada(motivo.trim()); onFechar(); }}>Marcar fora</button>
              </div>
            </div>
          </>
        )}
      </div>
    </div>
  );
}

/* ================================================================== */
/*  TELA DA TV                                                         */
/* ================================================================== */

function TelaTV({
  quadro, equipamentos, porFamilia, porId, frentes, frentePorId,
  data, turno, relogio, tvPorFrente, onAlternarVista, onSair,
}) {
  const preenchidos = Object.keys(quadro.alocacoes || {}).length;
  const parados = Object.keys(quadro.fora || {}).length;

  /* Uma linha por equipamento, no mesmo formato da aba de escalados. */
  const Linha = ({ e, mostrarFrente }) => {
    const op = porId.get(quadro.alocacoes?.[e.id]);
    const parado = quadro.fora?.[e.id];
    const hora = quadro.intervalos?.[e.id];
    const frente = frentePorId.get(quadro.locais?.[e.id]);
    return (
      <div className="tvl">
        <span className="tvl-eq" style={{ color: parado ? "var(--red)" : FAMILIAS[e.fam].cor }}>
          {e.nome}
        </span>
        <span className="tvl-op" style={{ color: parado ? "#F0A0AC" : op ? "#EDEFF9" : "#7A81BA" }}>
          {parado ? parado.motivo : op ? op.nome : "sem operador"}
        </span>
        {mostrarFrente && frente && (
          <span className="tvl-fr" style={{ color: frente.cor }}>{nomeFrente(frente)}</span>
        )}
        <span className="tvl-hr">{!parado && hora ? hora : ""}</span>
      </div>
    );
  };

  /* A TV mostra só o que está em operação: equipamento com operador, ou parado
     com motivo. Os livres ficam de fora para o quadro não poluir. */
  const emOperacao = (e) => !!quadro.alocacoes?.[e.id] || !!quadro.fora?.[e.id];
  const visiveis = equipamentos.filter(emOperacao);
  const ocultos = equipamentos.length - visiveis.length;

  const grupos = (tvPorFrente
    ? [
        ...frentes.map((f) => ({
          chave: f.id, titulo: nomeFrente(f), cor: f.cor,
          itens: visiveis.filter((e) => quadro.locais?.[e.id] === f.id),
        })),
        {
          chave: "sem", titulo: "Sem frente definida", cor: "#6E75AC",
          itens: visiveis.filter((e) => !quadro.locais?.[e.id]),
        },
      ]
    : ORDEM_FAM.map((f) => ({
        chave: f, titulo: FAMILIAS[f].nome, cor: FAMILIAS[f].cor,
        itens: porFamilia[f].filter(emOperacao),
      }))
  ).filter((g) => g.itens.length);

  return (
    <div className="jbs"><style>{CSS}</style>
      <div className="tv">
        <div className="tvh">
          <div>
            <div className="marca" style={{ border: 0, padding: 0, marginBottom: 4 }}>
              JBS <span>Terminais</span>
            </div>
            <div style={{ fontFamily: "Poppins", fontWeight: 700, fontSize: 40, lineHeight: 1.02 }}>
              Turno {rotuloTurno(turno)}
            </div>
            <div style={{ fontSize: 19, color: "#B4BAE6", marginTop: 2 }}>{rotuloData(data)}</div>
          </div>
          <div style={{ textAlign: "right" }}>
            <div style={{ fontFamily: "Poppins", fontWeight: 700, fontSize: 54, lineHeight: 1,
              fontVariantNumeric: "tabular-nums" }}>
              {horaMin(relogio)}
            </div>
            <div className="dim" style={{ fontSize: 15, marginTop: 4 }}>
              {preenchidos} de {equipamentos.length} com operador
              {parados > 0 && ` · ${parados} fora de operação`}
              {ocultos > parados && ` · ${ocultos - parados} livres não exibidos`}
              {quadro.por ? ` · por ${quadro.por}` : ""}
            </div>
            <div className="row" style={{ marginTop: 8, justifyContent: "flex-end" }}>
              <button className="b sm" onClick={onAlternarVista}>
                {tvPorFrente ? "ver por tipo" : "ver por frente"}
              </button>
              <button className="b sm" onClick={onSair}>sair · Esc</button>
            </div>
          </div>
        </div>

        {grupos.length === 0 && (
          <p style={{ fontSize: 22, color: "#7A81BA", marginTop: 40 }}>
            Nenhum equipamento em operação neste turno.
          </p>
        )}

        <div className="tvcols">
          {grupos.map((g) => (
            <section className="tvbloco" key={g.chave}>
              <div className="tvbh" style={{ borderColor: g.cor }}>
                <h3 style={{ color: g.cor }}>{g.titulo}</h3>
                <span className="dim" style={{ fontSize: 14 }}>{g.itens.length}</span>
              </div>
              {g.itens.map((e) => <Linha key={e.id} e={e} mostrarFrente={!tvPorFrente} />)}
            </section>
          ))}
        </div>

        {quadro.obs && (
          <div style={{ marginTop: 14, fontSize: 20, color: "var(--green)", fontWeight: 600 }}>
            {quadro.obs}
          </div>
        )}
      </div>
    </div>
  );
}

/* ================================================================== */
/*  ENTRADA                                                            */
/* ================================================================== */

function TelaEntrada({ onEntrar }) {
  const [nome, setNome] = useState("");
  const [conhecidos, setConhecidos] = useState([]);
  useEffect(() => { carregar("usuarios").then((l) => setConhecidos(l || [])); }, []);
  return (
    <div className="jbs"><style>{CSS}</style>
      <div style={{ maxWidth: 430, margin: "13vh auto", padding: 24 }}>
        <div style={{ fontFamily: "Poppins", fontWeight: 700, fontSize: 20, marginBottom: 22 }}>
          JBS <span style={{ color: "var(--cyan)" }}>Terminais</span>
        </div>
        <h1 style={{ fontSize: 30, marginBottom: 6 }}>Escala do pátio</h1>
        <p className="hint" style={{ marginTop: 0 }}>
          Diga quem está operando o quadro. O nome fica registrado em cada alteração do histórico.
        </p>
        {conhecidos.length > 0 && (
          <div className="chips" style={{ margin: "16px 0" }}>
            {conhecidos.map((n) => <button key={n} className="b" onClick={() => onEntrar(n)}>{n}</button>)}
          </div>
        )}
        <div className="row" style={{ marginTop: 10 }}>
          <input className="i" style={{ flex: 1 }} placeholder="Seu nome" value={nome}
            onChange={(e) => setNome(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && nome.trim() && onEntrar(nome)} />
          <button className="b pri" disabled={!nome.trim()} onClick={() => nome.trim() && onEntrar(nome)}>
            Entrar
          </button>
        </div>
        <p className="hint" style={{ marginTop: 18 }}>
          Isto identifica quem alterou o quê. Não é uma senha e não bloqueia o acesso de ninguém.
        </p>
      </div>
    </div>
  );
}

/* ================================================================== */
/*  TELA DE SENHA                                                      */
/* ================================================================== */

function TelaSenha({ cofre, onDestrancar }) {
  const nomes = Object.values(cofre.usuarios || {}).map((u) => u.nome).sort();
  const [nome, setNome] = useState(nomes.length === 1 ? nomes[0] : "");
  const [senha, setSenha] = useState("");
  const [lembrar, setLembrar] = useState(false);
  const [erro, setErro] = useState("");
  const [ocupado, setOcupado] = useState(false);

  async function tentar() {
    if (!nome || !senha) return;
    setOcupado(true); setErro("");
    const ok = await onDestrancar(nome, senha, lembrar);
    setOcupado(false);
    if (!ok) { setErro("Nome ou senha não conferem."); setSenha(""); }
  }

  return (
    <div className="jbs"><style>{CSS}</style>
      <div style={{ maxWidth: 400, margin: "13vh auto", padding: 24 }}>
        <div style={{ fontFamily: "Poppins", fontWeight: 700, fontSize: 20, marginBottom: 22 }}>
          JBS <span style={{ color: "var(--cyan)" }}>Terminais</span>
        </div>
        <h1 style={{ fontSize: 28, marginBottom: 6 }}>Escala do pátio</h1>
        <p className="hint" style={{ marginTop: 0 }}>
          Este quadro está protegido. Entre com o seu nome e senha.
        </p>

        {nomes.length > 1 && (
          <div className="chips" style={{ margin: "16px 0 4px" }}>
            {nomes.map((n) => (
              <button key={n} className={`b sm${nome === n ? " on" : ""}`} onClick={() => setNome(n)}>{n}</button>
            ))}
          </div>
        )}

        <div className="campo">
          <label>Nome</label>
          <input className="i" style={{ width: "100%" }} value={nome} autoComplete="username"
            onChange={(e) => setNome(e.target.value)} />
        </div>
        <div className="campo">
          <label>Senha</label>
          <input className="i" style={{ width: "100%" }} type="password" value={senha}
            autoComplete="current-password" autoFocus={!!nome}
            onChange={(e) => setSenha(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && tentar()} />
        </div>

        <label className="row" style={{ marginTop: 12, fontSize: 13, cursor: "pointer" }}>
          <input type="checkbox" checked={lembrar} onChange={(e) => setLembrar(e.target.checked)} />
          Lembrar neste computador
        </label>
        <p className="hint" style={{ marginTop: 4 }}>
          Marque na TV e no computador da sala, para não digitar a senha a cada recarga.
          Não marque em máquina compartilhada.
        </p>

        {erro && <p className="hint" style={{ color: "#F0A0AC", marginTop: 10 }}>{erro}</p>}

        <button className="b pri" style={{ marginTop: 14 }} disabled={ocupado || !nome || !senha} onClick={tentar}>
          {ocupado ? "Verificando…" : "Entrar"}
        </button>

        <p className="hint" style={{ marginTop: 20 }}>
          Não existe recuperação de senha: os dados são cifrados com ela e ninguém, nem a Anthropic,
          consegue abri-los sem uma senha válida. Se todas se perderem, o quadro precisa ser recomeçado.
        </p>
      </div>
    </div>
  );
}

/* ================================================================== */
/*  PAINEL DE ACESSO                                                   */
/* ================================================================== */

function PainelAcesso({ cofre, usuario, onFechar, onProteger, onAdicionar, onRemover, onAviso }) {
  const [nome, setNome] = useState(usuario?.nome || "");
  const [senha, setSenha] = useState("");
  const [conf, setConf] = useState("");
  const [erro, setErro] = useState("");
  const [ocupado, setOcupado] = useState(false);

  const fraca = senha.length > 0 && senha.length < 8;

  async function aplicar() {
    setErro("");
    if (!nome.trim()) return setErro("Informe o nome.");
    if (senha.length < 8) return setErro("A senha precisa de pelo menos 8 caracteres.");
    if (senha !== conf) return setErro("As duas senhas não são iguais.");
    setOcupado(true);
    const r = cofre ? await onAdicionar(nome, senha) : await onProteger(nome, senha);
    setOcupado(false);
    if (r?.erro) return setErro(r.erro);
    setSenha(""); setConf("");
    onAviso(r?.ok || "Pronto.");
    if (!cofre) onFechar();
  }

  const gente = Object.entries(cofre?.usuarios || {});

  return (
    <div className="modal" onClick={onFechar}>
      <div className="dlg" style={{ maxWidth: 520 }} onClick={(e) => e.stopPropagation()}>
        <div className="dlgh">
          <h3 style={{ flex: 1 }}>{cofre ? "Quem pode abrir o quadro" : "Proteger com senha"}</h3>
          <button className="b ghost sm" onClick={onFechar}>fechar</button>
        </div>

        {!temCripto() && (
          <p className="hint" style={{ color: "#F0A0AC" }}>
            Este navegador não expõe as funções de criptografia. Abra o app por https.
          </p>
        )}

        {!cofre ? (
          <p className="hint" style={{ marginTop: 0 }}>
            Hoje qualquer pessoa com o link vê o quadro inteiro. Ao proteger, todo o conteúdo passa a
            ser cifrado com uma chave que só as senhas cadastradas abrem — nomes, matrículas, histórico.
            Quem tiver o link e não tiver senha vê apenas dados embaralhados.
          </p>
        ) : (
          <p className="hint" style={{ marginTop: 0 }}>
            {gente.length} pessoa{gente.length > 1 ? "s" : ""} com acesso. Cada uma tem a própria senha,
            e o nome dela aparece no histórico de alterações.
          </p>
        )}

        {cofre && (
          <div style={{ marginTop: 12 }}>
            {gente.map(([k, u]) => (
              <div key={k} className="oc">
                <span className="dot" style={{ background: "var(--green)" }} />
                <span style={{ flex: 1, fontSize: 13, fontWeight: 600 }}>{u.nome}</span>
                {gente.length > 1 && u.nome !== usuario?.nome && (
                  <button className="b ghost sm" onClick={() => onRemover(k)}>tirar acesso</button>
                )}
              </div>
            ))}
          </div>
        )}

        <div className="campo">
          <label>{cofre ? "Liberar acesso para" : "Seu nome"}</label>
          <input className="i" style={{ width: "100%" }} value={nome} onChange={(e) => setNome(e.target.value)} />
        </div>
        <div className="campo">
          <label>Senha {cofre ? "dessa pessoa" : ""}</label>
          <input className="i" style={{ width: "100%" }} type="password" value={senha}
            autoComplete="new-password" onChange={(e) => setSenha(e.target.value)} />
          {fraca && <p className="hint" style={{ color: "#E8B24D" }}>Pelo menos 8 caracteres.</p>}
        </div>
        <div className="campo">
          <label>Repita a senha</label>
          <input className="i" style={{ width: "100%" }} type="password" value={conf}
            autoComplete="new-password" onChange={(e) => setConf(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && aplicar()} />
        </div>

        {erro && <p className="hint" style={{ color: "#F0A0AC", marginTop: 10 }}>{erro}</p>}

        <div className="row" style={{ marginTop: 16 }}>
          <button className="b pri" disabled={ocupado || !temCripto()} onClick={aplicar}>
            {ocupado ? "Aplicando…" : cofre ? "Liberar acesso" : "Proteger agora"}
          </button>
          <button className="b" onClick={onFechar}>Cancelar</button>
        </div>

        <p className="hint" style={{ marginTop: 16 }}>
          {cofre
            ? "Tirar o acesso impede novas entradas, mas quem já marcou “lembrar neste computador” continua com a chave naquela máquina até sair. Para cortar de vez, é preciso recriar o quadro com uma chave nova."
            : "Anote a senha em lugar seguro antes de continuar. Não há recuperação: os dados são cifrados com ela."}
        </p>
      </div>
    </div>
  );
}

/* ================================================================== */
/*  PAINÉIS                                                            */
/* ================================================================== */

function PainelIntervalos({ data, equipamentos, quadro, porId, onFechar, onAplicar }) {
  const padrao = intervaloPadrao(data);
  const [familias, setFamilias] = useState(["RS"]);
  const [inicio, setInicio] = useState("");
  const [duracao, setDuracao] = useState(padrao);
  const [duracaoOgmo, setDuracaoOgmo] = useState(INTERVALO_OGMO);
  const [porVez, setPorVez] = useState(1);
  const [sobrescrever, setSobrescrever] = useState(false);

  const alvos = equipamentos.filter(
    (e) => familias.includes(e.fam) && quadro.alocacoes?.[e.id] && !quadro.fora?.[e.id]);
  const aPreencher = sobrescrever ? alvos : alvos.filter((e) => !quadro.intervalos?.[e.id]);
  const durDe = (e) => (porId.get(quadro.alocacoes[e.id])?.ogmo ? duracaoOgmo : duracao);
  const nOgmo = aPreencher.filter((e) => porId.get(quadro.alocacoes[e.id])?.ogmo).length;

  const fmt = (min) =>
    `${String(Math.floor((min / 60) % 24)).padStart(2, "0")}:${String(min % 60).padStart(2, "0")}`;

  let fim = null;
  const previa = [];
  if (inicio && aPreencher.length) {
    const [h, m] = inicio.split(":").map(Number);
    let min = h * 60 + m;
    let grupo = [];
    aPreencher.forEach((e, idx) => {
      if (idx < 5) previa.push(`${e.nome} ${fmt(min)}`);
      grupo.push(durDe(e));
      if (grupo.length >= porVez) { min += Math.max(...grupo); grupo = []; }
    });
    if (grupo.length) min += Math.max(...grupo);
    fim = fmt(min);
  }

  const alternar = (f) => setFamilias(familias.includes(f) ? familias.filter((x) => x !== f) : [...familias, f]);

  return (
    <div className="modal" onClick={onFechar}>
      <div className="dlg" style={{ maxWidth: 520 }} onClick={(e) => e.stopPropagation()}>
        <div className="dlgh"><h3>Distribuir intervalos</h3>
          <button className="b ghost sm" onClick={onFechar}>fechar</button></div>
        <p className="hint" style={{ marginTop: 0 }}>
          Informe o horário do primeiro e o app escalona o resto.
          Pessoal próprio: {fimDeSemana(data) ? "60 minutos no fim de semana" : "15 minutos em dia útil"}.
          TPA do OGMO: {INTERVALO_OGMO} minutos, domingo a domingo.
        </p>

        <div className="campo">
          <label>Aplicar a</label>
          <div className="chips">
            {ORDEM_FAM.map((id) => (
              <button key={id} className={`b sm${familias.includes(id) ? " on" : ""}`} onClick={() => alternar(id)}>
                {FAMILIAS[id].nome}
              </button>
            ))}
          </div>
        </div>

        <div className="campo row" style={{ gap: 14, alignItems: "flex-end", flexWrap: "wrap" }}>
          <div><label style={{ display: "block", fontSize: 11.5, color: "var(--dim)", marginBottom: 5 }}>Primeiro</label>
            <input className="i" type="time" value={inicio} onChange={(e) => setInicio(e.target.value)} /></div>
          <div><label style={{ display: "block", fontSize: 11.5, color: "var(--dim)", marginBottom: 5 }}>Próprios (min)</label>
            <input className="i" type="number" min="5" max="120" step="5" style={{ width: 92 }}
              value={duracao} onChange={(e) => setDuracao(Math.max(5, Number(e.target.value) || padrao))} /></div>
          <div><label style={{ display: "block", fontSize: 11.5, color: "var(--dim)", marginBottom: 5 }}>OGMO (min)</label>
            <input className="i" type="number" min="5" max="120" step="5" style={{ width: 92 }}
              value={duracaoOgmo} onChange={(e) => setDuracaoOgmo(Math.max(5, Number(e.target.value) || INTERVALO_OGMO))} /></div>
          <div><label style={{ display: "block", fontSize: 11.5, color: "var(--dim)", marginBottom: 5 }}>Por vez</label>
            <input className="i" type="number" min="1" max="12" style={{ width: 78 }}
              value={porVez} onChange={(e) => setPorVez(Math.max(1, Number(e.target.value) || 1))} /></div>
        </div>

        <label className="row" style={{ marginTop: 12, fontSize: 13, cursor: "pointer" }}>
          <input type="checkbox" checked={sobrescrever} onChange={(e) => setSobrescrever(e.target.checked)} />
          Refazer também os que já têm horário
        </label>

        <p className="hint" style={{ marginTop: 12 }}>
          {aPreencher.length === 0
            ? "Nenhum equipamento com operador nessa seleção. Monte o quadro primeiro."
            : `${aPreencher.length} equipamentos${nOgmo ? `, sendo ${nOgmo} do OGMO` : ""}${fim ? `. Termina às ${fim}` : ""}.`}
          {previa.length > 0 && <><br />{previa.join(" · ")}{aPreencher.length > 5 ? " …" : ""}</>}
        </p>

        <div className="row" style={{ marginTop: 14 }}>
          <button className="b pri" disabled={!inicio || !aPreencher.length}
            onClick={() => onAplicar({ familias, inicio, duracao, duracaoOgmo, porVez, sobrescrever })}>Preencher</button>
          <button className="b" onClick={onFechar}>Cancelar</button>
        </div>
        <p className="hint" style={{ marginTop: 10 }}>
          Depois é só abrir o equipamento para corrigir um horário específico.
        </p>
      </div>
    </div>
  );
}

function PainelFrentes({ frentes, locais, onFechar, onSalvar }) {
  const [lista, setLista] = useState(frentes);
  const [nome, setNome] = useState("");
  const [ternos, setTernos] = useState(1);

  /* Um navio com N ternos vira N frentes: MSC Julia T1, T2, T3… */
  function adicionar() {
    if (!nome.trim()) return;
    const base = Date.now();
    const novos = [];
    const n = Math.max(0, Math.min(12, Number(ternos) || 0));
    if (n === 0) {
      novos.push({ id: `f${base}`, nome: nome.trim(), terno: "",
        cor: CORES_FRENTE[lista.length % CORES_FRENTE.length] });
    } else {
      for (let i = 1; i <= n; i++)
        novos.push({ id: `f${base}-${i}`, nome: nome.trim(), terno: String(i),
          cor: CORES_FRENTE[(lista.length + i - 1) % CORES_FRENTE.length] });
    }
    setLista([...lista, ...novos]);
    setNome(""); setTernos(1);
  }
  function salvar() {
    const ids = new Set(lista.map((f) => f.id));
    const l = {};
    Object.entries(locais).forEach(([k, v]) => { if (ids.has(v)) l[k] = v; });
    onSalvar(lista, l);
  }
  const usos = (id) => Object.values(locais).filter((v) => v === id).length;

  return (
    <div className="modal" onClick={onFechar}>
      <div className="dlg" style={{ maxWidth: 520 }} onClick={(e) => e.stopPropagation()}>
        <div className="dlgh"><h3>Frentes de trabalho</h3>
          <button className="b ghost sm" onClick={onFechar}>fechar</button></div>
        <p className="hint" style={{ marginTop: 0 }}>
          Diga o navio e quantos ternos ele tem: o app cria uma frente para cada um.
          Depois arraste os equipamentos para a frente, ou selecione vários e clique nela.
          Use zero ternos para locais como o pátio.
        </p>

        <div className="row" style={{ marginTop: 14, flexWrap: "wrap", alignItems: "flex-end" }}>
          <div style={{ flex: 2, minWidth: 170 }}>
            <label style={{ display: "block", fontSize: 11.5, color: "var(--dim)", marginBottom: 5 }}>Navio ou local</label>
            <input className="i" style={{ width: "100%" }} placeholder="ex.: MSC Julia"
              value={nome} onChange={(e) => setNome(e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && adicionar()} />
          </div>
          <div>
            <label style={{ display: "block", fontSize: 11.5, color: "var(--dim)", marginBottom: 5 }}>Ternos</label>
            <input className="i" type="number" min="0" max="12" style={{ width: 84 }}
              value={ternos} onChange={(e) => setTernos(e.target.value)} />
          </div>
          <button className="b" onClick={adicionar}>Adicionar</button>
        </div>
        {nome.trim() && Number(ternos) > 0 && (
          <p className="hint" style={{ marginTop: 8 }}>
            Serão criadas: {Array.from({ length: Math.min(12, Number(ternos)) },
              (_, i) => `${nome.trim()} · T${i + 1}`).join(", ")}
          </p>
        )}

        <div style={{ marginTop: 14 }}>
          {lista.map((f) => (
            <div key={f.id} className="oc">
              <span className="dot" style={{ background: f.cor }} />
              <span style={{ flex: 1 }}>
                <span style={{ display: "block", fontSize: 13, fontWeight: 600 }}>{nomeFrente(f)}</span>
                <span className="dim" style={{ fontSize: 11 }}>
                  {usos(f.id) ? `${usos(f.id)} equipamentos` : "nenhum equipamento ainda"}
                </span>
              </span>
              <button className="b ghost sm" onClick={() => setLista(lista.filter((x) => x.id !== f.id))}>tirar</button>
            </div>
          ))}
          {lista.length === 0 && <p className="hint">Nenhuma frente cadastrada.</p>}
        </div>

        <div className="row" style={{ marginTop: 16 }}>
          <button className="b pri" onClick={salvar}>Salvar</button>
          <button className="b" onClick={onFechar}>Cancelar</button>
        </div>
      </div>
    </div>
  );
}

function PainelAjuste({ op, data, onFechar, onAplicar }) {
  const [status, setStatus] = useState(op.ajuste?.status || op.status);
  const [motivo, setMotivo] = useState(op.ajuste?.motivo || "");
  return (
    <div className="modal" onClick={onFechar}>
      <div className="dlg" style={{ maxWidth: 440 }} onClick={(e) => e.stopPropagation()}>
        <div className="dlgh"><h3>{op.nome}</h3>
          <button className="b ghost sm" onClick={onFechar}>fechar</button></div>
        <p className="hint" style={{ marginTop: 0 }}>
          {FUNCOES[op.funcao]?.nome || "TPA"}{op.mat && ` · matrícula ${op.mat}`} · {rotuloData(data)}
          {op.codigo && ` · planilha diz ${op.codigo}`}
        </p>
        <div className="campo">
          <label>Situação hoje</label>
          <div className="chips">
            {AJUSTES.map((a) => (
              <button key={a.id} className={`b sm${status === a.id ? " on" : ""}`} onClick={() => setStatus(a.id)}>
                {a.rot}
              </button>
            ))}
          </div>
        </div>
        <div className="campo">
          <label>Motivo (opcional)</label>
          <input className="i" style={{ width: "100%" }} value={motivo} onChange={(e) => setMotivo(e.target.value)} />
        </div>
        <p className="hint" style={{ marginTop: 12 }}>
          Vale só para este dia. A planilha do mês continua intacta e o registro entra no histórico com o seu nome.
        </p>
        <div className="row" style={{ marginTop: 14 }}>
          <button className="b pri" onClick={() => onAplicar(status, motivo)}>Aplicar</button>
          {op.ajuste && <button className="b" onClick={() => onAplicar(null, "")}>Voltar para a planilha</button>}
          <button className="b" onClick={onFechar}>Cancelar</button>
        </div>
      </div>
    </div>
  );
}

function PainelHistorico({ data, historico, equipamentos, operadores, carregarDia, onFechar }) {
  const [aba, setAba] = useState("escalados");
  const [dia, setDia] = useState(data);
  const [turnos, setTurnos] = useState(null);
  const [copiado, setCopiado] = useState(false);

  useEffect(() => {
    let vivo = true;
    setTurnos(null);
    carregarDia(dia).then((r) => { if (vivo) setTurnos(r); });
    return () => { vivo = false; };
  }, [dia, carregarDia]);

  /* nomes vêm da escala importada e também dos TPAs gravados naquele turno */
  const nomeDe = (q, id) =>
    operadores.find((o) => o.id === id)?.nome ||
    (q.extras || []).find((x) => x.id === id)?.nome ||
    id;

  const linhasDoTurno = (q) =>
    equipamentos
      .filter((e) => q.alocacoes?.[e.id] || q.fora?.[e.id])
      .map((e) => ({
        eq: e,
        parado: q.fora?.[e.id],
        nome: q.fora?.[e.id] ? null : nomeDe(q, q.alocacoes[e.id]),
        hora: q.intervalos?.[e.id],
        frente: (q.frentes || []).find((f) => f.id === q.locais?.[e.id]),
      }));

  function copiarDia() {
    const l = [`ESCALADOS — ${rotuloData(dia)}`, ""];
    (turnos || []).forEach(({ turno, quadro: q }) => {
      l.push(`Turno ${rotuloTurno(turno)}`);
      linhasDoTurno(q).forEach((r) => {
        if (r.parado) { l.push(`  ${r.eq.nome}: FORA — ${r.parado.motivo}`); return; }
        const extra = [r.frente ? nomeFrente(r.frente) : null, r.hora ? `int. ${r.hora}` : null]
          .filter(Boolean).join(" · ");
        l.push(`  ${r.eq.nome}: ${r.nome}${extra ? ` (${extra})` : ""}`);
      });
      l.push("");
    });
    navigator.clipboard?.writeText(l.join("\n")).then(() => {
      setCopiado(true);
      setTimeout(() => setCopiado(false), 2500);
    });
  }

  return (
    <div className="modal" onClick={onFechar}>
      <div className="dlg" style={{ maxWidth: 660 }} onClick={(e) => e.stopPropagation()}>
        <div className="dlgh">
          <h3 style={{ flex: 1 }}>Registro do dia</h3>
          <button className="b ghost sm" onClick={onFechar}>fechar</button>
        </div>

        <div className="row" style={{ marginTop: 6, flexWrap: "wrap" }}>
          <div className="seg">
            <button aria-pressed={aba === "escalados"} onClick={() => setAba("escalados")}>Escalados</button>
            <button aria-pressed={aba === "alteracoes"} onClick={() => setAba("alteracoes")}>Alterações</button>
          </div>
          <div style={{ flex: 1 }} />
          <button className="b ghost sm" onClick={() => setDia(somarDias(dia, -1))}>‹</button>
          <input className="i" type="date" value={dia} onChange={(e) => setDia(e.target.value)} />
          <button className="b ghost sm" onClick={() => setDia(somarDias(dia, 1))}>›</button>
        </div>

        {aba === "escalados" ? (
          <div style={{ marginTop: 14 }}>
            {turnos === null ? <p className="hint">Carregando os turnos do dia…</p> :
              turnos.length === 0 ? <p className="hint">Nenhum quadro montado em {rotuloData(dia)}.</p> : (
                <>
                  {turnos.map(({ turno: t, quadro: q }) => {
                    const linhas = linhasDoTurno(q);
                    return (
                      <div key={t} style={{ marginBottom: 18 }}>
                        <div className="grph">
                          <h3 style={{ fontSize: 15, color: "var(--cyan)" }}>Turno {rotuloTurno(t)}</h3>
                          <span className="dim" style={{ fontSize: 12 }}>
                            {linhas.filter((r) => !r.parado).length} operadores
                            {q.por ? ` · por ${q.por}` : ""}
                          </span>
                          <div className="barra" />
                        </div>
                        {linhas.length === 0 ? <p className="hint">Quadro vazio.</p> :
                          linhas.map((r) => (
                            <div key={r.eq.id} style={{ display: "flex", gap: 10, alignItems: "baseline",
                              padding: "5px 0", borderBottom: "1px solid var(--lin)" }}>
                              <span style={{ width: 108, flex: "none", fontSize: 12.5, fontWeight: 600,
                                fontFamily: "Poppins,sans-serif", color: r.parado ? "var(--red)" : FAMILIAS[r.eq.fam].cor }}>
                                {r.eq.nome}
                              </span>
                              <span style={{ flex: 1, fontSize: 13, color: r.parado ? "#F0A0AC" : "inherit" }}>
                                {r.parado ? `fora — ${r.parado.motivo}` : r.nome}
                              </span>
                              {r.frente && (
                                <span style={{ fontSize: 11.5, color: r.frente.cor }}>{nomeFrente(r.frente)}</span>
                              )}
                              {r.hora && (
                                <span style={{ fontSize: 11.5, color: "var(--cyan)", fontVariantNumeric: "tabular-nums" }}>
                                  {r.hora}
                                </span>
                              )}
                            </div>
                          ))}
                        {q.obs && <p className="hint" style={{ color: "var(--green)" }}>{q.obs}</p>}
                      </div>
                    );
                  })}
                  <button className="b" onClick={copiarDia}>{copiado ? "Copiado" : "Copiar o dia inteiro"}</button>
                </>
              )}
            <p className="hint" style={{ marginTop: 12 }}>
              Mostra o quadro como está gravado hoje. Quem foi trocado durante o turno aparece na aba de alterações.
            </p>
          </div>
        ) : (
          <div style={{ marginTop: 14 }}>
            <p className="hint" style={{ marginTop: 0 }}>
              {dia === data
                ? "Tudo que mudou no dia, com hora e responsável."
                : "As alterações só são carregadas para o dia aberto no quadro."}
            </p>
            {dia !== data ? null : historico.length === 0 ?
              <p className="hint">Nenhuma alteração registrada nesta data.</p> :
              historico.map((h, i) => (
                <div key={i} style={{ display: "flex", gap: 10, padding: "7px 0", borderBottom: "1px solid var(--lin)" }}>
                  <span className="dim" style={{ fontSize: 12, width: 44, flex: "none", fontVariantNumeric: "tabular-nums" }}>
                    {horaMin(new Date(h.ts))}
                  </span>
                  <span style={{ flex: 1, fontSize: 13 }}>{h.acao}</span>
                  <span className="dim" style={{ fontSize: 11.5, whiteSpace: "nowrap" }}>{h.turno} · {h.por}</span>
                </div>
              ))}
          </div>
        )}
      </div>
    </div>
  );
}

function PainelZerar({ data, turno, resumo, onFechar, onZerar }) {
  const [op, setOp] = useState({ quadro: true, frentes: false, ajustes: false, escala: false, historico: false });
  const [confirmando, setConfirmando] = useState(false);
  const marcar = (k) => { setOp({ ...op, [k]: !op[k] }); setConfirmando(false); };
  const nada = !op.quadro && !op.ajustes && !op.escala && !op.historico;

  const Linha = ({ k, titulo, desc, alerta }) => (
    <label className="row" style={{ alignItems: "flex-start", gap: 9, padding: "9px 0",
      borderBottom: "1px solid var(--lin)", cursor: "pointer" }}>
      <input type="checkbox" checked={op[k]} onChange={() => marcar(k)} style={{ marginTop: 3 }} />
      <span style={{ flex: 1 }}>
        <span style={{ display: "block", fontSize: 13.5, fontWeight: 600,
          color: alerta ? "#F0A0AC" : "inherit" }}>{titulo}</span>
        <span className="hint">{desc}</span>
      </span>
    </label>
  );

  return (
    <div className="modal" onClick={onFechar}>
      <div className="dlg" style={{ maxWidth: 520 }} onClick={(e) => e.stopPropagation()}>
        <div className="dlgh"><h3>Zerar</h3>
          <button className="b ghost sm" onClick={onFechar}>fechar</button></div>
        <p className="hint" style={{ marginTop: 0 }}>
          Escolha o que apagar. Nada é desfeito depois, então confira antes de confirmar.
        </p>

        <div style={{ marginTop: 10 }}>
          <Linha k="quadro" titulo={`Quadro de ${rotuloData(data)}, turno ${rotuloTurno(turno)}`}
            desc={`${resumo.alocados} operadores alocados, ${resumo.parados} equipamentos fora, ${resumo.extras} TPAs do OGMO, intervalos e frentes atribuídas. Outros turnos e outros dias não são tocados.`} />
          {op.quadro && (
            <label className="row" style={{ alignItems: "flex-start", gap: 9, padding: "9px 0 9px 26px",
              borderBottom: "1px solid var(--lin)", cursor: "pointer" }}>
              <input type="checkbox" checked={op.frentes} onChange={() => marcar("frentes")} style={{ marginTop: 3 }} />
              <span style={{ flex: 1 }}>
                <span style={{ display: "block", fontSize: 13.5, fontWeight: 600 }}>
                  Apagar também os navios cadastrados
                </span>
                <span className="hint">
                  {resumo.frentes} frentes. Sem marcar, elas continuam disponíveis para o próximo quadro.
                </span>
              </span>
            </label>
          )}
          <Linha k="ajustes" titulo={`Atestados e faltas de ${rotuloData(data)}`}
            desc={`${resumo.ajustes} marcações manuais do dia. A planilha do mês volta a valer.`} />
          <Linha k="escala" alerta titulo="Escala importada do mês"
            desc={`Apaga os ${resumo.operadores} operadores lidos dos arquivos. Use quando a escala virar no dia 15 e você for importar as planilhas novas.`} />
          <Linha k="historico" alerta titulo={`Histórico de ${rotuloData(data)}`}
            desc="Apaga o registro de quem alterou o quê. Só faça isso se souber que não vai precisar auditar o dia." />
        </div>

        <div className="row" style={{ marginTop: 18 }}>
          {confirmando ? (
            <>
              <button className="b perigo" onClick={() => onZerar(op)}>Sim, apagar agora</button>
              <button className="b" onClick={() => setConfirmando(false)}>Voltar</button>
            </>
          ) : (
            <>
              <button className="b perigo" disabled={nada} onClick={() => setConfirmando(true)}>Zerar</button>
              <button className="b" onClick={onFechar}>Cancelar</button>
            </>
          )}
        </div>
        {confirmando && (
          <p className="hint" style={{ marginTop: 10, color: "#F0A0AC" }}>
            Confirmando, isto apaga definitivamente para todo mundo que usa o quadro, inclusive a TV.
          </p>
        )}
      </div>
    </div>
  );
}

function PainelOgmo({ extras, onFechar, onSalvar }) {
  const [texto, setTexto] = useState("");
  const [lista, setLista] = useState(extras);
  const [nome, setNome] = useState("");
  const [mat, setMat] = useState("");
  const [fn, setFn] = useState("III");
  const [erro, setErro] = useState("");

  function colar() {
    const novos = lerOgmo(texto);
    if (!novos.length) {
      setErro("Não identifiquei nenhum nome. Selecione a tabela inteira na tela do OGMO e cole aqui.");
      return;
    }
    const tem = new Set(lista.map((l) => l.nome));
    setLista([...lista, ...novos.filter((n) => !tem.has(n.nome))]);
    setTexto(""); setErro("");
  }

  return (
    <div className="modal" onClick={onFechar}>
      <div className="dlg" onClick={(e) => e.stopPropagation()}>
        <div className="dlgh"><h3>TPAs do OGMO neste turno</h3>
          <button className="b ghost sm" onClick={onFechar}>fechar</button></div>
        <p className="hint" style={{ marginTop: 0 }}>
          Na tela “Relação de TPAs Escalados”, selecione a tabela, copie e cole abaixo. O app aproveita
          função, matrícula e nome. CF, OP2 e NR35 são ignorados.
        </p>
        <textarea className="i" style={{ width: "100%", height: 118, marginTop: 12,
          fontFamily: "ui-monospace,monospace", fontSize: 12 }}
          value={texto} onChange={(e) => setTexto(e.target.value)}
          placeholder={"Operador III Pátio\t1\t190\t1168\t40839\tJOAO ALBERTO ZIMMERMANN\tOK\tOK\tOK"} />
        {erro && <p className="hint" style={{ color: "#E8B24D" }}>{erro}</p>}
        <button className="b pri" style={{ marginTop: 8 }} onClick={colar}>Ler lista colada</button>

        <div className="campo">
          <label>Ou adicione um de cada vez</label>
          <div className="row" style={{ flexWrap: "wrap" }}>
            <input className="i" style={{ flex: 2, minWidth: 150 }} placeholder="Nome"
              value={nome} onChange={(e) => setNome(e.target.value)} />
            <input className="i" style={{ width: 110 }} placeholder="Matrícula"
              value={mat} onChange={(e) => setMat(e.target.value)} />
            <select className="i" value={fn} onChange={(e) => setFn(e.target.value)}>
              <option value="III">Operador III</option>
              <option value="IV">Operador IV</option>
              <option value="TT">Operador TT</option>
            </select>
            <button className="b" onClick={() => {
              if (!nome.trim()) return;
              setLista([...lista, {
                id: `ogmo-man-${Date.now()}`, nome: nome.trim().toUpperCase(), mat: mat.trim(),
                funcao: fn, funcaoTxt: FUNCOES[fn].nome, ogmo: true, avulso: true, dias: {},
              }]);
              setNome(""); setMat("");
            }}>Adicionar</button>
          </div>
        </div>

        <div style={{ marginTop: 14 }}>
          {lista.length === 0 ? <p className="hint">Nenhum TPA adicionado ainda.</p> :
            lista.map((o, i) => (
              <div key={o.id} className="oc">
                <span className="dot" style={{ background: "var(--cyan)" }} />
                <span style={{ flex: 1 }}>
                  <span style={{ display: "block", fontSize: 13, fontWeight: 600 }}>{o.nome}</span>
                  <span className="dim" style={{ fontSize: 11 }}>
                    {o.funcaoTxt || FUNCOES[o.funcao]?.nome}{o.mat && ` · ${o.mat}`}
                  </span>
                </span>
                <button className="b ghost sm" onClick={() => setLista(lista.filter((_, j) => j !== i))}>tirar</button>
              </div>
            ))}
        </div>

        <div className="row" style={{ marginTop: 16 }}>
          <button className="b pri" onClick={() => onSalvar(lista)}>Salvar no turno</button>
          <button className="b" onClick={onFechar}>Cancelar</button>
        </div>
      </div>
    </div>
  );
}
