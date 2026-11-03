// Busca cronograma, indicadores e pendências no ClickUp, recalcula e recriptografa a base.
// Segredos do repositório: CLICKUP_TOKEN (token pessoal do ClickUp) e DASH_KEY (chave dos dados).
import { readFile, writeFile } from "node:fs/promises";

const TOKEN = process.env.CLICKUP_TOKEN, KEY_B64 = process.env.DASH_KEY;
if (!TOKEN || !KEY_B64) { console.error("Faltam os segredos CLICKUP_TOKEN e/ou DASH_KEY."); process.exit(1); }

// Fontes padrão. O mapa salvo no dashboard (Base de dados → Mapa das fontes) tem prioridade.
const PADRAO = [
  { tipo: "indicadores", nome: "Indicadores de engenharia (Linha de Base)", view: "8cktdwf-91333" },
  { tipo: "indicadores", nome: "Riscos por empreendimento", view: "li:901327399298" },
  { tipo: "indicadores", nome: "Desvio de LB de PE", view: "li:901329055549" },
  { tipo: "indicadores", nome: "Alterações de Projeto", view: "li:901327896268" },
  { tipo: "indicadores", nome: "Estouro de obra por conta de projetos", view: "li:901327896274" },
  { tipo: "indicadores", nome: "Economia de orçamento por conta de projetos", view: "li:901327896284" },
  { tipo: "indicadores", nome: "Qualidade de entrega", view: "li:901319437903" },
  { tipo: "pendencias", nome: "Pendências de gestão", view: "8cktdwf-98893" },
  { tipo: "cronograma", nome: "Lume", sigla: "LUM", bu: "Make It", view: "li:901305541456" },
  { tipo: "cronograma", nome: "Essence", sigla: "ESS", bu: "Make It", view: "li:901305651702" },
  { tipo: "cronograma", nome: "Now", sigla: "NOW", bu: "Make It", view: "li:901305540246" },
  { tipo: "cronograma", nome: "Edge", sigla: "EDG", bu: "Make It", view: "li:901305589843" },
  { tipo: "cronograma", nome: "Horizon", sigla: "HOR", bu: "Make It", view: "li:901308225891" },
  { tipo: "cronograma", nome: "Quadratto", sigla: "QUA", bu: "Make It", view: "li:901327828147" },
  { tipo: "cronograma", nome: "Flow", sigla: "FLO", bu: "Make It", view: "li:901305618860" },
  { tipo: "cronograma", nome: "Aura", sigla: "AUR", bu: "Make It", view: "li:901325334354" },
  { tipo: "cronograma", nome: "Alba", sigla: "ALB", bu: "Make It", view: "li:901305538823" },
  { tipo: "cronograma", nome: "Arena", sigla: "ARE", bu: "Make It", view: "li:901327999906" },
  { tipo: "cronograma", nome: "Fundação", sigla: "FUN", bu: "Make It", view: "li:901327965854" },
  { tipo: "cronograma", nome: "Zago", sigla: "ZAG", bu: "Make It", view: "li:901328072536" },
  { tipo: "cronograma", nome: "Al Nasser I", sigla: "AL1", bu: "Make It", view: "li:901326874097" },
  { tipo: "cronograma", nome: "Rafitec", sigla: "RAF", bu: "Make It", view: "li:1400450000002217" },
  { tipo: "cronograma", nome: "Vert", sigla: "VER", bu: "Make It", view: "li:901322097447" },
  { tipo: "cronograma", nome: "Al Nasser II", sigla: "AL2", bu: "Make It", view: "li:901327104307" },
  { tipo: "cronograma", nome: "Martim Afonso", sigla: "MAR", bu: "Gadens Inc", view: "li:901327721107" },
  { tipo: "cronograma", nome: "Core Evo (João Negrão)", sigla: "EVO", bu: "Gadens Inc", view: "li:901328360960" },
  { tipo: "cronograma", nome: "Guadalupe", sigla: "GUA", bu: "Gadens Inc", view: "li:1400450000008729" },
  { tipo: "cronograma", nome: "General Carneiro", sigla: "GEN", bu: "Gadens Inc", view: "f:901310141418" },
  { tipo: "cronograma", nome: "Pe Agostinho", sigla: "PAG", bu: "Gadens Inc", view: "f:901310219261" }
];
const X = await import("data:text/javascript;base64," + (await readFile("data/xlsx.js")).toString("base64"));

const subtle = globalThis.crypto.subtle;
const key = await subtle.importKey("raw", Buffer.from(KEY_B64.trim(), "base64"), { name: "AES-GCM" }, false, ["encrypt", "decrypt"]);
const b64 = b => Buffer.from(b).toString("base64");
async function abrir(p) {
  const pkt = JSON.parse(await readFile(p, "utf8"));
  const pt = await subtle.decrypt({ name: "AES-GCM", iv: Buffer.from(pkt.iv, "base64") }, key, Buffer.from(pkt.ct, "base64"));
  return JSON.parse(Buffer.from(pt).toString("utf8"));
}
async function fechar(p, obj) {
  const iv = globalThis.crypto.getRandomValues(new Uint8Array(12));
  const ct = await subtle.encrypt({ name: "AES-GCM", iv }, key, Buffer.from(JSON.stringify(obj)));
  await writeFile(p, JSON.stringify({ v: 1, iv: b64(iv), ct: b64(ct) }));
}

const espera = ms => new Promise(r => setTimeout(r, ms));
async function api(url) {
  for (;;) {
    const r = await fetch("https://api.clickup.com/api/v2/" + url, { headers: { Authorization: TOKEN } });
    if (r.status === 429) { await espera(60000); continue; }
    if (!r.ok) throw new Error(`ClickUp ${url.split("?")[0]}: HTTP ${r.status} ${(await r.text()).slice(0, 200)}`);
    return r.json();
  }
}
async function paginas(base) {
  const out = [];
  for (let p = 0; p < 400; p++) {
    const j = await api(base + (base.includes("?") ? "&" : "?") + "page=" + p);
    out.push(...(j.tasks || []));
    if (j.last_page || !(j.tasks || []).length) break;
  }
  return out;
}
// view: "8cktdwf-90873" (visualização) · "li:901305541456" (lista) · "f:901310141418" (pasta)
async function tarefasDaView(id) {
  const LST = "/task?include_closed=true&subtasks=true&archived=false";
  if (id.startsWith("li:")) return paginas("list/" + id.slice(3) + LST);
  if (id.startsWith("f:")) {
    const ls = (await api("folder/" + id.slice(2) + "/list?archived=false")).lists || [];
    const out = []; for (const l of ls) out.push(...await paginas("list/" + l.id + LST)); return out;
  }
  return paginas("view/" + id + "/task");
}

// Converte tarefas da API para o mesmo formato de linhas do export em Excel
const SUF = { drop_down: "drop down", labels: "labels", date: "date", number: "number", currency: "currency",
  text: "text", short_text: "text", checkbox: "checkbox", users: "people", email: "email", url: "url", formula: "formula" };
const serial = ms => (ms == null || ms === "") ? null : Number(ms) / 864e5 + 25569;
function valorCampo(f) {
  const v = f.value; if (v == null || v === "") return null;
  const op = (f.type_config && f.type_config.options) || [];
  if (f.type === "drop_down") { const o = op.find(o => o.orderindex === Number(v) || o.id === v); return o ? (o.name || o.label) : String(v); }
  if (f.type === "labels") return (Array.isArray(v) ? v : [v]).map(id => { const o = op.find(o => o.id === id); return o ? (o.label || o.name) : id; }).join(", ");
  if (f.type === "date") return serial(v);
  if (f.type === "users") return (v || []).map(u => u.username || u.email).join(", ");
  return typeof v === "object" ? JSON.stringify(v) : v;
}
function paraPlanilha(tasks) {
  const fixas = ["Task ID", "Task Name", "Parent ID", "Lists", "Obra Sigla", "Status", "Assignee", "Start Date", "Due Date", "Priority", "tags", "Task URL"];
  const extras = [];
  const linhas = tasks.map(t => {
    const r = { "Task ID": t.id, "Task Name": t.name, "Parent ID": t.parent || null, "Lists": t.list && t.list.name,
      "Status": String((t.status && t.status.status) || "").toUpperCase(),
      "Assignee": (t.assignees || []).map(a => a.username).join(", "),
      "Start Date": serial(t.start_date), "Due Date": serial(t.due_date),
      "Obra Sigla": t.__sigla || null, "Priority": t.priority && t.priority.priority, "tags": (t.tags || []).map(g => g.name).join(", "), "Task URL": t.url };
    (t.custom_fields || []).forEach(f => {
      const n = f.id === "b9709a86-fd65-4ab9-a874-f686b019aeb3" ? "Classe Datas Marco (drop down)" : f.name + " (" + (SUF[f.type] || f.type) + ")";
      if (fixas.indexOf(n) < 0 && extras.indexOf(n) < 0) extras.push(n);
      r[n] = valorCampo(f);
    });
    return r;
  });
  const H = fixas.concat(extras);
  return { Tasks: [H].concat(linhas.map(r => H.map(h => r[h] ?? null))) };
}

const hoje = new Date(Date.now() - 3 * 3600e3).toISOString().slice(0, 10); // horário de Brasília
const dias = (s, n) => new Date(Date.parse(s) + n * 864e5).toISOString().slice(0, 10);
const d = await abrir("data/dashboard.json.enc");
const log = []; let ok = 0;
let FONTES = PADRAO;
let CFG = null;
try { CFG = await abrir("data/config.json.enc"); if (CFG && Array.isArray(CFG.fontes) && CFG.fontes.some(f => f.view)) { FONTES = CFG.fontes.filter(f => f.view && f.ativo !== false && ["8cktdwf-90873", "8cktdwf-104773"].indexOf(f.view) < 0);
    if (!CFG.fontes.some(f => f.sigla)) FONTES = FONTES.concat(PADRAO.filter(p => p.sigla && !FONTES.some(f => f.view === p.view)));
    if (!CFG.fontes.some(f => /^li:|^f:/.test(f.view || "") && f.tipo === "indicadores")) FONTES = FONTES.concat(PADRAO.filter(p => p.tipo === "indicadores" && /^li:/.test(p.view) && !FONTES.some(f => f.view === p.view))); } }
catch (e) { /* sem config: usa o padrão */ }
const leitura = []; // status por fonte, exibido no Mapa das fontes
async function ler(f) {
  try { const t = await tarefasDaView(f.view); if (f.sigla) t.forEach(x => { x.__sigla = String(f.sigla).toUpperCase(); }); leitura.push({ view: f.view, nome: f.nome, tipo: f.tipo, n: t.length, erro: null }); return t; }
  catch (e) { leitura.push({ view: f.view, nome: f.nome, tipo: f.tipo, n: 0, erro: e.message.slice(0, 160) }); throw e; }
}
const VIEWS = { cronograma: FONTES.filter(f => f.tipo === "cronograma"),
  indicadores: FONTES.filter(f => f.tipo === "indicadores"), pendencias: FONTES.filter(f => f.tipo === "pendencias") };
async function lerTodas(lista) {
  const vistos = {}, todas = [];
  for (const f of lista) { try { (await ler(f)).forEach(t => { const v = vistos[t.id]; if (!v) { vistos[t.id] = t; todas.push(t); } else if (t.__sigla && !v.__sigla) v.__sigla = t.__sigla; }); } catch (e) {} }
  if (!todas.length) throw new Error("nenhuma visualização respondeu");
  return todas;
}

try {
  const wb = paraPlanilha(await lerTodas(VIEWS.cronograma));
  if (X.detectKind(wb) !== "cronograma") throw new Error("a visualização não traz a coluna Classe Datas Marco");
  // obras cadastradas no Mapa das fontes que ainda não existem na base
  VIEWS.cronograma.filter(f => f.sigla).forEach(f => {
    const sg = String(f.sigla).toUpperCase(), o = d.obras.find(x => x.sigla === sg);
    if (!o) { d.obras.push({ sigla: sg, nome: f.nome || sg, bu: f.bu || "Make It" }); log.push("obra nova: " + sg); }
    else if (f.bu && !o.bu) o.bu = f.bu;
  });
  d.obras = X.mapCronograma(wb, d.obras, hoje);
  let hist = []; try { hist = await abrir("data/historico.json.enc"); } catch (e) { hist = []; }
  hist = hist.filter(h => h.quando !== hoje).sort((a, b) => a.quando.localeCompare(b.quando));
  const snap = X.snapshotCompl(wb);
  const ref = hist.filter(h => h.quando <= dias(hoje, -7)).pop() || hist[0];
  if (ref) d.atrasos = X.comparaAtrasos(ref.snap, snap, d.obras, ref.quando, hoje);
  hist.push({ quando: hoje, snap });
  const alt7 = X.alteracoesJanela(hist, d.obras, hoje, 7); if (alt7) d.alteracoes7 = alt7;
  await fechar("data/historico.json.enc", hist.filter(h => h.quando >= dias(hoje, -14)));
  log.push("cronograma: " + (wb.Tasks.length - 1) + " tarefas"); ok++;
} catch (e) { log.push("cronograma NÃO atualizado: " + e.message); }

try {
  const I = X.mapIndicadores(paraPlanilha(await lerTodas(VIEWS.indicadores)));
  if (!I.riscos.length && !I.qualidade.length && !I.linhaBase.length) throw new Error("nenhum registro reconhecido");
  d.indicadores = Object.assign({}, d.indicadores, I);
  log.push("indicadores: " + I.riscos.length + " riscos, " + I.qualidade.length + " avaliações"); ok++;
} catch (e) { log.push("indicadores NÃO atualizados: " + e.message); }

try {
  const its = X.mapPendencias(paraPlanilha(await lerTodas(VIEWS.pendencias)));
  if (!its.length) throw new Error("nenhuma pendência reconhecida");
  d.pendencias = { arquivo: "ClickUp (automático)", quando: hoje, itens: its };
  log.push("pendências: " + its.length); ok++;
} catch (e) { log.push("pendências NÃO atualizadas: " + e.message); }

// Mantém o Supabase ativo: o plano gratuito pausa o projeto após 7 dias sem uso.
if (CFG && CFG.url && CFG.anon) {
  try {
    const r = await fetch(CFG.url.replace(/\/$/, "") + "/rest/v1/tap_base?select=id&limit=1",
      { headers: { apikey: CFG.anon, Authorization: "Bearer " + CFG.anon, "x-tap-chave": CFG.chave || "" } });
    log.push("supabase: " + (r.ok ? "ativo" : "respondeu " + r.status + (r.status === 540 || r.status >= 500 ? " (projeto pausado? use Resume project)" : "")));
  } catch (e) { log.push("supabase: sem resposta (" + e.message.slice(0, 80) + ")"); }
}
console.log(log.join("\n"));
if (!ok) process.exit(1);
d.geradoEm = hoje;
d.fontesUltima = Object.assign({}, d.fontesUltima);
leitura.forEach(l => { if (!l.erro && l.n) d.fontesUltima[l.view] = hoje; });
d.automatico = { quando: new Date().toISOString(), log, leitura };
await fechar("data/dashboard.json.enc", d);
