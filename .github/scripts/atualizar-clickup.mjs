/ Busca cronograma, indicadores e pendências no ClickUp, recalcula e recriptografa a base.
// Segredos do repositório: CLICKUP_TOKEN (token pessoal do ClickUp) e DASH_KEY (chave dos dados).
import { readFile, writeFile } from "node:fs/promises";

const TOKEN = process.env.CLICKUP_TOKEN, KEY_B64 = process.env.DASH_KEY;
if (!TOKEN || !KEY_B64) { console.error("Faltam os segredos CLICKUP_TOKEN e/ou DASH_KEY."); process.exit(1); }

// Fontes padrão. O mapa salvo no dashboard (Base de dados → Mapa das fontes) tem prioridade.
const PADRAO = [
  { tipo: "cronograma", nome: "Cronograma Make It", view: "8cktdwf-90873" },
  { tipo: "cronograma", nome: "Cronograma Gadens", view: "8cktdwf-104773" },
  { tipo: "indicadores", nome: "Indicadores de engenharia", view: "8cktdwf-91333" },
  { tipo: "pendencias", nome: "Pendências de gestão", view: "8cktdwf-98893" }
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
async function tarefasDaView(id) {
  const out = [];
  for (let p = 0; p < 400; p++) {
    const r = await fetch(`https://api.clickup.com/api/v2/view/${id}/task?page=${p}`, { headers: { Authorization: TOKEN } });
    if (r.status === 429) { await espera(60000); p--; continue; }
    if (!r.ok) throw new Error(`ClickUp view ${id}: HTTP ${r.status} ${(await r.text()).slice(0, 200)}`);
    const j = await r.json();
    out.push(...(j.tasks || []));
    if (j.last_page || !(j.tasks || []).length) break;
  }
  return out;
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
  const fixas = ["Task ID", "Task Name", "Parent ID", "Lists", "Status", "Assignee", "Start Date", "Due Date", "Priority", "tags", "Task URL"];
  const extras = [];
  const linhas = tasks.map(t => {
    const r = { "Task ID": t.id, "Task Name": t.name, "Parent ID": t.parent || null, "Lists": t.list && t.list.name,
      "Status": String((t.status && t.status.status) || "").toUpperCase(),
      "Assignee": (t.assignees || []).map(a => a.username).join(", "),
      "Start Date": serial(t.start_date), "Due Date": serial(t.due_date),
      "Priority": t.priority && t.priority.priority, "tags": (t.tags || []).map(g => g.name).join(", "), "Task URL": t.url };
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
try { CFG = await abrir("data/config.json.enc"); if (CFG && Array.isArray(CFG.fontes) && CFG.fontes.some(f => f.view)) FONTES = CFG.fontes.filter(f => f.view && f.ativo !== false); }
catch (e) { /* sem config: usa o padrão */ }
const leitura = []; // status por fonte, exibido no Mapa das fontes
async function ler(f) {
  try { const t = await tarefasDaView(f.view); leitura.push({ view: f.view, nome: f.nome, tipo: f.tipo, n: t.length, erro: null }); return t; }
  catch (e) { leitura.push({ view: f.view, nome: f.nome, tipo: f.tipo, n: 0, erro: e.message.slice(0, 160) }); throw e; }
}
const VIEWS = { cronograma: FONTES.filter(f => f.tipo === "cronograma"),
  indicadores: FONTES.filter(f => f.tipo === "indicadores"), pendencias: FONTES.filter(f => f.tipo === "pendencias") };
async function lerTodas(lista) {
  const vistos = {}, todas = [];
  for (const f of lista) { try { (await ler(f)).forEach(t => { if (!vistos[t.id]) { vistos[t.id] = 1; todas.push(t); } }); } catch (e) {} }
  if (!todas.length) throw new Error("nenhuma visualização respondeu");
  return todas;
}

try {
  const wb = paraPlanilha(await lerTodas(VIEWS.cronograma));
  if (X.detectKind(wb) !== "cronograma") throw new Error("a visualização não traz a coluna Classe Datas Marco");
  d.obras = X.mapCronograma(wb, d.obras, hoje);
  let hist = []; try { hist = await abrir("data/historico.json.enc"); } catch (e) { hist = []; }
  hist = hist.filter(h => h.quando !== hoje).sort((a, b) => a.quando.localeCompare(b.quando));
  const snap = X.snapshotCompl(wb);
  const ref = hist.filter(h => h.quando <= dias(hoje, -7)).pop() || hist[0];
  if (ref) d.atrasos = X.comparaAtrasos(ref.snap, snap, d.obras, ref.quando, hoje);
  hist.push({ quando: hoje, snap });
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
