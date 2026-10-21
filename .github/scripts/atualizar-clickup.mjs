// Busca cronograma, indicadores e pendências no ClickUp, recalcula e recriptografa a base.
// Segredos do repositório: CLICKUP_TOKEN (token pessoal do ClickUp) e DASH_KEY (chave dos dados).
import { readFile, writeFile } from "node:fs/promises";

const TOKEN = process.env.CLICKUP_TOKEN, KEY_B64 = process.env.DASH_KEY;
if (!TOKEN || !KEY_B64) { console.error("Faltam os segredos CLICKUP_TOKEN e/ou DASH_KEY."); process.exit(1); }

const VIEWS = { cronograma: "8cktdwf-90873", indicadores: "8cktdwf-91333", pendencias: "8cktdwf-98893" };
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
      const n = f.name + " (" + (SUF[f.type] || f.type) + ")";
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

try {
  const wb = paraPlanilha(await tarefasDaView(VIEWS.cronograma));
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
  const I = X.mapIndicadores(paraPlanilha(await tarefasDaView(VIEWS.indicadores)));
  if (!I.riscos.length && !I.qualidade.length && !I.linhaBase.length) throw new Error("nenhum registro reconhecido");
  d.indicadores = Object.assign({}, d.indicadores, I);
  log.push("indicadores: " + I.riscos.length + " riscos, " + I.qualidade.length + " avaliações"); ok++;
} catch (e) { log.push("indicadores NÃO atualizados: " + e.message); }

try {
  const its = X.mapPendencias(paraPlanilha(await tarefasDaView(VIEWS.pendencias)));
  if (!its.length) throw new Error("nenhuma pendência reconhecida");
  d.pendencias = { arquivo: "ClickUp (automático)", quando: hoje, itens: its };
  log.push("pendências: " + its.length); ok++;
} catch (e) { log.push("pendências NÃO atualizadas: " + e.message); }

console.log(log.join("\n"));
if (!ok) process.exit(1);
d.geradoEm = hoje;
d.automatico = { quando: new Date().toISOString(), log };
await fechar("data/dashboard.json.enc", d);
