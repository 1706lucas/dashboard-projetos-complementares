// Lançamento Financeiro: pedidos e aditivos com parcelas, controle de budget e contrato,
// formulário de estouro, e-mail de pedido e exportação no padrão da planilha SAÍDAS.
// Módulo isolado: erro aqui não afeta as outras abas. Gravação feita pelo painel (opts.api).

export const BUS = [{ bu: "Make It", nome: "Make It" }, { bu: "Gadens Inc", nome: "Gadens" }];
export const CAUSAS = ["Erro ou omissão de projeto", "Alteração de escopo", "Exigência legal ou de concessionária", "Orçamento subdimensionado",
  "Solicitação comercial ou de produto", "Reajuste do fornecedor", "Retrabalho por terceiros", "Outro"];
export const EMAIL_PADRAO = {
  para: "suprimentos@grupogadens.com.br", cc: "lucas@grupogadens.com.br",
  titulo: "Solicitação [Nº PEDIDO] - [OBRA]",
  corpo: "Solicitação [Nº PEDIDO], referente ao [ADITIVO/PEDIDO] [CAUSA RAIZ DO PEDIDO] - [OBRA].\n\nEMPREENDIMENTO: [OBRA]\nFORNECEDOR: [NOME FORNECEDOR]\nCNPJ: [Nº CNPJ]\nVALOR: [VALOR TOTAL]\nCOMPOSIÇÃO PARCELAS: [Nº DE PARCELAS] parcela(s)\n[PARCELAS]\n[AVISO ESTOURO]"
};
export const CAMPOS = ["[Nº PEDIDO]", "[PEDIDO ERP]", "[Nº CADASTRO]", "[OBRA]", "[ADITIVO/PEDIDO]", "[CAUSA RAIZ DO PEDIDO]", "[DISCIPLINA]", "[NOME FORNECEDOR]", "[Nº CNPJ]", "[VALOR TOTAL]",
  "[VALOR PARCELA]", "[Nº DE PARCELAS]", "[PARCELAS]", "[SPE]", "[CNPJ OBRA]", "[GDS]", "[ANALISTA]", "[DATA]", "[OBSERVAÇÃO]", "[AVISO ESTOURO]"];

const norm = s => String(s || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().replace(/ /g, " ")
  .replace(/\bii\b/g, "2").replace(/\bi\b/g, "1").replace(/[^a-z0-9]+/g, " ").trim();
const tok = s => norm(s).split(" ")[0] || "";
const copia = o => JSON.parse(JSON.stringify(o));
let seq = 0; const novoId = p => p + Date.now().toString(36) + (seq++).toString(36) + Math.random().toString(36).slice(2, 5);
const num = v => { if (v == null || v === "") return null; if (typeof v === "number") return v;
  const s = String(v).replace(/[^\d.,-]/g, ""); const x = parseFloat(s.indexOf(",") >= 0 ? s.replace(/\./g, "").replace(",", ".") : s); return isNaN(x) ? null : x; };
const moeda = v => v == null ? "—" : v.toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
const br = d => d ? d.slice(0, 10).split("-").reverse().join("/") : "—";
const hojeISO = () => new Date(Date.now() - 3 * 3600e3).toISOString().slice(0, 10);
const serial = v => { const n = parseFloat(v); return isNaN(n) || n < 20000 ? null : new Date(Date.UTC(1899, 11, 30) + Math.round(n) * 864e5).toISOString().slice(0, 10); };
const somaMes = (iso, k) => { const d = new Date(iso + "T12:00:00"); d.setMonth(d.getMonth() + k); return d.toISOString().slice(0, 10); };

// ── leitor de .xlsx tolerante a abas enormes (para quando a coluna A fica 300 linhas seguidas vazia) ──
export async function lerPlanilha(file) {
  const buf = new Uint8Array(await file.arrayBuffer()), v = new DataView(buf.buffer), dec = new TextDecoder();
  let e = buf.length - 22; while (e >= 0 && v.getUint32(e, true) !== 0x06054b50) e--; if (e < 0) throw new Error("arquivo não é um .xlsx válido");
  const n = v.getUint16(e + 10, true); let p = v.getUint32(e + 16, true); const ent = {};
  for (let i = 0; i < n; i++) { const met = v.getUint16(p + 10, true), csz = v.getUint32(p + 20, true), ln = v.getUint16(p + 28, true), lx = v.getUint16(p + 30, true), lc = v.getUint16(p + 32, true), off = v.getUint32(p + 42, true);
    ent[dec.decode(buf.subarray(p + 46, p + 46 + ln))] = { met, off, csz }; p += 46 + ln + lx + lc; }
  const ler = async nome => { const x = ent[nome]; if (!x) return ""; const ini = x.off + 30 + v.getUint16(x.off + 26, true) + v.getUint16(x.off + 28, true), raw = buf.subarray(ini, ini + x.csz);
    return x.met === 0 ? dec.decode(raw) : await new Response(new Blob([raw]).stream().pipeThrough(new DecompressionStream("deflate-raw"))).text(); };
  const un = t => t.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&#(\d+);/g, (m, c) => String.fromCharCode(+c)).replace(/&amp;/g, "&");
  const sst = [...(await ler("xl/sharedStrings.xml")).matchAll(/<si>([\s\S]*?)<\/si>/g)].map(m => un([...m[1].matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map(t => t[1]).join("")));
  const rels = {}; for (const m of (await ler("xl/_rels/workbook.xml.rels")).matchAll(/<Relationship [^>]*>/g)) { const id = (m[0].match(/Id="([^"]+)"/) || [])[1], tg = (m[0].match(/Target="([^"]+)"/) || [])[1]; if (id) rels[id] = String(tg).replace(/^\/?xl\//, ""); }
  const wb = {};
  for (const m of (await ler("xl/workbook.xml")).matchAll(/<sheet [^>]*\/>/g)) {
    const nome = un((m[0].match(/name="([^"]+)"/) || [])[1] || ""), rid = (m[0].match(/r:id="([^"]+)"/) || [])[1]; if (!rels[rid]) continue;
    const xml = await ler("xl/" + rels[rid]), rows = []; let pos = 0, vazias = 0;
    while (true) { const a = xml.indexOf("<row", pos); if (a < 0) break; const fe = xml.indexOf(">", a); if (xml[fe - 1] === "/") { pos = fe + 1; continue; }
      const b = xml.indexOf("</row>", fe); if (b < 0) break; const corpo = xml.slice(fe + 1, b); pos = b + 6; const cells = [];
      for (const c of corpo.matchAll(/<c ([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) { const at = c[1], ref = (at.match(/r="([A-Z]+)\d+"/) || [])[1]; if (!ref) continue;
        const t = (at.match(/t="(\w+)"/) || [])[1], bd = c[2] || ""; let val = null;
        if (t === "inlineStr") val = un([...bd.matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map(x => x[1]).join(""));
        else { const vm = bd.match(/<v>([\s\S]*?)<\/v>/); if (vm) val = t === "s" ? sst[+vm[1]] : un(vm[1]); }
        let k = 0; for (const ch of ref) k = k * 26 + ch.charCodeAt(0) - 64; if (val != null && val !== "") cells[k - 1] = val; }
      if (cells[0] == null || String(cells[0]).trim() === "") { if (++vazias > 300 && rows.length) break; } else vazias = 0;
      rows.push(cells); }
    wb[nome] = rows; }
  return wb;
}

// ── importação da planilha SAÍDAS (layout Make It ou Gadens) ──
const H = (row, ...res) => { const r = (row || []).map(x => norm(x)); for (const re of res) { const i = r.findIndex(x => re.test(x)); if (i >= 0) return i; } return -1; };
const achaCab = (rows, re) => { for (let i = 0; i < Math.min(rows.length, 8); i++) if ((rows[i] || []).some(x => re.test(norm(x)))) return i; return 0; };
export function importarSaidas(wb, bu, baseAtual) {
  const key = re => Object.keys(wb).find(k => re.test(k));
  const base = baseAtual ? copia(baseAtual) : baseVazia();
  const avisos = [];
  const apAnt = {}; base.obras.filter(o => o.bu === bu).forEach(o => { apAnt[norm(o.nome)] = o.apelidos || []; });
  base.obras = base.obras.filter(o => o.bu !== bu); base.contratos = base.contratos.filter(c => c.bu !== bu); base.historico = base.historico.filter(h => h.bu !== bu);
  const fornPorNome = {}; base.fornecedores.forEach(f => { fornPorNome[norm(f.nome)] = f; if (f.fantasia) fornPorNome[norm(f.fantasia)] = f; if (f.razao) fornPorNome[norm(f.razao)] = f; });
  const porCnpj = cnpj => { const c = String(cnpj || "").replace(/\D/g, ""); return c.length >= 11 ? base.fornecedores.find(f => String(f.cnpj || "").replace(/\D/g, "") === c) : null; };
  const forn = (nome, cnpj) => { const n = norm(nome); let f = porCnpj(cnpj); if (!n && !f) return null;
    f = f || fornPorNome[n] || base.fornecedores.find(x => tok(x.nome) === tok(nome) || (x.fantasia && tok(x.fantasia) === tok(nome)));
    if (!f) { f = { id: novoId("f"), nome: String(nome).replace(/ /g, " ").trim(), fantasia: "", razao: "", cnpj: "", mega: "", contato: "", telefone: "", email: "" }; base.fornecedores.push(f); }
    if (cnpj && !f.cnpj) f.cnpj = String(cnpj).trim(); if (n) fornPorNome[n] = f; return f; };
  // fornecedores
  const kc = key(/CNPJ CONTATOS/i);
  if (kc) { const rows = wb[kc], h = achaCab(rows, /^empresa$/), cab = rows[h] || [];
    const iE = H(cab, /^empresa$/), iF = H(cab, /^fantasia$/), iC = H(cab, /^cnpj$/), iM = H(cab, /mega|cadastro/), iCt = H(cab, /^contato$/), iT = H(cab, /^telefone$/), iEm = H(cab, /^e ?mail$/);
    rows.slice(h + 1).forEach(r => { const razao = r[iE], fant = iF >= 0 ? r[iF] : null; if (!razao && !fant) return;
      const f = forn(fant || razao, iC >= 0 ? r[iC] : null), tx = i => i >= 0 && r[i] != null ? String(r[i]).replace(/ /g, " ").trim() : "";
      if (razao && !f.razao) f.razao = String(razao).replace(/ /g, " ").trim(); if (!f.fantasia) f.fantasia = tx(iF);
      ["cnpj", "mega", "contato", "telefone", "email"].forEach((k, j) => { const v = tx([iC, iM, iCt, iT, iEm][j]); if (v && !f[k]) f[k] = v; }); }); }
  // obras
  const kcnpj = Object.keys(wb).find(k => /^CNPJ/i.test(k) && !/CONTATOS/i.test(k)), kor = key(/OR[ÇC]AMENTOS/i);
  const novaObra = (o) => { o.apelidos = apAnt[norm(o.nome)] || APELIDOS[norm(o.nome)] || []; base.obras.push(o); };
  if (kcnpj) { const rows = wb[kcnpj], cab = rows[0] || [];
    const ix = { n: H(cab, /empreendimento/), g: H(cab, /^gds$/), e: H(cab, /^empresa$/), c: H(cab, /^cnpj$/), en: H(cab, /endereco/), u: H(cab, /^unidades$/), m: H(cab, /m2 construido|m construido/), v: H(cab, /^vgv$/), b: H(cab, /budget/) };
    rows.slice(1).forEach(r => { if (!r[ix.n]) return; const t = i => i >= 0 && r[i] != null ? String(r[i]).trim() : "";
      novaObra({ id: novoId("o"), bu, nome: t(ix.n), gds: t(ix.g), spe: t(ix.e), cnpj: t(ix.c), endereco: t(ix.en), uh: num(r[ix.u]), m2: num(r[ix.m]), vgv: num(r[ix.v]), budget: num(r[ix.b]) }); }); }
  else if (kor) wb[kor].slice(1).forEach(r => { if (!r[0] || num(r[4]) == null && num(r[1]) == null) return;
    novaObra({ id: novoId("o"), bu, nome: String(r[0]).trim(), gds: "", spe: "", cnpj: "", endereco: "", uh: num(r[1]), m2: num(r[2]), vgv: num(r[3]), budget: num(r[4]) }); });
  if (!base.obras.some(o => o.bu === bu)) throw new Error("não encontrei a aba de obras (CNPJ MAKE IT ou ORÇAMENTOS)");
  const obraDe = nome => achaObra(base, bu, nome), semObra = {};
  const contrato = (obra, disc, f, valor, extra) => {
    let c = base.contratos.find(x => x.obraId === obra.id && norm(x.disciplina) === norm(disc) && x.fornecedorId === (f && f.id));
    if (!c) { c = { id: novoId("c"), bu, obraId: obra.id, disciplina: String(disc || "").replace(/\s+/g, " ").trim(), fornecedorId: f ? f.id : null, valor: valor || 0, pedido: "" }; base.contratos.push(c); }
    else if (valor != null) c.valor = Math.max(c.valor || 0, valor);
    if (extra && extra.pedido && !c.pedido) c.pedido = extra.pedido;
    return c; };
  const kct = key(/^contratado$/i), kn = key(/^NF$/i), kg = key(/^geral$/i);
  if (kct) { const rows = wb[kct], cab = rows[0] || [];
    const ix = { o: H(cab, /empreendimento/), p: H(cab, /^projeto$|disciplina/), f: H(cab, /^contratado/), c: H(cab, /^cnpj$/), m: H(cab, /cadastro/), v: H(cab, /valor total|^valor$/), pd: H(cab, /pedido/) };
    rows.slice(1).forEach(r => { if (!r[ix.o]) return; const o = obraDe(r[ix.o]); if (!o) { semObra[r[ix.o]] = 1; return; }
      const f = forn(r[ix.f], ix.c >= 0 ? r[ix.c] : null); if (f && ix.m >= 0 && r[ix.m] && !f.mega) f.mega = String(r[ix.m]).trim();
      contrato(o, r[ix.p], f, num(r[ix.v]), { pedido: ix.pd >= 0 && r[ix.pd] != null ? String(r[ix.pd]).trim() : "" }); }); }
  const hist = (o, disc, fnome, valor, data, status) => {
    const f = forn(fnome), c = base.contratos.find(x => x.obraId === o.id && norm(x.disciplina) === norm(disc) && (!f || x.fornecedorId === f.id)) ||
      base.contratos.find(x => x.obraId === o.id && norm(x.disciplina) === norm(disc));
    base.historico.push({ id: novoId("h"), bu, obraId: o.id, disciplina: String(disc || "").replace(/\s+/g, " ").trim(), fornecedorId: f ? f.id : (c ? c.fornecedorId : null),
      contratoId: c ? c.id : null, valor: valor || 0, data, status: String(status || "").trim().toUpperCase() || "" }); };
  if (kn) { const rows = wb[kn], cab = rows[0] || [];
    const ix = { o: H(cab, /empreendimento/), d: H(cab, /disciplina|projeto/), f: H(cab, /^contratado/), v: H(cab, /valor/), dt: H(cab, /data/), s: H(cab, /^status$/) };
    rows.slice(1).forEach(r => { if (!r[ix.o]) return; const o = obraDe(r[ix.o]); if (!o) { semObra[r[ix.o]] = 1; return; } hist(o, r[ix.d], r[ix.f], num(r[ix.v]), serial(r[ix.dt]), r[ix.s]); }); }
  else if (kg) wb[kg].slice(1).forEach(r => { if (!r[0]) return; const o = obraDe(r[0]); if (!o) { semObra[r[0]] = 1; return; }
    if (num(r[3]) != null) contrato(o, r[1], forn(r[2]), num(r[3])); hist(o, r[1], r[2], num(r[4]), serial(r[5]), r[6]); });
  const sem = Object.keys(semObra); if (sem.length) avisos.push("Linhas sem obra cadastrada (ignoradas): " + sem.join(", ") + ". Cadastre o nome como apelido da obra e importe de novo.");
  base.importado = Object.assign({}, base.importado, { [bu]: new Date().toISOString().slice(0, 10) });
  return { base, avisos, n: { obras: base.obras.filter(o => o.bu === bu).length, contratos: base.contratos.filter(c => c.bu === bu).length, historico: base.historico.filter(h => h.bu === bu).length } };
}
// nomes de cadastro conhecidos (planilha × painel)
const APELIDOS = { "leilao": ["Quadratto"], "alta tensao": ["Aura"], "follador": ["Vert (Follador)"], "core silva": ["Core Silva Jardim", "CORE Silva Jardim"], "padre agostinho": ["Pe. Agostinho", "Pe Agostinho"],
  "joao negrao": ["Core Evo", "Core Evo (João Negrão)"], "core palacio": ["Core Palácio"], "al nasser 1": ["AL NASSER 1"], "al nasser 2": ["AL NASSER 2"] };
export function baseVazia() { return { versao: 1, obras: [], fornecedores: [], contratos: [], historico: [], vinculos: {}, email: copia(EMAIL_PADRAO), importado: {} }; }
export function achaObra(base, bu, nome) {
  const n = norm(nome); if (!n) return null; const os = base.obras.filter(o => !bu || o.bu === bu);
  return os.find(o => norm(o.nome) === n) || os.find(o => (o.apelidos || []).some(a => norm(a) === n)) ||
    (os.filter(o => norm(o.nome).indexOf(n) === 0 || n.indexOf(norm(o.nome)) === 0).length === 1 ? os.find(o => norm(o.nome).indexOf(n) === 0 || n.indexOf(norm(o.nome)) === 0) : null);
}

// ── cálculos ──
export function linhas(base, lancs) {
  const L = base.historico.map(h => Object.assign({ origem: "planilha" }, h));
  lancs.filter(l => l.status === "finalizado").forEach(l => (l.parcelas || []).forEach((p, i) => L.push({ origem: "painel", id: l.id + ":" + i, lancId: l.id, numero: l.numero,
    bu: l.bu, obraId: l.obraId, disciplina: l.disciplina, fornecedorId: l.fornecedorId, contratoId: l.contratoId, valor: num(p.valor) || 0, data: p.data, status: p.status || "PREVISTO", tipo: l.tipo, parcela: i })));
  return L;
}
export function situacaoObra(base, lancs, obraId, extra) {
  const o = base.obras.find(x => x.id === obraId); if (!o) return null;
  const usado = linhas(base, lancs).filter(x => x.obraId === obraId).reduce((a, x) => a + x.valor, 0) + (extra || 0);
  return { budget: o.budget, usado, saldo: o.budget != null ? o.budget - usado : null };
}
export function situacaoContrato(base, lancs, contratoId, extra) {
  const c = base.contratos.find(x => x.id === contratoId); if (!c) return null;
  const usado = linhas(base, lancs).filter(x => x.contratoId === contratoId).reduce((a, x) => a + x.valor, 0) + (extra || 0);
  return { valor: c.valor, usado, saldo: c.valor - usado };
}
export function atrasadas(base, lancs) {
  const h = hojeISO();
  return linhas(base, lancs).filter(x => x.data && x.data < h && /^(PREVISTO|PREVISÃO|A PAGAR|PENDENTE)$/.test(x.status));
}
function resumoDe(base, lancs, est) { return { atrasadas: atrasadas(base, lancs).length, estPend: est.filter(e => e.status === "pendente").length }; }
export { resumoDe as resumo };

// ── e-mail ──
export function montaEmail(base, l, link) {
  const o = base.obras.find(x => x.id === l.obraId) || {}, f = base.fornecedores.find(x => x.id === l.fornecedorId) || {};
  const ps = l.parcelas || [], tot = ps.reduce((a, p) => a + (num(p.valor) || 0), 0);
  const est = l.estouro && (l.estouro.contrato > 0 || l.estouro.obra > 0) ? "\nATENÇÃO: este lançamento excede " +
    [l.estouro.contrato > 0 ? "o contrato em " + moeda(l.estouro.contrato) : "", l.estouro.obra > 0 ? "o budget da obra em " + moeda(l.estouro.obra) : ""].filter(Boolean).join(" e ") +
    ".\nFormulário de composição do estouro: " + link + "\n" : "";
  const ct = base.contratos.find(x => x.id === l.contratoId) || {};
  const M = { "[Nº PEDIDO]": l.numero || "XX", "[PEDIDO ERP]": ct.pedido || "", "[Nº CADASTRO]": f.mega || "", "[OBRA]": o.nome || "", "[ADITIVO/PEDIDO]": l.tipo === "Aditivo" ? "ADITIVO" : "PEDIDO", "[CAUSA RAIZ DO PEDIDO]": l.causa || "",
    "[DISCIPLINA]": l.disciplina || "", "[NOME FORNECEDOR]": f.nome || "", "[Nº CNPJ]": f.cnpj || "(CNPJ não cadastrado)", "[VALOR TOTAL]": moeda(tot),
    "[VALOR PARCELA]": ps.length === 1 ? moeda(num(ps[0].valor)) : ps.map(p => moeda(num(p.valor))).join(" + "), "[Nº DE PARCELAS]": String(ps.length),
    "[PARCELAS]": ps.map((p, i) => (i + 1) + "ª parcela: " + moeda(num(p.valor)) + " · pagamento em " + br(p.data)).join("\n"),
    "[SPE]": o.spe || "", "[CNPJ OBRA]": o.cnpj || "", "[GDS]": o.gds || "", "[ANALISTA]": l.criadoPor || "", "[DATA]": br(hojeISO()), "[OBSERVAÇÃO]": l.obs || "", "[AVISO ESTOURO]": est };
  const troca = t => CAMPOS.reduce((s, k) => s.split(k).join(M[k] != null ? M[k] : ""), String(t || ""));
  const e = base.email || EMAIL_PADRAO;
  return { para: troca(e.para), cc: troca(e.cc), titulo: troca(e.titulo), corpo: troca(e.corpo).replace(/\n{3,}/g, "\n\n").trim() };
}
export const mailto = m => "mailto:" + encodeURIComponent(m.para).replace(/%40/g, "@").replace(/%2C/g, ",") + "?" +
  (m.cc ? "cc=" + encodeURIComponent(m.cc).replace(/%40/g, "@").replace(/%2C/g, ",") + "&" : "") + "subject=" + encodeURIComponent(m.titulo) + "&body=" + encodeURIComponent(m.corpo);

// ── exportação .xlsx (padrão SAÍDAS) ──
let TAB = null;
function crc32(b) { if (!TAB) { TAB = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; TAB[n] = c >>> 0; } }
  let c = 0xffffffff; for (let i = 0; i < b.length; i++) c = TAB[(c ^ b[i]) & 0xff] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; }
function zip(arqs) {
  const enc = new TextEncoder(), partes = [], cent = []; let off = 0;
  arqs.forEach(a => { const nm = enc.encode(a.nome), d = typeof a.data === "string" ? enc.encode(a.data) : a.data, crc = crc32(d);
    const h = new DataView(new ArrayBuffer(30)); h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0x0800, true);
    h.setUint32(14, crc, true); h.setUint32(18, d.length, true); h.setUint32(22, d.length, true); h.setUint16(26, nm.length, true); partes.push(new Uint8Array(h.buffer), nm, d);
    const c = new DataView(new ArrayBuffer(46)); c.setUint32(0, 0x02014b50, true); c.setUint16(4, 20, true); c.setUint16(6, 20, true); c.setUint16(8, 0x0800, true);
    c.setUint32(16, crc, true); c.setUint32(20, d.length, true); c.setUint32(24, d.length, true); c.setUint16(28, nm.length, true); c.setUint32(42, off, true);
    cent.push(new Uint8Array(c.buffer), nm); off += 30 + nm.length + d.length; });
  const tam = cent.reduce((a, x) => a + x.length, 0), fim = new DataView(new ArrayBuffer(22));
  fim.setUint32(0, 0x06054b50, true); fim.setUint16(8, arqs.length, true); fim.setUint16(10, arqs.length, true); fim.setUint32(12, tam, true); fim.setUint32(16, off, true);
  return new Blob(partes.concat(cent, [new Uint8Array(fim.buffer)]), { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}
const xesc = s => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, "");
const colL = i => { let s = ""; i++; while (i) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; };
// tipos de coluna: t = texto, m = moeda, d = data, n = número, p = percentual
function planilha(cab, tipos, rows, larg) {
  const S = { t: 0, m: 3, d: 2, n: 4, p: 5 };
  const cel = (v, t, ref) => { if (v == null || v === "") return '<c r="' + ref + '" s="6"/>';
    if (t === "d") { const ms = Date.parse(String(v).slice(0, 10) + "T00:00:00Z"); return isNaN(ms) ? '<c r="' + ref + '" s="6"/>' : '<c r="' + ref + '" s="2"><v>' + (ms / 864e5 + 25569) + "</v></c>"; }
    if (t !== "t" && typeof v === "number") return '<c r="' + ref + '" s="' + S[t] + '"><v>' + v + "</v></c>";
    return '<c r="' + ref + '" s="6" t="inlineStr"><is><t xml:space="preserve">' + xesc(v) + "</t></is></c>"; };
  const fim = colL(cab.length - 1) + (rows.length + 1);
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<dimension ref="A1:' + fim + '"/><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>' +
    "<cols>" + cab.map((c, i) => '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + (larg[i] || 16) + '" customWidth="1"/>').join("") + "</cols><sheetData>" +
    '<row r="1" ht="22" customHeight="1">' + cab.map((c, i) => '<c r="' + colL(i) + '1" s="1" t="inlineStr"><is><t>' + xesc(c) + "</t></is></c>").join("") + "</row>" +
    rows.map((r, j) => '<row r="' + (j + 2) + '">' + r.map((v, i) => cel(v, tipos[i], colL(i) + (j + 2))).join("") + "</row>").join("") +
    '</sheetData><autoFilter ref="A1:' + fim + '"/></worksheet>';
}
const STYLES = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
  '<numFmts count="2"><numFmt numFmtId="164" formatCode="&quot;R$&quot; #,##0.00"/><numFmt numFmtId="165" formatCode="dd/mm/yyyy"/></numFmts>' +
  '<fonts count="2"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font></fonts>' +
  '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F497D"/><bgColor indexed="64"/></patternFill></fill></fills>' +
  '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FFD9D9D9"/></left><right style="thin"><color rgb="FFD9D9D9"/></right><top style="thin"><color rgb="FFD9D9D9"/></top><bottom style="thin"><color rgb="FFD9D9D9"/></bottom><diagonal/></border></borders>' +
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="7">' +
  '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
  '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>' +
  '<xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>' +
  '<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>' +
  '<xf numFmtId="4" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>' +
  '<xf numFmtId="10" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>' +
  '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>' +
  '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
function pasta(abas) {
  const ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
    abas.map((a, i) => '<Override PartName="/xl/worksheets/sheet' + (i + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>').join("") + "</Types>";
  const wbx = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' +
    abas.map((a, i) => '<sheet name="' + xesc(a.nome) + '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"/>').join("") + "</sheets>" +
    "<definedNames>" + abas.map((a, i) => '<definedName name="_xlnm._FilterDatabase" localSheetId="' + i + '" hidden="1">\'' + xesc(a.nome) + "'!$A$1:$" + colL(a.cab.length - 1) + "$" + (a.rows.length + 1) + "</definedName>").join("") + "</definedNames></workbook>";
  const rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    abas.map((a, i) => '<Relationship Id="rId' + (i + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + (i + 1) + '.xml"/>').join("") +
    '<Relationship Id="rId' + (abas.length + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>';
  const root = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>';
  return zip([{ nome: "[Content_Types].xml", data: ct }, { nome: "_rels/.rels", data: root }, { nome: "xl/workbook.xml", data: wbx }, { nome: "xl/_rels/workbook.xml.rels", data: rels },
    { nome: "xl/styles.xml", data: STYLES }].concat(abas.map((a, i) => ({ nome: "xl/worksheets/sheet" + (i + 1) + ".xml", data: planilha(a.cab, a.tipos, a.rows, a.larg) }))));
}
export function exportar(base, lancs, ests, bu) {
  const O = id => base.obras.find(o => o.id === id) || {}, F = id => base.fornecedores.find(f => f.id === id) || {};
  const obras = base.obras.filter(o => o.bu === bu), L = linhas(base, lancs).filter(x => x.bu === bu);
  const usado = id => L.filter(x => x.obraId === id).reduce((a, x) => a + x.valor, 0);
  const nomeBu = (BUS.find(b => b.bu === bu) || {}).nome || bu;
  const abas = [
    { nome: "CNPJ " + nomeBu.toUpperCase(), cab: ["NOME DO EMPREENDIMENTO", "GDS", "EMPRESA", "CNPJ", "ENDEREÇO", "UNIDADES", "m² CONSTRUÍDO", "VGV", "BUDGET DISP.", "% VGV", "VALOR POR UNIDADE", "VALOR POR m²", "VALOR UTILIZADO", "DESVIO"],
      tipos: ["t", "t", "t", "t", "t", "n", "n", "m", "m", "p", "m", "m", "m", "p"], larg: [22, 9, 34, 20, 40, 10, 13, 16, 15, 9, 15, 13, 16, 10],
      rows: obras.map(o => { const u = usado(o.id); return [o.nome, o.gds, o.spe, o.cnpj, o.endereco, o.uh, o.m2, o.vgv, o.budget, o.budget && o.vgv ? o.budget / o.vgv : null,
        o.budget && o.uh ? o.budget / o.uh : null, o.budget && o.m2 ? o.budget / o.m2 : null, u, o.budget ? u / o.budget - 1 : null]; }) },
    { nome: "CNPJ CONTATOS", cab: ["EMPRESA", "FANTASIA", "CNPJ", "CADASTRO NO MEGA", "Contato", "Telefone", "Email"], tipos: ["t", "t", "t", "t", "t", "t", "t"], larg: [36, 24, 20, 14, 22, 18, 32],
      rows: base.fornecedores.map(f => [f.razao || f.nome, f.fantasia || f.nome, f.cnpj, f.mega || "", f.contato, f.telefone, f.email]) },
    { nome: "NF", cab: ["Empreendimento", "Disciplina", "Contratado", "Valor da NF", "Data de pagamento programado", "Status", "Mês de Pagamento", "Ano de Pagamento", "Mês/ Ano", "Nº Pedido", "Tipo", "Origem"],
      tipos: ["t", "t", "t", "m", "d", "t", "n", "n", "t", "t", "t", "t"], larg: [20, 22, 24, 14, 14, 12, 10, 10, 10, 10, 10, 10],
      rows: L.sort((a, b) => String(O(a.obraId).nome).localeCompare(String(O(b.obraId).nome)) || String(a.data || "").localeCompare(String(b.data || ""))).map(x => {
        const m = x.data ? +x.data.slice(5, 7) : null, y = x.data ? +x.data.slice(0, 4) : null;
        return [O(x.obraId).nome, x.disciplina, F(x.fornecedorId).nome || "", x.valor, x.data, x.status, m, y, m ? m + "/" + y : "", x.numero ? String(x.numero) : "", x.tipo || "", x.origem]; }) },
    { nome: "Contratado", cab: ["Empreendimento", "Projeto", "Contratado", "Valor Total", "Lançado", "Saldo"], tipos: ["t", "t", "t", "m", "m", "m"], larg: [20, 26, 26, 15, 15, 15],
      rows: base.contratos.filter(c => c.bu === bu).map(c => { const u = L.filter(x => x.contratoId === c.id).reduce((a, x) => a + x.valor, 0); return [O(c.obraId).nome, c.disciplina, F(c.fornecedorId).nome || "", c.valor, u, c.valor - u]; }) },
    { nome: "Motivos de estouro", cab: ["Nº Pedido", "Empreendimento", "Disciplina", "Fornecedor", "Excesso contrato", "Excesso obra", "Causa raiz", "Descrição", "Composição", "Ação", "Analista", "Status", "Respondido em"],
      tipos: ["t", "t", "t", "t", "m", "m", "t", "t", "t", "t", "t", "t", "d"], larg: [10, 18, 18, 22, 14, 14, 26, 40, 40, 30, 18, 12, 13],
      rows: ests.filter(e => e.bu === bu).map(e => { const r = e.resposta || {};
        return [String(e.numero || ""), O(e.obraId).nome, e.disciplina, F(e.fornecedorId).nome || "", e.contrato || null, e.obra || null, r.causa || "", r.descricao || "",
          (r.composicao || []).map(c => c.item + ": " + moeda(num(c.valor))).join("; "), r.acao || "", e.analista || "", e.status, r.em || null]; }) }
  ];
  return { blob: pasta(abas), nome: "SAIDAS - " + nomeBu.toUpperCase() + " - " + hojeISO() + ".xlsx" };
}

// ── interface ──
const CSS = `
.fi{display:flex;flex-direction:column;gap:14px;color:var(--color-text,#0F172A)}
.fi *{box-sizing:border-box}
.fi-card{background:var(--card,#fff);border:1px solid var(--slate-100,#F1F5F9);border-radius:16px;box-shadow:var(--card-shadow);padding:16px 18px}
.fi-l{display:flex;flex-wrap:wrap;gap:8px;align-items:center}.fi-sb{justify-content:space-between}
.fi h3{margin:0;font-size:14.5px;font-weight:800}.fi h4{margin:0 0 8px;font-size:13px;font-weight:800}
.fi button{font:inherit;font-size:12.5px;font-weight:700;cursor:pointer;border:1px solid var(--slate-200,#E2E8F0);background:var(--card,#fff);color:var(--slate-600,#475569);border-radius:10px;padding:7px 12px}
.fi button.on,.fi button.pri{background:var(--blue-600,#2563EB);border-color:var(--blue-600,#2563EB);color:#fff}
.fi button.per{color:#B91C1C}.fi button.mini{padding:3px 8px;font-size:11.5px}
.fi button:disabled{opacity:.45;cursor:default}
.fi select,.fi input,.fi textarea{font:inherit;font-size:13px;padding:7px 9px;border-radius:9px;border:1px solid var(--slate-200,#E2E8F0);background:var(--card,#fff);color:var(--color-text,#0F172A);min-width:0}
.fi textarea{width:100%;resize:vertical}
.fi-grade{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:10px}
.fi-campo{display:flex;flex-direction:column;gap:4px;font-size:11px;font-weight:800;letter-spacing:.04em;text-transform:uppercase;color:var(--slate-400,#94A3B8)}
.fi-campo input,.fi-campo select,.fi-campo textarea{width:100%;text-transform:none;letter-spacing:0;font-weight:500}
.fi-kpi{display:flex;flex-direction:column;gap:2px;padding:12px 14px;border-radius:12px;background:var(--slate-50,#F8FAFC);border-left:4px solid var(--cor,#2563EB)}
.fi-kpi b{font-size:18px;font-weight:900}.fi-kpi span{font-size:11px;color:var(--slate-500,#64748B)}
button.fi-kpi{text-align:left;border:1px solid transparent;border-left:4px solid var(--cor,#2563EB);cursor:pointer}button.fi-kpi:hover{border-color:var(--slate-200,#E2E8F0)}
.fi-pill{font-size:10.5px;font-weight:800;padding:2px 9px;border-radius:99px;white-space:nowrap}
.fi-ok{background:#D1FAE5;color:#065F46}.fi-no{background:#FEE2E2;color:#991B1B}.fi-am{background:#FEF3C7;color:#92400E}.fi-ci{background:var(--slate-100,#F1F5F9);color:var(--slate-600,#475569)}
.fi table{width:100%;border-collapse:collapse;font-size:12.5px}
.fi th{text-align:left;font-size:10.5px;letter-spacing:.05em;text-transform:uppercase;color:var(--slate-400,#94A3B8);padding:6px 8px;border-bottom:1px solid var(--slate-200,#E2E8F0);white-space:nowrap;position:sticky;top:0;background:var(--card,#fff)}
.fi td{padding:6px 8px;border-bottom:1px solid var(--slate-100,#F1F5F9);vertical-align:middle}
.fi td.n{text-align:right;font-family:var(--font-mono,monospace);white-space:nowrap}
.fi tr.clic{cursor:pointer}.fi tr.clic:hover td{background:var(--slate-50,#F8FAFC)}
.fi-tab{overflow:auto;max-height:560px}
.fi-alerta{margin:0;font-size:12.5px;line-height:1.45;padding:10px 13px;border-radius:12px;background:#FEF2F2;color:#991B1B;border:1px solid #FECACA}
.fi-alerta.am{background:#FFFBEB;color:#92400E;border-color:#FDE68A}
.fi-msg{margin:0;font-size:12.5px;padding:9px 12px;border-radius:10px;background:#ECFDF5;color:#065F46}.fi-msg.err{background:#FEF2F2;color:#991B1B}
.fi-aviso{margin:0;font-size:12px;line-height:1.45;color:var(--slate-500,#64748B)}
.fi-parc{display:grid;grid-template-columns:40px 1fr 1fr 130px 34px;gap:6px;align-items:center}
.fi-chip{display:inline-flex;font-size:11.5px;padding:3px 8px;border-radius:8px;border:1px dashed var(--slate-300,#CBD5E1);background:transparent;cursor:pointer}
.fi-modal{position:fixed;inset:0;background:rgba(15,23,42,.45);display:flex;align-items:flex-start;justify-content:center;padding:40px 16px;z-index:9000;overflow:auto}
.fi-modal>.fi-card{width:100%;max-width:820px}
.fi-barra{height:8px;border-radius:99px;background:var(--slate-100,#F1F5F9);overflow:hidden;min-width:80px}.fi-barra i{display:block;height:100%;border-radius:99px}
.fi-dupla{display:grid;grid-template-columns:repeat(auto-fit,minmax(460px,1fr));gap:14px}.fi-dupla>.fi-card{overflow-x:auto;min-width:0}
.fi-filtros{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:8px}
.fi-nav{display:flex;flex-wrap:wrap;gap:6px}
.fi-dias{font-weight:800;color:#B91C1C;white-space:nowrap}
@media(max-width:700px){.fi-parc{grid-template-columns:30px 1fr 1fr}.fi-dupla{grid-template-columns:1fr}}
.fi{--fs1:#2a78d6;--fs2:#eb6834;--fs3:#1baf7a;--fs4:#eda100;--fi-bom:#10B981;--fi-aten:#F59E0B;--fi-crit:#E34948;--fi-neutro:#94A3B8;--fi-trilho:var(--slate-100,#F1F5F9)}
:root[data-tema="escuro"] .fi{--fs1:#3987e5;--fs2:#d95926;--fs3:#199e70;--fs4:#c98500;--fi-crit:#e66767;--fi-trilho:rgba(148,163,184,.18)}
.fi-top{display:flex;flex-direction:column;gap:12px}
.fi-seg{display:inline-flex;flex-wrap:wrap;gap:2px;padding:3px;border-radius:12px;background:var(--slate-100,#F1F5F9)}
.fi .fi-seg button{border:0;background:transparent;border-radius:9px;padding:7px 13px;color:var(--slate-500,#64748B);display:inline-flex;align-items:center;gap:6px}
.fi .fi-seg button.on{background:var(--card,#fff);color:var(--blue-700,#1D4ED8);box-shadow:0 1px 2px rgba(15,23,42,.1)}
.fi .fi-seg.mini button{padding:5px 11px;font-size:12px}
.fi-badge{font-size:10.5px;font-weight:800;min-width:18px;padding:1px 6px;border-radius:99px;background:var(--fi-crit);color:#fff}
.fi .fi-cta{background:var(--blue-600,#2563EB);border-color:var(--blue-600,#2563EB);color:#fff;border-radius:12px;padding:10px 18px;font-size:13.5px;box-shadow:0 4px 14px rgba(37,99,235,.28)}
.fi-fil{display:flex;flex-wrap:wrap;gap:8px;align-items:center}
.fi select.fi-pillsel{border-radius:99px;padding:6px 12px;font-size:12.5px;max-width:240px}.fi select.fi-pillsel.on{border-color:var(--blue-500,#3B82F6);background:var(--blue-50,#EFF6FF);color:var(--blue-700,#1D4ED8);font-weight:700}
.fi .fi-link{border:0;background:transparent;color:var(--blue-600,#2563EB);padding:4px 6px}
.fi-hero{display:grid;grid-template-columns:150px minmax(200px,1fr) minmax(0,2fr);gap:18px;align-items:center;background:var(--card,#fff);border:1px solid var(--slate-100,#F1F5F9);border-radius:20px;box-shadow:var(--card-shadow);padding:18px 22px}
.fi-hero-t{display:flex;flex-direction:column;gap:4px}.fi-hero-r{font-size:12px;font-weight:700;color:var(--slate-500,#64748B)}.fi-hero-v{font-size:30px;font-weight:900;letter-spacing:-.02em;line-height:1.05}
.fi-hero-s{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px}
.fi-anel{width:130px;height:130px}.fi-anel-v{font-size:24px;font-weight:900;fill:var(--color-text,#0F172A)}.fi-anel-r{font-size:10px;fill:var(--slate-500,#64748B)}
.fi-graf{width:100%;height:230px;display:block}.fi-grade-l{stroke:var(--slate-100,#F1F5F9);stroke-width:1}.fi-eixo{font-size:10.5px;fill:var(--slate-400,#94A3B8)}.fi-eixo.hoje,.fi-eixo.forte{fill:var(--color-text,#0F172A);font-weight:800}
.fi-hoje{stroke:var(--slate-400,#94A3B8);stroke-dasharray:3 4;stroke-width:1.2}.fi-ref{stroke:var(--fi-crit);stroke-dasharray:6 5;stroke-width:1.5}
.fi-leg{display:flex;flex-wrap:wrap;gap:12px;font-size:11.5px;color:var(--slate-500,#64748B)}.fi-leg span{display:inline-flex;align-items:center;gap:6px}.fi-leg b{width:10px;height:10px;border-radius:3px}
.fi-bh{display:flex;flex-direction:column;gap:2px}
.fi-bh-l{display:grid;grid-template-columns:minmax(90px,36%) 1fr 92px;gap:10px;align-items:center;padding:5px 6px;border-radius:9px;border:0;background:transparent;text-align:left;font:inherit;color:inherit;width:100%}
button.fi-bh-l{cursor:pointer}button.fi-bh-l:hover{background:var(--slate-50,#F8FAFC)}
.fi-bh-n{font-size:12px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;color:var(--slate-600,#475569)}
.fi-bh-t{position:relative;height:10px;border-radius:99px;background:var(--fi-trilho)}.fi-bh-t i{position:absolute;left:0;top:0;bottom:0;border-radius:99px}
.fi-bh-t em{position:absolute;top:-3px;bottom:-3px;width:2px;margin-left:-1px;background:var(--color-text,#0F172A);opacity:.55;border-radius:2px}
.fi-bh-v{font-size:11.5px;font-weight:700;text-align:right;font-family:var(--font-mono,monospace);white-space:nowrap}
.fi-bus{display:grid;grid-template-columns:1fr 1fr;gap:10px}
.fi-bu{display:flex;align-items:center;gap:12px;padding:10px;border-radius:14px;border:1px solid var(--slate-100,#F1F5F9);background:var(--slate-50,#F8FAFC);cursor:pointer;text-align:left;font:inherit;color:inherit}
.fi-bu .fi-anel{width:86px;height:86px}.fi-bu>div{display:flex;flex-direction:column;gap:2px}
.fi-dica{position:fixed;z-index:99999;pointer-events:none;background:#0F172A;color:#fff;font-size:12px;line-height:1.45;padding:8px 10px;border-radius:9px;box-shadow:0 8px 24px rgba(0,0,0,.25);display:none;max-width:280px}
.fi-lista{display:flex;flex-direction:column;gap:6px}.fi-li{display:grid;grid-template-columns:52px 1fr auto;gap:12px;align-items:center;padding:8px 10px;border-radius:12px;background:var(--slate-50,#F8FAFC)}
.fi-li-d{display:flex;flex-direction:column;align-items:center;line-height:1}.fi-li-d b{font-size:18px;font-weight:900}.fi-li-d span{font-size:10.5px;color:var(--slate-500,#64748B)}
.fi-li>div:nth-child(2){display:flex;flex-direction:column;gap:2px;min-width:0}.fi-li-v{display:flex;flex-direction:column;align-items:flex-end;gap:3px}
.fi-drawer-bg{position:fixed;inset:0;background:rgba(15,23,42,.42);z-index:9000;display:flex;justify-content:flex-end}
.fi-drawer{width:min(600px,100vw);height:100%;background:var(--bg,var(--card,#fff));background:var(--card,#fff);box-shadow:-12px 0 40px rgba(15,23,42,.2);display:flex;flex-direction:column;gap:0;animation:fiIn .18s ease-out}
@keyframes fiIn{from{transform:translateX(40px);opacity:.4}to{transform:none;opacity:1}}
.fi-dw-cab{padding:18px 22px 12px;border-bottom:1px solid var(--slate-100,#F1F5F9);display:flex;flex-direction:column;gap:12px}
.fi-passos{display:grid;grid-template-columns:repeat(4,1fr);gap:6px}.fi .fi-passos button{font-size:11.5px;padding:6px 4px;border-radius:99px;background:var(--slate-100,#F1F5F9);border:0;color:var(--slate-500,#64748B)}
.fi .fi-passos button.on{background:var(--blue-600,#2563EB);color:#fff}.fi .fi-passos button.feito{background:var(--blue-50,#EFF6FF);color:var(--blue-700,#1D4ED8)}
.fi-dw-corpo{flex:1;overflow:auto;padding:18px 22px;display:flex;flex-direction:column;gap:10px}
.fi-dw-rod{display:flex;justify-content:space-between;align-items:center;gap:8px;padding:14px 22px;border-top:1px solid var(--slate-100,#F1F5F9)}
.fi-pergunta{margin:6px 0 2px;font-size:16px;font-weight:900}
.fi-opcoes{display:grid;grid-template-columns:repeat(auto-fill,minmax(230px,1fr));gap:8px}
.fi .fi-op{display:flex;flex-direction:column;align-items:flex-start;gap:5px;text-align:left;padding:12px 14px;border-radius:14px;border:1px solid var(--slate-200,#E2E8F0);background:var(--card,#fff);font-weight:500}
.fi .fi-op b{font-size:13.5px;font-weight:800;color:var(--color-text,#0F172A)}.fi .fi-op.on{border-color:var(--blue-600,#2563EB);box-shadow:0 0 0 3px var(--blue-50,#EFF6FF)}
.fi .fi-op.novo{border-style:dashed}
.fi-mini-barra{display:block;width:100%;height:6px;border-radius:99px;background:var(--fi-trilho);overflow:hidden}.fi-mini-barra i{display:block;height:100%;border-radius:99px}
.fi-dividir{display:grid;grid-template-columns:1.3fr .8fr 1.2fr auto;gap:8px;align-items:end;padding:12px;border-radius:14px;background:var(--slate-50,#F8FAFC)}
.fi-parcelas{display:flex;flex-direction:column;gap:6px}.fi-drawer .fi-parc{grid-template-columns:36px 1fr 1fr 30px}
.fi-resumo{display:grid;grid-template-columns:1fr 1fr;gap:8px}.fi-resumo div{display:flex;flex-direction:column;gap:2px;padding:10px 12px;border-radius:12px;background:var(--slate-50,#F8FAFC)}
.fi-resumo span{font-size:11px;color:var(--slate-500,#64748B)}.fi-resumo b{font-size:13px}
.fi-imp{display:flex;flex-direction:column;gap:6px;padding:10px 0}.fi-imp-b{position:relative;height:14px;border-radius:99px;background:var(--fi-trilho)}
.fi-imp-b i{position:absolute;top:0;bottom:0;border-radius:99px}.fi-imp-b i.a{left:0;background:var(--fs1)}.fi-imp-b em{position:absolute;top:-4px;bottom:-4px;width:2px;background:var(--color-text,#0F172A)}
@media(max-width:900px){.fi-hero{grid-template-columns:1fr;justify-items:center;text-align:center}.fi-hero-s{width:100%}.fi-dividir{grid-template-columns:1fr 1fr}.fi-bus{grid-template-columns:1fr}}
`;
const el = (t, c, x) => { const e = document.createElement(t); if (c) e.className = c; if (x != null) e.textContent = x; return e; };
const bt = (x, f, c) => { const b = el("button", c, x); b.type = "button"; b.addEventListener("click", f); return b; };
const campo = (rot, ctrl) => { const l = el("label", "fi-campo", rot); l.appendChild(ctrl); return l; };
const sel = (ops, val, f) => { const s = el("select"); ops.forEach(o => { const op = el("option", null, o[1]); op.value = o[0]; op.selected = String(o[0]) === String(val == null ? "" : val); s.appendChild(op); });
  s.addEventListener("change", () => f(s.value)); return s; };
const inp = (val, f, ph, tipo) => { const i = el("input"); if (tipo) i.type = tipo; i.value = val == null ? "" : val; if (ph) i.placeholder = ph; i.addEventListener("input", () => f(i.value)); return i; };
const kpi = (rot, v, cor, sub, go) => { const k = el(go ? "button" : "div", "fi-kpi"); if (go) { k.type = "button"; k.addEventListener("click", go); }
  k.style.setProperty("--cor", cor); k.appendChild(el("span", null, rot)); k.appendChild(el("b", null, v)); if (sub) k.appendChild(el("span", null, sub)); return k; };
const tabela = (cab, linhasTr) => { const t = el("table"), h = el("tr"); cab.forEach(x => h.appendChild(el("th", null, x))); t.appendChild(h); linhasTr.forEach(r => t.appendChild(r)); return t; };
const td = (r, v, c) => { const x = el("td", c, v == null ? "—" : String(v)); r.appendChild(x); return x; };

export function montar(raiz, opts) {
  if (!document.getElementById("fi-css")) { const s = el("style"); s.id = "fi-css"; s.textContent = CSS; document.head.appendChild(s); }
  const api = opts.api, adm = !!opts.admin;
  const S = { base: opts.dados.base || baseVazia(), lanc: opts.dados.lanc || [], est: opts.dados.est || [], view: "inicio", f: { bu: "", obra: "", forn: "", disc: "" }, busca: "",
    form: null, estForm: null, fornSel: null, msg: opts.msg || "", erro: !!opts.erroMsg, ocupado: false, baseAba: "obras", fpag: "atrasadas", marcados: {} };
  if (!S.base.email) S.base.email = copia(EMAIL_PADRAO); if (!S.base.vinculos) S.base.vinculos = {};
  const O = id => S.base.obras.find(o => o.id === id) || {}, F = id => S.base.fornecedores.find(f => f.id === id) || {};
  const link = id => location.href.split("#")[0] + "#estouro=" + id;
  const abrir = u => (opts.abrirEmail ? opts.abrirEmail(u) : (window.location.href = u));
  const wrap = el("div", "fi"); raiz.textContent = ""; raiz.appendChild(wrap);
  const ok = m => { S.msg = m; S.erro = false; }, falha = m => { S.msg = m; S.erro = true; };
  const ir = v => { S.view = v; S.form = null; S.msg = ""; tudo(); window.scrollTo({ top: 0 }); };
  if (!opts.semEsc) document.addEventListener("keydown", ev => { if (ev.key === "Escape" && S.form && raiz.isConnected && !S.form.obraId) { S.form = null; tudo(); } });
  const discDe = x => x.disciplina || (S.base.contratos.find(c => c.id === x.contratoId) || {}).disciplina || "";
  // ── filtros globais ──
  const passa = x => { const f = S.f, o = O(x.obraId);
    return (!f.bu || (x.bu || o.bu) === f.bu) && (!f.obra || x.obraId === f.obra) && (!f.forn || x.fornecedorId === f.forn) && (!f.disc || norm(discDe(x)) === f.disc); };
  const obrasF = () => S.base.obras.filter(o => (!S.f.bu || o.bu === S.f.bu) && (!S.f.obra || o.id === S.f.obra));
  const vinculosDe = fid => { const sig = opts.siglas || {}, out = [];
    Object.keys(S.base.vinculos || {}).forEach(k => { if (S.base.vinculos[k] !== fid) return; const [sg, disc] = k.split("|"); const ob = achaObra(S.base, null, sig[sg] || sg);
      out.push({ sigla: sg, obra: sig[sg] || sg, obraId: ob ? ob.id : null, disciplina: disc }); }); return out; };

  const sincroniza = () => { const n = api.atual && api.atual(); if (n && n.base) S.base = copia(n.base); };
  async function gravaBase(txt) { S.ocupado = true; tudo(); const r = await api.salvarBase(copia(S.base)); sincroniza(); S.ocupado = false; r.ok ? ok(txt || r.msg) : falha(r.msg); tudo(); return r; }
  async function gravaLanc(l) { const r = await api.salvarLanc(copia(l)); const i = S.lanc.findIndex(x => x.id === l.id); if (i >= 0) S.lanc[i] = l; else S.lanc.push(l); return r; }
  async function gravaEst(e) { const r = await api.salvarEst(copia(e)); const i = S.est.findIndex(x => x.id === e.id); if (i >= 0) S.est[i] = e; else S.est.push(e); return r; }
  const L = () => linhas(S.base, S.lanc);
  const atr = () => { const ids = {}; atrasadas(S.base, S.lanc).forEach(x => { ids[x.id] = 1; }); return L().filter(x => ids[x.id]); };
  const diasAtraso = d => Math.round((Date.parse(hojeISO()) - Date.parse(d)) / 864e5);

  // ── barra superior: navegação enxuta + filtros em linha ──
  function topo() {
    const c = el("div", "fi-top"), l = el("div", "fi-l fi-sb"), nav = el("div", "fi-seg");
    const nAt = atr().filter(passa).length, nEst = S.est.filter(e => e.status === "pendente" && passa(e)).length;
    [["inicio", "Visão geral", ""], ["atr", "Pagamentos", nAt], ["lanc", "Lançamentos", ""], ["forn", "Fornecedores", ""], ["est", "Estouros", nEst]]
      .concat(adm ? [["base", "Base", ""], ["email", "E-mail", ""]] : [])
      .forEach(x => { const b = bt(x[1], () => ir(x[0]), S.view === x[0] || (x[0] === "atr" && S.view === "pag") ? "on" : ""); if (x[2]) b.appendChild(el("span", "fi-badge", String(x[2]))); nav.appendChild(b); });
    const nb = bt("+ Novo lançamento", novo, "fi-cta"); l.append(nav, nb); c.appendChild(l);
    const fl = el("div", "fi-fil");
    const setF = (k, v) => { S.f[k] = v; if (k === "bu") S.f.obra = ""; S.marcados = {}; tudo(); };
    const seg = el("div", "fi-seg mini"); [["", "Todas"]].concat(BUS.map(b => [b.bu, b.nome])).forEach(o => seg.appendChild(bt(o[1], () => setF("bu", o[0]), S.f.bu === o[0] ? "on" : ""))); fl.appendChild(seg);
    const so = sel([["", "Todas as obras"]].concat(S.base.obras.filter(o => !S.f.bu || o.bu === S.f.bu).sort((a, b) => a.nome.localeCompare(b.nome)).map(o => [o.id, o.nome])), S.f.obra, v => setF("obra", v));
    const sf = sel([["", "Todos os fornecedores"]].concat(S.base.fornecedores.slice().sort((a, b) => a.nome.localeCompare(b.nome)).map(f => [f.id, f.nome])), S.f.forn, v => setF("forn", v));
    const ds = {}; S.base.contratos.concat(S.base.historico).forEach(x => { if (x.disciplina) ds[norm(x.disciplina)] = ds[norm(x.disciplina)] || x.disciplina; });
    const sd = sel([["", "Todas as disciplinas"]].concat(Object.keys(ds).sort().map(k => [k, ds[k]])), S.f.disc, v => setF("disc", v));
    [so, sf, sd].forEach((s, i) => { s.className = "fi-pillsel" + (s.value ? " on" : ""); fl.appendChild(s); });
    if (S.f.bu || S.f.obra || S.f.forn || S.f.disc) fl.appendChild(bt("Limpar", () => { S.f = { bu: "", obra: "", forn: "", disc: "" }; tudo(); }, "fi-link"));
    c.appendChild(fl); return c;
  }

  // ── gráficos (SVG, sem bibliotecas) ──
  const SER = ["var(--fs1)", "var(--fs2)", "var(--fs3)", "var(--fs4)"];
  const curto = v => { const a = Math.abs(v); return (v < 0 ? "-" : "") + (a >= 1e6 ? (a / 1e6).toFixed(a >= 1e7 ? 0 : 1).replace(".", ",") + " mi" : a >= 1e3 ? Math.round(a / 1e3) + " mil" : String(Math.round(a))); };
  const svgEl = (t, at) => { const e = document.createElementNS("http://www.w3.org/2000/svg", t); Object.keys(at || {}).forEach(k => e.setAttribute(k, at[k])); return e; };
  let dica = null;
  const mostraDica = (ev, html) => { if (!dica) { dica = el("div", "fi-dica"); document.body.appendChild(dica); } dica.innerHTML = html; dica.style.display = "block";
    const w = dica.offsetWidth; dica.style.left = Math.min(window.innerWidth - w - 8, ev.clientX + 14) + "px"; dica.style.top = (ev.clientY + 14) + "px"; };
  const escondeDica = () => { if (dica) dica.style.display = "none"; };
  const esc = s => String(s == null ? "" : s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  // anel de consumo do budget
  function anel(p, rot, cor) {
    const s = svgEl("svg", { viewBox: "0 0 120 120", class: "fi-anel" }), r = 50, C = 2 * Math.PI * r, v = Math.max(0, Math.min(1, p || 0));
    s.appendChild(svgEl("circle", { cx: 60, cy: 60, r, fill: "none", stroke: "var(--fi-trilho)", "stroke-width": 12 }));
    s.appendChild(svgEl("circle", { cx: 60, cy: 60, r, fill: "none", stroke: cor, "stroke-width": 12, "stroke-linecap": "round", "stroke-dasharray": (C * v) + " " + C, transform: "rotate(-90 60 60)" }));
    const t1 = svgEl("text", { x: 60, y: 60, "text-anchor": "middle", class: "fi-anel-v" }); t1.textContent = p == null ? "—" : Math.round(p * 100) + "%"; s.appendChild(t1);
    const t2 = svgEl("text", { x: 60, y: 78, "text-anchor": "middle", class: "fi-anel-r" }); t2.textContent = rot; s.appendChild(t2); return s;
  }
  // colunas mensais empilhadas: pago × previsto × atrasado
  function colunasMes(ls, meses) {
    const h = hojeISO(), ini = somaMes(h.slice(0, 7) + "-01", -(meses - 6)), chaves = Array.from({ length: meses }, (_, i) => somaMes(ini, i).slice(0, 7));
    const atIds = {}; atr().forEach(x => { atIds[x.id] = 1; });
    const G = chaves.map(k => ({ k, pago: 0, prev: 0, atr: 0 }));
    ls.forEach(x => { if (!x.data) return; const g = G.find(y => y.k === x.data.slice(0, 7)); if (!g) return; if (x.status === "PAGO") g.pago += x.valor; else if (atIds[x.id]) g.atr += x.valor; else g.prev += x.valor; });
    const W = 720, H = 220, P = { l: 52, r: 8, t: 12, b: 26 }, max = Math.max(1, ...G.map(g => g.pago + g.prev + g.atr)), bw = (W - P.l - P.r) / G.length;
    const s = svgEl("svg", { viewBox: "0 0 " + W + " " + H, class: "fi-graf", preserveAspectRatio: "none" });
    [0, .5, 1].forEach(f => { const y = P.t + (H - P.t - P.b) * (1 - f); s.appendChild(svgEl("line", { x1: P.l, x2: W - P.r, y1: y, y2: y, class: "fi-grade-l" }));
      const t = svgEl("text", { x: P.l - 6, y: y + 4, "text-anchor": "end", class: "fi-eixo" }); t.textContent = curto(max * f); s.appendChild(t); });
    const esc2 = v => (H - P.t - P.b) * v / max, MES = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"];
    G.forEach((g, i) => { const x = P.l + i * bw + bw * .18, w = bw * .64; let y = H - P.b;
      [["pago", SER[0]], ["atr", "var(--fi-crit)"], ["prev", SER[1]]].forEach(([k, cor]) => { const hh = esc2(g[k]); if (hh < .5) return; y -= hh;
        s.appendChild(svgEl("rect", { x, y: y + 1, width: w, height: Math.max(0, hh - 2), rx: 3, fill: cor })); });
      const t = svgEl("text", { x: x + w / 2, y: H - 8, "text-anchor": "middle", class: "fi-eixo" + (g.k === h.slice(0, 7) ? " hoje" : "") }); t.textContent = MES[+g.k.slice(5) - 1] + (g.k.slice(5) === "01" ? "/" + g.k.slice(2, 4) : ""); s.appendChild(t);
      const alvo = svgEl("rect", { x: P.l + i * bw, y: P.t, width: bw, height: H - P.t - P.b, fill: "transparent" });
      alvo.addEventListener("mousemove", ev => mostraDica(ev, "<b>" + MES[+g.k.slice(5) - 1] + "/" + g.k.slice(0, 4) + "</b><br>Pago " + moeda(g.pago) + "<br>Previsto " + moeda(g.prev) + (g.atr ? "<br><span style='color:#EF4444'>Atrasado " + moeda(g.atr) + "</span>" : "")));
      alvo.addEventListener("mouseleave", escondeDica); s.appendChild(alvo); });
    const hx = P.l + (chaves.indexOf(h.slice(0, 7)) + .5) * bw; s.appendChild(svgEl("line", { x1: hx, x2: hx, y1: P.t, y2: H - P.b, class: "fi-hoje" }));
    return s;
  }
  // barras horizontais com rótulo e valor; opcional linha de referência em 100%
  function barrasH(itens, opt) {
    const c = el("div", "fi-bh"), max = opt.max || Math.max(1, ...itens.map(x => x.v));
    itens.forEach(x => { const r = el(opt.go ? "button" : "div", "fi-bh-l"); if (opt.go) { r.type = "button"; r.addEventListener("click", () => opt.go(x)); }
      r.appendChild(el("span", "fi-bh-n", x.n)); const tr = el("span", "fi-bh-t"), b = el("i"); b.style.width = Math.max(1.5, Math.min(100, x.v / max * 100)) + "%"; b.style.background = x.cor || SER[0]; tr.appendChild(b);
      if (opt.ref) { const rf = el("em"); rf.style.left = Math.min(100, opt.ref / max * 100) + "%"; tr.appendChild(rf); }
      r.appendChild(tr); r.appendChild(el("span", "fi-bh-v", x.txt != null ? x.txt : curto(x.v)));
      if (x.dica) { r.addEventListener("mousemove", ev => mostraDica(ev, x.dica)); r.addEventListener("mouseleave", escondeDica); }
      c.appendChild(r); });
    return c;
  }
  // curva acumulada do desembolso × budget
  function curva(ls, budget) {
    const pts = ls.filter(x => x.data).sort((a, b) => a.data.localeCompare(b.data)); if (!pts.length) return el("p", "fi-aviso", "Sem pagamentos com data.");
    let ac = 0; const serie = pts.map(x => { ac += x.valor; return { d: x.data, v: ac, x }; });
    const W = 720, H = 220, P = { l: 52, r: 10, t: 12, b: 26 }, d0 = Date.parse(serie[0].d), d1 = Math.max(Date.parse(serie[serie.length - 1].d), Date.parse(hojeISO())) + 864e5;
    const max = Math.max(ac, budget || 0) * 1.05, X = d => P.l + (W - P.l - P.r) * (Date.parse(d) - d0) / (d1 - d0), Y = v => P.t + (H - P.t - P.b) * (1 - v / max);
    const s = svgEl("svg", { viewBox: "0 0 " + W + " " + H, class: "fi-graf", preserveAspectRatio: "none" });
    [0, .5, 1].forEach(f => { const y = Y(max * f); s.appendChild(svgEl("line", { x1: P.l, x2: W - P.r, y1: y, y2: y, class: "fi-grade-l" })); const t = svgEl("text", { x: P.l - 6, y: y + 4, "text-anchor": "end", class: "fi-eixo" }); t.textContent = curto(max * f); s.appendChild(t); });
    if (budget) { s.appendChild(svgEl("line", { x1: P.l, x2: W - P.r, y1: Y(budget), y2: Y(budget), class: "fi-ref" })); const t = svgEl("text", { x: W - P.r - 4, y: Y(budget) - 6, "text-anchor": "end", class: "fi-eixo forte" }); t.textContent = "Budget " + curto(budget); s.appendChild(t); }
    let dp = "M" + X(serie[0].d) + " " + Y(0); serie.forEach(p => { dp += " H" + X(p.d) + " V" + Y(p.v); });
    s.appendChild(svgEl("path", { d: dp + " V" + Y(0) + " Z", fill: "var(--fs1)", opacity: .12 }));
    s.appendChild(svgEl("path", { d: dp, fill: "none", stroke: "var(--fs1)", "stroke-width": 2 }));
    const hx = X(hojeISO()); s.appendChild(svgEl("line", { x1: hx, x2: hx, y1: P.t, y2: H - P.b, class: "fi-hoje" }));
    [serie[0].d, hojeISO()].forEach((d, i) => { const t = svgEl("text", { x: i ? hx : P.l, y: H - 8, "text-anchor": i ? "middle" : "start", class: "fi-eixo" + (i ? " hoje" : "") }); t.textContent = i ? "hoje" : br(d); s.appendChild(t); });
    const alvo = svgEl("rect", { x: P.l, y: P.t, width: W - P.l - P.r, height: H - P.t - P.b, fill: "transparent" });
    alvo.addEventListener("mousemove", ev => { const r = s.getBoundingClientRect(), dx = d0 + (ev.clientX - r.left) / r.width * W; const tx = (dx - P.l) / (W - P.l - P.r) * (d1 - d0) + d0;
      let p = serie[0]; serie.forEach(q => { if (Date.parse(q.d) <= tx) p = q; }); mostraDica(ev, "<b>" + br(p.d) + "</b><br>Acumulado " + moeda(p.v) + (budget ? "<br>" + Math.round(p.v / budget * 100) + "% do budget" : "")); });
    alvo.addEventListener("mouseleave", escondeDica); s.appendChild(alvo); return s;
  }
  const corConsumo = p => p > 1 ? "var(--fi-crit)" : p > .85 ? "var(--fi-aten)" : "var(--fi-bom)";
  const cartao = (tit, sub, corpo, extra) => { const c = el("div", "fi-card"), t = el("div", "fi-l fi-sb"), tt = el("div"); tt.appendChild(el("h4", null, tit)); if (sub) tt.appendChild(el("p", "fi-aviso", sub));
    t.appendChild(tt); if (extra) t.appendChild(extra); c.appendChild(t); c.appendChild(corpo); return c; };
  const legenda = itens => { const l = el("div", "fi-leg"); itens.forEach(x => { const s = el("span"); const b = el("b"); b.style.background = x[1]; s.append(b, x[0]); l.appendChild(s); }); return l; };

  // ── visão geral (sem obra: carteira; com obra: a obra) ──
  function inicio() { return S.f.obra ? visaoObra() : visaoGeral(); }
  function heroi(budget, usado, extra) {
    const c = el("div", "fi-hero"), p = budget ? usado / budget : null;
    c.appendChild(anel(p, "do budget", p == null ? "var(--fs1)" : corConsumo(p)));
    const t = el("div", "fi-hero-t"); t.appendChild(el("span", "fi-hero-r", "Saldo de budget"));
    const v = el("b", "fi-hero-v", budget ? moeda(budget - usado) : "—"); if (budget && budget - usado < 0) v.style.color = "var(--fi-crit)"; t.appendChild(v);
    t.appendChild(el("span", "fi-aviso", "Budget " + moeda(budget) + " · utilizado " + moeda(usado))); c.appendChild(t);
    const st = el("div", "fi-hero-s"); extra.forEach(x => st.appendChild(kpi(x[0], x[1], x[2], x[3], x[4]))); c.appendChild(st); return c;
  }
  function visaoGeral() {
    const box = el("div", "fi"), h = hojeISO(), ls = L().filter(passa), obras = obrasF(), filtroLinha = !!(S.f.forn || S.f.disc);
    const budget = filtroLinha ? 0 : obras.reduce((a, o) => a + (o.budget || 0), 0), usado = ls.reduce((a, x) => a + x.valor, 0);
    const a30 = ls.filter(x => x.status !== "PAGO" && x.data && x.data >= h && x.data <= somaMes(h, 1)), at = atr().filter(passa), estP = S.est.filter(e => e.status === "pendente" && passa(e));
    box.appendChild(heroi(budget, usado, [
      ["A pagar em 30 dias", moeda(a30.reduce((a, x) => a + x.valor, 0)), "var(--fi-aten)", a30.length + " parcela(s)", () => { S.fpag = "30"; ir("pag"); }],
      ["Atrasados", moeda(at.reduce((a, x) => a + x.valor, 0)), at.length ? "var(--fi-crit)" : "var(--fi-bom)", at.length ? at.length + " parcela(s) · resolver" : "em dia", () => ir("atr")],
      ["Estouros pendentes", String(estP.length), estP.length ? "var(--fi-crit)" : "var(--fi-bom)", estP.length ? "aguardando analista" : "nenhum", () => ir("est")]]));
    const g1 = cartao("Fluxo de pagamentos", "Últimos 12 meses e próximos 6 · passe o mouse nas colunas", colunasMes(ls, 18), legenda([["Pago", SER[0]], ["Previsto", SER[1]], ["Atrasado", "var(--fi-crit)"]]));
    box.appendChild(g1);
    const dupla = el("div", "fi-dupla");
    if (!filtroLinha) {
      const rk = obras.filter(o => o.budget).map(o => { const u = L().filter(x => x.obraId === o.id).reduce((a, x) => a + x.valor, 0); return { o, u, p: u / o.budget }; }).sort((a, b) => b.p - a.p);
      dupla.appendChild(cartao("Consumo do budget por obra", rk.length + " obras com budget · linha = 100% · clique para abrir a obra",
        barrasH(rk.slice(0, 10).map(x => ({ n: x.o.nome, v: x.p * 100, txt: Math.round(x.p * 100) + "%", cor: corConsumo(x.p), id: x.o.id, dica: "<b>" + esc(x.o.nome) + "</b><br>Budget " + moeda(x.o.budget) + "<br>Utilizado " + moeda(x.u) })),
          { max: Math.max(120, ...rk.slice(0, 10).map(x => x.p * 100)), ref: 100, go: x => { S.f.obra = x.id; tudo(); } })));
    }
    const porD = {}; S.base.contratos.filter(passa).forEach(c => { const k = c.disciplina || "Sem disciplina"; porD[k] = (porD[k] || 0) + (c.valor || 0); });
    const dd = Object.keys(porD).map(k => ({ n: k, v: porD[k] })).sort((a, b) => b.v - a.v), topD = dd.slice(0, 8), outD = dd.slice(8).reduce((a, x) => a + x.v, 0); if (outD) topD.push({ n: "Outras", v: outD, cor: "var(--fi-neutro)" });
    dupla.appendChild(cartao("Contratado por disciplina", moeda(dd.reduce((a, x) => a + x.v, 0)) + " em contratos", barrasH(topD.map(x => Object.assign({ cor: x.cor || SER[0], txt: curto(x.v), dica: "<b>" + esc(x.n) + "</b><br>" + moeda(x.v) }, x)), {
      go: x => { if (x.n !== "Outras") { S.f.disc = norm(x.n); tudo(); } } })));
    const porF = {}; S.base.contratos.filter(passa).forEach(c => { if (!c.fornecedorId) return; porF[c.fornecedorId] = (porF[c.fornecedorId] || 0) + (c.valor || 0); });
    const tf = Object.keys(porF).map(k => ({ id: k, n: F(k).nome || "—", v: porF[k] })).sort((a, b) => b.v - a.v).slice(0, 8);
    dupla.appendChild(cartao("Maiores fornecedores", "Por valor contratado · clique para filtrar", barrasH(tf.map(x => Object.assign({ cor: SER[2], dica: "<b>" + esc(x.n) + "</b><br>" + moeda(x.v) }, x)), { go: x => { S.f.forn = x.id; tudo(); } })));
    if (!S.f.bu && !filtroLinha) { const c = el("div", "fi-bus");
      BUS.forEach((b, i) => { const os = obras.filter(o => o.bu === b.bu), bg = os.reduce((a, o) => a + (o.budget || 0), 0), us = ls.filter(x => x.bu === b.bu).reduce((a, x) => a + x.valor, 0);
        const it = el("button", "fi-bu"); it.type = "button"; it.addEventListener("click", () => { S.f.bu = b.bu; tudo(); }); it.appendChild(anel(bg ? us / bg : null, b.nome, bg ? corConsumo(us / bg) : "var(--fs1)"));
        const t = el("div"); t.append(el("b", null, b.nome), el("span", "fi-aviso", os.length + " obras · " + at.filter(x => x.bu === b.bu).length + " atrasados"), el("span", "fi-aviso", "Saldo " + moeda(bg - us))); it.appendChild(t); c.appendChild(it); });
      dupla.appendChild(cartao("Unidades de negócio", "Consumo do budget · clique para filtrar", c)); }
    box.appendChild(dupla);
    const pend = [];
    const semB = obras.filter(o => o.budget == null).length; if (semB && !filtroLinha) pend.push(semB + " obra(s) sem budget");
    const semC = S.base.fornecedores.filter(f => !f.cnpj && S.base.contratos.some(c => c.fornecedorId === f.id && passa(c))).length; if (semC) pend.push(semC + " fornecedor(es) com contrato e sem CNPJ");
    if (pend.length) box.appendChild(el("p", "fi-alerta am", "Pendências de cadastro: " + pend.join(" · ") + "."));
    if (!S.base.obras.length) box.appendChild(el("p", "fi-alerta am", adm ? "Base ainda não importada. Vá em Base → Importar planilha SAÍDAS." : "Base ainda não cadastrada pelo administrador."));
    return box;
  }
  function visaoObra() {
    const o = O(S.f.obra), box = el("div", "fi"), ls = L().filter(passa), h = hojeISO(), usado = ls.reduce((a, x) => a + x.valor, 0), filtroLinha = !!(S.f.forn || S.f.disc);
    const cts = S.base.contratos.filter(passa), at = atr().filter(passa), a30 = ls.filter(x => x.status !== "PAGO" && x.data && x.data >= h && x.data <= somaMes(h, 1));
    const cab = el("div", "fi-l fi-sb"); const ct = el("div"); ct.append(el("h3", null, o.nome), el("p", "fi-aviso", [(BUS.find(b => b.bu === o.bu) || {}).nome, o.gds, o.spe, o.uh ? o.uh + " UH" : "", o.m2 ? Math.round(o.m2).toLocaleString("pt-BR") + " m²" : ""].filter(Boolean).join(" · ")));
    cab.append(ct, bt("← Visão geral", () => { S.f.obra = ""; tudo(); }, "fi-link")); box.appendChild(cab);
    box.appendChild(heroi(filtroLinha ? 0 : o.budget || 0, usado, [
      ["Contratado", moeda(cts.reduce((a, c) => a + (c.valor || 0), 0)), "var(--fs1)", cts.length + " contrato(s)"],
      ["A pagar em 30 dias", moeda(a30.reduce((a, x) => a + x.valor, 0)), "var(--fi-aten)", a30.length + " parcela(s)", () => { S.fpag = "30"; ir("pag"); }],
      ["Atrasados", moeda(at.reduce((a, x) => a + x.valor, 0)), at.length ? "var(--fi-crit)" : "var(--fi-bom)", at.length ? at.length + " · resolver" : "em dia", () => ir("atr")]]));
    box.appendChild(cartao("Desembolso acumulado", "Soma das parcelas no tempo · linha tracejada = budget", curva(ls, filtroLinha ? null : o.budget)));
    const dupla = el("div", "fi-dupla");
    const cl = cts.map(c => { const s2 = situacaoContrato(S.base, S.lanc, c.id); return { c, s2, p: c.valor ? s2.usado / c.valor : 0 }; }).sort((a, b) => b.c.valor - a.c.valor);
    dupla.appendChild(cartao("Contratos da obra", "Lançado sobre o contratado · linha = 100%", barrasH(cl.map(x => ({ n: x.c.disciplina + " · " + (F(x.c.fornecedorId).nome || "—"), v: x.p * 100, txt: (x.c.valor > 0 ? "saldo " + curto(x.s2.saldo) : "sem valor"),
      cor: x.c.valor > 0 ? corConsumo(x.p) : "var(--fi-neutro)", dica: "<b>" + esc(x.c.disciplina) + "</b> · " + esc(F(x.c.fornecedorId).nome) + "<br>Contratado " + moeda(x.c.valor) + "<br>Lançado " + moeda(x.s2.usado) + "<br>Saldo " + moeda(x.s2.saldo) + (x.c.pedido ? "<br>Pedido ERP " + esc(x.c.pedido) : "") })),
      { max: Math.max(110, ...cl.map(x => x.p * 100)), ref: 100 })));
    const prox = ls.filter(x => x.status !== "PAGO").sort((a, b) => String(a.data || "9").localeCompare(String(b.data || "9"))).slice(0, 8);
    const lp = el("div", "fi-lista"); prox.forEach(x => { const r = el("div", "fi-li"), d = el("div", "fi-li-d"); d.append(el("b", null, x.data ? x.data.slice(8, 10) : "—"), el("span", null, x.data ? ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"][+x.data.slice(5, 7) - 1] + "/" + x.data.slice(2, 4) : ""));
      const t = el("div"); t.append(el("b", null, F(x.fornecedorId).nome || "—"), el("span", "fi-aviso", discDe(x)));
      const v = el("div", "fi-li-v"); v.append(el("b", null, moeda(x.valor)), el("span", "fi-pill " + (at.some(y => y.id === x.id) ? "fi-no" : "fi-am"), at.some(y => y.id === x.id) ? "atrasado" : "previsto"));
      r.append(d, t, v); lp.appendChild(r); });
    dupla.appendChild(cartao("Próximas parcelas", prox.length ? "" : "Nada em aberto", lp, bt("Ver todas", () => { S.fpag = "abertas"; ir("pag"); }, "fi-link")));
    box.appendChild(dupla);
    return box;
  }

  // ── pagamentos: atrasados e todos ──
  async function marcarPago(xs) {
    const h = hojeISO(); let base = false; const lancs = {};
    xs.forEach(x => { if (x.origem === "painel") { const lc = S.lanc.find(y => y.id === x.lancId); if (lc) { lc.parcelas[x.parcela].status = "PAGO"; lc.parcelas[x.parcela].pagoEm = h; lancs[lc.id] = lc; } }
      else { const hh = S.base.historico.find(y => y.id === x.id); if (hh) { hh.status = "PAGO"; base = true; } } });
    S.ocupado = true; tudo(); let falhou = "";
    for (const id of Object.keys(lancs)) { const r = await gravaLanc(lancs[id]); if (!r.ok) falhou = r.msg; }
    if (base) { const r = await api.salvarBase(copia(S.base)); sincroniza(); if (!r.ok) falhou = r.msg; }
    S.ocupado = false; S.marcados = {}; falhou ? falha(falhou) : ok(xs.length + " parcela(s) marcada(s) como paga(s)."); tudo();
  }
  async function novaData(x, d) {
    if (!d) return;
    if (x.origem === "painel") { const lc = S.lanc.find(y => y.id === x.lancId); lc.parcelas[x.parcela].data = d; const r = await gravaLanc(lc); r.ok ? ok("Data reprogramada para " + br(d) + ".") : falha(r.msg); }
    else { const hh = S.base.historico.find(y => y.id === x.id); hh.data = d; if (hh.status !== "PAGO") hh.status = "PREVISTO"; const r = await api.salvarBase(copia(S.base)); r.ok ? ok("Data reprogramada para " + br(d) + ".") : falha(r.msg); }
    tudo();
  }
  function pagamentos(soAtr) {
    const c = el("div", "fi-card"), h = hojeISO();
    const atIds = {}; atr().forEach(x => { atIds[x.id] = 1; });
    let ls = L().filter(passa);
    if (soAtr) ls = ls.filter(x => atIds[x.id]);
    else { const sl = sel([["abertas", "Em aberto"], ["30", "A pagar em 30 dias"], ["atrasadas", "Atrasadas"], ["pagas", "Pagas"], ["todas", "Todas"]], S.fpag, v => { S.fpag = v; tudo(); });
      const lt = el("div", "fi-l"); lt.append(el("h3", null, "Pagamentos"), sl); c.appendChild(lt);
      const f = S.fpag; ls = ls.filter(x => f === "todas" ? true : f === "pagas" ? x.status === "PAGO" : f === "atrasadas" ? atIds[x.id] : f === "30" ? x.status !== "PAGO" && x.data && x.data >= h && x.data <= somaMes(h, 1) : x.status !== "PAGO"); }
    ls.sort((a, b) => String(a.data || "9").localeCompare(String(b.data || "9")));
    if (soAtr) { const lt = el("div", "fi-l fi-sb"); lt.appendChild(el("h3", null, ls.length ? ls.length + " pagamento(s) atrasado(s) · " + moeda(ls.reduce((a, x) => a + x.valor, 0)) : "Nenhum pagamento atrasado"));
      c.appendChild(lt); c.appendChild(el("p", "fi-aviso", "Data de pagamento já passou e o status não é PAGO. Confirme com o financeiro: marque como pago ou escolha a nova data no calendário da linha (grava na hora).")); }
    const podeEd = x => x.origem === "painel" || adm;
    const marc = ls.filter(x => S.marcados[x.id] && podeEd(x));
    const ab = el("div", "fi-l"); ab.style.margin = "10px 0";
    const tudoSel = el("input"); tudoSel.type = "checkbox"; tudoSel.checked = ls.length > 0 && ls.filter(podeEd).every(x => S.marcados[x.id]);
    tudoSel.addEventListener("change", () => { ls.filter(podeEd).forEach(x => { if (tudoSel.checked) S.marcados[x.id] = 1; else delete S.marcados[x.id]; }); tudo(); });
    const lb = el("label", "fi-l"); lb.append(tudoSel, el("span", "fi-aviso", "Selecionar todos"));
    const mp = bt("Marcar " + marc.length + " como pago", () => { if (window.confirm("Marcar " + marc.length + " parcela(s) como PAGO? Confirme antes com o financeiro.")) marcarPago(marc); }, "pri"); mp.disabled = !marc.length || S.ocupado;
    ab.append(lb, mp); if (ls.some(x => x.status !== "PAGO")) c.appendChild(ab);
    const box = el("div", "fi-tab");
    box.appendChild(tabela(["", "Data", "Atraso", "Obra", "Disciplina", "Fornecedor", "Valor", "Status", "Origem", "Reprogramar", ""], ls.slice(0, 500).map(x => {
      const r = el("tr"), cb = el("td"); if (x.status !== "PAGO" && podeEd(x)) { const k = el("input"); k.type = "checkbox"; k.checked = !!S.marcados[x.id]; k.addEventListener("change", () => { if (k.checked) S.marcados[x.id] = 1; else delete S.marcados[x.id]; tudo(); }); cb.appendChild(k); } r.appendChild(cb);
      td(r, br(x.data)); const ta = td(r, atIds[x.id] ? diasAtraso(x.data) + " d" : ""); if (atIds[x.id]) ta.className = "fi-dias";
      td(r, O(x.obraId).nome); td(r, discDe(x) || "—"); td(r, F(x.fornecedorId).nome || "—"); td(r, moeda(x.valor), "n");
      const st = el("td"); st.appendChild(el("span", "fi-pill " + (x.status === "PAGO" ? "fi-ok" : atIds[x.id] ? "fi-no" : x.status ? "fi-am" : "fi-ci"), atIds[x.id] ? "Atrasado" : x.status || "sem status")); r.appendChild(st);
      td(r, x.origem === "painel" ? "Nº " + x.numero : "Planilha");
      const rp = el("td"); if (x.status !== "PAGO" && podeEd(x)) { const d = el("input"); d.type = "date"; d.value = x.data || ""; d.addEventListener("change", () => novaData(x, d.value)); rp.appendChild(d); } r.appendChild(rp);
      const ac = el("td"); if (x.status !== "PAGO" && podeEd(x)) ac.appendChild(bt("Pago ✓", () => marcarPago([x]), "mini")); r.appendChild(ac);
      return r; })));
    if (!ls.length) box.appendChild(el("p", "fi-aviso", "Nada neste filtro.")); if (ls.length > 500) box.appendChild(el("p", "fi-aviso", "Mostrando 500 de " + ls.length + ". Use os filtros."));
    if (!adm && ls.some(x => x.origem === "planilha" && x.status !== "PAGO")) box.appendChild(el("p", "fi-aviso", "Linhas vindas da planilha só o administrador atualiza."));
    c.appendChild(box); return c;
  }

  // ── fornecedores: cadastro único (Lançamento Financeiro + Gestão de Fornecedores) ──
  function fornecedores() {
    const c = el("div", "fi-card"), lt = el("div", "fi-l fi-sb");
    lt.append(el("h3", null, "Cadastro único de fornecedores"), bt("+ Fornecedor", () => { const f = { id: novoId("f"), nome: "", razao: "", fantasia: "", cnpj: "", mega: "", contato: "", telefone: "", email: "" }; S.fornSel = f; S.fornNovo = true; tudo(); }, "pri"));
    c.appendChild(lt);
    c.appendChild(el("p", "fi-aviso", "Mesma base da aba Gestão de Fornecedores: os vínculos obra × disciplina de lá aparecem aqui, junto com contratos e pagamentos."));
    const q = inp(S.busca, v => { S.busca = v; desenha(); }, "Buscar nome, CNPJ ou contato"); q.style.margin = "10px 0"; q.style.width = "100%"; c.appendChild(q);
    const box = el("div", "fi-tab"); c.appendChild(box);
    function desenha() {
      const qq = norm(S.busca), ls = L();
      const fs = S.base.fornecedores.filter(f => (!S.f.forn || f.id === S.f.forn) && (!qq || norm([f.nome, f.razao, f.fantasia, f.cnpj, f.contato, f.email].join(" ")).indexOf(qq) >= 0))
        .map(f => { const cts = S.base.contratos.filter(x => x.fornecedorId === f.id && passa(x)), vs = vinculosDe(f.id).filter(v => (!S.f.obra || v.obraId === S.f.obra) && (!S.f.disc || norm(v.disciplina) === S.f.disc));
          const pg = ls.filter(x => x.fornecedorId === f.id && passa(x)); return { f, cts, vs, pg, tot: cts.reduce((a, x) => a + (x.valor || 0), 0) }; })
        .filter(x => !(S.f.obra || S.f.disc || S.f.bu) || x.cts.length || x.vs.length || x.pg.length).sort((a, b) => b.tot - a.tot || a.f.nome.localeCompare(b.f.nome));
      box.textContent = "";
      box.appendChild(tabela(["Fornecedor", "CNPJ", "Cadastro Mega", "Contato", "E-mail", "Obras", "Contratado", "Vínculos (Gestão)", ""], fs.map(x => { const r = el("tr", "clic");
        r.addEventListener("click", () => { S.fornSel = x.f; S.fornNovo = false; tudo(); });
        td(r, x.f.nome); const tc = td(r, x.f.cnpj || "sem CNPJ"); if (!x.f.cnpj) tc.style.color = "#B45309"; td(r, x.f.mega || "—"); td(r, [x.f.contato, x.f.telefone].filter(Boolean).join(" · ") || "—"); td(r, x.f.email || "—");
        td(r, new Set(x.cts.map(c2 => c2.obraId)).size, "n"); td(r, moeda(x.tot), "n"); td(r, x.vs.length, "n"); td(r, "Abrir →"); return r; })));
      if (!fs.length) box.appendChild(el("p", "fi-aviso", "Nenhum fornecedor neste filtro."));
    }
    desenha(); return c;
  }
  function fichaFornecedor() {
    const f = S.fornSel, md = el("div", "fi-modal"), c = el("div", "fi-card fi");
    const t = el("div", "fi-l fi-sb"); t.append(el("h3", null, S.fornNovo ? "Novo fornecedor" : f.nome), bt("Fechar", () => { S.fornSel = null; tudo(); })); c.appendChild(t);
    const g = el("div", "fi-grade"), d = copia(f);
    [["nome", "Nome (como aparece nos lançamentos)"], ["razao", "Razão social"], ["fantasia", "Nome fantasia"], ["cnpj", "CNPJ"], ["mega", "Nº cadastro Mega"], ["contato", "Contato"], ["telefone", "Telefone"], ["email", "E-mail"]]
      .forEach(k => g.appendChild(campo(k[1], inp(d[k[0]], v => { d[k[0]] = v; })))); c.appendChild(g);
    const ac = el("div", "fi-l"); ac.style.justifyContent = "flex-end";
    ac.appendChild(bt("Salvar cadastro", async () => {
      if (!String(d.nome || "").trim()) { alert("Informe o nome."); return; }
      const dup = S.base.fornecedores.find(x => x.id !== f.id && (norm(x.nome) === norm(d.nome) || (d.cnpj && String(x.cnpj).replace(/\D/g, "") === String(d.cnpj).replace(/\D/g, ""))));
      if (dup && !window.confirm("Já existe " + dup.nome + (dup.cnpj ? " (" + dup.cnpj + ")" : "") + " com esse nome ou CNPJ. Salvar mesmo assim?")) return;
      Object.keys(d).forEach(k => { f[k] = typeof d[k] === "string" ? d[k].trim() : d[k]; }); if (S.fornNovo) S.base.fornecedores.push(f);
      S.fornNovo = false; S.fornSel = null; await gravaBase("Cadastro de " + f.nome + " salvo. Vale para o Lançamento Financeiro e a Gestão de Fornecedores."); }, "pri"));
    c.appendChild(ac);
    if (!S.fornNovo) {
      const cts = S.base.contratos.filter(x => x.fornecedorId === f.id);
      c.appendChild(el("h4", null, "Contratos · " + cts.length));
      c.appendChild(cts.length ? tabela(["Obra", "Disciplina", "Pedido ERP", "Contratado", "Lançado", "Saldo"], cts.map(x => { const s2 = situacaoContrato(S.base, S.lanc, x.id), r = el("tr");
        td(r, O(x.obraId).nome); td(r, x.disciplina); td(r, x.pedido || "—"); td(r, moeda(x.valor), "n"); td(r, moeda(s2.usado), "n"); const sd = td(r, moeda(s2.saldo), "n"); if (s2.saldo < 0) sd.style.color = "#B91C1C"; return r; })) : el("p", "fi-aviso", "Sem contratos."));
      const vs = vinculosDe(f.id); c.appendChild(el("h4", null, "Vínculos na Gestão de Fornecedores · " + vs.length));
      c.appendChild(vs.length ? tabela(["Obra", "Disciplina"], vs.map(v => { const r = el("tr"); td(r, v.obra + " (" + v.sigla + ")"); td(r, v.disciplina); return r; })) : el("p", "fi-aviso", "Sem vínculos."));
      const pg = L().filter(x => x.fornecedorId === f.id).sort((a, b) => String(b.data || "").localeCompare(String(a.data || ""))).slice(0, 12);
      c.appendChild(el("h4", null, "Últimos pagamentos"));
      c.appendChild(pg.length ? tabela(["Data", "Obra", "Disciplina", "Valor", "Status"], pg.map(x => { const r = el("tr"); td(r, br(x.data)); td(r, O(x.obraId).nome); td(r, discDe(x)); td(r, moeda(x.valor), "n"); td(r, x.status || "—"); return r; })) : el("p", "fi-aviso", "Sem pagamentos."));
      const emUso = S.base.contratos.some(x => x.fornecedorId === f.id) || S.base.historico.some(x => x.fornecedorId === f.id) || S.lanc.some(x => x.fornecedorId === f.id) || vs.length;
      if (adm) { const outros = S.base.fornecedores.filter(x => x.id !== f.id).sort((a, b) => a.nome.localeCompare(b.nome));
        const ms = sel([["", "Mesclar este cadastro em…"]].concat(outros.map(x => [x.id, x.nome + (x.cnpj ? " · " + x.cnpj : "")])), "", async v => { if (!v) return; const alvo = F(v);
          if (!window.confirm("Mesclar " + f.nome + " em " + alvo.nome + "? Contratos, pagamentos, lançamentos e vínculos passam para " + alvo.nome + " e este cadastro some. Dados vazios do destino são completados com os deste.")) { ms.value = ""; return; }
          ["razao", "fantasia", "cnpj", "mega", "contato", "telefone", "email"].forEach(k => { if (!alvo[k] && f[k]) alvo[k] = f[k]; });
          S.base.contratos.forEach(x => { if (x.fornecedorId === f.id) x.fornecedorId = v; }); S.base.historico.forEach(x => { if (x.fornecedorId === f.id) x.fornecedorId = v; });
          Object.keys(S.base.vinculos || {}).forEach(k => { if (S.base.vinculos[k] === f.id) S.base.vinculos[k] = v; });
          for (const l of S.lanc.filter(x => x.fornecedorId === f.id)) { l.fornecedorId = v; await gravaLanc(l); }
          for (const e of S.est.filter(x => x.fornecedorId === f.id)) { e.fornecedorId = v; await gravaEst(e); }
          S.base.fornecedores = S.base.fornecedores.filter(x => x !== f); S.fornSel = null; await gravaBase(f.nome + " mesclado em " + alvo.nome + "."); });
        ms.style.marginTop = "12px"; c.appendChild(ms); }
      if (adm && !emUso) c.appendChild(bt("Excluir fornecedor", async () => { if (!window.confirm("Excluir " + f.nome + "?")) return; S.base.fornecedores = S.base.fornecedores.filter(x => x !== f); S.fornSel = null; await gravaBase("Fornecedor excluído."); }, "per"));
    }
    md.appendChild(c); md.addEventListener("click", ev => { if (ev.target === md) { S.fornSel = null; tudo(); } }); return md;
  }

  // ── lançamentos ──
  function novo() {
    const bu = S.f.bu || (S.f.obra ? O(S.f.obra).bu : "Make It");
    S.form = { id: novoId("l"), bu, obraId: S.f.obra || "", tipo: "Pedido", contratoId: "", novoContrato: false, disciplina: "", fornecedorId: S.f.forn || "", valorContrato: null,
      causa: "", parcelas: [{ valor: null, data: somaMes(hojeISO(), 1), status: "PREVISTO" }], obs: "", status: "rascunho", criadoPor: opts.usuario || "", criadoEm: hojeISO() };
    S.msg = ""; tudo();
  }
  const totalForm = f => (f.parcelas || []).reduce((a, p) => a + (num(p.valor) || 0), 0);
  function estouroDe(f) {
    const tot = totalForm(f), outros = S.lanc.filter(x => x.id !== f.id), so = situacaoObra(S.base, outros, f.obraId, tot);
    const sc = f.contratoId ? situacaoContrato(S.base, outros, f.contratoId, tot) : (f.novoContrato && num(f.valorContrato) != null ? { valor: num(f.valorContrato), usado: tot, saldo: num(f.valorContrato) - tot } : null);
    const r2 = v => Math.round(v * 100) / 100;
    const scOk = sc && sc.valor > 0; // contrato com valor zerado na planilha = sem limite cadastrado
    return { so, sc: scOk ? sc : (sc ? Object.assign({}, sc, { valor: null, saldo: null }) : null), contrato: scOk && sc.saldo < -0.005 ? r2(-sc.saldo) : 0, obra: so && so.saldo != null && so.saldo < -0.005 ? r2(-so.saldo) : 0 };
  }
  async function finalizar() {
    const f = S.form, faltas = [];
    if (!f.obraId) faltas.push("obra"); if (!f.contratoId && !f.novoContrato) faltas.push("contrato");
    if (f.novoContrato && (!f.disciplina || !f.fornecedorId || num(f.valorContrato) == null)) faltas.push("dados do novo contrato");
    if (!String(f.causa || "").trim()) faltas.push("causa raiz do pedido");
    if (!f.parcelas.length || f.parcelas.some(p => num(p.valor) == null || !p.data)) faltas.push("valor e data de todas as parcelas");
    if (faltas.length) { falha("Preencha: " + faltas.join(", ") + "."); tudo(); return; }
    const e = estouroDe(f);
    if ((e.contrato || e.obra) && !window.confirm("Este lançamento excede " + [e.contrato ? "o contrato em " + moeda(e.contrato) : "", e.obra ? "o budget da obra em " + moeda(e.obra) : ""].filter(Boolean).join(" e ") +
      ".\n\nAo finalizar, abre um formulário de composição do estouro para o analista responder. Continuar?")) return;
    S.ocupado = true; tudo();
    try {
      if (f.novoContrato) { const c = { id: novoId("c"), bu: f.bu, obraId: f.obraId, disciplina: f.disciplina, fornecedorId: f.fornecedorId, valor: num(f.valorContrato), pedido: "" };
        S.base.contratos.push(c); f.contratoId = c.id; f.novoContrato = false; const rb = await api.salvarBase(copia(S.base)); if (!rb.ok) throw new Error(rb.msg); }
      const c = S.base.contratos.find(x => x.id === f.contratoId); if (c) { f.disciplina = c.disciplina; f.fornecedorId = c.fornecedorId; }
      if (api.recarregar) { const nv = await api.recarregar(); if (nv) { S.lanc = nv.lanc; S.est = nv.est; } }
      f.numero = S.lanc.reduce((m, x) => Math.max(m, x.numero || 0), 0) + 1;
      f.status = "finalizado"; f.finalizadoEm = hojeISO(); f.estouro = e.contrato || e.obra ? { contrato: e.contrato, obra: e.obra } : null;
      if (f.estouro) { const es = { id: novoId("e"), lancId: f.id, numero: f.numero, bu: f.bu, obraId: f.obraId, disciplina: f.disciplina, fornecedorId: f.fornecedorId,
          contrato: e.contrato, obra: e.obra, analista: f.criadoPor, criadoEm: hojeISO(), status: "pendente" }; f.estouroId = es.id; const r2 = await gravaEst(es); if (!r2.ok) throw new Error(r2.msg); }
      const r = await gravaLanc(f); if (!r.ok) throw new Error(r.msg);
      const m = montaEmail(S.base, f, f.estouroId ? link(f.estouroId) : "");
      S.form = null; S.ocupado = false; S.view = "lanc"; ok("Lançamento nº " + f.numero + " finalizado. Abrindo o e-mail do pedido…"); tudo();
      abrir(mailto(m));
    } catch (err) { S.ocupado = false; falha("Não foi possível finalizar: " + err.message); tudo(); }
  }
  // ── novo lançamento: painel lateral em 4 passos ──
  function telaForm() {
    const f = S.form; f.passo = f.passo || (f.obraId ? 2 : 1);
    const ov = el("div", "fi-drawer-bg"), dw = el("div", "fi-drawer fi");
    const fechar = () => { if (f.obraId && !window.confirm("Descartar este lançamento?")) return; S.form = null; tudo(); };
    ov.addEventListener("click", ev => { if (ev.target === ov) fechar(); });
    const cab = el("div", "fi-dw-cab"), l = el("div", "fi-l fi-sb"); l.append(el("h3", null, "Novo lançamento"), bt("✕", fechar, "fi-link")); cab.appendChild(l);
    const PASSOS = ["Obra", "Contrato", "Pagamento", "Revisão"], st = el("div", "fi-passos");
    PASSOS.forEach((p, i) => { const b = bt((i + 1) + " " + p, () => { if (podeIr(i + 1)) { f.passo = i + 1; S.msg = ""; tudo(); } }, i + 1 === f.passo ? "on" : i + 1 < f.passo ? "feito" : ""); st.appendChild(b); });
    cab.appendChild(st); dw.appendChild(cab);
    const corpo = el("div", "fi-dw-corpo"); dw.appendChild(corpo);
    const ob = O(f.obraId), ct = S.base.contratos.find(x => x.id === f.contratoId);
    const ok1 = () => !!f.obraId, ok2 = () => !!(f.contratoId || (f.novoContrato && String(f.disciplina || "").trim() && f.fornecedorId && num(f.valorContrato) != null)),
      ok3 = () => !!(String(f.causa || "").trim() && f.parcelas.length && f.parcelas.every(p => num(p.valor) != null && p.data));
    const podeIr = n => n <= 1 || (ok1() && (n <= 2 || (ok2() && (n <= 3 || ok3()))));
    if (f.passo === 1) {
      corpo.appendChild(el("p", "fi-pergunta", "Para qual obra é o lançamento?"));
      const seg = el("div", "fi-seg mini"); BUS.forEach(b => seg.appendChild(bt(b.nome, () => { f.bu = b.bu; f.obraId = ""; f.contratoId = ""; tudo(); }, f.bu === b.bu ? "on" : ""))); corpo.appendChild(seg);
      const q = inp(S.buscaObra || "", v => { S.buscaObra = v; desenha(); }, "Buscar obra"); q.style.margin = "10px 0"; q.style.width = "100%"; corpo.appendChild(q);
      const g = el("div", "fi-opcoes"); corpo.appendChild(g);
      const desenha = () => { g.textContent = ""; S.base.obras.filter(o => o.bu === f.bu && (!S.buscaObra || norm(o.nome).indexOf(norm(S.buscaObra)) >= 0)).sort((a, b) => a.nome.localeCompare(b.nome)).forEach(o => {
        const so = situacaoObra(S.base, S.lanc, o.id), p = o.budget ? so.usado / o.budget : null, b = el("button", "fi-op" + (f.obraId === o.id ? " on" : "")); b.type = "button";
        b.addEventListener("click", () => { f.obraId = o.id; f.contratoId = ""; f.novoContrato = false; f.passo = 2; S.msg = ""; tudo(); });
        b.appendChild(el("b", null, o.nome)); const br2 = el("span", "fi-mini-barra"), i = el("i"); i.style.width = Math.min(100, (p || 0) * 100) + "%"; i.style.background = p == null ? "var(--fi-neutro)" : corConsumo(p); br2.appendChild(i); b.appendChild(br2);
        b.appendChild(el("span", "fi-aviso", p == null ? "sem budget cadastrado" : "saldo " + moeda(so.saldo) + " · " + Math.round(p * 100) + "% usado")); g.appendChild(b); }); };
      desenha();
    } else if (f.passo === 2) {
      corpo.appendChild(el("p", "fi-pergunta", "Qual contrato de " + ob.nome + "?"));
      const g = el("div", "fi-opcoes"); corpo.appendChild(g);
      S.base.contratos.filter(x => x.obraId === f.obraId).sort((a, b) => String(a.disciplina).localeCompare(String(b.disciplina))).forEach(x => {
        const s2 = situacaoContrato(S.base, S.lanc.filter(y => y.id !== f.id), x.id), p = x.valor ? s2.usado / x.valor : 0, b = el("button", "fi-op" + (f.contratoId === x.id ? " on" : "")); b.type = "button";
        b.addEventListener("click", () => { f.contratoId = x.id; f.novoContrato = false; f.passo = 3; S.msg = ""; tudo(); });
        b.appendChild(el("b", null, x.disciplina)); b.appendChild(el("span", null, (F(x.fornecedorId).nome || "sem fornecedor") + (x.pedido ? " · pedido " + x.pedido : "")));
        const br2 = el("span", "fi-mini-barra"), i = el("i"); i.style.width = Math.min(100, p * 100) + "%"; i.style.background = corConsumo(p); br2.appendChild(i); b.appendChild(br2);
        b.appendChild(el("span", "fi-aviso", "saldo " + moeda(s2.saldo) + " de " + moeda(x.valor))); g.appendChild(b); });
      const nv = el("button", "fi-op novo" + (f.novoContrato ? " on" : "")); nv.type = "button"; nv.append(el("b", null, "+ Novo contrato"), el("span", "fi-aviso", "Disciplina e fornecedor ainda sem contrato nesta obra"));
      nv.addEventListener("click", () => { f.novoContrato = true; f.contratoId = ""; tudo(); }); g.appendChild(nv);
      if (f.novoContrato) { const n = el("div", "fi-card"); n.style.marginTop = "10px";
        const disc = [...new Set(S.base.contratos.map(x => x.disciplina).filter(Boolean))].sort();
        const di = inp(f.disciplina, v => { f.disciplina = v; }, "Ex.: Estrutura"); di.setAttribute("list", "fi-disc"); const dl = el("datalist"); dl.id = "fi-disc"; disc.forEach(d => { const o = el("option"); o.value = d; dl.appendChild(o); });
        const l1 = campo("Disciplina", di); l1.appendChild(dl); n.appendChild(l1);
        n.appendChild(campo("Fornecedor", sel([["", "Selecione"]].concat(S.base.fornecedores.slice().sort((a, b) => a.nome.localeCompare(b.nome)).map(x => [x.id, x.nome + (x.cnpj ? " · " + x.cnpj : "")])), f.fornecedorId, v => { f.fornecedorId = v; })));
        const vc = inp(f.valorContrato == null ? "" : String(f.valorContrato).replace(".", ","), v => { f.valorContrato = num(v); }, "R$ total contratado"); vc.inputMode = "decimal"; n.appendChild(campo("Valor do contrato", vc));
        n.appendChild(bt("Fornecedor não está na lista? Cadastrar", () => { S.fornSel = { id: novoId("f"), nome: "", razao: "", fantasia: "", cnpj: "", mega: "", contato: "", telefone: "", email: "" }; S.fornNovo = true; tudo(); }, "fi-link"));
        corpo.appendChild(n); }
    } else if (f.passo === 3) {
      corpo.appendChild(el("p", "fi-pergunta", "O que está sendo pago?"));
      const seg = el("div", "fi-seg"); [["Pedido", "Pedido"], ["Aditivo", "Aditivo"]].forEach(o => seg.appendChild(bt(o[1], () => { f.tipo = o[0]; tudo(); }, f.tipo === o[0] ? "on" : ""))); corpo.appendChild(seg);
      const ca = el("textarea"); ca.rows = 2; ca.value = f.causa || ""; ca.placeholder = "Ex.: 2ª etapa do projeto estrutural, revisão de fundação…"; ca.addEventListener("input", () => { f.causa = ca.value; });
      const lca = campo("Causa raiz do pedido", ca); lca.style.marginTop = "12px"; corpo.appendChild(lca);
      corpo.appendChild(el("p", "fi-pergunta", "Como vai ser pago?"));
      const dv = el("div", "fi-dividir"), nT = inp("", () => {}, "Valor total R$"), nN = inp("", () => {}, "Parcelas"), d0 = inp((f.parcelas[0] && f.parcelas[0].data) || somaMes(hojeISO(), 1), () => {}, "", "date");
      nT.inputMode = "decimal"; nN.inputMode = "numeric";
      dv.append(campo("Valor total", nT), campo("Nº de parcelas", nN), campo("1º pagamento", d0), bt("Gerar parcelas", () => { const t = num(nT.value), n = parseInt(nN.value, 10) || 1; if (!t) { alert("Informe o valor total."); return; }
        const b0 = Math.floor(t / n * 100) / 100; f.parcelas = Array.from({ length: n }, (_, i) => ({ valor: i === n - 1 ? Math.round((t - b0 * (n - 1)) * 100) / 100 : b0, data: somaMes(d0.value || hojeISO(), i), status: "PREVISTO" })); tudo(); }, "pri"));
      corpo.appendChild(dv);
      const pl = el("div", "fi-parcelas"); f.parcelas.forEach((p, i) => { const r = el("div", "fi-parc"); r.appendChild(el("b", null, (i + 1) + "ª"));
        const v = inp(p.valor == null ? "" : String(p.valor).replace(".", ","), x => { p.valor = num(x); }, "Valor R$"); v.inputMode = "decimal"; v.addEventListener("change", tudo);
        r.append(v, inp(p.data, x => { p.data = x; }, "", "date"), bt("×", () => { f.parcelas.splice(i, 1); tudo(); }, "fi-link")); pl.appendChild(r); });
      corpo.appendChild(pl);
      corpo.appendChild(bt("+ parcela", () => { const u = f.parcelas[f.parcelas.length - 1]; f.parcelas.push({ valor: null, data: u && u.data ? somaMes(u.data, 1) : hojeISO(), status: "PREVISTO" }); tudo(); }, "fi-link"));
      const obs = el("textarea"); obs.rows = 2; obs.value = f.obs || ""; obs.addEventListener("input", () => { f.obs = obs.value; }); const lo = campo("Observação (opcional)", obs); lo.style.marginTop = "10px"; corpo.appendChild(lo);
    } else {
      const e = estouroDe(f), c2 = ct || { disciplina: f.disciplina, fornecedorId: f.fornecedorId, valor: num(f.valorContrato) }, fo = F(c2.fornecedorId), tot = totalForm(f);
      corpo.appendChild(el("p", "fi-pergunta", "Confira antes de finalizar"));
      const rs = el("div", "fi-resumo");
      [["Obra", ob.nome], ["Contrato", c2.disciplina + " · " + (fo.nome || "—")], ["CNPJ", fo.cnpj || "não cadastrado"], ["Tipo", f.tipo], ["Causa raiz", f.causa], ["Total", moeda(tot) + " em " + f.parcelas.length + " parcela(s)"]]
        .forEach(x => { const r = el("div"); r.append(el("span", null, x[0]), el("b", null, x[1] || "—")); rs.appendChild(r); });
      corpo.appendChild(rs);
      const imp = (rot, antes, depois, total) => { const c = el("div", "fi-imp"); c.appendChild(el("span", "fi-campo", rot)); const b = el("div", "fi-imp-b"), a = el("i", "a"), n = el("i", "n");
        const mx = Math.max(total || 0, depois) || 1; a.style.width = Math.min(100, antes / mx * 100) + "%"; n.style.left = Math.min(100, antes / mx * 100) + "%"; n.style.width = Math.max(0, Math.min(100, depois / mx * 100) - antes / mx * 100) + "%";
        n.style.background = depois > (total || Infinity) ? "var(--fi-crit)" : "var(--fs2)"; b.append(a, n); if (total) { const m = el("em"); m.style.left = Math.min(100, total / mx * 100) + "%"; b.appendChild(m); } c.appendChild(b);
        c.appendChild(el("span", "fi-aviso", "Antes " + moeda(antes) + " · depois " + moeda(depois) + (total ? " · limite " + moeda(total) : " · sem limite cadastrado"))); return c; };
      if (e.sc) corpo.appendChild(imp("Contrato", e.sc.usado - tot, e.sc.usado, e.sc.valor));
      if (e.sc && e.sc.valor == null) corpo.appendChild(el("p", "fi-alerta am", "Este contrato está sem valor cadastrado: o controle de estouro do contrato não se aplica. Peça ao administrador para corrigir em Base → Contratos."));
      if (e.so) corpo.appendChild(imp("Budget da obra", e.so.usado - tot, e.so.usado, e.so.budget));
      if (e.contrato || e.obra) corpo.appendChild(el("p", "fi-alerta", "Estouro: " + [e.contrato ? "contrato em " + moeda(e.contrato) : "", e.obra ? "budget da obra em " + moeda(e.obra) : ""].filter(Boolean).join(" · ") +
        ". Ao finalizar, o e-mail sai com o aviso e o link do formulário de composição, e o coordenador vê o estouro pendente."));
      else corpo.appendChild(el("p", "fi-msg", "Dentro do contrato e do budget."));
      corpo.appendChild(el("p", "fi-aviso", "Ao finalizar, o Outlook abre com o pedido para " + (S.base.email || EMAIL_PADRAO).para + ". Confira e clique em Enviar."));
    }
    if (S.msg && S.erro) corpo.insertBefore(el("p", "fi-msg err", S.msg), corpo.firstChild);
    const rod = el("div", "fi-dw-rod");
    rod.appendChild(f.passo > 1 ? bt("← Voltar", () => { f.passo--; S.msg = ""; tudo(); }, "") : el("span"));
    const dir = el("div", "fi-l"); dir.appendChild(bt("Salvar rascunho", async () => { S.ocupado = true; tudo(); const r = await gravaLanc(S.form); S.ocupado = false; r.ok ? ok("Rascunho salvo em Lançamentos.") : falha(r.msg); S.form = null; tudo(); }, ""));
    if (f.passo < 4) { const nx = bt("Continuar →", () => { if (!podeIr(f.passo + 1)) { falha(["", "Escolha a obra.", "Escolha o contrato ou preencha o novo contrato.", "Preencha a causa raiz e o valor e a data de cada parcela."][f.passo]); tudo(); return; } S.msg = ""; f.passo++; tudo(); }, "pri"); dir.appendChild(nx); }
    else { const fz = bt("Finalizar e gerar e-mail", finalizar, "pri"); fz.disabled = S.ocupado; dir.appendChild(fz); }
    rod.appendChild(dir); dw.appendChild(rod);
    ov.appendChild(dw); return ov;
  }
  function lancamentos() {
    const c = el("div", "fi-card"), lt = el("div", "fi-l fi-sb");
    lt.append(el("h3", null, "Lançamentos feitos no painel"), inp(S.busca, v => { S.busca = v; desenha(); }, "Buscar nº, causa, fornecedor")); c.appendChild(lt);
    const box = el("div", "fi-tab"); box.style.marginTop = "10px"; c.appendChild(box);
    function desenha() {
      const q = norm(S.busca), ls = S.lanc.filter(x => x.status !== "excluido" && passa(x))
        .filter(x => !q || norm([x.numero, discDe(x), F(x.fornecedorId).nome, O(x.obraId).nome, x.causa].join(" ")).indexOf(q) >= 0).sort((a, b) => (b.numero || 1e9) - (a.numero || 1e9));
      box.textContent = "";
      if (!ls.length) { box.appendChild(el("p", "fi-aviso", "Nenhum lançamento neste filtro. O histórico importado da planilha está em Todos os pagamentos.")); return; }
      box.appendChild(tabela(["Nº", "Obra", "Disciplina", "Fornecedor", "Tipo", "Causa raiz", "Total", "Parcelas", "Situação", ""], ls.map(x => {
        const r = el("tr"), tot = (x.parcelas || []).reduce((a, p) => a + (num(p.valor) || 0), 0), pg = (x.parcelas || []).filter(p => p.status === "PAGO").length;
        td(r, x.numero || "—"); td(r, O(x.obraId).nome); td(r, discDe(x) || "—"); td(r, F(x.fornecedorId).nome || "—"); td(r, x.tipo); td(r, x.causa || "—"); td(r, moeda(tot), "n"); td(r, pg + "/" + (x.parcelas || []).length + " pagas");
        const st = el("td"); st.appendChild(el("span", "fi-pill " + (x.status === "finalizado" ? (x.estouro ? "fi-no" : "fi-ok") : x.status === "cancelado" ? "fi-ci" : "fi-am"),
          x.status === "finalizado" ? (x.estouro ? "Finalizado · estouro" : "Finalizado") : x.status === "cancelado" ? "Cancelado" : "Rascunho")); r.appendChild(st);
        const ac = el("td"); ac.style.whiteSpace = "nowrap";
        if (x.status === "rascunho") { ac.appendChild(bt("Editar", () => { S.form = copia(x); tudo(); }, "mini"));
          ac.appendChild(bt("Excluir", async () => { if (!window.confirm("Excluir este rascunho?")) return; x.status = "excluido"; await gravaLanc(x); tudo(); }, "mini per")); }
        else if (x.status === "finalizado") { ac.appendChild(bt("Reenviar e-mail", () => abrir(mailto(montaEmail(S.base, x, x.estouroId ? link(x.estouroId) : ""))), "mini"));
          if (x.estouroId) { const es = S.est.find(e => e.id === x.estouroId); if (es) ac.appendChild(bt(es.status === "pendente" ? "Responder estouro" : "Ver estouro", () => { S.estForm = es; tudo(); }, "mini")); }
          if (adm) ac.appendChild(bt("Cancelar", async () => { if (!window.confirm("Cancelar o lançamento nº " + x.numero + "? Ele sai do utilizado e das parcelas.")) return; x.status = "cancelado"; await gravaLanc(x); tudo(); }, "mini per")); }
        r.appendChild(ac); return r; })));
    }
    desenha(); return c;
  }

  // ── estouros ──
  function estouros() {
    const c = el("div", "fi-card"), es = S.est.filter(passa).sort((a, b) => (a.status === "pendente" ? 0 : 1) - (b.status === "pendente" ? 0 : 1) || (b.numero || 0) - (a.numero || 0));
    c.appendChild(el("h3", null, "Base de motivos de estouro"));
    const cont = {}; es.filter(e => e.resposta && e.resposta.causa).forEach(e => { cont[e.resposta.causa] = (cont[e.resposta.causa] || 0) + Math.max(e.contrato || 0, e.obra || 0); });
    const ks = Object.keys(cont).sort((a, b) => cont[b] - cont[a]);
    if (ks.length) { const k = el("div", "fi-grade"); k.style.marginTop = "10px"; ks.slice(0, 4).forEach(x => k.appendChild(kpi(x, moeda(cont[x]), "#EF4444", "maior excesso somado"))); c.appendChild(k); }
    const t = tabela(["Nº", "Obra", "Disciplina · fornecedor", "Excesso", "Analista", "Causa", "Status", ""], es.map(e => { const r = el("tr");
      td(r, e.numero || "—"); td(r, O(e.obraId).nome); td(r, e.disciplina + " · " + (F(e.fornecedorId).nome || ""));
      td(r, [e.contrato ? "contrato " + moeda(e.contrato) : "", e.obra ? "obra " + moeda(e.obra) : ""].filter(Boolean).join(" · ") || "—", "n"); td(r, e.analista || "—"); td(r, (e.resposta && e.resposta.causa) || "—");
      const st = el("td"); st.appendChild(el("span", "fi-pill " + (e.status === "pendente" ? "fi-no" : "fi-ok"), e.status === "pendente" ? "Aguardando analista" : "Respondido")); r.appendChild(st);
      const ac = el("td"); ac.style.whiteSpace = "nowrap"; ac.appendChild(bt(e.status === "pendente" ? "Responder" : "Ver", () => { S.estForm = e; tudo(); }, "mini"));
      ac.appendChild(bt("Copiar link", async () => { try { await navigator.clipboard.writeText(link(e.id)); ok("Link do formulário copiado."); } catch (x) { window.prompt("Copie o link:", link(e.id)); } tudo(); }, "mini"));
      r.appendChild(ac); return r; }));
    t.style.marginTop = "12px"; c.appendChild(es.length ? t : el("p", "fi-aviso", "Nenhum estouro neste filtro.")); return c;
  }
  function telaEstouro() {
    const e = S.estForm, r = e.resposta = e.resposta || { composicao: [{ item: "", valor: null }], causa: "", descricao: "", acao: "" };
    const md = el("div", "fi-modal"), c = el("div", "fi-card fi");
    const fecha = () => { S.estForm = null; if (location.hash.indexOf("#estouro=") === 0) history.replaceState(null, "", location.href.split("#")[0]); tudo(); };
    const top = el("div", "fi-l fi-sb"); top.append(el("h3", null, "Composição do estouro · lançamento nº " + (e.numero || "—")), bt("Fechar", fecha)); c.appendChild(top);
    const k = el("div", "fi-grade");
    k.appendChild(kpi("Empreendimento", O(e.obraId).nome || "—", "#2563EB", e.disciplina + " · " + (F(e.fornecedorId).nome || "")));
    if (e.contrato) k.appendChild(kpi("Acima do contrato", moeda(e.contrato), "#EF4444")); if (e.obra) k.appendChild(kpi("Acima do budget da obra", moeda(e.obra), "#EF4444"));
    c.appendChild(k);
    const resp = e.status === "respondido" && !adm;
    c.appendChild(el("div", "fi-campo", "Composição do valor excedente"));
    r.composicao.forEach((x, i) => { const l = el("div", "fi-l");
      const a = inp(x.item, v => { x.item = v; }, "Item (ex.: revisão de fundação após sondagem)"); a.style.flex = "1"; a.disabled = resp;
      const b = inp(x.valor == null ? "" : String(x.valor).replace(".", ","), v => { x.valor = num(v); }, "R$"); b.style.width = "130px"; b.inputMode = "decimal"; b.disabled = resp;
      l.append(a, b); if (!resp) l.appendChild(bt("×", () => { r.composicao.splice(i, 1); tudo(); }, "mini")); c.appendChild(l); });
    if (!resp) c.appendChild(bt("+ item", () => { r.composicao.push({ item: "", valor: null }); tudo(); }, "mini"));
    const cs = sel([["", "Selecione"]].concat(CAUSAS.map(x => [x, x])), r.causa, v => { r.causa = v; }); cs.disabled = resp; c.appendChild(campo("Causa raiz", cs));
    const ds = el("textarea"); ds.rows = 3; ds.value = r.descricao || ""; ds.disabled = resp; ds.addEventListener("input", () => { r.descricao = ds.value; }); c.appendChild(campo("O que aconteceu", ds));
    const ac = el("textarea"); ac.rows = 2; ac.value = r.acao || ""; ac.disabled = resp; ac.addEventListener("input", () => { r.acao = ac.value; }); c.appendChild(campo("Ação para não repetir", ac));
    if (!resp) { const l = el("div", "fi-l"); l.style.justifyContent = "flex-end"; l.appendChild(bt("Enviar resposta", async () => {
        if (!r.causa || !String(r.descricao || "").trim() || !r.composicao.some(x => x.item && num(x.valor) != null)) { alert("Preencha a composição, a causa raiz e o que aconteceu."); return; }
        e.status = "respondido"; r.em = hojeISO(); r.por = opts.usuario || ""; const x = await gravaEst(e); x.ok ? ok("Resposta registrada na base de motivos de estouro.") : falha(x.msg); fecha(); }, "pri")); c.appendChild(l); }
    md.appendChild(c); md.addEventListener("click", ev => { if (ev.target === md) fecha(); }); return md;
  }

  // ── base e orçamento (admin) ──
  function baseOrc() {
    const c = el("div", "fi-card"), l = el("div", "fi-l");
    [["obras", "Obras e budget"], ["contratos", "Contratos"], ["imp", "Importar planilha SAÍDAS"]].forEach(x => l.appendChild(bt(x[1], () => { S.baseAba = x[0]; tudo(); }, S.baseAba === x[0] ? "on" : "")));
    l.appendChild(el("span", "fi-aviso", "Fornecedores ficam no botão Fornecedores (cadastro único).")); c.appendChild(l);
    const box = el("div", "fi-tab"); box.style.marginTop = "12px"; c.appendChild(box);
    const nIn = (o, k, w) => { const i = inp(o[k] == null ? "" : String(o[k]).replace(".", ","), v => { o[k] = num(v); }); i.inputMode = "decimal"; i.style.width = (w || 120) + "px"; return i; };
    const tIn = (o, k, w) => { const i = inp(o[k], v => { o[k] = v; }); i.style.width = (w || 160) + "px"; return i; };
    const cel = (r, ctrl) => { const t = el("td"); t.appendChild(ctrl); r.appendChild(t); };
    if (S.baseAba === "obras") {
      box.appendChild(tabela(["Unidade", "Obra", "GDS", "SPE", "CNPJ da obra", "UH", "m²", "VGV", "Budget disp.", "Utilizado", "Apelidos (outros nomes)", ""], obrasF().map(o => { const r = el("tr"), so = situacaoObra(S.base, S.lanc, o.id);
        td(r, (BUS.find(b => b.bu === o.bu) || {}).nome); cel(r, tIn(o, "nome", 140)); cel(r, tIn(o, "gds", 70)); cel(r, tIn(o, "spe", 200)); cel(r, tIn(o, "cnpj", 150)); cel(r, nIn(o, "uh", 70)); cel(r, nIn(o, "m2", 90)); cel(r, nIn(o, "vgv", 120)); cel(r, nIn(o, "budget", 110));
        td(r, moeda(so.usado), "n");
        const ap = inp((o.apelidos || []).join("; "), v => { o.apelidos = v.split(";").map(x => x.trim()).filter(Boolean); }, "Ex.: Quadratto; QUA"); ap.style.width = "180px"; cel(r, ap);
        cel(r, bt("×", () => { if (!window.confirm("Excluir a obra " + o.nome + "?")) return; S.base.obras = S.base.obras.filter(x => x !== o); tudo(); }, "mini per")); return r; })));
      box.appendChild(bt("+ obra", () => { S.base.obras.push({ id: novoId("o"), bu: S.f.bu || "Make It", nome: "Nova obra", gds: "", spe: "", cnpj: "", uh: null, m2: null, vgv: null, budget: null, apelidos: [] }); tudo(); }, "mini"));
    } else if (S.baseAba === "contratos") {
      const fo = S.base.fornecedores.slice().sort((a, b) => a.nome.localeCompare(b.nome)).map(f => [f.id, f.nome]);
      const cs = S.base.contratos.filter(passa).sort((a, b) => String(O(a.obraId).nome).localeCompare(String(O(b.obraId).nome)) || String(a.disciplina).localeCompare(String(b.disciplina)));
      box.appendChild(tabela(["Obra", "Disciplina", "Fornecedor", "Pedido ERP", "Valor contratado", "Lançado", "Saldo", ""], cs.slice(0, 400).map(x => { const r = el("tr"), s2 = situacaoContrato(S.base, S.lanc, x.id);
        cel(r, sel(S.base.obras.filter(o => o.bu === x.bu).map(o => [o.id, o.nome]), x.obraId, v => { x.obraId = v; })); cel(r, tIn(x, "disciplina", 160));
        cel(r, sel([["", "—"]].concat(fo), x.fornecedorId, v => { x.fornecedorId = v; })); cel(r, tIn(x, "pedido", 80)); cel(r, nIn(x, "valor"));
        td(r, moeda(s2.usado), "n"); const sd = td(r, moeda(s2.saldo), "n"); if (s2.saldo < 0) sd.style.color = "#B91C1C";
        cel(r, bt("×", () => { if (!window.confirm("Excluir este contrato?")) return; S.base.contratos = S.base.contratos.filter(y => y !== x); tudo(); }, "mini per")); return r; })));
      if (cs.length > 400) box.appendChild(el("p", "fi-aviso", "Mostrando 400 de " + cs.length + ". Use os filtros."));
      box.appendChild(bt("+ contrato", () => { const o = obrasF()[0]; if (!o) return; S.base.contratos.push({ id: novoId("c"), bu: o.bu, obraId: S.f.obra || o.id, disciplina: "", fornecedorId: S.f.forn || "", valor: 0, pedido: "" }); tudo(); }, "mini"));
    } else {
      const bu = S.f.bu; box.appendChild(el("p", "fi-aviso", "Importa a planilha SAÍDAS de uma unidade: obras e budget, contratos (com CNPJ e Nº de pedido), fornecedores e pagamentos. Substitui obras, contratos e histórico da unidade; lançamentos do painel e apelidos de obra são mantidos; fornecedores são somados ao cadastro único."));
      const sb = sel(BUS.map(b => [b.bu, b.nome]), bu || "Make It", v => { S.impBu = v; }); S.impBu = S.impBu || bu || "Make It"; sb.value = S.impBu;
      const fi = el("input"); fi.type = "file"; fi.accept = ".xlsx"; const lg = el("div", "fi-l"); lg.style.marginTop = "8px"; lg.append(campo("Unidade", sb), fi); box.appendChild(lg);
      Object.keys(S.base.importado || {}).forEach(k => box.appendChild(el("p", "fi-aviso", "Última importação " + ((BUS.find(b => b.bu === k) || {}).nome || k) + ": " + br(S.base.importado[k]))));
      const bi = bt("Importar", async () => { const f = fi.files && fi.files[0]; if (!f) { alert("Escolha o arquivo .xlsx."); return; } const ub = S.impBu;
        if (S.base.obras.some(o => o.bu === ub) && !window.confirm("Substituir obras, contratos e histórico de " + ub + " pelos da planilha?")) return;
        try { S.ocupado = true; tudo(); const wb = await lerPlanilha(f); const r = importarSaidas(wb, ub, S.base); S.base = r.base;
          await gravaBase("Importado (" + ub + "): " + r.n.obras + " obras, " + r.n.contratos + " contratos e " + r.n.historico + " pagamentos." + (r.avisos.length ? " " + r.avisos.join(" ") : "")); }
        catch (e) { S.ocupado = false; falha("Não foi possível importar: " + e.message); tudo(); } }, "pri"); bi.style.marginTop = "8px"; box.appendChild(bi);
    }
    if (S.baseAba !== "imp") { const sv = el("div", "fi-l"); sv.style.cssText = "margin-top:12px;justify-content:flex-end"; const b = bt("Salvar base", () => gravaBase("Base salva. Já vale para todos."), "pri"); b.disabled = S.ocupado; sv.appendChild(b); c.appendChild(sv); }
    return c;
  }
  function modeloEmail() {
    const c = el("div", "fi-card"), e = S.base.email;
    c.appendChild(el("h3", null, "Modelo do e-mail de pedido"));
    c.appendChild(el("p", "fi-aviso", "Clique em um campo para inserir no corpo. [Nº PEDIDO] é o número sequencial do lançamento; [PEDIDO ERP] é o Nº de pedido do contrato na planilha."));
    const g = el("div", "fi-grade"); g.style.marginTop = "10px"; g.append(campo("Para", inp(e.para, v => { e.para = v; })), campo("Cc", inp(e.cc, v => { e.cc = v; }))); c.appendChild(g);
    const ti = inp(e.titulo, v => { e.titulo = v; }); const lt = campo("Título", ti); lt.style.marginTop = "10px"; c.appendChild(lt);
    const co = el("textarea"); co.rows = 12; co.value = e.corpo; co.addEventListener("input", () => { e.corpo = co.value; }); const lc = campo("Corpo", co); lc.style.marginTop = "10px"; c.appendChild(lc);
    const ch = el("div", "fi-l"); ch.style.marginTop = "8px";
    CAMPOS.forEach(k => ch.appendChild(bt(k, () => { const p = co.selectionStart || co.value.length; co.value = co.value.slice(0, p) + k + co.value.slice(co.selectionEnd || p); e.corpo = co.value; co.focus(); }, "fi-chip"))); c.appendChild(ch);
    const ex = S.lanc.filter(x => x.status === "finalizado").slice(-1)[0];
    if (ex) { const pv = montaEmail(S.base, ex, ex.estouroId ? link(ex.estouroId) : ""), p = el("pre"); p.style.cssText = "white-space:pre-wrap;font-size:12px;background:var(--slate-50,#F8FAFC);padding:12px;border-radius:10px;margin-top:12px";
      p.textContent = "Prévia com o lançamento nº " + ex.numero + "\nPara: " + pv.para + "\nCc: " + pv.cc + "\nTítulo: " + pv.titulo + "\n\n" + pv.corpo; c.appendChild(p); }
    const a = el("div", "fi-l"); a.style.cssText = "margin-top:12px;justify-content:flex-end";
    a.append(bt("Restaurar padrão", () => { S.base.email = copia(EMAIL_PADRAO); tudo(); }), bt("Salvar modelo", () => gravaBase("Modelo de e-mail salvo."), "pri")); c.appendChild(a); return c;
  }
  function exportarBloco() {
    const c = el("div", "fi-l"); c.style.justifyContent = "flex-end";
    BUS.filter(b => !S.f.bu || b.bu === S.f.bu).forEach(b => c.appendChild(bt("Exportar planilha " + b.nome + " (.xlsx)", () => {
      try { const r = exportar(S.base, S.lanc, S.est, b.bu); const a = document.createElement("a"); a.href = URL.createObjectURL(r.blob); a.download = r.nome; document.body.appendChild(a); a.click();
        setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000); ok("Planilha gerada: " + r.nome); } catch (e) { falha("Erro ao exportar: " + e.message); } tudo(); }, "mini")));
    return c;
  }
  function tudo() {
    wrap.textContent = ""; wrap.appendChild(topo());
    if (S.msg) wrap.appendChild(el("p", "fi-msg" + (S.erro ? " err" : ""), S.msg));
    const pend = S.est.filter(e => e.status === "pendente"), meus = pend.filter(e => e.analista && e.analista === opts.usuario && !adm);
    if (meus.length) wrap.appendChild(el("p", "fi-alerta", "Você tem " + meus.length + " formulário(s) de estouro para responder. Veja em Estouros."));
    wrap.appendChild(({ inicio, lanc: lancamentos, atr: () => pagamentos(true), pag: () => pagamentos(false), forn: fornecedores, est: estouros, base: adm ? baseOrc : inicio, email: adm ? modeloEmail : inicio }[S.view] || inicio)());
    if (["inicio", "pag", "atr", "lanc"].indexOf(S.view) >= 0) wrap.appendChild(exportarBloco());
    escondeDica();
    if (S.form) { const y = window.scrollY; wrap.appendChild(telaForm()); window.scrollTo(0, y); }
    if (S.fornSel) wrap.appendChild(fichaFornecedor());
    if (S.estForm) wrap.appendChild(telaEstouro());
  }
  const hs = location.hash.match(/^#estouro=([\w-]+)/);
  if (hs) { const e = S.est.find(x => x.id === hs[1]); if (e) { S.estForm = e; S.view = "est"; } else falha("Formulário de estouro não encontrado. Confira o link ou peça um novo ao coordenador."); }
  if (opts.view) S.view = opts.view;
  tudo();
}
