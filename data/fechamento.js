// Fechamento mensal dos indicadores (Projeto Next) e exportação no modelo do setor de Processos.
// Módulo isolado: erro aqui não afeta as outras abas. Só o admin edita; quem grava é o painel (opts.salvar).

export const PROCESSOS = [
  { nome: "Incorporação Make It", bu: "Make It", curto: "Make It" },
  { nome: "Incorporação Gadens", bu: "Gadens Inc", curto: "Gadens" }
];
export const INDICADORES = [
  { n: 1, nome: "Riscos de Projetos", sentido: "Maior é melhor", meta: 100,
    calculo: "Nota por empreendimento de PE em diante, pelos riscos em aberto com score (impacto x probabilidade) maior que 12: até 1 = 100; 2 a 3 = 50; 4 = 10; 5 ou mais = 0. Resultado = média das notas dos empreendimentos." },
  { n: 2, nome: "Atendimento a Linha de Base de Pré-executivo", sentido: "Maior é melhor", meta: 100,
    calculo: "Por empreendimento em Pré-executivo: dias de atraso registrados divididos pela duração da fase. Até 30% = 100; de 30% a 70% cai até 10; acima, cai até 0. Resultado = média das notas." },
  { n: 3, nome: "Alterações de projeto por erro", sentido: "Maior é melhor", meta: 100,
    calculo: "Por empreendimento em obra: até 10 alterações = 100 menos 2 por alteração; de 10 a 20 cai até 50; de 20 a 30 cai até 10; acima, cai até 0. Resultado = média das notas." },
  { n: 5, nome: "Qualidade do Projeto (EP)", sentido: "Maior é melhor", meta: 4,
    calculo: "Média das notas do checklist do projetista (1 a 5) por empreendimento. Resultado = média dos empreendimentos." }
];
const MESES = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"];
const rotulo = k => MESES[+k.slice(5, 7) - 1] + "/" + k.slice(0, 4);
const copia = o => JSON.parse(JSON.stringify(o));
const num = v => { if (v == null || v === "") return null; const x = typeof v === "number" ? v : parseFloat(String(v).replace(",", ".")); return isNaN(x) ? null : x; };
const fmt = v => v == null ? "—" : String(Math.round(v * 100) / 100).replace(".", ",");

// Out/2026: resultados enviados ao setor de Processos. Nº 3 segue o cálculo do painel.
export function semente() {
  const t = (fato, causa, acao) => ({ fato, causa, acao });
  return { versao: 1, meses: { "2026-10": { status: "fechado", por: "Planilha enviada a Processos", atualizado: "2026-11-07", proc: {
    "Incorporação Make It": { itens: {
      1: Object.assign({ resultado: 75, porObra: [{ e: "Alba", v: 50 }, { e: "Horizon", v: 100 }] },
        t("Empreendimento Alba trouxe o indicador para baixo", "Foi levado para executivo a incerteza relacionada ao projeto pré-executivo, o que impactou consideravelmente a nota", "Estamos resolvendo o quanto antes o risco para recuperar a nota")),
      2: Object.assign({ resultado: 71, porObra: [{ e: "Horizon", v: 71 }] },
        t("Empreendimento Horizon impactou na nota baixa", "Tivemos alteração da fundação em etapa de pré-executivo", "Teremos diretrizes de padrões de soluções para que não tenhamos essa tipologia de alteração em pré-executivo")),
      3: Object.assign({ resultado: null, porObra: [], usarCalculo: true },
        t("Empreendimento Essence impactou na nota baixa", "Tivemos várias alterações de projeto ao longo da obra Essence, o que impactou consideravelmente a nota", "Os comitês de projetos são para ajustar esse item de projeto")),
      5: Object.assign({ resultado: 3.93, porObra: [{ e: "Alba", v: 4.6 }, { e: "Uniq", v: 3.8 }, { e: "Lume", v: 3.2 }, { e: "Essence", v: 3 }, { e: "Now", v: 4.4 }, { e: "Horizon", v: 4.1 }, { e: "Al Nasser I", v: 4.2 }, { e: "Al Nasser II", v: 4.2 }] },
        t("Empreendimentos que abaixaram consideravelmente a nota da qualidade: Essence e Lume", "Muitas entregas insatisfatórias e fora do prazo de entrega, principalmente no Essence", "Estamos criando as diretrizes de projetos para que possamos deixar com maior clareza o que está sendo projetado"))
    } }
  } } } };
}

export function situacao(it, def) {
  if (it.situacao) return it.situacao;
  const r = num(it.resultado), m = num(it.meta != null ? it.meta : def.meta);
  if (r == null || m == null) return "Em implementação";
  return (def.sentido === "Menor é melhor" ? r <= m : r >= m) ? "Atingida" : "Não atingida";
}
function normaliza(d) {
  const o = d && d.meses ? copia(d) : semente();
  if (!o.meses["2026-10"]) o.meses["2026-10"] = semente().meses["2026-10"];
  return o;
}
function chaveMes(dt) { return dt.getFullYear() + "-" + String(dt.getMonth() + 1).padStart(2, "0"); }

// ── leitura e escrita de .xlsx (zip) ──
async function unzip(buf) {
  const v = new DataView(buf), u8 = new Uint8Array(buf), dec = new TextDecoder();
  let e = buf.byteLength - 22; while (e >= 0 && v.getUint32(e, true) !== 0x06054b50) e--;
  if (e < 0) throw new Error("modelo .xlsx inválido");
  const n = v.getUint16(e + 10, true); let p = v.getUint32(e + 16, true); const out = [];
  for (let i = 0; i < n; i++) {
    const met = v.getUint16(p + 10, true), csz = v.getUint32(p + 20, true), ln = v.getUint16(p + 28, true), lx = v.getUint16(p + 30, true), lc = v.getUint16(p + 32, true), off = v.getUint32(p + 42, true);
    const nome = dec.decode(u8.subarray(p + 46, p + 46 + ln));
    const ini = off + 30 + v.getUint16(off + 26, true) + v.getUint16(off + 28, true), raw = u8.subarray(ini, ini + csz);
    const data = met === 0 ? raw.slice() : new Uint8Array(await new Response(new Blob([raw]).stream().pipeThrough(new DecompressionStream("deflate-raw"))).arrayBuffer());
    if (!nome.endsWith("/")) out.push({ nome, data });
    p += 46 + ln + lx + lc;
  }
  return out;
}
let TAB = null;
function crc32(b) {
  if (!TAB) { TAB = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; TAB[n] = c >>> 0; } }
  let c = 0xffffffff; for (let i = 0; i < b.length; i++) c = TAB[(c ^ b[i]) & 0xff] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0;
}
function zip(arqs) {
  const enc = new TextEncoder(), partes = [], cent = []; let off = 0;
  arqs.forEach(a => {
    const nm = enc.encode(a.nome), d = a.data, crc = crc32(d);
    const h = new DataView(new ArrayBuffer(30)); h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0x0800, true);
    h.setUint32(14, crc, true); h.setUint32(18, d.length, true); h.setUint32(22, d.length, true); h.setUint16(26, nm.length, true);
    partes.push(new Uint8Array(h.buffer), nm, d);
    const c = new DataView(new ArrayBuffer(46)); c.setUint32(0, 0x02014b50, true); c.setUint16(4, 20, true); c.setUint16(6, 20, true); c.setUint16(8, 0x0800, true);
    c.setUint32(16, crc, true); c.setUint32(20, d.length, true); c.setUint32(24, d.length, true); c.setUint16(28, nm.length, true); c.setUint32(42, off, true);
    cent.push(new Uint8Array(c.buffer), nm); off += 30 + nm.length + d.length;
  });
  const tam = cent.reduce((a, x) => a + x.length, 0), fim = new DataView(new ArrayBuffer(22));
  fim.setUint32(0, 0x06054b50, true); fim.setUint16(8, arqs.length, true); fim.setUint16(10, arqs.length, true); fim.setUint32(12, tam, true); fim.setUint32(16, off, true);
  return new Blob(partes.concat(cent, [new Uint8Array(fim.buffer)]), { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}
const xesc = s => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const COLS = "ABCDEFGHIJKL";
function celula(col, r, v, s) {
  const ref = col + r;
  if (v == null || v === "") return '<c r="' + ref + '" s="' + s + '"/>';
  if (typeof v === "number") return '<c r="' + ref + '" s="' + s + '"><v>' + v + "</v></c>";
  return '<c r="' + ref + '" s="' + s + '" t="inlineStr"><is><t xml:space="preserve">' + xesc(v) + "</t></is></c>";
}
function preenche(xml, linhas, ncol, estilos) {
  const ult = Math.max(199, linhas.length + 2), cols = COLS.slice(0, ncol);
  let corpo = "";
  for (let r = 3; r <= ult; r++) {
    const l = linhas[r - 3];
    corpo += '<row r="' + r + '" spans="1:' + ncol + '"' + (l ? ' ht="30" customHeight="1"' : "") + ">" +
      cols.split("").map((c, i) => celula(c, r, l ? l[i] : null, l ? estilos[i] : 9)).join("") + "</row>";
  }
  const a = xml.indexOf("</row>", xml.indexOf('<row r="2"')) + 6, b = xml.indexOf("</sheetData>");
  xml = xml.slice(0, a) + corpo + xml.slice(b);
  const fimCol = cols[cols.length - 1];
  xml = xml.replace(/<dimension ref="[^"]*"\/>/, '<dimension ref="A1:' + fimCol + ult + '"/>')
    .replace(/(<autoFilter ref=")[^"]*"/, "$1A1:" + fimCol + ult + '"')
    .replace(/sqref="([A-L])\d+:[A-L]\d+(?: [A-L]\d+:[A-L]\d+)*"/g, (m, c) => 'sqref="' + c + "3:" + c + ult + '"');
  return xml;
}
export async function exportar(dados, mes, urlModelo) {
  const r = await fetch(urlModelo, { cache: "no-cache" });
  if (!r.ok) throw new Error("modelo da planilha não encontrado (" + urlModelo + ", HTTP " + r.status + ")");
  const arqs = await unzip(await r.arrayBuffer()), enc = new TextEncoder(), dec = new TextDecoder();
  const M = dados.meses[mes] || { proc: {} }, mesTxt = rotulo(mes), L1 = [], L2 = [];
  PROCESSOS.forEach(p => { const P = M.proc && M.proc[p.nome]; if (!P) return;
    INDICADORES.forEach(def => { const it = P.itens[def.n]; if (!it) return;
      const res = num(it.resultado), sit = situacao(it, def), dash = v => v && String(v).trim() ? String(v).trim() : "-";
      L1.push([p.nome, mesTxt, def.n, it.nome || def.nome, it.calculo || def.calculo, def.sentido, num(it.meta != null ? it.meta : def.meta),
        res == null ? "A implantar" : Math.round(res * 100) / 100, sit, dash(it.fato), dash(it.causa), dash(it.acao)]);
      (it.porObra || []).forEach(o => { const v = num(o.v); if (o.e && v != null) L2.push([p.nome, mesTxt, def.n, it.nome || def.nome, o.e, Math.round(v * 100) / 100]); });
    }); });
  if (!L1.length) throw new Error("não há lançamentos em " + mesTxt);
  const est1 = [9, 9, 9, 3, 3, 9, 9, 9, 9, 3, 3, 3], est2 = [9, 9, 9, 3, 9, 9];
  arqs.forEach(a => {
    if (a.nome === "xl/worksheets/sheet2.xml") a.data = enc.encode(preenche(dec.decode(a.data), L1, 12, est1));
    if (a.nome === "xl/worksheets/sheet3.xml") a.data = enc.encode(preenche(dec.decode(a.data), L2, 6, est2));
  });
  arqs.sort((x, y) => (x.nome === "[Content_Types].xml" ? -1 : y.nome === "[Content_Types].xml" ? 1 : 0));
  return { blob: zip(arqs), nome: "ProjetoNext_Indicadores_PROJETOS COMPLEMENTARES_" + mesTxt.replace("/", "-") + ".xlsx", linhas: L1.length, obras: L2.length };
}

// ── interface ──
const CSS = `
.fc{display:flex;flex-direction:column;gap:14px;color:var(--color-text,#0F172A)}
.fc *{box-sizing:border-box}
.fc-card{background:var(--card,#fff);border:1px solid var(--slate-100,#F1F5F9);border-radius:16px;box-shadow:var(--card-shadow);padding:16px 18px}
.fc-barra{display:flex;flex-wrap:wrap;gap:10px;align-items:center;justify-content:space-between}
.fc-g{display:flex;flex-wrap:wrap;gap:6px;align-items:center}
.fc h2{margin:0;font-size:17px;font-weight:900}
.fc h3{margin:0;font-size:14.5px;font-weight:800}
.fc button{font:inherit;font-size:12.5px;font-weight:700;cursor:pointer;border:1px solid var(--slate-200,#E2E8F0);background:var(--card,#fff);color:var(--slate-600,#475569);border-radius:10px;padding:7px 12px}
.fc button.on,.fc button.pri{background:var(--blue-600,#2563EB);border-color:var(--blue-600,#2563EB);color:#fff}
.fc button.per{color:#B91C1C}
.fc button:disabled{opacity:.45;cursor:default}
.fc select,.fc input,.fc textarea{font:inherit;font-size:13px;padding:7px 9px;border-radius:9px;border:1px solid var(--slate-200,#E2E8F0);background:var(--card,#fff);color:var(--color-text,#0F172A)}
.fc textarea{width:100%;resize:vertical}
.fc-pill{font-size:11px;font-weight:800;padding:3px 10px;border-radius:99px;white-space:nowrap}
.fc-ok{background:#D1FAE5;color:#065F46}.fc-no{background:#FEE2E2;color:#991B1B}.fc-imp{background:var(--slate-100,#F1F5F9);color:var(--slate-600,#475569)}
.fc-rasc{background:#FEF3C7;color:#92400E}.fc-fech{background:#DBEAFE;color:#1E40AF}
.fc-grade{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:10px;margin-top:12px}
.fc-campo{display:flex;flex-direction:column;gap:4px;font-size:11px;font-weight:800;letter-spacing:.04em;text-transform:uppercase;color:var(--slate-400,#94A3B8)}
.fc-campo input,.fc-campo select{width:100%}
.fc-val{font-size:20px;font-weight:900;text-transform:none;letter-spacing:0;color:var(--color-text,#0F172A)}
.fc-txt{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:10px;margin-top:10px}
.fc-txt p{margin:0;font-size:12.5px;line-height:1.45;color:var(--slate-600,#475569);text-transform:none;letter-spacing:0;font-weight:500;white-space:pre-wrap}
.fc-obras{margin-top:10px;display:flex;flex-wrap:wrap;gap:6px}
.fc-obra{display:inline-flex;align-items:center;gap:6px;font-size:12px;padding:4px 6px 4px 10px;border-radius:99px;background:var(--slate-50,#F8FAFC);border:1px solid var(--slate-100,#F1F5F9)}
.fc-obra b{font-family:var(--font-mono,monospace);font-size:12px}
.fc-obra input{width:64px;padding:3px 6px;font-size:12px}
.fc-obra input.nome{width:120px}
.fc-x{border:0!important;background:transparent!important;padding:0 4px!important;color:var(--slate-400,#94A3B8)!important}
.fc-aviso{margin:0;font-size:12px;line-height:1.45;color:var(--slate-500,#64748B)}
.fc-msg{margin:0;font-size:12.5px;padding:9px 12px;border-radius:10px;background:#ECFDF5;color:#065F46}.fc-msg.err{background:#FEF2F2;color:#991B1B}
.fc-calc{font-size:11.5px;color:var(--slate-500,#64748B);text-transform:none;letter-spacing:0;font-weight:600}
@media(max-width:900px){.fc-grade,.fc-txt{grid-template-columns:1fr 1fr}}
@media(max-width:560px){.fc-grade,.fc-txt{grid-template-columns:1fr}}
`;
const el = (t, c, x) => { const e = document.createElement(t); if (c) e.className = c; if (x != null) e.textContent = x; return e; };
const bt = (x, f, c) => { const b = el("button", c, x); b.type = "button"; b.addEventListener("click", f); return b; };
const pillSit = s => "fc-pill " + (s === "Atingida" ? "fc-ok" : s === "Não atingida" ? "fc-no" : "fc-imp");

export function montar(raiz, opts) {
  if (!document.getElementById("fc-css")) { const s = el("style"); s.id = "fc-css"; s.textContent = CSS; document.head.appendChild(s); }
  const hoje = new Date(), ant = new Date(hoje.getFullYear(), hoje.getMonth() - 1, 1);
  const S = { pub: normaliza(opts.dados), dados: null, mes: null, proc: 0, sujo: false, msg: opts.msg || "", erro: !!opts.erroMsg, ocupado: false };
  const calc = {}; const calcDe = bu => calc[bu] || (calc[bu] = opts.calcular(bu));
  // completa o nº 3 da semente com o cálculo do painel
  const resolve = D => Object.keys(D.meses).forEach(k => { const m = D.meses[k]; Object.keys(m.proc || {}).forEach(pn => {
    const p = PROCESSOS.find(x => x.nome === pn); if (!p) return;
    Object.keys(m.proc[pn].itens).forEach(n => { const it = m.proc[pn].itens[n]; if (!it.usarCalculo) return;
      const c = calcDe(p.bu)[n]; if (c) { it.resultado = c.resultado; it.porObra = copia(c.porObra); } delete it.usarCalculo; }); }); });
  resolve(S.pub);
  S.dados = opts.rascunho && opts.rascunho() ? normaliza(opts.rascunho()) : copia(S.pub); S.sujo = !!(opts.rascunho && opts.rascunho());
  S.mes = opts.mes && opts.mes() || (S.dados.meses[chaveMes(ant)] ? chaveMes(ant) : Object.keys(S.dados.meses).sort().pop() || chaveMes(ant));
  const guarda = () => { S.sujo = true; if (opts.setRascunho) opts.setRascunho(S.dados); };
  const P = () => PROCESSOS[S.proc];
  const M = () => S.dados.meses[S.mes];
  const fechado = () => !!(M() && M().status === "fechado");
  const pode = () => opts.admin && !fechado();
  if (opts.onDados) opts.onDados(S.pub);

  raiz.textContent = "";
  const wrap = el("div", "fc"); raiz.appendChild(wrap);
  function lancarDoCalculo(manterTextos) {
    const p = P(), c = calcDe(p.bu), m = S.dados.meses[S.mes] = S.dados.meses[S.mes] || { status: "rascunho", proc: {} };
    m.proc = m.proc || {}; const ant2 = (m.proc[p.nome] || {}).itens || {}, itens = {};
    INDICADORES.forEach(def => { const a = ant2[def.n] || {}, cc = c[def.n] || { resultado: null, porObra: [] };
      itens[def.n] = Object.assign({ meta: def.meta, fato: "", causa: "", acao: "" }, manterTextos ? a : {}, { resultado: cc.resultado, porObra: copia(cc.porObra), situacao: manterTextos ? a.situacao || "" : "" }); });
    m.proc[p.nome] = { itens }; guarda(); tudo();
  }
  async function salvar(status) {
    const m = M(); if (!m) return;
    if (status === "fechado") {
      const falta = []; PROCESSOS.forEach(p => { const pp = m.proc && m.proc[p.nome]; if (!pp) return;
        INDICADORES.forEach(def => { const it = pp.itens[def.n]; if (it && situacao(it, def) === "Não atingida" && !(String(it.fato || "").trim() && String(it.causa || "").trim() && String(it.acao || "").trim())) falta.push(p.curto + " nº " + def.n); }); });
      if (falta.length && !window.confirm("Indicadores não atingidos sem Fato, Causa ou Ação: " + falta.join(", ") + ".\n\nFechar o mês mesmo assim?")) return;
    }
    if (status) m.status = status;
    m.atualizado = new Date().toISOString().slice(0, 10); m.por = opts.usuario || "";
    S.ocupado = true; tudo();
    const r = await opts.salvar(copia(S.dados));
    S.ocupado = false; S.msg = r.msg; S.erro = !r.ok; S.pub = copia(S.dados); S.sujo = false; if (opts.setRascunho) opts.setRascunho(null);
    if (opts.onDados) opts.onDados(S.pub); tudo();
  }
  async function exportarMes() {
    try { S.ocupado = true; tudo(); const r = await exportar(S.dados, S.mes, opts.modelo);
      const a = document.createElement("a"); a.href = URL.createObjectURL(r.blob); a.download = r.nome; document.body.appendChild(a); a.click();
      setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000);
      S.msg = "Planilha gerada: " + r.linhas + " indicador(es) e " + r.obras + " linha(s) por empreendimento." + (S.sujo ? " Atenção: inclui alterações ainda não salvas." : ""); S.erro = false; }
    catch (e) { S.msg = "Não foi possível gerar a planilha: " + e.message; S.erro = true; }
    S.ocupado = false; tudo();
  }

  function barra() {
    const b = el("div", "fc-card fc-barra"), e = el("div", "fc-g"), d = el("div", "fc-g");
    const t = el("div"); t.appendChild(el("h2", null, "Fechamento mensal · Projeto Next"));
    t.appendChild(el("p", "fc-aviso", "Resultado oficial enviado ao setor de Processos. Exporta no modelo da planilha mensal."));
    const sel = el("select"), ks = new Set(Object.keys(S.dados.meses)); ks.add(chaveMes(ant)); ks.add(chaveMes(hoje));
    [...ks].sort().reverse().forEach(k => { const o = el("option", null, rotulo(k) + (S.dados.meses[k] ? (S.dados.meses[k].status === "fechado" ? " · fechado" : " · rascunho") : "")); o.value = k; o.selected = k === S.mes; sel.appendChild(o); });
    sel.addEventListener("change", () => { S.mes = sel.value; if (opts.setMes) opts.setMes(S.mes); S.msg = ""; tudo(); });
    const m = M(), st = !m ? ["Sem lançamento", "fc-imp"] : m.status === "fechado" ? ["Fechado", "fc-fech"] : ["Rascunho", "fc-rasc"];
    e.append(t); d.append(sel, el("span", "fc-pill " + st[1], st[0] + (S.sujo ? " · não salvo" : "")));
    PROCESSOS.forEach((p, i) => d.appendChild(bt(p.curto, () => { S.proc = i; tudo(); }, i === S.proc ? "on" : "")));
    if (opts.admin) {
      if (fechado()) d.appendChild(bt("Reabrir mês", () => { if (!window.confirm("Reabrir " + rotulo(S.mes) + " para edição?")) return; M().status = "rascunho"; guarda(); tudo(); }));
      else if (m) { const s1 = bt("Salvar", () => salvar(), ""); s1.disabled = S.ocupado || !S.sujo; const s2 = bt("Fechar mês", () => salvar("fechado"), "pri"); s2.disabled = S.ocupado; d.append(s1, s2); }
    }
    const ex = bt("Exportar planilha (.xlsx)", exportarMes, opts.admin ? "" : "pri"); ex.disabled = S.ocupado || !m; d.appendChild(ex);
    b.append(e, d); return b;
  }
  function campoNum(rot, val, set) { const l = el("label", "fc-campo", rot); const i = el("input"); i.inputMode = "decimal"; i.value = val == null ? "" : String(val).replace(".", ",");
    i.addEventListener("input", () => { set(num(i.value)); guarda(); }); i.addEventListener("change", () => tudo()); l.appendChild(i); return l; }
  function cartao(def) {
    const p = P(), m = M(), it = m && m.proc && m.proc[p.nome] && m.proc[p.nome].itens[def.n], c = calcDe(p.bu)[def.n] || {};
    const card = el("div", "fc-card"), top = el("div", "fc-barra");
    const tt = el("div"); tt.appendChild(el("h3", null, "Nº " + def.n + " · " + ((it && it.nome) || def.nome))); tt.appendChild(el("p", "fc-aviso", def.sentido + " · meta " + fmt(it && it.meta != null ? it.meta : def.meta)));
    const sit = it ? situacao(it, def) : "Em implementação";
    top.append(tt, el("span", pillSit(sit), sit)); card.appendChild(top);
    const g = el("div", "fc-grade"), ed = pode();
    if (ed) {
      g.appendChild(campoNum("Meta", it.meta != null ? it.meta : def.meta, v => { it.meta = v; }));
      g.appendChild(campoNum("Resultado", it.resultado, v => { it.resultado = v; }));
      const ls = el("label", "fc-campo", "Situação"), s = el("select");
      [["", "Automática (" + situacao(Object.assign({}, it, { situacao: "" }), def) + ")"], ["Atingida", "Atingida"], ["Não atingida", "Não atingida"], ["Em implementação", "Em implementação"]]
        .forEach(o => { const op = el("option", null, o[1]); op.value = o[0]; op.selected = (it.situacao || "") === o[0]; s.appendChild(op); });
      s.addEventListener("change", () => { it.situacao = s.value; guarda(); tudo(); }); ls.appendChild(s); g.appendChild(ls);
    } else {
      [["Meta", fmt(it && it.meta != null ? it.meta : def.meta)], ["Resultado", it ? (it.resultado == null ? "A implantar" : fmt(it.resultado)) : "—"], ["Situação", sit]]
        .forEach(x => { const l = el("div", "fc-campo", x[0]); l.appendChild(el("span", "fc-val", x[1])); g.appendChild(l); });
    }
    const lc = el("div", "fc-campo", "Cálculo do painel hoje"); const vc = el("span", "fc-val", fmt(c.resultado)); lc.appendChild(vc);
    if (ed && c.resultado != null && num(it.resultado) !== c.resultado) lc.appendChild(bt("Usar este valor", () => { it.resultado = c.resultado; it.porObra = copia(c.porObra || []); guarda(); tudo(); }));
    else lc.appendChild(el("span", "fc-calc", c.porObra && c.porObra.length ? c.porObra.length + " empreendimento(s) na base atual" : "sem dados na base atual"));
    g.appendChild(lc); card.appendChild(g);
    if (!it) { card.appendChild(el("p", "fc-aviso", "Sem lançamento neste mês.")); return card; }
    const tx = el("div", "fc-txt");
    [["fato", "Fato"], ["causa", "Causa"], ["acao", "Ação"]].forEach(k => { const l = el("label", "fc-campo", k[1]);
      if (ed) { const a = el("textarea"); a.rows = 3; a.value = it[k[0]] || ""; a.addEventListener("input", () => { it[k[0]] = a.value; guarda(); }); l.appendChild(a); }
      else l.appendChild(el("p", null, it[k[0]] || "-"));
      tx.appendChild(l); });
    card.appendChild(tx);
    if (ed) { const l = el("label", "fc-campo", "Cálculo (texto da planilha)"); const a = el("textarea"); a.rows = 2; a.value = it.calculo || def.calculo;
      a.addEventListener("input", () => { it.calculo = a.value; guarda(); }); l.appendChild(a); l.style.marginTop = "10px"; card.appendChild(l); }
    const ob = el("div", "fc-obras"); ob.appendChild(el("span", "fc-campo", "Por empreendimento"));
    (it.porObra || []).forEach((o, i) => { const ch = el("span", "fc-obra");
      if (ed) { const n = el("input", "nome"); n.value = o.e; n.addEventListener("input", () => { o.e = n.value; guarda(); });
        const v = el("input"); v.value = o.v == null ? "" : String(o.v).replace(".", ","); v.addEventListener("input", () => { o.v = num(v.value); guarda(); });
        ch.append(n, v, bt("×", () => { it.porObra.splice(i, 1); guarda(); tudo(); }, "fc-x")); }
      else { ch.append(o.e + " "); ch.appendChild(el("b", null, fmt(o.v))); }
      ob.appendChild(ch); });
    if (!(it.porObra || []).length) ob.appendChild(el("span", "fc-aviso", "nenhum"));
    if (ed) { ob.appendChild(bt("+ empreendimento", () => { (it.porObra = it.porObra || []).push({ e: "", v: null }); guarda(); tudo(); }));
      const vs = (it.porObra || []).map(o => num(o.v)).filter(v => v != null);
      if (vs.length) ob.appendChild(bt("Resultado = média (" + fmt(vs.reduce((a, b) => a + b, 0) / vs.length) + ")", () => { it.resultado = Math.round(vs.reduce((a, b) => a + b, 0) / vs.length * 100) / 100; guarda(); tudo(); })); }
    card.appendChild(ob);
    return card;
  }
  function tudo() {
    const foco = document.activeElement && raiz.contains(document.activeElement) ? null : null; void foco;
    wrap.textContent = ""; wrap.appendChild(barra());
    if (S.msg) { const p = el("p", "fc-msg" + (S.erro ? " err" : ""), S.msg); wrap.appendChild(p); }
    const m = M(), p = P();
    if (!m || !m.proc || !m.proc[p.nome]) {
      const c = el("div", "fc-card"); c.appendChild(el("h3", null, "Sem lançamento de " + p.curto + " em " + rotulo(S.mes)));
      c.appendChild(el("p", "fc-aviso", opts.admin ? "Comece pelo cálculo atual do painel e revise antes de fechar." : "O administrador ainda não lançou este mês."));
      if (opts.admin && !fechado()) c.appendChild(bt("Iniciar com o cálculo do painel", () => lancarDoCalculo(false), "pri"));
      wrap.appendChild(c);
    } else {
      if (pode()) { const a = el("div", "fc-g"); a.appendChild(bt("Recalcular com a base atual", () => {
          if (!window.confirm("Substituir Resultado e Por empreendimento de " + p.curto + " pelo cálculo atual? Fato, Causa e Ação são mantidos.")) return; lancarDoCalculo(true); }));
        a.appendChild(el("span", "fc-aviso", "A base não guarda histórico por mês: o cálculo é sempre o da data de hoje.")); wrap.appendChild(a); }
      INDICADORES.forEach(def => wrap.appendChild(cartao(def)));
      if (m.atualizado) wrap.appendChild(el("p", "fc-aviso", "Última gravação em " + m.atualizado.split("-").reverse().join("/") + (m.por ? " · " + m.por : "")));
    }
  }
  tudo();
}
