// Cadastro de Obras: fonte única das obras do painel (Supabase dash_obras, uma linha por obra).
// Qualquer usuário cadastra e edita. Só o admin: habilitar/desabilitar, excluir, alterar a sigla depois de criada, VGV e budget.
// Módulo isolado: erro aqui não derruba as outras abas. Quem grava é o painel (opts.api).

export const BUS = [["Make It", "Make It"], ["Gadens Inc", "Gadens"]];
export const FASES = ["Viabilidade", "Estudo Preliminar", "Anteprojeto", "Pré-executivo", "Executivo", "Obra", "Entregue"];
export const PADROES = ["Econômico", "Médio", "Médio-alto", "Alto", "Luxo"];
// Campos do formulário: [chave, rótulo, tipo, largura (1 a 4), opções/placeholder, só admin]
export const SECOES = [
  { t: "Identificação", d: "A sigla é a chave que liga a obra ao ClickUp, TAP, fornecedores e financeiro.", f: [
    ["sigla", "Sigla *", "sigla", 1, "Ex.: HOR"], ["nome", "Nome comercial *", "text", 2, "Ex.: Horizon"], ["bu", "Unidade de negócio *", "select", 1, BUS],
    ["apelidos", "Outros nomes (como aparece no ClickUp, SAÍDAS, TAP)", "lista", 4, "Separe por ponto e vírgula. Ex.: Leilão; QUA"],
    ["padrao", "Padrão", "select", 1, [["", "—"]].concat(PADROES.map(x => [x, x]))], ["fase", "Fase atual", "select", 1, [["", "—"]].concat(FASES.map(x => [x, x]))],
    ["lancamento", "Data de lançamento", "date", 1], ["inicioObra", "Início de obra previsto", "date", 1]] },
  { t: "Localização", d: "O endereço aparece no Geral com link para o mapa.", f: [
    ["endereco", "Endereço (rua e número)", "text", 2], ["bairro", "Bairro", "text", 1], ["cep", "CEP", "text", 1],
    ["cidade", "Cidade", "text", 2], ["uf", "UF", "text", 1, "PR"], ["coord", "Coordenadas (opcional)", "text", 1, "-25.43, -49.27"]] },
  { t: "Jurídico", d: "", f: [
    ["spe", "SPE / razão social", "text", 2], ["cnpj", "CNPJ da obra", "text", 1], ["gds", "GDS", "text", 1], ["matricula", "Matrícula(s)", "text", 2]] },
  { t: "Produto", d: "", f: [
    ["uh", "Unidades (UH)", "num", 1], ["torres", "Torres / blocos", "num", 1], ["pavimentos", "Pavimentos", "num", 1], ["subsolos", "Subsolos", "num", 1],
    ["vagas", "Vagas", "num", 1], ["aTerreno", "Área do terreno (m²)", "num", 1], ["aConst", "Área construída (m²)", "num", 1], ["aPriv", "Área privativa (m²)", "num", 1]] },
  { t: "Financeiro", d: "Só o administrador altera. Alimenta o Geral, o Financeiro e o controle de budget do Lançamento Financeiro.", f: [
    ["vgv", "VGV (R$)", "num", 2, "", true], ["budget", "Budget de projetos complementares (R$)", "num", 2, "", true]] },
  { t: "Responsáveis", d: "", f: [
    ["gerente", "Gerente do projeto", "text", 2], ["coordenador", "Coordenador de projetos", "text", 2],
    ["analistas", "Analista(s) responsável(is)", "lista", 2, "Separe por ponto e vírgula"], ["patrocinador", "Patrocinador / diretoria", "text", 2]] },
  { t: "Integrações", d: "O link do ClickUp define de onde o painel lê prazos, linha axial e tarefas da obra (próxima atualização automática).", f: [
    ["clickup", "Link do ClickUp (lista ou pasta do cronograma da obra) *", "url", 4, "https://app.clickup.com/…/v/li/…"],
    ["sharepoint", "Pasta de projetos (SharePoint)", "url", 2, "https://…"], ["construflow", "Construflow", "url", 2, "https://…"],
    ["autodoc", "Autodoc", "url", 2, "https://…"], ["outros", "Outro link", "url", 2, "https://…"], ["obs", "Observações", "area", 4]] }
];

export function viewDoLink(link) {
  const s = String(link || "").trim(); if (!s) return "";
  if (/^(li|f):\d+$/.test(s)) return s;
  const li = s.match(/\/v\/li\/(\d+)/), fo = s.match(/\/v\/f\/(\d+)/), m = s.match(/(8[a-z0-9]{5,}-\d{3,})/i);
  return li ? "li:" + li[1] : fo ? "f:" + fo[1] : m ? m[1] : "";
}
export const norm = s => String(s || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
const num = v => { if (v == null || v === "") return null; if (typeof v === "number") return v; const t = String(v).trim().replace(/\s/g, "").replace(/R\$/i, "");
  const x = /,\d{1,2}$/.test(t) ? t.replace(/\./g, "").replace(",", ".") : t.replace(/,/g, ""); const n = parseFloat(x); return isNaN(n) ? null : n; };
const moeda = v => v == null ? "—" : v >= 1e6 ? "R$ " + (v / 1e6).toFixed(1).replace(".", ",") + " mi" : v.toLocaleString("pt-BR", { style: "currency", currency: "BRL", maximumFractionDigits: 0 });
const hoje = () => new Date().toISOString().slice(0, 10);
const copia = o => JSON.parse(JSON.stringify(o));
const linkOk = s => /^https?:\/\/\S+$/i.test(String(s || "").trim());
export function nova() { return { id: "ob" + Date.now().toString(36) + Math.random().toString(36).slice(2, 6), sigla: "", nome: "", bu: "Make It", apelidos: [], ativa: true, analistas: [] }; }

const CSS = `
.ob{display:flex;flex-direction:column;gap:14px;color:var(--color-text,#0F172A)}
.ob *{box-sizing:border-box}
.ob-card{background:var(--card,#fff);border:1px solid var(--slate-100,#F1F5F9);border-radius:16px;box-shadow:var(--card-shadow);padding:16px 18px}
.ob-l{display:flex;align-items:center;gap:8px;flex-wrap:wrap}.ob-sb{justify-content:space-between}
.ob button{font:inherit;font-size:12.5px;font-weight:700;cursor:pointer;border:1px solid var(--slate-200,#E2E8F0);background:var(--card,#fff);color:var(--slate-600,#475569);border-radius:10px;padding:7px 12px}
.ob button:hover{border-color:var(--blue-500,#3B82F6)}.ob button.pri{background:var(--blue-600,#2563EB);border-color:var(--blue-600,#2563EB);color:#fff}
.ob button.per{color:#B91C1C}.ob button.on{background:var(--blue-50,#EFF6FF);border-color:var(--blue-500,#3B82F6);color:var(--blue-700,#1D4ED8)}
.ob button:disabled{opacity:.45;cursor:default}.ob button.mini{padding:4px 9px;font-size:11.5px;border-radius:8px}
.ob input,.ob select,.ob textarea{font:inherit;font-size:13px;width:100%;padding:8px 10px;border-radius:10px;border:1px solid var(--slate-200,#E2E8F0);background:var(--card,#fff);color:var(--color-text,#0F172A)}
.ob input:disabled,.ob select:disabled{background:var(--slate-50,#F8FAFC);color:var(--slate-500,#64748B)}
.ob-kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px}
.ob-kpi{background:var(--card,#fff);border:1px solid var(--slate-100,#F1F5F9);border-left:4px solid var(--cor);border-radius:14px;padding:12px 14px;display:flex;flex-direction:column;gap:2px}
.ob-kpi span{font-size:10px;letter-spacing:.05em;text-transform:uppercase;font-weight:800;color:var(--slate-400,#94A3B8)}.ob-kpi b{font-size:22px;font-family:var(--font-heading)}
.ob table{width:100%;border-collapse:collapse;font-size:12.5px}
.ob th{text-align:left;font-size:10px;letter-spacing:.06em;text-transform:uppercase;color:var(--slate-400,#94A3B8);font-weight:800;padding:8px;border-bottom:1px solid var(--slate-200,#E2E8F0);position:sticky;top:0;background:var(--card,#fff)}
.ob td{padding:8px;border-bottom:1px solid var(--slate-100,#F1F5F9);vertical-align:middle}.ob tr.inativa td{opacity:.5}
.ob .n{text-align:right;font-family:var(--font-mono);white-space:nowrap}.ob .sg{font-family:var(--font-mono);font-weight:800;color:var(--blue-600,#2563EB)}
.ob-pill{display:inline-block;font-size:10.5px;font-weight:800;padding:2px 8px;border-radius:99px;white-space:nowrap}
.ob-ok{background:#D1FAE5;color:#065F46}.ob-no{background:#FEE2E2;color:#991B1B}.ob-am{background:#FEF3C7;color:#92400E}.ob-ci{background:var(--slate-100,#F1F5F9);color:var(--slate-500,#64748B)}
.ob-sw{position:relative;width:38px;height:22px;border-radius:99px;border:0!important;padding:0!important;background:var(--slate-300,#CBD5E1)!important}
.ob-sw.on{background:#10B981!important}.ob-sw i{position:absolute;top:3px;left:3px;width:16px;height:16px;border-radius:99px;background:#fff;transition:left .15s}.ob-sw.on i{left:19px}
.ob-bg{position:fixed;inset:0;background:rgba(15,23,42,.42);z-index:9000;display:flex;justify-content:flex-end}
.ob-dw{width:min(760px,100vw);height:100%;background:var(--card,#fff);display:flex;flex-direction:column;box-shadow:-12px 0 40px rgba(15,23,42,.2)}
.ob-dw.entra{animation:obIn .18s ease-out}@keyframes obIn{from{transform:translateX(40px);opacity:.4}to{transform:none;opacity:1}}
.ob-dw-cab{padding:16px 20px;border-bottom:1px solid var(--slate-100,#F1F5F9)}.ob-dw-corpo{flex:1;overflow-y:auto;padding:16px 20px;display:flex;flex-direction:column;gap:18px}
.ob-dw-rod{padding:12px 20px;border-top:1px solid var(--slate-100,#F1F5F9);display:flex;justify-content:space-between;gap:8px}
.ob-sec h4{margin:0 0 2px;font-size:12px;letter-spacing:.05em;text-transform:uppercase;font-weight:800;color:var(--slate-500,#64748B)}
.ob-grade{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:10px;margin-top:8px}
.ob-campo{display:flex;flex-direction:column;gap:4px;font-size:11.5px;font-weight:700;color:var(--slate-500,#64748B)}
.ob-aviso{margin:0;font-size:11.5px;line-height:1.4;color:var(--slate-500,#64748B)}.ob-aviso.err{color:#B91C1C}.ob-aviso.ok{color:#047857}
.ob-msg{margin:0;font-size:12.5px;padding:9px 12px;border-radius:10px;background:#ECFDF5;color:#065F46}.ob-msg.err{background:#FEF2F2;color:#991B1B}
.ob-vinc{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:10px}
.ob-vinc>div{border:1px solid var(--slate-100,#F1F5F9);border-radius:12px;padding:10px 12px;background:var(--slate-50,#F8FAFC);display:flex;flex-direction:column;gap:3px;font-size:12px}
.ob-vinc b{font-size:11px;letter-spacing:.04em;text-transform:uppercase;color:var(--slate-500,#64748B)}.ob-vinc a{color:var(--blue-600,#2563EB);font-weight:700}
@media(max-width:700px){.ob-grade{grid-template-columns:repeat(2,minmax(0,1fr))}}
`;
const el = (t, c, x) => { const e = document.createElement(t); if (c) e.className = c; if (x != null) e.textContent = x; return e; };
const bt = (x, f, c) => { const b = el("button", c, x); b.type = "button"; b.addEventListener("click", f); return b; };

export function montar(raiz, opts) {
  if (!document.getElementById("ob-css")) { const s = el("style"); s.id = "ob-css"; s.textContent = CSS; document.head.appendChild(s); }
  const adm = !!opts.admin, api = opts.api;
  const S = { lista: copia(opts.obras || []), f: { bu: "", st: "ativas", q: "" }, form: null, msg: opts.msg || "", erro: !!opts.erroMsg, ocupado: false, dwId: null };
  const wrap = el("div", "ob"); raiz.textContent = ""; raiz.appendChild(wrap);
  const ok = m => { S.msg = m; S.erro = false; }, falha = m => { S.msg = m; S.erro = true; };
  const vinc = o => { try { return (opts.vinculos && opts.vinculos(o)) || {}; } catch (e) { return {}; } };

  function topo() {
    const c = el("div", "ob-card");
    const l = el("div", "ob-l ob-sb");
    const fl = el("div", "ob-l");
    [["", "Todas"]].concat(BUS).forEach(b => fl.appendChild(bt(b[1], () => { S.f.bu = b[0]; tudo(); }, S.f.bu === b[0] ? "on" : "")));
    const st = el("select"); [["ativas", "Ativas"], ["inativas", "Desabilitadas"], ["todas", "Todas"]].forEach(o => { const op = el("option", null, o[1]); op.value = o[0]; op.selected = S.f.st === o[0]; st.appendChild(op); });
    st.style.width = "150px"; st.addEventListener("change", () => { S.f.st = st.value; tudo(); }); fl.appendChild(st);
    const q = el("input"); q.placeholder = "Buscar sigla, nome, cidade"; q.value = S.f.q; q.style.width = "220px";
    q.addEventListener("input", () => { S.f.q = q.value; desenhaTabela(); }); fl.appendChild(q);
    const dir = el("div", "ob-l");
    if (adm && opts.importar) dir.appendChild(bt(S.lista.length ? "Trazer obras que faltam" : "Pré-preencher com as obras atuais", importar));
    dir.appendChild(bt("+ Nova obra", () => { S.form = nova(); S.form.bu = S.f.bu || "Make It"; tudo(); }, "pri"));
    l.append(fl, dir); c.appendChild(l); return c;
  }
  function kpis() {
    const k = el("div", "ob-kpis"), L = S.lista.filter(o => !S.f.bu || o.bu === S.f.bu);
    const at = L.filter(o => o.ativa !== false);
    [["Obras cadastradas", L.length, "#2563EB"], ["Ativas no painel", at.length, "#10B981"], ["Desabilitadas", L.length - at.length, "#94A3B8"],
     ["Ativas sem link do ClickUp", at.filter(o => !viewDoLink(o.clickup)).length, "#EF4444"], ["Ativas sem endereço", at.filter(o => !o.endereco).length, "#F59E0B"]]
      .forEach(x => { const d = el("div", "ob-kpi"); d.style.setProperty("--cor", x[2]); d.append(el("span", null, x[0]), el("b", null, String(x[1]))); k.appendChild(d); });
    return k;
  }
  let corpoTab = null;
  function tabela() {
    const c = el("div", "ob-card"); const w = el("div"); w.style.cssText = "overflow:auto;max-height:640px"; c.appendChild(w); corpoTab = w; desenhaTabela(); return c;
  }
  function desenhaTabela() {
    if (!corpoTab) return; corpoTab.textContent = "";
    const q = norm(S.f.q);
    const L = S.lista.filter(o => (!S.f.bu || o.bu === S.f.bu) && (S.f.st === "todas" || (S.f.st === "ativas" ? o.ativa !== false : o.ativa === false)))
      .filter(o => !q || norm([o.sigla, o.nome, o.cidade, (o.apelidos || []).join(" ")].join(" ")).indexOf(q) >= 0)
      .sort((a, b) => String(a.bu).localeCompare(String(b.bu)) || String(a.nome).localeCompare(String(b.nome)));
    if (!L.length) { corpoTab.appendChild(el("p", "ob-aviso", S.lista.length ? "Nenhuma obra neste filtro." : "Nenhuma obra cadastrada ainda." + (adm ? " Use \"Pré-preencher com as obras atuais\" para começar com o que o painel já tem." : ""))); return; }
    const t = el("table"), h = el("tr");
    ["Sigla", "Obra", "Unidade", "Fase", "Cidade", "UH", "VGV", "ClickUp", "Vínculos", "No painel", ""].forEach((x, i) => { const th = el("th", null, x); if (i === 5 || i === 6) th.style.textAlign = "right"; h.appendChild(th); });
    t.appendChild(h);
    L.forEach(o => {
      const r = el("tr", o.ativa === false ? "inativa" : ""), v = vinc(o);
      const td = (x, c) => { const d = el("td", c, x == null || x === "" ? "—" : String(x)); r.appendChild(d); return d; };
      td(o.sigla, "sg"); const tn = td(o.nome); tn.style.fontWeight = "700";
      td(o.bu === "Gadens Inc" ? "Gadens" : o.bu); td(o.fase); td(o.cidade); td(o.uh ? Number(o.uh).toLocaleString("pt-BR") : "", "n"); td(o.vgv ? moeda(num(o.vgv)) : "", "n");
      const tc = el("td"); tc.appendChild(el("span", "ob-pill " + (viewDoLink(o.clickup) ? "ob-ok" : "ob-no"), viewDoLink(o.clickup) ? "vinculado" : "sem link")); r.appendChild(tc);
      const tv = el("td"); tv.style.cssText = "font-size:11px;color:var(--slate-500,#64748B);white-space:nowrap";
      tv.textContent = [v.tarefas != null ? v.tarefas + " tarefas" : "", v.taps ? "TAP Rev. " + String(v.taps[v.taps.length - 1].rev).padStart(2, "0") : "", v.contratos ? v.contratos + " contratos" : "", v.riscos ? v.riscos + " riscos" : ""].filter(Boolean).join(" · ") || "—";
      r.appendChild(tv);
      const ta = el("td");
      if (adm) { const sw = bt("", async () => { const nv = Object.assign({}, o, { ativa: o.ativa === false }); await grava(nv, nv.ativa ? o.nome + " habilitada no painel." : o.nome + " desabilitada: deixa de aparecer no painel."); }, "ob-sw" + (o.ativa !== false ? " on" : ""));
        sw.title = o.ativa !== false ? "Desabilitar" : "Habilitar"; sw.appendChild(el("i")); ta.appendChild(sw); }
      else ta.appendChild(el("span", "ob-pill " + (o.ativa !== false ? "ob-ok" : "ob-ci"), o.ativa !== false ? "ativa" : "desabilitada"));
      r.appendChild(ta);
      const tb = el("td"); tb.style.whiteSpace = "nowrap"; tb.appendChild(bt("Editar", () => { S.form = copia(o); tudo(); }, "mini"));
      if (adm) tb.appendChild(bt("Excluir", () => excluir(o), "mini per"));
      r.appendChild(tb); t.appendChild(r);
    });
    corpoTab.appendChild(t);
  }
  async function grava(o, txt) {
    S.ocupado = true; o.alteradoEm = new Date().toISOString(); o.alteradoPor = opts.usuario || ""; if (!o.criadoEm) { o.criadoEm = o.alteradoEm; o.criadoPor = o.alteradoPor; }
    const r = await api.salvar(copia(o)); S.ocupado = false;
    if (r.ok) { const i = S.lista.findIndex(x => x.id === o.id); if (i >= 0) S.lista[i] = o; else S.lista.push(o); ok(txt || r.msg); } else falha(r.msg);
    tudo(); return r;
  }
  async function excluir(o) {
    const v = vinc(o);
    const aviso = [v.tarefas ? v.tarefas + " tarefas do ClickUp" : "", v.taps ? v.taps.length + " revisão(ões) de TAP" : "", v.contratos ? v.contratos + " contrato(s) no Lançamento Financeiro" : ""].filter(Boolean);
    if (!window.confirm("Excluir o cadastro de " + o.nome + " (" + o.sigla + ")?\n\n" + (aviso.length ? "Ligados a esta obra: " + aviso.join(", ") + ". Eles continuam guardados, mas a obra some do painel.\nSe a obra só não deve aparecer, prefira desabilitar.\n\n" : "") + "Essa ação não pode ser desfeita.")) return;
    const r = await api.excluir(o.id); if (r.ok) { S.lista = S.lista.filter(x => x.id !== o.id); ok(o.nome + " excluída do cadastro."); } else falha(r.msg); tudo();
  }
  async function importar() {
    const novas = (opts.importar() || []).filter(n => !S.lista.some(o => norm(o.sigla) === norm(n.sigla)));
    if (!novas.length) { ok("Todas as obras do painel já estão no cadastro."); tudo(); return; }
    if (!window.confirm("Trazer " + novas.length + " obra(s) com os dados que o painel já tem (ClickUp, Lançamento Financeiro e TAP)?\n\n" + novas.map(n => n.sigla + " · " + n.nome + (n.ativa === false ? " (desabilitada)" : "")).join("\n"))) return;
    S.ocupado = true; tudo(); let n = 0, erro = "";
    for (const o of novas) { o.criadoEm = new Date().toISOString(); o.criadoPor = (opts.usuario || "") + " (importação)"; const r = await api.salvar(copia(o)); if (r.ok) { S.lista.push(o); n++; } else { erro = r.msg; break; } }
    S.ocupado = false; erro ? falha(n + " obra(s) importada(s); parou com erro: " + erro) : ok(n + " obra(s) trazida(s) para o cadastro. Confira e complete o que faltar."); tudo();
  }

  // ── formulário ──
  function campo(o, c) {
    const [k, rot, tipo, larg, op, soAdm] = c, bloqueado = (soAdm && !adm) || (k === "sigla" && !adm && S.lista.some(x => x.id === o.id));
    const l = el("label", "ob-campo", rot); l.style.gridColumn = "span " + larg;
    let i;
    if (tipo === "select") { i = el("select"); op.forEach(x => { const e = el("option", null, x[1]); e.value = x[0]; e.selected = String(o[k] || "") === String(x[0]); i.appendChild(e); }); i.addEventListener("change", () => { o[k] = i.value; }); }
    else if (tipo === "area") { i = el("textarea"); i.rows = 3; i.value = o[k] || ""; i.addEventListener("input", () => { o[k] = i.value; }); }
    else { i = el("input"); if (tipo === "date") i.type = "date"; if (tipo === "num") i.inputMode = "decimal"; if (op && typeof op === "string") i.placeholder = op;
      i.value = tipo === "lista" ? (o[k] || []).join("; ") : o[k] == null ? "" : String(o[k]).replace(tipo === "num" ? "." : "\u0000", tipo === "num" ? "," : "\u0000");
      i.addEventListener("input", () => { o[k] = tipo === "lista" ? i.value.split(";").map(x => x.trim()).filter(Boolean) : tipo === "num" ? num(i.value) : tipo === "sigla" ? i.value.toUpperCase().replace(/[^A-Z0-9]/g, "").slice(0, 4) : i.value;
        if (tipo === "sigla") i.value = o[k]; if (tipo === "url") dica(); }); }
    i.disabled = bloqueado; l.appendChild(i);
    if (k === "sigla" && bloqueado) l.appendChild(el("span", "ob-aviso", "Só o administrador altera a sigla depois de criada."));
    if (soAdm && !adm) l.appendChild(el("span", "ob-aviso", "Só o administrador altera."));
    let dica = () => {};
    if (k === "clickup") { const a = el("span", "ob-aviso"); l.appendChild(a);
      dica = () => { const v = viewDoLink(o.clickup); a.className = "ob-aviso " + (o.clickup ? (v ? "ok" : "err") : ""); a.textContent = !o.clickup ? "Abra a lista do cronograma da obra no ClickUp e copie o endereço do navegador." : v ? "Reconhecido: " + (v.startsWith("li:") ? "lista " : v.startsWith("f:") ? "pasta " : "visualização ") + v.replace(/^(li|f):/, "") + ". Entra na próxima atualização automática." : "Link não reconhecido. Use o link de uma lista (…/v/li/…) ou pasta (…/v/f/…) do ClickUp."; }; dica(); }
    else if (tipo === "url") { const a = el("a", null, "abrir"); dica = () => { a.style.display = linkOk(o[k]) ? "" : "none"; a.href = o[k] || "#"; }; a.target = "_blank"; a.rel = "noopener"; a.style.cssText += ";font-size:11px;color:var(--blue-600)"; l.appendChild(a); dica(); }
    if (k === "endereco") { const a = el("a", null, "ver no mapa"); a.target = "_blank"; a.rel = "noopener"; a.style.cssText = "font-size:11px;color:var(--blue-600)";
      const at = () => { const q = [o.endereco, o.bairro, o.cidade, o.uf].filter(Boolean).join(", "); a.href = "https://www.google.com/maps/search/?api=1&query=" + encodeURIComponent(o.coord || q); a.style.display = o.endereco ? "" : "none"; };
      i.addEventListener("input", at); at(); l.appendChild(a); }
    return l;
  }
  function telaForm() {
    const o = S.form, novo = !S.lista.some(x => x.id === o.id);
    const bg = el("div", "ob-bg"), dw = el("div", "ob-dw ob" + (S.dwId !== o.id ? " entra" : "")); S.dwId = o.id;
    const fechar = () => { S.form = null; S.dwId = null; tudo(); };
    bg.addEventListener("click", ev => { if (ev.target === bg) fechar(); });
    const cab = el("div", "ob-dw-cab"), l = el("div", "ob-l ob-sb");
    l.append(el("h3", null, novo ? "Nova obra" : o.sigla + " · " + o.nome), bt("✕", fechar));
    cab.appendChild(l);
    if (!novo) cab.appendChild(el("p", "ob-aviso", "Criada por " + (o.criadoPor || "—") + (o.criadoEm ? " em " + o.criadoEm.slice(0, 10).split("-").reverse().join("/") : "") + (o.alteradoEm ? " · última alteração por " + (o.alteradoPor || "—") + " em " + o.alteradoEm.slice(0, 10).split("-").reverse().join("/") : "")));
    dw.appendChild(cab);
    const corpo = el("div", "ob-dw-corpo"); dw.appendChild(corpo);
    if (S.msg && S.erro) corpo.appendChild(el("p", "ob-msg err", S.msg));
    if (!novo) corpo.appendChild(vinculos(o));
    SECOES.forEach(sc => { const s = el("div", "ob-sec"); s.appendChild(el("h4", null, sc.t)); if (sc.d) s.appendChild(el("p", "ob-aviso", sc.d));
      const g = el("div", "ob-grade"); sc.f.forEach(c => g.appendChild(campo(o, c))); s.appendChild(g); corpo.appendChild(s); });
    const rod = el("div", "ob-dw-rod");
    const esq = el("div", "ob-l");
    if (adm) { const sw = bt("", () => { o.ativa = o.ativa === false; tudo(); }, "ob-sw" + (o.ativa !== false ? " on" : "")); sw.appendChild(el("i")); esq.append(sw, el("span", "ob-aviso", o.ativa !== false ? "Aparece no painel" : "Desabilitada: não aparece no painel")); }
    rod.appendChild(esq);
    const sv = bt(S.ocupado ? "Salvando…" : "Salvar", async () => {
      o.sigla = String(o.sigla || "").toUpperCase().trim(); o.nome = String(o.nome || "").trim();
      const faltas = []; if (!o.sigla || o.sigla.length < 2) faltas.push("sigla"); if (!o.nome) faltas.push("nome"); if (!o.bu) faltas.push("unidade de negócio");
      if (faltas.length) { falha("Preencha: " + faltas.join(", ") + "."); tudo(); return; }
      const dup = S.lista.find(x => x.id !== o.id && norm(x.sigla) === norm(o.sigla)); if (dup) { falha("A sigla " + o.sigla + " já é da obra " + dup.nome + "."); tudo(); return; }
      if (o.clickup && !viewDoLink(o.clickup)) { falha("O link do ClickUp não foi reconhecido. Corrija ou deixe em branco."); tudo(); return; }
      const antes = S.lista.find(x => x.id === o.id);
      if (antes && antes.sigla !== o.sigla && !window.confirm("Trocar a sigla de " + antes.sigla + " para " + o.sigla + "?\n\nTAPs, vínculos de fornecedores e lançamentos ficam ligados à sigla antiga até serem ajustados.")) return;
      const r = await grava(o, (novo ? "Obra " : "Cadastro de ") + o.sigla + " · " + o.nome + (novo ? " criada." : " atualizado.") + " Já vale para todos.");
      if (r.ok) { S.form = null; S.dwId = null; tudo(); }
    }, "pri"); sv.disabled = S.ocupado;
    const dir = el("div", "ob-l"); dir.append(bt("Cancelar", fechar), sv); rod.appendChild(dir); dw.appendChild(rod);
    bg.appendChild(dw); return bg;
  }
  function vinculos(o) {
    const v = vinc(o), s = el("div", "ob-sec"); s.appendChild(el("h4", null, "Onde esta obra aparece"));
    const g = el("div", "ob-vinc"); s.appendChild(g);
    const caixa = (t, linhas) => { const d = el("div"); d.appendChild(el("b", null, t)); (linhas.length ? linhas : ["—"]).forEach(x => d.appendChild(typeof x === "string" ? el("span", null, x) : x)); g.appendChild(d); };
    const ln = (txt, href) => { const a = el("a", null, txt); a.href = href; a.target = "_blank"; a.rel = "noopener"; return a; };
    caixa("Prazos (ClickUp)", v.tarefas != null ? [v.tarefas + " tarefas complementares · " + (v.pct != null ? v.pct + "% concluído" : ""), v.fase ? "Fase: " + v.fase : ""].filter(Boolean).concat(viewDoLink(o.clickup) ? [ln("Abrir no ClickUp", o.clickup)] : []) : viewDoLink(o.clickup) ? ["Link cadastrado: entra na próxima atualização", ln("Abrir no ClickUp", o.clickup)] : ["Sem link do ClickUp"]);
    caixa("TAP", v.taps && v.taps.length ? v.taps.map(t => "TAP." + t.sigla + ".001 · Rev. " + String(t.rev).padStart(2, "0") + (t.data ? " · " + t.data : "")).concat([ln("Abrir o TAP", "tap.html")]) : ["Nenhum TAP emitido com a sigla " + (o.sigla || "—")]);
    caixa("Lançamento Financeiro", v.contratos != null ? [v.contratos + " contrato(s)", "Utilizado " + moeda(v.utilizado || 0) + (o.budget ? " de " + moeda(num(o.budget)) : "")] : ["Sem contratos"]);
    caixa("Fornecedores vinculados", (v.fornecedores || []).slice(0, 6).map(f => f.disc + " · " + f.nome).concat((v.fornecedores || []).length > 6 ? ["+ " + (v.fornecedores.length - 6) + " outros"] : []));
    caixa("Riscos de engenharia", v.riscos != null ? [v.riscos + " risco(s) · " + (v.riscosAbertos || 0) + " em aberto"] : []);
    caixa("Equipe", (o.analistas || []).length ? o.analistas : v.equipe && v.equipe.length ? v.equipe.map(x => x + " (ClickUp)") : ["Sem analista definido"]);
    return s;
  }
  function tudo() {
    wrap.textContent = ""; wrap.appendChild(topo());
    if (S.msg && !(S.form && S.erro)) wrap.appendChild(el("p", "ob-msg" + (S.erro ? " err" : ""), S.msg));
    wrap.appendChild(kpis()); wrap.appendChild(tabela());
    if (S.form) wrap.appendChild(telaForm());
  }
  if (opts.abrirSigla) { const o = S.lista.find(x => x.sigla === opts.abrirSigla); if (o) S.form = copia(o); }
  tudo();
  return { atualizar: lista => { S.lista = copia(lista || []); if (!S.form) tudo(); } };
}
