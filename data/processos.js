// Aba Processos: fluxos em canvas livre. Módulo isolado do painel: erro aqui não afeta as outras abas.
// Só o admin edita. Quem grava é o painel (opts.salvar), no Supabase.

const SUPORTE = { documento: 1, formulario: 1, indicador: 1 };
const TIPOS = {
  atividade: { nome: "Atividade", w: 138, h: 66 },
  subprocesso: { nome: "Subprocesso", w: 138, h: 66 },
  documento: { nome: "Documento", w: 138, h: 58 },
  formulario: { nome: "Formulário", w: 138, h: 58 },
  indicador: { nome: "Indicador", w: 138, h: 58 },
  inicio: { nome: "Início", w: 96, h: 66 },
  fim: { nome: "Fim", w: 96, h: 66 }
};
const COR = { documento: "#0D9488", formulario: "#8B5CF6", indicador: "#D97706", subprocesso: "#64748B" };
const ICONE = {
  documento: '<path d="M6 2h8l5 5v15H6z" fill="none" stroke="currentColor" stroke-width="2"/><path d="M14 2v5h5" fill="none" stroke="currentColor" stroke-width="2"/>',
  formulario: '<rect x="4" y="3" width="16" height="18" rx="2" fill="none" stroke="currentColor" stroke-width="2"/><path d="M8 8h8M8 12h8M8 16h5" stroke="currentColor" stroke-width="2"/>',
  indicador: '<ellipse cx="12" cy="6" rx="8" ry="3" fill="none" stroke="currentColor" stroke-width="2"/><path d="M4 6v12c0 1.7 3.6 3 8 3s8-1.3 8-3V6M4 12c0 1.7 3.6 3 8 3s8-1.3 8-3" fill="none" stroke="currentColor" stroke-width="2"/>'
};

let seq = 0;
const novoId = p => p + Date.now().toString(36) + (seq++).toString(36);
const copia = o => JSON.parse(JSON.stringify(o));
const linkOk = s => /^https?:\/\/[^\s]+$/i.test(String(s || "").trim());
const suporte = n => !!SUPORTE[n.tipo];

function fluxo(id, nome, itens, setas) {
  const nos = itens.map((n, i) => ({ id: id + i, tipo: n[0], x: n[1], y: n[2] + (n[0] === "inicio" || n[0] === "fim" ? 18 : 0), texto: n[3], link: "" }));
  return { id, nome, nos, setas: setas.map((s, i) => ({ id: id + "s" + i, de: id + s[0], para: id + s[1] })) };
}
export function semente() {
  const y = 150, d = 300, a = 20;
  const px = i => 150 + i * 178;
  return { versao: 1, fases: [
    fluxo("ep", "Estudo Preliminar", [
      ["inicio", a, y, "Início de Estudo Preliminar"],
      ["atividade", px(0), y, "Preenchimento de TAP"], ["atividade", px(1), y, "Definição de Escopo de Projetos"],
      ["atividade", px(2), y, "Contratação de Projetistas"], ["atividade", px(3), y, "Reunião Kick-off com Projetistas"],
      ["atividade", px(4), y, "Elaboração de Estudo Preliminar"], ["atividade", px(5), y, "Análise de Projeto"],
      ["atividade", px(6), y, "Relatório de apontamentos"], ["atividade", px(7), y, "Comitê de Estudo Preliminar"],
      ["fim", px(8), y, "Início de AP"],
      ["documento", px(0), d, "POP Preenchimento TAP"], ["documento", px(1), d, "POP Definição de Escopo de Projetos"],
      ["subprocesso", px(2), d, "Subprocesso de Contratação de Projetista"], ["documento", px(3), d, "Regimento Reunião de Kick-off"],
      ["documento", px(4), d, "Diretrizes de Estudos Preliminares"], ["documento", px(5), d, "POP de Análise de Projetos"],
      ["documento", px(6), d, "POP de Criação de Apontamentos Construflow"], ["documento", px(7), d, "Regimento de Comitê de Estudo Preliminar"],
      ["indicador", px(4), 20, "Indicador de Qualidade de Projetos"]
    ], [[0,1],[1,2],[2,3],[3,4],[4,5],[5,6],[6,7],[7,8],[8,9],[10,1],[11,2],[12,3],[13,4],[14,5],[15,6],[16,7],[17,8],[18,5]]),
    fluxo("ap", "Anteprojeto", [
      ["inicio", a, y, "Início de Anteprojeto"],
      ["atividade", px(0), y, "Elaboração de Estudo Preliminar"], ["atividade", px(1), y, "Análise de Projeto"],
      ["atividade", px(2), y, "Relatório de apontamentos"], ["atividade", px(3), y, "Comitê Anteprojeto"],
      ["fim", px(4), y, "Início de Pré-executivo"],
      ["documento", px(0), d, "Diretrizes de Anteprojetos"], ["documento", px(1), d, "POP de Análise de Projetos"],
      ["documento", px(2), d, "POP de Criação de Apontamentos Construflow"], ["documento", px(3), d, "Regimento de Comitê de Anteprojeto"]
    ], [[0,1],[1,2],[2,3],[3,4],[4,5],[6,1],[7,2],[8,3],[9,4]]),
    fluxo("pe", "Pré-executivo", [
      ["inicio", a, 110, "Início de Pré-executivo"],
      ["subprocesso", 170, 30, "Projeto Macrologístico"], ["subprocesso", 348, 30, "Projeto de Canteiro de Obras"],
      ["atividade", 170, 190, "Elaboração de Pré-executivo"],
      ["atividade", 560, 110, "Análise de Projeto"], ["atividade", 738, 110, "Relatório de apontamentos"],
      ["atividade", 916, 110, "Comitê Pré-executivo"], ["fim", 1094, 110, "Início de Executivo"],
      ["documento", 170, 320, "Diretrizes de Pré-executivos"], ["documento", 560, 260, "POP de Análise de Projetos"],
      ["documento", 738, 260, "POP de Criação de Apontamentos Construflow"], ["documento", 916, 260, "Regimento de Comitê de Pré-executivo"]
    ], [[0,1],[1,2],[0,3],[2,4],[3,4],[4,5],[5,6],[6,7],[8,3],[9,4],[10,5],[11,6]]),
    fluxo("ex", "Executivo", [
      ["inicio", a, y, "Início Executivo"],
      ["atividade", px(0), y, "Elaboração de Executivo"], ["atividade", px(1), y, "Análise de Projeto"],
      ["atividade", px(2), y, "Relatório de Apontamentos"], ["atividade", px(3), y, "Comitê Executivo"],
      ["fim", px(4), y, "Início de Processo de Início de Obra"],
      ["documento", px(0), d, "Diretrizes de Executivos"], ["documento", px(1), d, "POP de Análise de Projetos"],
      ["documento", px(2), d, "POP de Criação de Apontamentos Construflow"], ["documento", px(3), d, "Regimento de Comitê de Pré-executivo"]
    ], [[0,1],[1,2],[2,3],[3,4],[4,5],[6,1],[7,2],[8,3],[9,4]])
  ] };
}

const CSS = `
.pc{display:flex;flex-direction:column;gap:14px;color:var(--color-text,#0F172A)}
.pc *{box-sizing:border-box}
.pc-card{background:var(--card,#fff);border:1px solid var(--slate-100,#F1F5F9);border-radius:16px;box-shadow:var(--card-shadow);padding:14px 16px}
.pc-barra{display:flex;flex-wrap:wrap;gap:8px;align-items:center;justify-content:space-between}
.pc-grupo{display:flex;flex-wrap:wrap;gap:6px;align-items:center}
.pc button{font:inherit;font-size:12.5px;font-weight:700;cursor:pointer;border:1px solid var(--slate-200,#E2E8F0);background:var(--card,#fff);color:var(--slate-600,#475569);border-radius:10px;padding:7px 12px}
.pc button:hover{border-color:var(--blue-500,#3B82F6)}
.pc button.on,.pc button.pri{background:var(--blue-600,#2563EB);border-color:var(--blue-600,#2563EB);color:#fff}
.pc button.per{color:#B91C1C}
.pc button:disabled{opacity:.45;cursor:default}
.pc-corpo{display:flex;gap:12px;align-items:stretch}
.pc-palco{position:relative;flex:1;min-width:0;height:560px;overflow:hidden;border:1px solid var(--slate-100,#F1F5F9);border-radius:16px;background-color:var(--card,#fff);background-image:radial-gradient(var(--slate-200,#E2E8F0) 1px,transparent 1px);background-size:20px 20px;cursor:grab;touch-action:none;user-select:none}
.pc-palco.arrasta{cursor:grabbing}
.pc-palco.pc-fixo{overflow:auto;cursor:default;touch-action:auto;user-select:text;background-image:none;background-color:var(--slate-50,#F8FAFC)}
.pc-palco.pc-fixo .pc-mundo{transform-origin:0 0}
.pc-mundo{position:absolute;left:0;top:0;transform-origin:0 0}
.pc-mundo svg{position:absolute;left:0;top:0;overflow:visible;pointer-events:none}
.pc-seta{fill:none;stroke:var(--slate-500,#64748B);stroke-width:1.8}
.pc-seta.sup{stroke:var(--slate-400,#94A3B8);stroke-dasharray:3 4}
.pc-seta.sel{stroke:var(--blue-600,#2563EB);stroke-width:2.6}
.pc-alvo{fill:none;stroke:transparent;stroke-width:14;pointer-events:stroke;cursor:pointer}
.pc-no{position:absolute;display:flex;align-items:center;justify-content:center;text-align:center;font-size:12px;line-height:1.25;font-weight:700;color:var(--color-text,#0F172A);text-decoration:none;cursor:default;padding:6px 8px;border-radius:12px;background:var(--card,#fff);border:2px solid var(--blue-600,#2563EB)}
.pc-no.subprocesso{padding-bottom:18px;border-color:#64748B;background:var(--slate-100,#F1F5F9)}
.pc-no.subprocesso::after{content:"+";position:absolute;bottom:2px;left:50%;margin-left:-7px;width:14px;height:14px;line-height:12px;font-size:11px;border:1px solid #64748B;border-radius:3px;background:var(--card,#fff)}
.pc-no.sup{justify-content:flex-start;gap:8px;text-align:left;font-weight:600;font-size:11.5px;border-width:1px;border-color:var(--slate-200,#E2E8F0);border-left:4px solid var(--cor)}
.pc-no.sup svg{position:static;flex:none;width:18px;height:18px;color:var(--cor)}
.pc-no.sup.vazio{border-style:dashed;border-left-style:solid;color:var(--slate-500,#64748B)}
a.pc-no{cursor:pointer}a.pc-no:hover{box-shadow:0 0 0 3px var(--blue-50,#EFF6FF);text-decoration:underline}
.pc-no.ini,.pc-no.fim{flex-direction:column;justify-content:flex-start;gap:5px;background:transparent;border:0;padding:0;font-weight:600;font-size:11px}
.pc-no.ini i,.pc-no.fim i{display:block;flex:none;width:30px;height:30px;border-radius:99px;border:3px solid #65A30D;background:#D9F99D}
.pc-no.fim i{border-color:#B91C1C;background:#FCA5A5}
.pc-edita .pc-no{cursor:move}
.pc-no.sel{outline:2px solid var(--blue-500,#3B82F6);outline-offset:2px}
.pc-no.mira{outline:3px solid #10B981;outline-offset:2px}
.pc-ponto{position:absolute;right:-9px;top:50%;margin-top:-9px;width:18px;height:18px;border-radius:99px;background:var(--blue-600,#2563EB);border:3px solid #fff;cursor:crosshair;box-shadow:0 1px 3px rgba(15,23,42,.35)}
.pc-lateral{flex:0 0 250px;display:flex;flex-direction:column;gap:9px;align-self:flex-start}
.pc-lateral h4{margin:0;font-size:10.5px;letter-spacing:.06em;text-transform:uppercase;font-weight:800;color:var(--slate-400,#94A3B8)}
.pc-lateral label{display:flex;flex-direction:column;gap:4px;font-size:11.5px;font-weight:700;color:var(--slate-500,#64748B)}
.pc input,.pc textarea,.pc select{font:inherit;font-size:13px;width:100%;padding:7px 9px;border-radius:9px;border:1px solid var(--slate-200,#E2E8F0);background:var(--card,#fff);color:var(--color-text,#0F172A)}
.pc-paleta{display:grid;grid-template-columns:1fr 1fr;gap:6px}
.pc-aviso{margin:0;font-size:11.5px;line-height:1.4;color:var(--slate-500,#64748B)}
.pc-aviso.err{color:#B91C1C}
.pc-msg{margin:0;font-size:12.5px;padding:9px 12px;border-radius:10px;background:#ECFDF5;color:#065F46}
.pc-msg.err{background:#FEF2F2;color:#991B1B}
.pc-legenda{display:flex;flex-wrap:wrap;gap:6px 16px;font-size:11.5px;color:var(--slate-500,#64748B)}
.pc-legenda span{display:inline-flex;align-items:center;gap:6px}
.pc-legenda b{width:10px;height:10px;border-radius:3px;background:var(--cor)}
.pc-apoio{display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:14px}
.pc-apoio h4{margin:0 0 8px;font-size:13px;font-weight:800}
.pc-apoio ul{margin:0;padding:0;list-style:none;display:flex;flex-direction:column;gap:6px;font-size:12.5px}
.pc-apoio a{color:var(--blue-600,#2563EB);font-weight:700}
.pc-apoio em{font-style:normal;font-size:11px;color:var(--slate-400,#94A3B8)}
@media(max-width:900px){.pc-corpo{flex-direction:column}.pc-lateral{flex:none;width:100%}.pc-palco{height:440px}}
`;

const el = (tag, cls, txt) => { const e = document.createElement(tag); if (cls) e.className = cls; if (txt != null) e.textContent = txt; return e; };
const botao = (txt, fn, cls) => { const b = el("button", cls, txt); b.type = "button"; b.addEventListener("click", fn); return b; };

function normaliza(d) {
  const ok = d && Array.isArray(d.fases) && d.fases.length ? copia(d) : semente();
  ok.fases.forEach(f => { f.nos = Array.isArray(f.nos) ? f.nos.filter(n => n && TIPOS[n.tipo]) : []; f.setas = Array.isArray(f.setas) ? f.setas : [];
    const ids = {}; f.nos.forEach(n => { ids[n.id] = 1; }); f.setas = f.setas.filter(s => ids[s.de] && ids[s.para]); });
  return ok;
}

// Pontas da seta. Fluxo sai à direita e entra à esquerda; apoio liga pelo lado mais próximo.
function caminho(a, b) {
  const A = TIPOS[a.tipo], B = TIPOS[b.tipo];
  const ax = a.x + A.w / 2, ay = a.y + A.h / 2, bx = b.x + B.w / 2, by = b.y + B.h / 2;
  const redondo = t => t === "inicio" || t === "fim";
  const direita = n => redondo(n.tipo) ? n.x + TIPOS[n.tipo].w / 2 + 16 : n.x + TIPOS[n.tipo].w;
  const esquerda = n => redondo(n.tipo) ? n.x + TIPOS[n.tipo].w / 2 - 16 : n.x;
  const cy = n => redondo(n.tipo) ? n.y + 15 : n.y + TIPOS[n.tipo].h / 2;
  if (suporte(a) || suporte(b) || a.tipo === "subprocesso" && Math.abs(ax - bx) < 40 || b.tipo === "subprocesso" && Math.abs(ax - bx) < 40) {
    if (Math.abs(ay - by) >= Math.abs(ax - bx) * .5) {
      const desce = by > ay, y1 = desce ? a.y + A.h : a.y, y2 = desce ? b.y : b.y + B.h;
      return "M" + ax + " " + y1 + " L" + bx + " " + y2;
    }
    const vai = bx > ax;
    return "M" + (vai ? direita(a) : esquerda(a)) + " " + cy(a) + " L" + (vai ? esquerda(b) : direita(b)) + " " + cy(b);
  }
  const x1 = direita(a), y1 = cy(a), x2 = esquerda(b), y2 = cy(b);
  if (x2 >= x1 + 12) { const m = Math.round((x1 + x2) / 2); return y1 === y2 ? "M" + x1 + " " + y1 + " L" + x2 + " " + y2 : "M" + x1 + " " + y1 + " H" + m + " V" + y2 + " H" + x2; }
  const baixo = Math.max(a.y + A.h, b.y + B.h) + 22;
  return "M" + x1 + " " + y1 + " h12 V" + baixo + " H" + (x2 - 12) + " V" + y2 + " H" + x2;
}

export function montar(raiz, opts) {
  if (!document.getElementById("pc-css")) { const st = el("style"); st.id = "pc-css"; st.textContent = CSS; document.head.appendChild(st); }
  const S = {
    pub: normaliza(opts.dados), edita: false, dados: null, hist: [], fase: 0, sel: null, msg: opts.msg || "", erro: !!opts.erroMsg,
    z: 1, px: 20, py: 20, salvando: false
  };
  S.dados = S.pub;
  if (opts.rascunho && opts.rascunho() && opts.admin) { S.dados = normaliza(opts.rascunho()); S.edita = true; }
  const fase = () => S.dados.fases[Math.min(S.fase, S.dados.fases.length - 1)];
  const guarda = () => { if (opts.setRascunho) opts.setRascunho(S.edita ? S.dados : null); };
  const marca = () => { S.hist.push(JSON.stringify(S.dados)); if (S.hist.length > 60) S.hist.shift(); };

  raiz.textContent = "";
  const wrap = el("div", "pc");
  const barra = el("div", "pc-card pc-barra"), msg = el("p"), corpo = el("div", "pc-corpo");
  const palco = el("div", "pc-palco"), mundo = el("div", "pc-mundo"), lateral = el("div", "pc-card pc-lateral");
  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  mundo.appendChild(svg); palco.appendChild(mundo); corpo.append(palco, lateral);
  const legenda = el("div", "pc-legenda"), apoio = el("div", "pc-card pc-apoio");
  wrap.append(barra, msg, corpo, legenda, apoio); raiz.appendChild(wrap);
  [["Atividade", "var(--blue-600,#2563EB)"], ["Subprocesso", COR.subprocesso], ["Documento", COR.documento], ["Formulário", COR.formulario], ["Indicador", COR.indicador]]
    .forEach(l => { const s = el("span"); const b = el("b"); b.style.setProperty("--cor", l[1]); s.append(b, l[0]); legenda.appendChild(s); });
  const dica = el("span", null, "Clique em um item de apoio para abrir. Use as barras de rolagem para ver todo o fluxo.");
  legenda.appendChild(dica);

  const aplicaVista = () => { palco.classList.toggle("pc-fixo", !S.edita); mundo.style.transform = "translate(" + S.px + "px," + S.py + "px) scale(" + S.z + ")";
    const p = mundo.querySelector(".pc-ponto"); if (p) p.style.transform = "scale(" + 1 / S.z + ")"; };
  // Leitura: escala fixa de 100% em todas as fases, com barras de rolagem. Zoom e arraste são só da edição.
  function fixar() {
    const ns = fase().nos; S.z = 1; palco.scrollLeft = 0; palco.scrollTop = 0;
    if (!ns.length) { S.px = S.py = 24; mundo.style.width = mundo.style.height = "0"; aplicaVista(); return; }
    const x0 = Math.min(...ns.map(n => n.x)), y0 = Math.min(...ns.map(n => n.y));
    const x1 = Math.max(...ns.map(n => n.x + TIPOS[n.tipo].w)), y1 = Math.max(...ns.map(n => n.y + TIPOS[n.tipo].h));
    S.px = 24 - x0; S.py = 24 - y0;
    mundo.style.width = (x1 - x0 + 48 - S.px) + "px"; mundo.style.height = (y1 - y0 + 60 - S.py) + "px";
    aplicaVista();
  }
  function ajustar() {
    if (!S.edita) { fixar(); return; }
    mundo.style.width = mundo.style.height = "";
    const ns = fase().nos, W = palco.clientWidth || 900, H = palco.clientHeight || 500;
    if (!ns.length) { S.z = 1; S.px = 20; S.py = 20; aplicaVista(); return; }
    const x0 = Math.min(...ns.map(n => n.x)), y0 = Math.min(...ns.map(n => n.y));
    const x1 = Math.max(...ns.map(n => n.x + TIPOS[n.tipo].w)), y1 = Math.max(...ns.map(n => n.y + TIPOS[n.tipo].h));
    S.z = Math.max(.35, Math.min(1.1, Math.min((W - 44) / (x1 - x0), (H - 44) / (y1 - y0))));
    S.px = (W - (x1 - x0) * S.z) / 2 - x0 * S.z; S.py = (H - (y1 - y0) * S.z) / 2 - y0 * S.z;
    aplicaVista();
  }
  const zoom = f => { const W = palco.clientWidth / 2, H = palco.clientHeight / 2, z = Math.max(.3, Math.min(2, S.z * f));
    S.px = W - (W - S.px) * z / S.z; S.py = H - (H - S.py) * z / S.z; S.z = z; aplicaVista(); };
  const mundoXY = e => { const r = palco.getBoundingClientRect(); return { x: (e.clientX - r.left - S.px) / S.z, y: (e.clientY - r.top - S.py) / S.z }; };

  function desenhaSetas() {
    const f = fase(), mapa = {}; f.nos.forEach(n => { mapa[n.id] = n; });
    let h = '<defs><marker id="pc-pt" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse"><path d="M0 0L10 5L0 10z" fill="var(--slate-500,#64748B)"/></marker></defs>';
    f.setas.forEach(s => { const a = mapa[s.de], b = mapa[s.para]; if (!a || !b) return;
      const d = caminho(a, b), sup = suporte(a) || suporte(b) || a.tipo === "subprocesso" && b.tipo === "atividade" && Math.abs(a.x - b.x) < 40;
      const sel = S.sel && S.sel.seta === s.id;
      h += '<path class="pc-seta' + (sup ? " sup" : "") + (sel ? " sel" : "") + '" d="' + d + '" marker-end="url(#pc-pt)"/>';
      if (S.edita) h += '<path class="pc-alvo" data-seta="' + s.id + '" d="' + d + '"/>'; });
    svg.innerHTML = h;
  }
  function desenhaNos() {
    mundo.querySelectorAll(".pc-no").forEach(n => n.remove());
    fase().nos.forEach(n => {
      const t = TIPOS[n.tipo], lk = linkOk(n.link), ehLink = lk && !S.edita;
      const d = el(ehLink ? "a" : "div", "pc-no " + n.tipo + (n.tipo === "inicio" ? " ini" : "") + (suporte(n) ? " sup" + (lk ? "" : " vazio") : ""));
      d.dataset.no = n.id; d.style.left = n.x + "px"; d.style.top = n.y + "px"; d.style.width = t.w + "px"; d.style.minHeight = t.h + "px";
      if (suporte(n)) d.style.setProperty("--cor", COR[n.tipo]);
      if (ehLink) { d.href = n.link.trim(); d.target = "_blank"; d.rel = "noopener noreferrer"; d.title = "Abrir"; }
      else if (suporte(n) && !lk) d.title = "Link ainda não cadastrado";
      if (ICONE[n.tipo]) { const s = document.createElementNS("http://www.w3.org/2000/svg", "svg"); s.setAttribute("viewBox", "0 0 24 24"); s.innerHTML = ICONE[n.tipo]; d.appendChild(s); }
      if (n.tipo === "inicio" || n.tipo === "fim") d.appendChild(el("i"));
      d.appendChild(el("span", null, n.texto || t.nome));
      if (S.edita && S.sel && S.sel.no === n.id) { d.classList.add("sel"); const pt = el("div", "pc-ponto"); pt.style.transform = "scale(" + 1 / S.z + ")"; d.appendChild(pt); }
      mundo.appendChild(d);
    });
  }
  function desenhaBarra() {
    barra.textContent = "";
    const g = el("div", "pc-grupo");
    S.dados.fases.forEach((f, i) => g.appendChild(botao(f.nome, () => { S.fase = i; S.sel = null; tudo(); ajustar(); }, i === S.fase ? "on" : "")));
    if (S.edita) g.appendChild(botao("+ Fase", () => { const nome = window.prompt("Nome da nova fase:"); if (!nome || !nome.trim()) return;
      marca(); S.dados.fases.push({ id: novoId("f"), nome: nome.trim(), nos: [], setas: [] }); S.fase = S.dados.fases.length - 1; S.sel = null; guarda(); tudo(); ajustar(); }));
    const d = el("div", "pc-grupo");
    if (S.edita) {
      d.append(botao("−", () => zoom(.85)), botao("+", () => zoom(1.18)), botao("Ajustar", ajustar));
      const des = botao("Desfazer", () => { const v = S.hist.pop(); if (!v) return; S.dados = JSON.parse(v); S.sel = null; guarda(); tudo(); }); des.disabled = !S.hist.length;
      const sal = botao(S.salvando ? "Salvando…" : "Salvar e publicar", salvar, "pri"); sal.disabled = S.salvando;
      d.append(des, botao("Cancelar", () => { if (S.hist.length && !window.confirm("Descartar as alterações?")) return;
        S.dados = S.pub; S.edita = false; S.sel = null; S.hist = []; guarda(); S.msg = ""; palco.classList.remove("pc-edita"); tudo(); ajustar(); }), sal);
    } else if (opts.admin) d.appendChild(botao("✎ Editar fluxo", () => { S.dados = copia(S.pub); S.edita = true; S.hist = []; S.msg = ""; guarda(); tudo(); ajustar(); }, "pri"));
    barra.append(g, d);
  }
  function campo(rotulo, ctrl) { const l = el("label", null, rotulo); l.appendChild(ctrl); return l; }
  function desenhaLateral() {
    lateral.style.display = S.edita ? "" : "none"; lateral.textContent = "";
    if (!S.edita) return;
    const f = fase();
    lateral.appendChild(el("h4", null, "Fase"));
    const nome = el("input"); nome.value = f.nome; nome.addEventListener("focus", marca);
    nome.addEventListener("input", () => { f.nome = nome.value; guarda(); }); nome.addEventListener("change", desenhaBarra);
    lateral.appendChild(campo("Nome", nome));
    const mv = el("div", "pc-grupo");
    const mover = dl => { const i = S.dados.fases.indexOf(f), j = i + dl; if (j < 0 || j >= S.dados.fases.length) return; marca(); const fs = S.dados.fases; fs.splice(i, 1); fs.splice(j, 0, f); S.fase = j; guarda(); tudo(); };
    mv.append(botao("← Mover", () => mover(-1)), botao("Mover →", () => mover(1)), botao("Excluir fase", () => {
      if (S.dados.fases.length < 2) { window.alert("O fluxo precisa de pelo menos uma fase."); return; }
      if (!window.confirm("Excluir a fase " + f.nome + " e todo o fluxo dela?")) return;
      marca(); S.dados.fases.splice(S.dados.fases.indexOf(f), 1); S.fase = 0; S.sel = null; guarda(); tudo(); ajustar(); }, "per"));
    lateral.appendChild(mv);
    lateral.appendChild(el("h4", null, "Adicionar"));
    const p = el("div", "pc-paleta");
    Object.keys(TIPOS).forEach(t => p.appendChild(botao(TIPOS[t].nome, () => { marca();
      const c = { x: (palco.clientWidth / 2 - S.px) / S.z - TIPOS[t].w / 2, y: (palco.clientHeight / 2 - S.py) / S.z - TIPOS[t].h / 2 };
      const n = { id: novoId("n"), tipo: t, x: Math.round(c.x / 10) * 10, y: Math.round(c.y / 10) * 10, texto: "", link: "" };
      f.nos.push(n); S.sel = { no: n.id }; guarda(); tudo(); })));
    lateral.appendChild(p);
    const n = S.sel && S.sel.no ? f.nos.find(x => x.id === S.sel.no) : null;
    if (n) {
      lateral.appendChild(el("h4", null, "Item selecionado"));
      const tx = el("textarea"); tx.rows = 3; tx.value = n.texto; tx.addEventListener("focus", marca);
      tx.addEventListener("input", () => { n.texto = tx.value; guarda(); const sp = mundo.querySelector('[data-no="' + n.id + '"] span'); if (sp) sp.textContent = n.texto || TIPOS[n.tipo].nome; });
      lateral.appendChild(campo("Texto", tx));
      const tp = el("select"); Object.keys(TIPOS).forEach(k => { const o = el("option", null, TIPOS[k].nome); o.value = k; o.selected = k === n.tipo; tp.appendChild(o); });
      tp.addEventListener("change", () => { marca(); n.tipo = tp.value; guarda(); tudo(); });
      lateral.appendChild(campo("Tipo", tp));
      const lk = el("input"); lk.placeholder = "https://… link do SharePoint"; lk.value = n.link || ""; lk.addEventListener("focus", marca);
      const av = el("p", "pc-aviso");
      const conf = () => { const v = lk.value.trim(); av.className = "pc-aviso" + (v && !linkOk(v) ? " err" : "");
        av.textContent = !v ? "Sem link: o item fica tracejado." : linkOk(v) ? "Link válido." : "Use um link que comece com https://"; };
      lk.addEventListener("input", () => { n.link = lk.value.trim(); guarda(); conf(); }); conf();
      lk.addEventListener("change", () => { desenhaNos(); desenhaApoio(); });
      lateral.append(campo("Link", lk), av);
      lateral.appendChild(el("p", "pc-aviso", "Para ligar: arraste o ponto azul até outro item."));
      lateral.appendChild(botao("Excluir item", excluir, "per"));
    } else if (S.sel && S.sel.seta) {
      lateral.appendChild(el("h4", null, "Seta selecionada"));
      lateral.appendChild(botao("Excluir seta", excluir, "per"));
    } else lateral.appendChild(el("p", "pc-aviso", "Selecione um item para editar. Arraste para mover. Delete exclui o selecionado."));
  }
  function desenhaApoio() {
    apoio.textContent = "";
    const f = fase();
    [["documento", "Documentos"], ["formulario", "Formulários"], ["indicador", "Indicadores"]].forEach(g => {
      const its = f.nos.filter(n => n.tipo === g[0]); if (!its.length) return;
      const b = el("div"); b.appendChild(el("h4", null, g[1] + " · " + its.length)); const ul = el("ul");
      its.forEach(n => { const li = el("li");
        if (linkOk(n.link)) { const a = el("a", null, n.texto || TIPOS[n.tipo].nome); a.href = n.link.trim(); a.target = "_blank"; a.rel = "noopener noreferrer"; li.appendChild(a); }
        else { li.append(n.texto || TIPOS[n.tipo].nome, " "); li.appendChild(el("em", null, "link pendente")); }
        ul.appendChild(li); });
      b.appendChild(ul); apoio.appendChild(b); });
    apoio.style.display = apoio.childNodes.length ? "" : "none";
  }
  function desenhaMsg() { msg.className = S.msg ? "pc-msg" + (S.erro ? " err" : "") : ""; msg.textContent = S.msg; msg.style.display = S.msg ? "" : "none"; }
  function tudo() { palco.classList.toggle("pc-edita", S.edita); dica.style.display = S.edita ? "none" : ""; desenhaBarra(); desenhaMsg(); desenhaSetas(); desenhaNos(); desenhaLateral(); desenhaApoio(); }

  function excluir() {
    const f = fase(); if (!S.sel) return; marca();
    if (S.sel.no) { f.nos = f.nos.filter(n => n.id !== S.sel.no); f.setas = f.setas.filter(s => s.de !== S.sel.no && s.para !== S.sel.no); }
    else f.setas = f.setas.filter(s => s.id !== S.sel.seta);
    S.sel = null; guarda(); tudo();
  }
  async function salvar() {
    const ruim = []; S.dados.fases.forEach(f => f.nos.forEach(n => { if (n.link && !linkOk(n.link)) ruim.push(n.texto || TIPOS[n.tipo].nome); }));
    if (ruim.length) { S.msg = "Corrija o link de: " + ruim.slice(0, 5).join(", "); S.erro = true; desenhaMsg(); return; }
    S.salvando = true; desenhaBarra();
    const r = await opts.salvar(copia(S.dados));
    S.salvando = false; S.msg = r.msg; S.erro = !r.ok;
    S.pub = S.dados; S.edita = false; S.hist = []; S.sel = null; guarda(); tudo(); ajustar();
  }

  // ── interações ──
  let g = null;
  palco.addEventListener("pointerdown", e => {
    if (e.button || !S.edita) return;
    const noEl = e.target.closest && e.target.closest(".pc-no"), seta = e.target.dataset && e.target.dataset.seta;
    if (S.edita && e.target.classList.contains("pc-ponto")) { const n = fase().nos.find(x => x.id === noEl.dataset.no); g = { tipo: "liga", n, mira: null }; palco.setPointerCapture(e.pointerId); e.preventDefault(); return; }
    if (S.edita && noEl) { const id = noEl.dataset.no, n = fase().nos.find(x => x.id === id), m = mundoXY(e);
      if (!S.sel || S.sel.no !== id) { S.sel = { no: id }; desenhaNos(); desenhaSetas(); desenhaLateral(); }
      g = { tipo: "move", n, dx: m.x - n.x, dy: m.y - n.y, moveu: false, el: mundo.querySelector('[data-no="' + id + '"]') }; palco.setPointerCapture(e.pointerId); e.preventDefault(); return; }
    if (S.edita && seta) { S.sel = { seta }; desenhaSetas(); desenhaLateral(); desenhaNos(); return; }
    if (noEl && noEl.tagName === "A") return;
    if (S.edita && S.sel) { S.sel = null; desenhaNos(); desenhaSetas(); desenhaLateral(); }
    g = { tipo: "pan", x: e.clientX - S.px, y: e.clientY - S.py }; palco.classList.add("arrasta"); palco.setPointerCapture(e.pointerId);
  });
  palco.addEventListener("pointermove", e => {
    if (!g) return;
    if (g.tipo === "pan") { S.px = e.clientX - g.x; S.py = e.clientY - g.y; aplicaVista(); return; }
    if (g.tipo === "move") { const m = mundoXY(e), x = Math.round((m.x - g.dx) / 10) * 10, y = Math.round((m.y - g.dy) / 10) * 10;
      if (x === g.n.x && y === g.n.y) return;
      if (!g.moveu) { marca(); g.moveu = true; }
      g.n.x = x; g.n.y = y; g.el.style.left = x + "px"; g.el.style.top = y + "px"; desenhaSetas(); return; }
    if (g.tipo === "liga") { const a = document.elementFromPoint(e.clientX, e.clientY), al = a && a.closest && a.closest(".pc-no");
      const id = al && al.dataset.no !== g.n.id ? al.dataset.no : null;
      if (id !== g.mira) { mundo.querySelectorAll(".mira").forEach(x => x.classList.remove("mira")); if (al && id) al.classList.add("mira"); g.mira = id; }
      desenhaSetas(); const m = mundoXY(e), t = TIPOS[g.n.tipo];
      const l = document.createElementNS("http://www.w3.org/2000/svg", "path"); l.setAttribute("class", "pc-seta sel"); l.setAttribute("d", "M" + (g.n.x + t.w) + " " + (g.n.y + t.h / 2) + " L" + m.x + " " + m.y); svg.appendChild(l); }
  });
  const fimGesto = () => {
    if (!g) return; const v = g; g = null; palco.classList.remove("arrasta");
    if (v.tipo === "move" && v.moveu) guarda();
    if (v.tipo === "liga") { const f = fase();
      if (v.mira && !f.setas.some(s => s.de === v.n.id && s.para === v.mira)) { marca(); f.setas.push({ id: novoId("s"), de: v.n.id, para: v.mira }); guarda(); }
      tudo(); }
  };
  palco.addEventListener("pointerup", fimGesto); palco.addEventListener("pointercancel", fimGesto);
  palco.addEventListener("wheel", e => { if (!e.ctrlKey || !S.edita) return; e.preventDefault(); zoom(e.deltaY < 0 ? 1.1 : .9); }, { passive: false });
  const tecla = e => { if (!raiz.isConnected) { document.removeEventListener("keydown", tecla); return; }
    if (!S.edita || !S.sel || /INPUT|TEXTAREA|SELECT/.test(e.target.tagName)) return;
    if (e.key === "Delete" || e.key === "Backspace") { e.preventDefault(); excluir(); } };
  document.addEventListener("keydown", tecla);

  tudo(); ajustar();
  return { pendentes: S.pub.fases.reduce((a, f) => a + f.nos.filter(n => suporte(n) && !linkOk(n.link)).length, 0) };
}
