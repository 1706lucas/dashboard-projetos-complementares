// Leitor de .xlsx no navegador + remapeamento das bases do dashboard.
async function unzip(blob) {
  const buf = new Uint8Array(await blob.arrayBuffer());
  const dv = new DataView(buf.buffer);
  let eocd = -1;
  for (let i = buf.length - 22; i >= 0; i--) { if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; } }
  if (eocd < 0) throw new Error("Arquivo não é um .xlsx válido");
  let cdOff = dv.getUint32(eocd + 16, true), n = dv.getUint16(eocd + 10, true);
  if (cdOff === 0xffffffff || n === 0xffff) {
    for (let i = eocd - 20; i >= 0; i--) {
      if (dv.getUint32(i, true) === 0x07064b50) {
        const z = Number(dv.getBigUint64(i + 8, true));
        cdOff = Number(dv.getBigUint64(z + 48, true)); n = Number(dv.getBigUint64(z + 32, true)); break;
      }
    }
  }
  const out = {}; let p = cdOff;
  for (let i = 0; i < n; i++) {
    const cm = dv.getUint16(p + 10, true), csz = dv.getUint32(p + 20, true);
    const fnl = dv.getUint16(p + 28, true), efl = dv.getUint16(p + 30, true), fcl = dv.getUint16(p + 32, true);
    let lho = dv.getUint32(p + 42, true), realCsz = csz, off = lho;
    const name = new TextDecoder().decode(buf.subarray(p + 46, p + 46 + fnl));
    if (csz === 0xffffffff || lho === 0xffffffff) {
      let e = p + 46 + fnl; const end = e + efl;
      while (e < end) {
        const id = dv.getUint16(e, true), sz = dv.getUint16(e + 2, true);
        if (id === 1) {
          let q = e + 4;
          if (dv.getUint32(p + 24, true) === 0xffffffff) q += 8;
          if (csz === 0xffffffff) { realCsz = Number(dv.getBigUint64(q, true)); q += 8; }
          if (lho === 0xffffffff) off = Number(dv.getBigUint64(q, true));
        }
        e += 4 + sz;
      }
    }
    p += 46 + fnl + efl + fcl;
    const lfnl = dv.getUint16(off + 26, true), lefl = dv.getUint16(off + 28, true);
    const ds = off + 30 + lfnl + lefl;
    const data = buf.subarray(ds, ds + realCsz);
    out[name] = cm === 0 ? new TextDecoder().decode(data)
      : await new Response(new Blob([data]).stream().pipeThrough(new DecompressionStream("deflate-raw"))).text();
  }
  return out;
}

const dec = s => s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&amp;/g, "&");
const shared = xml => !xml ? [] : [...xml.matchAll(/<si>([\s\S]*?)<\/si>/g)]
  .map(m => dec([...m[1].matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map(t => t[1]).join("")));
const colIdx = r => { const s = r.match(/[A-Z]+/)[0]; let n = 0; for (const c of s) n = n * 26 + (c.charCodeAt(0) - 64); return n - 1; };

function sheetRows(xml, ss) {
  const rows = [];
  for (const rm of xml.matchAll(/<row[^>]*>([\s\S]*?)<\/row>/g)) {
    const cells = [];
    for (const cm of rm[1].matchAll(/<c ([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
      const a = cm[1], b = cm[2] || "";
      const ref = (a.match(/r="([A-Z]+\d+)"/) || [])[1], t = (a.match(/t="(\w+)"/) || [])[1];
      let v = null;
      if (t === "inlineStr") v = dec([...b.matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map(x => x[1]).join(""));
      else { const vm = b.match(/<v>([\s\S]*?)<\/v>/); if (vm) { v = vm[1]; v = t === "s" ? ss[+v] : dec(v); } }
      if (ref && v != null && v !== "") cells[colIdx(ref)] = v;
    }
    if (cells.length) rows.push(cells);
  }
  return rows;
}

export async function readWorkbook(file) {
  const z = await unzip(file);
  const ss = shared(z["xl/sharedStrings.xml"]);
  const rels = {};
  for (const m of (z["xl/_rels/workbook.xml.rels"] || "").matchAll(/Id="(rId\d+)"[^>]*?Target="([^"]+)"/g))
    rels[m[1]] = m[2].replace(/^\/?xl\//, "").replace(/^\//, "");
  const sheets = [...(z["xl/workbook.xml"] || "").matchAll(/<sheet [^>]*\/>/g)].map(m => ({
    name: dec((m[0].match(/name="([^"]+)"/) || [])[1] || ""), rid: (m[0].match(/r:id="(rId\d+)"/) || [])[1]
  }));
  const wb = {};
  for (const s of sheets) { const xml = z["xl/" + (rels[s.rid] || "")]; if (xml) wb[s.name] = sheetRows(xml, ss); }
  return wb;
}

const ser = v => {
  if (v == null || v === "") return null;
  const n = parseFloat(v); if (isNaN(n)) return null;
  return new Date(Date.UTC(1899, 11, 30) + Math.round(n) * 864e5).toISOString().slice(0, 10);
};
const num = v => {
  if (v == null || v === "") return null;
  const n = parseFloat(String(v).replace(/[^\d.,-]/g, "").replace(",", ".")); return isNaN(n) ? null : n;
};
const norm = s => String(s || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase()
  .replace(/\bii\b/g, "2").replace(/\bi\b/g, "1").replace(/[^a-z0-9]+/g, " ").trim();
const clean = s => String(s || "").replace(/\s*·\s*/g, " · ").replace(/\s+/g, " ").trim();

// export do ClickUp agrupado: linhas-título de grupo + cabeçalho repetido por grupo
function achaCabecalho(rows) {
  for (let i = 0; i < Math.min(rows.length, 12); i++) {
    const r = (rows[i] || []).map(v => String(v == null ? "" : v));
    if (r.indexOf("Task Name") >= 0 || r.indexOf("Task ID") >= 0) return i;
  }
  return 0;
}
const colClasse = H => (H || []).findIndex(h => /^classes?\s+datas?\s+marco/i.test(String(h || "").trim()));
export function detectKind(wb) {
  const first = wb[Object.keys(wb)[0]] || [];
  const H0 = (first[achaCabecalho(first)] || []).map(String);
  if (H0.indexOf("Classe de Restrição (drop down)") >= 0) return "pendencias";
  const H = (first[0] || []).map(String);
  if (H.indexOf("Lists") >= 0 && H.some(h => /Impacto do Risco|Data final PE/.test(h))) return "indicadores";
  if (colClasse(H) >= 0 || (H.indexOf("Milestone") >= 0 && H.indexOf("Task Name") >= 0)) return "cronograma";
  if (Object.keys(wb).some(k => /contratado/i.test(k)) || Object.keys(wb).some(k => /^NF$/i.test(k))) return "saidas";
  if (Object.keys(wb).some(k => /^geral$/i.test(k)) && Object.keys(wb).some(k => /OR[ÇC]AMENTOS/i.test(k))) return "saidasGadens";
  if (Object.keys(wb).some(k => (wb[k][0] || []).join("|").toUpperCase().indexOf("NOME OBRA") >= 0)) return "siglas";
  return null;
}

// ── indicadores ──
export function mapIndicadores(wb) {
  const t = wb[Object.keys(wb)[0]], H = t[0], k = n => H.indexOf(n);
  const g = (r, n) => r[k(n)] ?? null;
  const rows = t.slice(1), pick = l => rows.filter(r => g(r, "Lists") === l);
  const notaCols = ["NOTA Clareza das Informações (drop down)", "NOTA Compatibilização (drop down)",
    "NOTA Solução Otimizada (drop down)", "NOTA Conformidade Técnica (drop down)", "NOTA Confiabilidade na entrga (drop down)"];  const linhaBase = pick("Linha de Base Oficializada").map(r => ({
    obra: g(r, "Obra (drop down)") || g(r, "Task Name"), empreendimento: g(r, "Task Name"),
    dataLB: ser(g(r, "Data final PE (date)")), status: g(r, "Status"), bu: g(r, "Unidade de Negócio (drop down)")
  }));
  return {
    riscos: pick("Riscos por empreendimento").map(r => ({
      obra: g(r, "Obra (drop down)") || g(r, "Empreendimento (drop down)"), tarefa: g(r, "Task Name"),
      disciplina: g(r, "Disciplina (labels)"), tipo: g(r, "Ameaça ou Oportunidade (drop down)"),
      dimensao: g(r, "Aferação (drop down)"), impacto: num(g(r, "Impacto do Risco (drop down)")),
      probabilidade: num(g(r, "Probabilidade de ocorrência do Risco (drop down)")),
      fase: g(r, "Fase do Projeto (drop down)"), status: g(r, "Status"),
      bu: g(r, "Unidade de Negócio (drop down)"), responsavel: g(r, "Assignee"),
      plano: g(r, "Plano de Ação (text)"), prazo: ser(g(r, "Due Date"))
    })),
    linhaBase,
    desvioPE: linhaBase.map(x => ({ obra: x.obra, dataPE: x.dataLB, bu: x.bu, status: x.status })),
    desvioRegistros: pick("Desvio de LB de PE").map(r => ({
      obra: g(r, "Obra (drop down)") || g(r, "Empreendimento (drop down)"), disciplina: g(r, "Disciplina (labels)"),
      responsavel: g(r, "Assignee"), fornecedor: g(r, "Fornecedor (drop down)"),
      dataPE: ser(g(r, "Data final PE (date)")), dias: num(g(r, "Dias de atraso (number)")),
      motivo: g(r, "Motivos de atraso (labels)"), plano: g(r, "Plano de Ação (text)"), registro: g(r, "Task Name")
    })),
    estouro: pick("Estouro de obra por conta de projetos").map(r => ({
      obra: g(r, "Obra (drop down)") || g(r, "Empreendimento (drop down)"), valor: num(g(r, "Estouro de Projeto (currency)")),
      justificativa: g(r, "Task Name"), disciplina: g(r, "Disciplina (labels)"),
      bu: g(r, "Unidade de Negócio (drop down)"), causaRaiz: null
    })),
    alteracoes: pick("Alterações de Projeto").map(r => ({
      obra: g(r, "Obra (drop down)") || g(r, "Empreendimento (drop down)"), disciplina: g(r, "Disciplina (labels)"),
      motivo: g(r, "Motivo da Alteração (drop down)"), gravidade: g(r, "Gravidade da alteração (drop down)"),
      descricao: g(r, "Descrição da Alteração (text)") || g(r, "Task Name"),
      n: num(g(r, "Número de alterações (number)")), bu: g(r, "Unidade de Negócio (drop down)")
    })),
    qualidade: pick("Qualidade de entrega").map(r => {
      const notas = notaCols.map(c => num(g(r, c))).filter(v => v != null);
      return {
        empreendimento: g(r, "Empreendimento (drop down)"), disciplina: g(r, "Disciplina (labels)"),
        fornecedor: g(r, "Fornecedor (drop down)"), fase: g(r, "Fase do Projeto (drop down)"),
        bu: g(r, "Unidade de Negócio (drop down)"), notas,
        media: notas.length ? +(notas.reduce((a, b) => a + b, 0) / notas.length).toFixed(2) : null
      };
    }),
    economia: rows.filter(r => /Economia/i.test(g(r, "Lists") || "") || num(g(r, "Valor Economizado (currency)"))).map(r => ({
      obra: g(r, "Obra (drop down)") || g(r, "Empreendimento (drop down)"), disciplina: g(r, "Disciplina (labels)"),
      valor: num(g(r, "Valor Economizado (currency)")),
      descricao: g(r, "Descrição de Economia (text)") || g(r, "Task Name"), bu: g(r, "Unidade de Negócio (drop down)")
    }))
  };
}

// ── cronograma: recalcula os campos por obra ──
const DONE = ["CONCLUÍDO", "CONCLUÍDO COM ATRASO", "COMPLETE"];
const PROG = ["EM PROGRESSO", "EM PROGRESSO FORA DO PRAZO", "EM ATRASO", "IN PROGRESS", "EM VALIDAÇÃO TÉCNICA", "PAUSADO"];
const kindOf = s => DONE.indexOf(s) >= 0 ? "done" : PROG.indexOf(s) >= 0 ? "prog" : s === "CANCELADO" ? "cancel" : "todo";
const ORDEM = ["Estudo Preliminar", "Anteprojeto", "Pré-executivo", "Executivo"];
const faseDe = l => {
  const t = norm(l);
  if (t.indexOf("estudo preliminar") >= 0) return "Estudo Preliminar";
  if (t.indexOf("anteprojeto") >= 0) return "Anteprojeto";
  if (t.indexOf("pre executivo") >= 0) return "Pré-executivo";
  if (t.indexOf("revisao executivo") >= 0 || t.indexOf("projeto executivo") >= 0 || /\bexecutivo\b/.test(t)) return "Executivo";
  return null;
};

const isCompl = t => String(t || "").split(",").some(x => /^\s*(eng\s*-\s*)?complementares\s*$/i.test(x));

// Fotografia das atividades complementares, para comparar bases entre uploads
export function snapshotCompl(wb) {
  const t = wb[Object.keys(wb)[0]], H = t[0], i = n => H.indexOf(n);
  const C = { id: i("Task ID"), name: i("Task Name"), st: i("Status"), sd: i("Start Date"), dd: i("Due Date"), tg: i("tags"), cl: colClasse(H) };
  const nm = r => (r[C.name] || "").replace(/^\s*\ud83d\udccc\s*/, "").trim();
  const MARCO = ["Marco Axial", "Marco Task"];
  const out = {};
  t.slice(1).forEach(r => {
    const id = r[C.id];
    if (!id || !isCompl(r[C.tg]) || MARCO.indexOf(r[C.cl]) < 0) return;
    out[id] = { nome: nm(r), status: r[C.st], start: ser(r[C.sd]), due: ser(r[C.dd]), classe: r[C.cl] };
  });
  return out;
}

// Compara a fotografia anterior com a atual e devolve o bloco de atrasos
export function comparaAtrasos(prevSnap, curSnap, obras, dataAnterior, dataAtual) {
  const sig = {};
  obras.forEach(o => { sig[o.sigla] = o; });
  const itens = [];
  let novas = 0;
  Object.keys(curSnap).forEach(id => {
    const c = curSnap[id], p = prevSnap ? prevSnap[id] : null;
    const pre = (c.nome || "").slice(0, 3).toUpperCase();
    const ob = sig[pre]; if (!ob) return;
    if (!p) novas++;
    const anterior = p ? p.due : null;
    const dias = anterior && c.due ? Math.round((new Date(c.due) - new Date(anterior)) / 864e5) : null;
    itens.push({ id, nome: c.nome, obra: ob.nome, sigla: pre, bu: ob.bu, status: c.status, classe: c.classe,
      start: c.start, anterior, atual: c.due, dias, url: "https://app.clickup.com/t/" + id });
  });
  const total = itens.length;
  const comAtraso = itens.filter(x => x.dias && x.dias > 0).sort((a, b) => b.dias - a.dias);
  return { baseAnterior: dataAnterior || null, baseAtual: dataAtual, total, novas,
    comAtraso: comAtraso.length, itens: comAtraso.slice(0, 300) };
}

export function mapCronograma(wb, obras, hoje) {
  const t = wb[Object.keys(wb)[0]], H = t[0], i = n => H.indexOf(n);
  const C = { name: i("Task Name"), st: i("Status"), sd: i("Start Date"), dd: i("Due Date"),
    tg: i("tags"), list: i("Lists"), cl: colClasse(H) };
  if (C.cl < 0) throw new Error("export sem a coluna Classe Datas Marco — refaça pela visão ATUALIZAÇÃO DE DATAS do ClickUp");
  const nm = r => (r[C.name] || "").replace(/^\s*📌\s*/, "").trim();
  const rows = t.slice(1), pre = r => nm(r).slice(0, 3).toUpperCase();
  const comClasse = rows.filter(r => r[C.cl]).length;
  const soMarcos = rows.length > 0 && comClasse / rows.length > 0.8;
  return obras.map(o0 => {
    const o = Object.assign({}, o0);
    const mine = rows.filter(r => pre(r) === o.sigla);
    if (!mine.length && o0.temDados) return Object.assign(o, { foraDoExport: true });
    delete o.foraDoExport;
    if (!soMarcos || !(o0.totalCompl > 0)) {
      const compl = mine.filter(r => isCompl(r[C.tg]));
      o.totalCompl = compl.length;
      o.doneCompl = compl.filter(r => DONE.indexOf(r[C.st]) >= 0).length;
      o.pct = o.totalCompl ? Math.round(o.doneCompl / o.totalCompl * 100) : 0;
      const fase = {}; compl.forEach(r => { const s = r[C.st] || "—"; fase[s] = (fase[s] || 0) + 1; });
      o.fases = Object.keys(fase).map(s => ({ status: s, n: fase[s], pct: o.totalCompl ? Math.round(fase[s] / o.totalCompl * 100) : 0 }))
        .sort((a, b) => b.n - a.n);
    }
    o.soMarcos = soMarcos;
    o.axial = mine.filter(r => r[C.cl] === "Marco Axial").map(r => ({
      label: clean(nm(r).split("|").slice(1).join(" · ")), status: r[C.st], kind: kindOf(r[C.st]),
      start: ser(r[C.sd]), end: ser(r[C.dd])
    })).sort((a, b) => String(a.end || a.start).localeCompare(String(b.end || b.start)));
    o.gantt = mine.filter(r => r[C.cl] === "Marco Task").map(r => ({
      label: clean(nm(r).split("|").slice(1).join(" · ")), grupo: clean((nm(r).split("|")[1] || "")),
      status: r[C.st], kind: kindOf(r[C.st]), start: ser(r[C.sd]), end: ser(r[C.dd]), tag: r[C.tg] || ""
    })).filter(g => g.start && g.end).sort((a, b) => a.start.localeCompare(b.start));
    o.lista = (mine.map(r => r[C.list]).filter(Boolean)[0]) || o.lista || "";
    o.temDados = mine.length > 0 || (soMarcos && !!o0.temDados);
    const pe = mine.find(r => r[C.cl] === "Marco Axial" && /pré-execu/i.test(nm(r)));
    o.dataPE = pe ? ser(pe[C.dd]) : null;
    // fases (marcos axiais)
    const fmap = {};
    mine.filter(r => r[C.cl] === "Marco Axial").forEach(r => {
      const f = faseDe(nm(r).split("|").slice(1).join(" ")); if (!f) return;
      const s = ser(r[C.sd]), d = ser(r[C.dd]); if (!s && !d) return;
      const cur = fmap[f] || (fmap[f] = { start: null, end: null, status: r[C.st] });
      if (s && (!cur.start || s < cur.start)) cur.start = s;
      if (d && (!cur.end || d > cur.end)) cur.end = d;
    });
    o.fasesLinha = ORDEM.filter(f => fmap[f]).map(f => {
      const x = fmap[f];
      return { fase: f, start: x.start, end: x.end, status: x.status,
        dias: x.start && x.end ? Math.round((new Date(x.end) - new Date(x.start)) / 864e5) : null };
    });
    const vs = mine.filter(r => /viabilid/i.test(nm(r))).map(r => ser(r[C.sd])).filter(Boolean).sort();
    const primeiro = o.fasesLinha.length ? o.fasesLinha[0].start : null;
    const inicio = [vs[0], primeiro].filter(Boolean).sort()[0] || null;
    const exec = o.fasesLinha.find(f => f.fase === "Executivo");
    o.vida = { inicio, fim: exec ? exec.end : null,
      origem: inicio && inicio === vs[0] ? "Estudo de viabilidade" : "Primeiro marco axial",
      dias: inicio && exec && exec.end ? Math.round((new Date(exec.end) - new Date(inicio)) / 864e5) : null };
    const pef = o.fasesLinha.find(f => f.fase === "Pré-executivo");
    o.pe = pef ? { start: pef.start, end: pef.end, dias: pef.dias } : null;
    o.pePlus = !!(pef && pef.start && pef.start <= hoje);
    o.inicioObra = mine.some(r => r[C.cl] === "Marco Axial" && norm(nm(r)).indexOf("inicio de obra") >= 0);
    return o;
  });
}

// ── SAÍDAS: contratado + NF + benchmark ──
export function mapSaidas(wb, bu) {
  const key = re => Object.keys(wb).find(k => re.test(k));
  const out = { contratado: [], nf: [], benchmark: [] };
  const tag = x => Object.assign(x, bu ? { bu } : {});

  // layout Make It: abas Contratado / NF / CNPJ
  const kc = key(/^contratado$/i) || key(/contratado/i);
  if (kc) out.contratado = wb[kc].slice(1).filter(r => r[0]).map(r => tag({
    empreendimento: r[0], projeto: r[1], contratado: r[2], valor: num(r[3])
  }));
  const kn = key(/^NF$/i);
  if (kn) out.nf = wb[kn].slice(1).filter(r => r[0]).map(r => tag({
    empreendimento: r[0], disciplina: r[1], contratado: r[2], valor: num(r[3]),
    status: r[5], mes: num(r[6]), ano: num(r[7])
  }));
  const kb = key(/CNPJ/i);
  if (kb) out.benchmark = wb[kb].slice(1).filter(r => r[0]).map(r => {
    const m2 = num(r[6]), uh = num(r[5]), bud = num(r[8]), usado = num(r[12]);
    return tag({ obra: r[0], unidades: uh, m2, budget: bud, utilizado: usado, vgv: num(r[7]),
      rM2: m2 && usado ? usado / m2 : null, rUH: uh && usado ? usado / uh : null,
      refM2: m2 && bud ? bud / m2 : null, refUH: uh && bud ? bud / uh : null });
  }).filter(x => x.rM2 != null || x.rUH != null);

  // layout Gadens: aba GERAL (pagamentos) + ORÇAMENTOS
  const kg = key(/^geral$/i);
  if (!kn && kg) {
    const rows = wb[kg].slice(1).filter(r => r[0]).map(r => ({
      empreendimento: r[0], disciplina: r[1], contratado: r[2] || null,
      valorTotal: num(r[3]), valor: num(r[4]), serial: num(r[5]), status: r[6],
      mes: num(r[7]), ano: num(r[8])
    }));
    out.nf = rows.map(x => tag({ empreendimento: x.empreendimento, disciplina: x.disciplina,
      contratado: x.contratado, valor: x.valor, status: x.status, mes: x.mes, ano: x.ano }));
    const best = {};
    rows.forEach(x => {
      if (x.valorTotal == null) return;
      const k = x.empreendimento + "|" + x.disciplina + "|" + (x.contratado || "");
      if (!best[k] || (x.serial || 0) > (best[k].serial || 0)) best[k] = x;
    });
    out.contratado = Object.keys(best).map(k => best[k]).map(x => tag({
      empreendimento: x.empreendimento, projeto: x.disciplina,
      contratado: x.contratado || "—", valor: x.valorTotal
    }));
  }
  const ko = key(/OR[ÇC]AMENTOS/i);
  if (!kb && ko) out.benchmark = wb[ko].slice(1).filter(r => r[0]).map(r => {
    const uh = num(r[1]), m2 = num(r[2]), vgv = num(r[3]), bud = num(r[4]), usado = num(r[8]);
    return tag({ obra: r[0], unidades: uh, m2, budget: bud, utilizado: usado, vgv,
      rM2: m2 && usado ? usado / m2 : null, rUH: uh && usado ? usado / uh : null,
      refM2: m2 && bud ? bud / m2 : null, refUH: uh && bud ? bud / uh : null });
  }).filter(x => x.rM2 != null || x.rUH != null);
  return out;
}

export function mapSiglas(wb) {
  const obras = [];
  const bu = { "MAKE IT": "Make It", "GADENS INC": "Gadens Inc" };
  Object.keys(wb).forEach(k => {
    const H = (wb[k][0] || []).join("|").toUpperCase();
    if (H.indexOf("NOME OBRA") < 0) return;
    wb[k].slice(1).forEach(r => {
      if (r[0] && r[1]) obras.push({ sigla: String(r[0]).trim(), nome: String(r[1]).trim(), bu: bu[k.toUpperCase()] || k });
    });
  });
  return obras;
}


// ── Pendências de Gestão (COMPL | ATIVIDADES - Lista Pendências Gestão) ──
const CLASSE_RESTR = c => {
  const t = norm(c);
  if (t.indexOf("incorpora") >= 0) return "inc";
  if (t.indexOf("obra") >= 0) return "obra";
  return "ger";
};
export function mapPendencias(wb) {
  const rows = wb[Object.keys(wb)[0]] || [];
  const out = []; let H = null;
  rows.forEach(r => {
    const cells = (r || []).map(v => v == null ? "" : String(v));
    if (cells.indexOf("Task Name") >= 0) { H = cells; return; }
    if (!H) return;
    const g = n => { const i = H.indexOf(n); return i >= 0 ? (r[i] ?? null) : null; };
    const nome = g("Task Name"); if (!nome) return;
    const st = String(g("Status") || "");
    out.push({
      url: g("Task URL") || g("URL") || (g("Task ID") ? "https://app.clickup.com/t/" + g("Task ID") : ""),
      id: g("Task ID"), tarefa: String(nome).replace(/^[A-Z0-9]{2,4}\s*-\s*/, "").trim(), titulo: nome,
      obra: g("Obra (drop down)"), classe: g("Classe de Restrição (drop down)"), tipo: CLASSE_RESTR(g("Classe de Restrição (drop down)")),
      prazo: ser(g("Due Date")), plano: String(g("Plano de ação (text)") || "").trim(),
      prioridade: g("Priority"), status: st,
      aberta: !/complete|conclu|closed|done|cancel/i.test(st)
    });
  });
  return out;
}
