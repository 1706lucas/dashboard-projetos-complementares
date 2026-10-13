// Exportação: .xlsx (Por Empreendimento) e apresentação HTML imprimível.
const TB = new TextEncoder();

let CRC_T = null;
function crcTable() {
  if (CRC_T) return CRC_T;
  CRC_T = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    CRC_T[n] = c >>> 0;
  }
  return CRC_T;
}
function crc32(buf) {
  const t = crcTable(); let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) c = t[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

// zip sem compressão (método store) — suficiente para .xlsx
function zipStore(files) {
  const parts = [], central = [];
  let offset = 0;
  const dt = new Date();
  const time = ((dt.getHours() << 11) | (dt.getMinutes() << 5) | (dt.getSeconds() / 2)) & 0xffff;
  const date = (((dt.getFullYear() - 1980) << 9) | ((dt.getMonth() + 1) << 5) | dt.getDate()) & 0xffff;
  for (const f of files) {
    const name = TB.encode(f.name), data = typeof f.data === "string" ? TB.encode(f.data) : f.data;
    const crc = crc32(data);
    const lh = new DataView(new ArrayBuffer(30));
    lh.setUint32(0, 0x04034b50, true); lh.setUint16(4, 20, true); lh.setUint16(6, 0, true);
    lh.setUint16(8, 0, true); lh.setUint16(10, time, true); lh.setUint16(12, date, true);
    lh.setUint32(14, crc, true); lh.setUint32(18, data.length, true); lh.setUint32(22, data.length, true);
    lh.setUint16(26, name.length, true); lh.setUint16(28, 0, true);
    parts.push(new Uint8Array(lh.buffer), name, data);
    const ch = new DataView(new ArrayBuffer(46));
    ch.setUint32(0, 0x02014b50, true); ch.setUint16(4, 20, true); ch.setUint16(6, 20, true);
    ch.setUint16(8, 0, true); ch.setUint16(10, 0, true); ch.setUint16(12, time, true); ch.setUint16(14, date, true);
    ch.setUint32(16, crc, true); ch.setUint32(20, data.length, true); ch.setUint32(24, data.length, true);
    ch.setUint16(28, name.length, true); ch.setUint16(30, 0, true); ch.setUint16(32, 0, true);
    ch.setUint16(34, 0, true); ch.setUint16(36, 0, true); ch.setUint32(38, 0, true); ch.setUint32(42, offset, true);
    central.push(new Uint8Array(ch.buffer), name);
    offset += 30 + name.length + data.length;
  }
  const cdSize = central.reduce((a, x) => a + x.length, 0);
  const eo = new DataView(new ArrayBuffer(22));
  eo.setUint32(0, 0x06054b50, true); eo.setUint16(8, files.length, true); eo.setUint16(10, files.length, true);
  eo.setUint32(12, cdSize, true); eo.setUint32(16, offset, true);
  return new Blob([...parts, ...central, new Uint8Array(eo.buffer)], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
}

const esc = s => String(s == null ? "" : s)
  .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const colRef = n => { let s = ""; n++; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = (n - r - 1) / 26; } return s; };

// rows: matriz de valores (string | number)
export function buildXlsx(sheetName, rows) {
  const sheetRows = rows.map((r, ri) => {
    const cells = r.map((v, ci) => {
      const ref = colRef(ci) + (ri + 1);
      if (v == null || v === "") return "";
      if (typeof v === "number") return '<c r="' + ref + '"><v>' + v + "</v></c>";
      return '<c r="' + ref + '" t="inlineStr"><is><t xml:space="preserve">' + esc(v) + "</t></is></c>";
    }).join("");
    return '<row r="' + (ri + 1) + '">' + cells + "</row>";
  }).join("");
  const sheet = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<cols><col min="1" max="1" width="22" customWidth="1"/><col min="2" max="2" width="11" customWidth="1"/>' +
    '<col min="3" max="3" width="12" customWidth="1"/><col min="4" max="4" width="58" customWidth="1"/>' +
    '<col min="5" max="5" width="22" customWidth="1"/><col min="6" max="6" width="12" customWidth="1"/></cols>' +
    "<sheetData>" + sheetRows + "</sheetData></worksheet>";
  const wb = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<sheets><sheet name="' + esc(sheetName) + '" sheetId="1" r:id="rId1"/></sheets></workbook>';
  const wbRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
    "</Relationships>";
  const rootRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
    "</Relationships>";
  const ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
    "</Types>";
  return zipStore([
    { name: "[Content_Types].xml", data: ct },
    { name: "_rels/.rels", data: rootRels },
    { name: "xl/workbook.xml", data: wb },
    { name: "xl/_rels/workbook.xml.rels", data: wbRels },
    { name: "xl/worksheets/sheet1.xml", data: sheet }
  ]);
}

export function download(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 4000);
}

// ── apresentação imprimível ──
export function buildApresentacao(ctx) {
  const { titulo, geradoEm, obras } = ctx;
  const page = o => {
    const axial = o.axial.map(a => '<div class="pin"><span class="dot ' + a.kind + '">' + (a.kind === "done" ? "✓" : "") +
      '</span><div class="pl">' + esc(a.label) + '</div><div class="pd">' + esc(a.date) + "</div></div>").join("");
    const fases = o.fases.map(f => '<tr><td>' + esc(f.fase) + "</td><td>" + esc(f.periodo) +
      '</td><td class="r">' + esc(f.dias) + "</td></tr>").join("");
    const mat = o.matriz.map(c => '<div class="mc" style="background:' + c.bg + ";color:" + c.fg + '">' + (c.n || "") + "</div>").join("");
    const inds = o.indicadores.map(i => '<tr><td>' + esc(i.label) + '</td><td class="r">' + esc(i.valor) +
      '</td><td class="r"><b style="color:' + i.cor + '">' + esc(i.nota) + "</b></td></tr>").join("");
    const riscos = o.riscosTop.map(r => '<tr><td>' + esc(r.tarefa) + "</td><td>" + esc(r.disciplina) +
      '</td><td class="r">' + esc(r.score) + "</td><td>" + esc(r.status) + "</td></tr>").join("");
    return '<section class="pg">' +
      '<header><div><div class="kick">' + esc(o.bu) + " · " + esc(o.sigla) + '</div><h1>' + esc(o.nome) + "</h1></div>" +
      '<div class="fase"><span>Fase atual</span><b>' + esc(o.faseAtual) + "</b></div></header>" +
      '<div class="kpis">' + o.kpis.map(k => '<div class="kpi"><span>' + esc(k.label) + "</span><b>" + esc(k.valor) +
        '</b><i>' + esc(k.sub) + "</i></div>").join("") + "</div>" +
      '<h2>Linha axial</h2><div class="axial">' + (axial || '<p class="empty">Sem marcos axiais cadastrados.</p>') + "</div>" +
      '<div class="cols"><div><h2>Fases de projeto</h2><table><thead><tr><th>Fase</th><th>Período</th><th class="r">Dias</th></tr></thead><tbody>' +
        (fases || '<tr><td colspan="3" class="empty">Sem fases cadastradas.</td></tr>') + "</tbody></table></div>" +
      '<div><h2>Matriz de risco</h2><div class="mwrap"><div class="mgrid">' + mat + "</div>" +
        '<div class="mleg"><span>Impacto ↑ · Probabilidade →</span><span>' + esc(o.matrizNota) + "</span></div></div></div></div>" +
      '<h2>Indicadores do empreendimento</h2><table><thead><tr><th>Indicador</th><th class="r">Valor</th><th class="r">Nota</th></tr></thead><tbody>' +
        (inds || '<tr><td colspan="3" class="empty">Sem indicadores.</td></tr>') + "</tbody></table>" +
      '<h2>Riscos em aberto de maior score</h2><table><thead><tr><th>Risco</th><th>Disciplina</th><th class="r">I × P</th><th>Status</th></tr></thead><tbody>' +
        (riscos || '<tr><td colspan="4" class="empty">Nenhum risco em aberto.</td></tr>') + "</tbody></table>" +
      '<footer>' + esc(titulo) + " · gerado em " + esc(geradoEm) + "</footer></section>";
  };
  return '<!DOCTYPE html><html lang="pt-BR"><head><meta charset="utf-8"><title>' + esc(titulo) + "</title>" +
    '<link href="https://fonts.googleapis.com/css2?family=Nunito:wght@400;600;700;800;900&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet">' +
    "<style>" +
    "*{box-sizing:border-box}body{margin:0;background:#F1F5F9;font-family:Nunito,sans-serif;color:#1E293B}" +
    ".pg{background:#fff;width:1100px;max-width:100%;margin:24px auto;padding:44px 48px;box-shadow:0 2px 10px rgba(15,23,42,.08)}" +
    "header{display:flex;justify-content:space-between;align-items:flex-start;gap:24px;border-bottom:2px solid #E2E8F0;padding-bottom:18px}" +
    ".kick{font-size:11px;letter-spacing:.08em;text-transform:uppercase;font-weight:800;color:#64748B}" +
    "h1{font-size:30px;margin:6px 0 0;font-weight:900;letter-spacing:-.01em}" +
    ".fase{text-align:right}.fase span{display:block;font-size:10px;letter-spacing:.06em;text-transform:uppercase;font-weight:800;color:#94A3B8}" +
    ".fase b{font-size:16px}" +
    "h2{font-size:14px;margin:28px 0 12px;font-weight:800}" +
    ".kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-top:22px}" +
    ".kpi{border:1px solid #E2E8F0;border-left:4px solid #2563EB;padding:12px 14px}" +
    ".kpi span{display:block;font-size:9.5px;letter-spacing:.06em;text-transform:uppercase;font-weight:800;color:#94A3B8}" +
    ".kpi b{display:block;font-size:19px;margin-top:4px}.kpi i{font-style:normal;font-size:11px;color:#64748B}" +
    ".axial{display:flex;gap:10px;flex-wrap:wrap}" +
    ".pin{flex:1;min-width:112px;text-align:center}" +
    ".dot{display:flex;align-items:center;justify-content:center;width:30px;height:30px;border-radius:99px;margin:0 auto 8px;font-size:13px;font-weight:800;border:2px solid #CBD5E1;color:#fff}" +
    ".dot.done{background:#2563EB;border-color:#2563EB}.dot.prog{background:#F59E0B;border-color:#F59E0B}.dot.todo{background:#fff;border-color:#CBD5E1}" +
    ".pl{font-size:11.5px;font-weight:700;line-height:1.25}.pd{font-family:'DM Mono',monospace;font-size:10.5px;color:#94A3B8;margin-top:3px}" +
    ".cols{display:grid;grid-template-columns:1fr 300px;gap:28px;align-items:start}" +
    "table{width:100%;border-collapse:collapse;font-size:12px}" +
    "th{text-align:left;font-size:9.5px;letter-spacing:.06em;text-transform:uppercase;color:#94A3B8;padding:7px 8px;border-bottom:1px solid #E2E8F0}" +
    "td{padding:8px;border-bottom:1px solid #F1F5F9;vertical-align:top}.r{text-align:right;font-family:'DM Mono',monospace}" +
    ".empty{color:#94A3B8;font-size:11.5px}" +
    ".mgrid{display:grid;grid-template-columns:repeat(5,1fr);gap:4px}" +
    ".mc{height:34px;display:flex;align-items:center;justify-content:center;font-weight:800;font-size:12px}" +
    ".mleg{margin-top:8px;display:flex;flex-direction:column;gap:4px;font-size:10px;color:#94A3B8}" +
    "footer{margin-top:32px;padding-top:12px;border-top:1px solid #E2E8F0;font-size:10px;color:#94A3B8}" +
    "@media print{body{background:#fff}.pg{margin:0;width:auto;box-shadow:none;padding:14mm;page-break-after:always}}" +
    "@page{size:A4 landscape;margin:0}" +
    "</style></head><body>" + obras.map(page).join("") + "</body></html>";
}


// ── material da Reunião de Diretoria ──
export function buildReuniao(ctx) {
  const pinCls = k => k === "done" ? "done" : k === "prog" ? "prog" : k === "atraso" ? "late" : "todo";
  const page = o => {
    const restr = (o.restricoes || []).map(l => '<div class="rl"><h3><i style="background:' + l.cor + '"></i>' + esc(l.titulo) +
      ' <span>' + l.itens.length + '</span></h3>' + (l.itens.length ? '<ul>' + l.itens.map(t => typeof t === "string" ? '<li>' + esc(t) + '</li>' :
        '<li>' + (t.url ? '<a href="' + esc(t.url) + '" target="_blank" style="color:#1D4ED8;font-weight:700;text-decoration:none">' + esc(t.t) + ' ↗</a>' : '<b>' + esc(t.t) + '</b>') +
        (t.det ? '<br><span style="color:#64748B">' + esc(t.det) + '</span>' : '') + '</li>').join('') + '</ul>'
      : '<p class="empty">Nenhuma restrição registrada.</p>') + '</div>').join('');
    const axial = o.axial.map(a => '<div class="pin"><span class="dot ' + pinCls(a.kind) + '">' + (a.kind === "done" ? "✓" : "") +
      '</span><div class="pl">' + esc(a.label) + '</div><div class="pd">' + esc(a.date) + '</div></div>').join('');
    const mat = o.matriz.map(c => '<div class="mc" style="background:' + c.bg + ';color:' + c.fg + '">' + (c.n || '') + '</div>').join('');
    const riscos = (o.riscosLista || []).map(r => '<tr><td>' + esc(r.tarefa) + '</td><td>' + esc(r.disciplina) +
      '</td><td class="r">' + esc(r.score) + '</td><td>' + esc(r.status) + '</td></tr>').join('');
    return '<section class="pg"><header><div><div class="kick">' + esc(o.bu) + ' · ' + esc(o.sigla) + '</div><h1>' + esc(o.nome) +
      '</h1></div><div class="fase"><span>Fase atual</span><b>' + esc(o.faseAtual) + '</b></div></header>' +
      '<div class="restr">' + restr + '</div>' +
      '<h2>Linha axial de projetos</h2><div class="axial">' + (axial || '<p class="empty">Sem marcos axiais cadastrados.</p>') + '</div>' +
      '<div class="cols"><div><h2>Matriz de risco</h2><div class="mgrid">' + mat + '</div><div class="mleg"><span>Impacto ↑ · Probabilidade →</span><span>' +
      esc(o.matrizNota) + '</span></div></div><div><h2>Riscos de engenharia</h2><table><thead><tr><th>Risco</th><th>Disciplina</th><th class="r">I × P</th><th>Status</th></tr></thead><tbody>' +
      (riscos || '<tr><td colspan="4" class="empty">Nenhum risco cadastrado.</td></tr>') + '</tbody></table></div></div>' +
      '<footer>' + esc(ctx.titulo) + ' · gerado em ' + esc(ctx.geradoEm) + '</footer></section>';
  };
  return '<!DOCTYPE html><html lang="pt-BR"><head><meta charset="utf-8"><title>' + esc(ctx.titulo) + '</title>' +
    '<link href="https://fonts.googleapis.com/css2?family=Nunito:wght@400;600;700;800;900&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet"><style>' +
    '*{box-sizing:border-box;-webkit-print-color-adjust:exact;print-color-adjust:exact}body{margin:0;background:#F1F5F9;font-family:Nunito,sans-serif;color:#1E293B}' +
    '.pg{background:#fff;width:1100px;max-width:100%;margin:24px auto;padding:36px 42px;box-shadow:0 2px 10px rgba(15,23,42,.08)}' +
    'header{display:flex;justify-content:space-between;align-items:flex-end;gap:24px;border-bottom:2px solid #E2E8F0;padding-bottom:14px}' +
    '.kick{font-size:11px;letter-spacing:.08em;text-transform:uppercase;font-weight:800;color:#64748B}h1{font-size:26px;margin:4px 0 0;font-weight:900}' +
    '.fase{text-align:right}.fase span{display:block;font-size:10px;letter-spacing:.06em;text-transform:uppercase;font-weight:800;color:#94A3B8}.fase b{font-size:14px}' +
    'h2{font-size:13px;margin:22px 0 10px;font-weight:800}' +
    '.restr{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-top:18px}.rl{border:1px solid #E2E8F0;border-radius:12px;padding:12px 14px}' +
    '.rl h3{display:flex;align-items:center;gap:7px;font-size:12.5px;margin:0 0 8px}.rl h3 i{width:8px;height:8px;border-radius:99px}.rl h3 span{margin-left:auto;font-family:"DM Mono",monospace;color:#94A3B8;font-weight:500}' +
    '.rl ul{margin:0;padding-left:16px;font-size:11.5px;line-height:1.5}.rl li{margin-bottom:3px}' +
    '.axial{display:flex;gap:8px}.pin{flex:1;text-align:center}.dot{display:flex;align-items:center;justify-content:center;width:24px;height:24px;border-radius:99px;margin:0 auto 6px;font-size:11px;font-weight:800;border:2px solid #CBD5E1;color:#fff}' +
    '.dot.done{background:#2563EB;border-color:#2563EB}.dot.prog{background:#F59E0B;border-color:#F59E0B}.dot.late{background:#EF4444;border-color:#EF4444}' +
    '.pl{font-size:10.5px;font-weight:700;line-height:1.2}.pd{font-family:"DM Mono",monospace;font-size:9.5px;color:#94A3B8;margin-top:2px}' +
    '.cols{display:grid;grid-template-columns:250px 1fr;gap:26px;align-items:start}.mgrid{display:grid;grid-template-columns:repeat(5,1fr);gap:3px}' +
    '.mc{height:30px;border-radius:6px;display:flex;align-items:center;justify-content:center;font-weight:800;font-size:11.5px}.mleg{margin-top:6px;display:flex;flex-direction:column;gap:3px;font-size:9.5px;color:#94A3B8}' +
    'table{width:100%;border-collapse:collapse;font-size:11px}th{text-align:left;font-size:9px;letter-spacing:.06em;text-transform:uppercase;color:#94A3B8;padding:6px;border-bottom:1px solid #E2E8F0}' +
    'td{padding:6px;border-bottom:1px solid #F1F5F9;vertical-align:top}.r{text-align:right;font-family:"DM Mono",monospace}.empty{color:#94A3B8;font-size:11px;margin:0}' +
    'footer{margin-top:22px;padding-top:10px;border-top:1px solid #E2E8F0;font-size:9.5px;color:#94A3B8}' +
    '@media print{body{background:#fff}.pg{margin:0;width:auto;box-shadow:none;padding:10mm 12mm;page-break-after:always}}@page{size:A4 landscape;margin:0}' +
    '</style></head><body>' + ctx.obras.map(page).join('') +
    '<script>window.addEventListener("load",function(){setTimeout(function(){window.print();},600);});<\/script></body></html>';
}

export function zipar(files) { const b = zipStore(files); return new Blob([b], { type: "application/zip" }); }
