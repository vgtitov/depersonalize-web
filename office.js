// Таблицы (CSV, XLSX) и документы (DOCX) локально. Нужен JSZip (vendor/jszip.min.js), движок engine.js. Без сети.
const Z = typeof JSZip !== "undefined" ? JSZip : require("./vendor/jszip.min.js");
const E = typeof depersonalize !== "undefined" ? { depersonalize, restore } : require("./engine.js");

// ---------- общее ----------
const xmlUnesc = s => s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(+n)).replace(/&#x([0-9a-f]+);/gi, (_, n) => String.fromCodePoint(parseInt(n, 16))).replace(/&amp;/g, "&");
const xmlEsc = s => String(s).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
// Ключ вида ЧЕЛОВЕК_001 в тексте ответа ИИ → исходное значение по карте.
function restoreKeys(text, map) {
  return E.restore(text, map).replace(/(?<![\wА-Яа-яЁё])[A-ZА-ЯЁ]+_\d{3,}(?![\wА-Яа-яЁё])/g, k => map.has(k) ? map.get(k) : k);
}

// ---------- CSV ----------
function parseCsv(text) {
  text = text.replace(/^\uFEFF/, "");
  const first = text.split("\n", 1)[0], cnt = c => first.split(c).length;
  const delim = [";", "\t", ","].sort((a, b) => cnt(b) - cnt(a))[0];
  const rows = []; let row = [], cell = "", q = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (q) { if (c === '"') { if (text[i + 1] === '"') { cell += '"'; i++ } else q = false } else cell += c }
    else if (c === '"') q = true;
    else if (c === delim) { row.push(cell); cell = "" }
    else if (c === "\n" || c === "\r") { if (c === "\r" && text[i + 1] === "\n") i++; row.push(cell); rows.push(row); row = []; cell = "" }
    else cell += c;
  }
  if (cell !== "" || row.length) { row.push(cell); rows.push(row) }
  return { rows, delim };
}
function toCsv(rows, delim = ";") {
  const q = v => { v = String(v ?? ""); return /["\n\r]/.test(v) || v.includes(delim) ? '"' + v.replace(/"/g, '""') + '"' : v };
  return "\uFEFF" + rows.map(r => r.map(q).join(delim)).join("\r\n") + "\r\n";
}

// ---------- XLSX ----------
const colIdx = ref => { const s = /^[A-Z]+/.exec(ref)[0]; let n = 0; for (const ch of s) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1 };
const colName = i => { let s = ""; for (i++; i > 0; i = Math.floor((i - 1) / 26)) s = String.fromCharCode(65 + (i - 1) % 26) + s; return s };
const textsOf = xml => [...xml.matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)].map(m => xmlUnesc(m[1])).join("");
async function readXlsx(buf) {
  const zip = await Z.loadAsync(buf);
  const rd = async p => zip.file(p) ? await zip.file(p).async("string") : null;
  const wb = await rd("xl/workbook.xml"); if (!wb) throw new Error("это не файл Excel (.xlsx)");
  const rels = await rd("xl/_rels/workbook.xml.rels") || "";
  const relTarget = {}; for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)) { const id = /\bId="([^"]*)"/.exec(m[0]), t = /\bTarget="([^"]*)"/.exec(m[0]); if (id && t) relTarget[id[1]] = t[1] }
  const ss = await rd("xl/sharedStrings.xml");
  const shared = ss ? [...ss.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)].map(m => textsOf(m[1].replace(/<rPh\b[\s\S]*?<\/rPh>/g, ""))) : [];
  const sheets = [];
  for (const m of wb.matchAll(/<sheet\b[^>]*>/g)) {
    const name = /\bname="([^"]*)"/.exec(m[0]), rid = /\br:id="([^"]*)"/.exec(m[0]);
    if (!name || !rid || !relTarget[rid[1]]) continue;
    let t = relTarget[rid[1]]; t = t.startsWith("/") ? t.slice(1) : "xl/" + t;
    const xml = await rd(t); if (!xml) continue;
    const rows = [];
    for (const c of xml.matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
      const ref = /\br="([A-Z]+)(\d+)"/.exec(c[1]); if (!ref) continue;
      const type = (/\bt="([^"]*)"/.exec(c[1]) || [])[1], inner = c[2] || "";
      const v = (/<v>([\s\S]*?)<\/v>/.exec(inner) || [])[1];
      let val = "";
      if (type === "s") val = shared[+v] ?? "";
      else if (type === "inlineStr") val = textsOf(inner);
      else if (type === "str") val = v == null ? "" : xmlUnesc(v);
      else if (type === "b") val = v === "1" ? "TRUE" : "FALSE";
      else if (type === "e") val = "";
      else if (v != null && v !== "" && Number.isFinite(+v)) val = +v;
      const r = +ref[2] - 1, k = colIdx(ref[1]);
      (rows[r] = rows[r] || [])[k] = val;
    }
    const w = Math.max(0, ...rows.filter(Boolean).map(r => r.length));
    sheets.push({ name: xmlUnesc(name[1]), rows: Array.from(rows, r => Array.from({ length: w }, (_, i) => (r && r[i] != null) ? r[i] : "")) });
  }
  if (!sheets.length) throw new Error("в файле нет листов");
  return sheets;
}
async function writeXlsx(rows, sheetName = "Лист1") {
  const zip = new Z();
  zip.file("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');
  zip.file("_rels/.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  zip.file("xl/workbook.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="' + xmlEsc(sheetName).replace(/"/g, "&quot;") + '" sheetId="1" r:id="rId1"/></sheets></workbook>');
  zip.file("xl/_rels/workbook.xml.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>');
  const body = rows.map((r, i) => '<row r="' + (i + 1) + '">' + r.map((v, k) => {
    const ref = colName(k) + (i + 1);
    if (v === "" || v == null) return "";
    return typeof v === "number" ? '<c r="' + ref + '"><v>' + v + "</v></c>" : '<c r="' + ref + '" t="inlineStr"><is><t xml:space="preserve">' + xmlEsc(v) + "</t></is></c>";
  }).join("") + "</row>").join("");
  zip.file("xl/worksheets/sheet1.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' + body + "</sheetData></worksheet>");
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}

// ---------- колонки таблицы ----------
const KINDS = ["ЧЕЛОВЕК", "ТЕЛЕФОН", "EMAIL", "ОРГ", "ЗНАЧЕНИЕ"];
const HINTS = [["EMAIL", /mail|почт/i], ["ТЕЛЕФОН", /тел|phone|моб/i], ["ЧЕЛОВЕК", /фио|фамил|\bимя\b|сотрудник|менеджер|ответствен/i], ["ОРГ", /организ|компан|контрагент|клиент|поставщик|юрлиц/i]];
// Подсказка по колонке; выбор остаётся за человеком.
function detectColumns(rows, hasHeader = true) {
  const head = hasHeader ? rows[0] || [] : [], body = rows.slice(hasHeader ? 1 : 0), w = Math.max(0, ...rows.map(r => r.length));
  return Array.from({ length: w }, (_, i) => {
    const name = String(head[i] ?? "") || "Колонка " + colName(i);
    const vals = body.map(r => r[i]).filter(v => v !== "" && v != null).map(String), sample = vals.slice(0, 40);
    let kind = "ЗНАЧЕНИЕ", mode = "keep";
    const full = {};
    for (const v of sample) { const o = E.depersonalize(v, []).text.trim(); const m = /^\[([А-ЯЁA-Z]+)_\d+\]$/.exec(o); if (m) full[m[1]] = (full[m[1]] || 0) + 1 }
    const top = Object.entries(full).sort((a, b) => b[1] - a[1])[0];
    if (top && top[1] >= 0.6 * sample.length) { kind = KINDS.includes(top[0]) ? top[0] : "ЗНАЧЕНИЕ"; mode = "key" }
    else { const h = HINTS.find(([, re]) => re.test(name)); if (h) { kind = h[0]; mode = "key" } }
    return { index: i, name, kind, mode, preview: vals.slice(0, 3) };
  });
}
// Применение выбора: keep / key (ключи вместо значений) / drop (убрать колонку) / scan (найти и скрыть внутри текста ячеек).
function anonymizeTable(rows, cols, { hasHeader = true, dict = [] } = {}) {
  const map = new Map(), reg = new Map(), n = {}, counts = {}, warn = [];
  const keyOf = (kind, v) => { const norm = String(v).trim().replace(/\s+/g, " ").toLowerCase(), rk = kind + "|" + norm;
    if (!reg.has(rk)) { n[kind] = (n[kind] || 0) + 1; counts[kind] = n[kind]; const k = kind + "_" + String(n[kind]).padStart(3, "0"); reg.set(rk, k); map.set(k, String(v).trim()) }
    return reg.get(rk) };
  const out = rows.map(r => r.slice());
  const start = hasHeader ? 1 : 0;
  for (const c of cols) {
    if (c.mode === "key") for (let r = start; r < out.length; r++) { const v = out[r][c.index]; if (v !== "" && v != null) out[r][c.index] = keyOf(c.kind, v) }
    if (c.mode === "scan") {
      const cells = []; for (let r = start; r < out.length; r++) cells.push(String(out[r][c.index] ?? ""));
      const res = E.depersonalize(cells.join("\n"), dict), parts = res.text.split("\n");
      for (const [t, v] of res.map) map.set(t, v);
      for (const [k, v] of Object.entries(res.counts)) counts[k] = (counts[k] || 0) + v;
      for (let r = start; r < out.length; r++) if (cells[r - start] !== "") out[r][c.index] = parts[r - start];
    }
  }
  const drop = new Set(cols.filter(c => c.mode === "drop").map(c => c.index));
  return { rows: out.map(r => r.filter((_, i) => !drop.has(i))), map, counts, warn };
}
function restoreRows(rows, map) { return rows.map(r => r.map(v => typeof v === "string" ? restoreKeys(v, map) : v)) }

// ---------- DOCX ----------
const PARTS = /^word\/(document|header\d*|footer\d*|footnotes|endnotes|comments)\.xml$/;
// Прогоняет текст каждого абзаца через fn(массив строк) → массив строк. Оформление внутри изменённого абзаца берётся от его первого фрагмента.
async function mapDocx(buf, fn) {
  const zip = await Z.loadAsync(buf);
  if (!zip.file("word/document.xml")) throw new Error("это не файл Word (.docx)");
  const items = []; // {part, xml, paras:[{start,end,text,runs}]}
  for (const name of Object.keys(zip.files).filter(n => PARTS.test(n))) {
    const xml = await zip.file(name).async("string"), paras = [];
    for (const m of xml.matchAll(/<w:p[ >][\s\S]*?<\/w:p>/g)) {
      const runs = [...m[0].matchAll(/(<w:(t|delText)\b(?![^>]*\/>)[^>]*>)([\s\S]*?)(<\/w:\2>)/g)];
      if (!runs.length) continue;
      paras.push({ start: m.index, end: m.index + m[0].length, src: m[0], runs, text: runs.map(r => xmlUnesc(r[3])).join("") });
    }
    items.push({ name, xml, paras });
  }
  const all = items.flatMap(i => i.paras), res = fn(all.map(p => p.text));
  if (res.length !== all.length) throw new Error("внутренняя ошибка: число абзацев изменилось");
  let k = 0, changed = 0;
  for (const it of items) {
    let out = "", pos = 0;
    for (const p of it.paras) {
      const nt = res[k++]; out += it.xml.slice(pos, p.start); pos = p.end;
      if (nt === p.text) { out += p.src; continue }
      changed++; let i = 0;
      out += p.src.replace(/(<w:(t|delText)\b(?![^>]*\/>))([^>]*>)([\s\S]*?)(<\/w:\2>)/g, (_, a, tag, attrs, _body, close) =>
        i++ === 0 ? a + " xml:space=\"preserve\"" + attrs.replace(/\s*xml:space="[^"]*"/, "") + xmlEsc(nt) + close : a + attrs + close);
    }
    zip.file(it.name, out + it.xml.slice(pos));
  }
  // автор и редактор в свойствах файла
  const core = zip.file("docProps/core.xml");
  if (core) zip.file("docProps/core.xml", (await core.async("string")).replace(/(<(dc:creator|cp:lastModifiedBy)>)[\s\S]*?(<\/\2>)/g, "$1$3"));
  return { data: await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), changed };
}
async function anonymizeDocx(buf, dict = []) {
  let r;
  let preview = [];
  const o = await mapDocx(buf, texts => { r = E.depersonalize(texts.join("\n"), dict); const out = r.text.split("\n");
    preview = texts.map((t, i) => [t, out[i]]).filter(([a, b]) => a !== b).slice(0, 6); return out });
  return { data: o.data, map: r.map, counts: r.counts, warn: r.warn, changed: o.changed, preview };
}
async function restoreDocx(buf, map) { return (await mapDocx(buf, texts => texts.map(t => restoreKeys(t, map)))).data }

// Простой .docx для примеров: абзацы и таблица. Открывается в Word и LibreOffice.
async function makeDocx({ paragraphs = [], table = [], author = "" }) {
  const p = t => '<w:p><w:r><w:t xml:space="preserve">' + xmlEsc(t) + "</w:t></w:r></w:p>";
  const tbl = table.length ? '<w:tbl><w:tblPr><w:tblBorders>' + ["top", "left", "bottom", "right", "insideH", "insideV"].map(b => '<w:' + b + ' w:val="single" w:sz="4" w:space="0" w:color="999999"/>').join("") + "</w:tblBorders></w:tblPr>" +
    table.map(r => "<w:tr>" + r.map(c => "<w:tc><w:tcPr><w:tcW w:w=\"2400\" w:type=\"dxa\"/></w:tcPr>" + p(c) + "</w:tc>").join("") + "</w:tr>").join("") + "</w:tbl><w:p/>" : "";
  const z = new Z();
  z.file("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>');
  z.file("_rels/.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>');
  z.file("docProps/core.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:creator>' + xmlEsc(author) + "</dc:creator><cp:lastModifiedBy>" + xmlEsc(author) + "</cp:lastModifiedBy></cp:coreProperties>");
  z.file("word/document.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' + paragraphs.map(p).join("") + tbl + "</w:body></w:document>");
  return z.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}

const api = { makeDocx, restoreKeys, parseCsv, toCsv, readXlsx, writeXlsx, detectColumns, anonymizeTable, restoreRows, anonymizeDocx, restoreDocx, mapDocx, KINDS };
if (typeof module !== "undefined") module.exports = api;
