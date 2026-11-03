// Таблицы (CSV, XLSX) и документы (DOCX) локально. Нужен JSZip (vendor/jszip.min.js), движок engine.js. Без сети.
const Z = typeof JSZip !== "undefined" ? JSZip : require("./vendor/jszip.min.js");
const E = typeof depersonalize !== "undefined" ? { depersonalize, restore, joinCells, splitCells } : require("./engine.js");

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
function isDateFmt(id, codes) {
  if ((id >= 14 && id <= 22) || (id >= 45 && id <= 47)) return true;
  const c = codes[id]; return !!c && /[dmyhs]/i.test(c.replace(/"[^"]*"|\[[^\]]*\]|\\./g, ""));
}
function serialToText(n) {
  const ms = Math.round((n - 25569) * 86400000), d = new Date(ms), p = (x, l = 2) => String(x).padStart(l, "0");
  const day = d.getUTCFullYear() + "-" + p(d.getUTCMonth() + 1) + "-" + p(d.getUTCDate());
  return n % 1 === 0 ? day : day + " " + p(d.getUTCHours()) + ":" + p(d.getUTCMinutes());
}
async function readXlsx(buf) {
  const zip = await Z.loadAsync(buf);
  const rd = async p => zip.file(p) ? await zip.file(p).async("string") : null;
  const wb = await rd("xl/workbook.xml"); if (!wb) throw new Error("это не файл Excel (.xlsx)");
  const rels = await rd("xl/_rels/workbook.xml.rels") || "";
  const relTarget = {}; for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)) { const id = /\bId="([^"]*)"/.exec(m[0]), t = /\bTarget="([^"]*)"/.exec(m[0]); if (id && t) relTarget[id[1]] = t[1] }
  const ss = await rd("xl/sharedStrings.xml");
  const shared = ss ? [...ss.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)].map(m => textsOf(m[1].replace(/<rPh\b[\s\S]*?<\/rPh>/g, ""))) : [];
  const st = await rd("xl/styles.xml") || "", codes = {};
  for (const m of st.matchAll(/<numFmt\b[^>]*>/g)) { const id = /numFmtId="(\d+)"/.exec(m[0]), c = /formatCode="([^"]*)"/.exec(m[0]); if (id && c) codes[+id[1]] = xmlUnesc(c[1]) }
  const xfs = (/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(st) || [, ""])[1];
  const dateXf = [...xfs.matchAll(/<xf\b[^>]*>/g)].map(m => { const f = /numFmtId="(\d+)"/.exec(m[0]); return !!f && isDateFmt(+f[1], codes) });
  const sheets = [];
  for (const m of wb.matchAll(/<sheet\b[^>]*>/g)) {
    const name = /\bname="([^"]*)"/.exec(m[0]), rid = /\br:id="([^"]*)"/.exec(m[0]);
    if (!name || !rid || !relTarget[rid[1]]) continue;
    let t = relTarget[rid[1]]; t = t.startsWith("/") ? t.slice(1) : "xl/" + t;
    const xml = await rd(t); if (!xml) continue;
    const rows = []; let rn = 0;
    for (const rw of xml.matchAll(/<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g)) {
      const rr = /\br="(\d+)"/.exec(rw[1]); rn = rr ? +rr[1] : rn + 1; let cn = -1;
      for (const c of (rw[2] || "").matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
        const ref = /\br="([A-Z]+)\d+"/.exec(c[1]); cn = ref ? colIdx(ref[1]) : cn + 1;
        const type = (/\bt="([^"]*)"/.exec(c[1]) || [])[1], sidx = (/\bs="(\d+)"/.exec(c[1]) || [])[1], inner = c[2] || "";
        const v = (/<v>([\s\S]*?)<\/v>/.exec(inner) || [])[1];
        let val = "";
        if (type === "s") val = shared[+v] ?? "";
        else if (type === "inlineStr") val = textsOf(inner);
        else if (type === "str" || type === "e") val = v == null ? "" : xmlUnesc(v);
        else if (type === "b") val = v === "1" ? "TRUE" : "FALSE";
        else if (v != null && v !== "" && Number.isFinite(+v)) val = sidx != null && dateXf[+sidx] ? serialToText(+v) : +v;
        if (val !== "") (rows[rn - 1] = rows[rn - 1] || [])[cn] = val;
      }
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
const HINTS = [["ЗНАЧЕНИЕ", /парол|password|pwd|токен|token|секрет|secret|\bpin\b|пин-?код/i, "drop"], ["ЗНАЧЕНИЕ", /инн|унп|паспорт|iban|карт[аы]|адрес|счёт|счет/i, "key"], ["EMAIL", /mail|почт/i], ["ТЕЛЕФОН", /тел|phone|моб/i], ["ЧЕЛОВЕК", /фио|фамил|\bимя\b|сотрудник|менеджер|ответствен/i], ["ОРГ", /организ|компан|контрагент|клиент|поставщик|юрлиц/i]];
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
    else { const h = HINTS.find(([, re]) => re.test(name)); if (h) { kind = h[0]; mode = h[2] || "key" } }
    return { index: i, name, kind, mode, preview: vals.slice(0, 3) };
  });
}
// Применение выбора: keep / key (ключи вместо значений) / drop (убрать колонку) / scan (найти и скрыть внутри текста ячеек).
function anonymizeTable(rows, cols, { hasHeader = true, dict = [] } = {}) {
  const map = new Map(), reg = new Map(), n = {}, counts = {}, warn = [];
  const keyOf = (kind, v) => { const norm = String(v).trim().replace(/\s+/g, " ").toLowerCase(), rk = kind + "|" + norm;
    if (!reg.has(rk)) { n[kind] = (n[kind] || 0) + 1; counts[kind] = n[kind]; const k = kind + "_" + String(n[kind]).padStart(3, "0"); reg.set(rk, k); map.set(k, String(v).trim()) }
    return reg.get(rk) };
  const out = rows.map(r => r.slice()), start = hasHeader ? 1 : 0, scan = [];
  for (const c of cols) {
    if (c.mode === "key") for (let r = start; r < out.length; r++) { const v = out[r][c.index]; if (v !== "" && v != null) out[r][c.index] = keyOf(c.kind, v) }
    if (c.mode === "scan") for (let r = start; r < out.length; r++) if (out[r][c.index] !== "" && out[r][c.index] != null) scan.push([r, c.index]);
    if (c.mode === "keep") for (let r = start; r < out.length; r++) if (/^[A-ZА-ЯЁ]+_\d{3,}$/.test(String(out[r][c.index] ?? ""))) { warn.push("в колонке «" + (rows[0] && hasHeader ? rows[0][c.index] : c.index + 1) + "» есть значения, похожие на коды (ЧЕЛОВЕК_001): при возврате они могут замениться"); break }
  }
  if (scan.length) { // все ячейки режима «внутри текста» одной пачкой: общая нумерация, переносы строк не мешают
    const res = E.depersonalize(E.joinCells(scan.map(([r, c]) => String(out[r][c]))), dict), parts = E.splitCells(res.text, scan.length);
    scan.forEach(([r, c], i) => { out[r][c] = parts[i] });
    for (const [t, v] of res.map) map.set(t, v);
    for (const [k, v] of Object.entries(res.counts)) counts[k] = (counts[k] || 0) + v;
  }
  const drop = new Set(cols.filter(c => c.mode === "drop").map(c => c.index));
  return { rows: out.map(r => r.filter((_, i) => !drop.has(i))), map, counts, warn };
}
function restoreRows(rows, map) {
  return rows.map(r => r.map(v => {
    if (typeof v !== "string") return v;
    const whole = /^[A-ZА-ЯЁ]+_\d{3,}$/.test(v) && map.get(v);
    if (whole && /^(?:0|[1-9]\d{0,14})$/.test(whole)) return Number(whole); // число, которое было числом
    return restoreKeys(v, map);
  }));
}

// ---------- DOCX ----------
const PARTS = /^(?:word|customXml)\/(?!_rels\/).*\.xml$/;
const attrEsc = s => xmlEsc(s).replace(/"/g, "&quot;");
// Группы текста в одной части: каждый абзац (в том числе вложенный в надпись) и каждый кусок внутри исправления (w:ins, w:del) — отдельно.
function textGroups(xml) {
  const re = /<(w|a):p[ >]|<\/(?:w|a):p>|<(w:ins|w:del)\b(?![^>]*\/>)[^>]*>|<\/(?:w:ins|w:del)>|<(w:t|w:delText|a:t)\b(?![^>]*\/>)([^>]*)>([\s\S]*?)<\/\3>/g;
  const paras = [], chg = [], groups = new Map(); let pid = 0, cid = 0;
  for (const m of xml.matchAll(re)) {
    if (m[1]) paras.push(++pid);
    else if (m[0].startsWith("</w:ins") || m[0].startsWith("</w:del")) chg.pop();
    else if (m[2]) chg.push(++cid);
    else if (m[0].startsWith("</")) paras.pop();
    else {
      const key = (paras[paras.length - 1] || 0) + ":" + (chg[chg.length - 1] || 0);
      const run = { open: m.index, ca: m.index + 1 + m[3].length + m[4].length + 1, name: m[3], attrs: m[4], raw: m[5] };
      run.cb = run.ca + m[5].length; run.end = run.cb + m[3].length + 3;
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(run);
    }
  }
  return [...groups.values()];
}
// Прогоняет текст каждой группы через fn(массив строк) → массив строк. Оформление внутри изменённого абзаца берётся от его первого фрагмента.
async function mapDocx(buf, fn, { clean = false } = {}) {
  const zip = await Z.loadAsync(buf);
  if (!zip.file("word/document.xml")) throw new Error("это не файл Word (.docx)");
  const items = [], warn = [], names = Object.keys(zip.files).filter(n => !zip.files[n].dir);
  for (const name of names) {
    if (PARTS.test(name)) {
      const xml = await zip.file(name).async("string"), groups = textGroups(xml);
      items.push({ name, xml, kind: "xml", groups, texts: groups.map(g => xmlUnesc(g.map(r => r.raw).join(""))) });
    } else if (/\.rels$/.test(name)) { // внешние ссылки (mailto:, адреса) тоже могут содержать данные
      const xml = await zip.file(name).async("string"), tags = [...xml.matchAll(/<Relationship\b[^>]*>/g)].filter(m => /TargetMode="External"/.test(m[0]) && /\bTarget="[^"]*"/.test(m[0]));
      if (tags.length) items.push({ name, xml, kind: "rel", tags, texts: tags.map(m => xmlUnesc(/\bTarget="([^"]*)"/.exec(m[0])[1])) });
    }
  }
  const all = items.flatMap(i => i.texts), res = fn(all);
  if (res.length !== all.length) throw new Error("внутренняя ошибка: число абзацев изменилось");
  let k = 0, changed = 0;
  for (const it of items) {
    const edits = []; // [от, до, новый текст]
    if (it.kind === "xml") for (const g of it.groups) {
      const nt = res[k], old = all[k]; k++;
      if (nt === old) continue; changed++;
      const f = g[0];
      edits.push([f.open, f.end, "<" + f.name + ' xml:space="preserve"' + f.attrs.replace(/\s*xml:space="[^"]*"/, "") + ">" + xmlEsc(nt) + "</" + f.name + ">"]);
      for (const r of g.slice(1)) edits.push([r.ca, r.cb, ""]);
    } else it.tags.forEach(m => {
      const nt = res[k], old = all[k]; k++;
      if (nt !== old) { changed++; edits.push([m.index, m.index + m[0].length, m[0].replace(/\bTarget="[^"]*"/, 'Target="' + attrEsc(nt) + '"')]) }
    });
    let xml = it.xml;
    for (const [a, b, t] of edits.sort((x, y) => y[0] - x[0])) xml = xml.slice(0, a) + t + xml.slice(b);
    if (clean && it.kind === "xml") xml = xml.replace(/\b(w:author|w15:author)="[^"]*"/g, '$1="Автор"').replace(/\bw:initials="[^"]*"/g, 'w:initials=""').replace(/<w15:person\b[\s\S]*?<\/w15:person>/g, "");
    if (xml !== it.xml) zip.file(it.name, xml);
  }
  if (clean) { // свойства файла: автор, редактор, название, организация
    const blank = async (p, tags) => { const f = zip.file(p); if (f) zip.file(p, (await f.async("string")).replace(new RegExp("(<(" + tags + ")(?:\\s[^>]*)?>)[\\s\\S]*?(</\\2>)", "g"), "$1$3")) };
    await blank("docProps/core.xml", "dc:creator|cp:lastModifiedBy|dc:title|dc:subject|dc:description|cp:keywords|cp:category");
    await blank("docProps/app.xml", "Company|Manager");
    const emb = names.filter(n => /^word\/(embeddings|activeX)\//.test(n)), img = names.filter(n => /^word\/media\//.test(n));
    if (emb.length) warn.push("в документе есть вложенные файлы (" + emb.length + "): они не обработаны, удалите их или проверьте отдельно");
    if (img.length) warn.push("в документе есть картинки (" + img.length + "): текст на них не обрабатывается");
  }
  return { data: await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), changed, warn };
}
async function anonymizeDocx(buf, dict = []) {
  let r, preview = [];
  const o = await mapDocx(buf, texts => { r = E.depersonalize(E.joinCells(texts), dict); const out = E.splitCells(r.text, texts.length);
    preview = texts.map((t, i) => [t, out[i]]).filter(([a, b]) => a !== b).slice(0, 6); return out }, { clean: true });
  return { data: o.data, map: r.map, counts: r.counts, warn: [...r.warn.filter(w => !/словарь пуст/.test(w)), ...o.warn], changed: o.changed, preview };
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
