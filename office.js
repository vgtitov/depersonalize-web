// Таблицы (CSV, XLSX) и документы (DOCX) локально. Нужен JSZip (vendor/jszip.min.js), движок engine.js. Без сети.
const Z = typeof JSZip !== "undefined" ? JSZip : require("./vendor/jszip.min.js");
const E = typeof depersonalize !== "undefined" ? { depersonalize, depersonalizeMany, restore, restoreStats } : require("./engine.js");

// ---------- общее ----------
const xmlUnesc = s => s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(+n)).replace(/&#x([0-9a-f]+);/gi, (_, n) => String.fromCodePoint(parseInt(n, 16))).replace(/&amp;/g, "&");
const xmlEsc = s => String(s).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
// Ключ вида ЧЕЛОВЕК_001 в тексте ответа ИИ → исходное значение по карте.
function restoreKeys(text, map) {
  return E.restore(text, map).replace(/(?<![\wА-Яа-яЁё])[A-ZА-ЯЁ]+_\d{3,}(?![\wА-Яа-яЁё])/g, k => map.has(k) ? map.get(k) : k);
}

// Имя безопасной копии не берётся из исходного файла: в имени файла бывают ФИО, а имя видит и ИИ при загрузке.
const SAFE_SHEET = "Лист1";
const safeName = kind => (kind === "docx" ? "документ" : "таблица") + "-безопасная-копия." + kind;

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
// Что за формат ячейки: date, datetime, time, elapsed (часы больше суток) или null.
function fmtKind(id, codes) {
  if (id >= 14 && id <= 17) return "date";
  if (id === 22) return "datetime";
  if (id === 46) return "elapsed";
  if ((id >= 18 && id <= 21) || id === 45 || id === 47) return "time";
  const c = codes[id]; if (!c) return null;
  if (/\[(?:h+|m+|s+)\]/i.test(c)) return "elapsed";
  const x = c.replace(/"[^"]*"|\[[^\]]*\]|\\./g, ""), t = /[hs]/i.test(x), d = /[dy]/i.test(x) || (!t && /m/i.test(x));
  return d && t ? "datetime" : d ? "date" : t ? "time" : null;
}
function serialToText(n, kind, d1904) {
  const p = (x, l = 2) => String(x).padStart(l, "0");
  if (kind === "elapsed") { const m = Math.round(n * 1440); return Math.floor(m / 60) + ":" + p(m % 60) }
  const ms = Math.round((n + (d1904 ? 1462 : 0) - 25569) * 86400000), d = new Date(ms), hm = p(d.getUTCHours()) + ":" + p(d.getUTCMinutes());
  if (kind === "time") return hm;
  const day = d.getUTCFullYear() + "-" + p(d.getUTCMonth() + 1) + "-" + p(d.getUTCDate());
  return n % 1 === 0 && kind === "date" ? day : day + " " + hm;
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
  const dateXf = [...xfs.matchAll(/<xf\b[^>]*>/g)].map(m => { const f = /numFmtId="(\d+)"/.exec(m[0]); return f ? fmtKind(+f[1], codes) : null }), d1904 = /<workbookPr\b[^>]*\bdate1904="(?:1|true)"/i.test(wb);
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
        else if (v != null && v !== "" && Number.isFinite(+v)) val = sidx != null && dateXf[+sidx] ? serialToText(+v, dateXf[+sidx], d1904) : +v;
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
const B = "(?<![а-яёa-z])"; // начало слова: «расчет» не «счёт», «показатель» не «тел»
const HINTS = [["ЗНАЧЕНИЕ", new RegExp(`${B}(?:парол|password|pwd|токен|token|секрет|secret|pin\\b|пин-?код)`, "i"), "drop"], ["ЗНАЧЕНИЕ", new RegExp(`${B}(?:инн|унп|паспорт|iban|карт[аы]|адрес|сч[её]т)`, "i"), "key"], ["EMAIL", new RegExp(`${B}(?:e-?mail|почт)`, "i")], ["ТЕЛЕФОН", new RegExp(`${B}(?:тел(?![а-яё])|телеф|phone|моб)`, "i")], ["ЧЕЛОВЕК", new RegExp(`${B}(?:фио|фамил|имя(?![а-яё])|сотрудник|менеджер|ответствен)`, "i")], ["ОРГ", new RegExp(`${B}(?:организ|компан|контрагент|клиент|поставщик|юрлиц)`, "i")]];
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
    else { const h = HINTS.find(([, re]) => re.test(name)); if (h) { kind = h[0]; mode = h[2] || "key" }
      else if (sample.some(v => !/^[\d\s.,:;\/+%$€£₽-]*$/.test(v) && /[A-Za-zА-Яа-яЁё]{3}/.test(v)) && !sample.every(v => /^\p{L}{0,3}[-_ ]?\d+$/u.test(v))) mode = "scan" } // текстовая колонка: имя или телефон в заметке уйдут, поэтому по умолчанию ищем внутри; числа, даты и короткие коды остаются
    return { index: i, name, kind, mode, preview: vals.slice(0, 3) };
  });
}
// Применение выбора: keep / key (ключи вместо значений) / drop (убрать колонку) / scan (найти и скрыть внутри текста ячеек).
function anonymizeTable(rows, cols, { hasHeader = true, dict = [] } = {}) {
  const map = new Map(), reg = new Map(), n = {}, counts = {}, warn = [];
  const keyOf = (kind, v) => { const num = typeof v === "number", norm = String(v).trim().replace(/\s+/g, " ").toLowerCase(), rk = kind + "|" + (num ? "#" : "") + norm;
    if (!reg.has(rk)) { n[kind] = (n[kind] || 0) + 1; counts[kind] = n[kind]; const k = kind + "_" + String(n[kind]).padStart(3, "0"); reg.set(rk, k); map.set(k, String(v).trim()); if (num) map.set("#число:" + k, "1") } // число помечаем, чтобы вернуть числом
    return reg.get(rk) };
  const out = rows.map(r => r.slice()), start = hasHeader ? 1 : 0, scan = [];
  if (hasHeader && out[0]) out[0].forEach((v, c) => { if (typeof v === "string" && v.trim()) scan.push([0, c]) }); // заголовки колонок тоже бывают ФИО
  for (const c of cols) {
    if (c.mode === "key") for (let r = start; r < out.length; r++) { const v = out[r][c.index]; if (v !== "" && v != null) out[r][c.index] = keyOf(c.kind, v) }
    if (c.mode === "scan") for (let r = start; r < out.length; r++) if (out[r][c.index] !== "" && out[r][c.index] != null) scan.push([r, c.index]);
    if (c.mode === "keep") for (let r = start; r < out.length; r++) if (/^[A-ZА-ЯЁ]+_\d{3,}$/.test(String(out[r][c.index] ?? ""))) { warn.push("в колонке «" + (rows[0] && hasHeader ? rows[0][c.index] : c.index + 1) + "» есть значения, похожие на коды (ЧЕЛОВЕК_001): при возврате они могут замениться"); break }
  }
  if (scan.length) { // все ячейки режима «внутри текста» одной пачкой: общая нумерация, переносы строк не мешают
    const res = E.depersonalizeMany(scan.map(([r, c]) => String(out[r][c])), dict);
    scan.forEach(([r, c], i) => { out[r][c] = res.texts[i] });
    for (const [t, v] of res.map) map.set(t, v);
    for (const [k, v] of Object.entries(res.counts)) counts[k] = (counts[k] || 0) + v;
  }
  map.meta = { id: Math.random().toString(16).slice(2, 8).padEnd(6, "0"), created: new Date().toISOString() };
  const drop = new Set(cols.filter(c => c.mode === "drop").map(c => c.index));
  return { rows: out.map(r => r.filter((_, i) => !drop.has(i))), map, counts, warn };
}
function restoreRows(rows, map) {
  return rows.map(r => r.map(v => {
    if (typeof v !== "string") return v;
    if (map.has("#число:" + v) && Number.isFinite(+map.get(v))) return Number(map.get(v)); // было числом
    return restoreKeys(v, map);
  }));
}

// ---------- DOCX ----------
const PARA = /^word\/(?!_rels\/|charts\/|diagrams\/).*\.xml$/;                                  // части с абзацами
const FLAT = /^(?:customXml\/(?!_rels\/).*\.xml|word\/(?:charts|diagrams)\/(?!_rels\/).*\.xml|docProps\/(?:custom|app)\.xml)$/; // части, где весь текст между тегами
const attrEsc = s => xmlEsc(s).replace(/"/g, "&quot;");
// Группы текста части: живой текст абзаца (с вставками w:ins) — одна группа; удалённый текст (w:delText) — своя группа на каждое w:del; w:instrText — по одному.
function textGroups(xml) {
  const re = /<w:br\b[^>]*\/>|<w:cr\s*\/>|<w:tab\s*\/>|<w:ptab\b[^>]*\/>|<(w|a):p[ >]|<\/(?:w|a):p>|<w:(?:del|moveFrom|moveTo)\b(?![^>]*\/>)[^>]*>|<\/w:(?:del|moveFrom|moveTo)>|<w:fldChar\b[^>]*>|<w:ins\b(?![^>]*\/>)[^>]*>|<\/w:ins>|<(w:t|w:delText|a:t|w:instrText)\b(?![^>]*\/>)([^>]*)>([\s\S]*?)<\/\2>/g;
  const paras = [], dels = []; let pid = 0, did = 0, ins = 0, fid = 0; const flds = []; let seg = 0; const groups = new Map();
  for (const m of xml.matchAll(re)) {
    const t = m[0];
    if (/^<w:(?:br|cr|tab|ptab)\b/.test(t)) { // разрыв строки и табуляция остаются на месте и разделяют слова
      const dl = dels[dels.length - 1], pr = paras[paras.length - 1] || 0, key = dl && dl.mv ? "m" + dl.id : dl ? "d" + dl.id : "p" + pr;
      if (!groups.has(key)) groups.set(key, { runs: [], items: [], chg: false });
      groups.get(key).items.push({ sep: /^<w:(?:tab|ptab)\b/.test(t) ? "\t" : "\n" }); continue;
    }
    if (m[1]) paras.push(++pid);
    else if (t === "</w:p>" || t === "</a:p>") paras.pop();
    else if (/^<w:(?:del|moveFrom|moveTo)[\s>]/.test(t)) dels.push({ id: ++did, mv: !/^<w:del[\s>]/.test(t) });
    else if (/^<\/w:(?:del|moveFrom|moveTo)>/.test(t)) dels.pop();
    else if (t.startsWith("<w:fldChar")) { const ty = /fldCharType="(\w+)"/.exec(t); if (ty && ty[1] === "begin") { flds.push(++fid); seg++ } else if (ty && ty[1] === "end") { flds.pop(); seg++ } }
    else if (/^<w:ins[\s>]/.test(t)) ins++;
    else if (t === "</w:ins>") ins--;
    else {
      const name = m[2], dl = dels[dels.length - 1], pr = paras[paras.length - 1] || 0;
      const key = dl && dl.mv ? "m" + dl.id : name === "w:instrText" ? (flds.length ? "f" + flds[flds.length - 1] + "." + seg : "i" + pr) : name === "w:delText" ? "d" + (dl ? dl.id : 0) : "p" + pr; // код поля может быть разбит на части: все части абзаца вместе
      const run = { open: m.index, ca: m.index + 1 + name.length + m[3].length + 1, name, attrs: m[3], raw: m[4] };
      run.cb = run.ca + m[4].length; run.end = run.cb + name.length + 3;
      if (!groups.has(key)) groups.set(key, { runs: [], items: [], chg: false });
      const g = groups.get(key); g.runs.push(run); g.items.push(run); if (ins > 0 || (dl && dl.mv)) g.chg = true;
    }
  }
  return [...groups.values()];
}
const groupText = g => g.items.map(i => i.sep || xmlUnesc(i.raw)).join("");
// Новый текст группы раскладывается обратно по фрагментам: разрывы и табуляции стоят на прежних местах. Не сошлось число разрывов: весь текст в первый фрагмент, разрывы остаются после него.
function groupEdits(g, nt) {
  const seps = g.items.filter(i => i.sep).length, pieces = nt.split(/[\n\t]/);
  const wr = (r, t) => "<" + r.name + ' xml:space="preserve"' + r.attrs.replace(/\s*xml:space=(?:"[^"]*"|'[^']*')/, "") + ">" + xmlEsc(t) + "</" + r.name + ">";
  const edits = [];
  if (pieces.length !== seps + 1 || !g.runs.length) { const f = g.runs[0]; if (f) { edits.push([f.open, f.end, wr(f, nt.replace(/[\n\t]/g, " "))]); for (const r of g.runs.slice(1)) edits.push([r.ca, r.cb, ""]) } return edits }
  let si = 0, first = true; // фрагменты между разрывами: текст куда в первый w:t сегмента, остальные пустые
  for (const it of g.items) {
    if (it.sep) { si++; first = true; continue }
    if (first) { edits.push([it.open, it.end, wr(it, pieces[si])]); first = false } else edits.push([it.ca, it.cb, ""]);
  }
  return edits;
}
// Атрибуты с текстом: подписи и подсказки, имена файлов и закладок, поля содержимого, переменные документа, коды простых полей.
const SLOT_ANY = /\b((?:w:|o:)?(?:descr|title|alt|tooltip|instr))="([^"]*)"/g, SLOT_TAG = /^<(?:[\w]+:)?(?:cNvPr|docPr|docVar|docPart|alias|tag|attr|smartTagPr|placeholder)\b/, SLOT_NAME = /\b((?:w:)?(?:name|val))="([^"]*)"/g;
function attrSlots(xml, all = false) {
  const out = [], seen = new Set();
  for (const m of xml.matchAll(/<[\w:]+\b[^>]*>/g)) {
    const tag = m[0], rx = all ? [/\b([\w:]+)="([^"]*)"/g] : SLOT_TAG.test(tag) ? [SLOT_ANY, SLOT_NAME] : [SLOT_ANY];
    for (const r of rx) for (const a of tag.matchAll(r)) {
      if (!a[2].trim() || /^xmlns/.test(a[1]) || (all && (/^\{.*\}$/.test(a[2]) || /^(?:ds:)?(?:uri|itemID)$/.test(a[1]) || /^https?:\/\/(?:schemas\.|www\.w3\.org|purl\.org)/.test(a[2]) || !/[A-Za-zА-Яа-яЁё\d]/.test(a[2])))) continue;
      const pos = m.index + a.index + a[0].indexOf('"') + 1;
      if (!seen.has(pos)) { seen.add(pos); out.push({ a: pos, b: pos + a[2].length, raw: a[2] }) }
    }
  }
  return out;
}
// Прогоняет тексты документа через fn(массив строк) → массив строк (общая нумерация меток по всем частям).
async function mapDocx(buf, fn, { clean = false } = {}) {
  const zip = await Z.loadAsync(buf);
  const doc = zip.file("word/document.xml");
  if (!doc || !/<w:(?:document|body)\b/.test(await doc.async("string"))) throw new Error("это не файл Word (.docx) или в нём нестандартная разметка");
  const items = [], warn = [], names = Object.keys(zip.files).filter(n => !zip.files[n].dir);
  for (const name of names) {
    if (PARA.test(name)) {
      const xml = await zip.file(name).async("string"), groups = textGroups(xml), slots = attrSlots(xml);
      items.push({ name, xml, kind: "xml", groups, slots, texts: [...groups.map(groupText), ...slots.map(x => xmlUnesc(x.raw))] });
    } else if (FLAT.test(name)) {
      const xml = await zip.file(name).async("string"), app = name === "docProps/app.xml", chart = /\/charts\//.test(name);
      const nodes = [...xml.matchAll(app ? /<vt:lpstr>([^<>]*[^<>\s][^<>]*)</g : /(?<=>)([^<>]*[^<>\s][^<>]*)(?=<)/g)].map(m => ({ a: m.index + m[0].indexOf(m[1]), b: m.index + m[0].indexOf(m[1]) + m[1].length, raw: m[1] })).filter(x => !(chart && /^[\d\s.,+-]+$/.test(x.raw)));
      if (!app) { for (const x of attrSlots(xml, true)) nodes.push(x); for (const m of xml.matchAll(/<!--([\s\S]*?)-->/g)) if (m[1].trim()) nodes.push({ a: m.index + 4, b: m.index + 4 + m[1].length, raw: m[1] }) }
      if (nodes.length) items.push({ name, xml, kind: "flat", nodes, texts: nodes.map(x => xmlUnesc(x.raw)) });
    } else if (/\.rels$/.test(name)) { // внешние ссылки (mailto:, адреса, пути к файлам) тоже могут содержать данные
      const xml = await zip.file(name).async("string"), tags = [...xml.matchAll(/<Relationship\b[^>]*>/g)].filter(m => /TargetMode="External"/.test(m[0]) && /\bTarget="[^"]*"/.test(m[0]));
      if (tags.length) items.push({ name, xml, kind: "rel", tags, texts: tags.map(m => xmlUnesc(/\bTarget="([^"]*)"/.exec(m[0])[1])) });
    }
  }
  const all = items.flatMap(i => i.texts), res = fn(all);
  if (res.length !== all.length) throw new Error("внутренняя ошибка: число абзацев изменилось");
  let k = 0, changed = 0, merged = false;
  for (const it of items) {
    const edits = []; // [от, до, новый текст]
    if (it.kind === "xml") {
      for (const g of it.groups) {
        const nt = res[k], old = all[k]; k++;
        if (nt === old) continue; changed++; if (g.chg) merged = true;
        edits.push(...groupEdits(g, nt));
      }
      for (const x of it.slots) { const nt = res[k], old = all[k]; k++; if (nt !== old) { changed++; edits.push([x.a, x.b, attrEsc(nt)]) } }
    } else if (it.kind === "flat") {
      for (const x of it.nodes) { const nt = res[k], old = all[k]; k++; if (nt !== old) { changed++; edits.push([x.a, x.b, xmlEsc(nt)]) } }
    } else it.tags.forEach(m => {
      const nt = res[k], old = all[k]; k++;
      if (nt !== old) { changed++; edits.push([m.index, m.index + m[0].length, m[0].replace(/\bTarget="[^"]*"/, 'Target="' + attrEsc(nt) + '"')]) }
    });
    let xml = it.xml;
    for (const [a, b, t] of edits.sort((x, y) => y[0] - x[0])) xml = xml.slice(0, a) + t + xml.slice(b);
    if (clean && it.kind === "xml") xml = xml.replace(/\b(w:author|w15:author)="[^"]*"/g, '$1="Автор"').replace(/\bw:initials="[^"]*"/g, 'w:initials=""').replace(/<w15:person\b[\s\S]*?<\/w15:person>/g, "");
    if (xml !== it.xml) zip.file(it.name, xml);
  }
  const bm = await Promise.all(names.filter(n => PARA.test(n)).map(async n => /<w:bookmarkStart\b[^>]*w:name="(?!_)/.test(await zip.file(n).async("string"))));
  if (bm.some(Boolean)) warn.push("в документе есть закладки с именами: имена не меняются, проверьте их вручную");
  if (merged) warn.push("в документе есть исправления (рецензирование): изменённые места сведены в обычный текст");
  if (clean) { // свойства файла: автор, редактор, название, организация
    const blank = async (p, tags) => { const f = zip.file(p); if (f) zip.file(p, (await f.async("string")).replace(new RegExp("(<(" + tags + ")(?:\\s[^>]*)?>)[\\s\\S]*?(</\\2>)", "g"), "$1$3")) };
    await blank("docProps/core.xml", "dc:creator|cp:lastModifiedBy|dc:title|dc:subject|dc:description|cp:keywords|cp:category");
    await blank("docProps/app.xml", "Company|Manager");
    const thumbs = names.filter(n => /^docProps\/thumbnail\./i.test(n));
    for (const n of thumbs) zip.remove(n); // миниатюра первой страницы показывает текст документа
    if (thumbs.length && zip.file("_rels/.rels")) zip.file("_rels/.rels", (await zip.file("_rels/.rels").async("string")).replace(/<Relationship\b[^>]*\/thumbnail"[^>]*\/>|<Relationship\b[^>]*Target="docProps\/thumbnail[^>]*\/>/g, ""));
    const alt = names.filter(n => /^word\/(?!_rels\/).*\.(?:mht|mhtml|htm|html|rtf)$/i.test(n));
    if (alt.length) warn.push("в документе есть вставки HTML или RTF (" + alt.length + "): их текст не обрабатывается, удалите вставки или проверьте отдельно");
    const emb = names.filter(n => /^word\/(embeddings|activeX)\//.test(n)), img = names.filter(n => /^word\/media\//.test(n));
    if (emb.length) warn.push("в документе есть вложенные файлы (" + emb.length + "): они не обработаны, удалите их или проверьте отдельно");
    if (img.length) warn.push("в документе есть картинки (" + img.length + "): текст на них не обрабатывается");
  }
  return { data: await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), changed, warn };
}
async function anonymizeDocx(buf, dict = []) {
  let r, preview = [];
  const o = await mapDocx(buf, texts => { r = E.depersonalizeMany(texts, dict); const out = r.texts;
    preview = texts.map((t, i) => [t, out[i]]).filter(([a, b]) => a !== b).slice(0, 6); return out }, { clean: true });
  return { data: o.data, map: r.map, counts: r.counts, warn: [...r.warn.filter(w => !/словарь пуст/.test(w)), ...o.warn], changed: o.changed, preview };
}
async function restoreDocx(buf, map, stats) { // stats (необязательно): сюда пишется, сколько кодов найдено и сколько не из этого ключа
  return (await mapDocx(buf, texts => { if (stats) { const u = new Set(); let f = 0, m = 0; for (const t of texts) { const q = E.restoreStats(t, map); f += q.found; m += q.matched; q.unknown.forEach(x => u.add(x)) } Object.assign(stats, { found: f, matched: m, unknown: [...u] }) } return texts.map(t => restoreKeys(t, map)) })).data }

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

const api = { safeName, SAFE_SHEET, makeDocx, restoreKeys, parseCsv, toCsv, readXlsx, writeXlsx, detectColumns, anonymizeTable, restoreRows, anonymizeDocx, restoreDocx, mapDocx, KINDS };
if (typeof module !== "undefined") module.exports = api;
