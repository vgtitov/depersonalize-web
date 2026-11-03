// Независимое ревью, 03.11: P1 1-2 и P2 3-9.
const assert = require("assert");
const Z = require("../vendor/jszip.min.js");
const O = require("../office.js"), E = require("../engine.js");
const W = '<w:document xmlns:w="w"><w:body>', WE = "</w:body></w:document>";
async function docx(files) { const d = new Z(); d.file("[Content_Types].xml", "<Types/>"); for (const [k, v] of Object.entries(files)) d.file(k, v); return d.generateAsync({ type: "uint8array" }) }
async function parts(data) { const z = await Z.loadAsync(data); const o = {}; for (const n of Object.keys(z.files)) if (!z.files[n].dir) o[n] = await z.file(n).async("string"); return o }
const run = async (files, dict = []) => { const r = await O.anonymizeDocx(await docx(files), dict); return { r, p: await parts(r.data) } };
const T = s => "<w:r><w:t>" + s + "</w:t></w:r>";
(async () => {
  // P1-1: разрыв строки, табуляция, w:cr не склеивают слова и не теряются
  for (const sep of ["<w:br/>", "<w:br w:type=\"textWrapping\"/>", "<w:tab/>", "<w:cr/>"]) {
    const { p } = await run({ "word/document.xml": W + "<w:p>" + T("Генеральный директор") + sep + T("Сидоров Иван Петрович") + sep + T("тел. 123") + "</w:p>" + WE });
    const x = p["word/document.xml"];
    assert.ok(!x.includes("Сидоров") && !x.includes("Петрович"), "имя осталось после " + sep + ": " + x);
    assert.ok(x.includes("Генеральный директор") && x.includes(sep), "разрыв потерян: " + x);
    assert.ok(x.indexOf("Генеральный") < x.indexOf(sep) && x.indexOf(sep) < x.indexOf("[ЧЕЛОВЕК_1]"), "порядок нарушен: " + x);
  }
  { // разрыв внутри одного w:r, несколько w:t вокруг
    const { p } = await run({ "word/document.xml": W + "<w:p><w:r><w:t>Исполнитель</w:t><w:br/><w:t>Петрова Анна Ивановна</w:t></w:r></w:p>" + WE });
    assert.ok(!p["word/document.xml"].includes("Петрова") && p["word/document.xml"].includes("<w:br/>"), p["word/document.xml"]);
  }
  { // разрыв не меняет абзац без находок; tab stops в pPr не принимаются за табуляцию
    const xml = W + '<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs></w:pPr>' + T("Обычный") + "<w:br/>" + T("текст") + "</w:p>" + WE;
    const { p } = await run({ "word/document.xml": xml });
    assert.strictEqual(p["word/document.xml"], xml);
  }

  // P1-2: имя безопасной копии не содержит имени исходного файла
  for (const kind of ["docx", "xlsx", "csv"]) { const n = O.safeName(kind); assert.ok(/безопасная-копия\.\w+$/.test(n) && n.endsWith("." + kind), n); assert.ok(!/Сидоров/.test(n)) }
  assert.strictEqual(O.safeName("docx", "Характеристика Сидоров Иван Петрович.docx"), O.safeName("docx"));

  // P2-3: имя листа и строка заголовков
  const rows = [["Сидоров Иван Петрович", "Янв"], [5, 7]];
  const cols = O.detectColumns(rows);
  const a = O.anonymizeTable(rows, cols);
  assert.ok(!a.rows[0].join("|").includes("Сидоров"), "заголовок не обработан: " + a.rows[0]);
  const xl = await O.writeXlsx(a.rows, O.SAFE_SHEET), px = await parts(xl);
  assert.ok(!px["xl/workbook.xml"].includes("Сидоров") && O.SAFE_SHEET === "Лист1");

  // P2-4: текстовая колонка по умолчанию просматривается; числа и даты остаются
  const r4 = [["Заметка", "Сумма", "Дата", "Код"], ["Звонил Олег Кравцов, тел 80291234567", 10, "2026-10-01", "A-17"], ["ок", 20, "2026-10-02", "A-18"]];
  const c4 = O.detectColumns(r4);
  assert.strictEqual(c4[0].mode, "scan"); assert.strictEqual(c4[1].mode, "keep"); assert.strictEqual(c4[2].mode, "keep");
  const a4 = O.anonymizeTable(r4, c4); assert.ok(!a4.rows[1][0].includes("80291234567") && !a4.rows[1][0].includes("Кравцов"), a4.rows[1][0]);

  // P2-5: фамилия без типичного окончания перед «Имя Отчество»
  for (const s of ["Кравец", "Жук", "Бондарь", "Мельник", "Новик"]) { const t = E.depersonalize(s + " Олег Иванович подписал", []).text; assert.strictEqual(t, "[ЧЕЛОВЕК_1] подписал", t) }
  for (const s of ["Директор", "Уважаемый", "Менеджер", "Покупатель", "Сотрудник"]) { const t = E.depersonalize(s + " Олег Иванович подписал", []).text; assert.ok(t.startsWith(s + " [ЧЕЛОВЕК_1]"), t) }

  // P2-6: паспорт, личный номер, дата рождения, адрес находятся (или явно описаны как ограничение в README)
  const cases = [["паспорт MP1234567 выдан", "ПАСПОРТ"], ["Паспорт серии KH 7654321", "ПАСПОРТ"], ["личный номер 3010190A001PB1", "ИДНОМЕР"], ["идентификационный номер: 3010190A001PB1.", "ИДНОМЕР"],
    ["Дата рождения: 12.03.1985", "РОЖДЕНИЕ"], ["родился 5 марта 1990 г.", "РОЖДЕНИЕ"], ["проживает: г. Минск, ул. Ленина, 15, кв. 3", "АДРЕС"], ["пр-т Независимости 44", "АДРЕС"]];
  for (const [t, label] of cases) { const o = E.depersonalize(t, []).text; assert.ok(o.includes("[" + label + "_1]"), t + " -> " + o) }
  assert.ok(!/MP1234567|3010190A001PB1|Ленина|12\.03\.1985/.test(E.depersonalize("MP1234567 3010190A001PB1 ул. Ленина, 15 дата рождения 12.03.1985", []).text));
  assert.ok(E.depersonalize("улица широкая, ночью тихо", []).text === "улица широкая, ночью тихо");

  // P2-7: вставки altChunk и миниатюра
  { const { r, p } = await run({ "word/document.xml": W + "<w:p>" + T("x") + "</w:p>" + WE, "word/afchunk.mht": "Иван Петрович Сидоров", "docProps/thumbnail.jpeg": "JPEGDATA",
      "_rels/.rels": '<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId9" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail" Target="docProps/thumbnail.jpeg"/></Relationships>' });
    assert.ok(r.warn.some(w => /вставк|altChunk|HTML|RTF/i.test(w)), r.warn.join("|"));
    assert.ok(!("docProps/thumbnail.jpeg" in p), "миниатюра осталась");
    assert.ok(!p["_rels/.rels"].includes("thumbnail") && p["_rels/.rels"].includes("officeDocument"), p["_rels/.rels"]) }

  // P2-8: ключ привязан к обработке, при возврате видно, сколько кодов найдено и сколько чужих
  const m1 = E.depersonalize("Иван Петрович Сидоров", []).map;
  assert.ok(m1.meta && /^[0-9a-f]{6}$/.test(m1.meta.id) && m1.meta.created);
  const back = E.mapFromJson(E.mapToJson(m1)); assert.strictEqual(back.meta.id, m1.meta.id);
  const st = E.restoreStats("[ЧЕЛОВЕК_1] и [ЧЕЛОВЕК_2] и ТЕЛЕФОН_007", back);
  assert.deepStrictEqual([st.found, st.matched, st.unknown.length], [3, 1, 2]);
  const m2 = E.depersonalize("Иван Петрович Сидоров", []).map; assert.notStrictEqual(m2.meta.id, m1.meta.id);
  assert.ok(O.anonymizeTable([["ФИО"], ["Иван Петрович Сидоров"]], O.detectColumns([["ФИО"], ["Иван Петрович Сидоров"]])).map.meta.id);

  // P2-9: слова короче трёх букв из словаря
  assert.strictEqual(E.depersonalize("Ли Вэй и Ян были. Придёт ли он, Ян?", ["Ли", "Ян"]).text, "[СЛОВО_1] Вэй и [СЛОВО_2] были. Придёт ли он, [СЛОВО_2]?");
  assert.ok(E.depersonalize("текст", ["Я"]).warn.some(w => /короче/.test(w)));
  console.log("ревью 03.11 (P1-2, P2-7): ok");
})().catch(e => { console.error(e); process.exit(1) });
