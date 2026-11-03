// Excel, CSV и Word: данные не остаются в выходе, возврат восстанавливает исходное. Файлы создаются здесь же.
const assert = require("assert");
const Z = require("../vendor/jszip.min.js");
const O = require("../office.js"), E = require("../engine.js");
const NAMES = ["Сидоров", "Иванов", "Петрова", "Рога и Копыта", "123-45-67", "ivan@example.test"];
async function allText(data) { const z = await Z.loadAsync(data); let s = ""; for (const n of Object.keys(z.files)) if (!z.files[n].dir) s += await z.file(n).async("string"); return s }
const noLeak = s => NAMES.forEach(n => assert.ok(!s.includes(n), "осталось: " + n));

(async () => {
  // CSV
  const csv = 'ФИО;Телефон;Заметка\r\n"Сидоров Иван Петрович";+375 29 123-45-67;"строка, с ""кавычками"""\r\nИванов Пётр;+375 29 123-45-67;ок\r\n';
  const p = O.parseCsv(csv); assert.strictEqual(p.delim, ";"); assert.strictEqual(p.rows[1][2], 'строка, с "кавычками"');
  assert.deepStrictEqual(O.parseCsv(O.toCsv(p.rows, ";")).rows, p.rows);

  // Excel: колонки, ключи, возврат
  const rows = [["ФИО", "Телефон", "Email", "Организация", "Сумма", "Комментарий", "Личное"],
    ["Сидоров Иван Петрович", "+375 29 123-45-67", "ivan@example.test", "ООО «Рога и Копыта»", 1500.5, "звонил Иванов Пётр", "секрет-1"],
    ["Иванов Пётр Сергеевич", "+375 44 765-43-21", "ivan@example.test", "ООО «Рога и Копыта»", 20, "ок", "секрет-2"],
    ["Сидоров Иван Петрович", "+375 29 123-45-67", "p@example.test", "ЗАО «Лютик»", 7, "", "секрет-3"]];
  const cols = O.detectColumns(rows);
  assert.deepStrictEqual(cols.slice(0, 4).map(c => [c.kind, c.mode]), [["ЧЕЛОВЕК", "key"], ["ТЕЛЕФОН", "key"], ["EMAIL", "key"], ["ОРГ", "key"]]);
  assert.strictEqual(cols[4].mode, "keep");
  cols[5].mode = "scan"; cols[6].mode = "drop";
  const a = O.anonymizeTable(rows, cols);
  assert.strictEqual(a.rows[1][0], "ЧЕЛОВЕК_001"); assert.strictEqual(a.rows[3][0], "ЧЕЛОВЕК_001"); assert.strictEqual(a.rows[2][0], "ЧЕЛОВЕК_002"); // связи сохраняются
  assert.strictEqual(a.rows[1][2], a.rows[2][2]);                       // одна почта — один ключ
  assert.strictEqual(a.rows[0].length, 6); assert.strictEqual(a.rows[1][4], 1500.5);
  const xl = await O.writeXlsx(a.rows);
  noLeak(await allText(xl)); assert.ok(!(await allText(xl)).includes("секрет-"));
  const back = (await O.readXlsx(xl))[0].rows;
  assert.deepStrictEqual(back, a.rows);
  const restored = O.restoreRows(back, a.map);
  assert.strictEqual(restored[1][0], "Сидоров Иван Петрович"); assert.strictEqual(restored[2][3], "ООО «Рога и Копыта»");
  assert.strictEqual(restored[1][5], "звонил Иванов Пётр");
  // ответ ИИ — обычный текст с ключами
  assert.strictEqual(O.restoreKeys("Самый крупный клиент: ЧЕЛОВЕК_001, ЧЕЛОВЕК_0012 не трогаем", a.map), "Самый крупный клиент: Сидоров Иван Петрович, ЧЕЛОВЕК_0012 не трогаем");

  // Excel «как из Excel»: общие строки и типы ячеек
  const z = new Z();
  z.file("xl/workbook.xml", '<workbook xmlns:r="x"><sheets><sheet name="Данные" sheetId="1" r:id="rId1"/></sheets></workbook>');
  z.file("xl/_rels/workbook.xml.rels", '<Relationships><Relationship Id="rId1" Type="t" Target="worksheets/sheet1.xml"/></Relationships>');
  z.file("xl/sharedStrings.xml", '<sst><si><t>ФИО</t></si><si><r><t>Сидоров </t></r><r><t>Иван</t></r></si><si><t>a &amp; b</t></si></sst>');
  z.file("xl/worksheets/sheet1.xml", '<worksheet><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="C1" t="s"><v>2</v></c></row><row r="2"><c r="A2" t="s"><v>1</v></c><c r="B2"><v>42</v></c><c r="C2" t="b"><v>1</v></c></row></sheetData></worksheet>');
  const sh = (await O.readXlsx(await z.generateAsync({ type: "uint8array" })))[0];
  assert.deepStrictEqual(sh.rows, [["ФИО", "", "a & b"], ["Сидоров Иван", 42, "TRUE"]]);
  await assert.rejects(() => O.readXlsx(new Uint8Array([1, 2, 3])));

  // Word: имена разбиты по фрагментам, таблица, колонтитул, комментарий, автор
  const w = (t) => '<w:p><w:r><w:t xml:space="preserve">' + t[0] + '</w:t></w:r>' + (t[1] ? '<w:r><w:rPr><w:b/></w:rPr><w:t>' + t[1] + '</w:t></w:r>' : '') + '</w:p>';
  const d = new Z();
  d.file("[Content_Types].xml", "<Types/>");
  d.file("word/document.xml", '<w:document xmlns:w="w"><w:body>' + w(["Договор с ООО «Рога ", "и Копыта» подписал Иван Петрович Сидоров"]) + '<w:tbl><w:tr><w:tc>' + w(["Телефон +375 29 123-", "45-67"]) + '</w:tc></w:tr></w:tbl><w:p><w:r><w:tab/></w:r></w:p>' + w(["Почта ivan@example.test & т. д."]) + "</w:body></w:document>");
  d.file("word/header1.xml", '<w:hdr xmlns:w="w">' + w(["Для Иванов Пётр Сергеевич"]) + "</w:hdr>");
  d.file("word/comments.xml", '<w:comments xmlns:w="w">' + w(["Петрова Анна Ивановна согласна"]) + "</w:comments>");
  d.file("docProps/core.xml", "<cp:coreProperties><dc:creator>Сидоров И. П.</dc:creator><cp:lastModifiedBy>Петрова</cp:lastModifiedBy></cp:coreProperties>");
  const dbuf = await d.generateAsync({ type: "uint8array" });
  const da = await O.anonymizeDocx(dbuf, ["Рога и Копыта"]);
  const dtxt = await allText(da.data);
  noLeak(dtxt); assert.ok(!dtxt.includes("Сидоров") && !dtxt.includes("Петрова"));
  assert.ok(dtxt.includes("[ОРГ_1]") && dtxt.includes("<w:tab/>"));
  assert.ok(!dtxt.includes("& т. д.") && dtxt.includes("&amp; т. д."), "амперсанд экранируется");
  const dr = await O.restoreDocx(da.data, da.map), rtxt = await allText(dr);
  assert.ok(rtxt.includes("Иван Петрович Сидоров") && rtxt.includes("Рога и Копыта") && rtxt.includes("+375 29 123-45-67") && rtxt.includes("Петрова Анна Ивановна"));
  await assert.rejects(() => O.anonymizeDocx(new Uint8Array([1, 2, 3])));
  console.log("Excel, CSV, Word: ok");
})().catch(e => { console.error(e); process.exit(1) });
