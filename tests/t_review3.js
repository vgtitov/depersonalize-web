// Случаи из ревью (раунд 3): Excel/Word и движок. Файлы собираются в тесте.
const assert = require("assert");
const Z = require("../vendor/jszip.min.js");
const O = require("../office.js"), E = require("../engine.js");
const eq = (t, d, want) => assert.strictEqual(E.depersonalize(t, d).text, want, t);
async function parts(data) { const z = await Z.loadAsync(data); const o = {}; for (const n of Object.keys(z.files)) if (!z.files[n].dir) o[n] = await z.file(n).async("string"); return o }
const all = o => Object.values(o).join("\n");
(async () => {
  // 1-2. режим «найти внутри текста»: несколько колонок, переносы строк
  let rows = [["Комментарий", "Заметка"], ["Звонил Иван Петрович Сидоров", "Встреча с Анна Сергеевна Козлова"]];
  let cols = O.detectColumns(rows); cols.forEach(c => c.mode = "scan");
  let r = O.anonymizeTable(rows, cols);
  assert.notStrictEqual(r.rows[1][0], r.rows[1][1]);
  assert.deepStrictEqual(O.restoreRows(r.rows, r.map)[1], rows[1]);
  rows = [["К"], ["Строка1\nИван Петрович Сидоров"], ["Анна Сергеевна Козлова"], ["третья"]];
  cols = O.detectColumns(rows); cols[0].mode = "scan"; r = O.anonymizeTable(rows, cols);
  assert.strictEqual(r.rows.length, 4); assert.strictEqual(r.rows[3][0], "третья"); assert.ok(r.rows[1][0].startsWith("Строка1\n"));
  assert.deepStrictEqual(O.restoreRows(r.rows, r.map), rows);
  // 3-5. Word: автор правок, mailto, свойства, надписи, исправления, вложенные абзацы
  const w = (inner) => '<w:document xmlns:w="w"><w:body>' + inner + "</w:body></w:document>";
  const d = new Z(); d.file("[Content_Types].xml", "<Types/>");
  d.file("word/document.xml", w(
    '<w:p><w:r><w:t>До</w:t></w:r><w:r><w:txbxContent><w:p><w:r><w:t>Внутри рамки</w:t></w:r></w:p></w:txbxContent></w:r><w:r><w:t> Иван Петрович Сидоров, тел. +375 29 123-45-67</w:t></w:r></w:p>' +
    '<w:p><w:del w:id="1" w:author="Козлова Мария"><w:r><w:delText>Ответственный Иван Сидоров</w:delText></w:r></w:del><w:ins w:id="2" w:author="Козлова Мария"><w:r><w:t>Пётр Иванов звонил</w:t></w:r></w:ins></w:p>' +
    '<w:p><w:r><w:t>Строка &#10; Анна Сергеевна Козлова</w:t></w:r></w:p>'));
  d.file("word/comments.xml", '<w:comments xmlns:w="w"><w:comment w:id="0" w:author="Козлова Мария" w:initials="КМ"><w:p><w:r><w:t>ок</w:t></w:r></w:p></w:comment></w:comments>');
  d.file("word/people.xml", '<w15:people xmlns:w15="x"><w15:person w15:author="Козлова Мария"><w15:presenceInfo w15:userId="maria@example.test" w15:providerId="None"/></w15:person></w15:people>');
  d.file("word/_rels/document.xml.rels", '<Relationships><Relationship Id="rId9" Type="h" Target="mailto:maria.kozlova@example.test" TargetMode="External"/><Relationship Id="rId1" Type="s" Target="styles.xml"/></Relationships>');
  d.file("docProps/app.xml", "<Properties><Company>ООО «Ромашка»</Company><Manager>Козлова Мария</Manager></Properties>");
  d.file("docProps/core.xml", "<cp:coreProperties><dc:title>Договор Сидорова</dc:title><dc:creator>Сидоров И. П.</dc:creator></cp:coreProperties>");
  d.file("word/embeddings/Microsoft_Excel_Sheet1.xlsx", "x");
  const da = await O.anonymizeDocx(await d.generateAsync({ type: "uint8array" }), []);
  const txt = all(await parts(da.data));
  for (const bad of ["Сидоров", "123-45-67", "Козлов", "Иванов", "kozlova", "maria@", "Ромашка", "Договор Сидорова"]) assert.ok(!txt.includes(bad), "осталось: " + bad);
  assert.ok(da.warn.some(x => /вложенн/.test(x)), "нет предупреждения про вложенные файлы");
  const ins = (await parts(da.data))["word/document.xml"];
  assert.ok(/<w:delText[^>]*>[^<]*<\/w:delText>/.test(ins) && /<w:ins [^>]*><w:r><w:t[^>]*>[^<]+<\/w:t>/.test(ins), "текст исправления остался в своём месте");
  const back = all(await parts(await O.restoreDocx(da.data, da.map)));
  assert.ok(back.includes("Иван Петрович Сидоров") && back.includes("maria.kozlova@example.test"));
  // 6. имена вне списка, один инициал, фамилия без суффикса
  eq("Ярослав Сидоров, Анатолий Петров, Тамара Иванова, Руслан Козлов", [], "[ЧЕЛОВЕК_1], [ЧЕЛОВЕК_2], [ЧЕЛОВЕК_3], [ЧЕЛОВЕК_4]");
  eq("И. Сидоров позвонит", [], "[ЧЕЛОВЕК_1] позвонит");
  eq("Иван Петрович Коваль приехал", [], "[ЧЕЛОВЕК_1] приехал");
  eq("Иван Мельник и Ольга Ткач", [], "[ЧЕЛОВЕК_1] и [ЧЕЛОВЕК_2]");
  eq("Приехал Иван в понедельник", [], "Приехал Иван в понедельник");
  // 7. белорусские формы собственности и полное название
  eq("СООО «Ромашка», ИООО «Лютик», ЧТУП «Василёк», РУП «Белсвязь»", [], "[ОРГ_1], [ОРГ_2], [ОРГ_3], [ОРГ_4]");
  eq("Общество с ограниченной ответственностью «Ромашка» прислало", [], "[ОРГ_1] прислало");
  // 8. формулировки пароля
  eq("Пароль от почты: Qwerty123", ["x"], "Пароль от почты: [СЕКРЕТ_1]");
  eq("пароль Qwerty123 дан", ["x"], "пароль [СЕКРЕТ_1] дан");
  eq("Пароль:\nQwerty123", ["x"], "Пароль:\n[СЕКРЕТ_1]");
  eq("PIN: 4821, пин-код: 4821", ["x"], "PIN: [СЕКРЕТ_1], пин-код: [СЕКРЕТ_1]");
  eq("пароль доступа выдан", ["x"], "пароль доступа выдан");
  // 9. колонка «Пароль»
  cols = O.detectColumns([["Пароль", "Адрес", "ИНН"], ["abc123", "ул. Мира 5", "123456789"]]);
  assert.strictEqual(cols[0].mode, "drop"); assert.strictEqual(cols[1].mode, "key"); assert.strictEqual(cols[2].mode, "key");
  // 10. xlsx: ячейка без r, ошибка, даты, дальний край
  const x = new Z();
  x.file("xl/workbook.xml", '<workbook xmlns:r="x"><sheets><sheet name="Л" sheetId="1" r:id="rId1"/></sheets></workbook>');
  x.file("xl/_rels/workbook.xml.rels", '<Relationships><Relationship Id="rId1" Type="t" Target="worksheets/sheet1.xml"/></Relationships>');
  x.file("xl/sharedStrings.xml", "<sst><si><t>А</t></si><si><t>Б</t></si></sst>");
  x.file("xl/styles.xml", '<styleSheet><numFmts><numFmt numFmtId="164" formatCode="dd.mm.yyyy"/></numFmts><cellXfs><xf numFmtId="0"/><xf numFmtId="14"/><xf numFmtId="164"/></cellXfs></styleSheet>');
  x.file("xl/worksheets/sheet1.xml", '<worksheet><sheetData><row r="1"><c t="s"><v>0</v></c><c t="s"><v>1</v></c><c r="C1" t="e"><v>#N/A</v></c></row><row r="2"><c r="A2" s="1"><v>45962</v></c><c r="B2" s="2"><v>45962.5</c></c><c r="XFD2" s="1"/></row></sheetData></worksheet>'.replace("45962.5</c></c>", "45962.5</v></c>"));
  const sh = (await O.readXlsx(await x.generateAsync({ type: "uint8array" })))[0].rows;
  assert.deepStrictEqual(sh, [["А", "Б", "#N/A"], ["2025-11-01", "2025-11-01 12:00", ""]]);
  // 12. значение, похожее на ключ, в оставленной колонке: предупреждение
  r = O.anonymizeTable([["Имя", "Код"], ["Сидоров Иван", "ЧЕЛОВЕК_001"]], [{ index: 0, mode: "key", kind: "ЧЕЛОВЕК" }, { index: 1, mode: "keep", kind: "ЗНАЧЕНИЕ" }]);
  assert.ok(r.warn.some(x => /похож/.test(x)));
  // 14. число в колонке с кодами возвращается числом
  r = O.anonymizeTable([["Тел"], [375291112233]], [{ index: 0, mode: "key", kind: "ТЕЛЕФОН" }]);
  assert.strictEqual(O.restoreRows(r.rows, r.map)[1][0], 375291112233);
  // 15. слова-ловушки и телефоны
  eq("Магазин Елена закрыт", [], "Магазин Елена закрыт");
  eq("Господин Иван пришёл", [], "Господин Иван пришёл");
  eq("Звоните 29 123-45-67 и (029)1234567", [], "Звоните [ТЕЛЕФОН_1] и [ТЕЛЕФОН_2]");
  console.log("ревью раунд 3: ok");
})().catch(e => { console.error(e); process.exit(1) });
