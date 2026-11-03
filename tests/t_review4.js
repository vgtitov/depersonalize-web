// Случаи из ревью, раунд 4.
const assert = require("assert");
const Z = require("../vendor/jszip.min.js");
const O = require("../office.js"), E = require("../engine.js");
const eq = (t, d, want) => assert.strictEqual(E.depersonalize(t, d).text, want, t);
async function parts(data) { const z = await Z.loadAsync(data); const o = {}; for (const n of Object.keys(z.files)) if (!z.files[n].dir) o[n] = await z.file(n).async("string"); return o }
const W = '<w:document xmlns:w="w"><w:body>', WE = "</w:body></w:document>";
async function docx(files) { const d = new Z(); d.file("[Content_Types].xml", "<Types/>"); for (const [k, v] of Object.entries(files)) d.file(k, v); return d.generateAsync({ type: "uint8array" }) }
const leak = async (files, bads) => { const r = await O.anonymizeDocx(await docx(files), []); const t = Object.values(await parts(r.data)).join("\n"); for (const b of bads) assert.ok(!t.includes(b), "осталось: " + b + " в " + Object.entries(await parts(r.data)).filter(([, v]) => v.includes(b)).map(([k, v]) => k + ": " + v.slice(Math.max(0, v.indexOf(b) - 60), v.indexOf(b) + 40)).join(" | ")); return r };
(async () => {
  // P1-1, P2-4: фамилия во вставленном фрагменте, порядок вокруг исправления
  let r = await leak({ "word/document.xml": W + '<w:p><w:r><w:t xml:space="preserve">Звонил Иван </w:t></w:r><w:ins w:id="1" w:author="x"><w:r><w:t>Сидоров</w:t></w:r></w:ins><w:r><w:t xml:space="preserve"> вчера</w:t></w:r></w:p>' + WE }, ["Сидоров"]);
  const ord = await O.anonymizeDocx(await docx({ "word/document.xml": W + '<w:p><w:r><w:t xml:space="preserve">Иван Сидоров </w:t></w:r><w:ins w:id="1" w:author="x"><w:r><w:t>срочно</w:t></w:r></w:ins><w:r><w:t xml:space="preserve"> звонил</w:t></w:r></w:p>' + WE }), []);
  const txt = (await parts(ord.data))["word/document.xml"].replace(/<[^>]+>/g, "");
  assert.strictEqual(txt, "[ЧЕЛОВЕК_1] срочно звонил");
  // P1-2: другие места
  await leak({
    "word/document.xml": W + '<w:p><w:r><w:instrText> HYPERLINK "mailto:ivan.sidorov@example.test" </w:instrText></w:r><w:r><w:drawing><wp:docPr id="1" descr="Фото: Иван Петрович Сидоров" title="Сидоров"/></w:drawing></w:r></w:p>' + WE,
    "customXml/item1.xml": "<root><Client>Иван Петрович Сидоров</Client><Mail>ivan@example.test</Mail></root>",
    "word/charts/chart1.xml": "<c:chart><c:pt><c:v>Иван Петрович Сидоров</c:v></c:pt></c:chart>",
    "docProps/custom.xml": "<Properties><vt:lpwstr>Иван Петрович Сидоров</vt:lpwstr></Properties>",
    "word/_rels/settings.xml.rels": '<Relationships><Relationship Id="r1" Type="t" Target="file:///C:\\Users\\isidorov\\AppData\\Normal.dotm" TargetMode="External"/></Relationships>',
  }, ["Сидоров", "sidorov", "isidorov", "ivan@"]);
  // P1-3 и P3-11, 12: пароли
  for (const [t, w] of [["Пароль от личного кабинета: Qwerty123", "Пароль от личного кабинета: [СЕКРЕТ_1]"], ["Пароль для входа в систему: Qwerty123", "Пароль для входа в систему: [СЕКРЕТ_1]"],
    ["пароль 12345678", "пароль [СЕКРЕТ_1]"], ["Пароль от Wi-Fi Qwerty123", "Пароль от Wi-Fi [СЕКРЕТ_1]"], ["пароль: Qw,erty;123", "пароль: [СЕКРЕТ_1]"],
    ["пароль: Qwerty123!, логин admin", "пароль: [СЕКРЕТ_1], логин admin"], ["Пароль — это набор символов", "Пароль — это набор символов"], ["Секрет успеха прост: работать", "Секрет успеха прост: работать"]]) eq(t, ["x"], w);
  r = await leak({ "word/document.xml": W + "<w:p><w:r><w:t>Пароль:</w:t></w:r></w:p><w:p><w:r><w:t>Qwerty123</w:t></w:r></w:p>" + WE }, []);
  // P2-5: ячейки с цифрами рядом с телефоном; пустой документ
  r = await O.anonymizeDocx(await docx({ "word/document.xml": W + "<w:p><w:r><w:t>Звоните +7 495 1234567</w:t></w:r></w:p><w:p><w:r><w:t>15 шт. к поставке</w:t></w:r></w:p>" + WE }), []);
  assert.ok((await parts(r.data))["word/document.xml"].includes("15 шт. к поставке"));
  let rows = [["К"], ["Звоните +7 495 1234567"], ["15 шт."]]; let cols = O.detectColumns(rows); cols[0].mode = "scan";
  assert.strictEqual(O.anonymizeTable(rows, cols).rows[2][0], "15 шт.");
  await O.anonymizeDocx(await docx({ "word/document.xml": W + '<w:p><w:r><w:drawing/></w:r></w:p>' + WE }), []);
  await assert.rejects(async () => O.anonymizeDocx(await docx({ "word/document.xml": '<ns0:document xmlns:ns0="x"><ns0:p><ns0:t>Иван Сидоров</ns0:t></ns0:p></ns0:document>' }), []), /Word/);
  // P2-7: ложные замены
  eq("Рада Вам сообщить", [], "Рада Вам сообщить"); eq("Марка Автомобиля: Тойота", [], "Марка Автомобиля: Тойота");
  eq("Вариант Б. Магазин закрыт. Раздел А. Бензин.", [], "Вариант Б. Магазин закрыт. Раздел А. Бензин.");
  eq("Иван Петрович Сказал, что придёт", [], "[ЧЕЛОВЕК_1] Сказал, что придёт");
  // P2-8: двойная фамилия
  eq("Иван Петрович Римский-Корсаков", [], "[ЧЕЛОВЕК_1]"); eq("Иван Петрович Коваль-Мельник", [], "[ЧЕЛОВЕК_1]");
  // P2-9: даты 1904, время
  const x = new Z();
  x.file("xl/workbook.xml", '<workbook xmlns:r="x"><workbookPr date1904="1"/><sheets><sheet name="Л" sheetId="1" r:id="rId1"/></sheets></workbook>');
  x.file("xl/_rels/workbook.xml.rels", '<Relationships><Relationship Id="rId1" Type="t" Target="worksheets/sheet1.xml"/></Relationships>');
  x.file("xl/styles.xml", '<styleSheet><numFmts><numFmt numFmtId="165" formatCode="[h]:mm"/></numFmts><cellXfs><xf numFmtId="0"/><xf numFmtId="14"/><xf numFmtId="20"/><xf numFmtId="165"/></cellXfs></styleSheet>');
  x.file("xl/worksheets/sheet1.xml", '<worksheet><sheetData><row r="1"><c r="A1" s="1"><v>44500</v></c><c r="B1" s="2"><v>0.5</v></c><c r="C1" s="3"><v>1.25</v></c></row></sheetData></worksheet>');
  assert.deepStrictEqual((await O.readXlsx(await x.generateAsync({ type: "uint8array" })))[0].rows, [["2025-11-01", "12:00", "30:00"]]);
  // P2-10: текстовый телефон остаётся текстом
  let a = O.anonymizeTable([["Тел"], ["375291112233"], [375291112244]], [{ index: 0, mode: "key", kind: "ТЕЛЕФОН" }]);
  const back = O.restoreRows(a.rows, a.map); assert.strictEqual(back[1][0], "375291112233"); assert.strictEqual(back[2][0], 375291112244);
  assert.ok(!JSON.stringify([...E.depersonalize("x", []).map]).includes("число"));
  // P3-13: подсказки колонок
  assert.strictEqual(O.detectColumns([["Показатель", "Расчет суммы", "Тел."], ["1", "2", "3"]]).map(c => c.mode).join(), "keep,keep,key");
  // P3-14: U+2028 в ячейке сохраняется
  a = O.anonymizeTable([["К"], ["a\u2028b Иван Сидоров"]], [{ index: 0, mode: "scan", kind: "ЗНАЧЕНИЕ" }]); assert.strictEqual(a.rows[1][0], "a\u2028b [ЧЕЛОВЕК_1]");
  console.log("ревью раунд 4: ok");
})().catch(e => { console.error(e); process.exit(1) });
