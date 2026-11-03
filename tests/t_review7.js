// Случаи из ревью, раунд 7: пароль с & вне URL, query-строка, таблица key+scan, ссылки в DOCX.
const assert = require("assert");
const { depersonalize, restore } = require("../engine.js");
const O = require("../office.js");
const JSZip = require("../vendor/jszip.min.js");
const eq = (t, want) => assert.strictEqual(depersonalize(t, []).text, want, t);
// 1. & внутри пароля вне URL не обрывает значение
eq("Пароль: Abc&admin=1", "Пароль: [СЕКРЕТ_1]");
eq("pwd=Abc&admin=1 и дальше", "pwd=[СЕКРЕТ_1] и дальше");
// 2. в query-строке & по-прежнему разделяет параметры
eq("запрос user=ivan&password=Qwerty123&mode=1", "запрос user=ivan&password=[СЕКРЕТ_1]&mode=1");
// 3. пароль не попадает в карту и не возвращается
const r = depersonalize("Пароль: Abc&admin=1", []);
assert.ok(!restore(r.text, r.map).includes("admin"));
(async () => {
  // 4. таблица: ключи колонки и метки поиска в тексте не пересекаются, возврат точный
  const rows = [["ФИО", "Заметка"], ["Иванов Иван", "звонил Петров Пётр"], ["Сидоров Олег", "писал Иванов Иван"]];
  const t = O.anonymizeTable(rows, [{ index: 0, kind: "ЧЕЛОВЕК", mode: "key" }, { index: 1, kind: "ТЕКСТ", mode: "scan" }], { hasHeader: true, dict: [] });
  assert.deepStrictEqual(O.restoreRows(t.rows, t.map), rows);
  // 5. DOCX: адреса внешних ссылок обезличиваются вместе с текстом
  const z = await JSZip.loadAsync(await O.makeDocx({ paragraphs: ["Ссылка"] }));
  z.file("word/_rels/document.xml.rels", '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="x" Target="https://host.example/reset?token=Qwerty123" TargetMode="External"/><Relationship Id="rId8" Type="x" Target="mailto:ivan@example.test" TargetMode="External"/></Relationships>');
  const o = await O.anonymizeDocx(await z.generateAsync({ type: "uint8array" }), []);
  const rels = await (await JSZip.loadAsync(o.data)).file("word/_rels/document.xml.rels").async("string");
  assert.ok(!/Qwerty123|ivan@example/.test(rels), rels);
  console.log("t_review7: ok");
})();
