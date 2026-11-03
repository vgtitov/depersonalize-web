// Случаи из ревью Opus, раунд 3.
const assert = require("assert");
const Z = require("../vendor/jszip.min.js");
const O = require("../office.js"), E = require("../engine.js");
const eq = (t, d, want) => assert.strictEqual(E.depersonalize(t, d).text, want, t);
async function parts(data) { const z = await Z.loadAsync(data); const o = {}; for (const n of Object.keys(z.files)) if (!z.files[n].dir) o[n] = await z.file(n).async("string"); return o }
const W = '<w:document xmlns:w="w"><w:body>', WE = "</w:body></w:document>";
async function docx(files) { const d = new Z(); d.file("[Content_Types].xml", "<Types/>"); for (const [k, v] of Object.entries(files)) d.file(k, v); return d.generateAsync({ type: "uint8array" }) }
const leak = async (files, bads) => { const r = await O.anonymizeDocx(await docx(files), []); const p = await parts(r.data), t = Object.values(p).join("\n"); for (const b of bads) assert.ok(!t.includes(b), "осталось: " + b + " в " + Object.entries(p).filter(([, v]) => v.includes(b)).map(([k, v]) => k + ": " + v.slice(Math.max(0, v.indexOf(b) - 60), v.indexOf(b) + 40)).join(" | ")); return { r, p } };
(async () => {
  // P1-1: код поля разбит; fldSimple
  await leak({ "word/document.xml": W + '<w:p><w:r><w:instrText> HYPERLINK "mailto:ivan.sid</w:instrText></w:r><w:r><w:instrText>orov@example.test" </w:instrText></w:r><w:fldSimple w:instr=" HYPERLINK &quot;mailto:ivan.sidorov@example.test&quot; "><w:r><w:t>x</w:t></w:r></w:fldSimple></w:p>' + WE }, ["ivan.sid", "sidorov"]);
  // P1-2: атрибуты и части
  await leak({
    "word/document.xml": W + '<w:p><w:hyperlink w:tooltip="Профиль: Иван Петрович Сидоров"><w:r><w:t>x</w:t></w:r></w:hyperlink><w:r><w:drawing><pic:cNvPr id="0" name="Сидоров Иван Петрович паспорт.jpg"/></w:drawing></w:r><w:pict><v:shape alt="Иван Петрович Сидоров"><v:imagedata o:title="Иван Сидоров"/></v:shape></w:pict><w:smartTag><w:smartTagPr><w:attr w:name="ProductID" w:val="Иван Петрович Сидоров"/></w:smartTagPr></w:smartTag><w:sdt><w:sdtPr><w:docPart w:val="Иван Петрович Сидоров"/></w:sdtPr></w:sdt></w:p>' + WE,
    "word/settings.xml": '<w:settings xmlns:w="w"><w:docVars><w:docVar w:name="Client" w:val="Иван Петрович Сидоров"/></w:docVars></w:settings>',
    "docProps/app.xml": "<Properties><TitlesOfParts><vt:vector><vt:lpstr>Договор с Иваном Петровичем Сидоровым</vt:lpstr></vt:vector></TitlesOfParts><Words>1234567</Words></Properties>",
    "customXml/item1.xml": '<root><Person name="Иван Петрович Сидоров" mail="ivan@example.test"/><!-- автор Иван Петрович Сидоров --></root>',
  }, ["Сидоров", "ivan@", "Сидоровым"]);
  // app.xml: числа не трогаем
  assert.ok((await parts((await O.anonymizeDocx(await docx({ "word/document.xml": W + "<w:p/>" + WE, "docProps/app.xml": "<Properties><Words>1234567</Words></Properties>" }), [])).data))["docProps/app.xml"].includes("1234567"));
  // P1-3: пароли
  for (const [t, w] of [["пин-код 4821, PIN 4821, CVV 123", "пин-код [СЕКРЕТ_1], PIN [СЕКРЕТ_1], CVV [СЕКРЕТ_2]"], ["Пароль на сейф 4821", "Пароль на сейф [СЕКРЕТ_1]"], ["Код доступа: 4821", "Код доступа: [СЕКРЕТ_1]"],
    ["Password is Qwerty123", "Password is [СЕКРЕТ_1]"], ["пароль 'Qwerty123'", "пароль '[СЕКРЕТ_1]'"], ["Пароль к Wi-Fi гостевой сети офиса на втором этаже: Qwerty123", "Пароль к Wi-Fi гостевой сети офиса на втором этаже: [СЕКРЕТ_1]"]]) eq(t, ["x"], w);
  // P2-4: перемещённый текст
  const mv = await leak({ "word/document.xml": W + '<w:p><w:moveFrom w:id="1"><w:r><w:t>Иван Петрович Сидоров</w:t></w:r></w:moveFrom><w:r><w:t> подписал</w:t></w:r></w:p>' + WE }, ["Сидоров"]);
  assert.ok(mv.p["word/document.xml"].includes(" подписал"), "соседний текст не должен уйти в перемещение");
  assert.ok(mv.r.warn.some(w => /исправлен/.test(w)));
  // P2-5: случайное слово не становится фамилией
  eq("Иван Петрович Директор подписал. Директор уехал.", [], "[ЧЕЛОВЕК_1] подписал. Директор уехал.");
  eq("Анна Банк просит акт. Банк закрыт.", [], "[ЧЕЛОВЕК_1] просит акт. Банк закрыт.");
  assert.deepStrictEqual(E.depersonalizeMany(["Ответственный: Ольга Зима", "План на Зима 2026"], []).texts[1], "План на Зима 2026");
  // P3-7: числа диаграммы остаются числами
  assert.ok((await leak({ "word/document.xml": W + "<w:p/>" + WE, "word/charts/c1.xml": "<c:numCache><c:pt><c:v>4111111111111111</c:v></c:pt></c:numCache>" }, [])).p["word/charts/c1.xml"].includes("4111111111111111"));
  // раунд 4
  eq("Ольга Ткач пришла. Ткач ушла. Иван Мельник звонил, Мельнику перезвонить", [], "[ЧЕЛОВЕК_1] пришла. [ЧЕЛОВЕК_1] ушла. [ЧЕЛОВЕК_2] звонил, [ЧЕЛОВЕК_2] перезвонить");
  const fl = await leak({ "word/document.xml": W + '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> HYPERLINK "mailto:ivan@example.test" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>' + WE }, ["ivan@"]);
  assert.ok(fl.p["word/document.xml"].includes("> PAGE <"), "второе поле должно остаться");
  const sc = await leak({ "word/document.xml": W + '<w:p><w:bookmarkStart w:id="1" w:name="Сидоров_Иван"/></w:p>' + WE, "customXml/itemProps1.xml": '<ds:datastoreItem ds:itemID="{ABC}"><ds:schemaRefs><ds:schemaRef ds:uri="http://schemas.openxmlformats.org/officeDocument/2006/bibliography"/></ds:schemaRefs></ds:datastoreItem>' }, []);
  assert.ok(sc.p["customXml/itemProps1.xml"].includes("officeDocument/2006/bibliography") && sc.p["customXml/itemProps1.xml"].includes("{ABC}"));
  assert.ok(sc.r.warn.some(w => /закладк/.test(w)));
  // раунд 5: вложенное поле не уезжает
  const nf = await leak({ "word/document.xml": W + '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> IF 1 = 1 "</w:instrText></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> MERGEFIELD x </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:instrText>" "ivan.sid</w:instrText></w:r><w:r><w:instrText>orov@example.test" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>' + WE }, ["ivan.sid"]);
  assert.ok(/IF 1 = 1 "<\/w:instrText>[\s\S]*MERGEFIELD x[\s\S]*\[EMAIL_1\]/.test(nf.p["word/document.xml"]), "порядок кусков кода сохранён");
  console.log("ревью раунд 5 (Opus 3): ok");
})().catch(e => { console.error(e); process.exit(1) });
