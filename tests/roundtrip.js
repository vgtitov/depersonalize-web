const assert = require("assert");
const { depersonalize, restore, mapToJson, mapFromJson } = require("../engine.js");
const t = "Иван Петрович Сидоров, +375 29 123-45-67, a@b.by, карта 4111 1111 1111 1111. Проект Альфа.";
const r = depersonalize(t, ["Проект Альфа"]);
const ans = "Свяжитесь с " + r.text.match(/\[ЧЕЛОВЕК_1\]/)[0] + " по " + r.text.match(/\[ТЕЛЕФОН_1\]/)[0];
const m2 = mapFromJson(mapToJson(r.map));                       // сохранить и загрузить карту
assert.strictEqual(restore(r.text, m2), t);                      // полный круг
assert.strictEqual(restore(ans, m2), "Свяжитесь с Иван Петрович Сидоров по +375 29 123-45-67");
assert.throws(() => mapFromJson('{"x":1}'));
assert.strictEqual(restore("[ЧЕЛОВЕК_9]", m2), "[ЧЕЛОВЕК_9]");   // неизвестная метка остаётся
console.log("круг анонимизация-деанонимизация: ok");
