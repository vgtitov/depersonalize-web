// Случаи из ревью Codex, раунд 2.
const assert = require("assert");
const { depersonalize, restore, mapToJson } = require("../engine.js");
const eq = (t, d, want) => assert.strictEqual(depersonalize(t, d).text, want, t);
// 1. словарь не перебивает секрет, значение не в карте
let r = depersonalize("Пароль: Qwerty123", ["Qwerty123"]);
assert.strictEqual(r.text, "Пароль: [СЕКРЕТ_1]");
assert.ok(!mapToJson(r.map).includes("Qwerty123"));
// 2-3. экранированные кавычки и длинные значения
eq('{"password":"abc\\"DEF123"}', ["x"], '{"password":"[СЕКРЕТ_1]"}');
eq('token: "' + "a".repeat(300) + '"', ["x"], 'token: "[СЕКРЕТ_1]"');
// 4. пароль в адресе вида user:pass@host
r = depersonalize("https://alice:Qwerty123@example.test/", []);
assert.strictEqual(r.text, "[URL_1]");
assert.ok(!mapToJson(r.map).includes("Qwerty123"));
// 5. прописная фамилия на -ович
eq("АБРАМОВИЧ ИВАН ИВАНОВИЧ", ["x"], "[ЧЕЛОВЕК_1]");
// 6. дефисные имена и фамилии
eq("Иван Петров-Сидоров", [], "[ЧЕЛОВЕК_1]");
eq("Анна-Мария Иванова", [], "[ЧЕЛОВЕК_1]");
// 7. разные инициалы — разные люди
eq("Иванов И. И.; Иванов П. П.", [], "[ЧЕЛОВЕК_1]; [ЧЕЛОВЕК_2]");
eq("Иванов И. И.; Иванов Иван Иванович", [], "[ЧЕЛОВЕК_1]; [ЧЕЛОВЕК_1]");
// 8. название без кавычек из нескольких слов; ЧП
eq("ООО Северный Ветер прислало", [], "[ОРГ_1] прислало");
eq("ООО Ромашка Иван Петров прислал", [], "[ОРГ_1] [ЧЕЛОВЕК_1] прислал");
eq("ЧП Ромашка закрыто", [], "[ОРГ_1] закрыто");
// 9. латинский словарь
eq("market grows. Mark's plan", ["Mark"], "market grows. [СЛОВО_1] plan");
// 10. несколько секретов в адресе
r = depersonalize("https://x.test/?token=aaaa1111&password=bbbb2222", []);
assert.strictEqual(restore(r.text, r.map), "https://x.test/?token=[СЕКРЕТ_1]&password=[СЕКРЕТ_2]");
// 11. много разных людей — время
let big = ""; for (let i = 0; i < 20000; i++) big += "Иван А" + i.toString(36).replace(/[0-9]/g, "б") + "ов\n";
const t0 = Date.now(); depersonalize(big, []);
assert.ok(Date.now() - t0 < 5000, "слишком долго: " + (Date.now() - t0) + " мс");
console.log("ревью раунд 2: ok");
