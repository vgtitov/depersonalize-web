// Пароли и токены заменяются на [СЕКРЕТ_N] и не попадают в карту.
const assert = require("assert");
const { depersonalize, restore, mapToJson } = require("../engine.js");
const run = t => depersonalize(t, ["x"]);
assert.strictEqual(run("Пароль: Qwerty123 для входа").text, "Пароль: [СЕКРЕТ_1] для входа");
assert.strictEqual(run("password=Zz9!abcd, pass: hunter22").text, "password=[СЕКРЕТ_1], pass: [СЕКРЕТ_2]");
assert.strictEqual(run("token: abcd1234efgh5678").text, "token: [СЕКРЕТ_1]");
assert.strictEqual(run("api key = k_live_51abcXYZ").text, "api key = [СЕКРЕТ_1]");
assert.strictEqual(run("API_KEY=k_live_51abcXYZ").text, "API_KEY=[СЕКРЕТ_1]");
assert.strictEqual(run("Authorization: Bearer abcDEF123456xyz").text, "Authorization: Bearer [СЕКРЕТ_1]");
assert.strictEqual(run("ключ ghp_abcdefghijklmnopqrstuvwxyz0123456789 в коде").text, "ключ [СЕКРЕТ_1] в коде");
// слова без значения не трогаем
assert.strictEqual(run("Нужен пароль от почты, токен выдали вчера.").text, "Нужен пароль от почты, токен выдали вчера.");
// значение секрета не пишется в карту и не возвращается
const r = run("Пароль: Qwerty123");
assert.ok(!mapToJson(r.map).includes("Qwerty123"));
assert.strictEqual(restore(r.text, r.map), "Пароль: [СЕКРЕТ_1]");
console.log("секреты: ok");
