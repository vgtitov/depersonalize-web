// ФИО прописными, в том числе с лишним словом перед ним.
const assert = require("assert");
const { depersonalize } = require("../engine.js");
const run = t => depersonalize(t, ["x"]).text;
assert.strictEqual(run("КЛИЕНТ ИВАНОВ ИВАН ИВАНОВИЧ"), "КЛИЕНТ [ЧЕЛОВЕК_1]");
assert.strictEqual(run("ИВАНОВ ИВАН ИВАНОВИЧ"), "[ЧЕЛОВЕК_1]");
assert.strictEqual(run("ДОГОВОР ПОДПИСАЛ ПЕТРОВ ПЁТР СЕРГЕЕВИЧ СЕГОДНЯ"), "ДОГОВОР ПОДПИСАЛ [ЧЕЛОВЕК_1] СЕГОДНЯ");
assert.strictEqual(run("ЗАЯВКА ИВАН ИВАНОВИЧ ИВАНОВ"), "ЗАЯВКА [ЧЕЛОВЕК_1]");
assert.strictEqual(run("ВАЖНО СРОЧНО ОТВЕТИТЬ"), "ВАЖНО СРОЧНО ОТВЕТИТЬ");
console.log("прописные ФИО: ok");
