// Организации с формой собственности заменяются без словаря.
const assert = require("assert");
const { depersonalize } = require("../engine.js");
const run = t => depersonalize(t, []).text;
assert.strictEqual(run("Письмо от ООО «Рога и Копыта» получено"), "Письмо от [ОРГ_1] получено");
assert.strictEqual(run('Договор с ЗАО "Ромашка" и ОАО «Лютик»'), "Договор с [ОРГ_1] и [ОРГ_2]");
assert.strictEqual(run("Поставщик УП «Север-Юг», ОДО «Альфа», АО «Бета»"), "Поставщик [ОРГ_1], [ОРГ_2], [ОРГ_3]");
assert.strictEqual(run("ИП «Ландыш» закрыт"), "[ОРГ_1] закрыт");
assert.strictEqual(run("ООО Ромашка прислала счёт"), "[ОРГ_1] прислала счёт");
assert.strictEqual(run("ООО «Рога и Копыта» и снова ООО «Рога и Копыта»"), "[ОРГ_1] и снова [ОРГ_1]");
assert.strictEqual(run("Слово ПОДРАЗДЕЛЕНИЕ ООО без названия"), "Слово ПОДРАЗДЕЛЕНИЕ ООО без названия");
assert.strictEqual(depersonalize("ООО «Рога и Копыта»", ["Рога и Копыта"]).text, "[ОРГ_1]");
console.log("организации: ok");
