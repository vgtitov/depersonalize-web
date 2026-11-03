// index.html содержит ровно те файлы, что проверяют тесты, и остаётся автономным.
const assert = require("assert"), fs = require("fs"), path = require("path");
const root = path.join(__dirname, "..");
const html = fs.readFileSync(path.join(root, "index.html"), "utf8");
for (const f of ["vendor/jszip.min.js", "engine.js", "office.js", "ui.js"]) {
  const src = fs.readFileSync(path.join(root, f), "utf8").replace(/<\/script/g, "<\\/script");
  assert.ok(html.includes("<script>\n" + src + "\n</script>"), f + " не вставлен в index.html: sh build.sh");
}
assert.ok(/Content-Security-Policy" content="default-src 'none'/.test(html), "нет CSP");
const scripts = html.match(/<script[^>]*>/g) || [];
assert.ok(scripts.every(s => s === "<script>"), "внешние скрипты запрещены");
assert.ok(!/<link\b|<iframe|<img[^>]+src="http/.test(html), "внешние ресурсы запрещены");
for (const f of ["engine.js", "office.js", "ui.js"]) {
  const src = fs.readFileSync(path.join(root, f), "utf8");
  assert.ok(!/\bfetch\s*\(|XMLHttpRequest|WebSocket|sendBeacon|localStorage|sessionStorage|indexedDB|document\.cookie/.test(src), f + ": сеть или хранение запрещены");
}
console.log("сборка, CSP, нет сети и хранения: ok");
// в браузере движок берётся из глобальных функций: набор должен совпадать с экспортом Node
const off = fs.readFileSync(path.join(root, "office.js"), "utf8"), eng = require("../engine.js");
const m = /\{ ([^}]*) \} : require/.exec(off);
assert.ok(m, "не нашёл привязку движка в office.js");
for (const n of m[1].split(", ")) assert.ok(typeof eng[n] === "function", n + " нет в экспорте engine.js");
for (const n of Object.keys(eng)) if (!m[1].split(", ").includes(n) && /depersonalize|restore/.test(n)) assert.fail(n + " не подключён в браузерной привязке office.js");
console.log("привязка движка в браузере: ok");
// блок «сообщить о проблеме» есть и ведёт только на страницу запросов репозитория
assert.ok(/id="help"/.test(html) && html.includes('href="https://github.com/vgtitov/depersonalize-web/issues/new/choose"'), "нет блока про сообщение о проблеме");
assert.ok(/настоящие данные/.test(html.split('id="help"')[1].split("</footer>")[0]), "в блоке нет предупреждения не присылать настоящие данные");
console.log("сообщить о проблеме: ok");
// переключатель темы: кнопка, обе темы заданы в CSS, без хранения выбора
assert.ok(/id="theme"/.test(html) && /:root\[data-theme=dark\]/.test(html) && /:root\[data-theme=light\]|:not\(\[data-theme=light\]\)/.test(html), "нет переключателя темы");
console.log("переключатель темы: ok");
