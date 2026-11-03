// index.html содержит ровно тот engine.js, что проверяют тесты.
const assert = require("assert"), fs = require("fs"), path = require("path");
const root = path.join(__dirname, "..");
const html = fs.readFileSync(path.join(root, "index.html"), "utf8");
const eng = fs.readFileSync(path.join(root, "engine.js"), "utf8");
assert.ok(html.includes("<!--ENGINE-->\n<script>\n" + eng + "</script>\n<!--/ENGINE-->"), "index.html не собран: sh build.sh");
assert.ok(/Content-Security-Policy" content="default-src 'none'/.test(html), "нет CSP");
assert.ok(!/<script[^>]+src=|<link[^>]+href=|https?:\/\//.test(html.replace(/<p class="note">[\s\S]*?<\/p>/, "")), "во внешних ссылках быть нечему");
console.log("сборка и CSP: ok");
