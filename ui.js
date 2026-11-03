// Интерфейс. Логика в engine.js и office.js, здесь только кнопки и показ.
(function () {
  const $ = id => document.getElementById(id);
  const dict = () => $("dict").value.split("\n");
  const M = { text: new Map(), xl: new Map(), doc: new Map() }; // у каждой вкладки свой ключ (в памяти страницы)
  const X = { rows: null, sheets: null, name: "таблица", csv: false, delim: ";", cols: [], result: null };
  const D = { buf: null, name: "документ", result: null };

  // ---- общее ----
  const say = (id, text, kind) => { const e = $(id); e.hidden = false; e.className = "msg " + (kind || ""); e.textContent = text };
  const fail = (id, e) => say(id, "Не получилось: " + friendly(e), "bad");
  const friendly = e => /не файл|не файл Excel|не файл Word|не карта|нет листов|Can't find end|corrupt|zip/i.test(String(e && e.message || e)) ? "файл не похож на нужный или повреждён. Проверьте, что это именно тот тип файла, и попробуйте ещё раз." : String(e && e.message || e);
  const download = (name, data, type) => { const a = document.createElement("a"); a.href = URL.createObjectURL(new Blob([data], { type })); a.download = name; document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 5000) };
  const baseName = f => f.replace(/\.[^.]+$/, "");
  const counts = c => Object.entries(c).map(([k, v]) => k + ": " + v).join(", ") || "ничего не найдено";
  const TYPES = { xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", csv: "text/csv;charset=utf-8", json: "application/json" };
  const saveKey = tab => { if (!M[tab].size) return false; download("ключ-обезличивания.json", mapToJson(M[tab]), TYPES.json); return true };
  async function loadKey(file, msgId, tab) {
    if (!file) return;
    try { M[tab] = mapFromJson(await file.text()); say(msgId, "Ключ загружен, значений: " + [...M[tab].keys()].filter(k => k[0] !== "#").length, "ok") }
    catch (e) { say(msgId, "Это не файл-ключ, который сохранила эта страница. Выберите файл «ключ-обезличивания.json».", "bad") }
  }
  const needKey = (id, tab) => { if (M[tab].size) return true; say(id, "Нет ключа. Сначала скройте данные на этой странице или загрузите файл-ключ.", "bad"); return false };
  async function readTable(file) {
    if (/\.csv$/i.test(file.name)) {
      const buf = await file.arrayBuffer(); let t = new TextDecoder("utf-8").decode(buf);
      if (t.includes("\uFFFD")) t = new TextDecoder("windows-1251").decode(buf);
      const p = parseCsv(t); return { csv: true, delim: p.delim, sheets: [{ name: "Лист1", rows: p.rows }] };
    }
    return { csv: false, sheets: await readXlsx(await file.arrayBuffer()) };
  }

  // ---- вкладки ----
  const tabs = ["text", "xl", "doc"];
  const select = k => tabs.forEach(t => { const on = t === k; $("t-" + t).setAttribute("aria-selected", on); $("t-" + t).tabIndex = on ? 0 : -1; $("p-" + t).hidden = !on });
  tabs.forEach((t, i) => { $("t-" + t).onclick = () => select(t); $("t-" + t).onkeydown = e => { const d = e.key === "ArrowRight" ? 1 : e.key === "ArrowLeft" ? -1 : 0; if (d) { const n = tabs[(i + d + 3) % 3]; select(n); $("t-" + n).focus() } } });

  // ---- текст ----
  const EXAMPLE_TEXT = "Здравствуйте! Иван Петрович Сидоров из ООО «Ромашка» просит согласовать счёт. Позвоните ему по +375 29 123-45-67 или напишите на ivan.sidorov@example.test. Пароль: Qwerty123. Директор Мария Козлова сказала Сидорову, что отчёт подготовит Пётр Иванов.";
  $("ex-text").onclick = () => { $("src").value = EXAMPLE_TEXT; $("go").click() };
  $("go").onclick = () => {
    if (!$("src").value.trim()) return say("rep", "Сначала вставьте текст в поле выше.", "bad");
    try {
      const r = depersonalize($("src").value, dict()); M.text = r.map; $("out").value = r.text;
      say("rep", "Заменено: " + counts(r.counts) + (r.warn.length ? "\nВнимание: " + r.warn.join("; ") : "") + "\nПроверьте копию глазами перед отправкой.", r.warn.length ? "warn" : "ok");
      $("copy").disabled = $("savekey-text").disabled = false;
    } catch (e) { fail("rep", e) }
  };
  $("copy").onclick = async () => { try { await navigator.clipboard.writeText($("out").value) } catch (e) { $("out").select(); document.execCommand("copy") } say("m-text", "Скопировано. Вставьте в ИИ.", "ok") };
  $("savekey-text").onclick = () => { if (saveKey("text")) say("m-text", "Ключ сохранён. Держите файл у себя: в нём настоящие данные.", "ok") };
  $("keyfile-text").onchange = e => loadKey(e.target.files[0], "m-text", "text");
  $("back").onclick = () => { if (!$("ans").value.trim()) return say("m-text", "Вставьте ответ ИИ.", "bad"); if (!needKey("m-text", "text")) return; $("res").value = restoreKeys($("ans").value, M.text); say("m-text", "Готово: настоящие значения возвращены.", "ok") };

  // ---- Excel ----
  const KIND_LABEL = { "ЧЕЛОВЕК": "ФИО", "ТЕЛЕФОН": "Телефон", "EMAIL": "Почта", "ОРГ": "Организация", "ЗНАЧЕНИЕ": "Другое" };
  const MODE_LABEL = { keep: "Оставить как есть", key: "Заменить на коды", drop: "Убрать колонку", scan: "Найти и скрыть внутри текста" };
  function showCols() {
    X.cols = detectColumns(X.rows, $("xl-head").checked);
    const tb = $("xl-cols").tBodies[0]; tb.textContent = "";
    for (const c of X.cols) {
      const tr = tb.insertRow(), c0 = tr.insertCell(); c0.textContent = c.name; c0.dataset.l = "Колонка";
      const pv = tr.insertCell(); pv.className = "pv"; pv.dataset.l = "Примеры"; pv.textContent = c.preview.join(" · ");
      const mk = (opts, cur, key, label) => { const td = tr.insertCell(), s = document.createElement("select"); td.dataset.l = label; s.setAttribute("aria-label", label + ": " + c.name);
        for (const [v, t] of Object.entries(opts)) { const o = new Option(t, v); s.add(o) } s.value = cur; s.onchange = () => { c[key] = s.value }; td.append(s) };
      mk(KIND_LABEL, c.kind, "kind", "Что это"); mk(MODE_LABEL, c.mode, "mode", "Что сделать");
    }
    $("xl-go").disabled = false; $("xl-dl").disabled = $("xl-key").disabled = true; $("xl-prev").textContent = "";
  }
  function loadSheet(i) { X.sheetIdx = i; X.rows = X.sheets[i].rows; showCols() }
  async function openTable(file) {
    try {
      const t = await readTable(file); Object.assign(X, { sheets: t.sheets, csv: t.csv, delim: t.delim || ";", name: baseName(file.name) });
      const s = $("xl-sheet"); s.textContent = ""; t.sheets.forEach((sh, i) => s.add(new Option(sh.name, i)));
      $("xl-sheetbox").hidden = t.sheets.length < 2;
      $("xl-info").textContent = t.csv ? "Файл CSV, строк: " + t.sheets[0].rows.length : "Листов: " + t.sheets.length + (t.sheets.length > 1 ? ". Обработается выбранный лист, остальные в копию не попадут." : "");
      loadSheet(0);
    } catch (e) { $("xl-go").disabled = true; fail("m-xl", e) }
  }
  $("xl-file").onchange = e => e.target.files[0] && openTable(e.target.files[0]);
  $("xl-sheet").onchange = e => loadSheet(+e.target.value);
  $("xl-head").onchange = () => X.rows && showCols();
  const prevTable = (rows, title) => { const w = document.createElement("div"); const h = document.createElement("div"); h.className = "small"; h.textContent = title; const tw = document.createElement("div"); tw.className = "tw"; const t = document.createElement("table");
    rows.slice(0, 6).forEach((r, i) => { const tr = t.insertRow(); r.forEach(v => { const c = document.createElement(i === 0 && $("xl-head").checked ? "th" : "td"); c.textContent = v; tr.append(c) }) }); tw.append(t); w.append(h, tw); return w };
  $("xl-go").onclick = () => {
    try {
      const r = anonymizeTable(X.rows, X.cols, { hasHeader: $("xl-head").checked, dict: dict() }); M.xl = r.map; X.result = r.rows;
      const p = $("xl-prev"); p.textContent = ""; const g = document.createElement("div"); g.className = "two"; g.append(prevTable(X.rows, "Было (первые строки)"), prevTable(r.rows, "Стало: это уйдёт в ИИ")); p.append(g);
      $("xl-dl").disabled = $("xl-key").disabled = false; say("m-xl", "Готово. Заменено значений: " + counts(r.counts) + ". Посмотрите таблицу «Стало» и скачайте копию." + (r.warn.length ? "\nВнимание: " + r.warn.join("; ") : ""), r.warn.length ? "warn" : "ok");
    } catch (e) { fail("m-xl", e) }
  };
  $("xl-dl").onclick = async () => { try { const n = X.name + "-безопасная-копия"; if (X.csv) download(n + ".csv", toCsv(X.result, X.delim), TYPES.csv); else download(n + ".xlsx", await writeXlsx(X.result, X.sheets[X.sheetIdx].name), TYPES.xlsx) } catch (e) { fail("m-xl", e) } };
  $("xl-key").onclick = () => { if (saveKey("xl")) say("m-xl", "Ключ сохранён. Держите файл у себя: в нём настоящие данные.", "ok") };
  $("keyfile-xl").onchange = e => loadKey(e.target.files[0], "m-xl", "xl");
  $("xl-back").onclick = async () => {
    if (!needKey("m-xl", "xl")) return;
    try {
      const f = $("xl-ansfile").files[0];
      if (f) { const t = await readTable(f), rows = restoreRows(t.sheets[0].rows, M.xl), n = baseName(f.name) + "-с-именами";
        if (t.csv) download(n + ".csv", toCsv(rows, t.delim), TYPES.csv); else download(n + ".xlsx", await writeXlsx(rows, t.sheets[0].name), TYPES.xlsx); say("m-xl", "Готово: файл с настоящими значениями скачан.", "ok") }
      else if ($("xl-anstext").value.trim()) { const r = $("xl-res"); r.hidden = false; r.value = restoreKeys($("xl-anstext").value, M.xl); say("m-xl", "Готово: настоящие значения возвращены в тексте.", "ok") }
      else say("m-xl", "Выберите файл с ответом ИИ или вставьте текст ответа.", "bad");
    } catch (e) { fail("m-xl", e) }
  };
  async function fakeTableExample() {
    const rows = [["ФИО", "Телефон", "Email", "Отдел", "Сумма", "Дата"]];
    const names = ["Сидоров Иван Петрович", "Козлова Мария Сергеевна", "Иванов Пётр Андреевич", "Петрова Анна Ивановна", "Смирнов Олег Николаевич"], deps = ["Продажи", "Закупки", "ИТ"];
    for (let i = 0; i < 16; i++) { const n = names[i % 5]; rows.push([n, "+375 29 " + (100 + i % 5) + "-" + (10 + i % 5) + "-" + (20 + i % 5), "user" + (i % 5) + "@example.test", deps[i % 3], 1000 + i * 137, "2026-10-" + String(1 + i).padStart(2, "0")]) }
    Object.assign(X, { sheets: [{ name: "Лист1", rows }], csv: false, name: "пример" }); $("xl-sheetbox").hidden = true; $("xl-info").textContent = "Пример: вымышленная таблица, 16 строк"; loadSheet(0);
  }
  $("ex-xl").onclick = fakeTableExample;

  // ---- Word ----
  async function openDoc(file) { try { D.buf = await file.arrayBuffer(); D.name = baseName(file.name); $("doc-go").disabled = false; $("doc-dl").disabled = $("doc-key").disabled = true; say("doc-rep", "Файл выбран: " + file.name + ". Нажмите «Скрыть данные».") } catch (e) { fail("doc-rep", e) } }
  $("doc-file").onchange = e => e.target.files[0] && openDoc(e.target.files[0]);
  $("doc-go").onclick = async () => {
    try {
      const r = await anonymizeDocx(D.buf, dict()); M.doc = r.map; D.result = r.data;
      const sample = r.preview.map(([a, b]) => "Было: " + a.slice(0, 160) + "\nСтало: " + b.slice(0, 160)).join("\n\n");
      say("doc-rep", "Заменено: " + counts(r.counts) + ". Изменено абзацев: " + r.changed + (r.warn.length ? "\nВнимание: " + r.warn.join("; ") : "") + (sample ? "\n\nНапример:\n" + sample : "") + "\n\nОткройте скачанную копию в Word и проверьте глазами.", r.warn.length ? "warn" : "ok");
      $("doc-dl").disabled = $("doc-key").disabled = false;
    } catch (e) { fail("doc-rep", e) }
  };
  $("doc-dl").onclick = () => download(D.name + "-безопасная-копия.docx", D.result, TYPES.docx);
  $("doc-key").onclick = () => { if (saveKey("doc")) say("m-doc", "Ключ сохранён. Держите файл у себя: в нём настоящие данные.", "ok") };
  $("keyfile-doc").onchange = e => loadKey(e.target.files[0], "m-doc", "doc");
  $("doc-back").onclick = async () => {
    if (!needKey("m-doc", "doc")) return; const f = $("doc-ansfile").files[0]; if (!f) return say("m-doc", "Выберите файл .docx с ответом ИИ.", "bad");
    try { download(baseName(f.name) + "-с-именами.docx", await restoreDocx(await f.arrayBuffer(), M.doc), TYPES.docx); say("m-doc", "Готово: файл с настоящими значениями скачан.", "ok") } catch (e) { fail("m-doc", e) }
  };
  $("ex-doc").onclick = async () => {
    const data = await makeDocx({ author: "Сидоров И. П.", paragraphs: ["Договор поставки", "Заказчик: ООО «Ромашка», контактное лицо Иван Петрович Сидоров, тел. +375 29 123-45-67, почта ivan.sidorov@example.test.", "Бухгалтер Козлова Мария Сергеевна подготовит акт до конца месяца. Сидорову нужно подписать его первым."], table: [["Сотрудник", "Телефон"], ["Пётр Иванов", "+375 44 765-43-21"]] });
    D.buf = data; D.name = "пример"; $("doc-go").disabled = false; await $("doc-go").onclick();
  };
})();
