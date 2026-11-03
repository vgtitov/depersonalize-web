"""Прогон страницы в настоящем браузере (Playwright, Chromium). Не входит в run_tests.sh: нужен браузер.
python3 tests/ui_check.py [каталог для скриншотов]"""
import sys, pathlib, re
from playwright.sync_api import sync_playwright

root = pathlib.Path(__file__).resolve().parent.parent
shots = pathlib.Path(sys.argv[1] if len(sys.argv) > 1 else "/tmp/depers-shots"); shots.mkdir(parents=True, exist_ok=True)
errors, requests = [], []
def check(cond, msg):
    print(("ok   " if cond else "FAIL ") + msg)
    if not cond: errors.append(msg)

with sync_playwright() as p:
    b = p.chromium.launch(); ctx = b.new_context(viewport={"width": 1100, "height": 900}, accept_downloads=True); pg = ctx.new_page()
    pg.on("request", lambda r: requests.append(r.url)); pg.on("pageerror", lambda e: errors.append("JS: " + str(e)))
    pg.on("console", lambda m: m.type == "error" and errors.append("console: " + m.text))
    pg.goto((root / "index.html").as_uri())
    check("Данные не уходят" in pg.inner_text("body"), "постоянная пометка про браузер видна")
    # Тема
    bg0 = pg.evaluate("getComputedStyle(document.body).backgroundColor"); pg.click("#theme")
    bg1 = pg.evaluate("getComputedStyle(document.body).backgroundColor"); th = pg.evaluate("document.documentElement.dataset.theme")
    check(bg0 != bg1 and th in ("dark", "light"), f"переключатель темы меняет фон: {bg0} -> {bg1} ({th})")
    pg.click("#theme"); check(pg.evaluate("getComputedStyle(document.body).backgroundColor") == bg0, "второй щелчок возвращает тему")
    # Текст
    pg.click("#ex-text"); out = pg.input_value("#out")
    check("[ЧЕЛОВЕК_1]" in out and "Сидоров" not in out and "Qwerty123" not in out and "[ОРГ_1]" in out, "текст: пример обезличен: " + out[:90].replace("\n", " "))
    pg.fill("#ans", "Свяжитесь с [ЧЕЛОВЕК_1] по [ТЕЛЕФОН_1]"); pg.click("#back")
    check(pg.input_value("#res").startswith("Свяжитесь с Иван Петрович Сидоров по +375"), "текст: имена возвращены")
    with pg.expect_download() as d: pg.click("#savekey-text")
    check(d.value.suggested_filename.endswith(".json"), "текст: ключ скачивается")
    pg.screenshot(path=str(shots / "1-text.png"), full_page=True)
    # Excel
    pg.click("#t-xl"); pg.click("#ex-xl")
    kinds = pg.eval_on_selector_all("#xl-cols tbody tr", "rs=>rs.map(r=>[...r.querySelectorAll('select')].map(s=>s.value).join('/'))")
    check(kinds[:3] == ["ЧЕЛОВЕК/key", "ТЕЛЕФОН/key", "EMAIL/key"], "excel: типы колонок угаданы: " + str(kinds))
    pg.click("#xl-go"); txt = pg.inner_text("#xl-prev")
    check("ЧЕЛОВЕК_001" in txt and "Сидоров" in txt.split("Стало")[0] and "Сидоров" not in txt.split("Стало")[1], "excel: до и после показаны")
    with pg.expect_download() as d: pg.click("#xl-dl")
    check(d.value.suggested_filename == "таблица-безопасная-копия.xlsx", "excel: имя копии нейтральное: " + d.value.suggested_filename)
    xl = pathlib.Path("/tmp/depers-copy.xlsx"); d.value.save_as(xl)
    pg.screenshot(path=str(shots / "2-excel.png"), full_page=True)
    pg.set_input_files("#xl-ansfile", str(xl))
    with pg.expect_download() as d2: pg.click("#xl-back")
    back = pathlib.Path("/tmp/depers-back.xlsx"); d2.value.save_as(back)
    import zipfile; z = zipfile.ZipFile(back).read("xl/worksheets/sheet1.xml").decode()
    check("Сидоров Иван Петрович" in z and "ЧЕЛОВЕК_0" not in z, "excel: имена вернулись в файл")
    check("Сидоров" not in zipfile.ZipFile(xl).read("xl/worksheets/sheet1.xml").decode(), "excel: в безопасной копии имён нет")
    # Word
    pg.click("#t-doc"); pg.click("#ex-doc"); pg.wait_for_selector("#doc-dl:not([disabled])"); rep = pg.inner_text("#doc-rep")
    check("Заменено" in rep and "ЧЕЛОВЕК" in rep, "word: пример обезличен")
    with pg.expect_download() as d: pg.click("#doc-dl")
    dx = pathlib.Path("/tmp/depers-copy.docx"); d.value.save_as(dx)
    zz = zipfile.ZipFile(dx); allx = "".join(zz.read(n).decode("utf8", "ignore") for n in zz.namelist())
    check(not re.search("Сидоров|Козлова|Ромашка|123-45-67|example.test", allx), "word: в безопасной копии данных нет (включая автора)")
    pg.set_input_files("#doc-ansfile", str(dx))
    with pg.expect_download() as d2: pg.click("#doc-back")
    d2.value.save_as("/tmp/depers-back.docx"); zb = zipfile.ZipFile("/tmp/depers-back.docx").read("word/document.xml").decode()
    check("Иван Петрович Сидоров" in zb and "Ромашка" in zb, "word: имена вернулись")
    pg.screenshot(path=str(shots / "3-word.png"), full_page=True)
    # клавиатура и узкий экран
    pg.click("#t-text"); pg.focus("#t-text"); pg.keyboard.press("ArrowRight")
    check(pg.get_attribute("#t-xl", "aria-selected") == "true", "клавиатура: стрелки переключают вкладки")
    pg.set_viewport_size({"width": 375, "height": 800}); pg.click("#t-xl")
    check(pg.evaluate("document.documentElement.scrollWidth <= window.innerWidth + 1"), "узкий экран: нет горизонтальной прокрутки страницы")
    print("шире окна:", pg.evaluate("[...document.querySelectorAll('body *')].filter(e=>e.getBoundingClientRect().right>window.innerWidth+1&&!e.closest('.tw')).slice(0,5).map(e=>e.tagName+'#'+e.id+'.'+e.className)"))
    pg.screenshot(path=str(shots / "4-mobile.png"), full_page=True)
    # тёмная тема
    pg.emulate_media(color_scheme="dark"); pg.screenshot(path=str(shots / "5-dark.png"))
    b.close()
check(all(u.startswith("file:") for u in requests), "сеть: запросов за пределы файла нет: " + str([u for u in requests if not u.startswith("file:")]))
print("\nОШИБКИ: " + str(errors) if errors else "\nвсе проверки интерфейса прошли")
sys.exit(1 if errors else 0)
