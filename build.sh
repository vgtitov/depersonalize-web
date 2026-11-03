#!/bin/sh
# Вставляет engine.js в index.html между <!--ENGINE--> и <!--/ENGINE-->, чтобы страница оставалась одним файлом без внешних скриптов.
set -e
cd "$(dirname "$0")"
python3 - <<'PY'
import re
h=open("index.html",encoding="utf-8").read()
e=open("engine.js",encoding="utf-8").read()
h=re.sub(r"<!--ENGINE-->.*?<!--/ENGINE-->",lambda m:"<!--ENGINE-->\n<script>\n"+e+"</script>\n<!--/ENGINE-->",h,flags=re.S)
open("index.html","w",encoding="utf-8").write(h)
PY
