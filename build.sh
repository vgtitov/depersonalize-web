#!/bin/sh
# Вставляет vendor/jszip.min.js, engine.js, office.js и ui.js в index.html между <!--BUILD--> и <!--/BUILD-->: страница остаётся одним файлом без внешних скриптов.
set -e
cd "$(dirname "$0")"
python3 - <<'PY'
import re
h=open("index.html",encoding="utf-8").read()
parts=[open(p,encoding="utf-8").read() for p in ("vendor/jszip.min.js","engine.js","office.js","ui.js")]
blk="<!--BUILD-->\n"+"".join("<script>\n"+p.replace("</script","<\\/script")+"\n</script>\n" for p in parts)+"<!--/BUILD-->"
h=re.sub(r"<!--BUILD-->.*?<!--/BUILD-->",lambda m:blk,h,flags=re.S)
open("index.html","w",encoding="utf-8").write(h)
PY
