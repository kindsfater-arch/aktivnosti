#!/usr/bin/env bash
# Собирает dist/index.html для GitHub Pages: тот же интерфейс, что и в Apps Script,
# но запросы идут через fetch на адрес веб-приложения (API_URL = https://script.google.com/macros/s/…/exec).
set -euo pipefail
: "${API_URL:?Укажите API_URL — адрес веб-приложения Apps Script (…/exec)}"
cd "$(dirname "$0")/.."
mkdir -p dist
python3 - "$API_URL" <<'EOF'
import json, re, sys
src = open('apps-script/Index.html', encoding='utf-8').read()
boot = json.dumps({'mode': 'web', 'apiUrl': sys.argv[1]})
out, n = re.subn(r'/\*BOOT\*/.*?/\*END\*/', lambda m: boot, src, count=1)
if n != 1:
    sys.exit('Не найден маркер /*BOOT*/ в Index.html')
open('dist/index.html', 'w', encoding='utf-8').write(out)
print('dist/index.html собран, API_URL =', sys.argv[1])
EOF
