#!/bin/sh
# Нужен только Node.js. Каждый файл падает с ошибкой, если проверка не прошла.
set -e
cd "$(dirname "$0")"
for t in tests/roundtrip.js tests/regress.js tests/t_org.js tests/t_person.js tests/t_secret.js tests/t_caps.js tests/t_review1.js tests/t_review2.js tests/t_office.js tests/t_review3.js tests/t_review4.js tests/t_review5.js tests/t_build.js; do node "$t"; done
echo "все проверки прошли"
