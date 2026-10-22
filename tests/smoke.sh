#!/usr/bin/env bash
# Full local check: python backend, extension tool layer, TypeScript types.
set -euo pipefail
cd "$(dirname "$0")/.."

echo "== python syntax =="
python3 -m py_compile python/pptkit.py

echo "== backend tests =="
python3 tests/test_backend.py

echo "== extension tests =="
if [ -d node_modules ]; then
	node --test tests/extension.test.ts
else
	echo "skip (run npm install first)"
fi

echo "== typecheck =="
if [ -d node_modules ]; then
	npx tsc --noEmit
else
	echo "skip (run npm install first)"
fi

echo "OK"
