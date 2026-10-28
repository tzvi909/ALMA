#!/usr/bin/env bash
# Run all ALMA checks.  Needs the venv (.venv) and, for the VBA run, docker with the
# alma-libreoffice image:  docker build -t alma-libreoffice tests/docker
set -euo pipefail
cd "$(dirname "$0")/.."
PY=.venv/bin/python
OUT=${ALMA_TEST_OUT:-$PWD/test_output}
rm -rf "$OUT"; mkdir -p "$OUT"
echo "== VBA project writer (olefile / oletools)";  ALMA_TMP="$OUT" $PY tests/test_vbaproject.py
echo "== VBA static checks";                         $PY tests/vba_static_check.py
echo "== build Excel files";                         $PY excel/build_excel.py
echo "== xlwings script, file mode (openpyxl)";      $PY tests/xlwings_file_mode.py "$OUT/xlw" > "$OUT/xlw.log"
$PY tests/compare_vba_python.py "$OUT/xlw/sheets" results/snapshot_2020-02-10/csv
echo "== xlwings script, Excel code path (fake xlwings)"; $PY tests/xlwings_adapter_test.py "$OUT/xlw2" > "$OUT/xlw2.log"
$PY tests/compare_vba_python.py "$OUT/xlw2/sheets" results/snapshot_2020-02-10/csv
if docker image inspect alma-libreoffice >/dev/null 2>&1; then
  echo "== VBA macros in LibreOffice (VBA compatibility mode)"
  mkdir -p "$OUT/lo"
  docker run --rm --user "$(id -u):$(id -g)" -e HOME=/out -v "$PWD":/work:ro -v "$OUT/lo":/out alma-libreoffice \
      /usr/bin/python3 /work/tests/lo_vba_harness.py /work/excel/ALMA.xlsm /out > "$OUT/lo.log" 2>&1
  grep -E "self|ERRORS" "$OUT/lo.log"
  $PY tests/compare_vba_python.py "$OUT/lo/sheets" results/snapshot_2020-02-10/csv
  $PY tests/check_vba_urls.py "$OUT/lo/vba_urls.json"
else
  echo "(skipping the LibreOffice VBA run: docker image alma-libreoffice not built)"
fi
