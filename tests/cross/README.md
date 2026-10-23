# xlspy cross-validation tests

Two-way interoperability tests between this TypeScript library and
[xlspy](https://github.com/justybase/xlspy) (Python):

- `xlspy_bridge.py` — JSON CLI over xlspy (`check` / `write` / `read` /
  `replace-sheet-data`). Node drives it via `child_process.spawnSync`.
- `xlspy-cross.test.ts` — the suite (run with `npm run test:cross`).

Coverage: golden fixtures authored by openpyxl/xlsxwriter (incl. the 1904
date system), TS-write → xlspy-read and xlspy-write → TS-read of a shared
value matrix (unicode, pre-1900 dates, gap rows, empty/unicode sheets,
streaming writes), updater round-trips in both directions for XLSX and XLSB,
and ListObject table growth across implementations.

## Requirements

- Python 3.12+ with the `xlspy` package importable
  (`pip install xlspy`, or `PYTHONPATH=/path/to/xlspy/src`).
- Python binary: `XLSPY_PYTHON` (default `python3`).
- Golden-fixture tests additionally need the xlspy repo checkout: `XLSPY_REPO`
  (default: `../xlspy` next to this repo).

If xlspy is unavailable the suite prints `SKIPPED` and exits 0, so `npm test`
stays green without Python. Missing repo checkout skips only the golden tests.
