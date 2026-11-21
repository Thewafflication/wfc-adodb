# wfc-adodb instructions

- Tests: `ctest` on the x64 build only; GitHub Actions covers other targets.
- `third_party/sqlite` is a vendored upstream release; do not edit it, replace
  it whole and update `third_party/sqlite/VERSION` and the README hash.
- The public surface is `include/wfc/adodb/adodb.hpp`; keep it mirroring the
  ADODB members so the WFC binding stays a thin table.
