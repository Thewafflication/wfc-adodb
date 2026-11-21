# wfc-adodb

ADODB for WFC, implemented as an interposer over SQLite. A Visual Basic 6
program that asks for `ADODB.Connection`, `ADODB.Command` or `ADODB.Recordset`
is answered from a SQLite database instead of the Microsoft data access
components, so it runs the same on every platform WFC runs on.

SQLite 3.54.1 is vendored as the public-domain amalgamation in
`third_party/sqlite` (SHA3-256 of the source archive, as published on
sqlite.org: `92281c86e2957f9b9305890f9ac04204caa7b677cc868dc326d942689df28e30`).

## Status (0.1.0)

The C++ core is in place and tested. WFC binds it (src/interpreter/ado_helper.cpp
and the VB classes in builtin_class_sources.hpp): when a program uses
`ADODB.Connection` or `ADODB.Command`, `CreateObject` and `As ADODB.*` give the
SQLite-backed classes in the interpreter and in compiled programs. A program that
uses only `ADODB.Recordset` (a disconnected recordset with `Fields.Append`) still
goes to COM until disconnected recordsets are implemented here.

| ADODB member | State |
| --- | --- |
| `Connection.Open`, `Close`, `State`, `Version` | done (`Data Source=`, `Database=`, `File Name=`, `DBQ=` or a bare path; `:memory:`) |
| `Connection.Execute` with `?` parameters, `RecordsAffected` | done |
| `BeginTrans`, `CommitTrans`, `RollbackTrans` (nested, as savepoints) | done |
| `Connection.Errors` | done (ADO numbers for syntax errors and constraint violations) |
| `Recordset` read-only static cursor: `Fields`, `MoveFirst/Last/Next/Previous`, `BOF`, `EOF`, `RecordCount`, `AbsolutePosition` | done |
| `Recordset.AddNew`, `Update`, `Delete`, `Find`, `Filter`, `Sort` | planned |
| `Command` object, `Parameters`, named parameters | planned |
| `CreateObject("ADODB.*")` binding in the interpreter and in native programs | done for Connection, Command (`Execute , Array(...)`), read-only Recordset |
| ADO type mapping beyond Integer, BigInt, Double, VarWChar, LongVarBinary; Date and Boolean | planned |

## Build and test

```powershell
cmake -S . -B build -G "Visual Studio 17 2022" -A x64
cmake --build build --config Debug
ctest --test-dir build -C Debug --output-on-failure
```

WFC mounts the library with `add_subdirectory(lib/wfc-adodb)`; the tests then
carry the prefix `wfc-adodb-`.
