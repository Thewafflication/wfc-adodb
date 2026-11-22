# wfc-adodb

ADODB for WFC, implemented as an interposer over SQLite. A Visual Basic 6
program that asks for `ADODB.Connection`, `ADODB.Command` or `ADODB.Recordset`
is answered from a SQLite database instead of the Microsoft data access
components, so it runs the same on every platform WFC runs on.

SQLite 3.54.1 is vendored as the public-domain amalgamation in
`third_party/sqlite` (SHA3-256 of the source archive, as published on
sqlite.org: `92281c86e2957f9b9305890f9ac04204caa7b677cc868dc326d942689df28e30`).

## Status (0.2.0)

The C++ core is complete for its scope and tested. WFC binds it
(src/interpreter/ado_helper.cpp and the VB classes in
builtin_class_sources.hpp): when a program uses `ADODB.Connection`,
`ADODB.Command` or `ADODB.Recordset`, `CreateObject` and `As ADODB.*` give the
SQLite-backed classes in the interpreter and in compiled programs.

| ADODB member | State |
| --- | --- |
| `Connection.Open`, `Close`, `State`, `Version` | done (`Data Source=`, `Database=`, `File Name=`, `DBQ=` or a bare path; `:memory:`) |
| `Connection.Open` through ODBC (`DSN=`, `FILEDSN=`, `Driver={...}` or `Provider=MSDASQL`): MySQL, SQL Server, Postgres, Access, anything with an ODBC driver | done on Windows (odbc32) and in Winelib builds (Wine's odbc32 over unixODBC); tested with the Access and SQLite drivers, not against MySQL, SQL Server or Postgres servers. The SQL is the database's own; `BeginTrans` does not nest |
| `Connection.Execute` with `?` parameters, `RecordsAffected` | done |
| `select @@identity` | done on every backend (SQLite `last_insert_rowid`; under ODBC the DBMS's own identity query: `@@IDENTITY`, `LAST_INSERT_ID()`, `lastval()`; none for Oracle) |
| `BeginTrans`, `CommitTrans`, `RollbackTrans` (nested, as savepoints) | done (SQLite) |
| `Connection.Errors` | done (ADO numbers for syntax errors and constraint violations) |
| `Command`: `CommandText`, `ActiveConnection`, `Execute`, `Parameters` (`Append`, `Count`, `Item`, `Delete`), `CreateParameter`; parameters bind to `?` in order | done; named parameters are not supported |
| `Recordset` static cursor: `Fields` (`Name`, `Type`, `DefinedSize`, `Value`), `MoveFirst/Last/Next/Previous`, `Move`, `BOF`, `EOF`, `RecordCount`, `AbsolutePosition` (get and set), `Bookmark`, `Filter`, `Find`, `Sort`, `GetRows`, `GetString`, `#date#` literals | done |
| Disconnected `Recordset` (`Fields.Append`, `Open` with no source, `AddNew`, `Update`, `CancelUpdate`, `Delete`) | done |
| Updating a query result: `rs.Open "select ... from t", cn, , adLockOptimistic` (a bare table name works too) writes `Update`, `AddNew` and `Delete` to the table at once | done for a SELECT of plain columns from one table (WHERE, ORDER BY, LIMIT allowed); joins, expressions, grouping and views stay read-only; `cn.Execute` results stay read-only, as in ADO; SQLite only |
| `CreateObject("ADODB.*")` binding in the interpreter and in native programs | done |
| Batch updates (`UpdateBatch`), `Clone`, `Save`/`Load`, shaped recordsets, server cursors | not planned |

## Build and test

```powershell
cmake -S . -B build -G "Visual Studio 17 2022" -A x64
cmake --build build --config Debug
ctest --test-dir build -C Debug --output-on-failure
```

WFC mounts the library with `add_subdirectory(lib/wfc-adodb)`; the tests then
carry the prefix `wfc-adodb-`.
