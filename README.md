# wfc-adodb

ADODB for WFC, implemented as an interposer over SQLite. A Visual Basic 6
program that asks for `ADODB.Connection`, `ADODB.Command` or `ADODB.Recordset`
is answered from a SQLite database instead of the Microsoft data access
components, so it runs the same on every platform WFC runs on.

SQLite 3.54.1 is vendored as the public-domain amalgamation in
`third_party/sqlite` (SHA3-256 of the source archive, as published on
sqlite.org: `92281c86e2957f9b9305890f9ac04204caa7b677cc868dc326d942689df28e30`).

## Status (0.4.0)

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
| `msado15.dll`, a COM server (see below) | done: Connection, Recordset, Command, Fields, Field, Parameters, Parameter, late-bound |
| Batch updates (`UpdateBatch`), `Clone`, `Save`/`Load`, shaped recordsets, server cursors | not planned |

## The COM server (`msado15.dll`)

With `-DWFC_ADODB_SHARED=ON` the library is also built as `msado15.dll`, a COM
server that answers `ADODB.Connection`, `ADODB.Recordset` and `ADODB.Command`
from the library, so any COM client (a Visual Basic 6 program, VBScript, a
script host) gets the SQLite-backed ADODB. `regsvr32 msado15.dll` registers it
under WFC's own ProgIDs (`WFC.ADODB.Connection` and so on) and class ids;
with `WFC_REGISTER_MICROSOFT_IDS=1` it also serves the Microsoft ProgIDs
(`ADODB.Connection`) and what was registered there before comes back with the
unregistration. `WFC_REGISTER_SCOPE=user|machine` picks the registry.

The classes are `Connection`, `Recordset`, `Command`, `Parameter`, `Record`
and `Stream`, with the `Fields`, `Field`, `Parameters`, `Properties`,
`Property`, `Errors` and `Error` objects they hand out. Every member of the
ADO 6.1 object model is there (the members of each object are the ones of
`src/adodb.idl`, with the DISPIDs and parameters of ADO), and the ones that
cannot mean anything over SQLite answer error 3251.

The DLL carries a type library (`ADODB`, version 6.1, with the enumerations
such as `adOpenKeyset` and `adInteger`), registered by `regsvr32`. The
interfaces in it are plain dispinterfaces, so a Visual Basic 6 project that
references it should get `Dim cn As ADODB.Connection`, `New ADODB.Recordset`,
the constants and the member lists, with calls going through `IDispatch` (VB6
is a 32-bit program and needs the x86 build of the DLL; this has been tried
with VBScript and a C++ client, not yet with VB6 itself). What it does
not offer is the vtable of ADO's dual interfaces: a program built against
Microsoft's `msado15.dll` type library calls through the vtable and cannot use
this one, and there are no connection-point events (`WithEvents`).

Behaviour worth knowing: `Errors` and `Properties` (a few properties of a
connection) work; `Clone`, `Requery`, paging (`PageSize`, `AbsolutePage`,
`PageCount`), `EditMode`, `Status`, `GetRows`, `GetString` and `Find` work;
`UpdateBatch` is a no-op because changes are applied at once (see updating a
query result, above); `Save` writes ADO's XML format whatever `PersistFormat`
says, and `Open` with `adCmdFile` reads that XML (not the binary ADTG format);
`Connection.OpenSchema` answers tables, columns, primary keys and indexes
over SQLite; `Stream` is complete in memory (text in `Unicode`, `utf-8` and
single-byte character sets, binary, files); a `Record` is the current row of a
recordset or the first row of a query, and has no folder operations.
`Seek`, `Fields.Delete`, stored procedures (`adCmdStoredProc`), `MaxRecords`
(stored but not applied) and named parameters are not supported.

```vbscript
Set cn = CreateObject("WFC.ADODB.Connection")
cn.Open "Data Source=:memory:"
cn.Execute "create table t(id integer primary key, name text)"
Set rs = CreateObject("WFC.ADODB.Recordset")
rs.Open "t", cn, adOpenKeyset, adLockOptimistic
```

The type library is built with midl (Windows SDK) from `src/adodb.idl`. The
IDL was written once from the member lists of the Microsoft type library by
`tools/tlb2idl.cpp` and `tools/make_idl.py` (they need the Microsoft
`msado15.dll`, which is not part of this repository); every identifier in it is
WFC's own.

## Installing with WPM

Each release is a signed WPM package, `wfc-adodb`, with the index of the
release (`index.json`) beside it. Trust the WFC release key once, add the
release as a WPM repository, and install:

```powershell
Invoke-WebRequest `
  https://raw.githubusercontent.com/Thewafflication/wfc/master/release_keys/wfc-release.public `
  -OutFile wfc-release.public
wpm trust add wfc-release.public
wpm repo add https://github.com/Thewafflication/wfc-adodb/releases/latest/download
wpm update
wpm install wfc-adodb
```

The install script puts `msado15.dll` for x86, x64 and ARM64 under
`lib\wfcadodb\<version>\<arch>` of the WFC installation and registers the
ones this Windows can run (`wpm remove wfc-adodb` unregisters them).

## Build and test

```powershell
cmake -S . -B build -G "Visual Studio 17 2022" -A x64
cmake --build build --config Debug
ctest --test-dir build -C Debug --output-on-failure
```

WFC mounts the library with `add_subdirectory(lib/wfc-adodb)`; the tests then
carry the prefix `wfc-adodb-`.
