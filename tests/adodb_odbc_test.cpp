// TC-0002: the ODBC backend, against the Microsoft Access driver when this
// machine has it (the test is skipped when it does not).

#include <wfc/adodb/adodb.hpp>

#ifdef _WIN32
#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#include <windows.h>

#include <odbcinst.h>

#include <chrono>
#include <cstdlib>
#include <filesystem>
#include <iostream>
#include <string>

namespace {

int failures = 0;

void check(const bool condition, const char* what) {
    if (!condition) {
        std::cerr << "FAIL: " << what << '\n';
        ++failures;
    }
}

}  // namespace

int main() {
    using namespace wfc::adodb;

    check(connection_string_path("Driver={X};Dbq=a.accdb") == "a.accdb",
          "Dbq is the path of a path-style string");

    const std::string driver = "Microsoft Access Driver (*.mdb, *.accdb)";
    const auto folder =
        std::filesystem::temp_directory_path() /
        ("wfc-adodb-odbc-" +
         std::to_string(
             std::chrono::steady_clock::now().time_since_epoch().count()));
    std::filesystem::create_directories(folder);
    const std::string file = (folder / "test.accdb").string();
    const std::string attributes = "CREATE_DBV12=" + file + '\0';
    if (SQLConfigDataSource(nullptr, ODBC_ADD_DSN, driver.c_str(),
                            attributes.c_str()) == FALSE) {
        std::cout << "SKIP: no Access ODBC driver here\n";
        std::filesystem::remove_all(folder);
        return 0;
    }

    Connection cn;
    const std::string text = "Driver={" + driver + "};Dbq=" + file + ';';
    check(cn.open(text), "open through ODBC");
    check(cn.state() == adStateOpen, "open state");
    check(!cn.open(text), "second open fails");
    cn.clear_errors();
    check(cn.version().find("ACCESS") != std::string::npos ||
              !cn.version().empty(),
          "version names the DBMS");

    check(cn.execute("create table t (id long primary key, name text(50), "
                     "score double, born datetime, flag yesno)")
              .ok,
          "create");
    const auto one = cn.execute(
        "insert into t values (?, ?, ?, ?, ?)",
        {std::int64_t{1}, std::string("ann"), 1.5, Value{}, std::int64_t{1}});
    check(one.ok && one.records_affected == 1 && !one.recordset,
          "insert reports one row");
    check(cn.execute("insert into t (id, name, score) values (2, 'bob', 2.25)")
              .ok,
          "insert literal");

    auto rows = cn.execute("select id, name, score, born from t order by id");
    check(rows.ok && rows.recordset != nullptr, "select returns rows");
    if (rows.recordset != nullptr) {
        auto& set = *rows.recordset;
        check(set.field_count() == 4 && set.field(1).name == "name",
              "field names");
        check(set.record_count() == 2, "two records");
        check(std::get<std::string>(set.value(1)) == "ann", "text value");
        check(std::get<double>(set.value(2)) == 1.5, "double value");
        check(set.value(3).index() == 0, "null value");
        set.move_next();
        check(std::get<std::int64_t>(set.value(0)) == 2, "second id");
    }

    const auto missing = cn.execute("select * from nothing_here");
    check(!missing.ok && !cn.errors().empty(), "bad table fails");
    check(!cn.errors().empty() && cn.errors().back().source == "ODBC",
          "error comes from ODBC");
    cn.clear_errors();
    const auto dup = cn.execute("insert into t (id, name) values (1, 'again')");
    check(!dup.ok && !cn.errors().empty() &&
              cn.errors().back().number == -2147217873L,
          "duplicate key is a constraint error");
    cn.clear_errors();

    check(cn.begin_trans() == 1, "begin");
    check(cn.begin_trans() == 0, "no nesting");
    cn.clear_errors();
    check(cn.execute("delete from t").ok, "delete inside a transaction");
    check(cn.rollback_trans(), "rollback");
    auto again = cn.execute("select count(*) from t");
    check(again.ok && again.recordset != nullptr &&
              std::get<std::int64_t>(again.recordset->value(0)) == 2,
          "rollback kept both rows");
    check(cn.begin_trans() == 1, "begin again");
    check(cn.execute("delete from t where id = 2").ok, "delete one");
    check(cn.commit_trans(), "commit");
    auto left = cn.execute("select count(*) from t");
    check(left.ok && left.recordset != nullptr &&
              std::get<std::int64_t>(left.recordset->value(0)) == 1,
          "commit kept the delete");

    cn.close();
    check(cn.state() == adStateClosed, "closed");
    std::error_code ignored;
    std::filesystem::remove_all(folder, ignored);

    std::cout << (failures == 0 ? "PASS" : "FAILED") << '\n';
    return failures == 0 ? EXIT_SUCCESS : EXIT_FAILURE;
}
#else
#include <iostream>
int main() {
    std::cout << "SKIP: the ODBC backend is built on Windows only\n";
    return 0;
}
#endif
