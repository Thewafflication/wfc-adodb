// TC-0001: the ADODB surface over SQLite.

#include <wfc/adodb/adodb.hpp>

#include <cstdlib>
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

    check(connection_string_path("Provider=SQLite;Data Source=a.db;") == "a.db",
          "path from Data Source");
    check(connection_string_path("x.db") == "x.db", "bare path");
    check(connection_string_path("Provider=Foo").empty(), "no path");

    Connection cn;
    check(cn.state() == adStateClosed, "closed at first");
    check(!cn.execute("select 1").ok, "execute on closed fails");
    check(cn.errors().size() == 1 && cn.errors()[0].number == 3704,
          "closed error 3704");
    cn.clear_errors();

    check(cn.open("Data Source=:memory:"), "open memory");
    check(cn.state() == adStateOpen, "open state");
    check(!cn.open(":memory:"), "second open fails");
    cn.clear_errors();

    check(cn.execute("create table t(id integer primary key, name text, "
                     "score real, pic blob)")
              .ok,
          "create");
    auto ins = cn.execute("insert into t(name, score) values (?, ?)",
                          {std::string("ann"), 1.5});
    check(ins.ok && ins.records_affected == 1 && !ins.recordset,
          "insert reports one row");
    cn.execute("insert into t(name, score) values (?, ?)",
               {std::string("bob"), std::monostate{}});
    check(cn.last_insert_id() == 2, "last insert id");

    auto q = cn.execute("select id, name, score, pic from t order by id");
    check(q.ok && q.recordset && q.recordset->record_count() == 2, "two rows");
    auto& rs = *q.recordset;
    check(rs.field_count() == 4 && rs.field(1).name == "name", "field names");
    check(rs.field(0).type == adInteger && rs.field(1).type == adVarWChar &&
              rs.field(2).type == adDouble &&
              rs.field(3).type == adLongVarBinary,
          "field types from declarations");
    check(rs.field_index("NAME") == 1, "field index ignores case");
    check(!rs.bof() && !rs.eof() && rs.absolute_position() == 1,
          "at first row");
    check(std::get<std::string>(rs.value(1)) == "ann", "first name");
    check(std::get<double>(rs.value(2)) == 1.5, "first score");
    check(rs.move_next() && std::get<std::string>(rs.value(1)) == "bob",
          "second row");
    check(rs.value(2).index() == 0, "Null score");
    check(rs.move_next() && rs.eof() && rs.absolute_position() == -2,
          "EOF after last");
    check(!rs.move_next(), "no move past EOF");
    rs.move_first();
    check(!rs.eof() && rs.absolute_position() == 1, "MoveFirst");
    check(!rs.move_previous() == false && rs.bof(), "MoveFirst then previous");
    rs.move_last();
    check(rs.absolute_position() == 2, "MoveLast");

    auto none = cn.execute("select id from t where id = 99");
    check(none.ok && none.recordset && none.recordset->bof() &&
              none.recordset->eof() && none.recordset->record_count() == 0,
          "empty result is BOF and EOF");

    // Transactions, nested.
    check(cn.begin_trans() == 1, "begin 1");
    cn.execute("insert into t(name) values ('cy')");
    check(cn.begin_trans() == 2, "begin 2");
    cn.execute("insert into t(name) values ('di')");
    check(cn.rollback_trans(), "rollback inner");
    check(cn.commit_trans(), "commit outer");
    auto count = cn.execute("select count(*) as n from t");
    check(std::get<std::int64_t>(count.recordset->value(0)) == 3,
          "inner insert rolled back, outer kept");
    check(!cn.commit_trans(), "commit without transaction fails");
    cn.clear_errors();

    // Errors.
    check(!cn.execute("selec nothing").ok &&
              cn.errors().back().number == -2147217900L,
          "syntax error");
    check(!cn.execute("insert into t(id, name) values (1, 'dup')").ok &&
              cn.errors().back().number == -2147217873L,
          "constraint error");
    check(!cn.execute("select ?").ok, "missing parameter");

    // A disconnected recordset: fields, rows, filter, find, delete.
    auto mem = Recordset::create();
    check(mem->append_field("id", adInteger, 0) &&
              mem->append_field("name", adVarWChar, 20) &&
              !mem->append_field("ID", adInteger, 0),
          "append fields, no duplicates");
    check(mem->state() == adStateClosed && mem->open_empty() &&
              mem->state() == adStateOpen,
          "open empty");
    check(mem->bof() && mem->eof() && mem->record_count() == 0, "starts empty");
    for (std::int64_t i = 1; i <= 4; ++i) {
        check(mem->add_new(), "add_new");
        check(
            mem->set_value(0, Value{i}) &&
                mem->set_value(1, Value{std::string("n") + std::to_string(i)}),
            "set values");
        check(mem->update(), "update");
    }
    check(mem->record_count() == 4 && !mem->eof(), "four rows");
    mem->move_first();
    check(std::get<std::string>(mem->value(1)) == "n1", "first row kept");
    check(mem->set_filter("id > 1 AND id <= 3") && mem->record_count() == 2 &&
              std::get<std::int64_t>(mem->value(0)) == 2,
          "filter with AND");
    check(mem->set_filter("name = 'N4' OR id = 1") && mem->record_count() == 2,
          "filter with OR, case-insensitive text");
    check(mem->set_filter("name like 'n*'") && mem->record_count() == 4,
          "filter with LIKE");
    check(!mem->set_filter("nonsense") && mem->record_count() == 4,
          "bad filter refused");
    check(mem->set_filter("") && mem->record_count() == 4, "filter removed");
    check(mem->find("id = 3") && std::get<std::int64_t>(mem->value(0)) == 3,
          "find");
    check(mem->find("id = 99") && mem->eof(), "find without a hit is EOF");
    mem->move_first();
    check(mem->delete_current() && mem->record_count() == 3 &&
              std::get<std::int64_t>(mem->value(0)) == 2,
          "delete leaves the cursor on the next row");
    mem->add_new();
    mem->set_value(0, Value{std::int64_t{9}});
    check(mem->cancel_update() && mem->record_count() == 3,
          "cancel_update drops the new row");
    check(
        mem->set_sort("id DESC") && std::get<std::int64_t>(mem->value(0)) == 4,
        "sort descending starts at the largest");
    mem->move_last();
    check(std::get<std::int64_t>(mem->value(0)) == 2, "sort descending ends");
    check(mem->set_sort("[name] asc, id") &&
              std::get<std::int64_t>(mem->value(0)) == 2,
          "sort by name ascending");
    check(!mem->set_sort("nosuch") && mem->sort() == "[name] asc, id",
          "unknown sort field refused");
    check(mem->set_sort("") && std::get<std::int64_t>(mem->value(0)) == 2,
          "sort removed");
    {
        auto dates = Recordset::create();
        dates->append_field("d", 7, 0);
        dates->open_empty();
        // 2020-01-01 is OLE day 43831; 2020-06-15 is 43997; with 18:00, .75.
        for (const double serial : {43831.0, 43997.0, 43997.75}) {
            dates->add_new();
            dates->set_value(0, Value{serial});
            dates->update();
        }
        check(dates->set_filter("d > #2020-01-01#") &&
                  dates->record_count() == 2,
              "date literal, ISO form");
        check(dates->set_filter("d >= #6/15/2020#") &&
                  dates->record_count() == 2,
              "date literal, US form");
        check(dates->set_filter("d = #6/15/2020 6:00 PM#") &&
                  dates->record_count() == 1,
              "date literal with a time");
        check(dates->set_filter("d < #2020-06-15 12:00#") &&
                  dates->record_count() == 2,
              "date literal with 24-hour time");
        check(!dates->set_filter("d = #not a date#"), "bad date refused");
    }
    {
        mem->move_first();
        const long first = mem->bookmark();
        mem->move_last();
        const long last = mem->bookmark();
        check(first > 0 && last > 0 && first != last, "bookmarks differ");
        check(mem->set_sort("id DESC") && mem->goto_bookmark(first) &&
                  mem->bookmark() == first,
              "bookmark survives sorting");
        mem->move_first();
        static_cast<void>(mem->delete_current());
        check(mem->goto_bookmark(first) || mem->goto_bookmark(last),
              "bookmarks survive deletes");
        check(!mem->goto_bookmark(9999), "unknown bookmark refused");
    }
    mem->close();
    check(mem->state() == adStateClosed, "closed");
    check(!cn.execute("select 1").recordset->editable(),
          "a query result is read-only");

    {
        // Write-back: a one-table SELECT opened updatable.
        cn.execute("create table w(id integer primary key, name text not "
                   "null, n integer default 7)");
        cn.execute("insert into w(name) values ('a'), ('b')");
        auto up = cn.execute_updatable("select id, name from w order by id");
        check(up.ok && up.recordset && up.recordset->editable() &&
                  up.recordset->linked(),
              "a simple select is linked");
        auto& w = *up.recordset;
        check(w.set_value(1, Value{std::string("A")}) && w.update(),
              "update writes");
        w.add_new();
        w.set_value(1, Value{std::string("c")});
        check(w.update() && std::get<std::int64_t>(w.value(0)) == 3,
              "insert reads the key back");
        w.move_first();
        check(w.delete_current(), "delete writes");
        auto seen = cn.execute("select id, name, n from w order by id");
        check(seen.recordset->record_count() == 2 &&
                  std::get<std::string>(seen.recordset->value(1)) == "b",
              "the table has the changes");
        seen.recordset->move_last();
        check(std::get<std::int64_t>(seen.recordset->value(2)) == 7,
              "the default was applied");
        w.add_new();
        Error failure;
        check(!w.update() && w.take_error(failure) && failure.number != 0,
              "a failing insert reports an error and is undone");
        check(!cn.execute_updatable("select upper(name) from w")
                   .recordset->editable(),
              "an expression is read-only");
        check(!cn.execute_updatable("select * from w a, w b")
                   .recordset->editable(),
              "a join is read-only");
        check(cn.execute_updatable("select * from w where id > 1 limit 5")
                  .recordset->linked(),
              "where and limit stay linked");
    }
    cn.close();
    check(cn.state() == adStateClosed, "closed again");

    if (failures == 0) {
        std::cout << "wfc-adodb: all checks passed (SQLite " << cn.version()
                  << ")\n";
    }
    return failures == 0 ? EXIT_SUCCESS : EXIT_FAILURE;
}
