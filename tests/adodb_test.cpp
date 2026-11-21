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
    mem->close();
    check(mem->state() == adStateClosed, "closed");
    check(!cn.execute("select 1").recordset->editable(),
          "a query result is read-only");

    cn.close();
    check(cn.state() == adStateClosed, "closed again");

    if (failures == 0) {
        std::cout << "wfc-adodb: all checks passed (SQLite " << cn.version()
                  << ")\n";
    }
    return failures == 0 ? EXIT_SUCCESS : EXIT_FAILURE;
}
