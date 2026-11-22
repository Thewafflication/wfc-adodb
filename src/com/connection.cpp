// ADODB.Connection as a COM object, with Errors and OpenSchema.

#include "objects.hpp"

#include <algorithm>

namespace wfc::adodb::com {

ConnectionObj::~ConnectionObj() {
    if (properties != nullptr) {
        properties->Release();
    }
}

std::shared_ptr<Connection> native_of(IUnknown* object, const VARIANT* text,
                                      Ref& keep) {
    if (auto* mine = unwrap<ConnectionObj>(object)) {
        keep = Ref(static_cast<IDispatch*>(mine));
        return mine->native;
    }
    if (text == nullptr) {
        return nullptr;
    }
    auto* created = new ConnectionObj();
    std::shared_ptr<Connection> native = created->native;
    std::string connection_string;
    const Value given = to_value(*text);
    if (const auto* string = std::get_if<std::string>(&given)) {
        connection_string = *string;
    }
    created->connection_string = connection_string;
    keep = Ref(static_cast<IDispatch*>(created));
    created->Release();
    if (!native->open(connection_string)) {
        fail_from(*native);
    }
    return native;
}

namespace {

// A disconnected recordset with the given text and number columns and rows.
struct Column {
    const char* name;
    int type;
};

std::shared_ptr<Recordset> make_table(
    const std::vector<Column>& columns,
    const std::vector<std::vector<Value>>& rows) {
    auto set = Recordset::create();
    for (const auto& column : columns) {
        set->append_field(column.name, column.type, 0);
    }
    set->open_empty();
    for (const auto& row : rows) {
        set->add_new();
        for (std::size_t i = 0; i < row.size(); ++i) {
            set->set_value(i, row[i]);
        }
        set->update();
    }
    set->move_first();
    return set;
}

// The text of restriction `index` of an OpenSchema Restrictions argument.
std::string restriction(const VARIANT& restrictions, const long index) {
    const Var plain = dereferenced(restrictions);
    const VARIANT& v = plain.ref();
    if ((v.vt & VT_ARRAY) == 0 || v.parray == nullptr ||
        (v.vt & VT_TYPEMASK) != VT_VARIANT) {
        return {};
    }
    long low = 0;
    long high = -1;
    SafeArrayGetLBound(v.parray, 1, &low);
    SafeArrayGetUBound(v.parray, 1, &high);
    const long at = low + index;
    if (at > high) {
        return {};
    }
    VARIANT element{};
    VariantInit(&element);
    SafeArrayGetElement(v.parray, const_cast<long*>(&at), &element);
    const Value value = to_value(element);
    VariantClear(&element);
    const auto* text = std::get_if<std::string>(&value);
    return text != nullptr ? *text : std::string{};
}

bool same_text(const std::string& a, const std::string& b) {
    return _stricmp(a.c_str(), b.c_str()) == 0;
}

// The ADO data type of a declared SQLite column type.
std::int64_t ado_type_of(const std::string& declared) {
    std::string t;
    for (const char c : declared) {
        t.push_back(
            static_cast<char>(std::tolower(static_cast<unsigned char>(c))));
    }
    const auto has = [&](const char* word) {
        return t.find(word) != std::string::npos;
    };
    if (has("int")) {
        return has("big") ? 20 : 3;
    }
    if (has("char") || has("clob") || has("text")) {
        return 202;
    }
    if (has("blob")) {
        return 205;
    }
    if (has("real") || has("floa") || has("doub")) {
        return 5;
    }
    if (has("bool")) {
        return 11;
    }
    if (has("date") || has("time")) {
        return 7;
    }
    return 12;
}

// The rows of a query on the native connection as plain values.
std::vector<std::vector<Value>> query_rows(Connection& native,
                                           const std::string& sql) {
    auto outcome = native.execute(sql);
    if (!outcome.ok) {
        fail_from(native);
    }
    std::vector<std::vector<Value>> rows;
    if (outcome.recordset == nullptr) {
        return rows;
    }
    auto& set = *outcome.recordset;
    set.move_first();
    while (!set.eof()) {
        std::vector<Value> row;
        for (std::size_t i = 0; i < set.field_count(); ++i) {
            row.push_back(set.value(i));
        }
        rows.push_back(std::move(row));
        set.move_next();
    }
    return rows;
}

std::string quote_name(const std::string& name) {
    std::string out = "\"";
    for (const char c : name) {
        if (c == '"') {
            out += '"';
        }
        out += c;
    }
    return out + "\"";
}

std::string text_of(const Value& v) {
    const auto* s = std::get_if<std::string>(&v);
    return s != nullptr ? *s : std::string{};
}

std::int64_t number_of(const Value& v) {
    const auto* n = std::get_if<std::int64_t>(&v);
    return n != nullptr ? *n : 0;
}

// Connection.OpenSchema over SQLite: the tables, columns, primary keys and
// indexes of the database.
std::shared_ptr<Recordset> open_schema(Connection& native, const long schema,
                                       const VARIANT& restrictions) {
    if (native.state() == adStateClosed) {
        throw ComError{3704, kClosedText};
    }
    const auto tables = [&] {
        const auto rows = query_rows(
            native,
            "select name, type from sqlite_master where type in ('table', "
            "'view') and name not like 'sqlite_%' order by name");
        std::vector<std::pair<std::string, std::string>> list;
        const std::string wanted = restriction(restrictions, 2);
        const std::string kind = restriction(restrictions, 3);
        for (const auto& row : rows) {
            const std::string name = text_of(row[0]);
            const std::string type =
                text_of(row[1]) == "view" ? "VIEW" : "TABLE";
            if ((wanted.empty() || same_text(wanted, name)) &&
                (kind.empty() || same_text(kind, type))) {
                list.emplace_back(name, type);
            }
        }
        return list;
    };
    switch (schema) {
        case 20: {  // adSchemaTables
            std::vector<std::vector<Value>> rows;
            for (const auto& [name, type] : tables()) {
                rows.push_back({Value{}, Value{}, Value{name}, Value{type},
                                Value{}, Value{}});
            }
            return make_table({{"TABLE_CATALOG", 202},
                               {"TABLE_SCHEMA", 202},
                               {"TABLE_NAME", 202},
                               {"TABLE_TYPE", 202},
                               {"TABLE_GUID", 202},
                               {"DESCRIPTION", 202}},
                              rows);
        }
        case 4: {  // adSchemaColumns
            std::vector<std::vector<Value>> rows;
            const std::string wanted_column = restriction(restrictions, 3);
            for (const auto& [name, type] : tables()) {
                const auto info = query_rows(
                    native, "pragma table_info(" + quote_name(name) + ")");
                for (const auto& column : info) {
                    const std::string column_name = text_of(column[1]);
                    if (!wanted_column.empty() &&
                        !same_text(wanted_column, column_name)) {
                        continue;
                    }
                    const std::string declared = text_of(column[2]);
                    rows.push_back(
                        {Value{}, Value{}, Value{name}, Value{column_name},
                         Value{number_of(column[0]) + 1}, column[4],
                         Value{static_cast<std::int64_t>(
                             number_of(column[3]) == 0 ? -1 : 0)},
                         Value{ado_type_of(declared)}, Value{declared}});
                }
            }
            return make_table({{"TABLE_CATALOG", 202},
                               {"TABLE_SCHEMA", 202},
                               {"TABLE_NAME", 202},
                               {"COLUMN_NAME", 202},
                               {"ORDINAL_POSITION", 3},
                               {"COLUMN_DEFAULT", 202},
                               {"IS_NULLABLE", 11},
                               {"DATA_TYPE", 3},
                               {"TYPE_NAME", 202}},
                              rows);
        }
        case 28: {  // adSchemaPrimaryKeys
            std::vector<std::vector<Value>> rows;
            for (const auto& [name, type] : tables()) {
                const auto info = query_rows(
                    native, "pragma table_info(" + quote_name(name) + ")");
                for (const auto& column : info) {
                    if (number_of(column[5]) > 0) {
                        rows.push_back({Value{}, Value{}, Value{name},
                                        column[1], Value{number_of(column[5])},
                                        Value{"PRIMARY"}});
                    }
                }
            }
            return make_table({{"TABLE_CATALOG", 202},
                               {"TABLE_SCHEMA", 202},
                               {"TABLE_NAME", 202},
                               {"COLUMN_NAME", 202},
                               {"ORDINAL", 3},
                               {"PK_NAME", 202}},
                              rows);
        }
        case 12: {  // adSchemaIndexes
            std::vector<std::vector<Value>> rows;
            for (const auto& [name, type] : tables()) {
                const auto indexes = query_rows(
                    native, "pragma index_list(" + quote_name(name) + ")");
                for (const auto& index : indexes) {
                    const std::string index_name = text_of(index[1]);
                    const auto columns =
                        query_rows(native, "pragma index_info(" +
                                               quote_name(index_name) + ")");
                    for (const auto& column : columns) {
                        rows.push_back(
                            {Value{}, Value{}, Value{name}, Value{index_name},
                             Value{static_cast<std::int64_t>(
                                 number_of(index[2]) != 0 ? -1 : 0)},
                             column[2], Value{number_of(column[0]) + 1}});
                    }
                }
            }
            return make_table({{"TABLE_CATALOG", 202},
                               {"TABLE_SCHEMA", 202},
                               {"TABLE_NAME", 202},
                               {"INDEX_NAME", 202},
                               {"UNIQUE", 11},
                               {"COLUMN_NAME", 202},
                               {"ORDINAL_POSITION", 3}},
                              rows);
        }
        default:
            throw ComError{3251, kUnsupportedText};
    }
}

}  // namespace

const Table<ConnectionObj>& ConnectionObj::members() {
    static const Table<ConnectionObj> table = {
        {L"Properties", 500,
         [](ConnectionObj& s, Call& c) {
             if (s.properties == nullptr) {
                 make_properties(s.properties);
                 Var number;
                 number.set_long(4);  // adPromptNever
                 s.properties->add("Prompt", number, 3);
                 number.set_long(1033);
                 s.properties->add("Locale Identifier", number, 3);
                 number.set_long(2);
                 s.properties->add("Mode", number, 3);
             }
             c.put_object(s.properties);
         },
         nullptr, true},
        {L"ConnectionString", DISPID_VALUE,
         read_text(&ConnectionObj::connection_string),
         write_text(&ConnectionObj::connection_string)},
        {L"CommandTimeout", 2, read_long(&ConnectionObj::command_timeout),
         write_long(&ConnectionObj::command_timeout, 30)},
        {L"ConnectionTimeout", 3, read_long(&ConnectionObj::connection_timeout),
         write_long(&ConnectionObj::connection_timeout, 15)},
        {L"Version", 4, [](ConnectionObj&, Call& c) { c.put_text("2.8"); },
         nullptr},
        {L"Close", 5,
         [](ConnectionObj& s, Call&) {
             if (s.native->state() == adStateClosed) {
                 throw ComError{3704, kClosedText};
             }
             s.native->close();
         },
         nullptr},
        {L"Execute", 6,
         [](ConnectionObj& s, Call& c) {
             if (s.native->state() == adStateClosed) {
                 throw ComError{3704, kClosedText};
             }
             std::string sql = c.to_text(0);
             const long options = c.to_long(2, -1);
             if (options > 0 && (options & 4) != 0) {
                 throw ComError{3251, kUnsupportedText};
             }
             if (options > 0 && ((options & 2) != 0 || (options & 512) != 0)) {
                 sql = "select * from " + sql;
             }
             s.native->clear_errors();
             auto outcome = s.native->execute(sql);
             if (!outcome.ok) {
                 fail_from(*s.native);
             }
             Var affected;
             affected.set_long(outcome.records_affected);
             c.write_back(1, affected);
             auto* set =
                 make_recordset(std::move(outcome.recordset),
                                Ref(static_cast<IDispatch*>(&s)), s.native, 1);
             set->source = sql;
             c.put_object(set);
             set->Release();
         },
         nullptr},
        {L"BeginTrans", 7,
         [](ConnectionObj& s, Call& c) {
             s.native->clear_errors();
             const int level = s.native->begin_trans();
             if (level == 0) {
                 fail_from(*s.native);
             }
             c.put_long(level);
         },
         nullptr},
        {L"CommitTrans", 8,
         [](ConnectionObj& s, Call&) {
             s.native->clear_errors();
             if (!s.native->commit_trans()) {
                 fail_from(*s.native);
             }
         },
         nullptr},
        {L"RollbackTrans", 9,
         [](ConnectionObj& s, Call&) {
             s.native->clear_errors();
             if (!s.native->rollback_trans()) {
                 fail_from(*s.native);
             }
         },
         nullptr},
        {L"Open", 10,
         [](ConnectionObj& s, Call& c) {
             if (s.native->state() != adStateClosed) {
                 throw ComError{3705, kOpenText};
             }
             if (c.has(0)) {
                 s.connection_string = c.to_text(0);
             }
             s.native->clear_errors();
             if (!s.native->open(s.connection_string)) {
                 fail_from(*s.native);
             }
         },
         nullptr},
        {L"Errors", 11,
         [](ConnectionObj& s, Call& c) {
             auto* errors = new ErrorsObj();
             errors->native = s.native;
             c.put_object(errors);
             errors->Release();
         },
         nullptr, true},
        {L"DefaultDatabase", 12, read_text(&ConnectionObj::default_database),
         write_text(&ConnectionObj::default_database)},
        {L"IsolationLevel", 13, read_long(&ConnectionObj::isolation_level),
         write_long(&ConnectionObj::isolation_level, 4096)},
        {L"Attributes", 14, read_long(&ConnectionObj::attributes),
         write_long(&ConnectionObj::attributes)},
        {L"CursorLocation", 15, read_long(&ConnectionObj::cursor_location),
         write_long(&ConnectionObj::cursor_location, 2)},
        {L"Mode", 16, read_long(&ConnectionObj::mode),
         write_long(&ConnectionObj::mode)},
        {L"Provider", 17, read_text(&ConnectionObj::provider),
         write_text(&ConnectionObj::provider)},
        {L"State", 18,
         [](ConnectionObj& s, Call& c) { c.put_long(s.native->state()); },
         nullptr},
        {L"OpenSchema", 19,
         [](ConnectionObj& s, Call& c) {
             auto set = open_schema(*s.native, c.to_long(0, 0),
                                    c.has(1) ? c.args[1] : VARIANT{});
             auto* object = make_recordset(
                 std::move(set), Ref(static_cast<IDispatch*>(&s)), s.native, 1);
             c.put_object(object);
             object->Release();
         },
         nullptr},
        {L"Cancel", 21, [](ConnectionObj&, Call&) {}, nullptr},
    };
    return table;
}

}  // namespace wfc::adodb::com
