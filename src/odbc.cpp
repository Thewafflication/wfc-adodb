// The ODBC backend. See src/odbc.hpp. Windows and Winelib (odbc32); other
// platforms get a backend that reports it is unavailable.

#include "odbc.hpp"

#include <algorithm>
#include <cctype>
#include <cstdint>
#include <cstdlib>
#include <utility>

#ifdef _WIN32
#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>

#include <sql.h>
#include <sqlext.h>
#include <sqlucode.h>

#ifdef WFC_WINELIB
// Wine's odbc32 exports this but its headers do not declare it.
extern "C" SQLRETURN WINAPI SQLDriverConnectW(
    SQLHDBC hdbc, SQLHWND hwnd, SQLWCHAR* in, SQLSMALLINT in_length,
    SQLWCHAR* out, SQLSMALLINT out_max, SQLSMALLINT* out_length,
    SQLUSMALLINT completion);
#endif
#endif

namespace wfc::adodb::detail {
namespace {

constexpr long kUnspecified = -2147467259L;  // E_FAIL
constexpr long kSyntaxError = -2147217900L;  // DB_E_ERRORSINCOMMAND
constexpr long kConstraint = -2147217873L;   // DB_E_INTEGRITYVIOLATION

[[nodiscard]] std::string lowered(std::string_view text) {
    std::string out(text);
    std::transform(out.begin(), out.end(), out.begin(), [](unsigned char c) {
        return static_cast<char>(std::tolower(c));
    });
    return out;
}

[[nodiscard]] std::string_view trimmed(std::string_view text) {
    while (!text.empty() &&
           std::isspace(static_cast<unsigned char>(text.front())) != 0) {
        text.remove_prefix(1);
    }
    while (!text.empty() &&
           std::isspace(static_cast<unsigned char>(text.back())) != 0) {
        text.remove_suffix(1);
    }
    return text;
}

// The "key=value" parts of a connection string. A value in braces may hold
// semicolons, as in Driver={ODBC Driver 17 for SQL Server}.
[[nodiscard]] std::vector<std::string> split_parts(std::string_view text) {
    std::vector<std::string> parts;
    std::string current;
    bool braces = false;
    for (const char c : text) {
        if (c == '{') {
            braces = true;
        } else if (c == '}') {
            braces = false;
        }
        if (c == ';' && !braces) {
            parts.push_back(current);
            current.clear();
        } else {
            current.push_back(c);
        }
    }
    parts.push_back(current);
    return parts;
}

[[nodiscard]] std::string key_of(const std::string& part) {
    const auto eq = part.find('=');
    return eq == std::string::npos
               ? std::string{}
               : lowered(trimmed(std::string_view(part).substr(0, eq)));
}

[[nodiscard]] std::string value_of(const std::string& part) {
    const auto eq = part.find('=');
    return eq == std::string::npos
               ? std::string{}
               : std::string(trimmed(std::string_view(part).substr(eq + 1)));
}

[[nodiscard]] long number_for(std::string_view sql_state) {
    if (sql_state.starts_with("23")) {
        return kConstraint;
    }
    if (sql_state.starts_with("42") || sql_state.starts_with("37")) {
        return kSyntaxError;
    }
    return kUnspecified;
}

}  // namespace

bool wants_odbc(const std::string_view connection_string) {
    for (const auto& part : split_parts(connection_string)) {
        const std::string key = key_of(part);
        if (key == "dsn" || key == "filedsn" || key == "driver") {
            return true;
        }
        if (key == "provider" && lowered(value_of(part)) == "msdasql") {
            return true;
        }
    }
    return false;
}

#ifdef _WIN32

namespace {

[[nodiscard]] std::u16string widen(const std::string_view text) {
    if (text.empty()) {
        return {};
    }
    const int size = MultiByteToWideChar(
        CP_UTF8, 0, text.data(), static_cast<int>(text.size()), nullptr, 0);
    std::u16string out(static_cast<std::size_t>(size), L'\0');
    MultiByteToWideChar(CP_UTF8, 0, text.data(), static_cast<int>(text.size()),
                        reinterpret_cast<WCHAR*>(out.data()), size);
    return out;
}

[[nodiscard]] std::string narrow(const char16_t* text, const std::size_t count) {
    if (count == 0) {
        return {};
    }
    const int size =
        WideCharToMultiByte(CP_UTF8, 0, reinterpret_cast<const WCHAR*>(text),
                            static_cast<int>(count), nullptr,
                            0, nullptr, nullptr);
    std::string out(static_cast<std::size_t>(size), '\0');
    WideCharToMultiByte(CP_UTF8, 0, reinterpret_cast<const WCHAR*>(text),
                        static_cast<int>(count), out.data(),
                        size, nullptr, nullptr);
    return out;
}

[[nodiscard]] SQLWCHAR* sql_text(std::u16string& text) {
    return reinterpret_cast<SQLWCHAR*>(text.data());
}

[[nodiscard]] SQLPOINTER as_pointer(const std::uintptr_t value) {
    return reinterpret_cast<SQLPOINTER>(value);
}

// Days from 1899-12-30 (the OLE automation date origin) to a civil date.
[[nodiscard]] double ole_days(const long year, const long month,
                              const long day) {
    const long y = month <= 2 ? year - 1 : year;
    const long era = (y >= 0 ? y : y - 399) / 400;
    const long yoe = y - era * 400;
    const long mp = (month + 9) % 12;
    const long doy = (153 * mp + 2) / 5 + day - 1;
    const long doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    const long since_1970 = era * 146097 + doe - 719468;
    return static_cast<double>(since_1970 + 25569);
}

enum class Kind { Integer, Real, Text, Binary, Date, Time, Stamp };

struct Column {
    int ado{0};
    Kind kind{Kind::Text};
};

[[nodiscard]] Column classify(const SQLSMALLINT type) {
    switch (type) {
        case SQL_SMALLINT:
            return {2, Kind::Integer};
        case SQL_INTEGER:
            return {3, Kind::Integer};
        case SQL_TINYINT:
            return {16, Kind::Integer};
        case SQL_BIGINT:
            return {20, Kind::Integer};
        case SQL_BIT:
            return {11, Kind::Integer};
        case SQL_REAL:
            return {4, Kind::Real};
        case SQL_FLOAT:
        case SQL_DOUBLE:
            return {5, Kind::Real};
        case SQL_DECIMAL:
        case SQL_NUMERIC:
            return {131, Kind::Real};
        case SQL_BINARY:
            return {128, Kind::Binary};
        case SQL_VARBINARY:
            return {204, Kind::Binary};
        case SQL_LONGVARBINARY:
            return {205, Kind::Binary};
        case SQL_TYPE_DATE:
            return {133, Kind::Date};
        case SQL_TYPE_TIME:
            return {134, Kind::Time};
        case SQL_TYPE_TIMESTAMP:
            return {135, Kind::Stamp};
        case SQL_LONGVARCHAR:
        case SQL_WLONGVARCHAR:
            return {203, Kind::Text};
        case SQL_GUID:
            return {72, Kind::Text};
        default:
            return {202, Kind::Text};
    }
}

}  // namespace

struct OdbcSession::Impl {
    SQLHENV env{SQL_NULL_HENV};
    SQLHDBC dbc{SQL_NULL_HDBC};
    bool open{false};
    int depth{0};
};

namespace {

// Adds the driver's diagnostic records for a handle to errors.
void collect(const SQLSMALLINT kind, const SQLHANDLE handle,
             std::vector<Error>& errors, const char* fallback) {
    const std::size_t before = errors.size();
    for (SQLSMALLINT record = 1;; ++record) {
        SQLWCHAR state[6] = {};
        SQLWCHAR message[1024] = {};
        SQLINTEGER native = 0;
        SQLSMALLINT length = 0;
        const SQLRETURN status = SQLGetDiagRecW(
            kind, handle, record, state, &native, message, 1024, &length);
        if (!SQL_SUCCEEDED(status)) {
            break;
        }
        Error error;
        error.sql_state = narrow(reinterpret_cast<const char16_t*>(state), 5U);
        error.number = number_for(error.sql_state);
        error.description = narrow(
            reinterpret_cast<const char16_t*>(message),
            static_cast<std::size_t>(std::min<SQLSMALLINT>(length, 1023)));
        error.source = "ODBC";
        errors.push_back(std::move(error));
    }
    if (errors.size() == before) {
        errors.push_back(Error{kUnspecified, fallback, "ODBC", {}});
    }
}

// One column of the current row as text (UTF-8) or bytes, in chunks.
enum class Fetch { Value, Null, Failed };

[[nodiscard]] Fetch fetch_chunks(const SQLHSTMT statement,
                                 const SQLUSMALLINT column,
                                 const SQLSMALLINT ctype, std::string& out) {
    constexpr SQLLEN kBuffer = 4096;
    const SQLLEN terminator = ctype == SQL_C_WCHAR ? 2 : 0;
    std::vector<char> buffer(static_cast<std::size_t>(kBuffer));
    std::string raw;
    bool any = false;
    for (;;) {
        SQLLEN indicator = 0;
        const SQLRETURN status = SQLGetData(statement, column, ctype,
                                            buffer.data(), kBuffer, &indicator);
        if (status == SQL_NO_DATA) {
            break;
        }
        if (!SQL_SUCCEEDED(status)) {
            return Fetch::Failed;
        }
        if (indicator == SQL_NULL_DATA) {
            return any ? Fetch::Value : Fetch::Null;
        }
        any = true;
        const bool truncated =
            indicator == SQL_NO_TOTAL || indicator >= kBuffer - terminator + 1;
        const SQLLEN take = truncated ? kBuffer - terminator : indicator;
        raw.append(buffer.data(), static_cast<std::size_t>(take));
        if (!truncated) {
            break;
        }
    }
    if (ctype == SQL_C_WCHAR) {
        out = narrow(reinterpret_cast<const char16_t*>(raw.data()),
                     raw.size() / 2U);
    } else {
        out = std::move(raw);
    }
    return Fetch::Value;
}

[[nodiscard]] bool read_column(const SQLHSTMT statement,
                               const SQLUSMALLINT column, const Kind kind,
                               Value& out) {
    switch (kind) {
        case Kind::Integer: {
            SQLBIGINT number = 0;
            SQLLEN indicator = 0;
            if (!SQL_SUCCEEDED(SQLGetData(statement, column, SQL_C_SBIGINT,
                                          &number, 0, &indicator))) {
                return false;
            }
            out = indicator == SQL_NULL_DATA
                      ? Value{}
                      : Value{static_cast<std::int64_t>(number)};
            return true;
        }
        case Kind::Real: {
            double number = 0;
            SQLLEN indicator = 0;
            if (!SQL_SUCCEEDED(SQLGetData(statement, column, SQL_C_DOUBLE,
                                          &number, 0, &indicator))) {
                return false;
            }
            out = indicator == SQL_NULL_DATA ? Value{} : Value{number};
            return true;
        }
        case Kind::Date: {
            SQL_DATE_STRUCT date{};
            SQLLEN indicator = 0;
            if (!SQL_SUCCEEDED(SQLGetData(statement, column, SQL_C_TYPE_DATE,
                                          &date, 0, &indicator))) {
                return false;
            }
            out = indicator == SQL_NULL_DATA
                      ? Value{}
                      : Value{ole_days(date.year, date.month, date.day)};
            return true;
        }
        case Kind::Time: {
            SQL_TIME_STRUCT time{};
            SQLLEN indicator = 0;
            if (!SQL_SUCCEEDED(SQLGetData(statement, column, SQL_C_TYPE_TIME,
                                          &time, 0, &indicator))) {
                return false;
            }
            out = indicator == SQL_NULL_DATA
                      ? Value{}
                      : Value{(time.hour * 3600.0 + time.minute * 60.0 +
                               time.second) /
                              86400.0};
            return true;
        }
        case Kind::Stamp: {
            SQL_TIMESTAMP_STRUCT stamp{};
            SQLLEN indicator = 0;
            if (!SQL_SUCCEEDED(SQLGetData(statement, column,
                                          SQL_C_TYPE_TIMESTAMP, &stamp, 0,
                                          &indicator))) {
                return false;
            }
            if (indicator == SQL_NULL_DATA) {
                out = Value{};
                return true;
            }
            const double seconds = stamp.hour * 3600.0 + stamp.minute * 60.0 +
                                   stamp.second + stamp.fraction / 1.0e9;
            out = Value{ole_days(stamp.year, stamp.month, stamp.day) +
                        seconds / 86400.0};
            return true;
        }
        case Kind::Binary: {
            std::string bytes;
            const Fetch got =
                fetch_chunks(statement, column, SQL_C_BINARY, bytes);
            if (got == Fetch::Failed) {
                return false;
            }
            out = got == Fetch::Null ? Value{}
                                     : Value{Blob(bytes.begin(), bytes.end())};
            return true;
        }
        case Kind::Text:
        default: {
            std::string text;
            const Fetch got =
                fetch_chunks(statement, column, SQL_C_WCHAR, text);
            if (got == Fetch::Failed) {
                return false;
            }
            out = got == Fetch::Null ? Value{} : Value{std::move(text)};
            return true;
        }
    }
}

// Parameter storage that stays put until the statement has run.
struct Parameter {
    std::int64_t integer{0};
    double real{0};
    std::u16string text;
    Blob bytes;
    SQLLEN indicator{0};
};

[[nodiscard]] bool bind_parameter(const SQLHSTMT statement,
                                  const SQLUSMALLINT slot, const Value& value,
                                  Parameter& store) {
    SQLRETURN status = SQL_SUCCESS;
    switch (value.index()) {
        case 1:
            store.integer = std::get<1>(value);
            // Not every driver has SQL_BIGINT (Access does not); a value
            // that fits a 32-bit integer goes as SQL_INTEGER.
            status = SQLBindParameter(
                statement, slot, SQL_PARAM_INPUT, SQL_C_SBIGINT,
                store.integer >= INT32_MIN && store.integer <= INT32_MAX
                    ? SQL_INTEGER
                    : SQL_BIGINT,
                0, 0, &store.integer, 0, &store.indicator);
            break;
        case 2:
            store.real = std::get<2>(value);
            status = SQLBindParameter(statement, slot, SQL_PARAM_INPUT,
                                      SQL_C_DOUBLE, SQL_DOUBLE, 0, 0,
                                      &store.real, 0, &store.indicator);
            break;
        case 3: {
            store.text = widen(std::get<3>(value));
            const auto chars = store.text.size();
            store.indicator = static_cast<SQLLEN>(chars * 2U);
            status = SQLBindParameter(
                statement, slot, SQL_PARAM_INPUT, SQL_C_WCHAR,
                chars > 4000U ? SQL_WLONGVARCHAR : SQL_WVARCHAR,
                std::max<SQLULEN>(1, chars), 0, store.text.data(),
                static_cast<SQLLEN>(chars * 2U), &store.indicator);
            break;
        }
        case 4: {
            store.bytes = std::get<4>(value);
            const auto size = store.bytes.size();
            store.indicator = static_cast<SQLLEN>(size);
            status = SQLBindParameter(
                statement, slot, SQL_PARAM_INPUT, SQL_C_BINARY,
                size > 8000U ? SQL_LONGVARBINARY : SQL_VARBINARY,
                std::max<SQLULEN>(1, size), 0, store.bytes.data(),
                static_cast<SQLLEN>(size), &store.indicator);
            break;
        }
        default:
            store.indicator = SQL_NULL_DATA;
            status = SQLBindParameter(statement, slot, SQL_PARAM_INPUT,
                                      SQL_C_CHAR, SQL_VARCHAR, 1, 0, nullptr, 0,
                                      &store.indicator);
    }
    return SQL_SUCCEEDED(status);
}

}  // namespace

OdbcSession::OdbcSession() : impl_(std::make_unique<Impl>()) {}

OdbcSession::~OdbcSession() {
    disconnect();
}

bool OdbcSession::connected() const noexcept {
    return impl_->open;
}

int OdbcSession::depth() const noexcept {
    return impl_->depth;
}

bool OdbcSession::connect(const std::string_view connection_string,
                          std::vector<Error>& errors) {
    // The Provider key means nothing to a driver; leave it out.
    std::string text;
    for (const auto& part : split_parts(connection_string)) {
        if (key_of(part) == "provider" || trimmed(part).empty()) {
            continue;
        }
        text += part;
        text += ';';
    }
    if (SQLAllocHandle(SQL_HANDLE_ENV, SQL_NULL_HANDLE, &impl_->env) !=
        SQL_SUCCESS) {
        errors.push_back(Error{kUnspecified, "Out of memory.", "ODBC", {}});
        return false;
    }
    SQLSetEnvAttr(impl_->env, SQL_ATTR_ODBC_VERSION, as_pointer(SQL_OV_ODBC3),
                  0);
    if (!SQL_SUCCEEDED(
            SQLAllocHandle(SQL_HANDLE_DBC, impl_->env, &impl_->dbc))) {
        collect(SQL_HANDLE_ENV, impl_->env, errors, "Cannot allocate.");
        disconnect();
        return false;
    }
    SQLSetConnectAttr(impl_->dbc, SQL_LOGIN_TIMEOUT, as_pointer(15), 0);
    std::u16string in = widen(text);
    SQLWCHAR out[1024] = {};
    SQLSMALLINT out_length = 0;
    const SQLRETURN status =
        SQLDriverConnectW(impl_->dbc, nullptr, sql_text(in), SQL_NTS, out, 1024,
                          &out_length, SQL_DRIVER_NOPROMPT);
    if (!SQL_SUCCEEDED(status)) {
        collect(SQL_HANDLE_DBC, impl_->dbc, errors, "Cannot connect.");
        disconnect();
        return false;
    }
    impl_->open = true;
    impl_->depth = 0;
    return true;
}

void OdbcSession::disconnect() noexcept {
    if (impl_->dbc != SQL_NULL_HDBC) {
        if (impl_->open) {
            if (impl_->depth > 0) {
                SQLEndTran(SQL_HANDLE_DBC, impl_->dbc, SQL_ROLLBACK);
            }
            SQLDisconnect(impl_->dbc);
        }
        SQLFreeHandle(SQL_HANDLE_DBC, impl_->dbc);
        impl_->dbc = SQL_NULL_HDBC;
    }
    if (impl_->env != SQL_NULL_HENV) {
        SQLFreeHandle(SQL_HANDLE_ENV, impl_->env);
        impl_->env = SQL_NULL_HENV;
    }
    impl_->open = false;
    impl_->depth = 0;
}

bool OdbcSession::execute(const std::string_view sql,
                          const std::vector<Value>& params, OdbcTable& out,
                          std::vector<Error>& errors) {
    SQLHSTMT statement = SQL_NULL_HSTMT;
    if (!SQL_SUCCEEDED(
            SQLAllocHandle(SQL_HANDLE_STMT, impl_->dbc, &statement))) {
        collect(SQL_HANDLE_DBC, impl_->dbc, errors, "Cannot allocate.");
        return false;
    }
    const auto fail = [&](const char* fallback) {
        collect(SQL_HANDLE_STMT, statement, errors, fallback);
        SQLFreeHandle(SQL_HANDLE_STMT, statement);
        return false;
    };

    std::u16string text = widen(sql);
    std::vector<Parameter> store(params.size());
    SQLRETURN status = SQL_SUCCESS;
    if (params.empty()) {
        status = SQLExecDirectW(statement, sql_text(text),
                                static_cast<SQLINTEGER>(text.size()));
    } else {
        status = SQLPrepareW(statement, sql_text(text),
                             static_cast<SQLINTEGER>(text.size()));
        for (std::size_t i = 0; SQL_SUCCEEDED(status) && i < params.size();
             ++i) {
            if (!bind_parameter(statement, static_cast<SQLUSMALLINT>(i + 1U),
                                params[i], store[i])) {
                return fail("Parameter object is improperly defined.");
            }
        }
        if (SQL_SUCCEEDED(status)) {
            status = SQLExecute(statement);
        }
    }
    if (!SQL_SUCCEEDED(status) && status != SQL_NO_DATA) {
        return fail("The statement failed.");
    }

    SQLLEN affected = -1;
    SQLSMALLINT columns = 0;
    for (;;) {
        SQLNumResultCols(statement, &columns);
        if (columns == 0 && affected < 0) {
            SQLRowCount(statement, &affected);
        }
        if (columns > 0 || !SQL_SUCCEEDED(SQLMoreResults(statement))) {
            break;
        }
    }
    out.affected = static_cast<long>(std::max<SQLLEN>(affected, 0));
    out.columns = columns > 0;
    std::vector<Kind> kinds;
    for (SQLSMALLINT c = 1; c <= columns; ++c) {
        SQLWCHAR name[256] = {};
        SQLSMALLINT name_length = 0;
        SQLSMALLINT type = 0;
        SQLULEN size = 0;
        SQLSMALLINT digits = 0;
        SQLSMALLINT nullable = 0;
        if (!SQL_SUCCEEDED(SQLDescribeColW(
                statement, static_cast<SQLUSMALLINT>(c), name, 256,
                &name_length, &type, &size, &digits, &nullable))) {
            return fail("Cannot describe a column.");
        }
        const Column column = classify(type);
        kinds.push_back(column.kind);
        out.fields.push_back(
            FieldInfo{narrow(reinterpret_cast<const char16_t*>(name),
                             static_cast<std::size_t>(
                                 std::min<SQLSMALLINT>(name_length, 255))),
                      column.ado, static_cast<long>(size)});
    }
    while (columns > 0) {
        status = SQLFetch(statement);
        if (status == SQL_NO_DATA) {
            break;
        }
        if (!SQL_SUCCEEDED(status)) {
            return fail("Cannot fetch a row.");
        }
        std::vector<Value> row(static_cast<std::size_t>(columns));
        for (std::size_t c = 0; c < kinds.size(); ++c) {
            if (!read_column(statement, static_cast<SQLUSMALLINT>(c + 1U),
                             kinds[c], row[c])) {
                return fail("Cannot read a column.");
            }
        }
        out.rows.push_back(std::move(row));
    }
    SQLFreeHandle(SQL_HANDLE_STMT, statement);
    return true;
}

int OdbcSession::begin(std::vector<Error>& errors) {
    if (impl_->depth > 0) {
        errors.push_back(Error{kUnspecified,
                               "ODBC does not support nested transactions.",
                               "ODBC",
                               {}});
        return 0;
    }
    if (!SQL_SUCCEEDED(SQLSetConnectAttr(impl_->dbc, SQL_ATTR_AUTOCOMMIT,
                                         as_pointer(SQL_AUTOCOMMIT_OFF), 0))) {
        collect(SQL_HANDLE_DBC, impl_->dbc, errors,
                "Transactions are not supported.");
        return 0;
    }
    impl_->depth = 1;
    return 1;
}

bool OdbcSession::end(const bool commit, std::vector<Error>& errors) {
    if (impl_->depth == 0) {
        errors.push_back(Error{3704, "No transaction is active.", "ODBC", {}});
        return false;
    }
    const bool done = SQL_SUCCEEDED(SQLEndTran(
        SQL_HANDLE_DBC, impl_->dbc, commit ? SQL_COMMIT : SQL_ROLLBACK));
    if (!done) {
        collect(SQL_HANDLE_DBC, impl_->dbc, errors, "The transaction failed.");
    }
    SQLSetConnectAttr(impl_->dbc, SQL_ATTR_AUTOCOMMIT,
                      as_pointer(SQL_AUTOCOMMIT_ON), 0);
    impl_->depth = 0;
    return done;
}

std::int64_t OdbcSession::last_insert_id() {
    if (!impl_->open) {
        return 0;
    }
    // Each DBMS has its own way to ask; the identity belongs to this session.
    const std::string name = lowered(dbms_name());
    const auto has = [&](const char* word) {
        return name.find(word) != std::string::npos;
    };
    const char* sql = nullptr;
    if (has("access") || has("sql server") || has("sybase") ||
        has("adaptive server")) {
        sql = "SELECT @@IDENTITY";
    } else if (has("mysql") || has("mariadb")) {
        sql = "SELECT LAST_INSERT_ID()";
    } else if (has("postgres")) {
        sql = "SELECT lastval()";
    } else if (has("sqlite")) {
        sql = "SELECT last_insert_rowid()";
    } else if (has("db2")) {
        sql = "VALUES IDENTITY_VAL_LOCAL()";
    }
    if (sql == nullptr) {
        return 0;
    }
    OdbcTable table;
    std::vector<Error> ignored;
    if (!execute(sql, {}, table, ignored) || table.rows.empty() ||
        table.rows.front().empty()) {
        return 0;
    }
    const Value& value = table.rows.front().front();
    if (const auto* whole = std::get_if<std::int64_t>(&value)) {
        return *whole;
    }
    if (const auto* real = std::get_if<double>(&value)) {
        return static_cast<std::int64_t>(*real);
    }
    if (const auto* text = std::get_if<std::string>(&value)) {
        return std::strtoll(text->c_str(), nullptr, 10);
    }
    return 0;
}

std::string OdbcSession::dbms_name() const {
    if (!impl_->open) {
        return {};
    }
    SQLWCHAR buffer[256] = {};
    SQLSMALLINT length = 0;
    if (!SQL_SUCCEEDED(SQLGetInfoW(impl_->dbc, SQL_DBMS_NAME, buffer,
                                   sizeof(buffer), &length))) {
        return {};
    }
    return narrow(reinterpret_cast<const char16_t*>(buffer),
                  static_cast<std::size_t>(length) / 2U);
}

std::string OdbcSession::version() const {
    if (!impl_->open) {
        return {};
    }
    const auto info = [&](const SQLUSMALLINT what) {
        SQLWCHAR buffer[256] = {};
        SQLSMALLINT length = 0;
        if (!SQL_SUCCEEDED(SQLGetInfoW(impl_->dbc, what, buffer, sizeof(buffer),
                                       &length))) {
            return std::string{};
        }
        return narrow(reinterpret_cast<const char16_t*>(buffer),
                      static_cast<std::size_t>(length) / 2U);
    };
    return info(SQL_DBMS_NAME) + " " + info(SQL_DBMS_VER);
}

#else  // not Windows

struct OdbcSession::Impl {};

OdbcSession::OdbcSession() : impl_(std::make_unique<Impl>()) {}
OdbcSession::~OdbcSession() = default;

bool OdbcSession::connected() const noexcept {
    return false;
}

int OdbcSession::depth() const noexcept {
    return 0;
}

bool OdbcSession::connect(std::string_view, std::vector<Error>& errors) {
    errors.push_back(Error{kUnspecified,
                           "The ODBC backend is only built on Windows and Winelib.",
                           "ODBC",
                           {}});
    return false;
}

void OdbcSession::disconnect() noexcept {}

bool OdbcSession::execute(std::string_view, const std::vector<Value>&,
                          OdbcTable&, std::vector<Error>&) {
    return false;
}

int OdbcSession::begin(std::vector<Error>&) {
    return 0;
}

bool OdbcSession::end(bool, std::vector<Error>&) {
    return false;
}

std::string OdbcSession::version() const {
    return {};
}

std::int64_t OdbcSession::last_insert_id() {
    return 0;
}

std::string OdbcSession::dbms_name() const {
    return {};
}

#endif

}  // namespace wfc::adodb::detail
