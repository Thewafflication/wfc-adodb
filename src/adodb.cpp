// The ADODB-to-SQLite interposer. See include/wfc/adodb/adodb.hpp.

#include <wfc/adodb/adodb.hpp>

#include <algorithm>
#include <cctype>
#include <cstring>
#include <utility>

#include "sqlite3.h"

namespace wfc::adodb {
namespace {

constexpr long kSyntaxError = -2147217900L;  // 0x80040E14, DB_E_ERRORSINCOMMAND
constexpr long kUnspecified = -2147467259L;  // 0x80004005, E_FAIL
constexpr long kNotOpen = 3704;              // operation not allowed when closed
constexpr long kAlreadyOpen = 3705;          // operation not allowed when open

[[nodiscard]] std::string lower(std::string_view text) {
    std::string out(text);
    std::transform(out.begin(), out.end(), out.begin(), [](unsigned char c) {
        return static_cast<char>(std::tolower(c));
    });
    return out;
}

[[nodiscard]] std::string_view trim(std::string_view text) {
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

[[nodiscard]] int type_from_declared(const char* declared) {
    if (declared == nullptr) {
        return adEmpty;
    }
    const std::string text = lower(declared);
    const auto has = [&](const char* part) {
        return text.find(part) != std::string::npos;
    };
    if (has("int")) {
        return has("big") ? adBigInt : adInteger;
    }
    if (has("char") || has("clob") || has("text")) {
        return adVarWChar;
    }
    if (has("blob")) {
        return adLongVarBinary;
    }
    if (has("real") || has("floa") || has("doub")) {
        return adDouble;
    }
    return adEmpty;
}

[[nodiscard]] int type_of_value(const Value& value) {
    switch (value.index()) {
        case 1:
            return adBigInt;
        case 2:
            return adDouble;
        case 3:
            return adVarWChar;
        case 4:
            return adLongVarBinary;
        default:
            return adEmpty;
    }
}

}  // namespace

std::string connection_string_path(std::string_view text) {
    if (text.find('=') == std::string_view::npos) {
        return std::string(trim(text));
    }
    std::string path;
    while (!text.empty()) {
        const auto cut = text.find(';');
        const std::string_view part = text.substr(0, cut);
        text = cut == std::string_view::npos ? std::string_view{}
                                             : text.substr(cut + 1);
        const auto eq = part.find('=');
        if (eq == std::string_view::npos) {
            continue;
        }
        const std::string key = lower(trim(part.substr(0, eq)));
        if (key == "data source" || key == "database" || key == "file name" ||
            key == "dbq") {
            path = std::string(trim(part.substr(eq + 1)));
        }
    }
    return path;
}

// ----- Recordset ------------------------------------------------------------------

int Recordset::field_index(std::string_view name) const noexcept {
    const std::string wanted = lower(name);
    for (std::size_t i = 0; i < fields_.size(); ++i) {
        if (lower(fields_[i].name) == wanted) {
            return static_cast<int>(i);
        }
    }
    return -1;
}

long Recordset::absolute_position() const noexcept {
    if (empty() || position_ < 0) {
        return -1;
    }
    if (position_ >= static_cast<long>(rows_.size())) {
        return -2;
    }
    return position_ + 1;
}

void Recordset::move_first() noexcept { position_ = 0; }

void Recordset::move_last() noexcept {
    position_ = rows_.empty() ? 0 : static_cast<long>(rows_.size()) - 1;
}

bool Recordset::move_next() noexcept {
    if (eof()) {
        return false;
    }
    ++position_;
    return true;
}

bool Recordset::move_previous() noexcept {
    if (bof()) {
        return false;
    }
    --position_;
    return true;
}

void Recordset::close() noexcept {
    state_ = adStateClosed;
    fields_.clear();
    rows_.clear();
    position_ = 0;
}

const Value& Recordset::value(const std::size_t index) const {
    static const Value null_value;
    if (state_ != adStateOpen || bof() || eof() || index >= fields_.size()) {
        return null_value;
    }
    return rows_[static_cast<std::size_t>(position_)][index];
}

// ----- Connection -----------------------------------------------------------------

struct Connection::Impl {
    sqlite3* db{nullptr};
    int depth{0};  // open BeginTrans levels
};

Connection::Connection() : impl_(std::make_unique<Impl>()) {}

Connection::~Connection() { close(); }

void Connection::fail(const long number, std::string description) {
    errors_.push_back(Error{number, std::move(description), "wfc-adodb", {}});
}

int Connection::state() const noexcept {
    return impl_->db != nullptr ? adStateOpen : adStateClosed;
}

bool Connection::open(const std::string_view connection_string) {
    if (impl_->db != nullptr) {
        fail(kAlreadyOpen, "Operation is not allowed when the object is open.");
        return false;
    }
    const std::string path = connection_string_path(connection_string);
    if (path.empty()) {
        fail(kUnspecified, "The connection string names no database.");
        return false;
    }
    sqlite3* db = nullptr;
    const int status = sqlite3_open_v2(
        path.c_str(), &db,
        SQLITE_OPEN_READWRITE | SQLITE_OPEN_CREATE | SQLITE_OPEN_URI, nullptr);
    if (status != SQLITE_OK) {
        fail(kUnspecified,
             db != nullptr ? sqlite3_errmsg(db) : "Out of memory.");
        sqlite3_close(db);
        return false;
    }
    impl_->db = db;
    impl_->depth = 0;
    sqlite3_busy_timeout(db, 5000);
    return true;
}

void Connection::close() noexcept {
    if (impl_->db == nullptr) {
        return;
    }
    // ADO rolls back what is still pending when a connection closes.
    sqlite3_exec(impl_->db, "ROLLBACK", nullptr, nullptr, nullptr);
    sqlite3_close_v2(impl_->db);
    impl_->db = nullptr;
    impl_->depth = 0;
}

std::string Connection::version() const { return sqlite3_libversion(); }

std::int64_t Connection::last_insert_id() const noexcept {
    return impl_->db != nullptr ? sqlite3_last_insert_rowid(impl_->db) : 0;
}

Result Connection::execute(const std::string_view sql,
                           const std::vector<Value>& params) {
    Result result;
    if (impl_->db == nullptr) {
        fail(kNotOpen, "Operation is not allowed when the object is closed.");
        return result;
    }
    sqlite3_stmt* stmt = nullptr;
    const char* tail = nullptr;
    if (sqlite3_prepare_v2(impl_->db, sql.data(), static_cast<int>(sql.size()),
                           &stmt, &tail) != SQLITE_OK) {
        fail(kSyntaxError, sqlite3_errmsg(impl_->db));
        return result;
    }
    if (stmt == nullptr) {  // only whitespace or a comment
        result.ok = true;
        return result;
    }
    const int wanted = sqlite3_bind_parameter_count(stmt);
    if (static_cast<std::size_t>(wanted) != params.size()) {
        sqlite3_finalize(stmt);
        fail(3708, "Parameter object is improperly defined.");
        return result;
    }
    for (int i = 0; i < wanted; ++i) {
        const Value& v = params[static_cast<std::size_t>(i)];
        const int slot = i + 1;
        switch (v.index()) {
            case 1:
                sqlite3_bind_int64(stmt, slot, std::get<1>(v));
                break;
            case 2:
                sqlite3_bind_double(stmt, slot, std::get<2>(v));
                break;
            case 3: {
                const auto& text = std::get<3>(v);
                sqlite3_bind_text(stmt, slot, text.data(),
                                  static_cast<int>(text.size()),
                                  SQLITE_TRANSIENT);
                break;
            }
            case 4: {
                const auto& blob = std::get<4>(v);
                sqlite3_bind_blob(stmt, slot, blob.data(),
                                  static_cast<int>(blob.size()),
                                  SQLITE_TRANSIENT);
                break;
            }
            default:
                sqlite3_bind_null(stmt, slot);
        }
    }

    auto set = std::make_shared<Recordset>();
    const int columns = sqlite3_column_count(stmt);
    for (int c = 0; c < columns; ++c) {
        set->fields_.push_back(FieldInfo{
            sqlite3_column_name(stmt, c) != nullptr ? sqlite3_column_name(stmt, c)
                                                    : "",
            type_from_declared(sqlite3_column_decltype(stmt, c))});
    }
    int status = SQLITE_OK;
    while ((status = sqlite3_step(stmt)) == SQLITE_ROW) {
        std::vector<Value> row;
        row.reserve(static_cast<std::size_t>(columns));
        for (int c = 0; c < columns; ++c) {
            switch (sqlite3_column_type(stmt, c)) {
                case SQLITE_INTEGER:
                    row.emplace_back(
                        static_cast<std::int64_t>(sqlite3_column_int64(stmt, c)));
                    break;
                case SQLITE_FLOAT:
                    row.emplace_back(sqlite3_column_double(stmt, c));
                    break;
                case SQLITE_TEXT:
                    row.emplace_back(std::string(
                        reinterpret_cast<const char*>(sqlite3_column_text(stmt, c)),
                        static_cast<std::size_t>(sqlite3_column_bytes(stmt, c))));
                    break;
                case SQLITE_BLOB: {
                    const auto* bytes = static_cast<const unsigned char*>(
                        sqlite3_column_blob(stmt, c));
                    const auto size =
                        static_cast<std::size_t>(sqlite3_column_bytes(stmt, c));
                    row.emplace_back(Blob(bytes, bytes + size));
                    break;
                }
                default:
                    row.emplace_back(std::monostate{});
            }
        }
        set->rows_.push_back(std::move(row));
    }
    if (status != SQLITE_DONE) {
        const std::string message = sqlite3_errmsg(impl_->db);
        sqlite3_finalize(stmt);
        fail(status == SQLITE_CONSTRAINT ? -2147217873L : kUnspecified, message);
        return result;
    }
    // A column without a declared type takes the type of its first value.
    for (std::size_t c = 0; c < set->fields_.size(); ++c) {
        if (set->fields_[c].type == adEmpty) {
            for (const auto& row : set->rows_) {
                if (row[c].index() != 0) {
                    set->fields_[c].type = type_of_value(row[c]);
                    break;
                }
            }
        }
    }
    sqlite3_finalize(stmt);
    result.ok = true;
    if (columns > 0) {
        set->state_ = adStateOpen;
        result.recordset = std::move(set);
        result.records_affected = -1;  // ADO reports -1 for a query
    } else {
        result.records_affected = sqlite3_changes(impl_->db);
    }
    return result;
}

int Connection::begin_trans() {
    if (impl_->db == nullptr) {
        fail(kNotOpen, "Operation is not allowed when the object is closed.");
        return 0;
    }
    const std::string sql =
        "SAVEPOINT wfc" + std::to_string(impl_->depth + 1);
    char* message = nullptr;
    if (sqlite3_exec(impl_->db, sql.c_str(), nullptr, nullptr, &message) !=
        SQLITE_OK) {
        fail(kUnspecified, message != nullptr ? message : "");
        sqlite3_free(message);
        return 0;
    }
    return ++impl_->depth;
}

bool Connection::commit_trans() {
    if (impl_->db == nullptr || impl_->depth == 0) {
        fail(kNotOpen, "No transaction is active.");
        return false;
    }
    const std::string sql = "RELEASE wfc" + std::to_string(impl_->depth);
    if (sqlite3_exec(impl_->db, sql.c_str(), nullptr, nullptr, nullptr) !=
        SQLITE_OK) {
        fail(kUnspecified, sqlite3_errmsg(impl_->db));
        return false;
    }
    --impl_->depth;
    return true;
}

bool Connection::rollback_trans() {
    if (impl_->db == nullptr || impl_->depth == 0) {
        fail(kNotOpen, "No transaction is active.");
        return false;
    }
    const std::string name = "wfc" + std::to_string(impl_->depth);
    const std::string sql =
        "ROLLBACK TO " + name + "; RELEASE " + name;
    if (sqlite3_exec(impl_->db, sql.c_str(), nullptr, nullptr, nullptr) !=
        SQLITE_OK) {
        fail(kUnspecified, sqlite3_errmsg(impl_->db));
        return false;
    }
    --impl_->depth;
    return true;
}

}  // namespace wfc::adodb
