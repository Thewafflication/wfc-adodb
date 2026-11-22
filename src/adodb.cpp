// The ADODB-to-SQLite interposer. See include/wfc/adodb/adodb.hpp.

#include <wfc/adodb/adodb.hpp>

#include <algorithm>
#include <cctype>
#include <cstdio>
#include <cstdlib>
#include <cstring>
#include <utility>

#include "odbc.hpp"
#include "sqlite3.h"

namespace wfc::adodb {
namespace {

constexpr long kSyntaxError = -2147217900L;  // 0x80040E14, DB_E_ERRORSINCOMMAND
constexpr long kUnspecified = -2147467259L;  // 0x80004005, E_FAIL
constexpr long kNotOpen = 3704;      // operation not allowed when closed
constexpr long kAlreadyOpen = 3705;  // operation not allowed when open

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

// ----- Recordset
// ------------------------------------------------------------------

namespace {

[[nodiscard]] bool numeric_of(const Value& value, double& out) {
    if (const auto* i = std::get_if<std::int64_t>(&value)) {
        out = static_cast<double>(*i);
        return true;
    }
    if (const auto* d = std::get_if<double>(&value)) {
        out = *d;
        return true;
    }
    return false;
}

// Case-insensitive match with % and * as the wildcards of LIKE.
[[nodiscard]] bool like(std::string_view text, std::string_view pattern) {
    if (pattern.empty()) {
        return text.empty();
    }
    if (pattern.front() == '%' || pattern.front() == '*') {
        for (std::size_t skip = 0; skip <= text.size(); ++skip) {
            if (like(text.substr(skip), pattern.substr(1))) {
                return true;
            }
        }
        return false;
    }
    return !text.empty() &&
           std::tolower(static_cast<unsigned char>(text.front())) ==
               std::tolower(static_cast<unsigned char>(pattern.front())) &&
           like(text.substr(1), pattern.substr(1));
}

[[nodiscard]] bool is_space(const char c) {
    return std::isspace(static_cast<unsigned char>(c)) != 0;
}

// Splits text at a keyword (AND / OR) that stands alone outside quotes.
[[nodiscard]] std::vector<std::string> split_keyword(std::string_view text,
                                                     std::string_view word) {
    std::vector<std::string> parts;
    std::string current;
    bool quoted = false;
    for (std::size_t i = 0; i < text.size(); ++i) {
        const char c = text[i];
        if (c == '\'') {
            quoted = !quoted;
        }
        if (!quoted && i > 0 && is_space(text[i - 1]) &&
            i + word.size() < text.size() &&
            lower(text.substr(i, word.size())) == word &&
            is_space(text[i + word.size()])) {
            parts.push_back(current);
            current.clear();
            i += word.size() - 1;
            continue;
        }
        current.push_back(c);
    }
    parts.push_back(current);
    return parts;
}

[[nodiscard]] bool is_operator_char(const char c) {
    return c == '=' || c == '<' || c == '>';
}

}  // namespace

std::shared_ptr<Recordset> Recordset::create() {
    auto set = std::make_shared<Recordset>();
    set->editable_ = true;
    return set;
}

bool Recordset::append_field(const std::string_view name, const int type,
                             const long defined_size) {
    if (!editable_ || state_ != adStateClosed || name.empty() ||
        field_index(name) >= 0) {
        return false;
    }
    fields_.push_back(FieldInfo{std::string(name), type, defined_size});
    return true;
}

bool Recordset::open_empty() {
    if (!editable_ || state_ != adStateClosed || fields_.empty()) {
        return false;
    }
    state_ = adStateOpen;
    rows_.clear();
    view_.clear();
    position_ = 0;
    return true;
}

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
    if (position_ >= static_cast<long>(view_.size())) {
        return -2;
    }
    return position_ + 1;
}

// Leaving a row with changes pending saves them, as ADO does.
void Recordset::move_first() noexcept {
    static_cast<void>(update());
    position_ = 0;
}

void Recordset::move_last() noexcept {
    static_cast<void>(update());
    position_ = view_.empty() ? 0 : static_cast<long>(view_.size()) - 1;
}

bool Recordset::move_next() noexcept {
    if (eof()) {
        return false;
    }
    static_cast<void>(update());
    ++position_;
    return true;
}

bool Recordset::move_previous() noexcept {
    if (bof()) {
        return false;
    }
    static_cast<void>(update());
    --position_;
    return true;
}

void Recordset::close() noexcept {
    static_cast<void>(update());
    state_ = adStateClosed;
    fields_.clear();
    rows_.clear();
    view_.clear();
    filter_.clear();
    position_ = 0;
}

const Value& Recordset::value(const std::size_t index) const {
    static const Value null_value;
    if (state_ != adStateOpen || bof() || eof() || index >= fields_.size()) {
        return null_value;
    }
    return rows_[view_[static_cast<std::size_t>(position_)]][index];
}

bool Recordset::add_new() {
    if (!editable_ || state_ != adStateOpen) {
        return false;
    }
    static_cast<void>(update());
    rows_.emplace_back(fields_.size());
    view_.push_back(rows_.size() - 1);
    position_ = static_cast<long>(view_.size()) - 1;
    pending_ = true;
    pending_new_ = true;
    return true;
}

bool Recordset::set_value(const std::size_t index, Value value) {
    if (!editable_ || state_ != adStateOpen || bof() || eof() ||
        index >= fields_.size()) {
        return false;
    }
    auto& row = rows_[view_[static_cast<std::size_t>(position_)]];
    if (!pending_) {
        saved_ = row;
        pending_ = true;
        pending_new_ = false;
    }
    row[index] = std::move(value);
    return true;
}

bool Recordset::update() {
    if (!pending_) {
        return false;
    }
    pending_ = false;
    pending_new_ = false;
    saved_.clear();
    return true;
}

bool Recordset::cancel_update() {
    if (!pending_) {
        return false;
    }
    if (pending_new_) {
        rows_.pop_back();
        view_.pop_back();
        position_ = view_.empty() ? 0 : static_cast<long>(view_.size()) - 1;
    } else if (!bof() && !eof()) {
        rows_[view_[static_cast<std::size_t>(position_)]] = saved_;
    }
    pending_ = false;
    pending_new_ = false;
    saved_.clear();
    return true;
}

bool Recordset::delete_current() {
    if (!editable_ || state_ != adStateOpen || bof() || eof()) {
        return false;
    }
    pending_ = false;
    pending_new_ = false;
    saved_.clear();
    const std::size_t at = view_[static_cast<std::size_t>(position_)];
    rows_.erase(rows_.begin() + static_cast<std::ptrdiff_t>(at));
    view_.erase(view_.begin() + position_);
    for (auto& index : view_) {
        if (index > at) {
            --index;
        }
    }
    return true;  // the cursor now rests on the row that followed
}

// One comparison: field op literal. Sets valid to false for a text that is not
// understood.
bool Recordset::matches(const std::vector<Value>& row,
                        const std::string_view criteria, bool& valid) const {
    for (const auto& disjunct : split_keyword(criteria, "or")) {
        bool all = true;
        for (const auto& conjunct : split_keyword(disjunct, "and")) {
            const std::string_view text = trim(conjunct);
            std::size_t cut = 0;
            while (cut < text.size() && !is_operator_char(text[cut]) &&
                   !(is_space(text[cut]) &&
                     lower(trim(text.substr(cut))).rfind("like", 0) == 0)) {
                ++cut;
            }
            if (cut == 0 || cut >= text.size()) {
                valid = false;
                return false;
            }
            std::string_view name = trim(text.substr(0, cut));
            if (name.size() >= 2 && name.front() == '[' && name.back() == ']') {
                name = name.substr(1, name.size() - 2);
            }
            std::string_view rest = trim(text.substr(cut));
            std::string op;
            if (lower(rest).rfind("like", 0) == 0) {
                op = "like";
                rest.remove_prefix(4);
            } else {
                while (!rest.empty() && is_operator_char(rest.front())) {
                    op.push_back(rest.front());
                    rest.remove_prefix(1);
                }
            }
            std::string literal(trim(rest));
            const int column = field_index(name);
            if (column < 0 || literal.empty() ||
                (op != "=" && op != "<>" && op != "<" && op != ">" &&
                 op != "<=" && op != ">=" && op != "like")) {
                valid = false;
                return false;
            }
            const bool quoted = literal.size() >= 2 &&
                                literal.front() == '\'' &&
                                literal.back() == '\'';
            if (quoted) {
                const std::string inner = literal.substr(1, literal.size() - 2);
                literal.clear();
                for (std::size_t i = 0; i < inner.size(); ++i) {
                    literal.push_back(inner[i]);
                    if (inner[i] == '\'' && i + 1 < inner.size() &&
                        inner[i + 1] == '\'') {
                        ++i;
                    }
                }
            }
            const Value& cell = row[static_cast<std::size_t>(column)];
            bool hit = false;
            if (cell.index() != 0) {
                double left = 0;
                if (op == "like") {
                    const auto* shown = std::get_if<std::string>(&cell);
                    hit = shown != nullptr && like(*shown, literal);
                } else {
                    char* end = nullptr;
                    const double right = std::strtod(literal.c_str(), &end);
                    int order = 0;
                    if (!quoted && numeric_of(cell, left) &&
                        end != literal.c_str() && *end == '\0') {
                        order = left < right ? -1 : (left > right ? 1 : 0);
                    } else {
                        std::string shown;
                        if (const auto* t = std::get_if<std::string>(&cell)) {
                            shown = *t;
                        } else if (numeric_of(cell, left)) {
                            char buffer[40];
                            std::snprintf(buffer, sizeof(buffer), "%.15g",
                                          left);
                            shown = buffer;
                        }
                        const int c = lower(shown).compare(lower(literal));
                        order = c < 0 ? -1 : (c > 0 ? 1 : 0);
                    }
                    if (op == "=") {
                        hit = order == 0;
                    } else if (op == "<>") {
                        hit = order != 0;
                    } else if (op == "<") {
                        hit = order < 0;
                    } else if (op == ">") {
                        hit = order > 0;
                    } else if (op == "<=") {
                        hit = order <= 0;
                    } else {
                        hit = order >= 0;
                    }
                }
            }
            all = all && hit;
        }
        if (all) {
            return true;
        }
    }
    return false;
}

void Recordset::rebuild_view() {
    view_.clear();
    bool valid = true;
    for (std::size_t i = 0; i < rows_.size(); ++i) {
        if (filter_.empty() || matches(rows_[i], filter_, valid)) {
            view_.push_back(i);
        }
    }
    position_ = 0;
}

bool Recordset::set_filter(const std::string_view criteria) {
    if (state_ != adStateOpen) {
        return false;
    }
    static_cast<void>(update());
    const std::string text(trim(criteria));
    if (!text.empty()) {
        // A row-shaped probe tells whether the text is understood.
        bool valid = true;
        const std::vector<Value> probe(fields_.size());
        static_cast<void>(matches(probe, text, valid));
        if (!valid) {
            return false;
        }
    }
    filter_ = text;
    rebuild_view();
    return true;
}

bool Recordset::find(const std::string_view criteria, const long skip,
                     const bool forward) {
    if (state_ != adStateOpen) {
        return false;
    }
    bool valid = true;
    const std::vector<Value> probe(fields_.size());
    static_cast<void>(matches(probe, criteria, valid));
    if (!valid) {
        return false;
    }
    const long count = static_cast<long>(view_.size());
    long at = (bof() || eof()) ? (forward ? 0 : count - 1)
                               : position_ + (forward ? skip : -skip);
    while (at >= 0 && at < count) {
        if (matches(rows_[view_[static_cast<std::size_t>(at)]], criteria,
                    valid)) {
            position_ = at;
            return true;
        }
        at += forward ? 1 : -1;
    }
    position_ = forward ? count : -1;
    return true;
}

// ----- Connection
// -----------------------------------------------------------------

struct Connection::Impl {
    sqlite3* db{nullptr};
    int depth{0};                               // open BeginTrans levels
    std::unique_ptr<detail::OdbcSession> odbc;  // set when ODBC is the backend
};

Connection::Connection() : impl_(std::make_unique<Impl>()) {}

Connection::~Connection() {
    close();
}

void Connection::fail(const long number, std::string description) {
    errors_.push_back(Error{number, std::move(description), "wfc-adodb", {}});
}

int Connection::state() const noexcept {
    return impl_->db != nullptr || impl_->odbc != nullptr ? adStateOpen
                                                          : adStateClosed;
}

bool Connection::open(const std::string_view connection_string) {
    if (state() == adStateOpen) {
        fail(kAlreadyOpen, "Operation is not allowed when the object is open.");
        return false;
    }
    if (detail::wants_odbc(connection_string)) {
        auto session = std::make_unique<detail::OdbcSession>();
        if (!session->connect(connection_string, errors_)) {
            return false;
        }
        impl_->odbc = std::move(session);
        return true;
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
    if (impl_->odbc != nullptr) {
        impl_->odbc->disconnect();
        impl_->odbc.reset();
        return;
    }
    if (impl_->db == nullptr) {
        return;
    }
    // ADO rolls back what is still pending when a connection closes.
    sqlite3_exec(impl_->db, "ROLLBACK", nullptr, nullptr, nullptr);
    sqlite3_close_v2(impl_->db);
    impl_->db = nullptr;
    impl_->depth = 0;
}

std::string Connection::version() const {
    if (impl_->odbc != nullptr) {
        return impl_->odbc->version();
    }
    return sqlite3_libversion();
}

std::int64_t Connection::last_insert_id() const noexcept {
    // ODBC has no portable way to ask; a program reads @@IDENTITY itself.
    return impl_->db != nullptr ? sqlite3_last_insert_rowid(impl_->db) : 0;
}

Result Connection::execute(const std::string_view sql,
                           const std::vector<Value>& params) {
    Result result;
    if (impl_->odbc != nullptr) {
        detail::OdbcTable table;
        if (!impl_->odbc->execute(sql, params, table, errors_)) {
            return result;
        }
        result.ok = true;
        if (table.columns) {
            auto set = std::make_shared<Recordset>();
            set->fields_ = std::move(table.fields);
            set->rows_ = std::move(table.rows);
            set->state_ = adStateOpen;
            set->rebuild_view();
            result.recordset = std::move(set);
            result.records_affected = -1;
        } else {
            result.records_affected = table.affected;
        }
        return result;
    }
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
        set->fields_.push_back(
            FieldInfo{sqlite3_column_name(stmt, c) != nullptr
                          ? sqlite3_column_name(stmt, c)
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
                    row.emplace_back(static_cast<std::int64_t>(
                        sqlite3_column_int64(stmt, c)));
                    break;
                case SQLITE_FLOAT:
                    row.emplace_back(sqlite3_column_double(stmt, c));
                    break;
                case SQLITE_TEXT:
                    row.emplace_back(
                        std::string(reinterpret_cast<const char*>(
                                        sqlite3_column_text(stmt, c)),
                                    static_cast<std::size_t>(
                                        sqlite3_column_bytes(stmt, c))));
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
        fail(status == SQLITE_CONSTRAINT ? -2147217873L : kUnspecified,
             message);
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
        set->rebuild_view();
        result.recordset = std::move(set);
        result.records_affected = -1;  // ADO reports -1 for a query
    } else {
        result.records_affected = sqlite3_changes(impl_->db);
    }
    return result;
}

int Connection::begin_trans() {
    if (impl_->odbc != nullptr) {
        return impl_->odbc->begin(errors_);
    }
    if (impl_->db == nullptr) {
        fail(kNotOpen, "Operation is not allowed when the object is closed.");
        return 0;
    }
    const std::string sql = "SAVEPOINT wfc" + std::to_string(impl_->depth + 1);
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
    if (impl_->odbc != nullptr) {
        return impl_->odbc->end(true, errors_);
    }
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
    if (impl_->odbc != nullptr) {
        return impl_->odbc->end(false, errors_);
    }
    if (impl_->db == nullptr || impl_->depth == 0) {
        fail(kNotOpen, "No transaction is active.");
        return false;
    }
    const std::string name = "wfc" + std::to_string(impl_->depth);
    const std::string sql = "ROLLBACK TO " + name + "; RELEASE " + name;
    if (sqlite3_exec(impl_->db, sql.c_str(), nullptr, nullptr, nullptr) !=
        SQLITE_OK) {
        fail(kUnspecified, sqlite3_errmsg(impl_->db));
        return false;
    }
    --impl_->depth;
    return true;
}

}  // namespace wfc::adodb
