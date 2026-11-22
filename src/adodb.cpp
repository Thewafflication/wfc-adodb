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

namespace detail {

// The SQLite handle a connection shares with the recordsets linked to it; it
// is cleared when the connection closes.
struct SqliteHandle {
    sqlite3* db{nullptr};
};

// Where an updatable recordset writes: the table as the query named it, and
// the table column behind each field.
struct SqliteLink {
    std::shared_ptr<SqliteHandle> handle;
    std::string table;
    std::vector<std::string> columns;
};

}  // namespace detail

namespace {

void bind_all(sqlite3_stmt* stmt, const std::vector<Value>& params) {
    for (std::size_t i = 0; i < params.size(); ++i) {
        const Value& v = params[i];
        const int slot = static_cast<int>(i) + 1;
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
}

[[nodiscard]] Value column_value(sqlite3_stmt* stmt, const int c) {
    switch (sqlite3_column_type(stmt, c)) {
        case SQLITE_INTEGER:
            return static_cast<std::int64_t>(sqlite3_column_int64(stmt, c));
        case SQLITE_FLOAT:
            return sqlite3_column_double(stmt, c);
        case SQLITE_TEXT:
            return std::string(
                reinterpret_cast<const char*>(sqlite3_column_text(stmt, c)),
                static_cast<std::size_t>(sqlite3_column_bytes(stmt, c)));
        case SQLITE_BLOB: {
            const auto* bytes =
                static_cast<const unsigned char*>(sqlite3_column_blob(stmt, c));
            const auto size =
                static_cast<std::size_t>(sqlite3_column_bytes(stmt, c));
            return Blob(bytes, bytes + size);
        }
        default:
            return std::monostate{};
    }
}

[[nodiscard]] std::string quoted_name(const std::string_view name) {
    std::string out = "\"";
    for (const char c : name) {
        if (c == '"') {
            out += '"';
        }
        out += c;
    }
    out += '"';
    return out;
}

// Runs one write against the linked table; false with error filled when it
// fails. changed is the number of rows it touched.
[[nodiscard]] bool run_write(const detail::SqliteLink& link,
                             const std::string& sql,
                             const std::vector<Value>& params,
                             std::int64_t& inserted, int& changed,
                             Error& error) {
    sqlite3* db = link.handle->db;
    if (db == nullptr) {
        error = Error{3704, "Operation is not allowed when the object is closed.",
                      "wfc-adodb", {}};
        return false;
    }
    sqlite3_stmt* stmt = nullptr;
    if (sqlite3_prepare_v2(db, sql.c_str(), -1, &stmt, nullptr) != SQLITE_OK) {
        error = Error{-2147217900L, sqlite3_errmsg(db), "wfc-adodb", {}};
        return false;
    }
    bind_all(stmt, params);
    const int status = sqlite3_step(stmt);
    const std::string message = sqlite3_errmsg(db);
    sqlite3_finalize(stmt);
    if (status != SQLITE_DONE) {
        error = Error{status == SQLITE_CONSTRAINT ? -2147217873L : -2147467259L,
                      message, "wfc-adodb", {}};
        return false;
    }
    inserted = sqlite3_last_insert_rowid(db);
    changed = sqlite3_changes(db);
    return true;
}

}  // namespace

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

// Days from 1899-12-30 (the OLE automation origin) to a civil date.
[[nodiscard]] double ole_days(const long year, const long month,
                              const long day) {
    const long y = month <= 2 ? year - 1 : year;
    const long era = (y >= 0 ? y : y - 399) / 400;
    const long yoe = y - era * 400;
    const long mp = (month + 9) % 12;
    const long doy = (153 * mp + 2) / 5 + day - 1;
    const long doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    return static_cast<double>(era * 146097 + doe - 719468 + 25569);
}

// Reads a run of digits from the front of text.
[[nodiscard]] bool read_number(std::string_view& text, long& out) {
    std::size_t used = 0;
    long value = 0;
    while (used < text.size() && text[used] >= '0' && text[used] <= '9' &&
           used < 9) {
        value = value * 10 + (text[used] - '0');
        ++used;
    }
    if (used == 0) {
        return false;
    }
    out = value;
    text.remove_prefix(used);
    return true;
}

[[nodiscard]] bool read_char(std::string_view& text, const char wanted) {
    if (text.empty() || text.front() != wanted) {
        return false;
    }
    text.remove_prefix(1);
    return true;
}

// The text of a #date# literal (M/D/YYYY or YYYY-MM-DD, then an optional time
// H:MM[:SS] [AM|PM]) as an OLE date; false when it is not one.
[[nodiscard]] bool parse_date_literal(std::string_view text, double& out) {
    text = trim(text);
    long a = 0;
    long b = 0;
    long c = 0;
    if (!read_number(text, a) || text.empty()) {
        return false;
    }
    const char separator = text.front();
    if ((separator != '/' && separator != '-') || !read_char(text, separator) ||
        !read_number(text, b) || !read_char(text, separator) ||
        !read_number(text, c)) {
        return false;
    }
    long year = c;
    long month = a;
    long day = b;
    if (separator == '-' && a > 31) {
        year = a;
        month = b;
        day = c;
    }
    if (month < 1 || month > 12 || day < 1 || day > 31) {
        return false;
    }
    double fraction = 0;
    text = trim(text);
    if (!text.empty() && (text.front() == 'T' || text.front() == 'T' + 32)) {
        text.remove_prefix(1);
    }
    if (!text.empty()) {
        long h = 0;
        long m = 0;
        long sec = 0;
        if (!read_number(text, h) || !read_char(text, ':') ||
            !read_number(text, m)) {
            return false;
        }
        if (read_char(text, ':') && !read_number(text, sec)) {
            return false;
        }
        const std::string ampm = lower(trim(text));
        if (!ampm.empty() && ampm != "am" && ampm != "pm") {
            return false;
        }
        if (ampm == "pm" && h < 12) {
            h += 12;
        } else if (ampm == "am" && h == 12) {
            h = 0;
        }
        fraction = (static_cast<double>(h) * 3600.0 +
                    static_cast<double>(m) * 60.0 + static_cast<double>(sec)) /
                   86400.0;
    }
    out = ole_days(year, month, day) + fraction;
    return true;
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
    ids_.clear();
    keys_.clear();
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
    ids_.clear();
    keys_.clear();
    link_.reset();
    has_error_ = false;
    view_.clear();
    filter_.clear();
    sort_.clear();
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
    ids_.push_back(++next_id_);
    keys_.push_back(0);
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

bool Recordset::take_error(Error& out) {
    if (!has_error_) {
        return false;
    }
    out = error_;
    has_error_ = false;
    return true;
}

// Writes the pending row to the linked table; a row of a disconnected
// recordset has nowhere to go and always succeeds.
bool Recordset::flush_row(const std::size_t at) {
    if (link_ == nullptr) {
        return true;
    }
    const auto& row = rows_[at];
    const auto& table = link_->table;
    std::string sql;
    std::vector<Value> params;
    if (pending_new_) {
        std::string names;
        std::string marks;
        for (std::size_t i = 0; i < row.size(); ++i) {
            if (row[i].index() == 0) {
                continue;
            }
            names += names.empty() ? "" : ", ";
            marks += marks.empty() ? "?" : ", ?";
            names += quoted_name(link_->columns[i]);
            params.push_back(row[i]);
        }
        sql = names.empty() ? "INSERT INTO " + table + " DEFAULT VALUES"
                            : "INSERT INTO " + table + " (" + names +
                                  ") VALUES (" + marks + ")";
    } else {
        std::string sets;
        for (std::size_t i = 0; i < row.size(); ++i) {
            if (i < saved_.size() && row[i] == saved_[i]) {
                continue;
            }
            sets += sets.empty() ? "" : ", ";
            sets += quoted_name(link_->columns[i]) + " = ?";
            params.push_back(row[i]);
        }
        if (sets.empty()) {
            return true;
        }
        sql = "UPDATE " + table + " SET " + sets + " WHERE _rowid_ = ?";
        params.emplace_back(keys_[at]);
    }
    std::int64_t inserted = 0;
    int changed = 0;
    Error error;
    if (!run_write(*link_, sql, params, inserted, changed, error)) {
        error_ = std::move(error);
        has_error_ = true;
        return false;
    }
    if (!pending_new_) {
        if (changed == 0) {
            error_ = Error{-2147217864L,
                           "Row cannot be located for updating. Some values "
                           "may have been changed since it was last read.",
                           "wfc-adodb", {}};
            has_error_ = true;
            return false;
        }
        return true;
    }
    keys_[at] = inserted;
    // Read the row back: the table fills in keys and defaults.
    std::string names;
    for (std::size_t i = 0; i < row.size(); ++i) {
        names += names.empty() ? "" : ", ";
        names += quoted_name(link_->columns[i]);
    }
    sqlite3* db = link_->handle->db;
    sqlite3_stmt* stmt = nullptr;
    const std::string select =
        "SELECT " + names + " FROM " + table + " WHERE _rowid_ = ?";
    if (sqlite3_prepare_v2(db, select.c_str(), -1, &stmt, nullptr) ==
        SQLITE_OK) {
        sqlite3_bind_int64(stmt, 1, inserted);
        if (sqlite3_step(stmt) == SQLITE_ROW) {
            for (std::size_t i = 0; i < rows_[at].size(); ++i) {
                rows_[at][i] = column_value(stmt, static_cast<int>(i));
            }
        }
        sqlite3_finalize(stmt);
    }
    return true;
}

bool Recordset::update() {
    if (!pending_) {
        return false;
    }
    if (link_ != nullptr) {
        const std::size_t at =
            pending_new_ ? rows_.size() - 1
                         : view_[static_cast<std::size_t>(position_)];
        if (!flush_row(at)) {
            static_cast<void>(cancel_update());
            return false;
        }
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
        ids_.pop_back();
        keys_.pop_back();
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

long Recordset::bookmark() const noexcept {
    if (state_ != adStateOpen || bof() || eof()) {
        return 0;
    }
    const std::size_t at = view_[static_cast<std::size_t>(position_)];
    return editable_ ? ids_[at] : static_cast<long>(at) + 1;
}

bool Recordset::goto_bookmark(const long mark) {
    if (state_ != adStateOpen || mark <= 0) {
        return false;
    }
    for (std::size_t i = 0; i < view_.size(); ++i) {
        const std::size_t at = view_[i];
        if ((editable_ ? ids_[at] : static_cast<long>(at) + 1) == mark) {
            static_cast<void>(update());
            position_ = static_cast<long>(i);
            return true;
        }
    }
    return false;
}

bool Recordset::delete_current() {
    if (!editable_ || state_ != adStateOpen || bof() || eof()) {
        return false;
    }
    pending_ = false;
    pending_new_ = false;
    saved_.clear();
    const std::size_t at = view_[static_cast<std::size_t>(position_)];
    if (link_ != nullptr && keys_[at] != 0) {
        std::int64_t inserted = 0;
        int changed = 0;
        Error error;
        if (!run_write(*link_,
                       "DELETE FROM " + link_->table + " WHERE _rowid_ = ?",
                       {Value{keys_[at]}}, inserted, changed, error)) {
            error_ = std::move(error);
            has_error_ = true;
            return false;
        }
    }
    rows_.erase(rows_.begin() + static_cast<std::ptrdiff_t>(at));
    ids_.erase(ids_.begin() + static_cast<std::ptrdiff_t>(at));
    keys_.erase(keys_.begin() + static_cast<std::ptrdiff_t>(at));
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
            std::string date_text;
            if (literal.size() >= 3 && literal.front() == '#' &&
                literal.back() == '#') {
                date_text = literal.substr(1, literal.size() - 2);
                double serial = 0;
                if (!parse_date_literal(date_text, serial)) {
                    valid = false;
                    return false;
                }
                char buffer[40];
                std::snprintf(buffer, sizeof(buffer), "%.15g", serial);
                literal = buffer;
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
                        const int c = lower(shown).compare(
                            lower(date_text.empty() ? literal : date_text));
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

namespace {

struct SortKey {
    std::size_t field{0};
    bool descending{false};
};

// -1, 0 or 1; Null sorts first, numbers by value, text without regard to case.
[[nodiscard]] int compare_values(const Value& a, const Value& b) {
    double x = 0;
    double y = 0;
    if (numeric_of(a, x) && numeric_of(b, y)) {
        return x < y ? -1 : (x > y ? 1 : 0);
    }
    if (const auto* s = std::get_if<std::string>(&a)) {
        if (const auto* t = std::get_if<std::string>(&b)) {
            const std::string l = lower(*s);
            const std::string r = lower(*t);
            return l < r ? -1 : (l > r ? 1 : 0);
        }
    }
    if (a.index() != b.index()) {
        return a.index() < b.index() ? -1 : 1;
    }
    if (const auto* p = std::get_if<Blob>(&a)) {
        const auto& q = std::get<Blob>(b);
        return *p < q ? -1 : (*p > q ? 1 : 0);
    }
    return 0;
}

// The sort spec as keys; false when a field is unknown.
[[nodiscard]] bool parse_sort(const Recordset& set, const std::string_view spec,
                              std::vector<SortKey>& keys) {
    std::string_view rest = spec;
    while (!trim(rest).empty()) {
        const auto cut = rest.find(',');
        std::string_view part = trim(rest.substr(0, cut));
        rest = cut == std::string_view::npos ? std::string_view{}
                                             : rest.substr(cut + 1);
        SortKey key;
        const std::string tail = lower(part);
        if (tail.size() > 5 && tail.ends_with(" desc")) {
            key.descending = true;
            part = trim(part.substr(0, part.size() - 5));
        } else if (tail.size() > 4 && tail.ends_with(" asc")) {
            part = trim(part.substr(0, part.size() - 4));
        }
        if (part.size() > 1 && part.front() == '[' && part.back() == ']') {
            part = part.substr(1, part.size() - 2);
        }
        const int index = set.field_index(part);
        if (index < 0) {
            return false;
        }
        key.field = static_cast<std::size_t>(index);
        keys.push_back(key);
    }
    return true;
}

}  // namespace

void Recordset::rebuild_view() {
    view_.clear();
    bool valid = true;
    for (std::size_t i = 0; i < rows_.size(); ++i) {
        if (filter_.empty() || matches(rows_[i], filter_, valid)) {
            view_.push_back(i);
        }
    }
    std::vector<SortKey> keys;
    if (!sort_.empty() && parse_sort(*this, sort_, keys)) {
        std::stable_sort(view_.begin(), view_.end(),
                         [&](const std::size_t a, const std::size_t b) {
                             for (const auto& key : keys) {
                                 const int c = compare_values(
                                     rows_[a][key.field], rows_[b][key.field]);
                                 if (c != 0) {
                                     return key.descending ? c > 0 : c < 0;
                                 }
                             }
                             return false;
                         });
    }
    position_ = 0;
}

bool Recordset::set_sort(const std::string_view spec) {
    if (state_ != adStateOpen) {
        return false;
    }
    static_cast<void>(update());
    std::vector<SortKey> keys;
    const std::string text(trim(spec));
    if (!parse_sort(*this, text, keys)) {
        return false;
    }
    sort_ = text;
    rebuild_view();
    return true;
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
    std::shared_ptr<detail::SqliteHandle> handle{
        std::make_shared<detail::SqliteHandle>()};  // shared with linked sets
};

namespace {

[[nodiscard]] bool is_word_char(const char c) {
    return std::isalnum(static_cast<unsigned char>(c)) != 0 || c == '_';
}

// The index after the quoted name or string that starts at i.
[[nodiscard]] std::size_t skip_quoted(const std::string_view text,
                                      const std::size_t i) {
    const char open = text[i];
    const char close = open == '[' ? ']' : open;
    std::size_t j = i + 1;
    while (j < text.size()) {
        if (text[j] == close) {
            if (close != ']' && j + 1 < text.size() && text[j + 1] == close) {
                j += 2;
                continue;
            }
            return j + 1;
        }
        ++j;
    }
    return text.size();
}

[[nodiscard]] bool is_quote(const char c) {
    return c == '\'' || c == '"' || c == '[' || c == '`';
}

[[nodiscard]] bool word_at(const std::string_view text, const std::size_t i,
                           const std::string_view word) {
    if (i + word.size() > text.size() || (i > 0 && is_word_char(text[i - 1]))) {
        return false;
    }
    for (std::size_t k = 0; k < word.size(); ++k) {
        if (std::tolower(static_cast<unsigned char>(text[i + k])) != word[k]) {
            return false;
        }
    }
    return i + word.size() == text.size() || !is_word_char(text[i + word.size()]);
}

[[nodiscard]] std::size_t skip_space(const std::string_view text,
                                     std::size_t i) {
    while (i < text.size() &&
           std::isspace(static_cast<unsigned char>(text[i])) != 0) {
        ++i;
    }
    return i;
}

// An identifier (bare or quoted) at i; sets end past it and returns its text
// as written, or empty when there is none.
[[nodiscard]] std::string read_identifier(const std::string_view text,
                                          const std::size_t i,
                                          std::size_t& end) {
    if (i >= text.size()) {
        return {};
    }
    if (is_quote(text[i]) && text[i] != '\'') {
        end = skip_quoted(text, i);
        return std::string(text.substr(i, end - i));
    }
    std::size_t j = i;
    while (j < text.size() && is_word_char(text[j])) {
        ++j;
    }
    end = j;
    return std::string(text.substr(i, j - i));
}

// The name an identifier stands for, without its quotes.
[[nodiscard]] std::string bare_name(const std::string& written) {
    if (written.size() >= 2 && is_quote(written.front())) {
        return written.substr(1, written.size() - 2);
    }
    return written;
}

[[nodiscard]] bool same_name(const std::string_view a, const std::string_view b) {
    return a.size() == b.size() &&
           std::equal(a.begin(), a.end(), b.begin(), [](char x, char y) {
               return std::tolower(static_cast<unsigned char>(x)) ==
                      std::tolower(static_cast<unsigned char>(y));
           });
}

struct LinkPlan {
    std::string sql;  // the query with the row id as its first column
    detail::SqliteLink link;
};

// A SELECT of plain columns from one table, with at most WHERE, ORDER BY and
// LIMIT after it, can be written back to that table. Anything else (joins,
// expressions, grouping, aliases of the table) is read-only.
[[nodiscard]] bool plan_link(sqlite3* db, std::string_view sql,
                             LinkPlan& plan) {
    while (!sql.empty() &&
           (sql.back() == ';' ||
            std::isspace(static_cast<unsigned char>(sql.back())) != 0)) {
        sql.remove_suffix(1);
    }
    std::size_t i = skip_space(sql, 0);
    if (!word_at(sql, i, "select")) {
        return false;
    }
    i += 6;
    const std::size_t list_begin = i;
    std::size_t from = std::string_view::npos;
    for (std::size_t j = i; j < sql.size();) {
        const char c = sql[j];
        if (is_quote(c)) {
            j = skip_quoted(sql, j);
            continue;
        }
        if (c == '(' || c == ')' || c == ';') {
            return false;
        }
        if (word_at(sql, j, "from")) {
            from = j;
            break;
        }
        ++j;
    }
    if (from == std::string_view::npos) {
        return false;
    }
    // The select list: * or plain columns, each with an optional alias.
    std::vector<std::string> sources;
    bool star = false;
    std::size_t at = skip_space(sql, list_begin);
    while (at < from) {
        if (sql[at] == '*') {
            star = true;
            at = skip_space(sql, at + 1);
            break;
        }
        std::size_t end = at;
        const std::string column = read_identifier(sql, at, end);
        if (column.empty() || word_at(sql, at, "distinct")) {
            return false;
        }
        sources.push_back(bare_name(column));
        at = skip_space(sql, end);
        if (word_at(sql, at, "as")) {
            at = skip_space(sql, at + 2);
        }
        if (at < from && sql[at] != ',') {
            if (read_identifier(sql, at, end).empty()) {
                return false;
            }
            at = skip_space(sql, end);
        }
        if (at < from && sql[at] == ',') {
            at = skip_space(sql, at + 1);
        } else if (at < from) {
            return false;
        }
    }
    if (at != from || (star && !sources.empty()) ||
        (!star && sources.empty())) {
        return false;
    }
    // The table, then nothing or WHERE / ORDER BY / LIMIT.
    std::size_t end = from + 4;
    const std::string table = read_identifier(sql, skip_space(sql, end), end);
    if (table.empty() || (end < sql.size() && sql[end] == '.')) {
        return false;
    }
    const std::size_t rest = skip_space(sql, end);
    if (rest < sql.size() && !word_at(sql, rest, "where") &&
        !word_at(sql, rest, "order") && !word_at(sql, rest, "limit")) {
        return false;
    }
    for (std::size_t j = rest; j < sql.size();) {
        if (is_quote(sql[j])) {
            j = skip_quoted(sql, j);
            continue;
        }
        if (sql[j] == ';') {
            return false;
        }
        for (const char* word : {"select", "join", "group", "having", "union",
                                 "intersect", "except", "window"}) {
            if (word_at(sql, j, word)) {
                return false;
            }
        }
        ++j;
    }
    // The table's own column names.
    std::vector<std::string> names;
    sqlite3_stmt* stmt = nullptr;
    const std::string pragma = "PRAGMA table_info(" + table + ")";
    if (sqlite3_prepare_v2(db, pragma.c_str(), -1, &stmt, nullptr) !=
        SQLITE_OK) {
        return false;
    }
    while (sqlite3_step(stmt) == SQLITE_ROW) {
        names.emplace_back(
            reinterpret_cast<const char*>(sqlite3_column_text(stmt, 1)));
    }
    sqlite3_finalize(stmt);
    if (names.empty()) {
        return false;
    }
    if (star) {
        sources = names;
    } else {
        for (auto& source : sources) {
            const auto hit =
                std::find_if(names.begin(), names.end(), [&](const auto& n) {
                    return same_name(n, source);
                });
            if (hit == names.end()) {
                return false;
            }
            source = *hit;
        }
    }
    plan.sql = "SELECT _rowid_ AS wfc_rowid, " +
               std::string(sql.substr(list_begin));
    plan.link.table = table;
    plan.link.columns = std::move(sources);
    return true;
}

}  // namespace

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
    impl_->handle->db = db;
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
    impl_->handle->db = nullptr;
    impl_->depth = 0;
}

std::string Connection::version() const {
    if (impl_->odbc != nullptr) {
        return impl_->odbc->version();
    }
    return sqlite3_libversion();
}

std::int64_t Connection::last_insert_id() const noexcept {
    if (impl_->odbc != nullptr) {
        try {
            return impl_->odbc->last_insert_id();
        } catch (...) {
            return 0;
        }
    }
    return impl_->db != nullptr ? sqlite3_last_insert_rowid(impl_->db) : 0;
}

namespace {

// True for "select @@identity" (any case, optional semicolon): the way ADO
// programs ask for the key of the row they just inserted.
[[nodiscard]] bool asks_for_identity(std::string_view sql) {
    while (!sql.empty() &&
           (sql.back() == ';' ||
            std::isspace(static_cast<unsigned char>(sql.back())) != 0)) {
        sql.remove_suffix(1);
    }
    std::size_t i = skip_space(sql, 0);
    if (!word_at(sql, i, "select")) {
        return false;
    }
    i = skip_space(sql, i + 6);
    const std::string_view rest = sql.substr(i);
    constexpr std::string_view wanted = "@@identity";
    return rest.size() == wanted.size() && same_name(rest, wanted);
}

}  // namespace

Result Connection::execute(const std::string_view sql,
                           const std::vector<Value>& params) {
    if (state() == adStateOpen && params.empty() && asks_for_identity(sql)) {
        auto set = std::make_shared<Recordset>();
        set->fields_.push_back(FieldInfo{"Expr1000", adInteger, 0});
        set->rows_.push_back({Value{last_insert_id()}});
        set->state_ = adStateOpen;
        set->rebuild_view();
        Result result;
        result.ok = true;
        result.records_affected = -1;
        result.recordset = std::move(set);
        return result;
    }
    return run(sql, params, false);
}

Result Connection::execute_updatable(const std::string_view sql,
                                     const std::vector<Value>& params) {
    LinkPlan plan;
    if (impl_->odbc == nullptr && impl_->db != nullptr &&
        plan_link(impl_->db, sql, plan)) {
        const std::size_t before = errors_.size();
        Result result = run(plan.sql, params, true);
        if (result.ok && result.recordset != nullptr) {
            auto& set = *result.recordset;
            plan.link.handle = impl_->handle;
            set.link_ = std::make_shared<detail::SqliteLink>(plan.link);
            set.editable_ = true;
            for (std::size_t i = 0; i < set.rows_.size(); ++i) {
                set.ids_.push_back(static_cast<long>(i) + 1);
            }
            set.next_id_ = static_cast<long>(set.rows_.size());
            return result;
        }
        errors_.resize(before);  // e.g. a view has no row id: read-only
    }
    return run(sql, params, false);
}

Result Connection::run(const std::string_view sql,
                       const std::vector<Value>& params, const bool keyed) {
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
    bind_all(stmt, params);

    auto set = std::make_shared<Recordset>();
    const int columns = sqlite3_column_count(stmt);
    for (int c = 0; c < columns; ++c) {
        set->fields_.push_back(
            FieldInfo{sqlite3_column_name(stmt, c) != nullptr
                          ? sqlite3_column_name(stmt, c)
                          : "",
                      type_from_declared(sqlite3_column_decltype(stmt, c))});
    }
    if (keyed) {
        set->fields_.erase(set->fields_.begin());
    }
    int status = SQLITE_OK;
    while ((status = sqlite3_step(stmt)) == SQLITE_ROW) {
        std::vector<Value> row;
        row.reserve(static_cast<std::size_t>(columns));
        for (int c = 0; c < columns; ++c) {
            row.push_back(column_value(stmt, c));
        }
        if (keyed) {
            const auto* key = std::get_if<std::int64_t>(&row.front());
            set->keys_.push_back(key != nullptr ? *key : 0);
            row.erase(row.begin());
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
