// wfc-adodb: the ADODB object model (Connection, Command, Recordset, Fields,
// Errors) implemented as an interposer over SQLite. A Visual Basic program asks
// for ADODB; this library answers from a SQLite database. Internal to the WFC
// runtime for now; the C++ surface mirrors the ADO members so the runtime
// binding stays a thin table.

#ifndef WFC_ADODB_ADODB_HPP
#define WFC_ADODB_ADODB_HPP

#include <cstddef>
#include <cstdint>
#include <memory>
#include <string>
#include <string_view>
#include <variant>
#include <vector>

namespace wfc::adodb {

// A field value. monostate is Null.
using Blob = std::vector<unsigned char>;
using Value =
    std::variant<std::monostate, std::int64_t, double, std::string, Blob>;

// DataTypeEnum values used when reporting a field.
enum DataType : int {
    adEmpty = 0,
    adInteger = 3,
    adDouble = 5,
    adBigInt = 20,
    adVarWChar = 202,
    adLongVarBinary = 205,
};

// ObjectStateEnum.
enum ObjectState : int { adStateClosed = 0, adStateOpen = 1 };

// One ADODB.Error (Connection.Errors).
struct Error {
    long number{0};
    std::string description;
    std::string source{"wfc-adodb"};
    std::string sql_state;
};

struct FieldInfo {
    std::string name;
    int type{adEmpty};
    long defined_size{0};
};

// A client-side static cursor: the whole result is read when it is opened, as
// ADO does for adUseClient. Positions are 0-based internally; AbsolutePosition
// is 1-based as in ADO. A recordset that comes from Connection::execute is
// read-only. One made with Recordset::create (fields appended, then
// open_empty) is disconnected and editable: add_new, set_value, update,
// cancel_update, delete_current.
class Recordset {
public:
    // A closed, disconnected recordset whose fields can be appended.
    [[nodiscard]] static std::shared_ptr<Recordset> create();

    [[nodiscard]] int state() const noexcept { return state_; }
    [[nodiscard]] bool editable() const noexcept { return editable_; }
    [[nodiscard]] std::size_t field_count() const noexcept {
        return fields_.size();
    }
    [[nodiscard]] const FieldInfo& field(std::size_t index) const {
        return fields_.at(index);
    }
    // Index of the field called name (case-insensitive), or -1.
    [[nodiscard]] int field_index(std::string_view name) const noexcept;

    // Fields.Append; only while the recordset is closed and disconnected.
    bool append_field(std::string_view name, int type, long defined_size);
    // Recordset.Open without a source: opens with the appended fields, no rows.
    bool open_empty();

    [[nodiscard]] bool bof() const noexcept { return empty() || position_ < 0; }
    [[nodiscard]] bool eof() const noexcept {
        return empty() || position_ >= static_cast<long>(view_.size());
    }
    // The rows that pass the filter.
    [[nodiscard]] long record_count() const noexcept {
        return static_cast<long>(view_.size());
    }
    // 1-based; -1 (adPosBOF) before the first row, -2 (adPosEOF) after the
    // last.
    [[nodiscard]] long absolute_position() const noexcept;

    void move_first() noexcept;
    void move_last() noexcept;
    // Moving past either end leaves the cursor at BOF/EOF; moving further is an
    // error (3021) in ADO, reported here as false.
    bool move_next() noexcept;
    bool move_previous() noexcept;
    void close() noexcept;

    // The value of a field of the current row; Null when there is none.
    [[nodiscard]] const Value& value(std::size_t index) const;

    // Editing (disconnected recordsets only; each returns false when it is not
    // allowed here or there is no current record).
    bool add_new();
    bool set_value(std::size_t index, Value value);
    bool update();
    bool cancel_update();
    bool delete_current();

    // Filter criteria: field op literal, joined by AND / OR; an empty text
    // removes the filter. False when the text is not understood.
    bool set_filter(std::string_view criteria);
    [[nodiscard]] const std::string& filter() const noexcept { return filter_; }
    // Sort: "field [ASC|DESC], field ..."; an empty text restores the original
    // order. False when a field is unknown. The cursor goes to the first row.
    bool set_sort(std::string_view spec);
    [[nodiscard]] const std::string& sort() const noexcept { return sort_; }
    // Find: moves to the first row at or after the current one (plus skip)
    // that meets the criteria; at EOF when there is none.
    bool find(std::string_view criteria, long skip = 0, bool forward = true);

private:
    friend class Connection;
    [[nodiscard]] bool empty() const noexcept { return view_.empty(); }
    [[nodiscard]] bool matches(const std::vector<Value>& row,
                               std::string_view criteria, bool& valid) const;
    void rebuild_view();

    int state_{adStateClosed};
    bool editable_{false};
    bool pending_{false};       // add_new or an edit not yet updated
    bool pending_new_{false};   // the pending row is a new one
    std::vector<Value> saved_;  // the row as it was before an edit
    std::vector<FieldInfo> fields_;
    std::vector<std::vector<Value>> rows_;
    std::vector<std::size_t> view_;  // indexes into rows_ in cursor order
    std::string filter_;
    std::string sort_;
    long position_{0};
};

struct Result {
    // Set when the statement returned columns; closed otherwise.
    std::shared_ptr<Recordset> recordset;
    long records_affected{0};
    bool ok{false};
};

class Connection {
public:
    Connection();
    ~Connection();
    Connection(const Connection&) = delete;
    Connection& operator=(const Connection&) = delete;

    // ConnectionString: a string with a DSN, FILEDSN or Driver key (or
    // Provider=MSDASQL) goes to the ODBC driver manager, so any database with
    // an ODBC driver works (Windows only for now). Otherwise SQLite: "Data
    // Source=path" or "Database=path" or "File Name=path" (other keys,
    // among them Provider, are accepted and ignored); a bare string with no
    // "=" is taken as the path; ":memory:" opens a private in-memory
    // database. Under ODBC the SQL is the database's own, BeginTrans does not
    // nest and last_insert_id is 0. Returns false and fills errors() on
    // failure.
    bool open(std::string_view connection_string);
    void close() noexcept;
    [[nodiscard]] int state() const noexcept;

    // Execute one statement with positional parameters (? placeholders).
    Result execute(std::string_view sql, const std::vector<Value>& params = {});

    // Transactions; BeginTrans returns the nesting level, 0 on failure.
    int begin_trans();
    bool commit_trans();
    bool rollback_trans();

    [[nodiscard]] const std::vector<Error>& errors() const noexcept {
        return errors_;
    }
    void clear_errors() noexcept { errors_.clear(); }
    [[nodiscard]] std::string version() const;
    // ADO's rowid of the last insert, for @@IDENTITY.
    [[nodiscard]] std::int64_t last_insert_id() const noexcept;

private:
    struct Impl;
    std::unique_ptr<Impl> impl_;
    std::vector<Error> errors_;
    void fail(long number, std::string description);
};

// The path of a connection string, or empty when it names none.
[[nodiscard]] std::string connection_string_path(std::string_view text);

}  // namespace wfc::adodb

#endif  // WFC_ADODB_ADODB_HPP
