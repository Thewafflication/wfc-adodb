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
using Value = std::variant<std::monostate, std::int64_t, double, std::string,
                           Blob>;

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
};

// A client-side static cursor: the whole result is read when it is opened, as
// ADO does for adUseClient. Positions are 0-based internally; AbsolutePosition
// is 1-based as in ADO.
class Recordset {
public:
    [[nodiscard]] int state() const noexcept { return state_; }
    [[nodiscard]] std::size_t field_count() const noexcept {
        return fields_.size();
    }
    [[nodiscard]] const FieldInfo& field(std::size_t index) const {
        return fields_.at(index);
    }
    // Index of the field called name (case-insensitive), or -1.
    [[nodiscard]] int field_index(std::string_view name) const noexcept;

    [[nodiscard]] bool bof() const noexcept { return empty() || position_ < 0; }
    [[nodiscard]] bool eof() const noexcept {
        return empty() || position_ >= static_cast<long>(rows_.size());
    }
    [[nodiscard]] long record_count() const noexcept {
        return static_cast<long>(rows_.size());
    }
    // 1-based; -1 (adPosBOF) before the first row, -2 (adPosEOF) after the last.
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

private:
    friend class Connection;
    [[nodiscard]] bool empty() const noexcept { return rows_.empty(); }

    int state_{adStateClosed};
    std::vector<FieldInfo> fields_;
    std::vector<std::vector<Value>> rows_;
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

    // ConnectionString: "Data Source=path" or "Database=path" or "File Name=path"
    // (other keys, among them Provider and Driver, are accepted and ignored). A
    // bare string with no "=" is taken as the path. ":memory:" opens a private
    // in-memory database. Returns false and fills errors() on failure.
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
