// The ODBC backend of wfc-adodb: a Connection whose connection string names a
// DSN or a Driver talks to that database through the ODBC driver manager
// (odbc32 on Windows) instead of SQLite. Internal; Connection owns one.

#ifndef WFC_ADODB_ODBC_HPP
#define WFC_ADODB_ODBC_HPP

#include <wfc/adodb/adodb.hpp>

#include <memory>
#include <string>
#include <string_view>
#include <vector>

namespace wfc::adodb::detail {

// What one statement produced, before it becomes a Recordset.
struct OdbcTable {
    std::vector<FieldInfo> fields;
    std::vector<std::vector<Value>> rows;
    long affected{0};
    bool columns{false};  // the statement returned a result set
};

// True when the connection string asks for ODBC: it has a DSN, FILEDSN or
// Driver key, or names the MSDASQL provider (OLE DB over ODBC).
[[nodiscard]] bool wants_odbc(std::string_view connection_string);

class OdbcSession {
public:
    OdbcSession();
    ~OdbcSession();
    OdbcSession(const OdbcSession&) = delete;
    OdbcSession& operator=(const OdbcSession&) = delete;

    // Each failing call appends the driver's diagnostics to errors.
    bool connect(std::string_view connection_string,
                 std::vector<Error>& errors);
    void disconnect() noexcept;
    [[nodiscard]] bool connected() const noexcept;

    bool execute(std::string_view sql, const std::vector<Value>& params,
                 OdbcTable& out, std::vector<Error>& errors);

    // BeginTrans returns the level (1), 0 on failure; ODBC has no nesting.
    int begin(std::vector<Error>& errors);
    bool end(bool commit, std::vector<Error>& errors);
    [[nodiscard]] int depth() const noexcept;

    [[nodiscard]] std::string version() const;

private:
    struct Impl;
    std::unique_ptr<Impl> impl_;
};

}  // namespace wfc::adodb::detail

#endif  // WFC_ADODB_ODBC_HPP
