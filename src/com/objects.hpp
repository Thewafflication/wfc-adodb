// The ADODB COM objects of msado15.dll: Connection, Recordset, Command, Stream,
// Record and the collections and members they hand out. The members are
// described by tables (dispatch.hpp) and match the dispinterfaces of
// adodb.idl, whose ids and parameters are those of ADO 6.1. Windows only;
// internal to the msado15.dll build.

#ifndef WFC_ADODB_COM_OBJECTS_HPP
#define WFC_ADODB_COM_OBJECTS_HPP

#include "dispatch.hpp"

#include <wfc/adodb/adodb.hpp>

#include <map>
#include <memory>
#include <string>
#include <vector>

// The interface ids midl writes from adodb.idl (adodb_i.c).
extern "C" {
extern const IID DIID__Connection;
extern const IID DIID__Recordset;
extern const IID DIID__Command;
extern const IID DIID__Parameter;
extern const IID DIID_Fields;
extern const IID DIID_Field;
extern const IID DIID_Parameters;
extern const IID DIID_Properties;
extern const IID DIID_Property;
extern const IID DIID_Errors;
extern const IID DIID_Error;
extern const IID DIID__Record;
extern const IID DIID__Stream;
}

namespace wfc::adodb::com {

inline constexpr const char* kClosedText =
    "Operation is not allowed when the object is closed.";
inline constexpr const char* kOpenText =
    "Operation is not allowed when the object is open.";
inline constexpr const char* kNoRecordText =
    "Either BOF or EOF is True, or the current record has been deleted. "
    "Requested operation requires a current record.";
inline constexpr const char* kNoConnectionText =
    "The connection cannot be used to perform this operation. It is either "
    "closed or invalid in this context.";
inline constexpr const char* kNotFoundText =
    "Item cannot be found in the collection corresponding to the requested "
    "name or ordinal.";
inline constexpr const char* kBadArgumentText =
    "Arguments are of the wrong type, are out of acceptable range, or are in "
    "conflict with one another.";
inline constexpr const char* kReadOnlyText =
    "Current Recordset does not support updating. This may be a limitation of "
    "the provider, or of the selected locktype.";
inline constexpr const char* kUnsupportedText =
    "The operation requested by the application is not supported by the "
    "provider.";

// Throws the last error of the connection as the call's error.
[[noreturn]] void fail_from(const Connection& connection);

// An owned reference to a COM object.
class Ref {
public:
    Ref() = default;
    explicit Ref(IUnknown* object) : object_(object) {
        if (object_ != nullptr) {
            object_->AddRef();
        }
    }
    Ref(const Ref& other) : Ref(other.object_) {}
    Ref& operator=(const Ref& other) {
        if (this != &other) {
            if (other.object_ != nullptr) {
                other.object_->AddRef();
            }
            if (object_ != nullptr) {
                object_->Release();
            }
            object_ = other.object_;
        }
        return *this;
    }
    ~Ref() {
        if (object_ != nullptr) {
            object_->Release();
        }
    }
    [[nodiscard]] IUnknown* get() const noexcept { return object_; }
    [[nodiscard]] bool empty() const noexcept { return object_ == nullptr; }

private:
    IUnknown* object_{nullptr};
};

// The object of class T behind a COM pointer, or null when the pointer is
// null or belongs to another class (not borrowed: the caller does not own it).
template <class T>
T* unwrap(IUnknown* object) {
    if (object == nullptr) {
        return nullptr;
    }
    void* found = nullptr;
    if (FAILED(object->QueryInterface(kPrivateObject, &found)) ||
        found == nullptr) {
        return nullptr;
    }
    auto* base = static_cast<ObjectBase*>(found);
    T* result =
        base->tag() == T::static_tag() ? static_cast<T*>(base) : nullptr;
    object->Release();  // the reference the query added
    return result;
}

// ----- Property and Properties ---------------------------------------------

class PropertyObj : public DispatchObject<PropertyObj> {
public:
    std::string name;
    Var value;
    long type{12};  // adVariant
    long attributes{0};
    static const IID& iid() { return DIID_Property; }
    static const Table<PropertyObj>& members();
};

class PropertiesObj : public DispatchObject<PropertiesObj> {
public:
    ~PropertiesObj() override;
    std::vector<PropertyObj*> items;  // each holds a reference
    void add(const std::string& name, const Var& value, long type = 12,
             long attributes = 0);
    static const IID& iid() { return DIID_Properties; }
    static const Table<PropertiesObj>& members();
};

// The collection an object hands out as its Properties: `owner` keeps it.
IDispatch* make_properties(PropertiesObj*& slot);

// ----- Error and Errors ------------------------------------------------------

class ErrorObj : public DispatchObject<ErrorObj> {
public:
    Error error;
    static const IID& iid() { return DIID_Error; }
    static const Table<ErrorObj>& members();
};

class ErrorsObj : public DispatchObject<ErrorsObj> {
public:
    std::shared_ptr<Connection> native;
    static const IID& iid() { return DIID_Errors; }
    static const Table<ErrorsObj>& members();
};

// ----- Connection ------------------------------------------------------------

class ConnectionObj : public DispatchObject<ConnectionObj> {
public:
    ~ConnectionObj() override;
    std::shared_ptr<Connection> native = std::make_shared<Connection>();
    std::string connection_string;
    std::string provider;
    std::string default_database;
    long cursor_location{2};
    long command_timeout{30};
    long connection_timeout{15};
    long isolation_level{4096};
    long attributes{0};
    long mode{0};
    PropertiesObj* properties{nullptr};
    static const IID& iid() { return DIID__Connection; }
    static const Table<ConnectionObj>& members();
};

// ----- Recordset, Fields and Field
// -----------------------------------------------

class RecordsetObj : public DispatchObject<RecordsetObj> {
public:
    ~RecordsetObj() override;
    std::shared_ptr<Recordset> set = Recordset::create();
    Ref connection;  // the Connection object
    std::shared_ptr<Connection> native;
    Ref command;                       // the Command that made it, if any
    std::string source;                // the SQL it was opened from
    std::vector<Value> source_values;  // and the parameters of that SQL
    std::string data_member;
    std::string index;
    long cursor_type{0};
    long lock_type{1};
    long cursor_location{2};
    long page_size{10};
    long cache_size{1};
    long max_records{0};
    long marshal_options{0};
    bool stay_in_sync{true};
    std::map<std::size_t, long> chunk_offsets;     // GetChunk, by field
    std::map<std::size_t, long> field_attributes;  // Field.Attributes, by field
    std::map<std::size_t, long> field_precision;
    std::map<std::size_t, long> field_scale;
    PropertiesObj* properties{nullptr};

    void need_open() const;
    void need_record() const;
    void need_editable() const;
    // Raises an error left by a write to the table while a call flushed a row.
    void raise_pending();
    // Runs the SQL on the native connection and takes the result.
    void run(const std::string& sql, const std::vector<Value>& values);
    static const IID& iid() { return DIID__Recordset; }
    static const Table<RecordsetObj>& members();
};

class FieldsObj : public DispatchObject<FieldsObj> {
public:
    explicit FieldsObj(RecordsetObj* owner);
    ~FieldsObj() override;
    RecordsetObj* owner;
    static const IID& iid() { return DIID_Fields; }
    static const Table<FieldsObj>& members();
};

class FieldObj : public DispatchObject<FieldObj> {
public:
    FieldObj(RecordsetObj* owner, std::size_t index);
    ~FieldObj() override;
    RecordsetObj* owner;
    std::size_t index;
    static const IID& iid() { return DIID_Field; }
    static const Table<FieldObj>& members();
};

// Creates a Recordset object around a native recordset (null for a closed one).
RecordsetObj* make_recordset(std::shared_ptr<Recordset> set,
                             const Ref& connection,
                             std::shared_ptr<Connection> native,
                             long lock_type);

// The position of a field named or numbered by an Item argument.
std::size_t field_index(const Recordset& set, const VARIANT& key);

// The native connection behind an ActiveConnection argument: a Connection
// object, or a connection string opened on the spot.
std::shared_ptr<Connection> native_of(IUnknown* object, const VARIANT* text,
                                      Ref& keep);

// Saving and loading a recordset (ADO's XML persistence format).
std::string save_xml(const Recordset& set);
std::shared_ptr<Recordset> load_xml(const std::string& text);

// ----- Command, Parameters and Parameter
// ---------------------------------------

class ParameterObj : public DispatchObject<ParameterObj> {
public:
    ~ParameterObj() override;
    std::string name;
    long type{0};
    long direction{1};
    long size{0};
    long attributes{0};
    long precision{0};
    long numeric_scale{0};
    Var value;
    PropertiesObj* properties{nullptr};
    static const IID& iid() { return DIID__Parameter; }
    static const Table<ParameterObj>& members();
};

class ParametersObj : public DispatchObject<ParametersObj> {
public:
    ~ParametersObj() override;
    std::vector<ParameterObj*> items;  // each holds a reference
    [[nodiscard]] std::size_t locate(const VARIANT& index) const;
    [[nodiscard]] std::vector<Value> values() const;
    static const IID& iid() { return DIID_Parameters; }
    static const Table<ParametersObj>& members();
};

class CommandObj : public DispatchObject<CommandObj> {
public:
    ~CommandObj() override;
    Ref connection;
    std::shared_ptr<Connection> native;
    std::string text;
    std::string name;
    std::string dialect;
    long type{8};
    long timeout{30};
    bool prepared{false};
    bool named_parameters{false};
    ParametersObj* parameters{nullptr};
    PropertiesObj* properties{nullptr};
    ParametersObj* params();
    static const IID& iid() { return DIID__Command; }
    static const Table<CommandObj>& members();
};

// ----- Stream and Record
// ------------------------------------------------------------

class StreamObj : public DispatchObject<StreamObj> {
public:
    Blob data;
    long position{0};
    long type{2};             // adTypeText
    long line_separator{-1};  // adCRLF
    long mode{0};
    long state{0};
    std::string charset{"Unicode"};
    static const IID& iid() { return DIID__Stream; }
    static const Table<StreamObj>& members();
};

class RecordObj : public DispatchObject<RecordObj> {
public:
    ~RecordObj() override;
    RecordsetObj* holder{nullptr};  // its one row, as a disconnected set
    Ref connection;
    std::string source;
    long mode{0};
    long state{0};
    PropertiesObj* properties{nullptr};
    static const IID& iid() { return DIID__Record; }
    static const Table<RecordObj>& members();
};

// ----- class factories
// -----------------------------------------------------------------

// The class factory of class `index` of the library's class table.
HRESULT make_class_factory(std::size_t index, REFIID iid, void** object);

// True when no object of the library lives and the server is not locked.
bool can_unload() noexcept;

}  // namespace wfc::adodb::com

#endif  // WFC_ADODB_COM_OBJECTS_HPP
