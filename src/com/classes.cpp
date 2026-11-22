// The ADODB classes as late-bound COM objects over the wfc-adodb core:
// Connection, Command, Parameters, Parameter, Recordset, Fields and Field,
// with the members the WFC built-in classes offer. Windows only; internal to
// the msado15.dll build. See classes.hpp.

#include "classes.hpp"

#include "dispatch.hpp"

#include <wfc/adodb/adodb.hpp>

#include <algorithm>
#include <cstdio>
#include <memory>
#include <new>
#include <string>
#include <vector>

namespace wfc::adodb::com {
namespace {

constexpr const char* kClosedText =
    "Operation is not allowed when the object is closed.";
constexpr const char* kOpenText =
    "Operation is not allowed when the object is open.";
constexpr const char* kNoRecordText =
    "Either BOF or EOF is True, or the current record has been deleted. "
    "Requested operation requires a current record.";
constexpr const char* kNoConnectionText =
    "The connection cannot be used to perform this operation. It is either "
    "closed or invalid in this context.";
constexpr const char* kNotFoundText =
    "Item cannot be found in the collection corresponding to the requested "
    "name or ordinal.";
constexpr const char* kBadArgumentText =
    "Arguments are of the wrong type, are out of acceptable range, or are in "
    "conflict with one another.";
constexpr const char* kReadOnlyText =
    "Current Recordset does not support updating. This may be a limitation of "
    "the provider, or of the selected locktype.";

// The private interfaces the objects answer to each other, so a Recordset can
// find the native connection behind a Connection object, and so on.
constexpr GUID kPrivateConnection = {
    0x6f0b0b5a,
    0x2c1e,
    0x4d7a,
    {0x9a, 0x01, 0x57, 0x6e, 0xa5, 0x0c, 0x11, 0x01}};
constexpr GUID kPrivateParameter = {
    0x6f0b0b5a,
    0x2c1e,
    0x4d7a,
    {0x9a, 0x01, 0x57, 0x6e, 0xa5, 0x0c, 0x11, 0x02}};
constexpr GUID kPrivateRecordset = {
    0x6f0b0b5a,
    0x2c1e,
    0x4d7a,
    {0x9a, 0x01, 0x57, 0x6e, 0xa5, 0x0c, 0x11, 0x03}};

[[noreturn]] void fail_from(const Connection& connection) {
    if (!connection.errors().empty()) {
        const auto& last = connection.errors().back();
        throw ComError{last.number, last.description};
    }
    throw ComError{5, "Invalid procedure call or argument"};
}

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

// Asks an object of ours for its private face; null when it is not ours.
template <class T>
T* unwrap(IUnknown* object, const GUID& iid) {
    if (object == nullptr) {
        return nullptr;
    }
    void* found = nullptr;
    if (FAILED(object->QueryInterface(iid, &found)) || found == nullptr) {
        return nullptr;
    }
    static_cast<IUnknown*>(static_cast<IDispatch*>(static_cast<T*>(found)))
        ->Release();
    return static_cast<T*>(found);
}

// ----- Parameter and Parameters -------------------------------------------

class ParameterObj : public DispatchObject<ParameterObj> {
public:
    std::string name;
    long type{0};
    long direction{1};
    long size{0};
    long attributes{0};
    long precision{0};
    long numeric_scale{0};
    Var value;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid,
                                             void** object) override {
        if (iid == kPrivateParameter) {
            *object = this;
            AddRef();
            return S_OK;
        }
        return DispatchObject::QueryInterface(iid, object);
    }
    static const std::vector<MemberEntry<ParameterObj>>& members();
};

const std::vector<MemberEntry<ParameterObj>>& ParameterObj::members() {
    static const std::vector<MemberEntry<ParameterObj>> table = {
        {L"Name", 1, [](ParameterObj& s, Call& c) { c.put_text(s.name); },
         [](ParameterObj& s, Call& c) { s.name = c.to_text(c.count() - 1); }},
        {L"Type", 2, [](ParameterObj& s, Call& c) { c.put_long(s.type); },
         [](ParameterObj& s, Call& c) {
             s.type = c.to_long(c.count() - 1, 0);
         }},
        {L"Direction", 3,
         [](ParameterObj& s, Call& c) { c.put_long(s.direction); },
         [](ParameterObj& s, Call& c) {
             s.direction = c.to_long(c.count() - 1, 1);
         }},
        {L"Size", 4, [](ParameterObj& s, Call& c) { c.put_long(s.size); },
         [](ParameterObj& s, Call& c) {
             s.size = c.to_long(c.count() - 1, 0);
         }},
        {L"Attributes", 5,
         [](ParameterObj& s, Call& c) { c.put_long(s.attributes); },
         [](ParameterObj& s, Call& c) {
             s.attributes = c.to_long(c.count() - 1, 0);
         }},
        {L"Precision", 6,
         [](ParameterObj& s, Call& c) { c.put_long(s.precision); },
         [](ParameterObj& s, Call& c) {
             s.precision = c.to_long(c.count() - 1, 0);
         }},
        {L"NumericScale", 7,
         [](ParameterObj& s, Call& c) { c.put_long(s.numeric_scale); },
         [](ParameterObj& s, Call& c) {
             s.numeric_scale = c.to_long(c.count() - 1, 0);
         }},
        {L"Value", DISPID_VALUE,
         [](ParameterObj& s, Call& c) { c.put_var(s.value); },
         [](ParameterObj& s, Call& c) {
             s.value = dereferenced(c.args.back());
         }},
    };
    return table;
}

class ParametersObj : public DispatchObject<ParametersObj> {
public:
    std::vector<ParameterObj*> items;  // each holds a reference

    ~ParametersObj() override {
        for (auto* item : items) {
            item->Release();
        }
    }
    [[nodiscard]] std::size_t locate(const VARIANT& index) const {
        const Value key = to_value(index);
        if (const auto* name = std::get_if<std::string>(&key)) {
            for (std::size_t i = 0; i < items.size(); ++i) {
                if (_stricmp(items[i]->name.c_str(), name->c_str()) == 0) {
                    return i;
                }
            }
            throw ComError{3265, kNotFoundText};
        }
        const auto* number = std::get_if<std::int64_t>(&key);
        if (number == nullptr || *number < 0 ||
            static_cast<std::size_t>(*number) >= items.size()) {
            throw ComError{3265, kNotFoundText};
        }
        return static_cast<std::size_t>(*number);
    }
    [[nodiscard]] std::vector<Value> values() const {
        std::vector<Value> out;
        for (const auto* item : items) {
            out.push_back(to_value(item->value.ref()));
        }
        return out;
    }
    static const std::vector<MemberEntry<ParametersObj>>& members();
};

const std::vector<MemberEntry<ParametersObj>>& ParametersObj::members() {
    static const std::vector<MemberEntry<ParametersObj>> table = {
        {L"Count", 1,
         [](ParametersObj& s, Call& c) {
             c.put_long(static_cast<long>(s.items.size()));
         },
         nullptr},
        {L"Item", DISPID_VALUE,
         [](ParametersObj& s, Call& c) {
             if (c.count() < 1) {
                 throw ComError{450, "Wrong number of arguments"};
             }
             c.put_object(s.items[s.locate(c.args[0])]);
         },
         nullptr},
        {L"Append", 2,
         [](ParametersObj& s, Call& c) {
             auto* parameter =
                 unwrap<ParameterObj>(c.object(0), kPrivateParameter);
             if (parameter == nullptr) {
                 throw ComError{3001, kBadArgumentText};
             }
             parameter->AddRef();
             s.items.push_back(parameter);
         },
         nullptr},
        {L"Delete", 3,
         [](ParametersObj& s, Call& c) {
             if (c.count() < 1) {
                 throw ComError{450, "Wrong number of arguments"};
             }
             const std::size_t at = s.locate(c.args[0]);
             s.items[at]->Release();
             s.items.erase(s.items.begin() + static_cast<std::ptrdiff_t>(at));
         },
         nullptr},
        {L"Refresh", 4, [](ParametersObj&, Call&) {}, nullptr},
        {L"_NewEnum", DISPID_NEWENUM,
         [](ParametersObj& s, Call& c) {
             std::vector<Var> list;
             for (auto* item : s.items) {
                 Var v;
                 v.set_object(item);
                 list.push_back(v);
             }
             auto* enumerator = new Enumerator(std::move(list));
             if (c.result != nullptr) {
                 c.result->vt = VT_UNKNOWN;
                 c.result->punkVal = enumerator;
             } else {
                 enumerator->Release();
             }
         },
         nullptr},
    };
    return table;
}

// ----- Connection -----------------------------------------------------------

class RecordsetObj;
IDispatch* make_recordset(std::shared_ptr<Recordset> set, const Ref& connection,
                          std::shared_ptr<Connection> native, long lock_type);

class ConnectionObj : public DispatchObject<ConnectionObj> {
public:
    std::shared_ptr<Connection> native = std::make_shared<Connection>();
    std::string connection_string;
    std::string provider;
    long cursor_location{2};
    long command_timeout{30};

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid,
                                             void** object) override {
        if (iid == kPrivateConnection) {
            *object = this;
            AddRef();
            return S_OK;
        }
        return DispatchObject::QueryInterface(iid, object);
    }
    static const std::vector<MemberEntry<ConnectionObj>>& members();
};

const std::vector<MemberEntry<ConnectionObj>>& ConnectionObj::members() {
    static const std::vector<MemberEntry<ConnectionObj>> table = {
        {L"ConnectionString", 1,
         [](ConnectionObj& s, Call& c) { c.put_text(s.connection_string); },
         [](ConnectionObj& s, Call& c) {
             s.connection_string = c.to_text(c.count() - 1);
         }},
        {L"Provider", 2,
         [](ConnectionObj& s, Call& c) { c.put_text(s.provider); },
         [](ConnectionObj& s, Call& c) {
             s.provider = c.to_text(c.count() - 1);
         }},
        {L"CursorLocation", 3,
         [](ConnectionObj& s, Call& c) { c.put_long(s.cursor_location); },
         [](ConnectionObj& s, Call& c) {
             s.cursor_location = c.to_long(c.count() - 1, 2);
         }},
        {L"CommandTimeout", 4,
         [](ConnectionObj& s, Call& c) { c.put_long(s.command_timeout); },
         [](ConnectionObj& s, Call& c) {
             s.command_timeout = c.to_long(c.count() - 1, 30);
         }},
        {L"State", 5,
         [](ConnectionObj& s, Call& c) { c.put_long(s.native->state()); },
         nullptr},
        {L"Version", 6, [](ConnectionObj&, Call& c) { c.put_text("2.8"); },
         nullptr},
        {L"Open", 7,
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
        {L"Close", 8,
         [](ConnectionObj& s, Call&) {
             if (s.native->state() == adStateClosed) {
                 throw ComError{3704, kClosedText};
             }
             s.native->close();
         },
         nullptr},
        {L"Execute", 9,
         [](ConnectionObj& s, Call& c) {
             s.native->clear_errors();
             auto outcome = s.native->execute(c.to_text(0));
             if (!outcome.ok) {
                 fail_from(*s.native);
             }
             Var affected;
             affected.set_long(outcome.records_affected);
             c.write_back(1, affected);
             IDispatch* set =
                 make_recordset(std::move(outcome.recordset),
                                Ref(static_cast<IDispatch*>(&s)), s.native, 1);
             c.put_object(set);
             set->Release();
         },
         nullptr},
        {L"BeginTrans", 10,
         [](ConnectionObj& s, Call& c) {
             s.native->clear_errors();
             const int level = s.native->begin_trans();
             if (level == 0) {
                 fail_from(*s.native);
             }
             c.put_long(level);
         },
         nullptr},
        {L"CommitTrans", 11,
         [](ConnectionObj& s, Call&) {
             s.native->clear_errors();
             if (!s.native->commit_trans()) {
                 fail_from(*s.native);
             }
         },
         nullptr},
        {L"RollbackTrans", 12,
         [](ConnectionObj& s, Call&) {
             s.native->clear_errors();
             if (!s.native->rollback_trans()) {
                 fail_from(*s.native);
             }
         },
         nullptr},
    };
    return table;
}

// The native connection behind an ActiveConnection argument: a Connection
// object, or a connection string that is opened on the spot.
std::shared_ptr<Connection> native_of(IDispatch* object, const VARIANT* text,
                                      Ref& keep) {
    if (auto* mine = unwrap<ConnectionObj>(object, kPrivateConnection)) {
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

// ----- Recordset, Fields, Field --------------------------------------------

class RecordsetObj : public DispatchObject<RecordsetObj> {
public:
    std::shared_ptr<Recordset> set = Recordset::create();
    Ref connection;
    std::shared_ptr<Connection> native;
    long cursor_type{0};
    long lock_type{1};
    long cursor_location{2};
    long page_size{10};

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid,
                                             void** object) override {
        if (iid == kPrivateRecordset) {
            *object = this;
            AddRef();
            return S_OK;
        }
        return DispatchObject::QueryInterface(iid, object);
    }

    void need_open() const {
        if (set->state() == adStateClosed) {
            throw ComError{3704, kClosedText};
        }
    }
    void need_record() const {
        need_open();
        if (set->bof() || set->eof()) {
            throw ComError{3021, kNoRecordText};
        }
    }
    void need_editable() const {
        need_open();
        if (!set->editable()) {
            throw ComError{3251, kReadOnlyText};
        }
    }
    // A write to the table that failed while a call flushed a pending row.
    void raise_pending() {
        Error error;
        if (set->take_error(error)) {
            throw ComError{error.number, error.description};
        }
    }
    static const std::vector<MemberEntry<RecordsetObj>>& members();
};

class FieldsObj : public DispatchObject<FieldsObj> {
public:
    explicit FieldsObj(RecordsetObj* owner) : owner(owner) { owner->AddRef(); }
    ~FieldsObj() override { owner->Release(); }
    RecordsetObj* owner;
    static const std::vector<MemberEntry<FieldsObj>>& members();
};

class FieldObj : public DispatchObject<FieldObj> {
public:
    FieldObj(RecordsetObj* owner, std::size_t index)
        : owner(owner), index(index) {
        owner->AddRef();
    }
    ~FieldObj() override { owner->Release(); }
    RecordsetObj* owner;
    std::size_t index;
    static const std::vector<MemberEntry<FieldObj>>& members();
};

IDispatch* make_field(RecordsetObj* owner, std::size_t index) {
    return new FieldObj(owner, index);
}

std::size_t field_index(RecordsetObj& owner, const VARIANT& key) {
    const Value value = to_value(key);
    if (const auto* name = std::get_if<std::string>(&value)) {
        const int index = owner.set->field_index(*name);
        if (index < 0) {
            throw ComError{3265, kNotFoundText};
        }
        return static_cast<std::size_t>(index);
    }
    const auto* number = std::get_if<std::int64_t>(&value);
    if (number == nullptr || *number < 0 ||
        static_cast<std::size_t>(*number) >= owner.set->field_count()) {
        throw ComError{3265, kNotFoundText};
    }
    return static_cast<std::size_t>(*number);
}

const std::vector<MemberEntry<FieldObj>>& FieldObj::members() {
    static const std::vector<MemberEntry<FieldObj>> table = {
        {L"Name", 1,
         [](FieldObj& s, Call& c) {
             c.put_text(s.owner->set->field(s.index).name);
         },
         nullptr},
        {L"Type", 2,
         [](FieldObj& s, Call& c) {
             c.put_long(s.owner->set->field(s.index).type);
         },
         nullptr},
        {L"DefinedSize", 3,
         [](FieldObj& s, Call& c) {
             c.put_long(s.owner->set->field(s.index).defined_size);
         },
         nullptr},
        {L"ActualSize", 4,
         [](FieldObj& s, Call& c) {
             s.owner->need_record();
             const Value& v = s.owner->set->value(s.index);
             long size = 0;
             if (const auto* text = std::get_if<std::string>(&v)) {
                 size = static_cast<long>(text->size());
             } else if (const auto* blob = std::get_if<Blob>(&v)) {
                 size = static_cast<long>(blob->size());
             } else if (v.index() != 0) {
                 size = 8;
             }
             c.put_long(size);
         },
         nullptr},
        {L"Value", DISPID_VALUE,
         [](FieldObj& s, Call& c) {
             s.owner->need_record();
             Var out;
             from_value(s.owner->set->value(s.index),
                        s.owner->set->field(s.index).type, out);
             c.put_var(out);
         },
         [](FieldObj& s, Call& c) {
             s.owner->need_editable();
             if (s.owner->set->bof() || s.owner->set->eof()) {
                 throw ComError{3021, kNoRecordText};
             }
             if (!s.owner->set->set_value(
                     s.index, coerced(c.args.back(),
                                      s.owner->set->field(s.index).type))) {
                 throw ComError{3001, kBadArgumentText};
             }
             s.owner->raise_pending();
         }},
    };
    return table;
}

const std::vector<MemberEntry<FieldsObj>>& FieldsObj::members() {
    static const std::vector<MemberEntry<FieldsObj>> table = {
        {L"Count", 1,
         [](FieldsObj& s, Call& c) {
             c.put_long(static_cast<long>(s.owner->set->field_count()));
         },
         nullptr},
        {L"Item", DISPID_VALUE,
         [](FieldsObj& s, Call& c) {
             if (c.count() < 1) {
                 throw ComError{450, "Wrong number of arguments"};
             }
             IDispatch* field =
                 make_field(s.owner, field_index(*s.owner, c.args[0]));
             c.put_object(field);
             field->Release();
         },
         nullptr},
        {L"Append", 2,
         [](FieldsObj& s, Call& c) {
             if (s.owner->set->state() != adStateClosed ||
                 !s.owner->set->append_field(c.to_text(0), c.to_long(1, 0),
                                             c.to_long(2, 0))) {
                 throw ComError{3705, kOpenText};
             }
         },
         nullptr},
        {L"Refresh", 3, [](FieldsObj&, Call&) {}, nullptr},
        {L"_NewEnum", DISPID_NEWENUM,
         [](FieldsObj& s, Call& c) {
             std::vector<Var> list;
             for (std::size_t i = 0; i < s.owner->set->field_count(); ++i) {
                 IDispatch* field = make_field(s.owner, i);
                 Var v;
                 v.set_object(field);
                 field->Release();
                 list.push_back(v);
             }
             auto* enumerator = new Enumerator(std::move(list));
             if (c.result != nullptr) {
                 c.result->vt = VT_UNKNOWN;
                 c.result->punkVal = enumerator;
             } else {
                 enumerator->Release();
             }
         },
         nullptr},
    };
    return table;
}

IDispatch* make_recordset(std::shared_ptr<Recordset> set, const Ref& connection,
                          std::shared_ptr<Connection> native,
                          const long lock_type) {
    auto* object = new RecordsetObj();
    if (set != nullptr) {
        object->set = std::move(set);
    }
    object->connection = connection;
    object->native = std::move(native);
    object->lock_type = lock_type;
    return object;
}

std::string trimmed(std::string text) {
    const auto first = text.find_first_not_of(" \t\r\n");
    if (first == std::string::npos) {
        return {};
    }
    const auto last = text.find_last_not_of(" \t\r\n");
    return text.substr(first, last - first + 1);
}

const std::vector<MemberEntry<RecordsetObj>>& RecordsetObj::members() {
    static const std::vector<MemberEntry<RecordsetObj>> table = {
        {L"ActiveConnection", 1,
         [](RecordsetObj& s, Call& c) {
             c.put_object(static_cast<IDispatch*>(unwrap<ConnectionObj>(
                 s.connection.get(), kPrivateConnection)));
         },
         [](RecordsetObj& s, Call& c) {
             Ref keep;
             auto native =
                 native_of(c.object(c.count() - 1),
                           c.count() > 0 && c.object(c.count() - 1) == nullptr
                               ? &c.args.back()
                               : nullptr,
                           keep);
             s.native = std::move(native);
             s.connection = keep;
         }},
        {L"Open", 2,
         [](RecordsetObj& s, Call& c) {
             if (s.set->state() != adStateClosed) {
                 throw ComError{3705, kOpenText};
             }
             s.cursor_type = c.to_long(2, s.cursor_type);
             s.lock_type = c.to_long(3, s.lock_type);
             if (!c.has(0)) {
                 if (!s.set->open_empty()) {
                     throw ComError{3705, kOpenText};
                 }
                 return;
             }
             if (c.has(1)) {
                 Ref keep;
                 auto native = native_of(
                     c.object(1), c.object(1) == nullptr ? &c.args[1] : nullptr,
                     keep);
                 if (native != nullptr) {
                     s.native = std::move(native);
                     s.connection = keep;
                 }
             }
             if (s.native == nullptr) {
                 throw ComError{3709, kNoConnectionText};
             }
             std::string sql = trimmed(c.to_text(0));
             if (c.to_long(4, -1) == 2 || sql.find(' ') == std::string::npos) {
                 sql = "select * from " + sql;
             }
             s.native->clear_errors();
             auto outcome = s.lock_type > 1 ? s.native->execute_updatable(sql)
                                            : s.native->execute(sql);
             if (!outcome.ok) {
                 fail_from(*s.native);
             }
             s.set = outcome.recordset != nullptr ? std::move(outcome.recordset)
                                                  : Recordset::create();
         },
         nullptr},
        {L"Close", 3,
         [](RecordsetObj& s, Call&) {
             s.need_open();
             s.set->close();
             s.raise_pending();
         },
         nullptr},
        {L"Fields", 4,
         [](RecordsetObj& s, Call& c) {
             auto* fields = new FieldsObj(&s);
             c.put_object(fields);
             fields->Release();
         },
         nullptr, true},
        {L"State", 5,
         [](RecordsetObj& s, Call& c) { c.put_long(s.set->state()); }, nullptr},
        {L"BOF", 6,
         [](RecordsetObj& s, Call& c) {
             s.need_open();
             c.put_bool(s.set->bof());
         },
         nullptr},
        {L"EOF", 7,
         [](RecordsetObj& s, Call& c) {
             s.need_open();
             c.put_bool(s.set->eof());
         },
         nullptr},
        {L"RecordCount", 8,
         [](RecordsetObj& s, Call& c) {
             s.need_open();
             c.put_long(s.set->record_count());
         },
         nullptr},
        {L"AbsolutePosition", 9,
         [](RecordsetObj& s, Call& c) {
             s.need_open();
             c.put_long(s.set->absolute_position());
         },
         [](RecordsetObj& s, Call& c) {
             s.need_open();
             const long position = c.to_long(c.count() - 1, 1);
             s.set->move_first();
             for (long i = 2; i <= position; ++i) {
                 s.set->move_next();
             }
             s.raise_pending();
         }},
        {L"Bookmark", 10,
         [](RecordsetObj& s, Call& c) {
             s.need_open();
             c.put_long(s.set->bookmark());
         },
         [](RecordsetObj& s, Call& c) {
             s.need_open();
             if (!s.set->goto_bookmark(c.to_long(c.count() - 1, 0))) {
                 throw ComError{3001, kBadArgumentText};
             }
             s.raise_pending();
         }},
        {L"Filter", 11,
         [](RecordsetObj& s, Call& c) { c.put_text(s.set->filter()); },
         [](RecordsetObj& s, Call& c) {
             s.need_open();
             if (!s.set->set_filter(c.to_text(c.count() - 1))) {
                 throw ComError{3001, kBadArgumentText};
             }
             s.raise_pending();
         }},
        {L"Sort", 12,
         [](RecordsetObj& s, Call& c) { c.put_text(s.set->sort()); },
         [](RecordsetObj& s, Call& c) {
             s.need_open();
             if (!s.set->set_sort(c.to_text(c.count() - 1))) {
                 throw ComError{3265, kNotFoundText};
             }
             s.raise_pending();
         }},
        {L"CursorType", 13,
         [](RecordsetObj& s, Call& c) { c.put_long(s.cursor_type); },
         [](RecordsetObj& s, Call& c) {
             s.cursor_type = c.to_long(c.count() - 1, 0);
         }},
        {L"LockType", 14,
         [](RecordsetObj& s, Call& c) { c.put_long(s.lock_type); },
         [](RecordsetObj& s, Call& c) {
             s.lock_type = c.to_long(c.count() - 1, 1);
         }},
        {L"CursorLocation", 15,
         [](RecordsetObj& s, Call& c) { c.put_long(s.cursor_location); },
         [](RecordsetObj& s, Call& c) {
             s.cursor_location = c.to_long(c.count() - 1, 2);
         }},
        {L"PageSize", 16,
         [](RecordsetObj& s, Call& c) { c.put_long(s.page_size); },
         [](RecordsetObj& s, Call& c) {
             s.page_size = c.to_long(c.count() - 1, 10);
         }},
        {L"MoveFirst", 17,
         [](RecordsetObj& s, Call&) {
             s.need_open();
             s.set->move_first();
             s.raise_pending();
         },
         nullptr},
        {L"MoveLast", 18,
         [](RecordsetObj& s, Call&) {
             s.need_open();
             s.set->move_last();
             s.raise_pending();
         },
         nullptr},
        {L"MoveNext", 19,
         [](RecordsetObj& s, Call&) {
             s.need_open();
             const bool moved = s.set->move_next();
             s.raise_pending();
             if (!moved) {
                 throw ComError{3021, kNoRecordText};
             }
         },
         nullptr},
        {L"MovePrevious", 20,
         [](RecordsetObj& s, Call&) {
             s.need_open();
             const bool moved = s.set->move_previous();
             s.raise_pending();
             if (!moved) {
                 throw ComError{3021, kNoRecordText};
             }
         },
         nullptr},
        {L"Move", 21,
         [](RecordsetObj& s, Call& c) {
             s.need_open();
             const long steps = c.to_long(0, 0);
             const long start = c.to_long(1, 0);
             if (start == 1) {
                 s.set->move_first();
             } else if (start == 2) {
                 s.set->move_last();
             }
             for (long i = 0; i < (steps < 0 ? -steps : steps); ++i) {
                 if (steps > 0) {
                     s.set->move_next();
                 } else {
                     s.set->move_previous();
                 }
             }
             s.raise_pending();
         },
         nullptr},
        {L"AddNew", 22,
         [](RecordsetObj& s, Call&) {
             s.need_editable();
             if (!s.set->add_new()) {
                 throw ComError{3001, kBadArgumentText};
             }
             s.raise_pending();
         },
         nullptr},
        {L"Update", 23,
         [](RecordsetObj& s, Call&) {
             s.need_editable();
             static_cast<void>(s.set->update());
             s.raise_pending();
         },
         nullptr},
        {L"CancelUpdate", 24,
         [](RecordsetObj& s, Call&) {
             s.need_editable();
             static_cast<void>(s.set->cancel_update());
         },
         nullptr},
        {L"Delete", 25,
         [](RecordsetObj& s, Call&) {
             s.need_editable();
             if (s.set->bof() || s.set->eof()) {
                 throw ComError{3021, kNoRecordText};
             }
             if (!s.set->delete_current()) {
                 s.raise_pending();
                 throw ComError{3001, kBadArgumentText};
             }
         },
         nullptr},
        {L"Find", 26,
         [](RecordsetObj& s, Call& c) {
             s.need_open();
             if (!s.set->find(c.to_text(0), c.to_long(1, 0),
                              c.to_long(2, 1) != -1)) {
                 throw ComError{3001, kBadArgumentText};
             }
             s.raise_pending();
         },
         nullptr},
        {L"Supports", 27, [](RecordsetObj&, Call& c) { c.put_bool(true); },
         nullptr},
        {L"GetRows", 28,
         [](RecordsetObj& s, Call& c) {
             s.need_record();
             const long fields = static_cast<long>(s.set->field_count());
             long rows = s.set->record_count() - s.set->absolute_position() + 1;
             const long asked = c.to_long(0, -1);
             if (asked >= 0 && asked < rows) {
                 rows = asked;
             }
             SAFEARRAYBOUND bounds[2] = {{static_cast<ULONG>(fields), 0},
                                         {static_cast<ULONG>(rows), 0}};
             SAFEARRAY* array = SafeArrayCreate(VT_VARIANT, 2, bounds);
             if (array == nullptr) {
                 throw ComError{7, "Out of memory"};
             }
             for (long row = 0; row < rows; ++row) {
                 for (long field = 0; field < fields; ++field) {
                     Var cell;
                     from_value(
                         s.set->value(static_cast<std::size_t>(field)),
                         s.set->field(static_cast<std::size_t>(field)).type,
                         cell);
                     long index[2] = {field, row};
                     SafeArrayPutElement(array, index, cell.get());
                 }
                 s.set->move_next();
             }
             if (c.result != nullptr) {
                 c.result->vt = VT_ARRAY | VT_VARIANT;
                 c.result->parray = array;
             } else {
                 SafeArrayDestroy(array);
             }
         },
         nullptr},
        {L"GetString", 29,
         [](RecordsetObj& s, Call& c) {
             s.need_open();
             const long limit = c.to_long(1, -1);
             const std::string column = c.has(2) ? c.to_text(2) : "\t";
             const std::string line = c.has(3) ? c.to_text(3) : "\r";
             const std::string empty = c.has(4) ? c.to_text(4) : "";
             std::string text;
             long done = 0;
             while (!s.set->bof() && !s.set->eof() &&
                    (limit < 0 || done < limit)) {
                 for (std::size_t i = 0; i < s.set->field_count(); ++i) {
                     const Value& v = s.set->value(i);
                     if (v.index() == 0) {
                         text += empty;
                     } else if (const auto* t = std::get_if<std::string>(&v)) {
                         text += *t;
                     } else if (const auto* n = std::get_if<std::int64_t>(&v)) {
                         text += std::to_string(*n);
                     } else if (const auto* d = std::get_if<double>(&v)) {
                         char buffer[40];
                         std::snprintf(buffer, sizeof(buffer), "%.15g", *d);
                         text += buffer;
                     }
                     if (i + 1U < s.set->field_count()) {
                         text += column;
                     }
                 }
                 text += line;
                 ++done;
                 s.set->move_next();
             }
             c.put_text(text);
         },
         nullptr},
        // rs("name"): the value of a field, as in ADO.
        {L"Collect", DISPID_VALUE,
         [](RecordsetObj& s, Call& c) {
             s.need_record();
             if (c.count() < 1) {
                 throw ComError{450, "Wrong number of arguments"};
             }
             const std::size_t i = field_index(s, c.args[0]);
             Var out;
             from_value(s.set->value(i), s.set->field(i).type, out);
             c.put_var(out);
         },
         [](RecordsetObj& s, Call& c) {
             s.need_editable();
             if (c.count() < 2) {
                 throw ComError{450, "Wrong number of arguments"};
             }
             if (s.set->bof() || s.set->eof()) {
                 throw ComError{3021, kNoRecordText};
             }
             const std::size_t i = field_index(s, c.args[0]);
             if (!s.set->set_value(
                     i, coerced(c.args.back(), s.set->field(i).type))) {
                 throw ComError{3001, kBadArgumentText};
             }
             s.raise_pending();
         }},
    };
    return table;
}

// ----- Command --------------------------------------------------------------

class CommandObj : public DispatchObject<CommandObj> {
public:
    Ref connection;
    std::shared_ptr<Connection> native;
    std::string text;
    long type{8};
    bool prepared{false};
    long timeout{30};
    ParametersObj* parameters{nullptr};

    ~CommandObj() override {
        if (parameters != nullptr) {
            parameters->Release();
        }
    }
    ParametersObj* params() {
        if (parameters == nullptr) {
            parameters = new ParametersObj();
        }
        return parameters;
    }
    static const std::vector<MemberEntry<CommandObj>>& members();
};

const std::vector<MemberEntry<CommandObj>>& CommandObj::members() {
    static const std::vector<MemberEntry<CommandObj>> table = {
        {L"ActiveConnection", 1,
         [](CommandObj& s, Call& c) {
             IDispatch* object =
                 s.connection.empty()
                     ? nullptr
                     : static_cast<IDispatch*>(unwrap<ConnectionObj>(
                           s.connection.get(), kPrivateConnection));
             c.put_object(object);
         },
         [](CommandObj& s, Call& c) {
             Ref keep;
             s.native = native_of(
                 c.object(c.count() - 1),
                 c.object(c.count() - 1) == nullptr ? &c.args.back() : nullptr,
                 keep);
             s.connection = keep;
         }},
        {L"CommandText", 2, [](CommandObj& s, Call& c) { c.put_text(s.text); },
         [](CommandObj& s, Call& c) { s.text = c.to_text(c.count() - 1); }},
        {L"CommandType", 3, [](CommandObj& s, Call& c) { c.put_long(s.type); },
         [](CommandObj& s, Call& c) { s.type = c.to_long(c.count() - 1, 8); }},
        {L"Prepared", 4, [](CommandObj& s, Call& c) { c.put_bool(s.prepared); },
         [](CommandObj& s, Call& c) {
             s.prepared = c.to_long(c.count() - 1, 0) != 0;
         }},
        {L"CommandTimeout", 5,
         [](CommandObj& s, Call& c) { c.put_long(s.timeout); },
         [](CommandObj& s, Call& c) {
             s.timeout = c.to_long(c.count() - 1, 30);
         }},
        {L"Parameters", 6,
         [](CommandObj& s, Call& c) { c.put_object(s.params()); }, nullptr,
         true},
        {L"CreateParameter", 7,
         [](CommandObj&, Call& c) {
             auto* parameter = new ParameterObj();
             parameter->name = c.to_text(0);
             parameter->type = c.to_long(1, 0);
             parameter->direction = c.to_long(2, 1);
             parameter->size = c.to_long(3, 0);
             if (c.has(4)) {
                 parameter->value = dereferenced(c.args[4]);
             }
             c.put_object(parameter);
             parameter->Release();
         },
         nullptr},
        {L"Execute", 8,
         [](CommandObj& s, Call& c) {
             if (s.native == nullptr) {
                 throw ComError{3709, kNoConnectionText};
             }
             std::vector<Value> values;
             if (c.has(1)) {
                 const Var plain = dereferenced(c.args[1]);
                 if ((plain.ref().vt & VT_ARRAY) != 0 &&
                     plain.ref().parray != nullptr &&
                     (plain.ref().vt & VT_TYPEMASK) == VT_VARIANT) {
                     long low = 0;
                     long high = -1;
                     SafeArrayGetLBound(plain.ref().parray, 1, &low);
                     SafeArrayGetUBound(plain.ref().parray, 1, &high);
                     for (long i = low; i <= high; ++i) {
                         VARIANT element{};
                         VariantInit(&element);
                         SafeArrayGetElement(plain.ref().parray, &i, &element);
                         values.push_back(to_value(element));
                         VariantClear(&element);
                     }
                 }
             } else if (s.parameters != nullptr) {
                 values = s.parameters->values();
             }
             s.native->clear_errors();
             auto outcome = s.native->execute(s.text, values);
             if (!outcome.ok) {
                 fail_from(*s.native);
             }
             Var affected;
             affected.set_long(outcome.records_affected);
             c.write_back(0, affected);
             IDispatch* set = make_recordset(std::move(outcome.recordset),
                                             s.connection, s.native, 1);
             c.put_object(set);
             set->Release();
         },
         nullptr},
    };
    return table;
}

// ----- class factories ------------------------------------------------------

using Maker = IDispatch* (*)();

class ClassFactory : public IClassFactory {
public:
    explicit ClassFactory(Maker maker) : maker_(maker) { ++g_objects; }
    ~ClassFactory() { --g_objects; }
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid,
                                             void** object) override {
        if (object == nullptr) {
            return E_POINTER;
        }
        if (iid == IID_IUnknown || iid == IID_IClassFactory) {
            *object = static_cast<IClassFactory*>(this);
            AddRef();
            return S_OK;
        }
        *object = nullptr;
        return E_NOINTERFACE;
    }
    ULONG STDMETHODCALLTYPE AddRef() override { return ++references_; }
    ULONG STDMETHODCALLTYPE Release() override {
        const long left = --references_;
        if (left == 0) {
            delete this;
        }
        return static_cast<ULONG>(left);
    }
    HRESULT STDMETHODCALLTYPE CreateInstance(IUnknown* outer, REFIID iid,
                                             void** object) override {
        if (object == nullptr) {
            return E_POINTER;
        }
        *object = nullptr;
        if (outer != nullptr) {
            return CLASS_E_NOAGGREGATION;
        }
        IDispatch* made = maker_();
        if (made == nullptr) {
            return E_OUTOFMEMORY;
        }
        const HRESULT result = made->QueryInterface(iid, object);
        made->Release();
        return result;
    }
    HRESULT STDMETHODCALLTYPE LockServer(BOOL lock) override {
        if (lock != FALSE) {
            ++g_locks;
        } else {
            --g_locks;
        }
        return S_OK;
    }

private:
    Maker maker_;
    std::atomic<long> references_{1};
};

IDispatch* new_connection() {
    return new (std::nothrow) ConnectionObj();
}
IDispatch* new_recordset() {
    return new (std::nothrow) RecordsetObj();
}
IDispatch* new_command() {
    return new (std::nothrow) CommandObj();
}

}  // namespace

HRESULT make_class_factory(const std::size_t index, REFIID iid, void** object) {
    static const Maker makers[] = {new_connection, new_recordset, new_command};
    if (object == nullptr) {
        return E_POINTER;
    }
    *object = nullptr;
    if (index >= std::size(makers)) {
        return CLASS_E_CLASSNOTAVAILABLE;
    }
    auto* factory = new (std::nothrow) ClassFactory(makers[index]);
    if (factory == nullptr) {
        return E_OUTOFMEMORY;
    }
    const HRESULT result = factory->QueryInterface(iid, object);
    factory->Release();
    return result;
}

bool can_unload() noexcept {
    return g_objects.load() == 0 && g_locks.load() == 0;
}

}  // namespace wfc::adodb::com
