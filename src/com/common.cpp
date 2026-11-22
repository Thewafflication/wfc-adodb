// Shared parts of the ADODB COM objects: the type library, Property and
// Properties, Error and Errors, and the class factories.

#include "objects.hpp"

#include <new>

namespace wfc::adodb::com {

namespace {

HMODULE this_module() {
    HMODULE module = nullptr;
    GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS |
                           GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
                       reinterpret_cast<LPCWSTR>(&this_module), &module);
    return module;
}

ITypeLib* load_type_library() {
    wchar_t path[MAX_PATH * 2]{};
    const DWORD length = GetModuleFileNameW(this_module(), path, MAX_PATH * 2U);
    if (length == 0) {
        return nullptr;
    }
    ITypeLib* library = nullptr;
    if (FAILED(LoadTypeLibEx(path, REGKIND_NONE, &library))) {
        return nullptr;
    }
    return library;
}

}  // namespace

ITypeLib* type_library() {
    static ITypeLib* const library = load_type_library();
    return library;
}

void fail_from(const Connection& connection) {
    if (!connection.errors().empty()) {
        const auto& last = connection.errors().back();
        throw ComError{last.number, last.description};
    }
    throw ComError{5, "Invalid procedure call or argument"};
}

// ----- Property and Properties -------------------------------------------

const Table<PropertyObj>& PropertyObj::members() {
    static const Table<PropertyObj> table = {
        {L"Value", DISPID_VALUE, read_var(&PropertyObj::value),
         write_var(&PropertyObj::value)},
        {L"Name", 1610743810, read_text(&PropertyObj::name), nullptr},
        {L"Type", 1610743811, read_long(&PropertyObj::type), nullptr},
        {L"Attributes", 1610743812, read_long(&PropertyObj::attributes),
         write_long(&PropertyObj::attributes)},
    };
    return table;
}

PropertiesObj::~PropertiesObj() {
    for (auto* item : items) {
        item->Release();
    }
}

void PropertiesObj::add(const std::string& name, const Var& value,
                        const long type, const long attributes) {
    auto* item = new PropertyObj();
    item->name = name;
    item->value = value;
    item->type = type;
    item->attributes = attributes;
    items.push_back(item);
}

IDispatch* make_properties(PropertiesObj*& slot) {
    if (slot == nullptr) {
        slot = new PropertiesObj();
    }
    return slot;
}

namespace {

// The index of an Item argument in a collection of named things.
template <class Items, class NameOf>
std::size_t locate_item(const Items& items, const VARIANT& index,
                        NameOf name_of) {
    const Value key = to_value(index);
    if (const auto* name = std::get_if<std::string>(&key)) {
        for (std::size_t i = 0; i < items.size(); ++i) {
            if (_stricmp(name_of(items[i]).c_str(), name->c_str()) == 0) {
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

}  // namespace

const Table<PropertiesObj>& PropertiesObj::members() {
    static const Table<PropertiesObj> table = {
        {L"Count", 1,
         [](PropertiesObj& s, Call& c) {
             c.put_long(static_cast<long>(s.items.size()));
         },
         nullptr},
        {L"Item", DISPID_VALUE,
         [](PropertiesObj& s, Call& c) {
             if (c.count() < 1) {
                 throw ComError{450, "Wrong number of arguments"};
             }
             const std::size_t at =
                 locate_item(s.items, c.args[0],
                             [](const PropertyObj* p) { return p->name; });
             c.put_object(s.items[at]);
         },
         nullptr},
        {L"Refresh", 2, [](PropertiesObj&, Call&) {}, nullptr},
        {L"_NewEnum", DISPID_NEWENUM,
         [](PropertiesObj& s, Call& c) {
             std::vector<Var> list;
             for (auto* item : s.items) {
                 Var v;
                 v.set_object(item);
                 list.push_back(v);
             }
             put_enumerator(c, std::move(list));
         },
         nullptr},
    };
    return table;
}

// ----- Error and Errors ------------------------------------------------------

const Table<ErrorObj>& ErrorObj::members() {
    static const Table<ErrorObj> table = {
        {L"Number", 1, [](ErrorObj& s, Call& c) { c.put_long(s.error.number); },
         nullptr},
        {L"Source", 2, [](ErrorObj& s, Call& c) { c.put_text(s.error.source); },
         nullptr},
        {L"Description", DISPID_VALUE,
         [](ErrorObj& s, Call& c) { c.put_text(s.error.description); },
         nullptr},
        {L"HelpFile", 3, [](ErrorObj&, Call& c) { c.put_text(""); }, nullptr},
        {L"HelpContext", 4, [](ErrorObj&, Call& c) { c.put_long(0); }, nullptr},
        {L"SQLState", 5,
         [](ErrorObj& s, Call& c) { c.put_text(s.error.sql_state); }, nullptr},
        {L"NativeError", 6,
         [](ErrorObj& s, Call& c) { c.put_long(s.error.number); }, nullptr},
    };
    return table;
}

const Table<ErrorsObj>& ErrorsObj::members() {
    static const Table<ErrorsObj> table = {
        {L"Count", 1,
         [](ErrorsObj& s, Call& c) {
             c.put_long(static_cast<long>(s.native->errors().size()));
         },
         nullptr},
        {L"Item", DISPID_VALUE,
         [](ErrorsObj& s, Call& c) {
             if (c.count() < 1) {
                 throw ComError{450, "Wrong number of arguments"};
             }
             const auto& all = s.native->errors();
             const Value key = to_value(c.args[0]);
             const auto* number = std::get_if<std::int64_t>(&key);
             if (number == nullptr || *number < 0 ||
                 static_cast<std::size_t>(*number) >= all.size()) {
                 throw ComError{3265, kNotFoundText};
             }
             auto* error = new ErrorObj();
             error->error = all[static_cast<std::size_t>(*number)];
             c.put_object(error);
             error->Release();
         },
         nullptr},
        {L"Clear", 1610809345,
         [](ErrorsObj& s, Call&) { s.native->clear_errors(); }, nullptr},
        {L"Refresh", 2, [](ErrorsObj&, Call&) {}, nullptr},
        {L"_NewEnum", DISPID_NEWENUM,
         [](ErrorsObj& s, Call& c) {
             std::vector<Var> list;
             for (const auto& entry : s.native->errors()) {
                 auto* error = new ErrorObj();
                 error->error = entry;
                 Var v;
                 v.set_object(error);
                 error->Release();
                 list.push_back(v);
             }
             put_enumerator(c, std::move(list));
         },
         nullptr},
    };
    return table;
}

// ----- class factories ------------------------------------------------------

namespace {

using Maker = IUnknown* (*)();

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
        IUnknown* made = maker_();
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

IUnknown* new_connection() {
    return static_cast<IDispatch*>(new (std::nothrow) ConnectionObj());
}
IUnknown* new_recordset() {
    return static_cast<IDispatch*>(new (std::nothrow) RecordsetObj());
}
IUnknown* new_command() {
    return static_cast<IDispatch*>(new (std::nothrow) CommandObj());
}
IUnknown* new_parameter() {
    return static_cast<IDispatch*>(new (std::nothrow) ParameterObj());
}
IUnknown* new_record() {
    return static_cast<IDispatch*>(new (std::nothrow) RecordObj());
}
IUnknown* new_stream() {
    return static_cast<IDispatch*>(new (std::nothrow) StreamObj());
}

}  // namespace

HRESULT make_class_factory(const std::size_t index, REFIID iid, void** object) {
    static const Maker makers[] = {new_connection, new_recordset, new_command,
                                   new_parameter,  new_record,    new_stream};
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
