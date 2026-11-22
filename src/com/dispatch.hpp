// A small late-bound COM layer for the ADODB classes: IDispatch objects whose
// members come from a table, VARIANT conversion to and from the library's
// values, and the error a Visual Basic client sees. Windows only; internal to
// the msado15.dll build.

#ifndef WFC_ADODB_COM_DISPATCH_HPP
#define WFC_ADODB_COM_DISPATCH_HPP

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>

#include <objbase.h>
#include <oleauto.h>

#include <wfc/adodb/adodb.hpp>

#include <algorithm>
#include <atomic>
#include <cmath>
#include <cstdint>
#include <cstring>
#include <limits>
#include <memory>
#include <new>
#include <string>
#include <utility>
#include <vector>

namespace wfc::adodb::com {

inline std::atomic<long> g_objects{0};
inline std::atomic<long> g_locks{0};

// A refused call: `number` is the Visual Basic or ADO error number.
struct ComError {
    long number;
    std::string text;
};

// ----- strings and values ---------------------------------------------------

[[nodiscard]] inline std::string to_utf8(const wchar_t* text, int length) {
    if (text == nullptr || length <= 0) {
        return {};
    }
    const int size = WideCharToMultiByte(CP_UTF8, 0, text, length, nullptr, 0,
                                         nullptr, nullptr);
    std::string out(static_cast<std::size_t>(size), '\0');
    WideCharToMultiByte(CP_UTF8, 0, text, length, out.data(), size, nullptr,
                        nullptr);
    return out;
}

[[nodiscard]] inline std::wstring to_wide(const std::string& text) {
    if (text.empty()) {
        return {};
    }
    const int size = MultiByteToWideChar(
        CP_UTF8, 0, text.data(), static_cast<int>(text.size()), nullptr, 0);
    std::wstring out(static_cast<std::size_t>(size), L'\0');
    MultiByteToWideChar(CP_UTF8, 0, text.data(), static_cast<int>(text.size()),
                        out.data(), size);
    return out;
}

// A VARIANT that clears itself.
class Var {
public:
    Var() { VariantInit(&value_); }
    Var(const Var& other) {
        VariantInit(&value_);
        VariantCopy(&value_, const_cast<VARIANT*>(&other.value_));
    }
    Var& operator=(const Var& other) {
        if (this != &other) {
            VariantClear(&value_);
            VariantCopy(&value_, const_cast<VARIANT*>(&other.value_));
        }
        return *this;
    }
    ~Var() { VariantClear(&value_); }
    [[nodiscard]] VARIANT* get() noexcept { return &value_; }
    [[nodiscard]] const VARIANT& ref() const noexcept { return value_; }

    void set_null() {
        VariantClear(&value_);
        value_.vt = VT_NULL;
    }
    void set_long(const long number) {
        VariantClear(&value_);
        value_.vt = VT_I4;
        value_.lVal = number;
    }
    void set_bool(const bool flag) {
        VariantClear(&value_);
        value_.vt = VT_BOOL;
        value_.boolVal = flag ? VARIANT_TRUE : VARIANT_FALSE;
    }
    void set_double(const double number) {
        VariantClear(&value_);
        value_.vt = VT_R8;
        value_.dblVal = number;
    }
    void set_date(const double serial) {
        VariantClear(&value_);
        value_.vt = VT_DATE;
        value_.date = serial;
    }
    void set_text(const std::string& text) {
        VariantClear(&value_);
        const std::wstring wide = to_wide(text);
        value_.vt = VT_BSTR;
        value_.bstrVal =
            SysAllocStringLen(wide.data(), static_cast<UINT>(wide.size()));
    }
    // Takes a reference to the object.
    void set_object(IDispatch* object) {
        VariantClear(&value_);
        value_.vt = VT_DISPATCH;
        value_.pdispVal = object;
        if (object != nullptr) {
            object->AddRef();
        }
    }

private:
    VARIANT value_;
};

// The VARIANT with a reference or a variant-in-variant removed.
[[nodiscard]] inline Var dereferenced(const VARIANT& variant) {
    Var out;
    if (variant.vt == (VT_BYREF | VT_VARIANT) && variant.pvarVal != nullptr) {
        VARIANT inner{};
        VariantInit(&inner);
        VariantCopyInd(&inner, variant.pvarVal);
        VariantCopy(out.get(), &inner);
        VariantClear(&inner);
    } else {
        VariantCopyInd(out.get(), const_cast<VARIANT*>(&variant));
    }
    return out;
}

[[nodiscard]] inline bool is_missing(const VARIANT& variant) {
    if (variant.vt == VT_ERROR) {
        return true;
    }
    return variant.vt == (VT_BYREF | VT_VARIANT) &&
           variant.pvarVal != nullptr && variant.pvarVal->vt == VT_ERROR;
}

[[nodiscard]] inline bool is_nothing(const VARIANT& variant) {
    const Var plain = dereferenced(variant);
    return plain.ref().vt == VT_EMPTY || plain.ref().vt == VT_NULL ||
           ((plain.ref().vt == VT_DISPATCH || plain.ref().vt == VT_UNKNOWN) &&
            plain.ref().pdispVal == nullptr);
}

[[nodiscard]] inline Value to_value(const VARIANT& variant) {
    const Var plain = dereferenced(variant);
    const VARIANT& v = plain.ref();
    switch (v.vt) {
        case VT_EMPTY:
        case VT_NULL:
        case VT_ERROR:
            return Value{};
        case VT_BOOL:
            return Value{
                static_cast<std::int64_t>(v.boolVal != VARIANT_FALSE ? -1 : 0)};
        case VT_I1:
            return Value{static_cast<std::int64_t>(v.cVal)};
        case VT_UI1:
            return Value{static_cast<std::int64_t>(v.bVal)};
        case VT_I2:
            return Value{static_cast<std::int64_t>(v.iVal)};
        case VT_UI2:
            return Value{static_cast<std::int64_t>(v.uiVal)};
        case VT_I4:
        case VT_INT:
            return Value{static_cast<std::int64_t>(v.lVal)};
        case VT_UI4:
        case VT_UINT:
            return Value{static_cast<std::int64_t>(v.ulVal)};
        case VT_I8:
            return Value{static_cast<std::int64_t>(v.llVal)};
        case VT_R4:
            return Value{static_cast<double>(v.fltVal)};
        case VT_R8:
            return Value{v.dblVal};
        case VT_DATE:
            return Value{v.date};
        case VT_CY:
            return Value{static_cast<double>(v.cyVal.int64) / 10000.0};
        case VT_BSTR:
            return Value{
                to_utf8(v.bstrVal, static_cast<int>(SysStringLen(v.bstrVal)))};
        default:
            break;
    }
    if ((v.vt & VT_ARRAY) != 0 && (v.vt & VT_TYPEMASK) == VT_UI1 &&
        v.parray != nullptr) {
        long low = 0;
        long high = -1;
        SafeArrayGetLBound(v.parray, 1, &low);
        SafeArrayGetUBound(v.parray, 1, &high);
        Blob bytes;
        void* data = nullptr;
        if (high >= low && SUCCEEDED(SafeArrayAccessData(v.parray, &data))) {
            bytes.assign(static_cast<unsigned char*>(data),
                         static_cast<unsigned char*>(data) + (high - low + 1));
            SafeArrayUnaccessData(v.parray);
        }
        return Value{std::move(bytes)};
    }
    VARIANT text{};
    VariantInit(&text);
    if (SUCCEEDED(
            VariantChangeType(&text, const_cast<VARIANT*>(&v), 0, VT_BSTR))) {
        Value result{to_utf8(text.bstrVal,
                             static_cast<int>(SysStringLen(text.bstrVal)))};
        VariantClear(&text);
        return result;
    }
    return Value{};
}

// A library value as the VARIANT its ADO data type names.
inline void from_value(const Value& value, const int type, Var& out) {
    switch (value.index()) {
        case 1: {
            const auto number = std::get<1>(value);
            if (type == 11) {
                out.set_bool(number != 0);
            } else if (type == 7 || type == 133 || type == 134 || type == 135) {
                out.set_date(static_cast<double>(number));
            } else if (number >= std::numeric_limits<long>::min() &&
                       number <= std::numeric_limits<long>::max()) {
                out.set_long(static_cast<long>(number));
            } else {
                out.set_double(static_cast<double>(number));
            }
            return;
        }
        case 2: {
            const double number = std::get<2>(value);
            if (type == 7 || type == 133 || type == 134 || type == 135) {
                out.set_date(number);
            } else if (type == 11) {
                out.set_bool(number != 0.0);
            } else {
                out.set_double(number);
            }
            return;
        }
        case 3:
            out.set_text(std::get<3>(value));
            return;
        case 4: {
            const auto& blob = std::get<4>(value);
            SAFEARRAY* array = SafeArrayCreateVector(
                VT_UI1, 0, static_cast<ULONG>(blob.size()));
            if (array != nullptr && !blob.empty()) {
                void* data = nullptr;
                if (SUCCEEDED(SafeArrayAccessData(array, &data))) {
                    std::memcpy(data, blob.data(), blob.size());
                    SafeArrayUnaccessData(array);
                }
            }
            VariantClear(out.get());
            out.get()->vt = VT_ARRAY | VT_UI1;
            out.get()->parray = array;
            return;
        }
        default:
            out.set_null();
    }
}

// A value assigned to a field, stored in the form its type calls for.
[[nodiscard]] inline Value coerced(const VARIANT& variant, const int type) {
    const Var plain = dereferenced(variant);
    const VARIANT& v = plain.ref();
    if (v.vt == VT_EMPTY || v.vt == VT_NULL || v.vt == VT_ERROR) {
        return Value{};
    }
    Value stored = to_value(v);
    const bool text_type = type == 200 || type == 201 || type == 202 ||
                           type == 203 || type == 8 || type == 129 ||
                           type == 130;
    if (text_type) {
        if (stored.index() == 3) {
            return stored;
        }
        if (v.vt == VT_BOOL) {
            return Value{
                std::string(v.boolVal != VARIANT_FALSE ? "True" : "False")};
        }
        VARIANT text{};
        VariantInit(&text);
        if (SUCCEEDED(VariantChangeType(&text, const_cast<VARIANT*>(&v), 0,
                                        VT_BSTR))) {
            Value result{to_utf8(text.bstrVal,
                                 static_cast<int>(SysStringLen(text.bstrVal)))};
            VariantClear(&text);
            return result;
        }
        return stored;
    }
    double number = 0;
    if (stored.index() == 3) {
        const std::string& text = std::get<3>(stored);
        char* end = nullptr;
        number = std::strtod(text.c_str(), &end);
        if (end == text.c_str() || *end != '\0') {
            return stored;  // not a number: kept as text
        }
    } else if (stored.index() == 1) {
        number = static_cast<double>(std::get<1>(stored));
    } else if (stored.index() == 2) {
        number = std::get<2>(stored);
    } else {
        return stored;
    }
    switch (type) {
        case 2:
        case 3:
        case 16:
        case 17:
        case 20:
        case 11:
            return Value{static_cast<std::int64_t>(
                type == 11 && number != 0.0 ? -1 : std::llround(number))};
        default:
            return Value{number};
    }
}

// ----- the call as a member sees it ----------------------------------------

struct Call {
    // The arguments in the order the program wrote them.
    std::vector<VARIANT> args;
    bool is_put{false};
    VARIANT* result{nullptr};

    [[nodiscard]] std::size_t count() const noexcept { return args.size(); }
    [[nodiscard]] bool has(const std::size_t index) const noexcept {
        return index < args.size() && !is_missing(args[index]);
    }
    [[nodiscard]] long to_long(const std::size_t index,
                               const long fallback) const {
        if (!has(index)) {
            return fallback;
        }
        const Value v = to_value(args[index]);
        if (const auto* i = std::get_if<std::int64_t>(&v)) {
            return static_cast<long>(*i);
        }
        if (const auto* d = std::get_if<double>(&v)) {
            return static_cast<long>(std::llround(*d));
        }
        if (const auto* s = std::get_if<std::string>(&v)) {
            return std::strtol(s->c_str(), nullptr, 10);
        }
        return fallback;
    }
    [[nodiscard]] std::string to_text(const std::size_t index) const {
        if (!has(index)) {
            return {};
        }
        const Value v = to_value(args[index]);
        if (const auto* s = std::get_if<std::string>(&v)) {
            return *s;
        }
        if (const auto* i = std::get_if<std::int64_t>(&v)) {
            return std::to_string(*i);
        }
        if (const auto* d = std::get_if<double>(&v)) {
            char buffer[40];
            std::snprintf(buffer, sizeof(buffer), "%.15g", *d);
            return buffer;
        }
        return {};
    }
    // The object in an argument (an IDispatch), or null.
    [[nodiscard]] IDispatch* object(const std::size_t index) const {
        if (index >= args.size()) {
            return nullptr;
        }
        const VARIANT* v = &args[index];
        if (v->vt == (VT_BYREF | VT_VARIANT) && v->pvarVal != nullptr) {
            v = v->pvarVal;
        }
        if (v->vt == VT_DISPATCH || v->vt == VT_UNKNOWN) {
            return v->pdispVal;
        }
        if (v->vt == (VT_BYREF | VT_DISPATCH) && v->ppdispVal != nullptr) {
            return *v->ppdispVal;
        }
        return nullptr;
    }
    // Writes an output argument that the caller passed by reference.
    void write_back(const std::size_t index, const Var& value) const {
        if (index < args.size() && args[index].vt == (VT_BYREF | VT_VARIANT) &&
            args[index].pvarVal != nullptr) {
            VariantClear(args[index].pvarVal);
            VariantCopy(args[index].pvarVal,
                        const_cast<VARIANT*>(&value.ref()));
        }
    }
    void put_long(const long number) const {
        if (result != nullptr) {
            VariantClear(result);
            result->vt = VT_I4;
            result->lVal = number;
        }
    }
    void put_bool(const bool flag) const {
        if (result != nullptr) {
            VariantClear(result);
            result->vt = VT_BOOL;
            result->boolVal = flag ? VARIANT_TRUE : VARIANT_FALSE;
        }
    }
    void put_text(const std::string& text) const {
        if (result != nullptr) {
            Var v;
            v.set_text(text);
            VariantClear(result);
            VariantCopy(result, v.get());
        }
    }
    void put_var(const Var& value) const {
        if (result != nullptr) {
            VariantClear(result);
            VariantCopy(result, const_cast<VARIANT*>(&value.ref()));
        }
    }
    void put_object(IDispatch* object) const {
        if (result != nullptr) {
            VariantClear(result);
            result->vt = VT_DISPATCH;
            result->pdispVal = object;
            if (object != nullptr) {
                object->AddRef();
            }
        }
    }
};

// ----- the late-bound object ------------------------------------------------

template <class Self>
struct MemberEntry {
    const wchar_t* name;
    DISPID id;
    void (*get)(Self&, Call&);  // a property read or a method
    void (*put)(Self&, Call&);  // a property write, or null
    // rs.Fields("name"): arguments given to a property that takes none go to
    // the default member of the object it returns.
    bool forward{false};
};

// A COM object whose members come from Self::members().
template <class Self>
class DispatchObject : public IDispatch {
public:
    DispatchObject() { ++g_objects; }
    DispatchObject(const DispatchObject&) = delete;
    DispatchObject& operator=(const DispatchObject&) = delete;
    virtual ~DispatchObject() { --g_objects; }

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid,
                                             void** object) override {
        if (object == nullptr) {
            return E_POINTER;
        }
        if (iid == IID_IUnknown || iid == IID_IDispatch) {
            *object = static_cast<IDispatch*>(this);
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

    HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT* count) override {
        if (count != nullptr) {
            *count = 0;
        }
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT, LCID, ITypeInfo**) override {
        return E_NOTIMPL;
    }
    HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID, LPOLESTR* names, UINT count,
                                            LCID, DISPID* ids) override {
        HRESULT result = S_OK;
        for (UINT i = 0; i < count; ++i) {
            ids[i] = DISPID_UNKNOWN;
            if (i > 0) {
                result = DISP_E_UNKNOWNNAME;  // named arguments: not offered
                continue;
            }
            bool found = false;
            for (const auto& entry : Self::members()) {
                if (_wcsicmp(entry.name, names[i]) == 0) {
                    ids[i] = entry.id;
                    found = true;
                    break;
                }
            }
            if (!found) {
                result = DISP_E_UNKNOWNNAME;
            }
        }
        return result;
    }
    HRESULT STDMETHODCALLTYPE Invoke(DISPID id, REFIID, LCID lcid, WORD flags,
                                     DISPPARAMS* params, VARIANT* result,
                                     EXCEPINFO* exception, UINT*) override {
        const MemberEntry<Self>* entry = nullptr;
        for (const auto& candidate : Self::members()) {
            if (candidate.id == id) {
                entry = &candidate;
                break;
            }
        }
        if (entry == nullptr) {
            return DISP_E_MEMBERNOTFOUND;
        }
        const bool wants_put =
            (flags & (DISPATCH_PROPERTYPUT | DISPATCH_PROPERTYPUTREF)) != 0;
        Call call;
        call.is_put = wants_put;
        call.result = result;
        UINT total = params != nullptr ? params->cArgs : 0;
        // The arguments arrive last first; a put also carries the new value
        // as a named argument.
        for (UINT i = total; i > 0; --i) {
            call.args.push_back(params->rgvarg[i - 1]);
        }
        if (wants_put) {
            // The new value arrives first, so it is the last element here.
            if (entry->put == nullptr) {
                return DISP_E_MEMBERNOTFOUND;
            }
        } else if (entry->get == nullptr) {
            return DISP_E_MEMBERNOTFOUND;
        }
        if (result != nullptr) {
            VariantInit(result);
        }
        try {
            (wants_put ? entry->put : entry->get)(static_cast<Self&>(*this),
                                                  call);
        } catch (const ComError& error) {
            return raise(exception, error);
        } catch (const std::bad_alloc&) {
            return E_OUTOFMEMORY;
        }
        if (entry->forward && !wants_put && !call.args.empty() &&
            result != nullptr && result->vt == VT_DISPATCH &&
            result->pdispVal != nullptr) {
            IDispatch* inner = result->pdispVal;
            VARIANT forwarded;
            VariantInit(&forwarded);
            const HRESULT outcome =
                inner->Invoke(DISPID_VALUE, IID_NULL, lcid, flags, params,
                              &forwarded, exception, nullptr);
            inner->Release();
            *result = forwarded;
            return outcome;
        }
        return S_OK;
    }

private:
    static HRESULT raise(EXCEPINFO* exception, const ComError& error) {
        if (exception != nullptr) {
            *exception = EXCEPINFO{};
            const long number = error.number;
            exception->scode =
                (number > 0 && number < 65536)
                    ? static_cast<SCODE>(0x800A0000UL |
                                         static_cast<unsigned long>(number))
                    : static_cast<SCODE>(number);
            exception->bstrSource = SysAllocString(L"ADODB");
            const std::wstring text = to_wide(error.text);
            exception->bstrDescription =
                SysAllocStringLen(text.data(), static_cast<UINT>(text.size()));
        }
        return DISP_E_EXCEPTION;
    }

    std::atomic<long> references_{1};
};

// An IEnumVARIANT over a list of values taken when For Each starts.
class Enumerator : public IEnumVARIANT {
public:
    explicit Enumerator(std::vector<Var> items) : items_(std::move(items)) {
        ++g_objects;
    }
    ~Enumerator() { --g_objects; }
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid,
                                             void** object) override {
        if (object == nullptr) {
            return E_POINTER;
        }
        if (iid == IID_IUnknown || iid == IID_IEnumVARIANT) {
            *object = static_cast<IEnumVARIANT*>(this);
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
    HRESULT STDMETHODCALLTYPE Next(ULONG count, VARIANT* out,
                                   ULONG* fetched) override {
        ULONG done = 0;
        while (done < count && position_ < items_.size()) {
            VariantInit(&out[done]);
            VariantCopy(&out[done],
                        const_cast<VARIANT*>(&items_[position_].ref()));
            ++done;
            ++position_;
        }
        if (fetched != nullptr) {
            *fetched = done;
        }
        return done == count ? S_OK : S_FALSE;
    }
    HRESULT STDMETHODCALLTYPE Skip(ULONG count) override {
        position_ = std::min<std::size_t>(items_.size(), position_ + count);
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE Reset() override {
        position_ = 0;
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE Clone(IEnumVARIANT** out) override {
        auto* copy = new (std::nothrow) Enumerator(items_);
        if (copy == nullptr) {
            return E_OUTOFMEMORY;
        }
        copy->position_ = position_;
        *out = copy;
        return S_OK;
    }

private:
    std::vector<Var> items_;
    std::size_t position_{0};
    std::atomic<long> references_{1};
};

}  // namespace wfc::adodb::com

#endif  // WFC_ADODB_COM_DISPATCH_HPP
