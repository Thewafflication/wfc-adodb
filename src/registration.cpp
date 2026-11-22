// Registration of the ADODB library with Windows: the entry points regsvr32
// calls, and DllGetClassObject for the classes Connection, Recordset, Command,
// Parameter, Record and Stream. The work of registering the classes is in
// wfc/ocx/registration.hpp (the copy every WFC library carries); the type
// library embedded in the DLL is registered here; the objects are in com/.

#include "com/objects.hpp"
#include "registration_info.hpp"

#include <wfc/ocx/registration.hpp>

#include <cwchar>
#include <iterator>
#include <string>

namespace {

const wfc::registration::Library kLibrary{L"adodb", L"" WFC_LIB_VERSION,
                                          L"" WFC_LIB_TARGET};

// WFC's own ProgIDs and class ids, and the Microsoft ones used when
// WFC_REGISTER_MICROSOFT_IDS=1 (the classes of msado15.dll). The order is the
// one of the class factories in com/common.cpp.
const wfc::registration::ComClass kClasses[] = {
    {L"Connection", L"{2A83196F-06D7-5B78-A8D4-D05306BADCD5}",
     L"WFC.ADODB.Connection", L"{00000514-0000-0010-8000-00AA006D2EA4}",
     L"ADODB.Connection"},
    {L"Recordset", L"{150FB6F7-076D-5F92-94E5-D26EF5FA235A}",
     L"WFC.ADODB.Recordset", L"{00000535-0000-0010-8000-00AA006D2EA4}",
     L"ADODB.Recordset"},
    {L"Command", L"{809901FD-185D-5CC0-B9FF-02A74E0FCFB3}",
     L"WFC.ADODB.Command", L"{00000507-0000-0010-8000-00AA006D2EA4}",
     L"ADODB.Command"},
    {L"Parameter", L"{AC191646-2F4E-5C18-B250-7A057E16F314}",
     L"WFC.ADODB.Parameter", L"{0000050B-0000-0010-8000-00AA006D2EA4}",
     L"ADODB.Parameter"},
    {L"Record", L"{2CB3716F-0DCF-57E9-B84B-B4BA056C4735}", L"WFC.ADODB.Record",
     L"{00000560-0000-0010-8000-00AA006D2EA4}", L"ADODB.Record"},
    {L"Stream", L"{FAA7CC41-96FE-5C75-B976-B93F78EE0F93}", L"WFC.ADODB.Stream",
     L"{00000566-0000-0010-8000-00AA006D2EA4}", L"ADODB.Stream"},
};

// The library id of adodb.idl, and its version.
const wchar_t kTypeLibText[] = L"{A5626876-1B8D-53E5-9D50-955E6AEB095C}";
const GUID kTypeLibId = {0xA5626876,
                         0x1B8D,
                         0x53E5,
                         {0x9D, 0x50, 0x95, 0x5E, 0x6A, 0xEB, 0x09, 0x5C}};
constexpr WORD kTypeLibMajor = 6;
constexpr WORD kTypeLibMinor = 1;

HMODULE this_module() {
    HMODULE module = nullptr;
    GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS |
                           GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
                       reinterpret_cast<LPCWSTR>(&this_module), &module);
    return module;
}

bool same_clsid(REFCLSID clsid, const wchar_t* text) {
    CLSID parsed{};
    return text != nullptr && SUCCEEDED(CLSIDFromString(text, &parsed)) &&
           IsEqualCLSID(clsid, parsed);
}

constexpr SYSKIND kSyskind = sizeof(void*) == 8 ? SYS_WIN64 : SYS_WIN32;

// Registers the type library embedded in the DLL, and points each class at it.
HRESULT register_type_library(HMODULE module) {
    const bool machine = wfc::registration::choose_machine_scope();
    const std::wstring path = wfc::registration::detail::module_path(module);
    ITypeLib* library = nullptr;
    HRESULT result = LoadTypeLibEx(path.c_str(), REGKIND_NONE, &library);
    if (FAILED(result)) {
        return result;
    }
    result = machine
                 ? RegisterTypeLib(library, path.c_str(), nullptr)
                 : RegisterTypeLibForUser(
                       library, const_cast<LPOLESTR>(path.c_str()), nullptr);
    library->Release();
    if (FAILED(result)) {
        return result;
    }
    const HKEY hive = machine ? HKEY_LOCAL_MACHINE : HKEY_CURRENT_USER;
    const std::wstring typelib = kTypeLibText;
    const std::wstring version = L"6.1";
    for (const auto& entry : kClasses) {
        std::vector<const wchar_t*> ids = {entry.wfc_clsid};
        if (wfc::registration::microsoft_ids() && entry.ms_clsid != nullptr) {
            ids.push_back(entry.ms_clsid);
        }
        for (const wchar_t* id : ids) {
            const std::wstring base =
                std::wstring(L"Software\\Classes\\CLSID\\") + id;
            if (!wfc::registration::detail::write_key(hive, base + L"\\TypeLib",
                                                      &typelib) ||
                !wfc::registration::detail::write_key(hive, base + L"\\Version",
                                                      &version)) {
                return SELFREG_E_CLASS;
            }
        }
    }
    return S_OK;
}

void unregister_type_library() {
    if (wfc::registration::choose_machine_scope()) {
        UnRegisterTypeLib(kTypeLibId, kTypeLibMajor, kTypeLibMinor, 0,
                          kSyskind);
    } else {
        UnRegisterTypeLibForUser(kTypeLibId, kTypeLibMajor, kTypeLibMinor, 0,
                                 kSyskind);
    }
}

}  // namespace

extern "C" {

HRESULT __stdcall DllRegisterServer() {
    const HRESULT classes = wfc::registration::register_library(
        this_module(), kLibrary, kClasses, std::size(kClasses));
    if (FAILED(classes)) {
        return classes;
    }
    return register_type_library(this_module());
}

HRESULT __stdcall DllUnregisterServer() {
    unregister_type_library();
    return wfc::registration::unregister_library(
        kLibrary, kClasses, std::size(kClasses), this_module());
}

HRESULT __stdcall DllCanUnloadNow() {
    return wfc::adodb::com::can_unload() ? S_OK : S_FALSE;
}

HRESULT __stdcall DllGetClassObject(REFCLSID clsid, REFIID iid, void** object) {
    for (std::size_t i = 0; i < std::size(kClasses); ++i) {
        if (same_clsid(clsid, kClasses[i].wfc_clsid) ||
            same_clsid(clsid, kClasses[i].ms_clsid)) {
            return wfc::adodb::com::make_class_factory(i, iid, object);
        }
    }
    if (object != nullptr) {
        *object = nullptr;
    }
    return CLASS_E_CLASSNOTAVAILABLE;
}

}  // extern "C"
