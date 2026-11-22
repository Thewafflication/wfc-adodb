// Registration of the ADODB library with Windows: the entry points regsvr32
// calls, and DllGetClassObject for the classes Connection, Recordset and
// Command. The work of registering is in wfc/ocx/registration.hpp (the copy
// every WFC library carries); the objects are in com/classes.cpp.

#include "com/classes.hpp"
#include "registration_info.hpp"

#include <wfc/ocx/registration.hpp>

#include <cwchar>
#include <iterator>

namespace {

const wfc::registration::Library kLibrary{L"adodb", L"" WFC_LIB_VERSION,
                                          L"" WFC_LIB_TARGET};

// WFC's own ProgIDs and class ids, and the Microsoft ones used when
// WFC_REGISTER_MICROSOFT_IDS=1 (the classes of msado15.dll).
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
};

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

}  // namespace

extern "C" {

HRESULT __stdcall DllRegisterServer() {
    return wfc::registration::register_library(this_module(), kLibrary,
                                               kClasses, std::size(kClasses));
}

HRESULT __stdcall DllUnregisterServer() {
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
