// Writes the IDL of the ADODB type library of WFC from the type library of the
// Microsoft msado15.dll of the machine it runs on: the interfaces become plain
// dispinterfaces (no vtable to implement), the members keep the names, the
// DISPIDs and the parameters of the Microsoft ones, and every identifier
// (library, interfaces, classes) is WFC's own: the text @IID:name@ stands for
// the id of `name`, which tools/make_idl.py replaces by a name-based UUID.
//
// This is a developer tool; adodb.idl is generated once and kept. Build:
//   cl /EHsc tlb2idl.cpp /link ole32.lib oleaut32.lib
// Run: tlb2idl "C:\Program Files\Common Files\System\ado\msado15.dll"

#define WIN32_LEAN_AND_MEAN
#include <oleauto.h>
#include <windows.h>

#include <cstdio>
#include <set>
#include <string>
#include <vector>

namespace {

std::string narrow(const wchar_t* text) {
    if (text == nullptr) {
        return {};
    }
    char buffer[1024];
    WideCharToMultiByte(CP_UTF8, 0, text, -1, buffer, sizeof(buffer), nullptr,
                        nullptr);
    return buffer;
}

std::string name_of(ITypeInfo* info) {
    BSTR name = nullptr;
    info->GetDocumentation(MEMBERID_NIL, &name, nullptr, nullptr, nullptr);
    const std::string text = narrow(name);
    SysFreeString(name);
    return text;
}

std::string vt_name(const VARTYPE vt) {
    switch (vt & VT_TYPEMASK) {
        case VT_I2:
            return "short";
        case VT_I4:
            return "long";
        case VT_R4:
            return "float";
        case VT_R8:
            return "double";
        case VT_CY:
            return "CURRENCY";
        case VT_DATE:
            return "DATE";
        case VT_BSTR:
            return "BSTR";
        case VT_DISPATCH:
            return "IDispatch*";
        case VT_ERROR:
            return "SCODE";
        case VT_BOOL:
            return "VARIANT_BOOL";
        case VT_VARIANT:
            return "VARIANT";
        case VT_UNKNOWN:
            return "IUnknown*";
        case VT_DECIMAL:
            return "DECIMAL";
        case VT_I1:
            return "char";
        case VT_UI1:
            return "unsigned char";
        case VT_UI2:
            return "unsigned short";
        case VT_UI4:
            return "unsigned long";
        case VT_I8:
            return "hyper";
        case VT_UI8:
            return "unsigned hyper";
        case VT_INT:
            return "int";
        case VT_UINT:
            return "unsigned int";
        case VT_VOID:
            return "void";
        case VT_HRESULT:
            return "HRESULT";
        default:
            return "long";
    }
}

std::string type_text(ITypeInfo* info, const TYPEDESC& type) {
    switch (type.vt) {
        case VT_PTR:
            return type_text(info, *type.lptdesc) + "*";
        case VT_SAFEARRAY:
            return "SAFEARRAY(" + type_text(info, *type.lptdesc) + ")";
        case VT_USERDEFINED: {
            ITypeInfo* ref = nullptr;
            if (FAILED(info->GetRefTypeInfo(type.hreftype, &ref))) {
                return "long";
            }
            TYPEATTR* attr = nullptr;
            ref->GetTypeAttr(&attr);
            std::string text = name_of(ref);
            if (attr->typekind == TKIND_ALIAS) {
                text = type_text(ref, attr->tdescAlias);
            } else if (attr->typekind == TKIND_DISPATCH ||
                       attr->typekind == TKIND_INTERFACE ||
                       attr->typekind == TKIND_COCLASS) {
                // The interfaces we keep are named here; the versioned ones
                // (Recordset15 and so on) fold into the main one.
                text = text;
            }
            ref->ReleaseTypeAttr(attr);
            ref->Release();
            return text;
        }
        default:
            return vt_name(type.vt);
    }
}

std::string default_text(const PARAMDESCEX* ex, const std::string& type) {
    VARIANT value;
    VariantInit(&value);
    if (ex == nullptr || FAILED(VariantChangeType(
                             &value, const_cast<VARIANT*>(&ex->varDefaultValue),
                             0, type == "BSTR" ? VT_BSTR : VT_I4))) {
        return type == "BSTR" ? "\"\"" : "0";
    }
    if (type == "BSTR") {
        return "\"" + narrow(value.bstrVal) + "\"";
    }
    return std::to_string(value.lVal);
}

// The dispinterface kept for an official interface name.
const char* const kKept[] = {
    "_Connection", "_Recordset", "_Command",   "_Parameter", "Fields",
    "Field",       "Parameters", "Properties", "Property",   "Errors",
    "Error",       "_Record",    "_Stream"};

bool kept(const std::string& name) {
    for (const char* k : kKept) {
        if (name == k) {
            return true;
        }
    }
    return false;
}

void emit_dispatch(ITypeInfo* info, const std::string& name) {
    TYPEATTR* attr = nullptr;
    info->GetTypeAttr(&attr);
    std::printf(
        "    [uuid(@IID:%s@)]\n    dispinterface %s {\n"
        "    properties:\n    methods:\n",
        name.c_str(), name.c_str());
    std::set<std::string> seen;
    for (UINT f = 0; f < attr->cFuncs; ++f) {
        FUNCDESC* fd = nullptr;
        info->GetFuncDesc(f, &fd);
        BSTR names[40] = {};
        UINT count = 0;
        info->GetNames(fd->memid, names, 40, &count);
        const std::string member = narrow(names[0]);
        if (member == "QueryInterface" || member == "AddRef" ||
            member == "Release" || member == "GetTypeInfoCount" ||
            member == "GetTypeInfo" || member == "GetIDsOfNames" ||
            member == "Invoke") {
            info->ReleaseFuncDesc(fd);
            continue;  // the IUnknown and IDispatch members of a dual interface
        }
        std::string attrs = "id(" + std::to_string(fd->memid) + ")";
        if (fd->invkind == INVOKE_PROPERTYGET) {
            attrs += ", propget";
        } else if (fd->invkind == INVOKE_PROPERTYPUT) {
            attrs += ", propput";
        } else if (fd->invkind == INVOKE_PROPERTYPUTREF) {
            attrs += ", propputref";
        }
        if ((fd->wFuncFlags & FUNCFLAG_FRESTRICTED) != 0) {
            attrs += ", restricted";
        }
        if ((fd->wFuncFlags & FUNCFLAG_FHIDDEN) != 0) {
            attrs += ", hidden";
        }
        std::string result = type_text(info, fd->elemdescFunc.tdesc);
        std::string params;
        for (int p = 0; p < fd->cParams; ++p) {
            const ELEMDESC& element = fd->lprgelemdescParam[p];
            const USHORT flags = element.paramdesc.wParamFlags;
            std::string pa;
            auto add = [&pa](const char* word) {
                pa += pa.empty() ? "" : ", ";
                pa += word;
            };
            if ((flags & PARAMFLAG_FIN) != 0) {
                add("in");
            }
            if ((flags & PARAMFLAG_FOUT) != 0) {
                add("out");
            }
            const std::string type = type_text(info, element.tdesc);
            // MIDL wants [optional] only on VARIANT parameters; a default
            // value alone makes the others optional.
            if ((flags & PARAMFLAG_FOPT) != 0 &&
                ((flags & PARAMFLAG_FHASDEFAULT) == 0 ||
                 (type == "VARIANT" || type == "VARIANT*"))) {
                add("optional");
            }
            if ((flags & PARAMFLAG_FHASDEFAULT) != 0) {
                add(("defaultvalue(" +
                     default_text(element.paramdesc.pparamdescex, type) + ")")
                        .c_str());
            }
            if (pa.empty()) {
                pa = "in";
            }
            std::string pname = (static_cast<UINT>(p) + 1U < count)
                                    ? narrow(names[p + 1])
                                    : "rhs";
            if (pname.empty() || pname == "?") {
                pname = "rhs";
            }
            params += (p ? ", " : "");
            params += "[" + pa + "] " + type + " " + pname;
        }
        // A property put returns nothing in a dispinterface.
        if (fd->invkind == INVOKE_PROPERTYPUT ||
            fd->invkind == INVOKE_PROPERTYPUTREF) {
            result = "void";
        }
        std::printf("        [%s] %s %s(%s);\n", attrs.c_str(), result.c_str(),
                    member.c_str(), params.c_str());
        info->ReleaseFuncDesc(fd);
    }
    std::printf("    };\n\n");
    info->ReleaseTypeAttr(attr);
}

void emit_enum(ITypeInfo* info, const std::string& name) {
    TYPEATTR* attr = nullptr;
    info->GetTypeAttr(&attr);
    std::printf("    typedef enum %s {\n", name.c_str());
    for (UINT v = 0; v < attr->cVars; ++v) {
        VARDESC* vd = nullptr;
        info->GetVarDesc(v, &vd);
        BSTR member = nullptr;
        UINT count = 0;
        info->GetNames(vd->memid, &member, 1, &count);
        VARIANT value;
        VariantInit(&value);
        VariantChangeType(&value, vd->lpvarValue, 0, VT_I4);
        std::printf("        %s = %ld%s\n", narrow(member).c_str(), value.lVal,
                    v + 1U < attr->cVars ? "," : "");
        info->ReleaseVarDesc(vd);
    }
    std::printf("    } %s;\n\n", name.c_str());
    info->ReleaseTypeAttr(attr);
}

}  // namespace

int main(int argc, char** argv) {
    if (argc < 2) {
        std::fprintf(stderr, "usage: tlb2idl <msado15.dll>\n");
        return 2;
    }
    CoInitialize(nullptr);
    wchar_t path[MAX_PATH];
    MultiByteToWideChar(CP_UTF8, 0, argv[1], -1, path, MAX_PATH);
    ITypeLib* lib = nullptr;
    if (FAILED(LoadTypeLib(path, &lib))) {
        std::fprintf(stderr, "cannot load the type library\n");
        return 1;
    }
    std::printf(
        "// ADODB type library of WFC. Generated by tools/tlb2idl.cpp and\n"
        "// tools/make_idl.py from the member lists of the ADO 6.1 type\n"
        "// library; every interface is a dispinterface (WFC's objects are\n"
        "// late-bound), and the identifiers are WFC's own.\n\n"
        "import \"oaidl.idl\";\nimport \"ocidl.idl\";\n\n"
        "[uuid(@IID:ADODB@), version(6.1), helpstring(\"ADODB for WFC\")]\n"
        "library ADODB {\n    importlib(\"stdole2.tlb\");\n\n");
    const UINT count = lib->GetTypeInfoCount();
    // Forward declarations.
    for (const char* k : kKept) {
        std::printf("    dispinterface %s;\n", k);
    }
    std::printf("\n");
    // Enums first.
    for (UINT i = 0; i < count; ++i) {
        ITypeInfo* info = nullptr;
        lib->GetTypeInfo(i, &info);
        TYPEATTR* attr = nullptr;
        info->GetTypeAttr(&attr);
        if (attr->typekind == TKIND_ENUM) {
            emit_enum(info, name_of(info));
        }
        info->ReleaseTypeAttr(attr);
        info->Release();
    }
    for (UINT i = 0; i < count; ++i) {
        ITypeInfo* info = nullptr;
        lib->GetTypeInfo(i, &info);
        TYPEATTR* attr = nullptr;
        info->GetTypeAttr(&attr);
        const std::string name = name_of(info);
        if (attr->typekind == TKIND_DISPATCH && kept(name)) {
            emit_dispatch(info, name);
        }
        info->ReleaseTypeAttr(attr);
        info->Release();
    }
    const char* const classes[][2] = {
        {"Connection", "_Connection"}, {"Recordset", "_Recordset"},
        {"Command", "_Command"},       {"Parameter", "_Parameter"},
        {"Record", "_Record"},         {"Stream", "_Stream"}};
    for (const auto& entry : classes) {
        std::printf(
            "    [uuid(@CLSID:%s@)]\n    coclass %s {\n"
            "        [default] dispinterface %s;\n    };\n\n",
            entry[0], entry[0], entry[1]);
    }
    std::printf("};\n");
    return 0;
}
