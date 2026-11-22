// TC-0003: msado15.dll as a COM server. Loads the DLL, checks its entry points
// and version resource, registers it for the current user, creates the ADODB
// classes through COM and drives them through IDispatch the way a late-bound
// client (a VB6 program, VBScript) does, then unregisters it.
//
// usage: adodb_dll_test <path to msado15.dll> <version>

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>

#include <objbase.h>
#include <oleauto.h>

#include <cstdlib>
#include <iostream>
#include <string>
#include <vector>

namespace {

int failures = 0;

void check(const bool condition, const char* what) {
    if (!condition) {
        std::cerr << "FAIL: " << what << '\n';
        ++failures;
    }
}

VARIANT number(const long value) {
    VARIANT v;
    VariantInit(&v);
    v.vt = VT_I4;
    v.lVal = value;
    return v;
}

VARIANT text(const wchar_t* value) {
    VARIANT v;
    VariantInit(&v);
    v.vt = VT_BSTR;
    v.bstrVal = SysAllocString(value);
    return v;
}

VARIANT real(const double value) {
    VARIANT v;
    VariantInit(&v);
    v.vt = VT_R8;
    v.dblVal = value;
    return v;
}

// An argument the program left out.
VARIANT nothing() {
    VARIANT v;
    VariantInit(&v);
    v.vt = VT_ERROR;
    v.scode = DISP_E_PARAMNOTFOUND;
    return v;
}

// A reference to an object as an argument (the call releases it).
VARIANT object_var(IDispatch* object) {
    VARIANT v;
    VariantInit(&v);
    v.vt = VT_DISPATCH;
    v.pdispVal = object;
    object->AddRef();
    return v;
}

struct Outcome {
    HRESULT hr{S_OK};
    VARIANT value{};
    EXCEPINFO exception{};
    Outcome() { VariantInit(&value); }
    Outcome(Outcome&& other) noexcept
        : hr(other.hr), value(other.value), exception(other.exception) {
        VariantInit(&other.value);
        other.exception = EXCEPINFO{};
    }
    Outcome(const Outcome&) = delete;
    Outcome& operator=(const Outcome&) = delete;
    ~Outcome() {
        VariantClear(&value);
        SysFreeString(exception.bstrSource);
        SysFreeString(exception.bstrDescription);
    }
};

// Calls a member by name: `args` in the order the program wrote them (for a
// property put the new value is the last one).
void invoke(IDispatch* object, const wchar_t* name, const WORD flags,
            std::vector<VARIANT> args, Outcome& out) {
    DISPID id = 0;
    LPOLESTR member = const_cast<LPOLESTR>(name);
    out.hr =
        object->GetIDsOfNames(IID_NULL, &member, 1, LOCALE_USER_DEFAULT, &id);
    if (FAILED(out.hr)) {
        return;
    }
    std::vector<VARIANT> reversed(args.rbegin(), args.rend());
    DISPPARAMS params{};
    params.cArgs = static_cast<UINT>(reversed.size());
    params.rgvarg = reversed.empty() ? nullptr : reversed.data();
    DISPID put = DISPID_PROPERTYPUT;
    if ((flags & DISPATCH_PROPERTYPUT) != 0) {
        params.cNamedArgs = 1;
        params.rgdispidNamedArgs = &put;
    }
    out.hr = object->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT, flags, &params,
                            &out.value, &out.exception, nullptr);
    for (auto& argument : args) {
        VariantClear(&argument);
    }
}

Outcome call(IDispatch* object, const wchar_t* name,
             std::vector<VARIANT> args = {}) {
    Outcome out;
    invoke(object, name, DISPATCH_METHOD | DISPATCH_PROPERTYGET,
           std::move(args), out);
    return out;
}

HRESULT put(IDispatch* object, const wchar_t* name, VARIANT value) {
    Outcome out;
    invoke(object, name, DISPATCH_PROPERTYPUT, {value}, out);
    return out.hr;
}

long as_long(const Outcome& out) {
    VARIANT copy;
    VariantInit(&copy);
    if (FAILED(VariantChangeType(&copy, const_cast<VARIANT*>(&out.value), 0,
                                 VT_I4))) {
        return -999999;
    }
    return copy.lVal;
}

std::wstring as_text(const Outcome& out) {
    VARIANT copy;
    VariantInit(&copy);
    if (FAILED(VariantChangeType(&copy, const_cast<VARIANT*>(&out.value), 0,
                                 VT_BSTR))) {
        return L"<not text>";
    }
    std::wstring result(copy.bstrVal);
    VariantClear(&copy);
    return result;
}

// The object an Outcome holds (caller releases), or null.
IDispatch* as_object(Outcome& out) {
    if (out.value.vt != VT_DISPATCH || out.value.pdispVal == nullptr) {
        return nullptr;
    }
    IDispatch* object = out.value.pdispVal;
    object->AddRef();
    return object;
}

// The type information of an object lists every member of its dispinterface;
// each one must be answered by the object under the same DISPID.
void check_members(IDispatch* object, const char* label) {
    UINT infos = 0;
    object->GetTypeInfoCount(&infos);
    check(infos == 1, label);
    ITypeInfo* info = nullptr;
    if (FAILED(object->GetTypeInfo(0, LOCALE_USER_DEFAULT, &info))) {
        check(false, label);
        return;
    }
    TYPEATTR* attr = nullptr;
    info->GetTypeAttr(&attr);
    int missing = 0;
    for (UINT f = 0; f < attr->cFuncs; ++f) {
        FUNCDESC* fd = nullptr;
        info->GetFuncDesc(f, &fd);
        BSTR name = nullptr;
        info->GetDocumentation(fd->memid, &name, nullptr, nullptr, nullptr);
        DISPID id = 0;
        LPOLESTR names[1] = {name};
        const HRESULT hr =
            object->GetIDsOfNames(IID_NULL, names, 1, LOCALE_USER_DEFAULT, &id);
        if (FAILED(hr) || id != fd->memid) {
            std::wcerr << L"member not answered: " << label << L"."
                       << (name != nullptr ? name : L"?") << L'\n';
            ++missing;
        }
        SysFreeString(name);
        info->ReleaseFuncDesc(fd);
    }
    info->ReleaseTypeAttr(attr);
    info->Release();
    check(missing == 0, label);
}

IDispatch* create(const wchar_t* progid) {
    IDispatch* object = nullptr;
    CLSID clsid{};
    if (FAILED(CLSIDFromProgID(progid, &clsid))) {
        return nullptr;
    }
    if (FAILED(CoCreateInstance(clsid, nullptr, CLSCTX_INPROC_SERVER,
                                IID_IDispatch,
                                reinterpret_cast<void**>(&object)))) {
        return nullptr;
    }
    return object;
}

long error_number(const Outcome& out) {
    return out.hr == DISP_E_EXCEPTION ? static_cast<long>(out.exception.scode)
                                      : 0;
}

}  // namespace

int main(int argc, char** argv) {
    if (argc < 3) {
        std::cerr << "usage: adodb_dll_test <msado15.dll> <version>\n";
        return EXIT_FAILURE;
    }
    const std::string path = argv[1];
    const std::string version = argv[2];

    HMODULE dll = LoadLibraryA(path.c_str());
    check(dll != nullptr, "the DLL loads");
    if (dll == nullptr) {
        return EXIT_FAILURE;
    }
    using Entry = HRESULT(__stdcall*)();
    const auto reg =
        reinterpret_cast<Entry>(GetProcAddress(dll, "DllRegisterServer"));
    const auto unreg =
        reinterpret_cast<Entry>(GetProcAddress(dll, "DllUnregisterServer"));
    const auto can_unload =
        reinterpret_cast<Entry>(GetProcAddress(dll, "DllCanUnloadNow"));
    check(reg != nullptr && unreg != nullptr && can_unload != nullptr &&
              GetProcAddress(dll, "DllGetClassObject") != nullptr,
          "the COM entry points are exported");

    // The version resource names the version of the library.
    DWORD handle = 0;
    const DWORD size = GetFileVersionInfoSizeA(path.c_str(), &handle);
    check(size > 0, "the DLL has a version resource");

    SetEnvironmentVariableW(L"WFC_REGISTER_SCOPE", L"user");
    CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    check(reg != nullptr && SUCCEEDED(reg()), "DllRegisterServer succeeds");

    IDispatch* connection = create(L"WFC.ADODB.Connection");
    check(connection != nullptr, "ADODB.Connection is created through COM");
    if (connection == nullptr) {
        return EXIT_FAILURE;
    }
    check(as_long(call(connection, L"State")) == 0, "closed at first");
    check(call(connection, L"Open", {text(L"Data Source=:memory:")}).hr == S_OK,
          "Open");
    check(as_long(call(connection, L"State")) == 1, "open state");
    check(error_number(call(connection, L"Open")) ==
              static_cast<long>(0x800A0000UL | 3705),
          "a second Open is error 3705");
    check(call(connection, L"Execute",
               {text(L"create table t(id integer primary key, name text, "
                     L"score real)")})
                  .hr == S_OK,
          "Execute DDL");
    check(call(connection, L"Execute",
               {text(L"insert into t(name, score) values ('ann', 1.5)")})
                  .hr == S_OK,
          "Execute insert");
    check(call(connection, L"Execute",
               {text(L"insert into t(name, score) values ('bob', 2.25)")})
                  .hr == S_OK,
          "Execute insert 2");

    // A query result: Fields, values, movement, For Each support.
    {
        Outcome made =
            call(connection, L"Execute",
                 {text(L"select id, name, score from t order by id")});
        IDispatch* rs = as_object(made);
        check(rs != nullptr, "Execute returns a Recordset");
        if (rs != nullptr) {
            check(as_long(call(rs, L"RecordCount")) == 2, "two records");
            Outcome fields = call(rs, L"Fields");
            IDispatch* list = as_object(fields);
            check(list != nullptr && as_long(call(list, L"Count")) == 3,
                  "three fields");
            if (list != nullptr) {
                Outcome item = call(list, L"Item", {number(1)});
                IDispatch* field = as_object(item);
                check(field != nullptr &&
                          as_text(call(field, L"Name")) == L"name",
                      "Fields(1).Name");
                check(field != nullptr &&
                          as_text(call(field, L"Value")) == L"ann",
                      "Fields(1).Value");
                if (field != nullptr) {
                    field->Release();
                }
                Outcome by_name = call(list, L"Item", {text(L"SCORE")});
                IDispatch* score = as_object(by_name);
                check(score != nullptr &&
                          as_text(call(score, L"Value")) == L"1.5",
                      "Fields(\"score\") by name, any case");
                if (score != nullptr) {
                    score->Release();
                }
                Outcome enumerated = call(list, L"_NewEnum");
                check(enumerated.value.vt == VT_UNKNOWN &&
                          enumerated.value.punkVal != nullptr,
                      "_NewEnum gives an enumerator");
                list->Release();
            }
            check(as_text(call(rs, L"Collect", {text(L"name")})) == L"ann",
                  "Collect(\"name\")");
            check(call(rs, L"MoveNext").hr == S_OK &&
                      as_text(call(rs, L"Collect", {number(1)})) == L"bob",
                  "MoveNext");
            check(call(rs, L"MoveNext").hr == S_OK &&
                      call(rs, L"EOF").value.boolVal == VARIANT_TRUE,
                  "EOF after the last row");
            check(error_number(call(rs, L"MoveNext")) ==
                      static_cast<long>(0x800A0000UL | 3021),
                  "MoveNext at EOF is error 3021");
            call(rs, L"MoveFirst");
            check(as_text(call(rs, L"GetString",
                               {number(2), number(-1), text(L","),
                                text(L";")})) == L"1,ann,1.5;2,bob,2.25;",
                  "GetString");
            check(put(rs, L"Sort", text(L"id DESC")) == S_OK &&
                      as_text(call(rs, L"Collect", {text(L"name")})) == L"bob",
                  "Sort");
            check(put(rs, L"Filter", text(L"score > 2")) == S_OK &&
                      as_long(call(rs, L"RecordCount")) == 1,
                  "Filter");
            check(error_number(call(rs, L"Delete")) ==
                      static_cast<long>(0x800A0000UL | 3251),
                  "a query result is read-only");
            check(call(rs, L"Close").hr == S_OK &&
                      as_long(call(rs, L"State")) == 0,
                  "Close");
            rs->Release();
        }
    }

    // An updatable recordset writes back to its table.
    {
        IDispatch* rs = create(L"WFC.ADODB.Recordset");
        check(rs != nullptr, "ADODB.Recordset is created");
        if (rs != nullptr) {
            Outcome open = call(rs, L"Open",
                                {text(L"select id, name from t order by id"),
                                 object_var(connection), number(1), number(3)});
            check(open.hr == S_OK, "Open with a lock type");
            check(call(rs, L"Fields").hr == S_OK, "Fields on an open set");
            Outcome field = call(rs, L"Fields");
            IDispatch* list = as_object(field);
            if (list != nullptr) {
                Outcome item = call(list, L"Item", {text(L"name")});
                IDispatch* name = as_object(item);
                check(name != nullptr &&
                          put(name, L"Value", text(L"ANN")) == S_OK,
                      "assign a field");
                if (name != nullptr) {
                    name->Release();
                }
                list->Release();
            }
            check(call(rs, L"Update").hr == S_OK, "Update");
            check(call(rs, L"AddNew").hr == S_OK, "AddNew");
            Outcome assigned;
            invoke(rs, L"Collect", DISPATCH_PROPERTYPUT,
                   {text(L"name"), text(L"cy")}, assigned);
            check(assigned.hr == S_OK, "assign through Collect");
            check(call(rs, L"Update").hr == S_OK, "Update the new row");
            check(as_long(call(rs, L"Collect", {text(L"id")})) == 3,
                  "the new row has its key");
            rs->Release();
        }
        Outcome seen = call(connection, L"Execute",
                            {text(L"select name from t where id = 1")});
        check(as_long(call(connection, L"BeginTrans")) == 1, "BeginTrans");
        check(call(connection, L"RollbackTrans").hr == S_OK, "RollbackTrans");
        IDispatch* check_set = as_object(seen);
        check(check_set != nullptr &&
                  as_text(call(check_set, L"Collect", {number(0)})) == L"ANN",
              "the table has the update");
        if (check_set != nullptr) {
            check_set->Release();
        }
    }

    // A Command with Parameters.
    {
        IDispatch* command = create(L"WFC.ADODB.Command");
        check(command != nullptr, "ADODB.Command is created");
        if (command != nullptr) {
            check(put(command, L"ActiveConnection", object_var(connection)) ==
                      S_OK,
                  "ActiveConnection");
            check(put(command, L"CommandText",
                      text(L"select name from t where id >= ?")) == S_OK,
                  "CommandText");
            Outcome parameters = call(command, L"Parameters");
            IDispatch* list = as_object(parameters);
            Outcome created = call(
                command, L"CreateParameter",
                {text(L"min"), number(3), number(1), number(0), number(3)});
            IDispatch* parameter = as_object(created);
            check(list != nullptr && parameter != nullptr &&
                      call(list, L"Append", {object_var(parameter)}).hr == S_OK,
                  "Parameters.Append");
            check(list != nullptr && as_long(call(list, L"Count")) == 1,
                  "Parameters.Count");
            Outcome run = call(command, L"Execute");
            IDispatch* rs = as_object(run);
            check(rs != nullptr && as_long(call(rs, L"RecordCount")) == 1 &&
                      as_text(call(rs, L"Collect", {number(0)})) == L"cy",
                  "Command.Execute binds the parameters");
            if (rs != nullptr) {
                rs->Release();
            }
            if (parameter != nullptr) {
                parameter->Release();
            }
            if (list != nullptr) {
                list->Release();
            }
            command->Release();
        }
    }

    // select @@identity answers on SQLite.
    {
        Outcome identity =
            call(connection, L"Execute", {text(L"select @@identity")});
        IDispatch* rs = as_object(identity);
        check(rs != nullptr && as_long(call(rs, L"Collect", {number(0)})) == 3,
              "select @@identity");
        if (rs != nullptr) {
            rs->Release();
        }
    }

    // An error carries the number and text.
    {
        Outcome bad = call(connection, L"Execute", {text(L"selec nothing")});
        check(bad.hr == DISP_E_EXCEPTION &&
                  bad.exception.scode == static_cast<SCODE>(-2147217900L),
              "a syntax error is DB_E_ERRORSINCOMMAND");
    }

    // The type library is registered, and every class has its type
    // information with all the members of the dispinterface.
    {
        ITypeLib* library = nullptr;
        check(SUCCEEDED(LoadRegTypeLib(
                  {0xA5626876,
                   0x1B8D,
                   0x53E5,
                   {0x9D, 0x50, 0x95, 0x5E, 0x6A, 0xEB, 0x09, 0x5C}},
                  6, 1, 0, &library)),
              "the type library is registered");
        if (library != nullptr) {
            library->Release();
        }
        check_members(connection, "Connection");
        Outcome errors = call(connection, L"Errors");
        IDispatch* list = as_object(errors);
        check(list != nullptr, "Errors");
        if (list != nullptr) {
            check_members(list, "Errors");
            list->Release();
        }
        Outcome properties = call(connection, L"Properties");
        IDispatch* props = as_object(properties);
        check(props != nullptr && as_long(call(props, L"Count")) >= 1,
              "Properties holds some properties");
        if (props != nullptr) {
            check_members(props, "Properties");
            Outcome item = call(props, L"Item", {number(0)});
            IDispatch* property = as_object(item);
            check(property != nullptr, "Properties.Item");
            if (property != nullptr) {
                check_members(property, "Property");
                property->Release();
            }
            props->Release();
        }
        for (const wchar_t* progid :
             {L"WFC.ADODB.Recordset", L"WFC.ADODB.Command",
              L"WFC.ADODB.Parameter", L"WFC.ADODB.Record",
              L"WFC.ADODB.Stream"}) {
            IDispatch* made = create(progid);
            check(made != nullptr, "every class is created through COM");
            if (made != nullptr) {
                check_members(made, "created class");
                made->Release();
            }
        }
    }

    // Errors: a failed call leaves its error in the collection.
    {
        call(connection, L"Execute", {text(L"selec nothing")});
        Outcome errors = call(connection, L"Errors");
        IDispatch* list = as_object(errors);
        check(list != nullptr && as_long(call(list, L"Count")) == 1,
              "Errors.Count after a failure");
        if (list != nullptr) {
            Outcome item = call(list, L"Item", {number(0)});
            IDispatch* error = as_object(item);
            check(error != nullptr &&
                      as_long(call(error, L"Number")) == -2147217900L,
                  "Errors(0).Number");
            if (error != nullptr) {
                check_members(error, "Error");
                check(!as_text(call(error, L"Description")).empty(),
                      "Errors(0).Description");
                error->Release();
            }
            check(call(list, L"Clear").hr == S_OK &&
                      as_long(call(list, L"Count")) == 0,
                  "Errors.Clear");
            list->Release();
        }
    }

    // Recordset members that were missing: Clone, Requery, EditMode, Status,
    // Supports, paging, GetRows, Find, Save and Open of a file.
    {
        Outcome made = call(connection, L"Execute",
                            {text(L"select id, name from t order by id")});
        IDispatch* rs = as_object(made);
        check(rs != nullptr, "Execute for the member checks");
        if (rs != nullptr) {
            check_members(rs, "Recordset");
            check(as_long(call(rs, L"PageCount")) == 1 &&
                      as_long(call(rs, L"AbsolutePage")) == 1,
                  "PageCount, AbsolutePage");
            check(as_long(call(rs, L"EditMode")) == 0, "EditMode none");
            check(call(rs, L"Supports", {number(0x2000)}).value.boolVal ==
                      VARIANT_TRUE,
                  "Supports adBookmark");
            Outcome copy = call(rs, L"Clone");
            IDispatch* twin = as_object(copy);
            check(twin != nullptr && as_long(call(twin, L"RecordCount")) ==
                                         as_long(call(rs, L"RecordCount")),
                  "Clone");
            if (twin != nullptr) {
                check(call(twin, L"MoveLast").hr == S_OK &&
                          as_long(call(rs, L"AbsolutePosition")) == 1,
                      "a clone moves on its own");
                twin->Release();
            }
            Outcome rows = call(rs, L"GetRows", {number(2)});
            check(rows.hr == S_OK && rows.value.vt == (VT_ARRAY | VT_VARIANT),
                  "GetRows");
            call(rs, L"MoveFirst");
            check(call(rs, L"Find", {text(L"name = 'bob'")}).hr == S_OK &&
                      as_text(call(rs, L"Collect", {text(L"name")})) == L"bob",
                  "Find");
            check(call(rs, L"Requery").hr == S_OK &&
                      as_long(call(rs, L"RecordCount")) >= 2,
                  "Requery");
            // Save and open the file again.
            wchar_t folder[MAX_PATH];
            GetTempPathW(MAX_PATH, folder);
            const std::wstring file =
                std::wstring(folder) + L"wfc-adodb-test-" +
                std::to_wstring(GetCurrentProcessId()) + L".xml";
            check(call(rs, L"Save", {text(file.c_str())}).hr == S_OK,
                  "Save to a file");
            rs->Release();
            IDispatch* loaded = create(L"WFC.ADODB.Recordset");
            check(loaded != nullptr && call(loaded, L"Open",
                                            {text(file.c_str()), nothing(),
                                             number(0), number(3), number(256)})
                                               .hr == S_OK,
                  "Open from a file");
            if (loaded != nullptr) {
                check(as_long(call(loaded, L"RecordCount")) >= 2,
                      "the saved rows come back");
                check(as_text(call(loaded, L"Collect", {text(L"name")})) !=
                          L"<not text>",
                      "the saved values come back");
                loaded->Release();
            }
            DeleteFileW(file.c_str());
        }
    }

    // OpenSchema.
    {
        Outcome tables = call(connection, L"OpenSchema", {number(20)});
        IDispatch* rs = as_object(tables);
        check(rs != nullptr && as_long(call(rs, L"RecordCount")) >= 1 &&
                  as_text(call(rs, L"Collect", {text(L"TABLE_NAME")})) == L"t",
              "OpenSchema(adSchemaTables)");
        if (rs != nullptr) {
            rs->Release();
        }
        Outcome columns = call(connection, L"OpenSchema", {number(4)});
        IDispatch* cs = as_object(columns);
        check(cs != nullptr && as_long(call(cs, L"RecordCount")) >= 3,
              "OpenSchema(adSchemaColumns)");
        if (cs != nullptr) {
            cs->Release();
        }
    }

    // Stream.
    {
        IDispatch* stream = create(L"WFC.ADODB.Stream");
        check(stream != nullptr, "ADODB.Stream is created");
        if (stream != nullptr) {
            check(error_number(call(stream, L"WriteText", {text(L"x")})) ==
                      static_cast<long>(0x800A0000UL | 3704),
                  "a closed Stream refuses writes");
            check(call(stream, L"Open").hr == S_OK, "Stream.Open");
            check(put(stream, L"Charset", text(L"utf-8")) == S_OK,
                  "Stream.Charset");
            check(
                call(stream, L"WriteText", {text(L"one"), number(1)}).hr ==
                        S_OK &&
                    call(stream, L"WriteText", {text(L"two"), number(1)}).hr ==
                        S_OK,
                "Stream.WriteText lines");
            check(put(stream, L"Position", number(0)) == S_OK,
                  "Stream.Position");
            check(as_text(call(stream, L"ReadText", {number(-2)})) == L"one",
                  "Stream.ReadText a line");
            check(as_text(call(stream, L"ReadText")) == L"two\r\n",
                  "Stream.ReadText the rest");
            check(call(stream, L"EOS").value.boolVal == VARIANT_TRUE,
                  "Stream.EOS");
            wchar_t folder[MAX_PATH];
            GetTempPathW(MAX_PATH, folder);
            const std::wstring file =
                std::wstring(folder) + L"wfc-adodb-stream-" +
                std::to_wstring(GetCurrentProcessId()) + L".txt";
            check(call(stream, L"SaveToFile", {text(file.c_str()), number(2)})
                          .hr == S_OK,
                  "Stream.SaveToFile");
            check(call(stream, L"Close").hr == S_OK, "Stream.Close");
            check(call(stream, L"Open").hr == S_OK &&
                      call(stream, L"LoadFromFile", {text(file.c_str())}).hr ==
                          S_OK &&
                      as_long(call(stream, L"Size")) > 6,
                  "Stream.LoadFromFile");
            check_members(stream, "Stream");
            stream->Release();
            DeleteFileW(file.c_str());
        }
    }

    // Record: the current row of a recordset.
    {
        Outcome made =
            call(connection, L"Execute", {text(L"select id, name from t")});
        IDispatch* rs = as_object(made);
        IDispatch* record = create(L"WFC.ADODB.Record");
        check(rs != nullptr && record != nullptr, "Record and Recordset");
        if (rs != nullptr && record != nullptr) {
            check(call(record, L"Open", {object_var(rs)}).hr == S_OK,
                  "Record.Open on a recordset");
            Outcome fields = call(record, L"Fields");
            IDispatch* list = as_object(fields);
            check(list != nullptr && as_long(call(list, L"Count")) == 2,
                  "Record.Fields");
            if (list != nullptr) {
                list->Release();
            }
            check(as_long(call(record, L"State")) == 1, "Record.State");
            check_members(record, "Record");
        }
        if (record != nullptr) {
            record->Release();
        }
        if (rs != nullptr) {
            rs->Release();
        }
    }

    check(call(connection, L"Close").hr == S_OK, "Close the connection");
    connection->Release();

    check(can_unload != nullptr && can_unload() == S_OK,
          "DllCanUnloadNow once every object is released");

    check(unreg != nullptr && SUCCEEDED(unreg()),
          "DllUnregisterServer succeeds");
    CLSID gone{};
    check(FAILED(CLSIDFromProgID(L"WFC.ADODB.Connection", &gone)),
          "the ProgID is gone after unregistering");
    CoUninitialize();

    if (failures == 0) {
        std::cout << "PASS (msado15.dll " << version << ")\n";
    }
    return failures == 0 ? EXIT_SUCCESS : EXIT_FAILURE;
}
