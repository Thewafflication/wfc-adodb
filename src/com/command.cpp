// ADODB.Command, Parameters and Parameter as COM objects.

#include "objects.hpp"

namespace wfc::adodb::com {

// ----- Parameter
// ------------------------------------------------------------------

ParameterObj::~ParameterObj() {
    if (properties != nullptr) {
        properties->Release();
    }
}

namespace {

void append_chunk(ParameterObj& s, Call& c) {
    if (c.count() < 1) {
        throw ComError{450, "Wrong number of arguments"};
    }
    const Value more = to_value(c.args[0]);
    Value current = to_value(s.value.ref());
    if (const auto* text = std::get_if<std::string>(&more)) {
        const std::string before = std::holds_alternative<std::string>(current)
                                       ? std::get<std::string>(current)
                                       : std::string{};
        Var joined;
        joined.set_text(before + *text);
        s.value = joined;
    } else if (const auto* bytes = std::get_if<Blob>(&more)) {
        Blob before = std::holds_alternative<Blob>(current)
                          ? std::get<Blob>(current)
                          : Blob{};
        before.insert(before.end(), bytes->begin(), bytes->end());
        Var joined;
        from_value(Value{std::move(before)}, 205, joined);
        s.value = joined;
    } else {
        throw ComError{3001, kBadArgumentText};
    }
}

}  // namespace

const Table<ParameterObj>& ParameterObj::members() {
    static const Table<ParameterObj> table = {
        {L"Properties", 500,
         [](ParameterObj& s, Call& c) {
             c.put_object(make_properties(s.properties));
         },
         nullptr, true},
        {L"Name", 1, read_text(&ParameterObj::name),
         write_text(&ParameterObj::name)},
        {L"Value", DISPID_VALUE, read_var(&ParameterObj::value),
         write_var(&ParameterObj::value)},
        {L"Type", 2, read_long(&ParameterObj::type),
         write_long(&ParameterObj::type)},
        {L"Direction", 3, read_long(&ParameterObj::direction),
         write_long(&ParameterObj::direction, 1)},
        {L"Precision", 4,
         [](ParameterObj& s, Call& c) {
             c.put_byte(static_cast<unsigned char>(s.precision));
         },
         write_long(&ParameterObj::precision)},
        {L"NumericScale", 5,
         [](ParameterObj& s, Call& c) {
             c.put_byte(static_cast<unsigned char>(s.numeric_scale));
         },
         write_long(&ParameterObj::numeric_scale)},
        {L"Size", 6, read_long(&ParameterObj::size),
         write_long(&ParameterObj::size)},
        {L"AppendChunk", 7, append_chunk, nullptr},
        {L"Attributes", 8, read_long(&ParameterObj::attributes),
         write_long(&ParameterObj::attributes)},
    };
    return table;
}

// ----- Parameters
// -----------------------------------------------------------------------

ParametersObj::~ParametersObj() {
    for (auto* item : items) {
        item->Release();
    }
}

std::size_t ParametersObj::locate(const VARIANT& index) const {
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

std::vector<Value> ParametersObj::values() const {
    std::vector<Value> out;
    for (const auto* item : items) {
        out.push_back(to_value(item->value.ref()));
    }
    return out;
}

const Table<ParametersObj>& ParametersObj::members() {
    static const Table<ParametersObj> table = {
        {L"Count", 1,
         [](ParametersObj& s, Call& c) {
             c.put_long(static_cast<long>(s.items.size()));
         },
         nullptr},
        {L"_NewEnum", DISPID_NEWENUM,
         [](ParametersObj& s, Call& c) {
             std::vector<Var> list;
             for (auto* item : s.items) {
                 Var v;
                 v.set_object(item);
                 list.push_back(v);
             }
             put_enumerator(c, std::move(list));
         },
         nullptr},
        {L"Refresh", 2, [](ParametersObj&, Call&) {}, nullptr},
        {L"Append", 1610809344,
         [](ParametersObj& s, Call& c) {
             auto* parameter = unwrap<ParameterObj>(c.object(0));
             if (parameter == nullptr) {
                 throw ComError{3001, kBadArgumentText};
             }
             parameter->AddRef();
             s.items.push_back(parameter);
         },
         nullptr},
        {L"Delete", 1610809345,
         [](ParametersObj& s, Call& c) {
             if (c.count() < 1) {
                 throw ComError{450, "Wrong number of arguments"};
             }
             const std::size_t at = s.locate(c.args[0]);
             s.items[at]->Release();
             s.items.erase(s.items.begin() + static_cast<std::ptrdiff_t>(at));
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
    };
    return table;
}

// ----- Command
// ----------------------------------------------------------------------------

CommandObj::~CommandObj() {
    if (parameters != nullptr) {
        parameters->Release();
    }
    if (properties != nullptr) {
        properties->Release();
    }
}

ParametersObj* CommandObj::params() {
    if (parameters == nullptr) {
        parameters = new ParametersObj();
    }
    return parameters;
}

namespace {

// The values of a Parameters argument of Execute: an array of Variants.
std::vector<Value> argument_values(const VARIANT& variant) {
    std::vector<Value> values;
    const Var plain = dereferenced(variant);
    const VARIANT& v = plain.ref();
    if ((v.vt & VT_ARRAY) == 0 || v.parray == nullptr ||
        (v.vt & VT_TYPEMASK) != VT_VARIANT) {
        return values;
    }
    long low = 0;
    long high = -1;
    SafeArrayGetLBound(v.parray, 1, &low);
    SafeArrayGetUBound(v.parray, 1, &high);
    for (long i = low; i <= high; ++i) {
        VARIANT element{};
        VariantInit(&element);
        SafeArrayGetElement(v.parray, &i, &element);
        values.push_back(to_value(element));
        VariantClear(&element);
    }
    return values;
}

}  // namespace

const Table<CommandObj>& CommandObj::members() {
    static const Table<CommandObj> table = {
        {L"Properties", 500,
         [](CommandObj& s, Call& c) {
             c.put_object(make_properties(s.properties));
         },
         nullptr, true},
        {L"ActiveConnection", 1,
         [](CommandObj& s, Call& c) { c.put_object(s.connection.get()); },
         [](CommandObj& s, Call& c) {
             Ref keep;
             IUnknown* object = c.object(c.count() - 1);
             s.native = native_of(
                 object, object == nullptr ? &c.args.back() : nullptr, keep);
             s.connection = keep;
         }},
        {L"CommandText", 2, read_text(&CommandObj::text),
         write_text(&CommandObj::text)},
        {L"CommandTimeout", 3, read_long(&CommandObj::timeout),
         write_long(&CommandObj::timeout, 30)},
        {L"Prepared", 4, read_bool(&CommandObj::prepared),
         write_bool(&CommandObj::prepared)},
        {L"Execute", 5,
         [](CommandObj& s, Call& c) {
             if (s.native == nullptr) {
                 throw ComError{3709, kNoConnectionText};
             }
             std::vector<Value> values;
             if (c.has(1)) {
                 values = argument_values(c.args[1]);
             } else if (s.parameters != nullptr) {
                 values = s.parameters->values();
             }
             const long options = c.to_long(2, -1);
             long type = options > 0 ? options : s.type;
             std::string sql = s.text;
             if ((type & 4) != 0) {  // adCmdStoredProc
                 throw ComError{3251, kUnsupportedText};
             }
             if ((type & 2) != 0 || (type & 512) != 0) {  // a table
                 sql = "select * from " + sql;
             }
             s.native->clear_errors();
             auto outcome = s.native->execute(sql, values);
             if (!outcome.ok) {
                 fail_from(*s.native);
             }
             Var affected;
             affected.set_long(outcome.records_affected);
             c.write_back(0, affected);
             auto* set = make_recordset(std::move(outcome.recordset),
                                        s.connection, s.native, 1);
             set->source = sql;
             set->source_values = values;
             set->command = Ref(static_cast<IDispatch*>(&s));
             c.put_object(set);
             set->Release();
         },
         nullptr},
        {L"CreateParameter", 6,
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
        {L"Parameters", DISPID_VALUE,
         [](CommandObj& s, Call& c) { c.put_object(s.params()); }, nullptr,
         true},
        {L"CommandType", 7, read_long(&CommandObj::type),
         write_long(&CommandObj::type, 8)},
        {L"Name", 8, read_text(&CommandObj::name),
         write_text(&CommandObj::name)},
        {L"State", 9, [](CommandObj&, Call& c) { c.put_long(0); }, nullptr},
        {L"Cancel", 10, [](CommandObj&, Call&) {}, nullptr},
        {L"CommandStream", 11,
         [](CommandObj&, Call& c) { c.put_object(nullptr); },
         [](CommandObj&, Call&) {}},
        {L"Dialect", 12, read_text(&CommandObj::dialect),
         write_text(&CommandObj::dialect)},
        {L"NamedParameters", 13, read_bool(&CommandObj::named_parameters),
         write_bool(&CommandObj::named_parameters)},
    };
    return table;
}

}  // namespace wfc::adodb::com
