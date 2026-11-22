// ADODB.Recordset, Fields and Field as COM objects.

#include "objects.hpp"

#include <algorithm>
#include <cstdio>
#include <fstream>
#include <iterator>
#include <sstream>

namespace wfc::adodb::com {

namespace {

// The items of a list argument: a scalar is a list of one, an array its
// elements.
std::vector<Var> array_items(const VARIANT& variant) {
    std::vector<Var> items;
    const Var plain = dereferenced(variant);
    const VARIANT& v = plain.ref();
    if ((v.vt & VT_ARRAY) == 0 || v.parray == nullptr) {
        items.push_back(plain);
        return items;
    }
    long low = 0;
    long high = -1;
    SafeArrayGetLBound(v.parray, 1, &low);
    SafeArrayGetUBound(v.parray, 1, &high);
    for (long i = low; i <= high; ++i) {
        Var item;
        if ((v.vt & VT_TYPEMASK) == VT_VARIANT) {
            SafeArrayGetElement(v.parray, &i, item.get());
        } else {
            VARIANT element{};
            VariantInit(&element);
            element.vt = static_cast<VARTYPE>(v.vt & VT_TYPEMASK);
            SafeArrayGetElement(v.parray, &i, &element.lVal);
            VariantCopy(item.get(), &element);
        }
        items.push_back(item);
    }
    return items;
}

std::string trimmed(std::string text) {
    const auto first = text.find_first_not_of(" \t\r\n");
    if (first == std::string::npos) {
        return {};
    }
    const auto last = text.find_last_not_of(" \t\r\n");
    return text.substr(first, last - first + 1);
}

// Sets the fields named in FieldList (a name, or an array of names or
// positions) to Values on the current row.
void set_fields(RecordsetObj& s, const VARIANT& names, const VARIANT& values) {
    const auto keys = array_items(names);
    const auto data = array_items(values);
    if (keys.size() != data.size()) {
        throw ComError{3001, kBadArgumentText};
    }
    for (std::size_t i = 0; i < keys.size(); ++i) {
        const std::size_t at = field_index(*s.set, keys[i].ref());
        if (!s.set->set_value(at,
                              coerced(data[i].ref(), s.set->field(at).type))) {
            throw ComError{3001, kBadArgumentText};
        }
    }
}

// The ADO data types that have a fixed precision.
long default_precision(const int type) {
    switch (type) {
        case 2:
            return 5;
        case 3:
            return 10;
        case 4:
            return 7;
        case 5:
            return 15;
        case 6:
            return 19;
        case 20:
            return 19;
        case 16:
        case 17:
            return 3;
        default:
            return 255;
    }
}

}  // namespace

// ----- RecordsetObj
// ----------------------------------------------------------------------------

RecordsetObj::~RecordsetObj() {
    if (properties != nullptr) {
        properties->Release();
    }
}

void RecordsetObj::need_open() const {
    if (set->state() == adStateClosed) {
        throw ComError{3704, kClosedText};
    }
}

void RecordsetObj::need_record() const {
    need_open();
    if (set->bof() || set->eof()) {
        throw ComError{3021, kNoRecordText};
    }
}

void RecordsetObj::need_editable() const {
    need_open();
    if (!set->editable()) {
        throw ComError{3251, kReadOnlyText};
    }
}

void RecordsetObj::raise_pending() {
    Error error;
    if (set->take_error(error)) {
        throw ComError{error.number, error.description};
    }
}

void RecordsetObj::run(const std::string& sql,
                       const std::vector<Value>& values) {
    if (native == nullptr) {
        throw ComError{3709, kNoConnectionText};
    }
    native->clear_errors();
    auto outcome = lock_type > 1 ? native->execute_updatable(sql, values)
                                 : native->execute(sql, values);
    if (!outcome.ok) {
        fail_from(*native);
    }
    set = outcome.recordset != nullptr ? std::move(outcome.recordset)
                                       : Recordset::create();
    chunk_offsets.clear();
}

RecordsetObj* make_recordset(std::shared_ptr<Recordset> set,
                             const Ref& connection,
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

std::size_t field_index(const Recordset& set, const VARIANT& key) {
    const Value value = to_value(key);
    if (const auto* name = std::get_if<std::string>(&value)) {
        const int index = set.field_index(*name);
        if (index < 0) {
            throw ComError{3265, kNotFoundText};
        }
        return static_cast<std::size_t>(index);
    }
    const auto* number = std::get_if<std::int64_t>(&value);
    if (number == nullptr || *number < 0 ||
        static_cast<std::size_t>(*number) >= set.field_count()) {
        throw ComError{3265, kNotFoundText};
    }
    return static_cast<std::size_t>(*number);
}

// ----- Fields and Field
// --------------------------------------------------------------------------

FieldsObj::FieldsObj(RecordsetObj* owner_object) : owner(owner_object) {
    owner->AddRef();
}
FieldsObj::~FieldsObj() {
    owner->Release();
}

FieldObj::FieldObj(RecordsetObj* owner_object, const std::size_t at)
    : owner(owner_object), index(at) {
    owner->AddRef();
}
FieldObj::~FieldObj() {
    owner->Release();
}

namespace {

void read_value(FieldObj& s, Call& c) {
    s.owner->need_record();
    Var out;
    from_value(s.owner->set->value(s.index), s.owner->set->field(s.index).type,
               out);
    c.put_var(out);
}

void write_value(FieldObj& s, Call& c) {
    s.owner->need_editable();
    if (s.owner->set->bof() || s.owner->set->eof()) {
        throw ComError{3021, kNoRecordText};
    }
    if (!s.owner->set->set_value(
            s.index, coerced(c.value(), s.owner->set->field(s.index).type))) {
        throw ComError{3001, kBadArgumentText};
    }
    s.owner->raise_pending();
}

IDispatch* field_object(RecordsetObj* owner, const std::size_t index) {
    return new FieldObj(owner, index);
}

}  // namespace

const Table<FieldObj>& FieldObj::members() {
    static const Table<FieldObj> table = {
        {L"Properties", 500,
         [](FieldObj& s, Call& c) {
             c.put_object(make_properties(s.owner->properties));
         },
         nullptr, true},
        {L"ActualSize", 1109,
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
        {L"Attributes", 1114,
         [](FieldObj& s, Call& c) {
             const auto found = s.owner->field_attributes.find(s.index);
             if (found != s.owner->field_attributes.end()) {
                 c.put_long(found->second);
                 return;
             }
             long attributes = 0x40;  // adFldIsNullable
             if (s.owner->set->editable()) {
                 attributes |= 0x4;  // adFldUpdatable
             }
             const int type = s.owner->set->field(s.index).type;
             if (type == 203 || type == 205) {
                 attributes |= 0x80;  // adFldLong
             }
             c.put_long(attributes);
         },
         [](FieldObj& s, Call& c) {
             s.owner->field_attributes[s.index] = c.to_long(c.count() - 1, 0);
         }},
        {L"DefinedSize", 1103,
         [](FieldObj& s, Call& c) {
             c.put_long(s.owner->set->field(s.index).defined_size);
         },
         [](FieldObj&, Call&) { throw ComError{3251, kUnsupportedText}; }},
        {L"Name", 1100,
         [](FieldObj& s, Call& c) {
             c.put_text(s.owner->set->field(s.index).name);
         },
         nullptr},
        {L"Type", 1102,
         [](FieldObj& s, Call& c) {
             c.put_long(s.owner->set->field(s.index).type);
         },
         [](FieldObj&, Call&) { throw ComError{3251, kUnsupportedText}; }},
        {L"Value", DISPID_VALUE, read_value, write_value},
        {L"Precision", 1112,
         [](FieldObj& s, Call& c) {
             const auto found = s.owner->field_precision.find(s.index);
             c.put_byte(static_cast<unsigned char>(
                 found != s.owner->field_precision.end()
                     ? found->second
                     : default_precision(s.owner->set->field(s.index).type)));
         },
         [](FieldObj& s, Call& c) {
             s.owner->field_precision[s.index] = c.to_long(c.count() - 1, 255);
         }},
        {L"NumericScale", 1113,
         [](FieldObj& s, Call& c) {
             const auto found = s.owner->field_scale.find(s.index);
             c.put_byte(static_cast<unsigned char>(
                 found != s.owner->field_scale.end() ? found->second : 255));
         },
         [](FieldObj& s, Call& c) {
             s.owner->field_scale[s.index] = c.to_long(c.count() - 1, 255);
         }},
        {L"AppendChunk", 1107,
         [](FieldObj& s, Call& c) {
             s.owner->need_editable();
             if (s.owner->set->bof() || s.owner->set->eof() || c.count() < 1) {
                 throw ComError{3021, kNoRecordText};
             }
             Value current = s.owner->set->value(s.index);
             const Value more = to_value(c.args[0]);
             if (const auto* text = std::get_if<std::string>(&more)) {
                 const std::string before =
                     std::holds_alternative<std::string>(current)
                         ? std::get<std::string>(current)
                         : std::string{};
                 current = Value{before + *text};
             } else if (const auto* bytes = std::get_if<Blob>(&more)) {
                 Blob before = std::holds_alternative<Blob>(current)
                                   ? std::get<Blob>(current)
                                   : Blob{};
                 before.insert(before.end(), bytes->begin(), bytes->end());
                 current = Value{std::move(before)};
             } else {
                 throw ComError{3001, kBadArgumentText};
             }
             s.owner->set->set_value(s.index, std::move(current));
             s.owner->raise_pending();
         },
         nullptr},
        {L"GetChunk", 1108,
         [](FieldObj& s, Call& c) {
             s.owner->need_record();
             const long length = c.to_long(0, 0);
             const Value& v = s.owner->set->value(s.index);
             long& offset = s.owner->chunk_offsets[s.index];
             Var out;
             if (const auto* text = std::get_if<std::string>(&v)) {
                 const std::string part =
                     static_cast<std::size_t>(offset) < text->size()
                         ? text->substr(static_cast<std::size_t>(offset),
                                        static_cast<std::size_t>(length))
                         : std::string{};
                 offset += static_cast<long>(part.size());
                 out.set_text(part);
             } else if (const auto* blob = std::get_if<Blob>(&v)) {
                 Blob part;
                 if (static_cast<std::size_t>(offset) < blob->size()) {
                     const auto end = std::min<std::size_t>(
                         blob->size(), static_cast<std::size_t>(offset) +
                                           static_cast<std::size_t>(length));
                     part.assign(blob->begin() + offset, blob->begin() + end);
                 }
                 offset += static_cast<long>(part.size());
                 from_value(Value{std::move(part)}, 205, out);
             } else {
                 out.set_null();
             }
             c.put_var(out);
         },
         nullptr},
        {L"OriginalValue", 1104, read_value, nullptr},
        {L"UnderlyingValue", 1105, read_value, nullptr},
        {L"DataFormat", 1115, [](FieldObj&, Call& c) { c.put_object(nullptr); },
         [](FieldObj&, Call&) {}},
        {L"Status", 1116, [](FieldObj&, Call& c) { c.put_long(0); }, nullptr},
    };
    return table;
}

const Table<FieldsObj>& FieldsObj::members() {
    const auto append = [](FieldsObj& s, Call& c) {
        if (s.owner->set->state() != adStateClosed ||
            !s.owner->set->append_field(c.to_text(0), c.to_long(1, 0),
                                        c.to_long(2, 0))) {
            throw ComError{3705, kOpenText};
        }
        if (c.has(3) && c.to_long(3, -1) >= 0) {
            s.owner->field_attributes[s.owner->set->field_count() - 1] =
                c.to_long(3, 0);
        }
    };
    static const Table<FieldsObj> table = {
        {L"Count", 1,
         [](FieldsObj& s, Call& c) {
             c.put_long(static_cast<long>(s.owner->set->field_count()));
         },
         nullptr},
        {L"_NewEnum", DISPID_NEWENUM,
         [](FieldsObj& s, Call& c) {
             std::vector<Var> list;
             for (std::size_t i = 0; i < s.owner->set->field_count(); ++i) {
                 IDispatch* field = field_object(s.owner, i);
                 Var v;
                 v.set_object(field);
                 field->Release();
                 list.push_back(v);
             }
             put_enumerator(c, std::move(list));
         },
         nullptr},
        {L"Refresh", 2, [](FieldsObj&, Call&) {}, nullptr},
        {L"Item", DISPID_VALUE,
         [](FieldsObj& s, Call& c) {
             if (c.count() < 1) {
                 throw ComError{450, "Wrong number of arguments"};
             }
             IDispatch* field =
                 field_object(s.owner, field_index(*s.owner->set, c.args[0]));
             c.put_object(field);
             field->Release();
         },
         nullptr},
        {L"_Append", 1610874880, append, nullptr},
        {L"Delete", 4,
         [](FieldsObj&, Call&) { throw ComError{3251, kUnsupportedText}; },
         nullptr},
        {L"Append", 3, append, nullptr},
        {L"Update", 5,
         [](FieldsObj& s, Call&) {
             s.owner->need_editable();
             static_cast<void>(s.owner->set->update());
             s.owner->raise_pending();
         },
         nullptr},
        {L"Resync", 6, [](FieldsObj&, Call&) {}, nullptr},
        {L"CancelUpdate", 7,
         [](FieldsObj& s, Call&) {
             s.owner->need_editable();
             static_cast<void>(s.owner->set->cancel_update());
         },
         nullptr},
    };
    return table;
}

// ----- Recordset members
// ----------------------------------------------------------------------------------

namespace {

void open_recordset(RecordsetObj& s, Call& c) {
    if (s.set->state() != adStateClosed) {
        throw ComError{3705, kOpenText};
    }
    if (c.to_long(2, -1) >= 0) {
        s.cursor_type = c.to_long(2, 0);
    }
    if (c.to_long(3, -1) >= 0) {
        s.lock_type = c.to_long(3, s.lock_type);
    }
    const long options = c.to_long(4, -1);
    if (!c.has(0) && s.source.empty()) {
        if (!s.set->open_empty()) {
            throw ComError{3705, kOpenText};
        }
        return;
    }
    // The source: a Command, or the text of a query, a table or a file.
    std::vector<Value> values;
    std::string sql;
    if (c.has(0) && c.object(0) != nullptr) {
        auto* command = unwrap<CommandObj>(c.object(0));
        if (command == nullptr) {
            throw ComError{3001, kBadArgumentText};
        }
        sql = command->text;
        if (command->parameters != nullptr) {
            values = command->parameters->values();
        }
        s.command = Ref(static_cast<IDispatch*>(command));
        if (!command->connection.empty()) {
            s.connection = command->connection;
            s.native = command->native;
        }
    } else if (c.has(0)) {
        sql = trimmed(c.to_text(0));
    } else {
        sql = s.source;
    }
    if (c.has(1)) {
        Ref keep;
        auto native = native_of(
            c.object(1), c.object(1) == nullptr ? &c.args[1] : nullptr, keep);
        if (native != nullptr) {
            s.native = std::move(native);
            s.connection = keep;
        }
    }
    if (options > 0 && (options & 256) != 0) {  // adCmdFile
        std::ifstream file(to_wide(sql).c_str(), std::ios::binary);
        if (!file) {
            throw ComError{3709, "The file could not be opened."};
        }
        const std::string text((std::istreambuf_iterator<char>(file)),
                               std::istreambuf_iterator<char>());
        s.set = load_xml(text);
        s.source = sql;
        return;
    }
    if (options > 0 && (options & 4) != 0) {  // adCmdStoredProc
        throw ComError{3251, kUnsupportedText};
    }
    const bool table =
        options > 0 && ((options & 2) != 0 || (options & 512) != 0);
    if (table || (sql.find(' ') == std::string::npos && !sql.empty())) {
        sql = "select * from " + sql;
    }
    s.source = sql;
    s.source_values = values;
    s.run(sql, values);
}

// The value or the object for a Source or an ActiveConnection read.
void read_source(RecordsetObj& s, Call& c) {
    Var out;
    out.set_text(s.source);
    c.put_var(out);
}

void move_by(RecordsetObj& s, const long steps) {
    for (long i = 0; i < (steps < 0 ? -steps : steps); ++i) {
        if (steps > 0) {
            s.set->move_next();
        } else {
            s.set->move_previous();
        }
    }
}

void write_item(RecordsetObj& s, Call& c) {
    s.need_editable();
    if (c.count() < 2) {
        throw ComError{450, "Wrong number of arguments"};
    }
    if (s.set->bof() || s.set->eof()) {
        throw ComError{3021, kNoRecordText};
    }
    const std::size_t i = field_index(*s.set, c.args[0]);
    if (!s.set->set_value(i, coerced(c.value(), s.set->field(i).type))) {
        throw ComError{3001, kBadArgumentText};
    }
    s.raise_pending();
}

void read_item(RecordsetObj& s, Call& c) {
    s.need_record();
    if (c.count() < 1) {
        throw ComError{450, "Wrong number of arguments"};
    }
    const std::size_t i = field_index(*s.set, c.args[0]);
    Var out;
    from_value(s.set->value(i), s.set->field(i).type, out);
    c.put_var(out);
}

void make_clone(RecordsetObj& s, Call& c) {
    s.need_open();
    auto* copy =
        make_recordset(s.set->clone(), s.connection, s.native, s.lock_type);
    copy->source = s.source;
    copy->source_values = s.source_values;
    c.put_object(static_cast<IDispatch*>(copy));
    copy->Release();
}

void save_recordset(RecordsetObj& s, Call& c) {
    s.need_open();
    const std::string xml = save_xml(*s.set);
    if (!c.has(0)) {
        throw ComError{3001, kBadArgumentText};
    }
    IUnknown* destination = c.object(0);
    if (auto* stream = unwrap<StreamObj>(destination)) {
        // Written as text in the stream's character set (UTF-8 here).
        stream->data.assign(xml.begin(), xml.end());
        stream->position = static_cast<long>(xml.size());
        stream->state = 1;
        return;
    }
    const std::wstring path = to_wide(c.to_text(0));
    HANDLE file = CreateFileW(path.c_str(), GENERIC_WRITE, 0, nullptr,
                              CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (file == INVALID_HANDLE_VALUE) {
        throw ComError{3004, "The file could not be written."};
    }
    DWORD written = 0;
    WriteFile(file, xml.data(), static_cast<DWORD>(xml.size()), &written,
              nullptr);
    CloseHandle(file);
}

}  // namespace

const Table<RecordsetObj>& RecordsetObj::members() {
    static const Table<RecordsetObj> table = {
        {L"Properties", 500,
         [](RecordsetObj& s, Call& c) {
             c.put_object(make_properties(s.properties));
         },
         nullptr, true},
        {L"AbsolutePosition", 1000,
         [](RecordsetObj& s, Call& c) {
             s.need_open();
             c.put_long(s.set->absolute_position());
         },
         [](RecordsetObj& s, Call& c) {
             s.need_open();
             const long position = c.to_long(c.count() - 1, 1);
             s.set->move_first();
             move_by(s, position - 1);
             s.raise_pending();
         }},
        {L"ActiveConnection", 1001,
         [](RecordsetObj& s, Call& c) { c.put_object(s.connection.get()); },
         [](RecordsetObj& s, Call& c) {
             Ref keep;
             IUnknown* object = c.object(c.count() - 1);
             s.native = native_of(
                 object, object == nullptr ? &c.args.back() : nullptr, keep);
             s.connection = keep;
         }},
        {L"BOF", 1002,
         [](RecordsetObj& s, Call& c) {
             s.need_open();
             c.put_bool(s.set->bof());
         },
         nullptr},
        {L"Bookmark", 1003,
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
        {L"CacheSize", 1004, read_long(&RecordsetObj::cache_size),
         write_long(&RecordsetObj::cache_size, 1)},
        {L"CursorType", 1005, read_long(&RecordsetObj::cursor_type),
         write_long(&RecordsetObj::cursor_type)},
        {L"EOF", 1006,
         [](RecordsetObj& s, Call& c) {
             s.need_open();
             c.put_bool(s.set->eof());
         },
         nullptr},
        // rs("name") and rs.Fields("name"): the Fields collection is the
        // default member, and the arguments go on to its Item.
        {L"Fields", DISPID_VALUE,
         [](RecordsetObj& s, Call& c) {
             auto* fields = new FieldsObj(&s);
             c.put_object(static_cast<IDispatch*>(fields));
             fields->Release();
         },
         write_item, true},
        {L"LockType", 1008, read_long(&RecordsetObj::lock_type),
         write_long(&RecordsetObj::lock_type, 1)},
        {L"MaxRecords", 1009, read_long(&RecordsetObj::max_records),
         write_long(&RecordsetObj::max_records)},
        {L"RecordCount", 1010,
         [](RecordsetObj& s, Call& c) {
             s.need_open();
             c.put_long(s.set->record_count());
         },
         nullptr},
        {L"Source", 1011, read_source,
         [](RecordsetObj& s, Call& c) {
             IUnknown* object = c.object(c.count() - 1);
             if (auto* command = unwrap<CommandObj>(object)) {
                 s.command = Ref(static_cast<IDispatch*>(command));
                 s.source = command->text;
             } else {
                 s.source = c.to_text(c.count() - 1);
             }
         }},
        {L"AddNew", 1012,
         [](RecordsetObj& s, Call& c) {
             s.need_editable();
             if (!s.set->add_new()) {
                 throw ComError{3001, kBadArgumentText};
             }
             if (c.has(0) && c.has(1)) {
                 set_fields(s, c.args[0], c.args[1]);
                 static_cast<void>(s.set->update());
             }
             s.raise_pending();
         },
         nullptr},
        {L"CancelUpdate", 1013,
         [](RecordsetObj& s, Call&) {
             s.need_editable();
             static_cast<void>(s.set->cancel_update());
         },
         nullptr},
        {L"Close", 1014,
         [](RecordsetObj& s, Call&) {
             s.need_open();
             s.set->close();
             s.chunk_offsets.clear();
             s.raise_pending();
         },
         nullptr},
        {L"Delete", 1015,
         [](RecordsetObj& s, Call& c) {
             s.need_editable();
             s.need_record();
             const long affect = c.to_long(0, 1);
             if (affect == 2 || affect == 3) {
                 s.set->move_first();
                 while (s.set->record_count() > 0) {
                     if (!s.set->delete_current()) {
                         break;
                     }
                 }
                 s.raise_pending();
                 return;
             }
             if (!s.set->delete_current()) {
                 s.raise_pending();
                 throw ComError{3001, kBadArgumentText};
             }
         },
         nullptr},
        {L"GetRows", 1016,
         [](RecordsetObj& s, Call& c) {
             s.need_record();
             if (c.has(1) && !s.set->goto_bookmark(c.to_long(1, 0))) {
                 throw ComError{3001, kBadArgumentText};
             }
             std::vector<std::size_t> columns;
             if (c.has(2)) {
                 for (const auto& item : array_items(c.args[2])) {
                     columns.push_back(field_index(*s.set, item.ref()));
                 }
             } else {
                 for (std::size_t i = 0; i < s.set->field_count(); ++i) {
                     columns.push_back(i);
                 }
             }
             long rows = s.set->record_count() - s.set->absolute_position() + 1;
             const long asked = c.to_long(0, -1);
             if (asked >= 0 && asked < rows) {
                 rows = asked;
             }
             SAFEARRAYBOUND bounds[2] = {
                 {static_cast<ULONG>(columns.size()), 0},
                 {static_cast<ULONG>(rows), 0}};
             SAFEARRAY* array = SafeArrayCreate(VT_VARIANT, 2, bounds);
             if (array == nullptr) {
                 throw ComError{7, "Out of memory"};
             }
             for (long row = 0; row < rows; ++row) {
                 for (std::size_t k = 0; k < columns.size(); ++k) {
                     Var cell;
                     from_value(s.set->value(columns[k]),
                                s.set->field(columns[k]).type, cell);
                     long index[2] = {static_cast<long>(k), row};
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
        {L"Move", 1017,
         [](RecordsetObj& s, Call& c) {
             s.need_open();
             if (c.has(1)) {
                 const long start = c.to_long(1, 0);
                 if (start == 1) {
                     s.set->move_first();
                 } else if (start == 2) {
                     s.set->move_last();
                 } else if (start > 2 && !s.set->goto_bookmark(start)) {
                     throw ComError{3001, kBadArgumentText};
                 }
             }
             move_by(s, c.to_long(0, 0));
             s.raise_pending();
         },
         nullptr},
        {L"MoveNext", 1018,
         [](RecordsetObj& s, Call&) {
             s.need_open();
             const bool moved = s.set->move_next();
             s.raise_pending();
             if (!moved) {
                 throw ComError{3021, kNoRecordText};
             }
         },
         nullptr},
        {L"MovePrevious", 1019,
         [](RecordsetObj& s, Call&) {
             s.need_open();
             const bool moved = s.set->move_previous();
             s.raise_pending();
             if (!moved) {
                 throw ComError{3021, kNoRecordText};
             }
         },
         nullptr},
        {L"MoveFirst", 1020,
         [](RecordsetObj& s, Call&) {
             s.need_open();
             s.set->move_first();
             s.raise_pending();
         },
         nullptr},
        {L"MoveLast", 1021,
         [](RecordsetObj& s, Call&) {
             s.need_open();
             s.set->move_last();
             s.raise_pending();
         },
         nullptr},
        {L"Open", 1022, open_recordset, nullptr},
        {L"Requery", 1023,
         [](RecordsetObj& s, Call&) {
             if (s.source.empty()) {
                 throw ComError{3704, kClosedText};
             }
             if (s.set->state() != adStateClosed) {
                 s.set->close();
             }
             s.run(s.source, s.source_values);
         },
         nullptr},
        {L"_xResync", 1610809378, [](RecordsetObj&, Call&) {}, nullptr},
        {L"Update", 1025,
         [](RecordsetObj& s, Call& c) {
             s.need_editable();
             if (c.has(0) && c.has(1)) {
                 if (s.set->bof() || s.set->eof()) {
                     throw ComError{3021, kNoRecordText};
                 }
                 set_fields(s, c.args[0], c.args[1]);
             }
             static_cast<void>(s.set->update());
             s.raise_pending();
         },
         nullptr},
        {L"AbsolutePage", 1047,
         [](RecordsetObj& s, Call& c) {
             s.need_open();
             const long position = s.set->absolute_position();
             const long size = std::max<long>(1, s.page_size);
             c.put_long(position > 0 ? (position - 1) / size + 1 : position);
         },
         [](RecordsetObj& s, Call& c) {
             s.need_open();
             const long page = c.to_long(c.count() - 1, 1);
             const long size = std::max<long>(1, s.page_size);
             s.set->move_first();
             move_by(s, (page - 1) * size);
             s.raise_pending();
         }},
        {L"EditMode", 1026,
         [](RecordsetObj& s, Call& c) {
             s.need_open();
             c.put_long(s.set->edit_mode());
         },
         nullptr},
        {L"Filter", 1030,
         [](RecordsetObj& s, Call& c) {
             Var out;
             out.set_text(s.set->filter());
             c.put_var(out);
         },
         [](RecordsetObj& s, Call& c) {
             s.need_open();
             const Value given = to_value(c.value());
             std::string criteria;
             if (const auto* text = std::get_if<std::string>(&given)) {
                 criteria = *text;
             }  // a number (adFilterNone and the others) clears the filter
             if (!s.set->set_filter(criteria)) {
                 throw ComError{3001, kBadArgumentText};
             }
             s.raise_pending();
         }},
        {L"PageCount", 1050,
         [](RecordsetObj& s, Call& c) {
             s.need_open();
             const long size = std::max<long>(1, s.page_size);
             c.put_long((s.set->record_count() + size - 1) / size);
         },
         nullptr},
        {L"PageSize", 1048, read_long(&RecordsetObj::page_size),
         write_long(&RecordsetObj::page_size, 10)},
        {L"Sort", 1031,
         [](RecordsetObj& s, Call& c) { c.put_text(s.set->sort()); },
         [](RecordsetObj& s, Call& c) {
             s.need_open();
             if (!s.set->set_sort(c.to_text(c.count() - 1))) {
                 throw ComError{3265, kNotFoundText};
             }
             s.raise_pending();
         }},
        {L"Status", 1029,
         [](RecordsetObj& s, Call& c) {
             const int mode = s.set->edit_mode();
             c.put_long(mode == 2 ? 1 : (mode == 1 ? 2 : 0));
         },
         nullptr},
        {L"State", 1054,
         [](RecordsetObj& s, Call& c) { c.put_long(s.set->state()); }, nullptr},
        {L"_xClone", 1610809392, make_clone, nullptr},
        {L"UpdateBatch", 1035,
         [](RecordsetObj& s, Call&) {
             s.need_open();
             static_cast<void>(s.set->update());
             s.raise_pending();
         },
         nullptr},
        {L"CancelBatch", 1049,
         [](RecordsetObj& s, Call&) {
             s.need_open();
             static_cast<void>(s.set->cancel_update());
         },
         nullptr},
        {L"CursorLocation", 1051, read_long(&RecordsetObj::cursor_location),
         write_long(&RecordsetObj::cursor_location, 2)},
        {L"NextRecordset", 1052,
         [](RecordsetObj& s, Call& c) {
             s.need_open();
             Var none;
             none.set_long(0);
             c.write_back(0, none);
             c.put_object(nullptr);
         },
         nullptr},
        {L"Supports", 1036,
         [](RecordsetObj& s, Call& c) {
             const long wanted = c.to_long(0, 0);
             const long always = 0x100 | 0x200 | 0x2000 | 0x4000 | 0x80000;
             long offered = always;
             if (s.set->editable()) {
                 offered |= 0x1000400 | 0x1000800 | 0x1008000 | 0x10000;
             }
             c.put_bool((wanted & ~offered) == 0 ||
                        ((wanted & 0xFFFF0000) != 0 &&
                         (wanted & ~(offered | 0x1000000)) == 0));
         },
         nullptr},
        {L"Collect", -8, read_item, write_item},
        {L"MarshalOptions", 1053, read_long(&RecordsetObj::marshal_options),
         write_long(&RecordsetObj::marshal_options)},
        {L"Find", 1058,
         [](RecordsetObj& s, Call& c) {
             s.need_open();
             if (c.has(3) && !s.set->goto_bookmark(c.to_long(3, 0))) {
                 throw ComError{3001, kBadArgumentText};
             }
             if (!s.set->find(c.to_text(0), c.to_long(1, 0),
                              c.to_long(2, 1) != -1)) {
                 throw ComError{3001, kBadArgumentText};
             }
             s.raise_pending();
         },
         nullptr},
        {L"Cancel", 1055, [](RecordsetObj&, Call&) {}, nullptr},
        {L"DataSource", 1056,
         [](RecordsetObj&, Call& c) { c.put_object(nullptr); },
         [](RecordsetObj&, Call&) {}},
        {L"_xSave", 1610874883, save_recordset, nullptr},
        {L"ActiveCommand", 1061,
         [](RecordsetObj& s, Call& c) { c.put_object(s.command.get()); },
         nullptr},
        {L"StayInSync", 1063, read_bool(&RecordsetObj::stay_in_sync),
         write_bool(&RecordsetObj::stay_in_sync)},
        {L"GetString", 1062,
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
        {L"DataMember", 1064, read_text(&RecordsetObj::data_member),
         write_text(&RecordsetObj::data_member)},
        {L"CompareBookmarks", 1065,
         [](RecordsetObj&, Call& c) {
             const long a = c.to_long(0, 0);
             const long b = c.to_long(1, 0);
             c.put_long(a == b ? 1 : (a < b ? 0 : 2));
         },
         nullptr},
        {L"Clone", 1034, make_clone, nullptr},
        {L"Resync", 1024, [](RecordsetObj&, Call&) {}, nullptr},
        {L"Seek", 1066,
         [](RecordsetObj&, Call&) { throw ComError{3251, kUnsupportedText}; },
         nullptr},
        {L"Index", 1067, read_text(&RecordsetObj::index),
         write_text(&RecordsetObj::index)},
        {L"Save", 1057, save_recordset, nullptr},
    };
    return table;
}

}  // namespace wfc::adodb::com
