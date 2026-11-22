// ADODB.Stream and ADODB.Record as COM objects.

#include "objects.hpp"

#include <algorithm>

namespace wfc::adodb::com {

namespace {

// ----- character sets
// -------------------------------------------------------------------

struct Encoding {
    bool utf16{false};
    UINT code_page{CP_ACP};
    bool bom{false};
};

Encoding encoding_of(const std::string& charset) {
    std::string name;
    for (const char c : charset) {
        name.push_back(
            static_cast<char>(std::tolower(static_cast<unsigned char>(c))));
    }
    Encoding result;
    if (name == "unicode" || name == "utf-16" || name == "utf-16le" ||
        name == "ucs-2") {
        result.utf16 = true;
        result.bom = true;
    } else if (name == "utf-8" || name == "utf8") {
        result.code_page = CP_UTF8;
        result.bom = true;
    } else if (name == "us-ascii" || name == "ascii") {
        result.code_page = 20127;
    } else if (name == "iso-8859-1" || name == "latin1") {
        result.code_page = 28591;
    } else if (name == "windows-1252") {
        result.code_page = 1252;
    }
    return result;
}

Blob encode(const std::string& utf8, const Encoding& encoding,
            const bool with_bom) {
    Blob out;
    const std::wstring wide = to_wide(utf8);
    if (encoding.utf16) {
        if (with_bom) {
            out.push_back(0xFF);
            out.push_back(0xFE);
        }
        const auto* bytes = reinterpret_cast<const unsigned char*>(wide.data());
        out.insert(out.end(), bytes, bytes + wide.size() * sizeof(wchar_t));
        return out;
    }
    if (with_bom && encoding.code_page == CP_UTF8) {
        out.push_back(0xEF);
        out.push_back(0xBB);
        out.push_back(0xBF);
    }
    if (wide.empty()) {
        return out;
    }
    const int size = WideCharToMultiByte(encoding.code_page, 0, wide.data(),
                                         static_cast<int>(wide.size()), nullptr,
                                         0, nullptr, nullptr);
    const std::size_t at = out.size();
    out.resize(at + static_cast<std::size_t>(size));
    WideCharToMultiByte(
        encoding.code_page, 0, wide.data(), static_cast<int>(wide.size()),
        reinterpret_cast<char*>(out.data() + at), size, nullptr, nullptr);
    return out;
}

// The text of data[from, to) as UTF-8, a byte-order mark at the start dropped.
std::string decode(const Blob& data, std::size_t from, const std::size_t to,
                   const Encoding& encoding) {
    if (from == 0 && encoding.utf16 && to >= 2 && data[0] == 0xFF &&
        data[1] == 0xFE) {
        from = 2;
    }
    if (from == 0 && !encoding.utf16 && encoding.code_page == CP_UTF8 &&
        to >= 3 && data[0] == 0xEF && data[1] == 0xBB && data[2] == 0xBF) {
        from = 3;
    }
    if (from >= to) {
        return {};
    }
    if (encoding.utf16) {
        const auto* wide = reinterpret_cast<const wchar_t*>(data.data() + from);
        return to_utf8(wide, static_cast<int>((to - from) / sizeof(wchar_t)));
    }
    const int size =
        MultiByteToWideChar(encoding.code_page, 0,
                            reinterpret_cast<const char*>(data.data() + from),
                            static_cast<int>(to - from), nullptr, 0);
    std::wstring wide(static_cast<std::size_t>(size), L'\0');
    MultiByteToWideChar(encoding.code_page, 0,
                        reinterpret_cast<const char*>(data.data() + from),
                        static_cast<int>(to - from), wide.data(), size);
    return to_utf8(wide.c_str(), static_cast<int>(wide.size()));
}

void need_open(const StreamObj& s) {
    if (s.state == 0) {
        throw ComError{3704, kClosedText};
    }
}

std::string separator_of(const StreamObj& s) {
    return s.line_separator == 10 ? "\n"
                                  : (s.line_separator == 13 ? "\r" : "\r\n");
}

void write_bytes(StreamObj& s, const Blob& bytes) {
    const auto at = static_cast<std::size_t>(s.position);
    if (s.data.size() < at) {
        s.data.resize(at);
    }
    if (s.data.size() < at + bytes.size()) {
        s.data.resize(at + bytes.size());
    }
    std::copy(bytes.begin(), bytes.end(),
              s.data.begin() + static_cast<std::ptrdiff_t>(at));
    s.position = static_cast<long>(at + bytes.size());
}

bool read_file(const std::wstring& path, Blob& out) {
    HANDLE file =
        CreateFileW(path.c_str(), GENERIC_READ, FILE_SHARE_READ, nullptr,
                    OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (file == INVALID_HANDLE_VALUE) {
        return false;
    }
    LARGE_INTEGER size{};
    GetFileSizeEx(file, &size);
    out.assign(static_cast<std::size_t>(size.QuadPart), 0);
    DWORD read = 0;
    if (!out.empty()) {
        ReadFile(file, out.data(), static_cast<DWORD>(out.size()), &read,
                 nullptr);
    }
    CloseHandle(file);
    return true;
}

}  // namespace

const Table<StreamObj>& StreamObj::members() {
    static const Table<StreamObj> table = {
        {L"Size", 1,
         [](StreamObj& s, Call& c) {
             need_open(s);
             c.put_long(static_cast<long>(s.data.size()));
         },
         nullptr},
        {L"EOS", 2,
         [](StreamObj& s, Call& c) {
             need_open(s);
             c.put_bool(static_cast<std::size_t>(s.position) >= s.data.size());
         },
         nullptr},
        {L"Position", 3,
         [](StreamObj& s, Call& c) {
             need_open(s);
             c.put_long(s.position);
         },
         [](StreamObj& s, Call& c) {
             need_open(s);
             const long position = c.to_long(c.count() - 1, 0);
             if (position < 0) {
                 throw ComError{3001, kBadArgumentText};
             }
             s.position = position;
         }},
        {L"Type", 4, read_long(&StreamObj::type),
         [](StreamObj& s, Call& c) {
             const long type = c.to_long(c.count() - 1, 2);
             if ((type != 1 && type != 2) ||
                 (s.state != 0 && s.position != 0 && !s.data.empty())) {
                 throw ComError{3001, kBadArgumentText};
             }
             s.type = type;
         }},
        {L"LineSeparator", 5, read_long(&StreamObj::line_separator),
         write_long(&StreamObj::line_separator, -1)},
        {L"State", 6, read_long(&StreamObj::state), nullptr},
        {L"Mode", 7, read_long(&StreamObj::mode), write_long(&StreamObj::mode)},
        {L"Charset", 8, read_text(&StreamObj::charset),
         write_text(&StreamObj::charset)},
        {L"Read", 9,
         [](StreamObj& s, Call& c) {
             need_open(s);
             const long wanted = c.to_long(0, -1);
             const auto from = static_cast<std::size_t>(s.position);
             std::size_t to = s.data.size();
             if (wanted >= 0) {
                 to = std::min(to, from + static_cast<std::size_t>(wanted));
             }
             Blob bytes;
             if (from < to) {
                 bytes.assign(
                     s.data.begin() + static_cast<std::ptrdiff_t>(from),
                     s.data.begin() + static_cast<std::ptrdiff_t>(to));
             }
             s.position = static_cast<long>(std::max(from, to));
             Var out;
             if (bytes.empty()) {
                 out.set_null();
             } else {
                 from_value(Value{std::move(bytes)}, 205, out);
             }
             c.put_var(out);
         },
         nullptr},
        {L"Open", 10,
         [](StreamObj& s, Call& c) {
             if (s.state != 0) {
                 throw ComError{3705, kOpenText};
             }
             s.data.clear();
             s.position = 0;
             if (c.has(1)) {
                 s.mode = c.to_long(1, 0);
             }
             if (c.has(0)) {
                 const std::string path = c.to_text(0);
                 if (!read_file(to_wide(path), s.data)) {
                     throw ComError{3004, "The file could not be opened."};
                 }
             }
             s.state = 1;
         },
         nullptr},
        {L"Close", 11,
         [](StreamObj& s, Call&) {
             need_open(s);
             s.data.clear();
             s.position = 0;
             s.state = 0;
         },
         nullptr},
        {L"SkipLine", 12,
         [](StreamObj& s, Call&) {
             need_open(s);
             const Encoding encoding = encoding_of(s.charset);
             const auto from = static_cast<std::size_t>(s.position);
             const std::string rest =
                 decode(s.data, from, s.data.size(), encoding);
             const std::string separator = separator_of(s);
             const auto found = rest.find(separator);
             const std::string skipped =
                 found == std::string::npos
                     ? rest
                     : rest.substr(0, found + separator.size());
             s.position = static_cast<long>(std::min(
                 s.data.size(),
                 from + encode(skipped, encoding, false).size() +
                     (from == 0 && encoding.bom ? (encoding.utf16 ? 2U : 3U)
                                                : 0U)));
         },
         nullptr},
        {L"Write", 13,
         [](StreamObj& s, Call& c) {
             need_open(s);
             if (c.count() < 1) {
                 throw ComError{450, "Wrong number of arguments"};
             }
             const Value given = to_value(c.args[0]);
             const auto* bytes = std::get_if<Blob>(&given);
             if (bytes == nullptr) {
                 throw ComError{3001, kBadArgumentText};
             }
             write_bytes(s, *bytes);
         },
         nullptr},
        {L"SetEOS", 14,
         [](StreamObj& s, Call&) {
             need_open(s);
             s.data.resize(static_cast<std::size_t>(s.position));
         },
         nullptr},
        {L"CopyTo", 15,
         [](StreamObj& s, Call& c) {
             need_open(s);
             auto* target = unwrap<StreamObj>(c.object(0));
             if (target == nullptr) {
                 throw ComError{3001, kBadArgumentText};
             }
             need_open(*target);
             const auto from = static_cast<std::size_t>(s.position);
             Blob bytes;
             if (from < s.data.size()) {
                 bytes.assign(
                     s.data.begin() + static_cast<std::ptrdiff_t>(from),
                     s.data.end());
             }
             write_bytes(*target, bytes);
             s.position = static_cast<long>(s.data.size());
         },
         nullptr},
        {L"Flush", 16, [](StreamObj&, Call&) {}, nullptr},
        {L"SaveToFile", 17,
         [](StreamObj& s, Call& c) {
             need_open(s);
             const std::wstring path = to_wide(c.to_text(0));
             const long options = c.to_long(1, 1);
             HANDLE file =
                 CreateFileW(path.c_str(), GENERIC_WRITE, 0, nullptr,
                             options == 2 ? CREATE_ALWAYS : CREATE_NEW,
                             FILE_ATTRIBUTE_NORMAL, nullptr);
             if (file == INVALID_HANDLE_VALUE) {
                 throw ComError{3004, "The file could not be written."};
             }
             DWORD written = 0;
             if (!s.data.empty()) {
                 WriteFile(file, s.data.data(),
                           static_cast<DWORD>(s.data.size()), &written,
                           nullptr);
             }
             CloseHandle(file);
         },
         nullptr},
        {L"LoadFromFile", 18,
         [](StreamObj& s, Call& c) {
             need_open(s);
             Blob bytes;
             if (!read_file(to_wide(c.to_text(0)), bytes)) {
                 throw ComError{3004, "The file could not be opened."};
             }
             s.data = std::move(bytes);
             s.position = 0;
         },
         nullptr},
        {L"ReadText", 19,
         [](StreamObj& s, Call& c) {
             need_open(s);
             const Encoding encoding = encoding_of(s.charset);
             const long count = c.to_long(0, -1);
             const auto from = static_cast<std::size_t>(s.position);
             const std::string rest =
                 decode(s.data, from, s.data.size(), encoding);
             const std::size_t bom =
                 from == 0 && encoding.bom && !s.data.empty() &&
                         ((encoding.utf16 && s.data.size() >= 2 &&
                           s.data[0] == 0xFF && s.data[1] == 0xFE) ||
                          (!encoding.utf16 && s.data.size() >= 3 &&
                           s.data[0] == 0xEF && s.data[1] == 0xBB))
                     ? (encoding.utf16 ? 2U : 3U)
                     : 0U;
             std::string taken = rest;
             if (count == -2) {  // adReadLine
                 const std::string separator = separator_of(s);
                 const auto found = rest.find(separator);
                 taken =
                     found == std::string::npos ? rest : rest.substr(0, found);
                 const std::size_t used = found == std::string::npos
                                              ? rest.size()
                                              : found + separator.size();
                 s.position = static_cast<long>(
                     from + bom +
                     encode(rest.substr(0, used), encoding, false).size());
             } else if (count >= 0) {
                 const std::wstring wide = to_wide(rest);
                 const std::wstring head =
                     wide.substr(0, static_cast<std::size_t>(count));
                 taken = to_utf8(head.c_str(), static_cast<int>(head.size()));
                 s.position = static_cast<long>(
                     from + bom + encode(taken, encoding, false).size());
             } else {
                 s.position = static_cast<long>(s.data.size());
             }
             if (s.position > static_cast<long>(s.data.size())) {
                 s.position = static_cast<long>(s.data.size());
             }
             c.put_text(taken);
         },
         nullptr},
        {L"WriteText", 20,
         [](StreamObj& s, Call& c) {
             need_open(s);
             const Encoding encoding = encoding_of(s.charset);
             std::string text = c.to_text(0);
             if (c.to_long(1, 0) == 1) {  // adWriteLine
                 text += separator_of(s);
             }
             const bool first = s.data.empty() && s.position == 0;
             write_bytes(s, encode(text, encoding, first && encoding.bom));
         },
         nullptr},
        {L"Cancel", 21, [](StreamObj&, Call&) {}, nullptr},
    };
    return table;
}

// ----- Record
// -------------------------------------------------------------------------

RecordObj::~RecordObj() {
    if (holder != nullptr) {
        holder->Release();
    }
    if (properties != nullptr) {
        properties->Release();
    }
}

namespace {

void need_record_open(const RecordObj& s) {
    if (s.state == 0 || s.holder == nullptr) {
        throw ComError{3704, kClosedText};
    }
}

}  // namespace

const Table<RecordObj>& RecordObj::members() {
    static const Table<RecordObj> table = {
        {L"Properties", 500,
         [](RecordObj& s, Call& c) {
             c.put_object(make_properties(s.properties));
         },
         nullptr, true},
        {L"ActiveConnection", 1,
         [](RecordObj& s, Call& c) { c.put_object(s.connection.get()); },
         [](RecordObj& s, Call& c) {
             Ref keep;
             IUnknown* object = c.object(c.count() - 1);
             native_of(object, object == nullptr ? &c.args.back() : nullptr,
                       keep);
             s.connection = keep;
         }},
        {L"State", 2, read_long(&RecordObj::state), nullptr},
        {L"Source", 3, read_text(&RecordObj::source),
         write_text(&RecordObj::source)},
        {L"Mode", 4, read_long(&RecordObj::mode), write_long(&RecordObj::mode)},
        {L"ParentURL", 5, [](RecordObj&, Call& c) { c.put_null(); }, nullptr},
        {L"MoveRecord", 6,
         [](RecordObj&, Call&) { throw ComError{3251, kUnsupportedText}; },
         nullptr},
        {L"CopyRecord", 7,
         [](RecordObj&, Call&) { throw ComError{3251, kUnsupportedText}; },
         nullptr},
        {L"DeleteRecord", 8,
         [](RecordObj&, Call&) { throw ComError{3251, kUnsupportedText}; },
         nullptr},
        {L"Open", 9,
         [](RecordObj& s, Call& c) {
             if (s.state != 0) {
                 throw ComError{3705, kOpenText};
             }
             if (c.has(1)) {
                 Ref keep;
                 native_of(c.object(1),
                           c.object(1) == nullptr ? &c.args[1] : nullptr, keep);
                 s.connection = keep;
             }
             std::shared_ptr<Recordset> row = Recordset::create();
             IUnknown* object = c.has(0) ? c.object(0) : nullptr;
             if (auto* source = unwrap<RecordsetObj>(object)) {
                 // The current row of a recordset.
                 source->need_record();
                 for (std::size_t i = 0; i < source->set->field_count(); ++i) {
                     row->append_field(source->set->field(i).name,
                                       source->set->field(i).type,
                                       source->set->field(i).defined_size);
                 }
                 row->open_empty();
                 row->add_new();
                 for (std::size_t i = 0; i < source->set->field_count(); ++i) {
                     row->set_value(i, source->set->value(i));
                 }
                 row->update();
                 row->move_first();
             } else if (c.has(0) && object == nullptr) {
                 // A query on the connection: its first row.
                 auto* link = unwrap<ConnectionObj>(s.connection.get());
                 if (link == nullptr) {
                     throw ComError{3709, kNoConnectionText};
                 }
                 s.source = c.to_text(0);
                 link->native->clear_errors();
                 auto outcome = link->native->execute(s.source);
                 if (!outcome.ok) {
                     fail_from(*link->native);
                 }
                 if (outcome.recordset == nullptr ||
                     outcome.recordset->record_count() == 0) {
                     throw ComError{3021, kNoRecordText};
                 }
                 auto& found = *outcome.recordset;
                 for (std::size_t i = 0; i < found.field_count(); ++i) {
                     row->append_field(found.field(i).name, found.field(i).type,
                                       found.field(i).defined_size);
                 }
                 row->open_empty();
                 row->add_new();
                 for (std::size_t i = 0; i < found.field_count(); ++i) {
                     row->set_value(i, found.value(i));
                 }
                 row->update();
                 row->move_first();
             } else {
                 throw ComError{3251, kUnsupportedText};
             }
             s.holder = new RecordsetObj();
             s.holder->set = std::move(row);
             s.mode = c.to_long(2, 0);
             s.state = 1;
         },
         nullptr},
        {L"Close", 10,
         [](RecordObj& s, Call&) {
             need_record_open(s);
             s.holder->Release();
             s.holder = nullptr;
             s.state = 0;
         },
         nullptr},
        {L"Fields", DISPID_VALUE,
         [](RecordObj& s, Call& c) {
             need_record_open(s);
             auto* fields = new FieldsObj(s.holder);
             c.put_object(static_cast<IDispatch*>(fields));
             fields->Release();
         },
         nullptr, true},
        {L"RecordType", 11, [](RecordObj&, Call& c) { c.put_long(0); },
         nullptr},
        {L"GetChildren", 12,
         [](RecordObj&, Call&) { throw ComError{3251, kUnsupportedText}; },
         nullptr},
        {L"Cancel", 13, [](RecordObj&, Call&) {}, nullptr},
    };
    return table;
}

}  // namespace wfc::adodb::com
