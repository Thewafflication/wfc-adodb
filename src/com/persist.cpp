// Saving and loading a Recordset in ADO's XML persistence format: the schema
// of the columns and the rows as attributes of z:row elements. adPersistADTG
// (the binary format) is written as this XML too, and the loader reads the
// XML only.

#include "objects.hpp"

#include <cstdio>
#include <cstdlib>
#include <map>
#include <string>
#include <vector>

namespace wfc::adodb::com {

namespace {

// ----- dates
// ---------------------------------------------------------------------------------

// Days from the Unix epoch to a civil date, and back.
long days_from_civil(long y, const long m, const long d) {
    y -= m <= 2 ? 1 : 0;
    const long era = (y >= 0 ? y : y - 399) / 400;
    const long yoe = y - era * 400;
    const long doy = (153 * (m + (m > 2 ? -3 : 9)) + 2) / 5 + d - 1;
    const long doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    return era * 146097 + doe - 719468;
}

void civil_from_days(long z, long& y, long& m, long& d) {
    z += 719468;
    const long era = (z >= 0 ? z : z - 146096) / 146097;
    const long doe = z - era * 146097;
    const long yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
    y = yoe + era * 400;
    const long doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    const long mp = (5 * doy + 2) / 153;
    d = doy - (153 * mp + 2) / 5 + 1;
    m = mp + (mp < 10 ? 3 : -9);
    y += m <= 2 ? 1 : 0;
}

std::string date_text(const double serial) {
    const long whole = static_cast<long>(std::floor(serial));
    long seconds = std::lround((serial - static_cast<double>(whole)) * 86400.0);
    long days = whole;
    if (seconds >= 86400) {
        seconds -= 86400;
        ++days;
    }
    long y = 0;
    long m = 0;
    long d = 0;
    civil_from_days(days - 25569, y, m, d);
    char buffer[40];
    std::snprintf(buffer, sizeof(buffer), "%04ld-%02ld-%02ldT%02ld:%02ld:%02ld",
                  y, m, d, seconds / 3600, (seconds / 60) % 60, seconds % 60);
    return buffer;
}

bool parse_date(const std::string& text, double& serial) {
    long y = 0;
    long m = 0;
    long d = 0;
    long hh = 0;
    long mm = 0;
    long ss = 0;
    const char* p = text.c_str();
    char* end = nullptr;
    y = std::strtol(p, &end, 10);
    if (*end != '-') {
        return false;
    }
    m = std::strtol(end + 1, &end, 10);
    if (*end != '-') {
        return false;
    }
    d = std::strtol(end + 1, &end, 10);
    if (*end == 'T' || *end == ' ') {
        hh = std::strtol(end + 1, &end, 10);
        if (*end == ':') {
            mm = std::strtol(end + 1, &end, 10);
        }
        if (*end == ':') {
            ss = std::strtol(end + 1, &end, 10);
        }
    }
    serial = static_cast<double>(days_from_civil(y, m, d) + 25569) +
             static_cast<double>(hh * 3600 + mm * 60 + ss) / 86400.0;
    return true;
}

// ----- XML text
// ---------------------------------------------------------------------------

std::string escape(const std::string& text) {
    std::string out;
    for (const unsigned char c : text) {
        switch (c) {
            case '&':
                out += "&amp;";
                break;
            case '<':
                out += "&lt;";
                break;
            case '>':
                out += "&gt;";
                break;
            case '\'':
                out += "&apos;";
                break;
            case '"':
                out += "&quot;";
                break;
            default:
                if (c < 0x20 && c != '\t') {
                    char buffer[16];
                    std::snprintf(buffer, sizeof(buffer), "&#x%X;", c);
                    out += buffer;
                } else {
                    out.push_back(static_cast<char>(c));
                }
        }
    }
    return out;
}

std::string unescape(const std::string& text) {
    std::string out;
    for (std::size_t i = 0; i < text.size(); ++i) {
        if (text[i] != '&') {
            out.push_back(text[i]);
            continue;
        }
        const auto semicolon = text.find(';', i);
        if (semicolon == std::string::npos) {
            out.push_back(text[i]);
            continue;
        }
        const std::string entity = text.substr(i + 1, semicolon - i - 1);
        i = semicolon;
        if (entity == "amp") {
            out.push_back('&');
        } else if (entity == "lt") {
            out.push_back('<');
        } else if (entity == "gt") {
            out.push_back('>');
        } else if (entity == "apos") {
            out.push_back('\'');
        } else if (entity == "quot") {
            out.push_back('"');
        } else if (!entity.empty() && entity[0] == '#') {
            const long code =
                entity.size() > 1 && (entity[1] == 'x' || entity[1] == 'X')
                    ? std::strtol(entity.c_str() + 2, nullptr, 16)
                    : std::strtol(entity.c_str() + 1, nullptr, 10);
            const wchar_t wide[2] = {static_cast<wchar_t>(code), L'\0'};
            out += to_utf8(wide, 1);
        }
    }
    return out;
}

// A column name as an XML name: what is not a letter, digit or underscore
// becomes _xHHHH_.
std::string xml_name(const std::string& name) {
    std::string out;
    const std::wstring wide = to_wide(name);
    for (const wchar_t c : wide) {
        if ((c >= L'a' && c <= L'z') || (c >= L'A' && c <= L'Z') ||
            (c >= L'0' && c <= L'9' && !out.empty()) || c == L'_') {
            out.push_back(static_cast<char>(c));
        } else {
            char buffer[16];
            std::snprintf(buffer, sizeof(buffer), "_x%04X_",
                          static_cast<unsigned>(c));
            out += buffer;
        }
    }
    return out;
}

std::string plain_name(const std::string& name) {
    std::string out;
    for (std::size_t i = 0; i < name.size(); ++i) {
        if (name[i] == '_' && i + 6 < name.size() + 0 && name[i + 1] == 'x' &&
            name[i + 6] == '_') {
            const long code =
                std::strtol(name.substr(i + 2, 4).c_str(), nullptr, 16);
            const wchar_t wide[2] = {static_cast<wchar_t>(code), L'\0'};
            out += to_utf8(wide, 1);
            i += 6;
        } else {
            out.push_back(name[i]);
        }
    }
    return out;
}

struct TypeInfo {
    const char* name;
    int length;
    bool long_text;
};

TypeInfo xml_type(const int type, const long size) {
    switch (type) {
        case 2:
            return {"i2", 2, false};
        case 3:
            return {"int", 4, false};
        case 4:
            return {"r4", 4, false};
        case 5:
            return {"float", 8, false};
        case 6:
            return {"i8", 8, false};
        case 7:
        case 133:
        case 135:
            return {"dateTime", 16, false};
        case 11:
            return {"boolean", 2, false};
        case 16:
        case 17:
            return {"ui1", 1, false};
        case 20:
            return {"i8", 8, false};
        case 72:
            return {"uuid", 16, false};
        case 128:
        case 204:
            return {"bin.hex", static_cast<int>(size > 0 ? size : 255), false};
        case 205:
            return {"bin.hex", 2147483647, true};
        case 201:
        case 203:
            return {"string", 2147483647, true};
        default:
            return {"string", static_cast<int>(size > 0 ? size : 255), false};
    }
}

std::string hex_of(const Blob& bytes) {
    static const char digits[] = "0123456789abcdef";
    std::string out;
    for (const unsigned char b : bytes) {
        out.push_back(digits[b >> 4]);
        out.push_back(digits[b & 15]);
    }
    return out;
}

Blob bytes_of(const std::string& hex) {
    Blob out;
    for (std::size_t i = 0; i + 1 < hex.size(); i += 2) {
        out.push_back(static_cast<unsigned char>(
            std::strtol(hex.substr(i, 2).c_str(), nullptr, 16)));
    }
    return out;
}

std::string value_text(const Value& value, const int type) {
    if (const auto* number = std::get_if<std::int64_t>(&value)) {
        if (type == 11) {
            return *number != 0 ? "true" : "false";
        }
        return std::to_string(*number);
    }
    if (const auto* real = std::get_if<double>(&value)) {
        if (type == 7 || type == 133 || type == 135) {
            return date_text(*real);
        }
        char buffer[40];
        std::snprintf(buffer, sizeof(buffer), "%.17g", *real);
        return buffer;
    }
    if (const auto* text = std::get_if<std::string>(&value)) {
        return *text;
    }
    if (const auto* blob = std::get_if<Blob>(&value)) {
        return hex_of(*blob);
    }
    return {};
}

}  // namespace

std::string save_xml(const Recordset& set) {
    std::string out =
        "<xml xmlns:s='uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882'\n"
        "\txmlns:dt='uuid:C2F41010-65B3-11d1-A29F-00AA00C14882'\n"
        "\txmlns:rs='urn:schemas-microsoft-com:rowset'\n"
        "\txmlns:z='#RowsetSchema'>\n"
        "<s:Schema id='RowsetSchema'>\n"
        "\t<s:ElementType name='row' content='eltOnly'>\n";
    for (std::size_t i = 0; i < set.field_count(); ++i) {
        const FieldInfo& field = set.field(i);
        const TypeInfo info = xml_type(field.type, field.defined_size);
        out += "\t\t<s:AttributeType name='" + xml_name(field.name) +
               "' rs:number='" + std::to_string(i + 1) + "'";
        if (set.editable()) {
            out += " rs:write='true'";
        }
        out += ">\n\t\t\t<s:datatype dt:type='";
        out += info.name;
        out += "' dt:maxLength='" + std::to_string(info.length) + "'";
        if (info.long_text) {
            out += " rs:long='true'";
        }
        out += " rs:maybenull='true'/>\n\t\t</s:AttributeType>\n";
    }
    out +=
        "\t\t<s:extends type='rs:rowbase'/>\n\t</s:ElementType>\n"
        "</s:Schema>\n<rs:data>\n";
    auto rows = set.clone();
    rows->move_first();
    while (!rows->eof()) {
        out += "\t<z:row";
        for (std::size_t i = 0; i < rows->field_count(); ++i) {
            const Value& value = rows->value(i);
            if (value.index() == 0) {
                continue;
            }
            out += " " + xml_name(rows->field(i).name) + "='" +
                   escape(value_text(value, rows->field(i).type)) + "'";
        }
        out += "/>\n";
        rows->move_next();
    }
    out += "</rs:data>\n</xml>";
    return out;
}

namespace {

using Attributes = std::vector<std::pair<std::string, std::string>>;

// The attributes of the tag that starts at text[at] ('<'); `end` is left
// after its '>'.
Attributes read_tag(const std::string& text, const std::size_t at,
                    std::string& name, std::size_t& end) {
    Attributes attributes;
    std::size_t i = at + 1;
    while (i < text.size() && text[i] != ' ' && text[i] != '>' &&
           text[i] != '/' && text[i] != '\t' && text[i] != '\n' &&
           text[i] != '\r') {
        ++i;
    }
    name = text.substr(at + 1, i - at - 1);
    while (i < text.size() && text[i] != '>') {
        if (std::isspace(static_cast<unsigned char>(text[i])) != 0 ||
            text[i] == '/') {
            ++i;
            continue;
        }
        const std::size_t key_start = i;
        while (i < text.size() && text[i] != '=' && text[i] != '>' &&
               std::isspace(static_cast<unsigned char>(text[i])) == 0) {
            ++i;
        }
        const std::string key = text.substr(key_start, i - key_start);
        while (i < text.size() &&
               std::isspace(static_cast<unsigned char>(text[i])) != 0) {
            ++i;
        }
        if (i < text.size() && text[i] == '=') {
            ++i;
            while (i < text.size() &&
                   std::isspace(static_cast<unsigned char>(text[i])) != 0) {
                ++i;
            }
            if (i < text.size() && (text[i] == '\'' || text[i] == '"')) {
                const char quote = text[i++];
                const std::size_t value_start = i;
                while (i < text.size() && text[i] != quote) {
                    ++i;
                }
                attributes.emplace_back(
                    key, unescape(text.substr(value_start, i - value_start)));
                ++i;
            }
        } else if (key.empty()) {
            ++i;
        }
    }
    end = i < text.size() ? i + 1 : i;
    return attributes;
}

std::string attribute(const Attributes& attributes, const std::string& name) {
    for (const auto& [key, value] : attributes) {
        if (key == name) {
            return value;
        }
    }
    return {};
}

int ado_type(const std::string& xml, const bool long_text) {
    if (xml == "i2") {
        return 2;
    }
    if (xml == "int" || xml == "i4") {
        return 3;
    }
    if (xml == "r4") {
        return 4;
    }
    if (xml == "float" || xml == "r8") {
        return 5;
    }
    if (xml == "i8") {
        return 20;
    }
    if (xml == "boolean") {
        return 11;
    }
    if (xml == "ui1") {
        return 17;
    }
    if (xml == "dateTime" || xml == "dateTime.tz") {
        return 7;
    }
    if (xml == "uuid") {
        return 72;
    }
    if (xml == "bin.hex") {
        return long_text ? 205 : 204;
    }
    if (xml == "number" || xml == "fixed.14.4") {
        return 131;
    }
    return long_text ? 203 : 202;
}

}  // namespace

std::shared_ptr<Recordset> load_xml(const std::string& text) {
    struct Column {
        std::string name;
        int type;
        long size;
    };
    std::vector<Column> columns;
    std::vector<Attributes> rows;
    std::size_t at = 0;
    bool schema = false;
    while ((at = text.find('<', at)) != std::string::npos) {
        if (text.compare(at, 2, "<?") == 0 ||
            text.compare(at, 4, "<!--") == 0) {
            at = text.find('>', at);
            if (at == std::string::npos) {
                break;
            }
            continue;
        }
        std::string tag;
        std::size_t end = 0;
        const Attributes attributes = read_tag(text, at, tag, end);
        if (tag == "s:AttributeType" || tag == "AttributeType") {
            schema = true;
            columns.push_back(
                Column{plain_name(attribute(attributes, "name")), 202, 255});
        } else if ((tag == "s:datatype" || tag == "datatype") &&
                   !columns.empty()) {
            const bool long_text = attribute(attributes, "rs:long") == "true";
            columns.back().type =
                ado_type(attribute(attributes, "dt:type"), long_text);
            const std::string length = attribute(attributes, "dt:maxLength");
            columns.back().size =
                length.empty() ? 0 : std::strtol(length.c_str(), nullptr, 10);
            if (long_text || columns.back().size > 1000000) {
                columns.back().size = 0;
            }
        } else if (tag == "z:row" || tag == "row") {
            rows.push_back(attributes);
        }
        at = end;
    }
    if (!schema || columns.empty()) {
        throw ComError{3000, "The file is not a persisted recordset."};
    }
    auto set = Recordset::create();
    for (const auto& column : columns) {
        set->append_field(column.name, column.type, column.size);
    }
    set->open_empty();
    for (const auto& row : rows) {
        set->add_new();
        for (std::size_t i = 0; i < columns.size(); ++i) {
            std::string raw;
            bool found = false;
            const std::string key = xml_name(columns[i].name);
            for (const auto& [name, value] : row) {
                if (name == key) {
                    raw = value;
                    found = true;
                    break;
                }
            }
            if (!found) {
                continue;
            }
            Value value;
            switch (columns[i].type) {
                case 2:
                case 3:
                case 16:
                case 17:
                case 20:
                    value = Value{static_cast<std::int64_t>(
                        std::strtoll(raw.c_str(), nullptr, 10))};
                    break;
                case 11:
                    value = Value{static_cast<std::int64_t>(
                        (raw == "true" || raw == "True" || raw == "1" ||
                         raw == "-1")
                            ? -1
                            : 0)};
                    break;
                case 4:
                case 5:
                case 6:
                case 131:
                    value = Value{std::strtod(raw.c_str(), nullptr)};
                    break;
                case 7:
                case 133:
                case 135: {
                    double serial = 0;
                    value =
                        parse_date(raw, serial) ? Value{serial} : Value{raw};
                    break;
                }
                case 128:
                case 204:
                case 205:
                    value = Value{bytes_of(raw)};
                    break;
                default:
                    value = Value{raw};
            }
            set->set_value(i, std::move(value));
        }
        set->update();
    }
    set->move_first();
    return set;
}

}  // namespace wfc::adodb::com
