// The ADODB COM classes of msado15.dll. Windows only; internal.

#ifndef WFC_ADODB_COM_CLASSES_HPP
#define WFC_ADODB_COM_CLASSES_HPP

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>

#include <cstddef>

namespace wfc::adodb::com {

// The class factory of class `index` of the library's class table
// (Connection, Recordset, Command).
HRESULT make_class_factory(std::size_t index, REFIID iid, void** object);

// True when no object of the library lives and the server is not locked.
bool can_unload() noexcept;

}  // namespace wfc::adodb::com

#endif  // WFC_ADODB_COM_CLASSES_HPP
