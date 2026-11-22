# wfc-adodb: ADODB (Connection, Command, Recordset) implemented over SQLite.
#
# SQLite is vendored as the amalgamation in third_party/sqlite (public domain).
# Parent project variables used when present: WFC_ADODB_TEST_PREFIX and the
# functions wfc_enable_strict_warnings and wsp_enable_hardening.

if(CMAKE_SOURCE_DIR STREQUAL CMAKE_CURRENT_SOURCE_DIR)
  cmake_minimum_required(VERSION 3.28)
  project(wfc-adodb LANGUAGES C CXX)
  if(WFC_WINELIB)
    # winegcc writes name.exe (a launcher) and name.exe.so; the platform module
    # resets the suffix when the project is declared, so it is set here.
    set(CMAKE_EXECUTABLE_SUFFIX ".exe")
  endif()
  option(WFC_BUILD_TESTS "Build the tests" ON)
  include(CTest)
  list(APPEND CMAKE_MODULE_PATH "${CMAKE_CURRENT_SOURCE_DIR}/cmake")
  include(WfcWarnings)
endif()

if(NOT DEFINED WFC_ADODB_TEST_PREFIX)
  if(CMAKE_SOURCE_DIR STREQUAL CMAKE_CURRENT_SOURCE_DIR)
    set(WFC_ADODB_TEST_PREFIX "")
  else()
    set(WFC_ADODB_TEST_PREFIX "wfc-adodb-")
  endif()
endif()

file(STRINGS "${CMAKE_CURRENT_SOURCE_DIR}/VERSION" WFC_ADODB_VERSION LIMIT_COUNT 1)
set(WFC_ADODB_VERSION "${WFC_ADODB_VERSION}" CACHE INTERNAL "wfc-adodb version")

# SQLite builds as its own C target: third-party code, no strict warnings.
enable_language(C)
add_library(wfc_adodb_sqlite STATIC third_party/sqlite/sqlite3.c)
target_include_directories(
  wfc_adodb_sqlite
  PUBLIC "${CMAKE_CURRENT_SOURCE_DIR}/third_party/sqlite"
)
target_compile_definitions(
  wfc_adodb_sqlite
  PRIVATE SQLITE_THREADSAFE=1 SQLITE_OMIT_LOAD_EXTENSION SQLITE_DQS=0
)
if(WFC_WINELIB)
  # A Winelib program is a Linux program: SQLite uses the POSIX file layer.
  target_compile_definitions(wfc_adodb_sqlite PRIVATE SQLITE_OS_UNIX=1 SQLITE_OS_WIN=0)
endif()
set_target_properties(wfc_adodb_sqlite PROPERTIES OUTPUT_NAME wfcadodb_sqlite)
if(NOT WIN32)
  find_package(Threads REQUIRED)
  target_link_libraries(wfc_adodb_sqlite PUBLIC Threads::Threads ${CMAKE_DL_LIBS})
endif()

add_library(wfc_adodb STATIC src/adodb.cpp src/odbc.cpp)
target_include_directories(wfc_adodb PUBLIC "${CMAKE_CURRENT_SOURCE_DIR}/include")
target_link_libraries(wfc_adodb PUBLIC wfc_adodb_sqlite)
if(WIN32 OR WFC_WINELIB)
  target_link_libraries(wfc_adodb PUBLIC odbc32)
  if(WFC_WINELIB)
    target_compile_definitions(wfc_adodb PRIVATE WFC_WINELIB=1)
  endif()
endif()
target_compile_features(wfc_adodb PUBLIC cxx_std_20)
set_target_properties(wfc_adodb PROPERTIES OUTPUT_NAME wfcadodb_static)
if(COMMAND wfc_enable_strict_warnings)
  wfc_enable_strict_warnings(wfc_adodb)
endif()
if(COMMAND wsp_enable_hardening)
  wsp_enable_hardening(wfc_adodb)
endif()

# The COM server: msado15.dll registers ADODB.Connection, ADODB.Recordset and
# ADODB.Command (under WFC's own ProgIDs, and under the Microsoft ones with
# WFC_REGISTER_MICROSOFT_IDS=1) and answers them from the library, so any COM
# client (a VB6 program, VBScript) gets the SQLite-backed ADODB. The static
# library above is still what WFC itself links.
if(NOT DEFINED WFC_OCX_SHARED)
  set(WFC_OCX_SHARED OFF)
endif()
option(
  WFC_ADODB_SHARED
  "Build the WFC ADODB library as msado15.dll"
  ${WFC_OCX_SHARED}
)
file(STRINGS "${CMAKE_CURRENT_SOURCE_DIR}/VERSION" WFC_ADODB_VERSION_TEXT
     LIMIT_COUNT 1)
if(WFC_ADODB_SHARED AND WIN32)
  enable_language(RC)
  string(REPLACE "." ";" version_parts "${WFC_ADODB_VERSION_TEXT}")
  list(GET version_parts 0 WFC_LIB_VERSION_MAJOR)
  list(GET version_parts 1 WFC_LIB_VERSION_MINOR)
  list(GET version_parts 2 WFC_LIB_VERSION_PATCH)
  set(WFC_LIB_VERSION "${WFC_ADODB_VERSION_TEXT}")
  file(STRINGS "${CMAKE_CURRENT_SOURCE_DIR}/TARGET" WFC_LIB_TARGET
       LIMIT_COUNT 1)
  configure_file(
    src/registration_info.hpp.in
    "${CMAKE_CURRENT_BINARY_DIR}/registration_info.hpp"
    @ONLY
  )
  set(WFC_LIB_NAME wfcadodb)
  set(WFC_LIB_FILE msado15.dll)
  set(WFC_LIB_DESCRIPTION "WFC ADODB")
  configure_file(
    version.rc.in
    "${CMAKE_CURRENT_BINARY_DIR}/wfcadodb.rc"
    @ONLY
  )
  add_library(
    wfc_adodb_dll
    SHARED
    src/registration.cpp
    src/com/classes.cpp
    src/adodb.def
    "${CMAKE_CURRENT_BINARY_DIR}/wfcadodb.rc"
  )
  target_include_directories(
    wfc_adodb_dll
    PRIVATE "${CMAKE_CURRENT_BINARY_DIR}" "${CMAKE_CURRENT_SOURCE_DIR}/src"
  )
  target_link_libraries(
    wfc_adodb_dll
    PRIVATE wfc_adodb advapi32 ole32 oleaut32
  )
  set_target_properties(
    wfc_adodb_dll
    PROPERTIES
      OUTPUT_NAME msado15
      SUFFIX ".dll"
      ARCHIVE_OUTPUT_NAME wfcadodb
      PDB_NAME wfcadodb
      VERSION "${WFC_ADODB_VERSION_TEXT}"
      RUNTIME_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}"
  )
  if(COMMAND wfc_enable_strict_warnings)
    wfc_enable_strict_warnings(wfc_adodb_dll)
  endif()
  if(COMMAND wsp_enable_hardening)
    wsp_enable_hardening(wfc_adodb_dll)
  endif()
endif()

if(BUILD_TESTING AND WFC_BUILD_TESTS)
  add_executable(adodb_test tests/adodb_test.cpp)
  target_link_libraries(adodb_test PRIVATE wfc_adodb)
  if(COMMAND wfc_enable_strict_warnings)
    wfc_enable_strict_warnings(adodb_test)
  endif()
  # The ODBC round trip needs the Windows driver manager and its Access
  # driver; a Winelib build compiles the backend but does not run this test.
  if(WIN32)
    add_executable(adodb_odbc_test tests/adodb_odbc_test.cpp)
    target_link_libraries(adodb_odbc_test PRIVATE wfc_adodb)
    if(WIN32)
      target_link_libraries(
        adodb_odbc_test
        PRIVATE odbccp32 legacy_stdio_definitions
      )
    endif()
    if(COMMAND wfc_enable_strict_warnings)
      wfc_enable_strict_warnings(adodb_odbc_test)
    endif()
    add_test(
      NAME ${WFC_ADODB_TEST_PREFIX}TC-0002-adodb-odbc
      COMMAND adodb_odbc_test
    )
    set_tests_properties(
      ${WFC_ADODB_TEST_PREFIX}TC-0002-adodb-odbc
      PROPERTIES LABELS "unit;adodb;odbc" TIMEOUT 120
    )
  endif()
  if(WIN32 AND WFC_ADODB_SHARED)
    add_executable(adodb_dll_test tests/adodb_dll_test.cpp)
    target_link_libraries(adodb_dll_test PRIVATE ole32 oleaut32 version)
    add_dependencies(adodb_dll_test wfc_adodb_dll)
    target_compile_features(adodb_dll_test PRIVATE cxx_std_20)
    if(COMMAND wfc_enable_strict_warnings)
      wfc_enable_strict_warnings(adodb_dll_test)
    endif()
    add_test(
      NAME ${WFC_ADODB_TEST_PREFIX}TC-0003-adodb-dll
      COMMAND
        adodb_dll_test "$<TARGET_FILE:wfc_adodb_dll>" "${WFC_ADODB_VERSION_TEXT}"
    )
    set_tests_properties(
      ${WFC_ADODB_TEST_PREFIX}TC-0003-adodb-dll
      PROPERTIES LABELS "integration;adodb" TIMEOUT 120
    )
  endif()
  add_test(NAME ${WFC_ADODB_TEST_PREFIX}TC-0001-adodb COMMAND adodb_test)
  set_tests_properties(
    ${WFC_ADODB_TEST_PREFIX}TC-0001-adodb
    PROPERTIES LABELS "unit;adodb" TIMEOUT 120
  )
endif()
