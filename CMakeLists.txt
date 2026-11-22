# wfc-adodb: ADODB (Connection, Command, Recordset) implemented over SQLite.
#
# SQLite is vendored as the amalgamation in third_party/sqlite (public domain).
# Parent project variables used when present: WFC_ADODB_TEST_PREFIX and the
# functions wfc_enable_strict_warnings and wsp_enable_hardening.

if(CMAKE_SOURCE_DIR STREQUAL CMAKE_CURRENT_SOURCE_DIR)
  cmake_minimum_required(VERSION 3.28)
  project(wfc-adodb LANGUAGES C CXX)
  option(WFC_BUILD_TESTS "Build the tests" ON)
  include(CTest)
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
set_target_properties(wfc_adodb_sqlite PROPERTIES OUTPUT_NAME wfcadodb_sqlite)
if(NOT WIN32)
  find_package(Threads REQUIRED)
  target_link_libraries(wfc_adodb_sqlite PUBLIC Threads::Threads ${CMAKE_DL_LIBS})
endif()

add_library(wfc_adodb STATIC src/adodb.cpp src/odbc.cpp)
target_include_directories(wfc_adodb PUBLIC "${CMAKE_CURRENT_SOURCE_DIR}/include")
target_link_libraries(wfc_adodb PUBLIC wfc_adodb_sqlite)
if(WIN32)
  target_link_libraries(wfc_adodb PUBLIC odbc32)
endif()
target_compile_features(wfc_adodb PUBLIC cxx_std_20)
set_target_properties(wfc_adodb PROPERTIES OUTPUT_NAME wfcadodb_static)
if(COMMAND wfc_enable_strict_warnings)
  wfc_enable_strict_warnings(wfc_adodb)
endif()
if(COMMAND wsp_enable_hardening)
  wsp_enable_hardening(wfc_adodb)
endif()

if(BUILD_TESTING AND WFC_BUILD_TESTS)
  add_executable(adodb_test tests/adodb_test.cpp)
  target_link_libraries(adodb_test PRIVATE wfc_adodb)
  if(COMMAND wfc_enable_strict_warnings)
    wfc_enable_strict_warnings(adodb_test)
  endif()
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
  add_test(NAME ${WFC_ADODB_TEST_PREFIX}TC-0001-adodb COMMAND adodb_test)
  set_tests_properties(
    ${WFC_ADODB_TEST_PREFIX}TC-0001-adodb
    PROPERTIES LABELS "unit;adodb" TIMEOUT 120
  )
endif()
