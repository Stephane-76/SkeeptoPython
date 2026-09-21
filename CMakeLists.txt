# Native Python extension for skeepto-engine.
# Links prebuilt static libs from SK_ENGINE_DIR (sibling ../skeepto-engine by default).
# Does not compile the engine and does not install into the system Python.
#
# Always use a venv. pip install -e . drives this CMakeLists via scikit-build-core
# (Release by default). Match CMAKE_BUILD_TYPE to the engine .a (Debug unix/lib
# vs Release): mixing NDEBUG / _DEBUGLeak layouts SIGSEGVs.
#
# 1. Engine (required first), from skeepto-engine:
#      cmake -B build-unix
#      cmake --build build-unix --parallel
#      cmake --build build-unix --target run-tests
#    Release engine (no leak counters):
#      cmake -B build-unix -DCMAKE_BUILD_TYPE=Release
#      cmake --build build-unix --parallel
#
# 2. Python venv + extension, from skeepto-python:
#      python3 -m venv .venv
#      source .venv/bin/activate
#      pip install -e ".[test]"
#      pytest
#    Optional: SK_ENGINE_DIR=/path/to/skeepto-engine pip install -e ".[test]"
#
# 3. Direct CMake (optional; pip is the usual path):
#      cmake -B build -DSK_ENGINE_DIR=../skeepto-engine
#      cmake --build build --parallel
#    Release extension against a Release engine:
#      cmake -B build -DCMAKE_BUILD_TYPE=Release -DSK_ENGINE_DIR=../skeepto-engine
#      cmake --build build --parallel

cmake_minimum_required(VERSION 3.16)
project(skeepto_python LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

if(APPLE)
	set(CMAKE_OSX_DEPLOYMENT_TARGET "14.0" CACHE STRING "Minimum macOS deployment version" FORCE)
endif()

# --- Engine (prebuilt) -------------------------------------------------------
set(SK_ENGINE_DIR "" CACHE PATH "Path to the skeepto-engine checkout")
if(SK_ENGINE_DIR STREQUAL "")
	if(DEFINED ENV{SK_ENGINE_DIR} AND NOT "$ENV{SK_ENGINE_DIR}" STREQUAL "")
		set(SK_ENGINE_DIR "$ENV{SK_ENGINE_DIR}")
	else()
		set(SK_ENGINE_DIR "${CMAKE_CURRENT_SOURCE_DIR}/../skeepto-engine")
	endif()
endif()
# Windows CMAKE_ARGS: D:\a\... backslashes are CMake escapes (\a eats the path).
file(TO_CMAKE_PATH "${SK_ENGINE_DIR}" SK_ENGINE_DIR)
get_filename_component(SK_ENGINE_DIR "${SK_ENGINE_DIR}" ABSOLUTE)

if(NOT EXISTS "${SK_ENGINE_DIR}/Libraries/SkSpreadSheet/include/SkApi.hpp")
	message(FATAL_ERROR
		"skeepto-engine not found at ${SK_ENGINE_DIR}. "
		"Pass -DSK_ENGINE_DIR=/path/to/skeepto-engine or keep the repos as siblings.")
endif()

if(WIN32)
	set(_SK_PLATFORM "windows")
	set(_SK_LIB_DIR "${SK_ENGINE_DIR}/windows/lib/Release")
	set(_SK_LIBS
		"${_SK_LIB_DIR}/SkSpreadSheet.lib"
		"${_SK_LIB_DIR}/SkFormat.lib"
		"${_SK_LIB_DIR}/SkRoot.lib")
elseif(APPLE AND EXISTS "${SK_ENGINE_DIR}/unix/lib/libSkSpreadSheet.a")
	set(_SK_PLATFORM "unix")
	set(_SK_LIB_DIR "${SK_ENGINE_DIR}/unix/lib")
	set(_SK_LIBS
		"${_SK_LIB_DIR}/libSkSpreadSheet.a"
		"${_SK_LIB_DIR}/libSkFormat.a"
		"${_SK_LIB_DIR}/libSkRoot.a")
elseif(APPLE AND EXISTS "${SK_ENGINE_DIR}/xcode/lib/Release/libSkSpreadSheet.a")
	set(_SK_PLATFORM "xcode")
	set(_SK_LIB_DIR "${SK_ENGINE_DIR}/xcode/lib/Release")
	set(_SK_LIBS
		"${_SK_LIB_DIR}/libSkSpreadSheet.a"
		"${_SK_LIB_DIR}/libSkFormat.a"
		"${_SK_LIB_DIR}/libSkRoot.a")
else()
	set(_SK_PLATFORM "unix")
	set(_SK_LIB_DIR "${SK_ENGINE_DIR}/unix/lib")
	set(_SK_LIBS
		"${_SK_LIB_DIR}/libSkSpreadSheet.a"
		"${_SK_LIB_DIR}/libSkFormat.a"
		"${_SK_LIB_DIR}/libSkRoot.a")
endif()

foreach(_sk_lib IN LISTS _SK_LIBS)
	if(NOT EXISTS "${_sk_lib}")
		message(FATAL_ERROR
			"Engine library missing: ${_sk_lib}\n"
			"Build skeepto-engine first, e.g.:\n"
			"  cd ${SK_ENGINE_DIR} && cmake -B build-unix && cmake --build build-unix --parallel")
	endif()
endforeach()

# tClass layout depends on NDEBUG (m_IndiceAlloc / leak counters). scikit-build
# compiles the extension as Release (-DNDEBUG) by default; unix/lib is typically
# a Debug engine. Mixing them SIGSEGVs (Apple PAC: KERN_INVALID_ADDRESS).
set(_SK_ENGINE_DEBUG FALSE)
if(NOT WIN32 AND EXISTS "${_SK_LIB_DIR}/libSkRoot.a")
	execute_process(
		COMMAND nm "${_SK_LIB_DIR}/libSkRoot.a"
		OUTPUT_VARIABLE _SK_ROOT_SYMS
		ERROR_QUIET
	)
	if(_SK_ROOT_SYMS MATCHES "GetNbAlloc")
		set(_SK_ENGINE_DEBUG TRUE)
	endif()
endif()

if(_SK_ENGINE_DEBUG)
	# Strip NDEBUG before the extension is compiled so SkTypes.hpp defines
	# _DEBUGLeak like the engine .a (m_IndiceAlloc on every tClass).
	foreach(_sk_var CMAKE_CXX_FLAGS CMAKE_CXX_FLAGS_RELEASE CMAKE_CXX_FLAGS_RELWITHDEBINFO CMAKE_CXX_FLAGS_MINSIZEREL)
		if(DEFINED ${_sk_var})
			string(REPLACE "-DNDEBUG" "" ${_sk_var} "${${_sk_var}}")
			string(REPLACE "/DNDEBUG" "" ${_sk_var} "${${_sk_var}}")
		endif()
	endforeach()
endif()

message(STATUS "skeepto-python: engine=${SK_ENGINE_DIR} platform=${_SK_PLATFORM} lib=${_SK_LIB_DIR} debug_abi=${_SK_ENGINE_DEBUG}")

# --- pybind11 (fetched into this build tree only) ----------------------------
include(FetchContent)
FetchContent_Declare(
	pybind11
	GIT_REPOSITORY https://github.com/pybind/pybind11.git
	GIT_TAG v2.13.6
)
FetchContent_MakeAvailable(pybind11)

# --- Extension ---------------------------------------------------------------
set(_SK_REACT "${SK_ENGINE_DIR}/SkReactSpreadSheet")
set(_SK_EXCEL "${SK_ENGINE_DIR}/SkExcel")
set(_SK_THIRD "${SK_ENGINE_DIR}/third-party")
set(_SK_EXCEL_SOURCES
	"${_SK_EXCEL}/source/SkExcel2SpreadSheet.cpp"
	"${_SK_EXCEL}/source/SkExcelPugiXMLReader.cpp"
	"${_SK_EXCEL}/source/SkExcelTools.cpp"
	"${_SK_EXCEL}/source/SkSpreadSheet2Excel.cpp"
	"${_SK_EXCEL}/source/SkExcelCollectCharts.cpp"
)
if(MSVC AND EXISTS "${_SK_THIRD}/pugixml/src/pugixml.cpp")
	list(APPEND _SK_EXCEL_SOURCES "${_SK_THIRD}/pugixml/src/pugixml.cpp")
endif()

pybind11_add_module(_core MODULE
	src/SkCore.cpp
	"${_SK_REACT}/source/SkUISpreadSheetApi.cpp"
	"${_SK_REACT}/source/SkJavascriptFunction.cpp"
	"${_SK_REACT}/source/SkJsonSelect.cpp"
	${_SK_EXCEL_SOURCES}
)

target_include_directories(_core PRIVATE
	"${SK_ENGINE_DIR}/Libraries/SkRoot/include"
	"${SK_ENGINE_DIR}/Libraries/SkFormat/include"
	"${SK_ENGINE_DIR}/Libraries/SkSpreadSheet/include"
	"${_SK_REACT}/include"
	"${_SK_EXCEL}/include"
	"${SK_ENGINE_DIR}/third-party/rapidjson/include"
	"${SK_ENGINE_DIR}/third-party/rapidxml"
	"${_SK_THIRD}/pugixml/src"
)

# libzip / pugixml — same search as skeepto-engine/SkExcel/CMakeLists.txt
if(APPLE AND EXISTS "/opt/homebrew/opt/libzip")
	set(_SK_LIBZIP "/opt/homebrew/opt/libzip")
elseif(EXISTS "${_SK_THIRD}/libzip/windows-install/include/zip.h")
	set(_SK_LIBZIP "${_SK_THIRD}/libzip/windows-install")
elseif(EXISTS "/usr/lib/x86_64-linux-gnu/libzip.a" OR EXISTS "/usr/lib/libzip.so" OR EXISTS "/usr/lib64/libzip.so")
	set(_SK_LIBZIP "/usr")
else()
	set(_SK_LIBZIP "/opt/homebrew/opt/libzip")
endif()
if(EXISTS "${_SK_LIBZIP}/include")
	target_include_directories(_core PRIVATE "${_SK_LIBZIP}/include")
endif()
if(WIN32 AND EXISTS "${_SK_THIRD}/libzip/zlib-install/include")
	target_include_directories(_core PRIVATE "${_SK_THIRD}/libzip/zlib-install/include")
endif()
if(EXISTS "/opt/homebrew/include")
	target_include_directories(_core PRIVATE "/opt/homebrew/include")
endif()

if(_SK_ENGINE_DEBUG)
	# Later -UNDEBUG wins over CMAKE_CXX_FLAGS_RELEASE -DNDEBUG.
	target_compile_options(_core PRIVATE -UNDEBUG)
endif()

# Circular refs between SkRoot / SkFormat / SkSpreadSheet: repeat the archives
# (Apple ld) or start-group (GNU ld).
if(CMAKE_CXX_COMPILER_ID STREQUAL "GNU")
	# Python sysconfig often injects -flto; engine .a is not LTO bitcode.
	target_compile_options(_core PRIVATE -fno-lto)
	target_link_options(_core PRIVATE -fno-lto)
	target_link_libraries(_core PRIVATE
		-Wl,--start-group ${_SK_LIBS} -Wl,--end-group)
else()
	target_link_libraries(_core PRIVATE ${_SK_LIBS} ${_SK_LIBS})
endif()

if(APPLE)
	target_link_libraries(_core PRIVATE "-framework CoreFoundation")
endif()

if(EXISTS "${_SK_LIBZIP}/lib/libzip.a")
	set(_SK_ZIP "${_SK_LIBZIP}/lib/libzip.a")
elseif(EXISTS "${_SK_LIBZIP}/lib/libzip.dylib")
	set(_SK_ZIP "${_SK_LIBZIP}/lib/libzip.dylib")
elseif(EXISTS "${_SK_LIBZIP}/lib/zip.lib")
	set(_SK_ZIP "${_SK_LIBZIP}/lib/zip.lib")
elseif(EXISTS "${_SK_LIBZIP}/lib/libzip.lib")
	set(_SK_ZIP "${_SK_LIBZIP}/lib/libzip.lib")
elseif(EXISTS "/usr/lib/x86_64-linux-gnu/libzip.a")
	set(_SK_ZIP "/usr/lib/x86_64-linux-gnu/libzip.a")
elseif(EXISTS "/usr/lib64/libzip.so")
	set(_SK_ZIP "/usr/lib64/libzip.so")
elseif(EXISTS "/usr/lib/libzip.so")
	set(_SK_ZIP "/usr/lib/libzip.so")
else()
	set(_SK_ZIP "zip")
endif()
# Windows libzip is built against zlib in third-party/libzip/zlib-install.
if(WIN32)
	set(_SK_ZLIB "${_SK_THIRD}/libzip/zlib-install")
	if(EXISTS "${_SK_ZLIB}/lib/zlibstatic.lib")
		set(_SK_Z "${_SK_ZLIB}/lib/zlibstatic.lib")
	elseif(EXISTS "${_SK_ZLIB}/lib/zlib.lib")
		set(_SK_Z "${_SK_ZLIB}/lib/zlib.lib")
	elseif(EXISTS "${_SK_ZLIB}/lib/libz.lib")
		set(_SK_Z "${_SK_ZLIB}/lib/libz.lib")
	endif()
endif()
if(EXISTS "/opt/homebrew/lib/libpugixml.a")
	set(_SK_PUGI "/opt/homebrew/lib/libpugixml.a")
elseif(EXISTS "/usr/local/lib/libpugixml.a")
	set(_SK_PUGI "/usr/local/lib/libpugixml.a")
elseif(EXISTS "${_SK_THIRD}/pugixml/build/libpugixml.a")
	set(_SK_PUGI "${_SK_THIRD}/pugixml/build/libpugixml.a")
elseif(NOT MSVC)
	set(_SK_PUGI "pugixml")
endif()
if(DEFINED _SK_ZIP)
	target_link_libraries(_core PRIVATE ${_SK_ZIP})
endif()
if(DEFINED _SK_Z)
	target_link_libraries(_core PRIVATE ${_SK_Z})
endif()
if(DEFINED _SK_PUGI)
	target_link_libraries(_core PRIVATE ${_SK_PUGI})
endif()
if(MSVC)
	target_link_libraries(_core PRIVATE Rpcrt4.lib Shlwapi.lib)
endif()

install(TARGETS _core DESTINATION skeepto)
