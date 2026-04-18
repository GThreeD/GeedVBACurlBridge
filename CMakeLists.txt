cmake_minimum_required(VERSION 3.24)
project(libcurl_vba_bridge C)

# ----------------------------
# Version
# ----------------------------
set(WSC_VERSION_MAJOR 1)
set(WSC_VERSION_MINOR 0)
set(WSC_VERSION_PATCH 2)
set(WSC_VERSION_BUILD 0)

set(WSC_BRIDGE_API_VERSION 1)

set(WSC_VERSION_STRING
    "${WSC_VERSION_MAJOR}.${WSC_VERSION_MINOR}.${WSC_VERSION_PATCH}.${WSC_VERSION_BUILD}"
)

# ----------------------------
# Paths
# ----------------------------
set(WSC_OUT_DIR "${CMAKE_SOURCE_DIR}/out")
set(WSC_OUT_BIN_DIR "${WSC_OUT_DIR}/bin")
set(WSC_OUT_VBA_DIR "${WSC_OUT_DIR}/vba")
set(WSC_VBA_SRC_DIR "${CMAKE_SOURCE_DIR}/src")

set(WSC_EXCEL_CURL_DIR "$ENV{HOME}/Library/Containers/com.microsoft.Excel/Data/curl")
set(WSC_DEPLOY_LIB "${WSC_EXCEL_CURL_DIR}/libcurl_vba_bridge.dylib")

file(MAKE_DIRECTORY "${WSC_OUT_BIN_DIR}")
file(MAKE_DIRECTORY "${WSC_OUT_VBA_DIR}")

# ----------------------------
# Dependencies
# ----------------------------
#set(CURL_USE_STATIC_LIBS ON)
find_package(CURL REQUIRED)

# ----------------------------
# Native library
# ----------------------------
add_library(libcurl_vba_bridge SHARED
    src/libcurl_vba_bridge.c
)

target_include_directories(libcurl_vba_bridge PRIVATE
    ${CURL_INCLUDE_DIRS}
)

target_link_libraries(libcurl_vba_bridge PRIVATE
    ${CURL_LIBRARIES}
)
#target_link_libraries(libcurl_vba_bridge PRIVATE CURL::libcurl)

target_compile_definitions(libcurl_vba_bridge PRIVATE
    WSC_BRIDGE_API_VERSION=${WSC_BRIDGE_API_VERSION}
    WSC_BRIDGE_VERSION_MAJOR=${WSC_VERSION_MAJOR}
    WSC_BRIDGE_VERSION_MINOR=${WSC_VERSION_MINOR}
    WSC_BRIDGE_VERSION_PATCH=${WSC_VERSION_PATCH}
    WSC_BRIDGE_VERSION_BUILD=${WSC_VERSION_BUILD}
    WSC_BRIDGE_VERSION_STRING="${WSC_VERSION_STRING}"
)

set_target_properties(libcurl_vba_bridge PROPERTIES
    PREFIX ""
    OUTPUT_NAME "libcurl_vba_bridge"
    LIBRARY_OUTPUT_DIRECTORY "${WSC_OUT_BIN_DIR}"
)

# ----------------------------
# VBA files -> out/vba
# ----------------------------
file(GLOB WSC_VBA_FILES CONFIGURE_DEPENDS
    "${WSC_VBA_SRC_DIR}/*.bas"
    "${WSC_VBA_SRC_DIR}/*.cls"
    "${WSC_VBA_SRC_DIR}/*.frm"
)

foreach(vba_file IN LISTS WSC_VBA_FILES)
    get_filename_component(vba_name "${vba_file}" NAME)
    configure_file(
        "${vba_file}"
        "${WSC_OUT_VBA_DIR}/${vba_name}"
        @ONLY
        NEWLINE_STYLE DOS
    )
    message(STATUS "Copied VBA file: ${vba_name}")
endforeach()

# ----------------------------
# Deploy to Excel + sign
# ----------------------------
add_custom_target(deploy
    COMMAND ${CMAKE_COMMAND} -E make_directory "${WSC_EXCEL_CURL_DIR}"
    COMMAND ${CMAKE_COMMAND} -E copy
            "${WSC_OUT_BIN_DIR}/libcurl_vba_bridge.dylib"
            "${WSC_DEPLOY_LIB}"
    COMMAND /usr/bin/codesign --force --sign - --timestamp=none "${WSC_DEPLOY_LIB}"
    DEPENDS libcurl_vba_bridge
    COMMENT "Deploying libcurl_vba_bridge.dylib to Excel curl folder"
)

# ----------------------------
# Info
# ----------------------------
message(STATUS "WSC version: ${WSC_VERSION_STRING}")
message(STATUS "WSC API version: ${WSC_BRIDGE_API_VERSION}")
message(STATUS "WSC output bin dir: ${WSC_OUT_BIN_DIR}")
message(STATUS "WSC output vba dir: ${WSC_OUT_VBA_DIR}")