cmake_minimum_required(VERSION 3.16)
project(NewFolderFromFiles VERSION 1.0.0 LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

if(NOT CMAKE_SIZEOF_VOID_P EQUAL 8)
    message(FATAL_ERROR "Only 64-bit builds are supported")
endif()

# Shell Extension DLL
add_library(NewFolderFromFiles SHARED
    src/dllmain.cpp
    src/NewFolderFromFilesClassFactory.cpp
    src/NewFolderFromFilesContextMenuHandler.cpp
    src/NewFolderFromFiles.def
)

target_include_directories(NewFolderFromFiles PRIVATE src)
target_link_libraries(NewFolderFromFiles PRIVATE Shlwapi Shell32 Ole32)

# Hotkey Helper App
add_executable(NewFolderFromFilesHotkey WIN32
    src/HotkeyHelper.cpp
)

target_link_libraries(NewFolderFromFilesHotkey PRIVATE Shell32 Ole32 Shlwapi)

# Output to build/bin
set_target_properties(NewFolderFromFiles NewFolderFromFilesHotkey PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}/bin"
    LIBRARY_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}/bin"
)
