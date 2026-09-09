cmake_minimum_required(VERSION 3.14)
project(excel_toolkit C CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

include_directories(${CMAKE_CURRENT_SOURCE_DIR})

set(SOURCES
    gui_main.cpp
    excel_toolkit_core.h
    excel_toolkit_core.cpp
    miniz.h
    miniz.c
    resource.rc
)

add_executable(excel-toolkit WIN32 ${SOURCES})

target_link_libraries(excel-toolkit
    comctl32
    shell32
    user32
    gdi32
    advapi32
    comdlg32
)
