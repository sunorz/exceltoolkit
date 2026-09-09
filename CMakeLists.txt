cmake_minimum_required(VERSION 3.15)
project(ExcelToolkit LANGUAGES C CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

if(MSVC)
    set(CMAKE_MSVC_RUNTIME_LIBRARY "MultiThreaded$<$<CONFIG:Debug>:Debug>")
    add_compile_options(/W3 /utf-8 /EHsc)
elseif(MINGW)
    add_compile_options(-Wall)
    add_link_options(-static -static-libgcc -static-libstdc++)
endif()

set(SOURCES
    main.cpp
    excel_toolkit_core.h
    excel_toolkit_core.cpp
    miniz.h
    miniz.c
    resource.rc
)

add_executable(exceltookit ${SOURCES})
