cmake_minimum_required(VERSION 3.16)

project(ThermalTool VERSION 1.1 LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_AUTOMOC ON)

find_package(QT NAMES Qt6 Qt5 REQUIRED COMPONENTS Core Gui Widgets)
find_package(Qt${QT_VERSION_MAJOR} REQUIRED COMPONENTS Core Gui Widgets)

set(QXLSX_DIR "" CACHE PATH "Optional path to a local QXlsx source directory")

if (QXLSX_DIR AND EXISTS "${QXLSX_DIR}/CMakeLists.txt")
    add_subdirectory("${QXLSX_DIR}" "${CMAKE_BINARY_DIR}/QXlsx")
else()
    include(FetchContent)
    FetchContent_Declare(
        QXlsx
        GIT_REPOSITORY https://github.com/QtExcel/QXlsx.git
        GIT_TAG master
        SOURCE_SUBDIR QXlsx
    )
    FetchContent_MakeAvailable(QXlsx)
endif()

add_executable(ThermalTool
    main.cpp
)

set_target_properties(ThermalTool PROPERTIES
    WIN32_EXECUTABLE TRUE
)

target_link_libraries(ThermalTool PRIVATE
    Qt${QT_VERSION_MAJOR}::Core
    Qt${QT_VERSION_MAJOR}::Widgets
    QXlsx::QXlsx
)

if (MSVC)
    target_compile_options(ThermalTool PRIVATE /W4)
else()
    target_compile_options(ThermalTool PRIVATE -Wall -Wextra -Wpedantic)
endif()
