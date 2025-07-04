cmake_minimum_required(VERSION 3.16)

project(Spreadsheet VERSION 1.0.0 LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

find_package(Qt6 REQUIRED COMPONENTS Core Widgets)

set(CMAKE_AUTOMOC ON)
set(CMAKE_AUTORCC ON)
set(CMAKE_AUTOUIC ON)

set(SOURCES
    main.cpp
    core/Cell.cpp
    core/Worksheet.cpp
    core/Workbook.cpp
    ui/MainWindow.cpp
    ui/SpreadsheetView.cpp

)

set(HEADERS
    core/Cell.h
    core/Worksheet.h
    core/Workbook.h
    ui/MainWindow.h
    ui/SpreadsheetView.h

)


add_executable(Spreadsheet 
    ${SOURCES} 
    ${HEADERS}
    ui/MainWindow.ui
    ui/CellDetailEditor.h ui/CellDetailEditor.cpp
    ui/WorksheetManager.h ui/WorksheetManager.cpp
    ui/SearchWidget.h ui/SearchWidget.cpp
    core/FileManager.h core/FileManager.cpp
)

target_link_libraries(Spreadsheet 
    Qt6::Core 
    Qt6::Widgets
)
