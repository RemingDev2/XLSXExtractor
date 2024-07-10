cmake_minimum_required(VERSION 3.28)
project(XLSXExtractor)

set(CMAKE_CXX_STANDARD 17)

############### OpenXLSX ###############
add_subdirectory(external/OpenXLSX)

############### SQLiteC++ ###############
add_subdirectory(external/SQLiteCpp-3.3.1)

add_executable(XLSXExtractor main.cpp
        external/Randomizer/Randomizer.cpp
        external/Randomizer/Randomizer.h
        src/files/XLSXWorker.cpp
        src/headers/XLSXWorker.h)

target_link_libraries(XLSXExtractor
        OpenXLSX::OpenXLSX
        SQLiteCpp sqlite3 pthread)
