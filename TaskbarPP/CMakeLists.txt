# CMakeList.txt : CMake project for TaskbarPP, include source and define
# project specific logic here.
#
cmake_minimum_required (VERSION 3.8)

# Add source to this project's executable.

find_package(GTest CONFIG REQUIRED)

set(CMAKE_CXX_STANDARD 17)
add_executable (TaskbarPP "TaskbarPP.cpp")
target_link_libraries(TaskbarPP PRIVATE GTest::gtest GTest::gtest_main)

# TODO: Add tests and install targets if needed.
