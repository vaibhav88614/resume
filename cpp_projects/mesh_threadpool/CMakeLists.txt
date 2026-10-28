cmake_minimum_required(VERSION 3.14)
project(mesh_threadpool LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

find_package(Threads REQUIRED)

add_executable(mesh_threadpool src/main.cpp)
target_link_libraries(mesh_threadpool PRIVATE Threads::Threads)

if(MSVC)
  target_compile_options(mesh_threadpool PRIVATE /W4)
else()
  target_compile_options(mesh_threadpool PRIVATE -Wall -Wextra -Wpedantic)
endif()
