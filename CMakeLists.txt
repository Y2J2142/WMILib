cmake_minimum_required(VERSION 3.10)

project(WMILib VERSION 1.0)


set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED True)


add_library(WMILib SHARED 
src/WMIlib.cpp 
src/WMIlib.hpp
)

target_include_directories(WMILib PUBLIC src/ )

#only used for basic testing
#add_executable(WMITest src/main.cpp)
#target_link_libraries(WMITest PUBLIC WMILib)