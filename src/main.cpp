#include "WMIlib.hpp"
#include <iostream>
int main() 
{ 
    std::cout << "TEST\n";
    auto api = WMIAPI{};
    if(!api.WmiInitialize()) {
        api.WmiUninitialize();
        return -1;
    }
    while (true) {
        std::wcout << "CPU usage : " << api.getTotalCPUUsage() << std::endl;
        std::wcout << "Total memory : " << api.getTotalMemory()  << "MB" << std::endl;
        std::wcout << "Free memory : " << api.getAvailableMemory()  << "MB" << std::endl;
        std::wcout << "OS : " << api.getOSName() << std::endl;
        std::wcout << "Mem speed : " << api.getMemClockSpeed() << std::endl;
        std::wcout << "Mem Type : " << api.getMemType() << std::endl;
        std::wcout << "Mem Voltage : " << api.getMemVoltage() << "mV" << std::endl;


    }
    api.WmiUninitialize();
    return 0; 
}