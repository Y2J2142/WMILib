#pragma once
//extern "C" __declspec(dllexport)
struct IWbemClassObject;
struct IWbemServices ;
struct  IWbemLocator;
#include <string>


enum class InitResult {
    Success = 0,
    ModeChanged,
    AlreadyStarted

};

struct WMIAPI {

    bool WmiInitialize();
    void WmiUninitialize();
    uint64_t getTotalCPUUsage();
    std::wstring getTotalMemory();
    std::wstring getAvailableMemory();
    std::wstring getOSName();
    uint32_t getMemClockSpeed();
    int32_t getMemType();
    int32_t getMemVoltage();
private:

    IWbemClassObject* pClass = nullptr;
    IWbemServices *pSvc = nullptr;
    IWbemLocator *pLoc = nullptr;
    bool isInitialized = false;
    
    InitResult InitializeCOM();
    bool GenerateCOMSecurity();
    bool ConnectToWMI();
    bool LocateWMI();
    bool SetProxySecurity();
    auto ExecQuery(const wchar_t *, const wchar_t *);
};

 