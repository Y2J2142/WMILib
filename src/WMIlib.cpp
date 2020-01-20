#define _WIN32_DCOM
#include <comdef.h>
#include <cstdlib>
#include <Wbemidl.h>
#include "WMIlib.hpp"
#include <iostream>
#include <string>
#pragma comment(lib, "wbemuuid.lib")
using namespace std::literals;

extern "C" __declspec( dllexport ) WMIAPI* __cdecl getWMIAPI() {
    return new WMIAPI();
}

extern "C" __declspec( dllexport ) bool __cdecl InitializeWMIAPI(WMIAPI* wmi) {
    return wmi->WmiInitialize();
}

extern "C" __declspec( dllexport ) uint64_t __cdecl getTotalCPUUsage(WMIAPI* wmi) {
    return wmi->getTotalCPUUsage();
}

extern "C" __declspec( dllexport ) BSTR __cdecl  getTotalMemory(WMIAPI* wmi) {
    return SysAllocString(wmi->getTotalMemory().c_str());
}
extern "C" __declspec ( dllexport ) BSTR __cdecl getAvailableMemory(WMIAPI* wmi) {
    return SysAllocString(wmi->getAvailableMemory().c_str());
}

extern "C" __declspec ( dllexport ) BSTR __cdecl getOSName(WMIAPI* wmi) {
    return SysAllocString(wmi->getOSName().c_str());
}

extern "C" __declspec ( dllexport ) uint32_t __cdecl getMemClockSpeed(WMIAPI* wmi) {
    return wmi->getMemClockSpeed();
}

extern "C" __declspec ( dllexport ) int32_t __cdecl getMemType(WMIAPI* wmi) {
    return wmi->getMemType();
}
extern "C" __declspec ( dllexport ) int32_t __cdecl getMemVoltage(WMIAPI* wmi) {
    return wmi->getMemVoltage();
}

extern "C" __declspec ( dllexport ) void __cdecl uninitializeWMIAPI(WMIAPI* wmi) {
    wmi->WmiUninitialize();
}





namespace {
    std::wstring InscpectVariant(VARIANT& v) {
        return [&v]() -> std::wstring{
            switch(v.vt) {
                case VT_ARRAY: 
	                return L"VT_ARRAY"s;
                case VT_BLOB: 
                    return L"VT_BLOB"s;
                case VT_BLOB_OBJECT: 
                    return L"VT_BLOB_OBJECT"s;
                case VT_BOOL: 
                    return L"VT_BOOL"s;
                case VT_BSTR: 
                    return L"VT_BSTR"s;
                case VT_BYREF: 
                    return L"VT_BYREF"s;
                case VT_CARRAY: 
                    return L"VT_CARRAY"s;
                case VT_CF: 
                    return L"VT_CF"s;
                case VT_CLSID: 
                    return L"VT_CLSID"s;
                case VT_CY: 
                    return L"VT_CY"s;
                case VT_DATE: 
                    return L"VT_DATE"s;
                case VT_DECIMAL: 
                    return L"VT_DECIMAL"s;
                case VT_DISPATCH: 
                    return L"VT_DISPATCH"s;
                case VT_EMPTY: 
                    return L"VT_EMPTY"s;
                case VT_ERROR: 
                    return L"VT_ERROR"s;
                case VT_FILETIME: 
                    return L"VT_FILETIME"s;
                case VT_HRESULT: 
                    return L"VT_HRESULT"s;
                case VT_I1: 
                    return L"VT_I1"s;
                case VT_I2: 
                    return L"VT_I2"s;
                case VT_I4: 
                    return L"VT_I4"s;
                case VT_I8: 
                    return L"VT_I8"s;
                case VT_ILLEGAL: 
                    return L"VT_ILLEGAL"s;
                case VT_ILLEGALMASKED: 
                    return L"VT_ILLEGALMASKED"s;
                case VT_INT: 
                    return L"VT_INT"s;
                case VT_LPSTR: 
                    return L"VT_LPSTR"s;
                case VT_LPWSTR: 
                    return L"VT_LPWSTR"s;
                case VT_NULL: 
                    return L"VT_NULL"s;
                case VT_PTR: 
                    return L"VT_PTR"s;
                case VT_R4: 
                    return L"VT_R4"s;
                case VT_R8: 
                    return L"VT_R8"s;
                case VT_RESERVED: 
                    return L"VT_RESERVED"s;
                case VT_SAFEARRAY: 
                    return L"VT_SAFEARRAY"s;
                case VT_STORAGE: 
                    return L"VT_STORAGE"s;
                case VT_STORED_OBJECT: 
                    return L"VT_STORED_OBJECT"s;
                case VT_STREAM: 
                    return L"VT_STREAM"s;
                case VT_STREAMED_OBJECT: 
                    return L"VT_STREAMED_OBJECT"s;
                case VT_UI1: 
                    return L"VT_UI1"s;
                case VT_UI2: 
                    return L"VT_UI2"s;
                case VT_UI4: 
                    return L"VT_UI4"s;
                case VT_UI8: 
                    return L"VT_UI8"s;
                case VT_UINT: 
                    return L"VT_UINT"s;
                case VT_UNKNOWN: 
                    return L"VT_UNKNOWN"s;
                case VT_USERDEFINED: 
                    return L"VT_USERDEFINED"s;
                case VT_VARIANT: 
                    return L"VT_VARIANT"s;
                case VT_VECTOR: 
                    return L"VT_VECTOR"s;
                case VT_VOID: 
                    return L"VT_VOID"s; 
                default :
                    return L"Fuck me if I know fam"s;
            }
        }();




    }
}



bool WMIAPI::WmiInitialize() {


    const auto res = InitializeCOM();
    if (res != InitResult::AlreadyStarted)
    {
        if (!GenerateCOMSecurity()) {
            std::cout << "COM security init failed\n";
            return false;
        }
    }

    if (!LocateWMI()) {
        std::cout << "WMI Locator failed\n";
        return false;
    }
    if (!ConnectToWMI()) {
        std::cout << "WMI connection failed\n";
        return false;
    }
    if (!SetProxySecurity()) {
        std::cout << "Set proxy security failed\n";
        return false;
    }
    std::cout << "Initialization complete\n";
    return true;


}

InitResult WMIAPI::InitializeCOM() {
    if (isInitialized) {
        return InitResult::Success;
    }
    auto hres = CoInitializeEx(0, COINIT_MULTITHREADED);
    isInitialized = true;
    switch(hres) {
        case S_OK:
            return InitResult::Success;
        case S_FALSE:
            return InitResult::AlreadyStarted;
        case RPC_E_CHANGED_MODE:
            return InitResult::ModeChanged;
    }
}

void WMIAPI::WmiUninitialize() {
    if(pLoc != nullptr) {
        pLoc->Release();
        pLoc = nullptr;
    }
    if(pSvc != nullptr) {
        pSvc->Release();
        pLoc = nullptr;
    }
    if(isInitialized) {
        CoUninitialize();
        isInitialized = false;
    }

}

bool WMIAPI::GenerateCOMSecurity() {
    auto hres = CoInitializeSecurity(
        nullptr,
        -1,
        nullptr,
        nullptr,
        RPC_C_AUTHN_LEVEL_DEFAULT,
        RPC_C_IMP_LEVEL_IMPERSONATE,
        nullptr,
        EOAC_NONE,
        nullptr
    );
    if(FAILED(hres)) {
        std::cout << "COM initialization failed\t Error: " << std::hex << hres << '\n'; 
        CoUninitialize();
        return false;
    }
    return true;
}

bool WMIAPI::LocateWMI() {
    auto hres = CoCreateInstance(
        CLSID_WbemLocator,
        0,
        CLSCTX_INPROC_SERVER,
        IID_IWbemLocator,
       (LPVOID*)&pLoc
    );

    if(FAILED(hres)){
        std::cout << "WMI locator failed\t Error: " << std::hex << hres << '\n'; 

        CoUninitialize();
        return false;
    }

    return true;
}

bool WMIAPI::ConnectToWMI() {
    auto hres = pLoc->ConnectServer(
        _bstr_t(L"root\\cimv2"),
        nullptr,
        nullptr,
        0,
        0,
        0,
        0,
        &pSvc
    );

    if(FAILED(hres)) {
        std::cout << "WMI connection failed\t Error: " << std::hex << hres << '\n'; 

        pLoc->Release();
        CoUninitialize();
        return false;
    }

    return true;

}

bool WMIAPI::SetProxySecurity() {
    auto hres = CoSetProxyBlanket(
        pSvc,
        RPC_C_AUTHN_WINNT,
        RPC_C_AUTHZ_NONE,
        nullptr,
        RPC_C_AUTHN_LEVEL_CALL,
        RPC_C_IMP_LEVEL_IMPERSONATE,
        nullptr,
        EOAC_NONE
    );
    if(FAILED(hres)) {
        std::cout << "Proxy blanket failed\t Error: " << std::hex << hres << '\n'; 

        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return false;
    }

    return true;
} 

auto WMIAPI::ExecQuery(const wchar_t* query, const wchar_t* name) {

    IEnumWbemClassObject* pEnumerator = nullptr;


    auto hr = pSvc->ExecQuery(
        bstr_t(L"WQL"),
        bstr_t(query),
        WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
        NULL,
        &pEnumerator
    );

    if (FAILED(hr)) {
        std::cout << "Query failed\t Error: " << std::hex << hr << '\n';
        return VARIANT{};
    }

    IWbemClassObject* pclsObj = NULL;
    ULONG uReturn = 0;


    HRESULT h = pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);

    if (uReturn == 0)  return VARIANT{};

    VARIANT vtProp;
    hr = pclsObj->Get(name, 0, &vtProp, 0, 0);
    pclsObj->Release();
    pEnumerator->Release();
    return vtProp;

}

uint64_t WMIAPI::getTotalCPUUsage() {

    auto variant = ExecQuery(L"select LoadPercentage from Win32_Processor", 
                        L"LoadPercentage" );
    return variant.vt == VT_I4 ? variant.intVal : 0;

}

std::wstring WMIAPI::getTotalMemory() {
    auto variant = ExecQuery(L"SELECT * FROM Win32_OperatingSystem ", L"TotalVisibleMemorySize");

    return variant.vt == VT_BSTR ? variant.bstrVal : L"Get Total memory returned wrong data type";

}

std::wstring WMIAPI::getAvailableMemory() {    
    auto variant = ExecQuery(L"SELECT * FROM Win32_OperatingSystem ", L"FreePhysicalMemory");

    return variant.vt == VT_BSTR ? variant.bstrVal : L"Get Available memory returned wrong data type";
}

std::wstring WMIAPI::getOSName() {

   auto variant = ExecQuery(L"SELECT Caption FROM Win32_OperatingSystem", L"Caption");
   
   return variant.vt == VT_BSTR ? variant.bstrVal : L"Get OS Name returned wrong data type";
}

uint32_t WMIAPI::getMemClockSpeed() {
    auto variant = ExecQuery(L"SELECT ConfiguredClockSpeed FROM Win32_PhysicalMemory", L"ConfiguredClockSpeed");
    auto memType = getMemType();
    //checking if DDR for frequency calculation
    auto multiplier = (memType >= 20 && memType <=24) ? 2 : 1;
    return variant.vt == VT_I4 ? variant.intVal * multiplier : 0;
}

int32_t WMIAPI::getMemType() {
    auto variant = ExecQuery(L"SELECT MemoryType FROM Win32_PhysicalMemory", L"MemoryType");
    return variant.vt == VT_I4 ? variant.intVal : 0;
}

int32_t WMIAPI::getMemVoltage() {
    auto variant = ExecQuery(L"SELECT ConfiguredVoltage FROM Win32_PhysicalMemory", L"ConfiguredVoltage");
    return variant.vt == VT_I4 ? variant.ulVal : 0;
}