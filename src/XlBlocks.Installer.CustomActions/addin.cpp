#include "pch.h"

UINT __stdcall GetUserProgramFilesPath(__in MSIHANDLE hInstall)
{
    HRESULT hr = S_OK;
    DWORD er = ERROR_SUCCESS;

    PWSTR path = nullptr;

    hr = WcaInitialize(hInstall, "GetUserProgramFilesPath");
    ExitOnFailure(hr, "Failed to initialize");

    WcaLog(LOGMSG_STANDARD, "GetUserProgramFilesPath - initialized.");

    hr = SHGetKnownFolderPath(FOLDERID_UserProgramFiles, 0, NULL, &path);
    ExitOnFailure(hr, "Failed to get known folder path");

    hr = WcaSetProperty(L"USER_PROGRAM_FILES_PATH", path);
    ExitOnFailure(hr, "Failed to set property USER_PROGRAM_FILES_PATH");

    WcaLog(LOGMSG_STANDARD, "GetUserProgramFilesPath - USER_PROGRAM_FILES_PATH set to: %ls", path);

LExit:
    if (path) 
    { 
        CoTaskMemFree(path); 
    }

    er = SUCCEEDED(hr) ? ERROR_SUCCESS : ERROR_INSTALL_FAILURE;
    return WcaFinalize(er);
}

UINT __stdcall GetOfficeInfo(
    __in MSIHANDLE hInstall)
{
    HRESULT hr = S_OK;
    DWORD er = ERROR_SUCCESS;

    HKEY hKeyOffice = NULL;
    HKEY hKeyExcelPath = NULL;
    WCHAR szRegKeyPath[MAX_PATH] = { 0 };
    WCHAR szpwOfficeVersion[16] = { 0 };
    WCHAR szExcelPath[MAX_PATH] = { 0 };
    WCHAR szOfficeBitness[16] = { 0 };
    DWORD dwType = 0;
    DWORD dwSize = sizeof(szExcelPath);
    BOOL bFoundOffice = false;
    INT iOfficeVersions[5] = { 11, 15, 14, 12, 16 }; // Office 2016+, 2013, 2010, 2007, 2003

    hr = WcaInitialize(hInstall, "GetOfficeInfo");
    ExitOnFailure(hr, "Failed to initialize");

    WcaLog(LOGMSG_STANDARD, "GetOfficeInfo initialized.");

    // Iterate through possible Office versions
    for (int i = 0; i < sizeof(iOfficeVersions) / sizeof(iOfficeVersions[0]); i++)
    {
        StringCchPrintf(szRegKeyPath, MAX_PATH, L"Software\\Microsoft\\Office\\%d.0\\Common\\InstallRoot", iOfficeVersions[i]);
        er = RegOpenKeyEx(HKEY_LOCAL_MACHINE, szRegKeyPath, 0, KEY_READ, &hKeyOffice);

        if (er == ERROR_SUCCESS)
        {
            // Office version found
            WcaLog(LOGMSG_STANDARD, "Found office version %d.0", iOfficeVersions[i]);
            StringCchPrintf(szpwOfficeVersion, sizeof(szpwOfficeVersion) / sizeof(szpwOfficeVersion[0]), L"%d", iOfficeVersions[i]);
            bFoundOffice = true;
            break;
        }
    }

    if (!bFoundOffice)
    {
        WcaLog(LOGMSG_STANDARD, "No compatible Office version found.");
        ExitFunction();
    }

    hr = WcaSetProperty(L"OFFICE_VERSION", szpwOfficeVersion);
    ExitOnFailure(hr, "Failed to set OFFICE_VERSION property");

    // Determine Office bitness by looking at excel exe itself vs trying to iterate the various registry possibilities
    StringCchPrintf(szRegKeyPath, MAX_PATH, L"Software\\Microsoft\\Windows\\CurrentVersion\\App Paths\\excel.exe");
    er = RegOpenKeyEx(HKEY_LOCAL_MACHINE, szRegKeyPath, 0, KEY_READ, &hKeyExcelPath);

    // Default to 32-bit office
    StringCchCopy(szOfficeBitness, sizeof(szOfficeBitness) / sizeof(szOfficeBitness[0]), L"x86");

    if (er != ERROR_SUCCESS)
    {
        WcaLog(LOGMSG_STANDARD, "No registered excel.exe path was found, unable to determine bitness of installed Office");
        ExitFunction();
    }

    er = RegQueryValueEx(hKeyExcelPath, NULL, NULL, &dwType, (LPBYTE)szExcelPath, &dwSize);
    if (er != ERROR_SUCCESS || dwType != REG_SZ)
    {
        WcaLog(LOGMSG_STANDARD, "Could not find excel.exe at registered path, unable to determine bitness of installed Office");
        ExitFunction();
    }

    WcaLog(LOGMSG_STANDARD, "excel.exe path was found in registry: '%ls'", szExcelPath);
    DWORD dwBinaryType;
    if (GetBinaryType(szExcelPath, &dwBinaryType))
    {
        if (SCS_64BIT_BINARY == dwBinaryType)
        {
            // Detected 64-bit Office
            StringCchCopy(szOfficeBitness, sizeof(szOfficeBitness) / sizeof(szOfficeBitness[0]), L"x64");
        }
    }

LExit:
    if (hKeyOffice)
    {
        RegCloseKey(hKeyOffice);
    }

    if (hKeyExcelPath)
    {
        RegCloseKey(hKeyExcelPath);
    }

    if (wcslen(szOfficeBitness) > 0)
    {
        hr = WcaSetProperty(L"OFFICE_BITNESS", szOfficeBitness);
    }

    WcaLog(LOGMSG_STANDARD, "OFFICE_VERSION set to: %ls", szpwOfficeVersion);
    WcaLog(LOGMSG_STANDARD, "OFFICE_BITNESS set to: %ls", szOfficeBitness);

    er = SUCCEEDED(hr) ? ERROR_SUCCESS : ERROR_INSTALL_FAILURE;
    return WcaFinalize(er);
}

UINT __stdcall RegisterAddIn(__in MSIHANDLE hInstall)
{
    HRESULT hr = S_OK; 
    DWORD er = ERROR_SUCCESS; 
    
    HKEY hKey = NULL; 
    PWSTR pwpwOfficeVersion = nullptr;
    PWSTR pwAddInFolder = nullptr;
    PWSTR pwAddInFileName = nullptr;

    WCHAR szRegKeyPath[MAX_PATH] = { 0 }; 
    WCHAR szValueName[32] = { 0 };
    WCHAR szExistingValue[MAX_PATH] = { 0 };
    WCHAR szAddInFullPath[MAX_PATH] = { 0 };
    DWORD dwIndex = 0;
    DWORD dwSize = sizeof(szExistingValue);
    DWORD dwType = 0;
    BOOL bFoundExisting = false;

    hr = WcaInitialize(hInstall, "RegisterAddIn");
    ExitOnFailure(hr, "Failed to initialize");

    WcaLog(LOGMSG_STANDARD, "RegisterAddIn initialized.");

    hr = WcaGetProperty(L"OFFICE_VERSION", &pwpwOfficeVersion);
    ExitOnFailure(hr, "Failed to get OFFICE_VERSION property");

    hr = WcaGetProperty(L"Directory_AddIn", &pwAddInFolder);
    ExitOnFailure(hr, "Failed to get Directory_AddIn property");

    hr = WcaGetProperty(L"ADDIN_FILE_NAME", &pwAddInFileName);
    ExitOnFailure(hr, "Failed to get ADDIN_FILE_NAME property");

    StringCchPrintf(szRegKeyPath, MAX_PATH, L"Software\\Microsoft\\Office\\%s.0\\Excel\\Options", pwpwOfficeVersion);
    er = RegOpenKeyEx(HKEY_CURRENT_USER, szRegKeyPath, 0, KEY_READ | KEY_WRITE, &hKey); 
    ExitOnFailure(hr = HRESULT_FROM_WIN32(er), "Failed to open registry key");

    while (true && dwIndex < 100)
    {
        if (dwIndex > 0)
        {
            StringCchPrintf(szValueName, sizeof(szValueName) / sizeof(szValueName[0]), L"OPEN%d", dwIndex);
        }
        else
        {
            StringCchPrintf(szValueName, sizeof(szValueName) / sizeof(szValueName[0]), L"OPEN");
        }

        er = RegQueryValueEx(hKey, szValueName, NULL, &dwType, (LPBYTE)szExistingValue, &dwSize);
        if (er == ERROR_FILE_NOT_FOUND)
        {

            // No more OPEN keys
            break;
        }

        if (er == ERROR_SUCCESS && dwType == REG_SZ && wcsstr(szExistingValue, pwAddInFileName) != NULL)
        {
            // Found existing addin
            bFoundExisting = true;
            WcaLog(LOGMSG_STANDARD, "Existing add-in registration was found with value %ls and path '%ls'", szValueName, szExistingValue);
            break;
        }
        dwIndex++;
    }

    if (!bFoundExisting)
    {
        StringCchPrintf(szAddInFullPath, MAX_PATH, L"%s%s", pwAddInFolder, pwAddInFileName);
        er = RegSetValueEx(hKey, szValueName, 0, REG_SZ, (LPBYTE)szAddInFullPath, (lstrlen(szAddInFullPath) + 1) * sizeof(WCHAR));
        ExitOnFailure(hr = HRESULT_FROM_WIN32(er), "Failed to set registry value");
        WcaLog(LOGMSG_STANDARD, "Add-in registered under %ls with value %ls", szRegKeyPath, pwAddInFolder);
    }

LExit:
    if (hKey)
        RegCloseKey(hKey);
    
    if (pwpwOfficeVersion)
        CoTaskMemFree(pwpwOfficeVersion);

    if (pwAddInFolder)
        CoTaskMemFree(pwAddInFolder);

    if (pwAddInFileName)
        CoTaskMemFree(pwAddInFileName);

    er = SUCCEEDED(hr) ? ERROR_SUCCESS : ERROR_INSTALL_FAILURE;
    return WcaFinalize(er);
}

UINT __stdcall UnregisterAddIn(__in MSIHANDLE hInstall)
{
    HRESULT hr = S_OK;
    DWORD er = ERROR_SUCCESS;

    HKEY hKey = NULL;
    PWSTR pwOfficeVersion = nullptr;
    PWSTR pwAddInFolder = nullptr;
    PWSTR pwAddInFileName = nullptr;

    WCHAR szRegKeyPath[MAX_PATH] = { 0 };
    WCHAR szValueName[32] = { 0 };
    WCHAR szExistingValue[MAX_PATH] = { 0 };
    WCHAR szAddInFullPath[MAX_PATH] = { 0 };
    DWORD dwIndex = 0;
    DWORD dwSize = sizeof(szExistingValue);
    DWORD dwType = 0;

    hr = WcaInitialize(hInstall, "UnregisterAddIn");
    ExitOnFailure(hr, "Failed to initialize");

    WcaLog(LOGMSG_STANDARD, "UnregisterAddIn initialized.");

    hr = WcaGetProperty(L"OFFICE_VERSION", &pwOfficeVersion);
    ExitOnFailure(hr, "Failed to get OFFICE_VERSION property");

    hr = WcaGetProperty(L"Directory_AddIn", &pwAddInFolder);
    ExitOnFailure(hr, "Failed to get Directory_AddIn property");

    hr = WcaGetProperty(L"ADDIN_FILE_NAME", &pwAddInFileName);
    ExitOnFailure(hr, "Failed to get ADDIN_FILE_NAME property");

    StringCchPrintf(szRegKeyPath, MAX_PATH, L"Software\\Microsoft\\Office\\%s.0\\Excel\\Options", pwOfficeVersion);
    er = RegOpenKeyEx(HKEY_CURRENT_USER, szRegKeyPath, 0, KEY_READ | KEY_WRITE, &hKey);
    ExitOnFailure(hr = HRESULT_FROM_WIN32(er), "Failed to open registry key");

    while (true && dwIndex < 100)
    {
        if (dwIndex > 0)
        {
            StringCchPrintf(szValueName, sizeof(szValueName) / sizeof(szValueName[0]), L"OPEN%d", dwIndex);
        }
        else
        {
            StringCchPrintf(szValueName, sizeof(szValueName) / sizeof(szValueName[0]), L"OPEN");
        }

        er = RegQueryValueEx(hKey, szValueName, NULL, &dwType, (LPBYTE)szExistingValue, &dwSize);
        if (er == ERROR_FILE_NOT_FOUND)
        {
            // No more OPEN keys
            break;
        }

        if (er == ERROR_SUCCESS && dwType == REG_SZ && wcsstr(szExistingValue, pwAddInFileName) != NULL)
        {
            // Found the matching entry, delete it
            er = RegDeleteValue(hKey, szValueName);
            ExitOnFailure(hr = HRESULT_FROM_WIN32(er), "Failed to delete registry value");

            WcaLog(LOGMSG_STANDARD, "Add-in unregistered from %ls with value %ls", szRegKeyPath, pwAddInFolder);
            break;
        }
        dwIndex++;
    }

LExit:
    if (hKey)
        RegCloseKey(hKey);
    
    if (pwOfficeVersion)
        CoTaskMemFree(pwOfficeVersion);

    if (pwAddInFolder)
        CoTaskMemFree(pwAddInFolder);

    if (pwAddInFileName)
        CoTaskMemFree(pwAddInFileName);

    er = SUCCEEDED(hr) ? ERROR_SUCCESS : ERROR_INSTALL_FAILURE;
    return WcaFinalize(er);
}

UINT __stdcall CheckIfExcelRunning(__in MSIHANDLE hInstall)
{
    HRESULT hr = S_OK;
    DWORD er = ERROR_SUCCESS;

    BOOL bExcelRunning = FALSE;
    HANDLE hProcessSnap = NULL;
    PROCESSENTRY32 pe32 = { 0 };

    hr = WcaInitialize(hInstall, "CheckIfExcelRunning");
    ExitOnFailure(hr, "Failed to initialize");

    WcaLog(LOGMSG_STANDARD, "CheckIfExcelRunning initialized.");

    // Get snapshot of all running processes
    hProcessSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
    if (hProcessSnap == INVALID_HANDLE_VALUE)
    {
        hr = HRESULT_FROM_WIN32(GetLastError());
        ExitOnFailure(hr, "Failed to take process snapshot");
    }

    pe32.dwSize = sizeof(PROCESSENTRY32);
    if (!Process32First(hProcessSnap, &pe32))
    {
        hr = HRESULT_FROM_WIN32(GetLastError());
        ExitOnFailure(hr, "Failed to retrieve first process information");
    }

    do
    {
        // Check for "EXCEL.EXE" process
        if (lstrcmpi(pe32.szExeFile, L"EXCEL.EXE") == 0)
        {
            bExcelRunning = TRUE;
            break;
        }
    } while (Process32Next(hProcessSnap, &pe32));

LExit:
    if (hProcessSnap != INVALID_HANDLE_VALUE)
    {
        CloseHandle(hProcessSnap);
    }

    hr = WcaSetProperty(L"EXCEL_RUNNING", bExcelRunning ? L"1" : L"0");
    WcaLog(LOGMSG_STANDARD, "EXCEL_RUNNING set to: %ls", bExcelRunning ? L"1" : L"0");

    er = SUCCEEDED(hr) ? ERROR_SUCCESS : ERROR_INSTALL_FAILURE;
    return WcaFinalize(er);
}
