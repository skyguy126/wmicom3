#pragma comment(lib, "wbemuuid.lib")

#define _WIN32_DCOM
#define WMICOM3_EXPORTS

#include "stdafx.h"
#include "wmicom3.h"
#include <regex>
#include <iostream>
#include <comdef.h>
#include <Wbemidl.h>

using namespace std;

extern "C" WMICOM3_API int wmicom(char *id)
{
	int port = -1;
	HRESULT hres;

	// Initialize COM.
	hres = CoInitializeEx(0, COINIT_MULTITHREADED);
	if (FAILED(hres))
	{
		cout << "Failed to initialize COM library. "
			<< "Error code = 0x"
			<< hex << hres << endl;
		return 1;
	}

	// Initialize 
	hres = CoInitializeSecurity(
		NULL,
		-1,      // COM negotiates service                  
		NULL,    // Authentication services
		NULL,    // Reserved
		RPC_C_AUTHN_LEVEL_DEFAULT,    // authentication
		RPC_C_IMP_LEVEL_IMPERSONATE,  // Impersonation
		NULL,             // Authentication info 
		EOAC_NONE,        // Additional capabilities
		NULL              // Reserved
	);


	if (FAILED(hres))
	{
		cout << "Failed to initialize security. "
			<< "Error code = 0x"
			<< hex << hres << endl;
		CoUninitialize();
		return -1;
	}

	// Obtain the initial locator to Windows Management
	// on a particular host computer.
	IWbemLocator *pLoc = 0;

	hres = CoCreateInstance(
		CLSID_WbemLocator,
		0,
		CLSCTX_INPROC_SERVER,
		IID_IWbemLocator, (LPVOID *)&pLoc
	);

	if (FAILED(hres))
	{
		cout << "Failed to create IWbemLocator object. "
			<< "Error code = 0x"
			<< hex << hres << endl;
		CoUninitialize();
		return -1;
	}

	IWbemServices *pSvc = 0;

	// Connect to the root\cimv2 namespace with the
	// current user and obtain pointer pSvc
	// to make IWbemServices calls.

	hres = pLoc->ConnectServer(
		_bstr_t(L"ROOT\\CIMV2"), // WMI namespace
		NULL,                    // User name
		NULL,                    // User password
		0,                       // Locale
		NULL,                    // Security flags                 
		0,                       // Authority       
		0,                       // Context object
		&pSvc                    // IWbemServices proxy
	);

	if (FAILED(hres))
	{
		cout << "Could not connect. Error code = 0x"
			<< hex << hres << endl;
		pLoc->Release();
		CoUninitialize();
		return -1;
	}

	// Set the IWbemServices proxy so that impersonation
	// of the user (client) occurs.
	hres = CoSetProxyBlanket(
		pSvc,                         // the proxy to set
		RPC_C_AUTHN_WINNT,            // authentication service
		RPC_C_AUTHZ_NONE,             // authorization service
		NULL,                         // Server principal name
		RPC_C_AUTHN_LEVEL_CALL,       // authentication level
		RPC_C_IMP_LEVEL_IMPERSONATE,  // impersonation level
		NULL,                         // client identity 
		EOAC_NONE                     // proxy capabilities     
	);

	if (FAILED(hres))
	{
		cout << "Could not set proxy blanket. Error code = 0x"
			<< hex << hres << endl;
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();
		return -1;
	}

	string base1 = "SELECT * FROM Win32_SerialPort WHERE PNPDeviceID LIKE '%";
	string idnum(id);
	string base2 = "%'";
	string query = base1 + idnum + base2;

	// For example, query for all the running processes
	IEnumWbemClassObject* pEnumerator = NULL;
	hres = pSvc->ExecQuery(
		bstr_t("WQL"),
		bstr_t(query.c_str()),
		WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
		NULL,
		&pEnumerator
	);

	if (FAILED(hres))
	{
		cout << "Query for processes failed. "
			<< "Error code = 0x"
			<< hex << hres << endl;
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();
		return -1;
	}

	IWbemClassObject *pclsObj;
	ULONG uReturn = 0;

	while (pEnumerator) 
	{

		hres = pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);

		if (uReturn == 0)
		{
			pSvc->Release();
			pLoc->Release();
			CoUninitialize();
			return -1;
		}

		VARIANT vtProp;

		// Get the value of the Name property
		hres = pclsObj->Get(L"Name", 0, &vtProp, 0, 0);
		if (!vtProp.bstrVal == NULL)
		{
			wstring deviceName(vtProp.bstrVal, SysStringLen(vtProp.bstrVal));
			wregex regex(L"[\\d][\\d]?[\\d]?");
			wsmatch match;
			regex_search(deviceName, match, regex);

			if (match.size() == 1)
			{
				port = stoi(match[0].str());
				break;
			}
		}

		VariantClear(&vtProp);
	}

	// Cleanup
	pSvc->Release();
	pLoc->Release();
	CoUninitialize();

	return port;
}
