#include <windows.h>
#include <strsafe.h>
#include <msi.h>
#include <winreg.h>
#include <stdlib.h>


/*
 *  MAPI Stub Utilities
 *
 *	Public Functions:
 *		
 *		GetPrivateMAPI()
 *			Obtain a handle to the MAPI DLL.  This function will load the MAPI DLL
 *			if it hasn't already been loaded
 *
 *		UnLoadPrivateMAPI()
 *			Forces the MAPI DLL to be unloaded.  This can cause problems if the code
 *			still has outstanding allocated MAPI memory, or unmatched calls to 
 *			MAPIInitialize/MAPIUninitialize
 *
 *		ForceOutlookMAPI()
 *			Instructs the stub code to always try loading the Outlook version of MAPI
 *			on the system, instead of respecting the system MAPI registration
 *			(HKLM\Software\Clients\Mail). This call must be made prior to any MAPI 
 *			function calls.
 */
HMODULE GetPrivateMAPI();
void UnLoadPrivateMAPI();
void ForceOutlookMAPI();


const WCHAR WszKeyNameMailClient[] = L"Software\\Clients\\Mail";
const WCHAR WszValueNameDllPathEx[] = L"DllPathEx";
const WCHAR WszValueNameDllPath[] = L"DllPath";

const CHAR SzValueNameMSI[] = "MSIComponentID";
const CHAR SzValueNameLCID[] = "MSIApplicationLCID";

const WCHAR WszOutlookMapiClientName[] = L"Microsoft Outlook";

const WCHAR WszMAPISystemPath[] = L"%s\\%s";

static const WCHAR WszPrivateMAPI[] = L"olmapi32.dll";
static const WCHAR WszPrivateMAPI_11[] = L"msmapi32.dll";
static const WCHAR WszMapi32[] = L"mapi32.dll";
static const WCHAR WszMapiStub[] = L"mapistub.dll";

static const CHAR SzFGetComponentPath[] = "FGetComponentPath";

// Sequence number which is incremented every time we set our MAPI handle which will
//  cause a re-fetch of all stored function pointers
volatile ULONG g_ulDllSequenceNum = 1;

// Whether or not we should ignore the system MAPI registration and always try to find
//  Outlook and its MAPI DLLs
static bool s_fForceOutlookMAPI = false;

static volatile HMODULE		g_hinstMAPI = NULL;

__inline HMODULE GetMAPIHandle()
{
	return g_hinstMAPI;
}

void SetMAPIHandle(HMODULE hinstMAPI)
{
	HMODULE	hinstNULL = NULL;
	HMODULE	hinstToFree = NULL;

	if (hinstMAPI == NULL)
	{
		hinstToFree = (HMODULE)InterlockedExchangePointer((PVOID*)&g_hinstMAPI, hinstNULL);
	}
	else
	{
		// Set the value only if the global is NULL
		HMODULE	hinstPrev;
		hinstPrev = (HMODULE)InterlockedCompareExchangePointer(reinterpret_cast<volatile PVOID*>(&g_hinstMAPI), hinstMAPI, hinstNULL);
		if (NULL != hinstPrev)
		{
			hinstToFree = hinstMAPI;
		}

		// If we've updated our MAPI handle, any previous addressed fetched via GetProcAddress are invalid, so we
		// have to increment a sequence number to signal that they need to be re-fetched
		InterlockedIncrement(reinterpret_cast<volatile LONG*>(&g_ulDllSequenceNum));
	}
	if (NULL != hinstToFree)
	{
		FreeLibrary(hinstToFree);
	}
}

/*
 *  RegQueryWszExpand
 *		Wrapper for RegQueryValueExW which automatically expands REG_EXPAND_SZ values
 */
DWORD RegQueryWszExpand(HKEY hKey, LPCWSTR lpValueName, LPWSTR lpValue, DWORD cchValueLen)
{
	DWORD dwErr = ERROR_SUCCESS;
	DWORD dwType = 0;

	WCHAR rgchValue[MAX_PATH];
	DWORD dwSize = sizeof(rgchValue);

	dwErr = RegQueryValueExW(hKey, lpValueName, 0, &dwType, (LPBYTE) &rgchValue, &dwSize);

	if (dwErr == ERROR_SUCCESS)
	{
		if (dwType == REG_EXPAND_SZ)
		{
			// Expand the strings
			DWORD cch = ExpandEnvironmentStringsW(rgchValue, lpValue, cchValueLen);
			if ((0 == cch) || (cch > cchValueLen))
			{
				dwErr = ERROR_INSUFFICIENT_BUFFER;
				goto Exit;
			}
		}
		else if (dwType == REG_SZ)
		{
			wcscpy_s(lpValue, cchValueLen, rgchValue);
		}
	}
Exit:
	return dwErr;
}

/*
 *  GetComponentPath
 *		Wrapper around mapi32.dll->FGetComponentPath which maps an MSI component ID to 
 *		a DLL location from the default MAPI client registration values
 *
 */
BOOL GetComponentPath(LPCSTR szComponent, LPSTR szQualifier, LPSTR szDllPath, DWORD cchBufferSize, BOOL fInstall)
{
	HMODULE hMapiStub = NULL;
	BOOL fReturn = FALSE;

	typedef BOOL (STDAPICALLTYPE *FGetComponentPathType)(LPCSTR, LPSTR, LPSTR, DWORD, BOOL);

	hMapiStub = LoadLibraryW(WszMapi32);
	if (!hMapiStub)
		hMapiStub = LoadLibraryW(WszMapiStub);

	if (hMapiStub)
	{
		FGetComponentPathType pFGetCompPath = (FGetComponentPathType)GetProcAddress(hMapiStub, SzFGetComponentPath);

		fReturn = pFGetCompPath(szComponent, szQualifier, szDllPath, cchBufferSize, fInstall);

		FreeLibrary(hMapiStub);
	}

	return fReturn;
}


/*
 *  LoadMailClientFromMSIData
 *		Attempt to locate the MAPI provider DLL via HKLM\Software\Clients\Mail\(provider)\MSIComponentID
 */
HMODULE LoadMailClientFromMSIData(HKEY hkeyMapiClient)
{
	HMODULE		hinstMapi = NULL;
	CHAR		rgchMSIComponentID[MAX_PATH];
	CHAR		rgchMSIApplicationLCID[MAX_PATH];
	CHAR		rgchComponentPath[MAX_PATH];
	DWORD		dwType = 0;

	DWORD dwSizeComponentID = sizeof(rgchMSIComponentID);
	DWORD dwSizeLCID = sizeof(rgchMSIApplicationLCID);

	if (ERROR_SUCCESS == RegQueryValueExA(	hkeyMapiClient, SzValueNameMSI,0,
											&dwType, (LPBYTE) &rgchMSIComponentID, &dwSizeComponentID)
		&& ERROR_SUCCESS == RegQueryValueExA(	hkeyMapiClient, SzValueNameLCID,0,
												&dwType, (LPBYTE) &rgchMSIApplicationLCID, &dwSizeLCID))
	{
		if (GetComponentPath(rgchMSIComponentID, rgchMSIApplicationLCID, 
			rgchComponentPath, _countof(rgchComponentPath), FALSE))
		{
			hinstMapi = LoadLibraryA(rgchComponentPath);
		}
	}
	return hinstMapi;
}

/*
 *  LoadMailClientFromDllPath
 *		Attempt to locate the MAPI provider DLL via HKLM\Software\Clients\Mail\(provider)\DllPathEx
 */
HMODULE LoadMailClientFromDllPath(HKEY hkeyMapiClient)
{
	HMODULE		hinstMapi = NULL;
	WCHAR		rgchDllPath[MAX_PATH];

	DWORD dwSizeDllPath = _countof(rgchDllPath);

	if (ERROR_SUCCESS == RegQueryWszExpand(	hkeyMapiClient, WszValueNameDllPathEx, rgchDllPath, dwSizeDllPath))
	{
		hinstMapi = LoadLibraryW(rgchDllPath);
	}

	if (!hinstMapi)
	{
		dwSizeDllPath = _countof(rgchDllPath);
		if (ERROR_SUCCESS == RegQueryWszExpand(	hkeyMapiClient, WszValueNameDllPath, rgchDllPath, dwSizeDllPath))
		{
			hinstMapi = LoadLibraryW(rgchDllPath);
		}
	}
	return hinstMapi;
}

/*
 *  LoadRegisteredMapiClient
 *		Read the registry to discover the registered MAPI client and attempt to load its MAPI DLL.
 *		
 *		If wzOverrideProvider is specified, this function will load that MAPI Provider instead of the 
 *		currently registered provider
 */
HMODULE LoadRegisteredMapiClient(LPCWSTR pwzProviderOverride)
{
	HMODULE		hinstMapi = NULL;
	DWORD		dwType;
	HKEY 		hkey = NULL, hkeyMapiClient = NULL;
	WCHAR		rgchMailClient[MAX_PATH];
	LPCWSTR		pwzProvider = pwzProviderOverride;

	// Open HKLM\Software\Clients\Mail
	if (ERROR_SUCCESS == RegOpenKeyExW( HKEY_LOCAL_MACHINE,
										WszKeyNameMailClient,
										0,
										KEY_READ,
										&hkey))
	{
		// If a specific provider wasn't specified, load the name of the default MAPI provider
		if (!pwzProvider)
		{
			// Get Outlook application path registry value
			DWORD dwSize = sizeof(rgchMailClient);
			if SUCCEEDED(RegQueryValueExW(	hkey,
				NULL,
				0,
				&dwType,
				(LPBYTE) &rgchMailClient,
				&dwSize))

				if (dwType != REG_SZ)
				goto Error;

			pwzProvider = rgchMailClient;
		}

		if (pwzProvider)
		{
			if SUCCEEDED(RegOpenKeyExW( hkey,
				pwzProvider,
				0,
				KEY_READ,
				&hkeyMapiClient))
			{
				hinstMapi = LoadMailClientFromMSIData(hkeyMapiClient);

				if (!hinstMapi)
					hinstMapi = LoadMailClientFromDllPath(hkeyMapiClient);
			}
		}
	}

Error:
	return hinstMapi;
}

/*
 *  LoadMAPIFromSystemDir
 *		Fall back for loading System32\Mapi32.dll if all else fails
 */
HMODULE LoadMAPIFromSystemDir()
{
	WCHAR szSystemDir[MAX_PATH] = {0};

	if (GetSystemDirectoryW(szSystemDir, MAX_PATH))
	{
		WCHAR szDLLPath[MAX_PATH] = {0};
		swprintf_s(szDLLPath, _countof(szDLLPath), WszMAPISystemPath, szSystemDir, WszMapi32);
		return LoadLibraryW(szDLLPath);
	}

	return NULL;
}


HMODULE GetDefaultMapiHandle()
{
	HMODULE	hinstMapi = NULL;

	// Try to respect the machine's default MAPI client settings.  If the active MAPI provider
	//  is Outlook, don't load and instead run the logic below
	if (s_fForceOutlookMAPI)
		hinstMapi = LoadRegisteredMapiClient(WszOutlookMapiClientName);
	else
		hinstMapi = LoadRegisteredMapiClient(NULL);

	// If MAPI still isn't loaded, load the stub from the system directory
	if (!hinstMapi && !s_fForceOutlookMAPI)
	{
		hinstMapi = LoadMAPIFromSystemDir();
	}

	return hinstMapi;
}


/*------------------------------------------------------------------------------
    Attach to wzMapiDll(olmapi32.dll/msmapi32.dll) if it is already loaded in the
	current process.
------------------------------------------------------------------------------*/
HMODULE AttachToMAPIDll(const WCHAR *wzMapiDll)
{
	HMODULE	hinstPrivateMAPI = NULL;
	GetModuleHandleExW(0UL, wzMapiDll, &hinstPrivateMAPI);
	return hinstPrivateMAPI;
}


void UnLoadPrivateMAPI()
{
	HMODULE hinstPrivateMAPI = NULL;

	hinstPrivateMAPI = GetMAPIHandle();
	if (NULL != hinstPrivateMAPI)
	{
		SetMAPIHandle(NULL);
	}
}

void ForceOutlookMAPI(bool fForce)
{
	s_fForceOutlookMAPI = fForce;
}


HMODULE GetPrivateMAPI()
{
	HMODULE hinstPrivateMAPI = GetMAPIHandle();

	if (NULL == hinstPrivateMAPI)
	{
		// First, try to attach to olmapi32.dll if it's loaded in the process
		hinstPrivateMAPI = AttachToMAPIDll(WszPrivateMAPI);

		// If that fails try msmapi32.dll, for Outlook 11 and below
		//  Only try this in the static lib, otherwise msmapi32.dll will attach to itself.
		if (NULL == hinstPrivateMAPI)
		{
			hinstPrivateMAPI = AttachToMAPIDll(WszPrivateMAPI_11);
		}

		// If MAPI isn't loaded in the process yet, then find the path to the DLL and
		// load it manually.
		if (NULL == hinstPrivateMAPI)
		{
			hinstPrivateMAPI = GetDefaultMapiHandle();
		}

		if (NULL != hinstPrivateMAPI)
		{
			SetMAPIHandle(hinstPrivateMAPI);
		}

		// Reason - if for any reason there is an instance already loaded, SetMAPIHandle()
		// will free the new one and reuse the old one
		// So we fetch the instance from the global again
		return GetMAPIHandle();
	}

	return hinstPrivateMAPI;
}
