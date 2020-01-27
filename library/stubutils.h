#pragma once
#include <functional>
#include <vector>
#include <string>

namespace mapistub
{
	extern std::function<void(LPCWSTR szMsg, va_list argList)> logLoadMapiCallback;
	extern std::function<void(LPCWSTR szMsg, va_list argList)> logLoadLibraryCallback;

	std::wstring GetSystemDirectory();

	_Check_return_ HMODULE LoadFromSystemDir(_In_ const std::wstring& szDLLName);
	HMODULE& GetHModMSI();
	HMODULE& GetHModKernel32();

	// Loads szModule at the handle given by hModule, then looks for szEntryPoint.
	// Will not load a module or entry point twice
	void LoadProc(_In_ const std::wstring& szModule, HMODULE& hModule, LPCSTR szEntryPoint, FARPROC& lpfn);

	extern volatile ULONG g_ulDllSequenceNum;
	// Keep this in sync with g_pszOutlookQualifiedComponents
	enum officeComponent
	{
		oqcOffice16 = 0,
		oqcOffice15 = 1,
		oqcOffice14 = 2,
		oqcOffice12 = 3,
		oqcOffice11 = 4,
		oqcOffice11Debug = 5
	};

	HMODULE GetMAPIHandle() noexcept;
	void UnloadPrivateMAPI();
	void ForceOutlookMAPI(bool fForce);
	void ForceSystemMAPI(bool fForce);
	void SetMAPIHandle(HMODULE hinstMAPI);
	HMODULE GetPrivateMAPI();
	std::wstring GetComponentPath(const std::wstring& szComponent, const std::wstring& szQualifier, bool fInstall);
	extern std::vector<std::wstring> g_pszOutlookQualifiedComponents;
	std::vector<std::wstring> GetMAPIPaths();
	// Looks up Outlook's path given its qualified component guid
	std::wstring GetOutlookPath(_In_ const std::wstring& szCategory, _Out_opt_ bool* lpb64);
	std::wstring GetInstalledOutlookMAPI(int iOutlook);
	std::wstring GetMAPISystemDir();
} // namespace mapistub
