#include "vbt.h"

/* todo: remove vbt_printf statements */
/* todo: Change generic ResultFromScode(E_FAIL) to better message */
/* where should new .tlb files be created? */

/* Set the given key and its value. */
HRESULT setKeyAndValue(LPCOLESTR pwPath, LPCOLESTR wSubkey, LPCOLESTR wValue);

/* Open a key and set a value. */
HRESULT setValueInKey(LPCOLESTR wKey, LPCOLESTR wNamedValue, LPCOLESTR wValue);

/* Delete wKeyChild and all of its descendents. */
HRESULT recursiveDeleteKey(HKEY hKeyParent, LPCOLESTR wKeyChild);

/* Size of a CLSID as a string */
#define CLSID_STRING_SIZE 39

/* Register the component in the registry. */
HRESULT RegisterServer(char * const szModuleName,  /* DLL module handle */
			BOOL InProcServer,	   /* Is this a local server? */
			REFCLSID clsid,            /* Class ID */
			LPCOLESTR wFriendlyName,   /* Friendly Name */
			LPCOLESTR wVerIndProgID,   /* Programmatic */
			LPCOLESTR wProgID,         /* IDs */
			REFGUID libid,		   /* TypeLib GUID */
			unsigned short libMajor,   /* TypeLib Major version */
			unsigned short libMinor,   /* TypeLib Minor version */
			LPCOLESTR wThreadingModel) /* ThreadingModel */
{
	/* Get server location. */
	OLECHAR wModule[512];
	HMODULE hModule;
	DWORD dwResult;
	OLECHAR wCLSID[64];
	OLECHAR wKey[64];
	HRESULT hr;

/* szModuleName can be NULL (creating executable) or wargv[0] or dll name */
	hModule = GetModuleHandleA(szModuleName);
vbt_printf("%s:%d h=%lx\n",__FILE__,__LINE__,hModule);
	if(hModule == NULL)
		return ResultFromScode(E_FAIL); /* need better error message */
	dwResult = GetModuleFileName(NULL, wModule, sizeof(wModule)/sizeof(OLECHAR));
vbt_printf("%s:%d dw=%lx\n",__FILE__,__LINE__,dwResult);
	if(dwResult == 0)
#ifdef NEVER
		return ResultFromScode(E_FAIL);
#else
wcscpy(wModule,L"project2");
#endif

	/* Convert a CLSID GUID into a wchar. */
 	hr = StringFromGUID2(clsid, wCLSID, sizeof(wCLSID));
vbt_printf("%s:%d\n",__FILE__,__LINE__);
	if(FAILED(hr))
		return hr;

	/* Build the key CLSID\\{...} */
	wcscpy(wKey, L"CLSID\\");
	wcscat(wKey, wCLSID);
  
	/* Add the CLSID to the registry. */
	hr = setKeyAndValue(wKey, NULL, wFriendlyName);
vbt_printf("%s:%d\n",__FILE__,__LINE__);
	if(FAILED(hr))
		return hr;

#ifdef NEVER /* when is this used? */
	hr = setKeyAndValue(wKey, L"AppID", wCLSID);
vbt_printf("%s:%d\n",__FILE__,__LINE__);
	if(FAILED(hr))
		return hr;
#endif

	hr = setKeyAndValue(wKey, L"Implemented Categories", NULL);
vbt_printf("%s:%d\n",__FILE__,__LINE__);
	if(FAILED(hr))
		return hr;

	hr = setKeyAndValue(wKey, L"Implemented Categories\\{40FC6ED5-2438-11CF-A3DB-080036F12502}", NULL);
vbt_printf("%s:%d\n",__FILE__,__LINE__);
	if(FAILED(hr))
		return hr;

	/* Add the server filename subkey under the CLSID key. */
vbt_printf("%s:%d\n",__FILE__,__LINE__);
	if(InProcServer)
	{
		OLECHAR wInproc[64];
		hr = setKeyAndValue(wKey, L"InprocServer32", wModule);
		if(FAILED(hr))
			return hr;
		wcscpy(wInproc, wKey);
		wcscat(wInproc, L"\\InprocServer32");
		hr = setValueInKey(wInproc, L"ThreadingModel", wThreadingModel);
		if(FAILED(hr))
			return hr;
	}
	else
	{
		hr = setKeyAndValue(wKey, L"LocalServer32", wModule);
		if(FAILED(hr))
			return hr;
	}

	/* Add the ProgID subkey under the CLSID key. */
	hr = setKeyAndValue(wKey, L"ProgID", wVerIndProgID);
	if(FAILED(hr))
		return hr;

	/* Add the Programmable subkey under the CLSID key. */
	hr = setKeyAndValue(wKey, L"Programmable", NULL);
	if(FAILED(hr))
		return hr;

	/* Add the TypeLib subkey under the CLSID key. */
	{
	/* Convert a TypeLib GUID into a wchar. */
	OLECHAR wLIBID[64];
	hr = StringFromGUID2(libid, wLIBID, sizeof(wLIBID));
	if(FAILED(hr))
		return hr;
	hr = setKeyAndValue(wKey, L"TypeLib", wLIBID);
	if(FAILED(hr))
		return hr;
	}

	/* Add the VERSION subkey under the CLSID key. */
	{
	wchar_t wVersion[16];
	swprintf(wVersion,L"%u.%u",libMajor,libMinor);
	_itow(libMajor,wVersion,10);
	wcscat(wVersion,L".");
	_itow(libMinor,wVersion+wcslen(wVersion),10);
	hr = setKeyAndValue(wKey, L"VERSION", wVersion);
	if(FAILED(hr))
		return hr;
	}

	/* Add the version-independent ProgID subkey under HKEY_CLASSES_ROOT. */
	hr = setKeyAndValue(wVerIndProgID, NULL, wFriendlyName); 
	if(FAILED(hr))
		return hr;
	hr = setKeyAndValue(wVerIndProgID, L"Clsid", wCLSID);
	if(FAILED(hr))
		return hr;

	return S_OK;
}

/* Remove the component from the registry. */
HRESULT UnregisterServer(REFCLSID clsid,             /* Class ID */
                      LPCOLESTR wVerIndProgID, /* Programmatic */
                      LPCOLESTR wProgID)       /* IDs */
{
	/* Convert the CLSID into a char. */
	OLECHAR wCLSID[64];
	OLECHAR wKey[64];
	HRESULT hr;

	/* Convert a CLSID GUID into a wchar. */
	hr = StringFromGUID2(clsid, wCLSID, sizeof(wCLSID));
	if(FAILED(hr))
		return hr;

	/* Build the key CLSID\\{...} */
	wcscpy(wKey, L"CLSID\\");
	wcscat(wKey, wCLSID);

	/* Delete the CLSID Key - CLSID\{...} */
	hr = recursiveDeleteKey(HKEY_CLASSES_ROOT, wKey);
	if(FAILED(hr))
		return ResultFromScode(E_FAIL);

	/* Delete the version-independent ProgID Key. */
	hr = recursiveDeleteKey(HKEY_CLASSES_ROOT, wVerIndProgID);
	if(FAILED(hr))
		return ResultFromScode(E_FAIL);

	/* Delete the ProgID key. */
	hr = recursiveDeleteKey(HKEY_CLASSES_ROOT, wProgID);
	if(FAILED(hr))
		return ResultFromScode(E_FAIL);

	return S_OK;
}

/* Delete a key and all of its descendents. */
HRESULT recursiveDeleteKey(HKEY hKeyParent,           /* Parent of key to delete */
                        LPCOLESTR lpwKeyChild)  /* Key to delete */
{
	/* Open the child. */
	HKEY hKeyChild;
	LONG lResult;
	FILETIME time;
	OLECHAR wBuffer[256];
	DWORD dwSize = 256;

	lResult = RegOpenKeyEx(hKeyParent, lpwKeyChild, 0, KEY_ALL_ACCESS, &hKeyChild);
	if(lResult == ERROR_FILE_NOT_FOUND)
		return S_OK;
 	else if(lResult != ERROR_SUCCESS)
		return ResultFromScode(E_FAIL);

	/* Enumerate all of the decendents of this child. */
	while(RegEnumKeyEx(hKeyChild, 0, wBuffer, &dwSize, NULL, NULL, NULL, &time) == ERROR_SUCCESS)
	{
		/* Delete the decendents of this child. */
		lResult = recursiveDeleteKey(hKeyChild, wBuffer);
		if(lResult != ERROR_SUCCESS)
		{
			/* Cleanup before exiting. */
			RegCloseKey(hKeyChild);
			return ResultFromScode(E_FAIL);
		}
		dwSize = 256;
	}

	/* Close the child. */
	lResult = RegCloseKey(hKeyChild);
	if(lResult != ERROR_SUCCESS)
		return ResultFromScode(E_FAIL);

	/* Delete this child. */
	lResult = RegDeleteKey(hKeyParent, lpwKeyChild);
	if(lResult != ERROR_SUCCESS)
		return ResultFromScode(E_FAIL);

	return S_OK;
}

/* Create a key and set its value. */
HRESULT setKeyAndValue(LPCOLESTR wKey, LPCOLESTR wSubkey, LPCOLESTR wValue)
{
	HKEY hKey;
	OLECHAR wKeyBuf[1024];
	LONG lResult;

	/* Copy keyname into buffer. */
	wcscpy(wKeyBuf, wKey);

	/* Add subkey name to buffer. */
	if(wSubkey != NULL)
	{
		wcscat(wKeyBuf, L"\\");
		wcscat(wKeyBuf, wSubkey );
	}

	/* Create and open key and subkey. */
	lResult = RegCreateKeyEx(HKEY_CLASSES_ROOT, wKeyBuf, 0, NULL, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hKey, NULL);
	if(lResult != ERROR_SUCCESS)
		return ResultFromScode(E_FAIL);

	/* Set the Value. */
	if(wValue != NULL)
	{
		lResult = RegSetValueEx(hKey, NULL, 0, REG_SZ, (BYTE *)wValue, wcslen(wValue)*sizeof(wValue[0])+1);
		if(lResult != ERROR_SUCCESS)
			return ResultFromScode(E_FAIL);
	}

	lResult = RegCloseKey(hKey);
	if(lResult != ERROR_SUCCESS)
		return ResultFromScode(E_FAIL);
	return S_OK;
}

/* Open a key and set a value. */
HRESULT setValueInKey(LPCOLESTR wKey, LPCOLESTR wNamedValue, LPCOLESTR wValue)
{
	HKEY hKey;
	OLECHAR wKeyBuf[1024];
	LONG lResult;

	/* Copy keyname into buffer. */
	wcscpy(wKeyBuf, wKey);

	/* Create and open key and subkey. */
	lResult = RegOpenKeyEx(HKEY_CLASSES_ROOT, wKeyBuf, 0, KEY_SET_VALUE, &hKey);
	if(lResult != ERROR_SUCCESS)
		return ResultFromScode(E_FAIL);

    /* Set the Value. */
	if(wValue != NULL)
	{
		lResult = RegSetValueEx(hKey, wNamedValue, 0, REG_SZ, (BYTE*)wValue, wcslen(wValue)*sizeof(wValue[0])+1);
		if(lResult != ERROR_SUCCESS)
			return ResultFromScode(E_FAIL);
	}

	lResult = RegCloseKey(hKey);
	if(lResult != ERROR_SUCCESS)
		return ResultFromScode(E_FAIL);
	return S_OK;
}

HRESULT ProcessRegistration(int argc, wchar_t *wargv[], ITypeLib **ppITypeLib)
{
HRESULT hr;
OLECHAR *p;
vbt_printf("%lx %S %S %S %lx\n",
	*(long *)g_ProjectTable.ClassTabs[0]->classGUID,
	g_ProjectTable.ClassTabs[0]->VIProgID,
	g_ProjectTable.ClassTabs[0]->VIProgID,
	g_ProjectTable.ClassTabs[0]->ProgID,
	*(long *)g_ProjectTable.TypeLibGUID);
/* Note that server registration always occurs */
	hr = RegisterServer(NULL,
			FALSE,
			g_ProjectTable.ClassTabs[0]->classGUID,
			g_ProjectTable.ClassTabs[0]->VIProgID,
			g_ProjectTable.ClassTabs[0]->VIProgID,
			g_ProjectTable.ClassTabs[0]->ProgID,
			g_ProjectTable.TypeLibGUID,
			g_ProjectTable.TypeLibMajorVersion,
			g_ProjectTable.TypeLibMinorVersion,
			NULL);
	vbt_printf("RegisterServer: hr=%lx\n",hr);
	if(FAILED(hr))
		return(hr);
/* VB uses "/" but seems to allow "-" switch on -embedding only */
#ifndef CE_TLINUX86
vbt_printf("%s:%d\n",__FILE__,__LINE__);
	if(argc > 1 && (p = wcstok(wargv[1], L"-/")) != NULL)
#else /* Linux gcc */
vbt_printf("%s:%d wargv[1]=%S\n",__FILE__,__LINE__,wargv[1]);
	/*if(argc > 1 && wcstok(wargv[1], L"-/", &p))*/
	if(argc > 1 && (p = wcschr(wargv[1], L'-')) != NULL && p++)
#endif
	{
vbt_printf("%s:%d p=%S\n",__FILE__,__LINE__,p);
		if(wcsicmp(p, L"RegServer") == 0)
		{
		OLECHAR wrel[512],wabs[512],*w;
			vbt_printf("%s:%d lib=%S wargv[0]=%S\n",__FILE__,__LINE__,g_ProjectTable.TypeLibFile,wargv[0]);
/*	mbstowcs(wrel,argv[0],sizeof(wrel)/sizeof(wrel[0]));*/
			wcsncpy(wrel,wargv[0],sizeof(wrel)/sizeof(wrel[0]));
/* fixme: fixup / vs. \\ logic */
#if SAG_COM
			w = wcsrchr(wrel,'/');
#else
			w = wcsrchr(wrel,'\\');
#endif
			if (w == NULL)
				w = wrel;
			*w = 0;
			vbt_printf("%s:%d wargv[0]=%S -- wrel=%S\n",__FILE__,__LINE__,wargv[0],wrel);
			_wfullpath(wabs,wrel,sizeof(wabs)/sizeof(wabs[0]));
			vbt_printf("%s:%d wargv[0]=%S -- wabs=%S\n",__FILE__,__LINE__,wargv[0],wabs);
#if 1
			hr = vbtCreateTypeLib(wabs);
			vbt_printf("%s:%d hr=%lx\n",__FILE__,__LINE__,hr);
			if (FAILED(hr))
				return(hr);
#endif
/* fixme: register HelpDir? - leaving NULL for now */
/*			hr = RegisterTypeLib(*ppITypeLib, wabs, NULL);*/
#if SAG_COM
			wcscat(wabs, L"/");
#else
			wcscat(wabs, L"\\");
#endif
			wcscat(wabs, g_ProjectTable.TypeLibFile);
			vbt_printf("%s:%d wargv[0]=%S -- wabs=%S\n",__FILE__,__LINE__,wargv[0],wabs);
			hr = LoadTypeLibEx(wabs, REGKIND_REGISTER, ppITypeLib);
			vbt_printf("RegisterTypeLib: hr=%lx\n",hr);
			if(FAILED(hr))
				return(hr);
			ITypeLib_Release(*ppITypeLib);
			vbt_printf("Component has been registered\n");
			return(-1);
		}
		else if(wcsicmp(p, L"UnregServer") == 0)
		{
			hr = UnRegisterTypeLib(g_ProjectTable.TypeLibGUID, g_ProjectTable.TypeLibMajorVersion, g_ProjectTable.TypeLibMinorVersion, LANG_NEUTRAL, SYS_WIN32);
			if(FAILED(hr) /*&& hr != TYPE_E_REGISTRYACCESS*/)
				vbt_printf("UnRegisterTypeLib failed: hr=%lx\n",hr);
			hr = UnregisterServer(g_ProjectTable.ClassTabs[0]->classGUID, g_ProjectTable.ClassTabs[0]->ProgID, g_ProjectTable.ClassTabs[0]->VIProgID);
			if(FAILED(hr) /*&& hr != TYPE_E_REGISTRYACCESS*/)
				vbt_printf("UnregisterServer failed: hr=%lx\n",hr);
			vbt_printf("Component has been unregistered\n");
			return(-1);
		}
		else if(wcsicmp(p, L"Embedding") == 0)
		{
			vbt_printf("Component is being embedded\n");
		}
		else
			vbt_printf("Unknown switch is ignored\n");
	}
	else
		vbt_printf("Component is being embedded by default\n");
vbt_printf("%s:%d  hr=%lx\n",__FILE__,__LINE__,hr);
#if 1
	hr = LoadRegTypeLib(g_ProjectTable.TypeLibGUID, g_ProjectTable.TypeLibMajorVersion, g_ProjectTable.TypeLibMinorVersion, LANG_NEUTRAL, ppITypeLib);
	vbt_printf("LoadRegTypeLib: hr=%lx\n",hr);
#else
#error wabs is undefined
	hr = LoadTypeLibEx(wabs, REGKIND_NONE, ppITypeLib);
	vbt_printf("LoadTypeLibEx: hr=%lx\n",hr);
#endif
	return(0); /* embed by default */
}
