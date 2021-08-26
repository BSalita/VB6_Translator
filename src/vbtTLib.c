#include "vbt.h"
#include "wchar.h" /* needed only for ICreateTypeLib2::SaveAllChanges path problem */

/* to do: check that all interface refs are released */
/* to do: release interfaces if error occurs */

HRESULT vbtCreateTypeLib(LPOLESTR wpath)
{
	HRESULT hr;
	ICreateTypeLib2* pCreateTypeLib2;
	ICreateTypeInfo* pCreateTypeInfoInterface;
	ICreateTypeInfo* pCreateTypeInfoCoClass;
	ITypeLib* pTypeLib;
	HREFTYPE hRefTypeIface;
	ITypeInfo* pTypeInfo;
	ITypeLib* pTypeLibStdOle;
	ITypeInfo* pTypeInfoDispatch;
	UINT nclass, niface, nmethod;
	HREFTYPE hRefType;
	wchar_t wolddir[512]; /* use symbolic constant */

/* _wgetcwd, _wchdir, _wchdir combination is to have CreateTypeLib2 put the file in the same directory as the executable */
	hr = CreateTypeLib2(SYS_WIN32, g_ProjectTable.TypeLibFile, &pCreateTypeLib2);
vbt_printf("%s:%d hr=%lx\n",__FILE__,__LINE__,hr);
	if (FAILED(hr)) return(hr);
	
	/* Set the type library LIBID  */
	hr = ICreateTypeLib2_SetGuid(pCreateTypeLib2, g_ProjectTable.TypeLibGUID);
vbt_printf("%s:%d hr=%lx\n",__FILE__,__LINE__,hr);
	if (FAILED(hr)) return(hr);
	
	/* Set the library version to 1.0 */
	hr = ICreateTypeLib2_SetVersion(pCreateTypeLib2, g_ProjectTable.TypeLibMajorVersion, g_ProjectTable.TypeLibMinorVersion);
vbt_printf("%s:%d hr=%lx\n",__FILE__,__LINE__,hr);
	if (FAILED(hr)) return(hr);
	
	/* Set the library helpstring */
	hr = ICreateTypeLib2_SetDocString(pCreateTypeLib2, L"Created by Softworks vbt");
vbt_printf("%s:%d hr=%lx\n",__FILE__,__LINE__,hr);
	if (FAILED(hr)) return(hr);
	
	/* Set the LCID */
	hr = ICreateTypeLib2_SetLcid(pCreateTypeLib2, LANG_NEUTRAL);
vbt_printf("%s:%d hr=%lx\n",__FILE__,__LINE__,hr);
	if (FAILED(hr)) return(hr);
	
	/* Set the library name */
	hr = ICreateTypeLib2_SetName(pCreateTypeLib2, (LPOLESTR)g_ProjectTable.TypeLibName); /* SetName should use LPCOLESTR */
vbt_printf("%s:%d hr=%lx\n",__FILE__,__LINE__,hr);
	if (FAILED(hr)) return(hr);
	
	/* Get a pointer to the ITypeLib interface for StdOLE */
	{
		GUID GUID_STDOLE2 = {0x00020430,0x00,0x00,{0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46}};
		hr = LoadRegTypeLib(&GUID_STDOLE2, STDOLE2_MAJORVERNUM, STDOLE2_MINORVERNUM, STDOLE2_LCID, &pTypeLibStdOle);
vbt_printf("%s:%d hr=%lx\n",__FILE__,__LINE__,hr);
		if (FAILED(hr)) return(hr);
	}
	
	for(nclass=0;nclass<1;nclass++) /* number of classes hardcoded to 1 */
	{
		for(niface=0;niface<*g_ProjectTable.ClassTabs[nclass]->InterfaceCount;niface++)
		{
			if (niface == 0)
			{
				/* Create the interface */
				hr = ICreateTypeLib2_CreateTypeInfo(pCreateTypeLib2, (LPOLESTR)g_ProjectTable.ClassTabs[nclass]->InterfaceDescs[niface]->InterfaceName, TKIND_INTERFACE, &pCreateTypeInfoInterface);
vbt_printf("%s:%d hr=%lx\n",__FILE__,__LINE__,hr);
				if (FAILED(hr)) return(hr);
				/* Set the IID */
				hr = ICreateTypeInfo_SetGuid(pCreateTypeInfoInterface, g_ProjectTable.ClassTabs[nclass]->InterfaceDescs[niface]->InterfaceGUID);
vbt_printf("%s:%d hr=%lx\n",__FILE__,__LINE__,hr);
				if (FAILED(hr)) return(hr);
				/* Set the Version */
				hr = ICreateTypeInfo_SetVersion(pCreateTypeInfoInterface, 1, 0);
vbt_printf("%s:%d hr=%lx\n",__FILE__,__LINE__,hr);
				if (FAILED(hr)) return(hr);
				/* Set the type flags */
				hr = ICreateTypeInfo_SetTypeFlags(pCreateTypeInfoInterface, TYPEFLAG_FHIDDEN | TYPEFLAG_FDUAL | TYPEFLAG_FNONEXTENSIBLE | TYPEFLAG_FOLEAUTOMATION);
vbt_printf("%s:%d hr=%lx\n",__FILE__,__LINE__,hr);
				if (FAILED(hr)) return(hr);
				/* Declare that the interface is derived from IDispatch */
				hr = ITypeLib_GetTypeInfoOfGuid(pTypeLibStdOle, &IID_IDispatch, &pTypeInfoDispatch);
vbt_printf("%s:%d hr=%lx\n",__FILE__,__LINE__,hr);
				if (FAILED(hr)) return(hr);
				hr = ICreateTypeInfo_AddRefTypeInfo(pCreateTypeInfoInterface, pTypeInfoDispatch, &hRefType);
vbt_printf("%s:%d hr=%lx\n",__FILE__,__LINE__,hr);
				if (FAILED(hr)) return(hr);
				hr = ICreateTypeInfo_AddImplType(pCreateTypeInfoInterface, 0, hRefType);
vbt_printf("%s:%d hr=%lx\n",__FILE__,__LINE__,hr);
				if (FAILED(hr)) return(hr);
				for(nmethod=0;nmethod<g_ProjectTable.ClassTabs[nclass]->InterfaceDescs[niface]->FuncCount;nmethod++)
				{
					/* Set function and parameter info */
					hr = ICreateTypeInfo_AddFuncDesc(pCreateTypeInfoInterface, nmethod, (LPFUNCDESC)g_ProjectTable.ClassTabs[nclass]->InterfaceDescs[niface]->Funcs[nmethod].FuncDescs);
vbt_printf("%s:%d hr=%lx\n",__FILE__,__LINE__,hr);
					if (FAILED(hr)) return(hr);
					/* Set names for the method and its parameters */
					hr = ICreateTypeInfo_SetFuncAndParamNames(pCreateTypeInfoInterface, nmethod, (LPOLESTR * const)g_ProjectTable.ClassTabs[nclass]->InterfaceDescs[niface]->Funcs[nmethod].FuncNames, g_ProjectTable.ClassTabs[nclass]->InterfaceDescs[niface]->Funcs[nmethod].FuncDescs->cParams+1);
vbt_printf("%s:%d hr=%lx\n",__FILE__,__LINE__,hr);
					if (FAILED(hr)) return(hr);
				} /* method loop */
				/* Create the coclass */
				hr = ICreateTypeLib2_CreateTypeInfo(pCreateTypeLib2, (LPOLESTR)g_ProjectTable.ClassTabs[nclass]->className, TKIND_COCLASS, &pCreateTypeInfoCoClass);
vbt_printf("%s:%d hr=%lx\n",__FILE__,__LINE__,hr);
				if (FAILED(hr)) return(hr);
				
				/* Set the CLSID */
				hr = ICreateTypeInfo_SetGuid(pCreateTypeInfoCoClass, g_ProjectTable.ClassTabs[nclass]->classGUID);
vbt_printf("%s:%d hr=%lx\n",__FILE__,__LINE__,hr);
				if (FAILED(hr)) return(hr);
				
				hr = ICreateTypeInfo_SetVersion(pCreateTypeInfoCoClass, 1, 0);
vbt_printf("%s:%d hr=%lx\n",__FILE__,__LINE__,hr);
				if (FAILED(hr)) return(hr);
				
				/* Specify that this coclass can be instantiated */
				hr = ICreateTypeInfo_SetTypeFlags(pCreateTypeInfoCoClass, TYPEFLAG_FCANCREATE);
vbt_printf("%s:%d hr=%lx\n",__FILE__,__LINE__,hr);
				if (FAILED(hr)) return(hr);
				
				/* Get a pointer to the ITypeLib interface */
				hr = ICreateTypeLib2_QueryInterface(pCreateTypeLib2, &IID_ITypeLib, (void**)&pTypeLib);
vbt_printf("%s:%d hr=%lx\n",__FILE__,__LINE__,hr);
				if (FAILED(hr)) return(hr);
				
				/* Get a pointer to the ITypeInfo interface */
				hr = ITypeLib_GetTypeInfoOfGuid(pTypeLib, g_ProjectTable.ClassTabs[nclass]->InterfaceDescs[niface]->InterfaceGUID, &pTypeInfo);
vbt_printf("%s:%d hr=%lx\n",__FILE__,__LINE__,hr);
				if (FAILED(hr)) return(hr);
				
				/* Trade in the ITypeInfo pointer for an HREFTYPE */
				hr = ICreateTypeInfo_AddRefTypeInfo(pCreateTypeInfoCoClass, pTypeInfo, &hRefTypeIface);
vbt_printf("%s:%d hr=%lx\n",__FILE__,__LINE__,hr);
				if (FAILED(hr)) return(hr);
				
				/* niface may not be correct choice */
				/* Insert the interface into the coclass */
				hr = ICreateTypeInfo_AddImplType(pCreateTypeInfoCoClass, niface, hRefTypeIface);
vbt_printf("%s:%d hr=%lx\n",__FILE__,__LINE__,hr);
				if (FAILED(hr)) return(hr);
				
				/* Set interface to be the default interface in coclass */
				hr = ICreateTypeInfo_SetImplTypeFlags(pCreateTypeInfoCoClass, niface, IMPLTYPEFLAG_FDEFAULT);
vbt_printf("%s:%d hr=%lx\n",__FILE__,__LINE__,hr);
				if (FAILED(hr)) return(hr);
			}
			else if (niface == 1)
			{
				ITypeInfo *pTypeInfoIUnknown;
				hr = ICreateTypeLib2_CreateTypeInfo(pCreateTypeLib2, (LPOLESTR)g_ProjectTable.ClassTabs[nclass]->InterfaceDescs[niface]->InterfaceName, TKIND_DISPATCH, &pCreateTypeInfoInterface);
vbt_printf("%s:%d hr=%lx\n",__FILE__,__LINE__,hr);
				if (FAILED(hr)) return(hr);
				hr = ICreateTypeInfo_SetGuid(pCreateTypeInfoInterface, g_ProjectTable.ClassTabs[nclass]->InterfaceDescs[niface]->InterfaceGUID);
vbt_printf("%s:%d hr=%lx\n",__FILE__,__LINE__,hr);
				if (FAILED(hr)) return(hr);
				hr = ICreateTypeInfo_SetVersion(pCreateTypeInfoInterface, 1, 0);
vbt_printf("%s:%d hr=%lx\n",__FILE__,__LINE__,hr);
				if (FAILED(hr)) return(hr);
				hr = ICreateTypeInfo_SetTypeFlags(pCreateTypeInfoInterface, TYPEFLAG_FHIDDEN | TYPEFLAG_FNONEXTENSIBLE);
vbt_printf("%s:%d hr=%lx\n",__FILE__,__LINE__,hr);
				if (FAILED(hr)) return(hr);
				/* Declare that the interface is derived from IDispatch */
				hr = ITypeLib_GetTypeInfoOfGuid(pTypeLibStdOle, &IID_IDispatch, &pTypeInfoDispatch);
vbt_printf("%s:%d hr=%lx\n",__FILE__,__LINE__,hr);
				if (FAILED(hr)) return(hr);
				hr = ICreateTypeInfo_AddRefTypeInfo(pCreateTypeInfoInterface, pTypeInfoDispatch, &hRefType);
vbt_printf("%s:%d hr=%lx\n",__FILE__,__LINE__,hr);
				if (FAILED(hr)) return(hr);
				hr = ICreateTypeInfo_AddImplType(pCreateTypeInfoInterface, 0, hRefType);
vbt_printf("%s:%d hr=%lx\n",__FILE__,__LINE__,hr);
				if (FAILED(hr)) return(hr);
				/* Get a pointer to the ITypeInfo interface for IUnknown */
				hr = ITypeLib_GetTypeInfoOfGuid(pTypeLibStdOle, &IID_IUnknown, &pTypeInfoIUnknown);
vbt_printf("%s:%d hr=%lx\n",__FILE__,__LINE__,hr);
				if (FAILED(hr)) return(hr);
				hr = ICreateTypeInfo_AddRefTypeInfo(pCreateTypeInfoInterface, pTypeInfoIUnknown, &hRefType);
vbt_printf("%s:%d hr=%lx\n",__FILE__,__LINE__,hr);
				if (FAILED(hr)) return(hr);
				/*		hr = ICreateTypeInfo_AddImplType(pCreateTypeInfoInterface, 1, hRefType);*/
				for(nmethod=0;nmethod<g_ProjectTable.ClassTabs[nclass]->InterfaceDescs[niface]->FuncCount;nmethod++)
				{
					hr = ICreateTypeInfo_AddFuncDesc(pCreateTypeInfoInterface, nmethod, (LPFUNCDESC)g_ProjectTable.ClassTabs[nclass]->InterfaceDescs[niface]->Funcs[nmethod].FuncDescs);
vbt_printf("%s:%d hr=%lx\n",__FILE__,__LINE__,hr);
					if (FAILED(hr)) return(hr);
					/* Set names for the method and its parameters */
					hr = ICreateTypeInfo_SetFuncAndParamNames(pCreateTypeInfoInterface, nmethod, (LPOLESTR * const)g_ProjectTable.ClassTabs[nclass]->InterfaceDescs[niface]->Funcs[nmethod].FuncNames, g_ProjectTable.ClassTabs[nclass]->InterfaceDescs[niface]->Funcs[nmethod].FuncDescs->cParams+1);
vbt_printf("%s:%d hr=%lx\n",__FILE__,__LINE__,hr);
					if (FAILED(hr)) return(hr);
				} /* method loop */
				/* Get a pointer to the ITypeInfo interface */
				hr = ITypeLib_GetTypeInfoOfGuid(pTypeLib, g_ProjectTable.ClassTabs[nclass]->InterfaceDescs[niface]->InterfaceGUID, &pTypeInfo);
vbt_printf("%s:%d hr=%lx\n",__FILE__,__LINE__,hr);
				if (FAILED(hr)) return(hr);
				
				/* Trade in the ITypeInfo pointer for an HREFTYPE */
				hr = ICreateTypeInfo_AddRefTypeInfo(pCreateTypeInfoCoClass, pTypeInfo, &hRefTypeIface);
vbt_printf("%s:%d hr=%lx\n",__FILE__,__LINE__,hr);
				if (FAILED(hr)) return(hr);
				
				/* niface may not be choice */
				/* (Insert the interface into the coclass */
				hr = ICreateTypeInfo_AddImplType(pCreateTypeInfoCoClass, niface, hRefTypeIface);
vbt_printf("%s:%d hr=%lx\n",__FILE__,__LINE__,hr);
				if (FAILED(hr)) return(hr);
				
				/* Set interface to be the default interface in coclass */
				hr = ICreateTypeInfo_SetImplTypeFlags(pCreateTypeInfoCoClass, niface, IMPLTYPEFLAG_FDEFAULT | IMPLTYPEFLAG_FSOURCE);
vbt_printf("%s:%d hr=%lx\n",__FILE__,__LINE__,hr);
				if (FAILED(hr)) return(hr);
			}
			else
				return ResultFromScode(E_FAIL);
		} /* interface loop */
	} /* class loop */
	
	/* Assign the v-table layout */
	hr = ICreateTypeInfo_LayOut(pCreateTypeInfoInterface);
vbt_printf("%s:%d hr=%lx\n",__FILE__,__LINE__,hr);
	if (FAILED(hr)) return(hr);
	
	/* Save changes to disk */
	if (*wpath)
	{
#if 1
		if (_wgetcwd(wolddir,sizeof(wolddir)) == NULL)
#else /* is this for some UNIX system? */
		if (_wgetcwd(wolddir,sizeof(wolddir)))
#endif
			return(ResultFromScode(E_FAIL));
		if (_wchdir(wpath))
			return(ResultFromScode(E_FAIL));
	}
	hr = ICreateTypeLib2_SaveAllChanges(pCreateTypeLib2);
vbt_printf("%s:%d hr=%lx\n",__FILE__,__LINE__,hr);
	if (*wpath && _wchdir(wolddir))
		return(ResultFromScode(E_FAIL));
	if (FAILED(hr)) return(hr);
	
	/* Release all references */
	/* rewrite so that all interfaces are properly released - both normally and upon error */
	ITypeInfo_Release(pTypeInfoDispatch);
	ITypeLib_Release(pTypeLibStdOle);
	ITypeInfo_Release(pTypeInfo);
	ITypeLib_Release(pTypeLib);
	ICreateTypeLib2_Release(pCreateTypeLib2);
	ICreateTypeInfo_Release(pCreateTypeInfoInterface);
	ICreateTypeInfo_Release(pCreateTypeInfoCoClass);
	
	return(S_OK);
}
