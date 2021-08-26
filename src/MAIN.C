#include "vbt.h"

HRESULT RegWaitUnReg(const ClassTable * const ct, ITypeLib *pITypeLib, HANDLE hEvent)
{
	HRESULT hr;
	Generic_IClassFactory *gThis;
	DWORD dw;

	if (ct < *g_ProjectTable.ClassTabs+*g_ProjectTable.ClassCount)
	{
		LPOLESTR iid;
		gThis = Generic_IClassFactory_Constructor(NULL, ct, pITypeLib, hEvent);
		if (gThis == NULL)
			return ResultFromScode(E_OUTOFMEMORY);
		StringFromIID(ct->classGUID, &iid);
vbt_printf("class=%S\n",iid);
		hr = CoRegisterClassObject(ct->classGUID, (IUnknown *)&gThis->iface, CLSCTX_SERVER, REGCLS_SUSPENDED|REGCLS_MULTIPLEUSE, &dw);
		if (FAILED(hr))
		{
			vbt_printf("CoRegisterClassObject failed hr=%lx\n", hr);
			return hr;
		}
vbt_printf("%s:%d\n",__FILE__,__LINE__);
		RegWaitUnReg(ct+1,pITypeLib, hEvent);
vbt_printf("%s:%d\n",__FILE__,__LINE__);
		hr = CoRevokeClassObject(dw);
vbt_printf("%s:%d\n",__FILE__,__LINE__);
		if (FAILED(hr))
			vbt_printf("CoRevokeClassObject failed hr=%d\n", hr);
		return hr;
	}
	else
	{
		hr = CoResumeClassObjects();
		if (FAILED(hr))
		{
			vbt_printf("CoResumeClassObject failed hr=%lx\n", hr);
			return hr;
		}

		vbt_printf("Waiting ...\n");
		dw = WaitForSingleObject(hEvent, INFINITE);
		vbt_printf("WaitForSingleObject: dw=%ld\n",dw);

		return S_OK;
	}
}

int g_argc;
wchar_t **g_wargv;
wchar_t **g_wenvp;
FILE *vbt_debugpf;

int main(int argc, char *argv[], char *envp[]) /* fix me - MSVC can't find main() - why? */
{
int i;
size_t l;
wchar_t **wargv;
wchar_t **wenvp;

	wargv = malloc(argc*sizeof(wchar_t *));
	for(i=0;i<argc;i++)
	{
		l = strlen(argv[i])+1;
		wargv[i] = malloc(l*sizeof(wchar_t));
		mbstowcs(wargv[i],argv[i],l);
	}
	wenvp = NULL; /* fix me */
/* wenvp not implemented!!! */
	return wmain(argc, wargv, wenvp);
}

int wmain(int argc, wchar_t *wargv[], wchar_t *wenvp[])
{
	HRESULT hr;
	ITypeLib *pITypeLib;
	HANDLE hEvent;
void STDMETHODCALLTYPE Main(void);

	vbt_debugpf = stdout;
	setbuf(stdout,NULL);
	g_argc = argc;
	g_wargv = wargv;
	g_wenvp = wenvp;

	/* note: CLSCTX_INPROC_HANDLER (Excel.Application) requires COINIT_MULTITHREADED to work */
	hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
	if (FAILED(hr))
		vbt_printf("CoInitializeEx failed hr=%d\n", hr);
#if 0 /* turned off because Excel and IConnectionPointAdvise didn't work */
	hr = CoInitializeSecurity(0, -1, 0, 0, RPC_C_AUTHN_LEVEL_NONE, RPC_C_IMP_LEVEL_ANONYMOUS, 0, 0, 0);
	if (FAILED(hr))
		vbt_printf("CoInitializeSecurity failed hr=%d\n", hr);
#endif	

	if (g_ProjectTable.StartupObject != NULL)
	{
		(*g_ProjectTable.StartupObject)();
		/*return 0;*/ goto e;
	}

	hr = ProcessRegistration(argc, wargv, &pITypeLib);
	if (FAILED(hr))
	{
		if (hr != -1)
			vbt_printf("ProcessRegistration failed hr=%lx\n", hr);
		/*Sleep(7000);*/
	}
	else
	{
		hEvent = CreateEvent(NULL, FALSE, FALSE, NULL);
		vbt_printf("CreateEvent: g_hEvent=%lx\n",hEvent);
		if (hEvent == NULL)
			vbt_printf("CreateEvent failed\n");
		
		hr = RegWaitUnReg(*g_ProjectTable.ClassTabs, pITypeLib, hEvent); /* nothing to do with return value */
		
vbt_printf("%s:%d\n",__FILE__,__LINE__);
		ITypeLib_Release(pITypeLib);
vbt_printf("%s:%d\n",__FILE__,__LINE__);
	}
e:
	CoUninitialize(); /* returns void */
vbt_printf("%s:%d\n",__FILE__,__LINE__);
	
	vbt_printf("Done\n");
	return (0);
}
