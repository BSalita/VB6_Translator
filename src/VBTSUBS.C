
/* make connection point a linked list */
/* have constructors return HRESULT for better error messages */
/* dump linked lists in Server routines */

#include "vbt.h"
#include <stdio.h> /* needed for printf for Linux */
#include <olectl.h> /* needed for CONNECT_E_ error messages */
#undef free
#define free(x)

int vbt_diag(int errnum,char *file,int line)
{
	vbt_printf("abort: %s:%d errnum=%d\n",__FILE__,__LINE__,errnum);
	abort();
}

#include <time.h> /* time_t */
#include <limits.h> /* INT_MAX */
/* converts local time VB julian date to local time tm */
struct tm *vbdatetotm(vbdate)
double vbdate;
{
time_t lt;
struct tm *tm;

/* init tm to 1 Jan 1970 */
	lt = 0;
	tm = gmtime(&lt);
/* adjust tm_mday to julian date */
	lt = (time_t)vbdate;
	if (lt)
		if (tm->tm_mday+(lt-SW_VB_1970) > INT_MAX)
			return(NULL);
		else
			tm->tm_mday += (int)(lt-SW_VB_1970);
	tm->tm_sec += (int)((vbdate - (double)lt) * (double)SW_SECS_PER_DAY + 0.5); /* includes rounding factor */
/* redo tm */
	if (mktime(tm) == -1)
		return(NULL);
	return(tm);
}

/* converts local time tm to local time VB format julian date */
double tmtovbdate(tm)
struct tm *tm;
{
	return((double)(mktime(tm)-timezone+tm->tm_isdst*3600)/(double)SW_SECS_PER_DAY+(double)SW_VB_1970);
}

wchar_t *wmemchr(s,c,n)
const wchar_t *s;
wchar_t c;
size_t n;
{
	while(n--)
		if (c == *s++)
			return((wchar_t *)--s);
	return(NULL);
}

INT wmemcmp(s1,s2,n)
const wchar_t *s1;
const wchar_t *s2;
size_t n;
{
	while(n--)
		if (*s1++ == *s2++)
			return(*--s1-*--s2);
	return(0);
}

wchar_t *wmemcpy(s1,s2,n)
wchar_t *s1;
const wchar_t *s2;
size_t n;
{
wchar_t *ss = s1;
	while(n--)
		*ss++ = *s2++;
	return(s1);
}

wchar_t *wmemmove(s1,s2,n)
wchar_t *s1;
const wchar_t *s2;
size_t n;
{
wchar_t *ss = s1;
	if (s1 < s2)
		while(n--)
			*s1++ = *s2++;
	else
		for(s1 += n, s2 += n;n--;)
			*--s1 = *--s2;
	return(ss);
}

wchar_t *wmemset(s,c,n)
wchar_t *s;
wchar_t c;
size_t n;
{
wchar_t *ss = s;
	while(n--)
		*ss++ = c;
	return(s);
}

static Generic_IUnknown *Generic_IUnknown_Constructor(IUnknown *This);
static void Generic_IUnknown_Destructor(IUnknown *This);

static Generic_IDispatch *Generic_IDispatch_Constructor(IUnknown *This);
static void Generic_IDispatch_Destructor(IDispatch *This);

static Generic_IConnectionPoint *Generic_IConnectionPoint_Constructor(IUnknown *This, const InterfaceDescTable * const idt);
static void Generic_IConnectionPoint_Destructor(IConnectionPoint *This);

static Generic_IConnectionPointContainer *Generic_IConnectionPointContainer_Constructor(IUnknown *This);
static void Generic_IConnectionPointContainer_Destructor(IConnectionPointContainer *This);

static Generic_IEnumConnections *Generic_IEnumConnections_Constructor(IUnknown *This, int cConn, CONNECTDATA* pConnData);
static void Generic_IEnumConnections_Destructor(IEnumConnections *This);

static Generic_IEnumConnectionPoints *Generic_IEnumConnectionPoints_Constructor(IUnknown* This, IConnectionPoint** rgpCP);
static void Generic_IEnumConnectionPoints_Destructor(IEnumConnectionPoints *This);

static void * Generic_Constructor(IUnknown *This, const InterfaceDescTable * const idt);
static void Generic_Destructor(IUnknown * This);

static HRESULT STDMETHODCALLTYPE Generic_IClassFactory_QueryInterface(IClassFactory * This, REFIID riid, void ** ppvObject);
static ULONG STDMETHODCALLTYPE Generic_IClassFactory_AddRef(IClassFactory * This);
static ULONG STDMETHODCALLTYPE Generic_IClassFactory_Release(IClassFactory * This);

static ULONG ServerAddRef(Generic_IUnknown *gThis);
static ULONG ServerRelease(Generic_IUnknown *gThis);

static ULONG ServerAddRef(Generic_IUnknown *gThis)
{
	ULONG ref;
	ref = CoAddRefServerProcess();
	vbt_printf("ServerAddRef: ref=%ld\n",ref);
	return ref;
}

static ULONG ServerRelease(Generic_IUnknown *gThis)
{
	ULONG ref;
	ref = CoReleaseServerProcess();
	vbt_printf("ServerRelease: ref=%ld\n",ref);
	if (ref == 0)
	{
		vbt_printf("ServerRelease: calling SetEvent\n",ref);
		SetEvent(gThis->gid.hEvent);
	}
	return ref;
}

/* add init parameter? */
static void * Generic_Constructor(IUnknown *This, const InterfaceDescTable * const idt) 
{
	Generic_IUnknown *gThis = (Generic_IUnknown *)This;
	Generic_IUnknown *newIUnk;
	newIUnk = calloc(1,idt->InterfaceDataSize ? idt->InterfaceDataSize : sizeof(Generic_IUnknown)); /* should zero be supported? */
	if (newIUnk == NULL)
		return NULL;
	newIUnk->iface = *(IUnknown *)&idt->iface;
	newIUnk->gid.idt = idt;
	if (gThis == NULL) /* Using NULL to create first node of circular queue (IClassFactory) */
	{
		newIUnk->gid.prev = newIUnk->gid.next = newIUnk;
	}
	else
	{
		newIUnk->gid.ct = gThis->gid.ct;
		newIUnk->gid.pITypeLib = gThis->gid.pITypeLib;
		newIUnk->gid.hEvent = gThis->gid.hEvent;
		/* insert node in circular queue */
		newIUnk->gid.prev = gThis;
		newIUnk->gid.next = gThis->gid.next;
		gThis->gid.next->gid.prev = newIUnk;
		gThis->gid.next = newIUnk;
		ServerAddRef(newIUnk);
	}
	return newIUnk;
}

static void Generic_Destructor( IUnknown * This )
{
	Generic_IUnknown *gThis = (Generic_IUnknown *)This;
	vbt_printf("Generic_Destructor called\n");
	/* remove node from circular queue */
	gThis->gid.prev->gid.next = gThis->gid.next;
	gThis->gid.next->gid.prev = gThis->gid.prev;
	ServerRelease(gThis);
	free(gThis);
}

HRESULT STDMETHODCALLTYPE Generic_IUnknown_QueryInterface(IUnknown * This, REFIID riid, void ** ppvObject)
{
	Generic_IUnknown *gThis = (Generic_IUnknown *)This;
	Generic_IUnknown *g = gThis;
	{
		LPOLESTR iid;
		StringFromIID(gThis->gid.idt->InterfaceGUID, &iid);
		vbt_printf("Generic_IUnknown_QueryInterface: owner=%s",wtos(iid));
		StringFromIID(riid, &iid);
		vbt_printf(", query=%s\n",wtos(iid));
	}
	*ppvObject = 0;
	do
	{
		if (IsEqualGUID(riid, g->gid.idt->InterfaceGUID))
		{
			*ppvObject = &g->iface;
			Generic_IUnknown_AddRef(&g->iface);
			vbt_printf("ppvObject=%lx\n",*ppvObject);
			return NOERROR;
		}
	} while ((g = g->gid.next) && g != gThis);
	vbt_printf("E_NOINTERFACE!\n");
	return ResultFromScode(E_NOINTERFACE);
}

ULONG STDMETHODCALLTYPE Generic_IUnknown_AddRef(IUnknown * This)
{
	Generic_IUnknown *gThis = (Generic_IUnknown *)This;
	{
		LPOLESTR iid;
		StringFromIID(gThis->gid.idt->InterfaceGUID, &iid);
		vbt_printf("Generic_IUnknown_AddRef: ref=%ld iid=%s\n",gThis->gid.m_cRef,wtos(iid));
	}
	return ++gThis->gid.m_cRef;
}

ULONG STDMETHODCALLTYPE Generic_IUnknown_Release(IUnknown * This)
{
	Generic_IUnknown *gThis = (Generic_IUnknown *)This;
	--gThis->gid.m_cRef;
	{
		LPOLESTR iid;
		StringFromIID(gThis->gid.idt->InterfaceGUID, &iid);
		vbt_printf("Generic_IUnknown_Release: ref=%ld iid=%s\n",gThis->gid.m_cRef,wtos(iid));
	}
	if (gThis->gid.m_cRef == 0)
	{
		Generic_IUnknown *gNext = gThis;
		do
		{
			gThis = gNext;
			gNext = gThis->gid.next;
			if (gThis->gid.idt->destructor == NULL)
				Generic_Destructor(&gThis->iface);
			else
				(*gThis->gid.idt->destructor)(&gThis->iface);
		} while (gNext != gThis);
		return 0;
	}
vbt_printf("%s:%d\n",__FILE__,__LINE__);
	return gThis->gid.m_cRef;
}

static const IUnknownVtbl IUnknownVtblInstance =
BEGIN_VTABLE
	VTABLE_ENTRY( Generic_IUnknown_QueryInterface ),
	VTABLE_ENTRY( Generic_IUnknown_AddRef ),
	VTABLE_ENTRY( Generic_IUnknown_Release ),
END_VTABLE;

HRESULT STDMETHODCALLTYPE Generic_IDispatch_GetTypeInfoCount(IDispatch * This, UINT *pctinfo)
{
	Generic_IDispatch *gThis = (Generic_IDispatch *)This;
	vbt_printf("Generic_IDispatch_GetTypeInfoCount called\n");
	return (*pctinfo = 1);
}

HRESULT STDMETHODCALLTYPE Generic_IDispatch_GetTypeInfo(IDispatch * This, UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
	Generic_IDispatch *gThis = (Generic_IDispatch *)This;
	vbt_printf("Generic_IDispatch_GetTypeInfo: iTInfo=%d lcid=%lx\n", iTInfo, lcid);
	*ppTInfo = NULL;
	if (iTInfo) return ResultFromScode(DISP_E_BADINDEX);
	ITypeInfo_AddRef(gThis->m_pITypeInfo);
	*ppTInfo = gThis->m_pITypeInfo;
	return NOERROR;
}

HRESULT STDMETHODCALLTYPE Generic_IDispatch_GetIDsOfNames(IDispatch * This, REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispID)
{
	Generic_IDispatch *gThis = (Generic_IDispatch *)This;
	{
		UINT u;
		vbt_printf("Generic_IDispatch_GetIDsOfNames: cNames=%d\n", cNames);
		for(u=0;u<cNames;u++)
		{
			vbt_printf("\tName %d: %s\n", u, wtos(rgszNames[u]) );
		}
	}
	return DispGetIDsOfNames(gThis->m_pITypeInfo, rgszNames, cNames, rgDispID);
}

HRESULT STDMETHODCALLTYPE Generic_IDispatch_Invoke(IDispatch * This, DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	Generic_IDispatch *gThis = (Generic_IDispatch *)This;
	HRESULT hr;
	vbt_printf("Generic_IDispatch_Invoke: dispID=%lx wFlags=%x\n", dispIdMember, wFlags);
	hr = DispInvoke(This, gThis->m_pITypeInfo, dispIdMember, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
	vbt_printf("Invoke: hr=%lx\n", hr);
	return hr;
}

static const IDispatchVtbl IDispatchVtblInstance =
BEGIN_VTABLE
	VTABLE_ENTRY( GENERIC_QUERYINTERFACE(IDispatch) ),
	VTABLE_ENTRY( GENERIC_ADDREF(IDispatch) ),
	VTABLE_ENTRY( GENERIC_RELEASE(IDispatch) ),
	VTABLE_ENTRY( Generic_IDispatch_GetTypeInfoCount ),
	VTABLE_ENTRY( Generic_IDispatch_GetTypeInfo ),
	VTABLE_ENTRY( Generic_IDispatch_GetIDsOfNames ),
	VTABLE_ENTRY( Generic_IDispatch_Invoke ),
END_VTABLE;

static HRESULT STDMETHODCALLTYPE Generic_IEnumConnections_Next(IEnumConnections * This, ULONG cConnections, LPCONNECTDATA rgcd, ULONG * pcFetched)
{
	Generic_IEnumConnections *IEnumC = (Generic_IEnumConnections *)This;
	UINT cReturn = 0;
	
	vbt_printf("Generic_IEnumConnections_Next called\n");
	if(pcFetched == NULL && cConnections != 1)
		return E_INVALIDARG;
	if(pcFetched != NULL)
		*pcFetched = 0;
	if(rgcd == NULL || IEnumC->m_iCur >= IEnumC->m_cConn)
		return S_FALSE;
	while(IEnumC->m_iCur < IEnumC->m_cConn && cConnections > 0)
	{
		*rgcd++ = IEnumC->m_rgConnData[IEnumC->m_iCur];
		IUnknown_AddRef(IEnumC->m_rgConnData[IEnumC->m_iCur++].pUnk);
		cReturn++;
		cConnections--;
	} 
	return S_OK;
}

static HRESULT STDMETHODCALLTYPE Generic_IEnumConnections_Skip(IEnumConnections * This, ULONG cConnections)
{
	Generic_IEnumConnections *IEnumC = (Generic_IEnumConnections *)This;
	
	vbt_printf("Generic_IEnumConnections_Skip called\n");
	if(IEnumC->m_iCur + cConnections >= IEnumC->m_cConn)
		return S_FALSE;
	IEnumC->m_iCur += cConnections;
	return S_OK;
}

static HRESULT STDMETHODCALLTYPE Generic_IEnumConnections_Reset(IEnumConnections * This)
{
	Generic_IEnumConnections *IEnumC = (Generic_IEnumConnections *)This;
	
	vbt_printf("Generic_IEnumConnections_Reset called\n");
	IEnumC->m_iCur = 0;
	return S_OK;
}

static HRESULT STDMETHODCALLTYPE Generic_IEnumConnections_Clone(IEnumConnections * This, IEnumConnections * * ppEnum)
{
	Generic_IEnumConnections *IEnumC = (Generic_IEnumConnections *)This;
	Generic_IEnumConnections *newIEnumC;
	
	vbt_printf("Generic_IEnumConnections_Clone called\n");
	if(ppEnum == NULL)
		return E_POINTER;
	*ppEnum = NULL;
	
	/* Create the clone */
	newIEnumC = Generic_IEnumConnections_Constructor((IUnknown *)This, IEnumC->m_cConn, IEnumC->m_rgConnData);
	if(newIEnumC == NULL)
		return E_OUTOFMEMORY;
	newIEnumC->m_iCur = IEnumC->m_iCur;
	newIEnumC->m_cConn = IEnumC->m_cConn;
	newIEnumC->m_rgConnData = IEnumC->m_rgConnData;
	IEnumConnections_AddRef(&newIEnumC->iface);
	*ppEnum = &newIEnumC->iface;
	
	return S_OK;
}

static const IEnumConnectionsVtbl Generic_IEnumConnectionsVtbl =
BEGIN_VTABLE
	VTABLE_ENTRY( GENERIC_QUERYINTERFACE(IEnumConnections) ),
	VTABLE_ENTRY( GENERIC_ADDREF(IEnumConnections) ),
	VTABLE_ENTRY( GENERIC_RELEASE(IEnumConnections) ),
	VTABLE_ENTRY( Generic_IEnumConnections_Next ),
	VTABLE_ENTRY( Generic_IEnumConnections_Skip ),
	VTABLE_ENTRY( Generic_IEnumConnections_Reset ),
	VTABLE_ENTRY( Generic_IEnumConnections_Clone )
END_VTABLE;

static const InterfaceDescTable Generic_IEnumConnections_InterfaceDesc =
{
	{ (IUnknownVtbl *)&Generic_IEnumConnectionsVtbl }, sizeof(Generic_IEnumConnections), (void *)Generic_IEnumConnections_Constructor, (void *)Generic_IEnumConnections_Destructor, L"IEnumConnections", &IID_IEnumConnections, 0, NULL
};

static Generic_IEnumConnections *Generic_IEnumConnections_Constructor(IUnknown *This, int cConn, CONNECTDATA* pConnData)
{
	Generic_IEnumConnections *IEnumC;
	Generic_IUnknown *gIUnk;
	int count;
	
	vbt_printf("Generic_IEnumConnections_Constructor called\n");
	gIUnk = Generic_IUnknown_Constructor(NULL);
	if (gIUnk == NULL)
		return NULL;
	IEnumC = (Generic_IEnumConnections *)Generic_Constructor(&gIUnk->iface, &Generic_IEnumConnections_InterfaceDesc);
	if (IEnumC == NULL)
		return NULL;
	IEnumC->m_rgConnData = malloc(cConn*sizeof(CONNECTDATA *));
	if (IEnumC->m_rgConnData == NULL)
		return NULL;
	if(IEnumC->m_rgConnData != NULL)
		for(count = 0; count < cConn; count++)
		{
			IEnumC->m_rgConnData[count] = pConnData[count];
			IUnknown_AddRef(IEnumC->m_rgConnData[count].pUnk);
		}
	return IEnumC;
}

static void Generic_IEnumConnections_Destructor(IEnumConnections *This)
{
	vbt_printf("Generic_IEnumConnections_Destructor called\n");
	Generic_Destructor((IUnknown *)This);
}

static HRESULT STDMETHODCALLTYPE Generic_IEnumConnectionPoints_Next(IEnumConnectionPoints * This, ULONG cConnections, LPCONNECTIONPOINT * ppCP, ULONG * pcFetched)
{
	Generic_IEnumConnectionPoints *IEnumCP = (Generic_IEnumConnectionPoints *)This;
	
	vbt_printf("Generic_IEnumConnectionPoints_Next called\n");
	if(ppCP == NULL)
		return E_POINTER;
	if(pcFetched == NULL && cConnections != 1)
		return E_INVALIDARG;
	if(pcFetched != NULL)
		*pcFetched = 0;
	
	while(IEnumCP->m_iCur < NUM_CONNECTION_POINTS && cConnections > 0)
	{
		*ppCP = IEnumCP->m_rgpCP[IEnumCP->m_iCur++];
		if(*ppCP != NULL)
			IConnectionPoint_AddRef(*ppCP);
		if(pcFetched != NULL)
			(*pcFetched)++;
		cConnections--;
		ppCP++;
	}
	return S_OK;
}

static HRESULT STDMETHODCALLTYPE Generic_IEnumConnectionPoints_Skip(IEnumConnectionPoints * This, ULONG cConnections)
{
	Generic_IEnumConnectionPoints *IEnumCP = (Generic_IEnumConnectionPoints *)This;

	vbt_printf("Generic_IEnumConnectionPoints_Skip called\n");
	if(IEnumCP->m_iCur + cConnections >= NUM_CONNECTION_POINTS)
		return S_FALSE;
	IEnumCP->m_iCur += cConnections;
	return S_OK;
}

static HRESULT STDMETHODCALLTYPE Generic_IEnumConnectionPoints_Reset(IEnumConnectionPoints * This)
{
	Generic_IEnumConnectionPoints *IEnumCP = (Generic_IEnumConnectionPoints *)This;

	vbt_printf("Generic_IEnumConnectionPoints_Reset called\n");
	IEnumCP->m_iCur = 0;
	return S_OK;
}

static HRESULT STDMETHODCALLTYPE Generic_IEnumConnectionPoints_Clone(IEnumConnectionPoints * This, IEnumConnectionPoints * * ppEnum)
{
	Generic_IEnumConnectionPoints *IEnumCP = (Generic_IEnumConnectionPoints *)This;
	Generic_IEnumConnectionPoints *newIEnumCP;

	vbt_printf("Generic_IEnumConnectionPoints_Clone called\n");
	if(ppEnum == NULL)
		return E_POINTER;
	*ppEnum = NULL;
	
	newIEnumCP = Generic_IEnumConnectionPoints_Constructor((IUnknown *)This, IEnumCP->m_rgpCP);
	if(newIEnumCP == NULL)
		return E_OUTOFMEMORY;
	IEnumConnectionPoints_AddRef(&newIEnumCP->iface);
	newIEnumCP->m_iCur = IEnumCP->m_iCur;
	*ppEnum = &newIEnumCP->iface;
	return S_OK;
}

static const IEnumConnectionPointsVtbl Generic_IEnumConnectionPointsVtbl =
BEGIN_VTABLE
	VTABLE_ENTRY( GENERIC_QUERYINTERFACE(IEnumConnectionPoints) ),
	VTABLE_ENTRY( GENERIC_ADDREF(IEnumConnectionPoints) ),
	VTABLE_ENTRY( GENERIC_RELEASE(IEnumConnectionPoints) ),
	VTABLE_ENTRY( Generic_IEnumConnectionPoints_Next ),
	VTABLE_ENTRY( Generic_IEnumConnectionPoints_Skip ),
	VTABLE_ENTRY( Generic_IEnumConnectionPoints_Reset ),
	VTABLE_ENTRY( Generic_IEnumConnectionPoints_Clone )
END_VTABLE;

static const InterfaceDescTable Generic_IEnumConnectionPoints_InterfaceDesc =
{
	{ (IUnknownVtbl *)&Generic_IEnumConnectionPointsVtbl }, sizeof(Generic_IEnumConnectionPoints), NULL, (void *)Generic_IEnumConnectionPoints_Destructor, L"IEnumConnectionPoints", &IID_IEnumConnectionPoints, 0, NULL
};

static Generic_IEnumConnectionPoints *Generic_IEnumConnectionPoints_Constructor(IUnknown* This, IConnectionPoint** rgpCP)
{
	Generic_IEnumConnectionPoints *IEnumCP;
	Generic_IUnknown *gIUnk;
	int count;
	
	vbt_printf("Generic_IEnumConnectionPoints_Constructor called\n");
	gIUnk = Generic_IUnknown_Constructor(NULL);
	if (gIUnk == NULL)
		return NULL;
	IEnumCP = (Generic_IEnumConnectionPoints *)Generic_Constructor(&gIUnk->iface, &Generic_IEnumConnectionPoints_InterfaceDesc);
	if (IEnumCP == NULL)
		return NULL;

	// m_rgpCP is a pointer to an array of IConnectionPoints or CConnectionPoints
	for(count = 0; count < NUM_CONNECTION_POINTS; count++)
		if (FAILED(IConnectionPoint_QueryInterface(rgpCP[count], &IID_IConnectionPoint, (void **)&IEnumCP->m_rgpCP[count])))
			return NULL;
	return IEnumCP;
}

static void Generic_IEnumConnectionPoints_Destructor(IEnumConnectionPoints *This)
{
	vbt_printf("Generic_IEnumConnectionPoints_Destructor called\n");
	Generic_Destructor((IUnknown *)This);
}

static HRESULT STDMETHODCALLTYPE Generic_IConnectionPoint_GetConnectionInterface(IConnectionPoint * This, IID * pIID)
{
	Generic_IConnectionPoint *ICP = (Generic_IConnectionPoint *)This;

	vbt_printf("Generic_IConnectionPoint_GetConnectionInterface called\n");
	if(pIID == NULL)
		return E_POINTER;
	*pIID = *(IID *)ICP->m_iid; /* move in IID - warning: uncasing a const */
	return S_OK;
}

static HRESULT STDMETHODCALLTYPE Generic_IConnectionPoint_GetConnectionPointContainer(IConnectionPoint * This, IConnectionPointContainer * * ppCPC)
{
	Generic_IConnectionPoint *ICP = (Generic_IConnectionPoint *)This;

	vbt_printf("Generic_IConnectionPoint_GetConnectionPointContainer called\n");
	return IConnectionPoint_QueryInterface(ICP->m_pObj, &IID_IConnectionPointContainer, (void**)ppCPC);
}

static HRESULT STDMETHODCALLTYPE Generic_IConnectionPoint_Advise(IConnectionPoint * This, IUnknown *pUnkSink, DWORD *pdwCookie)
{
	Generic_IConnectionPoint *ICP = (Generic_IConnectionPoint *)This;
	HRESULT hr;
	IUnknown* pSink;
	INT count;

	vbt_printf("Generic_IConnectionPoint_Advise called\n");
	*pdwCookie = 0;
	if(ICP->m_cConn == CCONNMAX)
		return CONNECT_E_ADVISELIMIT;
	hr = IUnknown_QueryInterface(pUnkSink, *ICP->m_iid, (void **)&pSink);
#if 0
	{ /* temp code to test firing of events */
		IDispatch *IDisp;
		DISPPARAMS disparams;
		VARIANTARG rgvarg[3];
		short i = 123;
		double d = 321.123;
		BSTR bstr;
		hr = IUnknown_QueryInterface(pSink, &IID_IDispatch, &IDisp);
		printf("IDispatch hr=%lx\n",hr);
		if (FAILED(hr))
			return hr;
		memset(&disparams,0,sizeof(DISPPARAMS));
		memset(rgvarg,0,sizeof(rgvarg));
		disparams.cArgs=1;
		disparams.rgvarg = rgvarg;
		bstr = SysAllocString(L"Hello World!");
		V_VT(rgvarg+0) = VT_BSTR | VT_BYREF;
		V_BSTRREF(rgvarg+0) = &bstr;
		hr = IDisp->lpVtbl->Invoke(IDisp, 0x00000000, &IID_NULL, 0, DISPATCH_METHOD, &disparams, NULL, NULL, NULL);
		printf("invoke hr=%lx\n",hr);
		if (FAILED(hr))
			return hr;
		V_VT(rgvarg+0) = VT_I2 | VT_BYREF;
		V_I2REF(rgvarg+0) = &i;
		hr = IDisp->lpVtbl->Invoke(IDisp, 0x00000001, &IID_NULL, 0, DISPATCH_METHOD, &disparams, NULL, NULL, NULL);
		printf("invoke hr=%lx\n",hr);
		if (FAILED(hr))
			return hr;
	}
#endif
	if (FAILED(hr))
		return hr;
	for(count = 0; count < CCONNMAX; count++)
		if(ICP->m_rgpUnknown[count] == NULL)
		{
			ICP->m_rgpUnknown[count] = pSink;
			ICP->m_rgnCookies[count] = ++ICP->m_nCookieNext;
			*pdwCookie = ICP->m_nCookieNext;
			break;
		}
	ICP->m_cConn++;

#ifdef NEVER
	/* Hack here to copy pointer to a global variable so that we can use it from main() */
	g_pOutGoing = (IOutGoing*)pSink;
#endif

	return NOERROR;
}

static HRESULT STDMETHODCALLTYPE Generic_IConnectionPoint_Unadvise(IConnectionPoint * This, DWORD dwCookie)
{
	Generic_IConnectionPoint *ICP = (Generic_IConnectionPoint *)This;
	INT count;

	vbt_printf("Generic_IConnectionPoint_Unadvise called\n");
	if(dwCookie == 0)
		return E_INVALIDARG;
	for(count = 0; count < CCONNMAX; count++)
		if(dwCookie == ICP->m_rgnCookies[count])
		{
			if(ICP->m_rgpUnknown[count] != NULL)
			{
				IUnknown_Release(ICP->m_rgpUnknown[count]);
				ICP->m_rgpUnknown[count] = NULL;
			}
			ICP->m_cConn--;
			return NOERROR;
		}
	return CONNECT_E_NOCONNECTION;
}

static HRESULT STDMETHODCALLTYPE Generic_IConnectionPoint_EnumConnections(IConnectionPoint * This, IEnumConnections * * ppEnum)
{
	Generic_IConnectionPoint *ICP = (Generic_IConnectionPoint *)This;
	Generic_IEnumConnections *IEnumC = (Generic_IEnumConnections *)This;
	INT count1, count2;
	CONNECTDATA *pCD;

	vbt_printf("Generic_IConnectionPoint_EnumConnections called\n");
	*ppEnum = NULL;
	pCD = malloc(sizeof(CONNECTDATA)*ICP->m_cConn);
	if (pCD == NULL)
		return ResultFromScode(E_OUTOFMEMORY);
	for(count1 = 0, count2 = 0; count1 < CCONNMAX; count1++)
		if(ICP->m_rgpUnknown[count1] != NULL)
		{
			pCD[count2].pUnk = ICP->m_rgpUnknown[count1];
			pCD[count2].dwCookie = ICP->m_rgnCookies[count1];
			count2++;
		}
	IEnumC = Generic_IEnumConnections_Constructor((IUnknown *)This, ICP->m_cConn, pCD);
	free(pCD);
	if (IEnumC == NULL)
		return ResultFromScode(E_OUTOFMEMORY);
	return IEnumConnections_QueryInterface(&IEnumC->iface, &IID_IEnumConnections, (void **)ppEnum);
}

static const IConnectionPointVtbl IConnectionPointVtblInstance =
BEGIN_VTABLE
	VTABLE_ENTRY( GENERIC_QUERYINTERFACE(IConnectionPoint) ),
	VTABLE_ENTRY( GENERIC_ADDREF(IConnectionPoint) ),
	VTABLE_ENTRY( GENERIC_RELEASE(IConnectionPoint) ),
	VTABLE_ENTRY( Generic_IConnectionPoint_GetConnectionInterface ),
	VTABLE_ENTRY( Generic_IConnectionPoint_GetConnectionPointContainer ),
	VTABLE_ENTRY( Generic_IConnectionPoint_Advise ),
	VTABLE_ENTRY( Generic_IConnectionPoint_Unadvise ),
	VTABLE_ENTRY( Generic_IConnectionPoint_EnumConnections )
END_VTABLE;

static HRESULT STDMETHODCALLTYPE Generic_IConnectionPointContainer_EnumConnectionPoints(IConnectionPointContainer * This, IEnumConnectionPoints * * ppEnum)
{
	Generic_IConnectionPointContainer *ICPC = (Generic_IConnectionPointContainer *)This;
	Generic_IEnumConnectionPoints *IEnumCP = (Generic_IEnumConnectionPoints *)This;

	vbt_printf("Generic_IConnectionPointContainer_EnumConnectionPoints called\n");
	IEnumCP = Generic_IEnumConnectionPoints_Constructor((IUnknown *)This, ICPC->m_rgpConnPt);
	if (IEnumCP == NULL)
		return ResultFromScode(E_OUTOFMEMORY);
	return IEnumConnectionPoints_QueryInterface(&IEnumCP->iface, &IID_IEnumConnectionPoints, (void**)ppEnum);
}

static HRESULT STDMETHODCALLTYPE Generic_IConnectionPointContainer_FindConnectionPoint(IConnectionPointContainer * This, REFIID riid, IConnectionPoint * * ppCP)
{
	Generic_IConnectionPointContainer *gThis = (Generic_IConnectionPointContainer *)This;
	UINT u;

	if (riid == NULL)
	{
		vbt_printf("Generic_IConnectionPointContainer_FindConnectionPoint: riid=NULL!!!\n");
		return E_NOINTERFACE; /* temp!! */
	}
	{
		LPOLESTR iid;
		StringFromIID(riid, &iid);
		vbt_printf("Generic_IConnectionPointContainer_FindConnectionPoint: riid=%s\n",wtos(iid));
	}
	for(u=0;u<gThis->m_cICP;u++)
	{
		if(IsEqualGUID(riid,*((Generic_IConnectionPoint *)gThis->m_rgpConnPt[u])->m_iid))
		{
			return IConnectionPoint_QueryInterface(&((Generic_IConnectionPoint *)gThis->m_rgpConnPt[u])->iface, &IID_IConnectionPoint, (void**)ppCP);
		}
	}
	return E_NOINTERFACE;
}

static const IConnectionPointContainerVtbl IConnectionPointContainerVtblInstance =
BEGIN_VTABLE
	VTABLE_ENTRY( GENERIC_QUERYINTERFACE(IConnectionPointContainer) ),
	VTABLE_ENTRY( GENERIC_ADDREF(IConnectionPointContainer) ),
	VTABLE_ENTRY( GENERIC_RELEASE(IConnectionPointContainer) ),
	VTABLE_ENTRY( Generic_IConnectionPointContainer_EnumConnectionPoints ),
	VTABLE_ENTRY( Generic_IConnectionPointContainer_FindConnectionPoint )
END_VTABLE;

/* make macros for these? */
static const InterfaceDescTable Generic_IUnknown_InterfaceDesc =
{
	{ (IUnknownVtbl *)&IUnknownVtblInstance }, sizeof(Generic_IUnknown), (void *)Generic_IUnknown_Constructor, (void *)Generic_IUnknown_Destructor, L"IUnknown", &IID_IUnknown, 0, NULL
};
static const InterfaceDescTable Generic_IDispatch_InterfaceDesc =
{
	{ (IUnknownVtbl *)&IDispatchVtblInstance }, sizeof(Generic_IDispatch), (void *)Generic_IDispatch_Constructor, (void *)Generic_IDispatch_Destructor, L"IDispatch", &IID_IDispatch, 0 , NULL
};
static const InterfaceDescTable Generic_IConnectionPoint_InterfaceDesc =
{
	{ (IUnknownVtbl *)&IConnectionPointVtblInstance }, sizeof(Generic_IConnectionPoint), (void *)Generic_IConnectionPoint_Constructor, (void *)Generic_IConnectionPoint_Destructor, L"IConnectionPoint", &IID_IConnectionPoint, 0, NULL
};
static const InterfaceDescTable Generic_IConnectionPointContainer_InterfaceDesc =
{
	{ (IUnknownVtbl *)&IConnectionPointContainerVtblInstance }, sizeof(Generic_IConnectionPointContainer), (void *)Generic_IConnectionPointContainer_Constructor, (void *)Generic_IConnectionPointContainer_Destructor, L"IConnectionPointContainer", &IID_IConnectionPointContainer, 0, NULL
};

static const InterfaceDescTable * const Generic_InternalInterfaceDescs[2] =
{
	&Generic_IDispatch_InterfaceDesc,
	&Generic_IConnectionPointContainer_InterfaceDesc
};

static UINT Generic_InternalInterfaceCount = 2;

static Generic_IUnknown *Generic_IUnknown_Constructor(IUnknown *This)
{
	vbt_printf("Generic_IUnknown_Constructor called\n");
	return Generic_Constructor(This, &Generic_IUnknown_InterfaceDesc);
}

static void Generic_IUnknown_Destructor(IUnknown *This)
{
	vbt_printf("Generic_IUnknown_Destructor called\n");
	Generic_Destructor(This);
}

static Generic_IDispatch *Generic_IDispatch_Constructor(IUnknown *This)
{
	Generic_IDispatch *gThis;
	HRESULT hr;
	vbt_printf("Generic_IDispatch_Constructor called\n");
	gThis =	Generic_Constructor(This, &Generic_IDispatch_InterfaceDesc);
	if (gThis == NULL)
		return NULL;
/* don't know how to get default IDispatch interface ITypeInfo - so kludge - assume first class interface is default */
	hr = ITypeLib_GetTypeInfoOfGuid(gThis->gid.pITypeLib, gThis->gid.ct->InterfaceDescs[0]->InterfaceGUID, &gThis->m_pITypeInfo);
	vbt_printf("ITypeLib_GetTypeInfoOfGuid: hr=%lx\n",hr);
	if (FAILED(hr))
		return NULL;
/* using class's first interface for Vtbl */
	gThis->iface = *(IDispatch *)&gThis->gid.ct->InterfaceDescs[0]->iface;
	return gThis;
}

static void Generic_IDispatch_Destructor(IDispatch *This)
{
	Generic_IDispatch *gThis = (Generic_IDispatch *)This;
	vbt_printf("Generic_IDispatch_Destructor called\n");
	if (gThis->m_pITypeInfo != NULL)
		ITypeInfo_Release(gThis->m_pITypeInfo);
	Generic_Destructor((IUnknown *)This);
}

static Generic_IConnectionPoint *Generic_IConnectionPoint_Constructor(IUnknown *This, const InterfaceDescTable * const idt)
{
	Generic_IConnectionPoint *gThis;
	vbt_printf("Generic_IConnectionPoint_Constructor called\n");
	gThis = Generic_Constructor(This, &Generic_IConnectionPoint_InterfaceDesc);
	if (gThis == NULL)
		return NULL;
	gThis->m_iid = &idt->InterfaceGUID;
	return gThis;
}

static void Generic_IConnectionPoint_Destructor(IConnectionPoint *This)
{
	vbt_printf("Generic_IConnectionPoint_Destructor called\n");
	Generic_Destructor((IUnknown *)This);
}

static Generic_IConnectionPointContainer *Generic_IConnectionPointContainer_Constructor(IUnknown *This)
{
	Generic_IConnectionPointContainer *gThis;
	UINT u;
	HRESULT hr;

	vbt_printf("Generic_IConnectionPointContainer_Constructor called\n");
	gThis = Generic_Constructor(This, &Generic_IConnectionPointContainer_InterfaceDesc);
	if (gThis == NULL)
		return NULL;
	gThis->m_cICP = *gThis->gid.ct->InterfaceCount;
	gThis->m_rgpConnPt = malloc(gThis->m_cICP*sizeof(IConnectionPoint *));
	if (gThis->m_rgpConnPt == NULL)
		return NULL;
	for(u=0;u<gThis->m_cICP;u++)
	{
		Generic_IUnknown *gIUnk;
		Generic_IConnectionPoint *gICP;
		gIUnk = Generic_IUnknown_Constructor(NULL);
		if (gIUnk == NULL)
			return NULL;
		gICP = Generic_IConnectionPoint_Constructor(&gIUnk->iface, gThis->gid.ct->InterfaceDescs[u]);
		if (gICP == NULL)
			return NULL;
		hr = IConnectionPoint_QueryInterface(&gICP->iface, &IID_IConnectionPoint, (void **)(gThis->m_rgpConnPt+u));
		if (FAILED(hr))
			return NULL;
	}
	return gThis;
}

void Generic_IConnectionPointContainer_Destructor(IConnectionPointContainer *This)
{
	Generic_IConnectionPointContainer *gThis = (Generic_IConnectionPointContainer *)This;
	UINT u;
	vbt_printf("Generic_IConnectionPointContainer_Destructor called\n");
	for(u=0;u<gThis->m_cICP;u++)
	{
		free(gThis->m_rgpConnPt[u]);
	}
	free(gThis->m_rgpConnPt);
	Generic_Destructor((IUnknown *)This);
vbt_printf("%s:%d\n",__FILE__,__LINE__);
}

static HRESULT STDMETHODCALLTYPE Generic_IClassFactory_QueryInterface(IClassFactory * This, REFIID riid, void ** ppvObject)
{
vbt_printf("Generic_IClassFactory_QI\n");
	{
		LPOLESTR iid;
		StringFromIID(riid, &iid);
		vbt_printf("Generic_IClassFactory_QueryInterface: riid=%s ",wtos(iid));
	}
	if ( IsEqualGUID(riid, &IID_IUnknown) || IsEqualGUID(riid, &IID_IClassFactory) )
	{
		*ppvObject = This;
		Generic_IClassFactory_AddRef(This);
		vbt_printf("ppvObject=%lx\n",*ppvObject);
		return NOERROR;
	}
	else
	{
		*ppvObject = NULL;
		vbt_printf("Not found! hr=%lx\n",ResultFromScode(E_NOINTERFACE));
		return ResultFromScode(E_NOINTERFACE);
	}
}

static ULONG STDMETHODCALLTYPE Generic_IClassFactory_AddRef(IClassFactory * This)
{
	vbt_printf("Generic_IClassFactory_AddRef called\n");
	return 2; /*ServerAddRef(This);*/
}

static ULONG STDMETHODCALLTYPE Generic_IClassFactory_Release(IClassFactory * This)
{
	vbt_printf("Generic_IClassFactory_Release called\n");
	return 1; /*ServerRelease(This);*/
}

static HRESULT STDMETHODCALLTYPE Generic_IClassFactory_LockServer(IClassFactory * This, BOOL fLock)
{
	vbt_printf("Generic_IClassFactory_LockServer called\n");
	if (fLock)
		ServerAddRef((Generic_IUnknown *)This);
	else
		ServerRelease((Generic_IUnknown *)This);
	return NOERROR;
}

static HRESULT STDMETHODCALLTYPE Generic_IClassFactory_CreateInstance(IClassFactory * This, IUnknown *pUnkOuter, REFIID riid, void **ppvObject)
{
	Generic_IClassFactory *gThis = (Generic_IClassFactory *)This;
	HRESULT hr;
	Generic_IUnknown *newIUnk;
	UINT u;
	{
		LPOLESTR iid;
		StringFromIID(riid, &iid);
		vbt_printf("Generic_IClassFactory_CreateInstance: iid=%s\n",wtos(iid));
	}
	*ppvObject = 0;
	if (pUnkOuter)
		return ResultFromScode(CLASS_E_NOAGGREGATION);
	for(u=0;u<*gThis->gid.ct->InterfaceCount;u++)
	{
		if (gThis->gid.ct->InterfaceDescs[u]->constructor == NULL)
			newIUnk = Generic_Constructor((IUnknown *)This, gThis->gid.ct->InterfaceDescs[u]);
		else
			newIUnk = (*Generic_InternalInterfaceDescs[u]->constructor)((IUnknown *)This);
		if (newIUnk == NULL)
		{
			Generic_IClassFactory_Destructor(gThis);
			return ResultFromScode(E_OUTOFMEMORY);
		}
	}
	for(u=0;u<Generic_InternalInterfaceCount;u++)
	{
		newIUnk = (*Generic_InternalInterfaceDescs[u]->constructor)((IUnknown *)This);
		if (newIUnk == NULL)
		{
			Generic_IClassFactory_Destructor(gThis);
			return ResultFromScode(E_OUTOFMEMORY);
		}
	}
	hr = IUnknown_QueryInterface(&newIUnk->iface, riid, ppvObject);
	if (FAILED(hr))
	{
		Generic_IClassFactory_Destructor(gThis);
		return hr;
	}
	return hr;
}

static const IClassFactoryVtbl Generic_IClassFactoryVtbl =
BEGIN_VTABLE
	VTABLE_ENTRY( Generic_IClassFactory_QueryInterface ),
	VTABLE_ENTRY( Generic_IClassFactory_AddRef ),
	VTABLE_ENTRY( Generic_IClassFactory_Release ),
	VTABLE_ENTRY( Generic_IClassFactory_CreateInstance ),
	VTABLE_ENTRY( Generic_IClassFactory_LockServer )
END_VTABLE;

static /*const*/ InterfaceDescTable Generic_IClassFactory_InterfaceDesc =
{
	{ (IUnknownVtbl *)&Generic_IClassFactoryVtbl }, sizeof(Generic_IClassFactory), (void *)Generic_IClassFactory_Constructor, (void *)Generic_IClassFactory_Destructor, L"IClassFactory", &IID_IClassFactory, 0, NULL
};

void *Generic_IClassFactory_Constructor(IUnknown *This, const ClassTable * const ct, ITypeLib *pITypeLib, HANDLE hEvent)
{
Generic_IUnknown *gIUnk;
Generic_IClassFactory *gThis;

	gIUnk = Generic_IUnknown_Constructor(NULL);
	if (gIUnk == NULL)
		return NULL;
	gThis = Generic_Constructor(&gIUnk->iface, &Generic_IClassFactory_InterfaceDesc);
	if (gThis == NULL)
		return NULL;
	gThis->gid.ct = ct;
	gThis->gid.pITypeLib = pITypeLib;
	gThis->gid.hEvent = hEvent;
	return gThis;
};

void Generic_IClassFactory_Destructor(Generic_IClassFactory *gThis)
{
	Generic_Destructor((IUnknown *)gThis);
};

char *wtos(w)
const wchar_t *w;
{
	static char s[128];
	size_t l;
	if (w == NULL)
		return(strcpy(s,"(null)"));
	l = wcstombs(s,w,sizeof(s)-1);
	if (l == -1)
		*s = 0;
	else
		s[l] = 0;
	return(s);
}

#ifdef va_dcl
INT vbt_printf(format, va_alist)
const TEXT *format;
va_dcl
#else
INT vbt_printf(const char *format, ...)
#endif
{
va_list ap;
INT ret;

/*printf("format=%s\n",format); */
	if (vbt_debugpf == NULL)
		return(-1);
#ifdef va_dcl
	va_start(ap);
#else
	va_start(ap, format);
#endif
	ret = vfprintf(vbt_debugpf,format,ap);
	va_end(ap);
	return(ret);
}

void MethodInitialize(const struct methodinfo_t *mi, void *lv)
{
UINT i;
	printf("MethodInitialize: i=%S m=%S lv=%lx\n",mi->mi_interfaceName, mi->mi_methodName, lv);
	for(i=0;i<mi->mi_nlv;i++)
	{
		printf("lv init: offset=%lx name=%S vt=%d\n",mi->mi_lv[i].lv_offset,mi->mi_lv[i].lv_name,mi->mi_lv[i].lv_vt);
		memset(lv,0,mi->mi_szlv);
	}
}

void MethodTerminate(const struct methodinfo_t *mi, void *lv)
{
UINT i;
	printf("MethodTerminate: i=%s",wtos(mi->mi_interfaceName));
	printf(" m=%s lv=%lx\n", wtos(mi->mi_methodName), lv);
	for(i=0;i<mi->mi_nlv;i++)
	{
		printf("lv init: offset=%lx name=%S vt=%d\n",mi->mi_lv[i].lv_offset,mi->mi_lv[i].lv_name,mi->mi_lv[i].lv_vt);
		printf("%x\n",*(INT *)((unsigned char *)lv+mi->mi_lv[i].lv_offset));
		if (mi->mi_lv[i].lv_vt == VT_BSTR)
			printf("%S\n",*(BSTR *)(unsigned char *)lv+mi->mi_lv[i].lv_offset);
		else
		{
		VARIANT vres,vsrc;
		VariantInit(&vres);
		VariantInit(&vsrc);
		V_VT(&vsrc) = mi->mi_lv[i].lv_vt | VT_BYREF;
		V_BYREF(&vsrc) = (unsigned char *)lv+mi->mi_lv[i].lv_offset;
		if (SUCCEEDED(VariantChangeType(&vres,&vsrc,0,VT_BSTR)))
			printf("%ls\n",V_BSTR(&vres));
		VariantClear(&vres);
		VariantClear(&vsrc);
		}
	}
}

BSTR vbtStrCat(const BSTR s1, const BSTR s2)
{
#if 1
BSTR res;
	if (VarBstrCat(s1,s2,&res))
		abort();
	return(res);
#else
size_t l1,l2;
BSTR bstr;

	l1 = SysStringLen(s1);
	l2 = SysStringLen(s2);
	bstr = SysAllocStringLen(s1,l1+l2);
	wcsncpy(bstr+l1,s2,l2);
	return(bstr);
#endif
}

VARIANT vbtBoolToVar(VARIANT_BOOL b)
{
VARIANT v;
	V_VT(&v) = VT_BOOL;
	V_BOOL(&v) = b;
	return v;
}

VARIANT vbtIntToVar(INT16 i16)
{
VARIANT v;
	V_VT(&v) = VT_I2;
	V_I2(&v) = i16;
	return v;
}

VARIANT vbtLngToVar(INT32 l)
{
VARIANT v;
	V_VT(&v) = VT_I4;
	V_I4(&v) = l;
	return v;
}

VARIANT vbtStrToVar(BSTR str)
{
VARIANT v;
	V_VT(&v) = VT_BSTR;
	V_BSTR(&v) = SysAllocString(str);
	return v;
}

#ifdef NEVER /* w stuff isn't resolved */
void StrLSet(LPOLESTR lv, size_t szlv, LPOLESTR v, size_t szv)
{
	if (szlv <= szv)
		wmemcpy(lv,v,szv);
	else
	{
		wmemcpy(lv,v,szv);
		wmemset(lv+szv,L' ',szlv-szv);
	}
}

void UDTLSet(unsigned char *lv, size_t szlv, unsigned char *v, size_t szv)
{
	if (szlv <= szv)
		memcpy(lv,v,szv);
	else
	{
		memcpy(lv,v,szv);
		memset(lv+szv,L' ',szlv-szv);
	}
}

void StrRSet(LPOLESTR lv, size_t szlv, LPOLESTR v, size_t szv)
{
	if (szlv <= szv)
		wmemcpy(lv,v,szv);
	else
	{
		wmemcpy(lv+szlv-szv,v,szv);
		wmemset(lv,L' ',szlv-szv);
	}
}
#endif

VARIANT vbtVarInt(VARIANT *arg)
{
#undef VarInt /* undef until vbt.h code gen stuff is separated from vbt defs */
	if (VarInt(arg, arg))
		abort();
	return(*arg);
}

INT16 vbtVarToInt(VARIANT v)
{
VARIANT Dest;
	VariantInit(&Dest);
	if (VariantChangeType(&Dest,&v,0,VT_I2))
		abort();
	return(V_I2(&Dest));
}

INT32 vbtVarToLng(VARIANT v)
{
VARIANT Dest;
	VariantInit(&Dest);
	if (VariantChangeType(&Dest,&v,0,VT_I4))
		abort();
	return(V_I4(&Dest));
}

DOUBLE vbtVarToDbl(VARIANT v)
{
VARIANT Dest;
	VariantInit(&Dest);
	if (VariantChangeType(&Dest,&v,0,VT_R8))
		abort();
	return(V_R8(&Dest));
}

BSTR vbtVarToStr(VARIANT v)
{
VARIANT Dest;
	VariantInit(&Dest);
	if (VariantChangeType(&Dest,&v,0,VT_BSTR))
		abort();
	return(V_BSTR(&Dest));
}

VARIANT vbtSngToVar(FLOAT f)
{
VARIANT Dest;
	V_VT(&Dest) = VT_R4;
	V_R4(&Dest) = f;
	return(Dest);
}

VARIANT vbtDblToVar(DOUBLE d)
{
VARIANT Dest;
	V_VT(&Dest) = VT_R8;
	V_R8(&Dest) = d;
	return(Dest);
}

BSTR vbtIntToStr(INT16 i16)
{
BSTR bstr;
	VarBstrFromI2(i16,0,0,&bstr);
	return(SysAllocString(bstr));
}

BSTR vbtSngToStr(FLOAT f)
{
BSTR bstr;
	VarBstrFromR4(f,0,0,&bstr);
	return(SysAllocString(bstr));
}

INT16 m_fn;
void vbtPrintStart(INT16 fn)
{
	m_fn = fn;
}

void vbtPrintExpr(VARIANT v)
{
	VariantChangeType(&v,&v,0,VT_BSTR);
	printf("%s",wtos(V_BSTR(&v)));
}

void vbtPrintComma()
{
	printf("\t");
}

void vbtPrintNL()
{
	printf("\n");
}

void vbtPrintEnd()
{
}

INT32 vbtMsgBox(VARIANT prompt, INT16 buttons, VARIANT title, VARIANT helpfile, VARIANT context)
{
	return(MessageBox(0,V_BSTR(&prompt),V_BSTR(&title),buttons));
}

#if 0
void vbtStrLSet(BSTR lv, size_t szlv, BSTR v)
{
size_t l;
	l = SysStringLen(v);
	if (szlv <= l)
		wmemcpy(lv,v,szlv);
	else
	{
		wmemcpy(lv,v,l);
		wmemset(lv+l,L' ',szlv-l);
	}
}

void vbtStrRSet(BSTR lv, size_t szlv, BSTR v)
{
size_t l;
	l = SysStringLen(v);
	if (szlv <= l)
		wmemcpy(lv,v,szlv);
	else
	{
		wmemset(lv,L' ',szlv-l);
		wmemcpy(lv+szlv-l,v,l);
	}
}
#endif

void vbtUDTLSet(void *lv, size_t szlv, void *v, size_t szv)
{
	if (szlv <= szv)
		memcpy(lv,v,szlv);
	else
	{
		memcpy(lv,v,szv);
		memset((UINT8 *)lv+szv,0,szlv-szv);
	}
}

INT16 IntByRefs[500000],*pIntByRefs = IntByRefs;
INT32 LngByRefs[500000],*pLngByRefs = LngByRefs;
FLOAT SngByRefs[500000],*pSngByRefs = SngByRefs;
DOUBLE DblByRefs[500000],*pDblByRefs = DblByRefs;
BSTR StrByRefs[500000],*pStrByRefs = StrByRefs;
VARIANT VarByRefs[500000],*pVarByRefs = VarByRefs;

void eos()
{
	pIntByRefs = IntByRefs;
	pLngByRefs = LngByRefs;
	pSngByRefs = SngByRefs;
	pDblByRefs = DblByRefs;
	pStrByRefs = StrByRefs;
	pVarByRefs = VarByRefs;
}

INT16 *IntByRef(INT16 i16)
{
	*pIntByRefs = i16;
	return(pIntByRefs++);
}

INT32 *LngByRef(INT32 i32)
{
	*pLngByRefs = i32;
	return(pLngByRefs++);
}

FLOAT *SngByRef(FLOAT f)
{
	*pSngByRefs = f;
	return(pSngByRefs++);
}

DOUBLE *DblByRef(DOUBLE d)
{
	*pDblByRefs = d;
	return(pDblByRefs++);
}

BSTR *StrByRef(BSTR bstr)
{
	*pStrByRefs = bstr;
	return(pStrByRefs++);
}

VARIANT *VarByRef(VARIANT v)
{
	*pVarByRefs = v;
	return(pVarByRefs++);
}

IDispatch *vbtObjNew(wchar_t *progID)
{
	/* fixme: transpose logic between here and VBA_CreateObject */
VARIANT WINAPI VBA_CreateObject(BSTR Class, BSTR ServerName);
VARIANT v;
	v = VBA_CreateObject(progID, NULL);
	return(V_DISPATCH(&v));
}

HRESULT vbtObjSet(IUnknown *ro, IUnknown **lo, LPWSTR wiid)
{
IID iid;
HRESULT hr;
	if (wiid)
	{
		hr = IIDFromString(wiid,&iid); /* could fail */
		if (SUCCEEDED(hr))
			hr = IUnknown_QueryInterface(ro,&iid,(void **)lo); /* could fail */
	}
	else
	{
		hr = IUnknown_QueryInterface(ro,&IID_IDispatch,(void **)lo); /* could fail */
	}
	return(hr);
}

void IntSelectExpr(int e)
{
static int e2;
	e2 = e;
}

VARIANT Invoke(IDispatch *pdisp, LPWSTR wfuncname, UINT dispatch_type, UINT nargs, ...)
{
va_list ap;
DISPPARAMS dispparams;
VARIANT v;
HRESULT hr;
UINT err;
EXCEPINFO excp;
DISPID dispid[1];
DISPID temp_dispid = DISPID_PROPERTYPUT;
UINT u;
VARIANTARG rgvarg[64]; /* use maximum number of args */

	va_start(ap, nargs);
	VariantInit(&v);
	V_VT(&v) = VT_ERROR;
	V_ERROR(&v) = 0x1234; /* ??? */
	vbt_printf("Invoke: wfuncname=%s pdisp=%lx &IID_NULL=%lx\n",wtos(wfuncname),pdisp,&IID_NULL);
	if (pdisp == NULL)
		return v;
	memset(&dispparams,0,sizeof(DISPPARAMS));
	dispparams.rgvarg = rgvarg;
    vbt_printf("Invoke: nargs=%u dispparams.rgvarg=%lx\n",nargs,dispparams.rgvarg);
	if (*wfuncname)
	{
		hr = IDispatch_GetIDsOfNames(pdisp,&IID_NULL,&wfuncname,1,LOCALE_SYSTEM_DEFAULT,dispid);
		vbt_printf("Invoke: GetIDsOfNames: hr=%lx dispid=%lx\n",hr,dispid[0]);
		if (hr)
			return v;
	}
	else
		dispid[0] = 0; /* default property */

	dispparams.cArgs = nargs;
	while(nargs--)
	{
		dispparams.rgvarg[nargs] = va_arg(ap, VARIANT);
	}
	/* invoke the member */
	switch (dispatch_type)
		{
		case DISPATCH_PROPERTYPUT:
			dispparams.rgdispidNamedArgs = &temp_dispid;
			dispparams.cNamedArgs = 1;
		case DISPATCH_METHOD:
		case DISPATCH_PROPERTYGET:
		case DISPATCH_METHOD | DISPATCH_PROPERTYGET:
			vbt_printf("Invoke: pdisp=%lx dispid[0]=%x dispatch_type=%d dispparams=%lx\n",pdisp,dispid[0],dispatch_type,dispparams);
			vbt_printf("\tdisparams: rgvarg=%lx rgdispidNamedArgs=%lx cArgs=%d cNamedArgs=%x\n",dispparams.rgvarg,dispparams.rgdispidNamedArgs,dispparams.cArgs,dispparams.cNamedArgs);
			for(u=0;u<dispparams.cArgs;u++)
				vbt_printf("\t\trgvarg[%d]: vt=%x I4=%x I4REF=%x\n",u,V_VT(dispparams.rgvarg+u),V_I4(dispparams.rgvarg+u),(0 && V_VT(dispparams.rgvarg+u) & VT_BYREF) ? *V_I4REF(dispparams.rgvarg+u) : 0);
			break;
		default:
			vbt_printf("Invoke: invalid dispatch: %u\n",dispatch_type);
		}

vbt_printf("Invoke: %d\n",__LINE__);
        hr = IDispatch_Invoke(pdisp,dispid[0],&IID_NULL,0,(WORD)dispatch_type,&dispparams,&v,&excp,&err);
vbt_printf("Invoke: %d\n",__LINE__);

#if 0 /* DOS or Windows only */ /* needed to reset fpu rouding */ /* reimplement!!!!!! */
	_fpreset();
#endif

	vbt_printf("\thr=%x err=%x\n",hr,err);
	return(v);
}

void FailedHR(HRESULT hr, LPWSTR name)
{
int len;
char buf[512]; /* is there a max length symbol? */

	len = sprintf(buf,"FailedHR: %s: hr=%lx ",wtos(name),hr);
	FormatMessageA(
		FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS | FORMAT_MESSAGE_MAX_WIDTH_MASK,
		NULL,
		hr,
		MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), /* Default language */
		buf+len,
		sizeof(buf)-len,
		NULL);
	vbt_printf("%s\n",buf);
	exit(0);	/* insert error recovery!!! - longjmp? */
}

VARIANT Null = {{{VT_NULL}}};
VARIANT _VarMissing = {{{VT_ERROR}}};
