
#ifndef _INC_VBT
#define _INC_VBT

#include "vbtcom.h"

#define SW_SECS_PER_DAY (24*60*60)
#define SW_VB_1970 25569L       /* 1 Jan 1970 in VB */
#define VB_SIGNIFICANCE 15
int vbt_diag(int errnum,char *file,int line);
#define diag(errnum) vbt_diag(0/*errnum*/,__FILE__,__LINE__)
extern int g_argc;
extern char *g_argv[];
extern WCHAR *g_env[];
#define diffptr(a,b) (size_t)((CHAR *)(a)-(CHAR *)(b))
wchar_t *wmemchr(const wchar_t *s, wchar_t c, size_t n);
int wmemcmp(const wchar_t *s1, const wchar_t *s2, size_t n);
wchar_t *wmemcpy(wchar_t *s1, const wchar_t *s2, size_t n);
wchar_t *wmemmove(wchar_t *s1, const wchar_t *s2, size_t n);
wchar_t *wmemset(wchar_t *s, wchar_t c, size_t n);
struct tm *vbdatetotm(double vbdate);
double tmtovbdate(struct tm *tm);
wchar_t *_wgetcwd( wchar_t *buffer, int maxlen );wchar_t *_wgetcwd( wchar_t *buffer, int maxlen );
int _wchdir( const wchar_t *dirname );
int wmain( int argc, wchar_t *argv[ ], wchar_t *envp[ ] );
wchar_t *_wfullpath( wchar_t *absPath, const wchar_t *relPath, size_t maxLength );

#define METHOD_INITIALIZE(pmi,plv) MethodInitialize(pmi,plv);
/* goto _method_terminate is used to suppress unref label warning message */
#define METHOD_TERMINATE(pmi,plv) MethodTerminate(pmi,plv); goto _method_terminate; _method_terminate: ;
#define LToStr(s) SysAllocString(s)
DATE vbtLToDate(wchar_t *datestring);
#define LToDate(s) vbtLToDate(s)
#define StrNToStr(s) SysAllocStringLen(s,(sizeof(s)+1)/2)
#define StrToStrN(s) (s)
#define StrRefToStrN(s) (*(s))
#define StrNLet(s,sn) wmemcpy(sn,s,sizeof(sn));
#define StrNRefLet(s,sn) wmemcpy(*(sn),s,sizeof(sn));

#ifdef NEVER
const struct literalpool_t
{
	UINT32 lp_size;
	OLECHAR lp_str[];
};
#endif
/* const */ struct localvars_t
{
	size_t lv_offset;
	LPCOLESTR lv_name;
	VARTYPE lv_vt;
	void * const lv_array;
};
/* const */ struct methodinfo_t
{
	LPCOLESTR mi_interfaceName;
	LPCOLESTR mi_methodName;
	const struct localvars_t *mi_lv;
	UINT mi_nlv;
	size_t mi_szlv;
}; 

#ifndef offsetof
#define offsetof(s,m) (size_t)&(((s *)0)->m)
#endif

typedef const struct tagFuncTable
{
    const FUNCDESC * const FuncDescs;
    const LPCOLESTR * const FuncNames;
} FuncTable;

typedef struct tagInterfaceDescTable
{
	const IUnknown iface;
	const size_t InterfaceDataSize;
	void *(*constructor)(IUnknown *iface, ...);
	void (*destructor)(void *iface);
	LPCOLESTR InterfaceName;
	REFIID InterfaceGUID;
	const UINT FuncCount;
	FuncTable * const Funcs;
} InterfaceDescTable;

typedef struct tagClassTable
{
	LPCOLESTR className;
	REFCLSID classGUID;
	LPCOLESTR ProgID;
	LPCOLESTR VIProgID;
	const UINT * const InterfaceCount;
	const InterfaceDescTable * const * const InterfaceDescs;
} ClassTable;

typedef struct tagApplicationData
{
	LPCOLESTR appName;
} ApplicationData;

typedef const struct tagProjectTable
{
	LPCOLESTR TypeLibName;
	REFIID TypeLibGUID;
	LPCOLESTR TypeLibFile;
	unsigned short TypeLibMajorVersion;
	unsigned short TypeLibMinorVersion;
	const UINT *ClassCount;
	const ClassTable * const * const ClassTabs;
	const ApplicationData * const AppData;
	/* const */ void (STDMETHODCALLTYPE * const StartupObject)(void);
} ProjectTable;

#define GENERIC_QUERYINTERFACE(iface) (HRESULT (STDMETHODCALLTYPE *)(iface *This,REFIID riid,void **ppvObject))Generic_IUnknown_QueryInterface
#define GENERIC_ADDREF(iface) (ULONG (STDMETHODCALLTYPE *)(iface *This))Generic_IUnknown_AddRef
#define GENERIC_RELEASE(iface) (ULONG (STDMETHODCALLTYPE *)(iface *This))Generic_IUnknown_Release
typedef struct tagGenericInterfaceData
{
	const InterfaceDescTable *idt;
	const ClassTable *ct;
	/* circular list of Generic_IUnknowns */
	struct tagGeneric_IUnknown *next;
	struct tagGeneric_IUnknown *prev;
	ITypeLib *pITypeLib;
	ULONG m_cRef;
	HANDLE hEvent; /* WaitForSingleObject event */
} GenericInterfaceData;

typedef struct tagGeneric_IUnknown
{
	IUnknown iface;
	GenericInterfaceData gid;
} Generic_IUnknown;

typedef struct tagGeneric_IClassFactory
{
	IClassFactory iface;
	GenericInterfaceData gid;
} Generic_IClassFactory;

typedef struct tagGeneric_IDispatch
{
	IDispatch iface;
	GenericInterfaceData gid;
	ITypeInfo *m_pITypeInfo;
} Generic_IDispatch;

typedef struct tagGeneric_IEnumConnections
{
	IEnumConnections iface;
	GenericInterfaceData gid;
	CONNECTDATA* m_rgConnData; /* Source of connections */
	UINT m_iCur;
	UINT m_cConn;
} Generic_IEnumConnections;

#define NUM_CONNECTION_POINTS 1

typedef struct tagGeneric_IEnumConnectionPoints
{
	IEnumConnectionPoints iface;
	GenericInterfaceData gid;
	UINT m_iCur;
	IConnectionPoint *m_rgpCP[NUM_CONNECTION_POINTS];  /* Array of connection points */
} Generic_IEnumConnectionPoints;

#define CCONNMAX 2

typedef struct tagGeneric_IConnectionPoint
{
	IConnectionPoint iface;
	GenericInterfaceData gid;
	REFIID *m_iid;
	IConnectionPoint *m_pObj;
	INT m_cConn;
	INT m_nCookieNext;
	UINT m_rgnCookies[CCONNMAX];
	IUnknown* m_rgpUnknown[CCONNMAX];
} Generic_IConnectionPoint;

typedef struct tagGeneric_IConnectionPointContainer
{
	IConnectionPointContainer iface;
	GenericInterfaceData gid;
	UINT m_cICP;
	IConnectionPoint **m_rgpConnPt;
} Generic_IConnectionPointContainer;

extern ProjectTable g_ProjectTable;

HRESULT ProcessRegistration(INT argc, wchar_t *wargv[], ITypeLib **ppITypeLib);
HRESULT vbtCreateTypeLib(LPOLESTR TypeLibFile);
HRESULT STDMETHODCALLTYPE Generic_IUnknown_QueryInterface(IUnknown * This, REFIID riid, void ** ppvObject);
ULONG STDMETHODCALLTYPE Generic_IUnknown_AddRef(IUnknown * This);
ULONG STDMETHODCALLTYPE Generic_IUnknown_Release(IUnknown * This);
HRESULT STDMETHODCALLTYPE Generic_IDispatch_GetTypeInfoCount(IDispatch * This, UINT *pctinfo);
HRESULT STDMETHODCALLTYPE Generic_IDispatch_GetTypeInfo(IDispatch * This, UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo);
HRESULT STDMETHODCALLTYPE Generic_IDispatch_GetIDsOfNames(IDispatch * This, REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispID);
HRESULT STDMETHODCALLTYPE Generic_IDispatch_Invoke(IDispatch * This, DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);
void *Generic_IClassFactory_Constructor(IUnknown *This, const ClassTable * const ct, ITypeLib *pITypeLib, HANDLE hEvent);
void Generic_IClassFactory_Destructor(Generic_IClassFactory *gThis);
INT vbt_printf(const char *format, ...);
char *wtos(const wchar_t *w);
void MethodInitialize(const struct methodinfo_t *mi, void * lv);
void MethodTerminate(const struct methodinfo_t *mi, void * lv);

/* Global variables */
extern FILE *vbt_debugpf;

/* typedefs */
typedef short INT16;
typedef int INT32;
typedef long Enum_t;

/* Declaration Section */
#define Public
#define Private
#define Const(name,type,value) /* const type name = value; - don't need to define */
#define Enum(name) /* enum name { */
#define EndEnum(name) /* }; */
#define Type(name) typedef struct __t_##name {
#define EndType(name) } _t_##name;
#define Declare
#define Function(name,type,args) WINUSERAPI type WINAPI name args;
#define Sub(name,args) Function(name,void,args)

/* Procedure Section */
/* l-value assignment defs */
#define ByteLet(v,lv) ((lv) = (v));
#define ByteRefLet(v,lv) (*(lv) = (v));
#define BoolLet(v,lv) ((lv) = (v));
#define BoolRefLet(v,lv) (*(lv) = (v));
#define IntLet(v,lv) ((lv) = (v));
#define IntRefLet(v,lv) (*(lv) = (v));
#define LngLet(v,lv) ((lv) = (v));
#define LngRefLet(v,lv) (*(lv) = (v));
#define SngLet(v,lv) ((lv) = (v));
#define SngRefLet(v,lv) (*(lv) = (v));
#define DblLet(v,lv) ((lv) = (v));
#define DblRefLet(v,lv) (*(lv) = (v));
#define DateLet(v,lv) ((lv) = (v));
#define DateRefLet(v,lv) (*(lv) = (v));
#define CurLet(v,lv) ((lv) = (v));
#define CurRefLet(v,lv) (*(lv) = (v));
#define StrLet(v,lv) ((lv) = (v));
#define StrRefLet(v,lv) (*(lv) = (v));
#define StrLetN(v,lv) wcsncpy(lv,v,sizeof(lv));
void vbtStrLSet(void *lv,size_t szlv, void *v);
void vbtStrRSet(void *lv,size_t szlv, void *v);
#define StrLSet(v,lv) vbtStrLSet(&lv,sizeof(lv),v);
#define StrRSet(v,lv) vbtStrRSet(&lv,sizeof(lv),v);
#define StrRSetN(v,lv) StrRSet(v,lv)
#define VarLet(v,lv) ((lv) = (v)); /* hmmm, VariantCopy or struct assignment? */
#define VarRefLet(v,lv) (*(lv) = (v)); /* hmmm, VariantCopy or struct assignment? */
#define UDTLet(v,lv) ((lv) = (v)); /* struct assignment */
#define UDTRefLet(v,lv) (*(lv) = (v)); /* struct assignment */
void vbtUDTLSet(void *lv, size_t szlv, void *v, size_t szv);
#define UDTLSet(v,lv) vbtUDTLSet(&lv,sizeof(lv),&v,sizeof(v));

/* Generic inline conversion defs */
#if 1
#define ToByte(v) (UINT8)(v)
#define ToBool(v) (INT16)(v)
#define ToInt(v) (INT16)(v)
#define ToLng(v) (INT32)(v)
#define ToSng(v) (FLOAT)(v)
#define ToDbl(v) (DOUBLE)(v)
#define ToDate(v) (DATE)(v)
#define ToINT(v) (int)(v) /* fixme: need more automation types */
#define ToUINT16(v) (UINT16)(v)
#define ToUINT32(v) (UINT32)(v)

/* Conversion defs */
/* fixme: many of these conversion require special routine names */
#define ByteToByte(v) ToByte(v)
#define ByteToBool(v) ToBool(v)
#define ByteRefToBool(v) ToBool(*(v))
#define ByteToInt(v) ToInt(v)
#define ByteRefToInt(v) ToInt(*(v))
#define ByteToLng(v) ToLng(v)
#define ByteRefToLng(v) ToLng(*(v))
#define ByteToSng(v) ToSng(v)
#define ByteRefToSng(v) ToSng(*(v))
#define ByteToDbl(v) ToDbl(v)
#define ByteRefToDbl(v) ToDbl(*(v))
#define ByteToDate(v) ToDate(v)
#define ByteToCur(v) ToCur(v)
#define ByteToVoidRef(v) ((void *)(LONG)(v)) /* fixme: As Any casts needed for all data types */

#define BoolToByte(v) ToByte(v)
#define BoolToBool(v) ToBool(v)
#define BoolToInt(v) ToInt(v)
#define BoolRefToInt(v) ToInt(*(v))
#define BoolToLng(v) ToLng(v)
#define BoolToSng(v) ToSng(v)
#define BoolToDbl(v) ToDbl(v)
#define BoolToDate(v) ToDate(v)
#define BoolToCur(v) ToCur(v)
#define BoolToVoidRef(v) ((void *)(LONG)(v)) /* fixme: As Any casts needed for all data types */

#define IntToByte(v) ToByte(v)
#define IntRefToByte(v) ToByte(*(v))
#define IntToBool(v) ToBool(v)
#define IntRefToBool(v) ToBool(*(v))
#define IntToInt(v) ToInt(v)
#define IntToLng(v) ToLng(v)
#define IntRefToLng(v) ToLng(*(v))
#define IntToSng(v) ToSng(v)
#define IntRefToSng(v) ToSng(*(v))
#define IntToDbl(v) ToDbl(v)
#define IntRefToDbl(v) ToDbl(*(v))
#define IntToDate(v) ToDate(v)
CY vbtIntToCur(INT16 v);
#define IntToCur(v) vbtIntToCur(v)
#define IntToVoidRef(v) ((void *)(LONG)(v)) /* fixme: As Any casts needed for all data types */
#define IntRefToVoidRef(v) ((void *)(LONG)(v)) /* fixme: As Any casts needed for all data types */

#define IntSAToVoidRef(v) ((void *)(LONG)(v)) /* fixme: create for all SA types */

#define IntToINT(v) ToINT(v) /* note: INT not Int or int */
#define INTToLng(v) ToLng(v) /* note: INT not Int or int */
VARIANT vbtINTToVar(INT i);
#define INTToVar(v) vbtINTToVar(v) /* note: INT not Int or int */

#define LngToByte(v) ToByte(v)
#define LngRefToByte(v) ToByte(*(v))
#define LngToBool(v) ToBool(v)
#define LngRefToBool(v) ToBool(*(v))
#define LngToInt(v) ToInt(v)
#define LngRefToInt(v) ToInt(*(v))
#define LngToLng(v) ToLng(v)
#define LngToSng(v) ToSng(v)
#define LngRefToSng(v) ToSng(*(v))
#define LngToDbl(v) ToDbl(v)
#define LngRefToDbl(v) ToDbl(*(v))
#define LngToDate(v) ToDate(v)
CY vbtLngToCur(INT32 v);
#define LngToCur(v) vbtLngToCur(v)
#define LngToVoidRef(v) ((void *)(LONG)(v)) /* fixme: As Any casts needed for all data types */
#define LngRefToVoidRef(v) ((void *)(LONG)(v)) /* fixme: As Any casts needed for all data types */
#define LngToINT(v) ToINT(v)
#define LngToUINT32(v) ToUINT32(v) /* fixme: create macros for all automation types? */

#define SngToByte(v) ToByte(v)
#define SngRefToByte(v) ToByte(*(v))
#define SngToBool(v) ToBool(v)
#define SngToInt(v) ToInt(v)
#define SngRefToInt(v) ToInt(*(v))
#define SngToLng(v) ToLng(v)
#define SngRefToLng(v) ToLng(*(v))
#define SngToSng(v) ToSng(v)
#define SngToDbl(v) ToDbl(v)
#define SngRefToDbl(v) ToDbl(*(v))
#define SngToDate(v) ToDate(v)
CY vbtSngToCur(FLOAT v);
#define SngToCur(v) vbtSngToCur(v)
#define SngToVoidRef(v) ((void *)(LONG)(v)) /* fixme: As Any casts needed for all data types */

#define DblToByte(v) ToByte(v)
#define DblToBool(v) ToBool(v)
#define DblToInt(v) ToInt(v)
#define DblRefToInt(v) ToInt(*(v))
#define DblToLng(v) ToLng(v)
#define DblRefToLng(v) ToLng(*(v))
#define DblToSng(v) ToSng(v)
#define DblRefToSng(v) ToSng(*(v))
#define DblToDbl(v) ToDbl(v)
#define DblToDate(v) ToDate(v)
CY vbtDblToCur(double v);
#define DblToCur(v) vbtDblToCur(v)
#define DblToVoidRef(v) ((void *)(LONG)(v)) /* fixme: As Any casts needed for all data types */

#define DateToByte(v) ToByte(v)
#define DateToBool(v) ToBool(v)
#define DateToInt(v) ToInt(v)
#define DateToLng(v) ToLng(v)
#define DateToSng(v) ToSng(v)
#define DateToDbl(v) ToDbl(v)
#define DateToDate(v) ToDate(v)
#define DateToCur(v) ToCur(v)
#define DateToVoidRef(v) ((void *)(LONG)(v)) /* fixme: As Any casts needed for all data types */

/* fixme: need explicit conversion routines for Cur */
#define CurToByte(v) ToByte((v).int64)
#define CurToBool(v) ToBool((v).int64)
#define CurToInt(v) ToInt((v).int64)
#define CurToLng(v) ToLng((v).int64)
#define CurToSng(v) ToSng((v).int64)
#define CurToDbl(v) ToDbl((v).int64)
#define CurToDate(v) ToDate((v).int64)
#define CurToCur(v) ToCur((v).int64)
#define CurToVoidRef(v) ((void *)(LONG)((v).int64)) /* fixme: As Any casts needed for all data types */
#define CurRefToVoidRef(v) ((void *)(LONG)((v)->int64)) /* fixme: As Any casts needed for all data types */

#else

#define ByteToByte(v) vbtByteToByte(v)
#define ByteToBool(v) vbtByteToBool(v)
#define ByteToInt(v) vbtByteToInt(v)
#define ByteToLng(v) vbtByteToLng(v)
#define ByteToSng(v) vbtByteToSng(v)
#define ByteToDbl(v) vbtByteToDbl(v)
#define ByteToDate(v) vbtByteToDate(v)
#define ByteToCur(v) vbtByteToCur(v)

#define BoolToByte(v) vbtBoolToByte(v)
#define BoolToBool(v) vbtBoolToBool(v)
#define BoolToInt(v) vbtBoolToInt(v)
#define BoolToLng(v) vbtBoolToLng(v)
#define BoolToSng(v) vbtBoolToSng(v)
#define BoolToDbl(v) vbtBoolToDbl(v)
#define BoolToDate(v) vbtBoolToDate(v)
#define BoolToCur(v) vbtBoolToCur(v)

#define IntToByte(v) vbtIntToByte(v)
#define IntToBool(v) vbtIntToBool(v)
#define IntToInt(v) vbtIntToInt(v)
#define IntToLng(v) vbtIntToLng(v)
#define IntToSng(v) vbtIntToSng(v)
#define IntToDbl(v) vbtIntToDbl(v)
#define IntToDate(v) vbtIntToDate(v)
#define IntToCur(v) vbtIntToCur(v)

#define LngToByte(v) vbtLngToByte(v)
#define LngToBool(v) vbtLngToBool(v)
#define LngToInt(v) vbtLngToInt(v)
#define LngToLng(v) vbtLngToLng(v)
#define LngToSng(v) vbtLngToSng(v)
#define LngToDbl(v) vbtLngToDbl(v)
#define LngToDate(v) vbtLngToDate(v)
#define LngToCur(v) vbtLngToCur(v)

#define SngToByte(v) vbtSngToByte(v)
#define SngToBool(v) vbtSngToBool(v)
#define SngToInt(v) vbtSngToInt(v)
#define SngToLng(v) vbtSngToLng(v)
#define SngToSng(v) vbtSngToSng(v)
#define SngToDbl(v) vbtSngToDbl(v)
#define SngToDate(v) vbtSngToDate(v)
#define SngToCur(v) vbtSngToCur(v)

#define DblToByte(v) vbtDblToByte(v)
#define DblToBool(v) vbtDblToBool(v)
#define DblToInt(v) vbtDblToInt(v)
#define DblToLng(v) vbtDblToLng(v)
#define DblToSng(v) vbtDblToSng(v)
#define DblToDbl(v) vbtDblToDbl(v)
#define DblToDate(v) vbtDblToDate(v)
#define DblToCur(v) vbtDblToCur(v)

#define DateToByte(v) vbtDateToByte(v)
#define DateToBool(v) vbtDateToBool(v)
#define DateToInt(v) vbtDateToInt(v)
#define DateToLng(v) vbtDateToLng(v)
#define DateToSng(v) vbtDateToSng(v)
#define DateToDbl(v) vbtDateToDbl(v)
#define DateToDate(v) vbtDateToDate(v)
#define DateToCur(v) vbtDateToCur(v)

#define CurToByte(v) vbtCurToByte(v)
#define CurToBool(v) vbtCurToBool(v)
#define CurToInt(v) vbtCurToInt(v)
#define CurToLng(v) vbtCurToLng(v)
#define CurToSng(v) vbtCurToSng(v)
#define CurToDbl(v) vbtCurToDbl(v)
#define CurToDate(v) vbtCurToDate(v)
#define CurToCur(v) vbtCurToCur(v)

#endif

#define ByteToStr(v) vbtByteToStr(v)
#define ByteRefToStr(v) vbtByteToStr(*(v))
#define ByteToVar(v) vbtByteToVar(v)
#define ByteRefToVar(v) vbtByteRefToVar(v)
#define ByteToDec(v) vbtByteToDec(v)

#define BoolToStr(v) vbtBoolToStr(v)
#define BoolRefToStr(v) vbtBoolToStr(*(v))
#define BoolToVar(v) vbtBoolToVar(v)
#define BoolRefToVar(v) vbtBoolToVar(*(v))
#define BoolToDec(v) vbtBoolToDec(v)

#define IntToStr(v) vbtIntToStr(v)
#define IntRefToStr(v) vbtIntRefToStr(v)
#define IntToVar(v) vbtIntToVar(v)
#define IntRefToVar(v) vbtIntRefToVar(v)
#define IntToDec(v) vbtIntToDec(v)

#define LngToStr(v) vbtLngToStr(v)
#define LngRefToStr(v) vbtLngRefToStr(v)
#define LngToVar(v) vbtLngToVar(v)
#define LngRefToVar(v) vbtLngRefToVar(v)
#define LngToDec(v) vbtLngToDec(v)

#define SngToStr(v) vbtSngToStr(v)
#define SngRefToStr(v) vbtSngToStr(*(v))
#define SngToVar(v) vbtSngToVar(v)
#define SngRefToVar(v) vbtSngToVar(*(v))
#define SngToDec(v) vbtSngToDec(v)

#define DblToStr(v) vbtDblToStr(v)
#define DblToVar(v) vbtDblToVar(v)
#define DblRefToVar(v) vbtDblToVar(*(v))
#define DblToDec(v) vbtDblToDec(v)

#define DateToStr(v) vbtDateToStr(v)
#define DateToVar(v) vbtDateToVar(v)
#define DateRefToVar(v) vbtDateRefToVar(v)
#define DateToDec(v) vbtDateToDec(v)

#define CurToStr(v) vbtCurToStr(v)
#define CurToVar(v) vbtCurToVar(v)
#define CurToDec(v) vbtCurToDec(v)

#define StrToByte(v) vbtStrToByte(v)
#define StrToByteSA(v) vbtStrToByteSA(v) /* Only StrToByteSA allowed? */
#define StrToBool(v) vbtStrToBool(v)
#define StrToInt(v) vbtStrToInt(v)
#define StrRefToInt(v) vbtStrToInt(*(v))
#define StrToLng(v) vbtStrToLng(v)
#define StrRefToLng(v) vbtStrToLng(*(v))
#define StrToSng(v) vbtStrToSng(v)
#define StrRefToSng(v) vbtStrToSng(*(v))
#define StrToDbl(v) vbtStrToDbl(v)
#define StrToDate(v) vbtStrToDate(v)
#define StrToCur(v) vbtStrToCur(v)
#define StrToStr(v) vbtStrToStr(v)
#define StrRefToVar(v) vbtStrToVar(*(v))
#define StrToVar(v) vbtStrToVar(v)
#define StrToDec(v) vbtStrToDec(v)
#define StrToVoidRef(v) ((void *)(LONG)(v)) /* fixme: As Any casts needed for all data types */
#define StrRefToVoidRef(v) ((void *)(LONG)(v)) /* fixme: As Any casts needed for all data types */

#define VarToByte(v) vbtVarToByte(v)
#define VarToBool(v) vbtVarToBool(v)
#define VarRefToBool(v) vbtVarToBool(*(v))
#define VarToInt(v) vbtVarToInt(v)
#define VarRefToInt(v) vbtVarToInt(*(v))
#define VarToLng(v) vbtVarToLng(v)
#define VarRefToLng(v) VarToLng(*(v))
#define VarToSng(v) vbtVarToSng(v)
#define VarRefToSng(v) vbtVarToSng(*(v))
#define VarToDbl(v) vbtVarToDbl(v)
#define VarRefToDbl(v) vbtVarToDbl(*(v))
#define VarToDate(v) vbtVarToDate(v)
#define VarToCur(v) vbtVarToCur(v)
#define VarToStr(v) vbtVarToStr(v)
#define VarToStrN(v) VarToStr(v) /* fixme: create more StrN defs. Treat StrN as BSTR??? */
#define VarRefToStr(v) VarToStr(*(v))
#define VarToVar(v) vbtVarToVar(v)
#define VarToDec(v) vbtVarToDec(v)
#define VarToUINT32(v) vbtVarToUINT32(v)
#define VarToVoidRef(v) ((void *)VarToLng(v)) /* fixme: & isn't correct for expr and constants /* fixme: As Any casts needed for all data types */
#define VarRefToVoidRef(v) ((void *)(LONG)(v))

#define DecToByte(v) vbtDecToByte(v)
#define DecToBool(v) vbtDecToBool(v)
#define DecToInt(v) vbtDecToInt(v)
#define DecToLng(v) vbtDecToLng(v)
#define DecToSng(v) vbtDecToSng(v)
#define DecToDbl(v) vbtDecToDbl(v)
#define DecToDate(v) vbtDecToDate(v)
#define DecToCur(v) vbtDecToCur(v)
#define DecToStr(v) vbtDecToStr(v)
#define DecToVar(v) vbtDecToVar(v)
#define DecToDec(v) vbtDecToDec(v)

#define UDTToVar(v) vbtUDTToVar(&v)
#define UDTToVoidRef(v) ((void *)(LONG)(&v))

BSTR vbtUINT32ToStr(UINT32 u32);
#define UINT32ToStr(u32) vbtUINT32ToStr(u32)
VARIANT vbtUINT32ToVar(UINT32 u32);
#define UINT32ToVar(u32) vbtUINT32ToVar(u32)
#define UINT32ToLng(v) ToLng(v) /* fixme: create macros for all automation types? */

/* vbt Conversion Prototypes */
typedef unsigned char UINT8;
UINT8 vbtByteToByte(UINT8 v);
INT16 vbtByteToBool(UINT8 v);
INT16 vbtByteToInt(UINT8 v);
INT32 vbtByteToLng(UINT8 v);
FLOAT vbtByteToSng(UINT8 v);
DOUBLE vbtByteToDbl(UINT8 v);
DATE vbtByteToDate(UINT8 v);
CURRENCY vbtByteToCur(UINT8 v);
BSTR vbtByteToStr(UINT8 v);
VARIANT vbtByteToVar(UINT8 v);
VARIANT vbtByteRefToVar(UINT8 *v); /* fixme: need more xxxRefToxxx */
DECIMAL vbtByteToDec(UINT8 v);

UINT8 vbtBoolToByte(INT16 v);
INT16 vbtBoolToBool(INT16 v);
INT16 vbtBoolToInt(INT16 v);
INT32 vbtBoolToLng(INT16 v);
FLOAT vbtBoolToSng(INT16 v);
DOUBLE vbtBoolToDbl(INT16 v);
DATE vbtBoolToDate(INT16 v);
CURRENCY vbtBoolToCur(INT16 v);
BSTR vbtBoolToStr(INT16 v);
VARIANT vbtBoolToVar(INT16 v);
DECIMAL vbtBoolToDec(INT16 v);

UINT8 vbtIntToByte(INT16 v);
INT16 vbtIntToBool(INT16 v);
INT16 vbtIntRefToBool(INT16 *v);
INT16 vbtIntToInt(INT16 v);
INT32 vbtIntToLng(INT16 v);
FLOAT vbtIntToSng(INT16 v);
DOUBLE vbtIntToDbl(INT16 v);
DATE vbtIntToDate(INT16 v);
CURRENCY vbtIntToCur(INT16 v);
BSTR vbtIntToStr(INT16 v);
BSTR vbtIntRefToStr(INT16 *v);
BSTR vbtIntToStr(INT16 v);
VARIANT vbtIntToVar(INT16 v);
VARIANT vbtIntRefToVar(INT16 *v);
DECIMAL vbtIntToDec(INT16 v);

UINT8 vbtLngToByte(INT32 v);
INT16 vbtLngToBool(INT32 v);
INT16 vbtLngToInt(INT32 v);
INT32 vbtLngToLng(INT32 v);
FLOAT vbtLngToSng(INT32 v);
DOUBLE vbtLngToDbl(INT32 v);
DATE vbtLngToDate(INT32 v);
CURRENCY vbtLngToCur(INT32 v);
BSTR vbtLngToStr(INT32 v);
BSTR vbtLngRefToStr(INT32 *v);
VARIANT vbtLngToVar(INT32 v);
VARIANT vbtLngRefToVar(INT32 *v);
DECIMAL vbtLngToDec(INT32 v);

UINT8 vbtSngToByte(FLOAT v);
INT16 vbtSngToBool(FLOAT v);
INT16 vbtSngToInt(FLOAT v);
INT32 vbtSngToLng(FLOAT v);
FLOAT vbtSngToSng(FLOAT v);
DOUBLE vbtSngToDbl(FLOAT v);
DATE vbtSngToDate(FLOAT v);
CURRENCY vbtSngToCur(FLOAT v);
BSTR vbtSngToStr(FLOAT v);
VARIANT vbtSngToVar(FLOAT v);
DECIMAL vbtSngToDec(FLOAT v);

UINT8 vbtDblToByte(DOUBLE v);
INT16 vbtDblToBool(DOUBLE v);
INT16 vbtDblToInt(DOUBLE v);
INT32 vbtDblToLng(DOUBLE v);
FLOAT vbtDblToSng(DOUBLE v);
DOUBLE vbtDblToDbl(DOUBLE v);
DATE vbtDblToDate(DOUBLE v);
CURRENCY vbtDblToCur(DOUBLE v);
BSTR vbtDblToStr(DOUBLE v);
VARIANT vbtDblToVar(DOUBLE v);
DECIMAL vbtDblToDec(DOUBLE v);

UINT8 vbtDateToByte(DATE v);
INT16 vbtDateToBool(DATE v);
INT16 vbtDateToInt(DATE v);
INT32 vbtDateToLng(DATE v);
FLOAT vbtDateToSng(DATE v);
DOUBLE vbtDateToDbl(DATE v);
DATE vbtDateToDate(DATE v);
CURRENCY vbtDateToCur(DATE v);
BSTR vbtDateToStr(DATE v);
VARIANT vbtDateToVar(DATE v);
VARIANT vbtDateRefToVar(DATE *v);
DECIMAL vbtDateToDec(DATE v);

UINT8 vbtCurToByte(CURRENCY v);
INT16 vbtCurToBool(CURRENCY v);
INT16 vbtCurToInt(CURRENCY v);
INT32 vbtCurToLng(CURRENCY v);
FLOAT vbtCurToSng(CURRENCY v);
DOUBLE vbtCurToDbl(CURRENCY v);
DATE vbtCurToDate(CURRENCY v);
CURRENCY vbtCurToCur(CURRENCY v);
BSTR vbtCurToStr(CURRENCY v);
VARIANT vbtCurToVar(CURRENCY v);
DECIMAL vbtCurToDec(CURRENCY v);

UINT8 vbtStrToByte(BSTR v);
INT16 vbtStrToBool(BSTR v);
INT16 vbtStrToInt(BSTR v);
INT32 vbtStrToLng(BSTR v);
FLOAT vbtStrToSng(BSTR v);
DOUBLE vbtStrToDbl(BSTR v);
DATE vbtStrToDate(BSTR v);
CURRENCY vbtStrToCur(BSTR v);
BSTR vbtStrToStr(BSTR v);
VARIANT vbtStrToVar(BSTR v);
DECIMAL vbtStrToDec(BSTR v);

UINT8 vbtVarToByte(VARIANT v);
INT16 vbtVarToBool(VARIANT v);
INT16 vbtVarToInt(VARIANT v);
INT32 vbtVarToLng(VARIANT v);
FLOAT vbtVarToSng(VARIANT v);
DOUBLE vbtVarToDbl(VARIANT v);
DATE vbtVarToDate(VARIANT v);
CURRENCY vbtVarToCur(VARIANT v);
BSTR vbtVarToStr(VARIANT v);
VARIANT vbtVarToVar(VARIANT v);
DECIMAL vbtVarToDec(VARIANT v);
UINT32 vbtVarToUINT32(VARIANT v);

UINT8 vbtDecToByte(DECIMAL v);
INT16 vbtDecToBool(DECIMAL v);
INT16 vbtDecToInt(DECIMAL v);
INT32 vbtDecToLng(DECIMAL v);
FLOAT vbtDecToSng(DECIMAL v);
DOUBLE vbtDecToDbl(DECIMAL v);
DATE vbtDecToDate(DECIMAL v);
CURRENCY vbtDecToCur(DECIMAL v);
BSTR vbtDecToStr(DECIMAL v);
VARIANT vbtDecToVar(DECIMAL v);
DECIMAL vbtDecToDec(DECIMAL v);

VARIANT vbtUDTToVar(void *v);

#ifdef NEVER
wchar_t wbuf[64];
#define SngToStr(v) SysAllocString(_itow((int)(v),wbuf,10))
/*#define StrToVar(v) (VARIANT)(v)*/
VARIANT *StrToVar(LPOLESTR str);
#endif

/* Operator defs */
#define ByteImp(b1,b2) (UINT8)(~(b1)|(b2))
#define ByteEqv(b1,b2) (UINT8)(~((b1)^(b2)))
#define ByteXor(b1,b2) (UINT8)((b1)^(b2))
#define ByteOr(b1,b2) (UINT8)((b1)|(b2))
#define ByteAnd(b1,b2) (UINT8)((b1)&(b2))
#define ByteAdd(b1,b2) (UINT8)((b1)+(b2))
#define ByteSub(b1,b2) (UINT8)((b1)-(b2)) /* MSVC6 needed a post-op cast to avoid "integral size mismatch" warning. Casted all other Byte operations just in case. */
#define ByteMul(b1,b2) (UINT8)((b1)*(b2))
#define ByteDiv(b1,b2) (UINT8)((b1)/(b2))
#define ByteIDiv(b1,b2) (UINT8)((b1)/(b2))
#define ByteLT(b1,b2) (VARIANT_BOOL)((b1)<(b2))
#define ByteLE(b1,b2) (VARIANT_BOOL)((b1)<=(b2))
#define ByteEQ(b1,b2) (VARIANT_BOOL)((b1)==(b2))
#define ByteNE(b1,b2) (VARIANT_BOOL)((b1)!=(b2))
#define ByteGT(b1,b2) (VARIANT_BOOL)((b1)>(b2))
#define ByteGE(b1,b2) (VARIANT_BOOL)((b1)>=(b2))
#define ByteAbs(b1) (UINT8)(b1)
#define ByteFix(b1) (UINT8)(b1)
#define ByteInt(b1) (UINT8)(b1)
#define ByteMod(b1,b2) (UINT8)((b1)%(b2))
#define BytePow(b1,b2) diag(ERR_NOT_IMPLEMENTED)
#define ByteNot(b1) (UINT8)(~(b1))
#define ByteNeg(b1) (UINT8)(-(b1))

/* haven't found any other Bool operations */
#define BoolNot(i1) (VARIANT_BOOL)~(i1) /* MSVC 6.0 warns int-to-short conversion, unless cast is used. Perhaps should use 32-bit instead of 16-bit for all param declares */

#define IntImp(i1,i2) (INT16)(~(i1)|(i2))
#define IntEqv(i1,i2) (INT16)(~((i1)^(i2)))
#define IntXor(i1,i2) (INT16)((i1)^(i2))
#define IntOr(i1,i2) (INT16)((i1)|(i2))
#define IntAnd(i1,i2) (INT16)((i1)&(i2))
#define IntAdd(i1,i2) (INT16)((i1)+(i2))
#define IntSub(i1,i2) (INT16)((i1)-(i2))
#define IntMul(i1,i2) (INT16)((i1)*(i2))
#define IntDiv(i1,i2) (INT16)((i1)/(i2))
#define IntIDiv(i1,i2) (INT16)((i1)/(i2))
#define IntLT(i1,i2) (VARIANT_BOOL)((i1)<(i2))
#define IntLE(i1,i2) (VARIANT_BOOL)((i1)<=(i2))
#define IntEQ(i1,i2) (VARIANT_BOOL)((i1)==(i2))
#define IntNE(i1,i2) (VARIANT_BOOL)((i1)!=(i2))
#define IntGT(i1,i2) (VARIANT_BOOL)((i1)>(i2))
#define IntGE(i1,i2) (VARIANT_BOOL)((i1)>=(i2))
#define IntAbs(i1) (INT16)((i1) < 0 ? -(i1) : (i1))
#define IntFix(i1)  (INT16)(i1)
#define IntInt(i1) (INT16)(i1)
#define IntMod(i1,i2) (INT16)((i1)%(i2))
#define IntPow(i1,i2) diag(ERR_NOT_IMPLEMENTED)
#define IntNot(i1) (INT16)(~(i1))
#define IntNeg(i1) (INT16)(-(i1))

#define LngImp(l1,l2) ~(l1)|(l2)
#define LngEqv(l1,l2) ~((l1)^(l2))
#define LngXor(l1,l2) (l1)^(l2)
#define LngOr(l1,l2) (l1)|(l2) /* why does MSVC give signed/unsigned mismatch when 2nd parameter is 0x80000000? */
#define LngAnd(l1,l2) (l1)&(l2)
#define LngAdd(l1,l2) (l1)+(l2)
#define LngSub(l1,l2) (l1)-(l2)
#define LngMul(l1,l2) (l1)*(l2)
#define LngDiv(l1,l2) (l1)/(l2)
#define LngIDiv(l1,l2) (l1)/(l2)
#define LngLT(l1,l2) (VARIANT_BOOL)((l1)<(l2))
#define LngLE(l1,l2) (VARIANT_BOOL)((l1)<=(l2))
#define LngEQ(l1,l2) (VARIANT_BOOL)((l1)==(l2))
#define LngNE(l1,l2) (VARIANT_BOOL)((l1)!=(l2))
#define LngGT(l1,l2) (VARIANT_BOOL)((l1)>(l2))
#define LngGE(l1,l2) (VARIANT_BOOL)((l1)>=(l2))
#define LngAbs(l1) (l1) < 0 ? -(l1) : (l1)
#define LngFix(l1)  (l1)
#define LngInt(l1) (l1)
#define LngMod(l1,l2) (l1)%(l2)
#define LngPow(l1,l2) diag(ERR_NOT_IMPLEMENTED)
#define LngNot(l1) ~(l1)
#define LngNeg(l1) -(l1)

#define SngImp(f1,f2) (FLOAT)(~(UINT32)(f1)|(UINT32)(f2))
#define SngEqv(f1,f2) (FLOAT)~((UINT32)(f1)^(UINT32)(f2))
#define SngXor(f1,f2) (FLOAT)((UINT32)(f1)^(UINT32)(f2))
#define SngOr(f1,f2) (FLOAT)((UINT32)(f1)|(UINT32)(f2))
#define SngAnd(f1,f2) (FLOAT)((UINT32)(f1)&(UINT32)(f2))
#define SngAdd(f1,f2) (f1)+(f2)
#define SngSub(f1,f2) (f1)-(f2)
#define SngMul(f1,f2) (f1)*(f2)
#define SngDiv(f1,f2) (f1)/(f2)
#define SngLT(f1,f2) (VARIANT_BOOL)((f1)<(f2))
#define SngLE(f1,f2) (VARIANT_BOOL)((f1)<=(f2))
#define SngEQ(f1,f2) (VARIANT_BOOL)((f1)==(f2))
#define SngNE(f1,f2) (VARIANT_BOOL)((f1)!=(f2))
#define SngGT(f1,f2) (VARIANT_BOOL)((f1)>(f2))
#define SngGE(f1,f2) (VARIANT_BOOL)((f1)>=(f2))
#define SngAbs(f1) (f1) < 0 ? -(f1) : (f1)
#define SngFix(f1) (FLOAT)floor((DOUBLE)(f1))
#define SngInt(f1) (FLOAT)floor((DOUBLE)(f1))
#define SngMod(f1,f2) (FLOAT)((UINT32)(f1)%(UINT32)(f2))
#define SngPow(f1,f2) diag(ERR_NOT_IMPLEMENTED)
#define SngNot(f1) (FLOAT)~(UINT32)(f1)
#define SngNeg(f1) -(f1)

#define DblImp(d1,d2) (DOUBLE)(~(UINT32)(d1)|(UINT32)(d2))
#define DblEqv(d1,d2) (DOUBLE)~((UINT32)(d1)^(UINT32)(d2))
#define DblXor(d1,d2) (DOUBLE)((UINT32)(d1)^(UINT32)(d2))
#define DblOr(d1,d2) (DOUBLE)((UINT32)(d1)|(UINT32)(d2))
#define DblAnd(d1,d2) (DOUBLE)((UINT32)(d1)&(UINT32)(d2))
#define DblAdd(d1,d2) (d1)+(d2)
#define DblSub(d1,d2) (d1)-(d2)
#define DblMul(d1,d2) (d1)*(d2)
#define DblDiv(d1,d2) (d1)/(d2)
#define DblLT(d1,d2) (VARIANT_BOOL)((d1)<(d2))
#define DblLE(d1,d2) (VARIANT_BOOL)((d1)<=(d2))
#define DblEQ(d1,d2) (VARIANT_BOOL)((d1)==(d2))
#define DblNE(d1,d2) (VARIANT_BOOL)((d1)!=(d2))
#define DblGT(d1,d2) (VARIANT_BOOL)((d1)>(d2))
#define DblGE(d1,d2) (VARIANT_BOOL)((d1)>=(d2))
#define DblAbs(d1) (d1) < 0 ? -(d1) : (d1)
#define DblFix(d1) (DOUBLE)floor((DOUBLE)(d1))
#define DblInt(d1) (DOUBLE)floor((DOUBLE)(d1))
#define DblMod(d1,d2) (DOUBLE)((UINT32)(d1)%(UINT32)(d2))
#define DblPow(d1,d2) diag(ERR_NOT_IMPLEMENTED)
#define DblNot(d1) (DOUBLE)~(UINT32)(d1)
#define DblNeg(d1) -(d1)

#if 0

#define ByteImp(b1,b2) vbtByteImp(b1,b2)
#define ByteEqv(b1,b2) vbtByteEqv(b1,b2)
#define ByteXor(b1,b2) vbtByteXor(b1,b2)
#define ByteOr(b1,b2) vbtByteOr(b1,b2)
#define ByteAnd(b1,b2) vbtByteAnd(b1,b2)
#define ByteAdd(b1,b2) vbtByteAdd(b1,b2)
#define ByteSub(b1,b2) vbtByteSub(b1,b2)
#define ByteMul(b1,b2) vbtByteMul(b1,b2)
#define ByteDiv(b1,b2) vbtByteDiv(b1,b2)
#define ByteIDiv(b1,b2) vbtByteIDiv(b1,b2)
#define ByteLT(b1,b2) vbtByteLT(b1,b2)
#define ByteLE(b1,b2) vbtByteLE(b1,b2)
#define ByteEQ(b1,b2) vbtByteEQ(b1,b2)
#define ByteNE(b1,b2) vbtByteNE(b1,b2)
#define ByteGT(b1,b2) vbtByteGT(b1,b2)
#define ByteGE(b1,b2) vbtByteGE(b1,b2)
#define ByteAbs(b1) vbtByteAbs(b1)
#define ByteFix(b1) vbtByteFix(b1)
#define ByteInt(b1) vbtByteInt(b1)
#define ByteMod(b1,b2) vbtByteMod(b1,b2)
#define BytePow(b1,b2) vbtBytePow(b1,b2)
#define ByteNot(b1) vbtByteNot(b1)
#define ByteNeg(b1) vbtByteNeg(b1)

#define BoolImp(b1,b2) vbtBoolImp(b1,b2)
#define BoolEqv(b1,b2) vbtBoolEqv(b1,b2)
#define BoolXor(b1,b2) vbtBoolXor(b1,b2)
#define BoolOr(b1,b2) vbtBoolOr(b1,b2)
#define BoolAnd(b1,b2) vbtBoolAnd(b1,b2)
#define BoolAdd(b1,b2) vbtBoolAdd(b1,b2)
#define BoolSub(b1,b2) vbtBoolSub(b1,b2)
#define BoolMul(b1,b2) vbtBoolMul(b1,b2)
#define BoolDiv(b1,b2) vbtBoolDiv(b1,b2)
#define BoolIDiv(b1,b2) vbtBoolIDiv(b1,b2)
#define BoolLT(b1,b2) vbtBoolLT(b1,b2)
#define BoolLE(b1,b2) vbtBoolLE(b1,b2)
#define BoolEQ(b1,b2) vbtBoolEQ(b1,b2)
#define BoolNE(b1,b2) vbtBoolNE(b1,b2)
#define BoolGT(b1,b2) vbtBoolGT(b1,b2)
#define BoolGE(b1,b2) vbtBoolGE(b1,b2)
#define BoolAbs(b1) vbtBoolAbs(b1)
#define BoolFix(b1) vbtBoolFix(b1)
#define BoolInt(b1) vbtBoolInt(b1)
#define BoolMod(b1,b2) vbtBoolMod(b1,b2)
#define BoolPow(b1,b2) vbtBoolPow(b1,b2)
#define BoolNot(b1) vbtBoolNot(b1)
#define BoolNeg(b1) vbtBoolNeg(b1)

#define IntImp(b1,b2) vbtIntImp(b1,b2)
#define IntEqv(b1,b2) vbtIntEqv(b1,b2)
#define IntXor(b1,b2) vbtIntXor(b1,b2)
#define IntOr(b1,b2) vbtIntOr(b1,b2)
#define IntAnd(b1,b2) vbtIntAnd(b1,b2)
#define IntAdd(b1,b2) vbtIntAdd(b1,b2)
#define IntSub(b1,b2) vbtIntSub(b1,b2)
#define IntMul(b1,b2) vbtIntMul(b1,b2)
#define IntDiv(b1,b2) vbtIntDiv(b1,b2)
#define IntIDiv(b1,b2) vbtIntIDiv(b1,b2)
#define IntLT(b1,b2) vbtIntLT(b1,b2)
#define IntLE(b1,b2) vbtIntLE(b1,b2)
#define IntEQ(b1,b2) vbtIntEQ(b1,b2)
#define IntNE(b1,b2) vbtIntNE(b1,b2)
#define IntGT(b1,b2) vbtIntGT(b1,b2)
#define IntGE(b1,b2) vbtIntGE(b1,b2)
#define IntAbs(b1) vbtIntAbs(b1)
#define IntFix(b1) vbtIntFix(b1)
#define IntInt(b1) vbtIntInt(b1)
#define IntMod(b1,b2) vbtIntMod(b1,b2)
#define IntPow(b1,b2) vbtIntPow(b1,b2)
#define IntNot(b1) vbtIntNot(b1)
#define IntNeg(b1) vbtIntNeg(b1)

#define LngImp(b1,b2) vbtLngImp(b1,b2)
#define LngEqv(b1,b2) vbtLngEqv(b1,b2)
#define LngXor(b1,b2) vbtLngXor(b1,b2)
#define LngOr(b1,b2) vbtLngOr(b1,b2)
#define LngAnd(b1,b2) vbtLngAnd(b1,b2)
#define LngAdd(b1,b2) vbtLngAdd(b1,b2)
#define LngSub(b1,b2) vbtLngSub(b1,b2)
#define LngMul(b1,b2) vbtLngMul(b1,b2)
#define LngDiv(b1,b2) vbtLngDiv(b1,b2)
#define LngIDiv(b1,b2) vbtLngIDiv(b1,b2)
#define LngLT(b1,b2) vbtLngLT(b1,b2)
#define LngLE(b1,b2) vbtLngLE(b1,b2)
#define LngEQ(b1,b2) vbtLngEQ(b1,b2)
#define LngNE(b1,b2) vbtLngNE(b1,b2)
#define LngGT(b1,b2) vbtLngGT(b1,b2)
#define LngGE(b1,b2) vbtLngGE(b1,b2)
#define LngAbs(b1) vbtLngAbs(b1)
#define LngFix(b1) vbtLngFix(b1)
#define LngInt(b1) vbtLngInt(b1)
#define LngMod(b1,b2) vbtLngMod(b1,b2)
#define LngPow(b1,b2) vbtLngPow(b1,b2)
#define LngNot(b1) vbtLngNot(b1)
#define LngNeg(b1) vbtLngNeg(b1)

#define SngImp(b1,b2) vbtSngImp(b1,b2)
#define SngEqv(b1,b2) vbtSngEqv(b1,b2)
#define SngXor(b1,b2) vbtSngXor(b1,b2)
#define SngOr(b1,b2) vbtSngOr(b1,b2)
#define SngAnd(b1,b2) vbtSngAnd(b1,b2)
#define SngAdd(b1,b2) vbtSngAdd(b1,b2)
#define SngSub(b1,b2) vbtSngSub(b1,b2)
#define SngMul(b1,b2) vbtSngMul(b1,b2)
#define SngDiv(b1,b2) vbtSngDiv(b1,b2)
#define SngIDiv(b1,b2) vbtSngIDiv(b1,b2)
#define SngLT(b1,b2) vbtSngLT(b1,b2)
#define SngLE(b1,b2) vbtSngLE(b1,b2)
#define SngEQ(b1,b2) vbtSngEQ(b1,b2)
#define SngNE(b1,b2) vbtSngNE(b1,b2)
#define SngGT(b1,b2) vbtSngGT(b1,b2)
#define SngGE(b1,b2) vbtSngGE(b1,b2)
#define SngAbs(b1) vbtSngAbs(b1)
#define SngFix(b1) vbtSngFix(b1)
#define SngInt(b1) vbtSngInt(b1)
#define SngMod(b1,b2) vbtSngMod(b1,b2)
#define SngPow(b1,b2) vbtSngPow(b1,b2)
#define SngNot(b1) vbtSngNot(b1)
#define SngNeg(b1) vbtSngNeg(b1)

#define DblImp(b1,b2) vbtDblImp(b1,b2)
#define DblEqv(b1,b2) vbtDblEqv(b1,b2)
#define DblXor(b1,b2) vbtDblXor(b1,b2)
#define DblOr(b1,b2) vbtDblOr(b1,b2)
#define DblAnd(b1,b2) vbtDblAnd(b1,b2)
#define DblAdd(b1,b2) vbtDblAdd(b1,b2)
#define DblSub(b1,b2) vbtDblSub(b1,b2)
#define DblMul(b1,b2) vbtDblMul(b1,b2)
#define DblDiv(b1,b2) vbtDblDiv(b1,b2)
#define DblIDiv(b1,b2) vbtDblIDiv(b1,b2)
#define DblLT(b1,b2) vbtDblLT(b1,b2)
#define DblLE(b1,b2) vbtDblLE(b1,b2)
#define DblEQ(b1,b2) vbtDblEQ(b1,b2)
#define DblNE(b1,b2) vbtDblNE(b1,b2)
#define DblGT(b1,b2) vbtDblGT(b1,b2)
#define DblGE(b1,b2) vbtDblGE(b1,b2)
#define DblAbs(b1) vbtDblAbs(b1)
#define DblFix(b1) vbtDblFix(b1)
#define DblInt(b1) vbtDblInt(b1)
#define DblMod(b1,b2) vbtDblMod(b1,b2)
#define DblPow(b1,b2) vbtDblPow(b1,b2)
#define DblNot(b1) vbtDblNot(b1)
#define DblNeg(b1) vbtDblNeg(b1)
#endif

#define DateImp(b1,b2) vbtDateImp(b1,b2)
#define DateEqv(b1,b2) vbtDateEqv(b1,b2)
#define DateXor(b1,b2) vbtDateXor(b1,b2)
#define DateOr(b1,b2) vbtDateOr(b1,b2)
#define DateAnd(b1,b2) vbtDateAnd(b1,b2)
#define DateAdd(b1,b2) vbtDateAdd(b1,b2)
#define DateSub(b1,b2) vbtDateSub(b1,b2)
#define DateMul(b1,b2) vbtDateMul(b1,b2)
#define DateDiv(b1,b2) vbtDateDiv(b1,b2)
#define DateIDiv(b1,b2) vbtDateIDiv(b1,b2)
#define DateLT(b1,b2) vbtDateLT(b1,b2)
#define DateLE(b1,b2) vbtDateLE(b1,b2)
#define DateEQ(b1,b2) vbtDateEQ(b1,b2)
#define DateNE(b1,b2) vbtDateNE(b1,b2)
#define DateGT(b1,b2) vbtDateGT(b1,b2)
#define DateGE(b1,b2) vbtDateGE(b1,b2)
#define DateAbs(b1) vbtDateAbs(b1)
#define DateFix(b1) vbtDateFix(b1)
#define DateInt(b1) vbtDateInt(b1)
#define DateMod(b1,b2) vbtDateMod(b1,b2)
#define DatePow(b1,b2) vbtDatePow(b1,b2)
#define DateNot(b1) vbtDateNot(b1)
#define DateNeg(b1) vbtDateNeg(b1)

#define CurImp(b1,b2) vbtCurImp(b1,b2)
#define CurEqv(b1,b2) vbtCurEqv(b1,b2)
#define CurXor(b1,b2) vbtCurXor(b1,b2)
#define CurOr(b1,b2) vbtCurOr(b1,b2)
#define CurAnd(b1,b2) vbtCurAnd(b1,b2)
#define CurAdd(b1,b2) vbtCurAdd(b1,b2)
#define CurSub(b1,b2) vbtCurSub(b1,b2)
#define CurMul(b1,b2) vbtCurMul(b1,b2)
#define CurDiv(b1,b2) vbtCurDiv(b1,b2)
#define CurIDiv(b1,b2) vbtCurIDiv(b1,b2)
#define CurLT(b1,b2) vbtCurLT(b1,b2)
#define CurLE(b1,b2) vbtCurLE(b1,b2)
#define CurEQ(b1,b2) vbtCurEQ(b1,b2)
#define CurNE(b1,b2) vbtCurNE(b1,b2)
#define CurGT(b1,b2) vbtCurGT(b1,b2)
#define CurGE(b1,b2) vbtCurGE(b1,b2)
#define CurAbs(b1) vbtCurAbs(b1)
#define CurFix(b1) vbtCurFix(b1)
#define CurInt(b1) vbtCurInt(b1)
#define CurMod(b1,b2) vbtCurMod(b1,b2)
#define CurPow(b1,b2) vbtCurPow(b1,b2)
#define CurNot(b1) vbtCurNot(b1)
#define CurNeg(b1) vbtCurNeg(b1)

/*#define StrCat(s1,s2) wcscat(SysAllocStringLen(s1,SysStringLen(s1)+SysStringLen(s2)),s2)*/
#define StrCat(s1,s2) vbtStrCat(s1,s2)
#define StrAdd(s1,s2) StrCat(s1,s2)
#define StrLT(b1,b2) vbtStrLT(b1,b2)
#define StrLE(b1,b2) vbtStrLE(b1,b2)
#define StrEQ(b1,b2) vbtStrEQ(b1,b2)
#define StrNE(b1,b2) vbtStrNE(b1,b2)
#define StrGT(b1,b2) vbtStrGT(b1,b2)
#define StrGE(b1,b2) vbtStrGE(b1,b2)

#define VarCat(b1,b2) vbtVarCat(b1,b2)
#define VarImp(b1,b2) vbtVarImp(b1,b2)
#define VarEqv(b1,b2) vbtVarEqv(b1,b2)
#define VarXor(b1,b2) vbtVarXor(b1,b2)
#define VarOr(b1,b2) vbtVarOr(b1,b2)
#define VarAnd(b1,b2) vbtVarAnd(b1,b2)
#define VarAdd(b1,b2) vbtVarAdd(b1,b2)
#define VarSub(b1,b2) vbtVarSub(b1,b2)
#define VarMul(b1,b2) vbtVarMul(b1,b2)
#define VarDiv(b1,b2) vbtVarDiv(b1,b2)
#define VarIDiv(b1,b2) vbtVarIDiv(b1,b2)
#define VarLT(b1,b2) vbtVarLT(b1,b2)
#define VarLE(b1,b2) vbtVarLE(b1,b2)
#define VarEQ(b1,b2) vbtVarEQ(b1,b2)
#define VarNE(b1,b2) vbtVarNE(b1,b2)
#define VarGT(b1,b2) vbtVarGT(b1,b2)
#define VarGE(b1,b2) vbtVarGE(b1,b2)
#define VarAbs(b1) vbtVarAbs(b1)
#define VarFix(b1) vbtVarFix(b1)
#define VarInt(b1) vbtVarInt(b1)
#define VarMod(b1,b2) vbtVarMod(b1,b2)
#define VarPow(b1,b2) vbtVarPow(b1,b2)
#define VarNot(b1) vbtVarNot(b1)
#define VarNeg(b1) vbtVarNeg(b1)

UINT8 vbtByteImp(UINT8 left, UINT8 right);
UINT8 vbtByteEqv(UINT8 left, UINT8 right);
UINT8 vbtByteXor(UINT8 left, UINT8 right);
UINT8 vbtByteOr(UINT8 left, UINT8 right);
UINT8 vbtByteAnd(UINT8 left, UINT8 right);
UINT8 vbtByteAdd(UINT8 left, UINT8 right);
UINT8 vbtByteSub(UINT8 left, UINT8 right);
UINT8 vbtByteMul(UINT8 left, UINT8 right);
UINT8 vbtByteDiv(UINT8 left, UINT8 right);
UINT8 vbtByteIDiv(UINT8 left, UINT8 right);
INT16 vbtByteLT(UINT8 left, UINT8 right);
INT16 vbtByteLE(UINT8 left, UINT8 right);
INT16 vbtByteEQ(UINT8 left, UINT8 right);
INT16 vbtByteNE(UINT8 left, UINT8 right);
INT16 vbtByteGT(UINT8 left, UINT8 right);
INT16 vbtByteGE(UINT8 left, UINT8 right);
UINT8 vbtByteAbs(UINT8 arg);
UINT8 vbtByteFix(UINT8 arg);
UINT8 vbtByteInt(UINT8 arg);
UINT8 vbtByteMod(UINT8 left, UINT8 right);
UINT8 vbtBytePow(UINT8 left, UINT8 right);
UINT8 vbtByteNot(UINT8 arg);
UINT8 vbtByteNeg(UINT8 arg);

INT16 vbtBoolImp(INT16 left, INT16 right);
INT16 vbtBoolEqv(INT16 left, INT16 right);
INT16 vbtBoolXor(INT16 left, INT16 right);
INT16 vbtBoolOr(INT16 left, INT16 right);
INT16 vbtBoolAnd(INT16 left, INT16 right);
INT16 vbtBoolAdd(INT16 left, INT16 right);
INT16 vbtBoolSub(INT16 left, INT16 right);
INT16 vbtBoolMul(INT16 left, INT16 right);
INT16 vbtBoolDiv(INT16 left, INT16 right);
INT16 vbtBoolIDiv(INT16 left, INT16 right);
INT16 vbtBoolLT(INT16 left, INT16 right);
INT16 vbtBoolLE(INT16 left, INT16 right);
INT16 vbtBoolEQ(INT16 left, INT16 right);
INT16 vbtBoolNE(INT16 left, INT16 right);
INT16 vbtBoolGT(INT16 left, INT16 right);
INT16 vbtBoolGE(INT16 left, INT16 right);
INT16 vbtBoolAbs(INT16 arg);
INT16 vbtBoolFix(INT16 arg);
INT16 vbtBoolInt(INT16 arg);
INT16 vbtBoolMod(INT16 left, INT16 right);
INT16 vbtBoolPow(INT16 left, INT16 right);
INT16 vbtBoolNot(INT16 arg);
INT16 vbtBoolNeg(INT16 arg);

INT16 vbtIntImp(INT16 left, INT16 right);
INT16 vbtIntEqv(INT16 left, INT16 right);
INT16 vbtIntXor(INT16 left, INT16 right);
INT16 vbtIntOr(INT16 left, INT16 right);
INT16 vbtIntAnd(INT16 left, INT16 right);
INT16 vbtIntAdd(INT16 left, INT16 right);
INT16 vbtIntSub(INT16 left, INT16 right);
INT16 vbtIntMul(INT16 left, INT16 right);
INT16 vbtIntDiv(INT16 left, INT16 right);
INT16 vbtIntIDiv(INT16 left, INT16 right);
INT16 vbtIntLT(INT16 left, INT16 right);
INT16 vbtIntLE(INT16 left, INT16 right);
INT16 vbtIntEQ(INT16 left, INT16 right);
INT16 vbtIntNE(INT16 left, INT16 right);
INT16 vbtIntGT(INT16 left, INT16 right);
INT16 vbtIntGE(INT16 left, INT16 right);
INT16 vbtIntAbs(INT16 arg);
INT16 vbtIntFix(INT16 arg);
INT16 vbtIntInt(INT16 arg);
INT16 vbtIntMod(INT16 left, INT16 right);
INT16 vbtIntPow(INT16 left, INT16 right);
INT16 vbtIntNot(INT16 arg);
INT16 vbtIntNeg(INT16 arg);

INT32 vbtLngImp(INT32 left, INT32 right);
INT32 vbtLngEqv(INT32 left, INT32 right);
INT32 vbtLngXor(INT32 left, INT32 right);
INT32 vbtLngOr(INT32 left, INT32 right);
INT32 vbtLngAnd(INT32 left, INT32 right);
INT32 vbtLngAdd(INT32 left, INT32 right);
INT32 vbtLngSub(INT32 left, INT32 right);
INT32 vbtLngMul(INT32 left, INT32 right);
INT32 vbtLngDiv(INT32 left, INT32 right);
INT32 vbtLngIDiv(INT32 left, INT32 right);
INT16 vbtLngLT(INT32 left, INT32 right);
INT16 vbtLngLE(INT32 left, INT32 right);
INT16 vbtLngEQ(INT32 left, INT32 right);
INT16 vbtLngNE(INT32 left, INT32 right);
INT16 vbtLngGT(INT32 left, INT32 right);
INT16 vbtLngGE(INT32 left, INT32 right);
INT32 vbtLngAbs(INT32 arg);
INT32 vbtLngFix(INT32 arg);
INT32 vbtLngInt(INT32 arg);
INT32 vbtLngMod(INT32 left, INT32 right);
INT32 vbtLngPow(INT32 left, INT32 right);
INT32 vbtLngNot(INT32 arg);
INT32 vbtLngNeg(INT32 arg);

FLOAT vbtSngImp(FLOAT left, FLOAT right);
FLOAT vbtSngEqv(FLOAT left, FLOAT right);
FLOAT vbtSngXor(FLOAT left, FLOAT right);
FLOAT vbtSngOr(FLOAT left, FLOAT right);
FLOAT vbtSngAnd(FLOAT left, FLOAT right);
FLOAT vbtSngAdd(FLOAT left, FLOAT right);
FLOAT vbtSngSub(FLOAT left, FLOAT right);
FLOAT vbtSngMul(FLOAT left, FLOAT right);
FLOAT vbtSngDiv(FLOAT left, FLOAT right);
FLOAT vbtSngIDiv(FLOAT left, FLOAT right);
INT16 vbtSngLT(FLOAT left, FLOAT right);
INT16 vbtSngLE(FLOAT left, FLOAT right);
INT16 vbtSngEQ(FLOAT left, FLOAT right);
INT16 vbtSngNE(FLOAT left, FLOAT right);
INT16 vbtSngGT(FLOAT left, FLOAT right);
INT16 vbtSngGE(FLOAT left, FLOAT right);
FLOAT vbtSngAbs(FLOAT arg);
FLOAT vbtSngFix(FLOAT arg);
FLOAT vbtSngInt(FLOAT arg);
FLOAT vbtSngMod(FLOAT left, FLOAT right);
FLOAT vbtSngPow(FLOAT left, FLOAT right);
FLOAT vbtSngNot(FLOAT arg);
FLOAT vbtSngNeg(FLOAT arg);

DOUBLE vbtDblImp(DOUBLE left, DOUBLE right);
DOUBLE vbtDblEqv(DOUBLE left, DOUBLE right);
DOUBLE vbtDblXor(DOUBLE left, DOUBLE right);
DOUBLE vbtDblOr(DOUBLE left, DOUBLE right);
DOUBLE vbtDblAnd(DOUBLE left, DOUBLE right);
DOUBLE vbtDblAdd(DOUBLE left, DOUBLE right);
DOUBLE vbtDblSub(DOUBLE left, DOUBLE right);
DOUBLE vbtDblMul(DOUBLE left, DOUBLE right);
DOUBLE vbtDblDiv(DOUBLE left, DOUBLE right);
DOUBLE vbtDblIDiv(DOUBLE left, DOUBLE right);
INT16 vbtDblLT(DOUBLE left, DOUBLE right);
INT16 vbtDblLE(DOUBLE left, DOUBLE right);
INT16 vbtDblEQ(DOUBLE left, DOUBLE right);
INT16 vbtDblNE(DOUBLE left, DOUBLE right);
INT16 vbtDblGT(DOUBLE left, DOUBLE right);
INT16 vbtDblGE(DOUBLE left, DOUBLE right);
DOUBLE vbtDblAbs(DOUBLE arg);
DOUBLE vbtDblFix(DOUBLE arg);
DOUBLE vbtDblInt(DOUBLE arg);
DOUBLE vbtDblMod(DOUBLE left, DOUBLE right);
DOUBLE vbtDblPow(DOUBLE left, DOUBLE right);
DOUBLE vbtDblNot(DOUBLE arg);
DOUBLE vbtDblNeg(DOUBLE arg);

DATE vbtDateImp(DATE left, DATE right);
DATE vbtDateEqv(DATE left, DATE right);
DATE vbtDateXor(DATE left, DATE right);
DATE vbtDateOr(DATE left, DATE right);
DATE vbtDateAnd(DATE left, DATE right);
DATE vbtDateAdd(DATE left, DATE right);
DATE vbtDateSub(DATE left, DATE right);
DATE vbtDateMul(DATE left, DATE right);
DATE vbtDateDiv(DATE left, DATE right);
DATE vbtDateIDiv(DATE left, DATE right);
INT16 vbtDateLT(DATE left, DATE right);
INT16 vbtDateLE(DATE left, DATE right);
INT16 vbtDateEQ(DATE left, DATE right);
INT16 vbtDateNE(DATE left, DATE right);
INT16 vbtDateGT(DATE left, DATE right);
INT16 vbtDateGE(DATE left, DATE right);
DATE vbtDateAbs(DATE arg);
DATE vbtDateFix(DATE arg);
DATE vbtDateInt(DATE arg);
DATE vbtDateMod(DATE left, DATE right);
DATE vbtDatePow(DATE left, DATE right);
DATE vbtDateNot(DATE arg);
DATE vbtDateNeg(DATE arg);

CURRENCY vbtCurImp(CURRENCY left, CURRENCY right);
CURRENCY vbtCurEqv(CURRENCY left, CURRENCY right);
CURRENCY vbtCurXor(CURRENCY left, CURRENCY right);
CURRENCY vbtCurOr(CURRENCY left, CURRENCY right);
CURRENCY vbtCurAnd(CURRENCY left, CURRENCY right);
CURRENCY vbtCurAdd(CURRENCY left, CURRENCY right);
CURRENCY vbtCurSub(CURRENCY left, CURRENCY right);
CURRENCY vbtCurMul(CURRENCY left, CURRENCY right);
CURRENCY vbtCurDiv(CURRENCY left, CURRENCY right);
CURRENCY vbtCurIDiv(CURRENCY left, CURRENCY right);
INT16 vbtCurLT(CURRENCY left, CURRENCY right);
INT16 vbtCurLE(CURRENCY left, CURRENCY right);
INT16 vbtCurEQ(CURRENCY left, CURRENCY right);
INT16 vbtCurNE(CURRENCY left, CURRENCY right);
INT16 vbtCurGT(CURRENCY left, CURRENCY right);
INT16 vbtCurGE(CURRENCY left, CURRENCY right);
CURRENCY vbtCurAbs(CURRENCY arg);
CURRENCY vbtCurFix(CURRENCY arg);
CURRENCY vbtCurInt(CURRENCY arg);
CURRENCY vbtCurMod(CURRENCY left, CURRENCY right);
CURRENCY vbtCurPow(CURRENCY left, CURRENCY right);
CURRENCY vbtCurNot(CURRENCY arg);
CURRENCY vbtCurNeg(CURRENCY arg);

BSTR vbtStrCat(BSTR left, BSTR right);
INT16 vbtStrLT(BSTR left, BSTR right);
INT16 vbtStrLE(BSTR left, BSTR right);
INT16 vbtStrEQ(BSTR left, BSTR right);
INT16 vbtStrNE(BSTR left, BSTR right);
INT16 vbtStrGT(BSTR left, BSTR right);
INT16 vbtStrGE(BSTR left, BSTR right);

VARIANT vbtVarCat(VARIANT left, VARIANT right);
VARIANT vbtVarImp(VARIANT left, VARIANT right);
VARIANT vbtVarEqv(VARIANT left, VARIANT right);
VARIANT vbtVarXor(VARIANT left, VARIANT right);
VARIANT vbtVarOr(VARIANT left, VARIANT right);
VARIANT vbtVarAnd(VARIANT left, VARIANT right);
VARIANT vbtVarAdd(VARIANT left, VARIANT right);
VARIANT vbtVarSub(VARIANT left, VARIANT right);
VARIANT vbtVarMul(VARIANT left, VARIANT right);
VARIANT vbtVarDiv(VARIANT left, VARIANT right);
VARIANT vbtVarIDiv(VARIANT left, VARIANT right);
VARIANT vbtVarLT(VARIANT left, VARIANT right); /* better to return INT16 (Boolean)? */
VARIANT vbtVarLE(VARIANT left, VARIANT right); /* better to return INT16 (Boolean)? */
VARIANT vbtVarEQ(VARIANT left, VARIANT right); /* better to return INT16 (Boolean)? */
VARIANT vbtVarNE(VARIANT left, VARIANT right); /* better to return INT16 (Boolean)? */
VARIANT vbtVarGT(VARIANT left, VARIANT right); /* better to return INT16 (Boolean)? */
VARIANT vbtVarGE(VARIANT left, VARIANT right); /* better to return INT16 (Boolean)? */
VARIANT vbtVarAbs(VARIANT arg);
VARIANT vbtVarFix(VARIANT arg);
VARIANT vbtVarInt(VARIANT *arg);
VARIANT vbtVarMod(VARIANT left, VARIANT right);
VARIANT vbtVarPow(VARIANT left, VARIANT right);
VARIANT vbtVarNot(VARIANT arg);
VARIANT vbtVarNeg(VARIANT arg);

/* Input Statement */
#define _InputStart(fn) ;
#define _InputVariable(v) ;
#define _InputEnd ;

/* Print Statement */
#define Print NULL /* Temp! - for Print "Hello" - Need to output correct Print object (Form1) */
void vbtPrintMethod(IDispatch *o);
void vbtPrintStart(INT16 fn); /* rename to vbtPrintFN */
void vbtPrintExpr(VARIANT v);
void vbtPrintComma(void);
void vbtPrintNL(void);
void vbtPrintSpc(INT16 n);
void vbtPrintTab0();
void vbtPrintTab1(INT16 t);
void vbtPrintEnd(void);
#define _PrintMethod(o) vbtPrintMethod((IDispatch *)(o));
#define _PrintStart(fn) vbtPrintStart(fn);
#define _PrintExpr(v) vbtPrintExpr(v);
#define _PrintComma vbtPrintComma();
#define _PrintNL vbtPrintNL();
#define _PrintSpc(n) vbtPrintSpc(n);
#define _PrintTab0() vbtPrintTab0();
#define _PrintTab1(t) vbtPrintTab1(t);
#define _PrintEnd vbtPrintEnd();

/* Debug Statement */
void vbtDebugAssert(VARIANT_BOOL b);
void vbtDebugPrint();
#define _DebugAssert(b) vbtDebugAssert(b);
#define _DebugPrint() vbtDebugPrint();

/* Write Statement */
void vbtWriteStart(INT16 fn);
void vbtWriteExpr(VARIANT v);
void vbtWriteComma(void);
void vbtWriteNL(void);
void vbtWriteEnd(void);
#define _WriteStart(fn) 
#define _WriteExpr(v) 
#define _WriteComma vbtWriteComma();
#define _WriteNL vbtWriteNL();
#define _WriteEnd vbtWriteEnd();

/* Do/Loop Statement */
#define Do do {
#define DoWhile(e) while(e) {
#define DoUntil(e) DoWhile(!(e))
#define ExitDo break;
#define Loop }
#define LoopInfinite } while(1);
#define LoopWhile(e) } while(e);
#define LoopUntil(e) LoopWhile(!(e))
#ifdef NEVER /* always boolean - right? */
#define ByteDoWhile(e) DoWhile(e)
#define BoolDoWhile(e) DoWhile(e)
#define IntDoWhile(e) DoWhile(e)
#define LngDoWhile(e) DoWhile(e)
#define SngDoWhile(e) DoWhile(e)
#define DateDoWhile(e) DoWhile(e)
#define DblDoWhile(e) DoWhile(e)
#define CurDoWhile(e) DoWhile(e)
#define StrDoWhile(e) DoWhile(e)
#define ObjDoWhile(e) DoWhile(e)
#define VarDoWhile(e) DoWhile(e)
#define DecDoWhile(e) DoWhile(e)
#define ByteDoUntil(e) DoUntil(e)
#define BoolDoUntil(e) DoUntil(e)
#define IntDoUntil(e) DoUntil(e)
#define LngDoUntil(e) DoUntil(e)
#define SngDoUntil(e) DoUntil(e)
#define DateDoUntil(e) DoUntil(e)
#define DblDoUntil(e) DoUntil(e)
#define CurDoUntil(e) DoUntil(e)
#define StrDoUntil(e) DoUntil(e)
#define ObjDoUntil(e) DoUntil(e)
#define VarDoUntil(e) DoUntil(e)
#define DecDoUntil(e) DoUntil(e)
#define ByteLoopWhile(e) LoopWhile(e)
#define BoolLoopWhile(e) LoopWhile(e)
#define IntLoopWhile(e) LoopWhile(e)
#define LngLoopWhile(e) LoopWhile(e)
#define SngLoopWhile(e) LoopWhile(e)
#define DateLoopWhile(e) LoopWhile(e)
#define DblLoopWhile(e) LoopWhile(e)
#define CurLoopWhile(e) LoopWhile(e)
#define StrLoopWhile(e) LoopWhile(e)
#define ObjLoopWhile(e) LoopWhile(e)
#define VarLoopWhile(e) LoopWhile(e)
#define DecLoopWhile(e) LoopWhile(e)
#define ByteLoopUntil(e) LoopUntil(e)
#define BoolLoopUntil(e) LoopUntil(e)
#define IntLoopUntil(e) LoopUntil(e)
#define LngLoopUntil(e) LoopUntil(e)
#define SngLoopUntil(e) LoopUntil(e)
#define DateLoopUntil(e) LoopUntil(e)
#define DblLoopUntil(e) LoopUntil(e)
#define CurLoopUntil(e) LoopUntil(e)
#define StrLoopUntil(e) LoopUntil(e)
#define ObjLoopUntil(e) LoopUntil(e)
#define VarLoopUntil(e) LoopUntil(e)
#define DecLoopUntil(e) LoopUntil(e)
#endif

/* Call Statement */
#define Call(c) c;
#define ByteCall(c) c;
#define BoolCall(c) c;
#define IntCall(c) c;
#define LngCall(c) c;
#define SngCall(c) c;
#define DateCall(c) c;
#define DblCall(c) c;
#define CurCall(c) c;
#define StrCall(c) c;
#define ObjCall(c) c;
#define VarCall(c) c;
#define DecCall(c) c;
#define VoidCall(c) c; /* for methods that return VT_VOID */
#define HResultCall(c) c; /* Form.Show */
#define IUnknownCall(c) c; /* _i_DxVBLib_DirectMusicPerformance_PlaySegment */
#define INTCall(c) c; /* fixme: arrrrgggg, need more defs */

/* ForEach Statement */
#define ForEach(e,g) for(;;) { /* not implemented */
#define ObjForEach(e,g) ForEach(e,g)
#define ObjRefForEach(e,g) ObjForEach(*(e),g)
#define VarForEach(e,g) ForEach(e,g)
#define VarRefForEach(e,g) VarForEach(*(e),g)
#define ByteSAForEach(e,g) ForEach(e,g)
#define BoolSAForEach(e,g) ForEach(e,g)
#define IntSAForEach(e,g) ForEach(e,g)
#define LngSAForEach(e,g) ForEach(e,g)
#define SngSAForEach(e,g) ForEach(e,g)
#define DateSAForEach(e,g) ForEach(e,g)
#define DblSAForEach(e,g) ForEach(e,g)
#define CurSAForEach(e,g) ForEach(e,g)
#define StrSAForEach(e,g) ForEach(e,g)
#define ObjSAForEach(e,g) ForEach(e,g)
#define ObjRefSAForEach(e,g) ForEach(*(e),g)
#define VarSAForEach(e,g) ForEach(e,g)
#define VarRefSAForEach(e,g) ForEach(*(e),g)
#define DecSAForEach(e,g) ForEach(e,g)

/* For/Next Statement */
#define For(v,start,stop,step) for((v)=(start);(v)<=(stop);(v)+=(step)) {
#define ByteFor(v,start,stop,step) For(v,start,stop,step)
#define ByteRefFor(_v,start,stop,step) For(*(_v),start,stop,step)
#define BoolFor(v,start,stop,step) For(v,start,stop,step)
#define BoolRefFor(_v,start,stop,step) For(*(_v),start,stop,step)
#define IntFor(v,start,stop,step) For(v,start,stop,step)
#define IntRefFor(_v,start,stop,step) For(*(_v),start,stop,step)
#define LngFor(v,start,stop,step) For(v,start,stop,step)
#define LngRefFor(_v,start,stop,step) For(*(_v),start,stop,step)
#define SngFor(v,start,stop,step) For(v,start,stop,step)
#define SngRefFor(_v,start,stop,step) For(*(_v),start,stop,step)
#define DateFor(v,start,stop,step) For(v,start,stop,step)
#define DateRefFor(_v,start,stop,step) For(*(_v),start,stop,step)
#define DblFor(v,start,stop,step) For(v,start,stop,step)
#define DblRefFor(_v,start,stop,step) For(*(_v),start,stop,step)
#if 1 /* fixme: test CurFor */
#define CurFor(v,start,stop,step) For((v).int64,start.int64,stop.int64,step.int64)
#define CurRefFor(v,start,stop,step) For((v)->int64,start.int64,stop.int64,step.int64)
#else
int vbtCurFor(CY *v,CY start,CY stop,CY step);
#define CurFor(v,start,stop,step) while(vbtCurFor(&v,start,stop,step)) {
#define CurRefFor(v,start,stop,step) while(vbtCurFor(v,start,stop,step)) {
#endif
#define StrFor(v,start,stop,step) For(v,start,stop,step)
#define StrRefFor(_v,start,stop,step) For(*(_v),start,stop,step)
#define ObjFor(v,start,stop,step) For(v,start,stop,step)
#define ObjRefFor(_v,start,stop,step) For(*(_v),start,stop,step)
/*#define VarFor(v,start,stop,step) for((v)=(start);VarLE(v,stop).boolVal;VarAdd(v,step)) {*/
int vbtVarFor(VARIANT *v,VARIANT start,VARIANT stop,VARIANT step);
#define VarFor(v,start,stop,step) while(vbtVarFor(&v,start,stop,step)) {
#define VarRefFor(v,start,stop,step) while(vbtVarFor(v,start,stop,step)) {
#define DecFor(v,start,stop,step) For(v,start,stop,step)
#define DecRefFor(_v,start,stop,step) For(*(_v),start,stop,step)

#define Next }
#define NextV } /* is this necessary? NextV(v)? */
#define ByteNext Next
#define BoolNext Next
#define IntNext Next
#define LngNext Next
#define SngNext Next
#define DateNext Next
#define DblNext Next
#define CurNext Next
#define StrNext Next
#define ObjNext Next
#define VarNext Next
#define DecNext Next
#define ByteRefNext Next
#define BoolRefNext Next
#define IntRefNext Next
#define LngRefNext Next
#define SngRefNext Next
#define DateRefNext Next
#define DblRefNext Next
#define CurRefNext Next
#define StrRefNext Next
#define ObjRefNext Next
#define VarRefNext Next
#define DecRefNext Next

#define ExitFor break;

/* GoSub Statement */
#define GoSub(l) goto l;	/* not implemented */

/* GoTo Statement */
#define GoTo(l) goto l;

/* On GoSub Statement */	/* not implemented */
#define OnGoSub(e) switch(e) { default:
#define IntOnGoSub(e) OnGoSub(e)
#define LabelOnGoSub(c,l) case c: GoSub(l)
#define EndOnGoSub }
#define Return 0; /* use GoSubReturn to avoid confusion with return(0) ? */

/* On GoTo Statement */		/* not implemented */
#define OnGoTo(e) switch(e) { default:
#define IntOnGoTo(e) OnGoTo(e)
#define LabelOnGoTo(c,l) case c: GoTo(l)
#define EndOnGoTo }

/* If/ElseIf/Else/Endif Statement */
#define If(e) if (e) {
#define Else } else {
#define EndIfElse } } else {
#define ElseIf(e) } else if (e) {
#define EndIf }

/* Mid, MidB */
void vbtStrMid(BSTR StringVar, INT32 Start, INT32 Length, BSTR StringExpr);
void vbtStrMidB(BSTR StringVar, INT32 Start, INT32 Length, BSTR StringExpr);
#define StrMid(StringVar, Start, Length, StringExpr) vbtStrMid(StringVar,Start,Length,StringExpr);
#define StrMidB(StringVar, Start, Length, StringExpr) vbtStrMid(StringVar,Start,Length,StringExpr);
#define StrRefMid(StringVar, Start, Length, StringExpr) vbtStrMid(*(StringVar),Start,Length,StringExpr);
#define StrRefMidB(StringVar, Start, Length, StringExpr) vbtStrMid(*(StringVar),Start,Length,StringExpr);
void vbtStrNMid(LPWSTR StringVar, INT32 Start, INT32 Length, BSTR StringExpr);
void vbtStrNMidB(LPWSTR StringVar, INT32 Start, INT32 Length, BSTR StringExpr);
#define StrNMid(StringVar, Start, Length, StringExpr) vbtStrNMid(StringVar,Start,Length,StringExpr);
#define StrNMidB(StringVar, Start, Length, StringExpr) vbtStrNMid(StringVar,Start,Length,StringExpr);

/* RaiseEvent */
#define RaiseEvent(c) c;

/* Select Statement */
void vbtByteSelectCase(UINT8 e);
/* if (0) { needed when only Case Else appears */
#define ByteSelectCase(e) vbtByteSelectCase(e); if (0) {
void vbtBoolSelectCase(VARIANT_BOOL e);
#define BoolSelectCase(e) vbtBoolSelectCase(e); if (0) {
void vbtIntSelectCase(INT16 e);
#define IntSelectCase(e) vbtIntSelectCase(e); if (0) {
void vbtLngSelectCase(INT32 e);
#define LngSelectCase(e) vbtLngSelectCase(e); if (0) {
void vbtSngSelectCase(FLOAT e);
#define SngSelectCase(e) vbtSngSelectCase(e); if (0) {
void vbtDateSelectCase(DATE e);
#define DateSelectCase(e) vbtDateSelectCase(e); if (0) {
void vbtDblSelectCase(DOUBLE e);
#define DblSelectCase(e) vbtDblSelectCase(e); if (0) {
void vbtCurSelectCase(CURRENCY e);
#define CurSelectCase(e) vbtCurSelectCase(e); if (0) {
void vbtStrSelectCase(BSTR e);
#define StrSelectCase(e) vbtStrSelectCase(e); if (0) {
/*void vbtObjSelectCase(IDispatch *e);
#define ObjSelectCase(e) vbtObjSelectCase(e); if (0) {*/
void vbtVarSelectCase(VARIANT e);
#define VarSelectCase(e) vbtVarSelectCase(e); if (0) {
void vbtDecSelectCase(DECIMAL e);
#define DecSelectCase(e) vbtDecSelectCase(e); if (0) {
/* Need to finish implementing Case stuff */
#define Case(e) } if (e) {
#define ByteCase(e) ByteCaseIsEQ(e)
#define BoolCase(e) BoolCaseIsEQ(e)
#define IntCase(e) IntCaseIsEQ(e)
#define LngCase(e) LngCaseIsEQ(e)
#define SngCase(e) SngCaseIsEQ(e)
#define DateCase(e) DateCaseIsEQ(e)
#define DblCase(e) DblCaseIsEQ(e)
#define CurCase(e) CurCaseIsEQ(e)
#define StrCase(e) StrCaseIsEQ(e)
/*#define ObjCase(e) ObjCaseIsEQ(e)*/
#define VarCase(e) VarCaseIsEQ(e)
#define DecCase(e) DecCaseIsEQ(e)
#define ByteCaseIsLT(e) (1)
#define BoolCaseIsLT(e) (1)
#define IntCaseIsLT(e) (1)
#define LngCaseIsLT(e) (1)
#define SngCaseIsLT(e) (1)
#define DateCaseIsLT(e) (1)
#define DblCaseIsLT(e) (1)
#define CurCaseIsLT(e) (1)
#define StrCaseIsLT(e) (1)
/*#define ObjCaseIsLT(e) (1)*/
#define VarCaseIsLT(e) (1)
#define DecCaseIsLT(e) (1)
#define ByteCaseIsLE(e) (1)
#define BoolCaseIsLE(e) (1)
#define IntCaseIsLE(e) (1)
#define LngCaseIsLE(e) (1)
#define SngCaseIsLE(e) (1)
#define DateCaseIsLE(e) (1)
#define DblCaseIsLE(e) (1)
#define CurCaseIsLE(e) (1)
#define StrCaseIsLE(e) (1)
/*#define ObjCaseIsLE(e) (1)*/
#define VarCaseIsLE(e) (1)
#define DecCaseIsLE(e) (1)
#define ByteCaseIsEQ(e) (1)
#define BoolCaseIsEQ(e) (1)
#define IntCaseIsEQ(e) (1)
#define LngCaseIsEQ(e) (1)
#define SngCaseIsEQ(e) (1)
#define DateCaseIsEQ(e) (1)
#define DblCaseIsEQ(e) (1)
#define CurCaseIsEQ(e) (1)
#define StrCaseIsEQ(e) (1)
/*#define ObjCaseIsEQ(e) (1)*/
#define VarCaseIsEQ(e) (1)
#define DecCaseIsEQ(e) (1)
#define ByteCaseIsGT(e) (1)
#define BoolCaseIsGT(e) (1)
#define IntCaseIsGT(e) (1)
#define LngCaseIsGT(e) (1)
#define SngCaseIsGT(e) (1)
#define DateCaseIsGT(e) (1)
#define DblCaseIsGT(e) (1)
#define CurCaseIsGT(e) (1)
#define StrCaseIsGT(e) (1)
/*#define ObjCaseIsGT(e) (1)*/
#define VarCaseIsGT(e) (1)
#define DecCaseIsGT(e) (1)
#define ByteCaseIsGE(e) (1)
#define BoolCaseIsGE(e) (1)
#define IntCaseIsGE(e) (1)
#define LngCaseIsGE(e) (1)
#define SngCaseIsGE(e) (1)
#define DateCaseIsGE(e) (1)
#define DblCaseIsGE(e) (1)
#define CurCaseIsGE(e) (1)
#define StrCaseIsGE(e) (1)
/*#define ObjCaseIsGE(e) (1)*/
#define VarCaseIsGE(e) (1)
#define DecCaseIsGE(e) (1)
#define ByteCaseIsNE(e) (1)
#define BoolCaseIsNE(e) (1)
#define IntCaseIsNE(e) (1)
#define LngCaseIsNE(e) (1)
#define SngCaseIsNE(e) (1)
#define DateCaseIsNE(e) (1)
#define DblCaseIsNE(e) (1)
#define CurCaseIsNE(e) (1)
#define StrCaseIsNE(e) (1)
/*#define ObjCaseIsNE(e) (1)*/
#define VarCaseIsNE(e) (1)
#define DecCaseIsNE(e) (1)
#define ByteCaseTo(e1,e2) (1)
#define BoolCaseTo(e1,e2) (1)
#define IntCaseTo(e1,e2) (1)
#define LngCaseTo(e1,e2) (1)
#define SngCaseTo(e1,e2) (1)
#define DateCaseTo(e1,e2) (1)
#define DblCaseTo(e1,e2) (1)
#define CurCaseTo(e1,e2) (1)
#define StrCaseTo(e1,e2) (1)
/*#define ObjCaseTo(e1,e2) (1)*/
#define VarCaseTo(e1,e2) (1)
#define DecCaseTo(e1,e2) (1)
#define CaseElse } else {
#define EndSelect }
#define Or ||

/* While/Wend Statement */
#define BoolWhile(e) while(e) {
#define BoolWend }

/* don't need {, w is temp With var, e is object expression */
#define ObjWith(w,e) {
#define ObjRefWith(w,e) {
#define IUnknownWith(w,e) {
#define IUnknownRefWith(w,e) {
#define UDTWith(w,e) {
#define UDTRefWith(w,e) {
#define VarWith(w,e) {
#define VarRefWith(w,e) {
#define ObjEndWith(w) }
#define ObjRefEndWith(w) }
#define IUnknownEndWith(w) }
#define IUnknownRefEndWith(w) }
#define UDTEndWith(w) }
#define UDTRefEndWith(w) }
#define VarEndWith(w) }
#define VarRefEndWith(w) }
/* fixme: need explicit type (generate in .h?) or cast instead of using void
   as in Form1WithValue1 or Form1WithValue(1) */
void *ObjWithValue(int n);
void *IUnknownWithValue(int n);
//void **ObjRefWithValue(int n);
VARIANT VarWithValue(int n);
//VARIANT *VarRefWithValue(int n);
void *UDTWithValue(int n);
//void *UDTRefWithValue(int n);

#define True 1 /* use OLE definition instead of 1 */
#define False 0

#define Erase(args) vbtErase args ;
#define ReDim(args) vbtReDim args ;
#define ReDimPreserve(args) vbtReDimPreserve args ;

#define ExitFunction goto _method_terminate;
#define ExitProperty goto _method_terminate;
#define ExitSub goto _method_terminate;

void vbtErase(int n, ...);
void vbtReDim(int n, ...);
void vbtReDimPreserve(int n, ...);
void vbtCloseFile(int n, ...);
void vbtGet(INT16 FileNumber, INT32 RecNumber, INT16 vt, void *Variable);
void vbtPut(INT16 FileNumber, INT32 RecNumber, INT16 vt, ...);
void vbtStop(void);
void vbtInput(INT16 FileNumber, void *Variable);
void vbtLineInput(INT16 FileNumber, void *Variable);
void vbtLock(INT16 FileNumber, INT32 RecStart, INT32 RecEnd);
void vbtName(BSTR OldFileName, BSTR NewFileName);
void vbtOpen(BSTR FileName, INT16 Flags, INT16 FileNumber, INT16  RecLen);
void vbtSeek(INT16 FileNumber, INT32 Position);
void vbtUnlock(INT16 FileNumber, INT32 RecStart, INT32 RecEnd);
void vbtWidth(INT16 FileNumber, UINT8 Width);

#define CloseFile(args) vbtCloseFile args ;
#define Get(FileNumber, RecNumber, vt, Variable) vbtGet(FileNumber, RecNumber, vt, &Variable);
#define Put(FileNumber, RecNumber, vt, Variable) vbtPut(FileNumber, RecNumber, vt, Variable);
#define Stop vbtStop();
#define Input(FileNumber, Variable) vbtInput(FileNumber, &Variable);
#define LineInput(FileNumber, Variable) vbtLineInput(FileNumber, &Variable);
#define Lock(FileNumber, RecStart, RecEnd) vbtLock(FileNumber, RecStart, RecEnd);
#define Name(OldFileName, NewFileName) vbtName(OldFileName,NewFileName);
#define Open(FileName, Flags, FileNumber,  RecLen) vbtOpen(FileName, Flags, FileNumber, RecLen);
#define Seek(FileNumber, Position) vbtSeek(FileNumber, Position);
#define Unlock(FileNumber, RecStart, RecEnd) vbtUnlock(FileNumber, RecStart, RecEnd);
#define Width(FileNumber, Width) vbtWidth(FileNumber, Width);

#define SingleIf(b) if (b) {
#define SingleIfEndIf() }
#define Resume(l) goto l; /* not completed!!! */
#define Resume0 ; /* not completed!!! */
#define ResumeNext ; /* not completed!!! */
#define OnErrorResumeNext ; /* not completed!!! */
#define OnErrorGoTo0 ; /* not completed!!! */
#define OnErrorGoTo(l) ; /* not completed!!! */
#define VBA_Int(e) vbtVarInt(e)
/*#define OnError(l)*/
void vbtError(INT32 n); /* note: Doc says n is UINT16. So do what you gotta do within vbtError, instead of VBT.(?) */
#define Error(n) vbtError(n);

/* fixme: needs reorging */
#define _IUnknownQI(o) ((IUnknown *)(o)) /* fixme: still required but s/b stdole_IUnknownQI */
IDispatch *_IDispatchQI(void *o); /* fixme: still required but s/b stdole_IDispatchQI */
void vbtIUnknownSetNothing(void *o); /* Does an IUnknown_Release and zeros o */
#define IUnknownRefToObj(v) _IDispatchQI(*(v)) /* tlib.IUnkMember Is Nothing */
/*#define ObjIUnknownNothing(n,o) vbtObjIUnknownNothing(o);*/
#define IUnknownSetNothing(o) vbtIUnknownSetNothing(o);
#define IUnknownRefSetNothing(o) vbtIUnknownSetNothing(*(o));
#define IUnknownRef(o) ObjRef(o)
VARIANT vbtIUnknownVar(IUnknown *o);
#define IUnknownToVar(o) vbtIUnknownVar((IUnknown *)(o))
#define IUnknownRefToVar(o) IUnknownToVar(*(o))
void vbtObjSetNothing(void *o); /* Does an IUnknown_Release and zeros o */
#define ObjSetNothing(o) vbtObjSetNothing(o);
#define ObjRefSetNothing(o) vbtObjSetNothing(*(o));
void vbtVarSetNothing(VARIANT v);
#define VarSetNothing(v) vbtVarSetNothing(v);
#define VarRefSetNothing(v) vbtVarSetNothing(*(v));
#define ObjPropertySetNothing(o) (o);
#define IUnknownPropertySetNothing(o) (o);
#define VarPropertySetNothing(v) (v);
#if 0 /* probably have to turn on? */
IDispatch *vbtVarToObj(VARIANT v);
#define VarToObj(v) vbtVarToObj(v)
#else
#ifndef SAG_COM
#define VarToObj(v) v.pdispVal /* not good for temp expression */
#define VarRefToObj(v) v->pdispVal
#else
#define VarToObj(v) v.n1.n2.n3.pdispVal
#define VarRefToObj(v) v->n1.n2.n3.pdispVal
#endif
#endif
HRESULT vbtIUnknownSet(void *ro, void **lo,LPWSTR iid); /* use IUnknown instead of void? */
#define IUnknownSet(ro,lo,iid) vbtIUnknownSet(ro,&lo,iid);
#define IUnknownRefSet(ro,lo,iid) vbtIUnknownSet(ro,lo,iid);
HRESULT vbtObjSet(void *ro, void **lo,LPWSTR iid); /* use IDispatch instead of void? */
#define ObjSet(ro,lo,iid) vbtObjSet(ro,&lo,iid);
#define ObjRefSet(ro,lo,iid) vbtObjSet(ro,lo,iid);
/* fixme: Instead of Unknown/Dispatch Set, generate Set for each interface??? */
void vbtIUnknownLet(void *ro, void **lo); /* use IUnknown instead of void? */
#define IUnknownLet(ro,lo) vbtIUnknownLet(ro,&lo);
#define IUnknownRefLet(ro,lo) vbtIUnknownLet(ro,lo);
void vbtObjLet(void *ro, void **lo); /* use IDispatch instead of void? */
#define ObjLet(ro,lo) vbtObjLet(ro,&lo);
#define ObjRefLet(ro,lo) vbtObjLet(ro,lo);

UINT8 vbtObjByte(IDispatch *o);
#define ObjToByte(o) vbtObjByte((IDispatch *)(o))
INT16 vbtObjBool(IDispatch *o);
#define ObjToBool(o) vbtObjBool((IDispatch *)(o))
INT16 vbtObjInt(IDispatch *o);
#define ObjToInt(o) vbtObjInt((IDispatch *)(o))
INT32 vbtObjLng(IDispatch *o);
#define ObjToLng(o) vbtObjLng((IDispatch *)(o))
FLOAT vbtObjSng(IDispatch *o);
#define ObjToSng(o) vbtObjSng((IDispatch *)(o))
DOUBLE vbtObjDbl(IDispatch *o);
#define ObjToDbl(o) vbtObjDbl((IDispatch *)(o))
DATE vbtObjDate(IDispatch *o);
#define ObjToDate(o) vbtObjDate((IDispatch *)(o))
CY vbtObjCy(IDispatch *o);
#define ObjToCy(o) vbtObjCy((IDispatch *)(o))
BSTR vbtObjStr(IDispatch *o);
#define ObjToStr(o) vbtObjStr((IDispatch *)(o))
VARIANT vbtObjVar(IDispatch *o);
#define ObjToVar(o) vbtObjVar((IDispatch *)(o))
#define ObjRefToVar(o) ObjToVar(*(o))
DECIMAL vbtObjDec(IDispatch *o);
#define ObjToDec(o) vbtObjDec((IDispatch *)(o))

/* fixme: Need defines for all VT_xxx? */
#define VarToIUnknown(o) _IUnknownQI(VarToObj(o))
#define VarRefToIUnknown(o) _IUnknownQI(VarRefToObj(o))
#define ObjToIUnknown(o) _IUnknownQI(o)
#define ObjRefToIUnknown(o) _IUnknownQI(*(o))
#define IUnknownToObj(o) _IDispatchQI(o)

void vbtVarSet(VARIANT ro, VARIANT *lo);
#define VarSet(ro,lo) vbtVarSet(ro,&lo);
void vbtVarRefSet(VARIANT ro, VARIANT *lo);
#define VarRefSet(ro,lo) vbtVarSet(ro,lo);
IDispatch *vbtObjNew(wchar_t *progID); /* always return IDispatch or IUnknown? */
#define ObjNew(progID) vbtObjNew(progID)
/*void *vbtGetObj(void *obj);*/
#define GetObj(obj) /*vbtGetObj(obj)*/ (obj)

UINT8 *SaToByte(SAFEARRAY *sa,int nDims,...);
INT16 *SaToBool(SAFEARRAY *sa,int nDims,...);
INT16 *SaToInt(SAFEARRAY *sa,int nDims,...);
INT32 *SaToLng(SAFEARRAY *sa,int nDims,...);
FLOAT *SaToSng(SAFEARRAY *sa,int nDims,...);
DATE *SaToDate(SAFEARRAY *sa,int nDims,...);
DOUBLE *SaToDbl(SAFEARRAY *sa,int nDims,...);
CURRENCY *SaToCur(SAFEARRAY *sa,int nDims,...);
BSTR *SaToStr(SAFEARRAY *sa,int nDims,...);
BSTR *SaToStrN(SAFEARRAY *sa,int nDims,...);
void *SaToObj(SAFEARRAY *sa,int nDims,...);
VARIANT *SaToVar(SAFEARRAY *sa,int nDims,...);
DECIMAL *SaToDec(SAFEARRAY *sa,int nDims,...);
/*void *UDTSA(SAFEARRAY *sa,int nDims,...);*/ /* obsolete this by generating specific UDT SA protos */

VARIANT VarArgsGet(VARIANT v,SAFEARRAY *sa);
VARIANT VarArgsLet(VARIANT lhs,SAFEARRAY *sa, VARIANT rhs);
VARIANT VarArgsSet(VARIANT lhs,SAFEARRAY *sa, VARIANT rhs);
VARIANT *VarRefArgsGet(VARIANT *v,SAFEARRAY *sa);
VARIANT *VarRefArgsLet(VARIANT *lhs,SAFEARRAY *sa, VARIANT rhs);
VARIANT *VarRefArgsSet(VARIANT *lhs,SAFEARRAY *sa, VARIANT rhs);

SAFEARRAY *VarArg(int nArgs,...); /* VarArg (ParamArray) can only be Variant */

SAFEARRAY *vbtVarToUDTSA(VARIANT v);
#define VarToUDTSA(v) vbtVarToUDTSA(v)
void vbtByteToSALet(UINT8 b,SAFEARRAY *sa); /* base641a\project1.vbp */
#define ByteToSALet(v,lv) vbtByteSALet(v,lv);
void vbtDblToSALet(DOUBLE d,SAFEARRAY *sa);
#define DblToSALet(v,lv) vbtDblSALet(v,lv);

void vbtSALet(SAFEARRAY *v,SAFEARRAY *lv); /* fixme: Change VBT to recognize array assignment. Don't deref RHS (unless ByRef) */
#define ByteSALet(v,lv) vbtSALet(v,lv); /* fixme: need more RefSA Let defs */
#define IntSALet(v,lv) vbtSALet(v,lv); /* fixme: need more RefSA Let defs */
#define StrSALet(v,lv) vbtSALet(v,lv); /* fixme: need more RefSA Let defs */
#define DblSALet(v,lv) vbtSALet(v,lv); /* fixme: need more SA Let defs */
#define VarSALet(v,lv) vbtSALet(v,lv); /* fixme: need more SA Let defs */
#define UDTSALet(v,lv) vbtSALet(v,lv); /* fixme: need more SA Let defs */

VARIANT ByteSAToVar(SAFEARRAY *sa); /* fixme: need more xxxSAToVar */
VARIANT BoolArrayToVar(SAFEARRAY *sa);
VARIANT IntSAToVar(SAFEARRAY *sa);
VARIANT LngSAToVar(SAFEARRAY *sa);
VARIANT SngSAToVar(SAFEARRAY *sa);
VARIANT DateSAToVar(SAFEARRAY *sa);
VARIANT DblSAToVar(SAFEARRAY *sa);
VARIANT CurSAToVar(SAFEARRAY *sa);
VARIANT StrSAToVar(SAFEARRAY *sa);
VARIANT ObjSAToVar(SAFEARRAY *sa);
VARIANT VarSAToVar(SAFEARRAY *sa);
VARIANT DecSAToVar(SAFEARRAY *sa);

VARIANT *ByteSAToVarRef(SAFEARRAY *sa); /* fixme: need more xxxSAToVar */
VARIANT *BoolSAToVarRef(SAFEARRAY *sa);
VARIANT *IntSAToVarRef(SAFEARRAY *sa);
VARIANT *LngSAToVarRef(SAFEARRAY *sa);
VARIANT *SngSAToVarRef(SAFEARRAY *sa);
VARIANT *DateSAToVarRef(SAFEARRAY *sa);
VARIANT *DblSAToVarRef(SAFEARRAY *sa);
VARIANT *CurSAToVarRef(SAFEARRAY *sa);
VARIANT *StrSAToVarRef(SAFEARRAY *sa);
VARIANT *ObjSAToVarRef(SAFEARRAY *sa);
VARIANT *VarSAToVarRef(SAFEARRAY *sa);
VARIANT *DecSAToVarRef(SAFEARRAY *sa);

SAFEARRAY *VarToByteSA(VARIANT v); /* fixme: need more VarToxxxSA( */
SAFEARRAY *VarToBoolSA(VARIANT v);
SAFEARRAY *VarToIntSA(VARIANT v);
SAFEARRAY *VarToLngSA(VARIANT v);
SAFEARRAY *VarToSngSA(VARIANT v);
SAFEARRAY *VarToDateSA(VARIANT v);
SAFEARRAY *VarToDblSA(VARIANT v);
SAFEARRAY *VarToCurArray(VARIANT v);
SAFEARRAY *VarToStrSA(VARIANT v);
SAFEARRAY *VarToObjSA(VARIANT v);
SAFEARRAY *VarToVarSA(VARIANT v);
SAFEARRAY *VarToDecSA(VARIANT v);

SAFEARRAY *VarRefToByteSA(VARIANT *v); /* fixme: need more VarRefToxxxSA( */
SAFEARRAY *VarRefToBoolSA(VARIANT *v);
SAFEARRAY *VarRefToIntSA(VARIANT *v);
SAFEARRAY *VarRefToLngSA(VARIANT *v);
SAFEARRAY *VarRefToSngSA(VARIANT *v);
SAFEARRAY *VarRefToDateSA(VARIANT *v);
SAFEARRAY *VarRefToDblSA(VARIANT *v);
SAFEARRAY *VarRefToCurSA(VARIANT *v);
SAFEARRAY *VarRefToStrSA(VARIANT *v);
SAFEARRAY *VarRefToObjSA(VARIANT *v);
SAFEARRAY *VarRefToVarSA(VARIANT *v);
SAFEARRAY *VarRefToDecSA(VARIANT *v);

/*#define ByRef(e) (&(e))*/
UINT8 *ByteRef(UINT8 u8);
INT16 *BoolRef(INT16 i16);
INT16 *IntRef(INT16 i16);
INT32 *LngRef(INT32 i32);
FLOAT *SngRef(FLOAT f);
DATE *DateRef(DATE date);
DOUBLE *DblRef(DOUBLE d);
CURRENCY *CurRef(CURRENCY cy);
BSTR *StrRef(BSTR bstr);
BSTR *StrNRef(wchar_t *str);
/*void *UDTRef(void *);*/
/* UDTRef not implemented - needs pointerlevel */
#define UDTRef(udt) (NULL) /* appbar4\appbar.vbp - pass UDT as ByRef parameter */
VARIANT vbtUDTRefToVar(void *udt); /* audiol1a\project1.vbp */
#define UDTRefToVar(udt) vbtUDTRefToVar(udt)
#define UDTRefToVoidRef(udt) ((void *)(LONG)(udt))
/* void *VoidRef(void *); */ /* for the "As Any" data type */
/* parameter void * could be IDispatch * but requires casting */
/* return type of void * could be IDispatch ** but requires casting */
void *vbtObjRef(void *pdisp);
#define ObjRef(o) vbtObjRef(o)
VARIANT *VarRef(VARIANT v);
DECIMAL *DecRef(DECIMAL d);

/* Miscelleneous other conversions */
/* fixme: Ugh, how many others are needed? Perhaps outputting "ToInt" is needed to generalize? */
#define INTToInt(v) ToInt(v) /* converts int [sic] to Int [INT16] - Help.vbp */

SAFEARRAY *ByteSA(SAFEARRAY *sa);
SAFEARRAY *BoolSA(SAFEARRAY *sa);
SAFEARRAY *IntSA(SAFEARRAY *sa);
SAFEARRAY *LngSA(SAFEARRAY *sa);
SAFEARRAY *SngSA(SAFEARRAY *sa);
SAFEARRAY *DateSA(SAFEARRAY *sa);
SAFEARRAY *DblSA(SAFEARRAY *sa);
SAFEARRAY *CurSA(SAFEARRAY *sa);
SAFEARRAY *StrSA(SAFEARRAY *sa);
SAFEARRAY *ObjSA(SAFEARRAY *sa);
SAFEARRAY *VarSA(SAFEARRAY *sa);
SAFEARRAY *DecSA(SAFEARRAY *sa);


/* Missing isn't fully implemented */
#define _ByteMissing (UINT8)0
#define _BoolMissing (INT16)0
#define _IntMissing (INT16)0
#define _LngMissing (INT32)0
#define _SngMissing (FLOAT)0
#define _DateMissing (DATE)0
#define _DblMissing (DOUBLE)0
#define _CurMissing (CURRENCY)0
#define _StrMissing (BSTR)NULL
#define _ObjMissing (IDispatch *)Nothing
/* fixme: probably need to make _Var macros into functions that return VARIANT */
extern VARIANT _VarMissing; /* returns VT_ERROR with DISP_E_PARAMNOTFOUND value */
extern VARIANT _VarNull;
extern VARIANT _VarEmpty;
extern SAFEARRAY *_VarSAMissing;
#define _DecMissing (DECIMAL)0
VARIANT Empty;

/* void * could be IDispatch * but requires casting */
/* replace wfuncname with MemID where possible */
VARIANT ObjInvoke(IDispatch *pdisp, LPWSTR wfuncname, UINT dispatch_type, UINT nargs, ...);
VARIANT ObjRefInvoke(IDispatch **pdisp, LPWSTR wfuncname, UINT dispatch_type, UINT nargs, ...);
/* using void for IUnknown Invoke, useless to coerce to IUnknown, although could coerce to IDispatch */
/*VARIANT IUnknownInvoke(void *pdisp, LPWSTR wfuncname, UINT dispatch_type, UINT nargs, ...);*/
/*VARIANT IUnknownRefInvoke(void *pdisp, LPWSTR wfuncname, UINT dispatch_type, UINT nargs, ...);*/
VARIANT VarInvoke(VARIANT pdisp, LPWSTR wfuncname, UINT dispatch_type, UINT nargs, ...);
/*IDispatch * ObjInvoke(IDispatch *pdisp, LPWSTR wfuncname, UINT dispatch_type, UINT nargs, ...);*/
DISPID GetVtbl(void *obj, LPOLESTR membername);
#define PropertyLet(i) i;
#define PropertySet(i) i;
VARIANT_BOOL BoolIs(void *e1,void *e2);
#define UDTTypeOf(udt,guid) vbtTypeOf(&udt,guid)
#define UDTRefTypeOf(udt,guid) vbtTypeOf(udt,guid)
#define ObjTypeOf(obj,guid) vbtTypeOf(obj,guid)
#define ObjRefTypeOf(obj,guid) vbtTypeOf(*(obj),guid)
VARIANT_BOOL vbtTypeOf(void *e1,wchar_t *guid); /* fixme: optimize out - constant expression */
VARIANT_BOOL VarTypeOf(VARIANT e1,wchar_t *guid); /* distinguish between Obj and UDT??? */
VARIANT_BOOL VarRefTypeOf(VARIANT *e1,wchar_t *guid); /* distinguish between Obj and UDT??? */
VARIANT VarLike(VARIANT v1,VARIANT v2);
VARIANT_BOOL BoolLike(BSTR s1, BSTR s2);

/* LBound, UBound */
INT32 vbtVarLBound(VARIANT v,VARIANT nDim);
INT32 vbtVarUBound(VARIANT v,VARIANT nDim);

#define VarLBound(v,n) vbtVarLBound(v,n)
#define VarRefLBound(v,n) vbtVarLBound(*(v),n)

#define VarUBound(v,n) vbtVarUBound(v,n)
#define VarRefUBound(v,n) vbtVarUBound(*(v),n)

INT32 vbtSaLBound(SAFEARRAY *sa,VARIANT nDim);
INT32 vbtSaUBound(SAFEARRAY *sa,VARIANT nDim);

#define ByteSALBound(a,n) vbtSaLBound(a,n)
#define BoolSALBound(a,n) vbtSaLBound(a,n)
#define IntSALBound(a,n) vbtSaLBound(a,n)
#define LngSALBound(a,n) vbtSaLBound(a,n)
#define SngSALBound(a,n) vbtSaLBound(a,n)
#define DateSALBound(a,n) vbtSaLBound(a,n)
#define DblSALBound(a,n) vbtSaLBound(a,n)
#define CurSALBound(a,n) vbtSaLBound(a,n)
#define StrSALBound(a,n) vbtSaLBound(a,n)
#define ObjSALBound(a,n) vbtSaLBound(a,n)
#define VarSALBound(a,n) vbtSaLBound(a,n)
#define DecSALBound(a,n) vbtSaLBound(a,n)
#define UDTSALBound(a,n) vbtSaLBound(a,n)

#define ByteSAUBound(a,n) vbtSaUBound(a,n)
#define BoolSAUBound(a,n) vbtSaUBound(a,n)
#define IntSAUBound(a,n) vbtSaUBound(a,n)
#define LngSAUBound(a,n) vbtSaUBound(a,n)
#define SngSAUBound(a,n) vbtSaUBound(a,n)
#define DateSAUBound(a,n) vbtSaUBound(a,n)
#define DblSAUBound(a,n) vbtSaUBound(a,n)
#define CurSAUBound(a,n) vbtSaUBound(a,n)
#define StrSAUBound(a,n) vbtSaUBound(a,n)
#define ObjSAUBound(a,n) vbtSaUBound(a,n)
#define VarSAUBound(a,n) vbtSaUBound(a,n)
#define DecSAUBound(a,n) vbtSaUBound(a,n)
#define UDTSAUBound(a,n) vbtSaUBound(a,n)

/* void * matches any interface */
extern void *ObjNothing;
extern void *ObjRefNothing;
extern void *IUnknownNothing;
extern void *IUnknownRefNothing;
extern VARIANT VarNothing;
extern VARIANT *VarRefNothing;
IUnknown *vbtCreateInstance(LPOLESTR classid);
/*IDispatch *_IDispatchQI(This);*/
#define DeParend(a,b) a,b
#define LngVTable(Class,method,This,args) ((This)->lpVtbl->method(C_ADJUST_THIS(Class,This,method),DeParend args,&l),l);
/*#define New(o) (o)*/ /* temp - make ObjRef? put New flag in lv area? */
#ifdef NEVER
/*#undef LngVTable
#define LngVTable(iface,This,method,args) (This)->lpVtbl->method(C_ADJUST_THIS(iface,This,method))
	LngVTable(vbtTest__Class1,lv.c,msg1,(&b,1))*/
#define x(iface,This,method) xx((This)->lpVtbl->method,C_ADJUST_THIS(iface,This,method))
	LngVTable(x(vbtTest__Class1,lv.c,msg1),&b,1,&l);
#undef LngVTable
#define LngVTable(This,method) (This)->lpVtbl->method
	LngVTable(lv.c,msg1)(C_ADJUST_THIS(vbtTest__Class1,lv.c,msg1),&b,1,&l),l;
#undef LngVTable
#define LngVTable(This,method,args,result) (This)->lpVtbl->method args, result
#endif
#endif /* _INC_VBT */
