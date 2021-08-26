
#ifndef _INC_VBTCOM
#define _INC_VBTCOM

#ifdef _MSC_VER
#define UNICODE
#define COBJMACROS /* needed for IUnknown_Release macro */
#else /* _MSC_VER */
#ifdef SAG_COM
/* none */
#else /* SAG_COM */
/* grab #defines from SAG's makefile.incl - requires updating on every version */
#ifdef __linux__
#define CE_TLINUX86
#define SAG_COM 1
#define _RENTRANT
#define _WIN32
#define WIN32 100
#define _WIN32_WINNT 0x400
#define INC_OLE2
#define FLAT
#define UNICODE
#define WIN32_LEAN_AND_MEAN
#define _POSIX_MAPPED_FILES
#define SECURITY_ENABLED 1
#define COBJMACROS
#else /* __linux__ */
#ifdef sun
#define CE_TSUNSOL2
#error vbtcom.h -- missing sun defines
#endif /* sun */
#endif /* __linux__ */
#endif /* SAG_COM */
#endif /* _MSC_VER */

/*#define _WIN32_DCOM*/

#include <windows.h>
#include <ocidl.h> /* Connection Point defs */
#include <stdio.h>

#ifndef SAG_COM

#define INTERFACE_PROLOGUE(iface) BEGIN_INTERFACE
#define INTERFACE_PROLOGUE_(iface, baseiface) BEGIN_INTERFACE
#define STDMETHODEX( method, params ) STDMETHOD(method) params ;
#define STDMETHODEX_( type, method, params ) STDMETHOD_( type, method ) params ;
#define INTERFACE_EPILOGUE(iface)  END_INTERFACE
#define INTERFACE_EPILOGUE_(iface, baseiface)  END_INTERFACE

#define BEGIN_VTABLE {
#define VTABLE_ENTRY(method) method
#define END_VTABLE }

#define ADJUST_THIS( iface, This, method ) ( This )

#endif

#define ChkHR(v,name) HRESULT __hr; if (__hr=FAILED(v)) FailedHR(__hr,name)
#define NoChkHR(v,name) (v)
void FailedHR(HRESULT hr, LPWSTR name);
#undef ERROR
typedef long ERROR;
/* what to do with interfaces that aren't defined??? */
/*typedef void StdFormat_IDataFormatDisp,StdFormat_IStdDataFormatDisp,MSDATASRC_DataSource,DAO_Database,DAO_Recordset;*/
#endif /* _INC_VBTCOM */
